# -*- coding: utf-8 -*-
# synthDrivers/flexvoice.py
#
# FlexVoice 3.01 via fvwrap.dll (native wrapper)
#
# Goals:
# - Make Rate/Volume/Pitch appear in NVDA Settings Ring reliably (BestSpeech-style).
# - Restore IndexCommand behavior even if wrapper index events are missing:
#   We split text into segments at IndexCommand boundaries and emit indexes after
#   segment audio has fully played (player.sync / idle).
#
# Runtime model:
# - Single worker thread owns WavePlayer and ALL wrapper calls that can block (stop/read loops).
# - Main thread only enqueues commands and updates desired settings fields.

from __future__ import annotations

import os
import re
import ctypes
import threading
import queue
import time

# audioop is stdlib and implemented in C (fast) - useful for PCM silence analysis.
# It may be missing in some future Python builds; keep a soft fallback.
try:
	import audioop
except Exception:
	audioop = None

from dataclasses import dataclass
from collections import deque

from . import _flexvoice

import nvwave
import config
from logHandler import log
from synthDriverHandler import (
	SynthDriver as BaseSynthDriver,
	VoiceInfo,
	synthIndexReached,
	synthDoneSpeaking,
)
try:
	from synthDriverHandler import LanguageInfo
except Exception:
	LanguageInfo = None


def _dropVoiceListCaches(synth) -> None:
	"""
	Make the next availableVoices read hit the driver again.

	SynthDriver._get_availableVoices caches its result in the instance
	attribute ``_availableVoices`` (``_availableVariants`` for variants) -
	verified in NVDA's own synthDriverHandler; invalidateCache() does not
	clear those.

	The voice list no longer changes with language, so this is not needed for
	the language switch any more. It is kept for the config-restore path,
	where the set of shipped languages can differ from the last run.
	"""
	for attr in ("_availableVoices", "_availableVariants"):
		try:
			delattr(synth, attr)
		except AttributeError:
			pass
	try:
		synth.invalidateCache()
	except Exception:
		pass


def _scheduleVoiceSettingsRefresh() -> None:
	"""
	Ask an open NVDA settings dialog to re-read our settings.

	Only the *selection* of the Voice combo can be moved this way, never its
	items - but since the voice list is the same for every language, moving
	the selection is all that is required.

	The panel is found by walking the wx window tree rather than
	NVDASettingsDialog.catIdToInstanceMap: that map holds category panels, and
	the AutoSettingsMixin panel carrying the driver settings is nested inside
	one, so an isinstance check against the map's values silently matched
	nothing.

	Runs only where wx exists - NVDA's own process. Inside the 32-bit bridge
	host the import fails and this is a no-op.
	"""
	try:
		import wx
		from autoSettingsUtils.autoSettings import AutoSettings  # noqa: F401
		from gui.settingsDialogs import AutoSettingsMixin
	except Exception:
		return

	def _walk(win, depth=0):
		if depth > 8:
			return
		yield win
		try:
			children = win.GetChildren()
		except Exception:
			return
		for child in children:
			yield from _walk(child, depth + 1)

	def _do():
		try:
			for top in wx.GetTopLevelWindows():
				if not top.IsShown():
					continue
				for win in _walk(top):
					if not isinstance(win, AutoSettingsMixin):
						continue
					# Only OUR panel. BrailleSettingsSubPanel is an
					# AutoSettingsMixin too, and refreshing it would make the
					# braille display driver re-read its settings for no
					# reason - and would create controls for a "language"
					# setting it does not have.
					try:
						if getattr(win.getSettings(), "name", None) != SynthDriver.name:
							continue
					except Exception:
						continue
					try:
						win.updateDriverSettings(changedSetting="language")
					except Exception:
						log.debugWarning("FlexVoice: updateDriverSettings failed", exc_info=True)
		except Exception:
			log.debugWarning("FlexVoice: settings panel refresh failed", exc_info=True)

	try:
		wx.CallAfter(_do)
	except Exception:
		pass

from speech.commands import IndexCommand

try:
	from speech.commands import EndUtteranceCommand
except Exception:
	EndUtteranceCommand = None

# BestSpeech-style settings objects (for settings ring visibility).
# If unavailable (older NVDA), we fall back to BaseSynthDriver.*Setting().
try:
	from autoSettingsUtils.driverSetting import NumericDriverSetting
except Exception:
	NumericDriverSetting = None


# FlexVoice language IDs (MM_TTSAPI::Language, see ttsapi/FVLanguage.h)
LNG_ENGLISH = 0x0409
LNG_HUNGARIAN = 0x040E

# The languages this add-on ships data for.
#   nvdaLang  -> what NVDA calls it ("en" / "hu"); also the setting value
#   dirName   -> the data folder next to fvwrap.dll
#   engineId  -> the numeric id handed to fvwrap_create()
_LANGUAGES = (
	("en", "English", LNG_ENGLISH),
	("hu", "Hungarian", LNG_HUNGARIAN),
)
_LANG_BY_ID = {nvdaLang: (dirName, engineId) for nvdaLang, dirName, engineId in _LANGUAGES}
_DEFAULT_LANG = "en"

# Tim and Kim are NOT aliases of Tom and Julie, despite sharing their diphone
# banks (Tim.tav carries voiceDescr="Tom", Kim.tav carries voiceDescr="Julie").
# The voice shaping differs substantially -- Tim is pitch 105 at 163 wpm with
# tilt 0, Tom is pitch 95 at 137 wpm with tilt 31, and their equalizer curves
# differ in both Q and gain. Tim.tav is in fact byte-identical to Ben.tav apart
# from voiceDescr, i.e. Ben's voice driving Tom's bank, which is why Tim and Ben
# sound so alike. Both ship, so long-time users keep the voice they know.

# Wrapper stream item types (from _flexvoice)
FVWRAP_ITEM_NONE = _flexvoice.FVWRAP_ITEM_NONE
FVWRAP_ITEM_AUDIO = _flexvoice.FVWRAP_ITEM_AUDIO
FVWRAP_ITEM_INDEX = _flexvoice.FVWRAP_ITEM_INDEX
FVWRAP_ITEM_DONE = _flexvoice.FVWRAP_ITEM_DONE
FVWRAP_ITEM_ERROR = _flexvoice.FVWRAP_ITEM_ERROR


# --- Optional PCM silence trimming ---
#
# FlexVoice often adds leading/trailing silence around each engine request.
# During say-all, NVDA can end up sending many short requests (e.g. per line),
# which makes these silences audible as "gaps".
#
# We can shave off *small* amounts of near-zero PCM at the start/end of each
# segment to make line-to-line flow smoother.
#
# This is deliberately conservative and only enabled for indexed utterances
# (i.e. sequences containing IndexCommand), so normal speech isn't altered.
_TRIM_SILENCE_ON_INDEXED_UTTERANCES = True

# Max silence we will trim at the *start* of a segment (ms)
_TRIM_LEADING_SILENCE_MS = 180

# Hold back this much audio (ms) so we can safely trim trailing silence.
# Setting this to 0 disables trailing trimming (leading trim still works).
_TRIM_TAIL_HOLDBACK_MS = 350

# Max silence we will trim at the *end* of a segment (ms)
# Should be <= _TRIM_TAIL_HOLDBACK_MS.
_TRIM_TRAILING_SILENCE_MS = 350

# Keep a tiny tail (ms) even if it's silent, to avoid "hard" joins/clicks.
_TRIM_KEEP_TRAILING_MS = 10

# Silence threshold in int16 absolute amplitude.
# 0 is only true digital silence; 30-80 is a good "near-zero" range.
_TRIM_SILENCE_THRESHOLD = 40

# Debug logging (NVDA log) for trimmed ms amounts.
_TRIM_DEBUG_LOG = False


def _ms_to_frames(sr: int, ms: int) -> int:
	if sr <= 0 or ms <= 0:
		return 0
	# integer math, rounded
	return int((sr * ms) // 1000)


def _trim_leading_silence_pcm16(
	data: bytes,
	channels: int,
	thresh: int,
	max_frames: int,
) -> tuple[bytes, int, bool]:
	"""Trim leading silence from 16-bit PCM.

	Returns: (trimmedBytes, framesTrimmed, hitNonSilent)
	"""
	if not data or channels <= 0 or max_frames <= 0:
		return data, 0, False

	frame_bytes = channels * 2
	frame_count = len(data) // frame_bytes
	if frame_count <= 0:
		return b"", 0, False

	max_frames = min(max_frames, frame_count)
	# Ensure we only view whole frames
	view_bytes = max_frames * frame_bytes
	mv = memoryview(data)[:view_bytes]
	if (len(mv) % 2) != 0:
		mv = mv[:-1]
	try:
		samples = mv.cast('h')
	except Exception:
		# Shouldn't happen on NVDA's Windows Python, but be safe.
		return data, 0, False

	frames_trimmed = 0
	for f in range(max_frames):
		base = f * channels
		# Any channel above threshold counts as non-silence.
		for ch in range(channels):
			if abs(int(samples[base + ch])) > thresh:
				# Found first non-silent frame.
				start_byte = f * frame_bytes
				return data[start_byte:], frames_trimmed, True
		frames_trimmed += 1

	# All silent in the scanned window.
	return data[frames_trimmed * frame_bytes:], frames_trimmed, False


def _trim_trailing_silence_pcm16(
	data: bytes,
	channels: int,
	thresh: int,
	max_frames: int,
	keep_frames: int,
) -> bytes:
	"""Trim trailing silence from 16-bit PCM (conservative).

	Only trims within the last max_frames. Keeps keep_frames after the last
	non-silent frame to avoid overly tight joins.
	"""
	if not data or channels <= 0 or max_frames <= 0:
		return data

	frame_bytes = channels * 2
	total_frames = len(data) // frame_bytes
	if total_frames <= 0:
		return b""

	max_frames = min(max_frames, total_frames)
	keep_frames = max(0, min(keep_frames, total_frames))

	# View whole frames
	mv = memoryview(data)[: total_frames * frame_bytes]
	if (len(mv) % 2) != 0:
		mv = mv[:-1]
	try:
		samples = mv.cast('h')
	except Exception:
		return data

	scan_start = total_frames - max_frames
	last_non_silent = None
	for f in range(total_frames - 1, scan_start - 1, -1):
		base = f * channels
		for ch in range(channels):
			if abs(int(samples[base + ch])) > thresh:
				last_non_silent = f
				break
		if last_non_silent is not None:
			break

	if last_non_silent is None:
		# Entire scan window is silent.
		new_end = scan_start
	else:
		new_end = last_non_silent + 1

	new_end = min(total_frames, new_end + keep_frames)
	return data[: new_end * frame_bytes]


def _clampPercent(v: int) -> int:
	try:
		v = int(v)
	except Exception:
		return 0
	if v < 0:
		return 0
	if v > 100:
		return 100
	return v


def _findLanguageDir(base_dir: str, dirName: str) -> str:
	"""
	Find the folder containing 'VoiceList.tvl' for one language.
	"""
	addon_root = os.path.abspath(os.path.join(base_dir, os.pardir))
	candidates = [
		os.path.join(base_dir, dirName),
		os.path.join(addon_root, dirName),
	]
	for p in candidates:
		if os.path.isdir(p) and os.path.isfile(os.path.join(p, "VoiceList.tvl")):
			return p
	return ""


def _findLanguageDirs(base_dir: str) -> dict[str, str]:
	"""
	Map nvda language id -> data folder, for every language actually present.
	"""
	out: dict[str, str] = {}
	for nvdaLang, dirName, _engineId in _LANGUAGES:
		d = _findLanguageDir(base_dir, dirName)
		if d:
			out[nvdaLang] = d
	return out


def _findEnglishDir(base_dir: str) -> str:
	"""
	Back-compat helper (used by check()). Also tolerates the old flat layout
	where VoiceList.tvl sat directly beside fvwrap.dll.
	"""
	d = _findLanguageDir(base_dir, "English")
	if d:
		return d
	addon_root = os.path.abspath(os.path.join(base_dir, os.pardir))
	for p in (base_dir, addon_root):
		if os.path.isdir(p) and os.path.isfile(os.path.join(p, "VoiceList.tvl")):
			return p
	return ""


def _parseVoiceList(voiceListPath: str, langDir: str) -> dict[str, str]:
	out: dict[str, str] = {}
	try:
		with open(voiceListPath, "r", encoding="mbcs", errors="replace") as f:
			for line in f:
				line = line.strip()
				if not line or line.startswith(";"):
					continue
				m = re.match(r'^\s*"([^"]+)"\s*"([^"]+)"', line)
				if not m:
					continue
				name = m.group(1).strip()
				rel = m.group(2).strip()
				if not name or not rel:
					continue
				full = os.path.normpath(rel if os.path.isabs(rel) else os.path.join(langDir, rel))
				# Stock VoiceList.tvl files list voices that the demo does not ship
				# (Bill, Jack, Jill, Kit, Julius). Offering them would hand the
				# engine a missing .tav and fail at create() time.
				if not os.path.isfile(full):
					continue
				out[name] = full
	except Exception:
		log.error("FlexVoice(wrapper): failed to parse VoiceList.tvl", exc_info=True)
	return out


# --- Text Sanitization
_TRANSLATE = {
	"\u00a0": " ",
	"\u2018": "'", "\u2019": "'",
	"\u201c": '"', "\u201d": '"',
	"\u2013": "-", "\u2014": "-",
	"\u2026": "...",
}


def _sanitizeTextForEngine(s: str, keepNonLatin1: bool = False) -> str:
	"""
	Strip control characters and normalize typographic punctuation.

	For English the wrapper converts UTF-8 -> Latin-1-ish, so dropping anything
	above 0xFF here avoids surprises.

	For Hungarian that dropping is actively harmful: o-double-acute (U+0151) and
	u-double-acute (U+0171) are both above 0xFF, so the blanket rule replaced
	them with spaces and turned "tükörfúrógép" into "t kr...". The wrapper's
	Hungarian path converts to CP1250 and approximates anything it cannot map,
	so pass those characters through untouched.
	"""
	if not s:
		return ""
	for k, v in _TRANSLATE.items():
		s = s.replace(k, v)

	out = []
	for ch in s:
		o = ord(ch)
		if ch in ("\r", "\n", "\t"):
			# FlexVoice tends to insert a noticeable pause on hard line breaks.
			# Converting them to spaces keeps flow smoother for say-all.
			out.append(" ")
			continue
		if o < 0x20 or o == 0x7F:
			out.append(" ")
			continue
		if o > 0xFF and not keepNonLatin1:
			out.append(" ")
			continue
		out.append(ch)
	return "".join(out)


@dataclass
class _Seg:
	text: str
	idxAfter: list[int]


@dataclass
class _Utt:
	uttId: int
	token: int
	leadingIndexes: list[int]
	segments: list[_Seg]
	expectedIndexes: list[int]


_CMD_SPEAK = 1
_CMD_CANCEL = 2
_CMD_RECREATE = 3
_CMD_PAUSE = 4
_CMD_TERMINATE = 5


class SynthDriver(BaseSynthDriver):
	name = "flexvoice"
	description = "Mindmaker FlexVoice 3"

	# NOTE: 'language' is deliberately NOT a class attribute any more. It is a
	# real setting now (see _get_language/_set_language), which also keeps
	# NVDA's synth.language honest about what is actually being spoken.

	# LanguageSetting() is documented for every supported NVDA version, but stay
	# defensive: a missing factory must not stop the synth from loading at all.
	try:
		_languageSetting = (BaseSynthDriver.LanguageSetting(),)
	except Exception:
		_languageSetting = ()

	# IMPORTANT:
	# We explicitly declare rate/volume/pitch as NumericDriverSetting with
	# availableInSettingsRing=True to match BestSpeech’s pattern and to avoid
	# NVDA-version-dependent defaults.
	if NumericDriverSetting is not None:
		supportedSettings = _languageSetting + (
			BaseSynthDriver.VoiceSetting(),
			NumericDriverSetting("rate", "&Rate", defaultVal=50, availableInSettingsRing=True),
			NumericDriverSetting("pitch", "&Pitch", defaultVal=50, availableInSettingsRing=True),
			NumericDriverSetting("volume", "&Volume", defaultVal=100, availableInSettingsRing=True),
			# Suggested by FlexVoice's author: the voices sound more natural
			# with the high frequencies lifted. 50 is the voice exactly as it
			# was authored, so nobody's voice changes until they ask.
			NumericDriverSetting("clarity", "&Clarity", defaultVal=50, availableInSettingsRing=True),
		)
	else:
		# Fallback for older NVDA: classic supported settings.
		supportedSettings = _languageSetting + (
			BaseSynthDriver.VoiceSetting(),
			BaseSynthDriver.RateSetting(),
			BaseSynthDriver.PitchSetting(),
			BaseSynthDriver.VolumeSetting(),
		)

	supportedCommands = {IndexCommand}
	supportedNotifications = {synthIndexReached, synthDoneSpeaking}

	@classmethod
	def check(cls):
		base = os.path.dirname(__file__)
		if not os.path.isfile(os.path.join(base, "fvwrap.dll")):
			addon_root = os.path.abspath(os.path.join(base, os.pardir))
			if not os.path.isfile(os.path.join(addon_root, "fvwrap.dll")):
				return False
		if not _findLanguageDirs(base) and not _findEnglishDir(base):
			return False
		if _flexvoice.IS_64BIT:
			host_exe = os.path.join(base, "flexvoice_host32.exe")
			host_script = os.path.join(base, "host_flexvoice32.py")
			if not os.path.exists(host_exe) and not os.path.exists(host_script):
				return False
		return True

	def __init__(self):
		# Ensure config.pre_configSave exists (bridge host compat)
		if not hasattr(config, 'pre_configSave'):
			import extensionPoints
			config.pre_configSave = extensionPoints.Action()
		super().__init__()

		# 1) Threading state
		self._cancelTokenLock = threading.Lock()
		self._cancelToken = 0

		self._uttCounterLock = threading.Lock()
		self._uttCounter = 0

		self._cmdQ: queue.Queue[tuple[int, object | None]] = queue.Queue()
		self._shutdown = threading.Event()

		self._needsRecreateLock = threading.Lock()
		self._needsRecreate = False

		# 2) Defaults (percent)
		self._rate = 50
		self._volume = 100
		self._pitch = 50
		self._clarity = 50

		# Track applied settings to avoid churn
		self._appliedRate = None
		self._appliedVol = None
		self._appliedPitch = None
		self._appliedClarity = None
		self._hasPitch = False
		self._hasClarity = False

		# 3) Runtime handles / audio format
		self._sr = 16000
		self._bits = 16
		self._channels = 1
		self._player = None
		self._playerFormat = None
		self._outputDevice = self._getOutputDevice()

		# 4) File system paths
		self._baseDir = os.path.dirname(__file__)
		self._addonRoot = os.path.abspath(os.path.join(self._baseDir, os.pardir))

		p1 = os.path.join(self._baseDir, "fvwrap.dll")
		p2 = os.path.join(self._addonRoot, "fvwrap.dll")
		self._wrapperPath = p1 if os.path.isfile(p1) else p2
		if not os.path.isfile(self._wrapperPath):
			raise RuntimeError("FlexVoice: fvwrap.dll not found")

		self._langDirs = _findLanguageDirs(self._baseDir)
		if not self._langDirs:
			# Tolerate the old flat layout (VoiceList.tvl beside fvwrap.dll).
			legacy = _findEnglishDir(self._baseDir)
			if not legacy:
				raise RuntimeError("FlexVoice: no language data found")
			self._langDirs = {"en": legacy}

		# One voice list spanning every language, built once and never rebuilt.
		# An open NVDA combo box cannot have its items replaced (see
		# _getAvailableVoices), so the list has to be stable; the language of a
		# voice travels with it instead.
		self._voiceLang: dict[str, str] = {}   # voice id -> nvda language
		self._voiceMap: dict[str, str] = {}    # voice id -> .tav path
		for nvdaLang, langDir in self._langDirs.items():
			for vid, path in _parseVoiceList(os.path.join(langDir, "VoiceList.tvl"), langDir).items():
				self._voiceMap[vid] = path
				self._voiceLang[vid] = nvdaLang
		self._voiceIds = sorted(self._voiceMap.keys(), key=lambda s: s.lower())
		if not self._voiceIds:
			raise RuntimeError("FlexVoice: no voices found")

		self._language = _DEFAULT_LANG if _DEFAULT_LANG in self._langDirs else sorted(self._langDirs)[0]
		self._voice = self._defaultVoiceFor(self._language)

		# 5) Engine init (via _flexvoice abstraction)
		self._engine = _flexvoice.create_engine(self._wrapperPath, [self._baseDir, self._addonRoot])
		self._createEngineOrDie()

		# 6) Player init
		self._ensurePlayer()

		# 7) Worker thread
		self._pendingUtts = deque()
		self._worker = threading.Thread(target=self._workerLoop, name="flexvoiceWorker", daemon=True)
		self._worker.start()

	# ---------------- NVDA API ----------------
	def terminate(self):
		# Hardening: cancel -> terminate -> only destroy wrapper if worker is stopped.
		self._shutdown.set()

		try:
			with self._cancelTokenLock:
				self._cancelToken += 1
				tok = self._cancelToken
			self._cmdQ.put((_CMD_CANCEL, tok))
		except Exception:
			pass

		self._cmdQ.put((_CMD_TERMINATE, None))

		try:
			self._worker.join(2.0)
		except Exception:
			pass

		if getattr(self, "_worker", None) is not None and self._worker.is_alive():
			log.error("FlexVoice(wrapper): worker did not terminate in time; skipping fvwrap_destroy to avoid crash")
			return

		try:
			if self._player:
				self._player.stop()
		except Exception:
			pass
		self._player = None

		try:
			self._engine.destroy()
		except Exception:
			pass

	def cancel(self):
		with self._cancelTokenLock:
			self._cancelToken += 1
			tok = self._cancelToken
		self._cmdQ.put((_CMD_CANCEL, tok))

	def speak(self, speechSequence):
		utt = self._buildUtterance(speechSequence)
		self._cmdQ.put((_CMD_SPEAK, utt))

	def pause(self, switch):
		self._cmdQ.put((_CMD_PAUSE, bool(switch)))

	# ---------------- Settings ----------------
	# Voice picked when a language is selected for the first time.
	_PREFERRED_VOICE = {"en": "Tom", "hu": "Zita"}

	def _defaultVoiceFor(self, nvdaLang: str) -> str:
		"""First choice of voice when *nvdaLang* becomes active."""
		want = self._PREFERRED_VOICE.get(nvdaLang)
		if want in self._voiceMap and self._voiceLang.get(want) == nvdaLang:
			return want
		for vid in self._voiceIds:
			if self._voiceLang.get(vid) == nvdaLang:
				return vid
		return self._voiceIds[0]

	def _getAvailableVoices(self):
		"""
		Every voice of every language, always, in one stable list.

		This is deliberate, and it is the only arrangement that works. An open
		combo box in NVDA's settings cannot have its items replaced: read from
		gui/settingsDialogs, _makeStringSettingControl builds the items once and
		stores the VoiceInfo list on the panel as _voices, _updateValueForControl
		only ever calls SetSelection, and the EVT_CHOICE handler maps a click
		through panel._voices[GetSelection()]. Filtering this list per language
		therefore did not merely look stale - clicking the visibly-Hungarian
		list resolved to an English VoiceInfo and bounced the language back.

		With one fixed list none of that can happen: NVDA only ever has to move
		the selection, which it does correctly by itself. Each VoiceInfo carries
		its own language, and picking a voice switches the engine to it. This is
		also how eSpeak and OneCore present themselves, and how TGSpeechBox
		sidesteps the same limitation.
		"""
		return {
			vid: VoiceInfo(vid, vid, self._voiceLang.get(vid, self._language))
			for vid in self._voiceIds
		}

	def _get_availableLanguages(self):
		# LanguageInfo derives the display name from the locale, so the combo
		# reads "English" / "Hungarian" in the user's own language.
		out = {}
		for nvdaLang, _dirName, _engineId in _LANGUAGES:
			if nvdaLang not in self._langDirs:
				continue
			if LanguageInfo is not None:
				out[nvdaLang] = LanguageInfo(nvdaLang)
			else:
				out[nvdaLang] = VoiceInfo(nvdaLang, nvdaLang, nvdaLang)
		return out

	def _get_language(self):
		return self._language

	def _set_language(self, val):
		if not val or val == self._language:
			return
		if val not in self._langDirs:
			log.warning(f"FlexVoice: no data for language {val!r}, ignoring")
			return
		self._language = val
		# Move to a voice that actually speaks it. The voice list itself does
		# not change, so NVDA only has to move the combo's selection - which
		# StringDriverSettingChanger already asks it to do after this returns.
		self._voice = self._defaultVoiceFor(val)
		self._restartEngine()

	def _get_voice(self):
		return self._voice

	def _set_voice(self, val):
		if val == "Default" or val not in self._voiceMap:
			# Configs from older builds can name "Default", or a voice that is
			# no longer shipped. Fall back rather than refusing to speak.
			if val not in self._voiceMap:
				val = self._defaultVoiceFor(self._language)
		if val == self._voice:
			return
		self._voice = val
		# The language travels with the voice: choosing Zita means Hungarian.
		self._language = self._voiceLang.get(val, self._language)
		self._restartEngine()

	def _restartEngine(self):
		self._needsRecreate = True
		self.cancel()
		self._cmdQ.put((_CMD_RECREATE, None))

	# Ensure these match the setting ids "rate", "pitch", "volume"
	def _get_rate(self):
		return int(self._rate)

	def _set_rate(self, val):
		self._rate = _clampPercent(val)

	def _get_volume(self):
		return int(self._volume)

	def _set_volume(self, val):
		self._volume = _clampPercent(val)

	def _get_pitch(self):
		return int(self._pitch)

	def _set_pitch(self, val):
		self._pitch = _clampPercent(val)

	def _get_clarity(self):
		return int(self._clarity)

	def _set_clarity(self, val):
		self._clarity = _clampPercent(val)

	# ---------------- Internals ----------------
	def _speakerPathForVoice(self, voiceName: str) -> str:
		p = self._voiceMap.get(voiceName)
		if p and os.path.isfile(p):
			return p
		langDir = self._langDirs.get(self._language) or ""
		# Both spellings occur in the wild: English ships default.tav, the
		# Hungarian data ships Default.tav.
		for name in ("default.tav", "Default.tav"):
			cand = os.path.join(langDir, name)
			if os.path.isfile(cand):
				return cand
		return os.path.join(langDir, "default.tav")

	def _createEngineOrDie(self):
		# The voice decides the language: they can only disagree transiently,
		# and the voice is what the user actually hears.
		self._language = self._voiceLang.get(self._voice, self._language)
		langDir = self._langDirs.get(self._language) or ""
		# EngineFactory wants the folder that CONTAINS the per-language folders;
		# it appends the language itself. Handing it a single language folder
		# happens to work when English is the only one present, but starts
		# picking the wrong data as soon as a sibling language exists.
		dataRoot = os.path.dirname(langDir)
		speakerPath = self._speakerPathForVoice(self._voice)
		engineLang = _LANG_BY_ID.get(self._language, ("", LNG_ENGLISH))[1]

		dataCandidates = [
			(dataRoot, "dataPath=data root"),
			(langDir, "dataPath=language dir"),
			(None, "dataPath=NULL"),
		]

		tryParams = [(16000, 16), (11025, 16), (8000, 16)]

		for dataPath, label in dataCandidates:
			# Set FLEXVOICE_DATA env var (used by engine on 32-bit)
			if not _flexvoice.IS_64BIT:
				if dataPath:
					os.environ["FLEXVOICE_DATA"] = dataPath
				else:
					os.environ["FLEXVOICE_DATA"] = dataRoot or langDir

			for sr, bits in tryParams:
				try:
					ok = self._engine.create(dataPath, speakerPath,
											 int(engineLang), int(sr), int(bits))
					if ok:
						self._sr = sr
						self._bits = bits
						self._channels = 1
						self._hasPitch = self._engine.has_pitch
						self._hasClarity = getattr(self._engine, "has_clarity", False)

						self._appliedRate = None
						self._appliedVol = None
						self._appliedPitch = None
						self._appliedClarity = None

						log.info(f"FlexVoice: created engine ({label}, {sr}Hz)")
						return
				except Exception:
					pass

		raise RuntimeError("FlexVoice: Failed to create engine")

	def _getOutputDevice(self):
		try:
			return config.conf["audio"]["outputDevice"]
		except Exception:
			return config.conf["speech"]["outputDevice"]

	def _ensurePlayer(self):
		fmt = (self._channels, self._sr, self._bits)
		if self._player and self._playerFormat == fmt:
			return

		try:
			if self._player:
				self._player.stop()
		except Exception:
			pass

		try:
			self._player = nvwave.WavePlayer(fmt[0], fmt[1], fmt[2], outputDevice=self._outputDevice)
		except Exception:
			self._player = nvwave.WavePlayer(fmt[0], fmt[1], fmt[2])

		self._playerFormat = fmt

	def _playerFeed(self, data: bytes):
		if not data or not self._player:
			return
		try:
			try:
				self._player.feed(data, len(data))
			except TypeError:
				self._player.feed(data)
		except Exception:
			log.error("FlexVoice(wrapper): WavePlayer.feed failed", exc_info=True)

	def _playerStop(self):
		try:
			if self._player:
				self._player.stop()
		except Exception:
			pass

	def _playerPause(self, pause: bool):
		try:
			if self._player:
				self._player.pause(bool(pause))
		except Exception:
			pass

	def _playerSync(self):
		if not self._player:
			return
		sync = getattr(self._player, "sync", None)
		if callable(sync):
			try:
				sync()
				return
			except Exception:
				pass
		try:
			self._player.idle()
		except Exception:
			pass

	def _nextUtterId(self) -> int:
		with self._uttCounterLock:
			self._uttCounter += 1
			return self._uttCounter

	def _getCancelToken(self) -> int:
		with self._cancelTokenLock:
			return int(self._cancelToken)

	def _buildUtterance(self, speechSequence) -> _Utt:
		"""
		Build segments split by IndexCommand boundaries.

		- leadingIndexes: indexes that occur before any speakable text.
		- segments: list of text segments; each segment has idxAfter indexes that should fire
		  once that segment's audio has finished playing.
		- expectedIndexes: all indexes in the utterance (for fallback / dedupe).
		"""
		uttId = self._nextUtterId()
		token = self._getCancelToken()

		leading: list[int] = []
		segments: list[_Seg] = []
		expected: list[int] = []

		textBuf: list[str] = []

		def flushTextAsSegment():
			nonlocal textBuf
			if not textBuf:
				return False
			raw = "".join(textBuf)
			textBuf = []
			s = _sanitizeTextForEngine(raw, keepNonLatin1=(self._language != "en"))
			# Drop whitespace-only segments (prevents "silent segments" causing weird timing)
			if not s or not s.strip():
				return False
			segments.append(_Seg(text=s, idxAfter=[]))
			return True

		for item in speechSequence:
			if isinstance(item, str):
				textBuf.append(item)
				continue

			if isinstance(item, IndexCommand):
				idx = int(item.index)
				expected.append(idx)

				# If we haven't produced any text segment yet and we have no buffered text,
				# treat it as leading.
				if not segments and not textBuf:
					leading.append(idx)
					continue

				# If we currently have buffered text, finalize it into a segment first.
				if textBuf:
					flushTextAsSegment()

				# Attach index to the last segment if possible, else leading.
				if segments:
					segments[-1].idxAfter.append(idx)
				else:
					leading.append(idx)
				continue

			if EndUtteranceCommand is not None and isinstance(item, EndUtteranceCommand):
				break

			# Ignore other command types (language switches, etc.) for now.

		# Final trailing text
		if textBuf:
			flushTextAsSegment()

		# If there is no speakable text at all, everything is effectively "leading".
		if not segments and expected and not leading:
			leading = list(expected)

		return _Utt(uttId=uttId, token=token, leadingIndexes=leading, segments=segments, expectedIndexes=expected)

	# ---------------- WORKER LOGIC ----------------
	def _hardStop(self):
		self._playerStop()
		try:
			self._engine.stop()
		except Exception:
			pass
		self._flushWrapperOutput(40)

	def _applySettingsIfChanged(self):
		if not self._engine.is_created:
			return

		rate = _clampPercent(self._rate)
		vol = _clampPercent(self._volume)
		pit = _clampPercent(self._pitch)
		clr = _clampPercent(self._clarity)

		if self._appliedRate != rate:
			try:
				self._engine.set_rate_percent(int(rate))
				self._appliedRate = rate
			except Exception:
				pass

		if self._appliedVol != vol:
			try:
				self._engine.set_volume_percent(int(vol))
				self._appliedVol = vol
			except Exception:
				pass

		if self._hasPitch and self._appliedPitch != pit:
			try:
				self._engine.set_pitch_percent(int(pit))
				self._appliedPitch = pit
			except Exception:
				pass

		if self._hasClarity and self._appliedClarity != clr:
			try:
				self._engine.set_clarity_percent(int(clr))
				self._appliedClarity = clr
			except Exception:
				pass

	def _flushWrapperOutput(self, maxMs: int = 40):
		if not self._engine.is_created:
			return
		self._engine.flush_output(maxMs)

	def _runWrapperTextSegment(self, text: str, tokenSnapshot: int, sentIndexes: set[int], trimSilence: bool = False) -> bool:
		"""
		Speak one plain text segment through the wrapper, draining audio until DONE.

		Returns False if canceled or shutdown.
		"""
		if not self._engine.is_created:
			return False

		self._applySettingsIfChanged()

		try:
			self._engine.begin()
			b = (text or "").encode("utf-8", "replace")
			if b:
				self._engine.add_text_utf8(b)
			log.debug("FlexVoice(wrapper): committing text segment (%d bytes)", len(b) if b else 0)
			rc = self._engine.commit(1)
			if rc != 0:
				log.warning("FlexVoice(wrapper): commit returned %d", rc)
				return False
			log.debug("FlexVoice(wrapper): commit ok, entering read loop")
		except Exception:
			log.error("FlexVoice(wrapper): begin/add/commit failed", exc_info=True)
			return False

		# Optional (say-all) silence trimming state.
		doTrim = bool(trimSilence) and int(self._bits) == 16 and int(self._channels) > 0
		leadFramesLeft = _ms_to_frames(self._sr, _TRIM_LEADING_SILENCE_MS) if doTrim else 0
		tailHoldMs = int(_TRIM_TAIL_HOLDBACK_MS) if doTrim else 0
		trailTrimFrames = _ms_to_frames(self._sr, min(_TRIM_TRAILING_SILENCE_MS, _TRIM_TAIL_HOLDBACK_MS)) if (doTrim and tailHoldMs > 0) else 0
		keepTrailFrames = _ms_to_frames(self._sr, _TRIM_KEEP_TRAILING_MS) if doTrim else 0
		thresh = int(_TRIM_SILENCE_THRESHOLD) if doTrim else 0

		frameBytes = int(self._channels) * 2
		bytesPerSec = int(self._sr) * frameBytes if self._sr else 0
		tailHoldBytes = ((bytesPerSec * tailHoldMs) // 1000) if (doTrim and bytesPerSec > 0 and tailHoldMs > 0) else 0
		# Keep alignment.
		if tailHoldBytes and frameBytes:
			tailHoldBytes -= (tailHoldBytes % frameBytes)

		tailBuf = bytearray()  # last tailHoldBytes we haven't fed yet
		trimmedLeadFrames = 0
		trimmedTrailFrames = 0

		seenDone = False
		lastActivity = time.time()

		while not self._shutdown.is_set():
			if tokenSnapshot != self._getCancelToken():
				return False

			tp, val, chunk = self._engine.read()

			if tp == FVWRAP_ITEM_AUDIO:
				if chunk:

					# 1) Trim leading near-zero PCM (only at the very start of the segment).
					if doTrim and leadFramesLeft > 0:
						beforeLen = len(chunk)
						chunk, framesTrimmed, hitNonSilent = _trim_leading_silence_pcm16(
							chunk, int(self._channels), thresh, leadFramesLeft
						)
						leadFramesLeft -= framesTrimmed
						trimmedLeadFrames += framesTrimmed
						if hitNonSilent:
							leadFramesLeft = 0
						# If we trimmed everything and still haven't hit non-silence, this
						# chunk may become empty; just wait for the next chunk.
						if not chunk:
							lastActivity = time.time()
							continue

					# 2) Hold back a small tail so we can trim trailing silence safely.
					if doTrim and tailHoldBytes > 0:
						tailBuf.extend(chunk)
						if len(tailBuf) > tailHoldBytes:
							emit = bytes(tailBuf[:-tailHoldBytes])
							if emit:
								self._playerFeed(emit)
								lastActivity = time.time()
							del tailBuf[:-tailHoldBytes]
					else:
						self._playerFeed(chunk)
						lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_INDEX:
				# If wrapper emits any indexes (some builds do), pass them through
				# but dedupe with driver-emitted ones.
				idx = val
				if idx not in sentIndexes:
					sentIndexes.add(idx)
					try:
						synthIndexReached.notify(synth=self, index=idx)
					except Exception:
						pass
				lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_DONE:
				seenDone = True
				lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_ERROR:
				# Treat as DONE-ish: stop waiting forever, but still flush/sync audio.
				seenDone = True
				lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_NONE:
				# Give a tiny tail window for late audio, then return.
				# IMPORTANT: Do NOT sync the WavePlayer here. Syncing drains the playback buffer,
				# which forces a gap between chunks (very noticeable during say-all).
				# By returning as soon as we've *queued* all audio, we allow natural pipelining.
				if seenDone and (time.time() - lastActivity > 0.05):
					# Flush any held-back tail audio, optionally trimming trailing silence.
					if tailBuf:
						out = bytes(tailBuf)
						if doTrim and trailTrimFrames > 0:
							beforeFrames = len(out) // frameBytes if frameBytes else 0
							out = _trim_trailing_silence_pcm16(
								out, int(self._channels), thresh, trailTrimFrames, keepTrailFrames
							)
							afterFrames = len(out) // frameBytes if frameBytes else 0
							trimmedTrailFrames = max(0, beforeFrames - afterFrames)
						if out:
							self._playerFeed(out)
							tailBuf.clear()

						if doTrim and _TRIM_DEBUG_LOG:
							try:
								leadMs = int((trimmedLeadFrames * 1000) // int(self._sr)) if self._sr else 0
								trailMs = int((trimmedTrailFrames * 1000) // int(self._sr)) if self._sr else 0
								log.debug(f"FlexVoice(trim): lead={leadMs}ms trail={trailMs}ms")
							except Exception:
								pass
					return True
				time.sleep(0.002)
				continue

		return False

	def _runWrapperUtteranceComposite(self, utt: _Utt, tokenSnapshot: int, sentIndexes: set[int], trimSilence: bool = False) -> bool:
		"""Run a single wrapper commit containing multiple text chunks and explicit indexes.

		This reduces per-chunk engine resets during say-all, which helps avoid the
		"gap" FlexVoice can insert between short requests.
		"""
		if not self._engine.is_created:
			return False

		self._applySettingsIfChanged()

		try:
			self._engine.begin()

			for seg in utt.segments:
				b = (seg.text or "").encode("utf-8", "replace")
				if b:
					self._engine.add_text_utf8(b)

				# Tell the wrapper where NVDA index boundaries are.
				for idx in seg.idxAfter:
					try:
						self._engine.add_index(int(idx))
					except Exception:
						pass

			log.debug("FlexVoice(wrapper): committing composite (%d segs)", len(utt.segments))
			rc = self._engine.commit(1)
			if rc != 0:
				log.warning("FlexVoice(wrapper): composite commit returned %d", rc)
				return False
			log.debug("FlexVoice(wrapper): composite commit ok, entering read loop")
		except Exception:
			log.error("FlexVoice(wrapper): composite begin/add/commit failed", exc_info=True)
			return False

		# Optional (say-all) silence trimming state.
		doTrim = bool(trimSilence) and int(self._bits) == 16 and int(self._channels) > 0
		leadFramesLeft = _ms_to_frames(self._sr, _TRIM_LEADING_SILENCE_MS) if doTrim else 0
		tailHoldMs = int(_TRIM_TAIL_HOLDBACK_MS) if doTrim else 0
		trailTrimFrames = _ms_to_frames(self._sr, min(_TRIM_TRAILING_SILENCE_MS, _TRIM_TAIL_HOLDBACK_MS)) if (doTrim and tailHoldMs > 0) else 0
		keepTrailFrames = _ms_to_frames(self._sr, _TRIM_KEEP_TRAILING_MS) if doTrim else 0
		thresh = int(_TRIM_SILENCE_THRESHOLD) if doTrim else 0

		frameBytes = int(self._channels) * 2
		bytesPerSec = int(self._sr) * frameBytes if self._sr else 0
		tailHoldBytes = ((bytesPerSec * tailHoldMs) // 1000) if (doTrim and bytesPerSec > 0 and tailHoldMs > 0) else 0
		# Keep alignment.
		if tailHoldBytes and frameBytes:
			tailHoldBytes -= (tailHoldBytes % frameBytes)

		tailBuf = bytearray()
		trimmedLeadFrames = 0
		trimmedTrailFrames = 0

		seenDone = False
		lastActivity = time.time()

		while not self._shutdown.is_set():
			if tokenSnapshot != self._getCancelToken():
				return False

			tp, val, chunk = self._engine.read()

			if tp == FVWRAP_ITEM_AUDIO:
				if chunk:

					# 1) Trim leading near-zero PCM (only at the very start of the utterance).
					if doTrim and leadFramesLeft > 0:
						chunk, framesTrimmed, hitNonSilent = _trim_leading_silence_pcm16(
							chunk, int(self._channels), thresh, leadFramesLeft
						)
						leadFramesLeft -= framesTrimmed
						trimmedLeadFrames += framesTrimmed
						if hitNonSilent:
							leadFramesLeft = 0
						if not chunk:
							lastActivity = time.time()
							continue

					# 2) Hold back a small tail so we can trim trailing silence safely.
					if doTrim and tailHoldBytes > 0:
						tailBuf.extend(chunk)
						if len(tailBuf) > tailHoldBytes:
							emit = bytes(tailBuf[:-tailHoldBytes])
							if emit:
								self._playerFeed(emit)
								lastActivity = time.time()
							del tailBuf[:-tailHoldBytes]
					else:
						self._playerFeed(chunk)
						lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_INDEX:
				idx = val
				if idx not in sentIndexes:
					sentIndexes.add(idx)
					try:
						synthIndexReached.notify(synth=self, index=idx)
					except Exception:
						pass
				lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_DONE:
				seenDone = True
				lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_ERROR:
				seenDone = True
				lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_NONE:
				# Same policy as _runWrapperTextSegment: don't sync here for indexed utterances.
				if seenDone and (time.time() - lastActivity > 0.05):
					if tailBuf:
						out = bytes(tailBuf)
						if doTrim and trailTrimFrames > 0 and frameBytes:
							beforeFrames = len(out) // frameBytes
							out = _trim_trailing_silence_pcm16(
								out, int(self._channels), thresh, trailTrimFrames, keepTrailFrames
							)
							afterFrames = len(out) // frameBytes
							trimmedTrailFrames = max(0, beforeFrames - afterFrames)
						if out:
							self._playerFeed(out)
							tailBuf.clear()

						if doTrim and _TRIM_DEBUG_LOG:
							try:
								leadMs = int((trimmedLeadFrames * 1000) // int(self._sr)) if self._sr else 0
								trailMs = int((trimmedTrailFrames * 1000) // int(self._sr)) if self._sr else 0
								log.debug(f"FlexVoice(trim): lead={leadMs}ms trail={trailMs}ms")
							except Exception:
								pass

					return True
				time.sleep(0.002)
				continue

		return False

	def _speakOne(self, utt: _Utt) -> bool:
		"""
		Speak a full NVDA utterance with driver-managed indexing.
		Returns True if finished normally, False if canceled/shutdown.
		"""
		if not self._engine.is_created:
			return False

		# Leading indexes: fire immediately.
		sentIndexes: set[int] = set()
		for idx in utt.leadingIndexes:
			if utt.token != self._getCancelToken() or self._shutdown.is_set():
				return False
			if idx in sentIndexes:
				continue
			sentIndexes.add(idx)
			try:
				synthIndexReached.notify(synth=self, index=int(idx))
			except Exception:
				pass

		# No speakable text: just ensure remaining indexes are fired, then done.
		if not utt.segments:
			for idx in utt.expectedIndexes:
				if idx in sentIndexes:
					continue
				sentIndexes.add(idx)
				try:
					synthIndexReached.notify(synth=self, index=int(idx))
				except Exception:
					pass
			return True


		trim = bool(_TRIM_SILENCE_ON_INDEXED_UTTERANCES and utt.expectedIndexes)

		# For say-all / caret tracking (indexed utterances), send one *combined* wrapper request
		# with explicit index markers. This avoids restarting the engine for each segment.
		if utt.expectedIndexes:
			ok = self._runWrapperUtteranceComposite(utt, utt.token, sentIndexes, trimSilence=trim)
			if not ok:
				return False
		else:
			# Non-indexed speech: keep the existing simple behavior.
			for seg in utt.segments:
				if utt.token != self._getCancelToken() or self._shutdown.is_set():
					return False
				ok = self._runWrapperTextSegment(seg.text, utt.token, sentIndexes, trimSilence=False)
				if not ok:
					return False

		# Fallback: emit any missing expected indexes (should be none in normal operation).

		for idx in utt.expectedIndexes:
			if idx in sentIndexes:
				continue
			sentIndexes.add(idx)
			try:
				synthIndexReached.notify(synth=self, index=int(idx))
			except Exception:
				pass

		# If there were no IndexCommands at all, we assume NVDA is using synthDoneSpeaking
		# as the boundary for this utterance. In that case, keep behavior accurate by
		# waiting until the WavePlayer buffer drains.
		#
		# If IndexCommands are present (say-all / caret tracking), we intentionally skip
		# syncing here to allow pipelining across chunks and avoid line-to-line pauses.
		if not utt.expectedIndexes:
			self._playerSync()

		return True


	def _workerLoop(self):
		log.debug("FlexVoice(wrapper): worker thread started")
		while not self._shutdown.is_set():
			# Drain pending utt queue first
			if self._pendingUtts:
				utt = self._pendingUtts.popleft()
				if utt.token == self._getCancelToken():
					try:
						log.debug("FlexVoice(wrapper): _speakOne (segs=%d, idx=%d)",
								  len(utt.segments), len(utt.expectedIndexes))
						ok = self._speakOne(utt)
					except Exception:
						log.error("FlexVoice(wrapper): _speakOne CRASHED", exc_info=True)
						ok = False
					if ok and utt.token == self._getCancelToken():
						try:
							synthDoneSpeaking.notify(synth=self)
						except Exception:
							pass
				else:
					log.debug("FlexVoice(wrapper): utt dropped (token mismatch: utt=%d cur=%d)",
							  utt.token, self._getCancelToken())
				continue

			# Wait for command
			try:
				cmd, payload = self._cmdQ.get(timeout=0.1)
			except queue.Empty:
				continue

			if cmd == _CMD_TERMINATE:
				break

			# Batch process commands (coalesce rapid events)
			batch = [(cmd, payload)]
			while True:
				try:
					batch.append(self._cmdQ.get_nowait())
				except queue.Empty:
					break

			recreate_requested = False
			cancel_requested = False
			last_pause = None
			speaks_to_queue: list[_Utt] = []

			for c, p in batch:
				if c == _CMD_TERMINATE:
					self._shutdown.set()
					break
				if c == _CMD_RECREATE:
					recreate_requested = True
					cancel_requested = True
					speaks_to_queue.clear()
					continue
				if c == _CMD_CANCEL:
					cancel_requested = True
					speaks_to_queue.clear()
					continue
				if c == _CMD_PAUSE:
					last_pause = bool(p)
					continue
				if c == _CMD_SPEAK and isinstance(p, _Utt):
					speaks_to_queue.append(p)
					continue

			if self._shutdown.is_set():
				break

			if last_pause is not None:
				self._playerPause(last_pause)

			if cancel_requested:
				self._pendingUtts.clear()
				self._hardStop()

			if recreate_requested:
				try:
					self._createEngineOrDie()
					self._ensurePlayer()
				except Exception:
					pass

			if speaks_to_queue:
				if cancel_requested:
					# Only keep the last speak in the batch after a cancel (best effort).
					self._pendingUtts.append(speaks_to_queue[-1])
				else:
					self._pendingUtts.extend(speaks_to_queue)

		try:
			self._playerStop()
		except Exception:
			pass


# ---------------------------------------------------------------------------
# 64-bit NVDA 2026.1+: use the built-in bridge to run the full driver in a
# 32-bit host process.  Audio plays from the host directly via nvwave.
# On 32-bit (including the bridge host), this block is skipped and the
# SynthDriver class defined above is used as-is.
# ---------------------------------------------------------------------------
try:
	from _bridge.clients.synthDriverHost32.synthDriver import SynthDriverProxy32 as _Proxy32
except ImportError:
	# NVDA older than 2026.1 has no built-in bridge. On 64-bit that leaves the
	# legacy host process path in _flexvoice.HostEngine; on 32-bit the driver
	# above runs in-process anyway. Either way, importing this module must not
	# fail, or the synth disappears from the list entirely.
	_Proxy32 = None

if ctypes.sizeof(ctypes.c_void_p) == 8 and _Proxy32 is not None:

	class SynthDriver(_Proxy32):
		name = "flexvoice"
		description = "Mindmaker FlexVoice 3"
		synthDriver32Path = os.path.dirname(__file__)
		synthDriver32Name = "flexvoice"

		@classmethod
		def check(cls):
			if not super().check():
				return False
			base = os.path.dirname(__file__)
			if not os.path.isfile(os.path.join(base, "fvwrap.dll")):
				addon_root = os.path.abspath(os.path.join(base, os.pardir))
				if not os.path.isfile(os.path.join(addon_root, "fvwrap.dll")):
					return False
			# Any shipped language will do; English is no longer mandatory.
			if not _findLanguageDirs(base) and not _findEnglishDir(base):
				return False
			return True

		# SynthDriverProxy only forwards voice, variant, rate, rateBoost, pitch
		# and volume. 'language' is ours, so it has to be forwarded by hand or it
		# silently does nothing under 64-bit NVDA while working fine in a
		# 32-bit test harness.
		#: Mirrors the host's language. NVDA reads synth.language while building
		#: every utterance, and each read would otherwise be a round trip to
		#: another process.
		_cachedLanguage = None

		def _get_language(self):
			if self._cachedLanguage is not None:
				return self._cachedLanguage
			try:
				self._cachedLanguage = self._remoteService.getParam("language")
			except Exception:
				log.debugWarning("FlexVoice: could not read language from host", exc_info=True)
				self._cachedLanguage = _DEFAULT_LANG
			return self._cachedLanguage

		def _set_language(self, value):
			try:
				self._remoteService.setParam("language", value)
			except Exception:
				log.error("FlexVoice: could not set language on host", exc_info=True)
				return
			# Read back rather than assuming the set took. The remote setter
			# ignores a language it has no data for, and that refusal does not
			# raise across the bridge, so caching `value` blind would report a
			# language the host is not actually speaking.
			try:
				self._cachedLanguage = self._remoteService.getParam("language")
			except Exception:
				self._cachedLanguage = None
			# The voice list is language-independent, so it does not need
			# rebuilding - but the host moved to a different voice, and the
			# open dialog should show that. Only the selection changes, which
			# updateDriverSettings can do.
			_scheduleVoiceSettingsRefresh()

		def _set_voice(self, value):
			# The language travels with the voice on the host side, so this
			# side's cached language has to follow - otherwise NVDA keeps
			# tagging speech with the previous language.
			super()._set_voice(value)
			try:
				self._cachedLanguage = self._remoteService.getParam("language")
			except Exception:
				self._cachedLanguage = None

		# Clarity is ours too, so the bridge does not forward it either.
		def _get_clarity(self):
			try:
				return int(self._remoteService.getParam("clarity"))
			except Exception:
				return 50

		def _set_clarity(self, value):
			try:
				self._remoteService.setParam("clarity", int(value))
			except Exception:
				log.debugWarning("FlexVoice: could not set clarity on host", exc_info=True)

		def _get_availableLanguages(self):
			# Computed locally rather than over the wire: the data folders are on
			# the same disk, and LanguageInfo objects would have to be serialized.
			out = {}
			for nvdaLang in _findLanguageDirs(os.path.dirname(__file__)):
				if LanguageInfo is not None:
					out[nvdaLang] = LanguageInfo(nvdaLang)
				else:
					out[nvdaLang] = VoiceInfo(nvdaLang, nvdaLang, nvdaLang)
			return out
