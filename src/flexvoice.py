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
from dataclasses import dataclass
from collections import deque
from ctypes import (
	c_void_p, c_char_p, c_int, c_ubyte, POINTER, byref
)

import nvwave
import config
from logHandler import log
from synthDriverHandler import (
	SynthDriver as BaseSynthDriver,
	VoiceInfo,
	synthIndexReached,
	synthDoneSpeaking,
)
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


# FlexVoice language IDs
LNG_ENGLISH = 0x0409

# Wrapper stream item types (must match fvwrap.h)
FVWRAP_ITEM_NONE = 0
FVWRAP_ITEM_AUDIO = 1
FVWRAP_ITEM_INDEX = 2
FVWRAP_ITEM_DONE = 3
FVWRAP_ITEM_ERROR = 4


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


def _mbcs(s: str | None) -> bytes | None:
	if not s:
		return None
	return str(s).encode("mbcs", errors="replace")


def _findEnglishDir(base_dir: str) -> str:
	"""
	Robustly find the folder containing 'VoiceList.tvl'.
	"""
	candidates = [
		base_dir,
		os.path.join(base_dir, "English"),
		os.path.abspath(os.path.join(base_dir, os.pardir)),
		os.path.abspath(os.path.join(base_dir, os.pardir, "English")),
	]
	# Check addon root explicitly
	addon_root = os.path.abspath(os.path.join(base_dir, os.pardir))
	candidates.append(addon_root)
	candidates.append(os.path.join(addon_root, "English"))

	for p in candidates:
		if os.path.isdir(p) and os.path.isfile(os.path.join(p, "VoiceList.tvl")):
			return p
	return ""


def _parseVoiceList(voiceListPath: str, englishDir: str) -> dict[str, str]:
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
				full = os.path.normpath(rel if os.path.isabs(rel) else os.path.join(englishDir, rel))
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


def _sanitizeTextForEngine(s: str) -> str:
	"""
	Wrapper-side conversion is UTF-8 -> Latin-1-ish.
	To reduce surprises we pre-sanitize to <= 0xFF and strip control chars.
	"""
	if not s:
		return ""
	for k, v in _TRANSLATE.items():
		s = s.replace(k, v)

	out = []
	for ch in s:
		o = ord(ch)
		if ch in ("\r", "\n", "\t"):
			out.append(ch)
			continue
		if o < 0x20 or o == 0x7F:
			out.append(" ")
			continue
		if o > 0xFF:
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
	language = "en"

	# IMPORTANT:
	# We explicitly declare rate/volume/pitch as NumericDriverSetting with
	# availableInSettingsRing=True to match BestSpeech’s pattern and to avoid
	# NVDA-version-dependent defaults.
	if NumericDriverSetting is not None:
		supportedSettings = (
			BaseSynthDriver.VoiceSetting(),
			NumericDriverSetting("rate", "&Rate", defaultVal=50, availableInSettingsRing=True),
			NumericDriverSetting("pitch", "&Pitch", defaultVal=50, availableInSettingsRing=True),
			NumericDriverSetting("volume", "&Volume", defaultVal=100, availableInSettingsRing=True),
		)
	else:
		# Fallback for older NVDA: classic supported settings.
		supportedSettings = (
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
		return bool(_findEnglishDir(base))

	def __init__(self):
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

		# Track applied settings to avoid churn
		self._appliedRate = None
		self._appliedVol = None
		self._appliedPitch = None

		# 3) Runtime handles / audio format
		self._handle: c_void_p | None = None
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

		self._englishDir = _findEnglishDir(self._baseDir)
		if not self._englishDir:
			raise RuntimeError("FlexVoice: English data not found")

		self._voiceMap = _parseVoiceList(os.path.join(self._englishDir, "VoiceList.tvl"), self._englishDir)
		self._voiceIds = ["Default"] + sorted(self._voiceMap.keys(), key=lambda s: s.lower())
		self._voice = "Default"
		if "Tim" in self._voiceIds:
			self._voice = "Tim"

		# 5) Engine init
		self._setupDllSearchPath()
		self._dll = ctypes.CDLL(self._wrapperPath)
		self._bindWrapper()
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

		if self._handle:
			try:
				self._dll.fvwrap_destroy(self._handle)
			except Exception:
				pass
		self._handle = None

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
	def _getAvailableVoices(self):
		return {name: VoiceInfo(name, name) for name in self._voiceIds}

	def _get_voice(self):
		return self._voice

	def _set_voice(self, val):
		if val == self._voice:
			return
		self._voice = val
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

	# ---------------- Internals ----------------
	def _setupDllSearchPath(self):
		dirs = [self._baseDir, self._addonRoot]
		try:
			if hasattr(os, "add_dll_directory"):
				for d in dirs:
					os.add_dll_directory(d)
		except Exception:
			pass
		os.environ["PATH"] = os.pathsep.join(dirs) + os.pathsep + os.environ.get("PATH", "")

		# Env var fallback used by EngineFactory(NULL) logic.
		parentOfEnglish = os.path.dirname(self._englishDir)
		os.environ["FLEXVOICE_DATA"] = parentOfEnglish or self._englishDir

	def _bindWrapper(self):
		d = self._dll
		d.fvwrap_create.restype = c_void_p
		d.fvwrap_create.argtypes = (c_char_p, c_char_p, c_int, c_int, c_int)

		d.fvwrap_destroy.restype = None
		d.fvwrap_destroy.argtypes = (c_void_p,)

		d.fvwrap_stop.restype = None
		d.fvwrap_stop.argtypes = (c_void_p,)

		d.fvwrap_setRatePercent.restype = c_int
		d.fvwrap_setRatePercent.argtypes = (c_void_p, c_int)

		d.fvwrap_setVolumePercent.restype = c_int
		d.fvwrap_setVolumePercent.argtypes = (c_void_p, c_int)

		self._hasPitch = hasattr(d, "fvwrap_setPitchPercent")
		if self._hasPitch:
			d.fvwrap_setPitchPercent.restype = c_int
			d.fvwrap_setPitchPercent.argtypes = (c_void_p, c_int)

		d.fvwrap_begin.restype = None
		d.fvwrap_begin.argtypes = (c_void_p,)

		d.fvwrap_addTextUtf8.restype = None
		d.fvwrap_addTextUtf8.argtypes = (c_void_p, c_char_p)

		# We keep these bound but (by design) do not rely on wrapper index events.
		d.fvwrap_addIndex.restype = None
		d.fvwrap_addIndex.argtypes = (c_void_p, c_int)

		d.fvwrap_commit.restype = c_int
		d.fvwrap_commit.argtypes = (c_void_p, c_int)

		d.fvwrap_read.restype = c_int
		d.fvwrap_read.argtypes = (c_void_p, POINTER(c_int), POINTER(c_int), POINTER(c_ubyte), c_int)

	def _speakerPathForVoice(self, voiceName: str) -> str:
		if voiceName and voiceName != "Default":
			p = self._voiceMap.get(voiceName)
			if p and os.path.isfile(p):
				return p
		return os.path.join(self._englishDir, "default.tav")

	def _createEngineOrDie(self):
		englishDir = self._englishDir
		parentOfEnglish = os.path.dirname(englishDir)
		speakerPath = self._speakerPathForVoice(self._voice)

		dataCandidates = [
			(englishDir, "dataPath=English dir"),
			(parentOfEnglish, "dataPath=parent(English)"),
			(None, "dataPath=NULL"),
		]

		tryParams = [(16000, 16), (11025, 16), (8000, 16)]

		for dataPath, label in dataCandidates:
			if dataPath:
				os.environ["FLEXVOICE_DATA"] = dataPath
			else:
				os.environ["FLEXVOICE_DATA"] = parentOfEnglish or englishDir

			for sr, bits in tryParams:
				try:
					h = self._dll.fvwrap_create(
						_mbcs(dataPath),
						_mbcs(speakerPath),
						int(LNG_ENGLISH),
						int(sr),
						int(bits)
					)
					if h:
						if self._handle:
							try:
								self._dll.fvwrap_destroy(self._handle)
							except Exception:
								pass
						self._handle = h
						self._sr = sr
						self._bits = bits
						self._channels = 1

						self._appliedRate = None
						self._appliedVol = None
						self._appliedPitch = None

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
			s = _sanitizeTextForEngine(raw)
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
			if self._handle:
				self._dll.fvwrap_stop(self._handle)
		except Exception:
			pass
		self._flushWrapperOutput(40)

	def _applySettingsIfChanged(self):
		h = self._handle
		if not h:
			return

		rate = _clampPercent(self._rate)
		vol = _clampPercent(self._volume)
		pit = _clampPercent(self._pitch)

		if self._appliedRate != rate:
			try:
				self._dll.fvwrap_setRatePercent(h, int(rate))
				self._appliedRate = rate
			except Exception:
				pass

		if self._appliedVol != vol:
			try:
				self._dll.fvwrap_setVolumePercent(h, int(vol))
				self._appliedVol = vol
			except Exception:
				pass

		if self._hasPitch and self._appliedPitch != pit:
			try:
				self._dll.fvwrap_setPitchPercent(h, int(pit))
				self._appliedPitch = pit
			except Exception:
				pass

	def _flushWrapperOutput(self, maxMs: int = 40):
		h = self._handle
		if not h:
			return
		bufSize = 4096
		audioBuf = (c_ubyte * bufSize)()
		outType = c_int()
		outValue = c_int()
		deadline = time.time() + (maxMs / 1000.0)

		while time.time() < deadline:
			_ = int(self._dll.fvwrap_read(h, byref(outType), byref(outValue), audioBuf, bufSize))
			if int(outType.value) == FVWRAP_ITEM_NONE:
				return

	def _runWrapperTextSegment(self, text: str, tokenSnapshot: int, sentIndexes: set[int]) -> bool:
		"""
		Speak one plain text segment through the wrapper, draining audio until DONE.

		Returns False if canceled or shutdown.
		"""
		h = self._handle
		if not h:
			return False

		self._applySettingsIfChanged()

		try:
			self._dll.fvwrap_begin(h)
			b = (text or "").encode("utf-8", "replace")
			if b:
				self._dll.fvwrap_addTextUtf8(h, b)
			rc = int(self._dll.fvwrap_commit(h, 1))
			if rc != 0:
				return False
		except Exception:
			return False

		bufSize = 8192
		buf = (c_ubyte * bufSize)()
		outTp, outVal = c_int(), c_int()

		seenDone = False
		lastActivity = time.time()

		while not self._shutdown.is_set():
			if tokenSnapshot != self._getCancelToken():
				return False

			n = int(self._dll.fvwrap_read(h, byref(outTp), byref(outVal), buf, bufSize))
			tp = int(outTp.value)

			if tp == FVWRAP_ITEM_AUDIO:
				if n > 0:
					self._playerFeed(bytes(buf[:n]))
					lastActivity = time.time()
				continue

			if tp == FVWRAP_ITEM_INDEX:
				# If wrapper emits any indexes (some builds do), pass them through
				# but dedupe with driver-emitted ones.
				idx = int(outVal.value)
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
				# Give a tiny tail window for late audio, then sync and return.
				if seenDone and (time.time() - lastActivity > 0.05):
					self._playerSync()
					return True
				time.sleep(0.002)
				continue

		return False

	def _speakOne(self, utt: _Utt) -> bool:
		"""
		Speak a full NVDA utterance with driver-managed indexing.
		Returns True if finished normally, False if canceled/shutdown.
		"""
		if not self._handle:
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

		# Speak each segment; after segment completes (audio fully played), fire indexes after it.
		for seg in utt.segments:
			if utt.token != self._getCancelToken() or self._shutdown.is_set():
				return False

			ok = self._runWrapperTextSegment(seg.text, utt.token, sentIndexes)
			if not ok:
				return False

			# After audio is synced for this segment, emit the boundary indexes.
			for idx in seg.idxAfter:
				if utt.token != self._getCancelToken() or self._shutdown.is_set():
					return False
				if idx in sentIndexes:
					continue
				sentIndexes.add(idx)
				try:
					synthIndexReached.notify(synth=self, index=int(idx))
				except Exception:
					pass

		# Fallback: emit any missing expected indexes (should be none in normal operation).
		for idx in utt.expectedIndexes:
			if idx in sentIndexes:
				continue
			sentIndexes.add(idx)
			try:
				synthIndexReached.notify(synth=self, index=int(idx))
			except Exception:
				pass

		return True

	def _workerLoop(self):
		while not self._shutdown.is_set():
			# Drain pending utt queue first
			if self._pendingUtts:
				utt = self._pendingUtts.popleft()
				if utt.token == self._getCancelToken():
					ok = self._speakOne(utt)
					if ok and utt.token == self._getCancelToken():
						try:
							synthDoneSpeaking.notify(synth=self)
						except Exception:
							pass
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
