"""Engine abstraction for FlexVoice fvwrap.dll.

On 32-bit Python (NVDA 2025.3 and earlier): loads fvwrap.dll directly
via ctypes -- no host process needed.

On 64-bit Python (NVDA 2026.1+): communicates with a 32-bit host process
over IPC so the 32-bit wrapper DLL can be used from 64-bit NVDA.

Adapted from Fastfinge's Eloquence 64 project with permission.
Original: https://github.com/Fastfinge/eloquence_64
"""
from __future__ import annotations

import ctypes
import itertools
import logging
import os
import queue
import subprocess
import threading
import time
from typing import Any, Dict, List, Optional, Tuple

IS_64BIT = ctypes.sizeof(ctypes.c_void_p) == 8

if IS_64BIT:
	# Only the legacy host-process path uses _ipc, and that path is unreachable
	# on any NVDA that is actually 64-bit (2026.1+ ships the built-in bridge).
	# An unguarded import here means a missing _ipc.py removes the whole synth
	# from NVDA's list instead of just disabling a dead code path.
	try:
		from . import _ipc
	except ImportError:
		_ipc = None

# Use NVDA's logger when available (so messages appear in NVDA log),
# fall back to standard logging for standalone host process.
try:
	from logHandler import log as LOGGER
except ImportError:
	LOGGER = logging.getLogger(__name__)

# Stream item types (must match fvwrap.h)
FVWRAP_ITEM_NONE = 0
FVWRAP_ITEM_AUDIO = 1
FVWRAP_ITEM_INDEX = 2
FVWRAP_ITEM_DONE = 3
FVWRAP_ITEM_ERROR = 4

HOST_EXECUTABLE = "flexvoice_host32.exe"
HOST_SCRIPT = "host_flexvoice32.py"
AUTH_KEY_BYTES = 16


# ---------------------------------------------------------------------------
# 32-bit DirectEngine (loads fvwrap.dll in-process)
# ---------------------------------------------------------------------------

class DirectEngine:
	"""Direct ctypes access to fvwrap.dll for 32-bit NVDA."""

	def __init__(self, wrapper_path: str, dll_search_dirs: List[str]):
		self._wrapper_path = wrapper_path
		self._dll_search_dirs = dll_search_dirs
		self._dll = None
		self._handle = None
		self.has_pitch = False
		# Read buffer
		self._buf_size = 8192
		self._audio_buf = None
		self._out_type = None
		self._out_value = None

	@property
	def is_created(self) -> bool:
		return self._handle is not None

	def create(self, data_path: Optional[str], speaker_path: Optional[str],
			   language: int, sample_rate: int, bits_per_sample: int) -> bool:
		"""Create the engine. Loads DLL on first call. Returns True on success."""
		if self._dll is None:
			self._setup_dll_search()
			self._dll = ctypes.CDLL(self._wrapper_path)
			self._bind_ctypes()
			self._audio_buf = (ctypes.c_ubyte * self._buf_size)()
			self._out_type = ctypes.c_int(0)
			self._out_value = ctypes.c_int(0)

		data_bytes = data_path.encode("mbcs", "replace") if data_path else None
		speaker_bytes = speaker_path.encode("mbcs", "replace") if speaker_path else None

		try:
			h = self._dll.fvwrap_create(data_bytes, speaker_bytes,
										int(language), int(sample_rate), int(bits_per_sample))
		except Exception:
			return False
		if not h:
			return False

		if self._handle:
			try:
				self._dll.fvwrap_destroy(self._handle)
			except Exception:
				pass
		self._handle = h
		return True

	def destroy(self) -> None:
		if self._handle:
			try:
				self._dll.fvwrap_destroy(self._handle)
			except Exception:
				pass
			self._handle = None

	def stop(self) -> None:
		if self._handle:
			try:
				self._dll.fvwrap_stop(self._handle)
			except Exception:
				pass

	def flush_output(self, max_ms: int = 40) -> None:
		if not self._handle:
			return
		deadline = time.time() + (max_ms / 1000.0)
		while time.time() < deadline:
			tp, val, data = self.read()
			if tp == FVWRAP_ITEM_NONE:
				return

	def set_rate_percent(self, val: int) -> None:
		if self._handle:
			self._dll.fvwrap_setRatePercent(self._handle, int(val))

	def set_volume_percent(self, val: int) -> None:
		if self._handle:
			self._dll.fvwrap_setVolumePercent(self._handle, int(val))

	def set_pitch_percent(self, val: int) -> None:
		if self._handle and self.has_pitch:
			self._dll.fvwrap_setPitchPercent(self._handle, int(val))

	def begin(self) -> None:
		if self._handle:
			self._dll.fvwrap_begin(self._handle)

	def add_text_utf8(self, text_bytes: bytes) -> None:
		if self._handle and text_bytes:
			self._dll.fvwrap_addTextUtf8(self._handle, text_bytes)

	def add_index(self, index: int) -> None:
		if self._handle:
			self._dll.fvwrap_addIndex(self._handle, int(index))

	def commit(self, repeat_count: int = 1) -> int:
		if not self._handle:
			return -1
		return int(self._dll.fvwrap_commit(self._handle, int(repeat_count)))

	def read(self) -> Tuple[int, int, bytes]:
		"""Read next item. Returns (item_type, value, audio_bytes)."""
		if not self._handle:
			return (FVWRAP_ITEM_NONE, 0, b"")
		n = int(self._dll.fvwrap_read(
			self._handle,
			ctypes.byref(self._out_type),
			ctypes.byref(self._out_value),
			self._audio_buf,
			self._buf_size,
		))
		tp = int(self._out_type.value)
		val = int(self._out_value.value)
		data = bytes(self._audio_buf[:n]) if n > 0 else b""
		return (tp, val, data)

	def _setup_dll_search(self) -> None:
		for d in self._dll_search_dirs:
			try:
				if hasattr(os, "add_dll_directory"):
					os.add_dll_directory(d)
			except Exception:
				pass
		os.environ["PATH"] = os.pathsep.join(self._dll_search_dirs) + os.pathsep + os.environ.get("PATH", "")

	def _bind_ctypes(self) -> None:
		d = self._dll
		d.fvwrap_create.restype = ctypes.c_void_p
		d.fvwrap_create.argtypes = (ctypes.c_char_p, ctypes.c_char_p, ctypes.c_int, ctypes.c_int, ctypes.c_int)
		d.fvwrap_destroy.restype = None
		d.fvwrap_destroy.argtypes = (ctypes.c_void_p,)
		d.fvwrap_stop.restype = None
		d.fvwrap_stop.argtypes = (ctypes.c_void_p,)
		d.fvwrap_setRatePercent.restype = ctypes.c_int
		d.fvwrap_setRatePercent.argtypes = (ctypes.c_void_p, ctypes.c_int)
		d.fvwrap_setVolumePercent.restype = ctypes.c_int
		d.fvwrap_setVolumePercent.argtypes = (ctypes.c_void_p, ctypes.c_int)
		self.has_pitch = hasattr(d, "fvwrap_setPitchPercent")
		if self.has_pitch:
			d.fvwrap_setPitchPercent.restype = ctypes.c_int
			d.fvwrap_setPitchPercent.argtypes = (ctypes.c_void_p, ctypes.c_int)
		d.fvwrap_begin.restype = None
		d.fvwrap_begin.argtypes = (ctypes.c_void_p,)
		d.fvwrap_addTextUtf8.restype = None
		d.fvwrap_addTextUtf8.argtypes = (ctypes.c_void_p, ctypes.c_char_p)
		d.fvwrap_addIndex.restype = None
		d.fvwrap_addIndex.argtypes = (ctypes.c_void_p, ctypes.c_int)
		d.fvwrap_commit.restype = ctypes.c_int
		d.fvwrap_commit.argtypes = (ctypes.c_void_p, ctypes.c_int)
		d.fvwrap_read.restype = ctypes.c_int
		d.fvwrap_read.argtypes = (
			ctypes.c_void_p,
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_ubyte),
			ctypes.c_int,
		)


# ---------------------------------------------------------------------------
# 64-bit HostEngine (IPC to 32-bit host process)
# ---------------------------------------------------------------------------

if IS_64BIT:
	from dataclasses import dataclass

	@dataclass
	class HostProcess:
		process: subprocess.Popen
		connection: _ipc.IpcConnection
		listener: Any  # socket

	class HostEngine:
		"""Communicates with a 32-bit host process via IPC."""

		def __init__(self, wrapper_path: str, dll_search_dirs: List[str]):
			self._wrapper_path = wrapper_path
			self._dll_search_dirs = dll_search_dirs
			self._host: Optional[HostProcess] = None
			self._pending: Dict[int, threading.Event] = {}
			self._responses: Dict[int, Dict[str, Any]] = {}
			self._receiver: Optional[threading.Thread] = None
			self._id_counter = itertools.count(1)
			self._send_lock = threading.Lock()
			self._event_queue: queue.Queue = queue.Queue()
			self._utt_parts: List[Tuple[str, Any]] = []
			self.has_pitch = False
			self._created = False

		@property
		def is_created(self) -> bool:
			return self._created

		def _ensure_host(self) -> None:
			if self._host:
				return
			addon_dir = os.path.abspath(os.path.dirname(self._wrapper_path))
			authkey = os.urandom(AUTH_KEY_BYTES)
			listener = _ipc.create_listener()
			port = listener.getsockname()[1]

			cmd = list(self._resolve_host_executable(addon_dir))
			cmd.extend([
				"--address", f"127.0.0.1:{port}",
				"--authkey", authkey.hex(),
				"--log-dir", addon_dir,
			])
			LOGGER.info("Launching FlexVoice host: %s", cmd)
			proc = subprocess.Popen(cmd, cwd=addon_dir, creationflags=subprocess.CREATE_NO_WINDOW)
			conn = _ipc.accept_authenticated(listener, authkey)
			self._host = HostProcess(process=proc, connection=conn, listener=listener)

			self._receiver = threading.Thread(target=self._receiver_loop, daemon=True,
											  name="FlexVoiceReceiver")
			self._receiver.start()

		def _resolve_host_executable(self, addon_dir: str):
			override = os.environ.get("FLEXVOICE_HOST_COMMAND")
			if override:
				import shlex
				return shlex.split(override)
			exe_path = os.path.join(addon_dir, HOST_EXECUTABLE)
			if os.path.exists(exe_path):
				return [exe_path]
			script_path = os.path.join(addon_dir, HOST_SCRIPT)
			if os.path.exists(script_path):
				return ["py", "-3.14-32", script_path]
			raise RuntimeError("FlexVoice host executable not found")

		def create(self, data_path: Optional[str], speaker_path: Optional[str],
				   language: int, sample_rate: int, bits_per_sample: int) -> bool:
			self._ensure_host()
			try:
				result = self._send_command(
					"create",
					wrapperPath=self._wrapper_path,
					dllDirs=self._dll_search_dirs,
					dataPath=data_path or "",
					speakerPath=speaker_path or "",
					language=language,
					sampleRate=sample_rate,
					bitsPerSample=bits_per_sample,
				)
				self.has_pitch = result.get("hasPitch", False)
				self._created = True
				LOGGER.info("HostEngine: create succeeded (hasPitch=%s)", self.has_pitch)
				return True
			except Exception:
				LOGGER.exception("FlexVoice create failed")
				return False

		def destroy(self) -> None:
			self._created = False
			if not self._host:
				return
			LOGGER.info("HostEngine: destroying")
			try:
				self._send_command("delete", timeout=3.0)
			except Exception:
				LOGGER.exception("Failed to send delete")
			if self._receiver:
				self._receiver.join(timeout=2)
				self._receiver = None
			try:
				self._host.connection.close()
			except Exception:
				pass
			try:
				self._host.listener.close()
			except Exception:
				pass
			try:
				self._host.process.terminate()
				self._host.process.wait(timeout=2)
			except Exception:
				try:
					self._host.process.kill()
				except Exception:
					pass
			self._host = None

		def stop(self) -> None:
			if not self._host:
				return
			# Clear event queue
			while not self._event_queue.empty():
				try:
					self._event_queue.get_nowait()
				except queue.Empty:
					break
			# Fire-and-forget stop
			try:
				msg_id = next(self._id_counter)
				with self._send_lock:
					self._host.connection.send({
						"type": "command", "id": msg_id,
						"command": "stop", "payload": {},
					})
			except Exception:
				LOGGER.exception("Stop command failed")

		def flush_output(self, max_ms: int = 40) -> None:
			deadline = time.time() + (max_ms / 1000.0)
			while time.time() < deadline:
				tp, val, data = self.read()
				if tp == FVWRAP_ITEM_NONE:
					return

		def set_rate_percent(self, val: int) -> None:
			if self._host:
				try:
					self._send_command("setRatePercent", value=int(val))
				except Exception:
					LOGGER.exception("setRatePercent failed")

		def set_volume_percent(self, val: int) -> None:
			if self._host:
				try:
					self._send_command("setVolumePercent", value=int(val))
				except Exception:
					LOGGER.exception("setVolumePercent failed")

		def set_pitch_percent(self, val: int) -> None:
			if self._host and self.has_pitch:
				try:
					self._send_command("setPitchPercent", value=int(val))
				except Exception:
					LOGGER.exception("setPitchPercent failed")

		def begin(self) -> None:
			self._utt_parts = []

		def add_text_utf8(self, text_bytes: bytes) -> None:
			if text_bytes:
				self._utt_parts.append(("text", text_bytes))

		def add_index(self, index: int) -> None:
			self._utt_parts.append(("index", int(index)))

		def commit(self, repeat_count: int = 1) -> int:
			if not self._host:
				LOGGER.warning("commit: no host")
				return -1
			# Clear event queue (new synthesis)
			while not self._event_queue.empty():
				try:
					self._event_queue.get_nowait()
				except queue.Empty:
					break
			# Fire-and-forget synthesize
			try:
				msg_id = next(self._id_counter)
				LOGGER.info("commit: sending synthesize (id=%d, %d parts)", msg_id, len(self._utt_parts))
				with self._send_lock:
					self._host.connection.send({
						"type": "command", "id": msg_id,
						"command": "synthesize",
						"payload": {"parts": self._utt_parts, "repeatCount": repeat_count},
					})
			except Exception:
				LOGGER.exception("synthesize command failed")
				return -1
			self._utt_parts = []
			return 0

		def read(self) -> Tuple[int, int, bytes]:
			"""Read next event from the host. Returns (item_type, value, audio_bytes)."""
			try:
				event = self._event_queue.get(timeout=0.005)
			except queue.Empty:
				return (FVWRAP_ITEM_NONE, 0, b"")
			etype = event.get("event")
			payload = event.get("payload", {})
			if etype == "audio":
				return (FVWRAP_ITEM_AUDIO, 0, payload.get("data", b""))
			elif etype == "index":
				return (FVWRAP_ITEM_INDEX, payload.get("index", 0), b"")
			elif etype == "done":
				return (FVWRAP_ITEM_DONE, 0, b"")
			elif etype == "error":
				return (FVWRAP_ITEM_ERROR, payload.get("code", 0), b"")
			elif etype == "stopped":
				# Drain this; not a real audio event
				return (FVWRAP_ITEM_NONE, 0, b"")
			LOGGER.debug("read: unknown event type %s", etype)
			return (FVWRAP_ITEM_NONE, 0, b"")

		# ---- IPC internals ----

		def _receiver_loop(self) -> None:
			conn = self._host.connection if self._host else None
			if conn is None:
				return
			LOGGER.debug("Receiver loop started")
			while True:
				try:
					message = conn.recv()
				except (EOFError, ConnectionAbortedError, OSError):
					LOGGER.info("Host connection closed")
					for msg_id, event in list(self._pending.items()):
						self._responses[msg_id] = {"error": "connectionClosed"}
						event.set()
					self._pending.clear()
					break
				except Exception:
					LOGGER.exception("Unexpected error in receiver loop")
					for msg_id, event in list(self._pending.items()):
						self._responses[msg_id] = {"error": "receiverException"}
						event.set()
					self._pending.clear()
					break

				msg_type = message.get("type")
				if msg_type == "response":
					msg_id = message["id"]
					self._responses[msg_id] = message
					event = self._pending.pop(msg_id, None)
					if event:
						event.set()
					else:
						LOGGER.debug("Receiver: response id=%s (no pending waiter)", msg_id)
				elif msg_type == "event":
					self._event_queue.put(message)
				else:
					LOGGER.warning("Unknown message type %s", msg_type)
			LOGGER.debug("Receiver loop exited")

		def _send_command(self, command: str, timeout: float = 10.0, **payload: Any) -> Dict[str, Any]:
			if not self._host:
				raise RuntimeError("Host not started")
			msg_id = next(self._id_counter)
			event = threading.Event()
			self._pending[msg_id] = event
			LOGGER.debug("_send_command: %s (id=%d)", command, msg_id)
			with self._send_lock:
				try:
					self._host.connection.send({
						"type": "command", "id": msg_id,
						"command": command, "payload": payload,
					})
				except Exception:
					self._pending.pop(msg_id, None)
					raise
			if not event.wait(timeout=timeout):
				self._pending.pop(msg_id, None)
				LOGGER.error("_send_command: %s (id=%d) TIMED OUT after %.1fs", command, msg_id, timeout)
				raise RuntimeError(f"Command {command} timed out")
			response = self._responses.pop(msg_id, {"error": "no response"})
			if "error" in response:
				LOGGER.debug("_send_command: %s (id=%d) error: %s", command, msg_id, response.get("error"))
				raise RuntimeError(response["error"])
			LOGGER.debug("_send_command: %s (id=%d) OK", command, msg_id)
			return response.get("payload", {})


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def create_engine(wrapper_path: str, dll_search_dirs: List[str]):
	"""Create the appropriate engine for the current Python bitness."""
	if IS_64BIT:
		return HostEngine(wrapper_path, dll_search_dirs)
	else:
		return DirectEngine(wrapper_path, dll_search_dirs)
