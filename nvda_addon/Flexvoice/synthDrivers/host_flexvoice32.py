"""32-bit host process for FlexVoice speech synthesis.

This module runs as a separate 32-bit Python process.  It loads
fvwrap.dll (which in turn loads FlexVoice_3_01_001.dll) and exposes
an RPC protocol over a TCP socket so that 64-bit NVDA can use the
synthesizer.

Adapted from Fastfinge's Eloquence 64 project with permission.
Original: https://github.com/Fastfinge/eloquence_64
"""
from __future__ import annotations

import argparse
import ctypes
import logging
import os
import threading
import time
from typing import Any, Dict, List, Optional, Tuple

import _ipc

# Stream item types (must match fvwrap.h)
FVWRAP_ITEM_NONE = 0
FVWRAP_ITEM_AUDIO = 1
FVWRAP_ITEM_INDEX = 2
FVWRAP_ITEM_DONE = 3
FVWRAP_ITEM_ERROR = 4

LOGGER = logging.getLogger("flexvoice.host")


def configure_logging(log_dir: Optional[str]) -> None:
	logging.basicConfig(
		filename=os.path.join(log_dir, "flexvoice-host.log") if log_dir else None,
		level=logging.DEBUG,
		format="%(asctime)s %(levelname)s %(message)s",
	)


class FlexVoiceRuntime:
	"""Wraps access to the 32-bit fvwrap.dll."""

	def __init__(self, conn: _ipc.IpcConnection):
		self._conn = conn
		self._dll = None
		self._handle = None
		self._has_pitch = False
		self._should_stop = False
		# Read buffer
		self._buf_size = 8192
		self._audio_buf = None
		self._out_type = None
		self._out_value = None

	def _send_event(self, event: str, **payload: object) -> None:
		try:
			self._conn.send({"type": "event", "event": event, "payload": payload})
		except Exception:
			LOGGER.exception("Failed to send event %s", event)

	def start(self, wrapper_path: str, dll_dirs: List[str],
			  data_path: str, speaker_path: str,
			  language: int, sample_rate: int, bits_per_sample: int) -> Dict[str, Any]:
		"""Load DLL and create engine. Returns result dict."""
		# Setup DLL search paths
		for d in dll_dirs:
			try:
				if hasattr(os, "add_dll_directory"):
					os.add_dll_directory(d)
			except Exception:
				pass
		os.environ["PATH"] = os.pathsep.join(dll_dirs) + os.pathsep + os.environ.get("PATH", "")

		# Set FLEXVOICE_DATA env var
		if data_path:
			if os.path.isdir(data_path):
				os.environ["FLEXVOICE_DATA"] = data_path
			else:
				os.environ["FLEXVOICE_DATA"] = os.path.dirname(data_path) or data_path

		LOGGER.info("Loading fvwrap.dll from %s", wrapper_path)
		self._dll = ctypes.CDLL(wrapper_path)
		self._bind_ctypes()

		self._audio_buf = (ctypes.c_ubyte * self._buf_size)()
		self._out_type = ctypes.c_int(0)
		self._out_value = ctypes.c_int(0)

		data_bytes = data_path.encode("mbcs", "replace") if data_path else None
		speaker_bytes = speaker_path.encode("mbcs", "replace") if speaker_path else None

		LOGGER.info("Creating engine: data=%s, speaker=%s, lang=%d, sr=%d, bps=%d",
					data_path, speaker_path, language, sample_rate, bits_per_sample)
		self._handle = self._dll.fvwrap_create(
			data_bytes, speaker_bytes, int(language), int(sample_rate), int(bits_per_sample))
		if not self._handle:
			raise RuntimeError("fvwrap_create returned NULL")

		self._has_pitch = hasattr(self._dll, "fvwrap_setPitchPercent")
		LOGGER.info("Engine created (hasPitch=%s)", self._has_pitch)
		return {"hasPitch": self._has_pitch}

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
		if hasattr(d, "fvwrap_setPitchPercent"):
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

	def synthesize(self, parts: List[Tuple[str, Any]], repeat_count: int = 1) -> None:
		"""Build utterance from parts, commit, and pump read loop."""
		LOGGER.info("synthesize: %d parts, repeat=%d", len(parts), repeat_count)
		self._should_stop = False
		self._dll.fvwrap_begin(self._handle)
		for ptype, pdata in parts:
			if ptype == "text":
				if isinstance(pdata, str):
					pdata = pdata.encode("utf-8", "replace")
				if pdata:
					LOGGER.debug("  addTextUtf8(%d bytes)", len(pdata))
					self._dll.fvwrap_addTextUtf8(self._handle, pdata)
			elif ptype == "index":
				self._dll.fvwrap_addIndex(self._handle, int(pdata))
		rc = int(self._dll.fvwrap_commit(self._handle, int(repeat_count)))
		if rc != 0:
			LOGGER.error("fvwrap_commit returned %d", rc)
			self._send_event("error", code=rc)
			self._send_event("done")
			return
		LOGGER.debug("synthesize: commit ok, entering read loop")
		self._read_loop()
		LOGGER.debug("synthesize: read loop finished")

	def _read_loop(self) -> None:
		"""Pull items from wrapper and push as events to client.

		The wrapper can report DONE *before* the engine has delivered its audio:
		fvwrap's worker calls engine->wait() and then finishRequest(), and the
		engine's put() callbacks can still arrive a few milliseconds later. So
		DONE is treated as "no more work is coming", not "stop reading now" --
		returning on it truncated the utterance, usually to complete silence.
		This mirrors the policy the in-process driver already uses.
		"""
		audio_chunks = 0
		total_bytes = 0
		seen_done = False
		last_activity = time.time()
		while not self._should_stop:
			try:
				n = int(self._dll.fvwrap_read(
					self._handle,
					ctypes.byref(self._out_type),
					ctypes.byref(self._out_value),
					self._audio_buf,
					self._buf_size,
				))
			except Exception:
				LOGGER.exception("fvwrap_read crashed")
				self._send_event("done")
				return

			tp = int(self._out_type.value)
			val = int(self._out_value.value)

			if tp == FVWRAP_ITEM_AUDIO and n > 0:
				audio_chunks += 1
				total_bytes += n
				self._send_event("audio", data=bytes(self._audio_buf[:n]))
				last_activity = time.time()
			elif tp == FVWRAP_ITEM_INDEX:
				self._send_event("index", index=val)
				last_activity = time.time()
			elif tp == FVWRAP_ITEM_DONE:
				seen_done = True
				last_activity = time.time()
			elif tp == FVWRAP_ITEM_ERROR:
				LOGGER.error("Wrapper error %d", val)
				self._send_event("error", code=val)
				seen_done = True
				last_activity = time.time()
			elif tp == FVWRAP_ITEM_NONE:
				if seen_done and (time.time() - last_activity) > 0.05:
					LOGGER.info("read loop done: %d audio chunks, %d bytes total",
								audio_chunks, total_bytes)
					self._send_event("done")
					return
				time.sleep(0.001)

	def stop(self) -> None:
		self._should_stop = True
		if self._handle:
			self._dll.fvwrap_stop(self._handle)
		self._send_event("stopped")

	def set_rate_percent(self, val: int) -> None:
		if self._handle:
			self._dll.fvwrap_setRatePercent(self._handle, int(val))

	def set_volume_percent(self, val: int) -> None:
		if self._handle:
			self._dll.fvwrap_setVolumePercent(self._handle, int(val))

	def set_pitch_percent(self, val: int) -> None:
		if self._handle and self._has_pitch:
			self._dll.fvwrap_setPitchPercent(self._handle, int(val))

	def delete(self) -> None:
		if self._handle:
			LOGGER.info("Destroying FlexVoice engine")
			try:
				self._dll.fvwrap_destroy(self._handle)
			except Exception:
				LOGGER.exception("fvwrap_destroy failed")
			self._handle = None


class HostController:
	"""Receives commands from 64-bit NVDA client and dispatches them.

	The 'synthesize' command runs in a worker thread so that stop can
	interrupt it from the main recv loop.
	"""

	def __init__(self, conn: _ipc.IpcConnection):
		self._conn = conn
		self._runtime: Optional[FlexVoiceRuntime] = None
		self._should_exit = False
		self._synth_thread: Optional[threading.Thread] = None
		self._handlers = {
			"create": self._handle_create,
			"synthesize": self._handle_synthesize,
			"stop": self._handle_stop,
			"setRatePercent": self._handle_set_rate,
			"setVolumePercent": self._handle_set_volume,
			"setPitchPercent": self._handle_set_pitch,
			"delete": self._handle_delete,
		}

	def serve_forever(self) -> None:
		LOGGER.info("Host controller waiting for commands")
		while not self._should_exit:
			try:
				message = self._conn.recv()
			except (EOFError, ConnectionError, OSError) as exc:
				LOGGER.info("Connection closed: %s", exc)
				break
			if not isinstance(message, dict):
				LOGGER.warning("Unexpected message %r", message)
				continue
			msg_type = message.get("type")
			if msg_type != "command":
				LOGGER.warning("Unsupported message type %s", msg_type)
				continue
			msg_id = message.get("id")
			command = message.get("command")
			LOGGER.debug("Received command: %s (id=%s)", command, msg_id)
			handler = self._handlers.get(command)
			if handler is None:
				LOGGER.error("Unknown command %s", command)
				self._conn.send({"type": "response", "id": msg_id, "error": "unknownCommand"})
				continue

			if command == "synthesize":
				self._wait_for_synth_thread()
				self._synth_thread = threading.Thread(
					target=self._run_blocking_handler,
					args=(msg_id, handler, message.get("payload", {})),
					daemon=True,
				)
				self._synth_thread.start()
			else:
				try:
					payload = handler(**message.get("payload", {}))
					self._conn.send({"type": "response", "id": msg_id, "payload": payload or {}})
					if command == "delete" and self._should_exit:
						break
				except Exception as exc:
					LOGGER.exception("Command %s failed", command)
					self._conn.send({"type": "response", "id": msg_id, "error": str(exc)})

	def _run_blocking_handler(self, msg_id: int, handler, payload: Dict[str, Any]) -> None:
		try:
			result = handler(**payload)
			self._conn.send({"type": "response", "id": msg_id, "payload": result or {}})
		except Exception as exc:
			LOGGER.exception("Blocking command failed")
			self._conn.send({"type": "response", "id": msg_id, "error": str(exc)})

	def _handle_create(self, wrapperPath: str, dllDirs: List[str],
					   dataPath: str = "", speakerPath: str = "",
					   language: int = 0x0409, sampleRate: int = 16000,
					   bitsPerSample: int = 16, **_kw) -> Dict:
		self._runtime = FlexVoiceRuntime(self._conn)
		return self._runtime.start(wrapperPath, dllDirs, dataPath, speakerPath,
								   language, sampleRate, bitsPerSample)

	def _handle_synthesize(self, parts: List = None, repeatCount: int = 1, **_kw) -> Dict:
		self._runtime.synthesize(parts or [], repeatCount)
		return {"status": "ok"}

	def _handle_stop(self, **_kw) -> Dict:
		if self._runtime:
			self._runtime.stop()
		self._wait_for_synth_thread()
		return {"status": "ok"}

	def _handle_set_rate(self, value: int, **_kw) -> Dict:
		if self._runtime:
			self._runtime.set_rate_percent(value)
		return {"status": "ok"}

	def _handle_set_volume(self, value: int, **_kw) -> Dict:
		if self._runtime:
			self._runtime.set_volume_percent(value)
		return {"status": "ok"}

	def _handle_set_pitch(self, value: int, **_kw) -> Dict:
		if self._runtime:
			self._runtime.set_pitch_percent(value)
		return {"status": "ok"}

	def _handle_delete(self, **_kw) -> Dict:
		self._wait_for_synth_thread()
		if self._runtime:
			self._runtime.delete()
		self._should_exit = True
		return {"status": "ok"}

	def _wait_for_synth_thread(self) -> None:
		if self._synth_thread and self._synth_thread.is_alive():
			self._synth_thread.join(timeout=30)
		self._synth_thread = None


def main() -> None:
	parser = argparse.ArgumentParser(description="FlexVoice 32-bit helper")
	parser.add_argument("--address", required=True)
	parser.add_argument("--authkey", required=True)
	parser.add_argument("--log-dir", default=None)
	args = parser.parse_args()

	configure_logging(args.log_dir)
	LOGGER.info("Connecting to controller at %s", args.address)

	host, port_str = args.address.split(":")
	address = (host, int(port_str))
	authkey = bytes.fromhex(args.authkey)
	conn = _ipc.connect_to_listener(address, authkey)
	controller = HostController(conn)
	controller.serve_forever()


if __name__ == "__main__":
	main()
