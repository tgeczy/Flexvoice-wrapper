"""Standalone test: try different configurations to get fvwrap audio.

Run with: py -3.14-32 test_fvwrap2.py
"""
import ctypes
import os
import sys
import time

FVWRAP_ITEM_NONE  = 0
FVWRAP_ITEM_AUDIO = 1
FVWRAP_ITEM_INDEX = 2
FVWRAP_ITEM_DONE  = 3
FVWRAP_ITEM_ERROR = 4

BASE = os.path.join(os.path.dirname(__file__), "nvda_addon", "Flexvoice", "synthDrivers")
WRAPPER = os.path.join(BASE, "fvwrap.dll")
ENGLISH = os.path.join(BASE, "English")
SPEAKER_TIM = os.path.join(BASE, "English", "Voices", "Tim.tav")
SPEAKER_KIM = os.path.join(BASE, "English", "Voices", "Kim.tav")
SPEAKER_DEFAULT = os.path.join(BASE, "English", "default.tav")

print(f"Python: {sys.version}")
print(f"Pointer size: {ctypes.sizeof(ctypes.c_void_p)} bytes\n")

os.add_dll_directory(BASE)
os.environ["PATH"] = BASE + os.pathsep + os.environ.get("PATH", "")

dll = ctypes.CDLL(WRAPPER)

dll.fvwrap_create.restype = ctypes.c_void_p
dll.fvwrap_create.argtypes = (ctypes.c_char_p, ctypes.c_char_p, ctypes.c_int, ctypes.c_int, ctypes.c_int)
dll.fvwrap_destroy.restype = None
dll.fvwrap_destroy.argtypes = (ctypes.c_void_p,)
dll.fvwrap_stop.restype = None
dll.fvwrap_stop.argtypes = (ctypes.c_void_p,)
dll.fvwrap_setRatePercent.restype = ctypes.c_int
dll.fvwrap_setRatePercent.argtypes = (ctypes.c_void_p, ctypes.c_int)
dll.fvwrap_setVolumePercent.restype = ctypes.c_int
dll.fvwrap_setVolumePercent.argtypes = (ctypes.c_void_p, ctypes.c_int)
dll.fvwrap_setPitchPercent.restype = ctypes.c_int
dll.fvwrap_setPitchPercent.argtypes = (ctypes.c_void_p, ctypes.c_int)
dll.fvwrap_begin.restype = None
dll.fvwrap_begin.argtypes = (ctypes.c_void_p,)
dll.fvwrap_addTextUtf8.restype = None
dll.fvwrap_addTextUtf8.argtypes = (ctypes.c_void_p, ctypes.c_char_p)
dll.fvwrap_addIndex.restype = None
dll.fvwrap_addIndex.argtypes = (ctypes.c_void_p, ctypes.c_int)
dll.fvwrap_commit.restype = ctypes.c_int
dll.fvwrap_commit.argtypes = (ctypes.c_void_p, ctypes.c_int)
dll.fvwrap_read.restype = ctypes.c_int
dll.fvwrap_read.argtypes = (
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_int),
    ctypes.POINTER(ctypes.c_int),
    ctypes.POINTER(ctypes.c_ubyte),
    ctypes.c_int,
)

buf_size = 8192
audio_buf = (ctypes.c_ubyte * buf_size)()
out_type = ctypes.c_int(0)
out_value = ctypes.c_int(0)


def try_synthesize(handle, text, add_index=False):
    """Try to synthesize text and return (audio_chunks, total_bytes, elapsed)."""
    dll.fvwrap_begin(handle)
    dll.fvwrap_addTextUtf8(handle, text.encode("utf-8"))
    if add_index:
        dll.fvwrap_addIndex(handle, 1)
    rc = int(dll.fvwrap_commit(handle, 1))
    if rc != 0:
        return (0, 0, 0, f"commit returned {rc}")

    audio_chunks = 0
    total_bytes = 0
    start = time.time()
    events = []

    while True:
        n = int(dll.fvwrap_read(handle, ctypes.byref(out_type), ctypes.byref(out_value), audio_buf, buf_size))
        tp = int(out_type.value)
        val = int(out_value.value)

        if tp == FVWRAP_ITEM_AUDIO and n > 0:
            audio_chunks += 1
            total_bytes += n
        elif tp == FVWRAP_ITEM_INDEX:
            events.append(f"INDEX({val})")
        elif tp == FVWRAP_ITEM_DONE:
            break
        elif tp == FVWRAP_ITEM_ERROR:
            events.append(f"ERROR({val})")
            break
        elif tp == FVWRAP_ITEM_NONE:
            if time.time() - start > 5.0:
                events.append("TIMEOUT")
                break
            time.sleep(0.001)

    elapsed = time.time() - start
    return (audio_chunks, total_bytes, elapsed, ", ".join(events) if events else "")


def try_create(data_path, speaker_path, sr=16000, flexvoice_data=None):
    """Try to create engine. Returns handle or None."""
    if flexvoice_data is not None:
        os.environ["FLEXVOICE_DATA"] = flexvoice_data
    elif "FLEXVOICE_DATA" in os.environ:
        del os.environ["FLEXVOICE_DATA"]

    data_bytes = data_path.encode("mbcs") if data_path else None
    speaker_bytes = speaker_path.encode("mbcs") if speaker_path else None
    h = dll.fvwrap_create(data_bytes, speaker_bytes, 0x0409, sr, 16)
    return h if h else None


# Test configurations
configs = [
    # (data_path, speaker, sr, FLEXVOICE_DATA, description)
    (BASE, SPEAKER_TIM, 16000, BASE, "data=synthDrivers, FVD=synthDrivers, Tim"),
    (BASE, SPEAKER_TIM, 16000, ENGLISH, "data=synthDrivers, FVD=English, Tim"),
    (BASE, SPEAKER_TIM, 16000, None, "data=synthDrivers, FVD=unset, Tim"),
    (BASE, SPEAKER_DEFAULT, 16000, BASE, "data=synthDrivers, FVD=synthDrivers, default.tav"),
    (BASE, SPEAKER_DEFAULT, 16000, ENGLISH, "data=synthDrivers, FVD=English, default.tav"),
    (BASE, SPEAKER_KIM, 16000, BASE, "data=synthDrivers, FVD=synthDrivers, Kim"),
    (BASE, SPEAKER_KIM, 16000, ENGLISH, "data=synthDrivers, FVD=English, Kim"),
    (BASE, SPEAKER_TIM, 11025, BASE, "data=synthDrivers, FVD=synthDrivers, Tim, 11025"),
    (BASE, SPEAKER_TIM, 8000, BASE, "data=synthDrivers, FVD=synthDrivers, Tim, 8000"),
]

for data_path, speaker, sr, fvd, desc in configs:
    print(f"--- {desc} ---")
    h = try_create(data_path, speaker, sr, fvd)
    if not h:
        print(f"  CREATE FAILED\n")
        continue

    dll.fvwrap_setRatePercent(h, 50)
    dll.fvwrap_setVolumePercent(h, 100)
    dll.fvwrap_setPitchPercent(h, 50)

    # Test 1: plain text, no index
    chunks, nbytes, elapsed, evts = try_synthesize(h, "Hello world")
    print(f"  Plain text:  {chunks} chunks, {nbytes} bytes, {elapsed:.3f}s  {evts}")

    # Test 2: text with index
    chunks, nbytes, elapsed, evts = try_synthesize(h, "This is a longer test sentence.", add_index=True)
    print(f"  With index:  {chunks} chunks, {nbytes} bytes, {elapsed:.3f}s  {evts}")

    # Test 3: very simple text
    chunks, nbytes, elapsed, evts = try_synthesize(h, "a")
    print(f"  Single 'a':  {chunks} chunks, {nbytes} bytes, {elapsed:.3f}s  {evts}")

    dll.fvwrap_destroy(h)
    print()

print("All tests complete.")
