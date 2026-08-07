"""Standalone test: does fvwrap.dll produce audio?

Run with: py -3.14-32 test_fvwrap.py
"""
import ctypes
import os
import sys
import time

# Item types
FVWRAP_ITEM_NONE  = 0
FVWRAP_ITEM_AUDIO = 1
FVWRAP_ITEM_INDEX = 2
FVWRAP_ITEM_DONE  = 3
FVWRAP_ITEM_ERROR = 4

TYPE_NAMES = {0: "NONE", 1: "AUDIO", 2: "INDEX", 3: "DONE", 4: "ERROR"}

BASE = os.path.join(os.path.dirname(__file__), "nvda_addon", "Flexvoice", "synthDrivers")
WRAPPER = os.path.join(BASE, "fvwrap.dll")
ENGLISH = os.path.join(BASE, "English")
SPEAKER = os.path.join(BASE, "English", "Voices", "Tim.tav")

print(f"Python: {sys.version}")
print(f"Pointer size: {ctypes.sizeof(ctypes.c_void_p)} bytes")
print(f"Base dir: {BASE}")
print(f"Wrapper: {WRAPPER} (exists={os.path.isfile(WRAPPER)})")
print(f"English dir: {ENGLISH} (exists={os.path.isdir(ENGLISH)})")
print(f"Speaker: {SPEAKER} (exists={os.path.isfile(SPEAKER)})")

# DLL search paths
os.add_dll_directory(BASE)
os.environ["PATH"] = BASE + os.pathsep + os.environ.get("PATH", "")

print("\nLoading fvwrap.dll...")
dll = ctypes.CDLL(WRAPPER)

# Bind
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

# Try different data paths (same as _createEngineOrDie)
data_candidates = [
    (ENGLISH, "English dir"),
    (BASE, "parent(English) = synthDrivers"),
    (None, "NULL"),
]
sample_rates = [16000, 11025, 8000]

handle = None
used_sr = None
used_label = None

for data_path, label in data_candidates:
    if data_path:
        os.environ["FLEXVOICE_DATA"] = data_path
        print(f"\nFLEXVOICE_DATA = {data_path}")

    for sr in sample_rates:
        data_bytes = data_path.encode("mbcs") if data_path else None
        speaker_bytes = SPEAKER.encode("mbcs")
        print(f"  Trying fvwrap_create({label}, Tim.tav, 0x0409, {sr}, 16)...", end=" ")
        h = dll.fvwrap_create(data_bytes, speaker_bytes, 0x0409, sr, 16)
        if h:
            print(f"OK! handle=0x{h:08x}")
            handle = h
            used_sr = sr
            used_label = label
            break
        else:
            print("NULL (failed)")
    if handle:
        break

if not handle:
    print("\nFATAL: Could not create engine with any data path / sample rate")
    sys.exit(1)

print(f"\nEngine created with: {used_label}, {used_sr}Hz")

# Set parameters
dll.fvwrap_setRatePercent(handle, 50)
dll.fvwrap_setVolumePercent(handle, 100)
dll.fvwrap_setPitchPercent(handle, 50)
print("Parameters set: rate=50, volume=100, pitch=50")

# Synthesize test text
test_text = b"Hello world. This is a test of FlexVoice."
print(f"\nSynthesizing: {test_text.decode()}")

dll.fvwrap_begin(handle)
dll.fvwrap_addTextUtf8(handle, test_text)
dll.fvwrap_addIndex(handle, 1)
rc = int(dll.fvwrap_commit(handle, 1))
print(f"commit returned: {rc}")

if rc != 0:
    print("commit FAILED")
else:
    # Read loop
    buf_size = 8192
    audio_buf = (ctypes.c_ubyte * buf_size)()
    out_type = ctypes.c_int(0)
    out_value = ctypes.c_int(0)

    audio_chunks = 0
    total_bytes = 0
    none_count = 0
    start = time.time()
    max_wait = 5.0  # seconds

    print("Entering read loop...")
    while True:
        n = int(dll.fvwrap_read(handle, ctypes.byref(out_type), ctypes.byref(out_value), audio_buf, buf_size))
        tp = int(out_type.value)
        val = int(out_value.value)

        if tp == FVWRAP_ITEM_AUDIO and n > 0:
            audio_chunks += 1
            total_bytes += n
            if audio_chunks <= 5 or audio_chunks % 50 == 0:
                print(f"  AUDIO: {n} bytes (chunk #{audio_chunks}, total={total_bytes})")
        elif tp == FVWRAP_ITEM_INDEX:
            print(f"  INDEX: {val}")
        elif tp == FVWRAP_ITEM_DONE:
            elapsed = time.time() - start
            print(f"  DONE after {elapsed:.3f}s")
            break
        elif tp == FVWRAP_ITEM_ERROR:
            print(f"  ERROR: code={val}")
            break
        elif tp == FVWRAP_ITEM_NONE:
            none_count += 1
            if time.time() - start > max_wait:
                print(f"  TIMEOUT after {max_wait}s ({none_count} NONE items)")
                break
            time.sleep(0.001)

    print(f"\nResults: {audio_chunks} audio chunks, {total_bytes} bytes, {none_count} NONE items")
    if total_bytes > 0:
        duration_ms = (total_bytes / (used_sr * 2)) * 1000
        print(f"Audio duration: ~{duration_ms:.0f}ms")

# Cleanup
print("\nDestroying engine...")
dll.fvwrap_destroy(handle)
print("Done.")
