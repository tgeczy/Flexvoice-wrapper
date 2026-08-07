"""Test: compare bin/fvwrap.dll vs addon fvwrap.dll audio output.

Run with: py -3.14-32 test_fvwrap3.py
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
BIN_DIR = os.path.join(os.path.dirname(__file__), "bin")
ENGLISH = os.path.join(BASE, "English")
SPEAKER_TIM = os.path.join(BASE, "English", "Voices", "Tim.tav")

# The FlexVoice SDK DLLs are in the addon dir
os.add_dll_directory(BASE)
os.environ["PATH"] = BASE + os.pathsep + os.environ.get("PATH", "")

def load_and_bind(dll_path):
    print(f"Loading: {dll_path} ({os.path.getsize(dll_path)} bytes)")
    d = ctypes.CDLL(dll_path)
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
    return d


def test_dll(dll, label):
    print(f"\n=== Testing: {label} ===")

    buf_size = 8192
    audio_buf = (ctypes.c_ubyte * buf_size)()
    out_type = ctypes.c_int(0)
    out_value = ctypes.c_int(0)

    os.environ["FLEXVOICE_DATA"] = BASE

    data_bytes = BASE.encode("mbcs")
    speaker_bytes = SPEAKER_TIM.encode("mbcs")

    h = dll.fvwrap_create(data_bytes, speaker_bytes, 0x0409, 16000, 16)
    if not h:
        print("  CREATE FAILED (NULL)")
        return

    print(f"  Created: handle=0x{h:08x}")
    dll.fvwrap_setRatePercent(h, 50)
    dll.fvwrap_setVolumePercent(h, 100)
    dll.fvwrap_setPitchPercent(h, 50)

    for text in ["Hello world", "This is a test of FlexVoice speech synthesis", "a b c"]:
        dll.fvwrap_begin(h)
        dll.fvwrap_addTextUtf8(h, text.encode("utf-8"))
        rc = int(dll.fvwrap_commit(h, 1))

        audio_chunks = 0
        total_bytes = 0
        start = time.time()
        events = []

        while True:
            n = int(dll.fvwrap_read(h, ctypes.byref(out_type), ctypes.byref(out_value), audio_buf, buf_size))
            tp = int(out_type.value)
            val = int(out_value.value)

            if tp == FVWRAP_ITEM_AUDIO and n > 0:
                audio_chunks += 1
                total_bytes += n
            elif tp == FVWRAP_ITEM_INDEX:
                events.append(f"IDX({val})")
            elif tp == FVWRAP_ITEM_DONE:
                break
            elif tp == FVWRAP_ITEM_ERROR:
                events.append(f"ERR({val})")
                break
            elif tp == FVWRAP_ITEM_NONE:
                if time.time() - start > 5.0:
                    events.append("TIMEOUT")
                    break
                time.sleep(0.001)

        elapsed = time.time() - start
        evts = f"  [{', '.join(events)}]" if events else ""
        print(f"  '{text[:30]}': {audio_chunks} chunks, {total_bytes} bytes, {elapsed:.3f}s{evts}")

    dll.fvwrap_destroy(h)


print(f"Python: {sys.version}")
print(f"Pointer size: {ctypes.sizeof(ctypes.c_void_p)} bytes")

# We can't load both DLLs in the same process since they have the same symbol names.
# Test whichever one the user wants via command line arg.
if len(sys.argv) > 1 and sys.argv[1] == "bin":
    dll = load_and_bind(os.path.join(BIN_DIR, "fvwrap.dll"))
    test_dll(dll, "bin/fvwrap.dll (older)")
else:
    dll = load_and_bind(os.path.join(BASE, "fvwrap.dll"))
    test_dll(dll, "addon/fvwrap.dll (current)")
