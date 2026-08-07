"""Test: try different CWDs and data paths to get fvwrap audio.

Run with: py -3.14-32 test_fvwrap4.py
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

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nvda_addon", "Flexvoice", "synthDrivers")
WRAPPER = os.path.join(BASE, "fvwrap.dll")
ENGLISH = os.path.join(BASE, "English")
SPEAKER_TIM = os.path.join(BASE, "English", "Voices", "Tim.tav")
SPEAKER_DEFAULT = os.path.join(BASE, "English", "default.tav")
INSTALLED_DATA = r"C:\Program Files (x86)\Common Files\Mindmaker\FlexVoice\Data"

print(f"Python: {sys.version}")
print(f"Pointer size: {ctypes.sizeof(ctypes.c_void_p)} bytes")

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


def synthesize_test(handle, text="Hello world"):
    dll.fvwrap_begin(handle)
    dll.fvwrap_addTextUtf8(handle, text.encode("utf-8"))
    rc = int(dll.fvwrap_commit(handle, 1))
    if rc != 0:
        return f"commit={rc}"

    chunks = 0
    nbytes = 0
    start = time.time()
    while True:
        n = int(dll.fvwrap_read(handle, ctypes.byref(out_type), ctypes.byref(out_value), audio_buf, buf_size))
        tp = int(out_type.value)
        if tp == FVWRAP_ITEM_AUDIO and n > 0:
            chunks += 1
            nbytes += n
        elif tp == FVWRAP_ITEM_DONE:
            break
        elif tp == FVWRAP_ITEM_ERROR:
            return f"ERROR code={int(out_value.value)}"
        elif tp == FVWRAP_ITEM_NONE:
            if time.time() - start > 5.0:
                return "TIMEOUT"
            time.sleep(0.001)
    elapsed = time.time() - start
    return f"{chunks} chunks, {nbytes} bytes, {elapsed:.3f}s"


def try_config(data_path, speaker_path, sr, fvd, cwd, desc):
    print(f"\n--- {desc} ---")
    print(f"  data={data_path}")
    print(f"  speaker={speaker_path}")
    print(f"  FVD={fvd}")
    print(f"  cwd={cwd}")

    if cwd:
        os.chdir(cwd)
    if fvd:
        os.environ["FLEXVOICE_DATA"] = fvd
    elif "FLEXVOICE_DATA" in os.environ:
        del os.environ["FLEXVOICE_DATA"]

    data_bytes = data_path.encode("mbcs") if data_path else None
    speaker_bytes = speaker_path.encode("mbcs") if speaker_path else None
    h = dll.fvwrap_create(data_bytes, speaker_bytes, 0x0409, sr, 16)
    if not h:
        print("  CREATE FAILED")
        return

    print(f"  Created OK (handle=0x{h:08x})")
    dll.fvwrap_setRatePercent(h, 50)
    dll.fvwrap_setVolumePercent(h, 100)
    dll.fvwrap_setPitchPercent(h, 50)
    print(f"  Synth: {synthesize_test(h)}")
    dll.fvwrap_destroy(h)


# Test 1: CWD = synthDrivers, data = synthDrivers
try_config(BASE, SPEAKER_TIM, 16000, BASE, BASE, "CWD=synthDrivers, data=synthDrivers")

# Test 2: CWD = English, data = synthDrivers
try_config(BASE, SPEAKER_TIM, 16000, ENGLISH, ENGLISH, "CWD=English, data=synthDrivers, FVD=English")

# Test 3: data = synthDrivers, FVD = installed data path
try_config(BASE, SPEAKER_TIM, 16000, INSTALLED_DATA, BASE, "data=synthDrivers, FVD=installed_data")

# Test 4: data = installed FlexVoice Data, speaker from addon
fv_dir = os.path.dirname(INSTALLED_DATA)  # FlexVoice dir
try_config(INSTALLED_DATA, SPEAKER_TIM, 16000, INSTALLED_DATA, INSTALLED_DATA, "data=installed_data, FVD=installed_data, Tim")
try_config(fv_dir, SPEAKER_TIM, 16000, fv_dir, fv_dir, "data=FlexVoice_dir, FVD=FlexVoice_dir, Tim")

# Test 5: NULL data path (let SDK find data via registry/env)
try_config(None, SPEAKER_TIM, 16000, BASE, BASE, "data=NULL, FVD=synthDrivers")
try_config(None, SPEAKER_TIM, 16000, ENGLISH, BASE, "data=NULL, FVD=English")
try_config(None, SPEAKER_TIM, 16000, None, BASE, "data=NULL, FVD=unset")

# Test 6: default.tav speaker
try_config(BASE, SPEAKER_DEFAULT, 16000, BASE, BASE, "data=synthDrivers, default.tav")

# Test 7: NULL speaker
try_config(BASE, None, 16000, BASE, BASE, "data=synthDrivers, speaker=NULL")

print("\nAll tests complete.")
