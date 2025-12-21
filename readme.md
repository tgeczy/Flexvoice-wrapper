# MindMaker FlexVoice 3.01 wrapper + NVDA synth driver

This repo contains:

- **fvwrap.dll** (source): a small C++ wrapper around the MindMaker FlexVoice 3.01 SDK that:
  - streams PCM audio in small chunks
  - provides **reliable NVDA IndexCommand support** (indent/nav beeps, etc.)
  - applies **rate / volume / pitch** safely
  - normalizes “fragile” tokens so the engine stays stable (digits, letter runs like “NVDA”, single consonant letters)

- **NVDA synth driver** (Python): an NVDA add-on synth driver that uses `fvwrap.dll` to speak through FlexVoice.

This project exists because FlexVoice is old but still has a unique voice quality, and NVDA needs modern behavior (indexing, stability, safe cancel/restart) to use it comfortably.

> **Important:** This repo does **not** include MindMaker’s proprietary headers or libraries. You must obtain the FlexVoice 3.01 SDK separately to build the wrapper.

---

## Status

✅ Stable speech without process crashes  
✅ IndexCommand support (NVDA indexing works correctly)  
✅ No “chipmunk” effect when changing rate (uses engine-native `speechRate`)  
✅ Digits and “NVDA-style” acronyms speak reliably  
✅ Pitch works via speaker parameters (`defaultPitch`)

---

## How it works

### 1) Wrapper (`fvwrap.dll`)
The wrapper exposes a tiny C API:

- `fvwrap_create(...)` / `fvwrap_destroy(...)`
- `fvwrap_begin()` / `fvwrap_addTextUtf8()` / `fvwrap_addIndex()` / `fvwrap_commit()`
- `fvwrap_read(...)` to pull a stream of:
  - audio chunks (`FVWRAP_ITEM_AUDIO`)
  - index events (`FVWRAP_ITEM_INDEX`)
  - done (`FVWRAP_ITEM_DONE`)
  - error (`FVWRAP_ITEM_ERROR`)

Internally it:
- runs the FlexVoice engine on a single worker thread
- applies settings only at safe boundaries
- normalizes input text to avoid known engine fragility
- implements bounded buffering to avoid runaway memory and timing issues

### 2) NVDA synth driver
The NVDA driver:
- loads `fvwrap.dll` with `ctypes`
- sends speech to the wrapper
- reads the wrapper stream and feeds audio to `nvwave.WavePlayer`
- forwards index events to NVDA via `synthIndexReached`
- sends `synthDoneSpeaking` when finished

---

## Requirements

### Runtime
- Windows
- NVDA (recent versions)
- A working **FlexVoice 3.01 runtime install** (English data + voice `.tav` files)

### Build (wrapper)
- Visual Studio (MSVC) with CMake (or your preferred build system)
- MindMaker FlexVoice 3.01 SDK:
  - headers (e.g. `Engine.h`, `Speaker.h`, `Bookmark.h`, …)
  - import libs (e.g. `FlexVoice_3_01_001.lib`, etc.)

This repo includes *only* the wrapper’s own header(s) and source; you must point your build at the SDK include/lib locations yourself.

---

## Folder layout (suggested)

