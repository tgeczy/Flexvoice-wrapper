# MindMaker FlexVoice 3.01 wrapper + NVDA synth driver

This repo contains:

- **fvwrap.dll** (source): a small C++ wrapper around the MindMaker FlexVoice 3.01 SDK that:
  - streams PCM audio in small chunks
  - provides **reliable NVDA IndexCommand support** (indent/nav beeps, etc.)
  - applies **rate / volume / pitch** safely
  - normalizes “fragile” tokens so the engine stays stable (digits, letter runs like “NVDA”, single consonant letters)

- **NVDA synth driver** (Python): an NVDA add-on synth driver that uses `fvwrap.dll` to speak through FlexVoice, in **English and Hungarian**.

This project exists because FlexVoice is old but still has a unique voice quality, and NVDA needs modern behavior (indexing, stability, safe cancel/restart) to use it comfortably.

> **Important:** This repo does **not** include MindMaker’s proprietary headers, libraries, engine binaries or voice data. You must obtain the FlexVoice 3.01 SDK separately to build the wrapper, and the voice data separately to run it.

---

## Status — v2.0

✅ **Hungarian** (Zita, Péter) alongside English (Tom, Julie, Ben, Leslie, Tim, Kim)  
✅ Fourteen voices in one list, plus a **Language** setting; the language follows the voice  
✅ Runs on 64-bit NVDA 2026.1+ through NVDA's built-in 32-bit bridge, and in-process on 32-bit NVDA  
✅ Stable speech without process crashes  
✅ IndexCommand support (NVDA indexing works correctly)  
✅ No “chipmunk” effect when changing rate (uses engine-native `speechRate`)  
✅ Digits and “NVDA-style” acronyms speak reliably, in both languages  
✅ Pitch works via speaker parameters (`defaultPitch`)

The Hungarian voice data came from **Király József**, who worked on FlexVoice
and supplied the 3.01.108 demo from his own archive. The engine could always
speak Hungarian — `LNG_HUNGARIAN` sits beside `LNG_ENGLISH` in the SDK headers —
only the data and a correct text path were missing.

See **[docs/engine-notes.md](docs/engine-notes.md)** for what we learned about
the engine itself: the data layout, the two text-submission paths and why only
one of them is safe, the embedded-command syntax, the foreign-word detector, and
the open questions we could not answer.

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
- NVDA 2024.4 or later (tested on 2026.1 64-bit; 32-bit builds run the driver in-process)
- FlexVoice 3.01 voice data, staged into `nvda_addon/Flexvoice/synthDrivers/` as
  `English/` and `Hungarian/` sibling folders, alongside the engine DLLs.
  `EngineFactory` is handed the folder that *contains* them — see
  [docs/engine-notes.md](docs/engine-notes.md) §2.

### Build (wrapper)
- Visual Studio (MSVC) with CMake (or your preferred build system)
- MindMaker FlexVoice 3.01 SDK:
  - headers (e.g. `Engine.h`, `Speaker.h`, `Bookmark.h`, …)
  - import libs (e.g. `FlexVoice_3_01_001.lib`, etc.)

This repo includes *only* the wrapper’s own header(s) and source; you must point your build at the SDK include/lib locations yourself.

---

## Folder layout (suggested)

