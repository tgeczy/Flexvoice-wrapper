# FlexVoice 3.01 engine notes

What we learned about MindMaker's FlexVoice engine while adding Hungarian to the
NVDA driver, written down so it does not have to be rediscovered. Everything
here is measured against the shipped 3.01.001 engine (`FlexVoice_3_01_001.dll`,
2,568,192 bytes) unless stated otherwise.

Sources: the FlexVoice 3.01 SDK, the FlexVoice 3.01.108 demo installer supplied
by Király József, and disassembly/strings of the engine DLL.

---

## 1. Languages

`ttsapi/FVLanguage.h` declares three ids:

```c
const Language LNG_INVALID   = 0x0000;
const Language LNG_ENGLISH   = 0x0409;
const Language LNG_HUNGARIAN = 0x040e;
```

Hungarian was never a separate build — the shipped engine has always supported
it. Only the data was missing. `Say.exe` additionally accepts `-lang malay`, so
a Malay build existed; no data for it has surfaced.

## 2. Data layout

`EngineFactory(path)` expects the directory that **contains** the per-language
folders, not a language folder itself:

```
<dataRoot>/
    English/    VoiceList.tvl, *.bin, *.tav, prosody models, L2P tables
    Hungarian/  VoiceList.tvl, Zita*.bin, *.tav, prosody models, MorphSet.dat
```

Passing a single language folder appears to work while only one language is
present, then silently loads the wrong data once a sibling exists.

The 3.01.108 demo installs its data to
`C:\Program Files (x86)\Common Files\Mindmaker\FlexVoice\Data_3_01_001\`.
`C:\Program Files (x86)\Mindmaker\` holds only the GUI stub — the engine DLL
itself is never installed, and no COM/SAPI server is registered.

### Files a language folder needs

Hungarian ships no `L2P.dat`, `L2P.id3`, `OnsetClusters.dat`, `PRCE.dat`,
`PRCS.dat` or `NotPRTR.dat` — Hungarian orthography is near-phonemic, so
letter-to-phoneme is rule-based (`MorphSet.dat`, `ptp.md`) rather than a lookup
table. English ships both `CHL2.dat` and `RHL2.dat`; Hungarian only `RHL2.dat`.
`CHL2.dat` appears optional — the demo's English data omits it and still loads.

`langdet_eng.dat`, `langdet_ger.dat` and `langdet_hun.dat` are **mandatory**.
Removing the foreign ones to disable detection makes `EngineFactory` fail for
*every* language, silently, with no audio at all.

## 3. Voices

A `.tav` is a plain text parameter file. `voiceDescr` selects which `.bin`
diphone bank the voice plays through — the file name does not:

| .tav | voiceDescr | bank used |
|---|---|---|
| `Tim.tav` | `Tom` | Tom.bin |
| `Kim.tav` | `Julie` | Julie.bin |
| `Peter.tav` | `Zita` | Zita.bin |
| `*Phone.tav` | `<name>8` | 8 kHz bank |

So several "voices" are one bank with different shaping. Two findings worth
keeping:

- `Tim.tav` is **byte-identical to `Ben.tav` except for `voiceDescr`**. Tim is
  Ben's voice driving Tom's bank, which is why they sound alike.
- Tim/Kim are *not* renames of Tom/Julie. They share banks but differ
  substantially in shaping (Tim: pitch 105, 163 wpm, tilt 0, 12-band EQ at
  Q=1.5; Tom: pitch 95, 137 wpm, tilt 31, Q=0.75).

`Zita_R.tav` cannot be loaded by this engine build: it requests
`loudnessModelType="relative"` and the engine reports
`Can not create feature LogVolumeModelRelative1Point`. The 3.01.108 engine
inside the demo installer might support it — untested.

Stock `VoiceList.tvl` files name voices that were never shipped (Bill, Jack,
Jill, Kit, Julius). Entries whose `.tav` is absent must be skipped or
`createEngine` fails.

## 4. Two ways to submit text, and only one of them is solid

```c
engine->speakRequest(text, count);          // whole-text path
engine->addFragment(text); ... speakRequest(count);  // fragment path
```

`Say.exe` uses the whole-text path. The fragment path is the one a screen
reader needs, because index bookmarks (`addBookmark`) have to interleave with
text. **The fragment path is markedly more fragile.** Everything below behaves
correctly through `speakRequest(text)` and misbehaves through fragments:

| input | fragment-path behaviour |
|---|---|
| raw digits | access violation |
| ALL-CAPS runs | near-silence |
| intra-word hyphen (`e-mail`, `NVDA-val`) | 2 bytes of audio |
| consonant-only tokens (`vlc`, `cmd`) | not spoken |
| words the foreign-word detector flags | mangled, see §6 |

The English side of this wrapper had already worked around digits years ago —
the comment "or digits will slip through and crash FlexVoice again" refers to
exactly this. It is a property of the fragment path, not of English.

### DONE arrives before the audio

`Engine::wait()` is documented to block until the request is processed, but the
output site's `put()` callbacks were observed trailing it consistently: DONE at
~2 ms after commit, first audio at 7–17 ms, on *every* utterance. A consumer
that treats DONE as end-of-utterance will drop speech whenever the engine
starts slowly. `fvwrap` now holds `finishRequest()` until audio has actually
been enqueued, with a 500 ms grace for text that legitimately renders to
nothing.

## 5. Embedded commands

The engine accepts inline commands. **The syntax is backslash-delimited**, as
used by the SDK's own `samples/C++/Say/PhonemeFileWriter.cpp`:

```
\vce=speaker="Zita"\ \rspd=1.0\ \pit=100\ \rpit=1.0\ \langdet=0\
```

Note the trailing backslash *and* a space. The DLL's string table also contains
`[:phone arpa MM]`-style tags; those are **not** parsed and are spoken aloud
verbatim — a red herring that cost several attempts.

Command names found in the DLL string table:

```
style=  age=  gender=  dialect=  accent=  language=  langdet=  lang=  wlang=
phstr=  speaker=  POS=  wordmode=  sayas=  spell  wspell  aspell  waspell
endlang  endspell  endsayas
```

Only `langdet`, `lang`, `wlang`, `rspd`, `pit`, `rpit` and `vce` have been
exercised. Their exact value grammars are not documented anywhere we have.

## 6. The foreign-word detector

The Hungarian data includes English and German n-gram tables and the engine
uses them to detect foreign words and pronounce them in that language. It
demonstrably fires:

```
set       ->  S ` e t            Hungarian /ʃ/
settings  ->  s ` e . t i N z    English /s/, English final /z/
```

Through the fragment path the foreign branch is broken. Envelope correlation
against the same text spoken natively (10 ms frames):

| word | corr | note |
|---|---|---|
| `stop` | 0.546 | ran 60 % long |
| `desktop` | 0.573 | heard as "dekeiop" |
| `desk` | 0.836 | |
| `top` | 0.866 | |
| `kap` `kép` `csop` `sapka` `lak` `pad` `hat` `alma` | 0.997–1.000 | Hungarian, unaffected |
| `stopa` `topa` `deska` | 0.998 | one added vowel repairs it entirely |

That last row is the proof it is detection and not phonetics: same letters, same
clusters, but no longer English-looking.

Fix: prefix each Hungarian utterance with `\langdet=0\ `. Measured after:
`stop` 0.999, `desktop` 1.000, `top` 1.000, `desk` 0.996, `settings` 0.992,
Hungarian unchanged at 1.000, and the command is consumed rather than spoken.
Foreign words then read with Hungarian letter-to-sound (`desktop` →
"deszktop"), which is what a Hungarian screen-reader user expects.

`\lang=eng\` is the reverse dial if a specific token should be anglicised.

## 7. Text encoding

Hungarian text must reach the engine as **CP1250** (identical to ISO-8859-2 at
every Hungarian letter). Verified with `Say.exe -lang hun -phout`:

```
ASCII   "Arvizturo"   ->  a r . v i s . t u . r o        all short, wrong
CP1250  "Árvíztűrő"   ->  A: r . v i: s . t U: . r O:    correct
UTF-8   raw bytes     ->  U ...                          garbage
Latin-1 with '?'      ->  A: r . v i: s t _ _ _          destroyed
```

On an unmappable character, fall back to an ASCII approximation — never `'?'`,
which is what wrecked the Latin-1 attempt.

Two traps on the way in: a Python-side sanitiser that clamps to ≤ 0xFF silently
eats `ő` (U+0151) and `ű` (U+0171); and English text normalisation must not run
on Hungarian, or `123` is spoken as "one two three" with Hungarian phonemes.

## 8. Hungarian text normalisation is excellent

Native, no help required — as long as the English normaliser is kept away:

| input | spoken |
|---|---|
| `123` | százhuszonhárom |
| `45 forint` | negyvenöt forint |
| `2026. augusztus 6.` | kétezer-huszonhat augusztus **hatodika** |
| `12,5%` | tizenkettő egész öt tized százalék |

Caveat: this richer expansion is only reachable through `speakRequest(text)`.
On the fragment path digits crash, so `fvwrap` expands them to Hungarian number
words itself and loses the ordinal/date cleverness.

Single letters are rendered as bare phonemes (`k` → /k/, nearly inaudible), so
the wrapper substitutes Hungarian letter names for character review.

---

## Open questions

Things we could not answer from the SDK, the binaries or experiment. Most of
these are only answerable by someone who worked on the engine.

1. **Is the fragment path's fragility known and bounded?** We found four
   triggers by experiment (digits, ALL-CAPS, intra-word hyphens, consonant-only
   tokens). Is there a list? Was `addFragment` intended for production use, or
   mainly for the bookmark demo?
2. **Why does `Engine::wait()` return before the audio callbacks?** Is there a
   supported way to be notified of true end-of-utterance — `Notifier.h` /
   `INotify` with a particular bookmark type?
3. **`langdet` value grammar.** `\langdet=0\` works. Are there other values
   (per-language, thresholds)? Can detection be configured rather than switched
   off, so English words are anglicised *correctly* instead of disabled?
4. **The full embedded-command grammar**, especially `sayas=`, `wordmode=`,
   `POS=` and the `spell` family — these look directly useful for a screen
   reader.
5. **`Zita_R` and `LogVolumeModelRelative1Point`** — which engine build
   supports the relative loudness model? Is the 3.01.108 engine inside the demo
   installer newer in that respect?
6. **Whose voices are the banks?** Tom, Julie, Ben, Leslie, Zita — recorded by
   whom, and when? (Asked; József no longer recalls.)
7. **Malay.** `Say.exe` accepts `-lang malay`. Did that data ever ship?
8. **The user dictionary API** (`insertDictionary`, `UserDictionary.h`) — is it
   a practical route for fixing individual mispronunciations, and what is the
   entry format?
