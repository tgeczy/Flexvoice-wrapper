# FlexVoice for NVDA 2.0 — Hungarian arrives

FlexVoice now speaks **Hungarian**. Zita and Péter join the English voices, and
a **Language** combo in NVDA's Voice settings switches between them.

Fourteen voices in total:

| Language | Voices |
|---|---|
| English (US) | Tom, Julie, Ben, Leslie, Tim, Kim + four 8 kHz "Phone" variants |
| Hungarian | Zita, Péter + two "Phone" variants |

All fourteen voices appear in one list, whichever language is selected — the
language travels with the voice, so choosing Zita switches to Hungarian and
choosing Leslie switches back. The **Language** combo is a shortcut: it moves
you to that language's default voice (Tom for English, Zita for Hungarian).

Tested on NVDA 2026.1 (64-bit, through NVDA's built-in 32-bit bridge). It should
also work on 32-bit NVDA from 2024.4 onwards, where the driver runs in-process
and needs no bridge — but that has not been tested, so reports are welcome.

## Why this was possible

The engine we already shipped could speak Hungarian all along — `LNG_HUNGARIAN`
is declared right next to `LNG_ENGLISH` in the 3.01 SDK headers. What was
missing was the voice data, and a text path that did not destroy Hungarian on
the way in.

The data came from **Király József**, who wrote PC-TALKER and worked on
FlexVoice, and who sent over the Hungarian demo build from his own archive.
Without that, none of this exists. Köszönöm.

## Hungarian text handling

Vowel length and umlauts are phonemic in Hungarian — *tükör* and *tukor* are
different words — but the wrapper folded every accent to ASCII, which is fine
for "Beyoncé" and fatal for Hungarian. Text is now converted to CP1250 with the
accents intact. Checked against the engine's own phoneme output:

```
ASCII   "Arvizturo"   ->  a r . v i s . t u . r o        (all short, wrong)
CP1250  "Árvíztűrő"   ->  A: r . v i: s . t U: . r O:    (correct)
```

The English number/acronym expansion is skipped for Hungarian, because the
Hungarian engine already does it, and does it better:

- `123` → *százhuszonhárom*
- `2026. augusztus 6.` → *kétezer-huszonhat augusztus hatodika*
- `12,5%` → *tizenkettő egész öt tized százalék*

Single letters now speak Hungarian letter names, so reviewing text character by
character is audible — previously a lone `k` was rendered as a bare /k/.

## About Tim and Kim

They are still here, and they are not duplicates. Tim and Kim share diphone
banks with Tom and Julie, but the voice shaping is quite different — Tim sits at
105 Hz and 163 wpm against Tom's 95 Hz and 137 wpm, with a different equalizer
entirely. If Tim is the voice you have used for years, it is unchanged.

(A curiosity found along the way: `Tim.tav` is byte-identical to `Ben.tav` apart
from which bank it points at. Tim is Ben's voice driving Tom's bank, which is
why the two sound so alike.)

## Fixes

- Three crashes of the engine's *fragment* path — the submission mode the
  wrapper must use so NVDA's index bookmarks work — were found and worked
  around for Hungarian: raw digits (an access violation), ALL-CAPS runs and
  intra-word hyphens (`e-mail`, `NVDA-val` — rendered as silence). Digits are
  expanded to Hungarian number words in the wrapper (`123` →
  *százhuszonhárom*), capitals are lowercased (the engine reads them
  identically), and intra-word hyphens become spaces, which is how the engine's
  own normalizer reads them anyway.
- `DONE` could reach the reader before the first audio chunk, so an utterance
  that started slowly — right after a language switch, or with the host under
  load — was dropped as silence: the "skipping CD" bug. `DONE` is now held
  until audio has actually been queued.
- Switching language now refreshes the voice list in an open settings dialog;
  previously it kept offering the old language's voices.
- The 32-bit host returned on the first `DONE`, but the wrapper emits `DONE`
  before the engine's audio arrives — utterances could be truncated to silence.
- A missing `_ipc.py` removed the whole synthesizer from NVDA's list instead of
  disabling one unused code path.
- On 64-bit NVDA without the built-in bridge, the driver failed to import at all.
- Voice lists no longer offer voices whose data is absent — the stock list names
  five that were never shipped.
- The engine is now given the folder that *contains* the language folders;
  passing a single one worked only while English was the only language present.

---

# FlexVoice NVDA-hoz 2.0 — megérkezett a magyar

A FlexVoice mostantól **magyarul is beszél**. Zita és Péter csatlakozik az angol
hangokhoz, és az NVDA hangbeállításaiban egy **Nyelv** lenyíló listával lehet
váltani közöttük.

Összesen tizennégy hang:

| Nyelv | Hangok |
|---|---|
| Angol (amerikai) | Tom, Julie, Ben, Leslie, Tim, Kim + négy 8 kHz-es „Phone" változat |
| Magyar | Zita, Péter + két „Phone" változat |

Mind a tizennégy hang egyetlen listában jelenik meg, bármelyik nyelv is aktív —
a nyelv a hanggal együtt jár, tehát Zitát választva magyarra, Leslie-t választva
angolra vált. A **Nyelv** lenyíló lista gyorsítás: az adott nyelv alapértelmezett
hangjára ugrik (angolnál Tom, magyarnál Zita).

Kipróbálva NVDA 2026.1 alatt (64 bites, az NVDA beépített 32 bites hídján
keresztül). Elvben a 2024.4-től kezdődő 32 bites NVDA-kon is működik, ahol a
meghajtó híd nélkül, közvetlenül fut — ez azonban nincs kipróbálva, így a
visszajelzéseket szívesen fogadjuk.

## Miért most

A motor, amit eddig is szállítottunk, végig tudott magyarul — a `LNG_HUNGARIAN`
ott áll a `LNG_ENGLISH` mellett a 3.01-es SDK fejléceiben. A hangadatok
hiányoztak, és egy olyan szövegút, amely nem teszi tönkre a magyart útközben.

Az adatok **Király Józseftől** származnak, aki a PC-TALKER szerzője és a
FlexVoice fejlesztésében is részt vett, és aki a saját archívumából küldte el a
magyar demót. Nélküle ez nem létezne. Köszönöm.

## Magyar szövegkezelés

A magánhangzó-hosszúság és az ékezetek jelentésmegkülönböztetők a magyarban — a
*tükör* és a *tukor* két különböző szó —, a csomagoló mégis minden ékezetet
ASCII-ra egyszerűsített. Ez a „Beyoncé" esetében rendben van, a magyarnál
végzetes. A szöveg mostantól CP1250 kódolással, ékezetekkel együtt jut el a
motorhoz. A motor saját fonémakimenetével ellenőrizve:

```
ASCII   "Arvizturo"   ->  a r . v i s . t u . r o        (csupa rövid, hibás)
CP1250  "Árvíztűrő"   ->  A: r . v i: s . t U: . r O:    (helyes)
```

Az angol szám- és betűszó-kifejtés magyarnál kimarad, mert a magyar motor ezt
eleve elvégzi, méghozzá jobban:

- `123` → *százhuszonhárom*
- `2026. augusztus 6.` → *kétezer-huszonhat augusztus hatodika*
- `12,5%` → *tizenkettő egész öt tized százalék*

Az egyedülálló betűk mostantól a magyar betűnevüket mondják, így a szöveg
karakterenkénti áttekintése hallható — korábban egy magányos `k` csak egy
puszta /k/ hang volt.

## Tim és Kim

Megmaradtak, és nem duplikátumok. Tim és Kim ugyanazt a diád-készletet
használja, mint Tom és Julie, a hangformálás azonban jócskán eltér — Tim 105
Hz-en és 163 szó/percen szól Tom 95 Hz-e és 137 szó/perce ellenében, teljesen
más hangszínszabályzóval. Ha évek óta Tim a megszokott hang, az változatlan.

## Javítások

- A motor *fragment*-útvonalának három összeomlása — ezt a beviteli módot kell
  használnunk, hogy az NVDA indexkönyvjelzői működjenek — vált ismertté és
  került megkerülésre a magyarnál: a nyers számjegyek (memóriahiba), a csupa
  nagybetűs sorok és a szóközi kötőjelek (`e-mail`, `NVDA-val` — némaságként
  szólaltak meg). A számjegyeket a csomagoló magyar számnevekké alakítja
  (`123` → *százhuszonhárom*), a nagybetűket kisbetűsíti (a motor azonosan
  olvassa őket), a szón belüli kötőjelből pedig szóköz lesz — a motor saját
  normalizálója is így olvasná.
- A `DONE` a hang első csomagja előtt érhetett az olvasóhoz, így a lassan
  induló mondat — nyelvváltás után, vagy terhelt gazdafolyamatnál — némaságként
  veszett el: az „ugráló CD" hiba. A `DONE` mostantól megvárja a hangot.
- A nyelvváltás frissíti a nyitott beállítópanel hanglistáját; korábban a
  régi nyelv hangjait kínálta tovább.
- A 32 bites gazdafolyamat az első `DONE` jelre visszatért, a csomagoló viszont
  a hang megérkezése *előtt* küldi a `DONE`-t — a mondatok némaságra
  csonkulhattak.
- A hiányzó `_ipc.py` az egész beszélőt eltüntette az NVDA listájából ahelyett,
  hogy csak egy használaton kívüli ágat tiltott volna le.
- A beépített híd nélküli 64 bites NVDA-n a meghajtó be sem töltődött.
- A hanglisták már nem kínálnak olyan hangot, amelynek az adata hiányzik.
- A motor mostantól a nyelvi mappákat *tartalmazó* könyvtárat kapja meg.
