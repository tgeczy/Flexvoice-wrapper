# FlexVoice for NVDA 2.5.1

A small fix. No changes to speech, voices or settings — 2.5 users can upgrade
without anything sounding different.

## Fixed

Changing language refreshed **every** driver settings panel NVDA had open, not
just FlexVoice's own. NVDA's braille display settings are built on the same
mechanism, so a language change also made the braille display driver re-read
its settings, and NVDA would try to build a control for a `language` setting
that braille drivers do not have.

Nothing visible went wrong, which is why it survived since 2.0 — but a speech
synthesizer has no business touching the braille subsystem, so it now refreshes
only its own panel.

Found thanks to a report from **Borris (@BorrisInABox@fwoof.space)**, who was
chasing something else entirely.

---

# FlexVoice NVDA-hoz 2.5.1

Apró javítás. A beszéd, a hangok és a beállítások változatlanok — a 2.5
felhasználói úgy frissíthetnek, hogy semmi nem szól másképp.

## Javítva

A nyelvváltás az NVDA **összes** nyitott meghajtó-beállítási paneljét
frissítette, nem csak a FlexVoice sajátját. Az NVDA braille-kijelző
beállításai ugyanarra a mechanizmusra épülnek, így a nyelvváltás a
braille-meghajtót is arra késztette, hogy újraolvassa a beállításait, az NVDA
pedig megpróbált vezérlőt létrehozni egy olyan `language` beállításhoz, amilyen
a braille-meghajtóknak nincs is.

Láthatóan semmi nem romlott el — ezért is maradt észrevétlen a 2.0 óta —, de
egy beszédszintetizátornak semmi keresnivalója a braille-alrendszerben, így
mostantól csak a saját paneljét frissíti.

Köszönet **Borris**nak (@BorrisInABox@fwoof.space) a bejelentésért, aki
egészen mást keresett.
