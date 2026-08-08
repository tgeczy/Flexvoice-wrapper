# FlexVoice for NVDA 2.5 — the filter Mindmaker never shipped

One new control: **Clarity**, in Voice settings and the settings ring.

It lifts the high frequencies, which makes the voices noticeably more natural
and easier to follow. **50 is the voice exactly as it was authored**, so nothing
changes until you move it. Below 50 is warmer, above 50 is brighter and crisper.
It works on every voice, English and Hungarian alike.

## Where this came from

This was **Király József's** idea, not ours. He worked on FlexVoice, and after
trying the 2.0 add-on he wrote back:

> *"egy egyszerű trükkel még emberibbé és természetesebbé lehetne tenni az Ő és
> a többi beszélő hangját. Egy-két egyszerű hang-szűrőn kellene átereszteni a
> hangot, kiemelve a magas hangokat… Annak idején terveztük ezt beépíteni, de a
> mostani verzióba már nem került be."*

They designed the filter at Mindmaker but never gave it a control. It turned out
the mechanism was in the voice data all along — every `.tav` carries an
equalizer curve, the engine honours it, and the SDK exposes it at runtime. So
Clarity is not a new effect bolted on top; it is the engine's own filter, finally
given a knob. Twenty-odd years late.

## How it behaves

Nothing below 1 kHz moves, so the body of the voice is untouched; the tilt ramps
in above that and reaches full strength by 6 kHz, where the consonant detail
that carries intelligibility lives. It is a tilt rather than a boost — highs come
up and lows come down together — so the loudness stays roughly constant and it
never clips at any setting.

Measured high-frequency content across the range:

| clarity | 0 | 25 | 50 | 75 | 100 |
|---|---|---|---|---|---|
| Zita (Hungarian) | 0.366 | 0.420 | **0.492** | 0.677 | 0.917 |
| Tom (English) | 0.338 | 0.415 | **0.509** | 0.644 | 0.813 |

Everything else is unchanged from 2.0.

---

# FlexVoice NVDA-hoz 2.5 — a szűrő, ami annak idején kimaradt

Egyetlen új vezérlő: **Tisztaság** (Clarity), a hangbeállításokban és a
beállításgyűrűben.

A magas hangokat emeli ki, amitől a hangok érezhetően természetesebbek és
könnyebben követhetők lesznek. **Az 50 pontosan az a hang, ahogyan eredetileg
elkészült**, tehát semmi nem változik, amíg el nem mozdítja. Az 50 alatt
melegebb, fölötte világosabb és tisztább. Minden hangra működik, angolra és
magyarra egyaránt.

## Honnan jött

Az ötlet **Király Józsefé**, nem a miénk. Ő dolgozott a FlexVoice-on, és miután
kipróbálta a 2.0-t, ezt írta:

> *„egy egyszerű trükkel még emberibbé és természetesebbé lehetne tenni az Ő és
> a többi beszélő hangját. Egy-két egyszerű hang-szűrőn kellene átereszteni a
> hangot, kiemelve a magas hangokat… Annak idején terveztük ezt beépíteni, de a
> mostani verzióba már nem került be."*

A szűrőt a Mindmakernél megtervezték, de sosem kapott vezérlőt. Kiderült, hogy a
mechanizmus végig ott volt a hangadatokban: minden `.tav` tartalmaz egy
hangszínszabályzó-görbét, a motor figyelembe is veszi, és az SDK futásidőben
hozzáférhetővé teszi. A Tisztaság tehát nem egy ráaggatott új effekt, hanem a
motor saját szűrője, amely végre gombot kapott. Bő húsz év késéssel.

## Hogyan viselkedik

1 kHz alatt semmi nem mozdul, így a hang teste érintetlen marad; a döntés efölött
kezdődik, és 6 kHz-re éri el a teljes erejét — ott, ahol az érthetőséget vivő
mássalhangzó-részletek vannak. Ez döntés, nem kiemelés: a magasak fel, a mélyek
le, együtt, így a hangerő nagyjából állandó marad, és egyetlen állásban sem
torzul.

Minden más változatlan a 2.0-hoz képest.
