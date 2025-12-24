#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#pragma warning(disable : 4250)

#include <windows.h>
#include <stdint.h>

#include <atomic>
#include <condition_variable>
#include <cstring>
#include <deque>
#include <mutex>
#include <string>
#include <thread>
#include <vector>
#include <memory>
#include <cmath>
#include <chrono>
#include <climits>

#include "fvwrap.h"

// FlexVoice SDK headers
#include "Engine.h"
#include "Speaker.h"
#include "Bookmark.h"
#include "Attribute.h"
#include "FVLanguage.h"

using namespace MM_TTSAPI;

// ============================================================
// Helpers
// ============================================================

static int clampInt(int v, int lo, int hi) {
    if (v < lo) return lo;
    if (v > hi) return hi;
    return v;
}

static double ratePercentToSpeechRate(int p) {
    p = clampInt(p, 0, 100);
    if (p <= 50) return 0.7 + (p / 50.0) * 0.3;        // 0.7 .. 1.0
    return 1.0 + ((p - 50) / 50.0) * 0.7;              // 1.0 .. 1.7
}

static inline bool isAsciiAlpha(char c) {
    return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}
static inline bool isAsciiDigit(char c) {
    return (c >= '0' && c <= '9');
}
static inline bool isAsciiAlnum(char c) {
    return isAsciiAlpha(c) || isAsciiDigit(c);
}
static inline char toLowerAscii(char c) {
    if (c >= 'A' && c <= 'Z') return (char)(c - 'A' + 'a');
    return c;
}
static inline bool isSpaceLike(char c) {
    return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

static const char* digitWord1(char d) {
    switch (d) {
    case '0': return "zero";
    case '1': return "one";
    case '2': return "two";
    case '3': return "three";
    case '4': return "four";
    case '5': return "five";
    case '6': return "six";
    case '7': return "seven";
    case '8': return "eight";
    case '9': return "nine";
    default:  return "";
    }
}

// Consonants only (so we don't break article "a" or pronoun "I")
static const char* consonantLetterName(char lc) {
    switch (lc) {
    case 'b': return "bee";
    case 'c': return "see";
    case 'd': return "dee";
    case 'f': return "eff";
    case 'g': return "gee";
    case 'h': return "aitch";
    case 'j': return "jay";
    case 'k': return "kay";
    case 'l': return "ell";
    case 'm': return "em";
    case 'n': return "en";
    case 'p': return "pee";
    case 'q': return "cue";
    case 'r': return "arr";
    case 's': return "ess";
    case 't': return "tee";
    case 'v': return "vee";
    case 'w': return "double u";
    case 'x': return "ex";
    case 'y': return "why";
    case 'z': return "zee";
    default:  return nullptr;
    }
}

// Full letter names (for spelling acronyms like NVDA)
static const char* letterName(char lc) {
    switch (lc) {
    case 'a': return "ay";
    case 'b': return "bee";
    case 'c': return "see";
    case 'd': return "dee";
    case 'e': return "ee";
    case 'f': return "eff";
    case 'g': return "gee";
    case 'h': return "aitch";
    case 'i': return "eye";
    case 'j': return "jay";
    case 'k': return "kay";
    case 'l': return "ell";
    case 'm': return "em";
    case 'n': return "en";
    case 'o': return "oh";
    case 'p': return "pee";
    case 'q': return "cue";
    case 'r': return "arr";
    case 's': return "ess";
    case 't': return "tee";
    case 'u': return "you";
    case 'v': return "vee";
    case 'w': return "double u";
    case 'x': return "ex";
    case 'y': return "why";
    case 'z': return "zee";
    default:  return nullptr;
    }
}

static std::string spellLetters(const std::string& tok) {
    std::string out;
    for (char c : tok) {
        if (!isAsciiAlpha(c)) continue;
        const char* nm = letterName(toLowerAscii(c));
        if (!nm) continue;
        if (!out.empty()) out.push_back(' ');
        out += nm;
    }
    return out;
}

static bool parseULL_10(const std::string& digits, unsigned long long& outVal) {
    unsigned long long v = 0;
    if (digits.empty()) return false;
    for (char c : digits) {
        if (!isAsciiDigit(c)) return false;
        unsigned int d = (unsigned int)(c - '0');
        if (v > (ULLONG_MAX - d) / 10ULL) return false;
        v = v * 10ULL + (unsigned long long)d;
    }
    outVal = v;
    return true;
}

static std::string wordsBelow100(unsigned int n) {
    static const char* below20[] = {
        "zero","one","two","three","four","five","six","seven","eight","nine",
        "ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"
    };
    static const char* tens[] = {
        "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"
    };
    if (n < 20) return below20[n];
    unsigned int t = n / 10;
    unsigned int u = n % 10;
    std::string out = tens[t];
    if (u) {
        out.push_back(' ');
        out += below20[u];
    }
    return out;
}

static std::string wordsBelow1000(unsigned int n) {
    if (n < 100) return wordsBelow100(n);
    static const char* below20[] = {
        "zero","one","two","three","four","five","six","seven","eight","nine",
        "ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"
    };
    unsigned int h = n / 100;
    unsigned int r = n % 100;
    std::string out = below20[h];
    out += " hundred";
    if (r) {
        out.push_back(' ');
        out += wordsBelow100(r);
    }
    return out;
}

static std::string numberToWordsULL(unsigned long long n) {
    if (n == 0ULL) return "zero";

    struct Scale { unsigned long long v; const char* name; };
    static const Scale scales[] = {
        { 1000000000000ULL, "trillion" },
        { 1000000000ULL,    "billion"  },
        { 1000000ULL,       "million"  },
        { 1000ULL,          "thousand" },
    };

    std::string out;

    for (const auto& sc : scales) {
        if (n >= sc.v) {
            unsigned long long chunk = n / sc.v;
            n = n % sc.v;

            std::string chunkWords = numberToWordsULL(chunk);
            if (!out.empty()) out.push_back(' ');
            out += chunkWords;
            out.push_back(' ');
            out += sc.name;
        }
    }

    if (n > 0ULL) {
        std::string tail = wordsBelow1000((unsigned int)n);
        if (!out.empty()) out.push_back(' ');
        out += tail;
    }

    return out;
}

static std::string digitsToWordsByDigit(const std::string& digits) {
    std::string out;
    for (char c : digits) {
        const char* w = digitWord1(c);
        if (!w || !w[0]) continue;
        if (!out.empty()) out.push_back(' ');
        out += w;
    }
    return out;
}

static std::string convertDigitRunToWords(const std::string& digits) {
    if (digits.empty()) return std::string();
    if (digits.size() == 1) return digitWord1(digits[0]);

    // Leading zeros: preserve by speaking digits individually
    if (digits[0] == '0') {
        return digitsToWordsByDigit(digits);
    }

    // Too large: speak digit-by-digit to avoid overflow
    if (digits.size() > 18) {
        return digitsToWordsByDigit(digits);
    }

    unsigned long long val = 0;
    if (!parseULL_10(digits, val)) {
        return digitsToWordsByDigit(digits);
    }

    return numberToWordsULL(val);
}

// Expand digits (crash fix), single consonant tokens ("k" -> "kay"),
// and short ALL-CAPS tokens ("NVDA" -> "en vee dee ay") so acronyms don't disappear.
static std::string normalizeFragileTokens(const std::string& in) {
    std::string out;
    out.reserve(in.size() * 2);

    size_t i = 0;
    while (i < in.size()) {
        char c = in[i];

        // Apostrophes inside words (contractions / possessives):
        // Without this, "it's" becomes "it" + "ess" (because 's' becomes a single-letter token).
        if (c == '\'') {
            char prevCh = (i > 0) ? in[i - 1] : '\0';
            char n1 = (i + 1 < in.size()) ? in[i + 1] : '\0';
            char n2 = (i + 2 < in.size()) ? in[i + 2] : '\0';
            char n3 = (i + 3 < in.size()) ? in[i + 3] : '\0';

            // Only treat as a joiner/contractor when it's between letters.
            if (isAsciiAlpha(prevCh) && isAsciiAlpha(n1)) {
                char l1 = toLowerAscii(n1);
                char l2 = toLowerAscii(n2);

                const bool n2Alpha = isAsciiAlpha(n2);
                const bool n3Alpha = isAsciiAlpha(n3);

                // "'re" -> " are"
                if (l1 == 'r' && l2 == 'e' && n2Alpha && !n3Alpha) {
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += "are";
                    i += 3; // skip: ' r e
                    continue;
                }

                // "'ve" -> " have"
                if (l1 == 'v' && l2 == 'e' && n2Alpha && !n3Alpha) {
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += "have";
                    i += 3;
                    continue;
                }

                // "'ll" -> " will"
                if (l1 == 'l' && l2 == 'l' && n2Alpha && !n3Alpha) {
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += "will";
                    i += 3;
                    continue;
                }

                // "'s" -> join as a real letter (keeps "it's" -> "its", stops "ess")
                if (l1 == 's' && !n2Alpha) {
                    out.push_back('s');
                    i += 2; // skip: ' s
                    continue;
                }

                // "'t" -> join (keeps "can't" -> "cant", stops "tee")
                if (l1 == 't' && !n2Alpha) {
                    out.push_back('t');
                    i += 2;
                    continue;
                }

                // "'m" -> " am" (nice for "I'm")
                if (l1 == 'm' && !n2Alpha) {
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += "am";
                    i += 2;
                    continue;
                }

                // "'d" -> " would" (imperfect, but much better than spelling "d")
                if (l1 == 'd' && !n2Alpha) {
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += "would";
                    i += 2;
                    continue;
                }

                // Default: drop apostrophe between letters (prevents the engine from splitting weirdly)
                i++;
                continue;
            }

            // If it's not between letters, keep it as-is.
            out.push_back(c);
            i++;
            continue;
        }

        // Hyphen handling:
        // FlexVoice can sometimes truncate/stop when '-' appears inside tokens like "IBM-TTS".
        // Treat '-' as a word boundary so both sides always get spoken.
        //
        // Special case: leading "-123" => "minus one hundred ..." (useful, and still safe).
        if (c == '-') {
            char prevCh = (i > 0) ? in[i - 1] : '\0';
            char nextCh = (i + 1 < in.size()) ? in[i + 1] : '\0';

            // If it's clearly a minus sign before a digit run, speak "minus".
            // We only do this at a token boundary to avoid messing with hyphenated words.
            const bool boundaryBefore =
                (i == 0) || isSpaceLike(prevCh) || prevCh == '(' || prevCh == '[' || prevCh == '{' || prevCh == ',' || prevCh == ':';

            if (boundaryBefore && isAsciiDigit(nextCh)) {
                if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                out += "minus";
                out.push_back(' ');
                i++;
                continue;
            }

            // Default: treat hyphen as a separator (space). This fixes "IBM-TTS", "ETI-Eloquence", etc.
            if (out.empty() || !isSpaceLike(out.back())) out.push_back(' ');
            i++;
            continue;
        }

        // Digit run -> words
        if (isAsciiDigit(c)) {
            size_t j = i;
            while (j < in.size() && isAsciiDigit(in[j])) j++;
            std::string run = in.substr(i, j - i);
            std::string repl = convertDigitRunToWords(run);

            char nextCh = (j < in.size()) ? in[j] : '\0';

            if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
            out += repl;
            if (nextCh && (isAsciiAlpha(nextCh) || isAsciiDigit(nextCh))) out.push_back(' ');

            i = j;
            continue;
        }

        // Alpha token
        if (isAsciiAlpha(c)) {
            size_t j = i;
            while (j < in.size() && isAsciiAlpha(in[j])) j++;
            size_t len = j - i;

            // ALL-CAPS short token -> spell letters
            bool allCaps = true;
            for (size_t k = i; k < j; k++) {
                char ch = in[k];
                if (!(ch >= 'A' && ch <= 'Z')) { allCaps = false; break; }
            }
            if (allCaps && len >= 2 && len <= 6) {
                std::string spelled = spellLetters(in.substr(i, len));
                char nextCh = (j < in.size()) ? in[j] : '\0';
                if (!spelled.empty()) {
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += spelled;
                    if (nextCh && (isAsciiAlpha(nextCh) || isAsciiDigit(nextCh))) out.push_back(' ');
                    i = j;
                    continue;
                }
            }

            // Single-letter consonant tokens -> letter name
            if (len == 1) {
                char lc = toLowerAscii(in[i]);
                if (const char* nm = consonantLetterName(lc)) {
                    char nextCh = (j < in.size()) ? in[j] : '\0';
                    if (!out.empty() && !isSpaceLike(out.back())) out.push_back(' ');
                    out += nm;
                    if (nextCh && (isAsciiAlpha(nextCh) || isAsciiDigit(nextCh))) out.push_back(' ');
                    i = j;
                    continue;
                }
            }

            // Otherwise keep token
            out.append(in, i, len);
            i = j;
            continue;
        }

        // Other chars
        out.push_back(c);
        i++;
    }

    while (!out.empty() && out.back() == ' ') out.pop_back();
    return out;
}

// UTF-8 -> ISO-8859-1 safe-ish conversion + control sanitization + fragile-token normalization
static std::string utf8ToLatin1Safe(const char* sUtf8) {
    if (!sUtf8) return std::string();

    // Fast path: pure ASCII (common for key echo / UI).
    // IMPORTANT: still normalizeFragileTokens here, or digits will slip through and crash FlexVoice again.
    {
        bool allAscii = true;
        for (const unsigned char* p = (const unsigned char*)sUtf8; *p; ++p) {
            if (*p >= 0x80) { allAscii = false; break; }
        }

        if (allAscii) {
            std::string out(sUtf8);

            for (char& c : out) {
                unsigned char uc = (unsigned char)c;
                if (uc < 0x20 && c != '\r' && c != '\n' && c != '\t') c = ' ';
                else if (uc == 0x7F) c = ' ';
            }

            out = normalizeFragileTokens(out);
            return out;
        }
    }

    int wlen = MultiByteToWideChar(CP_UTF8, 0, sUtf8, -1, nullptr, 0);
    if (wlen <= 0) return std::string();

    std::wstring ws;
    ws.resize((size_t)wlen);
    MultiByteToWideChar(CP_UTF8, 0, sUtf8, -1, &ws[0], wlen);

    BOOL usedDefault = FALSE;
    int alen = WideCharToMultiByte(28591, 0, ws.c_str(), -1, nullptr, 0, "?", &usedDefault);
    if (alen <= 0) return std::string();

    std::string out;
    out.resize((size_t)alen);
    WideCharToMultiByte(28591, 0, ws.c_str(), -1, &out[0], alen, "?", &usedDefault);

    if (!out.empty() && out.back() == '\0') out.pop_back();

    for (char& c : out) {
        unsigned char uc = (unsigned char)c;
        if (uc < 0x20 && c != '\r' && c != '\n' && c != '\t') c = ' ';
        else if (uc == 0x7F || (uc >= 0x80 && uc <= 0x9F)) c = ' ';
        else if (uc == 0xA0) c = ' ';
    }

    out = normalizeFragileTokens(out);
    return out;
}

// ============================================================
// Stream Items
// ============================================================

struct StreamItem {
    int type;                  // FVWRAP_ITEM_*
    int value;                 // index / error
    uint32_t gen;              // per-utterance generation
    std::vector<uint8_t> data; // audio bytes if AUDIO
    size_t offset;

    StreamItem(int t = FVWRAP_ITEM_NONE, int v = 0, uint32_t g = 0)
        : type(t), value(v), gen(g), offset(0) {}
};

// ============================================================
// Bookmark for Indexing (BM_USER)
// ============================================================

class FvwrapIndexBookmark : public Bookmark {
public:
    FvwrapIndexBookmark(uint32_t g, int idx)
        : gen(g), index(idx), magic(0x46565749u) // 'FVWI'
    {
        type = BM_USER;
        id = magic;
    }

    virtual Bookmark* clone() const override {
        return new FvwrapIndexBookmark(*this);
    }

    uint32_t gen;
    int index;
    uint32_t magic;
};

#ifndef FVWRAP_DELETE_FOREIGN_BOOKMARKS
#define FVWRAP_DELETE_FOREIGN_BOOKMARKS 1
#endif

// ============================================================
// Output Site
// ============================================================

class MemoryWaveOutputSite : public IWaveOutputSite {
public:
    MemoryWaveOutputSite(int sampleRate, int bitsPerSample)
        : _fmt(sampleRate, bitsPerSample, WaveOutputFormat::WC_PCM_SIGNED)
    {
        const int maxBufferedMs = 500;
        size_t bytesPerSample = (bitsPerSample > 0) ? (size_t)((bitsPerSample + 7) / 8) : (size_t)2;
        size_t bytesPerSec = (sampleRate > 0) ? (size_t)sampleRate * bytesPerSample : (size_t)32000;
        _maxBufferedBytes = (bytesPerSec * (size_t)maxBufferedMs) / 1000;
        if (_maxBufferedBytes < 4096) _maxBufferedBytes = 4096;

        _maxQueueItems = 8192;
    }

    void gateOff() {
        _activeGen.store(0, std::memory_order_relaxed);
        _cvNotFull.notify_all();
    }

    void beginRequest(uint32_t gen, const std::vector<int>& expectedIdx) {
        {
            std::lock_guard<std::mutex> g(_mtx);
            _activeGen.store(gen, std::memory_order_relaxed);
            _paused = false;

            _q.clear();
            _expectedIdx.clear();
            _queuedAudioBytes = 0;

            for (int v : expectedIdx) _expectedIdx.push_back(v);
        }
        _cvNotFull.notify_all();
    }

    void clearQueue() {
        {
            std::lock_guard<std::mutex> g(_mtx);
            _q.clear();
            _expectedIdx.clear();
            _queuedAudioBytes = 0;
        }
        _cvNotFull.notify_all();
    }

    void pushIndex(uint32_t gen, int idx) {
        std::lock_guard<std::mutex> g(_mtx);
        if (_activeGen.load(std::memory_order_relaxed) != gen) return;

        if (!_expectedIdx.empty()) {
            if (_expectedIdx.front() == idx) {
                _expectedIdx.pop_front();
            } else {
                for (auto it = _expectedIdx.begin(); it != _expectedIdx.end(); ++it) {
                    if (*it == idx) { _expectedIdx.erase(it); break; }
                }
            }
        }
        _q.push_back(StreamItem(FVWRAP_ITEM_INDEX, idx, gen));
    }

    void pushError(uint32_t gen, int err) {
        std::lock_guard<std::mutex> g(_mtx);
        if (_activeGen.load(std::memory_order_relaxed) != gen) return;
        _q.push_back(StreamItem(FVWRAP_ITEM_ERROR, err, gen));
    }

    void finishRequest(uint32_t gen) {
        std::lock_guard<std::mutex> g(_mtx);
        if (_activeGen.load(std::memory_order_relaxed) != gen) return;

        while (!_expectedIdx.empty()) {
            int idx = _expectedIdx.front();
            _expectedIdx.pop_front();
            _q.push_back(StreamItem(FVWRAP_ITEM_INDEX, idx, gen));
        }

        _q.push_back(StreamItem(FVWRAP_ITEM_DONE, 0, gen));
    }

    const IOutputFormat& getOutputFormat() const override { return _fmt; }

    void pause() override {
        {
            std::lock_guard<std::mutex> g(_mtx);
            _paused = true;
        }
        _cvNotFull.notify_all();
    }

    void play() override {
        {
            std::lock_guard<std::mutex> g(_mtx);
            _paused = false;
        }
        _cvNotFull.notify_all();
    }

    void clear() override {
        gateOff();
        clearQueue();
    }

    void setBookmark(Bookmark* bookmark) override {
        if (!bookmark) return;

        if (auto* ib = dynamic_cast<FvwrapIndexBookmark*>(bookmark)) {
            const uint32_t cur = _activeGen.load(std::memory_order_relaxed);
            if (cur != 0 && ib->gen == cur && ib->id == ib->magic) {
                pushIndex(ib->gen, ib->index);
            }
            delete ib;
            return;
        }

#if FVWRAP_DELETE_FOREIGN_BOOKMARKS
        delete bookmark;
#else
        (void)bookmark;
#endif
    }

    void sendBookmark(Bookmark& /*bookmark*/) override {}

    void put(unsigned char* buffer, unsigned int size) override {
        if (!buffer || size == 0) return;

        uint32_t gen = _activeGen.load(std::memory_order_relaxed);
        if (gen == 0) return;

        std::vector<uint8_t> copied(size);
        std::memcpy(copied.data(), buffer, size);

        enqueueAudio(gen, std::move(copied));
    }

    int read(int* outType, int* outValue, uint8_t* outAudio, int outCap) {
        if (outType) *outType = FVWRAP_ITEM_NONE;
        if (outValue) *outValue = 0;
        if (!outAudio || outCap < 0) outCap = 0;

        bool notify = false;

        std::lock_guard<std::mutex> g(_mtx);

        uint32_t gen = _activeGen.load(std::memory_order_relaxed);
        if (gen == 0) {
            if (!_q.empty() || _queuedAudioBytes != 0) notify = true;
            _q.clear();
            _expectedIdx.clear();
            _queuedAudioBytes = 0;
            if (notify) _cvNotFull.notify_all();
            return 0;
        }

        while (!_q.empty() && _q.front().gen != gen) {
            StreamItem& it = _q.front();
            if (it.type == FVWRAP_ITEM_AUDIO) {
                size_t remaining = (it.data.size() > it.offset) ? (it.data.size() - it.offset) : 0;
                if (_queuedAudioBytes >= remaining) _queuedAudioBytes -= remaining;
                else _queuedAudioBytes = 0;
                notify = true;
            }
            _q.pop_front();
        }

        if (_q.empty()) {
            if (notify) _cvNotFull.notify_all();
            return 0;
        }

        StreamItem& front = _q.front();
        if (outType) *outType = front.type;
        if (outValue) *outValue = front.value;

        if (front.type == FVWRAP_ITEM_AUDIO) {
            size_t remainingSz = (front.data.size() > front.offset) ? (front.data.size() - front.offset) : 0;
            int remaining = (remainingSz > (size_t)INT32_MAX) ? INT32_MAX : (int)remainingSz;
            int n = remaining;
            if (n > outCap) n = outCap;

            if (n > 0) {
                std::memcpy(outAudio, front.data.data() + front.offset, (size_t)n);
                front.offset += (size_t)n;

                if (_queuedAudioBytes >= (size_t)n) _queuedAudioBytes -= (size_t)n;
                else _queuedAudioBytes = 0;

                notify = true;
            }

            if (front.offset >= front.data.size()) {
                _q.pop_front();
            }

            if (notify) _cvNotFull.notify_all();
            return n;
        }

        _q.pop_front();
        if (notify) _cvNotFull.notify_all();
        return 0;
    }

    // Stubs
    void registerNotify(INotify*, const BookmarkTypeList&) override {}
    void unregisterNotify(INotify*) override {}
    void addBookmarkTypes(INotify*, const BookmarkTypeList&) override {}
    void removeBookmarkTypes(INotify*, const BookmarkTypeList&) override {}
    bool set(const char*, bool, int = 0) override { return false; }
    bool get(const char*, bool&, int = 0) const override { return false; }
    bool set(const char*, char, int = 0) override { return false; }
    bool get(const char*, char&, int = 0) const override { return false; }
    bool set(const char*, int, int = 0) override { return false; }
    bool get(const char*, int&, int = 0) const override { return false; }
    bool set(const char*, double, int = 0) override { return false; }
    bool get(const char*, double&, int = 0) const override { return false; }
    bool set(const char*, const char*, int = 0) override { return false; }
    bool get(const char*, char*, int, int*, int = 0) const override { return false; }
    bool setArraySize(const char*, int) override { return false; }
    bool getArraySize(const char*, int&) const override { return false; }
    const IAttribute* getAttribute(const char*) const override { return nullptr; }
    IAttribute* getAttribute(const char*) override { return nullptr; }

private:
    void enqueueAudio(uint32_t gen, std::vector<uint8_t>&& bytes) {
        if (bytes.empty()) return;

        std::unique_lock<std::mutex> lk(_mtx);

        if (_paused) return;
        if (_activeGen.load(std::memory_order_relaxed) != gen) return;

        size_t limit = _maxBufferedBytes;
        if (bytes.size() > limit) limit = bytes.size();

        while (
            (_queuedAudioBytes + bytes.size() > limit || _q.size() >= _maxQueueItems) &&
            !_paused &&
            _activeGen.load(std::memory_order_relaxed) == gen
        ) {
            _cvNotFull.wait_for(lk, std::chrono::milliseconds(20));
        }

        if (_paused) return;
        if (_activeGen.load(std::memory_order_relaxed) != gen) return;

        StreamItem it(FVWRAP_ITEM_AUDIO, 0, gen);
        it.data = std::move(bytes);
        it.offset = 0;

        _queuedAudioBytes += it.data.size();
        _q.push_back(std::move(it));
    }

private:
    WaveOutputFormat _fmt;
    std::atomic<uint32_t> _activeGen{ 0 };

    std::mutex _mtx;
    std::condition_variable _cvNotFull;

    std::deque<StreamItem> _q;
    std::deque<int> _expectedIdx;

    bool _paused = false;

    size_t _queuedAudioBytes = 0;
    size_t _maxBufferedBytes = 0;
    size_t _maxQueueItems = 0;
};

// ============================================================
// Command/State
// ============================================================

struct Part {
    enum Kind { TEXT, INDEX } kind;
    std::string text;
    int index = 0;

    static Part makeText(std::string s) { Part p; p.kind = TEXT; p.text = std::move(s); return p; }
    static Part makeIndex(int i) { Part p; p.kind = INDEX; p.index = i; return p; }
};

struct Cmd {
    enum Type { CMD_SPEAK, CMD_QUIT } type;
    uint32_t cancelSnapshot;
    std::vector<Part> parts;
};

struct WrapState {
    EngineFactory* factory = nullptr;
    Engine* engine = nullptr;
    MemoryWaveOutputSite* output = nullptr;

    // Base (loaded) speaker; we re-apply pitch by starting from this to avoid compounding.
    Speaker speaker;
    Language lang = (Language)0;

    std::atomic<uint32_t> cancelToken{ 1 };
    std::atomic<uint32_t> genCounter{ 1 };

    // Settings (store-only)
    std::atomic<int> desiredRate{ 50 };
    std::atomic<int> desiredVol{ 100 };
    std::atomic<int> desiredPitch{ 50 };

    // Applied caches (worker thread only)
    int appliedRate = -1;
    int appliedVol = -1;
    int appliedPitch = -1;

    std::mutex buildMtx;
    std::string buildText;
    std::vector<Part> buildParts;

    std::mutex engineMtx;

    std::mutex qMtx;
    std::condition_variable qCv;
    std::deque<Cmd> q;
    bool quitting = false;
    std::thread worker;
};

static WrapState* asState(FVWRAP_HANDLE h) { return (WrapState*)h; }

static void collectExpectedIndexes(const std::vector<Part>& parts, std::vector<int>& outIdx) {
    outIdx.clear();
    for (const auto& p : parts) {
        if (p.kind == Part::INDEX) outIdx.push_back(p.index);
    }
}

static bool trySetSpeakerDouble(Speaker& sp, const char* name, double v) {
    try { return sp.set(name, v); } catch (...) { return false; }
}
static bool tryGetSpeakerDouble(const Speaker& sp, const char* name, double& v) {
    try { return sp.get(name, v); } catch (...) { return false; }
}
static bool trySetSpeakerInt(Speaker& sp, const char* name, int v) {
    try { return sp.set(name, v); } catch (...) { return false; }
}
static bool tryGetSpeakerInt(const Speaker& sp, const char* name, int& v) {
    try { return sp.get(name, v); } catch (...) { return false; }
}

static bool applyPitchToSpeakerLocked(WrapState* st, int pitchPercent) {
    if (!st || !st->engine) return false;

    pitchPercent = clampInt(pitchPercent, 0, 100);

    // Start from base speaker each time (prevents compounding).
    Speaker sp = st->speaker;

    // Speaker.h gives us the correct names:
    // "defaultPitch" int, "pitchMin" int, "pitchMax" int
    int defPitch = 0;
    int pMin = 0;
    int pMax = 0;

    bool hasDef = tryGetSpeakerInt(sp, "defaultPitch", defPitch);
    bool hasMin = tryGetSpeakerInt(sp, "pitchMin", pMin);
    bool hasMax = tryGetSpeakerInt(sp, "pitchMax", pMax);

    // Some builds might store it as double internally; handle that too.
    if (!hasDef) {
        double d = 0.0;
        if (tryGetSpeakerDouble(sp, "defaultPitch", d)) {
            defPitch = (int)std::lround(d);
            hasDef = true;
        }
    }

    if (!hasDef) {
        // We can't affect pitch if speaker doesn't expose defaultPitch.
        // Still set the (unmodified) base speaker to keep behavior consistent.
        try { st->engine->setSpeaker(sp, true); } catch (...) {}
        return false;
    }

    int baseDef = defPitch;
    int newDef = baseDef;

    // If pitchMin/pitchMax exist, map NVDA percent around the *defaultPitch*:
    // - 50 => base defaultPitch
    // - 0  => pitchMin
    // - 100=> pitchMax
    if (hasMin && hasMax && pMin < pMax) {
        // Clamp base defaultPitch into [min, max] just in case.
        if (baseDef < pMin) baseDef = pMin;
        if (baseDef > pMax) baseDef = pMax;

        if (pitchPercent == 50) {
            newDef = baseDef;
        } else if (pitchPercent < 50) {
            // Move down toward pitchMin
            const double t = (50.0 - (double)pitchPercent) / 50.0; // 0..1
            newDef = (int)std::lround((double)baseDef - t * (double)(baseDef - pMin));
        } else {
            // Move up toward pitchMax
            const double t = ((double)pitchPercent - 50.0) / 50.0; // 0..1
            newDef = (int)std::lround((double)baseDef + t * (double)(pMax - baseDef));
        }

        if (newDef < pMin) newDef = pMin;
        if (newDef > pMax) newDef = pMax;
    } else {
        // Fallback if min/max are missing: semitone-ish scaling of defaultPitch.
        // This is less ideal, but better than nothing.
        const double semis = ((double)pitchPercent - 50.0) / 50.0 * 8.0; // +/- 8 semitones
        const double ratio = std::pow(2.0, semis / 12.0);
        newDef = (int)std::lround((double)baseDef * ratio);

        // If min/max exist but were malformed earlier, clamp anyway.
        if (hasMin && hasMax && pMin < pMax) {
            if (newDef < pMin) newDef = pMin;
            if (newDef > pMax) newDef = pMax;
        }
    }

    bool ok = false;
    ok |= trySetSpeakerInt(sp, "defaultPitch", newDef);
    if (!ok) {
        // Some builds might accept double instead of int.
        ok |= trySetSpeakerDouble(sp, "defaultPitch", (double)newDef);
    }

    // Apply speaker to engine.
    try { st->engine->setSpeaker(sp, true); } catch (...) {}

    return ok;
}

// Apply engine parameters at a safe boundary (worker thread, under engineMtx)
static void applyEngineParamsLocked(WrapState* st) {
    if (!st || !st->engine) return;

    // Rate -> "speechRate"
    const int rp = clampInt(st->desiredRate.load(std::memory_order_relaxed), 0, 100);
    if (st->appliedRate != rp) {
        try {
            double sr = ratePercentToSpeechRate(rp);
            if (sr < 0.5) sr = 0.5;
            if (sr > 2.2) sr = 2.2;
            st->engine->attribute().set("speechRate", sr);
        } catch (...) {}
        st->appliedRate = rp;
    }

    // Volume -> "volume"
    const int vp = clampInt(st->desiredVol.load(std::memory_order_relaxed), 0, 100);
    if (st->appliedVol != vp) {
        try {
            double vol = (double)vp / 100.0;
            if (vol < 0.0) vol = 0.0;
            if (vol > 1.0) vol = 1.0;
            st->engine->attribute().set("volume", vol);
        } catch (...) {}
        st->appliedVol = vp;
    }

    // Pitch via speaker
    const int pp = clampInt(st->desiredPitch.load(std::memory_order_relaxed), 0, 100);
    if (st->appliedPitch != pp) {
        (void)applyPitchToSpeakerLocked(st, pp);
        st->appliedPitch = pp;
    }
}

static void workerLoop(WrapState* st) {
    if (!st || !st->engine || !st->output) return;

    // Make the worker a bit more responsive under load.
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_ABOVE_NORMAL);

    while (true) {
        Cmd cmd;
        {
            std::unique_lock<std::mutex> lk(st->qMtx);
            st->qCv.wait(lk, [&]() { return st->quitting || !st->q.empty(); });
            if (st->quitting) return;

            cmd = std::move(st->q.front());
            st->q.pop_front();
        }

        if (cmd.type == Cmd::CMD_QUIT) return;

        if (st->cancelToken.load(std::memory_order_relaxed) != cmd.cancelSnapshot) {
            continue;
        }

        uint32_t gen = st->genCounter.fetch_add(1, std::memory_order_relaxed);

        st->output->gateOff();
        st->output->clearQueue();

        try {
            std::lock_guard<std::mutex> g(st->engineMtx);
            st->engine->stop();
        } catch (...) {}
        try {
            st->engine->wait();
        } catch (...) {}

        std::vector<int> expectedAll;
        collectExpectedIndexes(cmd.parts, expectedAll);

        st->output->beginRequest(gen, expectedAll);

        bool hasText = false;
        for (const auto& p : cmd.parts) {
            if (p.kind == Part::TEXT && !p.text.empty()) { hasText = true; break; }
        }
        if (!hasText) {
            st->output->finishRequest(gen);
            continue;
        }

        bool canceled = false;

        // Pre-normalize outside engineMtx so the lock is held for less time.
        struct Prepared {
            bool isIndex = false;
            int index = 0;
            std::string text;
        };

        std::vector<Prepared> prepared;
        prepared.reserve(cmd.parts.size());

        for (const auto& p : cmd.parts) {
            if (st->cancelToken.load(std::memory_order_relaxed) != cmd.cancelSnapshot) {
                canceled = true;
                break;
            }

            if (p.kind == Part::TEXT) {
                if (!p.text.empty()) {
                    Prepared pr;
                    pr.isIndex = false;
                    pr.text = normalizeFragileTokens(p.text); // KEEP ALL NORMALIZER RULES
                    if (!pr.text.empty()) {
                        prepared.push_back(std::move(pr));
                    }
                }
            } else {
                Prepared pr;
                pr.isIndex = true;
                pr.index = p.index;
                prepared.push_back(std::move(pr));
            }
        }

        try {
            {
                std::lock_guard<std::mutex> g(st->engineMtx);

                applyEngineParamsLocked(st);

                if (!canceled) {
                    for (const auto& pr : prepared) {
                        if (st->cancelToken.load(std::memory_order_relaxed) != cmd.cancelSnapshot) {
                            canceled = true;
                            break;
                        }

                        if (!pr.isIndex) {
                            st->engine->addFragment(pr.text.c_str());
                        } else {
                            st->engine->addBookmark(new FvwrapIndexBookmark(gen, pr.index));
                        }
                    }
                }

                if (!canceled) {
                    st->engine->speakRequest(1);
                }
            }

            if (!canceled) {
                st->engine->wait();
            }
        } catch (...) {
            st->output->pushError(gen, 901);
        }

        if (st->cancelToken.load(std::memory_order_relaxed) != cmd.cancelSnapshot) {
            canceled = true;
        }

        if (canceled) {
            st->output->gateOff();
            st->output->clearQueue();
            continue;
        }

        st->output->finishRequest(gen);
    }
}

// ============================================================
// C API
// ============================================================

extern "C" {

FVWRAP_API FVWRAP_HANDLE __cdecl fvwrap_create(
    const char* dataPath,
    const char* speakerTavPath,
    int language,
    int sampleRate,
    int bitsPerSample
) {
    auto st = new WrapState();

    try {
        st->lang = (Language)language;

        st->output = new MemoryWaveOutputSite(sampleRate, bitsPerSample);

        st->factory = new EngineFactory(dataPath);
        st->factory->loadLanguage(st->lang);

        if (speakerTavPath && speakerTavPath[0]) {
            st->speaker.load(speakerTavPath);
        }

        std::auto_ptr<Engine> eng = st->factory->createEngine(st->output, st->speaker, st->lang);
        st->engine = eng.release();

        st->worker = std::thread(workerLoop, st);
        return (FVWRAP_HANDLE)st;
    } catch (...) {
        try {
            st->quitting = true;
            st->qCv.notify_all();
            if (st->worker.joinable()) st->worker.join();
        } catch (...) {}

        if (st->engine) { delete st->engine; st->engine = nullptr; }
        if (st->factory) { delete st->factory; st->factory = nullptr; }
        if (st->output) { delete st->output; st->output = nullptr; }
        delete st;
        return nullptr;
    }
}

FVWRAP_API void __cdecl fvwrap_destroy(FVWRAP_HANDLE h) {
    auto st = asState(h);
    if (!st) return;

    st->cancelToken.fetch_add(1, std::memory_order_relaxed);

    if (st->output) {
        st->output->gateOff();
        st->output->clearQueue();
    }

    try {
        std::lock_guard<std::mutex> g(st->engineMtx);
        if (st->engine) st->engine->stop();
    } catch (...) {}

    {
        std::lock_guard<std::mutex> lk(st->qMtx);
        st->quitting = true;
        st->q.clear();
    }
    st->qCv.notify_all();
    if (st->worker.joinable()) st->worker.join();

    try {
        if (st->engine) st->engine->wait();
    } catch (...) {}

    if (st->engine) { delete st->engine; st->engine = nullptr; }
    if (st->factory) { delete st->factory; st->factory = nullptr; }
    if (st->output) { delete st->output; st->output = nullptr; }

    delete st;
}

FVWRAP_API void __cdecl fvwrap_stop(FVWRAP_HANDLE h) {
    auto st = asState(h);
    if (!st) return;

    st->cancelToken.fetch_add(1, std::memory_order_relaxed);

    if (st->output) st->output->gateOff();

    {
        std::lock_guard<std::mutex> lk(st->qMtx);
        st->q.clear();
    }
    st->qCv.notify_all();

    {
        std::lock_guard<std::mutex> bk(st->buildMtx);
        st->buildText.clear();
        st->buildParts.clear();
    }

    try {
        std::lock_guard<std::mutex> g(st->engineMtx);
        if (st->engine) st->engine->stop();
    } catch (...) {}

    if (st->output) st->output->clearQueue();
}

FVWRAP_API int __cdecl fvwrap_setRatePercent(FVWRAP_HANDLE h, int ratePercent) {
    auto st = asState(h);
    if (!st) return 1;
    st->desiredRate.store(clampInt(ratePercent, 0, 100), std::memory_order_relaxed);
    return 0;
}

FVWRAP_API int __cdecl fvwrap_setVolumePercent(FVWRAP_HANDLE h, int volPercent) {
    auto st = asState(h);
    if (!st) return 1;
    st->desiredVol.store(clampInt(volPercent, 0, 100), std::memory_order_relaxed);
    return 0;
}

FVWRAP_API int __cdecl fvwrap_setPitchPercent(FVWRAP_HANDLE h, int pitchPercent) {
    auto st = asState(h);
    if (!st) return 1;
    st->desiredPitch.store(clampInt(pitchPercent, 0, 100), std::memory_order_relaxed);
    return 0;
}

FVWRAP_API void __cdecl fvwrap_begin(FVWRAP_HANDLE h) {
    auto st = asState(h);
    if (!st) return;
    std::lock_guard<std::mutex> g(st->buildMtx);
    st->buildText.clear();
    st->buildParts.clear();
}

FVWRAP_API void __cdecl fvwrap_addTextUtf8(FVWRAP_HANDLE h, const char* textUtf8) {
    auto st = asState(h);
    if (!st || !textUtf8) return;

    // Converts + sanitizes + expands digits + expands single consonants + spells short ALL-CAPS tokens.
    std::string s = utf8ToLatin1Safe(textUtf8);
    if (s.empty()) return;

    std::lock_guard<std::mutex> g(st->buildMtx);
    st->buildText.append(s);
}

FVWRAP_API void __cdecl fvwrap_addIndex(FVWRAP_HANDLE h, int index) {
    auto st = asState(h);
    if (!st) return;

    std::lock_guard<std::mutex> g(st->buildMtx);

    if (!st->buildText.empty()) {
        st->buildParts.push_back(Part::makeText(std::move(st->buildText)));
        st->buildText.clear();
    }
    st->buildParts.push_back(Part::makeIndex(index));
}

FVWRAP_API int __cdecl fvwrap_commit(FVWRAP_HANDLE h, int /*repeatCount*/) {
    auto st = asState(h);
    if (!st || !st->engine || !st->output) return 1;

    std::vector<Part> parts;
    {
        std::lock_guard<std::mutex> g(st->buildMtx);

        if (!st->buildText.empty()) {
            st->buildParts.push_back(Part::makeText(std::move(st->buildText)));
            st->buildText.clear();
        }

        parts = std::move(st->buildParts);
        st->buildParts.clear();
    }

    uint32_t snap = st->cancelToken.load(std::memory_order_relaxed);

    {
        std::lock_guard<std::mutex> lk(st->qMtx);
        Cmd cmd;
        cmd.type = Cmd::CMD_SPEAK;
        cmd.cancelSnapshot = snap;
        cmd.parts = std::move(parts);
        st->q.push_back(std::move(cmd));
    }
    st->qCv.notify_one();

    return 0;
}

FVWRAP_API int __cdecl fvwrap_read(
    FVWRAP_HANDLE h,
    int* outType,
    int* outValue,
    uint8_t* outAudio,
    int outCap
) {
    auto st = asState(h);
    if (!st || !st->output) {
        if (outType) *outType = FVWRAP_ITEM_NONE;
        if (outValue) *outValue = 0;
        return 0;
    }
    return st->output->read(outType, outValue, outAudio, outCap);
}

} // extern "C"
