#!/usr/bin/env python3
"""PPTX-Narrator: automated narration of PowerPoint presenter notes.

Pipeline: note extraction -> technical-term scanning -> dictionary / SI-unit
normalization -> (optional) machine translation -> voice-cloned TTS
(GPT-SoVITS or Qwen3-TTS) -> ASR round-trip verification -> PPTX repackaging.
"""
import os
import re
import unicodedata
from collections import Counter
import difflib
import hashlib
import sys
import time
import shutil
import logging
import argparse
import tempfile
import zipfile
import csv
import json
import requests
from pptx import Presentation
from deep_translator import GoogleTranslator
from pydub import AudioSegment
import nltk
from pydub.effects import compress_dynamic_range, normalize

__version__ = "1.0.0"


def _ensure_nltk_data():
    """Download the NLTK corpora used by --scan on first use (not at import time)."""
    for corpus in ("stopwords", "words"):
        try:
            nltk.data.find(f"corpora/{corpus}")
        except LookupError:
            nltk.download(corpus, quiet=True)
            try:
                nltk.data.find(f"corpora/{corpus}")
            except LookupError:
                logger.error(
                    f"NLTK corpus '{corpus}' is not available and could not be downloaded. "
                    "Install it manually: python -m nltk.downloader stopwords words"
                )
                sys.exit(1)

MODELS_CONFIG = {
    "v2ProPlus": {
        "gpt": "./GPT_SoVITS/pretrained_models/s1v3.ckpt",
        "sovits": "./GPT_SoVITS/pretrained_models/v2Pro/s2Gv2ProPlus.pth"
    },
    "v4": {
        "gpt": "./GPT_SoVITS/pretrained_models/s1v3.ckpt",
        "sovits": "./GPT_SoVITS/pretrained_models/v4/s2Gv4.pth"
    },
    "v1_clear": {
        "gpt": "./GPT_SoVITS/pretrained_models/s1bert25hz-2kh-longer-epoch=68e-step=50232.ckpt",
        "sovits": "./GPT_SoVITS/pretrained_models/s2G488k.pth"
    }
}

# ==========================================
# Configuration & Constants
# ==========================================
# Workspace file names carry the language of the text: slide_N_<lang>.txt, with <lang>
# the Google Translate code (slide_3_ja.txt, slide_3_en.txt, slide_3_zh-CN.txt). Files
# written by pre-release versions (slide_N.txt for Japanese, slide_N_eng.txt for English)
# are still recognized when a workspace is read.
CJK_LANGS = {"ja", "zh", "yue", "ko"}

# Languages accepted by the TTS engines (ISO 639-1 base code -> value passed to the engine)
QWEN3_LANGUAGES = {
    "zh": "Chinese", "en": "English", "ja": "Japanese", "ko": "Korean", "de": "German",
    "fr": "French", "ru": "Russian", "pt": "Portuguese", "es": "Spanish", "it": "Italian",
}
GPT_SOVITS_LANGUAGES = {"zh": "zh", "en": "en", "ja": "ja", "ko": "ko", "yue": "yue"}

_TEXT_FILE_RE = re.compile(r"^slide_(\d+)(?:_([A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,4})?))?\.txt$")


def normalize_lang(code):
    """Normalize a language code to the Google Translate spelling.

    'JA' -> 'ja', 'zh_cn' -> 'zh-CN', 'zh' -> 'zh-CN', 'eng' -> 'en', 'he' -> 'iw'.
    """
    if code is None:
        return None
    code = code.strip().replace("_", "-")
    if not code:
        return None
    if code.lower() == "auto":
        return "auto"
    base, _, region = code.partition("-")
    base = {"eng": "en", "jpn": "ja", "he": "iw", "jv": "jw"}.get(base.lower(), base.lower())
    if not region:
        return "zh-CN" if base == "zh" else base
    return f"{base}-{region.upper() if len(region) == 2 else region.title()}"


def base_lang(code):
    return code.split("-")[0].lower()


def lang_suffix(lang):
    return f"_{lang}"


def text_filename(slide_num, lang):
    return f"slide_{slide_num}{lang_suffix(lang)}.txt"


def audio_filename(slide_num, lang, model_label):
    return f"slide_{slide_num}{lang_suffix(lang)}.{model_label}.m4a"


def spoken_filename(slide_num, lang, model_label):
    return f"slide_{slide_num}{lang_suffix(lang)}.{model_label}.spoken.txt"


def qwen3_language(lang):
    return QWEN3_LANGUAGES.get(base_lang(lang))


def gpt_sovits_language(lang):
    return GPT_SOVITS_LANGUAGES.get(base_lang(lang))


# Language identification of the notes (--source-lang auto)
_LANGID_TO_GOOGLE = {"he": "iw", "jv": "jw", "yue": "zh-TW", "wuu": "zh-CN"}
_language_identifier = None


def _get_language_identifier():
    """py3langid classifier restricted to languages that Google Translate accepts."""
    global _language_identifier
    if _language_identifier is None:
        from py3langid import langid
        from deep_translator.constants import GOOGLE_LANGUAGES_TO_CODES
        identifier = langid.LanguageIdentifier.from_modelpath(
            langid.MODEL_DIR / langid.MODEL_FILE, norm_probs=True)
        google = set(GOOGLE_LANGUAGES_TO_CODES.values())
        identifier.set_languages([c for c in set(identifier.nb_classes)
                                  if _LANGID_TO_GOOGLE.get(c, c) in google or c == "zh"])
        _language_identifier = identifier
    return _language_identifier


def _chinese_variant(text):
    """zh-CN if all Han characters exist in GB2312 (simplified), otherwise zh-TW."""
    han = "".join(re.findall(r'[一-鿿]', text))
    try:
        han.encode("gb2312")
        return "zh-CN"
    except UnicodeEncodeError:
        return "zh-TW"


def detect_language(text):
    """Identify the language of one note. Returns (language code, confident)."""
    letters = re.sub(r'[\W\d_]', '', text)
    if not letters:
        return "en", False
    kana = len(re.findall(r'[぀-ヿ]', letters))
    hangul = len(re.findall(r'[가-힯ᄀ-ᇿ]', letters))
    han = len(re.findall(r'[一-鿿]', letters))
    if kana and (kana + han) >= 0.3 * len(letters):
        return "ja", True
    if hangul and (hangul + han) >= 0.3 * len(letters):
        return "ko", True
    if han == len(letters):
        # Kanji/hanzi only: Chinese or a kanji-only Japanese note
        return _chinese_variant(text), False
    lang, prob = _get_language_identifier().classify(text)
    lang = _LANGID_TO_GOOGLE.get(lang, lang)
    if lang == "zh":
        lang = _chinese_variant(text)
    return lang, (len(letters) >= 20 and prob >= 0.9)


def detect_note_languages(texts, known=None):
    """Identify the language of every note ({slide: text} -> {slide: code}).

    Notes that are too short or ambiguous to identify reliably ("Thank you.",
    kanji-only titles, formulas) are assigned the most frequent language among the
    confidently identified notes of the deck.
    """
    first = {n: detect_language(t) for n, t in texts.items()}
    votes = Counter(lang for lang, confident in first.values() if confident)
    votes.update((known or {}).values())
    majority = votes.most_common(1)[0][0] if votes else None
    result = {}
    for n, (lang, confident) in first.items():
        if not confident and majority and lang != majority:
            logger.info(f"Slide #{n}: short/ambiguous note ('{lang}'); using the deck language '{majority}'")
            lang = majority
        result[n] = lang
    return result


def find_source_text(workspace_dir, slide_num, source_lang, exclude_lang=None):
    """Return (lang, path) of the note text of a slide in the source language, or (None, None).

    With source_lang 'auto', any existing slide text is used (Japanese first, then
    English, then other languages), skipping the file of exclude_lang.
    """
    if source_lang and source_lang != "auto":
        if exclude_lang and lang_suffix(source_lang) == lang_suffix(exclude_lang):
            return None, None
        path = os.path.join(workspace_dir, text_filename(slide_num, source_lang))
        return (source_lang, path) if os.path.exists(path) else (None, None)
    candidates = []
    for name in os.listdir(workspace_dir):
        m = _TEXT_FILE_RE.match(name)
        if m and int(m.group(1)) == slide_num:
            candidates.append(normalize_lang(m.group(2)) if m.group(2) else "ja")
    candidates.sort(key=lambda lang: ({"ja": 0, "en": 1}.get(base_lang(lang), 2), lang))
    for lang in candidates:
        if exclude_lang and lang_suffix(lang) == lang_suffix(exclude_lang):
            continue
        return lang, os.path.join(workspace_dir, text_filename(slide_num, lang))
    return None, None

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s', datefmt='%H:%M:%S')
logger = logging.getLogger(__name__)

# ------------------------------------------
# Structured notes (translated narration + original note)
# ------------------------------------------
# When translated narration is written back into the slide notes, the note is laid out as
#
#   === pptx-narrator: narration [en] from [ja] #3f2a9c0d1e ===
#   <English narration>
#
#   === pptx-narrator: source [ja] ===
#   <original Japanese note>
#
# The hash identifies the version of the source note that was translated. --extract
# recognizes this layout: the source part becomes the note to translate, and the
# narration part is restored as the existing translation only if the source is unchanged.
_NARRATION_MARK_RE = re.compile(
    r"^[ \t]*=+[ \t]*pptx-narrator:[ \t]*narration[ \t]*\[([^\]\n]+)\][ \t]*from[ \t]*\[([^\]\n]+)\]"
    r"[ \t]*#([0-9a-fA-F]{6,40})([ \t]+spoken)?[ \t]*=+[ \t]*$", re.MULTILINE)
_SOURCE_MARK_RE = re.compile(
    r"^[ \t]*=+[ \t]*pptx-narrator:[ \t]*source[ \t]*\[([^\]\n]+)\][ \t]*=+[ \t]*$", re.MULTILINE)
TRANSLATION_MANIFEST = "translations.json"


def text_fingerprint(text):
    """Short hash of a note, insensitive to line-ending and trailing-space differences."""
    norm = unicodedata.normalize("NFC", text).replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")
    norm = "\n".join(line.rstrip() for line in norm.strip().split("\n"))
    return hashlib.sha1(norm.encode("utf-8")).hexdigest()[:10]


def compose_structured_note(narration_lang, narration_text, source_lang, source_text,
                            fingerprint=None, spoken=False):
    fingerprint = fingerprint or text_fingerprint(source_text)
    head = (f"=== pptx-narrator: narration [{narration_lang}] from [{source_lang}] "
            f"#{fingerprint}{' spoken' if spoken else ''} ===")
    return (f"{head}\n{narration_text.strip()}\n\n"
            f"=== pptx-narrator: source [{source_lang}] ===\n{source_text.strip()}")


def parse_structured_note(text):
    """Split a structured note; returns None for ordinary notes."""
    text = text.replace("\r\n", "\n").replace("\v", "\n")
    m_src = _SOURCE_MARK_RE.search(text)
    if not m_src:
        return None
    info = {"source_lang": normalize_lang(m_src.group(1)), "source_text": text[m_src.end():].strip(),
            "narration_lang": None, "narration_text": "", "fingerprint": None, "spoken": False}
    m_nar = _NARRATION_MARK_RE.search(text, 0, m_src.start())
    if m_nar:
        info.update(narration_lang=normalize_lang(m_nar.group(1)),
                    narration_text=text[m_nar.end():m_src.start()].strip(),
                    fingerprint=m_nar.group(3).lower(), spoken=bool(m_nar.group(4)))
    return info


def _load_manifest(workspace_dir):
    path = os.path.join(workspace_dir, TRANSLATION_MANIFEST)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"Could not read {path}: {e}")
    return {}


def record_translation(workspace_dir, slide_num, target_lang, source_lang, source_text):
    """Remember which version of the source note a translation was made from."""
    manifest = _load_manifest(workspace_dir)
    manifest.setdefault(str(slide_num), {})[target_lang] = {
        "source_lang": source_lang, "source_fingerprint": text_fingerprint(source_text)}
    with open(os.path.join(workspace_dir, TRANSLATION_MANIFEST), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1, sort_keys=True)


# ==========================================
# Helper Functions
# ==========================================
def is_japanese(text):
    return bool(re.search(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]', text))

def parse_slide_ranges(range_str, max_slides):
    slides = set()
    if not range_str: return set(range(1, max_slides + 1))
    for part in range_str.split(','):
        part = part.strip()
        if not part: continue
        try:
            if '-' in part:
                s, e = part.split('-', 1)
                slides.update(range(int(s) if s else 1, (int(e) if e else max_slides) + 1))
            else: slides.add(int(part))
        except ValueError: pass
    return slides

def load_letter_map(json_path: str) -> dict:
    """Loads letter pronunciation mappings from a JSON file."""
    if not json_path or not os.path.exists(json_path):
        return {}
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            logger.info(f"Loaded letter map from: {json_path}")
            return json.load(f)
    except Exception as e:
        logger.error(f"Failed to load letter map '{json_path}': {e}")
        return {}

def spell_out_letters(s: str, letter_map: dict = None) -> str:
    """Spells out characters using the supplied letter_map."""
    if not letter_map:
        return s
    return "".join(letter_map.get(ch.upper(), ch) for ch in s)

# ==========================================
# Dictionary & Term Scanning Logic
# ==========================================
def _is_single_letter(term):
    """A lone Latin letter or digit (e.g. the A of a base pair) is never a useful candidate."""
    return bool(re.fullmatch(r'[A-Za-z0-9]', term.strip()))


_KANJI_RE = re.compile(r'[一-鿿々]')


def japanese_compounds(text):
    """Compounds of several words, with the reading a Japanese front end assembles for them.

    A compound the front end has to put together (二|本|鎖) is where a reading can go wrong:
    鎖 is クサリ on its own but サ in 二本鎖. Returns {compound: assembled reading}; the
    reading is a proposal to correct, not an answer.
    """
    try:
        import pyopenjtalk
    except ImportError:
        return {}
    try:
        tokens = pyopenjtalk.run_frontend(text)
    except Exception:
        return {}
    out, run = {}, []

    def flush():
        if len(run) > 1:
            word = "".join(t["string"] for t in run)
            try:
                out[word] = unicodedata.normalize("NFKC", pyopenjtalk.g2p(word, kana=True))
            except Exception:
                pass
        run.clear()

    for token in tokens:
        if token.get("pos") == "名詞" and _KANJI_RE.search(token.get("string", "")):
            run.append(token)
        else:
            flush()
    flush()
    return out


def step_scan_and_update_dict(workspace_dir, dict_path, entries, requested_slides, target_lang,
                              source_lang="auto", for_translation=False, propose_compounds=False):
    """Collect candidate terms and append them to the dictionary file of the run.

    The scanned text is the text the dictionary will be applied to: the narration text in
    --target-lang when it exists, with provisional readings filled in; otherwise (or when
    combined with --translate) the notes to be translated, with blank replacements for review.
    """
    logger.info("--- [Option: Scan] Scanning text for technical terms ---")

    _ensure_nltk_data()
    from nltk.corpus import stopwords, words
    stop_words = set(stopwords.words('english'))
    common_english = set(w.lower() for w in words.words())

    narration_texts = [(os.path.join(workspace_dir, text_filename(n, target_lang)), target_lang)
                       for n in requested_slides
                       if os.path.exists(os.path.join(workspace_dir, text_filename(n, target_lang)))]
    if narration_texts and not for_translation:
        texts = narration_texts
    else:
        # nothing in --target-lang yet (or translating in this run): scan the notes to be translated
        texts = []
        for slide_num in requested_slides:
            src_lang, src_p = find_source_text(workspace_dir, slide_num, source_lang, exclude_lang=target_lang)
            if src_p is not None:
                texts.append((src_p, src_lang))
    if not texts:
        logger.info("No note texts to scan.")
        return

    jp_char = r'[぀-ヿ一-鿿]'
    latin = 'A-Za-zÀ-ÖØ-öø-ɏ'
    generic_pattern = rf'(?<![{latin}])[{latin}]{{2,}}(?![{latin}])|[゠-ヿ]{{3,}}'
    bypass_pattern = (
        r'\d+(?:\.\d+)?\s*(?:%|℃|°C|nm|μm|mm|cm|km|kg|mg|µg|ng|pg|ml|μl|kb|Mb|Gb|bp|kDa|Å)'
        r'|\b(?:I|II|III|IV|V|VI|VII|VIII|IX|X)\b'
    )

    existing_terms = {t.lower() for t, _, _ in entries or []}
    candidates = {}  # term -> is narration text (reading candidate)
    for p, file_lang in texts:
        with open(p, 'r', encoding='utf-8') as f:
            text = f.read()
        narration_text = lang_suffix(file_lang) == lang_suffix(target_lang)
        # English word lists only help for English words (English notes, or Latin words
        # embedded in Japanese). For other languages keep acronyms / mixed-case terms only.
        english_filter = base_lang(file_lang) in ("en", "ja")

        generic_found = re.findall(generic_pattern, text)
        bypass_found = set(re.findall(bypass_pattern, text))
        for m in re.finditer(rf'[A-Za-z0-9]+(?={jp_char})|(?<={jp_char})[A-Za-z0-9]+', text):
            tok = m.group()
            if re.search(r'[A-Za-z]', tok):
                bypass_found.add(tok)

        for term in generic_found:
            term_lower = term.lower()
            if _is_single_letter(term) or term_lower in existing_terms or term_lower in stop_words:
                continue
            if re.fullmatch(rf'[{latin}]+', term):
                acronym_like = term.isupper() or bool(re.search(r'[a-z][A-Z]', term))
                if len(term) < 3 and not term.isupper():
                    continue
                if english_filter and not term.isupper() and term_lower in common_english:
                    continue
                if not english_filter and not acronym_like:
                    continue
            candidates[term] = narration_text
            existing_terms.add(term_lower)

        for term in bypass_found:
            term_lower = term.lower()
            if _is_single_letter(term) or term_lower in existing_terms:
                continue
            candidates[term] = narration_text
            existing_terms.add(term_lower)

    if not candidates:
        logger.info("No new terms found.")
        return

    ROMAN_NUMERAL_READINGS = {
        "I": "いち", "II": "に", "III": "さん", "IV": "よん", "V": "ご",
        "VI": "ろく", "VII": "なな", "VIII": "はち", "IX": "きゅう", "X": "じゅう",
    }
    tgt_base = base_lang(target_lang)
    translator = None
    new_entries = []
    for term in sorted(candidates):
        repl = ""
        if candidates[term]:
            if tgt_base == "ja" and term in ROMAN_NUMERAL_READINGS:
                repl = ROMAN_NUMERAL_READINGS[term]
            elif re.match(r'^[A-Z]+$', term):
                repl = " ".join(term)
            elif re.match(r'^[a-zA-Z]+$', term) and tgt_base != "en":
                if translator is None:
                    try:
                        translator = GoogleTranslator(source='en', target=target_lang)
                    except Exception as e:
                        logger.warning(f"Reading guesses by translation disabled for '{target_lang}': {e}")
                        translator = False
                if translator:
                    try:
                        guess = translator.translate(term) or ""
                    except Exception:
                        guess = ""
                    if tgt_base == "ja":
                        repl = guess if (guess != term and is_japanese(guess)) else ""
                    else:
                        repl = guess if guess.strip().lower() != term.lower() else ""
            elif tgt_base == "ja" and is_japanese(term):
                repl = term
        new_entries.append((term, repl))
        logger.info(f"New term: {term} -> {repl or '(blank)'}")

    append_dictionary_entries(dict_path, new_entries)

    if propose_compounds and base_lang(target_lang) == "ja":
        known = {t for t, _, _ in entries or []} | {t for t, _ in new_entries}
        proposals = {}
        for path, file_lang in texts:
            if base_lang(file_lang) != "ja":
                continue
            with open(path, "r", encoding="utf-8") as f:
                for word, reading in japanese_compounds(f.read()).items():
                    if word not in known:
                        proposals[word] = reading
        if proposals:
            with open(dict_path, "a", encoding="utf-8", newline="") as f:
                f.write("\n; Compounds of the notes with the reading a Japanese front end assembles for them.\n"
                        "; They are comments, so they do nothing until the ';' is removed; correct the reading\n"
                        "; first (鎖 is read クサリ on its own but サ in 二本鎖) or leave the line as it is.\n")
                for word in sorted(proposals):
                    f.write(f";{word},{proposals[word]},\n")
            logger.info(f"{len(proposals)} compound(s) proposed as comment lines in {dict_path}")
    logger.info(f"Added {len(new_entries)} entries to {dict_path}")

SI_PREFIXES = {
    "p": "ピコ", "n": "ナノ", "μ": "マイクロ", "µ": "マイクロ", "m": "ミリ",
    "c": "センチ", "k": "キロ", "K": "キロ", "M": "メガ", "G": "ギガ", "T": "テラ",
}
BASE_UNITS = {
    "bp": "えんきつい",
    "b": "ベース",
    "m": "メートル",
    "g": "グラム",
    "l": "リットル",
    "L": "リットル",
    "Da": "ダルトン",
}
SPECIAL_UNITS = {
    "%": "パーセント",
    "℃": "ど",
    "°C": "ど",
    "Å": "オングストローム",
    "rpm": "アールピーエム",
}

def _unit_reading(unit, extra_units, use_builtin=True):
    if unit in extra_units:
        return extra_units[unit]
    if not use_builtin:
        return None
    if unit in SPECIAL_UNITS:
        return SPECIAL_UNITS[unit]
    for split in range(1, len(unit)):
        prefix, base = unit[:split], unit[split:]
        if prefix in SI_PREFIXES and base in BASE_UNITS:
            return SI_PREFIXES[prefix] + BASE_UNITS[base]
    if unit in BASE_UNITS:
        return BASE_UNITS[unit]
    return None

def normalize_units(text, extra_units=None, letter_map=None, lang="ja"):
    """Rewrite number + unit symbol into a spoken reading.

    Built-in SI-prefix readings exist for Japanese only; for other languages, readings
    come from dictionary entries of Type 'unit' (and the optional letter map).
    """
    extra_units = extra_units or {}
    use_builtin = base_lang(lang) == "ja"
    if not use_builtin and not extra_units and not letter_map:
        return text
    sep = "" if base_lang(lang) in CJK_LANGS else " "

    def _replace(m):
        number, letters = m.group(1), m.group(2)
        reading = _unit_reading(letters, extra_units, use_builtin)
        if reading is not None:
            return number + sep + reading
        if letter_map:
            return number + sep + spell_out_letters(letters, letter_map)
        return number + letters if use_builtin else m.group(0)

    pattern = r'(\d+(?:\.\d+)?)\s*([A-Za-zμµÅ°%℃]+)(?![A-Za-z0-9_])'
    return re.sub(pattern, _replace, text)

# ------------------------------------------
# Dictionaries
# ------------------------------------------
# A dictionary is a list of plain string replacements, one per CSV row:
#
#     string,replacement,type
#     Gbp,ギガベースペア,unit
#     mRNA,メッセンジャーアールエヌエー,
#
# It is applied to the text processed in the run: with --translate, to the notes
# before translation (e.g. to fix how terms are translated); with --tts, to the
# narration text before synthesis (e.g. readings); with both, at both points.
# Translation and synthesis can be run separately to use different dictionaries. type 'unit' marks unit symbols
# that follow a number. The header row is optional. Files in the v1.x layout
# (Term,Japanese_Reading,English_Reading,Type) are read using the column of the
# narration language (ja or en).
DICT_HEADER = ["string", "replacement", "type"]
_HEADER_NAMES = {"string", "term", "replacement", "reading", "type"}


COMMENT_MARKER = ";"


def _strip_dictionary_comment(line):
    """Drop a ';' comment from a dictionary line.

    The marker starts a comment only at the start of the line or after a space, and not
    inside double quotes, so a term that contains or begins with ';' is still possible."""
    in_quotes = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_quotes = not in_quotes
        elif ch == COMMENT_MARKER and not in_quotes and (i == 0 or line[i - 1].isspace()):
            return line[:i].strip()
    return line.strip()


def read_dictionary_file(path, lang=None):
    """Return [(string, replacement, type)] from one dictionary CSV."""
    with open(path, 'r', encoding='utf-8', newline='') as f:
        # Everything from a '#' to the end of the line is a comment, so a line can explain or
        # switch off an entry; a line with nothing left in front of the '#' is skipped. A '#'
        # inside double quotes is part of the term (e.g. "#1").
        lines = [_strip_dictionary_comment(line) for line in f]
        rows = [r for r in csv.reader([ln for ln in lines if ln]) if any(c.strip() for c in r)]
    if not rows:
        return []
    head = [c.strip().lower() for c in rows[0]]
    if "japanese_reading" in head or "english_reading" in head:
        column = {"ja": "japanese_reading", "en": "english_reading"}.get(base_lang(lang)) if lang else None
        if column not in head:
            logger.warning(f"{path} is a v1.x dictionary without a column for '{lang}'; it is not used.")
            return []
        term_col, repl_col = 0, head.index(column)
        type_col = head.index("type") if "type" in head else None
        body = rows[1:]
    else:
        term_col, repl_col, type_col = 0, 1, 2
        body = rows[1:] if len(head) >= 2 and head[0] in _HEADER_NAMES and head[1] in _HEADER_NAMES else rows
    entries = []
    for r in body:
        if len(r) <= repl_col or not r[term_col].strip():
            continue
        if "\\" in r[term_col]:
            logger.warning(f"{path}: the entry '{r[term_col].strip()}' contains a backslash and is matched literally; "
                           "write the string as it appears in the notes.")
        typ = r[type_col].strip().lower() if type_col is not None and len(r) > type_col else ""
        entries.append((r[term_col].strip(), r[repl_col].strip(), "unit" if typ == "unit" else ""))
    return entries


def load_dictionaries(paths, lang=None):
    """Read the given dictionary files in order; later files override earlier entries."""
    merged = {}
    for p in paths or []:
        if not os.path.exists(p):
            logger.info(f"Dictionary {p} does not exist yet.")
            continue
        entries = read_dictionary_file(p, lang)
        for term, repl, typ in entries:
            merged[(term, typ)] = repl
        logger.info(f"Dictionary {p}: {len(entries)} entries")
    return [(term, repl, typ) for (term, typ), repl in merged.items()]


def append_dictionary_entries(path, new_entries):
    """Append (string, replacement) rows; create the file with a header if needed."""
    is_new = not os.path.exists(path) or os.path.getsize(path) == 0
    needs_newline = False
    if not is_new:
        with open(path, 'rb') as fb:
            fb.seek(-1, os.SEEK_END)
            needs_newline = fb.read(1) not in (b'\n', b'\r')
    with open(path, 'a', encoding='utf-8', newline='') as f:
        if needs_newline:
            f.write('\n')
        writer = csv.writer(f)
        if is_new:
            writer.writerow(DICT_HEADER)
        for term, repl in new_entries:
            writer.writerow([term, repl, ""])


# Apostrophes and primes come in several shapes: PowerPoint turns ' into ’ while
# notes pasted from a paper may use ′. A dictionary entry written with one of them should
# still match the others, so both text and terms are mapped to the plain forms.
_LOOKALIKES = {ord(c): "'" for c in "‘’‛′ʹ´＇"}
_LOOKALIKES.update({ord(c): '"' for c in "“”‟″＂"})
_LOOKALIKES.update({ord(c): "-" for c in "‐‑‒–−－"})


def normalize_lookalikes(text):
    """Map typographic quotes, primes and dashes to their plain ASCII forms."""
    return text.translate(_LOOKALIKES)


def _replace_term(text, term, replacement):
    if re.match(r'^[a-zA-Z0-9_ \-]+$', term):
        pattern = rf'(?<![A-Za-z0-9_]){re.escape(term)}(?![A-Za-z0-9_])'
        return re.subn(pattern, lambda _m: replacement, text)
    return text.replace(term, replacement), text.count(term)


def apply_dictionary(text, entries, lang, letter_map=None, builtin_units=True):
    """Apply the replacements (longest string first), then rewrite number + unit expressions.

    builtin_units=False (used before translation) disables the built-in Japanese unit
    readings and the letter map, so that only the dictionary's own unit entries apply.
    """
    text = normalize_lookalikes(unicodedata.normalize('NFC', text))
    units, terms = {}, {}
    for term, repl, typ in entries or []:
        term = normalize_lookalikes(unicodedata.normalize('NFC', term))
        repl = unicodedata.normalize('NFC', repl)
        if repl:
            (units if typ == "unit" else terms)[term] = repl
    for term, repl in sorted(terms.items(), key=lambda x: len(x[0]), reverse=True):
        text, _ = _replace_term(text, term, repl)
    if builtin_units:
        return normalize_units(text, units, letter_map=letter_map, lang=lang)
    return normalize_units(text, units, letter_map=None, lang="und")

# ==========================================
# Pipeline Steps
# ==========================================
def step_extract_notes(pptx_path, workspace_dir, requested_slides, source_lang="auto"):
    logger.info("--- [Option: Extract] Extracting Notes ---")
    prs = Presentation(pptx_path)
    plain, structured = {}, {}
    for slide_num in requested_slides:
        if slide_num > len(prs.slides): continue
        slide = prs.slides[slide_num - 1]

        if slide._element.get('show') == '0':
            logger.info(f"Slide #{slide_num} is a hidden slide (skipped)")
            continue

        raw_text = ""
        if slide.has_notes_slide:
            text_list = []
            for shape in slide.notes_slide.shapes:
                if shape.has_text_frame:
                    text = shape.text.strip()
                    if text and not text.isdigit():
                        text_list.append(text)
            raw_text = "\n".join(text_list).strip()

        clean_text = raw_text.replace('​', '').replace('‌', '').replace('‍', '')
        txt = re.sub(r'\d{4}/\d+/\d+', '', clean_text).strip()
        if not txt:
            logger.info(f"Slide #{slide_num} has no notes (skipped)")
            continue
        info = parse_structured_note(txt)
        if info and info["source_text"]:
            structured[slide_num] = info
        else:
            plain[slide_num] = txt

    if source_lang == "auto":
        languages = detect_note_languages(plain, known={n: i["source_lang"] for n, i in structured.items()})
    else:
        # An explicit language is a selector. Do not relabel a note merely because
        # the caller requested a language; retain only notes whose detected language
        # agrees with the request. Short/ambiguous notes are accepted when the
        # detector cannot make a confident distinction and the requested language
        # is the only explicit choice.
        languages = {}
        for n, txt in plain.items():
            detected, confident = detect_language(txt)
            if confident and base_lang(detected) != base_lang(source_lang):
                logger.info(f"Slide #{n}: note detected as '{detected}', not requested '{source_lang}' (skipped)")
                continue
            languages[n] = source_lang

    written = 0
    for slide_num, txt in plain.items():
        if slide_num not in languages:
            continue
        name = text_filename(slide_num, languages[slide_num])
        with open(os.path.join(workspace_dir, name), "w", encoding="utf-8") as f:
            f.write(txt)
        written += 1
        logger.info(f"Slide #{slide_num}: note extracted to {name} (language: {languages[slide_num]}).")

    for slide_num, info in structured.items():
        if source_lang != "auto" and base_lang(info["source_lang"]) != base_lang(source_lang):
            logger.info(f"Slide #{slide_num}: structured note source is '{info['source_lang']}', not requested '{source_lang}' (skipped)")
            continue
        src_lang = info["source_lang"] if source_lang == "auto" else source_lang
        languages[slide_num] = src_lang
        src_name = text_filename(slide_num, src_lang)
        with open(os.path.join(workspace_dir, src_name), "w", encoding="utf-8") as f:
            f.write(info["source_text"])
        written += 1
        logger.info(f"Slide #{slide_num}: structured note; source part extracted to {src_name} (language: {src_lang}).")

        narr_lang = info["narration_lang"]
        if not narr_lang or not info["narration_text"] or info["spoken"] or lang_suffix(narr_lang) == lang_suffix(src_lang):
            continue
        tgt_path = os.path.join(workspace_dir, text_filename(slide_num, narr_lang))
        if info["fingerprint"] == text_fingerprint(info["source_text"]):
            with open(tgt_path, "w", encoding="utf-8") as f:
                f.write(info["narration_text"])
            record_translation(workspace_dir, slide_num, narr_lang, src_lang, info["source_text"])
            logger.info(f"Slide #{slide_num}: existing {narr_lang} narration restored to {os.path.basename(tgt_path)}.")
        else:
            stale_path = os.path.join(workspace_dir, f"slide_{slide_num}{lang_suffix(narr_lang)}.stale.txt")
            if os.path.exists(tgt_path):
                os.remove(tgt_path)
            with open(stale_path, "w", encoding="utf-8") as f:
                f.write(info["narration_text"])
            logger.warning(f"Slide #{slide_num}: the source note was edited after the {narr_lang} narration was "
                           f"translated; it will be translated again by --translate "
                           f"(previous narration kept in {os.path.basename(stale_path)}).")

    if source_lang == "auto" and languages:
        summary = Counter(languages.values()).most_common()
        logger.info("Note languages: " + ", ".join(f"{l} ({c} slides)" for l, c in summary))
    return written

# Google Translate accepts about five requests per second. The notes are translated
# line by line, which is well within that for one deck, but a long deck run back to
# back with another can cross it, so the requests are spaced and a refusal is waited out.
TRANSLATE_MIN_INTERVAL = 0.25
TRANSLATE_RETRIES = 4
_last_translate_call = [0.0]


def _translate_line(translator, text, attempts=TRANSLATE_RETRIES):
    """One translation request, spaced out and retried when the service refuses."""
    delay = 1.0
    for attempt in range(1, attempts + 1):
        wait = TRANSLATE_MIN_INTERVAL - (time.monotonic() - _last_translate_call[0])
        if wait > 0:
            time.sleep(wait)
        try:
            _last_translate_call[0] = time.monotonic()
            return translator.translate(text) or ""
        except Exception as e:
            if attempt == attempts or "too many requests" not in str(e).lower():
                raise
            logger.warning(f"Translation rate limit reached; waiting {delay:.0f}s "
                           f"(attempt {attempt} of {attempts - 1})")
            time.sleep(delay)
            delay *= 2
    return ""


def step_translate_notes(workspace_dir, requested_slides, source_lang, target_lang,
                         dictionary=None, overwrite=False):
    logger.info(f"--- [Option: Translate] Translating notes into '{target_lang}' ---")
    translators = {}
    sources_found = 0
    for slide_num in requested_slides:
        tgt_p = os.path.join(workspace_dir, text_filename(slide_num, target_lang))
        if not overwrite and os.path.exists(tgt_p) and os.path.getsize(tgt_p) > 0:
            sources_found += 1
            continue
        src_lang, src_p = find_source_text(workspace_dir, slide_num, source_lang, exclude_lang=target_lang)
        if src_p is None:
            continue
        sources_found += 1
        with open(src_p, "r", encoding="utf-8") as f:
            text = f.read().strip()
        if not text:
            continue

        key = (src_lang, target_lang)
        if key not in translators:
            try:
                translators[key] = GoogleTranslator(source=src_lang, target=target_lang)
            except Exception as e:
                logger.error(f"Translation {src_lang} -> {target_lang} is not available: {e}")
                return
        translator = translators[key]

        translated_lines = []
        for line in text.split('\n'):
            if not line.strip():
                translated_lines.append("")
                continue
            line_in = line.strip()
            if dictionary:
                line_in = apply_dictionary(line_in, dictionary, src_lang, builtin_units=False)
            try:
                translated_lines.append(_translate_line(translator, line_in))
            except Exception as e:
                logger.error(f"Slide {slide_num} translation error: {e}")
                translated_lines.append("")

        if any(translated_lines):
            with open(tgt_p, "w", encoding="utf-8") as out_f:
                out_f.write('\n'.join(translated_lines))
            record_translation(workspace_dir, slide_num, target_lang, src_lang, text)
            logger.info(f"Slide #{slide_num}: translated {src_lang} -> {target_lang}.")
    return sources_found

def step_generate_audio(
    workspace_dir,
    requested_slides,
    lang,
    ref_wav,
    ref_text_f,
    ref_lang,
    api_url,
    dictionaries,
    model_label,
    enable_drc=False,
    drc_threshold=-20.0,
    drc_ratio=3.0,
    letter_map=None,
):
    logger.info("--- [Option: TTS / engine=gpt_sovits] Generating Audio ---")
    text_lang = gpt_sovits_language(lang)
    prompt_lang = gpt_sovits_language(ref_lang)
    if text_lang is None or prompt_lang is None:
        logger.error(f"GPT-SoVITS supports only: {', '.join(GPT_SOVITS_LANGUAGES)} "
                     f"(narration '{lang}', reference '{ref_lang}')")
        return
    with open(ref_text_f, "r", encoding="utf-8") as f:
        ref_txt = f.read().strip()
    tts_url = api_url.rstrip("/") + "/tts"
    stats = SynthesisStats("gpt_sovits", api_url)

    for slide_num in requested_slides:
        txt_p = os.path.join(workspace_dir, text_filename(slide_num, lang))
        if not os.path.exists(txt_p):
            continue

        with open(txt_p, "r", encoding="utf-8") as f:
            spoken_text = apply_dictionary(f.read().strip(), dictionaries, lang, letter_map=letter_map)
        if not spoken_text:
            continue

        with open(os.path.join(workspace_dir, spoken_filename(slide_num, lang, model_label)),
                  "w", encoding="utf-8") as f:
            f.write(spoken_text)

        payload = {
            "text": spoken_text,
            "text_lang": text_lang,
            "ref_audio_path": os.path.abspath(ref_wav),
            "prompt_text": ref_txt,
            "prompt_lang": prompt_lang,
            "text_split_method": "cut5",
            "media_type": "wav",
        }

        started = time.time()
        try:
            res = requests.post(tts_url, json=payload, timeout=600)
            if res.status_code == 200:
                wav_p = os.path.join(workspace_dir, "temp.wav")
                with open(wav_p, "wb") as f:
                    f.write(res.content)

                audio = AudioSegment.from_file(wav_p)
                if enable_drc:
                    audio = compress_dynamic_range(audio, threshold=drc_threshold, ratio=drc_ratio)
                    audio = normalize(audio)
                audio.export(os.path.join(workspace_dir, audio_filename(slide_num, lang, model_label)),
                             format="ipod")
                os.remove(wav_p)
                stats.add(slide_num, time.time() - started, len(audio))
            else:
                logger.error(f"Slide {slide_num} TTS failed! Server returned [{res.status_code}]: {res.text}")
        except Exception as e:
            logger.error(f"Slide {slide_num} TTS connection error: {e}")
    stats.report()

class SynthesisStats:
    """Collects synthesis times so that a run reports its cost (see --tts)."""

    def __init__(self, engine, device):
        self.engine, self.device = engine, device
        self.slides = 0
        self.seconds = 0.0
        self.audio_seconds = 0.0

    def add(self, slide_num, seconds, audio_ms):
        self.slides += 1
        self.seconds += seconds
        self.audio_seconds += audio_ms / 1000.0
        logger.info(f"Slide {slide_num}: {audio_ms / 1000.0:.1f} s of audio synthesized in {seconds:.1f} s "
                    f"({seconds / max(audio_ms / 1000.0, 1e-9):.2f} x audio time).")

    def report(self):
        if not self.slides:
            return
        logger.info(f"Synthesis summary ({self.engine}, {self.device}): {self.slides} slides, "
                    f"{self.audio_seconds / 60:.1f} min of audio in {self.seconds / 60:.1f} min "
                    f"({self.audio_seconds / max(self.seconds, 1e-9):.2f} x real time, "
                    f"{self.seconds / self.slides:.1f} s per slide).")


def split_into_chunks(text):
    """Split text into sentences: after 。！？, or after . ! ? followed by whitespace."""
    parts = re.split(r'(?<=[。！？])|(?<=[.!?])\s+', text)
    return [p.strip() for p in parts if p and p.strip()]

# Qwen3-TTS 12Hz models emit 12 codec frames per second of audio, so a cap on the
# number of new tokens is a cap on the duration.
QWEN3_CODEC_HZ = 12
RUNAWAY_FACTOR = 3.0


def chunk_token_budget(chunk_text, chars_per_sec, hz=QWEN3_CODEC_HZ, factor=RUNAWAY_FACTOR):
    """Most codec tokens a chunk of text can plausibly need, with room to spare.

    The same margin the anomaly check uses, so a generation is cut off only where it
    would have been rejected anyway.
    """
    expected_sec = max(len(chunk_text) / chars_per_sec, 1.0)
    return int(round(max(expected_sec * factor, 20.0) * hz)) + hz


_qwen3_model_cache = {}

def resolve_torch_device(device):
    """Resolve 'auto' to cuda:0, mps or cpu depending on what is available."""
    if device != "auto":
        return device
    try:
        import torch
    except ImportError:
        return "cpu"
    if torch.cuda.is_available():
        return "cuda:0"
    mps = getattr(torch.backends, "mps", None)
    if mps is not None and mps.is_available():
        return "mps"
    return "cpu"

def _load_qwen3_model(model_size, device):
    device = resolve_torch_device(device)
    key = (model_size, device)
    if key not in _qwen3_model_cache:
        import torch
        from qwen_tts import Qwen3TTSModel
        try:
            import transformers
            transformers.logging.set_verbosity_error()
        except Exception:
            pass
        model_name = f"Qwen/Qwen3-TTS-12Hz-{model_size}-Base"
        logger.info(f"Loading Qwen3-TTS model ({model_name}, device={device})...")
        base_kwargs = dict(device_map=device)
        if str(device).startswith("cuda"):
            import importlib.util
            if importlib.util.find_spec("flash_attn") is not None:
                base_kwargs["attn_implementation"] = "flash_attention_2"

        model = None
        last_err = None
        for dtype in (torch.bfloat16, torch.float16, torch.float32):
            try:
                model = Qwen3TTSModel.from_pretrained(model_name, dtype=dtype, **base_kwargs)
                logger.info(f"Successfully loaded model with dtype={dtype}")
                break
            except Exception as e:
                last_err = e
                logger.warning(f"Failed to load with dtype={dtype}: {e}")
        if model is None:
            raise last_err
        _qwen3_model_cache[key] = model
    return _qwen3_model_cache[key]

def step_generate_audio_qwen3(
    workspace_dir,
    requested_slides,
    lang,
    ref_wav,
    ref_text_f,
    dictionaries,
    model_label,
    qwen3_model_size,
    qwen3_device,
    enable_drc=False,
    drc_threshold=-20.0,
    drc_ratio=3.0,
    letter_map=None,
):
    logger.info("--- [Option: TTS / engine=qwen3] Generating Audio ---")
    language = qwen3_language(lang)
    if language is None:
        logger.error(f"Qwen3-TTS does not support '{lang}' (supported: {', '.join(QWEN3_LANGUAGES)})")
        return
    try:
        import numpy as np
        import soundfile as sf
    except ImportError as e:
        logger.error(f"Missing required libraries: {e} (pip install 'pptx-narrator[qwen3]')")
        return

    with open(ref_text_f, "r", encoding="utf-8") as f:
        ref_txt = f.read().strip()

    model = _load_qwen3_model(qwen3_model_size, qwen3_device)
    stats = SynthesisStats(f"qwen3/{qwen3_model_size}", resolve_torch_device(qwen3_device))

    voice_clone_prompt = None
    try:
        logger.info("Analyzing reference audio (create_voice_clone_prompt)...")
        voice_clone_prompt = model.create_voice_clone_prompt(ref_audio=ref_wav, ref_text=ref_txt)
    except AttributeError:
        logger.warning(
            "This qwen_tts version lacks create_voice_clone_prompt; falling"
            " back to passing ref_audio/ref_text per invocation."
        )

    # rough speaking rate used to detect runaway generation
    chars_per_sec = 6.0 if base_lang(lang) in CJK_LANGS else 14.0

    for slide_num in requested_slides:
        txt_p = os.path.join(workspace_dir, text_filename(slide_num, lang))
        if not os.path.exists(txt_p):
            continue

        with open(txt_p, "r", encoding="utf-8") as f:
            spoken_text = apply_dictionary(f.read().strip(), dictionaries, lang, letter_map=letter_map)
        if not spoken_text:
            continue

        with open(os.path.join(workspace_dir, spoken_filename(slide_num, lang, model_label)),
                  "w", encoding="utf-8") as f:
            f.write(spoken_text)

        chunks = split_into_chunks(spoken_text)
        started = time.time()
        try:
            wavs = []
            sr = None

            def _gen_chunk(chunk_text):
                # Stop the model at a length the text cannot justify instead of letting a
                # runaway generation finish and then discarding it: the codec runs at a
                # fixed frame rate, so the budget follows from the expected duration.
                kw = {"max_new_tokens": chunk_token_budget(chunk_text, chars_per_sec)}
                if voice_clone_prompt is not None:
                    return model.generate_voice_clone(
                        text=chunk_text, language=language, voice_clone_prompt=voice_clone_prompt, **kw
                    )
                return model.generate_voice_clone(
                    text=chunk_text, language=language, ref_audio=ref_wav, ref_text=ref_txt, **kw
                )

            def _is_anomalous(chunk_text, wav, sample_rate):
                expected_sec = max(len(chunk_text) / chars_per_sec, 1.0)
                actual_sec = len(wav) / sample_rate
                return actual_sec > max(expected_sec * 3, 20)

            for chunk in chunks:
                w, sr = _gen_chunk(chunk)
                if _is_anomalous(chunk, w[0], sr):
                    logger.warning(
                        f"Slide {slide_num}: Chunk '{chunk[:20]}...' abnormally long"
                        f" ({len(w[0])/sr:.1f}s). Regenerating..."
                    )
                    w, sr = _gen_chunk(chunk)
                    if _is_anomalous(chunk, w[0], sr):
                        logger.error(
                            f"Slide {slide_num}: Chunk remains abnormally long"
                            f" ({len(w[0])/sr:.1f}s): '{chunk}'"
                        )
                wavs.append(w[0])
            combined = np.concatenate(wavs) if len(wavs) > 1 else wavs[0]

            wav_p = os.path.join(workspace_dir, "temp_qwen3.wav")
            sf.write(wav_p, combined, sr)

            audio = AudioSegment.from_file(wav_p)
            if enable_drc:
                audio = compress_dynamic_range(audio, threshold=drc_threshold, ratio=drc_ratio)
                audio = normalize(audio)
            audio.export(os.path.join(workspace_dir, audio_filename(slide_num, lang, model_label)),
                         format="ipod")

            os.remove(wav_p)
            stats.add(slide_num, time.time() - started, len(audio))
        except Exception as e:
            logger.error(f"Slide {slide_num} Qwen3-TTS generation error: {e}")
    stats.report()

_KANA_PUNCT_RE = re.compile(r"[\s、。，．,.!?！？「」『』（）()・…]")

def levenshtein_distance(a, b):
    """Minimum number of insertions, deletions and substitutions turning a into b."""
    if len(a) < len(b):
        a, b = b, a
    previous = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        current = [i]
        for j, cb in enumerate(b, 1):
            current.append(min(previous[j] + 1,
                               current[j - 1] + 1,
                               previous[j - 1] + (ca != cb)))
        previous = current
    return previous[-1]

def character_error_rate(reference, hypothesis):
    """CER = edit_distance(reference, hypothesis) / len(reference)."""
    if not reference:
        return 0.0 if not hypothesis else 1.0
    return levenshtein_distance(reference, hypothesis) / len(reference)

def kana_sequence_scores(kana_ref, kana_hyp):
    """Return (similarity, CER) for two katakana strings (punctuation/whitespace ignored)."""
    kana_ref = _KANA_PUNCT_RE.sub("", kana_ref)
    kana_hyp = _KANA_PUNCT_RE.sub("", kana_hyp)
    similarity = difflib.SequenceMatcher(None, kana_ref, kana_hyp, autojunk=False).ratio()
    return similarity, character_error_rate(kana_ref, kana_hyp)

def kana_scores(text_intended, text_asr):
    """Convert both texts to katakana with pyopenjtalk; return (similarity, CER, kana_intended, kana_asr)."""
    import pyopenjtalk
    kana_a = _KANA_PUNCT_RE.sub("", pyopenjtalk.g2p(text_intended, kana=True))
    kana_b = _KANA_PUNCT_RE.sub("", pyopenjtalk.g2p(text_asr, kana=True))
    similarity, cer = kana_sequence_scores(kana_a, kana_b)
    return similarity, cer, kana_a, kana_b

_NON_WORD_RE = re.compile(r"[\W_]+")

def normalize_for_comparison(text):
    """NFKC, case folding, and removal of punctuation, symbols and whitespace."""
    return _NON_WORD_RE.sub("", unicodedata.normalize("NFKC", text).casefold())

def difference_list(a, b, context=8, minimum=1):
    """The places where two sequences differ: (position, intended, recognized, before, after).

    This is what a reader of the report actually works with: the fragment that was meant
    and the fragment the ASR heard, in the order they occur, with a little context."""
    out = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag != "equal" and max(i2 - i1, j2 - j1) >= minimum:
            out.append((i1, a[i1:i2], b[j1:j2], a[max(0, i1 - context):i1], a[i2:i2 + context]))
    return out


def difference_runs(a, b):
    """(number of differing stretches, characters in the longest one) for two sequences.

    The ratio-based scores say how much of a slide matched; these say how many places
    differ and how long the longest difference is, which is what points at a skipped
    phrase rather than at scattered recognition differences."""
    runs = [max(i2 - i1, j2 - j1)
            for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes()
            if tag != "equal"]
    return len(runs), max(runs, default=0)


def text_scores(text_intended, text_asr):
    """Character-level similarity and CER on normalized text; return (similarity, CER, norm_intended, norm_asr)."""
    a = normalize_for_comparison(text_intended)
    b = normalize_for_comparison(text_asr)
    similarity = difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()
    return similarity, character_error_rate(a, b), a, b

def step_verify_audio(workspace_dir, requested_slides, lang, model_label,
                       asr_model_size="small", asr_device="cpu", threshold=0.85,
                       cer_threshold=None, max_difference=40, min_difference=4):
    logger.info("--- [Option: Verify] ASR round-trip check ---")
    is_ja = base_lang(lang) == "ja"
    try:
        from faster_whisper import WhisperModel
        if is_ja:
            import pyopenjtalk  # noqa: F401
    except ImportError as e:
        logger.error(f"Missing dependency: {e} (pip install 'pptx-narrator[verify]')")
        return

    logger.info(f"Loading Whisper model ({asr_model_size}, device={asr_device})...")
    asr_model = WhisperModel(asr_model_size, device=asr_device, compute_type="int8")

    results, differences = [], []
    for slide_num in requested_slides:
        audio_p = os.path.join(workspace_dir, audio_filename(slide_num, lang, model_label))
        spoken_p = os.path.join(workspace_dir, spoken_filename(slide_num, lang, model_label))
        if not (os.path.exists(audio_p) and os.path.exists(spoken_p)):
            continue

        with open(spoken_p, "r", encoding="utf-8") as f:
            intended_text = f.read().strip()
        if not intended_text:
            continue

        try:
            segments, _ = asr_model.transcribe(audio_p, language=base_lang(lang))
            asr_text = "".join(seg.text for seg in segments).strip()
        except Exception as e:
            logger.error(f"Slide {slide_num}: transcription failed: {e}")
            continue

        try:
            if is_ja:
                score, cer, norm_intended, norm_asr = kana_scores(intended_text, asr_text)
            else:
                score, cer, norm_intended, norm_asr = text_scores(intended_text, asr_text)
        except Exception as e:
            logger.error(f"Slide {slide_num}: comparison failed: {e}")
            continue

        n_runs, worst_run = difference_runs(norm_intended, norm_asr)
        for pos, said, heard, before, after in difference_list(norm_intended, norm_asr, minimum=min_difference):
            differences.append((slide_num, max(len(said), len(heard)), pos, said, heard, before, after))

        # Latin-script words left in Japanese narration cannot be compared reliably as kana
        has_latin = is_ja and bool(re.search(r'[A-Za-z]{2,}', intended_text))
        long_difference = max_difference is not None and worst_run > max_difference
        failed = (score < threshold or long_difference
                  or (cer_threshold is not None and cer > cer_threshold))
        status = "ENGLISH" if has_latin else ("FLAGGED" if failed else "OK")

        results.append((slide_num, round(score, 4), round(cer, 4), status, n_runs, worst_run,
                        intended_text, asr_text, norm_intended, norm_asr))
        logger.info(f"Slide {slide_num}: similarity={score:.2f}, CER={cer:.2f}, "
                    f"{n_runs} difference(s), longest {worst_run} characters [{status}]")

    if not results:
        # The file names carry the engine and model, so looking for the wrong label is
        # the usual reason nothing is found. Say which labels the workspace does hold.
        present = sorted({m.group(1) for m in
                          (re.match(r"^slide_\d+(?:_[^.]+)?\.(.+)\.m4a$", n)
                           for n in os.listdir(workspace_dir)) if m})
        logger.info(f"No slides with both audio and spoken text for '{model_label}' in {lang} "
                    "-- nothing to verify.")
        if present:
            logger.info("The workspace holds audio for: " + ", ".join(present)
                        + ". Give the engine that produced it, e.g. --engine qwen3.")
        return

    report_p = os.path.join(workspace_dir, f"verify_report{lang_suffix(lang)}.{model_label}.csv")
    with open(report_p, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["slide", "similarity", "cer", "status", "differences", "longest_difference",
                    "intended_text", "asr_text", "intended_normalized", "asr_normalized"])
        for row in sorted(results, key=lambda r: r[1]):
            w.writerow(row)

    n_flagged = sum(1 for r in results if r[3] == "FLAGGED")
    n_latin = sum(1 for r in results if r[3] == "ENGLISH")
    criterion = f"similarity < {threshold}"
    if max_difference is not None:
        criterion += f", a difference longer than {max_difference} characters"
    if cer_threshold is not None:
        criterion += f" or CER > {cer_threshold}"
    logger.info(f"Done: {n_flagged}/{len(results)} slide(s) flagged for review ({criterion}).")
    if n_latin:
        logger.info(f"{n_latin} slide(s) contain un-converted Latin-script words and need a listen.")
    logger.info(f"Report saved to: {os.path.basename(report_p)} (sorted worst-first)")

    if differences:
        diff_p = os.path.join(workspace_dir, f"verify_differences{lang_suffix(lang)}.{model_label}.csv")
        with open(diff_p, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["slide", "length", "position", "intended", "recognized", "before", "after"])
            for row in sorted(differences, key=lambda r: (-r[1], r[0], r[2])):
                w.writerow(row)
        logger.info(f"{len(differences)} difference(s) of at least {min_difference} characters listed in: "
                    f"{os.path.basename(diff_p)} "
                    "(longest first)")

_P14_MEDIA_RE = re.compile(r'<(p14:media)\b([^>]*?)(/?)>(?:(.*?)</p14:media>)?', re.DOTALL)
_P14_PLAYBACK_CHILD_RE = re.compile(
    r'<p14:(?:trim|fade)\b[^>]*/>|<p14:(?:trim|fade)\b[^>]*>.*?</p14:(?:trim|fade)>|<p14:bmkLst\b.*?</p14:bmkLst>|<p14:bmkLst\b[^>]*/>',
    re.DOTALL)

def clear_media_playback_settings(slide_xml, media_rel_ids):
    """Remove trim, fade and bookmark settings from the p14:media elements of the given relationships.

    Returns (new_xml, number of media elements changed)."""
    changed = 0

    def _fix(m):
        nonlocal changed
        attrs, self_closing, body = m.group(2), m.group(3), m.group(4) or ""
        embed = re.search(r'r:embed="([^"]+)"', attrs)
        if self_closing or not embed or embed.group(1) not in media_rel_ids:
            return m.group(0)
        new_body = _P14_PLAYBACK_CHILD_RE.sub("", body)
        if new_body == body:
            return m.group(0)
        changed += 1
        return f"<p14:media{attrs}/>" if not new_body.strip() else f"<p14:media{attrs}>{new_body}</p14:media>"

    return _P14_MEDIA_RE.sub(_fix, slide_xml), changed

def _read_text(path):
    if path and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""

RECORDED_CHOICES = {"all": ("pointer", "events"), "pointer": ("pointer",), "events": ("events",), "none": ()}


def step_pack_pptx(original_pptx, output_pptx, workspace_dir, requested_slides, lang, model_label,
                   source_lang="auto", writeback_notes=False,
                   remove_recorded=("pointer", "events"), icon_outside=True, pause_ms=1000):
    logger.info("--- [Option: Pack] Rebuilding PPTX ---")
    prs = Presentation(original_pptx)
    manifest = _load_manifest(workspace_dir)
    notes_written = 0
    for i, slide in enumerate(prs.slides):
        s_num = i + 1
        if s_num not in requested_slides:
            continue

        if writeback_notes:
            src_lang, src_p = find_source_text(workspace_dir, s_num, source_lang, exclude_lang=lang)
            text_p = os.path.join(workspace_dir, text_filename(s_num, lang))
            narration = _read_text(text_p)
            if src_p is not None:
                original, original_lang = _read_text(src_p), src_lang
            else:
                original, original_lang = _read_text(text_p), lang

            if not narration and not original:
                continue
            if narration and original and narration != original:
                fingerprint = None
                entry = manifest.get(str(s_num), {}).get(lang)
                if src_p is not None and entry and entry.get("source_fingerprint"):
                    fingerprint = entry["source_fingerprint"]
                    if fingerprint != text_fingerprint(original):
                        logger.warning(f"Slide #{s_num}: the {lang} narration was translated from an older version "
                                       f"of the source note (run --translate --retranslate to update it).")
                note_text = compose_structured_note(lang, narration, original_lang, original,
                                                    fingerprint=fingerprint)
            else:
                note_text = narration or original

            logger.info(f"Slide #{s_num}: Updating notes...")
            notes_slide = slide.notes_slide
            if notes_slide.notes_text_frame is not None:
                notes_slide.notes_text_frame.text = note_text
                notes_written += 1
            else:
                logger.warning(f"Slide #{s_num}: No text frame in notes slide (skipped)")

    tmp_pptx = os.path.join(workspace_dir, "tmp.pptx")
    prs.save(tmp_pptx)
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(tmp_pptx, 'r') as z:
            z.extractall(tmpdir)
        slide_w, slide_h = _slide_size(tmpdir)
        slide_parts = _slide_part_paths(tmpdir)
        embedded = 0
        for s_num in requested_slides:
            m4a_p = os.path.join(workspace_dir, audio_filename(s_num, lang, model_label))
            if os.path.exists(m4a_p) and s_num <= len(slide_parts):
                embed_slide_narration(tmpdir, s_num, m4a_p, len(AudioSegment.from_file(m4a_p)),
                                      slide_w, slide_h, slide_part=slide_parts[s_num - 1],
                                      remove_recorded=remove_recorded, icon_outside=icon_outside,
                                      pause_ms=pause_ms)
                embedded += 1
        _ensure_default_content_types(tmpdir, {"m4a": "audio/mp4", "png": "image/png"})
        _remove_unreferenced_media(tmpdir)
        archive_path = _zip_package(tmpdir, os.path.splitext(output_pptx)[0] + ".packing.zip")
        os.replace(archive_path, output_pptx)
    os.remove(tmp_pptx)
    logger.info(f"Packed {embedded} slide(s); wrote back notes on {notes_written} slide(s).")
    logger.info(f"Final output saved to: {output_pptx}")
    return embedded, notes_written


# ------------------------------------------
# Embedding narration audio in slide XML
# ------------------------------------------
# The structure written here follows what PowerPoint itself writes for recorded slide
# narration: a p:pic with a:audioFile / p14:media, audio + media + icon relationships,
# a p:timing that starts the audio with the slide (isNarration="1"), and the slide
# transition's advance time (advTm) set to the audio length.
REL_AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
REL_MEDIA = "http://schemas.microsoft.com/office/2007/relationships/media"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_RELS_EMPTY = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
NARRATION_ICON = "pptx_narrator_audio.png"

_NARRATION_TIMING = (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
    '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
    '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0">'
    '<p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst>'
    '<p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="1" presetClass="mediacall" '
    'presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
    '<p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="6" dur="1" fill="hold"/>'
    '<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>'
    '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
    '</p:seq><p:audio isNarration="1"><p:cMediaNode vol="80000" showWhenStopped="0"><p:cTn id="7" fill="hold" '
    'display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" '
    'delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/>'
    '</p:tgtEl></p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
)


def _rel_elements(rels_xml):
    return [(m.group(0), dict(re.findall(r'(\w+)="([^"]*)"', m.group(0))))
            for m in re.finditer(r'<Relationship\b[^>]*?/>', rels_xml)]


def _set_rel_target(rels_xml, rel_id, target):
    def _fix(m):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(0)))
        if attrs.get("Id") != rel_id:
            return m.group(0)
        return f'<Relationship Id="{rel_id}" Type="{attrs["Type"]}" Target="{target}"/>'
    return re.sub(r'<Relationship\b[^>]*?/>', _fix, rels_xml)


def _new_rel_id(rels_xml):
    used = {int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)}
    n = 1
    while n in used:
        n += 1
    return f"rId{n}"


def _add_rel(rels_xml, rel_type, target):
    rel_id = _new_rel_id(rels_xml)
    rel = f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>'
    return rels_xml.replace("</Relationships>", rel + "</Relationships>"), rel_id


def _slide_size(pkg_dir):
    try:
        with open(os.path.join(pkg_dir, "ppt", "presentation.xml"), encoding="utf-8") as f:
            m = re.search(r'<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', f.read())
        if m:
            return int(m.group(1)), int(m.group(2))
    except OSError:
        pass
    return 12192000, 6858000


def _speaker_icon_png(size=64):
    """A small grey loudspeaker icon (only visible in the editor; hidden during the slide show)."""
    rows = []
    for y in range(size):
        row = bytearray([0])
        for x in range(size):
            u, v = x / size, y / size
            body = 0.18 <= u <= 0.36 and 0.38 <= v <= 0.62
            cone = 0.36 <= u <= 0.58 and abs(v - 0.5) <= 0.12 + (u - 0.36) * 1.1
            r = ((u - 0.58) ** 2 + (v - 0.5) ** 2) ** 0.5
            wave = u > 0.62 and abs(v - 0.5) < (u - 0.5) and (0.12 < r < 0.17 or 0.24 < r < 0.29)
            row += bytes((90, 90, 90, 255)) if (body or cone or wave) else bytes((0, 0, 0, 0))
        rows.append(bytes(row))
    import struct
    import zlib

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows), 9)) + chunk(b"IEND", b""))


def _insert_after_transition_or_clrmap(slide_xml, fragment):
    m = re.search(r'<mc:AlternateContent\b(?:(?!</mc:AlternateContent>).)*?<p:transition\b.*?</mc:AlternateContent>',
                  slide_xml, re.DOTALL)
    if not m:
        m = re.search(r'<p:transition\b[^>]*/>|<p:transition\b.*?</p:transition>', slide_xml, re.DOTALL)
    if not m:
        m = re.search(r'</p:clrMapOvr>|<p:clrMapOvr\b[^>]*/>', slide_xml) or re.search(r'</p:cSld>', slide_xml)
    return slide_xml[:m.end()] + fragment + slide_xml[m.end():]


def _set_advance_time(slide_xml, dur_ms):
    if re.search(r'<p:transition\b', slide_xml):
        def _fix(m):
            tag = m.group(0)
            if 'advTm="' in tag:
                return re.sub(r'advTm="\d+"', f'advTm="{dur_ms}"', tag)
            return re.sub(r'^<p:transition\b', f'<p:transition advTm="{dur_ms}"', tag)
        return re.sub(r'<p:transition\b[^>]*>', _fix, slide_xml)
    return _insert_after_transition_or_clrmap(slide_xml, f'<p:transition advTm="{dur_ms}"/>')


def _slide_part_paths(pkg_dir):
    """Slide part paths (relative to the package root) in presentation order."""
    with open(os.path.join(pkg_dir, "ppt", "presentation.xml"), encoding="utf-8") as f:
        pres_xml = f.read()
    with open(os.path.join(pkg_dir, "ppt", "_rels", "presentation.xml.rels"), encoding="utf-8") as f:
        targets = {a.get("Id"): a.get("Target", "") for _, a in _rel_elements(f.read())}
    paths = []
    for rel_id in re.findall(r'<p:sldId\b[^>]*\br:id="([^"]+)"', pres_xml):
        target = targets.get(rel_id, "")
        target = target.lstrip("/") if target.startswith("/") else os.path.normpath(os.path.join("ppt", target))
        paths.append(target.replace(os.sep, "/"))
    return paths


# Data recorded while a slide show was presented: the laser-pointer path
# (p14:laserTraceLst, one time-stamped point per sample) and the media events of the
# recording (p14:showEvtLst: play, pause, seek, stop). Both are timed against the audio
# that was recorded with them, so they are meaningless once the audio is replaced.
_RECORDED_SHOW_DATA = {
    "pointer": ("<p14:laserTraceLst", "laser-pointer path"),
    "events": ("<p14:showEvtLst", "recorded playback events"),
}


def remove_recorded_show_data(slide_xml, kinds=("pointer", "events")):
    """Drop the p:ext elements holding the given kinds of recorded show data.

    Returns (new_xml, [descriptions of what was removed])."""
    removed = []
    markers = {_RECORDED_SHOW_DATA[k][0]: _RECORDED_SHOW_DATA[k][1] for k in kinds if k in _RECORDED_SHOW_DATA}

    out, pos = [], 0
    for m in re.finditer(r'<p:ext\b[^>]*[^/]>', slide_xml):
        if m.start() < pos:
            continue
        end, depth = m.end(), 1
        for t in re.finditer(r'<p:ext\b[^>]*[^/]>|</p:ext>', slide_xml[m.end():]):
            depth += 1 if t.group(0).startswith("<p:ext") else -1
            if depth == 0:
                end = m.end() + t.end()
                break
        element = slide_xml[m.start():end]
        what = next((w for marker, w in markers.items() if marker in element), None)
        if what:
            removed.append(what)
            out.append(slide_xml[pos:m.start()])
            pos = end
    out.append(slide_xml[pos:])
    slide_xml = "".join(out)
    if removed:
        slide_xml = re.sub(r'<p:extLst\s*>\s*</p:extLst>', "", slide_xml)
    return slide_xml, removed


ICON_GAP = 228600  # 0.25 inch between the slide edge and the icon parked next to it


def _picture_extent(pic_xml, default=812800):
    m = re.search(r'<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', pic_xml)
    return (int(m.group(1)), int(m.group(2))) if m else (default, default)


def _set_picture_offset(pic_xml, x, y):
    """Move a picture; returns the picture unchanged if it carries no position."""
    if not re.search(r'<a:off\b', pic_xml):
        return pic_xml
    return re.sub(r'<a:off\b[^>]*/>', f'<a:off x="{x}" y="{y}"/>', pic_xml, count=1)


def embed_slide_narration(pkg_dir, s_num, audio_path, dur_ms, slide_w=12192000, slide_h=6858000, slide_part=None,
                          remove_recorded=("pointer", "events"), icon_outside=True, pause_ms=1000):
    """Put the narration audio into slide s_num of an unzipped PPTX package.

    Each slide gets its own media file, so slides that shared one audio clip (e.g. copied
    slides) no longer overwrite each other. An existing narration object is re-pointed to
    it and its trim, fade and bookmarks are removed; a slide without audio gets a new
    auto-playing narration object. Returns "replaced", "inserted" or None.
    """
    slide_part = slide_part or f"ppt/slides/slide{s_num}.xml"
    slide_p = os.path.join(pkg_dir, *slide_part.split("/"))
    rels_p = os.path.join(os.path.dirname(slide_p), "_rels", os.path.basename(slide_p) + ".rels")
    if not os.path.exists(slide_p):
        return None
    with open(slide_p, encoding="utf-8") as f:
        slide_xml = f.read()
    rels_xml = _RELS_EMPTY
    if os.path.exists(rels_p):
        with open(rels_p, encoding="utf-8") as f:
            rels_xml = f.read()

    slide_xml, removed = remove_recorded_show_data(slide_xml, remove_recorded)
    for what in removed:
        logger.info(f"Slide #{s_num}: removed the {what} of the previous recording.")
    if "<p:contentPart" in slide_xml:
        logger.warning(f"Slide #{s_num}: the slide contains ink annotations; they are kept and may no longer "
                       "match the new narration.")

    media_dir = os.path.join(pkg_dir, "ppt", "media")
    os.makedirs(media_dir, exist_ok=True)
    media_name = f"pptx_narrator_slide{s_num}.m4a"
    shutil.copy(audio_path, os.path.join(media_dir, media_name))
    target = f"../media/{media_name}"

    # Prefer the object PowerPoint marks as narration; otherwise take the first audio object.
    narration_ids = set(re.findall(r'<p:audio\b[^>]*isNarration="1".*?<p:spTgt spid="(\d+)"', slide_xml, re.DOTALL))
    pic = None
    for m in re.finditer(r'<p:pic\b.*?</p:pic>', slide_xml, re.DOTALL):
        if '<a:audioFile' in m.group(0) or '<p14:media' in m.group(0):
            pic_id = re.search(r'<p:cNvPr\b[^>]*\bid="(\d+)"', m.group(0))
            if pic is None or (pic_id and pic_id.group(1) in narration_ids):
                pic = m.group(0)
            if pic_id and pic_id.group(1) in narration_ids:
                break

    if pic is not None:
        rel_ids = set(re.findall(r'<a:audioFile\b[^>]*r:link="([^"]+)"', pic))
        rel_ids |= set(re.findall(r'<p14:media\b[^>]*r:embed="([^"]+)"', pic))
        for rel_id in rel_ids:
            rels_xml = _set_rel_target(rels_xml, rel_id, target)
        if icon_outside:
            cx, cy = _picture_extent(pic)
            moved = _set_picture_offset(pic, -(cx + ICON_GAP), 0)
            if moved != pic:
                slide_xml = slide_xml.replace(pic, moved, 1)
                pic = moved
                logger.info(f"Slide #{s_num}: moved the audio icon next to the slide, out of the visible area.")
        slide_xml, n_cleared = clear_media_playback_settings(slide_xml, rel_ids)
        if n_cleared:
            logger.info(f"Slide #{s_num}: removed trim/fade/bookmark settings of the previous audio.")
        status = "replaced"
    else:
        if '<p:timing' in slide_xml:
            logger.warning(f"Slide #{s_num}: the slide has animations but no audio object; narration not inserted "
                           "(insert any audio clip on this slide in PowerPoint and pack again).")
            return None
        icon_p = os.path.join(media_dir, NARRATION_ICON)
        if not os.path.exists(icon_p):
            with open(icon_p, "wb") as f:
                f.write(_speaker_icon_png())
        rels_xml, rid_media = _add_rel(rels_xml, REL_MEDIA, target)
        rels_xml, rid_audio = _add_rel(rels_xml, REL_AUDIO, target)
        rels_xml, rid_icon = _add_rel(rels_xml, REL_IMAGE, f"../media/{NARRATION_ICON}")
        spid = max([int(i) for i in re.findall(r'<p:cNvPr\b[^>]*\bid="(\d+)"', slide_xml)] + [1]) + 1
        size = 812800
        if icon_outside:
            x, y = -(size + ICON_GAP), 0
        else:
            x, y = max(slide_w - size - 215900, 0), max(slide_h - size - 215900, 0)
        pic_xml = (
            f'<p:pic><p:nvPicPr><p:cNvPr id="{spid}" name="Narration {spid}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/>'
            f'</p:cNvPicPr><p:nvPr><a:audioFile r:link="{rid_audio}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}">'
            f'<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="{rid_media}"/>'
            f'</p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="{rid_icon}"/><a:stretch><a:fillRect/>'
            f'</a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{size}" cy="{size}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        )
        slide_xml = slide_xml.replace("</p:spTree>", pic_xml + "</p:spTree>", 1)
        slide_xml = _set_advance_time(slide_xml, dur_ms + pause_ms)
        slide_xml = _insert_after_transition_or_clrmap(slide_xml, _NARRATION_TIMING.replace("{spid}", str(spid)))
        status = "inserted"

    if status == "replaced":
        slide_xml = _set_advance_time(slide_xml, dur_ms + pause_ms)
    with open(slide_p, "w", encoding="utf-8") as f:
        f.write(slide_xml)
    os.makedirs(os.path.dirname(rels_p), exist_ok=True)
    with open(rels_p, "w", encoding="utf-8") as f:
        f.write(rels_xml)
    logger.info(f"Slide #{s_num}: narration {status} ({dur_ms} ms).")
    return status


def _ensure_default_content_types(pkg_dir, defaults):
    path = os.path.join(pkg_dir, "[Content_Types].xml")
    with open(path, encoding="utf-8") as f:
        xml = f.read()
    for ext, ctype in defaults.items():
        if not re.search(rf'<Default\b[^>]*Extension="{ext}"', xml, re.IGNORECASE):
            xml = xml.replace("<Default ", f'<Default Extension="{ext}" ContentType="{ctype}"/><Default ', 1)
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml)


def _remove_unreferenced_media(pkg_dir):
    media_dir = os.path.join(pkg_dir, "ppt", "media")
    if not os.path.isdir(media_dir):
        return
    referenced = set()
    for root, _, files in os.walk(pkg_dir):
        for name in files:
            if name.endswith(".rels"):
                with open(os.path.join(root, name), encoding="utf-8") as f:
                    referenced.update(os.path.basename(t) for t in re.findall(r'Target="([^"]+)"', f.read()))
    removed = [name for name in os.listdir(media_dir) if name not in referenced]
    for name in removed:
        os.remove(os.path.join(media_dir, name))
    if removed:
        ct_path = os.path.join(pkg_dir, "[Content_Types].xml")
        with open(ct_path, encoding="utf-8") as f:
            ct_xml = f.read()
        for name in removed:
            ct_xml = re.sub(rf'<Override\b[^>]*PartName="/ppt/media/{re.escape(name)}"[^>]*/>', "", ct_xml)
        with open(ct_path, "w", encoding="utf-8") as f:
            f.write(ct_xml)


def _zip_package(pkg_dir, zip_path):
    """Zip an unzipped OOXML package with [Content_Types].xml as the first entry."""
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(pkg_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for root, _, files in os.walk(pkg_dir):
            for name in sorted(files):
                full = os.path.join(root, name)
                arc = os.path.relpath(full, pkg_dir).replace(os.sep, "/")
                if arc != "[Content_Types].xml":
                    z.write(full, arc)
    return zip_path

# ==========================================
# Main CLI
# ==========================================
DESCRIPTION = """PPTX-Narrator: automated narration of PowerPoint presenter notes

Usage:
  pptx-narrator COMMAND [INPUT] [OPTIONS]

Commands:
  extract      Extract presenter notes from a PPTX into language-tagged text files.
  scan         Scan text input for technical terms and update a dictionary.
  translate   Translate text input from --in-lang to --out-lang.
  synthesize  Generate voice-cloned narration from text input.
  verify      ASR round-trip verification of generated narration.
  pack        Embed generated narration into a PPTX.

INPUT may be a file or directory. A directory is processed using the file types
accepted by the command. If INPUT is omitted, the last explicit input for that
command is reused only after its identity has been verified with SHA-256.
"""


# Configuration and execution-state files are deliberately separate.  The config
# is user-authored; the state is maintained by the program and records the last
# explicit inputs and their identities.
STATE_FILE = ".pptx_narrator_state.json"
RESOLVED_CONFIG_FILE = ".pptx_narrator_resolved.toml"


def _load_toml(path):
    if not path:
        return {}
    try:
        try:
            import tomllib
        except ModuleNotFoundError:
            import tomli as tomllib
    except ModuleNotFoundError:
        raise RuntimeError("TOML configuration requires Python 3.11+ or the 'tomli' package")
    with open(path, "rb") as f:
        return _normalize_config_keys(tomllib.load(f))


def _normalize_config_keys(data):
    """Accept hyphenated keys in the configuration file (dict-file as well as dict_file).

    TOML allows both spellings, and the command line accepts both, so the file should too.
    """
    if not isinstance(data, dict):
        return data
    return {str(k).replace("-", "_"): _normalize_config_keys(v) for k, v in data.items()}


def _deep_update(base, update):
    for key, value in update.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            _deep_update(base[key], value)
        else:
            base[key] = value
    return base


def _load_config(path):
    """Load explicit config, or the conventional local config when present."""
    if path:
        if not os.path.exists(path):
            raise RuntimeError(f"config file not found: {path}")
        return _load_toml(path), os.path.abspath(path)
    default = os.path.abspath("pptx_narrator.toml")
    if os.path.exists(default):
        return _load_toml(default), default
    return {}, None


def _load_state(path=STATE_FILE):
    if not os.path.exists(path):
        return {"version": 1, "commands": {}}
    try:
        with open(path, "r", encoding="utf-8") as f:
            state = json.load(f)
        state.setdefault("version", 1)
        state.setdefault("commands", {})
        return state
    except Exception as e:
        raise RuntimeError(f"could not read state file '{path}': {e}")


def _save_state(state, path=STATE_FILE):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2, sort_keys=True)
    os.replace(tmp, path)


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _suggest_command(command, input_path, workspace_hint, in_lang=None):
    """The command the user most likely meant, when the one given cannot work.

    Everything needed is already known -- which command was asked for, what was
    handed to it, and which workspace is in play -- so the correction is worked
    out rather than guessed at.
    """
    def _lang():
        return f" --in-lang {in_lang}" if in_lang else " --in-lang <lang>"

    if command in {"scan", "translate", "synthesize", "verify"}:
        if input_path and input_path.lower().endswith(".pptx"):
            # A deck was handed to a workspace command.
            if workspace_hint:
                return f"pptx-narrator {command} {workspace_hint}{_lang()}"
            guess = os.path.join(os.path.dirname(input_path) or ".",
                                 "workspace_" + os.path.splitext(os.path.basename(input_path))[0])
            if os.path.isdir(guess):
                return f"pptx-narrator {command} {os.path.relpath(guess)}{_lang()}"
            return (f"pptx-narrator extract {input_path} --workspace ws  # first, then\n"
                    f"  pptx-narrator {command} ws{_lang()}")
        if not input_path and not workspace_hint:
            return f"pptx-narrator {command} <workspace>{_lang()}"
    if command in {"extract", "pack"} and input_path and os.path.isdir(input_path):
        # A workspace was handed to a command that takes the deck.
        decks = sorted(n for n in os.listdir(os.path.dirname(input_path) or ".")
                       if n.lower().endswith(".pptx"))
        deck = decks[0] if len(decks) == 1 else "<deck>.pptx"
        return f"pptx-narrator {command} {deck} --workspace {input_path}"
    return None


def _did_you_mean(suggestion):
    return f"\n\nDid you mean:\n  {suggestion}" if suggestion else ""


def _is_workspace_file(name):
    """Whether a file name is one a workspace command can be pointed at."""
    return bool(_TEXT_FILE_RE.match(name)
                or re.match(r"^slide_\d+_[^.]+\.(.+)\.(m4a|spoken\.txt)$", name))


def _file_role(name):
    """What a workspace file is, for the record of a directory input."""
    if name.endswith(".spoken.txt"):
        return "spoken text"
    if name.endswith(".stale.txt"):
        return "set-aside narration"
    if name.startswith("verify_report"):
        return "verification report"
    if name.startswith("verify_differences"):
        return "verification differences"
    if name == TRANSLATION_MANIFEST:
        return "translation manifest"
    if _TEXT_FILE_RE.match(name):
        return "note text"
    if name.lower().endswith((".m4a", ".wav")):
        return "audio"
    if name.lower().endswith(".csv"):
        return "dictionary"
    if name.lower().endswith(".pptx"):
        return "deck"
    return "other"


def _input_snapshot(path, command):
    """Return a content-identity snapshot for a file or a command-relevant directory."""
    path = os.path.abspath(path)
    if os.path.isfile(path):
        return {
            "path": path,
            "kind": "file",
            "mtime": os.path.getmtime(path),
            "sha256": _sha256_file(path),
        }
    if os.path.isdir(path):
        extensions = {
            "scan": {".txt"},
            "translate": {".txt"},
            "synthesize": {".txt"},
            "verify": {".m4a", ".txt"},
        }.get(command, None)
        files = []
        for name in sorted(os.listdir(path)):
            full = os.path.join(path, name)
            if not os.path.isfile(full):
                continue
            if extensions is not None and os.path.splitext(name)[1].lower() not in extensions:
                continue
            files.append({
                "path": os.path.abspath(full),
                "role": _file_role(name),
                "mtime": os.path.getmtime(full),
                "sha256": _sha256_file(full),
            })
        return {"path": path, "kind": "directory", "files": files}
    raise RuntimeError(f"input not found: {path}")


def _snapshot_matches(saved, current):
    if not saved or saved.get("kind") != current.get("kind") or saved.get("path") != current.get("path"):
        return False
    if current["kind"] == "file":
        return saved.get("sha256") == current.get("sha256")
    old = {x["path"]: x.get("sha256") for x in saved.get("files", [])}
    new = {x["path"]: x.get("sha256") for x in current.get("files", [])}
    return old == new


def _resolve_input(command, explicit_input, state, workspace_dir=None):
    """Resolve explicit input or a verified last input for this command.

    A recorded input inside the workspace is stored relative to it, so it is
    resolved against the workspace rather than against the current directory.
    """
    if explicit_input:
        snap = _input_snapshot(explicit_input, command)
        return snap["path"], snap
    saved = state.get("commands", {}).get(command, {}).get("input")
    if not saved:
        raise RuntimeError(f"no previous input is recorded for '{command}'; specify INPUT explicitly")
    saved_path = saved["path"]
    if not os.path.isabs(saved_path) and workspace_dir:
        saved_path = os.path.normpath(os.path.join(workspace_dir, saved_path))
    current = _input_snapshot(saved_path, command)
    if not _snapshot_matches(_relativize_snapshot(saved, workspace_dir),
                             _relativize_snapshot(current, workspace_dir)):
        raise RuntimeError(
            f"the previously used input for '{command}' has changed; specify INPUT explicitly"
        )
    return current["path"], current




def _validate_input_language(path, command, in_lang):
    """Reject an explicitly language-tagged input whose language contradicts --in-lang."""
    if not in_lang or command == "extract":
        return
    langs = [in_lang] if isinstance(in_lang, str) else list(in_lang)
    wanted = {base_lang(x) for x in langs}
    candidates = []
    if os.path.isfile(path):
        candidates = [os.path.basename(path)]
    elif os.path.isdir(path):
        candidates = [n for n in os.listdir(path) if os.path.isfile(os.path.join(path, n))]
    tagged = []
    matched = []
    for name in candidates:
        m = _TEXT_FILE_RE.match(name)
        if not m or not m.group(2):
            # Audio filenames use the same slide_N_<lang> convention.
            m2 = re.match(r"^slide_\d+_([A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,4})?)\.", name)
            if not m2:
                continue
            actual = base_lang(normalize_lang(m2.group(1)))
        else:
            actual = base_lang(normalize_lang(m.group(2)))
        tagged.append((name, actual))
        if actual in wanted:
            matched.append(name)
    if os.path.isfile(path) and tagged and not matched:
        name, actual = tagged[0]
        raise RuntimeError(
            f"input language mismatch: '{name}' is '{actual}', but --in-lang is '{in_lang}'"
        )
    if os.path.isdir(path) and tagged and not matched:
        raise RuntimeError(
            f"no input data matching --in-lang '{in_lang}' was found in '{path}'"
        )


def _record_input(state, command, snapshot, extra=None):
    entry = dict(extra or {})
    entry["input"] = snapshot
    state.setdefault("commands", {})[command] = entry


def _toml_scalar(value):
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return None
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        vals = [_toml_scalar(x) for x in value]
        if any(x is None for x in vals):
            return None
        return "[" + ", ".join(vals) + "]"
    return None


def _write_toml_table(f, data, prefix=()):
    scalars = []
    tables = []
    for key, value in data.items():
        if isinstance(value, dict):
            tables.append((key, value))
        else:
            scalars.append((key, value))
    if prefix:
        f.write("[" + ".".join(prefix) + "]\n")
    for key, value in scalars:
        rendered = _toml_scalar(value)
        if rendered is not None:
            f.write(f"{key} = {rendered}\n")
    if scalars and tables:
        f.write("\n")
    for i, (key, value) in enumerate(tables):
        _write_toml_table(f, value, prefix + (key,))
        if i != len(tables) - 1:
            f.write("\n")


def _save_resolved_config(resolved, path=RESOLVED_CONFIG_FILE):
    with open(path, "w", encoding="utf-8") as f:
        f.write("# Resolved PPTX-Narrator configuration; all values are effective values.\n")
        _write_toml_table(f, resolved)


def _rel_to_workspace(path, workspace_dir):
    """A path inside the workspace, written relative to it.

    A workspace is moved and copied often, so what it records about its own
    contents must not depend on where it happens to sit. Anything outside it
    keeps its absolute path, because nothing else identifies it.
    """
    if not path or not workspace_dir:
        return path
    abs_path = os.path.abspath(path)
    abs_ws = os.path.abspath(workspace_dir)
    if abs_path == abs_ws:
        return "."
    if abs_path.startswith(abs_ws + os.sep):
        return os.path.relpath(abs_path, abs_ws)
    return abs_path


def _relativize_snapshot(snapshot, workspace_dir):
    """Rewrite the paths of an input record relative to the workspace."""
    if not snapshot:
        return snapshot
    out = dict(snapshot)
    out["path"] = _rel_to_workspace(out.get("path"), workspace_dir)
    if out.get("files"):
        out["files"] = [dict(f, path=_rel_to_workspace(f.get("path"), workspace_dir))
                        for f in out["files"]]
    return out


def _workspace_path(workspace_dir, filename):
    """Return the path of workspace-owned execution metadata."""
    return os.path.join(workspace_dir, filename)


def _append_history(workspace_dir, entry):
    """Append a compact, machine-readable record of one completed command."""
    path = _workspace_path(workspace_dir, ".pptx_narrator_history.jsonl")
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False, sort_keys=True) + "\n")


def _config_for_command(config, command):
    common = config.get("common", {}) if isinstance(config.get("common", {}), dict) else {}
    section = config.get(command, {}) if isinstance(config.get(command, {}), dict) else {}
    out = {}
    _deep_update(out, common)
    _deep_update(out, section)
    return out


def _add(group, *names, **kwargs):
    """Register an option under its hyphenated name plus the underscore spelling.

    A switch (store_true) also gets a --no-... counterpart, so that a value set to true
    in the configuration file can be taken back on the command line.
    """
    flags = list(names)
    for name in names:
        legacy = name.replace("-", "_").replace("__", "--", 1)
        if legacy != name and legacy not in flags:
            flags.append(legacy)
    group.add_argument(*flags, **kwargs)
    if kwargs.get("action") == "store_true":
        off = dict(kwargs)
        off["action"] = "store_false"
        off["help"] = argparse.SUPPRESS
        neg = ["--no-" + names[0][2:]]
        neg.append(neg[0].replace("-", "_").replace("__", "--", 1))
        group.add_argument(*dict.fromkeys(neg), **off)


def _add_common_options(parser, config_values):
    parser.add_argument("--config", default=None,
                        help="TOML configuration file (default: ./pptx_narrator.toml when present)")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    # These are intentionally not parser defaults: config values and built-in defaults
    # must be distinguishable when resolving the three configuration layers.
    parser.set_defaults(_config_values=config_values)


def _apply_cli_config_defaults(namespace, config_values):
    """Apply command-section values only where argparse did not receive an explicit value."""
    for key, value in config_values.items():
        if hasattr(namespace, key) and getattr(namespace, key) is None:
            setattr(namespace, key, value)


def build_parser(config_values=None):
    parser = argparse.ArgumentParser(
        prog="pptx-narrator",
        description=DESCRIPTION,
        formatter_class=argparse.RawTextHelpFormatter,
    )
    _add_common_options(parser, config_values or {})

    sub = parser.add_subparsers(dest="command", metavar="COMMAND")

    # ---- extract ---------------------------------------------------------
    p = sub.add_parser("extract", help="Extract presenter notes from a PPTX")
    p.add_argument("input", nargs="?", help="Input PPTX file")
    _add(p, "--in-lang", dest="in_lang", type=normalize_lang, default=None,
         help="Language(s) of notes to extract, e.g. ja or ja,en; omitted = all recognized languages")
    _add(p, "--workspace", dest="workspace", default=None,
         help="Output workspace directory (default: workspace_<filename>)")
    _add(p, "--slides", dest="slides", default=None, help="Slide range, e.g. 1-5 or 1,3,5- (default: all)")

    # ---- directory/text commands ---------------------------------------
    p = sub.add_parser("scan", help="Scan text input for technical terms")
    p.add_argument("input", nargs="?", help="Input text file or directory")
    _add(p, "--in-lang", dest="in_lang", type=normalize_lang, default=None, help="Language of the input text")
    _add(p, "--dict-file", dest="dict_file", action="append", default=None,
         help="Dictionary CSV to update; repeatable")
    _add(p, "--scan-compounds", dest="scan_compounds", action="store_true", default=None,
         help="With Japanese input, propose Japanese compounds in the dictionary")
    _add(p, "--slides", dest="slides", default=None,
         help="Slide range, e.g. 1-5 or 1,3,5- (default: all matching slides)")
    _add(p, "--workspace", dest="workspace", default=None,
         help="Optional workspace associated with the input (normally not needed)")

    p = sub.add_parser("translate", help="Translate text input")
    p.add_argument("input", nargs="?", help="Input text file or directory")
    _add(p, "--in-lang", dest="in_lang", type=normalize_lang, default=None, help="Language of the input text")
    _add(p, "--out-lang", dest="out_lang", type=normalize_lang, default=None, help="Language of the translated output")
    _add(p, "--dict-file", dest="dict_file", action="append", default=None,
         help="Dictionary CSV(s) applied before translation")
    _add(p, "--retranslate", dest="retranslate", action="store_true", default=None,
         help="Overwrite existing translations")
    _add(p, "--workspace", dest="workspace", default=None,
         help="Workspace containing input/output text")

    p = sub.add_parser("synthesize", help="Generate voice-cloned narration")
    p.add_argument("input", nargs="?", help="Input text file or directory")
    _add(p, "--in-lang", dest="in_lang", type=normalize_lang, default=None, help="Language of the input narration text")
    _add(p, "--dict-file", dest="dict_file", action="append", default=None,
         help="Dictionary CSV(s) applied before synthesis")
    _add(p, "--letter-map", dest="letter_map", default=None,
         help="JSON mapping of letters to readings")
    _add(p, "--engine", choices=["gpt_sovits", "qwen3"], default=None, help="TTS engine")
    _add(p, "--ref-wav", dest="ref_wav", default=None, help="Reference recording (.wav)")
    _add(p, "--ref-text-file", dest="ref_text_file", default=None, help="Transcript of --ref-wav")
    _add(p, "--ref-lang", dest="ref_lang", type=normalize_lang, default=None,
         help="Language of reference recording for GPT-SoVITS")
    _add(p, "--api-url", dest="api_url", default=None, help="GPT-SoVITS API server URL")
    _add(p, "--model", default=None, help="GPT-SoVITS model")
    _add(p, "--qwen3-model-size", dest="qwen3_model_size", choices=["0.6B", "1.7B"], default=None)
    _add(p, "--qwen3-device", dest="qwen3_device", default=None, help="auto / cuda:0 / mps / cpu")
    _add(p, "--enable-drc", dest="enable_drc", action="store_true", default=None)
    _add(p, "--drc-threshold", dest="drc_threshold", type=float, default=None)
    _add(p, "--drc-ratio", dest="drc_ratio", type=float, default=None)
    _add(p, "--workspace", dest="workspace", default=None,
         help="Workspace containing input text and generated audio")

    p = sub.add_parser("verify", help="Verify generated narration with ASR")
    p.add_argument("input", nargs="?", help="Input audio/text directory or file")
    _add(p, "--in-lang", dest="in_lang", type=normalize_lang, default=None, help="Language of the narration")
    _add(p, "--model", default=None, help="TTS model label used in filenames")
    _add(p, "--engine", choices=["gpt_sovits", "qwen3"], default=None,
         help="TTS engine, used to derive the model label")
    _add(p, "--qwen3-model-size", dest="qwen3_model_size", choices=["0.6B", "1.7B"], default=None)
    _add(p, "--asr-model", dest="asr_model", default=None)
    _add(p, "--asr-device", dest="asr_device", default=None)
    _add(p, "--verify-threshold", dest="verify_threshold", type=float, default=None)
    _add(p, "--min-difference", dest="min_difference", type=int, default=None)
    _add(p, "--max-difference", dest="max_difference", type=int, default=None)
    _add(p, "--cer-threshold", dest="cer_threshold", type=float, default=None)
    _add(p, "--workspace", dest="workspace", default=None,
         help="Workspace containing audio and spoken-text sidecars")

    # ---- pack ------------------------------------------------------------
    p = sub.add_parser("pack", help="Embed generated narration into a PPTX")
    p.add_argument("input", nargs="?", help="Original/input PPTX file")
    _add(p, "--workspace", dest="workspace", default=None,
         help="Workspace containing generated text/audio (required unless recoverable from state)")
    _add(p, "--out", default=None, help="Output PPTX path")
    _add(p, "--in-lang", dest="in_lang", type=normalize_lang, default=None, help="Language of the narration data being packed")
    _add(p, "--model", default=None, help="TTS model label used in filenames")
    _add(p, "--engine", choices=["gpt_sovits", "qwen3"], default=None)
    _add(p, "--qwen3-model-size", dest="qwen3_model_size", choices=["0.6B", "1.7B"], default=None)
    _add(p, "--slides", dest="slides", default=None, help="Slide range")
    _add(p, "--writeback-notes", dest="writeback_notes", action="store_true", default=None)
    _add(p, "--slide-pause", dest="slide_pause", type=float, default=None)
    _add(p, "--keep-audio-icon", dest="keep_audio_icon", action="store_true", default=None)
    _add(p, "--remove-recorded", dest="remove_recorded", choices=["all", "pointer", "events", "none"], default=None)


    # --config is read from anywhere in the command line; accept it after the
    # command as well, which is where people naturally write it.
    for _sp in sub.choices.values():
        _sp.add_argument("--config", default=None, help=argparse.SUPPRESS)

    return parser


def _defaults():
    return {
        "extract": {"workspace": None, "slides": None},
        "scan": {"scan_compounds": False, "slides": None},
        "translate": {"retranslate": False, "dict_file": None},
        "synthesize": {
            "engine": "qwen3", "ref_lang": "ja", "api_url": "http://127.0.0.1:9880/",
            "model": "v2ProPlus", "qwen3_model_size": "1.7B", "qwen3_device": "auto",
            "enable_drc": False, "drc_threshold": -20.0, "drc_ratio": 3.0, "dict_file": None,
            "letter_map": None,
        },
        "verify": {
            "model": "v2ProPlus", "engine": "qwen3", "qwen3_model_size": "1.7B",
            "asr_model": "small", "asr_device": "cpu", "verify_threshold": 0.85,
            "min_difference": 4, "max_difference": 40, "cer_threshold": None,
        },
        "pack": {
            "workspace": None, "out": "output.pptx", "model": "v2ProPlus", "engine": "qwen3",
            "qwen3_model_size": "1.7B", "slides": None, "writeback_notes": False,
            "slide_pause": 1.0, "keep_audio_icon": False,
            "remove_recorded": "all",
        },
    }


def _merge_effective(command, args, config, parser):
    values = _defaults().get(command, {}).copy()
    cfg = _config_for_command(config, command)
    _deep_update(values, cfg)
    for key, value in vars(args).items():
        if key.startswith("_") or key in {"command", "input", "config"}:
            continue
        # argparse uses None for omitted optional values; boolean flags also use None here.
        if value is None:
            continue
        # An empty value on the command line clears what the configuration file set,
        # since there is otherwise no way to take a parameter back: --dict-file ''.
        if value == "" or (isinstance(value, list) and value and not any(value)):
            values[key] = None
            continue
        values[key] = value
    return values


def _inherit_synthesis_settings(command, values, args, config, state):
    """Use the last successful synthesis in this workspace when not overridden.

    CLI and TOML values remain authoritative.  This prevents a Qwen3 synthesis
    followed by a bare ``pack`` from silently looking for GPT-SoVITS filenames.
    """
    if command not in {"verify", "pack"}:
        return values
    prior = state.get("commands", {}).get("synthesize", {}).get("resolved_config", {})
    if not prior:
        return values
    configured = _config_for_command(config, command)
    for key in ("in_lang", "engine", "model", "qwen3_model_size"):
        if getattr(args, key, None) is None and key not in configured and prior.get(key) is not None:
            values[key] = prior[key]
    return values


def _audio_model_labels(workspace_dir, lang):
    """Return model labels of audio files available for one narration language."""
    pattern = re.compile(r"^slide_\d+_" + re.escape(lang) + r"\.(.+)\.m4a$")
    return sorted({m.group(1) for name in os.listdir(workspace_dir)
                   if (m := pattern.match(name))})


def _text_slides(workspace_dir, lang):
    return [slide for slide in sorted(_slides_from_workspace(workspace_dir))
            if os.path.exists(os.path.join(workspace_dir, text_filename(slide, lang)))]


def _validate_and_normalize(command, v, parser):
    if command == "extract":
        if v.get("in_lang"):
            raw = v["in_lang"]
            if isinstance(raw, str):
                raw = raw.split(",")
            v["in_lang"] = [normalize_lang(x) for x in raw if normalize_lang(x)]
    elif command == "translate":
        if not v.get("in_lang") or not v.get("out_lang"):
            parser.error("translate requires --in-lang and --out-lang")
        if v["in_lang"] == "auto" or v["out_lang"] == "auto":
            parser.error("--in-lang/--out-lang cannot be 'auto'")
        if lang_suffix(v["in_lang"]) == lang_suffix(v["out_lang"]):
            parser.error("translate requires different --in-lang and --out-lang")
    elif command in {"scan", "synthesize", "verify", "pack"}:
        if not v.get("in_lang"):
            parser.error(f"{command} requires --in-lang")
        if v["in_lang"] == "auto":
            parser.error("--in-lang cannot be 'auto'")
    if command in {"scan", "translate", "synthesize"} and isinstance(v.get("dict_file"), str):
        v["dict_file"] = [v["dict_file"]]
    if command == "scan" and not v.get("dict_file"):
        parser.error("scan requires --dict-file")
    if command == "synthesize":
        missing = [flag for flag, key in (("--ref-wav", "ref_wav"), ("--ref-text-file", "ref_text_file")) if not v.get(key)]
        if missing:
            parser.error("synthesize requires " + " and ".join(missing))
        if v["engine"] == "qwen3" and qwen3_language(v["in_lang"]) is None:
            parser.error(f"Qwen3-TTS does not support '{v['in_lang']}' (supported: {', '.join(QWEN3_LANGUAGES)})")
        if v["engine"] == "gpt_sovits":
            for flag, code in (("--in-lang", v["in_lang"]), ("--ref-lang", v["ref_lang"])):
                if gpt_sovits_language(code) is None:
                    parser.error(f"GPT-SoVITS does not support {flag} '{code}' (supported: {', '.join(GPT_SOVITS_LANGUAGES)})")
            if v["model"] not in MODELS_CONFIG:
                parser.error(f"unknown GPT-SoVITS model '{v['model']}' (available: {', '.join(MODELS_CONFIG)})")
    if command == "verify":
        if v["engine"] == "qwen3":
            v["model"] = f"qwen3-{v['qwen3_model_size']}"
    if command == "pack" and v["engine"] == "qwen3":
        v["model"] = f"qwen3-{v['qwen3_model_size']}"
    return v


def main(argv=None):
    # Parse command first without loading a config so --config can be honored cleanly.
    bootstrap = argparse.ArgumentParser(add_help=False)
    bootstrap.add_argument("--config")
    bootstrap.add_argument("command", nargs="?")
    bootstrap.add_argument("input", nargs="?")
    boot, _ = bootstrap.parse_known_args(argv)
    try:
        config, config_path = _load_config(boot.config)
    except RuntimeError as e:
        raise SystemExit(str(e))

    parser = build_parser(config)
    # Nothing runs without a command, but say what the commands are instead of
    # only complaining that one is missing.
    if not boot.command:
        parser.print_help()
        raise SystemExit(1)
    args = parser.parse_args(argv)
    command = args.command
    effective = _merge_effective(command, args, config, parser)

    # Execution state belongs to the workspace, never to whatever directory
    # happened to be current when a command was run.  A directory INPUT remains
    # accepted as a backwards-compatible spelling of --workspace.
    workspace_hint = effective.get("workspace")
    if not workspace_hint and command in {"scan", "translate", "synthesize", "verify"} and args.input:
        if os.path.isdir(args.input) or not os.path.isfile(args.input):
            workspace_hint = args.input
    if command == "extract" and workspace_hint:
        workspace_hint = os.path.abspath(workspace_hint)
        os.makedirs(workspace_hint, exist_ok=True)
    elif command != "extract":
        if not workspace_hint:
            parser.error(f"{command} requires --workspace (or a workspace directory as INPUT)"
                         + _did_you_mean(_suggest_command(command, args.input, None,
                                                          effective.get("in_lang"))))
        workspace_hint = os.path.abspath(workspace_hint)
        if not os.path.isdir(workspace_hint):
            parser.error(f"workspace does not exist: {workspace_hint}\n\nCheck the path, or create it first:\n"
                         f"  pptx-narrator extract <deck>.pptx --workspace {workspace_hint}")

    state = _load_state(_workspace_path(workspace_hint, STATE_FILE)) if workspace_hint else _load_state()
    effective = _inherit_synthesis_settings(command, effective, args, config, state)
    effective = _validate_and_normalize(command, effective, parser)
    if command in {"verify", "pack"}:
        configured = _config_for_command(config, command)
        model_is_explicit = any(getattr(args, key, None) is not None
                                for key in ("engine", "model", "qwen3_model_size"))
        model_is_configured = any(key in configured for key in ("engine", "model", "qwen3_model_size"))
        labels = _audio_model_labels(workspace_hint, effective["in_lang"])
        if not model_is_explicit and not model_is_configured and len(labels) > 1:
            parser.error("multiple audio model labels are present in this workspace: "
                         + ", ".join(labels) + ". Specify --engine (and model size if applicable).")
    try:
        # Workspace commands need no redundant positional INPUT; the workspace
        # itself is their explicit data source.
        explicit_input = args.input or (workspace_hint if command in {"scan", "translate", "synthesize", "verify"} else None)
        input_path, input_snapshot = _resolve_input(command, explicit_input, state, workspace_hint)
        _validate_input_language(input_path, command, effective.get("in_lang"))
    except RuntimeError as e:
        parser.error(str(e))

    # extract writes a workspace; all text/audio commands operate directly on their
    # INPUT directory. pack consumes a PPTX and a workspace containing generated assets.
    workspace_dir = effective.get("workspace")
    if command == "extract":
        if not os.path.isfile(input_path) or not input_path.lower().endswith(".pptx"):
            parser.error("extract INPUT must be a PPTX file")
        if not workspace_dir:
            workspace_dir = os.path.join(os.path.dirname(input_path),
                                         "workspace_" + os.path.splitext(os.path.basename(input_path))[0])
        workspace_dir = os.path.abspath(workspace_dir)
        os.makedirs(workspace_dir, exist_ok=True)
        prs = Presentation(input_path)
        req_slides = sorted(parse_slide_ranges(effective.get("slides"), len(prs.slides)))
        langs = effective.get("in_lang")
        if not langs:
            step_extract_notes(input_path, workspace_dir, req_slides, "auto")
        else:
            # Extract each requested language independently. A PPTX may legitimately
            # contain notes in several languages; each requested language is therefore
            # a selector, not a claim that every note in the deck has that language.
            # A requested language the deck does not contain is an error: the run would
            # otherwise report success while producing nothing for that language.
            missing = []
            for lang in langs:
                if not step_extract_notes(input_path, workspace_dir, req_slides, lang):
                    missing.append(lang)
            if missing:
                parser.error("no note in " + ", ".join(missing) + " was found in " + os.path.basename(input_path))
        _record_input(state, command, input_snapshot, {"workspace": workspace_dir})
    else:
        file_workspace_tmp = None
        original_file_input = None
        if command in {"scan", "translate", "synthesize", "verify"}:
            if os.path.isdir(input_path):
                workspace_dir = input_path
            elif os.path.isfile(input_path):
                # A single file names one slide of a workspace. Anything else -- a deck,
                # most obviously -- would silently become a workspace of its own and find
                # nothing, so it is refused here rather than reported as missing audio.
                if not _is_workspace_file(os.path.basename(input_path)):
                    parser.error(f"{command} INPUT must be a workspace directory or one of its files "
                                 f"(slide_N_<lang>.txt or slide_N_<lang>.<model>.m4a), not "
                                 f"{os.path.basename(input_path)}"
                                 + _did_you_mean(_suggest_command(command, input_path, workspace_hint,
                                                                  effective.get("in_lang"))))
                owner = os.path.abspath(os.path.dirname(input_path) or ".")
                if workspace_hint and os.path.abspath(workspace_hint) != owner:
                    parser.error(f"INPUT {input_path} is not in --workspace {workspace_hint}; "
                                 "name the file inside that workspace, or drop --workspace")
                original_file_input = input_path
                file_workspace_tmp = _prepare_file_workspace(command, input_path, effective.get("in_lang"))
                workspace_dir = file_workspace_tmp
            else:
                parser.error(f"{command} INPUT must be a file or directory")
        elif command == "pack":
            if not os.path.isfile(input_path) or not input_path.lower().endswith(".pptx"):
                parser.error("pack INPUT must be a PPTX file")
            if not workspace_dir:
                # Prefer the workspace associated with the most recent extract of this PPTX.
                extract_state = state.get("commands", {}).get("extract", {})
                if extract_state.get("input") and _snapshot_matches(extract_state["input"], input_snapshot):
                    workspace_dir = extract_state.get("workspace")
                    if workspace_dir:
                        logger.info(f"Using the workspace of the last extract of this deck: {workspace_dir} "
                                    "(give --workspace to choose another).")
            if not workspace_dir or not os.path.isdir(workspace_dir):
                parser.error("pack requires --workspace unless a matching extract workspace is available")
            workspace_dir = os.path.abspath(workspace_dir)

    if command == "extract":
        effective["workspace"] = workspace_dir
    logger.info(f"Workspace: {os.path.abspath(workspace_dir)}")
    if command in {"scan", "translate", "synthesize", "verify", "pack"}:
        # For directories, the command's INPUT is the processing workspace; for pack it is the PPTX.
        pass

    # The remaining command implementations operate on the existing workspace-based functions.
    if command == "scan":
        dictionaries = load_dictionaries(effective["dict_file"], effective["in_lang"])
        available_slides = _text_slides(workspace_dir, effective["in_lang"])
        if not available_slides:
            parser.error(f"no {effective['in_lang']} slide text was found in workspace: {workspace_dir}")
        selected = sorted(parse_slide_ranges(effective.get("slides"), max(available_slides)))
        slides = [n for n in selected if n in set(available_slides)]
        if not slides:
            parser.error(f"no workspace slide matches --slides {effective.get('slides')!r}")
        step_scan_and_update_dict(workspace_dir, effective["dict_file"][0], dictionaries,
                                  slides, effective["in_lang"],
                                  source_lang=effective["in_lang"], for_translation=False,
                                  propose_compounds=effective["scan_compounds"])
    elif command == "translate":
        slides = sorted(_slides_from_workspace(workspace_dir))
        dictionaries = load_dictionaries(effective.get("dict_file"), effective["in_lang"])
        if not step_translate_notes(workspace_dir, slides, effective["in_lang"], effective["out_lang"],
                                    dictionary=dictionaries, overwrite=effective["retranslate"]):
            parser.error(f"no {effective['in_lang']} text was found in "
                         f"{os.path.basename(original_file_input or input_path)}")
    elif command == "synthesize":
        slides = _text_slides(workspace_dir, effective["in_lang"])
        if not slides:
            parser.error(f"no {effective['in_lang']} text was found in workspace: {workspace_dir}")
        dictionaries = load_dictionaries(effective.get("dict_file"), effective["in_lang"])
        letter_map_data = load_letter_map(effective.get("letter_map"))
        if effective["engine"] == "gpt_sovits":
            config_model = MODELS_CONFIG[effective["model"]]
            base_url = effective["api_url"].rstrip("/")
            logger.info(f"Switching GPT-SoVITS weights to {effective['model']}...")
            try:
                requests.get(f"{base_url}/set_gpt_weights", params={"weights_path": config_model["gpt"]}, timeout=300)
                requests.get(f"{base_url}/set_sovits_weights", params={"weights_path": config_model["sovits"]}, timeout=300)
            except requests.RequestException as e:
                parser.error(f"Could not reach the GPT-SoVITS API server at {effective['api_url']}: {e}")
        model_label = effective["model"] if effective["engine"] == "gpt_sovits" else f"qwen3-{effective['qwen3_model_size']}"
        if effective["engine"] == "qwen3":
            step_generate_audio_qwen3(workspace_dir, slides, effective["in_lang"], effective["ref_wav"],
                                      effective["ref_text_file"], dictionaries, model_label,
                                      effective["qwen3_model_size"], effective["qwen3_device"],
                                      enable_drc=effective["enable_drc"], drc_threshold=effective["drc_threshold"],
                                      drc_ratio=effective["drc_ratio"], letter_map=letter_map_data)
        else:
            step_generate_audio(workspace_dir, slides, effective["in_lang"], effective["ref_wav"],
                                effective["ref_text_file"], effective["ref_lang"], effective["api_url"],
                                dictionaries, model_label, enable_drc=effective["enable_drc"],
                                drc_threshold=effective["drc_threshold"], drc_ratio=effective["drc_ratio"],
                                letter_map=letter_map_data)
    elif command == "verify":
        model_label = effective["model"] if effective["engine"] == "gpt_sovits" else f"qwen3-{effective['qwen3_model_size']}"
        slides = [slide for slide in _text_slides(workspace_dir, effective["in_lang"])
                  if os.path.exists(os.path.join(workspace_dir, audio_filename(slide, effective["in_lang"], model_label)))]
        if not slides:
            available = _audio_model_labels(workspace_dir, effective["in_lang"])
            parser.error(f"verify found no {effective['in_lang']} audio for model '{model_label}'. "
                         + (f"Available model labels: {', '.join(available)}." if available else ""))
        step_verify_audio(workspace_dir, slides, effective["in_lang"], model_label,
                          effective["asr_model"], effective["asr_device"], effective["verify_threshold"],
                          effective["cer_threshold"], max_difference=effective["max_difference"] or None,
                          min_difference=effective["min_difference"])
    elif command == "pack":
        prs = Presentation(input_path)
        slides = sorted(parse_slide_ranges(effective.get("slides"), len(prs.slides)))
        model_label = effective["model"] if effective["engine"] == "gpt_sovits" else f"qwen3-{effective['qwen3_model_size']}"
        audio_paths = [os.path.join(workspace_dir, audio_filename(s, effective["in_lang"], model_label))
                       for s in slides]
        if not any(os.path.exists(p) for p in audio_paths):
            available = _audio_model_labels(workspace_dir, effective["in_lang"])
            suffix = f" Available model labels: {', '.join(available)}." if available else " No matching audio files exist."
            parser.error(f"pack found no {effective['in_lang']} audio for model '{model_label}'."
                         f" Expected {os.path.basename(audio_paths[0]) if audio_paths else 'slide_N_<lang>.<model>.m4a'}.{suffix}")
        output = os.path.abspath(effective["out"])
        # pack uses the narration language as its input-data language.
        step_pack_pptx(input_path, output, workspace_dir, slides, effective["in_lang"], model_label,
                       source_lang="auto", writeback_notes=effective["writeback_notes"],
                       remove_recorded=RECORDED_CHOICES[effective["remove_recorded"]],
                       icon_outside=not effective["keep_audio_icon"],
                       pause_ms=int(round(effective["slide_pause"] * 1000)))

    if command in {"scan", "translate", "synthesize", "verify"} and '"'"'file_workspace_tmp'"'"' in locals() and file_workspace_tmp:
        _sync_file_workspace(file_workspace_tmp, original_file_input, os.path.basename(original_file_input))
        shutil.rmtree(file_workspace_tmp, ignore_errors=True)

    # Record the input as it stands after the run. translate and synthesize write their
    # results into the directory they read, so recording the state from before the run
    # would make the next reuse fail on this run's own output.
    try:
        input_snapshot = _input_snapshot(input_path, command)
    except RuntimeError:
        pass
    input_snapshot = _relativize_snapshot(input_snapshot, workspace_dir)
    extra = {}
    if workspace_dir:
        extra["workspace"] = workspace_dir
        # pack may take the workspace from the previous extract rather than from
        # --workspace; record what was actually read so the choice is not implicit.
        if command == "pack":
            try:
                extra["workspace_snapshot"] = _input_snapshot(workspace_dir, "verify")
            except RuntimeError:
                pass
    _record_input(state, command, input_snapshot, extra)
    state["commands"][command]["resolved_config"] = effective
    state["commands"][command]["software_version"] = __version__
    state["commands"][command]["generated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    if command == "synthesize":
        model_label = (effective["model"] if effective["engine"] == "gpt_sovits"
                       else f"qwen3-{effective['qwen3_model_size']}")
        state["commands"][command]["audio_files"] = [
            os.path.basename(path) for path in sorted(
                os.path.join(workspace_dir, audio_filename(slide, effective["in_lang"], model_label))
                for slide in _slides_from_workspace(workspace_dir)
            ) if os.path.exists(path)
        ]
    state_dir = workspace_hint or workspace_dir
    _save_state(state, _workspace_path(state_dir, STATE_FILE))

    # Everything inside the workspace is recorded relative to it, so that the
    # workspace can be moved or copied without invalidating its own record.
    rel_snapshot = _relativize_snapshot(input_snapshot, state_dir)
    metadata = {
        "software_version": __version__,
        "generated_at": state["commands"][command]["generated_at"],
        "command": command,
        "workspace": os.path.abspath(state_dir) if state_dir else "",
        "paths_relative_to": "workspace",
        "config_file": _rel_to_workspace(config_path, state_dir) if config_path else "",
        "input_path": rel_snapshot.get("path", ""),
        "input_kind": rel_snapshot.get("kind", ""),
    }
    if rel_snapshot.get("kind") == "file":
        metadata["input_sha256"] = rel_snapshot.get("sha256", "")
    else:
        metadata["input_files_json"] = json.dumps(rel_snapshot.get("files", []), ensure_ascii=False, sort_keys=True)
    resolved = {"metadata": metadata, command: effective}
    _save_resolved_config(resolved, _workspace_path(state_dir, RESOLVED_CONFIG_FILE))
    _append_history(state_dir, {
        "command": command,
        "generated_at": metadata["generated_at"],
        "input": rel_snapshot,
        "effective": effective,
        "status": "success",
    })


def _prepare_file_workspace(command, input_path, in_lang=None):
    """Create an isolated one-file workspace so file INPUT never processes neighbors."""
    tmp = tempfile.mkdtemp(prefix="pptx_narrator_")
    name = os.path.basename(input_path)
    shutil.copy2(input_path, os.path.join(tmp, name))

    # Commands that need a paired artifact (verify) get only the corresponding
    # slide/model files from the original directory. Translation also gets an
    # existing target file and manifest so its normal overwrite/skip semantics
    # are preserved.
    if command == "verify":
        m = re.match(r"^slide_(\d+)(?:_[^.]+)?\.[^.]+$", name)
        if m:
            prefix = f"slide_{m.group(1)}_"
            for sibling in os.listdir(os.path.dirname(input_path)):
                if sibling.startswith(prefix) and os.path.isfile(os.path.join(os.path.dirname(input_path), sibling)):
                    if sibling != name:
                        shutil.copy2(os.path.join(os.path.dirname(input_path), sibling), os.path.join(tmp, sibling))
    elif command == "translate":
        m = _TEXT_FILE_RE.match(name)
        if m:
            prefix = f"slide_{m.group(1)}_"
            for sibling in os.listdir(os.path.dirname(input_path)):
                if sibling.startswith(prefix) and sibling != name and os.path.isfile(os.path.join(os.path.dirname(input_path), sibling)):
                    shutil.copy2(os.path.join(os.path.dirname(input_path), sibling), os.path.join(tmp, sibling))
        manifest = os.path.join(os.path.dirname(input_path), TRANSLATION_MANIFEST)
        if os.path.exists(manifest):
            shutil.copy2(manifest, os.path.join(tmp, TRANSLATION_MANIFEST))
    return tmp


def _sync_file_workspace(tmp, original_path, original_input_name):
    """Copy generated artifacts from an isolated file workspace back beside INPUT."""
    dest_dir = os.path.dirname(original_path) or "."
    for name in os.listdir(tmp):
        if name == original_input_name:
            continue
        src = os.path.join(tmp, name)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(dest_dir, name))


def _slides_from_workspace(workspace_dir):
    """Return slide numbers represented by language-tagged text/audio files in a workspace."""
    slides = set()
    if not os.path.isdir(workspace_dir):
        return slides
    for name in os.listdir(workspace_dir):
        m = _TEXT_FILE_RE.match(name)
        if m:
            slides.add(int(m.group(1)))
            continue
        m = re.match(r"^slide_(\d+)(?:_[^.]+)?\.[^.]+$", name)
        if m:
            slides.add(int(m.group(1)))
    return slides


if __name__ == "__main__":
    main()
