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
import xml.etree.ElementTree as ET
import zipfile
import csv
import json
import requests
from pptx import Presentation
from deep_translator import GoogleTranslator
from pydub import AudioSegment
import nltk
from pydub.effects import compress_dynamic_range, normalize

__version__ = "1.2.0"


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
# Workspace file names: Japanese keeps the v1.0 names (slide_N.txt), English keeps
# slide_N_eng.txt, and every other language uses slide_N_<lang>.txt.
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
    base = base_lang(lang)
    if base == "ja":
        return ""
    if base == "en":
        return "_eng"
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
def step_scan_and_update_dict(workspace_dir, dict_path, entries, requested_slides, target_lang,
                              source_lang="auto", for_translation=False):
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
            if term_lower in existing_terms or term_lower in stop_words:
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
            if term_lower in existing_terms:
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
# narration text before synthesis (e.g. readings). Translation and synthesis are run
# separately when they need different dictionaries. type 'unit' marks unit symbols
# that follow a number. The header row is optional. Files in the v1.x layout
# (Term,Japanese_Reading,English_Reading,Type) are read using the column of the
# narration language (ja or en).
DICT_HEADER = ["string", "replacement", "type"]
_HEADER_NAMES = {"string", "term", "replacement", "reading", "type"}


def read_dictionary_file(path, lang=None):
    """Return [(string, replacement, type)] from one dictionary CSV."""
    with open(path, 'r', encoding='utf-8', newline='') as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
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
    text = unicodedata.normalize('NFC', text)
    units, terms = {}, {}
    for term, repl, typ in entries or []:
        term, repl = unicodedata.normalize('NFC', term), unicodedata.normalize('NFC', repl)
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
        languages = {n: source_lang for n in plain}

    for slide_num, txt in plain.items():
        name = text_filename(slide_num, languages[slide_num])
        with open(os.path.join(workspace_dir, name), "w", encoding="utf-8") as f:
            f.write(txt)
        logger.info(f"Slide #{slide_num}: note extracted to {name} (language: {languages[slide_num]}).")

    for slide_num, info in structured.items():
        src_lang = info["source_lang"] if source_lang == "auto" else source_lang
        languages[slide_num] = src_lang
        src_name = text_filename(slide_num, src_lang)
        with open(os.path.join(workspace_dir, src_name), "w", encoding="utf-8") as f:
            f.write(info["source_text"])
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

def step_translate_notes(workspace_dir, requested_slides, source_lang, target_lang,
                         dictionary=None, overwrite=False):
    logger.info(f"--- [Option: Translate] Translating notes into '{target_lang}' ---")
    translators = {}
    for slide_num in requested_slides:
        tgt_p = os.path.join(workspace_dir, text_filename(slide_num, target_lang))
        if not overwrite and os.path.exists(tgt_p) and os.path.getsize(tgt_p) > 0:
            continue
        src_lang, src_p = find_source_text(workspace_dir, slide_num, source_lang, exclude_lang=target_lang)
        if src_p is None:
            continue
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
                translated_lines.append(translator.translate(line_in) or "")
            except Exception as e:
                logger.error(f"Slide {slide_num} translation error: {e}")
                translated_lines.append("")

        if any(translated_lines):
            with open(tgt_p, "w", encoding="utf-8") as out_f:
                out_f.write('\n'.join(translated_lines))
            record_translation(workspace_dir, slide_num, target_lang, src_lang, text)
            logger.info(f"Slide #{slide_num}: translated {src_lang} -> {target_lang}.")

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
                logger.info(f"Slide {slide_num}: Audio generated successfully.")
            else:
                logger.error(f"Slide {slide_num} TTS failed! Server returned [{res.status_code}]: {res.text}")
        except Exception as e:
            logger.error(f"Slide {slide_num} TTS connection error: {e}")

def split_into_chunks(text):
    """Split text into sentences: after 。！？, or after . ! ? followed by whitespace."""
    parts = re.split(r'(?<=[。！？])|(?<=[.!?])\s+', text)
    return [p.strip() for p in parts if p and p.strip()]

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
        try:
            wavs = []
            sr = None

            def _gen_chunk(chunk_text):
                if voice_clone_prompt is not None:
                    return model.generate_voice_clone(
                        text=chunk_text, language=language, voice_clone_prompt=voice_clone_prompt,
                    )
                return model.generate_voice_clone(
                    text=chunk_text, language=language, ref_audio=ref_wav, ref_text=ref_txt,
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
            logger.info(f"Slide {slide_num}: Audio generated successfully (qwen3, {len(chunks)} chunks).")
        except Exception as e:
            logger.error(f"Slide {slide_num} Qwen3-TTS generation error: {e}")

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

def text_scores(text_intended, text_asr):
    """Character-level similarity and CER on normalized text; return (similarity, CER, norm_intended, norm_asr)."""
    a = normalize_for_comparison(text_intended)
    b = normalize_for_comparison(text_asr)
    similarity = difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()
    return similarity, character_error_rate(a, b), a, b

def step_verify_audio(workspace_dir, requested_slides, lang, model_label,
                       asr_model_size="small", asr_device="cpu", threshold=0.85,
                       cer_threshold=None):
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

    results = []
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

        # Latin-script words left in Japanese narration cannot be compared reliably as kana
        has_latin = is_ja and bool(re.search(r'[A-Za-z]{2,}', intended_text))
        failed = score < threshold or (cer_threshold is not None and cer > cer_threshold)
        status = "ENGLISH" if has_latin else ("FLAGGED" if failed else "OK")

        results.append((slide_num, round(score, 4), round(cer, 4), status,
                        intended_text, asr_text, norm_intended, norm_asr))
        logger.info(f"Slide {slide_num}: similarity={score:.2f}, CER={cer:.2f} [{status}]")

    if not results:
        logger.info("No slides with both audio and spoken-text files found -- nothing to verify.")
        return

    report_p = os.path.join(workspace_dir, f"verify_report{lang_suffix(lang)}.{model_label}.csv")
    with open(report_p, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["slide", "similarity", "cer", "status", "intended_text", "asr_text",
                    "intended_normalized", "asr_normalized"])
        for row in sorted(results, key=lambda r: r[1]):
            w.writerow(row)

    n_flagged = sum(1 for r in results if r[3] == "FLAGGED")
    n_latin = sum(1 for r in results if r[3] == "ENGLISH")
    criterion = f"similarity < {threshold}"
    if cer_threshold is not None:
        criterion += f" or CER > {cer_threshold}"
    logger.info(f"Done: {n_flagged}/{len(results)} slide(s) flagged for review ({criterion}).")
    if n_latin:
        logger.info(f"{n_latin} slide(s) contain un-converted Latin-script words and need a listen.")
    logger.info(f"Report saved to: {report_p} (sorted worst-first)")

def _read_text(path):
    if path and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""

def step_pack_pptx(original_pptx, output_pptx, workspace_dir, requested_slides, lang, model_label,
                   source_lang="auto", writeback_notes=False, use_spoken_notes=False):
    logger.info("--- [Option: Pack] Rebuilding PPTX ---")
    prs = Presentation(original_pptx)
    manifest = _load_manifest(workspace_dir)
    for i, slide in enumerate(prs.slides):
        s_num = i + 1
        if s_num not in requested_slides:
            continue

        if writeback_notes:
            src_lang, src_p = find_source_text(workspace_dir, s_num, source_lang, exclude_lang=lang)
            text_p = os.path.join(workspace_dir, text_filename(s_num, lang))
            spoken_p = os.path.join(workspace_dir, spoken_filename(s_num, lang, model_label))
            narration = _read_text(spoken_p if use_spoken_notes else text_p)
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
                                                    fingerprint=fingerprint, spoken=use_spoken_notes)
            else:
                note_text = narration or original

            logger.info(f"Slide #{s_num}: Updating notes...")
            notes_slide = slide.notes_slide
            if notes_slide.notes_text_frame is not None:
                notes_slide.notes_text_frame.text = note_text
            else:
                logger.warning(f"Slide #{s_num}: No text frame in notes slide (skipped)")

    tmp_pptx = os.path.join(workspace_dir, "tmp.pptx")
    prs.save(tmp_pptx)
    rel_ns = '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(tmp_pptx, 'r') as z:
            z.extractall(tmpdir)
        for s_num in requested_slides:
            m4a_p = os.path.join(workspace_dir, audio_filename(s_num, lang, model_label))
            if not os.path.exists(m4a_p):
                continue
            rels_p = os.path.join(tmpdir, "ppt", "slides", "_rels", f"slide{s_num}.xml.rels")
            replaced = False
            if os.path.exists(rels_p):
                tree = ET.parse(rels_p)
                for rel in tree.getroot().findall(rel_ns):
                    target = rel.get('Target', '')
                    if not target.endswith(('.m4a', '.wav')):
                        continue
                    shutil.copy(m4a_p, os.path.join(tmpdir, "ppt", "media", os.path.basename(target)))
                    dur = len(AudioSegment.from_file(m4a_p))
                    xml_p = os.path.join(tmpdir, "ppt", "slides", f"slide{s_num}.xml")
                    with open(xml_p, "r", encoding="utf-8") as f:
                        xml_c = f.read()
                    if not re.search(r'advTm="\d+"', xml_c):
                        logger.warning(f"Slide #{s_num}: no automatic slide timing (advTm) found; display duration left unchanged.")
                    with open(xml_p, "w", encoding="utf-8") as f:
                        f.write(re.sub(r'advTm="\d+"', f'advTm="{dur}"', xml_c))
                    replaced = True
                    logger.info(f"Slide #{s_num}: audio replaced ({dur} ms).")
                    break
            if not replaced:
                logger.warning(
                    f"Slide #{s_num}: generated audio exists but the slide has no embedded audio "
                    "object to replace (skipped). Add a placeholder audio clip in PowerPoint first."
                )
        archive_path = shutil.make_archive(os.path.splitext(output_pptx)[0] + ".packing", 'zip', tmpdir)
        os.replace(archive_path, output_pptx)
    os.remove(tmp_pptx)
    logger.info(f"Final output saved to: {output_pptx}")

# ==========================================
# Main CLI
# ==========================================
DESCRIPTION = """\
PPTX-Narrator: automated narration of PowerPoint presenter notes
-----------------------------------------------------------------
[Narration in the language of the notes]
 1. pptx-narrator --pptx deck.pptx --workspace ws --target-lang ja --extract --scan --dict-file readings_ja.csv
 2. Review ws/slide_N.txt and readings_ja.csv by hand
 3. pptx-narrator --pptx deck.pptx --workspace ws --target-lang ja --dict-file readings_ja.csv \\
      --tts --verify --pack --out narrated.pptx --ref-wav ref.wav --ref-text-file ref.txt
[Translated narration] translation and synthesis are separate runs with their own dictionaries:
 1. pptx-narrator --pptx deck.pptx --workspace ws --target-lang de --extract --scan --dict-file terms_ja_de.csv
 2. (review terms_ja_de.csv) pptx-narrator ... --target-lang de --translate --dict-file terms_ja_de.csv
 3. pptx-narrator --pptx deck.pptx --workspace ws --target-lang de --scan --dict-file readings_de.csv
 4. (review readings_de.csv) pptx-narrator --pptx deck.pptx --workspace ws --target-lang de --dict-file readings_de.csv \\
      --tts --verify --pack --writeback-notes --out deck_de.pptx --ref-wav ref.wav --ref-text-file ref.txt
Options are written with hyphens; the underscore spellings of v1.0
(e.g. --dict_file, --verify_threshold) are still accepted.
"""


def _add(group, *names, **kwargs):
    """Register an option under its hyphenated name plus the legacy underscore alias."""
    flags = list(names)
    for name in names:
        legacy = name.replace("-", "_").replace("__", "--", 1)
        if legacy != name and legacy not in flags:
            flags.append(legacy)
    group.add_argument(*flags, **kwargs)


def build_parser():
    parser = argparse.ArgumentParser(
        prog="pptx-narrator",
        description=DESCRIPTION,
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    g_io = parser.add_argument_group("input / output")
    g_io.add_argument("--pptx", required=True, help="Path to the input PPTX file (required)")
    g_io.add_argument("--out", default="output.pptx", help="Output PPTX path for --pack (default: output.pptx)")
    g_io.add_argument("--workspace",
                      help="Workspace directory for intermediate files\n"
                           "(default: creates workspace_<filename>_<timestamp>)")
    g_io.add_argument("--slides", help="Slide range to process, e.g. '1-5' or '1,3,5-' (default: all)")

    g_steps = parser.add_argument_group("pipeline steps (combine as needed; executed in this order)")
    g_steps.add_argument("--extract", action="store_true",
                         help="Extract presenter notes (hidden slides are skipped)")
    g_steps.add_argument("--scan", action="store_true",
                         help="Scan for acronyms / technical terms / units and append new\n"
                              "candidates to --dict-file. Scans the notes when combined with\n"
                              "--translate, otherwise the narration text in --target-lang")
    g_steps.add_argument("--translate", action="store_true",
                         help="Translate the notes into --target-lang with Google Translate,\n"
                              "after applying --dict-file to the notes")
    g_steps.add_argument("--retranslate", action="store_true",
                         help="With --translate, overwrite existing translations")
    g_steps.add_argument("--tts", action="store_true", help="Synthesize narration audio with the selected engine")
    g_steps.add_argument("--verify", action="store_true",
                         help="ASR round-trip check: transcribe the audio with faster-whisper and\n"
                              "report similarity and character error rate (CER) against the\n"
                              "intended text (kana-level for Japanese, normalized characters otherwise)")
    g_steps.add_argument("--pack", action="store_true",
                         help="Replace the embedded audio of each slide and set slide timings")

    g_lang = parser.add_argument_group("languages")
    _add(g_lang, "--source-lang", dest="source_lang", type=normalize_lang, default="auto",
         help="Language of the presenter notes, e.g. ja, en, zh-CN, de (default: auto =\n"
              "identified per note from its text; short notes take the deck's main language)")
    _add(g_lang, "--target-lang", dest="target_lang", type=normalize_lang, default=None,
         help="Narration language. Any Google Translate language for --translate;\n"
              "for --tts it must be supported by the engine:\n"
              "  qwen3: " + ", ".join(QWEN3_LANGUAGES) + "\n"
              "  gpt_sovits: " + ", ".join(GPT_SOVITS_LANGUAGES) + "\n"
              "(default: --source-lang if given, otherwise en)")

    g_text = parser.add_argument_group("text normalization")
    _add(g_text, "--dict-file", dest="dict_file", action="append", default=None,
         help="Dictionary CSV of string replacements (string,replacement,type), repeatable.\n"
              "Applied to the notes before --translate, or to the narration text before --tts;\n"
              "run translation and synthesis separately to use different dictionaries")
    _add(g_text, "--letter-map", dest="letter_map",
         help="JSON mapping of letters to readings in the narration language, used for\n"
              "unknown unit symbols (e.g. examples/letter_map_ja.json)")

    g_tts = parser.add_argument_group("speech synthesis")
    g_tts.add_argument("--engine", choices=["gpt_sovits", "qwen3"], default="gpt_sovits",
                       help="TTS engine for --tts (default: gpt_sovits)")
    _add(g_tts, "--ref-wav", dest="ref_wav", help="Reference recording (.wav) of the voice to clone")
    _add(g_tts, "--ref-text-file", dest="ref_text_file", help="Text file containing the transcript of --ref-wav")
    _add(g_tts, "--ref-lang", dest="ref_lang", type=normalize_lang, default="ja",
         help="Language of the reference recording, GPT-SoVITS only (default: ja)")
    _add(g_tts, "--api-url", dest="api_url", default="http://127.0.0.1:9880/",
         help="GPT-SoVITS API server URL (default: http://127.0.0.1:9880/)")
    g_tts.add_argument("--model", default="v2ProPlus",
                       help="GPT-SoVITS model: " + ", ".join(MODELS_CONFIG) + " (default: v2ProPlus)")
    _add(g_tts, "--qwen3-model-size", dest="qwen3_model_size", choices=["0.6B", "1.7B"], default="1.7B",
         help="Qwen3-TTS model size (default: 1.7B)")
    _add(g_tts, "--qwen3-device", dest="qwen3_device", default="auto",
         help="Device for Qwen3-TTS: auto / cuda:0 / mps / cpu (default: auto)")
    _add(g_tts, "--enable-drc", dest="enable_drc", action="store_true",
         help="Apply dynamic range compression to avoid volume drop at sentence ends")
    _add(g_tts, "--drc-threshold", dest="drc_threshold", type=float, default=-20.0,
         help="DRC threshold in dBFS (default: -20.0)")
    _add(g_tts, "--drc-ratio", dest="drc_ratio", type=float, default=3.0,
         help="DRC ratio (default: 3.0)")

    g_ver = parser.add_argument_group("verification")
    _add(g_ver, "--asr-model", dest="asr_model", default="small",
         help="faster-whisper model size: tiny/base/small/medium/large-v3 (default: small)")
    _add(g_ver, "--asr-device", dest="asr_device", default="cpu", help="Device for the ASR model: cpu/cuda (default: cpu)")
    _add(g_ver, "--verify-threshold", dest="verify_threshold", type=float, default=0.85,
         help="Flag a slide when similarity (0-1) is below this value (default: 0.85)")
    _add(g_ver, "--cer-threshold", dest="cer_threshold", type=float, default=None,
         help="Additionally flag a slide when CER exceeds this value (default: not used)")

    g_pack = parser.add_argument_group("packing")
    _add(g_pack, "--writeback-notes", dest="writeback_notes", action="store_true",
         help="Write the narration back into the slide notes. Translated (or spoken-form)\n"
              "narration is written first and the original note is kept below it, separated\n"
              "by '=== pptx-narrator: ... ===' marker lines that --extract recognizes")
    _add(g_pack, "--use-spoken-notes", dest="use_spoken_notes", action="store_true",
         help="With --writeback-notes, write the dictionary-normalized reading text instead")
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    args.source_lang = args.source_lang or "auto"
    if args.target_lang is None:
        args.target_lang = args.source_lang if args.source_lang != "auto" else "en"
        logger.info(f"[Target Lang] --target-lang was not specified. Defaulting to '{args.target_lang}'.")
    else:
        logger.info(f"[Target Lang] Target language explicitly set to: '{args.target_lang}'")
    if args.target_lang == "auto":
        parser.error("--target-lang cannot be 'auto'")

    steps = [args.extract, args.scan, args.translate, args.tts, args.verify, args.pack]
    if not any(steps):
        parser.error("no pipeline step selected; use at least one of "
                     "--extract, --scan, --translate, --tts, --verify, --pack")
    if not os.path.exists(args.pptx):
        parser.error(f"input PPTX not found: {args.pptx}")
    if args.dict_file and args.translate and args.tts:
        parser.error("--dict-file would apply to both --translate and --tts; run translation and "
                     "synthesis separately, each with its own dictionary")
    if args.scan and not args.dict_file:
        parser.error("--scan needs --dict-file (the dictionary new candidates are added to)")
    if args.translate and args.source_lang != "auto" and lang_suffix(args.source_lang) == lang_suffix(args.target_lang):
        parser.error("--translate needs --target-lang to differ from --source-lang")
    if args.tts:
        missing = [flag for flag, value in (("--ref-wav", args.ref_wav),
                                            ("--ref-text-file", args.ref_text_file)) if not value]
        if missing:
            parser.error("--tts requires " + " and ".join(missing))
        if args.engine == "qwen3" and qwen3_language(args.target_lang) is None:
            parser.error(f"Qwen3-TTS does not support '{args.target_lang}' "
                         f"(supported: {', '.join(QWEN3_LANGUAGES)})")
        if args.engine == "gpt_sovits":
            for flag, code in (("--target-lang", args.target_lang), ("--ref-lang", args.ref_lang)):
                if gpt_sovits_language(code) is None:
                    parser.error(f"GPT-SoVITS does not support {flag} '{code}' "
                                 f"(supported: {', '.join(GPT_SOVITS_LANGUAGES)})")
    if args.engine == "gpt_sovits" and args.model not in MODELS_CONFIG:
        parser.error(f"unknown GPT-SoVITS model '{args.model}' (available: {', '.join(MODELS_CONFIG)})")

    if args.tts and args.engine == "gpt_sovits":
        config = MODELS_CONFIG[args.model]
        base_url = args.api_url.rstrip("/")
        logger.info(f"Switching GPT-SoVITS weights to {args.model}...")
        try:
            requests.get(f"{base_url}/set_gpt_weights", params={"weights_path": config["gpt"]}, timeout=300)
            requests.get(f"{base_url}/set_sovits_weights", params={"weights_path": config["sovits"]}, timeout=300)
        except requests.RequestException as e:
            logger.error(f"Could not reach the GPT-SoVITS API server at {args.api_url}: {e}")
            sys.exit(1)

    if args.workspace:
        workspace_dir = args.workspace
    else:
        base_name = os.path.splitext(os.path.basename(args.pptx))[0]
        workspace_dir = f"workspace_{base_name}_{int(time.time())}"
    os.makedirs(workspace_dir, exist_ok=True)

    prs = Presentation(args.pptx)
    req_slides = sorted(parse_slide_ranges(args.slides, len(prs.slides)))

    model_label = f"qwen3-{args.qwen3_model_size}" if args.engine == "qwen3" else args.model
    lang = args.target_lang
    letter_map_data = load_letter_map(args.letter_map)
    # With --translate the dictionary rewrites the notes; otherwise the narration text.
    dictionaries = load_dictionaries(args.dict_file, None if args.translate else lang)

    if args.extract:
        step_extract_notes(args.pptx, workspace_dir, req_slides, args.source_lang)
    if args.scan:
        step_scan_and_update_dict(workspace_dir, args.dict_file[0], dictionaries, req_slides, lang,
                                  source_lang=args.source_lang, for_translation=args.translate)
    if args.translate:
        step_translate_notes(workspace_dir, req_slides, args.source_lang, lang,
                             dictionary=dictionaries, overwrite=args.retranslate)
    if args.tts:
        if args.engine == "qwen3":
            step_generate_audio_qwen3(
                workspace_dir, req_slides, lang, args.ref_wav, args.ref_text_file,
                dictionaries, model_label, args.qwen3_model_size, args.qwen3_device,
                enable_drc=args.enable_drc, drc_threshold=args.drc_threshold, drc_ratio=args.drc_ratio,
                letter_map=letter_map_data,
            )
        else:
            step_generate_audio(
                workspace_dir, req_slides, lang, args.ref_wav, args.ref_text_file, args.ref_lang,
                args.api_url, dictionaries, model_label,
                enable_drc=args.enable_drc, drc_threshold=args.drc_threshold, drc_ratio=args.drc_ratio,
                letter_map=letter_map_data,
            )
    if args.verify:
        step_verify_audio(workspace_dir, req_slides, lang, model_label,
                          args.asr_model, args.asr_device, args.verify_threshold, args.cer_threshold)
    if args.pack:
        step_pack_pptx(
            args.pptx, args.out, workspace_dir, req_slides, lang, model_label,
            source_lang=args.source_lang,
            writeback_notes=args.writeback_notes, use_spoken_notes=args.use_spoken_notes,
        )


if __name__ == "__main__":
    main()
