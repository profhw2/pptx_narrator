#!/usr/bin/env python3
"""PPTX-Narrator: automated narration of PowerPoint presenter notes.

Pipeline: note extraction -> technical-term scanning -> dictionary / SI-unit
normalization -> (optional) ja->en translation -> voice-cloned TTS
(GPT-SoVITS or Qwen3-TTS) -> ASR round-trip verification -> PPTX repackaging.
"""
import os
import re
import unicodedata
import difflib
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

__version__ = "1.1.0"

DICT_HEADER = ["Term", "Japanese_Reading", "English_Reading", "Type"]


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
TXT_JA_TPL = "slide_{slide_num}.txt"
TXT_EN_TPL = "slide_{slide_num}_eng.txt"
AUDIO_JA_TPL = "slide_{slide_num}.{model_label}.m4a"
AUDIO_EN_TPL = "slide_{slide_num}_eng.{model_label}.m4a"
SPOKEN_TXT_JA_TPL = "slide_{slide_num}.{model_label}.spoken.txt"
SPOKEN_TXT_EN_TPL = "slide_{slide_num}_eng.{model_label}.spoken.txt"

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s', datefmt='%H:%M:%S')
logger = logging.getLogger(__name__)

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
def step_scan_and_update_dict(workspace_dir, dict_file, requested_slides):
    logger.info("--- [Option: Scan] Scanning text. Filtering out common English words. ---")
    
    existing_terms = set()
    if os.path.exists(dict_file):
        with open(dict_file, 'r', encoding='utf-8') as f:
            for row in csv.reader(f):
                if row: existing_terms.add(row[0].strip().lower())
    
    candidates = set()
    _ensure_nltk_data()
    from nltk.corpus import stopwords, words
    stop_words = set(stopwords.words('english'))
    common_english = set(w.lower() for w in words.words())
    
    for slide_num in requested_slides:
        for suffix in ["", "_eng"]:
            p = os.path.join(workspace_dir, f"slide_{slide_num}{suffix}.txt")
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8') as f:
                    text = f.read()
                    jp_char = r'[\u3040-\u30FF\u4E00-\u9FFF]'
                    generic_pattern = r'[A-Z]{2,}|[a-zA-Z]{3,}|[\u30A0-\u30FF]{3,}'
                    bypass_pattern = (
                        r'\d+(?:\.\d+)?\s*(?:%|℃|°C|nm|μm|mm|cm|km|kg|mg|µg|ng|pg|ml|μl|kb|Mb|Gb|bp|kDa|Å)'
                        r'|\b(?:I|II|III|IV|V|VI|VII|VIII|IX|X)\b'
                    )

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
                        if not term.isupper() and term_lower in common_english:
                            continue
                        candidates.add(term)
                        existing_terms.add(term_lower)

                    for term in bypass_found:
                        term_lower = term.lower()
                        if term_lower in existing_terms:
                            continue
                        candidates.add(term)
                        existing_terms.add(term_lower)

    new_entries = []
    translator = GoogleTranslator(source='en', target='ja')
    
    ROMAN_NUMERAL_READINGS = {
        "I": "いち", "II": "に", "III": "さん", "IV": "よん", "V": "ご",
        "VI": "ろく", "VII": "なな", "VIII": "はち", "IX": "きゅう", "X": "じゅう",
    }

    for term in sorted(candidates):
        ja_pron = ""
        en_pron = ""

        if term in ROMAN_NUMERAL_READINGS:
            ja_pron = ROMAN_NUMERAL_READINGS[term]
        elif re.match(r'^[A-Z]+$', term):
            ja_pron = " ".join(list(term))
        elif re.match(r'^[a-zA-Z]+$', term):
            try:
                guess = translator.translate(term)
                ja_pron = guess if (guess != term and is_japanese(guess)) else ""
            except Exception:
                ja_pron = ""
        elif is_japanese(term):
            ja_pron = term

        new_entries.append([term, ja_pron, en_pron, ""])
        logger.info(f"New technical term found: {term} -> {'(blank)' if not ja_pron else ja_pron}")

    if new_entries:
        is_new_file = not os.path.exists(dict_file) or os.path.getsize(dict_file) == 0
        needs_newline = False
        if not is_new_file:
            with open(dict_file, 'rb') as fb:
                fb.seek(-1, os.SEEK_END)
                needs_newline = fb.read(1) not in (b'\n', b'\r')
        with open(dict_file, 'a', encoding='utf-8', newline='') as f:
            if needs_newline:
                f.write('\n')
            writer = csv.writer(f)
            if is_new_file:
                writer.writerow(DICT_HEADER)
            writer.writerows(new_entries)
        logger.info(f"Added {len(new_entries)} purely technical terms.")

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

def _unit_reading(unit, extra_units):
    if unit in extra_units:
        return extra_units[unit]
    if unit in SPECIAL_UNITS:
        return SPECIAL_UNITS[unit]
    for split in range(1, len(unit)):
        prefix, base = unit[:split], unit[split:]
        if prefix in SI_PREFIXES and base in BASE_UNITS:
            return SI_PREFIXES[prefix] + BASE_UNITS[base]
    if unit in BASE_UNITS:
        return BASE_UNITS[unit]
    return None

def normalize_units(text, extra_units=None, letter_map=None):
    extra_units = extra_units or {}

    def _replace(m):
        number, letters = m.group(1), m.group(2)
        reading = _unit_reading(letters, extra_units)
        if reading is not None:
            return number + reading
        return number + spell_out_letters(letters, letter_map)

    pattern = r'(\d+(?:\.\d+)?)\s*([A-Za-zμµÅ°%℃]+)(?![A-Za-z0-9_])'
    return re.sub(pattern, _replace, text)

def apply_dictionary(text, dict_file, is_english, letter_map=None):
    text = unicodedata.normalize('NFC', text)
    dict_units = {}
    try:
        if dict_file and os.path.exists(dict_file):
            with open(dict_file, mode='r', encoding='utf-8') as f:
                all_rows = [
                    row for row in csv.reader(f)
                    if len(row) >= 3 and row[0].strip() and row[0].strip() != DICT_HEADER[0]
                ]

            rows = []
            for row in all_rows:
                term = unicodedata.normalize('NFC', row[0].strip())
                target = unicodedata.normalize('NFC', row[2 if is_english else 1].strip())
                if not target:
                    continue
                if len(row) >= 4 and row[3].strip().lower() == "unit":
                    dict_units[term] = target
                else:
                    rows.append((term, target))

            rows.sort(key=lambda x: len(x[0]), reverse=True)
            
            for term, target in rows:
                if re.match(r'^[a-zA-Z0-9_ \-]+$', term):
                    pattern = rf'(?<![A-Za-z0-9_]){re.escape(term)}(?![A-Za-z0-9_])'
                    text = re.sub(pattern, target, text)
                else:
                    text = text.replace(term, target)

    except Exception as e:
        logger.error(f"Dictionary error: {e}")

    if not is_english:
        text = normalize_units(text, dict_units, letter_map=letter_map)

    return text

# ==========================================
# Pipeline Steps
# ==========================================
def step_extract_notes(pptx_path, workspace_dir, requested_slides):
    logger.info("--- [Option: Extract] Extracting Notes ---")
    prs = Presentation(pptx_path)
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

        if raw_text:
            print(f"[INFO] Note of slide #{slide_num} has been extracted.")
        else:
            print(f"[INFO] Slide #{slide_num} has no notes (skipped)")

        clean_text = raw_text.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
        txt = re.sub(r'\d{4}/\d+/\d+', '', clean_text).strip()
        
        if txt:
            name = TXT_JA_TPL if is_japanese(txt) else TXT_EN_TPL
            with open(os.path.join(workspace_dir, name.format(slide_num=slide_num)), "w", encoding="utf-8") as f:
                f.write(txt)
            logger.info(f"Note of slide #{slide_num} has been extracted to {txt}.")

def step_ensure_english_translation(workspace_dir, requested_slides):
    logger.info("--- [Option: English] Translating to English ---")
    translator = GoogleTranslator(source='ja', target='en')
    for slide_num in requested_slides:
        ja_p = os.path.join(workspace_dir, TXT_JA_TPL.format(slide_num=slide_num))
        en_p = os.path.join(workspace_dir, TXT_EN_TPL.format(slide_num=slide_num))
        
        if os.path.exists(ja_p) and (not os.path.exists(en_p) or os.path.getsize(en_p) == 0):
            with open(ja_p, "r", encoding="utf-8") as f:
                text = f.read().strip()
                
            if not text:
                continue
                
            translated_lines = []
            for line in text.split('\n'):
                if line.strip():
                    try:
                        translated_lines.append(translator.translate(line.strip()))
                    except Exception as e:
                        logger.error(f"Slide {slide_num} translation error: {e}")
                        translated_lines.append("")
                else:
                    translated_lines.append("")
                    
            if any(translated_lines):
                with open(en_p, "w", encoding="utf-8") as out:
                    out.write('\n'.join(translated_lines))
                logger.info(f"Note of slide #{slide_num} has been translated.")

def step_generate_audio(
    workspace_dir,
    requested_slides,
    is_english,
    ref_wav,
    ref_text_f,
    ref_lang,
    api_url,
    dict_file,
    model_label,
    enable_drc=False,
    drc_threshold=-20.0,
    drc_ratio=3.0,
    letter_map=None,
):
    logger.info("--- [Option: TTS] Generating Audio ---")
    with open(ref_text_f, "r", encoding="utf-8") as f:
        ref_txt = f.read().strip()
    tts_url = api_url.rstrip("/") + "/tts"

    for slide_num in requested_slides:
        txt_n = (TXT_EN_TPL if is_english else TXT_JA_TPL).format(
            slide_num=slide_num, model_label=model_label
        )
        m4a_n = (AUDIO_EN_TPL if is_english else AUDIO_JA_TPL).format(
            slide_num=slide_num, model_label=model_label
        )
        txt_p = os.path.join(workspace_dir, txt_n)
        if not os.path.exists(txt_p):
            continue

        with open(txt_p, "r", encoding="utf-8") as f:
            spoken_text = apply_dictionary(f.read().strip(), dict_file, is_english, letter_map=letter_map)

        if not spoken_text:
            continue

        spoken_n = (SPOKEN_TXT_EN_TPL if is_english else SPOKEN_TXT_JA_TPL).format(
            slide_num=slide_num, model_label=model_label
        )
        with open(
            os.path.join(workspace_dir, spoken_n), "w", encoding="utf-8"
        ) as f:
            f.write(spoken_text)

        payload = {
            "text": spoken_text,
            "text_lang": "en" if is_english else "ja",
            "ref_audio_path": os.path.abspath(ref_wav),
            "prompt_text": ref_txt,
            "prompt_lang": "ja" if ref_lang.lower() in ("ja", "japanese") else "en",
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
                    audio = compress_dynamic_range(
                        audio, threshold=drc_threshold, ratio=drc_ratio
                    )
                    audio = normalize(audio)
                audio.export(os.path.join(workspace_dir, m4a_n), format="ipod")

                os.remove(wav_p)
                logger.info(
                    f"Slide {slide_num}: Audio generated successfully."
                )
            else:
                logger.error(
                    f"Slide {slide_num} TTS failed! Server returned"
                    f" [{res.status_code}]: {res.text}"
                )
        except Exception as e:
            logger.error(f"Slide {slide_num} TTS connection error: {e}")

def split_into_chunks(text):
    parts = re.split(r'(?<=[。！？])', text)
    return [p.strip() for p in parts if p.strip()]

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
    is_english,
    ref_wav,
    ref_text_f,
    dict_file,
    model_label,
    qwen3_model_size,
    qwen3_device,
    enable_drc=False,
    drc_threshold=-20.0,
    drc_ratio=3.0,
    letter_map=None,
):
    logger.info("--- [Option: TTS / engine=qwen3] Generating Audio ---")
    try:
        import numpy as np
        import soundfile as sf
        from pydub import AudioSegment
    except ImportError as e:
        logger.error(
            f"Missing required libraries: {e} (pip install -U qwen-tts torch"
            " soundfile numpy --break-system-packages)"
        )
        return

    with open(ref_text_f, "r", encoding="utf-8") as f:
        ref_txt = f.read().strip()

    model = _load_qwen3_model(qwen3_model_size, qwen3_device)

    voice_clone_prompt = None
    try:
        logger.info("Analyzing reference audio (create_voice_clone_prompt)...")
        voice_clone_prompt = model.create_voice_clone_prompt(
            ref_audio=ref_wav, ref_text=ref_txt
        )
    except AttributeError:
        logger.warning(
            "This qwen_tts version lacks create_voice_clone_prompt; falling"
            " back to passing ref_audio/ref_text per invocation."
        )

    for slide_num in requested_slides:
        txt_n = (TXT_EN_TPL if is_english else TXT_JA_TPL).format(
            slide_num=slide_num
        )
        m4a_n = (AUDIO_EN_TPL if is_english else AUDIO_JA_TPL).format(
            slide_num=slide_num, model_label=model_label
        )
        txt_p = os.path.join(workspace_dir, txt_n)
        if not os.path.exists(txt_p):
            continue

        with open(txt_p, "r", encoding="utf-8") as f:
            spoken_text = apply_dictionary(f.read().strip(), dict_file, is_english, letter_map=letter_map)
        if not spoken_text:
            continue

        spoken_n = (SPOKEN_TXT_EN_TPL if is_english else SPOKEN_TXT_JA_TPL).format(
            slide_num=slide_num, model_label=model_label
        )
        with open(
            os.path.join(workspace_dir, spoken_n), "w", encoding="utf-8"
        ) as f:
            f.write(spoken_text)

        chunks = split_into_chunks(spoken_text)
        try:
            wavs = []
            sr = None

            def _gen_chunk(chunk_text):
                if voice_clone_prompt is not None:
                    return model.generate_voice_clone(
                        text=chunk_text,
                        language="English" if is_english else "Japanese",
                        voice_clone_prompt=voice_clone_prompt,
                    )
                else:
                    return model.generate_voice_clone(
                        text=chunk_text,
                        language="English" if is_english else "Japanese",
                        ref_audio=ref_wav,
                        ref_text=ref_txt,
                    )

            def _is_anomalous(chunk_text, wav, sample_rate):
                expected_sec = max(len(chunk_text) / 6.0, 1.0)
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
                audio = compress_dynamic_range(
                    audio, threshold=drc_threshold, ratio=drc_ratio
                )
                audio = normalize(audio)
            audio.export(os.path.join(workspace_dir, m4a_n), format="ipod")

            os.remove(wav_p)
            logger.info(
                f"Slide {slide_num}: Audio generated successfully (qwen3,"
                f" {len(chunks)} chunks)."
            )
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

def step_verify_audio(workspace_dir, requested_slides, is_english, model_label,
                       asr_model_size="small", asr_device="cpu", threshold=0.85,
                       cer_threshold=None):
    logger.info("--- [Option: Verify] ASR round-trip check ---")
    try:
        import pyopenjtalk  # noqa: F401
        from faster_whisper import WhisperModel
    except ImportError as e:
        logger.error(f"Missing dependency: {e} (pip install pyopenjtalk faster-whisper --break-system-packages)")
        return

    if is_english:
        logger.error("--verify currently only supports Japanese (kana-level comparison). Skipping.")
        return

    logger.info(f"Loading Whisper model ({asr_model_size}, device={asr_device})...")
    asr_model = WhisperModel(asr_model_size, device=asr_device, compute_type="int8")

    results = []
    for slide_num in requested_slides:
        audio_n = AUDIO_JA_TPL.format(slide_num=slide_num, model_label=model_label)
        spoken_n = SPOKEN_TXT_JA_TPL.format(slide_num=slide_num, model_label=model_label)
        audio_p = os.path.join(workspace_dir, audio_n)
        spoken_p = os.path.join(workspace_dir, spoken_n)
        if not (os.path.exists(audio_p) and os.path.exists(spoken_p)):
            continue

        with open(spoken_p, "r", encoding="utf-8") as f:
            intended_text = f.read().strip()
        if not intended_text:
            continue

        segments, _ = asr_model.transcribe(audio_p, language="ja")
        asr_text = "".join(seg.text for seg in segments).strip()

        try:
            score, cer, kana_intended, kana_asr = kana_scores(intended_text, asr_text)
        except Exception as e:
            logger.error(f"Slide {slide_num}: kana comparison failed: {e}")
            continue

        has_english = bool(re.search(r'[A-Za-z]{2,}', intended_text))
        failed = score < threshold or (cer_threshold is not None and cer > cer_threshold)
        status = "ENGLISH" if has_english else ("FLAGGED" if failed else "OK")

        results.append((slide_num, round(score, 4), round(cer, 4), status,
                        intended_text, asr_text, kana_intended, kana_asr))
        logger.info(f"Slide {slide_num}: similarity={score:.2f}, CER={cer:.2f} [{status}]")

    if not results:
        logger.info("No slides with both audio and spoken-text files found -- nothing to verify.")
        return

    report_p = os.path.join(workspace_dir, f"verify_report.{model_label}.csv")
    with open(report_p, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["slide", "similarity", "cer", "status", "intended_text", "asr_text", "intended_kana", "asr_kana"])
        for row in sorted(results, key=lambda r: r[1]):
            w.writerow(row)

    n_flagged = sum(1 for r in results if r[3] == "FLAGGED")
    n_english = sum(1 for r in results if r[3] == "ENGLISH")
    criterion = f"similarity < {threshold}"
    if cer_threshold is not None:
        criterion += f" or CER > {cer_threshold}"
    logger.info(f"Done: {n_flagged}/{len(results)} slide(s) flagged for review ({criterion}).")
    if n_english:
        logger.info(f"{n_english} slide(s) contain un-converted English and need a listen.")
    logger.info(f"Report saved to: {report_p} (sorted worst-first)")

def step_pack_pptx(original_pptx, output_pptx, workspace_dir, requested_slides, is_english, model_label, writeback_notes=False, use_spoken_notes=False):
    logger.info("--- [Option: Pack] Rebuilding PPTX ---")
    prs = Presentation(original_pptx)
    for i, slide in enumerate(prs.slides):
        s_num = i + 1
        if s_num not in requested_slides: continue

        if writeback_notes:
            if use_spoken_notes:
                en_p = os.path.join(workspace_dir, SPOKEN_TXT_EN_TPL.format(slide_num=s_num, model_label=model_label))
                ja_p = os.path.join(workspace_dir, SPOKEN_TXT_JA_TPL.format(slide_num=s_num, model_label=model_label))
            else:
                en_p = os.path.join(workspace_dir, TXT_EN_TPL.format(slide_num=s_num))
                ja_p = os.path.join(workspace_dir, TXT_JA_TPL.format(slide_num=s_num))

            en_t = open(en_p, "r", encoding="utf-8").read().strip() if os.path.exists(en_p) else ""
            ja_t = open(ja_p, "r", encoding="utf-8").read().strip() if os.path.exists(ja_p) else ""

            if en_t or ja_t:
                logger.info(f"Slide #{s_num}: Updating notes...")
                notes_slide = slide.notes_slide
                if notes_slide.notes_text_frame is not None:
                    notes_slide.notes_text_frame.text = f"{en_t}\n\n{ja_t}" if (is_english and en_t and ja_t) else (en_t or ja_t)
                else:
                    logger.warning(f"Slide #{s_num}: No text frame in notes slide (skipped)")

    tmp_pptx = os.path.join(workspace_dir, "tmp.pptx")
    prs.save(tmp_pptx)
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(tmp_pptx, 'r') as z: z.extractall(tmpdir)
        for s_num in requested_slides:
            m4a_n = (AUDIO_EN_TPL if is_english else AUDIO_JA_TPL).format(slide_num=s_num, model_label=model_label)
            m4a_p = os.path.join(workspace_dir, m4a_n)
            if os.path.exists(m4a_p):
                rels_p = os.path.join(tmpdir, "ppt", "slides", "_rels", f"slide{s_num}.xml.rels")
                replaced = False
                if os.path.exists(rels_p):
                    tree = ET.parse(rels_p)
                    for rel in tree.getroot().findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        if rel.get('Target').endswith(('.m4a', '.wav')):
                            shutil.copy(m4a_p, os.path.join(tmpdir, "ppt", "media", os.path.basename(rel.get('Target'))))
                            dur = len(AudioSegment.from_file(m4a_p))
                            xml_p = os.path.join(tmpdir, "ppt", "slides", f"slide{s_num}.xml")
                            with open(xml_p, "r", encoding="utf-8") as f: xml_c = f.read()
                            if not re.search(r'advTm="\d+"', xml_c):
                                logger.warning(f"Slide #{s_num}: no automatic slide timing (advTm) found; display duration left unchanged.")
                            with open(xml_p, "w", encoding="utf-8") as f: f.write(re.sub(r'advTm="\d+"', f'advTm="{dur}"', xml_c))
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
[Recommended workflow]
 1. Extract & scan : pptx-narrator --pptx deck.pptx --workspace ws --extract --scan
 2. Review         : edit ws/slide_N.txt and dict.csv by hand
 3. Synthesize     : pptx-narrator --pptx deck.pptx --workspace ws --target-lang ja \\
                       --tts --verify --ref-wav ref.wav --ref-text-file ref.txt
 4. Pack           : pptx-narrator --pptx deck.pptx --workspace ws --target-lang ja \\
                       --pack --out narrated.pptx
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
                         help="Scan notes for acronyms / technical terms / units and append\n"
                              "new candidates to the dictionary CSV")
    g_steps.add_argument("--translate", action="store_true",
                         help="Translate Japanese notes (slide_N.txt) into English\n"
                              "(slide_N_eng.txt) with Google Translate")
    g_steps.add_argument("--tts", action="store_true", help="Synthesize narration audio with the selected engine")
    g_steps.add_argument("--verify", action="store_true",
                         help="ASR round-trip check (Japanese only): transcribe the audio with\n"
                              "faster-whisper, convert both texts to kana with pyopenjtalk and\n"
                              "report similarity and character error rate (CER)")
    g_steps.add_argument("--pack", action="store_true",
                         help="Replace the embedded audio of each slide and set slide timings")

    g_text = parser.add_argument_group("text normalization")
    _add(g_text, "--target-lang", dest="target_lang", type=str.lower, choices=["ja", "en"], default=None,
         help="Narration language: 'ja' uses slide_N.txt, 'en' uses slide_N_eng.txt (default: en)")
    _add(g_text, "--dict-file", dest="dict_file", default="dict.csv",
         help="Pronunciation dictionary CSV (default: dict.csv)")
    _add(g_text, "--letter-map", dest="letter_map",
         help="JSON mapping of letters to readings used for unknown unit symbols\n"
              "(e.g. examples/letter_map_ja.json)")

    g_tts = parser.add_argument_group("speech synthesis")
    g_tts.add_argument("--engine", choices=["gpt_sovits", "qwen3"], default="gpt_sovits",
                       help="TTS engine for --tts (default: gpt_sovits)")
    _add(g_tts, "--ref-wav", dest="ref_wav", help="Reference recording (.wav) of the voice to clone")
    _add(g_tts, "--ref-text-file", dest="ref_text_file", help="Text file containing the transcript of --ref-wav")
    _add(g_tts, "--ref-lang", dest="ref_lang", type=str.lower, choices=["ja", "en"], default="ja",
         help="Language of the reference recording (GPT-SoVITS only, default: ja)")
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
         help="Flag a slide when kana similarity (0-1) is below this value (default: 0.85)")
    _add(g_ver, "--cer-threshold", dest="cer_threshold", type=float, default=None,
         help="Additionally flag a slide when kana CER exceeds this value (default: not used)")

    g_pack = parser.add_argument_group("packing")
    _add(g_pack, "--writeback-notes", dest="writeback_notes", action="store_true",
         help="Write the (edited) slide_N.txt files back into the slide notes")
    _add(g_pack, "--use-spoken-notes", dest="use_spoken_notes", action="store_true",
         help="With --writeback-notes, write the dictionary-normalized reading text instead")
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.target_lang is None:
        args.target_lang = "en"
        logger.info("[Target Lang] --target-lang was not specified. Defaulting to 'en'.")
    else:
        logger.info(f"[Target Lang] Target language explicitly set to: '{args.target_lang}'")

    steps = [args.extract, args.scan, args.translate, args.tts, args.verify, args.pack]
    if not any(steps):
        parser.error("no pipeline step selected; use at least one of "
                     "--extract, --scan, --translate, --tts, --verify, --pack")
    if not os.path.exists(args.pptx):
        parser.error(f"input PPTX not found: {args.pptx}")
    if args.tts:
        missing = [flag for flag, value in (("--ref-wav", args.ref_wav),
                                            ("--ref-text-file", args.ref_text_file)) if not value]
        if missing:
            parser.error("--tts requires " + " and ".join(missing))
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
    is_english = (args.target_lang == "en")
    letter_map_data = load_letter_map(args.letter_map)

    if args.extract:
        step_extract_notes(args.pptx, workspace_dir, req_slides)
    if args.scan:
        step_scan_and_update_dict(workspace_dir, args.dict_file, req_slides)
    if args.translate:
        step_ensure_english_translation(workspace_dir, req_slides)
    if args.tts:
        if args.engine == "qwen3":
            step_generate_audio_qwen3(
                workspace_dir, req_slides, is_english, args.ref_wav, args.ref_text_file,
                args.dict_file, model_label, args.qwen3_model_size, args.qwen3_device,
                enable_drc=args.enable_drc, drc_threshold=args.drc_threshold, drc_ratio=args.drc_ratio,
                letter_map=letter_map_data,
            )
        else:
            step_generate_audio(
                workspace_dir, req_slides, is_english, args.ref_wav, args.ref_text_file, args.ref_lang,
                args.api_url, args.dict_file, model_label,
                enable_drc=args.enable_drc, drc_threshold=args.drc_threshold, drc_ratio=args.drc_ratio,
                letter_map=letter_map_data,
            )
    if args.verify:
        step_verify_audio(workspace_dir, req_slides, is_english, model_label,
                          args.asr_model, args.asr_device, args.verify_threshold, args.cer_threshold)
    if args.pack:
        step_pack_pptx(
            args.pptx, args.out, workspace_dir, req_slides, is_english, model_label,
            writeback_notes=args.writeback_notes, use_spoken_notes=args.use_spoken_notes,
        )


if __name__ == "__main__":
    main()
