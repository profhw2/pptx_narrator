# PPTX-Narrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22812409.svg)](https://doi.org/10.5281/zenodo.22812409)

**PPTX-Narrator** turns the presenter notes of a PowerPoint deck into narration spoken in a cloned voice, in the language of the notes or translated into another language, and writes the audio back into the deck. It is designed for technical and scientific lectures that are revised frequently: when a note changes, the corresponding narration is regenerated instead of re-recorded.

## Features

- **Note extraction** – presenter notes are exported per slide as editable text files; the language of each note is identified automatically; hidden slides are skipped.
- **Translation** – notes can be translated into any language supported by Google Translate (via `deep-translator`), with a per-language-pair glossary for technical terms.
- **Technical-term scanning** – acronyms, domain terms and number–unit expressions are collected into dictionaries for manual review.
- **Reading normalization** – per-language rewrite dictionaries, plus built-in reading of SI-prefixed units for Japanese (e.g. `5 mg` → 5ミリグラム).
- **Voice cloning** – [Qwen3-TTS](https://github.com/QwenLM/Qwen3-TTS) (in-process) or [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) (via its API server).
- **ASR round-trip verification** – audio is transcribed with `faster-whisper` and compared with the intended text (kana level via `pyopenjtalk` for Japanese, normalized characters otherwise); similarity and character error rate (CER) are reported per slide.
- **PPTX repackaging** – the embedded audio of each slide is replaced and the automatic slide advance time (`advTm`) is set to the audio duration. Translated narration can be written into the notes above the original note, in a marked layout that later extractions recognize.

## Supported languages

| Step | Languages |
|---|---|
| Note language identification | Kana/Hangul rules and [py3langid](https://github.com/adbar/py3langid), restricted to Google Translate languages |
| `--translate` | Any source/target pair supported by Google Translate |
| `--tts --engine qwen3` | zh, en, ja, ko, de, fr, ru, pt, es, it |
| `--tts --engine gpt_sovits` | zh, en, ja, ko, yue (Cantonese) |
| `--verify` | Japanese: kana-level comparison; other languages: character-level comparison of normalized text |
| Built-in unit readings | Japanese only (other languages: `unit` entries in the dictionaries) |

Languages are written as Google Translate codes such as `ja`, `en`, `de`, `zh-CN`. With `--source-lang auto` (default), the language of each note is identified from its text; notes too short or ambiguous to identify reliably ("Thank you.", a kanji-only title) are assigned the main language of the deck. The detected language is shown in the log and in the file name; give `--source-lang` to override it.

## Requirements

- Python 3.10 or later
- [FFmpeg](https://ffmpeg.org/) on `PATH` (used by `pydub` to write `.m4a`)
- One TTS engine:
  - **Qwen3-TTS**: installed with the `qwen3` extra below. Model weights (`Qwen/Qwen3-TTS-12Hz-{0.6B,1.7B}-Base`) are downloaded from Hugging Face on first use. A CUDA GPU or Apple Silicon (MPS) is strongly recommended.
  - **GPT-SoVITS**: a local GPT-SoVITS installation with its API server running, e.g. `python api_v2.py -a 127.0.0.1 -p 9880` in the GPT-SoVITS directory.
- Network access for `--translate` (Google Translate).

## Installation

```bash
git clone https://github.com/profhw2/pptx_narrator.git
cd pptx_narrator
pip install -e .            # core pipeline (GPT-SoVITS engine, translation)
pip install -e ".[qwen3]"   # + Qwen3-TTS engine
pip install -e ".[verify]"  # + ASR verification
pip install -e ".[all]"     # everything
```

This installs the `pptx-narrator` command. Running `python pptx_narrator.py ...` without installing also works (`pip install -r requirements.txt`).
The `--scan` step downloads the NLTK `stopwords` and `words` corpora on first use; if that fails behind a proxy, run `python -m nltk.downloader stopwords words`.

## Quick start

The pipeline is split into steps so that text can be reviewed before synthesis. Intermediate files live in a workspace directory; dictionaries live in the current directory (or `--dict-dir`).

### Narration in the language of the notes

```bash
# 1. Extract notes and collect candidate technical terms into dict_ja_ja.csv
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja --extract --scan

# 2. Review by hand: ws/slide_N.txt (notes) and dict_ja_ja.csv (rewrites)

# 3. Synthesize with a cloned voice and verify
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja \
  --tts --engine qwen3 --ref-wav my_voice.wav --ref-text-file my_voice.txt \
  --verify --cer-threshold 0.15

# 4. Listen to the slides flagged in ws/verify_report.qwen3-1.7B.csv, fix the
#    dictionary or notes, re-run step 3 for those slides (--slides 4,7), then pack
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja --engine qwen3 \
  --pack --out lecture_narrated.pptx
```

### Translated narration

```bash
# 1. Extract the (e.g. Japanese) notes; terms found in them go to the glossary dict_ja_de.csv
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de --extract --scan
# 2. Fill in German translations of technical terms in dict_ja_de.csv, then translate;
#    scanning the German text adds rewrite candidates to dict_de_de.csv
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de --translate --scan
# 3. Review ws/slide_N_de.txt and dict_de_de.csv, then synthesize, verify and pack,
#    writing the German narration above the original note
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de \
  --tts --engine qwen3 --ref-wav my_voice.wav --ref-text-file my_voice.txt \
  --verify --pack --writeback-notes --out lecture_de.pptx
```

Existing translations are not overwritten; use `--retranslate` after changing the glossary.

### Workspace files

| File | Content |
|---|---|
| `slide_N.txt`, `slide_N_eng.txt`, `slide_N_<lang>.txt` | Note text in Japanese, English, or another language (e.g. `slide_3_de.txt`) |
| `slide_N<suffix>.<model>.spoken.txt` | Rewritten text actually sent to the TTS engine |
| `slide_N<suffix>.<model>.m4a` | Generated audio |
| `verify_report<suffix>.<model>.csv` | Verification report |
| `translations.json` | Which version of each source note a translation was made from |
| `slide_N<suffix>.stale.txt` | Previous narration set aside because the source note changed |

### Reference voice

`--ref-wav` is a short, clean recording of the target speaker (a few seconds to about ten seconds) and `--ref-text-file` contains its exact transcript. For GPT-SoVITS, give its language with `--ref-lang` (default `ja`).

## Dictionaries

Each dictionary is a CSV file for one language pair. The header names the source and target language, optionally followed by a `type` column; each row maps a string in the source language to a string in the target language:

```csv
ja,ja,type
CRISPR-Cas9,クリスパーキャスナイン,
mRNA,メッセンジャーアールエヌエー,
Gbp,ギガベースペア,unit
```

```csv
ja,de,type
塩基対,Basenpaare,
```

- **Same language (`ja,ja`, `de,de`, …)**: rewrites applied to the narration text right before synthesis, e.g. readings of acronyms and units. `type` = `unit` marks unit symbols that follow a number (`3 Gbp`).
- **Different languages (`ja,de`, …)**: glossary used by `--translate`. The source terms are replaced by placeholders before machine translation and by the given target strings afterwards. If the translator alters a placeholder, the line is translated without the glossary and a warning is shown.

Dictionaries are loaded from `dict_<source>_<target>.csv` files in `--dict-dir` (default: current directory) and from any files given with `--dict-file` (repeatable); the header, not the file name, determines the pair. `--scan` appends new candidates to the file of the corresponding pair and creates `dict_<source>_<target>.csv` if none exists: terms found in notes of language L go to `L,<target-lang>`. Rewrites are applied longest first; alphanumeric terms only match on word boundaries; rows with an empty target string are ignored. See the files in [`examples/`](examples/) (`--dict-dir examples`).

A v1.x `dict.csv` (`Term,Japanese_Reading,English_Reading,Type`) in `--dict-dir` is still read, as the pairs `ja,ja` and `en,en`, but not modified. Unit symbols not resolved otherwise can be spelled out letter by letter with a letter map for the narration language, e.g. `--letter-map examples/letter_map_ja.json`.

## Packing and notes write-back

`--pack` *replaces* audio that is already embedded in a slide; it does not insert new audio objects. Prepare the deck once, e.g. by recording a slide show narration in PowerPoint or inserting any audio clip on each slide to be narrated. To have slides advance automatically, set the slide transition to advance *After* a time; that time is overwritten with the length of the generated audio. Slides without embedded audio are reported and left unchanged.

With `--writeback-notes`, the narration is written into the notes. When it differs from the original note (translated narration, or the rewritten text with `--use-spoken-notes`), the note keeps both parts:

```
=== pptx-narrator: narration [en] from [ja] #ae82d4f1fc ===
Today we talk about genome editing.

=== pptx-narrator: source [ja] ===
今日はゲノム編集について話します。
```

Keep the marker lines when editing such notes in PowerPoint. When `--extract` finds this layout, only the source part is used as the note to translate. The narration part is restored as the existing translation, including edits made in PowerPoint, as long as the source part is unchanged (the hash identifies the translated version). If the source part was edited, the old narration is set aside as `slide_N<suffix>.stale.txt` and `--translate` produces a new translation. Spoken-form narration (`spoken` in the marker) is never restored.

## Command-line reference

| Option | Description |
|---|---|
| `--pptx` | Input deck (required) |
| `--workspace` | Workspace directory (default: `workspace_<name>_<timestamp>`) |
| `--slides` | Slide selection, e.g. `1-5`, `1,3,5-` |
| `--extract` / `--translate` / `--scan` / `--tts` / `--verify` / `--pack` | Pipeline steps, executed in this order |
| `--retranslate` | With `--translate`, overwrite existing translations |
| `--source-lang` | Language of the notes (default: `auto`) |
| `--target-lang` | Narration language (default: `--source-lang` if given, otherwise `en`) |
| `--dict-dir` | Directory of `dict_<source>_<target>.csv` files (default: `.`) |
| `--dict-file` | Additional dictionary file (repeatable) |
| `--letter-map` | JSON letter-reading map |
| `--engine {gpt_sovits,qwen3}` | TTS engine (default: `gpt_sovits`) |
| `--ref-wav`, `--ref-text-file` | Reference recording and its transcript (required with `--tts`) |
| `--ref-lang` | Language of the reference recording, GPT-SoVITS only (default: `ja`) |
| `--api-url`, `--model` | GPT-SoVITS server URL and model (`v2ProPlus`, `v4`, `v1_clear`) |
| `--qwen3-model-size {0.6B,1.7B}`, `--qwen3-device` | Qwen3-TTS model and device (`auto`, `cuda:0`, `mps`, `cpu`) |
| `--enable-drc`, `--drc-threshold`, `--drc-ratio` | Dynamic range compression of the output |
| `--asr-model`, `--asr-device` | faster-whisper model size and device |
| `--verify-threshold` | Flag slides with similarity below this value (default: 0.85) |
| `--cer-threshold` | Also flag slides with CER above this value (default: off) |
| `--out` | Output deck for `--pack` (default: `output.pptx`) |
| `--writeback-notes`, `--use-spoken-notes` | Write the narration (or its rewritten form) into the notes on `--pack` |

Run `pptx-narrator --help` for details. Underscore spellings from v1.0 (`--dict_file`, `--target_lang`, …) remain accepted.

### Verification report

`verify_report<suffix>.<model>.csv` lists, worst first: `slide`, `similarity` (difflib ratio, 0–1), `cer` (Levenshtein distance / length of the intended sequence), `status` (`OK`, `FLAGGED`, or `ENGLISH` when Latin-script words remain in Japanese narration), the intended and recognized text, and the two normalized sequences that were compared (katakana for Japanese; case-folded text without punctuation or spaces otherwise).

## Limitations

- Kana comparison cannot detect pitch-accent errors, and character comparison cannot detect prosody errors. ASR errors, and numbers or units written differently by the ASR (e.g. "5 mg" vs. "five milligrams"), can cause false flags; flagged slides should be checked by listening.
- Automatic language identification can fail for notes mixing several languages or for very short notes in decks without other notes; check the file names after `--extract` or give `--source-lang`.
- Machine translation should be reviewed before synthesis. Glossary placeholders are usually preserved by the translator, but this is not guaranteed.
- Voice cloning should only be used with the consent of the speaker whose voice is cloned.

## Tests

```bash
python tests/smoke_test.py
```

The smoke test mocks the TTS, ASR and translation back-ends, so no models or network access are needed (requires `numpy`, `soundfile` and FFmpeg).

## Citation

If you use PPTX-Narrator, please cite the archived release on Zenodo (see the DOI badge above; metadata in [`CITATION.cff`](CITATION.cff)). Changes between versions are listed in [`CHANGELOG.md`](CHANGELOG.md).

## License

[MIT](LICENSE) © 2026 Hidemi Watanabe
