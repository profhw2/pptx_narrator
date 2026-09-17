# PPTX-Narrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22812409.svg)](https://doi.org/10.5281/zenodo.22812409)

**PPTX-Narrator** turns the presenter notes of a PowerPoint deck into narration spoken in a cloned voice, optionally in another language, and writes the audio back into the deck. It is designed for technical and scientific lectures that are revised frequently: when a note changes, the corresponding narration is regenerated instead of re-recorded.

## Features

- **Note extraction** – presenter notes are exported per slide as editable text files; the language of each note is identified automatically; hidden slides are skipped.
- **Translation** – notes can be translated from any source language into any target language supported by Google Translate (via `deep-translator`).
- **Technical-term scanning** – acronyms, domain terms and number–unit expressions are collected into a pronunciation dictionary (CSV) with one reading column per language, for manual review.
- **Reading normalization** – dictionary substitution per narration language, plus built-in reading of SI-prefixed units for Japanese (e.g. `5 mg` → 5ミリグラム).
- **Voice cloning** – [Qwen3-TTS](https://github.com/QwenLM/Qwen3-TTS) (in-process) or [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) (via its API server); the voice of a short reference recording is reused across languages.
- **ASR round-trip verification** – audio is transcribed with `faster-whisper` and compared with the intended text (kana level via `pyopenjtalk` for Japanese, normalized characters for other languages); similarity and character error rate (CER) are reported per slide.
- **PPTX repackaging** – the embedded audio of each slide is replaced and the automatic slide advance time (`advTm`) is set to the audio duration.

## Supported languages

| Step | Languages |
|---|---|
| `--translate` | Any source/target pair supported by Google Translate |
| `--tts --engine qwen3` | zh, en, ja, ko, de, fr, ru, pt, es, it |
| `--tts --engine gpt_sovits` | zh, en, ja, ko, yue (Cantonese) |
| `--verify` | Japanese: kana-level comparison; other languages: character-level comparison of normalized text (any language recognized by Whisper) |
| Built-in unit readings | Japanese only (other languages: `unit` entries in the dictionary) |

Languages are given as Google Translate codes such as `ja`, `en`, `de`, `zh-CN`. With `--source-lang auto` (default), the language of each note is identified from its text: kana and Hangul are recognized directly, other languages with [py3langid](https://github.com/adbar/py3langid) (restricted to the languages Google Translate accepts). Notes too short or ambiguous to identify reliably, such as "Thank you." or a kanji-only title, are assigned the main language of the deck. The detected language of each slide is shown in the log and in the file name; give `--source-lang` to override it.

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

The pipeline is split into steps so that text can be reviewed before synthesis. All intermediate files live in a workspace directory.

### Narration in the language of the notes

```bash
# 1. Extract notes and collect candidate technical terms (readings for Japanese)
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja --extract --scan

# 2. Review by hand: ws/slide_N.txt (notes) and dict.csv (readings)

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
# Notes in any language (identified automatically) -> German narration in the same (cloned) voice
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de \
  --extract --translate --scan
# review ws/slide_N_de.txt and the Reading_de column of dict.csv, then:
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de \
  --tts --engine qwen3 --ref-wav my_voice.wav --ref-text-file my_voice.txt \
  --verify --pack --out lecture_de.pptx
```

### Workspace files

| File | Content |
|---|---|
| `slide_N.txt`, `slide_N_eng.txt`, `slide_N_<lang>.txt` | Note text in Japanese, English, or another language (e.g. `slide_3_de.txt`) |
| `slide_N<suffix>.<model>.spoken.txt` | Normalized text actually sent to the TTS engine |
| `slide_N<suffix>.<model>.m4a` | Generated audio |
| `verify_report<suffix>.<model>.csv` | Verification report |

### Reference voice

`--ref-wav` is a short, clean recording of the target speaker (a few seconds to about ten seconds) and `--ref-text-file` contains its exact transcript. For GPT-SoVITS, give its language with `--ref-lang` (default `ja`).

### Requirements for `--pack`

`--pack` *replaces* audio that is already embedded in a slide; it does not insert new audio objects. Prepare the deck once, e.g. by recording a slide show narration in PowerPoint or inserting any audio clip on each slide to be narrated. To have slides advance automatically, set the slide transition to advance *After* a time; that time is overwritten with the length of the generated audio. Slides without embedded audio are reported and left unchanged. With `--writeback-notes`, the narration text (followed by the source-language text, if different) is written into the notes.

## Pronunciation dictionary

`dict.csv` has a `Term` column, one reading column per narration language and an optional `Type` column:

| Column | Meaning |
|---|---|
| `Term` | Text to be replaced (matched case-sensitively; alphanumeric terms only on word boundaries) |
| `Japanese_Reading`, `English_Reading` | Readings for `ja` and `en` (v1.x column names) |
| `Reading_<lang>` | Reading for any other language, e.g. `Reading_de`, `Reading_zh-CN` |
| `Type` | Empty for ordinary terms; `unit` for unit symbols that follow a number (e.g. `3 Gbp`) |

`--scan` adds the `Reading_<lang>` column for `--target-lang` when it is missing. Longer terms are applied first; empty readings are ignored. Dictionaries from v1.x work unchanged. See [`examples/dict_example.csv`](examples/dict_example.csv):

```csv
Term,Japanese_Reading,English_Reading,Type,Reading_de
CRISPR-Cas9,クリスパーキャスナイン,crisper cas nine,,
mRNA,メッセンジャーアールエヌエー,messenger R N A,,Boten-RNA
Gbp,ギガベースペア,gigabase pairs,unit,Gigabasenpaare
```

Unit symbols not resolved otherwise can be spelled out letter by letter with a letter map for the narration language, e.g. `--letter-map examples/letter_map_ja.json`.

## Command-line reference

| Option | Description |
|---|---|
| `--pptx` | Input deck (required) |
| `--workspace` | Workspace directory (default: `workspace_<name>_<timestamp>`) |
| `--slides` | Slide selection, e.g. `1-5`, `1,3,5-` |
| `--extract` / `--scan` / `--translate` / `--tts` / `--verify` / `--pack` | Pipeline steps, executed in this order |
| `--source-lang` | Language of the notes (default: `auto`) |
| `--target-lang` | Narration language (default: `--source-lang` if given, otherwise `en`) |
| `--dict-file` | Dictionary CSV (default: `dict.csv`) |
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
| `--writeback-notes`, `--use-spoken-notes` | Write narration (or normalized) text back into the notes on `--pack` |

Run `pptx-narrator --help` for details. Underscore spellings from v1.0 (`--dict_file`, `--target_lang`, …) remain accepted.

### Verification report

`verify_report<suffix>.<model>.csv` lists, worst first: `slide`, `similarity` (difflib ratio, 0–1), `cer` (Levenshtein distance / length of the intended sequence), `status` (`OK`, `FLAGGED`, or `ENGLISH` when Latin-script words remain in Japanese narration), the intended and recognized text, and the two normalized sequences that were compared (katakana for Japanese; case-folded text without punctuation or spaces otherwise).

## Limitations

- Kana comparison cannot detect pitch-accent errors, and character comparison cannot detect prosody errors. ASR errors, and numbers or units written differently by the ASR (e.g. "5 mg" vs. "five milligrams"), can cause false flags; flagged slides should be checked by listening.
- Automatic language identification can fail for very short notes in decks without other notes, and for notes mixing several languages; check the file names after `--extract` or give `--source-lang`.
- Machine translation should be reviewed before synthesis, especially for technical terms.
- Voice cloning should only be used with the consent of the speaker whose voice is cloned.

## Tests

```bash
python tests/smoke_test.py
```

The smoke test mocks the TTS, ASR and translation back-ends, so no models or network access are needed (requires `numpy`, `soundfile` and FFmpeg).

## Citation

If you use PPTX-Narrator, please cite the archived release on Zenodo (see the DOI badge above; metadata in [`CITATION.cff`](CITATION.cff)). The changes between versions are listed in [`CHANGELOG.md`](CHANGELOG.md).

## License

[MIT](LICENSE) © 2026 Hidemi Watanabe
