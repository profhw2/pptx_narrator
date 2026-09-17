# PPTX-Narrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22812409.svg)](https://doi.org/10.5281/zenodo.22812409)

**PPTX-Narrator** turns the presenter notes of a PowerPoint deck into narration spoken in a cloned voice, and writes the audio back into the deck. It is designed for technical and scientific lectures that are revised frequently: when a note changes, the corresponding narration is regenerated instead of re-recorded.

## Features

- **Note extraction** – presenter notes are exported per slide as editable text files; hidden slides are skipped.
- **Technical-term scanning** – acronyms, domain terms and number–unit expressions are collected into a pronunciation dictionary (CSV) for manual review.
- **Reading normalization** – dictionary substitution plus rule-based reading of SI-prefixed units (e.g. `5 mg` → 5ミリグラム, `2 kDa` → 2キロダルトン).
- **Voice cloning** – two interchangeable engines: [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) (via its API server) and [Qwen3-TTS](https://github.com/QwenLM/Qwen3-TTS) (in-process).
- **ASR round-trip verification** (Japanese) – audio is transcribed with `faster-whisper`, both texts are converted to kana with `pyopenjtalk`, and a kana similarity and character error rate (CER) are reported per slide in `verify_report.<model>.csv`.
- **PPTX repackaging** – the embedded audio of each slide is replaced and the automatic slide advance time (`advTm`) is set to the audio duration.
- Optional Japanese→English translation of notes (Google Translate via `deep-translator`).

## Requirements

- Python 3.10 or later
- [FFmpeg](https://ffmpeg.org/) on `PATH` (used by `pydub` to write `.m4a`)
- One TTS engine:
  - **GPT-SoVITS**: a local GPT-SoVITS installation with its API server running, e.g. `python api_v2.py -a 127.0.0.1 -p 9880` in the GPT-SoVITS directory; or
  - **Qwen3-TTS**: installed with the `qwen3` extra below. Model weights (`Qwen/Qwen3-TTS-12Hz-{0.6B,1.7B}-Base`) are downloaded from Hugging Face on first use. A CUDA GPU or Apple Silicon (MPS) is strongly recommended.

## Installation

```bash
git clone https://github.com/profhw2/pptx_narrator.git
cd pptx_narrator
pip install -e .            # core pipeline (GPT-SoVITS engine)
pip install -e ".[qwen3]"   # + Qwen3-TTS engine
pip install -e ".[verify]"  # + ASR verification
pip install -e ".[all]"     # everything
```

This installs the `pptx-narrator` command. Running `python pptx_narrator.py ...` without installing also works (`pip install -r requirements.txt`).
The `--scan` step downloads the NLTK `stopwords` and `words` corpora on first use; if that fails behind a proxy, run `python -m nltk.downloader stopwords words`.

## Quick start

The pipeline is split into steps so that text can be reviewed before synthesis. All intermediate files live in a workspace directory.

```bash
# 1. Extract notes and collect candidate technical terms into dict.csv
pptx-narrator --pptx lecture.pptx --workspace ws --extract --scan

# 2. Review by hand: ws/slide_N.txt (notes) and dict.csv (readings)

# 3. Synthesize Japanese narration with a cloned voice and verify it
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja \
  --tts --engine qwen3 --ref-wav my_voice.wav --ref-text-file my_voice.txt \
  --verify --cer-threshold 0.15

# 4. Listen to the slides flagged in ws/verify_report.qwen3-1.7B.csv, fix the
#    dictionary or notes, re-run step 3 for those slides (--slides 4,7), then pack
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja --engine qwen3 \
  --pack --out lecture_narrated.pptx
```

For English narration, use `--translate` (Japanese notes → `slide_N_eng.txt`) or write English notes directly, and pass `--target-lang en`. ASR verification is currently available for Japanese only.

### Reference voice

`--ref-wav` is a short, clean recording of the target speaker (a few seconds to about ten seconds) and `--ref-text-file` contains its exact transcript.

### Requirements for `--pack`

`--pack` *replaces* audio that is already embedded in a slide; it does not insert new audio objects. Prepare the deck once, e.g. by recording a slide show narration in PowerPoint or inserting any audio clip on each slide to be narrated. To have slides advance automatically, set the slide transition to advance *After* a time; that time is overwritten with the length of the generated audio. Slides without embedded audio are reported and left unchanged.

## Pronunciation dictionary

`dict.csv` has four columns:

| Column | Meaning |
|---|---|
| `Term` | Text to be replaced (matched case-sensitively; alphanumeric terms only on word boundaries) |
| `Japanese_Reading` | Replacement used for `--target-lang ja` |
| `English_Reading` | Replacement used for `--target-lang en` |
| `Type` | Empty for ordinary terms; `unit` for unit symbols that follow a number (e.g. `3 Gbp`) |

Longer terms are applied first; empty readings are ignored. See [`examples/dict_example.csv`](examples/dict_example.csv):

```csv
Term,Japanese_Reading,English_Reading,Type
CRISPR-Cas9,クリスパーキャスナイン,crisper cas nine,
mRNA,メッセンジャーアールエヌエー,messenger R N A,
Gbp,ギガベースペア,,unit
```

Unit symbols not found in the dictionary or the built-in SI rules can be spelled out letter by letter with `--letter-map examples/letter_map_ja.json`.

## Command-line reference

| Option | Description |
|---|---|
| `--pptx` | Input deck (required) |
| `--workspace` | Workspace directory (default: `workspace_<name>_<timestamp>`) |
| `--slides` | Slide selection, e.g. `1-5`, `1,3,5-` |
| `--extract` / `--scan` / `--translate` / `--tts` / `--verify` / `--pack` | Pipeline steps, executed in this order |
| `--target-lang {ja,en}` | Narration language (default: `en`) |
| `--dict-file` | Dictionary CSV (default: `dict.csv`) |
| `--letter-map` | JSON letter-reading map |
| `--engine {gpt_sovits,qwen3}` | TTS engine (default: `gpt_sovits`) |
| `--ref-wav`, `--ref-text-file` | Reference recording and its transcript (required with `--tts`) |
| `--ref-lang {ja,en}` | Language of the reference recording, GPT-SoVITS only (default: `ja`) |
| `--api-url`, `--model` | GPT-SoVITS server URL and model (`v2ProPlus`, `v4`, `v1_clear`) |
| `--qwen3-model-size {0.6B,1.7B}`, `--qwen3-device` | Qwen3-TTS model and device (`auto`, `cuda:0`, `mps`, `cpu`) |
| `--enable-drc`, `--drc-threshold`, `--drc-ratio` | Dynamic range compression of the output |
| `--asr-model`, `--asr-device` | faster-whisper model size and device |
| `--verify-threshold` | Flag slides with kana similarity below this value (default: 0.85) |
| `--cer-threshold` | Also flag slides with kana CER above this value (default: off) |
| `--out` | Output deck for `--pack` (default: `output.pptx`) |
| `--writeback-notes`, `--use-spoken-notes` | Write edited (or normalized) text back into the notes on `--pack` |

Run `pptx-narrator --help` for details. Underscore spellings from v1.0 (`--dict_file`, `--verify_threshold`, …) remain accepted.

### Verification report

`verify_report.<model>.csv` lists, worst first: `slide`, `similarity` (difflib ratio of the two kana strings, 0–1), `cer` (Levenshtein distance / length of the intended kana), `status` (`OK`, `FLAGGED`, or `ENGLISH` when untranslated Latin text remains), the intended and recognized text, and both kana strings.

## Limitations

- ASR verification supports Japanese only.
- Kana comparison cannot detect pitch-accent errors, and ASR errors can cause false flags; flagged slides should be checked by listening.
- Translation relies on the Google Translate web service and requires network access.
- Voice cloning should only be used with the consent of the speaker whose voice is cloned.

## Citation

If you use PPTX-Narrator, please cite the archived release [doi:10.5281/zenodo.22812409](https://doi.org/10.5281/zenodo.22812409) (metadata in [`CITATION.cff`](CITATION.cff)).

## License

[MIT](LICENSE) © 2026 Hidemi Watanabe
