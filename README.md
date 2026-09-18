# PPTX-Narrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

**PPTX-Narrator** turns the presenter notes of a PowerPoint deck into narration spoken in the presenter's own cloned voice, in the language of the notes or translated into another language, and writes the audio back into the deck. The whole deck keeps one consistent voice however often it is revised, even when the presenter cannot speak (e.g. because of a cold).

## Features

- **Note extraction** – presenter notes are exported per slide as editable text files; the language of each note is identified automatically; hidden slides are skipped.
- **Translation** – notes can be translated into any language supported by Google Translate (via `deep-translator`); a dictionary applied to the notes beforehand fixes how technical terms are translated.
- **Technical-term scanning** – acronyms, domain terms and number–unit expressions are collected into a dictionary for manual review.
- **Reading normalization** – a dictionary of string replacements applied to the narration text, plus built-in reading of SI-prefixed units for Japanese (e.g. `5 mg` → 5ミリグラム).
- **Voice cloning** – [Qwen3-TTS](https://github.com/QwenLM/Qwen3-TTS) (in-process) or [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) (via its API server). Which one sounds closer to the speaker is a matter of judgement and changes with engine versions; in the author's use Qwen3-TTS reproduces Japanese and English closely from a single Japanese reference recording, so it is worth trying first.
- **ASR screening (auxiliary)** – audio can be transcribed with `faster-whisper` and compared with the intended text (kana level via `pyopenjtalk` for Japanese, normalized characters otherwise); similarity and character error rate (CER) per slide help decide which slides to listen to first.
- **PPTX repackaging** – the narration audio is embedded in each slide (replacing earlier audio or inserted as a new narration object) and the automatic slide advance time (`advTm`) is set to the audio duration, so the deck plays as a self-running show and can be exported as a video. Translated narration can be written into the notes above the original note, in a marked layout that later extractions recognize.

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

To try the tool without preparing a deck, build the sample deck first; its notes contain acronyms, gene names and number-unit expressions in Japanese and English:

```bash
python examples/make_sample_deck.py sample_lecture.pptx
```

Any few seconds of clear speech with its exact transcript can serve as the reference voice (`--ref-wav` / `--ref-text-file`).

The pipeline is split into steps so that text can be reviewed before synthesis. Intermediate files live in a workspace directory. A dictionary (`--dict-file`) is a list of string replacements applied to the text processed in the run: to the notes before translation and to the narration text before synthesis.

### Narration in the language of the notes

```bash
# 1. Extract notes and collect candidate technical terms into readings_ja.csv
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja \
  --extract --scan --dict-file readings_ja.csv

# 2. Review by hand: ws/slide_N.txt (notes) and readings_ja.csv (readings)

# 3. Synthesize with a cloned voice and verify
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja --dict-file readings_ja.csv \
  --tts --engine qwen3 --ref-wav my_voice.wav --ref-text-file my_voice.txt \
  --verify --cer-threshold 0.15

# 4. Listen to the slides flagged in ws/verify_report.qwen3-1.7B.csv, fix the
#    dictionary or notes, re-run step 3 for those slides (--slides 4,7), then pack
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang ja --engine qwen3 \
  --pack --out lecture_narrated.pptx
```

### Translated narration

Translation is usually done once, while synthesis is repeated while readings are refined, so it is convenient to run them separately, each with its own dictionary:

```bash
# 1. Extract the notes (language identified automatically) and collect terms to be translated
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de \
  --extract --scan --dict-file terms_ja_de.csv
# 2. Fill in the German terms in terms_ja_de.csv, then translate
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de \
  --translate --dict-file terms_ja_de.csv
# 3. Review ws/slide_N_de.txt; collect candidate readings from the German text
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de \
  --scan --dict-file readings_de.csv
# 4. Review readings_de.csv, then synthesize, verify and pack (repeat as needed),
#    writing the German narration above the original note
pptx-narrator --pptx lecture.pptx --workspace ws --target-lang de --dict-file readings_de.csv \
  --tts --engine qwen3 --ref-wav my_voice.wav --ref-text-file my_voice.txt \
  --verify --pack --writeback-notes --out lecture_de.pptx
```

Existing translations are not overwritten; use `--retranslate` after changing the translation dictionary. In a single run with both `--translate` and `--tts`, the dictionaries are applied before translation and again before synthesis. This works, but replacements meant as readings (e.g. `CRISPR,C R I S P R`) then also end up in the translated text and in the written-back notes, and may be altered by the translator.

### Workspace files

| File | Content |
|---|---|
| `slide_N.txt`, `slide_N_eng.txt`, `slide_N_<lang>.txt` | Note text in Japanese, English, or another language (e.g. `slide_3_de.txt`) |
| `slide_N<suffix>.<model>.spoken.txt` | Narration text after dictionary replacement, as sent to the TTS engine |
| `slide_N<suffix>.<model>.m4a` | Generated audio |
| `verify_report<suffix>.<model>.csv` | Verification report |
| `translations.json` | Which version of each source note a translation was made from |
| `slide_N<suffix>.stale.txt` | Previous narration set aside because the source note changed |

### Reference voice

`--ref-wav` is a short, clean recording (a few seconds to about ten seconds) and `--ref-text-file` contains its exact transcript; this recording alone defines the voice. It is typically the presenter's own voice, but any voice can be used with the speaker's consent, for example a native speaker's voice for a translated version. Keep reference recordings private. For GPT-SoVITS, give the language of the recording with `--ref-lang` (default `ja`).

## Dictionaries

A dictionary is a CSV file of plain string replacements. The header row is optional:

```csv
string,replacement,type
CRISPR-Cas9,クリスパーキャスナイン,
mRNA,メッセンジャーアールエヌエー,
Gbp,ギガベースペア,unit
```

- It is applied to the text processed in the run: with `--translate`, to each note before it is sent to the translator (e.g. `塩基対,Basenpaare` to fix a German term); with `--tts`, to the narration text before synthesis (readings); with both, at both points. The note files themselves are not changed; the rewritten narration is saved as `*.spoken.txt`.
- Longer strings are replaced first; alphanumeric strings only match on word boundaries; rows with an empty replacement are ignored.
- `type` = `unit` marks unit symbols that follow a number (`3 Gbp`). For Japanese narration, built-in rules additionally read SI-prefixed units.
- `--dict-file` can be given several times (e.g. a shared and a deck-specific file); later files take precedence.
- `--scan` appends new candidates to the first `--dict-file` (created if missing). It scans the narration text in `--target-lang` when it exists, filling in provisional readings; otherwise, or together with `--translate`, it scans the notes to be translated and leaves the replacements blank.

A v1.x dictionary (`Term,Japanese_Reading,English_Reading,Type`) can still be given with `--dict-file`; the column of the narration language (Japanese or English) is used. See the files in [`examples/`](examples/); [`examples/readings_ja_molbio.csv`](examples/readings_ja_molbio.csv) is a working dictionary from the author's molecular-biology lectures (109 entries: gene and organism names, researchers' names, units and Japanese words the engine misreads), which can be used as a starting point for that field. A dictionary belongs to a subject area rather than to a deck: once the terms of a course are in it, later decks need few new entries. An entry with an empty replacement leaves the term as it is and stops `--scan` from proposing it again. Unit symbols not resolved otherwise can be spelled out letter by letter with a letter map for the narration language, e.g. `--letter-map examples/letter_map_ja.json`.

## Packing and notes write-back

`--pack` embeds the generated audio in the same structure PowerPoint uses for recorded narration: the audio starts with the slide, is hidden during the show, and the slide advances automatically after the audio length. Slides without audio get a new narration object (a small speaker icon at the bottom right, visible only in the editor). If a slide already has audio (e.g. an earlier recording), that object is reused: it points to the new audio, and its trim, fade and bookmarks are removed. Each slide gets its own media file, so copied slides that shared one clip no longer overwrite each other, and clips no longer used are removed from the file. Slides that have animations but no audio are reported and left unchanged; insert any audio clip on such a slide in PowerPoint and pack again. The slide advances by itself one second after the narration ends, so that the last word is not clipped in a video; `--slide-pause` changes that pause. The audio icon of a narrated slide is parked next to the slide, outside the visible area, so that it does not cover the slide content while editing; `--keep-audio-icon` leaves it where it is. The icon is hidden during the slide show either way.

Data recorded together with the old audio is also removed from a narrated slide, because its timing belongs to that audio: the laser-pointer path (`p14:laserTraceLst`) and the recorded play/pause/seek events (`p14:showEvtLst`). `--remove-recorded pointer|events|none` narrows or disables this (default: `all`). Ink annotations are kept and reported. The packed deck can be exported as MP4 with PowerPoint's *Export* (use recorded timings and narrations).

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
| `--extract` / `--scan` / `--translate` / `--tts` / `--verify` / `--pack` | Pipeline steps, executed in this order |
| `--retranslate` | With `--translate`, overwrite existing translations |
| `--source-lang` | Language of the notes (default: `auto`) |
| `--target-lang` | Narration language (default: `--source-lang` if given, otherwise `en`) |
| `--dict-file` | Dictionary of string replacements (repeatable); applied to the notes before `--translate` and to the narration text before `--tts` |
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
| `--slide-pause` | Seconds between the end of the narration and the automatic slide advance (default: 1.0) |
| `--keep-audio-icon` | Leave the audio icon on the slide instead of parking it outside the visible area |
| `--remove-recorded` | `all` (default), `pointer`, `events` or `none`: what to remove of the data recorded with the previous slide show |

Run `pptx-narrator --help` for details. Underscore spellings from v1.0 (`--dict_file`, `--target_lang`, …) remain accepted.

### Verification report

The report is a listening aid, not a pass/fail test: high similarity usually means the narration is fine, while many low-scoring slides sound natural and only reflect recognition errors. `verify_report<suffix>.<model>.csv` lists, worst first: `slide`, `similarity` (difflib ratio, 0–1), `cer` (Levenshtein distance / length of the intended sequence), `status` (`OK`, `FLAGGED`, or `ENGLISH` when Latin-script words remain in Japanese narration), the intended and recognized text, and the two normalized sequences that were compared (katakana for Japanese; case-folded text without punctuation or spaces otherwise).

## Limitations

- Kana comparison cannot detect pitch-accent errors, and character comparison cannot detect prosody errors. ASR errors, and numbers or units written differently by the ASR (e.g. "5 mg" vs. "five milligrams"), can cause false flags; flagged slides should be checked by listening.
- Automatic language identification can fail for notes mixing several languages or for very short notes in decks without other notes; check the file names after `--extract` or give `--source-lang`.
- Machine translation should be reviewed before synthesis; terms replaced before translation can still be altered by the translator.
- Voice cloning should only be used with the consent of the speaker whose voice is cloned.

## Tests

```bash
python tests/smoke_test.py
```

The smoke test mocks the TTS, ASR and translation back-ends, so no models or network access are needed (requires `numpy`, `soundfile` and FFmpeg).

## Support and contributions

Questions and bug reports are welcome as GitHub issues. Please include the command you ran, the log output, and, if possible, a small deck that reproduces the problem (`examples/make_sample_deck.py` builds one). Do not attach reference recordings or unpublished course material.

The tool is maintained alongside teaching and research, so replies can take a while and new features are added as they become necessary for the author's own lecture material. Pull requests are welcome; small, self-contained changes are the easiest to review, and `python tests/smoke_test.py` should pass.

## Citation

If you use PPTX-Narrator, please cite the archived release on Zenodo; its DOI is added here and to [`CITATION.cff`](CITATION.cff) with the v1.0.0 release. Changes between versions are listed in [`CHANGELOG.md`](CHANGELOG.md).

## License

[MIT](LICENSE) © 2026 Hidemi Watanabe
