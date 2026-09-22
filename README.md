# PPTX-Narrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

**PPTX-Narrator** turns the presenter notes of a PowerPoint deck into narration spoken in the presenter's own cloned voice, in the language of the notes or translated into another language, and writes the audio back into the deck. The whole deck keeps one consistent voice however often it is revised, even when the presenter cannot speak (e.g. because of a cold).

## Features

- **Note extraction** – presenter notes are exported per slide as editable text files; the language of each note is identified automatically; hidden slides are skipped.
- **Translation** – notes can be translated into any language supported by Google Translate (via `deep-translator`); a dictionary applied to the notes beforehand fixes how technical terms are translated.
- **Technical-term scanning** – acronyms, domain terms and number–unit expressions are collected into a dictionary for manual review.
- **Reading normalization** – a dictionary of string replacements applied to the narration text, plus built-in reading of SI-prefixed units for Japanese (e.g. `5 mg` → 5ミリグラム).
- **Voice cloning** – [Qwen3-TTS](https://github.com/QwenLM/Qwen3-TTS) (in-process) or [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) (via its API server). Which one sounds closer to the speaker is a matter of judgement and changes with engine versions; in the author's use Qwen3-TTS reproduces Japanese and English closely from a single Japanese reference recording, which is why it is the default; any default can be set in the configuration file.
- **ASR screening (auxiliary)** – audio can be transcribed with `faster-whisper` and compared with the intended text (kana level via `pyopenjtalk` for Japanese, normalized characters otherwise); per-slide similarity, the longest single stretch of disagreement and the character error rate (CER) flag the slides most likely to be misread.
- **PPTX repackaging** – the narration audio is embedded in each slide (replacing earlier audio or inserted as a new narration object) and the automatic slide advance time (`advTm`) is set to the audio duration plus a short pause, so the deck plays as a self-running show and can be exported as a video. Translated narration can be written into the notes above the original note, in a marked layout that later extractions recognize.

## Supported languages

| Step | Languages |
|---|---|
| Note language identification | Kana/Hangul rules and [py3langid](https://github.com/adbar/py3langid), restricted to Google Translate languages |
| `translate` | Any source/target pair supported by Google Translate |
| `synthesize --engine qwen3` | zh, en, ja, ko, de, fr, ru, pt, es, it |
| `synthesize --engine gpt_sovits` | zh, en, ja, ko, yue (Cantonese) |
| `verify` | Japanese: kana-level comparison; other languages: character-level comparison of normalized text |
| Built-in unit readings | Japanese only (other languages: `unit` entries in the dictionaries) |

Languages are written as Google Translate codes such as `ja`, `en`, `de`, `zh-CN`. Without `--in-lang`, `extract` identifies the language of each note from its text; notes too short or ambiguous to identify reliably ("Thank you.", a kanji-only title) are assigned the main language of the deck. The detected language is shown in the log and in the file name. `--in-lang` selects which languages to extract: a note in another language is reported and left out, and a requested language the deck does not contain stops the run with an error.

## Requirements

- Python 3.10 or later
- [FFmpeg](https://ffmpeg.org/) on `PATH` (used by `pydub` to write `.m4a`)
- One TTS engine:
  - **Qwen3-TTS**: installed with the `qwen3` extra below. Model weights (`Qwen/Qwen3-TTS-12Hz-{0.6B,1.7B}-Base`) are downloaded from Hugging Face on first use. A CUDA GPU or Apple Silicon (MPS) is strongly recommended.
  - **GPT-SoVITS**: a local GPT-SoVITS installation with its API server running, e.g. `python api_v2.py -a 127.0.0.1 -p 9880` in the GPT-SoVITS directory.
- Network access for `translate` (Google Translate).

## Installation

```bash
git clone https://github.com/profhw2/pptx_narrator.git
cd pptx_narrator
pip install -e .            # pipeline without a local TTS engine
pip install -e ".[qwen3]"   # + Qwen3-TTS, the default engine
pip install -e ".[verify]"  # + ASR verification
pip install -e ".[all]"     # everything
```

This installs the `pptx-narrator` command. Running `python pptx_narrator.py ...` without installing also works (`pip install -r requirements.txt`).
The `scan` command downloads the NLTK `stopwords` and `words` corpora on first use; if that fails behind a proxy, run `python -m nltk.downloader stopwords words`.

## Quick start

To try the tool without preparing a deck, build the sample deck first; its notes contain acronyms, gene names and number-unit expressions in Japanese and English:

```bash
python examples/make_sample_deck.py sample_lecture.pptx
```

Any few seconds of clear speech with its exact transcript can serve as the reference voice (`--ref-wav` / `--ref-text`).

The pipeline is split into steps so that text can be reviewed before synthesis. Intermediate files live in a workspace directory. A dictionary (`--dict-file`) is a list of string replacements applied to the text processed in the run: to the notes before translation and to the narration text before synthesis.

### Narration in the language of the notes

```bash
# 1. Extract the notes, then collect candidate readings into readings_ja.csv
pptx-narrator extract lecture.pptx --workspace ws --in-lang ja
pptx-narrator scan ws --in-lang ja --dict-file readings_ja.csv

# 2. Review by hand: ws/slide_N_ja.txt (notes) and readings_ja.csv (readings)

# 3. Synthesize with a cloned voice, then screen the result
pptx-narrator synthesize ws --in-lang ja --dict-file readings_ja.csv \
  --engine qwen3 --ref-wav my_voice.wav --ref-text my_voice.txt
pptx-narrator verify ws --in-lang ja --engine qwen3 --cer-threshold 0.15

# 4. Listen to the slides flagged in ws/verify_report_ja.qwen3-1.7B.csv, fix the
#    dictionary or notes, re-run step 3 for one slide by naming its file
#    (pptx-narrator synthesize ws/slide_4_ja.txt ...), then pack
pptx-narrator pack lecture.pptx --workspace ws --in-lang ja --engine qwen3 \
  --out lecture_narrated.pptx
```

### Translated narration

Translation is usually done once, while synthesis is repeated while readings are refined, so it is convenient to run them separately, each with its own dictionary:

```bash
# 1. Extract the notes (language identified automatically) and collect terms to be translated
pptx-narrator extract lecture.pptx --workspace ws
pptx-narrator scan ws --in-lang ja --dict-file terms_ja_de.csv
# 2. Fill in the German terms in terms_ja_de.csv, then translate
pptx-narrator translate ws --in-lang ja --out-lang de --dict-file terms_ja_de.csv
# 3. Review ws/slide_N_de.txt; collect candidate readings from the German text
pptx-narrator scan ws --in-lang de --dict-file readings_de.csv
# 4. Review readings_de.csv, then synthesize, verify and pack (repeat as needed),
#    writing the German narration above the original note
pptx-narrator synthesize ws --in-lang de --dict-file readings_de.csv \
  --engine qwen3 --ref-wav my_voice.wav --ref-text my_voice.txt
pptx-narrator verify ws --in-lang de --engine qwen3
pptx-narrator pack lecture.pptx --workspace ws --in-lang de --engine qwen3 \
  --writeback-notes --out lecture_de.pptx
```

Existing translations are not overwritten; use `--retranslate` after changing the translation dictionary. Because `translate` and `synthesize` are separate commands, each is given its own dictionary: replacements meant as readings (e.g. `CRISPR,C R I S P R`) stay out of the translated text and of the written-back notes.

### Workspace files

| File | Content |
|---|---|
| `slide_N_<lang>.txt` | Note text, in the language named in the file name (`slide_3_ja.txt`, `slide_3_en.txt`, `slide_3_de.txt`) |
| `slide_N_<lang>.<model>.spoken.txt` | Narration text after dictionary replacement, as sent to the TTS engine |
| `slide_N_<lang>.<model>.m4a` | Generated audio |
| `verify_report_<lang>.<model>.csv` | Verification report, one row per slide |
| `verify_differences_<lang>.<model>.csv` | Every place where the narration and the transcript disagree |
| `translations.json` | Which version of each source note a translation was made from |
| `slide_N_<lang>.stale.txt` | Previous narration set aside because the source note changed |

### Reference voice

`--ref-wav` is a short, clean recording (a few seconds to about ten seconds) and `--ref-text` contains its exact transcript; this recording alone defines the voice. It is typically the presenter's own voice, but any voice can be used with the speaker's consent, for example a native speaker's voice for a translated version. Keep reference recordings private. For GPT-SoVITS, give the language of the recording with `--ref-lang` (default `ja`).

## Dictionaries

A dictionary is a CSV file of plain string replacements. The header row is optional:

```csv
string,replacement,type
CRISPR-Cas9,クリスパーキャスナイン,
mRNA,メッセンジャーアールエヌエー,
Gbp,ギガベースペア,unit
```

- It is applied to the text the command processes: with `translate`, to each note before it is sent to the translator (e.g. `塩基対,Basenpaare` to fix a German term); with `synthesize`, to the narration text before synthesis (readings). The note files themselves are not changed; the rewritten narration is saved as `*.spoken.txt`.
- Longer strings are replaced first; alphanumeric strings only match on word boundaries; rows with an empty replacement are ignored.
- `type` = `unit` marks unit symbols that follow a number (`3 Gbp`). For Japanese narration, built-in rules additionally read SI-prefixed units.
- `--dict-file` can be given several times (e.g. a shared and a deck-specific file); later files take precedence.
- `scan` appends new candidates to the first `--dict-file` (created if missing). It scans the text in `--in-lang`, filling in provisional readings when that is the narration language and leaving the replacements blank when the text is still to be translated. With `--scan-compounds` and Japanese narration, it additionally writes every compound of the notes, with the reading a Japanese front end assembles for it, as **comment lines** (`;二本鎖,ニホンクサリ,`). Comments do nothing, so the list can be long; compounds are where readings are unsettled (鎖 is クサリ alone but サ in 二本鎖), and an entry is activated by correcting the reading and removing the `;`. Needs `pyopenjtalk` (the `verify` extra).

A four-column dictionary (`Term,Japanese_Reading,English_Reading,Type`) from an earlier development version can still be given with `--dict-file`; the column of the narration language (Japanese or English) is used. See the files in [`examples/`](examples/); [`examples/readings_ja_molbio.csv`](examples/readings_ja_molbio.csv) is a working dictionary from the author's molecular-biology lectures (104 entries: gene and organism names, researchers' names, units and Japanese words the engine misreads), which can be used as a starting point for that field. A dictionary belongs to a subject area rather than to a deck: once the terms of a course are in it, later decks need few new entries. An entry with an empty replacement leaves the term as it is and stops `scan` from proposing it again; lone letters and digits are never proposed. A `;` at the start of a line or after a space begins a comment that runs to the end of the line, so an entry can be annotated or switched off without deleting it, and a line that is only a comment is skipped. Terms containing `#` or `;` are unaffected (`C#`); a term that begins with `;` is written in double quotes. Unit symbols not resolved otherwise can be spelled out letter by letter with a letter map for the narration language, e.g. `--letter-map examples/letter_map_ja.json`.

## Packing and notes write-back

`pack` embeds the generated audio in the same structure PowerPoint uses for recorded narration: the audio starts with the slide, is hidden during the show, and the slide advances automatically after the audio length. Slides without audio get a new narration object (a small speaker icon at the bottom right, visible only in the editor). If a slide already has audio (e.g. an earlier recording), that object is reused: it points to the new audio, and its trim, fade and bookmarks are removed. Each slide gets its own media file, so copied slides that shared one clip no longer overwrite each other, and clips no longer used are removed from the file. Slides that have animations but no audio are reported and left unchanged; insert any audio clip on such a slide in PowerPoint and pack again. The slide advances by itself one second after the narration ends, so that the last word is not clipped in a video; `--slide-pause` changes that pause. The audio icon of a narrated slide is parked next to the slide, outside the visible area, so that it does not cover the slide content while editing; `--keep-audio-icon` leaves it where it is. The icon is hidden during the slide show either way.

Data recorded together with the old audio is also removed from a narrated slide, because its timing belongs to that audio: the laser-pointer path (`p14:laserTraceLst`) and the recorded play/pause/seek events (`p14:showEvtLst`). `--remove-recorded pointer|events|none` narrows or disables this (default: `all`). Ink annotations are kept and reported. The packed deck can be exported as MP4 with PowerPoint's *Export* (use recorded timings and narrations).

With `--writeback-notes`, the human-editable narration text is written into the notes. When it differs from the original note (for example, translated narration), the note keeps both parts:

```
=== pptx-narrator: narration [en] from [ja] #ae82d4f1fc ===
Today we talk about genome editing.

=== pptx-narrator: source [ja] ===
今日はゲノム編集について話します。
```

Keep the marker lines when editing such notes in PowerPoint. When `extract` finds this layout, only the source part is used as the note to translate. The narration part is restored as the existing translation, including edits made in PowerPoint, as long as the source part is unchanged (the hash identifies the translated version). If the source part was edited, the old narration is set aside as `slide_N_<lang>.stale.txt` and `translate` produces a new translation. Spoken-form narration (`spoken` in the marker) is never restored.

## Command-line reference

```
pptx-narrator COMMAND [INPUT] [OPTIONS]
```

`COMMAND` is one of `extract`, `scan`, `translate`, `synthesize`, `verify`, `pack`; nothing runs unless a command
says so. `INPUT` is a file or a directory: a directory is processed as a whole, a file on its own, so one slide is
redone by naming its file (`pptx-narrator synthesize ws/slide_4_ja.txt`), or by selecting it with `--slides`
(`pptx-narrator synthesize ws --in-lang ja --slides 4,7-9`), which every command accepts. `extract` and `pack` take the deck;
`scan`, `translate`, `synthesize` and `verify` take the workspace or one of its files, and refuse a deck. If `INPUT` is omitted, the input recorded
by the previous run of that command is reused, but only after its SHA-256 still matches; a changed input has to be
named again.

The two language options mean the same thing in every command: `--in-lang` is the language of the data the command
reads, `--out-lang` the language of the data it writes. Each command accepts only the options it needs.

| Command | INPUT | Options |
|---|---|---|
| `extract` | PPTX | `--in-lang` (languages to extract, e.g. `ja` or `ja,en`; omitted = every language found), `--workspace`, `--slides` |
| `scan` | text file or workspace | `--in-lang`, `--dict-file`, `--scan-compounds`, `--slides`, `--workspace` |
| `translate` | text file or workspace | `--in-lang`, `--out-lang`, `--dict-file`, `--retranslate`, `--slides`, `--workspace` |
| `synthesize` | text file or workspace | `--in-lang`, `--dict-file`, `--letter-map`, `--engine {gpt_sovits,qwen3}` (default qwen3), `--ref-wav`, `--ref-text`, `--ref-lang`, `--api-url`, `--model`, `--qwen3-model-size {0.6B,1.7B}` (default 1.7B), `--qwen3-device`, `--enable-drc`, `--drc-threshold`, `--drc-ratio`, `--slides`, `--workspace` |
| `verify` | audio/text file or workspace | `--in-lang`, `--engine`, `--model`, `--qwen3-model-size`, `--asr-model`, `--asr-device`, `--verify-threshold` (default 0.85), `--min-difference` (default 4), `--max-difference` (default 40; `0` disables), `--cer-threshold` (default off), `--slides`, `--workspace` |
| `pack` | PPTX | `--workspace`, `--out` (default `output.pptx`), `--in-lang`, `--engine`, `--model`, `--qwen3-model-size`, `--slides`, `--writeback-notes`, `--slide-pause` (default 1.0 s), `--keep-audio-icon`, `--remove-recorded {all,pointer,events,none}` (default `all`) |

`--config FILE` and `--version` are accepted before the command. Run `pptx-narrator COMMAND --help` for the full
list. Underscore spellings (`--dict_file`, `--in_lang`, …) are also accepted, and any option may be typed as short
as it stays unambiguous (`--work ws` for `--workspace ws`; `--ref-t my_voice.txt` for `--ref-text`, since
`--ref-wav` and `--ref-lang` also start with `--ref-`).

## Configuration

Parameters that stay the same from run to run go into a TOML file, so that a command line carries only what changes.
`--config FILE` names it; otherwise `pptx_narrator.toml` in the current directory is used when it exists. A `[common]`
section applies to every command, and a section named after a command applies to that command and overrides `[common]`:

```toml
[common]
in_lang = "ja"
workspace = "ws"
engine = "qwen3"          # synthesize, verify and pack all need it

[synthesize]
ref_wav = "my_voice.wav"
ref_text = "my_voice.txt"
dict_file = "readings_ja.csv"

[verify]
cer_threshold = 0.15
```

`engine` (and `qwen3-model-size`, if it is not the default) belongs in `[common]`, because the generated file names
carry the engine and model: `synthesize` writes `slide_3_ja.qwen3-1.7B.m4a`, and `verify` and `pack` look for that same
name. Setting the engine under `[synthesize]` alone leaves the other two looking for `v2ProPlus` files and finding
nothing.

With that file, step 3 of the quick start is `pptx-narrator synthesize ws` and `pptx-narrator verify ws`.

Keys may be written with hyphens or underscores (`dict-file` and `dict_file` both work), as on the command line.

Values are resolved in one order: **built-in defaults → configuration file → command line**, the command line winning.
To take a configured value back for one run, give it empty (`--dict-file ''`); to turn off a switch the file sets, use
its `--no-` form (`--no-writeback-notes`). `--config` may be written before or after the command.
After every run the result is written to `.pptx_narrator_resolved.toml`: every parameter with the value actually used,
the version of the tool, the time, and the path, size and SHA-256 of each input file. That file is itself a valid
configuration file, so a run can be repeated later from it, and it records what produced a given narration.
`.pptx_narrator_state.json` holds the last input of each command and is what the SHA-256 check above compares against.

### Verification report

The step writes two files. `verify_differences_<lang>.<model>.csv` lists every place where the narration and the transcript disagree, longest first: the slide, the length, the position, what the text said, what the ASR heard, and a few characters of context on each side. This is the list to read: it says where to listen, often what happened, and it can be sorted or filtered as you like; `--min-difference` sets how short a difference is still listed (default 4 characters). `verify_report_<lang>.<model>.csv` summarizes the same comparison per slide, worst first, and marks a slide `FLAGGED` when its similarity falls below `--verify-threshold` or when one stretch of disagreement is longer than `--max-difference` characters (default 40): the first catches a small error in a short note, the second a dropped phrase in a note of any length. Neither is a pass/fail test; many low-scoring slides sound natural and only reflect recognition errors. The report lists, worst first: `slide`, `similarity` (difflib ratio, 0–1), `cer` (Levenshtein distance / length of the intended sequence), `status` (`OK`, `FLAGGED`, or `ENGLISH` when Latin-script words remain in Japanese narration), `differences` and `longest_difference` (how many places differ and how long the longest stretch is, which is what distinguishes a skipped phrase from scattered recognition differences), the intended and recognized text, and the two normalized sequences that were compared (katakana for Japanese; case-folded text without punctuation or spaces otherwise).

## Limitations

- Kana comparison cannot detect pitch-accent errors, and character comparison cannot detect prosody errors. ASR errors, and numbers or units written differently by the ASR (e.g. "5 mg" vs. "five milligrams"), can cause false flags; flagged slides should be checked by listening.
- Automatic language identification can fail for notes mixing several languages or for very short notes in decks without other notes; check the file names after `extract` or give `--in-lang`.
- Machine translation should be reviewed before synthesis; terms replaced before translation can still be altered by the translator.
- Voice cloning should only be used with the consent of the speaker whose voice is cloned.

## Checking the ASR screening

`examples/screening_check.py` measures which narration errors the check actually notices. The transcript of a slide depends only on its audio, so it is produced once (or read from an existing report) and any number of hypothetical errors can then be scored against it. The script injects one error of a known size into the sequence the check compares (katakana for Japanese, normalized text otherwise) -- a run of characters deleted, as when a phrase is skipped, or replaced by other characters, as when a term is misread -- and reports how often the check notices, by the size of the error and the length of the note. Sizes are in characters of that sequence, so 3 characters is about one short term and 50 is about one sentence:

```bash
python examples/screening_check.py --workspace ws --in-lang ja --asr-model small --asr-device cuda
```

The workspace is not modified; `--sizes` and `--repeats` set the error sizes in characters and the number of random positions per slide and size. The result says for which note lengths a given error is large enough to cross the threshold, which is what the threshold has to be chosen against. If faster-whisper cannot load the CUDA libraries (`libcublas.so.12 is not found`), use `--asr-device cpu` or install `nvidia-cublas-cu12` and `nvidia-cudnn-cu12`.

## Tests

```bash
python tests/smoke_test.py
```

The smoke test mocks the TTS, ASR and translation back-ends, so no models or network access are needed (requires `numpy`, `soundfile` and FFmpeg). It covers note extraction and language detection, dictionary reading and replacement, unit normalization, term scanning, translation and the hash that marks a translation as outdated, the structured notes and their round trip, the ASR comparison scores, packing (insertion, replacement, trim and pointer removal, slide timings) and the command-line checks.

## Support and contributions

Questions and bug reports are welcome as GitHub issues. Please include the command you ran, the log output, and, if possible, a small deck that reproduces the problem (`examples/make_sample_deck.py` builds one). Do not attach reference recordings or unpublished course material.

The tool is maintained alongside teaching and research, so replies can take a while and new features are added as they become necessary for the author's own lecture material. Pull requests are welcome; small, self-contained changes are the easiest to review, and `python tests/smoke_test.py` should pass.

## Citation

If you use PPTX-Narrator, please cite the archived release on Zenodo; its DOI is added here and to [`CITATION.cff`](CITATION.cff) with the v1.0.0 release. Changes between versions are listed in [`CHANGELOG.md`](CHANGELOG.md).

## License

[MIT](LICENSE) © 2026 Hidemi Watanabe
