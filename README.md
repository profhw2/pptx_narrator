# PPTX-Narrator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.txt)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22796544.svg)](https://doi.org/10.5281/zenodo.22796544)

**PPTX-Narrator** turns the presenter notes of a PowerPoint deck into narration spoken in the presenter's own cloned voice, in the language of the notes or translated into another language, and writes the audio back into the deck. The whole deck keeps one consistent voice however often it is revised, even when the presenter cannot speak (e.g. because of a cold).

## Features

- **Note extraction** – presenter notes are exported per slide as editable text files; the language of each note is identified automatically; hidden slides are skipped. Nothing is overwritten unless asked, and every run is logged and reported in the workspace.
- **Translation** – notes can be translated by a language model run locally (default [Qwen3-4B](https://huggingface.co/Qwen/Qwen3-4B)); each note is translated whole, and a translation dictionary tells the model how to translate particular terms. The result is a text file to be reviewed like any other note.
- **Technical-term scanning** – acronyms, domain terms and number–unit expressions are collected into a dictionary for manual review.
- **Reading normalization** – a dictionary of string replacements applied to the narration text, plus built-in reading of SI-prefixed units for Japanese (e.g. `5 mg` → 5ミリグラム).
- **Voice cloning** – [Qwen3-TTS](https://github.com/QwenLM/Qwen3-TTS) (in-process) or [GPT-SoVITS](https://github.com/RVC-Boss/GPT-SoVITS) (via its API server). Which one sounds closer to the speaker is a matter of judgement and changes with engine versions; in the author's use Qwen3-TTS reproduces Japanese and English closely from a single Japanese reference recording, which is why it is the default; any default can be set in the configuration file.
- **ASR screening (auxiliary)** – audio can be transcribed with `faster-whisper` and compared with the intended text (kana level via `pyopenjtalk` for Japanese, normalized characters otherwise); per-slide similarity, the longest single stretch of disagreement and the character error rate (CER) flag the slides most likely to be misread.
- **PPTX repackaging** – the narration audio is embedded in each slide (replacing earlier audio or inserted as a new narration object) and the automatic slide advance time (`advTm`) is set to the audio duration plus a short pause, so the deck plays as a self-running show and can be exported as a video. The texts are written back into the notes, all languages of a slide under headings of one form when there are several, the text of the audio on top; a part that is not rewritten keeps its formatting, and later extractions read each part back.

## Supported languages

| Step | Languages |
|---|---|
| Note language identification | Kana/Hangul rules and [py3langid](https://github.com/adbar/py3langid) |
| `translate` | Whatever the translation model handles (Qwen3: over 100 languages and dialects according to its model card) |
| `synthesize --engine qwen3` | zh, en, ja, ko, de, fr, ru, pt, es, it |
| `synthesize --engine gpt_sovits` | zh, en, ja, ko, yue (Cantonese) |
| `verify` | Japanese: kana-level comparison; other languages: character-level comparison of normalized text |
| Built-in unit readings | Japanese only (other languages: `unit` entries in the dictionaries) |

Languages are written as codes such as `ja`, `en`, `de`, `zh-CN`. Without `--lang`, `extract` identifies the language of each note from its text; notes too short or ambiguous to identify reliably ("Thank you.", a kanji-only title) are assigned the main language of the deck. The detected language is shown in the log and in the file name. `--lang` selects which languages to extract: a note in another language is reported and left out, and a requested language the deck does not contain stops the run with an error.

## Requirements

- Python 3.10 or later
- [FFmpeg](https://ffmpeg.org/) on `PATH` (used by `pydub` to write `.m4a`)
- One TTS engine:
  - **Qwen3-TTS**: installed with the `qwen3` extra below. Model weights (`Qwen/Qwen3-TTS-12Hz-{0.6B,1.7B}-Base`) are downloaded from Hugging Face on first use. A CUDA GPU or Apple Silicon (MPS) is strongly recommended.
  - **GPT-SoVITS**: a local GPT-SoVITS installation with its API server running, e.g. `python api_v2.py -a 127.0.0.1 -p 9880` in the GPT-SoVITS directory.
- For `translate`: the translation model (default `Qwen/Qwen3-4B`, about 8 GB) is downloaded from Hugging Face on first use; `transformers` 4.51 or later. A CUDA GPU or Apple Silicon (MPS) is strongly recommended.

## Installation

```bash
git clone https://github.com/profhw2/pptx_narrator.git
cd pptx_narrator
pip install -e .            # pipeline without a local TTS engine
pip install -e ".[qwen3]"   # + Qwen3-TTS, the default engine
pip install -e ".[verify]"  # + ASR verification
pip install -e ".[all]"     # everything
```

This installs the `pptx-narrator` command. Running `python src/pptx_narrator.py ...` without installing also works (`pip install -r requirements.txt`).
The `scan` command downloads the NLTK `stopwords` and `words` corpora on first use; if that fails behind a proxy, run `python -m nltk.downloader stopwords words`.

## Setting up the TTS engines

### Qwen3-TTS (default)

`pip install -e ".[qwen3]"` installs the `qwen-tts` package and PyTorch. With an NVIDIA GPU, install the PyTorch build that matches your CUDA version first (see pytorch.org). The model weights (`Qwen/Qwen3-TTS-12Hz-0.6B-Base` or `-1.7B-Base`, chosen with `--qwen3-model-size`) are downloaded from Hugging Face on first use. `--qwen3-device` selects `cuda:0`, `mps` (Apple Silicon) or `cpu`; the default `auto` uses `cuda:0` if available, otherwise `mps`, otherwise `cpu`.

### GPT-SoVITS

GPT-SoVITS runs as a separate program, which PPTX-Narrator reaches through its HTTP API server.

1. Install GPT-SoVITS by following its own instructions (https://github.com/RVC-Boss/GPT-SoVITS), preferably in an environment of its own, and download its pretrained models. The weights that `--model` selects must be in the GPT-SoVITS directory: `v2ProPlus` (the default) uses `GPT_SoVITS/pretrained_models/s1v3.ckpt` and `GPT_SoVITS/pretrained_models/v2Pro/s2Gv2ProPlus.pth`; `v4` and `v1_clear` use the files listed in `MODELS_CONFIG` in `src/pptx_narrator.py`.
2. Start the API server in the GPT-SoVITS directory: `python api_v2.py -a 127.0.0.1 -p 9880`.
3. Run `synthesize --engine gpt_sovits`; `--api-url` gives the server's address (default `http://127.0.0.1:9880/`). PPTX-Narrator switches the server to the weights of `--model` before synthesizing.

The reference recording is passed to the server by its absolute path, and the server reads the file itself. When the server runs on another computer, the file given as `--ref-wav` must therefore exist at the same absolute path on that computer (a shared folder, or a copy at the same path).

The GPT-SoVITS API server has no authentication of its own and, by default, accepts connections only from the same computer (127.0.0.1). Using it from another computer securely, for example over an SSH tunnel, is the user's responsibility; consult your network administrator where needed.

## Tab completion

With [shtab](https://github.com/iterative/shtab) installed (`pip install -e ".[completion]"`, or `pip install shtab`), `pptx-narrator --print-completion SHELL` prints a completion script, so that the Tab key offers the workspace directory, the commands, their options, and the files each argument takes (`.pptx` for a deck, `.csv` for a dictionary, `.wav` and `.txt` for the reference voice). Write the script once and let the shell read it; run the same commands again after updating PPTX-Narrator. Completion works with the installed `pptx-narrator` command, not with `python src/pptx_narrator.py`.

`pptx-narrator --completion-setup` shows the commands to run and the lines to add to the configuration of the current shell (or of the shell given, e.g. `--completion-setup bash`), with where they typically go; it changes nothing itself. The same steps, per shell:

- zsh (macOS):
  ```bash
  mkdir -p ~/.zfunc
  pptx-narrator --print-completion zsh > ~/.zfunc/_pptx-narrator
  ```
  and add to `~/.zshrc`, before any existing `compinit`:
  ```bash
  fpath=(~/.zfunc $fpath)
  autoload -Uz compinit && compinit
  ```
- bash (Linux; needs the bash-completion package):
  ```bash
  mkdir -p ~/.local/share/bash-completion/completions
  pptx-narrator --print-completion bash > ~/.local/share/bash-completion/completions/pptx-narrator
  ```
- PowerShell (Windows): add `pptx-narrator --print-completion powershell | Out-String | Invoke-Expression` to your profile (`notepad $PROFILE`).
- fish: `pptx-narrator --print-completion fish > ~/.config/fish/completions/pptx-narrator.fish`

Open a new terminal afterwards.

## Quick start

To try the tool without preparing a deck, build the sample deck first; its notes contain acronyms, gene names and number-unit expressions in Japanese and English:

```bash
python examples/make_sample_deck.py sample_lecture.pptx
```

Any few seconds of clear speech with its exact transcript can serve as the reference voice (`--ref-wav` / `--ref-text`).

The pipeline is split into steps so that text can be reviewed before synthesis. Intermediate files live in a workspace directory. A dictionary (`--dict-file`) is a list of string replacements applied to the text processed in the run: to the notes before translation and to the narration text before synthesis.

### Narration in the language of the notes

```bash
# 1. Extract the notes into the workspace ws, then collect candidate readings into readings_ja.csv
pptx-narrator ws extract lecture.pptx --lang ja
pptx-narrator ws scan readings_ja.csv --lang ja

# 2. Review by hand: ws/slide_N_ja.txt (the texts) and readings_ja.csv (readings)

# 3. Synthesize with a cloned voice, then screen the result
pptx-narrator ws synthesize --lang ja --dict-file readings_ja.csv \
  --engine qwen3 --ref-wav my_voice.wav --ref-text my_voice.txt
pptx-narrator ws verify --lang ja --engine qwen3

# 4. Listen to the slides flagged in ws/verify_report_ja.qwen3-1.7B.csv, fix the
#    dictionary or the texts, synthesize those slides again (--slides 4,7), then pack
pptx-narrator ws pack lecture.pptx lecture_narrated.pptx --lang ja --engine qwen3
```

The workspace comes first and the command second; the command's options follow it, in any order. Every run ends with the commands that can come next, ready to copy, and a report of the files it wrote; `pptx-narrator ws history` lists what has been run in the workspace.

### Translated narration

Translation is optional. A translated text is written next to the text it came from and is then a text like any other: it is reviewed, synthesized and packed in the same way. Translation is usually done once, while synthesis is repeated while readings are refined, so each step has its own dictionary:

```bash
# 1. Extract the notes (language identified automatically)
pptx-narrator ws extract lecture.pptx
# 2. Translate; terms_ja_en.csv (optional) lists terms and how to translate them
pptx-narrator ws translate --in-lang ja --out-lang en --dict-file terms_ja_en.csv
# 3. Review and correct ws/slide_N_en.txt; collect candidate readings from the English text
pptx-narrator ws scan readings_en.csv --lang en
# 4. Review readings_en.csv, then synthesize, verify and pack (repeat as needed)
pptx-narrator ws synthesize --lang en --dict-file readings_en.csv \
  --engine qwen3 --ref-wav my_voice.wav --ref-text my_voice.txt
pptx-narrator ws verify --lang en --engine qwen3
pptx-narrator ws pack lecture.pptx lecture_en.pptx --lang en --engine qwen3
```

Each note is given to the translation model whole, so every sentence is translated in the context of the note. The note is passed unchanged; the entries of the translation dictionary that occur in it are given to the model as instructions ("translate X as Y"). An existing translation is kept: `--update` translates again only where the source text has changed since (a translation edited by hand is still kept), and `--overwrite` translates again every selected slide.

`examples/` has the dictionaries of this example (`terms_ja_en.csv`, `readings_en.csv`) and, for another target language, German ones (`terms_ja_de.csv`, `readings_de.csv`); any language of the TTS engine is handled in the same way.

Once written, a translated text is a note text like any other: review and correct it before synthesis. Machine translation, including by language models, makes mistakes that read fluently — a technical term rendered as a similar-sounding word, a sentence whose meaning is reversed, a detail left out.

#### Translation models

`--translate-model` takes the Hugging Face id of the model (default `Qwen/Qwen3-4B`). Which one to use depends on the machine and on how much correction the output may need:

| Choice | Notes |
|---|---|
| `Qwen/Qwen3-4B` (default) | Runs on a laptop with Apple Silicon or a modest GPU (about 8 GB of weights). |
| `Qwen/Qwen3-8B`, `Qwen/Qwen3-14B`, `Qwen/Qwen3-32B` | Larger models of the same family; more memory and time, and usually fewer errors. |
| Another instruction-tuned model with a chat template | May work through the same option; not tested. |
| A hosted service (DeepL, Google Cloud Translation, a commercial language model, …) | Not built in. Translate the notes there and save the results as `ws/slide_N_<lang>.txt`; the rest of the pipeline uses them like any other note. |

### Workspace files

Files made for each workspace:

- `slide_N_<lang>.txt`: the text of a slide, in the language named in the file name (`slide_3_ja.txt`, `slide_3_en.txt`, `slide_3_de.txt`). Texts are treated alike whether they were extracted, translated or written by hand.
- `slide_N_<lang>.<model>.m4a`: generated audio.
- `slide_N_<lang>.<model>.spoken.txt`: the text after dictionary replacement, as sent to the TTS engine.
- `verify_report_<lang>.<model>.csv` and `verify_differences_<lang>.<model>.csv`: the verification report (see below).
- `note_baseline.json`: the fingerprint of each note as `extract` read it and as `pack` wrote it, per language, so that an edit in the deck or in the workspace can be told apart.
- `translations.json`: which version of the source text each translation was made from, and when.
- `audio_sources.json`: which version of the text (and of the dictionaries) each audio file was made from.
- `slide_map.json`: the number and ID of each slide of the deck the workspace follows (see below); `map_archive/`: what `map --apply` set aside.
- `pptx_narrator.log`: everything each run logged, appended run by run.
- `.pptx_narrator_history.jsonl` and `.pptx_narrator_resolved.toml`: the record of the runs (see Configuration).

Paths inside the workspace are recorded relative to it, so a workspace can be moved or copied as a whole.

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

- With `translate`, the entries that occur in a note are given to the model as instructions (e.g. `塩基対,base pairs`: translate 塩基対 as base pairs); the note itself is passed unchanged. With `synthesize`, the entries replace strings in the narration text before synthesis (readings). The note files themselves are not changed; the rewritten narration is saved as `*.spoken.txt`.
- Longer strings are replaced first; alphanumeric strings only match on word boundaries; rows with an empty replacement are ignored.
- Case: a string that is an ordinary word (three letters or more, lower case after the first letter) matches with its first letter in either case, so `Prophase I` also covers `prophase I` and `transcription` also `Transcription` at the start of a sentence. Acronyms, mixed-case and short strings (`DNA`, `mRNA`, `Mg`), whose case can carry meaning, and unit symbols match only as written. Full-width digits and letters in the text (４, ＤＮＡ) match entries written in the usual form.
- `type` = `unit` marks unit symbols that follow a number (`3 Gbp`). For Japanese narration, built-in rules additionally read SI-prefixed units.
- `--dict-file` can be given several times (e.g. a shared and a deck-specific file); later files take precedence.
- `scan` writes new candidates into the dictionary named on its command line: a new file, or with `--append` an existing one (with `--overwrite` it is made again from scratch). It scans the text in `--lang`, filling in provisional readings when that is the narration language and leaving the replacements blank when the text is still to be translated. With `--scan-compounds` and Japanese narration, it additionally writes every compound of the notes, with the reading a Japanese front end assembles for it, as **comment lines** (`;二本鎖,ニホンクサリ,`). Comments do nothing, so the list can be long; compounds are where readings are unsettled (鎖 is クサリ alone but サ in 二本鎖), and an entry is activated by correcting the reading and removing the `;`. Needs `pyopenjtalk` (the `verify` extra).

A four-column dictionary (`Term,Japanese_Reading,English_Reading,Type`) from an earlier development version can still be given with `--dict-file`; the column of the narration language (Japanese or English) is used. See the files in [`examples/`](examples/); [`examples/readings_ja_molbio.csv`](examples/readings_ja_molbio.csv) is a working dictionary from the author's molecular-biology lectures (104 entries: gene and organism names, researchers' names, units and Japanese words the engine misreads), which can be used as a starting point for that field. A dictionary belongs to a subject area rather than to a deck: once the terms of a course are in it, later decks need few new entries. An entry with an empty replacement leaves the term as it is and stops `scan` from proposing it again; lone letters and digits are never proposed. A `;` at the start of a line or after a space begins a comment that runs to the end of the line, so an entry can be annotated or switched off without deleting it, and a line that is only a comment is skipped. A `;` inside a term is kept (`A;B`); a term that begins with `;` is written in double quotes. Unit symbols not resolved otherwise can be spelled out letter by letter with a letter map for the narration language, e.g. `--letter-map examples/letter_map_ja.json`.

## Packing and notes write-back

`pack` writes into a copy of the deck: `pptx-narrator ws pack DECK OUT` reads `DECK` and saves `OUT`, which must be a different file; `DECK` is never changed. An existing `OUT` is replaced only with `--update` or `--overwrite`. `--data-type text|audio|all` chooses what is written (default `all`).

Audio is embedded in the same structure PowerPoint uses for recorded narration: the audio starts with the slide, is hidden during the show, and the slide advances automatically after the audio length. Slides without audio get a new narration object (a small speaker icon, visible only in the editor). If a slide already has audio (e.g. an earlier recording), that object is reused: it points to the new audio, and its trim, fade and bookmarks are removed. Audio is always written when it is asked for, since it is made from the text and never edited by hand; with `--update`, a slide that already plays the same audio is left as it is. Each slide gets its own media file, so copied slides that shared one clip no longer overwrite each other, and clips no longer used are removed from the file. Slides that have animations but no audio are reported and left unchanged; insert any audio clip on such a slide in PowerPoint and pack again. The slide advances by itself one second after the narration ends, so that the last word is not clipped in a video; `--slide-pause` changes that pause. The audio icon of a narrated slide is parked next to the slide, outside the visible area, so that it does not cover the slide content while editing; `--keep-audio-icon` leaves it where it is. The icon is hidden during the slide show either way.

A slide plays one audio. When the workspace holds audio of several languages, `--lang` chooses it; when it holds none of the chosen language, `pack` says so and writes the texts only (a slide can then be recorded in PowerPoint). If a text was edited after its audio was made, `pack` warns that the audio does not say the text.

Data recorded together with the old audio is also removed from a narrated slide, because its timing belongs to that audio: the laser-pointer path (`p14:laserTraceLst`) and the recorded play/pause/seek events (`p14:showEvtLst`). `--remove-recorded pointer|events|none` narrows or disables this (default: `all`). Ink annotations are kept and reported: ink drawn during a recording cannot be told apart from ink drawn while editing, which belongs to the slide, so the present version removes neither. The packed deck can be exported as MP4 with PowerPoint's *Export* (use recorded timings and narrations).

The texts of the workspace are written into the notes. Without `--lang`, every language the workspace has for a slide is written; with `--lang`, only that one. For each language, `pack` compares the text of the workspace with that part of the note in the deck and with what `extract` read or `pack` last wrote:

- the same text is not written again, so the note keeps its formatting and its struck-through text;
- a text edited in the workspace is written with `--update` or `--overwrite`;
- a note edited in the deck since (or one `extract` never read) is written over only with `--overwrite`;
- a language the note does not have yet is added.

The text of the language whose audio the slide plays is always on top, since it is what the presenter reads; `pack` moves it there, unchanged, if needed. Below it, newer parts are above older ones: what a run writes goes above what it leaves.

Without the option a text needs, `pack` leaves that part of the note and says why. When a note holds texts of several languages, each is preceded by a heading of the same form; a note of one language has none:

```
=== pptx-narrator: [en] translated from [ja] 2026-09-25T14:02 #ae82d4f1fc ===
Today we talk about genome editing.

=== pptx-narrator: [ja] ===
今日はゲノム編集について話します。
```

The heading says how a text was made: `translated from` the source language, when, and the fingerprint of the source version; `edited` and the time, when the text was changed after it was made. A part that is not rewritten keeps its formatting. Keep the heading lines when editing such notes in PowerPoint: `extract` reads each part into the text of its language. Notes written by earlier versions (`narration [..] from [..]` / `source [..]`) are still read.

`extract` follows the same rule: a text already in the workspace is replaced only with `--update` (when the note in the deck changed and the text was not edited in the workspace) or `--overwrite`; otherwise it is kept, with a warning.

## Slides inserted, deleted or reordered

The texts, audio and records of a workspace are kept by slide number, and inserting, deleting or reordering slides in PowerPoint changes the numbers of the slides after them. PowerPoint gives every slide an ID that stays the same when other slides move, and `extract`, `pack` and `map` record these IDs. Before they read or write a deck, `extract` and `pack` check that its slides still correspond to the workspace; if they do not, they stop without changing anything.

`pptx-narrator ws map lecture.pptx` shows how the slides of the deck correspond to the workspace: which slides moved, which are no longer in the deck (deleted, or hidden), and which are new. Nothing is changed. With `--apply`, the workspace is renumbered to follow the deck: the texts, audio and records of moved slides get their new numbers, and the files of slides no longer in the deck, together with the log and the ASR reports that speak of the old numbers, are set aside in `map_archive/<date_time>/` with a `mapping.txt` that lists what was done. Nothing is deleted. The new slides are then taken in with commands that `map` suggests (`extract --slides …` and so on).

A workspace made before slide IDs were recorded is matched by the fingerprints of its notes, and what is left by the similarity of the texts. To write the narration into another version of the deck, such as one with the slide text in another language, make that version with *Save As*, which keeps the slide IDs.

## Command-line reference

```
pptx-narrator WS COMMAND [INPUT] [OUTPUT] [OPTIONS]
```

`WS` is the workspace directory of one deck. `INPUT` and `OUTPUT` are the files outside the workspace that a command reads or writes; the files inside it are chosen with `--lang`, `--slides` and the model options, not by path. Nothing is carried over from an earlier run: what a command needs is given on its command line or in the configuration file. Existing files, notes and audio are not overwritten unless `--update` (write what has changed) or `--overwrite` (write everything selected) is given; `pack` writes the audio of the workspace into the deck it is told to write. `--dry-run` (extract, scan, translate, synthesize, pack) shows what a command would write, slide by slide with the reason, and changes nothing; `map` shows its changes unless `--apply` is given.

`--lang` gives the language of the texts a command works on; it is short for giving `--in-lang` and `--out-lang` the same language. `translate` reads `--in-lang` and writes `--out-lang`. Every command accepts `--slides` (e.g. `4,7-9`).

- `extract DECK`: the notes of `DECK` into the workspace (created if needed). `--lang` (languages to extract, e.g. `ja` or `ja,en`; omitted = every language found), `--update`, `--overwrite`.
- `scan DICT`: candidate terms of the texts into the dictionary `DICT`. `--lang`, `--append` (add to an existing `DICT`), `--overwrite` (make it again), `--dict-file` (other dictionaries whose terms need not be proposed again), `--scan-compounds`.
- `translate`: `--in-lang`, `--out-lang`, `--dict-file` (translation dictionary), `--translate-model`, `--translate-device`, `--update`, `--overwrite`.
- `synthesize`: `--lang`, `--dict-file` (readings), `--letter-map`, `--ref-wav`, `--ref-text` (a text file with the transcript), `--ref-lang`, `--engine {gpt_sovits,qwen3}` (default qwen3), `--model`, `--qwen3-model-size {0.6B,1.7B}` (default 1.7B), `--qwen3-device`, `--api-url`, `--enable-drc`, `--drc-threshold`, `--drc-ratio`. Only slides without audio are synthesized unless `--update` (slides whose reading changed since their audio was made: the text was edited, or a dictionary changed how it is read; add `--edited-texts-only` to leave out the latter) or `--overwrite` (every selected slide) is given. `--dry-run` lists the slides that would be synthesized, with the reason, and synthesizes nothing. With Qwen3-TTS, the sentences of a paragraph (text separated by a blank line; one sentence per line is still one paragraph) are synthesized together, up to `--chunk-chars` characters (default 200; a longer paragraph is split between sentences, a longer sentence is synthesized on its own), so that the model reads them as connected speech and a sentence end inside is not cut short. Each chunk is followed by a pause of fixed length: `--paragraph-pause` (default 0.5 s) after a paragraph, `--sentence-pause` (default 0.5 s) where a long paragraph was split. `--chunk-chars 0` synthesizes one sentence at a time.

Japanese with Qwen3-TTS: the model reads the characters themselves without dividing the sentence into words, so a character whose reading depends on its word (中 in 授業中 and 一日中, 内, 外, 毎, …) and a number before a counter (3割) are often misread. The reading assist lets OpenJTalk (pyopenjtalk) divide the text as written into words, before the dictionaries replace terms by their readings, and writes in katakana as read every suffix together with the word before it (授業中 → ジュギョウチュウ, 学生数 → ガクセイスウ) and the other words containing such a character (一日中 → イチニチジュウ), and numbers with their counters as kanji numerals where they are read as written (7割 → 七割) or in katakana as pronounced where the sound changes or the word has its own reading (三日目 → ミッカメ, 10分後 → ジュップンゴ); a number before a unit symbol becomes kanji numerals after the unit reading (3.5 mL → 三点五ミリリットル); Latin-script words such as `Cas9` are left as they are. It is on by default; `--no-reading-assist` turns it off, and `--reading-assist-chars` sets the characters (default `中内外毎間上下前後目方所分`). Entries in a reading dictionary take precedence, since they are applied first. A changed pause applies to audio made again with `--overwrite`.
- `verify`: `--lang`, `--engine`, `--model`, `--qwen3-model-size`, `--asr-model`, `--asr-device`, `--verify-threshold` (default 0.85), `--min-difference` (default 4), `--max-difference` (default 40; `0` disables), `--cer-threshold` (default off).
- `pack DECK OUT`: `--lang`, `--data-type {text,audio,all}` (default all), `--engine`, `--model`, `--qwen3-model-size`, `--update`, `--overwrite`, `--slide-pause` (default 1.0 s), `--keep-audio-icon`, `--remove-recorded {all,pointer,events,none}` (default `all`).
- `map DECK`: how the slides of `DECK` correspond to the workspace; `--apply` renumbers the workspace to follow `DECK`.
- `history`: the commands run in the workspace; `--dates` adds when.

The options of a command come after the command, in any order. `--config FILE` and `--version` may also come before it. Run `pptx-narrator WS COMMAND --help` for the full list. Underscore spellings (`--dict_file`, `--in_lang`, …) are also accepted, and any option may be typed as short as it stays unambiguous (`--data text` for `--data-type text`; `--ref-t my_voice.txt` for `--ref-text`, since `--ref-wav` and `--ref-lang` also start with `--ref-`).

## Configuration

Parameters that stay the same from run to run go into a TOML file, so that a command line carries only what changes.
`--config FILE` names it; otherwise `pptx_narrator.toml` in the current directory is used when it exists. A `[common]`
section applies to every command, and a section named after a command applies to that command and overrides `[common]`:

```toml
[common]
lang = "ja"
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

With that file, step 3 of the quick start is `pptx-narrator ws synthesize` and `pptx-narrator ws verify`.

Keys may be written with hyphens or underscores (`dict-file` and `dict_file` both work), as on the command line.

Values are resolved in one order: **built-in defaults → configuration file → command line**, the command line winning.
To take a configured value back for one run, give it empty (`--dict-file ''`); to turn off a switch the file sets, use
its `--no-` form (`--no-update`). `--config` may be written before or after the command.
After every run the result is written to `.pptx_narrator_resolved.toml`: every parameter with the value actually used,
the version of the tool, the time, and the path, size and SHA-256 of each input file. That file is itself a valid
configuration file, so a run can be repeated later from it, and it records what produced a given narration.
Each run is also appended to `.pptx_narrator_history.jsonl` (the settings, the files read, the files written), which
`pptx-narrator ws history` lists, and what it logged to `pptx_narrator.log`. Nothing in these records is read back to
fill in a later command.

### Verification report

The step writes three files. `verify_differences_<lang>.<model>.csv` lists every place where the narration and the transcript disagree, longest first: the slide, the length, the position, what the text said, what the ASR heard, and a few characters of context on each side. This is the list to read: it says where to listen, often what happened, and it can be sorted or filtered as you like; `--min-difference` sets how short a difference is still listed (default 4 characters). `verify_report_<lang>.<model>.csv` summarizes the same comparison per slide, worst first, and marks a slide `FLAGGED` when its similarity falls below `--verify-threshold` or when one stretch of disagreement is longer than `--max-difference` characters (default 40): the first catches a small error in a short note, the second a dropped phrase in a note of any length. Neither is a pass/fail test; many low-scoring slides sound natural and only reflect recognition errors. The report lists, worst first: `slide`, `similarity` (difflib ratio, 0–1), `cer` (Levenshtein distance / length of the intended sequence), `status` (`OK` or `FLAGGED`), `differences` and `longest_difference` (how many places differ and how long the longest stretch is, which is what distinguishes a skipped phrase from scattered recognition differences), `recognizer_differences` (how many were put down to the recognizer, see below), the intended and recognized text, and the two normalized sequences that were compared (katakana for Japanese; case-folded text without punctuation or spaces otherwise).

For Japanese narration, a difference is put down to the recognizer rather than to the narration, and listed in `verify_asr_differences_<lang>.<model>.csv` with its kind instead, when it falls on a Latin-script word of the narration (such a word becomes letter names in katakana, ピーアールオー…, while the recognizer writes what it heard, プロフェーズ) or when the recognizer wrote the words with kanji that the narration does not have (聖正 for 精製, 線粒体 for 染色体), which are then read in their own way, or when it wrote a number in digits, the same number as in the text of the slide, whose digits are read otherwise (２３つい read ニジューミッツイ). Places at most two characters apart are judged together, and a place within three characters of a Latin-script word counts as that word's. Before comparing, kana that sound alike are written alike (a long vowel written with イ or ウ as ー, ヲ as オ). The similarity, the CER, the longest difference and the flag count the remaining differences only, so that `verify_differences` lists the places most likely to be the narration's errors. The rule can misjudge a real misreading that the recognizer wrote in other kanji, which is why those differences are listed rather than dropped.

## Limitations

- Kana comparison cannot detect pitch-accent errors, and character comparison cannot detect prosody errors. ASR errors, and numbers or units written differently by the ASR (e.g. "5 mg" vs. "five milligrams"), can cause false flags; flagged slides should be checked by listening.
- Automatic language identification can fail for notes mixing several languages or for very short notes in decks without other notes; check the file names after `extract` or give `--lang`.
- Machine translation should be reviewed before synthesis; a term given to the model as an instruction can still be rendered otherwise.
- Writing the notes and embedding the audio edit the slide XML directly and rely on python-pptx, including one of its internal functions; a change in python-pptx or in the file format may require changes to the tool.
- Voice cloning should only be used with the consent of the speaker whose voice is cloned.
- The tool is designed not to overwrite or delete anything the user has made unless asked, and it never changes the input deck; still, as with any software that edits files, keep a copy of the decks and workspaces you cannot afford to lose.

## Planned work

- Narration of slides that have animations but no audio object (the audio has to be placed in the slide's animation timeline).
- Writing into a deck whose slides differ from those of the workspace, only the slides matched by slide ID, when asked with an option of `pack`.
- Translation with the whole deck as context, which favours a free translation that keeps the meaning; the translation dictionary would then serve mainly to keep terminology consistent.
- Comparison on the reading for Chinese (e.g. pinyin), as is done with katakana for Japanese, so that simplified or traditional characters and homophones written by the ASR are not counted as differences.

## Checking the ASR screening

`examples/screening_check.py` measures which narration errors the check actually notices. The transcript of a slide depends only on its audio, so it is produced once (or read from an existing report) and any number of hypothetical errors can then be scored against it. The script injects one error of a known size into the sequence the check compares (katakana for Japanese, normalized text otherwise) -- a run of characters deleted, as when a phrase is skipped, or replaced by other characters, as when a term is misread -- and reports how often the check notices, by the size of the error and the length of the note. Sizes are in characters of that sequence, so 3 characters is about one short term and 50 is about one sentence:

```bash
python examples/screening_check.py ws --lang ja --asr-model small --asr-device cuda
```

The workspace is not modified; `--sizes` and `--repeats` set the error sizes in characters and the number of random positions per slide and size. The result says for which note lengths a given error is large enough to cross the threshold, which is what the threshold has to be chosen against. If faster-whisper cannot load the CUDA libraries (`libcublas.so.12 is not found`), use `--asr-device cpu` or install `nvidia-cublas-cu12` and `nvidia-cudnn-cu12`.

### Tuning the check

How well the check works depends on the TTS engine, the ASR model and the material, so its criteria are options of `verify`, to be chosen on one's own decks:

- `--verify-threshold` (default 0.85) looks at the share of the slide that agrees with the transcript (similarity) and flags a slide below it. It works on short notes, where even a small error lowers the share noticeably, but in a long note an error is diluted and almost never lowers it enough: on a test deck, an error of 25 characters was caught in 180 of 200 trials on notes below 200 characters and in none of 480 on notes above 600.
- `--max-difference` (default 40 characters; `0` disables it) looks at the longest single stretch of disagreement and flags a slide above it. It catches a dropped, repeated or invented phrase whatever the length of the note, and so covers the long notes that the similarity misses; the two are used together, and a slide is flagged if either applies. A value just above the longest stretch found on correct slides flags such errors without flagging correct slides.
- `--cer-threshold` (default: not used) looks at the character error rate (the character edits that turn the transcript into the text, divided by the length of the text). Being a share as well, it is diluted in long notes in the same way.
- `--min-difference` (default 4 characters) does not change which slides are flagged; it is the shortest difference listed in `verify_differences_*.csv`, so that small differences in how the ASR writes a word are left out of the list.
- `--asr-model` (default `small`) chooses the faster-whisper model: a larger one makes fewer recognition errors, and so raises fewer false flags, but takes longer.

`examples/screening_check.py` shows, on a given deck, which errors each setting catches.

## Tests

```bash
python tests/smoke_test.py
```

The smoke test mocks the TTS, ASR and translation back-ends, so no models or network access are needed (requires `numpy`, `soundfile` and FFmpeg). It covers note extraction and language detection, dictionary reading and replacement, unit normalization, term scanning, translation and the hash that marks a translation as outdated, the notes written per language and their round trip, the ASR comparison scores, packing (insertion, replacement, trim and pointer removal, slide timings) and the command-line checks.

## Support and contributions

Questions and bug reports are welcome as GitHub issues. Please include the command you ran, the log output, and, if possible, a small deck that reproduces the problem (`examples/make_sample_deck.py` builds one). Do not attach reference recordings or unpublished course material.

The tool is maintained alongside teaching and research, so replies can take a while and new features are added as they become necessary for the author's own lecture material. Pull requests are welcome; small, self-contained changes are the easiest to review, and `python tests/smoke_test.py` should pass.

## Citation

If you use PPTX-Narrator, please cite the archived release on Zenodo: [10.5281/zenodo.22796544](https://doi.org/10.5281/zenodo.22796544) cites all versions and resolves to the latest; the DOI of each version is on that page and in [`CITATION.cff`](CITATION.cff). Changes between versions are listed in [`CHANGELOG.md`](CHANGELOG.md).

## License

[MIT](LICENSE.txt) © 2026 Hidemi Watanabe
