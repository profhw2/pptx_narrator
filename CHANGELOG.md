# Changelog

## Unreleased

### Changed
- The command line is now `pptx-narrator WS COMMAND [INPUT] [OUTPUT] [OPTIONS]`: the workspace comes first and the command second. `INPUT` and `OUTPUT` are the files outside the workspace that a command reads or writes (`extract DECK`, `scan DICT`, `pack DECK OUT`); the files inside the workspace are chosen with `--lang`, `--slides` and the model options. `--workspace`, `--out` and the positional target of `pack` are removed, as is naming one file of a workspace as `INPUT` (use `--slides`).
- Nothing is carried over from an earlier run any more: `INPUT` can no longer be left out, and `verify` and `pack` no longer take the engine and language of the last `synthesize`. `.pptx_narrator_state.json`, which served that purpose, is no longer written.
- `--lang` gives the language of the texts a command works on; it is short for giving `--in-lang` and `--out-lang` the same language.
- Existing files and notes are not overwritten unless `--update` (write what has changed) or `--overwrite` (write everything selected) is given; generated audio is the exception, since it is made from the text and not edited by hand. This applies to `extract` (texts in the workspace), `scan` (an existing dictionary needs `--append` or `--overwrite`), `translate` (`--retranslate` is replaced by `--overwrite`; `--update` translates again where the source changed, keeping translations edited by hand) and `pack` (`--forceupdate` is replaced by `--overwrite`; an existing `OUT` needs `--update` or `--overwrite`, and `OUT` must differ from `DECK`).
- `pack --data-type text|audio|all` replaces the positional target. Without `--lang`, `pack` writes the texts of every language of the workspace; a slide plays one audio, so audio of several languages needs `--lang`, and missing audio is reported rather than an error.
- Notes are written per language: a note with texts of several languages has one heading of the same form per language (`=== pptx-narrator: [en] translated from [ja] <time> #<hash>; edited <time> ===`), no language being treated specially, and a note of one language has none. For each language, the text is compared with that part of the note and with what `extract` read or `pack` wrote; the same text is not written again, so a part that is not rewritten keeps its formatting and struck-through text, and a language the note lacks is added. The text of the language whose audio the slide plays is always on top, where the presenter reads, and is moved there unchanged if needed; below it, what a run writes goes above what it leaves. Notes in the earlier layout (`narration` / `source`) are still read, and `extract` reads each part into the text of its language.
- `extract` no longer sets aside a translation whose source changed (`*.stale.txt`); `translate --update` handles that case.

### Added
- `synthesize` records the text and dictionaries each audio was made from (`audio_sources.json`); `pack` and `verify` warn when a text was edited after its audio was made.
- `translations.json` also records when a translation was made and the translation itself, so that a translation edited by hand is recognized.
- Every run appends what it logged to `pptx_narrator.log` in the workspace, ends with the commands that can come next and a report of the files it wrote, and records those files in `.pptx_narrator_history.jsonl`; `pptx-narrator WS history [--dates]` lists the runs.
- Settings that are paths are recorded relative to the workspace when inside it, and absolute otherwise.

### Fixed
- `pack` rewrote a note even when its text was unchanged, losing its formatting and struck-through text.
- A single file named as `INPUT` of `scan`, `translate`, `synthesize` or `verify` did not have its results copied back into the workspace (the feature is removed).

## 1.0.0 – 2026-09-18

First public release. Versions 1.1.0 and 1.2.0 were development numbers used before this release and are not published.

### Features
- The command line is organized as commands: `extract`, `scan`, `translate`, `synthesize`, `verify` and `pack`. Nothing runs implicitly; each command names what it does and takes its data as an argument.
- `INPUT` is a file or a directory. A directory is processed as a whole, a file on its own, so one slide is redone by naming its file. When `INPUT` is omitted, the input recorded by the previous run of that command is reused only if its SHA-256 still matches, and a changed input has to be named again (`.pptx_narrator_state.json`).
- `--in-lang` is the language of the data a command reads and `--out-lang` the language it writes; `--out-lang` belongs to `translate`. For `extract`, `--in-lang` selects the languages to extract, a note of another language is reported and left out, and a requested language the deck does not contain is an error.
- Parameters that stay the same across runs are read from a TOML file (`--config`, or `pptx_narrator.toml` in the current directory), with a `[common]` section and one section per command. Values are resolved as built-in defaults, then the configuration file, then the command line.
- Every run writes `.pptx_narrator_resolved.toml`: the effective value of every parameter, the version of the tool, the time, and the path and SHA-256 of each input. It is a valid configuration file, so a run can be repeated from it.
- What a workspace records about its own contents is written relative to the workspace, so that it can be moved or copied without invalidating its record; only the workspace itself, and inputs outside it, keep absolute paths.
- An error that comes of pointing a command at the wrong thing says what was most likely meant, as a command to run.
- `pptx-narrator --help` lists each command once, says that `pptx-narrator COMMAND --help` gives the options of one command, and shows every option under its hyphenated spelling only, the underscore spelling remaining accepted. Each command describes itself at the head of its own help, and every option has help text.
- `pptx-narrator --version` prints the version instead of the general help, and leaves with a success status, as does `--help`; running the tool with no command at all remains an error.
- `--ref-text-file` is renamed `--ref-text`, matching the brevity of `--ref-wav`.
- Every option may be abbreviated on the command line as far as it stays unambiguous (`--work` for `--workspace`).
- The underscore spelling of an option (`--in_lang`) is now accepted by normalizing the command line, instead of registering it as a second, separate option string. Registering both had made an abbreviation ambiguous against an option's own underscore spelling (`--dic` for `--dict-file` was rejected as ambiguous with `--dict_file`); normalizing removes that self-collision, and every option is abbreviated by the same rule that applies to any other pair of options that happen to share a prefix.
- `pack` now checks, for every slide it packs, whether the audio it is about to embed still reflects the deck's own note, across all three steps between them: whether the deck's own note has been edited in PowerPoint since it was last extracted (`note_sources.json`, new, records the fingerprint `extract` itself saw); whether, if the narration is a translation, the source note it was translated from was edited since without `translate --retranslate` (the existing `translations.json` record, previously consulted only under `--writeback-notes`); and whether the narration text was edited, or replaced by a later `synthesize`, after this audio was generated from it (`audio_sources.json`, new, records the fingerprint of the text an audio file came from). Each is a warning naming the slide, and the summary line lists which slides were affected; the audio is still packed, since it is the only audio there is. The first check compares the deck's live note against what `extract` last saw rather than against the workspace text file, so that hand-editing the workspace copy -- the documented way to work with it -- is never itself mistaken for staleness. All three checks run whether or not `--writeback-notes` is given, since packing already reads the workspace either way; previously, only the translation case ran, and only under `--writeback-notes` -- a deck note edited directly, or audio synthesized from since-edited text, was packed silently either way.
- `--slides` selects which slides a command works on, e.g. `--slides 4` or `--slides 1,3,5-`; it is accepted by every command, so `translate`, `synthesize` and `verify` can be re-run for part of a workspace without naming each file. A selection that matches no slide of the workspace is an error that lists the slides it has.
- Translation requests are spaced out and retried when the service refuses them for rate, instead of failing the slide.
- A configured value is taken back for one run by giving the option empty (`--dict-file ''`), and a switch set in the file by its `--no-` form (`--no-writeback-notes`); configuration keys may be written with hyphens or underscores, and `--config` may come before or after the command.
- The input recorded for a command is the input as it stands after the run, so that `translate` and `synthesize`, which write into the directory they read, can still reuse it; an edit made afterwards is still detected.
- `pack` says which workspace it is using when `--workspace` is not given, and records that workspace and its contents alongside the deck.
- `translate` stops with an error when no text in `--in-lang` is found, instead of doing nothing and reporting success.
- The record of a directory input names what each file is (note text, spoken text, audio, report, dictionary).
- Narration of the presenter notes of a PowerPoint deck with a voice-cloning TTS engine (Qwen3-TTS in-process, GPT-SoVITS through its HTTP API); the voice is defined by a few seconds of reference speech and its transcript.
- Multilingual: the note language is identified automatically (kana/Hangul rules plus py3langid), notes can be machine-translated into any language offered by Google Translate, and narration can be synthesized in any language of the selected engine.
- Dictionaries are plain string-replacement lists (`string,replacement,type`) applied to the notes before translation and to the narration text before synthesis; `--dict-file` is repeatable and `--scan` proposes candidate terms.
- Built-in SI-unit readings for Japanese narration, with an optional letter map for spelling out symbols.
- `--verify` transcribes the narration with faster-whisper and reports similarity and character error rate (kana-based for Japanese) so that listening can focus on the flagged slides.
- `--pack` embeds the audio in the structure PowerPoint writes for recorded narration, inserting a narration object where a slide has none, giving every slide its own media file, setting the slide advance time, and removing the settings of a previous recording (trim, fade, bookmarks, laser-pointer path, play/pause/seek events; `--remove-recorded`).
- `--pack` leaves a pause (default one second, `--slide-pause`) between the end of the narration and the automatic slide advance.
- `--pack` parks the audio icon next to the slide, outside the visible area, so that it does not cover the slide content in the editor (`--keep-audio-icon` to switch this off).
- `--writeback-notes` writes translated or spoken-form narration above the original note between `=== pptx-narrator: … ===` marker lines that `--extract` recognizes; `translations.json` records which source version each translation came from.
- The verification report gives, besides the ratio-based scores, the number of places where the narration and the text differ and the length of the longest such stretch.
- `--tts` reports the synthesis time per slide and a summary for the run.
- Qwen3-TTS generation is capped at the number of codec tokens the text can plausibly need, so a runaway is cut off rather than generated in full and then discarded.
- `tests/smoke_test.py` covers the pipeline with the back-ends mocked.
- `examples/make_sample_deck.py` builds a small deck with notes for trying the pipeline.
- `examples/screening_check.py`, which injects narration errors of a known size into the intended text and reports how often the ASR check notices them, by error size and note length.
- `examples/readings_ja_molbio.csv`, a working reading dictionary from the author's molecular-biology lectures.
- In a dictionary, a `;` at the start of a line or after a space starts a comment that runs to the end of the line, for annotating an entry or switching it off; a line that is only a comment is skipped. Terms containing `#` or `;` are unaffected.
- `--scan-compounds` writes the compounds of Japanese notes, with the reading a Japanese front end assembles for them, as comment lines in the dictionary; they do nothing until the reading is corrected and the `;` removed.
- `--scan` no longer proposes single letters or digits, which have no useful reading of their own.
- Workspace files always name the language of their text (`slide_3_ja.txt`, `slide_3_en.txt`, `slide_3_de.txt`, and likewise for the spoken text, the audio and the reports), instead of leaving Japanese unmarked and writing English as `_eng`. Files written by pre-release versions are still recognized when a workspace is read.
- Typographic apostrophes, primes, quotation marks and dashes in the notes (e.g. the ’ PowerPoint inserts, or ′ pasted from a paper) match a dictionary entry written with the plain ASCII character, so an entry such as `5',ごだっしゅ` applies to all of its shapes.
- Dictionary entries containing a backslash are reported, because they are matched literally and usually come from shell escaping.
