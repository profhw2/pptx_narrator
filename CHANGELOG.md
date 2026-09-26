# Changelog

## 1.0.0 – 2026-09-26

First public release. Versions 1.1.0 and 1.2.0 were development numbers used before this release and are not published.

### Command line
- The command line is `pptx-narrator WS COMMAND [INPUT] [OUTPUT] [OPTIONS]`: the workspace of one deck comes first, then one of the commands `extract`, `scan`, `translate`, `synthesize`, `verify`, `pack` and `history`. `INPUT` and `OUTPUT` are the files outside the workspace that a command reads or writes (`extract DECK`, `scan DICT`, `pack DECK OUT`); the files inside the workspace are chosen with `--lang`, `--slides` and the model options, not by path.
- Nothing runs implicitly and nothing is carried over from an earlier run: what a command needs is given on its command line or in the configuration file.
- Nothing that exists is overwritten unless `--update` (write what has changed) or `--overwrite` (write everything selected) is given; generated audio is the exception, since it is made from the text and not edited by hand. `scan` adds to an existing dictionary only with `--append`, and `pack` never changes `DECK` and replaces an existing `OUT` only on request.
- `--lang` gives the language of the texts a command works on and is short for giving `--in-lang` and `--out-lang` the same language; `translate` reads `--in-lang` and writes `--out-lang`. For `extract`, `--lang` selects the languages to extract; a note of another language is reported and left out, and a requested language the deck does not contain is an error.
- Every command accepts `--slides` (e.g. `4,7-9`).
- Every run ends with the commands that can come next, ready to copy, and a report of the files it wrote; what it logged is appended to `pptx_narrator.log` in the workspace, and `history` lists the runs of a workspace (`--dates` adds when).
- A command written before the workspace, or a misspelt command, is answered with the correct form.
- `pptx-narrator --help` lists each command once and says how to see the options of one; every option has help text and is listed under its hyphenated spelling, the underscore spelling (`--in_lang`) being accepted too. Every option may be abbreviated as far as it stays unambiguous. `--version` prints the version.

### Configuration and records
- Parameters that stay the same across runs are read from a TOML file (`--config`, or `pptx_narrator.toml` in the current directory), with a `[common]` section and one section per command. Values are resolved as built-in defaults, then the configuration file, then the command line. A configured value is taken back for one run by giving the option empty (`--dict-file ''`), a switch by its `--no-` form; keys may be written with hyphens or underscores.
- Every run writes `.pptx_narrator_resolved.toml` (the effective value of every parameter, the version of the tool, the time, and the path and SHA-256 of each input; itself a valid configuration file) and appends to `.pptx_narrator_history.jsonl` (the settings, the files read and the files written). These records are never read back to fill in a later command.
- Paths inside the workspace are recorded relative to it, so a workspace can be moved or copied as a whole; paths outside it are absolute.

### Notes and texts
- Presenter notes are extracted per slide as editable text files named after the slide and the language (`slide_3_ja.txt`, `slide_3_en.txt`); the language is identified automatically (kana/Hangul rules plus py3langid). Hidden slides, slide-number and date placeholders and zero-width characters are left out, and so is struck-through text, which the author has deleted.
- Texts are treated alike whether they were extracted, translated or written by hand.
- `extract` replaces a text already in the workspace only with `--update` (when the note in the deck changed and the text was not edited in the workspace) or `--overwrite`.
- Typographic apostrophes, primes, quotation marks and dashes in the notes match a dictionary entry written with the plain ASCII character.

### Translation
- Optional: a text can be translated by an instruction-tuned language model run locally (default Qwen/Qwen3-4B, chosen with `--translate-model`), each note whole; the entries of a translation dictionary that occur in it are given to the model as instructions. The result is a text like any other, to be reviewed before synthesis.
- An existing translation is kept; `--update` translates again where the source changed (keeping a translation edited by hand), `--overwrite` translates again. `translations.json` records which version of the source each translation was made from, when, and the translation as made.

### Dictionaries and readings
- Dictionaries are plain string-replacement lists (`string,replacement,type`) applied to the notes before translation and to the text before synthesis; `--dict-file` is repeatable, later files taking precedence. Longer strings are replaced first, alphanumeric strings only on word boundaries, and entries with an empty replacement do nothing.
- A `;` at the start of a line or after a space starts a comment; entries containing a backslash are reported.
- `scan` proposes candidate terms (acronyms, Latin and katakana words, number–unit expressions, Roman numerals), never single letters or digits; `--scan-compounds` writes the compounds of Japanese notes, with the reading a Japanese front end assembles, as comment lines.
- Built-in SI-unit readings for Japanese narration, with an optional letter map for spelling out symbols.
- `examples/readings_ja_molbio.csv`, a working reading dictionary from the author's molecular-biology lectures.

### Synthesis and screening
- Narration with a voice-cloning TTS engine (Qwen3-TTS in-process, GPT-SoVITS through its HTTP API); the voice is defined by a few seconds of reference speech and its transcript (`--ref-wav`, `--ref-text`). Generated files carry the engine and model in their names.
- Qwen3-TTS narration is synthesized sentence by sentence, and generation is capped at the number of codec tokens the text can plausibly need; the synthesis time is reported per slide.
- `synthesize` records the text and dictionaries each audio was made from (`audio_sources.json`); `pack` and `verify` warn when a text was edited after its audio was made.
- `verify` transcribes the narration with faster-whisper and compares it with the text (kana-based for Japanese); a slide is flagged by similarity, by the longest single stretch of disagreement, or optionally by character error rate. `verify_differences_*.csv` lists every difference, longest first.
- `examples/screening_check.py` injects narration errors of a known size and reports how often the check notices them.

### Packing
- `pack` embeds the audio in the structure PowerPoint writes for recorded narration, inserting a narration object where a slide has none, giving every slide its own media file, setting the slide advance time (audio length plus `--slide-pause`, default one second), parking the audio icon outside the visible area (`--keep-audio-icon`), and removing the settings of a previous recording (trim, fade, bookmarks, laser-pointer path, play/pause/seek events; `--remove-recorded`). `--data-type text|audio|all` chooses what is written.
- The texts are written into the notes, per language. A note with texts of several languages has one heading of the same form per language (`=== pptx-narrator: [en] translated from [ja] <time> #<hash>; edited <time> ===`); a note of one language has none. The text of the language whose audio the slide plays is on top, where the presenter reads; below it, newer parts are above older ones.
- For each language, the text is compared with that part of the note and with what `extract` read or `pack` wrote (`note_baseline.json`): the same text is not written again, so a part that is not rewritten keeps its formatting and struck-through text; a text edited in the workspace is written with `--update` or `--overwrite`; a note edited in the deck is written over only with `--overwrite`; a language the note lacks is added. `extract` reads each part back into the text of its language.
- A slide plays one audio: audio of several languages in the workspace needs `--lang`, and missing audio is reported while the texts are still written.

### Tests and examples
- `tests/smoke_test.py` covers the pipeline with the back-ends mocked.
- `examples/make_sample_deck.py` builds a small deck with notes for trying the pipeline.
