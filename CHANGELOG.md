# Changelog

## Unreleased

- `verify` (Japanese) tells the narration's own differences from those of the recognizer: a difference on a Latin-script word of the narration (compared as letter names, while the recognizer writes what it heard) or one where the recognizer wrote the words with kanji the narration does not have (聖正 for 精製) is listed in `verify_asr_differences_<lang>.<model>.csv`, with its kind, and not counted in the similarity, the CER, the longest difference or the flag. A misread number always counts. The `LATIN` status is gone, as Latin-script words no longer spoil the scores. On a real lecture deck, 80 of 120 differences fell on English terms and 35 on other kanji, leaving 4 to listen to.
- Full-width digits and Latin letters (４つ, ＤＮＡ) are read as their usual forms by the dictionaries and the reading assist: a dictionary entry `DNA` now matches ＤＮＡ, and ４つ is no longer passed as written (read よんつ). Full-width punctuation is left as it is.
- Reading assist: a word with a number in it (4つ, 2つ目, 四つ) is written as pronounced (ヨッツ, フタツメ).
- A dictionary term that is an ordinary word (three letters or more, lower case after the first letter) now matches with its first letter in either case: `Prophase I` also matches `prophase I`, `transcription` also `Transcription` at the start of a sentence, and the replacement follows the case of the text where it begins with the same letter. Acronyms, mixed-case and short terms (`DNA`, `mRNA`, `Mg`) still match only as written, and so do units. `scan` counts a term as already in the dictionary by the same rule (it compared without case before, so a term it left out could still go unreplaced).
- Every option is registered under one name; other spellings (`--dryrun`, the underscore spellings, and in `examples/screening_check.py` `--in-lang` and `--threshold`) are read as that name before parsing, so that a shortened option such as `--dry`, `--in` or `--model` is never ambiguous between two spellings of the same option. A test checks this for every option of the tool and of the example script.
- Japanese reading assist (Qwen3-TTS, on by default): OpenJTalk divides the text as written into words, before the dictionaries replace terms by readings (which would mislead the division: せんしょくたい中 is divided as せんしょく/たい/中), leaving the dictionary terms to the dictionaries; a number with the counters after it, and a word that is a number with a counter, are written in katakana as pronounced (三日目 → ミッカメ, 10分後 → ジュップンゴ, 4日間 → ヨッカカン; Qwen3-TTS read 三日目 as みつかめ), every suffix, whose reading depends on the word before it, together with that word (授業中 → ジュギョウチュウ, 学生数 → ガクセイスウ, 科学的 → カガクテキ; 数 in 染色体数 was misread), and a number before a unit symbol keeps its digits for the unit reading and becomes kanji numerals after it (3.5 mL → 三点五ミリリットル); words containing a character whose reading depends on the word (`--reading-assist-chars`, default `中内外毎間上下前後目方所分`) are written in katakana as read, and numbers as kanji numerals, which the model then reads correctly (一日中 was read いちにちちゅう, 毎 まい, 3割 せんわり). Latin-script words and their numbers are left as they are. `--no-reading-assist` turns it off. pyopenjtalk is now part of the `qwen3` extra.
- Information messages of other libraries (such as "code_predictor_config is None ..." from qwen_tts) are no longer shown; their warnings and errors are, as are the messages of PPTX-Narrator.
- `synthesize` writes the audio of a slide and the reading it was made from (`.spoken.txt`) together, under temporary names first and renamed only once the audio is complete, and records each slide in `audio_sources.json` as soon as it is made. A run stopped half-way, or a slide that fails, leaves the old audio with its old reading, never a new reading beside old audio or a half-written file; the slides made before the stop are not made again by the next `--update`.
- `scan` wrote dictionary lines with Windows line ends (CRLF; editors show `^M`) mixed with plain ones; dictionaries and the `verify` reports are now written with plain line ends. Reading accepts either.
- `synthesize` no longer makes existing audio again unless asked, like the other commands. `--update` makes again the slides whose reading changed since their audio was made — the text was edited, or a dictionary changed how it is read, both seen in the fingerprint of the reading now recorded in `audio_sources.json`; `--edited-texts-only` leaves out the slides changed only by a dictionary. `--overwrite` makes every selected slide again. Audio made by an earlier version, without a record, is left as it is with `--update`, unless its reading (`.spoken.txt`) is newer than the audio — earlier versions wrote the reading first, so a run stopped or failed half-way left such a pair — in which case it is made again, with that reason shown.
- `synthesize --dry-run` lists the slides that would be synthesized, with the reason, and synthesizes nothing. `extract`, `scan`, `translate` and `pack` take `--dry-run` too: the command decides slide by slide as usual, on a copy of the workspace (and of the deck or dictionary it would write), reports what it would create, rewrite or remove, and changes nothing; `translate` does not load the model for it.
- Tab completion: with shtab installed (`pip install -e ".[completion]"`), `pptx-narrator --print-completion SHELL` prints a completion script for zsh, bash, fish or PowerShell, covering the workspace, the commands, their options and the files each argument takes. `pptx-narrator --completion-setup [SHELL]` shows how to set it up for the current or a given shell; it changes no file.
- `pptx-narrator --help` ends with examples: narration in the language of the notes, an English version, updating after edits, and renumbering after slides were inserted, deleted or reordered. The help shown when no command is given leaves them out and says where to find them. `pptx-narrator WS COMMAND --help` ends with examples of that command.
- With Qwen3-TTS, the sentences of a paragraph are synthesized together, up to `--chunk-chars` characters (default 200; 0 for one sentence at a time): a sentence end that is also the end of a generation was sometimes cut short by the model, and the paragraph now sounds connected.
- With Qwen3-TTS, a pause of fixed length now follows each chunk, so that a sentence is never run into the next however little silence the model left (the silence before a sentence's first sound is cut; its end, which often fades out softly, is kept): `--sentence-pause` (default 0.5 s) and `--paragraph-pause` (default 0.5 s; a paragraph ends at a blank line in the text).

## 1.1.0 – 2026-09-29

- `map DECK` shows how the slides of a deck correspond to the workspace, matched by the slide IDs PowerPoint keeps when slides are inserted, deleted or reordered (a workspace made before the IDs were recorded is matched by the fingerprints of its notes and the similarity of its texts). `--apply` renumbers the workspace to follow the deck, setting aside the files of slides no longer in it, and the log and ASR reports that speak of the old numbers, in `map_archive/<date_time>/`.
- `extract`, `pack` and `map --apply` record the slide IDs of the deck (`slide_map.json`). `extract` and `pack` stop, changing nothing, when the slides of the deck no longer correspond to the workspace.

## 1.0.1 – 2026-09-28

- The synthesis log reports one ratio, the synthesis time divided by the duration of the audio ("synthesis took 1.93 x the audio duration"); the summary line used to give the inverse ("x real time").
- `verify` marks a Japanese slide whose text still contains Latin-script words as `LATIN` (was `ENGLISH`): the kana comparison cannot score such words, whatever their language.
- The instruction to the translation model speaks of the notes of a slide deck, not of a lecture.
- `examples/terms_ja_en.csv`: a translation dictionary for Japanese notes narrated in English.
- `examples/screening_check.py` takes the workspace first and `--lang`, like the commands (`screening_check.py WS --lang ja`); `--verify-threshold` names its threshold as in `verify`. The old option names still work.
- README: how to tune the ASR check on one's own decks, option by option; how to set up the TTS engines, including a GPT-SoVITS server on another computer.
- The source is in `src/` (`src/pptx_narrator.py`) and the license in `LICENSE.txt`, as SoftwareX asks of the repositories of its articles; the installed command is unchanged.

## 1.0.0 – 2026-09-27

First public release.

### Command line
- An error names the command it concerns and shows that command's usage; a failure during a run says which command failed and why, keeps the details in the log of the workspace, and is recorded in its history. Files named by options (`--ref-wav`, `--ref-text`, `--dict-file`, `--letter-map`) are checked before a run starts.
- The command line is `pptx-narrator WS COMMAND [INPUT] [OUTPUT] [OPTIONS]`: the workspace of one deck comes first, then one of the commands `extract`, `scan`, `translate`, `synthesize`, `verify`, `pack` and `history`. `INPUT` and `OUTPUT` are the files outside the workspace that a command reads or writes (`extract DECK`, `scan DICT`, `pack DECK OUT`); the files inside the workspace are chosen with `--lang`, `--slides` and the model options, not by path.
- Nothing runs implicitly and nothing is carried over from an earlier run: what a command needs is given on its command line or in the configuration file.
- Nothing that exists is overwritten unless `--update` (write what has changed) or `--overwrite` (write everything selected) is given; generated audio is the exception, since it is made from the text and not edited by hand. `scan` adds to an existing dictionary only with `--append`, and `pack` never changes `DECK` and replaces an existing `OUT` only on request.
- `--lang` gives the language of the texts a command works on and is short for giving `--in-lang` and `--out-lang` the same language; `translate` reads `--in-lang` and writes `--out-lang`. For `extract`, `--lang` selects the languages to extract; a note of another language is reported and left out, and a requested language the deck does not contain is an error.
- Every command accepts `--slides` (e.g. `4,7-9`).
- Every run ends with the commands that can come next, ready to copy, and a report of the files it wrote; what it logged is appended to `pptx_narrator.log` in the workspace, and `history` lists the runs of a workspace (`--dates` adds when), showing a failed run with its cause (`[FAILED: ...]`) and a run that wrote nothing as such.
- A command written before the workspace, or a misspelt command, is answered with the correct form.
- `pptx-narrator --help` lists each command once and says how to see the options of one; every option has help text and is listed under its hyphenated spelling, the underscore spelling (`--in_lang`) being accepted too. Every option may be abbreviated as far as it stays unambiguous. `--version` prints the version.

### Configuration and records
- Parameters that stay the same across runs are read from a TOML file (`--config`, or `pptx_narrator.toml` in the current directory), with a `[common]` section and one section per command. Values are resolved as built-in defaults, then the configuration file, then the command line. A configured value is taken back for one run by giving the option empty (`--dict-file ''`), a switch by its `--no-` form; keys may be written with hyphens or underscores.
- Every run writes `.pptx_narrator_resolved.toml` (the effective value of every parameter, the version of the tool, the time, and the path and SHA-256 of each input; itself a valid configuration file) and appends to `.pptx_narrator_history.jsonl` (the settings, the files read and the files written). These records are never read back to fill in a later command.
- Paths inside the workspace are recorded relative to it, so a workspace can be moved or copied as a whole; paths outside it are absolute.

### Notes and texts
- Presenter notes are extracted per slide as editable text files named after the slide and the language (`slide_3_ja.txt`, `slide_3_en.txt`); the language is identified automatically (kana/Hangul rules plus py3langid). Hidden slides, the date, slide-number, header and footer placeholders and fields that PowerPoint fills in, and zero-width characters are left out, and so is struck-through text, which the author has deleted; a date the author typed is kept.
- Texts are treated alike whether they were extracted, translated or written by hand.
- `extract` replaces a text already in the workspace only with `--update` (when the note in the deck changed and the text was not edited in the workspace) or `--overwrite`.
- Typographic apostrophes, primes, quotation marks and dashes in the notes match a dictionary entry written with the plain ASCII character.

### Translation
- Optional: a text can be translated by an instruction-tuned language model run locally (default Qwen/Qwen3-4B, chosen with `--translate-model`), each note whole; the entries of a translation dictionary that occur in it are given to the model as instructions. The result is a text like any other, to be reviewed before synthesis.
- An existing translation is kept; `--update` translates again where the source changed (keeping a translation edited by hand), `--overwrite` translates again. `translations.json` records which version of the source each translation was made from, when, and the translation as made.

### Dictionaries and readings
- Dictionaries are plain string-replacement lists (`string,replacement,type`) given to the translation model as instructions for the terms that occur in a note (the note itself is not changed) and applied to the text before synthesis; `--dict-file` is repeatable, later files taking precedence. Longer strings are replaced first, alphanumeric strings only on word boundaries, and entries with an empty replacement do nothing.
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
