# Changelog

## 1.0.0 – 2026-09-18

First public release. Versions 1.1.0 and 1.2.0 were development numbers used before this release and are not published.

### Features
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
- `tests/smoke_test.py` covers the pipeline with the back-ends mocked.
- `examples/make_sample_deck.py` builds a small deck with notes for trying the pipeline.
- `examples/screening_check.py`, which injects narration errors of a known size into the intended text and reports how often the ASR check notices them, by error size and note length.
- `examples/readings_ja_molbio.csv`, a working reading dictionary from the author's molecular-biology lectures.
- In a dictionary, a `;` or `#` at the start of a line or after a space starts a comment that runs to the end of the line, for annotating an entry or switching it off; a line that is only a comment is skipped. A `#` inside a term (`C#`, or `"#1"` in quotes) is kept.
- `--scan-compounds` writes the compounds of Japanese notes, with the reading a Japanese front end assembles for them, as comment lines in the dictionary; they do nothing until the reading is corrected and the `;` removed.
- `--scan` no longer proposes single letters or digits, which have no useful reading of their own.
- Typographic apostrophes, primes, quotation marks and dashes in the notes (e.g. the ’ PowerPoint inserts, or ′ pasted from a paper) match a dictionary entry written with the plain ASCII character, so an entry such as `5',ごだっしゅ` applies to all of its shapes.
- Dictionary entries containing a backslash are reported, because they are matched literally and usually come from shell escaping.
