# Changelog

## 1.2.0 – 2026-09-17

### Added
- Multilingual pipeline: `--source-lang` and free `--target-lang` codes (e.g. `de`, `zh-CN`).
- Automatic identification of the note language (`--source-lang auto`, default): kana/Hangul rules plus py3langid restricted to Google Translate languages; short or ambiguous notes take the deck's majority language. New dependency: `py3langid`.
- `--translate` works for any source/target pair supported by Google Translate (previously Japanese → English only); `--retranslate` overwrites existing translations.
- Dictionaries are plain string-replacement lists (`string,replacement,type`) applied to the text processed in the run: to the notes before `--translate` and to the narration text before `--tts`. Translation and synthesis can be run separately to use different dictionaries; `--dict-file` is repeatable.
- Structured notes write-back: translated or spoken-form narration is written above the original note between `=== pptx-narrator: … ===` marker lines; `--extract` recognizes the layout, uses only the source part, restores the narration as the existing translation when the source is unchanged (hash in the marker), and sets it aside as `.stale.txt` otherwise. `translations.json` records which source version each translation came from.
- TTS in all languages of the selected engine (Qwen3-TTS: zh, en, ja, ko, de, fr, ru, pt, es, it; GPT-SoVITS: zh, en, ja, ko, yue), with validation before synthesis. `--ref-lang` accepts any GPT-SoVITS language.
- `--verify` for languages other than Japanese (character-level similarity and CER on normalized text).
- `tests/smoke_test.py` (back-ends mocked).

### Changed
- There is no implicit default dictionary: give `--dict-file` (required by `--scan`). A v1.x `dict.csv` can be passed with `--dict-file`; the Japanese or English column is used according to the narration language.
- Workspace files for languages other than Japanese/English are named `slide_N_<lang>.txt`; the Japanese (`slide_N.txt`) and English (`slide_N_eng.txt`) names are unchanged.
- Verification report: kana columns renamed to `intended_normalized` / `asr_normalized`; report file includes the language suffix for non-Japanese narration.
- `--writeback-notes` keeps the original note (instead of overwriting it with the narration or appending it without markers); `--use-spoken-notes` no longer replaces the original note with its spoken form.
- Dictionary `unit` entries are applied to any narration language; built-in SI-unit readings remain Japanese only.
- Sentence splitting for Qwen3-TTS also handles `.`, `!`, `?` followed by whitespace.
- Term scanning no longer splits words containing accented Latin letters, and ignores ordinary words in non-English notes.

### Fixed
- Dictionary replacements containing backslashes were interpreted as regex escapes.

## 1.1.0 – 2026-09-17
- Clean repository; `pptx-narrator` entry point; hyphenated options; kana CER in verification; dictionary, README and LICENSE fixes; Zenodo metadata.
