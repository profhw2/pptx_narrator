# Changelog

## 1.2.0 – 2026-09-17

### Added
- Multilingual pipeline: `--source-lang` and free `--target-lang` codes (e.g. `de`, `zh-CN`).
- Automatic identification of the note language (`--source-lang auto`, default): kana/Hangul rules plus py3langid restricted to Google Translate languages; short or ambiguous notes take the deck's majority language. New dependency: `py3langid`.
- `--translate` works for any source/target pair supported by Google Translate (previously Japanese → English only); `--retranslate` overwrites existing translations.
- Dictionaries per language pair: CSV files whose header names the pair (`ja,ja,type`, `ja,de,type`, …). Same-language files rewrite the narration text before synthesis; different-language files are glossaries applied during translation through placeholders. Loaded from `dict_<source>_<target>.csv` in `--dict-dir` and from repeatable `--dict-file`; `--scan` appends candidates to the file of the pair (terms from notes in language L go to `L,<target-lang>`).
- Structured notes write-back: translated or spoken-form narration is written above the original note between `=== pptx-narrator: … ===` marker lines; `--extract` recognizes the layout, uses only the source part, restores the narration as the existing translation when the source is unchanged (hash in the marker), and sets it aside as `.stale.txt` otherwise. `translations.json` records which source version each translation came from.
- TTS in all languages of the selected engine (Qwen3-TTS: zh, en, ja, ko, de, fr, ru, pt, es, it; GPT-SoVITS: zh, en, ja, ko, yue), with validation before synthesis. `--ref-lang` accepts any GPT-SoVITS language.
- `--verify` for languages other than Japanese (character-level similarity and CER on normalized text).
- `tests/smoke_test.py` (back-ends mocked).

### Changed
- Step order is now extract → translate → scan → tts → verify → pack, so that one `--translate --scan` run can collect rewrite candidates from the translated text.
- `dict.csv` (v1.x layout) is still read as `ja,ja` and `en,en` but no longer written; new entries go to `dict_<source>_<target>.csv`. The repository's `dict.csv` became `dict_ja_ja.csv`.
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
