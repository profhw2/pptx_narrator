# Changelog

## 1.2.0 – 2026-09-17

### Added
- Multilingual pipeline: `--source-lang` (default `auto`) and free `--target-lang` codes (e.g. `de`, `zh-CN`).
- `--translate` works for any source/target pair supported by Google Translate (previously Japanese → English only).
- TTS in all languages of the selected engine (Qwen3-TTS: zh, en, ja, ko, de, fr, ru, pt, es, it; GPT-SoVITS: zh, en, ja, ko, yue), with validation of unsupported languages before synthesis. `--ref-lang` accepts any GPT-SoVITS language.
- Dictionary reading columns `Reading_<lang>`; `--scan` adds the column for `--target-lang` and writes provisional readings for it.
- `--verify` for languages other than Japanese (character-level similarity and CER on normalized text).
- `tests/smoke_test.py` (engines mocked).

### Changed
- Workspace files for languages other than Japanese/English are named `slide_N_<lang>.txt`; the Japanese (`slide_N.txt`) and English (`slide_N_eng.txt`) names are unchanged.
- Verification report: kana columns renamed to `intended_normalized` / `asr_normalized`; report file includes the language suffix for non-Japanese narration.
- `--scan` writes provisional readings for `--target-lang` (in v1.1, always Japanese). Pass `--target-lang ja` to keep the previous behaviour.
- Dictionary `unit` entries are also applied to non-Japanese narration; built-in SI-unit readings remain Japanese only.
- `--writeback-notes` writes the narration text followed by the source-language text.
- Sentence splitting for Qwen3-TTS also handles `.`, `!`, `?` followed by whitespace.
- Term scanning no longer splits words containing accented Latin letters, and ignores ordinary words in non-English notes.

### Fixed
- Dictionary readings containing backslashes were interpreted as regex escapes.

## 1.1.0 – 2026-09-17
- Clean repository; `pptx-narrator` entry point; hyphenated options; kana CER in verification; dictionary, README and LICENSE fixes; Zenodo metadata.
