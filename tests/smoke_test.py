"""Smoke tests for PPTX-Narrator (no TTS/ASR models or network needed; back-ends are mocked).

Run from the repository root:  python tests/smoke_test.py
Requires: python-pptx, pydub (+ FFmpeg), numpy, soundfile, py3langid.
"""
import os, sys, types, tempfile, csv, zipfile, re, io, contextlib
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import pptx_narrator as pn
from pptx import Presentation
from pydub import AudioSegment
import numpy as np


def ok(cond, msg):
    print(("PASS " if cond else "FAIL ") + msg)
    if not cond:
        sys.exit(1)


def write(path, text):
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


d = tempfile.mkdtemp()

# ---------------------------------------------------------------- languages
ok(pn.normalize_lang("ZH_cn") == "zh-CN" and pn.normalize_lang("zh") == "zh-CN" and pn.normalize_lang("eng") == "en"
   and pn.normalize_lang("he") == "iw" and pn.normalize_lang("Auto") == "auto", "normalize_lang")
ok(pn.text_filename(3, "ja") == "slide_3.txt" and pn.text_filename(3, "en") == "slide_3_eng.txt" and pn.text_filename(3, "de") == "slide_3_de.txt",
   "text filenames (v1.x names kept for ja/en)")
ok(pn.audio_filename(1, "ja", "v4") == "slide_1.v4.m4a" and pn.spoken_filename(1, "en", "v4") == "slide_1_eng.v4.spoken.txt",
   "audio/spoken filenames")
ok(pn.qwen3_language("de") == "German" and pn.qwen3_language("zh-TW") == "Chinese" and pn.qwen3_language("nl") is None, "Qwen3 languages")
ok(pn.gpt_sovits_language("yue") == "yue" and pn.gpt_sovits_language("de") is None, "GPT-SoVITS languages")
ok([pn.detect_language(t)[0] for t in ["今日はDNAの話です", "안녕하세요 여러분", "На этом слайде показаны результаты.",
                                       "Heute sprechen wir über die mRNA und ihre Rolle.", "这张幻灯片显示了实验结果。", "這張投影片顯示了實驗結果。"]]
   == ["ja", "ko", "ru", "de", "zh-CN", "zh-TW"], "language identification")
langs = pn.detect_note_languages({1: "Willkommen zur heutigen Vorlesung über Genomeditierung.", 2: "Fragen?", 3: "Vielen Dank.",
                                  4: "Hier vergleichen wir die Effizienz von drei Methoden in Zellen."})
ok(set(langs.values()) == {"de"}, f"short notes follow the deck language {langs}")
ok(pn.detect_note_languages({1: "本日はゲノム編集について説明します。", 2: "結論"}) == {1: "ja", 2: "ja"}, "kanji-only note in a Japanese deck")
ok(pn.split_into_chunks("Hello world. It costs 3.5 mg! Next?") == ["Hello world.", "It costs 3.5 mg!", "Next?"], "sentence split (Latin)")
ok(pn.split_into_chunks("今日は晴れ。明日は雨！") == ["今日は晴れ。", "明日は雨！"], "sentence split (CJK)")

# ---------------------------------------------------------------- dictionaries
readings_ja = os.path.join(d, "readings_ja.csv")
write(readings_ja, "string,replacement,type\nGbp,ギガベースペア,unit\nDNA,ディーエヌエー,\nCRISPR-Cas9,クリスパーキャスナイン,\n")
readings_de = os.path.join(d, "readings_de.csv")
write(readings_de, "mRNA,Boten-RNA,\nkDa,Kilodalton,unit\nX1,a\\1b,\n")          # no header
terms_ja_de = os.path.join(d, "terms_ja_de.csv")
write(terms_ja_de, "string,replacement\n塩基対,Basenpaare\n3 Gbp,drei Gigabasenpaare\n")
legacy = os.path.join(d, "legacy_dict.csv")
write(legacy, "Term,Japanese_Reading,English_Reading,Type\nGbp,ギガベースペア,gigabase pairs,unit\nRNA,アールエヌエー,R N A,\n")

R_ja = pn.load_dictionaries([readings_ja, legacy], "ja")
ok(pn.apply_dictionary("DNA 3Gbp 5mg RNA", R_ja, "ja") == "ディーエヌエー 3ギガベースペア 5ミリグラム アールエヌエー",
   "replacements + built-in units + v1.x file (Japanese column)")
ok(pn.apply_dictionary("RNA is 3 Gbp and 5 mg", pn.load_dictionaries([legacy], "en"), "en") == "R N A is 3 gigabase pairs and 5 mg",
   "v1.x file, English column")
ok(pn.load_dictionaries([legacy], "de") == [], "v1.x file has no column for other languages")
R_de = pn.load_dictionaries([readings_de], "de")
ok(pn.apply_dictionary("Die mRNA hat 2 kDa. X1", R_de, "de") == "Die Boten-RNA hat 2 Kilodalton. a\\1b", "headerless file, literal backslash")
ok(len(pn.load_dictionaries([readings_ja, os.path.join(d, "missing.csv")], "ja")) == 3, "missing dictionary file tolerated")
ok(pn.apply_dictionary("mRNA 5 mg", [], "fr") == "mRNA 5 mg", "no dictionary -> unchanged")
T = pn.load_dictionaries([terms_ja_de])
ok(pn.apply_dictionary("塩基対は3 Gbpで5 mg", T, "ja", builtin_units=False) == "Basenpaareはdrei Gigabasenpaareで5 mg",
   "pre-translation replacement without built-in unit readings")
both = pn.load_dictionaries([terms_ja_de, readings_de], "de")
pre = pn.apply_dictionary("塩基対とmRNA", both, "ja", builtin_units=False)
ok(pre == "BasenpaareとBoten-RNA" and pn.apply_dictionary("Basenpaare und Boten-RNA, 2 kDa", both, "de") == "Basenpaare und Boten-RNA, 2 Kilodalton",
   "one dictionary applied before translation and again before synthesis")

# ---------------------------------------------------------------- extraction
prs = Presentation(); layout = prs.slide_layouts[5]
for t in ["今日はDNAの話です。", "Today we talk about DNA and its structure in the cell.", "안녕하세요 여러분 DNA"]:
    s = prs.slides.add_slide(layout); s.notes_slide.notes_text_frame.text = t
deck = os.path.join(d, "deck.pptx"); prs.save(deck)
ws = os.path.join(d, "ws"); os.makedirs(ws)
pn.step_extract_notes(deck, ws, [1, 2, 3], "auto")
ok(sorted(os.listdir(ws)) == ["slide_1.txt", "slide_2_eng.txt", "slide_3_ko.txt"], f"auto extraction {sorted(os.listdir(ws))}")
ws2 = os.path.join(d, "ws2"); os.makedirs(ws2)
pn.step_extract_notes(deck, ws2, [2], "de")
ok(os.listdir(ws2) == ["slide_2_de.txt"], "explicit --source-lang")
ok(pn.find_source_text(ws, 1, "auto", exclude_lang="de")[0] == "ja" and pn.find_source_text(ws, 3, "auto", exclude_lang="de")[0] == "ko"
   and pn.find_source_text(ws, 2, "auto", exclude_lang="en") == (None, None), "find_source_text")

# ---------------------------------------------------------------- translation
class FakeTr:
    calls = []
    def __init__(self, source, target):
        if target == "xx":
            raise ValueError("unsupported language")
        self.s, self.t = source, target
    def translate(self, text):
        FakeTr.calls.append(text)
        return f"[{self.s}->{self.t}] {text}"

pn.GoogleTranslator = FakeTr
pn.step_translate_notes(ws, [1, 2, 3], "auto", "de")
ok(read(os.path.join(ws, "slide_1_de.txt")).startswith("[ja->de]") and read(os.path.join(ws, "slide_2_de.txt")).startswith("[en->de]")
   and read(os.path.join(ws, "slide_3_de.txt")).startswith("[ko->de]"), "translate ja/en/ko -> de")
write(os.path.join(ws, "slide_1.txt"), "塩基対の話です。")
pn.step_translate_notes(ws, [1], "auto", "de", dictionary=T)
ok("Basenpaare" not in read(os.path.join(ws, "slide_1_de.txt")), "existing translation kept without --retranslate")
pn.step_translate_notes(ws, [1], "auto", "de", dictionary=T, overwrite=True)
ok(FakeTr.calls[-1] == "Basenpaareの話です。" and read(os.path.join(ws, "slide_1.txt")) == "塩基対の話です。",
   "--retranslate: dictionary applied to the note sent to the translator; note file unchanged")
pn.step_translate_notes(ws2, [2], "de", "ja")
ok(read(os.path.join(ws2, "slide_2.txt")).startswith("[de->ja]"), "translate de -> ja")
pn.step_translate_notes(ws2, [2], "de", "xx")
ok(not os.path.exists(os.path.join(ws2, "slide_2_xx.txt")), "unsupported target language handled")

# ---------------------------------------------------------------- scan
pn._ensure_nltk_data = lambda: None
sys.modules["nltk.corpus"] = types.SimpleNamespace(
    stopwords=types.SimpleNamespace(words=lambda *a: ["the", "is", "about", "we", "and", "its", "in"]),
    words=types.SimpleNamespace(words=lambda: ["today", "talk", "structure", "cell", "slide"]))
ws3 = os.path.join(d, "ws3"); os.makedirs(ws3)
write(os.path.join(ws3, "slide_1.txt"), "今日はDNAとPCRとGFPの話です。")
write(os.path.join(ws3, "slide_2.txt"), "CRISPRについて。")
write(os.path.join(ws3, "slide_2_de.txt"), "Heute zeigen wir, wie CRISPR und mRNA in Zellen wirken, 5 kDa groß.")
terms_file = os.path.join(d, "new_terms.csv")
pn.step_scan_and_update_dict(ws3, terms_file, [], [1, 2], "de", for_translation=True)
rows = list(csv.reader(open(terms_file, encoding="utf-8")))
ok(rows[0] == ["string", "replacement", "type"] and sorted(r[0] for r in rows[1:]) == ["CRISPR", "DNA", "GFP", "PCR"]
   and all(r[1] == "" for r in rows[1:]), f"scan before translation: notes scanned, blank replacements {rows[1:]}")
pn.step_scan_and_update_dict(ws3, readings_de, pn.load_dictionaries([readings_de], "de"), [1, 2], "de")
rows = list(csv.reader(open(readings_de, encoding="utf-8")))
added = {r[0]: r[1] for r in rows[3:]}
ok(added.get("CRISPR") == "C R I S P R" and "kDa" not in added and "DNA" not in added,
   f"scan for synthesis: only the German narration text is scanned, with readings {added}")
pn.step_scan_and_update_dict(ws3, readings_ja, pn.load_dictionaries([readings_ja], "ja"), [1], "ja")
ok(["PCR", "P C R", ""] in list(csv.reader(open(readings_ja, encoding="utf-8"))), "scan of Japanese narration text")

# ---------------------------------------------------------------- synthesis (engines mocked)
logs = []
class Capture(pn.logging.Handler):
    def emit(self, record):
        logs.append(record.getMessage())
pn.logger.addHandler(Capture())
reftxt = os.path.join(d, "ref.txt"); write(reftxt, "x")
pn.step_generate_audio_qwen3(ws, [1], "nl", "r.wav", reftxt, R_de, "qwen3-1.7B", "1.7B", "cpu")
ok(any("does not support 'nl'" in m for m in logs), "Qwen3: unsupported language rejected before loading the model")
pn.step_generate_audio(ws, [1], "de", "r.wav", reftxt, "ja", "http://127.0.0.1:1/", R_de, "v4")
ok(any("GPT-SoVITS supports only" in m for m in logs), "GPT-SoVITS: unsupported language rejected")

sent = {}
class Resp:
    status_code = 500
    text = "mock"
pn.requests = types.SimpleNamespace(post=lambda url, json, timeout: (sent.update(json), Resp())[1], RequestException=Exception)
pn.step_generate_audio(ws, [3], "ko", "r.wav", reftxt, "ja", "http://x/", R_de, "v4")
ok(sent.get("text_lang") == "ko" and sent.get("prompt_lang") == "ja" and os.path.exists(os.path.join(ws, "slide_3_ko.v4.spoken.txt")),
   "GPT-SoVITS request languages")

seen = {}
class FakeModel:
    def create_voice_clone_prompt(self, ref_audio, ref_text):
        return "P"
    def generate_voice_clone(self, text, language, **kw):
        seen.setdefault("languages", set()).add(language)
        return [np.zeros(2400, dtype="float32")], 24000
pn._load_qwen3_model = lambda size, device: FakeModel()
pn.step_generate_audio_qwen3(ws, [1, 2], "de", "r.wav", reftxt, R_de, "qwen3-1.7B", "1.7B", "auto")
ok(seen.get("languages") == {"German"} and os.path.exists(os.path.join(ws, "slide_1_de.qwen3-1.7B.m4a")), "Qwen3 German synthesis")

# ---------------------------------------------------------------- verification (ASR mocked)
fw = types.ModuleType("faster_whisper")
class Segment:
    def __init__(self, text):
        self.text = text
class WhisperModel:
    def __init__(self, *a, **k):
        pass
    def transcribe(self, path, language=None):
        seen["asr_language"] = language
        return [Segment(" ganz anderer Text")], None
fw.WhisperModel = WhisperModel
sys.modules["faster_whisper"] = fw
pn.step_verify_audio(ws, [1, 2], "de", "qwen3-1.7B", threshold=0.5, cer_threshold=0.2)
report = list(csv.reader(open(os.path.join(ws, "verify_report_de.qwen3-1.7B.csv"), encoding="utf-8")))
ok(report[0][:4] == ["slide", "similarity", "cer", "status"] and seen["asr_language"] == "de" and report[1][3] == "FLAGGED",
   "verification report for German")
ok(pn.text_scores("Hello, World!", "hello world")[1] == 0 and pn.text_scores("Héllo", "hello")[1] > 0, "text normalization for comparison")
pj = types.ModuleType("pyopenjtalk")
pj.g2p = lambda t, kana=True: {"今日は晴れです。": "キョーワ、ハレデス。", "今日は晴れでした": "キョーワハレデシタ"}.get(t, t)
sys.modules["pyopenjtalk"] = pj
ok([round(x, 3) if isinstance(x, float) else x for x in pn.kana_scores("今日は晴れです。", "今日は晴れでした")]
   == [0.824, 0.25, "キョーワハレデス", "キョーワハレデシタ"], "kana similarity and CER")

# ---------------------------------------------------------------- packing
z_in = zipfile.ZipFile(deck)
packed = os.path.join(d, "deck_audio.pptx")
z_out = zipfile.ZipFile(packed, "w", zipfile.ZIP_DEFLATED)
for item in z_in.infolist():
    data = z_in.read(item.filename)
    if item.filename == "[Content_Types].xml":
        data = data.replace(b"<Default ", b'<Default Extension="m4a" ContentType="audio/mp4"/><Default ', 1)
    if item.filename == "ppt/slides/_rels/slide1.xml.rels":
        data = data.replace(b"</Relationships>", b'<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/media1.m4a"/></Relationships>')
    if item.filename == "ppt/slides/slide1.xml":
        data = re.sub(rb"</p:cSld>", b'</p:cSld><p:transition advTm="500"/>', data, count=1)
    z_out.writestr(item, data)
AudioSegment.silent(duration=300).export(os.path.join(d, "m.m4a"), format="ipod")
z_out.writestr("ppt/media/media1.m4a", open(os.path.join(d, "m.m4a"), "rb").read())
z_out.close()
AudioSegment.silent(duration=1800).export(os.path.join(ws, "slide_1_de.v4.m4a"), format="ipod")
out_deck = os.path.join(d, "out.pptx")
pn.step_pack_pptx(packed, out_deck, ws, [1, 2], "de", "v4", source_lang="auto", writeback_notes=True)
notes = Presentation(out_deck).slides[0].notes_slide.notes_text_frame.text
info = pn.parse_structured_note(notes)
ok(notes.startswith("=== pptx-narrator: narration [de] from [ja] #") and info and info["narration_lang"] == "de"
   and info["narration_text"].startswith("[ja->de]") and info["source_lang"] == "ja" and info["source_text"] == "塩基対の話です。"
   and info["fingerprint"] == pn.text_fingerprint("塩基対の話です。"), "write-back: marked narration part followed by the original note")
notes2 = Presentation(out_deck).slides[1].notes_slide.notes_text_frame.text
ok(pn.parse_structured_note(notes2)["source_lang"] == "en", "write-back for an English-note slide")
adv = re.findall(rb'advTm="(\d+)"', zipfile.ZipFile(out_deck).read("ppt/slides/slide1.xml"))
ok(adv and 1700 < int(adv[0]) < 2000, f"slide timing set to the audio length {adv}")

xml_trim = ('<p14:media r:embed="rId2"><p14:trim st="1200" end="800"/><p14:fade in="500"/><p14:bmkLst><p14:bmk name="a" time="1"/></p14:bmkLst></p14:media>'
            '<p14:media r:embed="rId9"><p14:trim st="10"/></p14:media>')
cleared, n_cleared = pn.clear_media_playback_settings(xml_trim, {"rId2"})
ok(n_cleared == 1 and cleared.startswith('<p14:media r:embed="rId2"/>') and '<p14:trim st="10"/>' in cleared,
   "trim/fade/bookmarks of the replaced audio removed, other media untouched")

# ---------------------------------------------------------------- structured notes round trip
note = pn.compose_structured_note("en", "Today we talk about DNA.", "ja", "今日はDNAの話です。")
info = pn.parse_structured_note(note)
ok(info["narration_lang"] == "en" and info["narration_text"] == "Today we talk about DNA." and info["source_text"] == "今日はDNAの話です。"
   and not info["spoken"], "compose/parse structured note")
ok(pn.parse_structured_note("普通のノート") is None, "ordinary notes are not structured")
ok(pn.text_fingerprint("a\r\nb  \n") == pn.text_fingerprint("a\nb"), "fingerprint ignores line endings and trailing spaces")

# extract the packed deck again: source part is extracted, unchanged translation restored
ws4 = os.path.join(d, "ws4"); os.makedirs(ws4)
pn.step_extract_notes(out_deck, ws4, [1, 2, 3], "auto")
ok(read(os.path.join(ws4, "slide_1.txt")) == "塩基対の話です。" and read(os.path.join(ws4, "slide_1_de.txt")).startswith("[ja->de]"),
   "re-extraction: source part + restored translation")
ok(pn._load_manifest(ws4)["1"]["de"]["source_fingerprint"] == pn.text_fingerprint("塩基対の話です。"), "manifest written on restore")
FakeTr.calls.clear()
pn.step_translate_notes(ws4, [1], "auto", "de", dictionary=T)
ok(FakeTr.calls == [], "restored translation is not translated again")

# edit the source part inside PowerPoint -> narration is stale
prs_e = Presentation(out_deck)
tf = prs_e.slides[0].notes_slide.notes_text_frame
tf.text = tf.text.replace("塩基対の話です。", "塩基対とRNAの話です。")
edited = os.path.join(d, "edited.pptx"); prs_e.save(edited)
ws5 = os.path.join(d, "ws5"); os.makedirs(ws5)
write(os.path.join(ws5, "slide_1_de.txt"), "old narration in the workspace")
pn.step_extract_notes(edited, ws5, [1], "auto")
ok(read(os.path.join(ws5, "slide_1.txt")) == "塩基対とRNAの話です。" and not os.path.exists(os.path.join(ws5, "slide_1_de.txt"))
   and read(os.path.join(ws5, "slide_1_de.stale.txt")).startswith("[ja->de]"), "edited source -> stale narration set aside")
pn.step_translate_notes(ws5, [1], "auto", "de", dictionary=T)
ok("RNA" in read(os.path.join(ws5, "slide_1_de.txt")), "stale narration is translated again")

# spoken-form narration blocks are not restored as translations
spoken_note = pn.compose_structured_note("de", "Boten-RNA", "ja", "mRNAの話", spoken=True)
prs_s = Presentation(); s_ = prs_s.slides.add_slide(prs_s.slide_layouts[5]); s_.notes_slide.notes_text_frame.text = spoken_note
sp = os.path.join(d, "spoken.pptx"); prs_s.save(sp)
ws6 = os.path.join(d, "ws6"); os.makedirs(ws6)
pn.step_extract_notes(sp, ws6, [1], "auto")
ok(sorted(os.listdir(ws6)) == ["slide_1.txt"], "spoken narration block ignored on extraction")

# pack warns when the translation is older than the source note
logs.clear()
write(os.path.join(ws, "slide_1.txt"), "塩基対の話を変更しました。")
pn.step_pack_pptx(packed, os.path.join(d, "out2.pptx"), ws, [1], "de", "v4", source_lang="auto", writeback_notes=True)
note_old = pn.parse_structured_note(Presentation(os.path.join(d, "out2.pptx")).slides[0].notes_slide.notes_text_frame.text)
ok(any("older version of the source note" in m for m in logs) and note_old["fingerprint"] != pn.text_fingerprint(note_old["source_text"]),
   "stale translation flagged at pack time and marked by its original fingerprint")

# ---------------------------------------------------------------- command line
parser = pn.build_parser()
a = parser.parse_args(["--pptx", deck, "--source_lang", "JA", "--target-lang", "zh_cn", "--translate",
                       "--dict-file", "a.csv", "--dict_file", "b.csv"])
ok(a.source_lang == "ja" and a.target_lang == "zh-CN" and a.dict_file == ["a.csv", "b.csv"], "CLI normalization, aliases, repeatable --dict-file")

def expect_error(argv, text):
    buf = io.StringIO()
    try:
        with contextlib.redirect_stderr(buf):
            pn.main(argv)
    except SystemExit:
        pass
    ok(text in buf.getvalue(), f"CLI error: {text}")

expect_error(["--pptx", deck, "--tts", "--engine", "qwen3", "--target-lang", "nl", "--ref-wav", "a", "--ref-text-file", "b"], "Qwen3-TTS does not support 'nl'")
expect_error(["--pptx", deck, "--tts", "--target-lang", "de", "--ref-wav", "a", "--ref-text-file", "b"], "GPT-SoVITS does not support --target-lang 'de'")
expect_error(["--pptx", deck, "--translate", "--source-lang", "de", "--target-lang", "de"], "--translate needs")
expect_error(["--pptx", deck, "--scan"], "--scan needs --dict-file")
print("ALL TESTS PASSED")
