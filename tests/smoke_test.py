"""Smoke tests for PPTX-Narrator (no TTS/ASR models or network needed; engines are mocked).

Run from the repository root:  python tests/smoke_test.py
Requires: python-pptx, pydub (+ FFmpeg), numpy, soundfile.
"""
import os, sys, types, tempfile, csv, shutil, zipfile, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import pptx_narrator as pn

ok = lambda c, m: print(("PASS " if c else "FAIL ") + m) or (c or sys.exit(1))

# --- language helpers / legacy filenames
ok(pn.normalize_lang("ZH_cn") == "zh-CN" and pn.normalize_lang("eng") == "en" and pn.normalize_lang("Auto") == "auto", "normalize_lang")
ok(pn.text_filename(3, "ja") == "slide_3.txt" and pn.text_filename(3, "en") == "slide_3_eng.txt" and pn.text_filename(3, "de") == "slide_3_de.txt", "text filenames (legacy ja/en)")
ok(pn.audio_filename(1, "ja", "v4") == "slide_1.v4.m4a" and pn.spoken_filename(1, "en", "v4") == "slide_1_eng.v4.spoken.txt", "audio/spoken legacy names")
ok(pn.qwen3_language("de") == "German" and pn.qwen3_language("zh-TW") == "Chinese" and pn.qwen3_language("nl") is None, "qwen3 languages")
ok(pn.gpt_sovits_language("yue") == "yue" and pn.gpt_sovits_language("de") is None, "gpt-sovits languages")
ok([pn.detect_language(t)[0] for t in ["今日はDNAの話です", "안녕하세요 여러분", "На этом слайде показаны результаты.", "Heute sprechen wir über die mRNA und ihre Rolle.", "这张幻灯片显示了实验结果。", "這張投影片顯示了實驗結果。"]] == ["ja", "ko", "ru", "de", "zh-CN", "zh-TW"], "language identification")
deck_langs = pn.detect_note_languages({1: "Willkommen zur heutigen Vorlesung über Genomeditierung.", 2: "Fragen?", 3: "Vielen Dank.", 4: "Hier vergleichen wir die Effizienz von drei Methoden in Zellen."})
ok(set(deck_langs.values()) == {"de"}, f"short notes follow deck language {deck_langs}")
ok(pn.detect_note_languages({1: "本日はゲノム編集について説明します。", 2: "結論"}) == {1: "ja", 2: "ja"}, "kanji-only note in Japanese deck")
ok(pn.split_into_chunks("Hello world. It costs 3.5 mg! Next?") == ["Hello world.", "It costs 3.5 mg!", "Next?"], "split latin")
ok(pn.split_into_chunks("今日は晴れ。明日は雨！") == ["今日は晴れ。", "明日は雨！"], "split cjk")

d = tempfile.mkdtemp()
# --- dictionary: legacy header, headerless, Reading_<lang>
legacy = os.path.join(d, "legacy.csv")
open(legacy, "w", encoding="utf-8").write("Term,Japanese_Reading,English_Reading,Type\nGbp,ギガベースペア,gigabase pairs,unit\nDNA,ディーエヌエー,D N A,\nTerm,,,\n")
ok(pn.apply_dictionary("DNA 3Gbp 5mg Term", legacy, "ja") == "ディーエヌエー 3ギガベースペア 5ミリグラム Term", "legacy dict ja + builtin units")
ok(pn.apply_dictionary("DNA is 3 Gbp and 5 mg", legacy, "en") == "D N A is 3 gigabase pairs and 5 mg", "legacy dict en (unit entry, no builtin, spacing kept)")
headerless = os.path.join(d, "nohdr.csv")
open(headerless, "w", encoding="utf-8").write("DNA,ディーエヌエー,D N A\n")
ok(pn.apply_dictionary("DNA", headerless, "ja") == "ディーエヌエー", "headerless v1.0 dict")
multi = os.path.join(d, "multi.csv")
open(multi, "w", encoding="utf-8").write("Term,Japanese_Reading,English_Reading,Type,Reading_de,Reading_zh-CN\nmRNA,メッセンジャーアールエヌエー,messenger R N A,,Boten-RNA,信使RNA\nkDa,キロダルトン,kilodaltons,unit,Kilodalton,千道尔顿\n")
ok(pn.apply_dictionary("Die mRNA hat 2 kDa.", multi, "de") == "Die Boten-RNA hat 2 Kilodalton.", "Reading_de")
ok(pn.apply_dictionary("mRNA 2kDa", multi, "zh-CN") == "信使RNA 2千道尔顿", "Reading_zh-CN with CJK unit join")
ok(pn.apply_dictionary("mRNA", multi, "fr") == "mRNA", "missing column -> unchanged")
ok(pn.apply_dictionary(r"X1", _p := os.path.join(d, "bs.csv"), "en") == "X1" , "no dict file")
open(_p, "w", encoding="utf-8").write("Term,English_Reading\nX1,a\\1b\n")
ok(pn.apply_dictionary("X1", _p, "en") == "a\\1b", "backslash in reading is literal")

# --- extraction with explicit/auto language
from pptx import Presentation
prs = Presentation(); lay = prs.slide_layouts[5]
for t in ["今日はDNAの話です。", "Today we talk about DNA and its structure in the cell.", "안녕하세요 여러분 DNA"]:
    s = prs.slides.add_slide(lay); s.notes_slide.notes_text_frame.text = t
deck = os.path.join(d, "deck.pptx"); prs.save(deck)
ws = os.path.join(d, "ws"); os.makedirs(ws)
pn.step_extract_notes(deck, ws, [1, 2, 3], "auto")
ok(sorted(os.listdir(ws)) == ["slide_1.txt", "slide_2_eng.txt", "slide_3_ko.txt"], f"auto extraction names {sorted(os.listdir(ws))}")
ws2 = os.path.join(d, "ws2"); os.makedirs(ws2)
pn.step_extract_notes(deck, ws2, [2], "de")
ok(os.listdir(ws2) == ["slide_2_de.txt"], "explicit --source-lang de")

# --- find_source_text
ok(pn.find_source_text(ws, 1, "auto", exclude_lang="de")[0] == "ja", "find source auto ja")
ok(pn.find_source_text(ws, 3, "auto", exclude_lang="de")[0] == "ko", "find source auto ko")
ok(pn.find_source_text(ws, 2, "auto", exclude_lang="en") == (None, None), "exclude target")
ok(pn.find_source_text(ws2, 2, "de", exclude_lang="fr")[0] == "de", "explicit source")

# --- translation (mocked GoogleTranslator)
calls = []
class FakeTr:
    def __init__(self, source, target):
        if target == "xx": raise ValueError("unsupported")
        self.s, self.t = source, target; calls.append((source, target))
    def translate(self, text): return f"[{self.s}->{self.t}] {text}"
pn.GoogleTranslator = FakeTr
pn.step_translate_notes(ws, [1, 2, 3], "auto", "de")
ok(open(os.path.join(ws, "slide_1_de.txt"), encoding="utf-8").read().startswith("[ja->de]"), "translate ja->de")
ok(open(os.path.join(ws, "slide_2_de.txt"), encoding="utf-8").read().startswith("[en->de]"), "translate en->de")
ok(open(os.path.join(ws, "slide_3_de.txt"), encoding="utf-8").read().startswith("[ko->de]"), "translate ko->de")
pn.step_translate_notes(ws2, [2], "de", "ja")
ok(open(os.path.join(ws2, "slide_2.txt"), encoding="utf-8").read().startswith("[de->ja]"), "translate de->ja (legacy ja name)")
pn.step_translate_notes(ws2, [2], "de", "xx")
ok(not os.path.exists(os.path.join(ws2, "slide_2_xx.txt")), "unsupported target handled")

# --- scan adds Reading_<lang> column, keeps legacy rows
pn._ensure_nltk_data = lambda: None
fake_corpus = types.SimpleNamespace(words=lambda *a: ["the", "is", "about", "we"])
sys.modules["nltk.corpus"] = types.SimpleNamespace(stopwords=fake_corpus, words=types.SimpleNamespace(words=lambda: ["heute", "sprechen", "notes"]))
scan_dict = os.path.join(d, "scan.csv"); shutil.copy(legacy, scan_dict)
pn.step_scan_and_update_dict(ws, scan_dict, [1, 2], reading_lang="de")
rows = list(csv.reader(open(scan_dict, encoding="utf-8")))
ok(rows[0] == ["Term", "Japanese_Reading", "English_Reading", "Type", "Reading_de"], f"scan header {rows[0]}")
ok(any(r[0] == "Gbp" and r[1] == "ギガベースペア" for r in rows), "legacy rows preserved")
new = {r[0]: r for r in rows[4:]}
ok("ber" not in new and "DNA" not in new, f"scan candidates sane {sorted(new)}")

# --- TTS language validation (engines not called)
logs = []
class H(pn.logging.Handler):
    def emit(self, rec): logs.append(rec.getMessage())
pn.logger.addHandler(H())
reftxt = os.path.join(d, "ref.txt"); open(reftxt, "w").write("x")
pn.step_generate_audio_qwen3(ws, [1], "nl", "r.wav", reftxt, None, "qwen3-1.7B", "1.7B", "cpu")
ok(any("does not support 'nl'" in m for m in logs), "qwen3 unsupported lang rejected before model load")
pn.step_generate_audio(ws, [1], "de", "r.wav", reftxt, "ja", "http://127.0.0.1:1/", None, "v4")
ok(any("GPT-SoVITS supports only" in m for m in logs), "gpt-sovits unsupported lang rejected")

# GPT-SoVITS payload for Korean
sent = {}
class Resp: status_code = 500; text = "mock"
pn.requests = types.SimpleNamespace(post=lambda url, json, timeout: (sent.update(json), Resp())[1], RequestException=Exception)
pn.step_generate_audio(ws, [3], "ko", "r.wav", reftxt, "ja", "http://x/", multi, "v4")
ok(sent.get("text_lang") == "ko" and sent.get("prompt_lang") == "ja", f"gpt-sovits payload langs {sent.get('text_lang')},{sent.get('prompt_lang')}")
ok(os.path.exists(os.path.join(ws, "slide_3_ko.v4.spoken.txt")), "spoken file for ko")

# Qwen3 with fake model
import numpy as np
fake_qwen = {}
class FakeModel:
    def create_voice_clone_prompt(self, ref_audio, ref_text): return "P"
    def generate_voice_clone(self, text, language, **kw):
        fake_qwen.setdefault("langs", set()).add(language); return [np.zeros(2400, dtype="float32")], 24000
pn._load_qwen3_model = lambda s, dvc: FakeModel()
pn.step_generate_audio_qwen3(ws, [1, 2], "de", "r.wav", reftxt, multi, "qwen3-1.7B", "1.7B", "auto")
ok(fake_qwen.get("langs") == {"German"} and os.path.exists(os.path.join(ws, "slide_1_de.qwen3-1.7B.m4a")), "qwen3 German synthesis + filename")

# --- verify for German with mocked faster-whisper
fw = types.ModuleType("faster_whisper")
class Seg:
    def __init__(s, t): s.text = t
class WM:
    def __init__(s, *a, **k): pass
    def transcribe(s, p, language=None):
        fake_qwen["asr_lang"] = language
        return [Seg(" [JA->DE] heute ist Wetter")], None
fw.WhisperModel = WM; sys.modules["faster_whisper"] = fw
pn.step_verify_audio(ws, [1, 2], "de", "qwen3-1.7B", threshold=0.5, cer_threshold=0.2)
rep = os.path.join(ws, "verify_report_de.qwen3-1.7B.csv")
r = list(csv.reader(open(rep, encoding="utf-8")))
ok(r[0][:4] == ["slide", "similarity", "cer", "status"] and fake_qwen["asr_lang"] == "de", "verify report de + whisper language")
ok(pn.text_scores("Héllo, World!", "hello world")[1] > 0 and pn.text_scores("Hello, World!", "hello world")[1] == 0, "text normalization")

# --- pack with writeback (target de + source ja)
# scan on a German file keeps only acronym-like terms
ws3 = os.path.join(d, "ws3"); os.makedirs(ws3)
open(os.path.join(ws3, "slide_1_de.txt"), "w", encoding="utf-8").write("Heute zeigen wir, wie CRISPR und mRNA in Zellen wirken, 5 kDa groß.")
dd = os.path.join(d, "de.csv")
pn.step_scan_and_update_dict(ws3, dd, [1], reading_lang="de")
terms = sorted(r[0] for r in list(csv.reader(open(dd, encoding="utf-8")))[1:])
ok(terms == ["5 kDa", "CRISPR", "kDa", "mRNA"], f"German scan terms {terms}")

# pack: writeback target de + source ja, audio replaced for de
from pydub import AudioSegment
prs = Presentation(deck)
prs.save(deck)
z_in = zipfile.ZipFile(deck); packed = os.path.join(d, "deck_audio.pptx"); z_out = zipfile.ZipFile(packed, "w", zipfile.ZIP_DEFLATED)
for it in z_in.infolist():
    data = z_in.read(it.filename)
    if it.filename == "[Content_Types].xml":
        data = data.replace(b"<Default ", b'<Default Extension="m4a" ContentType="audio/mp4"/><Default ', 1)
    if it.filename == "ppt/slides/_rels/slide1.xml.rels":
        data = data.replace(b"</Relationships>", b'<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/media1.m4a"/></Relationships>')
    if it.filename == "ppt/slides/slide1.xml":
        data = re.sub(rb"</p:cSld>", b'</p:cSld><p:transition advTm="500"/>', data, count=1)
    z_out.writestr(it, data)
AudioSegment.silent(duration=300).export(os.path.join(d, "m.m4a"), format="ipod")
z_out.writestr("ppt/media/media1.m4a", open(os.path.join(d, "m.m4a"), "rb").read()); z_out.close()
AudioSegment.silent(duration=1800).export(os.path.join(ws, "slide_1_de.v4.m4a"), format="ipod")
out = os.path.join(d, "out.pptx")
pn.step_pack_pptx(packed, out, ws, [1, 2], "de", "v4", source_lang="auto", writeback_notes=True)
p2 = Presentation(out)
n1 = p2.slides[0].notes_slide.notes_text_frame.text
ok(n1.startswith("[ja->de]") and "今日はDNA" in n1, f"writeback de + ja notes: {n1[:40]!r}")
adv = re.findall(rb'advTm="(\d+)"', zipfile.ZipFile(out).read("ppt/slides/slide1.xml"))
ok(adv and 1700 < int(adv[0]) < 2000, f"advTm set {adv}")

# CLI parsing and validation
parser = pn.build_parser()
a = parser.parse_args(["--pptx", deck, "--source_lang", "JA", "--target-lang", "zh_cn", "--translate"])
ok(a.source_lang == "ja" and a.target_lang == "zh-CN", "CLI lang normalization + legacy alias")
def expect_exit(argv, text):
    import io, contextlib
    buf = io.StringIO()
    try:
        with contextlib.redirect_stderr(buf):
            pn.main(argv)
    except SystemExit:
        pass
    ok(text in buf.getvalue(), f"CLI error '{text}'")
expect_exit(["--pptx", deck, "--tts", "--engine", "qwen3", "--target-lang", "nl", "--ref-wav", "a", "--ref-text-file", "b"], "Qwen3-TTS does not support 'nl'")
expect_exit(["--pptx", deck, "--tts", "--target-lang", "de", "--ref-wav", "a", "--ref-text-file", "b"], "GPT-SoVITS does not support --target-lang 'de'")
expect_exit(["--pptx", deck, "--translate", "--source-lang", "de", "--target-lang", "de"], "--translate needs")
print("ALL TESTS PASSED")
