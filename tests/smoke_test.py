"""Smoke tests for PPTX-Narrator (no TTS/ASR models or network needed; back-ends are mocked).

Run from the repository root:  python tests/smoke_test.py
Requires: python-pptx, pydub (+ FFmpeg), numpy, soundfile, py3langid.
"""
import os, sys, types, tempfile, csv, zipfile, re, io, contextlib, argparse
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
ok(pn.text_filename(3, "ja") == "slide_3_ja.txt" and pn.text_filename(3, "en") == "slide_3_en.txt" and pn.text_filename(3, "de") == "slide_3_de.txt",
   "text file names carry the language")
ok(pn.audio_filename(1, "ja", "v4") == "slide_1_ja.v4.m4a" and pn.spoken_filename(1, "en", "v4") == "slide_1_en.v4.spoken.txt",
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
ok(sorted(os.listdir(ws)) == ["slide_1_ja.txt", "slide_2_en.txt", "slide_3_ko.txt"], f"auto extraction {sorted(os.listdir(ws))}")
ws2 = os.path.join(d, "ws2"); os.makedirs(ws2)
pn.step_extract_notes(deck, ws2, [2], "de")
ok(os.listdir(ws2) == [], "--in-lang is a selector: a note of another language is left out")
ws2b = os.path.join(d, "ws2b"); os.makedirs(ws2b)
pn.step_extract_notes(deck, ws2b, [1, 2, 3], "en")
ok(os.listdir(ws2b) == ["slide_2_en.txt"], "--in-lang keeps only the notes of that language")
ok(pn.step_extract_notes(deck, ws2b, [1, 2, 3], "fr") == 0
   and pn.step_extract_notes(deck, ws2b, [1, 2, 3], "en") == 1,
   "extract reports what this run wrote, not what the workspace already held")
ok(pn._file_role("slide_3_ja.txt") == "note text"
   and pn._file_role("slide_3_ja.v4.spoken.txt") == "spoken text"
   and pn._file_role("slide_3_ja.v4.m4a") == "audio"
   and pn._file_role("verify_report_ja.v4.csv") == "verification report",
   "a directory record says what each file is")
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
write(os.path.join(ws, "slide_1_ja.txt"), "塩基対の話です。")
pn.step_translate_notes(ws, [1], "auto", "de", dictionary=T)
ok("Basenpaare" not in read(os.path.join(ws, "slide_1_de.txt")), "existing translation kept without --retranslate")
pn.step_translate_notes(ws, [1], "auto", "de", dictionary=T, overwrite=True)
ok(FakeTr.calls[-1] == "Basenpaareの話です。" and read(os.path.join(ws, "slide_1_ja.txt")) == "塩基対の話です。",
   "--retranslate: dictionary applied to the note sent to the translator; note file unchanged")
write(os.path.join(ws2, "slide_2_de.txt"), "Heute sprechen wir über Basenpaare.")
pn.step_translate_notes(ws2, [2], "de", "ja")
ok(read(os.path.join(ws2, "slide_2_ja.txt")).startswith("[de->ja]"), "translate de -> ja")
pn.step_translate_notes(ws2, [2], "de", "xx")
ok(not os.path.exists(os.path.join(ws2, "slide_2_xx.txt")), "unsupported target language handled")

# ---------------------------------------------------------------- scan
pn._ensure_nltk_data = lambda: None
sys.modules["nltk.corpus"] = types.SimpleNamespace(
    stopwords=types.SimpleNamespace(words=lambda *a: ["the", "is", "about", "we", "and", "its", "in"]),
    words=types.SimpleNamespace(words=lambda: ["today", "talk", "structure", "cell", "slide"]))
ws3 = os.path.join(d, "ws3"); os.makedirs(ws3)
write(os.path.join(ws3, "slide_1_ja.txt"), "今日はDNAとPCRとGFPの話です。")
write(os.path.join(ws3, "slide_2_ja.txt"), "CRISPRについて。")
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
        seen["max_new_tokens"] = kw.get("max_new_tokens")
        return [np.zeros(2400, dtype="float32")], 24000
pn._load_qwen3_model = lambda size, device: FakeModel()
pn.step_generate_audio_qwen3(ws, [1, 2], "de", "r.wav", reftxt, R_de, "qwen3-1.7B", "1.7B", "auto")
ok(seen.get("languages") == {"German"} and os.path.exists(os.path.join(ws, "slide_1_de.qwen3-1.7B.m4a")), "Qwen3 German synthesis")
ok(seen.get("max_new_tokens") is not None
   and pn.chunk_token_budget("a" * 140, 14.0) < pn.chunk_token_budget("a" * 700, 14.0)
   and pn.chunk_token_budget("", 14.0) >= 20 * pn.QWEN3_CODEC_HZ,
   "generation is capped at a length the text can justify")

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

ok(pn.apply_dictionary("5\u2019末端と5\u2032末端", [("5'", "ごだっしゅ", "")], "ja") == "ごだっしゅ末端とごだっしゅ末端",
   "typographic apostrophes and primes match a plain dictionary entry")

cmt = os.path.join(d, "commented.csv")
write(cmt, "; readings for the DNA lecture\nstring,replacement,type\nDNA,ディーエヌエー,   ; only in the first slide\n;PCR,ピーシーアール,\nGbp,ギガベースペア,unit\n   ; indented\n")
ok(pn.read_dictionary_file(cmt, "ja") == [("DNA", "ディーエヌエー", ""), ("Gbp", "ギガベースペア", "unit")],
   "; starts a comment anywhere in a dictionary line")
quoted = os.path.join(d, "quoted.csv")
write(quoted, '"#1",ナンバーワン,  ; a label in the figure\nC#,シーシャープ,\n; C,シー,\n')
ok([t for t, _, _ in pn.read_dictionary_file(quoted, "ja")] == ["#1", "C#"],
   "a # inside a term is kept now that only ; starts a comment")

ok(pn.difference_runs("abcdefghij", "abXXXXghij") == (1, 4) and pn.difference_runs("abc", "abc") == (0, 0),
   "difference runs count the places that differ and the longest one")
ok(pn.difference_list("0123456789abcdefghij", "0123456789XXXXefghij", context=3, minimum=2)
   == [(10, "abcd", "XXXX", "789", "efg")], "differences are listed with what was meant, what was heard and context")
_kana = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモ" * 4
ok(pn.kana_sequence_scores(_kana, _kana[:300] + _kana[350:])[0] > 0.85, "a long note survives a 50-character gap by similarity")

c = pn.japanese_compounds("二本鎖の話")
ok(c == {} or "二本鎖" in c, "compounds are proposed with the reading the front end assembles (skipped without pyopenjtalk)")

# ---------------------------------------------------------------- packing
# Deck as PowerPoint writes it: slide 1 has no audio; slide 2 has a recorded narration with trim;
# slide 3 is a copy of slide 2 that shares its media file (its audio link is "NULL").
P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
def audio_pic(link, embed, trim):
    return ('<p:pic><p:nvPicPr><p:cNvPr id="4" name="Audio 3"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
            f'<p:nvPr><a:audioFile r:link="{link}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}">'
            f'<p14:media {P14} r:embed="{embed}">{trim}</p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="rId90"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="812800" cy="812800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
transition = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice '
              f'{P14} Requires="p14"><p:transition spd="slow" p14:dur="2000" advTm="4289"/></mc:Choice><mc:Fallback>'
              '<p:transition spd="slow" advTm="4289"/></mc:Fallback></mc:AlternateContent>')
REL = '<Relationship Id="{}" Type="{}" Target="{}"{}/>'
media_rels = {
    2: [("rId91", pn.REL_AUDIO, "../media/media1.m4a", ""), ("rId92", pn.REL_MEDIA, "../media/media1.m4a", "")],
    3: [("rId91", pn.REL_AUDIO, "NULL", ' TargetMode="External"'), ("rId92", pn.REL_MEDIA, "../media/media1.m4a", "")],
}
z_in = zipfile.ZipFile(deck)
packed = os.path.join(d, "deck_audio.pptx")
z_out = zipfile.ZipFile(packed, "w", zipfile.ZIP_DEFLATED)
for item in z_in.infolist():
    data = z_in.read(item.filename)
    m = re.match(r"ppt/slides/(_rels/)?slide(\d)\.xml(\.rels)?$", item.filename)
    if item.filename == "[Content_Types].xml":
        data = data.replace(b"<Default ", b'<Default Extension="m4a" ContentType="audio/mp4"/><Default Extension="png" ContentType="image/png"/><Default ', 1)
    elif m and int(m.group(2)) in media_rels:
        n = int(m.group(2))
        if m.group(1):
            rels = "".join(REL.format(*r) for r in media_rels[n] + [("rId90", pn.REL_IMAGE, "../media/image1.png", "")])
            data = data.replace(b"</Relationships>", rels.encode() + b"</Relationships>")
        else:
            trim = '<p14:trim st="700.2867" end="1499.8366"/>'
            timing = pn._NARRATION_TIMING.replace("{spid}", "4")
            data = data.replace(b"</p:spTree>", audio_pic("rId91", "rId92", trim).encode() + b"</p:spTree>")
            data = re.sub(rb"(</p:clrMapOvr>)", lambda mm: mm.group(1) + transition.encode() + timing.encode(), data, count=1)
            recorded = ('<p:extLst><p:ext uri="{3A86A75C-4F4B-4683-9AE1-C65F6400EC91}"><p14:laserTraceLst ' + P14 + '>'
                        '<p14:tracePtLst><p14:tracePt t="48796" x="6062662" y="3259137"/></p14:tracePtLst>'
                        '</p14:laserTraceLst></p:ext><p:ext uri="{E180D4A7-C9FB-4DFB-919C-405C955672EB}">'
                        '<p14:showEvtLst ' + P14 + '><p14:playEvt time="12722" objId="4"/><p14:seekEvt time="38839" '
                        'objId="4" seek="10379"/></p14:showEvtLst></p:ext><p:ext uri="{BB962C8B}">'
                        '<p14:creationId ' + P14 + ' val="1"/></p:ext></p:extLst>')
            data = data.replace(b"</p:sld>", recorded.encode() + b"</p:sld>")
    z_out.writestr(item, data)
AudioSegment.silent(duration=300).export(os.path.join(d, "m.m4a"), format="ipod")
z_out.writestr("ppt/media/media1.m4a", open(os.path.join(d, "m.m4a"), "rb").read())
z_out.writestr("ppt/media/image1.png", pn._speaker_icon_png(8))
z_out.close()
for n, dur in [(1, 1800), (2, 2500), (3, 3200)]:
    AudioSegment.silent(duration=dur).export(os.path.join(ws, f"slide_{n}_de.v4.m4a"), format="ipod")
out_deck = os.path.join(d, "out.pptx")
pn.step_pack_pptx(packed, out_deck, ws, [1, 2, 3], "de", "v4", source_lang="auto", writeback_notes=True)
notes = Presentation(out_deck).slides[0].notes_slide.notes_text_frame.text
info = pn.parse_structured_note(notes)
ok(notes.startswith("=== pptx-narrator: narration [de] from [ja] #") and info and info["narration_lang"] == "de"
   and info["narration_text"].startswith("[ja->de]") and info["source_lang"] == "ja" and info["source_text"] == "塩基対の話です。"
   and info["fingerprint"] == pn.text_fingerprint("塩基対の話です。"), "write-back: marked narration part followed by the original note")
notes2 = Presentation(out_deck).slides[1].notes_slide.notes_text_frame.text
ok(pn.parse_structured_note(notes2)["source_lang"] == "en", "write-back for an English-note slide")

zo = zipfile.ZipFile(out_deck)
names = zo.namelist()
sx = {n: zo.read(f"ppt/slides/slide{n}.xml").decode() for n in (1, 2, 3)}
rx = {n: zo.read(f"ppt/slides/_rels/slide{n}.xml.rels").decode() for n in (1, 2, 3)}
import xml.dom.minidom
for name in names:
    if name.endswith((".xml", ".rels")):
        xml.dom.minidom.parseString(zo.read(name))
ok(names[0] == "[Content_Types].xml", "packed file is a well-formed package")
adv = {n: [int(a) for a in re.findall(r'advTm="(\d+)"', sx[n])] for n in (1, 2, 3)}
ok(all(adv[n] and all(abs(a - dur - 1000) < 100 for a in adv[n]) for n, dur in [(1, 1800), (2, 2500), (3, 3200)]),
   f"slide advance times set to the audio length plus the pause {adv}")
ok('<a:audioFile r:link=' in sx[1] and 'isNarration="1"' in sx[1] and "pptx_narrator_slide1.m4a" in rx[1]
   and re.search(r"</p:clrMapOvr><p:transition[^>]*/><p:timing>", sx[1]), "narration inserted into a slide without audio")
ok("pptx_narrator_slide2.m4a" in rx[2] and "pptx_narrator_slide3.m4a" in rx[3] and "NULL" not in rx[3]
   and "ppt/media/media1.m4a" not in names, "copied slides get their own audio; the unused old clip is removed")
ok("p14:trim" not in sx[2] and "p14:trim" not in sx[3] and sx[2].count("<p:pic>") == 1, "trim of the previous recording removed")
offs = {n: re.findall(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', sx[n]) for n in (1, 2, 3)}
ok(all(any(int(x) < 0 for x, _ in offs[n]) for n in (1, 2, 3)),
   f"audio icon parked outside the slide area {offs}")
pkg2 = tempfile.mkdtemp()
os.makedirs(os.path.join(pkg2, "ppt", "slides"))
write(os.path.join(pkg2, "ppt", "slides", "slide1.xml"),
      '<p:sld><p:cSld><p:spTree>' + audio_pic("rId91", "rId92", "") + '</p:spTree></p:cSld></p:sld>')
pn.embed_slide_narration(pkg2, 1, os.path.join(d, "m.m4a"), 300, icon_outside=False)
ok('<a:off x="0" y="0"/>' in read(os.path.join(pkg2, "ppt", "slides", "slide1.xml")),
   "--keep-audio-icon leaves the icon where it is")
ok(all("laserTraceLst" not in sx[n] and "showEvtLst" not in sx[n] for n in (2, 3)) and "creationId" in sx[2],
   "laser-pointer path and recorded playback events removed, other slide extensions kept")
kept = pn.remove_recorded_show_data(open(os.path.join(d, "recorded.xml"), encoding="utf-8").read()
                                    if False else '<p:sld><p:extLst><p:ext uri="{x}"><p14:laserTraceLst/></p:ext>'
                                    '<p:ext uri="{y}"><p14:showEvtLst/></p:ext></p:extLst></p:sld>', ("pointer",))
ok("laserTraceLst" not in kept[0] and "showEvtLst" in kept[0] and kept[1] == ["laser-pointer path"],
   "--remove-recorded pointer keeps the recorded playback events")
ok(len(Presentation(out_deck).slides) == 3, "packed deck opens with python-pptx")

logs.clear()
pkg = tempfile.mkdtemp()
os.makedirs(os.path.join(pkg, "ppt", "slides"))
write(os.path.join(pkg, "ppt", "slides", "slide1.xml"), '<p:sld><p:cSld><p:spTree></p:spTree></p:cSld><p:timing></p:timing></p:sld>')
ok(pn.embed_slide_narration(pkg, 1, os.path.join(d, "m.m4a"), 300) is None and any("animations" in m for m in logs),
   "slide with animations but no audio object is left alone with a warning")

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
ok(read(os.path.join(ws4, "slide_1_ja.txt")) == "塩基対の話です。" and read(os.path.join(ws4, "slide_1_de.txt")).startswith("[ja->de]"),
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
ok(read(os.path.join(ws5, "slide_1_ja.txt")) == "塩基対とRNAの話です。" and not os.path.exists(os.path.join(ws5, "slide_1_de.txt"))
   and read(os.path.join(ws5, "slide_1_de.stale.txt")).startswith("[ja->de]"), "edited source -> stale narration set aside")
pn.step_translate_notes(ws5, [1], "auto", "de", dictionary=T)
ok("RNA" in read(os.path.join(ws5, "slide_1_de.txt")), "stale narration is translated again")

# spoken-form narration blocks are not restored as translations
spoken_note = pn.compose_structured_note("de", "Boten-RNA", "ja", "mRNAの話", spoken=True)
prs_s = Presentation(); s_ = prs_s.slides.add_slide(prs_s.slide_layouts[5]); s_.notes_slide.notes_text_frame.text = spoken_note
sp = os.path.join(d, "spoken.pptx"); prs_s.save(sp)
ws6 = os.path.join(d, "ws6"); os.makedirs(ws6)
pn.step_extract_notes(sp, ws6, [1], "auto")
ok(sorted(os.listdir(ws6)) == ["slide_1_ja.txt"], "spoken narration block ignored on extraction")

# pack warns when the translation is older than the source note
logs.clear()
write(os.path.join(ws, "slide_1_ja.txt"), "塩基対の話を変更しました。")
pn.step_pack_pptx(packed, os.path.join(d, "out2.pptx"), ws, [1], "de", "v4", source_lang="auto", writeback_notes=True)
note_old = pn.parse_structured_note(Presentation(os.path.join(d, "out2.pptx")).slides[0].notes_slide.notes_text_frame.text)
ok(any("older version of the source note" in m for m in logs) and note_old["fingerprint"] != pn.text_fingerprint(note_old["source_text"]),
   "stale translation flagged at pack time and marked by its original fingerprint")

# ---------------------------------------------------------------- command line
# The CLI writes .pptx_narrator_state.json and .pptx_narrator_resolved.toml into the
# working directory, and "no previous input" only holds where no state exists yet, so
# this section runs in a directory of its own.
deck = os.path.abspath(deck)
cli_cwd = os.path.join(d, "cli"); os.makedirs(cli_cwd); os.chdir(cli_cwd)
# A workspace must exist, and hold text, for the commands that read one.
os.makedirs(os.path.join(cli_cwd, "ws"))
for _lang in ("ja", "de", "nl"):
    write(os.path.join(cli_cwd, "ws", f"slide_1_{_lang}.txt"), "text")
parser = pn.build_parser()
a = parser.parse_args(["translate", "ws", "--in_lang", "JA", "--out-lang", "zh_cn",
                       "--dict-file", "a.csv", "--dict_file", "b.csv"])
ok(a.command == "translate" and a.in_lang == "ja" and a.out_lang == "zh-CN"
   and a.dict_file == ["a.csv", "b.csv"], "CLI normalization, aliases, repeatable --dict-file")

ok([c for c in ("extract", "scan", "translate", "synthesize", "verify", "pack")
    if parser.parse_args([c, deck] if c in ("extract", "pack") else [c, "ws"]).command != c] == [],
   "every pipeline step is a command of its own")

ok(parser.parse_args(["extract", deck]).input == deck
   and parser.parse_args(["extract"]).input is None, "INPUT is positional and optional")

def expect_error(argv, text):
    buf = io.StringIO()
    try:
        with contextlib.redirect_stderr(buf):
            pn.main(argv)
    except SystemExit:
        pass
    ok(text in buf.getvalue(), f"CLI error: {text}")

expect_error(["synthesize", "ws", "--in-lang", "nl", "--engine", "qwen3",
              "--ref-wav", "a", "--ref-text", "b"], "Qwen3-TTS does not support 'nl'")
expect_error(["synthesize", "ws", "--in-lang", "de", "--engine", "gpt_sovits",
              "--ref-wav", "a", "--ref-text", "b"], "GPT-SoVITS does not support")
expect_error(["translate", "ws", "--in-lang", "de", "--out-lang", "de"], "translate requires different")
expect_error(["scan", "ws"], "--in-lang")
expect_error(["extract", deck, "--in-lang", "fr"], "no note in fr")
expect_error(["scan", "--in-lang", "ja", "--dict-file", "d.csv"],
             "requires --workspace")

# built-in defaults -> configuration file -> command line
args = parser.parse_args(["pack", deck])
eff = pn._merge_effective("pack", args, {}, parser)
ok(eff["remove_recorded"] == "all" and eff["slide_pause"] == 1.0
   and pn.RECORDED_CHOICES["none"] == (), "built-in defaults are used when nothing else sets a value")
eff = pn._merge_effective("pack", args, {"pack": {"slide_pause": 2.5}}, parser)
ok(eff["slide_pause"] == 2.5, "the configuration file overrides the built-in default")
eff = pn._merge_effective("pack", args, {"common": {"slide_pause": 3.5}}, parser)
ok(eff["slide_pause"] == 3.5, "a [common] section applies to every command")
args = parser.parse_args(["pack", deck, "--slide-pause", "0.5"])
eff = pn._merge_effective("pack", args, {"pack": {"slide_pause": 2.5}}, parser)
ok(eff["slide_pause"] == 0.5, "the command line overrides the configuration file")
eff = pn._merge_effective("pack", parser.parse_args(["pack", deck, "--out", ""]),
                          {"pack": {"out": "from_config.pptx"}}, parser)
ok(eff["out"] is None, "an empty value on the command line takes a configured value back")
cfg_true = {"pack": {"writeback_notes": True}}
ok(pn._merge_effective("pack", parser.parse_args(["pack", deck]), cfg_true, parser)["writeback_notes"] is True
   and pn._merge_effective("pack", parser.parse_args(["pack", deck, "--no-writeback-notes"]),
                           cfg_true, parser)["writeback_notes"] is False,
   "--no-... takes back a switch set in the configuration file")
ok(pn._normalize_config_keys({"a-b": {"c-d": 1}}) == {"a_b": {"c_d": 1}},
   "hyphenated keys in the configuration file are accepted")
ok(parser.parse_args(["scan", "ws", "--config", "x.toml"]).config == "x.toml",
   "--config may follow the command")
# ---------------------------------------------- workspace state and inheritance
wsi = os.path.join(cli_cwd, "wsi"); os.makedirs(wsi)
write(os.path.join(wsi, "slide_1_ja.txt"), "\u30c6\u30b9\u30c8")
AudioSegment.silent(duration=300).export(os.path.join(wsi, "slide_1_ja.qwen3-1.7B.m4a"), format="ipod")
write(os.path.join(wsi, "slide_1_ja.qwen3-1.7B.spoken.txt"), "\u30c6\u30b9\u30c8")
pn._save_state({"version": 1, "commands": {"synthesize": {
    "resolved_config": {"in_lang": "ja", "engine": "qwen3", "qwen3_model_size": "1.7B",
                        "model": "qwen3-1.7B"}}}}, pn._workspace_path(wsi, pn.STATE_FILE))
st = pn._load_state(pn._workspace_path(wsi, pn.STATE_FILE))
for cmd in ("verify", "pack"):
    a = parser.parse_args([cmd, deck, "--workspace", wsi] if cmd == "pack" else [cmd, wsi])
    eff = pn._inherit_synthesis_settings(cmd, pn._merge_effective(cmd, a, {}, parser), a, {}, st)
    ok(eff["engine"] == "qwen3" and eff["qwen3_model_size"] == "1.7B",
       f"{cmd} inherits the engine of the last synthesis in this workspace")
a = parser.parse_args(["pack", deck, "--workspace", wsi, "--engine", "gpt_sovits"])
eff = pn._inherit_synthesis_settings("pack", pn._merge_effective("pack", a, {}, parser), a, {}, st)
ok(eff["engine"] == "gpt_sovits", "an explicit engine still wins over the workspace state")

ok(os.path.exists(pn._workspace_path(wsi, pn.STATE_FILE))
   and not os.path.exists(os.path.join(cli_cwd, pn.STATE_FILE)),
   "execution state lives in the workspace, not in the current directory")

AudioSegment.silent(duration=300).export(os.path.join(wsi, "slide_1_ja.v2ProPlus.m4a"), format="ipod")
ok(pn._audio_model_labels(wsi, "ja") == ["qwen3-1.7B", "v2ProPlus"],
   "the audio of a workspace is reported by model label")
expect_error(["pack", deck, "--workspace", wsi], "multiple audio model labels")

wse = os.path.join(cli_cwd, "wse"); os.makedirs(wse)
write(os.path.join(wse, "slide_1_ja.txt"), "\u30c6\u30b9\u30c8")
expect_error(["pack", deck, "--workspace", wse, "--in-lang", "ja", "--engine", "qwen3"],
             "found no ja audio")

expect_error(["scan", os.path.join(cli_cwd, "no_such_ws"), "--in-lang", "ja",
              "--dict-file", "d.csv"], "workspace does not exist")
expect_error(["verify", deck, "--workspace", wse, "--in-lang", "ja"],
             "must be a workspace directory or one of its files")
expect_error(["verify", os.path.join(wse, "slide_1_ja.txt"), "--workspace", cli_cwd,
              "--in-lang", "ja"], "is not in --workspace")
ok(pn._is_workspace_file("slide_4_ja.txt")
   and pn._is_workspace_file("slide_4_ja.qwen3-1.7B.m4a")
   and not pn._is_workspace_file("lecture.pptx"),
   "a deck is not a workspace file")

ok(not hasattr(parser.parse_args(["pack", deck]), "use_spoken_notes")
   and "--use-spoken-notes" not in open(os.path.join(os.path.dirname(__file__), "..", "README.md"),
                                        encoding="utf-8").read(),
   "the spoken text is never written into the notes")

ok(pn._rel_to_workspace("/w/ws/slide_1_ja.txt", "/w/ws") == "slide_1_ja.txt"
   and pn._rel_to_workspace("/w/ws", "/w/ws") == "."
   and pn._rel_to_workspace("/elsewhere/deck.pptx", "/w/ws") == "/elsewhere/deck.pptx",
   "a path inside the workspace is recorded relative to it")
snap = {"kind": "directory", "path": "/w/ws",
        "files": [{"path": "/w/ws/slide_1_ja.txt", "sha256": "x"}]}
ok(pn._relativize_snapshot(snap, "/w/ws")["files"][0]["path"] == "slide_1_ja.txt",
   "the files of a directory record are relative too")

ok(pn._suggest_command("verify", "lecture.pptx", "ws", "ja") == "pptx-narrator verify ws --in-lang ja",
   "a deck given to a workspace command suggests the workspace")
ok("extract" in (pn._suggest_command("verify", os.path.join(cli_cwd, "nowhere.pptx"), None, "ja") or ""),
   "with no workspace in sight, extract is suggested first")

ok(all("--slides" in [a.option_strings[0] for a in sp._actions if a.option_strings]
       for name, sp in [(n, sp) for act in parser._actions
                        if isinstance(act, argparse._SubParsersAction)
                        for n, sp in act.choices.items()]),
   "every command can be limited to a slide selection")
ok(pn._select_slides([1, 2, 3, 7], "2,7", parser) == [2, 7]
   and pn._select_slides([1, 2, 3], None, parser) == [1, 2, 3],
   "--slides narrows the slides of a workspace")
expect_error(["synthesize", wse, "--in-lang", "ja", "--slides", "99",
              "--ref-wav", "a", "--ref-text", "b"], "no slide of this workspace matches")

def run_cli(argv):
    """Run the CLI and return (stdout, exit status)."""
    out = io.StringIO()
    status = 0
    try:
        with contextlib.redirect_stdout(out):
            pn.main(argv)
    except SystemExit as e:
        status = e.code if isinstance(e.code, int) else 1
    return out.getvalue(), status

top, status = run_cli(["--help"])
ok(status == 0 and top.count("extract") == 1,
   "--help lists every command once and is not an error")
ok(run_cli(["--version"])[0].strip().endswith(pn.__version__)
   and run_cli(["--version"])[1] == 0,
   "--version prints the version instead of the help")
ok(run_cli([])[1] != 0, "no command at all is still an error")
ok("COMMAND --help" in top, "the help says how to see the options of a command")

sub_help = run_cli(["verify", "--help"])[0]
ok("--in_lang" not in sub_help and "--in-lang" in sub_help,
   "an option is listed once, under its hyphenated spelling")
ok(parser.parse_args(["verify", "ws", "--in_lang", "ja"]).in_lang == "ja",
   "the underscore spelling still works although it is not listed")
undocumented = []
for act in parser._actions:
    if isinstance(act, argparse._SubParsersAction):
        for name, sp in act.choices.items():
            ok(bool(sp.description), f"{name} describes itself in its help")
            undocumented += [f"{name} {a.option_strings[0]}" for a in sp._actions
                             if a.option_strings and a.help is None]
ok(undocumented == [], f"every option has help text (missing: {undocumented})")

ok(parser.parse_args(["synthesize", "ws", "--ref-wav", "a.wav", "--ref-text", "a.txt"]).ref_text == "a.txt",
   "the reference transcript option is --ref-text, matching --ref-wav")

ok(parser.parse_args(["synthesize", "ws", "--ref-w", "a.wav", "--ref-t", "a.txt"]).ref_wav == "a.wav"
   and parser.parse_args(["synthesize", "ws", "--ref-l", "ja"]).ref_lang == "ja"
   and parser.parse_args(["pack", deck, "--out", "x.pptx", "--work", "ws"]).workspace == "ws",
   "an option may be abbreviated as far as it stays unambiguous")
expect_error(["synthesize", "ws", "--ref-", "a"], "ambiguous option")

ok(parser.parse_args(["synthesize", "ws", "--dic", "x.csv"]).dict_file == ["x.csv"],
   "an abbreviation is not ambiguous against the option's own underscore spelling")
ok(parser.parse_args(["synthesize", "ws", "--dict-file=z.csv"]).dict_file == ["z.csv"],
   "--option=value works together with the underscore spelling")

print("ALL TESTS PASSED")
