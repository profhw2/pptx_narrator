#!/usr/bin/env python3
"""Measure which narration errors the ASR check detects.

The transcript of a slide depends only on its audio, so it has to be produced once; any
number of hypothetical narration errors can then be scored against it without
synthesizing or transcribing anything again. This script injects one error of a known
size into the intended text of each slide -- a run of characters deleted, as when the
engine skips a phrase, or replaced by other words, as when it misreads a term -- and
reports how often the check notices, as a function of the size of the error and the
length of the note.

    python examples/screening_check.py --workspace ws --target-lang ja \
        --asr-model small --asr-device cuda

If the workspace already holds a verification report, its transcripts are reused and no
ASR model is loaded (--report to give another one). The workspace is not modified.
"""
import argparse
import csv
import difflib
import os
import random
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import pptx_narrator as pn  # noqa: E402

FILLER = "それからこの場合においてもおよそ同じように考えることができるという点がここでは重要になります"


def delete_run(text, size, rng):
    """Drop `size` characters, as when a phrase is skipped."""
    if len(text) <= size + 10:
        return None
    i = rng.randrange(0, len(text) - size)
    return text[:i] + text[i + size:]


def replace_run(text, size, rng):
    """Replace `size` characters with other words, as when a term is misread."""
    if len(text) <= size + 10:
        return None
    i = rng.randrange(0, len(text) - size)
    filler = (FILLER * (size // len(FILLER) + 1))[:size]
    return text[:i] + filler + text[i + size:]


KINDS = {"deleted": delete_run, "misread": replace_run}


def transcripts_from_report(path):
    with open(path, encoding="utf-8") as f:
        return {int(r["slide"]): r["asr_text"] for r in csv.DictReader(f) if r.get("asr_text")}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workspace", required=True)
    ap.add_argument("--target-lang", "--target_lang", dest="target_lang", default="ja")
    ap.add_argument("--model-label", "--model_label", dest="model_label", default=None)
    ap.add_argument("--report", default=None, help="Verification report to take the transcripts from")
    ap.add_argument("--asr-model", "--asr_model", dest="asr_model", default="small")
    ap.add_argument("--asr-device", "--asr_device", dest="asr_device", default="cpu")
    ap.add_argument("--threshold", type=float, default=0.85)
    ap.add_argument("--sizes", default="3,6,12,25,50,100", help="Error sizes in characters")
    ap.add_argument("--repeats", type=int, default=20, help="Random positions per slide and size")
    ap.add_argument("--seed", type=int, default=1)
    args = ap.parse_args()

    lang, ws = args.target_lang.lower(), args.workspace
    suffix = pn.lang_suffix(lang)
    labels = {m.group(1) for f in os.listdir(ws)
              for m in [re.fullmatch(rf'slide_\d+{re.escape(suffix)}\.(.+)\.m4a', f)] if m}
    label = args.model_label or (sorted(labels)[0] if len(labels) == 1 else None)
    if label is None:
        sys.exit(f"give --model-label; the workspace holds {sorted(labels) or 'no audio'}")

    report = args.report or os.path.join(ws, f"verify_report{suffix}.{label}.csv")
    if not os.path.exists(report):
        print(f"no report at {report}; transcribing the audio first")
        slides = sorted(int(m.group(1)) for f in os.listdir(ws)
                        for m in [re.fullmatch(rf'slide_(\d+){re.escape(suffix)}\.{re.escape(label)}\.m4a', f)] if m)
        pn.step_verify_audio(ws, slides, lang, label, args.asr_model, args.asr_device, args.threshold)
    asr = transcripts_from_report(report)

    # The comparison works on normalized sequences (katakana for Japanese), and the audio
    # never changes, so each text is converted once and every trial is then a string operation.
    is_ja = pn.base_lang(lang) == "ja"
    if is_ja:
        import pyopenjtalk

        def normalize(text):
            return pn._KANA_PUNCT_RE.sub("", pyopenjtalk.g2p(text, kana=True))
    else:
        normalize = pn.normalize_for_comparison

    intended, hypothesis = {}, {}
    for n in sorted(asr):
        p = os.path.join(ws, pn.spoken_filename(n, lang, label))
        if os.path.exists(p):
            text = open(p, encoding="utf-8").read().strip()
            if text:
                intended[n] = normalize(text)
                hypothesis[n] = normalize(asr[n])
    if not intended:
        sys.exit("no spoken-text files found next to the report")

    def similarity(ref, hyp):
        return difflib.SequenceMatcher(None, ref, hyp, autojunk=False).ratio()

    def longest_run(ref, hyp):
        return pn.difference_runs(ref, hyp)[1]

    baseline = {n: similarity(intended[n], hypothesis[n]) for n in intended}
    below = [n for n, s in baseline.items() if s < args.threshold]
    print(f"{len(intended)} slides, {min(len(t) for t in intended.values())}-"
          f"{max(len(t) for t in intended.values())} characters of narration each "
          f"(as compared: {'katakana' if is_ja else 'normalized text'})")
    print(f"without any injected error, {len(below)} slide(s) score below {args.threshold}"
          + (f": {below}" if below else ""))
    noise = sorted(longest_run(intended[n], hypothesis[n]) for n in intended)
    print(f"longest difference without any injected error: median {noise[len(noise) // 2]}, "
          f"worst {noise[-1]} characters -- the recognition noise a length-independent rule must clear")

    for n in below:  # a slide already below the threshold would "detect" everything
        intended.pop(n, None)
    if not intended:
        sys.exit("every slide is already below the threshold; nothing to measure")

    rng = random.Random(args.seed)
    sizes = [int(x) for x in args.sizes.split(",")]
    buckets = [(0, 200), (200, 600), (600, 10 ** 9)]
    trials = []
    for n, text in intended.items():
        for size in sizes:
            for kind, fn in KINDS.items():
                for _ in range(args.repeats):
                    altered = fn(text, size, rng)
                    if altered is None:
                        continue
                    s = similarity(altered, hypothesis[n])
                    trials.append(dict(slide=n, note_length=len(text), size=size, kind=kind,
                                       similarity=round(s, 4), baseline=round(baseline[n], 4),
                                       drop=round(baseline[n] - s, 4), detected=s < args.threshold,
                                       longest_difference=longest_run(altered, hypothesis[n])))

    def rate(rows):
        return f"{sum(r['detected'] for r in rows):>4}/{len(rows):<4}" if rows else "   -    "

    print("\ndetected (similarity < %.2f) by size of the error and length of the note:" % args.threshold)
    print(f"{'error':>7} | " + " | ".join(f"{lo}-{hi if hi < 10**9 else ''} chars".rjust(16) for lo, hi in buckets)
          + " |         all")
    for size in sizes:
        row = [t for t in trials if t["size"] == size]
        cells = [rate([t for t in row if lo <= t["note_length"] < hi]) for lo, hi in buckets]
        print(f"{size:>5} c | " + " | ".join(c.rjust(16) for c in cells) + " | " + rate(row))
    for kind in KINDS:
        rows = [t for t in trials if t["kind"] == kind]
        print(f"{kind:>10}: {rate(rows)}")
    run_threshold = max(noise) + 1
    print(f"\nthe same trials judged by the longest difference instead (more than {run_threshold} characters, "
          f"which no unaltered slide reaches):")
    print(f"{'error':>7} | " + " | ".join(f"{lo}-{hi if hi < 10**9 else ''} chars".rjust(16) for lo, hi in buckets)
          + " |         all")
    for size in sizes:
        row = [t for t in trials if t["size"] == size]
        cells = [f"{sum(t['longest_difference'] > run_threshold for t in row if lo <= t['note_length'] < hi):>4}/"
                 f"{len([t for t in row if lo <= t['note_length'] < hi]):<4}" for lo, hi in buckets]
        hit = sum(t["longest_difference"] > run_threshold for t in row)
        print(f"{size:>5} c | " + " | ".join(c.rjust(16) for c in cells) + f" | {hit:>4}/{len(row):<4}")

    print("\nmedian fall in similarity caused by the error:")
    for size in sizes:
        drops = sorted(t["drop"] for t in trials if t["size"] == size)
        if drops:
            print(f"{size:>5} c | {drops[len(drops) // 2]:.3f}")

    out = os.path.join(os.path.dirname(os.path.abspath(report)), "screening_check.csv")
    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(trials[0].keys()))
        w.writeheader()
        w.writerows(trials)
    print(f"\n{len(trials)} trials written to {out}")


if __name__ == "__main__":
    main()
