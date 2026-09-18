#!/usr/bin/env python3
"""Check how the ASR screening ranks narration that does not match its text.

The check reuses audio that has already been synthesized. For a subset of slides it
alters the *intended* text of the narration (the `.spoken.txt` file that `--verify`
compares against the transcript), which puts the text and the audio out of step in the
same way a synthesis error would, and then runs the ASR check on the result. Since the
altered slides are known, the report can be read as a ranking problem: do the slides
whose audio no longer matches their text end up at the bottom of the report?

    python examples/screening_check.py --workspace ws --target-lang ja \
        --asr-model small --asr-device cuda

It writes a copy of the workspace (<workspace>_screening) and a summary CSV next to it;
the original workspace is not modified.
"""
import argparse
import csv
import os
import random
import re
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import pptx_narrator as pn  # noqa: E402

SENTENCE_END = re.compile(r'(?<=[。！？])|(?<=[.!?])\s+')


def drop_sentence(text, rng):
    parts = [p for p in SENTENCE_END.split(text) if p and p.strip()]
    if len(parts) < 2:
        return None, None
    i = rng.randrange(len(parts))
    return "".join(parts[:i] + parts[i + 1:]), f"dropped a sentence of {len(parts[i])} characters"


def swap_term(text, rng):
    """Replace one long word with another word of the text (a misread technical term)."""
    words = sorted({w for w in re.findall(r'[A-Za-z]{3,}|[ァ-ヴー]{3,}|[一-龥]{2,}', text)}, key=str)
    if len(words) < 2:
        return None, None
    a, b = rng.sample(words, 2)
    if a not in text:
        return None, None
    return text.replace(a, b, 1), f"read '{a}' as '{b}'"


def change_number(text, rng):
    numbers = re.findall(r'\d+', text)
    if not numbers:
        return None, None
    n = rng.choice(numbers)
    other = str((int(n) + 1 + rng.randrange(8)) % (10 ** len(n)))
    return text.replace(n, other, 1), f"read the number {n} as {other}"


def drop_clause(text, rng):
    """Drop a run of about 20 characters (a phrase swallowed by the engine)."""
    if len(text) < 60:
        return None, None
    i = rng.randrange(0, len(text) - 30)
    return text[:i] + text[i + 20:], "dropped 20 characters"


PERTURBATIONS = [drop_sentence, swap_term, change_number, drop_clause]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workspace", required=True)
    ap.add_argument("--target-lang", "--target_lang", dest="target_lang", default="ja")
    ap.add_argument("--model-label", "--model_label", dest="model_label", default=None,
                    help="Engine/model part of the audio file names (inferred by default)")
    ap.add_argument("--asr-model", "--asr_model", dest="asr_model", default="small")
    ap.add_argument("--asr-device", "--asr_device", dest="asr_device", default="cpu")
    ap.add_argument("--threshold", type=float, default=0.85)
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--perturb", type=int, default=None, help="How many slides to alter (default: half)")
    ap.add_argument("--dry-run", action="store_true", help="Only show what would be altered")
    args = ap.parse_args()

    lang, ws = args.target_lang.lower(), args.workspace
    suffix = pn.lang_suffix(lang)
    labels = {m.group(1) for f in os.listdir(ws)
              for m in [re.fullmatch(rf'slide_\d+{re.escape(suffix)}\.(.+)\.m4a', f)] if m}
    label = args.model_label or (sorted(labels)[0] if len(labels) == 1 else None)
    if label is None:
        sys.exit(f"give --model-label; the workspace holds {sorted(labels) or 'no audio'}")

    slides = sorted(int(m.group(1)) for f in os.listdir(ws)
                    for m in [re.fullmatch(rf'slide_(\d+){re.escape(suffix)}\.{re.escape(label)}\.m4a', f)]
                    if m and os.path.exists(os.path.join(ws, pn.spoken_filename(int(m.group(1)), lang, label))))
    if not slides:
        sys.exit("no slides with both audio and spoken text were found")

    rng = random.Random(args.seed)
    n_perturb = args.perturb if args.perturb is not None else max(1, len(slides) // 2)
    texts = {n: open(os.path.join(ws, pn.spoken_filename(n, lang, label)), encoding="utf-8").read().strip()
             for n in slides}

    # Alter as many slides as asked for, skipping those whose text is too short to alter.
    altered, new_texts = {}, {}
    for n in rng.sample(slides, len(slides)):
        if len(altered) >= n_perturb:
            break
        for fn in rng.sample(PERTURBATIONS, len(PERTURBATIONS)):
            candidate, what = fn(texts[n], rng)
            if candidate and candidate != texts[n]:
                altered[n], new_texts[n] = what, candidate
                break
    too_short = [n for n in slides if n not in altered and len(texts[n]) < 40]

    out = os.path.abspath(ws.rstrip("/") + "_screening")
    os.makedirs(out, exist_ok=True)
    for n in slides:
        audio = pn.audio_filename(n, lang, label)
        shutil.copy(os.path.join(ws, audio), os.path.join(out, audio))
        open(os.path.join(out, pn.spoken_filename(n, lang, label)), "w", encoding="utf-8").write(
            new_texts.get(n, texts[n]))

    print(f"{len(slides)} slides, {len(altered)} of them altered"
          + (f" ({len(too_short)} too short to alter)" if too_short else "") + ":")
    for n, what in sorted(altered.items()):
        print(f"  slide {n}: {what}")
    if args.dry_run:
        return

    pn.step_verify_audio(out, slides, lang, label, args.asr_model, args.asr_device, args.threshold)

    report = os.path.join(out, f"verify_report{suffix}.{label}.csv")
    rows = list(csv.DictReader(open(report, encoding="utf-8")))
    ranking = [int(r["slide"]) for r in rows]  # worst first
    scores = {int(r["slide"]): float(r["similarity"]) for r in rows}
    flagged = {int(r["slide"]) for r in rows if r["status"] == "FLAGGED"}

    hit = sorted(altered) and [n for n in ranking[:len(altered)] if n in altered]
    summary = os.path.join(out, "screening_check.csv")
    with open(summary, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["slide", "altered", "how", "similarity", "rank_worst_first", "flagged"])
        for n in ranking:
            w.writerow([n, n in altered, altered.get(n, ""), f'{scores[n]:.3f}',
                        ranking.index(n) + 1, n in flagged])

    def median(v):
        v = sorted(v)
        return v[len(v) // 2] if v else float("nan")

    kept = [n for n in slides if n not in altered]
    print(f"\naltered slides in the {len(altered)} worst positions: {len(hit)}/{len(altered)}")
    print(f"flagged (similarity < {args.threshold}): {len(flagged & set(altered))}/{len(altered)} altered, "
          f"{len(flagged & set(kept))}/{len(kept)} unaltered")
    print(f"median similarity: altered {median([scores[n] for n in altered]):.3f}, "
          f"unaltered {median([scores[n] for n in kept]):.3f}")
    print(f"per-slide detail: {summary}")


if __name__ == "__main__":
    main()
