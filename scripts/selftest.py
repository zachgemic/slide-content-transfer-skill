#!/usr/bin/env python3
"""Self-test for check_pairs.py: builds a deck of hidden-original/duplicate pairs with known,
seeded transfer errors and asserts the checker reports exactly the expected verdicts.

Run this before trusting a report on a real deck (and after editing check_pairs.py):
    python selftest.py
Also writes fixture decks for the skill's evals:
    python selftest.py --write-fixtures DIR
"""
import argparse
import copy
import io
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

import check_pairs as cp

BASE = dict(
    title="Revenue grew 42% in 2023",
    bullets=["Enterprise demand drove the increase.", "Churn fell to 3.1% across all segments."],
    footnote="Source: Company filings, FY2023.",
    table=[["Region", "Revenue"], ["EMEA", "$12.4M"], ["APAC", "$8.9M"]],
    chart=[10, 20, 30],
    image="red",
    link=("Read the full report", "https://example.com/report"),
    notes="Say: this is the headline number. Pause for questions.",
    group_text="Callout: Best quarter on record",
)


def png(color):
    buf = io.BytesIO()
    Image.new("RGB", (40, 40), color).save(buf, "PNG")
    buf.seek(0)
    return buf


def add_slide(prs, spec, hidden=False, alt=False, merge=False):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    if hidden:
        s._element.set("show", "0")
    dx = 0.6 if alt else 0.0       # the designer moves things around
    tb = lambda x, y, w, h: s.shapes.add_textbox(Inches(x + dx), Inches(y), Inches(w), Inches(h)).text_frame

    if spec.get("title") is not None:
        tb(0.5, 0.3, 8, 0.8).text = spec["title"]
    bullets = list(spec.get("bullets") or [])
    if bullets:
        if merge and len(bullets) > 1:
            bullets = [" ".join(bullets)]
        tf = tb(0.5, 1.2, 5, 1.5)
        tf.text = bullets[0]
        for b in bullets[1:]:
            tf.add_paragraph().text = b
    if spec.get("footnote") is not None:
        tb(0.5, 6.8, 8, 0.4).text = spec["footnote"]
    if spec.get("table"):
        rows = spec["table"]
        t = s.shapes.add_table(len(rows), len(rows[0]), Inches(5.5 + dx), Inches(1.2), Inches(3.5), Inches(1.2)).table
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                t.cell(r, c).text = v
    if spec.get("chart"):
        cd = CategoryChartData()
        cd.categories = ["Q1", "Q2", "Q3"]
        cd.add_series("Revenue", spec["chart"])
        s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5 + dx), Inches(3.0), Inches(4), Inches(3), cd)
    if spec.get("image"):
        pic = s.shapes.add_picture(png(spec["image"]), Inches(8 + dx), Inches(0.3), Inches(1), Inches(1))
        pic._element.xpath(".//p:cNvPr")[0].set("name", "Logo")
    if spec.get("link"):
        text, url = spec["link"]
        tf = tb(5.5, 5.0, 4, 0.4)
        run = tf.paragraphs[0].add_run()
        run.text = text
        if url:
            run.hyperlink.address = url
    if spec.get("group_text"):
        g = s.shapes.add_group_shape()
        g.shapes.add_textbox(Inches(5.5 + dx), Inches(5.6), Inches(4), Inches(0.5)).text_frame.text = spec["group_text"]
    if spec.get("notes") is not None:
        s.notes_slide.notes_text_frame.text = spec["notes"]
    return s


def mutate(**kw):
    spec = copy.deepcopy(BASE)
    spec.update(kw)
    return spec


# (name, duplicate spec, merge_bullets, expected verdict, categories that must appear)
CASES = [
    ("clean-redesign", mutate(), True, "CLEAN", []),
    ("dropped-footnote", mutate(footnote=None), False, "FAIL", ["text-missing", "words-missing"]),
    ("changed-number-in-title", mutate(title="Revenue grew 24% in 2023"), False, "FAIL", ["text-altered", "numbers-missing"]),
    ("dropped-word-in-bullet", mutate(bullets=["Enterprise demand drove the increase.", "Churn fell to 3.1% across segments."]),
     False, "FAIL", ["text-altered", "words-missing"]),
    ("dropped-notes", mutate(notes=""), False, "FAIL", ["notes-changed"]),
    ("lost-hyperlink", mutate(link=("Read the full report", None)), False, "FAIL", ["link-missing"]),
    ("replaced-image", mutate(image="blue"), False, "REVIEW", ["image-missing"]),
    ("table-cell-changed", mutate(table=[["Region", "Revenue"], ["EMEA", "$12.9M"], ["APAC", "$8.9M"]]), False, "FAIL",
     ["numbers-missing"]),
    ("chart-value-changed", mutate(chart=[10, 25, 30]), False, "FAIL", ["numbers-missing"]),
    ("smart-quote-swap", mutate(group_text="Callout: Best quarter on record!"), False, "REVIEW", ["text-cosmetic"]),
    ("added-text", mutate(footnote="Source: Company filings, FY2023.\nDraft - not for distribution"), False, "REVIEW", ["text-added"]),
    ("dropped-group-text", mutate(group_text=None), False, "FAIL", ["text-missing"]),
]


def build(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
    expected = []
    for name, dup, merge, verdict, cats in CASES:
        orig_spec = mutate()
        if name == "smart-quote-swap":
            orig_spec["group_text"] = "Callout: Best quarter on record"
            dup = mutate(group_text="Callout: Best quarter on record")
            dup["title"] = BASE["title"]
            orig_spec["bullets"] = ["It’s the “best” quarter.", BASE["bullets"][1]]
            dup["bullets"] = ["It's the \"best\" quarter.", BASE["bullets"][1]]
        if name == "added-text":
            dup["footnote"] = BASE["footnote"]
            dup["extra"] = "Draft - not for distribution"
        add_slide(prs, orig_spec, hidden=True)
        add_slide(prs, dup, alt=True, merge=merge)
        if name == "added-text":
            prs.slides[-1].shapes.add_textbox(Inches(0.5), Inches(7), Inches(4), Inches(0.3)).text_frame.text = dup["extra"]
        expected.append((len(prs.slides) - 1, len(prs.slides), name, verdict, cats))
    prs.save(path)
    return expected


def selftest():
    ok = True
    with tempfile.TemporaryDirectory() as td:
        path = os.path.join(td, "selftest.pptx")
        expected = build(path)
        results, notes, _ = cp.run(path)
        by_pair = {(r["original"], r["duplicate"]): r for r in results}
        for a, b, name, verdict, cats in expected:
            r = by_pair.get((a, b))
            found = {f["category"] for f in r["findings"]} if r else set()
            good = r is not None and r["verdict"] == verdict and all(c in found for c in cats)
            ok &= good
            print("%-4s %-26s expected %-6s got %-6s %s" % ("ok" if good else "BAD", name, verdict,
                                                           r["verdict"] if r else "none", sorted(found) if not good else ""))
    print("\nSELFTEST", "PASSED" if ok else "FAILED")
    return ok


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--write-fixtures", metavar="DIR")
    args = ap.parse_args()
    if args.write_fixtures:
        os.makedirs(args.write_fixtures, exist_ok=True)
        build(os.path.join(args.write_fixtures, "all-cases.pptx"))
        print("wrote", args.write_fixtures)
    sys.exit(0 if selftest() else 1)
