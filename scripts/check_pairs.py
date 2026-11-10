#!/usr/bin/env python3
"""Deterministic content-transfer check for hidden-original / visible-duplicate slide pairs.

Workflow this supports: a strategist's slide is duplicated, the original is hidden, and the
designer redesigns the duplicate that sits right after it. This script pairs every hidden
slide with the visible slide that follows it, extracts every piece of content from both
(text, tables, charts, SmartArt, notes, hyperlinks, images, alt text, media) and diffs them
with several independent checks, so a miss in one check is likely caught by another.

Usage:
  check_pairs.py deck.pptx [--pairs 3:4,7:8] [--json out.json] [--md out.md] [--render DIR]

Slide numbers are 1-based positions in the deck, hidden slides included (as in PowerPoint's
thumbnail pane). Exit code: 0 = all CLEAN, 1 = something needs REVIEW, 2 = a FAIL.
"""
import argparse
import collections
import difflib
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

from lxml import etree
from pptx import Presentation

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def q(ns, tag):
    return "{%s}%s" % (ns, tag)


# --------------------------------------------------------------------------- normalising

QUOTES = {"‘": "'", "’": "'", "“": '"', "”": '"',
          "–": "-", "—": "-", "−": "-", "…": "..."}


def strict(s):
    """Whitespace-insensitive but otherwise verbatim."""
    s = (s or "").replace(" ", " ").replace("​", "").replace("\v", " ")
    return re.sub(r"\s+", " ", s).strip()


def unify(s):
    for k, v in QUOTES.items():
        s = s.replace(k, v)
    return s


def loose(s):
    """Ignores case, punctuation, quote style and dash style."""
    s = unify(strict(s)).casefold()
    return re.sub(r"[^\w$%€£]+", " ", s).strip()


TOKEN_RE = re.compile(r"[$€£]?\d+(?:,\d{3})*(?:\.\d+)?%?|\w+")


def tokens(s):
    return TOKEN_RE.findall(unify(strict(s)).casefold())


def is_number(tok):
    return bool(re.match(r"^[$€£]?\d", tok))


# --------------------------------------------------------------------------- extraction

def in_fallback(el):
    return any(a.tag == q(MC, "Fallback") for a in el.iterancestors())


def shape_name(el):
    for anc in el.iterancestors():
        if anc.tag in (q(P, "sp"), q(P, "pic"), q(P, "graphicFrame"), q(P, "cxnSp")):
            c = anc.find(".//" + q(P, "cNvPr"))
            if c is not None:
                return c.get("name") or ""
    return ""


def para_text(p):
    out = []
    for ch in p:
        t = etree.QName(ch).localname
        if t == "r":
            out.append(ch.findtext(q(A, "t")) or "")
        elif t == "br":
            out.append(" ")
        elif t == "fld":
            if (ch.get("type") or "").startswith("slidenum"):
                continue  # auto slide numbers legitimately differ
            out.append(ch.findtext(q(A, "t")) or "")
    return "".join(out)


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if hasattr(sh, "shapes"):
            yield from iter_shapes(sh.shapes)


def extract_slide(slide):
    root = slide._element
    inv = {"text": [], "notes": "", "links": [], "images": [], "alt_text": [], "objects": []}

    # all paragraphs anywhere on the slide (text boxes, groups, tables, placeholders)
    for p in root.iter(q(A, "p")):
        if in_fallback(p):
            continue
        txt = para_text(p)
        if not strict(txt):
            continue
        kind = "table" if any(a.tag == q(A, "tc") for a in p.iterancestors()) else "text"
        inv["text"].append({"kind": kind, "where": shape_name(p), "text": strict(txt)})

    # charts
    for sh in iter_shapes(slide.shapes):
        if getattr(sh, "has_chart", False) and sh.has_chart:
            try:
                ch = sh.chart
                if ch.has_title and ch.chart_title.has_text_frame and strict(ch.chart_title.text_frame.text):
                    inv["text"].append({"kind": "chart", "where": sh.name, "text": strict(ch.chart_title.text_frame.text)})
                for plot in ch.plots:
                    cats = [str(c) for c in plot.categories]
                    for s in plot.series:
                        inv["text"].append({"kind": "chart", "where": sh.name, "text": "series %s" % s.name})
                        for c, v in zip(cats, s.values):
                            inv["text"].append({"kind": "chart", "where": sh.name,
                                                "text": "%s | %s = %s" % (s.name, c, "%g" % v if v is not None else "blank")})
            except Exception as e:  # keep going, but make the gap visible
                inv["text"].append({"kind": "chart", "where": sh.name, "text": "[chart unreadable: %s]" % e})

    # SmartArt text lives in a separate diagram-data part
    for gd in root.iter(q(A, "graphicData")):
        if (gd.get("uri") or "").endswith("/diagram") and not in_fallback(gd):
            rel = gd.find(q(DGM, "relIds"))
            if rel is None:
                continue
            try:
                part = slide.part.related_part(rel.get(q(R, "dm")))
                dm = etree.fromstring(part.blob)
                for p in dm.iter(q(A, "p")):
                    txt = para_text(p)
                    if strict(txt):
                        inv["text"].append({"kind": "smartart", "where": shape_name(gd), "text": strict(txt)})
            except Exception as e:
                inv["text"].append({"kind": "smartart", "where": shape_name(gd), "text": "[smartart unreadable: %s]" % e})

    # hyperlinks (text-level and shape-level)
    for el in root.iter(q(A, "hlinkClick")):
        if in_fallback(el):
            continue
        rid, action, target = el.get(q(R, "id")), el.get("action"), None
        if rid:
            rel = slide.part.rels[rid]
            target = rel.target_ref if rel.is_external else "internal:" + str(rel.target_part.partname)
        elif action:
            target = action
        if not target:
            continue
        parent = el.getparent()
        anchor = ""
        if parent is not None and etree.QName(parent).localname in ("rPr", "endParaRPr"):
            run = parent.getparent()
            anchor = strict(run.findtext(q(A, "t")) or "") if run is not None else ""
        inv["links"].append({"url": target, "anchor": anchor or shape_name(el), "where": shape_name(el)})

    # images, including pictures used as shape fills
    for blip in root.iter(q(A, "blip")):
        rid = blip.get(q(R, "embed"))
        if not rid or in_fallback(blip):
            continue
        try:
            sha = hashlib.sha1(slide.part.related_part(rid).blob).hexdigest()
        except Exception:
            continue
        c = None
        for anc in blip.iterancestors():
            c = anc.find(".//" + q(P, "cNvPr"))
            if c is not None:
                break
        inv["images"].append({"sha1": sha, "where": (c.get("name") if c is not None else ""),
                              "alt": (c.get("descr") or "") if c is not None else ""})

    # alt text on any shape
    for c in root.iter(q(P, "cNvPr")):
        if strict(c.get("descr") or "") and not in_fallback(c):
            inv["alt_text"].append({"where": c.get("name") or "", "text": strict(c.get("descr"))})

    # media / embedded objects
    for tag, label in (("videoFile", "video"), ("audioFile", "audio"), ("wavAudioFile", "audio"), ("oleObj", "embedded-object")):
        for el in root.iter(q(A, tag) if tag != "oleObj" else q(P, tag)):
            if not in_fallback(el):
                inv["objects"].append({"kind": label, "where": shape_name(el)})

    # speaker notes
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
        inv["notes"] = strict(" ".join(para_text(p._p) for p in slide.notes_slide.notes_text_frame.paragraphs))
    return inv


# --------------------------------------------------------------------------- comparison

def word_diff(a, b):
    """Inline diff of two strings: {-removed-}{+added+}."""
    aw, bw = a.split(), b.split()
    out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, aw, bw).get_opcodes():
        if op == "equal":
            out.append(" ".join(aw[i1:i2]))
        else:
            if i2 > i1:
                out.append("{-%s-}" % " ".join(aw[i1:i2]))
            if j2 > j1:
                out.append("{+%s+}" % " ".join(bw[j1:j2]))
    return " ".join(out)


def contains_phrase(hay, needle):
    return bool(needle) and bool(re.search(r"(?<!\w)" + re.escape(needle) + r"(?!\w)", hay))


def match_text(ot, dt):
    """Match every original text block to the duplicate. Returns per-block outcomes."""
    ds = [strict(x["text"]) for x in dt]
    dl = [loose(x["text"]) for x in dt]
    os_ = [strict(x["text"]) for x in ot]
    ol = [loose(x["text"]) for x in ot]
    used = [False] * len(dt)
    status = [None] * len(ot)
    res = {"matched": 0, "cosmetic": [], "restructured": [], "altered": [], "missing": [], "added": []}

    for i, s in enumerate(os_):                      # 1) verbatim
        for j, t in enumerate(ds):
            if not used[j] and s == t:
                used[j], status[i] = True, "match"
                res["matched"] += 1
                break
    for i in range(len(ot)):                         # 2) same words, different case/punctuation
        if status[i] or not ol[i]:
            continue
        for j in range(len(dt)):
            if not used[j] and ol[i] == dl[j]:
                used[j], status[i] = True, "cosmetic"
                res["cosmetic"].append({"original": ot[i], "duplicate": dt[j]})
                break
    pool = [j for j in range(len(dt)) if not used[j]]
    joined = " ".join(dl[j] for j in pool)
    for i in range(len(ot)):                         # 3) verbatim text re-flowed / split / merged
        if not status[i] and ol[i] and contains_phrase(joined, ol[i]):
            status[i] = "restructured"
            res["restructured"].append(ot[i])
    for i in range(len(ot)):                         # 4) closest remaining block = altered, else missing
        if status[i]:
            continue
        if not ol[i]:
            continue
        best, bj = 0.0, None
        for j in pool:
            if not used[j]:
                r = difflib.SequenceMatcher(None, ol[i], dl[j]).ratio()
                if r > best:
                    best, bj = r, j
        if bj is not None and best >= 0.5:
            used[bj], status[i] = True, "altered"
            res["altered"].append({"original": ot[i], "duplicate": dt[bj], "similarity": round(best, 2),
                                   "diff": word_diff(os_[i], ds[bj])})
        else:
            status[i] = "missing"
            res["missing"].append(ot[i])
    rjoined = " ".join(ol[i] for i in range(len(ot)) if status[i] == "restructured")
    for j in range(len(dt)):                         # leftover duplicate blocks
        if not used[j] and dl[j] and not contains_phrase(rjoined, dl[j]):
            res["added"].append(dt[j])
    return res


def compare_pair(o, d):
    F = []

    def add(sev, cat, msg, **kw):
        F.append(dict(severity=sev, category=cat, message=msg, **kw))

    m = match_text(o["text"], d["text"])
    for x in m["missing"]:
        add("FAIL", "text-missing", "Text from the original is not in the duplicate", original=x["text"], where=x["where"], kind=x["kind"])
    for x in m["altered"]:
        add("FAIL", "text-altered", "Text was changed", original=x["original"]["text"], duplicate=x["duplicate"]["text"],
            diff=x["diff"], where=x["original"]["where"], kind=x["original"]["kind"])
    for x in m["cosmetic"]:
        add("REVIEW", "text-cosmetic", "Same words, different case/punctuation/quotes",
            original=x["original"]["text"], duplicate=x["duplicate"]["text"], where=x["original"]["where"])
    for x in m["added"]:
        add("REVIEW", "text-added", "Text in the duplicate that is not in the original", duplicate=x["text"], where=x["where"], kind=x["kind"])

    # independent check: words and numbers counted across the whole slide
    otoks = collections.Counter(t for x in o["text"] for t in tokens(x["text"]))
    dtoks = collections.Counter(t for x in d["text"] for t in tokens(x["text"]))
    lost, gained = otoks - dtoks, dtoks - otoks
    ln = {k: v for k, v in lost.items() if is_number(k)}
    lw = {k: v for k, v in lost.items() if not is_number(k)}
    gn = {k: v for k, v in gained.items() if is_number(k)}
    gw = {k: v for k, v in gained.items() if not is_number(k)}
    if ln:
        add("FAIL", "numbers-missing", "Numbers in the original that do not appear (same count) in the duplicate", items=ln)
    if lw:
        add("FAIL", "words-missing", "Words in the original that do not appear (same count) in the duplicate", items=lw)
    if gn:
        add("REVIEW", "numbers-added", "Numbers in the duplicate that are not in the original", items=gn)
    if gw:
        add("REVIEW", "words-added", "Words in the duplicate that are not in the original", items=gw)

    # hyperlinks
    ou, du = collections.Counter(x["url"] for x in o["links"]), collections.Counter(x["url"] for x in d["links"])
    for url, n in (ou - du).items():
        anchors = [x["anchor"] for x in o["links"] if x["url"] == url]
        add("FAIL", "link-missing", "Hyperlink in the original is not in the duplicate", url=url, anchor=anchors[0], count=n)
    for url, n in (du - ou).items():
        add("REVIEW", "link-added", "Hyperlink in the duplicate that is not in the original", url=url)
    for url in set(ou) & set(du):
        oa = {loose(x["anchor"]) for x in o["links"] if x["url"] == url}
        da = {loose(x["anchor"]) for x in d["links"] if x["url"] == url}
        if oa.isdisjoint(da):
            add("REVIEW", "link-anchor", "Hyperlink kept but attached to different text/shape", url=url,
                original=sorted(oa), duplicate=sorted(da))

    # images (bytes identical -> certainly the same asset; anything else must be confirmed by eye)
    oi, di = collections.Counter(x["sha1"] for x in o["images"]), collections.Counter(x["sha1"] for x in d["images"])
    for sha, n in (oi - di).items():
        meta = next(x for x in o["images"] if x["sha1"] == sha)
        add("REVIEW", "image-missing", "Image from the original is not present byte-for-byte in the duplicate "
            "(replaced, redrawn, or dropped?) - confirm visually", image=meta["where"], alt=meta["alt"], count=n)

    # alt text
    da = {loose(x["text"]) for x in d["alt_text"]} | {loose(x["text"]) for x in d["text"]}
    for x in o["alt_text"]:
        if loose(x["text"]) not in da:
            add("REVIEW", "alt-text-missing", "Alt text from the original is not in the duplicate", where=x["where"], original=x["text"])

    # media / embedded objects
    oo, do = collections.Counter(x["kind"] for x in o["objects"]), collections.Counter(x["kind"] for x in d["objects"])
    for k, n in (oo - do).items():
        add("REVIEW", "object-missing", "%s from the original is not in the duplicate" % k, count=n)

    # speaker notes
    if strict(o["notes"]) != strict(d["notes"]):
        if strict(o["notes"]):
            add("FAIL", "notes-changed", "Speaker notes differ from the original", original=o["notes"], duplicate=d["notes"],
                diff=word_diff(o["notes"], d["notes"]))
        else:
            add("REVIEW", "notes-added", "Duplicate has speaker notes that the original did not", duplicate=d["notes"])

    verdict = "FAIL" if any(f["severity"] == "FAIL" for f in F) else ("REVIEW" if F else "CLEAN")
    overlap_o = {t for x in o["text"] for t in tokens(x["text"])}
    overlap_d = {t for x in d["text"] for t in tokens(x["text"])}
    low_overlap = len(overlap_o) >= 3 and len(overlap_o & overlap_d) / max(1, len(overlap_o)) < 0.3
    stats = {
        "text_blocks_checked": len(o["text"]),
        "text_verbatim": m["matched"],
        "text_reflowed": len(m["restructured"]),
        "numbers_in_original": sum(v for k, v in otoks.items() if is_number(k)),
        "words_in_original": sum(v for k, v in otoks.items() if not is_number(k)),
        "links_checked": len(o["links"]), "images_checked": len(o["images"]),
        "alt_texts_checked": len(o["alt_text"]), "objects_checked": len(o["objects"]),
        "notes_chars": len(o["notes"]),
    }
    return {"verdict": verdict, "findings": F, "stats": stats, "reflowed": m["restructured"],
            "pairing_warning": "Very little text overlap - these two slides may not be a real original/duplicate pair" if low_overlap else None}


# --------------------------------------------------------------------------- pairing

def find_pairs(prs, manual=None):
    slides = list(prs.slides)
    hidden = [s._element.get("show") == "0" for s in slides]
    notes = []
    if manual:
        return [(a, b) for a, b in manual], notes
    pairs = []
    for i, h in enumerate(hidden):
        if not h:
            continue
        if i + 1 < len(slides) and not hidden[i + 1]:
            pairs.append((i + 1, i + 2))
        else:
            notes.append("Slide %d is hidden but is not followed by a visible slide - no duplicate to check." % (i + 1))
    paired = {b for _, b in pairs}
    for i, h in enumerate(hidden):
        if not h and (i + 1) not in paired:
            notes.append("Slide %d is visible and has no hidden original before it (not checked)." % (i + 1))
    return pairs, notes


# --------------------------------------------------------------------------- reporting

def fmt_finding(f):
    bits = ["- **%s - %s.**" % (f["severity"], f["message"])]
    for key in ("original", "duplicate", "diff", "url", "anchor", "image", "alt", "where", "items", "count"):
        if key in f and f[key] not in ("", None, {}, []):
            v = f[key]
            if key == "items":
                v = ", ".join("`%s`%s" % (k, " x%d" % n if n > 1 else "") for k, n in sorted(v.items()))
            elif isinstance(v, str):
                v = "`%s`" % v
            bits.append("  - %s: %s" % (key, v))
    return "\n".join(bits)


def markdown(path, results, notes):
    c = collections.Counter(r["verdict"] for r in results)
    out = ["# Slide content check - %s" % os.path.basename(path), "",
           "**%d pair(s) checked: %d CLEAN, %d REVIEW, %d FAIL**" % (len(results), c["CLEAN"], c["REVIEW"], c["FAIL"]), "",
           "| Original | Duplicate | Verdict | FAIL | REVIEW |", "|---|---|---|---|---|"]
    for r in results:
        nf = sum(f["severity"] == "FAIL" for f in r["findings"])
        out.append("| %d | %d | %s | %d | %d |" % (r["original"], r["duplicate"], r["verdict"], nf, len(r["findings"]) - nf))
    for r in results:
        s = r["stats"]
        out += ["", "## Slide %d -> %d: %s" % (r["original"], r["duplicate"], r["verdict"])]
        if r.get("pairing_warning"):
            out.append("> WARNING: " + r["pairing_warning"])
        out.append("Checked: %(text_blocks_checked)d text blocks (%(text_verbatim)d verbatim, %(text_reflowed)d re-flowed), "
                   "%(numbers_in_original)d numbers, %(words_in_original)d words, %(links_checked)d links, "
                   "%(images_checked)d images, %(alt_texts_checked)d alt texts, %(notes_chars)d chars of notes." % s)
        for f in sorted(r["findings"], key=lambda f: f["severity"] != "FAIL"):
            out.append(fmt_finding(f))
        if r["reflowed"]:
            out.append("- INFO - verbatim text present but split/merged/moved between boxes: " + "; ".join("`%s`" % x["text"] for x in r["reflowed"]))
    if notes:
        out += ["", "## Not checked"] + ["- " + n for n in notes]
    return "\n".join(out) + "\n"


# --------------------------------------------------------------------------- rendering

def render(path, outdir, pairs):
    """Render every slide (hidden ones included) and stitch original|duplicate images side by side."""
    os.makedirs(outdir, exist_ok=True)
    tmp = tempfile.mkdtemp()
    try:
        prs = Presentation(path)
        for s in prs.slides:
            if "show" in s._element.attrib:
                del s._element.attrib["show"]
        deck = os.path.join(tmp, "all.pptx")
        prs.save(deck)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, deck],
                       check=True, capture_output=True, timeout=300)
        subprocess.run(["pdftoppm", "-r", "80", "-png", os.path.join(tmp, "all.pdf"), os.path.join(tmp, "s")], check=True)
        imgs = sorted(glob.glob(os.path.join(tmp, "s-*.png")))
        from PIL import Image, ImageDraw
        written = []
        for a, b in pairs:
            ia, ib = Image.open(imgs[a - 1]), Image.open(imgs[b - 1])
            w = ia.width + ib.width + 30
            h = max(ia.height, ib.height) + 30
            canvas = Image.new("RGB", (w, h), "white")
            canvas.paste(ia, (0, 30))
            canvas.paste(ib, (ia.width + 30, 30))
            d = ImageDraw.Draw(canvas)
            d.text((5, 8), "ORIGINAL (slide %d)" % a, fill="black")
            d.text((ia.width + 35, 8), "DUPLICATE (slide %d)" % b, fill="black")
            p = os.path.join(outdir, "pair-%02d-%02d.png" % (a, b))
            canvas.save(p)
            written.append(p)
        return written
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------------------------- main

def run(path, manual=None):
    prs = Presentation(path)
    pairs, notes = find_pairs(prs, manual)
    slides = list(prs.slides)
    results = []
    for a, b in pairs:
        o, d = extract_slide(slides[a - 1]), extract_slide(slides[b - 1])
        r = compare_pair(o, d)
        r.update(original=a, duplicate=b, inventory={"original": o, "duplicate": d})
        results.append(r)
    return results, notes, pairs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--pairs", help="manual pairing, e.g. 3:4,7:8 (original:duplicate, 1-based)")
    ap.add_argument("--json")
    ap.add_argument("--md")
    ap.add_argument("--render", help="directory for side-by-side PNGs of each pair (needs soffice + pdftoppm)")
    ap.add_argument("--only", help="only check pairs whose ORIGINAL slide is in this range/list, e.g. 5-12 or 3,7,9-11")
    ap.add_argument("--list", action="store_true", help="just list the detected pairs (with slide titles) for the user to confirm; no comparison")
    args = ap.parse_args()
    manual = [tuple(int(x) for x in p.split(":")) for p in args.pairs.split(",")] if args.pairs else None
    if args.list:
        prs = Presentation(args.deck)
        pairs, notes = find_pairs(prs, manual)
        slides = list(prs.slides)
        title = lambda s: next((strict(para_text(p)) for p in s._element.iter(q(A, "p")) if strict(para_text(p))), "(no text)")[:60]
        print("%d slides, %d hidden" % (len(slides), sum(s._element.get("show") == "0" for s in slides)))
        for a, b in pairs:
            print("hidden original %d [%s]  ->  designed slide %d [%s]" % (a, title(slides[a - 1]), b, title(slides[b - 1])))
        for n in notes:
            print("NOTE:", n)
        return
    results, notes, pairs = run(args.deck, manual)
    if args.only:
        keep = set()
        for part in args.only.split(","):
            lo, _, hi = part.partition("-")
            keep.update(range(int(lo), int(hi or lo) + 1))
        results = [r for r in results if r["original"] in keep]
        pairs = [p for p in pairs if p[0] in keep]
    md = markdown(args.deck, results, notes)
    if args.json:
        with open(args.json, "w") as f:
            json.dump({"deck": args.deck, "results": results, "notes": notes}, f, indent=2)
    if args.md:
        with open(args.md, "w") as f:
            f.write(md)
    print(md)
    if args.render:
        try:
            for p in render(args.deck, args.render, pairs):
                print("rendered", p)
        except Exception as e:
            print("RENDER FAILED (%s) - visual check must be done another way" % e, file=sys.stderr)
    verdicts = {r["verdict"] for r in results}
    sys.exit(2 if "FAIL" in verdicts else 1 if "REVIEW" in verdicts or not results else 0)


if __name__ == "__main__":
    main()
