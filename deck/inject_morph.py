#!/usr/bin/env python3
"""
Inject PowerPoint Morph transitions into a .pptx.

Neither pptxgenjs nor python-pptx exposes slide transitions, so they are added
here as raw OOXML. <p:transition> must sit inside <p:sld>, after <p:clrMapOvr>
and before </p:sld>. Morph itself lives in the p159 (PowerPoint 2016+)
namespace, so it is wrapped in <mc:AlternateContent> with a plain <p:fade/>
fallback — Keynote, Google Slides and pre-2016 PowerPoint read the fallback and
still open the deck.

Morph matches shapes by NAME (<p:cNvPr name="...">, set via pptxgenjs
objectName), not by position, which is why the repeating cards in this deck
share objectNames across slides.

Usage:  python inject_morph.py deck.pptx [-o out.pptx]
"""
import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

# slide number (1-based) -> morph option: byObject | byWord | byChar
MORPH = {
    3: "byWord",    # broken Caesar line re-assembles into the real one
    5: "byWord",    # broken JFK line re-assembles into the real one
    11: "byObject",  # five-item grid collapses into the persistent rail
    12: "byObject",  # rail highlight slides down, panes swap
    13: "byObject",
    14: "byObject",
    15: "byObject",
    20: "byObject",  # rail unfolds back into the grid — the bookend
}
# every other slide gets a quiet fade so the deck feels of one piece
FADE_DEFAULT = True

SPEED = {"byWord": "med", "byObject": "med", "byChar": "med"}

NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/main"


def morph_xml(option: str) -> str:
    dur = 1400 if option == "byWord" else 900
    return (
        '<mc:AlternateContent xmlns:mc="{mc}">'
        '<mc:Choice xmlns:p159="{p159}" Requires="p159">'
        '<p:transition spd="{spd}" p14:dur="{dur}" '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
        '<p159:morph option="{opt}"/>'
        "</p:transition>"
        "</mc:Choice>"
        "<mc:Fallback>"
        '<p:transition spd="med"><p:fade/></p:transition>'
        "</mc:Fallback>"
        "</mc:AlternateContent>"
    ).format(mc=NS_MC, p159=NS_P159, spd=SPEED.get(option, "med"), dur=dur, opt=option)


FADE_XML = (
    '<p:transition spd="med" p14:dur="500" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
    "<p:fade/></p:transition>"
)


def insert(xml: str, transition: str) -> str:
    """Place the transition after <p:clrMapOvr .../>, else before </p:sld>."""
    if "<p:transition" in xml or "p159:morph" in xml:
        raise SystemExit("slide already carries a transition")
    m = re.search(r"<p:clrMapOvr\b[^>]*/>|<p:clrMapOvr\b.*?</p:clrMapOvr>", xml, re.S)
    if m:
        return xml[: m.end()] + transition + xml[m.end():]
    m = re.search(r"</p:sld>\s*$", xml)
    if not m:
        raise SystemExit("no </p:sld> found")
    return xml[: m.start()] + transition + xml[m.start():]


# pptxgenjs auto-names unnamed shapes "Shape 12", "Text 7", ... and those names
# collide across slides by accident. Morph matches on NAME, so a stray collision
# makes PowerPoint tween two unrelated objects. Prefixing every auto name with
# its slide number leaves only the deliberate objectNames matchable.
AUTO_NAME = re.compile(r'(<p:cNvPr id="\d+" name=")((?:Shape|Text|Image|Chart|Table|Media) \d+)(")')


def namespace_auto_names(xml: str, slide_no: int) -> str:
    return AUTO_NAME.sub(lambda m: f"{m.group(1)}s{slide_no} {m.group(2)}{m.group(3)}", xml)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out")
    args = ap.parse_args()

    src = Path(args.pptx)
    out = Path(args.out) if args.out else src
    tmp = src.with_suffix(".morph.tmp.pptx")

    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        slide_nums = sorted(
            int(re.search(r"slide(\d+)\.xml$", n).group(1))
            for n in names
            if re.match(r"ppt/slides/slide\d+\.xml$", n)
        )
        applied = {}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
                if m:
                    n = int(m.group(1))
                    data = namespace_auto_names(data.decode("utf-8"), n).encode("utf-8")
                    opt = MORPH.get(n)
                    if opt:
                        data = insert(data.decode("utf-8"), morph_xml(opt)).encode("utf-8")
                        applied[n] = opt
                    elif FADE_DEFAULT:
                        data = insert(data.decode("utf-8"), FADE_XML).encode("utf-8")
                        applied[n] = "fade"
                zout.writestr(item, data)

    shutil.move(str(tmp), str(out))
    unknown = set(MORPH) - set(slide_nums)
    if unknown:
        print("WARNING: MORPH targets missing slides:", sorted(unknown), file=sys.stderr)
    for n in sorted(applied):
        print(f"  slide {n:>2}  ->  {applied[n]}")
    print(f"transitions written to {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
