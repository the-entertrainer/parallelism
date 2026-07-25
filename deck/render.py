#!/usr/bin/env python3
"""Render every slide of a .pptx to PNG for visual QA.

LibreOffice -> PDF -> PyMuPDF -> PNG (this sandbox has no pdftoppm).
Renders STATIC frames only: slide transitions, including Morph, cannot be
previewed here. Consecutive renders are the before/after keyframes of a morph,
never the animation itself.
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

SOFFICE = Path("/root/.claude/skills/pptx/scripts/office/soffice.py")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--outdir", default=None)
    ap.add_argument("--dpi", type=int, default=110)
    args = ap.parse_args()

    pptx = Path(args.pptx).resolve()
    outdir = Path(args.outdir) if args.outdir else pptx.parent / "render"
    if outdir.exists():
        shutil.rmtree(outdir)
    outdir.mkdir(parents=True)

    subprocess.run(
        [sys.executable, str(SOFFICE), "--headless", "--convert-to", "pdf",
         "--outdir", str(outdir), str(pptx)],
        check=True, capture_output=True,
    )
    pdf = outdir / (pptx.stem + ".pdf")
    if not pdf.exists():
        print("PDF conversion failed", file=sys.stderr)
        return 1

    import fitz

    doc = fitz.open(pdf)
    zoom = args.dpi / 72
    for i, page in enumerate(doc, start=1):
        page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(outdir / f"slide-{i:02d}.png")
    print(f"{doc.page_count} slides -> {outdir}")
    print("NOTE: static frames only — Morph animation is not previewable here.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
