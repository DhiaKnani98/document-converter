"""
pdf2docx_convert.py
Usage: python pdf2docx_convert.py <input.pdf> <output.docx>

Converts a PDF to DOCX using pdf2docx, preserving layout at best fidelity.
Install: pip install pdf2docx
"""

import sys
import os


def convert(pdf_path: str, docx_path: str) -> None:
    from pdf2docx import Converter

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Input not found: {pdf_path}")

    out_dir = os.path.dirname(docx_path)
    if out_dir and not os.path.isdir(out_dir):
        raise NotADirectoryError(f"Output directory does not exist: {out_dir}")

    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path, start=0, end=None)
    finally:
        cv.close()


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: pdf2docx_convert.py <input.pdf> <output.docx>", file=sys.stderr)
        sys.exit(1)

    pdf_in  = sys.argv[1]
    docx_out = sys.argv[2]

    try:
        convert(pdf_in, docx_out)
        print(f"SUCCESS: {docx_out}")
        sys.exit(0)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
