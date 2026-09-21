import sys
import os

# Empecher le crash 'NoneType' has no attribute 'write' en mode sans console
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

from pathlib import Path
from pdf2docx import Converter
from docx2pdf import convert

def lancer_conversion(file_path_str):
    p = Path(file_path_str)
    if not p.exists():
        return

    ext = p.suffix.lower()

    if ext == ".pdf":
        out_path = p.with_suffix(".docx")
        cv = Converter(str(p))
        cv.convert(str(out_path))
        cv.close()

    elif ext == ".docx":
        out_path = p.with_suffix(".pdf")
        convert(str(p), str(out_path))

if __name__ == "__main__":
    if len(sys.argv) > 1:
        lancer_conversion(sys.argv[1])
