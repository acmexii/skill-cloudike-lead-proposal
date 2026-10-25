# -*- coding: utf-8 -*-
"""DOCX → PDF → 페이지 PNG. 만든 제안서를 '눈으로' 확인하기 위한 도구.

PDF 변환은 ① Windows 의 MS Word(COM, pywin32 또는 PowerShell) ② LibreOffice(soffice) 순으로 시도한다.
PNG 렌더는 PyMuPDF(pymupdf) 를 쓴다(poppler 불필요).

예)  python verify_docx.py 제안서.docx            → 제안서.pdf + 제안서_pages/p01.png ...
     그다음 Read 툴로 PNG 를 열어 표 붕괴·글자 겹침·빈 페이지를 확인한다.
"""
import os, shutil, subprocess, sys


def to_pdf(docx):
    docx = os.path.abspath(docx)
    pdf = os.path.splitext(docx)[0] + ".pdf"
    if os.name == "nt":
        q = lambda s: s.replace("'", "''")  # PowerShell 작은따옴표 이스케이프
        ps = ("$w=New-Object -ComObject Word.Application;$w.Visible=$false;"
              f"$d=$w.Documents.Open('{q(docx)}',$false,$true);$d.ExportAsFixedFormat('{q(pdf)}',17);"
              "$d.Close($false);$w.Quit()")
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
        if os.path.exists(pdf):
            return pdf
        print("Word 변환 실패:", r.stderr.strip()[:300], file=sys.stderr)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(docx), docx],
                       capture_output=True)
        if os.path.exists(pdf):
            return pdf
    sys.exit("PDF 변환 도구가 없습니다 (MS Word 또는 LibreOffice 필요).")


def to_png(pdf, zoom=1.4):
    import pymupdf
    outdir = os.path.splitext(pdf)[0] + "_pages"
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):  # 이전 실행의 PNG 가 남아 쪽수를 착각하지 않게
        if f.startswith("p") and f.endswith(".png"):
            os.remove(os.path.join(outdir, f))
    doc = pymupdf.open(pdf)
    paths = []
    for i, page in enumerate(doc, 1):
        p = os.path.join(outdir, f"p{i:02d}.png")
        page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).save(p)
        paths.append(p)
    return paths


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    pdf = to_pdf(sys.argv[1])
    pages = to_png(pdf)
    print(pdf)
    print(f"{len(pages)} pages →", os.path.dirname(pages[0]))
