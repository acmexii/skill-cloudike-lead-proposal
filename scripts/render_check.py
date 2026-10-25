# -*- coding: utf-8 -*-
"""PPTX(또는 DOCX) → PDF → 페이지 PNG. 만든 제안서를 '눈으로' 확인하고, PDF 는 고객 송부용으로도 쓴다.

PDF 변환 순서: ① Windows MS PowerPoint/Word (COM, PowerShell) ② LibreOffice(soffice).
PNG 렌더는 PyMuPDF(pymupdf) — poppler 불필요.

예)  python render_check.py 제안서.pptx      → 제안서.pdf + 제안서_pages/p01.png ...
     그다음 Read 로 PNG 를 열어 넘침·겹침·정렬을 확인한다.
"""
import os, shutil, subprocess, sys


def _ps(path):
    return path.replace("'", "''")  # PowerShell 작은따옴표 이스케이프


def to_pdf(src):
    src = os.path.abspath(src)
    pdf = os.path.splitext(src)[0] + ".pdf"
    if os.path.exists(pdf):
        os.remove(pdf)
    ext = os.path.splitext(src)[1].lower()
    if os.name == "nt":
        if ext == ".pptx":
            ps = ("$a=New-Object -ComObject PowerPoint.Application;"
                  f"$p=$a.Presentations.Open('{_ps(src)}',-1,0,0);$p.SaveAs('{_ps(pdf)}',32);$p.Close();$a.Quit()")
        else:
            ps = ("$w=New-Object -ComObject Word.Application;$w.Visible=$false;"
                  f"$d=$w.Documents.Open('{_ps(src)}',$false,$true);$d.ExportAsFixedFormat('{_ps(pdf)}',17);"
                  "$d.Close($false);$w.Quit()")
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
        if os.path.exists(pdf):
            return pdf
        print("Office 변환 실패:", r.stderr.strip()[:300], file=sys.stderr)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(src), src],
                       capture_output=True)
        if os.path.exists(pdf):
            return pdf
    sys.exit("PDF 변환 도구가 없습니다 (MS PowerPoint/Word 또는 LibreOffice 필요).")


def to_png(pdf, zoom=1.6):
    import pymupdf
    # 리드 폴더에 _work 가 있으면 확인용 PNG 는 그 안에 둔다 (리드 폴더에는 PPTX·PDF 만)
    base = os.path.splitext(os.path.basename(pdf))[0] + "_pages"
    work = os.path.join(os.path.dirname(pdf), "_work")
    outdir = os.path.join(work, base) if os.path.isdir(work) else os.path.splitext(pdf)[0] + "_pages"
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
