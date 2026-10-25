# -*- coding: utf-8 -*-
"""인바운드 리드 관리 대장(.xlsx/.csv)을 JSON 으로 읽는다.

헤더 행을 자동으로 찾고(‘고객사’가 들어 있는 첫 행), 빈 열은 버린다.
Windows 콘솔(cp949)에서 한글이 깨지지 않도록 UTF-8 로 출력한다.

예)  python read_leads.py "리드 관리 대장.xlsx"                 # 전체
     python read_leads.py "리드 관리 대장.xlsx" --company 에이아이랩스   # 부분 일치 1건
"""
import argparse, csv, datetime, json, sys


def rows_from(path):
    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig") as f:
            return list(csv.reader(f))
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = []
    for ws in wb:
        out += [list(r) for r in ws.iter_rows(values_only=True)]
    return out


def norm(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    return "" if v is None else str(v).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--company")
    a = ap.parse_args()
    rows = rows_from(a.path)
    hi = next(i for i, r in enumerate(rows) if any("고객사" in norm(c) for c in r))
    header = [norm(c) for c in rows[hi]]
    keep = [i for i, h in enumerate(header) if h]
    leads = []
    for r in rows[hi + 1:]:
        rec = {header[i]: norm(r[i]) if i < len(r) else "" for i in keep}
        if any(rec.values()):
            leads.append(rec)
    if a.company:
        key = next(h for h in header if "고객사" in h)
        leads = [l for l in leads if a.company in l.get(key, "")]
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(leads, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
