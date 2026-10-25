# -*- coding: utf-8 -*-
"""클라우다이크 공개 요금표 기반 요금제 추천·견적 초안.

공개 단가가 있는 범위만 계산하고, 3TB 초과나 구축형 후보는 금액을 만들지 않고
'[견적 협의]' 로 표시한다. 제안서의 가격 표는 이 스크립트 출력을 그대로 옮긴다.

예)  python quote.py --users 150 --storage-tb 10 --growth-users 180 --growth-tb 15
     python quote.py --users 40 --storage-tb 2 --sme      (바우처 참고 금액 포함)
"""
import argparse, json, sys

# references/cloudike-facts.md 4절과 동일하게 유지할 것 (VAT 별도, 원)
PLANS = [
    # name, tier, monthly, annual_month, annual_total, storage_gb, members(None=무제한)
    ("BASIC", "Team", 48000, 43200, 518400, 50, 10),
    ("STANDARD", "Business", 68000, 61200, 734400, 100, 50),
    ("PROFESSIONAL", "Business", 170000, 153000, 1836000, 300, 100),
    ("PROFESSIONAL+", "Business", 270000, 243000, 2916000, 500, None),
    ("ENTERPRISE1", "Enterprise", 462000, 415800, 4989600, 1024, None),
    ("ENTERPRISE2", "Enterprise", 726000, 653400, 7840800, 2048, None),
    ("ENTERPRISE3", "Enterprise", 957000, 861300, 10335600, 3072, None),
]
PRICE_DATE = "2026-10-25"
ONPREM_USERS = 300  # 이 인원 초과 또는 사내 보관 요건이면 구축형 비교안을 붙인다


def pick(users, storage_gb):
    for p in PLANS:
        name, tier, m, am, at, gb, mem = p
        if gb >= storage_gb and (mem is None or mem >= users):
            return dict(plan=name, tier=tier, monthly=m, annual_month=am, annual_total=at,
                        storage_gb=gb, members=mem or "무제한")
    return None


def scenario(label, users, tb):
    gb = tb * 1024
    p = pick(users, gb)
    s = dict(label=label, users=users, storage_tb=tb)
    if p:
        s.update(p)
        s["per_user_month_annual"] = round(p["annual_month"] / users)
    else:
        s.update(plan="CUSTOM", tier="Enterprise", note="공개 요금 범위(3TB) 초과 → [견적 협의]. 금액을 추정해 쓰지 말 것")
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--users", type=int, required=True)
    ap.add_argument("--storage-tb", type=float, required=True)
    ap.add_argument("--growth-users", type=int)
    ap.add_argument("--growth-tb", type=float)
    ap.add_argument("--onprem", action="store_true", help="고객이 사내 보관·폐쇄망을 언급")
    ap.add_argument("--sme", action="store_true", help="중소기업 → 바우처 참고 금액 표시")
    a = ap.parse_args()

    out = dict(price_basis=f"cloudike.kr/price/all/ 공개 요금, {PRICE_DATE} 기준, VAT 별도",
               scenarios=[scenario("도입 시점", a.users, a.storage_tb)])
    if a.growth_users or a.growth_tb:
        out["scenarios"].append(scenario("증설 시점", a.growth_users or a.users, a.growth_tb or a.storage_tb))
    if a.onprem or a.users > ONPREM_USERS or (a.growth_users or 0) > ONPREM_USERS:
        out["onprem"] = ("구축형 비교안 필요: 사용자·용량 무제한 라이선스, 반복 과금 없음. "
                         "금액은 인프라·커스터마이징 범위에 따라 [견적 협의]")
    if a.sme:
        for s in out["scenarios"]:
            if "annual_total" in s:
                s["voucher_reference"] = ("바우처 선정 시 이용료 약 80% 지원(사이트 안내), 수요기업당 최대 6,910만원. "
                                          f"참고: 연 {s['annual_total']:,}원 → 자부담 약 {round(s['annual_total']*0.2):,}원. "
                                          "선정·공고 일정에 따라 달라짐")
    out["always_add"] = ["추가 계정·용량 증설 단가: [영업팀 확정]",
                         "견적 유효기간·결제 조건: [영업팀 확정]"]
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
