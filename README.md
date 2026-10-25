# skill-cloudike-lead-proposal

클라우다이크(Cloudike) 같은 기업용 클라우드 스토리지·파일공유 솔루션을 인바운드 리드에 제안하는
**맞춤 사업 제안서(DOCX/PDF)** 를 리드 관리 대장(xlsx)과 미팅 녹취록으로부터 만드는 Claude Code 스킬입니다.
[skill-rnd-proposal](https://github.com/acmexii/skill-rnd-proposal) 의 원칙(요구사항 대비표 · 확인하지 못한 범위 · Stage-Gate · 측정 가능한 지표)을 영업 제안에 옮겼습니다.

## 설치

```bash
git clone https://github.com/acmexii/skill-cloudike-lead-proposal.git ~/.claude/skills/cloudike-lead-proposal
```

## 사용

Claude Code 에서 리드 대장과 녹취록을 주고 요청합니다.

> 에이아이랩스 미팅 녹취록이랑 리드 관리 대장 보고, 10/29에 메일로 보낼 맞춤 제안서 만들어줘.

또는 `/cloudike-lead-proposal`.

## 구성

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 진행 순서와 원칙 |
| `references/cloudike-facts.md` | 제품·인증·요금·기능 팩트시트 (근거 등급 A/B/C, 출처, 기준일) |
| `references/lead-intake.md` | 대장·녹취 해체 — 고객 요구 대응표, 약속 이행 대조표 등 5개 표 |
| `references/structure.md` | 제안서 표준 목차, 절별 규칙, 문체 |
| `references/competitors.md` | 경쟁·대체 솔루션 비교, 포지션, 페인포인트 |
| `scripts/read_leads.py` | 리드 대장 xlsx/csv → JSON |
| `scripts/quote.py` | 공개 요금 기반 요금제 추천·견적 (공개 범위 초과는 `[견적 협의]`) |
| `scripts/docx_builder.py` | DOCX 빌더 (표지·콜아웃·쪽번호, 표 머리행 반복) |
| `scripts/verify_docx.py` | DOCX → PDF → 페이지 PNG (MS Word 또는 LibreOffice + PyMuPDF) |

요구 패키지: `python-docx`, `openpyxl`, `pymupdf`.

## 주의

- 요금·인증 정보는 2026-10-25 cloudike.kr 기준입니다. 스킬은 작업 때마다 공식 페이지를 다시 확인하도록 되어 있습니다.
- 생성물은 **발송 전 검토본**입니다. `[견적 협의]`, `[영업팀 확정]` 자리와 「시연 확인」 항목은 사람이 채우고 확인한 뒤 보내야 합니다.
- 다른 제품에 쓰려면 `references/cloudike-facts.md` 와 `scripts/quote.py` 의 `PLANS` 를 바꾸면 됩니다.
