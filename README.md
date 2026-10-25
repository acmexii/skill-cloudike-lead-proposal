# skill-cloudike-lead-proposal

클라우다이크(Cloudike) 같은 기업용 클라우드 스토리지·파일공유 솔루션을 인바운드 리드에 제안하는
**맞춤 사업 제안서(DOCX/PDF)** 를 만드는 Claude Code 스킬입니다.
**리드 관리 대장(xlsx)만 있어도** 제안서를 쓰고, 리드별 폴더에 녹취록·회의록·메일·메모·요구사항 문서가 있으면 찾아서 반영합니다.
[skill-rnd-proposal](https://github.com/acmexii/skill-rnd-proposal) 의 원칙(요구사항 대비표 · 확인하지 못한 범위 · Stage-Gate · 측정 가능한 지표)을 영업 제안에 옮겼습니다.

## 설치

```bash
git clone https://github.com/acmexii/skill-cloudike-lead-proposal.git ~/.claude/skills/cloudike-lead-proposal
```

## 사용

Claude Code 에서 리드 대장이 있는 폴더를 열고 고객사를 지목합니다.

> 리드 관리 대장 보고 대한바이오 제안서 만들어줘.

### 리드 폴더 (선택)

참조 자료는 대장 옆에 고객사 이름이 들어간 폴더를 만들어 넣으면 자동으로 찾습니다(깊이 2까지, `(주)`·날짜 접두어 무관).
대장 폴더 바로 아래에 고객사명이 든 파일(`녹취록_에이아이랩스_….txt`)도 함께 잡힙니다.

```
인바운드 영업 리드 관리 대장.xlsx
에이아이랩스/
  녹취록_도입미팅_20261025.txt
  회의록_보안팀_20261105.docx
  고객메일_요구사항.pdf
```

제안서는 리드 폴더(없으면 새로 만든 `<고객사>/`)에 저장되어, 다음 회차에는 이전 제안서도 참조 자료가 됩니다.
자료가 없으면 대장의 기록으로 쓰고, 채울 수 없는 부분은 「가정」과 「첫 미팅에서 확인할 질문」으로 밝힙니다.

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
| `scripts/find_lead_materials.py` | 리드 폴더·관련 파일 탐색, 유형 분류(녹취록·회의록·메일·RFP 등), 텍스트 추출 |
| `scripts/quote.py` | 공개 요금 기반 요금제 추천·견적 (공개 범위 초과는 `[견적 협의]`) |
| `scripts/docx_builder.py` | DOCX 빌더 (표지·콜아웃·쪽번호, 표 머리행 반복) |
| `scripts/verify_docx.py` | DOCX → PDF → 페이지 PNG (MS Word 또는 LibreOffice + PyMuPDF) |

요구 패키지: `python-docx`, `openpyxl`, `pymupdf`, `pypdf` (pptx 자료가 있으면 `python-pptx`).

## 주의

- 요금·인증 정보는 2026-10-25 cloudike.kr 기준입니다. 스킬은 작업 때마다 공식 페이지를 다시 확인하도록 되어 있습니다.
- 생성물은 **발송 전 검토본**입니다. `[견적 협의]`, `[영업팀 확정]` 자리와 「시연 확인」 항목은 사람이 채우고 확인한 뒤 보내야 합니다.
- 다른 제품에 쓰려면 `references/cloudike-facts.md` 와 `scripts/quote.py` 의 `PLANS` 를 바꾸면 됩니다.
