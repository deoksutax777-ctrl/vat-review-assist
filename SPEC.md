# 부가세 결재 보조 도구 v2 — 설계서 (SPEC)

산출물: `C:\Users\deoks\.local\bin\vat_review_v2\index.html` **단일 파일**.
사용자: 세무사 본인 1인 (결재권자). 로컬 브라우저에서 실행.

## 0. 포지셔닝

기존 Python `vat_review`(정밀 정량검증 + 결재조서 엑셀 출력)와 역할 분담:

| 계층 | 역할 | 실행 |
|---|---|---|
| 1층 | 매입매출장 xlsx 로컬 파싱·유형별 집계 + 신고서 수기입력 간이대사 | 브라우저 JS, 무료 |
| 2층 | 공제·불공제 정성 스크리닝 (접대비성, 비영업용승용차 등 의심 항목 플래그) | Claude Sonnet API, 버튼 클릭 시에만 |

AI 결과는 **참고용 확인 목록**이지 결론이 아님 — UI에 명시할 것.

## 1. 기술 스택

- 단일 `index.html`, **바닐라 JS** (React/Babel 사용 금지 — 규모가 작고 검토 용이성 우선)
- 외부 의존은 SheetJS 1개만: `<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>`
- **파서는 순수 함수로 분리** — `parseLedger(rows2D)` 는 2차원 배열(시트 전체)을 받아 결과 객체를 반환. DOM/window 의존 금지. 이유: Node 테스트 하니스에서 동일 함수를 검증하기 위함 (§7).
  - HTML 안에 `<script id="parser-core">` 블록으로 격리해 두고, 테스트 하니스는 이 블록을 추출해 실행한다.

## 2. 화면 구성 (단일 페이지, 카드 3스텝)

### 헤더
- 타이틀: `부가세 신고서 결재 보조`
- 업체명(텍스트), 과세기간(select: `YYYY년 1기 예정/확정, 2기 예정/확정`), 사업자유형(select: 법인/개인), 업종(텍스트, 예: 음식점업), 면세사업 겸영 여부(체크)
- 우측 상단 ⚙ 버튼 → API 키 설정 모달 (localStorage key: `vat2_api_key`). 키는 마스킹 표시(`sk-ant-***…뒤4자`), 삭제 버튼 포함.

### STEP 1 — 매입매출장 업로드
- 드래그&드롭 존 + 파일 선택 버튼 (.xlsx)
- 파싱 성공 시: dialect(위하고/세무사랑), 총 행수, 날짜 범위 표시
- **유형별 집계표**: 매출부/매입부 구분, 컬럼 = `코드 | 유형명 | 건수 | 공급가액 | 부가세`. 각 부 소계 행.

### STEP 2 — 신고서 수기 입력 + 간이 대사
- 좌측: 신고서 항목 입력칸 (콤마 자동 포맷, §4 표의 입력 항목)
- 우측: 대사 결과 테이블 `항목 | 신고서 | 장부 집계 | 차이 | 판정`
- 판정: 차이 0 → `적정`(파랑 #00338D), 차이 ≠ 0 → `확인필요`(빨강) + 차이금액 표시
- **음수는 절대 0으로 클램프하지 말 것. 음수 그대로 빨간색 표시.**
- 하단 참고 문구: "예정신고누락분·대손세액 등으로 정당한 차이가 있을 수 있음"

### STEP 3 — AI 공제·불공제 스크리닝
- 옵션: `거래처명 마스킹` 체크박스 (앞 2자 + `***`)
- `[Sonnet 스크리닝 실행]` 버튼 — 클릭 전 예상 비용 표시 (§5 비용 추정)
- 실행 중 스피너, 완료 시 결과 테이블: `심각도 | 거래처 | 계정과목 | 유형 | 금액 | 이슈 | 사유 | 확인 포인트`
- 심각도 색: high=빨강, medium=주황(#C5A44E 계열), low=회색
- 상단에 summary 문단 표시 + "AI 판단은 참고용입니다. 최종 판단은 세무사 검토로 확정하세요." 고지 문구

### 하단 — 결재 요약 카드
- 업체명/기간, STEP2 대사 판정 요약(적정 n건/확인 n건), STEP3 플래그 요약(high n건…)
- `[인쇄]` 버튼 → `window.print()` (인쇄 CSS: 입력 UI 숨기고 결과만)

## 3. 파서 사양 (Python `vat_review/parse_ledger_xlsx.py` 이식 — 로직 동일해야 함)

### 3.1 시트 읽기
- `XLSX.read(arrayBuffer, {type:'array'})` → 첫 워크시트 → `XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null})` 로 2차원 배열 생성
- 참고: 세무사랑 파일은 openpyxl에서 styles.xml 문제가 있었으나 SheetJS는 무관. **반드시 실제 샘플 2종으로 검증**(§7).

### 3.2 헤더 행 탐색
- 위에서부터 스캔, 셀 문자열 trim 후 `'구분'` 이 있고 `'공급가액'`(또는 `'공급 가액'`)이 있는 **첫 행** = 헤더 행
- col_map: 각 헤더 셀에서 공백 전부 제거(`replace(/\s+/g,'')`) 후 `{컬럼명: 인덱스}` (중복 시 첫 것 유지)

### 3.3 dialect 판정
- `col_map['유형'] === 0 && col_map['구분'] === 1` → `'samu'`(세무사랑), 아니면 `'wehago'`(위하고)

### 3.4 컬럼 매핑

| 필드 | 세무사랑(samu) | 위하고(wehago) — `col(name, fallback인덱스)` |
|---|---|---|
| 유형명 | `유형` | (없음 — 매입매출유형에서 추출) |
| 구분(매출/매입) | `구분` | `구분`, 0 |
| 일자 | `일자` | `전표일자`, 1 |
| 품목 | `품목` | `품명`, 5 |
| 공급가액 | `공급가액` | `공급가액`, 6 |
| 부가세 | `부가세` | `부가세`, 7 |
| 합계 | `합계` | `합계`, 8 |
| 전자 | `전자` | `전자세금`, 10 |
| 거래처 | `거래처` | `거래처`, 3 |
| 사업자번호 | (없음) | `사업자번호`, 4 |
| 매입매출유형 | (없음) | `매입매출유형`, 12 |
| 계정과목 | `계정과목명` | `계정과목`, 15 |
| 카드사 | `신용카드사명` | `신용카드사`, 25 |
| 카드번호 | `카드(가맹점)번호` 또는 `카드가맹점번호` | `카드번호`, 26 |

### 3.5 행 필터
1. 전부 null인 행 스킵
2. **합계행 스킵**: 앞 5개 셀 중 문자열에 (공백 제거 후) `합계`/`월계`/`누계` 포함 시 스킵
3. `구분` 셀 trim 값이 `'매출'` 또는 `'매입'`이 아니면 스킵 (위하고 보조헤더도 이걸로 자연 스킵)

### 3.6 유형 코드 결정
- **위하고**: `매입매출유형` 셀을 정규식 `/^(\d{2})\.([가-힣]+)/` 로 파싱 → `(코드, 유형명)`. 매치 실패 시 해당 행 스킵.
- **세무사랑**: `(구분, 유형명)` → 코드 매핑표:

```
매출: 과세11 영세12 면세13 건별14 간이15 수출16 카과17 카면18 카영19 면건20 전자21 현과22 현면23 현영24
매입: 과세51 영세52 면세53 불공54 수입55 카과57 카면58 카영59 면건60 현과61 현면62
```

매핑 실패(코드 undefined) 시 행 스킵.

### 3.7 숫자/집계
- 숫자 변환: null→0, number→정수, 문자열→콤마 제거 후 parseInt (실패 시 0)
- 행 객체: `{side, date, vendor, bizNo, item, amount, tax, total, electronic, typeCode, typeName, account, cardCompany, cardNo}`
- 집계: `byType[code] = {count, amount, tax}` / `totals = {sales:{count,amount,tax}, purchases:{...}}` (code < 50 → sales)
- 반환: `{dialect, headerRowIndex(1-based), rows, byType, totals}`

### 3.8 검증 기준값
`ground_truth.json` (같은 폴더) — Python 파서가 뽑은 정답:
- 위하고: rows=230, [11] n=2/4,200,000/420,000, [51] n=3/1,160,000/116,000, [57] n=225/12,718,026/12,845
- 세무사랑: rows=1292, [11] n=36/5,329,159,277/532,915,927, [51] n=169/1,071,695,631/107,169,551, [53] n=24/26,562,850/0, [54] n=6/79,464,545/7,946,455, [57] n=1055/63,678,463/6,310,246, [58] n=2/260,000/0

**모든 코드의 count/amount/tax가 완전 일치해야 함.**

## 4. 간이 대사 룰 (STEP 2)

신고서 입력 항목과 장부 대응 (장부 = byType 공급가액 합, 세액 항목은 tax 합):

| # | 입력 항목 (신고서) | 장부 대응 | 비고 |
|---|---|---|---|
| (1) | 과세 세금계산서 발급분 공급가액 | [11] amount | |
| (3) | 과세 신용카드·현금영수증 발행분 공급가액 | [17]+[22] amount | |
| (4) | 과세 기타(정규영수증 외) 공급가액 | [14] amount | |
| (5)+(6) | 영세율 합계 공급가액 | [12]+[16] amount | |
| (10)+(12) | 세금계산서수취분 일반+고정자산 공급가액 | **[51]+[54]** amount | 불공 포함이 정답 — [51]만 비교하면 틀림 |
| (41)+(42) | 신용카드매출전표등 수취 공급가액 | [57]+[61] amount | 그밖의공제매입세액 내 신카·현영수증 매입 |
| (50) | 공제받지못할매입세액 해당 공급가액 | [54] amount | |
| (43) | 의제매입세액 공제 세액 | (대응 없음 — 입력값만 요약카드에 표시) | 정밀검증은 Python 툴 몫 |

- 입력하지 않은 칸(공란)은 대사 스킵 (판정 `—`)
- 차이 = 신고서 − 장부. 표시 형식 콤마, 음수 빨강.

## 5. AI 스크리닝 사양 (STEP 3)

### 5.1 전송 데이터 (토큰 절약이 목표)
- 대상: 매입 행(typeCode ≥ 50) 전부 + 매출은 유형별 합계만
- 매입은 `거래처 × 계정과목 × typeCode` 로 그룹핑: `{vendor, account, typeCode, count, amount, tax, items:[대표 품목 최대 2개]}`
- 그룹 500개 초과 시 amount 절대값 상위 500개만 전송하고 UI에 "하위 n개 그룹 생략됨" 명시 (조용한 누락 금지)
- 마스킹 체크 시 vendor를 `앞2자+'***'` 로 치환
- 함께 전송: 업체 업종, 법인/개인, 면세겸영 여부, 과세기간

### 5.2 API 호출

```js
const MODEL = 'claude-sonnet-5';
const res = await fetch('https://api.anthropic.com/v1/messages', {
  method: 'POST',
  headers: {
    'x-api-key': apiKey,
    'anthropic-version': '2023-06-01',
    'content-type': 'application/json',
    'anthropic-dangerous-direct-browser-access': 'true',
  },
  body: JSON.stringify({
    model: MODEL,
    max_tokens: 16000,           // Sonnet 5는 adaptive thinking 기본 ON — 사고 토큰도 max_tokens에 포함되므로 넉넉히
    system: SYSTEM_PROMPT,
    output_config: { effort: 'medium', format: { type: 'json_schema', schema: RESULT_SCHEMA } },
    messages: [{ role: 'user', content: userPayloadJson }],
  }),
});
```

- 응답: `data.content` 중 첫 `type==='text'` 블록의 text를 `JSON.parse` (structured outputs이므로 유효 JSON 보장)
- `data.stop_reason === 'max_tokens'` 이면 "결과가 잘렸습니다" 경고 표시
- 에러 분기: 401(키 확인 안내), 429(잠시 후 재시도 안내), 그 외 status+message 표시. 네트워크 실패 시 "file:// 환경에서 차단될 수 있음 — 로컬 서버(`python -m http.server`)로 열어보세요" 안내 포함.

### 5.3 RESULT_SCHEMA (structured outputs)

```json
{
  "type": "object",
  "properties": {
    "summary": {"type": "string"},
    "flags": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "vendor": {"type": "string"},
          "account": {"type": "string"},
          "type_code": {"type": "integer"},
          "amount": {"type": "integer"},
          "issue_type": {"type": "string", "enum": ["접대비성", "비영업용승용차", "면세관련", "사업무관의심", "불공제누락", "공제누락", "기타"]},
          "severity": {"type": "string", "enum": ["high", "medium", "low"]},
          "reason": {"type": "string"},
          "check_point": {"type": "string"}
        },
        "required": ["vendor", "account", "type_code", "amount", "issue_type", "severity", "reason", "check_point"],
        "additionalProperties": false
      }
    }
  },
  "required": ["summary", "flags"],
  "additionalProperties": false
}
```

### 5.4 SYSTEM_PROMPT 요지 (한국어로 작성)
역할: 한국 부가가치세 신고서 결재를 보조하는 세무 검토자. 매입 내역에서 매입세액 공제·불공제 처리의 적정성을 스크리닝.

판단 기준 (프롬프트에 명시):
1. [57][61] 카과·현과 매입 중 접대비성 의심 — 골프장, 유흥주점, 백화점/마트 상품권, 선물세트 등 거래처·품목 패턴
2. 주유소·정비소·렌터카 + 차량 관련 계정 → 비영업용 소형승용차(개별소비세 과세대상) 매입세액 불공제 검토 필요 플래그
3. [51] 과세매입 중 접대비·기업업무추진비 계정인데 불공[54] 처리 안 된 것 → 불공제누락
4. [54] 불공 행의 사유 타당성 (계정·거래처로 추정)
5. 면세겸영 표시가 있으면 공통매입세액 안분 확인 포인트 제시
6. 반대로 명백히 공제 가능한데 불공 처리된 것 → 공제누락
7. 확신 없는 것은 low로라도 보고 (누락보다 과다보고가 나음 — 하류에서 세무사가 거름)

출력: 스키마대로. reason은 1문장, check_point는 세무사가 원장·증빙에서 확인할 구체 행동 1문장.

### 5.5 비용 추정 (호출 전 표시)
- 추정 입력 토큰 ≈ `JSON.stringify(payload).length / 2` (한글 보수적), 출력 ≈ 3,000 고정 가정
- 비용(원) ≈ `(inTok/1e6*3 + outTok/1e6*15) * 1400` — "약 n원 예상" 표기 (표준단가 기준, 넉넉히)

## 6. 디자인

- 메인 `#00338D`, 포인트 `#C5A44E`, 배경 `#F5F6FA`, 카드 흰색 + 얇은 그림자
- 폰트: `'Malgun Gothic', 'Segoe UI', sans-serif`
- 숫자 셀 우측정렬 + `tabular-nums`, 천단위 콤마
- 헤더 바: #00338D 배경 흰 글자, 포인트 라인 #C5A44E
- 데스크톱 우선(1200px 컨테이너), 인쇄 CSS 별도(@media print — 결과 테이블만)

## 7. 검증 절차 (제작자가 반드시 수행)

1. 임시 폴더에 `npm init -y && npm i xlsx` (SheetJS Node판)
2. 테스트 하니스 `test_parser.mjs`:
   - `index.html`에서 `<script id="parser-core">…</script>` 블록 텍스트를 정규식으로 추출 → `new Function` 또는 임시 .mjs로 저장해 로드
   - 샘플 2종 (`매입매출장_위하고.xlsx`, `매입매출장_세무사랑.xlsx`)을 SheetJS로 읽어 2차원 배열 → `parseLedger()` 실행
   - `ground_truth.json`과 dialect/row_count/byType(count·amount·tax 전 항목) 비교, 불일치는 전부 출력
3. **전 항목 일치할 때까지 수정 반복.** 최종 실행 로그를 결과 보고에 포함.
4. API 실호출 테스트는 하지 않음 (키 없음). 호출부는 코드 정합성만.

## 8. 산출물

- `index.html` — 본체
- (검증용 임시 파일은 scratchpad에, 폴더에 남기지 말 것)
