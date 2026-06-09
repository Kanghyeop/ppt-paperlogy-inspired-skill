# ppt-skill 디자인 시스템 · 스펙 스키마

`scripts/render.py`가 강제하는 단일 디자인 시스템과, 너(Claude)가 만들어야 하는 **스펙 JSON**의 정확한 형식.

> 아래 값은 **현재 시점의 디자인 스냅샷**이다. 한 번의 빌드 안에서는 단일 렌더러가 강제(고정)하지만, 디자인 자체는 사용자 피드백을 받으며 계속 진화한다(`render.py`를 다듬는다). "고정"은 통일 방식이지 영구 불변이 아니다.

## 디자인 상수 (현재 스냅샷 — render.py가 강제)
- 배경: `#F4F4F4` (마스터 상속) · 카드: 흰색 `#FFFFFF`
- 텍스트: 잉크 `#222222` · 슬레이트 `#45515E` · 회색 `#8E8E93`
- 포인트(accent): `#1456F0` (config로 교체 가능) · 폰트: **Pretendard** (한글 자간 -5%, 영문 0)
- 16:9 · 프레임: 좌상단 헤더태그, 좌하단 페이지번호, 우하단 캡션, 우상단 your-logo(자동)
- 본문 1장 = 블록 1개. 긴 텍스트는 렌더러가 자동 축소(fit-to-box)하지만, 넘치면 슬라이드를 나눠라.

## 최상위 스펙 형식
```json
{
  "meta": {
    "eyebrow": "REPORT (표지 상단 라벨, 영문 권장)",
    "title1": "표지 큰 제목 1줄",
    "title2": "표지 큰 제목 2줄 (선택)",
    "subtitle": "표지 부제 한 줄",
    "foot": "2026 · 작성/출처",
    "cover": true,
    "thanks": true
  },
  "slides": [ { 본문 슬라이드 객체 }, ... ]
}
```
- `meta`로 표지·감사가 **자동 생성**된다. `cover:false`/`thanks:false`로 끌 수 있고, `thanks`에 `{ "title","subtitle","foot" }`를 주면 문구 변경.
- 본문 슬라이드만 페이지 번호(`NN / total`)가 붙는다.

## 본문 슬라이드 객체
```json
{
  "header": "03  섹션명",      // 2자리 번호 + 공백 2칸 + 한글/영문 섹션명
  "title": "제목 한 줄",        // ≤ 26자 권장
  "subtitle": "부제 한 줄",      // ≤ 42자 권장 (선택)
  "caption": "하단 우측 한 줄",   // 결론/출처 ≤ 40자 (선택)
  "block": { "type": "...", ... }  // 아래 블록 중 하나
}
```

## 블록 타입 (값은 전부 원문에 실제로 있는 것만)

### stat_cards — 핵심 지표 카드 3~4개
```json
{"type":"stat_cards","cards":[
  {"value":"₩2.4조","label":"항목명","sub":"보조설명"} ]}
```
value는 짧은 숫자/금액 권장(길면 자동 축소). 카드 ≤ 4.

### numbered — 번호 목록 4~5개
```json
{"type":"numbered","items":[{"title":"제목","desc":"설명"}]}
```

### bullets — 2열 불릿 (각 열 2~3개)
```json
{"type":"bullets","columns":[[{"head":"소제목","body":"설명"}],[{"head":"..","body":".."}]]}
```

### timeline — 시점 시퀀스 4~6개 (원문에 시기가 있을 때)
```json
{"type":"timeline","steps":[{"label":"2023","title":"이벤트","desc":"설명"}]}
```

### process — 단계 흐름 3~5개
```json
{"type":"process","steps":[{"title":"단계","desc":"설명"}]}
```

### comparison — 2분할 대비
```json
{"type":"comparison","left":{"title":"A","points":["..","..",".."]},
 "right":{"title":"B","points":["..","..",".."]}}
```

### table — 네이티브 표 (마크다운 표 → 이걸로)
```json
{"type":"table","headers":["열1","열2","열3"],"rows":[["a","b","c"]]}
```
헤더 3~5열, 행 ≤ 6, 모든 행 길이 = 헤더 길이.

### kpi_progress — 비율 막대 3~5개 (원문에 % 있을 때만)
```json
{"type":"kpi_progress","items":[{"label":"항목","pct":78,"note":"보조"}]}
```
pct는 number 0~100.

### callout — 강조 한 문장 + 근거 3개
```json
{"type":"callout","big":"핵심 한 문장","points":["근거1","근거2","근거3"]}
```

### column_chart / bar_chart — 막대 (네이티브, 편집 가능)
```json
{"type":"column_chart","unit":"%","categories":["2021","2022"],
 "series":[{"name":"계열","values":[12,34]}]}
```
세로=column, 가로=bar. 계열 1개 권장(여러 개도 가능). values 길이 = categories 길이. **숫자는 원문 데이터만.**

### line_chart — 추이 (1~3계열)
```json
{"type":"line_chart","unit":"지수","categories":["1월","2월"],
 "series":[{"name":"A","values":[10,20]}]}
```

### donut_chart — 구성비
```json
{"type":"donut_chart","unit":"%","labels":["A","B","C"],"values":[40,35,25]}
```
4~6조각. 합 100 권장.

### image — 이미지 (마크다운 ![](path) 위치)
```json
{"type":"image","path":"images/figure1.png","caption":"그림 설명(선택)"}
```
`path`는 **spec.json이 있는 폴더 기준 상대경로**. build.py가 자동으로 절대경로 해석. 종횡비 유지하며 본문 영역에 맞춤.

## 충실도 철칙 (위반 금지)
1. 사실·수치·고유명사·인용·인과·순서 **추가/삭제/변형 0**. 원문에 없는 값 날조 금지.
2. 허용: 슬라이드용 압축·제목화·재배열(의미 보존).
3. 차트/표/카드 숫자는 **원문에 실제로 있는 값만**. 없으면 텍스트 블록(bullets/numbered/callout)으로.
4. 한 슬라이드 한 메시지. 넘치면 분할(디자인 타협 금지).

## 매핑 가이드 (원문 → 블록)
| 원문 요소 | 블록 |
|---|---|
| 마크다운 표 | table (수치 비교 강조면 bar/column_chart) |
| 시계열 수치 | line_chart · 구성비 | donut_chart |
| 핵심 지표 숫자 묶음 | stat_cards |
| 단계/절차 | process · 연/시점 | timeline |
| A vs B | comparison |
| 일반 목록 | bullets · 우선순위/순번 | numbered |
| 한 줄 결론/주장 | callout |
| 비율(% 명시) | kpi_progress |
| 이미지 | image |
