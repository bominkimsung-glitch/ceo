# PCT 국내단계 진입 패키지 자동화

PCT 출원번호와 진입국가를 입력하면 Order Letter / Information Sheet를 생성하고,
번역문 등 첨부파일과 함께 zip으로 묶어주는 도구입니다.

## 처음 설치 (최초 1회)

```bash
cd tools/pct-national-phase
npm install
npx playwright install chromium
```

그리고 `templates/order-letter/<국가코드>.docx`, `templates/information-sheet/<국가코드>.docx`를
준비해주세요 (템플릿 만드는 방법은 아래 "템플릿 추가/수정" 참고). 저장소에는 포함되어 있지 않습니다
(공개 저장소라 실제 사무소 양식을 커밋하지 않도록 되어 있습니다).

## 사용 순서

### 1. ePCT에서 서류 받아오기

```bash
npm run fetch -- "PCT/KR2025/099656"
```

브라우저 창이 뜨면:
1. WIPO 계정으로 로그인 (아이디/비밀번호/2단계 인증 직접 입력)
2. 해당 PCT 출원의 문서함으로 이동
3. 필요한 서류를 평소처럼 다운로드 (자동으로 `output/PCT_KR2025_099656/downloads/`에 정리됨)
4. 화면에 보이는 서지사항을 `output/PCT_KR2025_099656/case-data.json` 파일에 직접 입력/확인
   (아래 "case-data.json 형식" 참고)
5. 터미널로 돌아와 Enter

### 2. Order Letter / Information Sheet 생성

```bash
npm run generate -- PCT_KR2025_099656 US,JP,EP
```

`case-data.json`의 값을 각 국가 템플릿에 채워서 `output/PCT_KR2025_099656/generated/`에
`<Our Ref>_Order_Letter.docx`, `<Our Ref>_Information_Sheet.docx`를 만듭니다.
해당 국가의 템플릿이 없으면 그 국가만 건너뛰고 계속 진행합니다.

### 3. zip 패키징

```bash
npm run package -- PCT_KR2025_099656 US,JP,EP
```

국가별로 `output/PCT_KR2025_099656/<Our Ref>-<국가>_package.zip`을 만듭니다.
각 zip에는 해당 국가의 Order Letter, Information Sheet, 그리고 `downloads/`에 있던
공통 첨부파일(번역문, IB 폼, ISR 등)이 전부 들어갑니다. 국가별로 불필요한 첨부파일이
섞여 있으면 보내기 전에 사람이 확인 후 빼주세요.

## case-data.json 형식

```jsonc
{
  "pctApplicationNumber": "PCT/KR2025/099656",
  "internationalFilingDate": "March 11, 2025",
  "letterDate": "August 14, 2026",           // Order Letter 상단에 찍히는 발송일
  "refBase": "PN25005-KINE",                  // 국가코드 붙기 전 사건번호 (Our Ref = refBase-국가코드)
  "applicant": { "name": "...", "address": "..." },
  "inventors": [                              // 최대 3명까지 템플릿에 반영됨 (아래 "알려진 한계" 참고)
    { "name": "...", "address": "..." }
  ],
  "titleOfInvention": "...",
  "priorityText": "Korean Patent Application No. ... filed on ... (DAS CODE: ...)",
  "entity": "Small entity",
  "countries": {
    "US": { "deadlineDate": "September 11, 2026", "requestedFilingDate": "On or before September 11, 2026" },
    "JP": { "deadlineDate": "September 11, 2026", "requestedFilingDate": "On or before September 11, 2026" },
    "EP": { "deadlineDate": "October 11, 2026", "requestedFilingDate": "On or before October 11, 2026" }
  }
}
```

## 템플릿 추가/수정

각 템플릿은 일반 .docx이며 `{fieldName}` 형태의 토큰만 실제 값으로 치환됩니다
(docxtemplater 사용). 국가별 문서 목록, 인사말, 수신처, 서명란 등은 그 나라를 담당하는
해외 대리인이 바뀌지 않는 한 건마다 바뀌지 않는 보일러플레이트라서 토큰화하지 않았습니다 —
바뀌면 템플릿 파일을 직접 열어 수정하면 됩니다.

사용 가능한 토큰:
- Order Letter: `{letterDate}` `{pctNo}` `{pctFilingDate}` `{applicantName}` `{ourRef}` `{deadlineDate}` `{requestedFilingDate}`
- Information Sheet: 위와 동일 + `{title}` `{priorityText}` `{entity}` `{applicantAddress}` `{inventor1Name}` `{inventor1Address}` `{inventor2Name}` `{inventor2Address}` `{inventor3Name}` `{inventor3Address}`

새 국가를 추가하려면 그 국가로 보냈던 기존 Order Letter/Information Sheet 파일을 복사해서
case-data.json 예시의 위 값들을 위 토큰으로 바꿔 `templates/order-letter/<국가코드>.docx`,
`templates/information-sheet/<국가코드>.docx`로 저장하면 됩니다.

## 알려진 한계

- 발명자는 템플릿에 3명 슬롯이 고정되어 있습니다. 발명자가 3명보다 많거나 적은 사건은
  템플릿 파일을 열어 표의 행을 직접 추가/삭제해야 합니다.
- 국내단계 진입 마감일(30개월/31개월 등)은 자동 계산하지 않습니다. 국가마다 예외가 많아
  실수 위험이 크므로, `case-data.json`에 사람이 직접 확인한 날짜를 입력하도록 설계했습니다.
- "Bypass Continuation Application" 안내 문구(미국)처럼 사건마다 포함 여부가 달라지는
  문장은 자동으로 켜고 끄지 않습니다. 해당 사건에 필요 없으면 생성된 문서에서 직접 지워주세요.

## 현재 상태

- [x] ePCT 다운로드 보조 스크립트 (`src/fetch-epct.js`)
- [x] 문서 생성 (`src/generate-docs.js`)
- [x] zip 패키징 (`src/build-package.js`)
- [ ] EP Information Sheet 템플릿 미보유 — 추가되면 `templates/information-sheet/EP.docx`로 저장

## 보안 참고

- 이 저장소는 public입니다. 실제 사건 데이터, 다운로드한 서류, 채워진 문서, 실제 사무소 템플릿은
  `.gitignore`에 의해 커밋되지 않도록 되어 있습니다 (`output/`, `templates/**/*.docx`).
- WIPO 계정 로그인은 항상 사람이 직접 수행하며, 이 스크립트는 로그인 정보를 저장하거나 자동 입력하지 않습니다.
