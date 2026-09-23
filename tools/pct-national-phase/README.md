# PCT 국내단계 진입 패키지 자동화

PCT 출원번호와 진입국가를 입력하면 Order Letter / Information Sheet를 생성하고,
번역문 등 첨부파일과 함께 zip으로 묶어주는 도구입니다.

## 처음 설치 (최초 1회)

```bash
cd tools/pct-national-phase
npm install
npx playwright install chromium
```

## 사용 순서

### 1. ePCT에서 서류 받아오기

```bash
npm run fetch -- "PCT/KR2024/012345"
```

브라우저 창이 뜨면:
1. WIPO 계정으로 로그인 (아이디/비밀번호/2단계 인증 직접 입력)
2. 해당 PCT 출원의 문서함으로 이동
3. 필요한 서류를 평소처럼 다운로드 (자동으로 `output/PCT_KR2024_012345/downloads/`에 정리됨)
4. 화면에 보이는 서지사항을 `output/PCT_KR2024_012345/bibliographic-data.json` 파일에 직접 입력
5. 터미널로 돌아와 Enter

### 2. Order Letter / Information Sheet 생성 (예정 — 템플릿 업로드 후 구현)

### 3. zip 패키징 (예정 — 템플릿 업로드 후 구현)

## 현재 상태

- [x] ePCT 다운로드 보조 스크립트 (`src/fetch-epct.js`)
- [ ] 문서 생성 (`src/generate-docs.js`) — 실제 Order Letter / Information Sheet 템플릿 필요 (`templates/` 폴더 참고)
- [ ] zip 패키징 (`src/build-package.js`)

## 보안 참고

- 이 저장소는 public입니다. 실제 사건 데이터, 다운로드한 서류, 채워진 문서, 실제 사무소 템플릿은
  `.gitignore`에 의해 커밋되지 않도록 되어 있습니다 (`output/`, `templates/*.docx`).
- WIPO 계정 로그인은 항상 사람이 직접 수행하며, 이 스크립트는 로그인 정보를 저장하거나 자동 입력하지 않습니다.
