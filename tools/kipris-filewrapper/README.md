# KIPRIS 해외 file wrapper 정리 도구

KIPRIS 해외특허 → **심사정보조회**에서 받은 zip 파일들을 한 폴더에 모아두고 실행하면:

1. 각 zip 안의 PDF 텍스트에서 **출원번호를 자동으로 찾아** 사건별 폴더로 풉니다
   (zip 파일명에는 출원번호가 없기 때문)
2. 파일명을 `2024-04-10_서류명.pdf` 형식으로 정리합니다
3. **사건목록 / 서류목록 엑셀**(`output/filewrapper_db.xlsx`)을 만듭니다

## 처음 설치 (최초 1회)

```bash
cd tools/kipris-filewrapper
npm install
```

## 사용 방법

1. KIPRIS에서 사건마다 심사정보조회 → 전체 선택 → 다운로드
2. 받은 zip을 전부 `tools/kipris-filewrapper/input/` 폴더에 넣기 (파일명은 그대로 둬도 됨)
3. 실행

```bash
npm run organize
```

결과:

```
output/
  filewrapper_db.xlsx          # 사건목록 / 서류목록 시트 (필터, 파일 링크 포함)
  cases/
    CN_202380012345.6/
      2024-03-15_Request for substantive examination.pdf
      2024-04-10_Notice of international application for entering the national phase of China.pdf
      ...
    _미식별/                    # 출원번호를 못 찾은 zip (아래 참고)
```

- 실행할 때마다 `input/`의 zip 전체로 `output/`을 **새로 만듭니다**. 엑셀을 직접 고칠 거면 다른 곳에 복사해서 쓰세요.
- 같은 사건을 나중에 다시 받아 넣어도 한 폴더로 합쳐지고, 같은 서류는 한 번만 들어갑니다.
  그래서 새 서류가 나오면 그 사건 zip만 새로 받아 `input/`에 추가하고 다시 실행하면 됩니다.

## 자동으로 찾는 정보

| 항목 | 찾는 방법 |
|---|---|
| 출원번호 | PDF 텍스트에서 CN(`申请号 202380012345.6`), US(`17/123,456`), EP(`Application No. 23123456.7`), JP(`特願2023-123456`) 형식을 찾아 가장 많이 나온 번호 채택 |
| PCT번호 | `PCT/KR2023/000001` 형식 |
| 관리번호(추정) | 명세서 머리글 등에 찍힌 사무소 관리번호 (기본 형식: `ABCD240001KRA`처럼 영문+숫자+영문) |
| 서류구분 | 서류명에 등록 관련 단어(grant, allowance, certificate, annual fee 등)가 있으면 **등록**, 사건의 최초 일자 서류는 **출원**, 나머지는 **중간** |

서류구분은 규칙에 따른 자동 추정이므로 엑셀에서 필터로 한 번 훑어보는 것을 권장합니다.

## 출원번호를 못 찾은 경우 (`_미식별`)

스캔 이미지로만 된 PDF(특히 미국 서류)는 텍스트가 없어 번호를 못 찾을 수 있습니다.
이때는 `input/mapping.csv`를 만들어 직접 적어주고 다시 실행하세요
(엑셀에서 "CSV UTF-8"로 저장해도 됩니다). `mapping.example.csv` 참고.

```csv
zip파일명,국가,출원번호,관리번호
20260924110000000.zip,US,"17/123,456",ABCD240001USA
```

mapping.csv에 적은 값은 자동 인식 결과보다 우선합니다 (잘못 인식된 경우 수정용으로도 사용).

## 설정 바꾸기 (선택)

`config.json`을 만들면 기본값을 바꿀 수 있습니다.

```jsonc
{
  // 사무소 관리번호 정규식 (예: PN25005-KINE 형식이면 아래처럼)
  "refPattern": "\\bP[NA][A-Z]{0,2}\\d{5,6}(?:-[A-Z]{2,5})?\\b",
  // "등록"으로 분류할 서류명 단어 (소문자). 지정하면 기본 목록을 대체
  "grantKeywords": ["grant", "allowance", "certificate", "annual fee"]
}
```

## 알려진 한계

- 출원번호 자동 인식은 CN / US / EP / JP 형식만 지원합니다. 그 외 국가는 mapping.csv로 지정하세요.
- 각 PDF의 앞 2페이지만 읽습니다 (속도 때문).
- 이 저장소는 public입니다. `input/`, `output/`은 .gitignore 되어 있으니 다른 곳에 고객 자료를 두고 커밋하지 마세요.
