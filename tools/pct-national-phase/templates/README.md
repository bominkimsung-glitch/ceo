# 템플릿 폴더

국가별 Order Letter / Information Sheet 템플릿(.docx)을 아래 구조로 넣습니다.

```
templates/
  order-letter/
    US.docx
    JP.docx
    EP.docx
  information-sheet/
    US.docx
    JP.docx
```

각 파일은 `{fieldName}` 토큰을 포함한 일반 .docx입니다. 토큰 목록과 만드는 방법은
`../README.md`의 "템플릿 추가/수정" 절 참고.

**주의**: 이 저장소는 public이므로 실제 사무소 양식 파일은 `.gitignore`에 의해 커밋되지 않습니다.
로컬에만 보관하세요.
