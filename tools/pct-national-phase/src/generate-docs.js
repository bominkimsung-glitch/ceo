// case-data.json + 국가별 템플릿(templates/order-letter/<국가>.docx, templates/information-sheet/<국가>.docx)을
// 읽어서 국가별 Order Letter / Information Sheet를 생성한다.
//
// 템플릿에는 {fieldName} 형태의 토큰이 들어 있고, docxtemplater가 실제 값으로 치환한다.
// 템플릿 자체의 문구(수신처, 인사말, 서류 목록 등)는 국가마다 안정적으로 재사용되는 보일러플레이트이므로
// 건별로 값이 바뀌는 항목만 토큰화되어 있다 (자세한 내용은 README 참고).
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const TEMPLATES_DIR = path.join(__dirname, "..", "templates");

function loadTemplate(kind, countryCode) {
  const file = path.join(TEMPLATES_DIR, kind, `${countryCode}.docx`);
  if (!fs.existsSync(file)) {
    throw new Error(`템플릿을 찾을 수 없습니다: templates/${kind}/${countryCode}.docx`);
  }
  const content = fs.readFileSync(file, "binary");
  const zip = new PizZip(content);
  return new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
}

function renderToBuffer(doc, data) {
  doc.render(data);
  return doc.getZip().generate({ type: "nodebuffer" });
}

function buildOrderLetterData(caseData, countryCode) {
  const c = caseData.countries[countryCode];
  return {
    letterDate: caseData.letterDate,
    pctNo: caseData.pctApplicationNumber,
    pctFilingDate: caseData.internationalFilingDate,
    applicantName: caseData.applicant.name,
    ourRef: `${caseData.refBase}-${countryCode}`,
    deadlineDate: c.deadlineDate,
    requestedFilingDate: c.requestedFilingDate,
  };
}

function buildInfoSheetData(caseData, countryCode) {
  const c = caseData.countries[countryCode];
  const inv = caseData.inventors || [];
  const get = (i, field) => (inv[i] ? inv[i][field] || "" : "");
  return {
    ourRef: `${caseData.refBase}-${countryCode}`,
    title: caseData.titleOfInvention,
    priorityText: caseData.priorityText,
    entity: caseData.entity,
    inventor1Name: get(0, "name"),
    inventor1Address: get(0, "address"),
    inventor2Name: get(1, "name"),
    inventor2Address: get(1, "address"),
    inventor3Name: get(2, "name"),
    inventor3Address: get(2, "address"),
    applicantName: caseData.applicant.name,
    applicantAddress: caseData.applicant.address,
    deadlineDate: c.deadlineDate,
    requestedFilingDate: c.requestedFilingDate,
  };
}

function main() {
  const safeCaseId = process.argv[2];
  const countriesArg = process.argv[3];
  if (!safeCaseId || !countriesArg) {
    console.error("사용법: npm run generate -- <case폴더명> <국가코드,콤마구분>");
    console.error("예:   npm run generate -- PCT_KR2025_099656 US,JP,EP");
    process.exit(1);
  }
  const countries = countriesArg.split(",").map((s) => s.trim().toUpperCase());
  const caseDir = path.join(__dirname, "..", "output", safeCaseId);
  const caseDataPath = path.join(caseDir, "case-data.json");
  if (!fs.existsSync(caseDataPath)) {
    console.error(`case-data.json이 없습니다: ${caseDataPath}`);
    console.error("먼저 npm run fetch 를 실행하거나, 해당 폴더에 case-data.json을 직접 만들어주세요.");
    process.exit(1);
  }
  const caseData = JSON.parse(fs.readFileSync(caseDataPath, "utf-8"));
  const outDir = path.join(caseDir, "generated");
  fs.mkdirSync(outDir, { recursive: true });

  for (const code of countries) {
    if (!caseData.countries || !caseData.countries[code]) {
      console.error(`  [건너뜀] ${code}: case-data.json의 countries.${code} 정보가 없습니다 (deadlineDate, requestedFilingDate 필요).`);
      continue;
    }

    try {
      const ol = loadTemplate("order-letter", code);
      const olBuf = renderToBuffer(ol, buildOrderLetterData(caseData, code));
      const olPath = path.join(outDir, `${caseData.refBase}-${code}_Order_Letter.docx`);
      fs.writeFileSync(olPath, olBuf);
      console.log(`  [생성됨] ${olPath}`);
    } catch (e) {
      console.error(`  [실패] ${code} Order Letter: ${e.message}`);
    }

    try {
      const is = loadTemplate("information-sheet", code);
      const isBuf = renderToBuffer(is, buildInfoSheetData(caseData, code));
      const isPath = path.join(outDir, `${caseData.refBase}-${code}_Information_Sheet.docx`);
      fs.writeFileSync(isPath, isBuf);
      console.log(`  [생성됨] ${isPath}`);
    } catch (e) {
      console.error(`  [실패] ${code} Information Sheet: ${e.message}`);
    }
  }
}

main();
