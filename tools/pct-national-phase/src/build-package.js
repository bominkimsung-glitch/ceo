// 국가별 생성된 Order Letter / Information Sheet + output/<case>/downloads/ 안의
// 공통 첨부파일(번역문, IB 폼, ISR 등)을 하나의 zip으로 묶는다.
const fs = require("fs");
const path = require("path");
const archiver = require("archiver");

function zipCountry(caseDir, refBase, code) {
  return new Promise((resolve, reject) => {
    const outPath = path.join(caseDir, `${refBase}-${code}_package.zip`);
    const output = fs.createWriteStream(outPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", () => resolve(outPath));
    archive.on("warning", (err) => console.warn(`  [경고] ${err.message}`));
    archive.on("error", reject);
    archive.pipe(output);

    const genDir = path.join(caseDir, "generated");
    const olPath = path.join(genDir, `${refBase}-${code}_Order_Letter.docx`);
    const isPath = path.join(genDir, `${refBase}-${code}_Information_Sheet.docx`);
    if (fs.existsSync(olPath)) archive.file(olPath, { name: path.basename(olPath) });
    if (fs.existsSync(isPath)) archive.file(isPath, { name: path.basename(isPath) });

    const downloadsDir = path.join(caseDir, "downloads");
    if (fs.existsSync(downloadsDir)) {
      archive.directory(downloadsDir, false);
    }

    archive.finalize();
  });
}

async function main() {
  const safeCaseId = process.argv[2];
  const countriesArg = process.argv[3];
  if (!safeCaseId || !countriesArg) {
    console.error("사용법: npm run package -- <case폴더명> <국가코드,콤마구분>");
    console.error("예:   npm run package -- PCT_KR2025_099656 US,JP,EP");
    process.exit(1);
  }
  const countries = countriesArg.split(",").map((s) => s.trim().toUpperCase());
  const caseDir = path.join(__dirname, "..", "output", safeCaseId);
  const caseDataPath = path.join(caseDir, "case-data.json");
  if (!fs.existsSync(caseDataPath)) {
    console.error(`case-data.json이 없습니다: ${caseDataPath}`);
    process.exit(1);
  }
  const caseData = JSON.parse(fs.readFileSync(caseDataPath, "utf-8"));

  for (const code of countries) {
    const outPath = await zipCountry(caseDir, caseData.refBase, code);
    console.log(`  [압축 완료] ${outPath}`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
