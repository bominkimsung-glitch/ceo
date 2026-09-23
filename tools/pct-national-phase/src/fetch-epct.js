// ePCT에서 서류를 받아오는 스크립트.
//
// 설계 원칙: ePCT 화면 구조(버튼 위치 등)를 스크립트가 추측해서 자동으로 누르지 않는다.
// 로그인/2단계 인증/실제 다운로드 버튼 클릭은 전부 사람이 직접 한다.
// 이 스크립트는 "브라우저를 띄우고, 사용자가 무엇을 다운로드하든 자동으로 정해진 폴더에 저장"하는 역할만 한다.
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");
const readline = require("readline");

const EPCT_URL = "https://pct.wipo.int/ePCT/";

async function main() {
  const pctNumber = process.argv[2];
  if (!pctNumber) {
    console.error("사용법: npm run fetch -- \"PCT/KR2024/012345\"");
    process.exit(1);
  }

  const safeName = pctNumber.replace(/[\/\\]/g, "_");
  const caseDir = path.join(__dirname, "..", "output", safeName);
  const downloadDir = path.join(caseDir, "downloads");
  fs.mkdirSync(downloadDir, { recursive: true });

  const caseDataPath = path.join(caseDir, "case-data.json");
  if (!fs.existsSync(caseDataPath)) {
    fs.writeFileSync(
      caseDataPath,
      JSON.stringify(
        {
          pctApplicationNumber: pctNumber,
          internationalFilingDate: "",
          letterDate: "",
          refBase: "",
          applicant: { name: "", address: "" },
          inventors: [
            { name: "", address: "" },
            { name: "", address: "" },
            { name: "", address: "" },
          ],
          titleOfInvention: "",
          priorityText: "",
          entity: "",
          countries: {
            US: { deadlineDate: "", requestedFilingDate: "" },
          },
        },
        null,
        2
      )
    );
  }

  console.log("브라우저를 엽니다...");
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  let downloadCount = 0;
  const attachDownloadHandler = (p) => {
    p.on("download", async (download) => {
      const suggested = download.suggestedFilename() || `document_${Date.now()}`;
      const dest = path.join(downloadDir, suggested);
      await download.saveAs(dest);
      downloadCount += 1;
      console.log(`  [저장됨] ${suggested}`);
    });
  };
  context.on("page", attachDownloadHandler);
  attachDownloadHandler(page);

  await page.goto(EPCT_URL);

  console.log("");
  console.log("=========================================================");
  console.log(" 1) 브라우저 창에서 WIPO 계정으로 로그인해주세요.");
  console.log(` 2) 로그인 후 PCT 출원 "${pctNumber}"의 문서함으로 이동하세요.`);
  console.log(" 3) 필요한 서류(명세서, 우선권서류, 국제조사보고서/IPRP, 번역문 등)를");
  console.log("    평소처럼 하나씩 다운로드하세요.");
  console.log("    -> 다운로드되는 파일은 자동으로 아래 폴더에 저장됩니다:");
  console.log(`       ${downloadDir}`);
  console.log(" 4) 화면에 보이는 서지사항(출원인/발명자/우선권 등)을 아래 파일에");
  console.log("    직접 입력/수정해주세요 (자동으로 채워지지 않습니다 - 정확성을 위해");
  console.log("    사람이 직접 확인하고 입력하도록 설계했습니다):");
  console.log(`       ${caseDataPath}`);
  console.log(" 5) 다 받으셨으면 이 터미널로 돌아와서 Enter 키를 눌러주세요.");
  console.log("=========================================================");
  console.log("");

  await waitForEnter();

  console.log(`\n총 ${downloadCount}개 파일을 저장했습니다: ${downloadDir}`);
  await browser.close();
}

function waitForEnter() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question("완료되셨으면 Enter를 눌러주세요...\n", () => {
      rl.close();
      resolve();
    });
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
