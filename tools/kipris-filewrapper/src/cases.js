// 사건 폴더별로 KIPRIS file wrapper 이력(이력.txt)과 직접 받은 주요 서류 파일을 정리한다.
//
//   <사건루트>/
//     ABCD240001EPA_EP_22123456/      <- 폴더명: 관리번호_국가_출원번호 (관리번호는 생략 가능)
//       이력.txt                        <- KIPRIS 심사목록 표를 복사해 붙여넣은 것 (또는 캡처를 옮겨 적은 것)
//       20240321_Communication ....pdf <- KIPRIS에서 받은 주요 서류 (zip도 가능)
//
// 실행하면 각 사건 폴더에 정리/ 폴더를 만들어 "제출일_국가번호_서류명.pdf"로 복사하고(원본은 그대로),
// 사건루트에 전체 사건의 이력 엑셀(file_wrapper_이력.xlsx)을 만든다.
const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const { entryName, parseDocFileName, safeFileName, loadConfig, FILING_KEYWORDS } = require("./organize");

const HISTORY_FILE = "이력.txt";
const OUT_DIR = "정리";
const XLSX_NAME = "file_wrapper_이력.xlsx";

// ---------------------------------------------------------------------------
// 이력.txt 읽기

// KIPRIS 날짜 표기(2024.02.07 / 2024-02-07 / 20240207)를 2024-02-07 로
function normDate(s) {
  const m = s.match(/(\d{4})[.\-/]?\s?(\d{2})[.\-/]?\s?(\d{2})/);
  return m ? `${m[1]}-${m[2]}-${m[3]}` : "";
}

const HEADER_LABELS = {
  appNo: "출원번호", regNo: "등록번호", pubNo: "공개번호", filed: "출원일자",
  title: "발명의 명칭", applicant: "출원인",
};
const LABEL_SPLIT = new RegExp(`(${Object.values(HEADER_LABELS).join("|")})(?=\\s*[:：\\t]|\\s{2,}|\\s*$)`);

// 상단 서지 정보 줄: "출원번호<TAB>EP.22123456.A<TAB>등록번호<TAB>" 처럼 한 줄에 여러 항목이 올 수 있다.
// 라벨 바로 뒤에 탭/콜론/공백 두 칸이 있어야 서지 정보로 본다 ("출원인변경신고서" 같은 서류명과 구분).
function parseHeaderLine(line, info) {
  if (!new RegExp(`^${LABEL_SPLIT.source}`).test(line)) return false;
  const parts = line.split(LABEL_SPLIT);
  const keyOf = Object.fromEntries(Object.entries(HEADER_LABELS).map(([k, v]) => [v, k]));
  for (let i = 1; i < parts.length; i += 2) {
    const value = (parts[i + 1] || "").replace(/^[\s:：]+/, "").trim();
    if (value) info[keyOf[parts[i]]] = value;
  }
  return true;
}

// 심사목록 표를 드래그 복사하면 "문서<TAB>제출일<TAB>문서그룹<TAB>원문" 형태가 된다.
// 탭이 없으면(캡처를 옮겨 적은 경우 등) 날짜를 기준으로 앞은 서류명, 뒤는 문서그룹으로 나눈다.
function parseHistory(text) {
  const info = {};
  const rows = [];
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.replace(/\u00a0/g, " ").trim();
    if (!line || line.startsWith("#")) continue;
    if (parseHeaderLine(line, info)) continue;

    let title, date, group;
    if (line.includes("\t")) {
      const cells = line.split("\t").map((c) => c.trim());
      const di = cells.findIndex((c) => /^\d{4}[.\-/]\d{2}[.\-/]\d{2}$/.test(c));
      if (di <= 0) continue;
      title = cells.slice(0, di).join(" ");
      date = normDate(cells[di]);
      group = cells.slice(di + 1).filter((c) => c && c !== "원문").join(", ");
    } else {
      const m = line.match(/^(.*?)\s+(\d{4}[.\-/]\d{2}[.\-/]\d{2})\s*(.*)$/);
      if (!m) continue;
      title = m[1];
      date = normDate(m[2]);
      group = m[3].replace(/\s*원문\s*$/, "");
    }
    if (title.trim() === "문서") continue; // 표 머리글
    rows.push({ title: title.trim(), date, group: group.trim() });
  }
  return { info, rows };
}

// ---------------------------------------------------------------------------
// 사건 식별 / 분류

// 폴더명 "ABCD240001EPA_EP_22123456" 또는 "EP_22123456" 또는 이력.txt의 "EP.22123456.A"
function parseCaseId(folderName, info) {
  const fromFolder = folderName.match(/(?:^|_)([A-Z]{2})_([0-9][0-9A-Za-z./,-]*)$/);
  let ref = "";
  let country = "";
  let appNo = "";
  if (fromFolder) {
    country = fromFolder[1];
    appNo = fromFolder[2];
    ref = folderName.slice(0, fromFolder.index).replace(/_$/, "");
  }
  if (!appNo && info.appNo) {
    const m = info.appNo.match(/^([A-Z]{2})[.\s]?(.+?)(?:\.[A-Z]\d?)?$/);
    if (m) {
      country = m[1];
      appNo = m[2];
    } else {
      appNo = info.appNo;
    }
  }
  return { ref, country, appNo };
}

// 서류명 비교용: KIPRIS 파일명은 / ( ) 등 특수문자가 빠진 채로 저장된다
function normTitle(s) {
  return s.toLowerCase().replace(/[^a-z0-9가-힣぀-ヿ一-鿿]/g, "");
}

function classifyRows(rows, grantKeywords) {
  const isFiling = (r) =>
    r.group.includes("출원문서") || FILING_KEYWORDS.some((k) => r.title.toLowerCase().includes(k));
  const filingDates = rows.filter(isFiling).map((r) => r.date).filter(Boolean).sort();
  const filedOn = filingDates[0] || "";
  for (const r of rows) {
    const t = r.title.toLowerCase();
    if (filedOn && r.date < filedOn) r.category = "PCT/출원 전";
    else if (r.date === filedOn) r.category = "출원";
    else if (grantKeywords.some((k) => t.includes(k))) r.category = "등록";
    else r.category = "중간";
  }
}

// ---------------------------------------------------------------------------
// 파일 모으기 / 매칭

// 사건 폴더의 pdf 목록. zip은 안의 pdf를 꺼내 쓴다 (zip 자체는 그대로 둔다).
function collectFiles(caseDir) {
  const files = [];
  for (const name of fs.readdirSync(caseDir)) {
    const full = path.join(caseDir, name);
    if (!fs.statSync(full).isFile()) continue;
    if (/\.pdf$/i.test(name)) {
      files.push({ name, data: () => fs.readFileSync(full) });
    } else if (/\.zip$/i.test(name)) {
      for (const entry of new AdmZip(full).getEntries()) {
        const entryFile = path.basename(entryName(entry));
        if (!entry.isDirectory && /\.pdf$/i.test(entryFile)) {
          files.push({ name: entryFile, data: () => entry.getData(), fromZip: name });
        }
      }
    }
  }
  return files;
}

// 파일명(YYYYMMDD_서류명)으로 이력의 행을 찾는다: 날짜+서류명 -> 서류명만(유일할 때)
function matchRow(file, rows) {
  const parsed = parseDocFileName(file.name);
  const key = normTitle(parsed.title);
  const free = rows.filter((r) => !r.file);
  let hit = free.find((r) => r.date === parsed.date && normTitle(r.title) === key);
  if (!hit && key) {
    const byTitle = free.filter((r) => normTitle(r.title) === key);
    if (byTitle.length === 1) hit = byTitle[0];
  }
  return { parsed, row: hit };
}

// ---------------------------------------------------------------------------
// 사건 1개 처리

function processCase(caseDir, cfg) {
  const folderName = path.basename(caseDir);
  const historyPath = path.join(caseDir, HISTORY_FILE);
  const { info, rows } = fs.existsSync(historyPath)
    ? parseHistory(fs.readFileSync(historyPath, "utf8").replace(/^﻿/, ""))
    : { info: {}, rows: [] };
  const id = parseCaseId(folderName, info);
  const idLabel = `${id.country}${id.appNo.replace(/[^0-9A-Za-z]/g, "")}`;
  rows.sort((a, b) => a.date.localeCompare(b.date));
  classifyRows(rows, cfg.grantKeywords);

  const outDir = path.join(caseDir, OUT_DIR);
  fs.rmSync(outDir, { recursive: true, force: true });
  const files = collectFiles(caseDir);
  if (files.length) fs.mkdirSync(outDir, { recursive: true });

  const extras = []; // 이력에서 못 찾은 파일
  const used = new Set();
  for (const file of files) {
    const { parsed, row } = matchRow(file, rows);
    const date = row ? row.date : parsed.date;
    const title = row ? row.title : parsed.title;
    const stem = safeFileName([date || "날짜미상", idLabel, title].filter(Boolean).join("_"));
    let fileName = `${stem}.pdf`;
    for (let i = 2; used.has(fileName.toLowerCase()); i++) fileName = `${stem} (${i}).pdf`;
    used.add(fileName.toLowerCase());
    fs.writeFileSync(path.join(outDir, fileName), file.data());

    const rel = path.posix.join(folderName, OUT_DIR, fileName);
    if (row) row.file = rel;
    else extras.push({ title, date, group: "", category: "", file: rel, note: "이력에 없는 파일 - 확인 필요" });
  }

  return { folderName, id, info, rows, extras, fileCount: files.length, hasHistory: fs.existsSync(historyPath) };
}

// ---------------------------------------------------------------------------
// 엑셀

async function writeWorkbook(xlsxPath, cases) {
  const wb = new ExcelJS.Workbook();
  const head = (ws) => {
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDE7F3" } };
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columns.length } };
  };

  const cs = wb.addWorksheet("사건목록");
  cs.columns = [
    { header: "관리번호", key: "ref", width: 16 },
    { header: "국가", key: "country", width: 6 },
    { header: "출원번호", key: "appNo", width: 16 },
    { header: "KIPRIS 표기", key: "kipris", width: 16 },
    { header: "출원일", key: "filed", width: 12 },
    { header: "등록번호", key: "regNo", width: 14 },
    { header: "발명의 명칭", key: "title", width: 50 },
    { header: "출원인", key: "applicant", width: 20 },
    { header: "이력 건수", key: "rowCount", width: 9 },
    { header: "정리한 파일", key: "fileCount", width: 10 },
    { header: "최근 서류일", key: "lastDate", width: 12 },
    { header: "최근 서류", key: "lastDoc", width: 45 },
    { header: "폴더", key: "folder", width: 28 },
    { header: "확인 필요", key: "status", width: 40 },
  ];
  head(cs);

  const ds = wb.addWorksheet("이력");
  ds.columns = [
    { header: "관리번호", key: "ref", width: 16 },
    { header: "국가", key: "country", width: 6 },
    { header: "출원번호", key: "appNo", width: 16 },
    { header: "제출일", key: "date", width: 12 },
    { header: "서류명", key: "title", width: 60 },
    { header: "문서그룹(KIPRIS)", key: "group", width: 36 },
    { header: "구분", key: "category", width: 11 },
    { header: "파일", key: "file", width: 60 },
    { header: "비고", key: "note", width: 28 },
  ];
  head(ds);

  for (const c of cases) {
    const last = c.rows[c.rows.length - 1];
    const status = [
      !c.hasHistory && "이력.txt 없음",
      !c.id.appNo && "출원번호 모름 (폴더명을 관리번호_국가_출원번호로)",
      c.extras.length && `이력에 없는 파일 ${c.extras.length}개`,
    ].filter(Boolean).join(" / ");
    const row = cs.addRow({
      ref: c.id.ref, country: c.id.country, appNo: c.id.appNo, kipris: c.info.appNo || "",
      filed: normDate(c.info.filed || "") || c.info.filed || "", regNo: c.info.regNo || "",
      title: c.info.title || "", applicant: c.info.applicant || "",
      rowCount: c.rows.length, fileCount: c.fileCount,
      lastDate: last ? last.date : "", lastDoc: last ? last.title : "",
      folder: c.folderName, status,
    });
    row.getCell("folder").value = { text: c.folderName, hyperlink: c.folderName };

    for (const r of [...c.rows, ...c.extras]) {
      const dr = ds.addRow({ ref: c.id.ref, country: c.id.country, appNo: c.id.appNo, ...r, file: r.file || "" });
      if (r.file) {
        dr.getCell("file").value = { text: path.posix.basename(r.file), hyperlink: r.file };
        dr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF4C2" } };
      }
    }
  }
  await wb.xlsx.writeFile(xlsxPath);
}

// ---------------------------------------------------------------------------

async function main() {
  // 사건 루트: 실행 인자 > config.json 의 casesRoot > tools/kipris-filewrapper/cases
  const cfgPath = path.join(__dirname, "..", "config.json");
  const casesRoot = fs.existsSync(cfgPath) ? JSON.parse(fs.readFileSync(cfgPath, "utf8")).casesRoot : "";
  const root = path.resolve(process.argv[2] || casesRoot || path.join(__dirname, "..", "cases"));
  if (!fs.existsSync(root)) {
    console.error(`사건 폴더가 없습니다: ${root}`);
    process.exit(1);
  }
  const cfg = loadConfig();
  const caseDirs = fs.readdirSync(root, { withFileTypes: true })
    .filter((e) => e.isDirectory() && !e.name.startsWith(".") && !e.name.startsWith("_"))
    .map((e) => path.join(root, e.name))
    .sort();
  if (caseDirs.length === 0) {
    console.error(`사건 폴더(하위 폴더)가 없습니다: ${root}`);
    process.exit(1);
  }

  const cases = [];
  for (const dir of caseDirs) {
    const c = processCase(dir, cfg);
    cases.push(c);
    const matched = c.rows.filter((r) => r.file).length;
    const warn = [!c.hasHistory && "이력.txt 없음", c.extras.length && `이력에 없는 파일 ${c.extras.length}개`]
      .filter(Boolean).join(", ");
    console.log(`- ${c.folderName}: 이력 ${c.rows.length}건, 파일 ${c.fileCount}개 (이력과 연결 ${matched})${warn ? `  [확인: ${warn}]` : ""}`);
  }

  const xlsxPath = path.join(root, XLSX_NAME);
  await writeWorkbook(xlsxPath, cases);
  console.log("");
  console.log(`엑셀: ${xlsxPath}`);
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { parseHistory, parseCaseId, normTitle };
