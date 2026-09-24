// KIPRIS 해외 "심사정보조회"에서 받은 file wrapper zip들을 사건별 폴더로 풀고,
// 사건목록/서류목록 엑셀(DB)을 만든다.
//
// zip 파일명에는 출원번호가 없으므로, 안의 PDF 텍스트에서 출원번호를 찾아 사건을 식별한다.
// 미국·유럽 서류처럼 전부 스캔 이미지라 텍스트가 없으면, 번호가 적혀 있을 만한 서류 몇 개의
// 첫 페이지를 OCR 한다 (결과는 output/ocr-cache.json 에 저장해 다음 실행 때 재사용).
// 그래도 식별이 안 되는 zip은 _미식별 폴더로 보내고, input/mapping.csv 에 적어주면 반영된다.
const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const iconv = require("iconv-lite");
const ExcelJS = require("exceljs");

const ROOT = path.join(__dirname, "..");
const PDFJS_DIR = path.dirname(require.resolve("pdfjs-dist/package.json"));
const TESS_LANG_DIR = path.join(path.dirname(require.resolve("@tesseract.js-data/eng/package.json")), "4.0.0_best_int");

// OCR 할 서류 우선순위 (서류명에 포함된 단어, 소문자). 출원번호가 큼직하게 찍히는 관청 통지서 위주.
const OCR_PRIORITY = [
  "filing receipt", "bibliographic data", "non-final rejection", "nonfinal rejection", "final rejection",
  "notice of allowance", "notice of publication", "communication", "notification", "search opinion",
  "notice", "receipt",
];
const OCR_MAX_DOCS = 5;
const UNIDENTIFIED_DIR = "_미식별";

// 사무소 관리번호 형식 (예: ABCD240001KRA). 다르면 config.json 의 refPattern 으로 바꾼다.
const DEFAULT_REF_PATTERN = "\\b[A-Z]{2,5}\\d{5,6}[A-Z]{0,4}\\b";

// "등록"으로 분류할 서류명 키워드 (소문자로 비교)
const DEFAULT_GRANT_KEYWORDS = [
  "grant", "registration", "register", "allowance", "certificate", "letters patent",
  "patent right", "annual fee", "annuity", "renewal", "issue fee", "issue notification",
  "등록결정", "특허결정", "설정등록", "등록료", "연차료", "등록증", "특허증",
];

// 출원일을 정하는 데 쓰는 출원 서류 키워드. 이 서류들이 처음 나온 날짜의 서류를 "출원"으로 본다.
// (국내단계 진입 전에 명의변경 통지 등이 먼저 올 수 있어서 단순히 가장 이른 날짜를 쓰면 안 된다)
const FILING_KEYWORDS = [
  "description", "specification", "claims", "abstract", "national entry", "application form",
  "특허출원서", "명세서", "청구범위", "요약서",
];

// ---------------------------------------------------------------------------
// 설정 / 수동 매핑

function loadConfig() {
  const p = path.join(ROOT, "config.json");
  const cfg = fs.existsSync(p) ? JSON.parse(fs.readFileSync(p, "utf8")) : {};
  return {
    refPattern: new RegExp(cfg.refPattern || DEFAULT_REF_PATTERN, "g"),
    // 한국 서식의 참조번호 = 우리 사무소 관리번호
    refLabelPattern: /【참조번호】\s*([A-Za-z0-9-]{4,20})/g,
    // 일본 서식의 整理番号 = 현지 대리인 관리번호
    agentRefLabelPattern: /(?:【整理番号】|\[Reference Number\])\s*([A-Za-z0-9-]{4,20})/g,
    grantKeywords: (cfg.grantKeywords || DEFAULT_GRANT_KEYWORDS).map((k) => k.toLowerCase()),
  };
}

// mapping.csv: zip파일명,국가,출원번호,관리번호  (첫 줄은 헤더)
function loadMapping(inputDir) {
  const p = path.join(inputDir, "mapping.csv");
  const map = new Map();
  if (!fs.existsSync(p)) return map;
  const text = fs.readFileSync(p, "utf8").replace(/^﻿/, "");
  for (const line of text.split(/\r?\n/).slice(1)) {
    if (!line.trim()) continue;
    const [zipName, country, appNo, ref] = splitCsvLine(line).map((s) => (s || "").trim());
    if (zipName) map.set(zipName, { country: country.toUpperCase(), appNo, ref });
  }
  return map;
}

// 엑셀에서 저장한 CSV는 쉼표가 든 값(예: "17/123,456")을 따옴표로 감싼다.
function splitCsvLine(line) {
  const out = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (quoted && ch === '"' && line[i + 1] === '"') {
      cur += '"';
      i++;
    } else if (ch === '"') {
      quoted = !quoted;
    } else if (ch === "," && !quoted) {
      out.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

// ---------------------------------------------------------------------------
// zip / pdf 읽기

// 한국어 윈도우에서 만든 zip은 파일명이 CP949라서 UTF-8 플래그가 없으면 CP949로 디코딩한다.
function entryName(entry) {
  const raw = entry.rawEntryName;
  if (entry.header.flags & 0x800) return raw.toString("utf8");
  const asUtf8 = raw.toString("utf8");
  if (!asUtf8.includes("�")) return asUtf8;
  return iconv.decode(raw, "cp949");
}

let pdfjsPromise;
function openPdf(pdfjs, buffer) {
  return pdfjs.getDocument({
    data: new Uint8Array(buffer),
    cMapUrl: path.join(PDFJS_DIR, "cmaps") + path.sep,
    cMapPacked: true,
    standardFontDataUrl: path.join(PDFJS_DIR, "standard_fonts") + path.sep,
    // 스캔 PDF의 팩스(CCITT)·JBIG2 이미지 디코더. 없으면 빈 페이지로 렌더링된다
    wasmUrl: path.join(PDFJS_DIR, "wasm") + path.sep,
    disableFontFace: true,
    verbosity: 0,
  });
}

async function pdfText(buffer, maxPages = 2) {
  pdfjsPromise = pdfjsPromise || import("pdfjs-dist/legacy/build/pdf.mjs");
  const pdfjs = await pdfjsPromise;
  let task;
  try {
    task = openPdf(pdfjs, buffer);
    const doc = await task.promise;
    let text = "";
    for (let i = 1; i <= Math.min(maxPages, doc.numPages); i++) {
      const page = await doc.getPage(i);
      const content = await page.getTextContent();
      text += content.items.map((it) => it.str).join(" ") + "\n";
    }
    return { text, pages: doc.numPages };
  } catch (err) {
    console.warn(`\n  [경고] PDF 텍스트 추출 실패: ${err.message}`);
    return { text: "", pages: null };
  } finally {
    if (task) await task.destroy();
  }
}

// 스캔 PDF 첫 페이지를 이미지로 그려 OCR 한다. OCR 엔진은 처음 필요할 때 한 번만 띄운다.
let ocrWorkerPromise;
async function ocrFirstPage(buffer) {
  const { createWorker } = require("tesseract.js");
  const { createCanvas } = require("@napi-rs/canvas");
  ocrWorkerPromise = ocrWorkerPromise || createWorker("eng", 1, { langPath: TESS_LANG_DIR, gzip: true, cacheMethod: "none" });
  pdfjsPromise = pdfjsPromise || import("pdfjs-dist/legacy/build/pdf.mjs");
  const [worker, pdfjs] = await Promise.all([ocrWorkerPromise, pdfjsPromise]);
  const task = openPdf(pdfjs, buffer);
  try {
    const doc = await task.promise;
    const page = await doc.getPage(1);
    const viewport = page.getViewport({ scale: 200 / 72 });
    const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, canvas, viewport }).promise;
    const { data } = await worker.recognize(canvas.toBuffer("image/png"));
    return data.text;
  } catch (err) {
    console.warn(`\n  [경고] OCR 실패: ${err.message}`);
    return "";
  } finally {
    await task.destroy();
  }
}

async function closeOcr() {
  if (ocrWorkerPromise) await (await ocrWorkerPromise).terminate();
}

// ---------------------------------------------------------------------------
// 출원번호 찾기

// 일본 서류의 전각 숫자(２０２４－５１６９８５)를 반각으로 바꾸고 공백을 정리한다
function normalizeText(s) {
  return s.normalize("NFKC").replace(/\s+/g, " ");
}

// 텍스트에서 국가별 출원번호 후보를 뽑는다. 결과: [{country, appNo}]
function findAppNumbers(text) {
  const found = [];
  const push = (country, appNo) => found.push({ country, appNo });
  let m;

  // 중국: 202380012345.6 (12자리 + 체크디지트)
  const cn = /(?:申请号|专利号|Application No\.?)[^0-9]{0,15}((?:19|20)\d{10}\s?\.\s?[\dX])/g;
  while ((m = cn.exec(text))) push("CN", m[1].replace(/\s/g, ""));

  // 미국: 17/123,456
  const us = /\b(\d{2})\s?\/\s?(\d{3}),?(\d{3})\b/g;
  while ((m = us.exec(text))) push("US", `${m[1]}/${m[2]},${m[3]}`);

  // 유럽: 23123456.7 (중국 번호 안의 숫자와 겹치지 않게 앞뒤 숫자 금지)
  const ep = /(?:Application (?:No|number)\.?|Anmeldenummer|N° de demande)[^0-9]{0,15}(?<!\d)(\d{2}\s?\d{3}\s?\d{3}\s?\.\s?\d)(?!\d)/gi;
  while ((m = ep.exec(text))) push("EP", m[1].replace(/\s/g, ""));
  const epPrefixed = /\bEP\s?(\d{2}\s?\d{3}\s?\d{3}\s?\.\s?\d)(?!\d)/g;
  while ((m = epPrefixed.exec(text))) push("EP", m[1].replace(/\s/g, ""));

  // 일본: 特願2023-123456
  const jp = /(?:特願|Japanese Patent Application No\.?)\s?(\d{4})\s?[-‐―]\s?(\d{1,6})/g;
  while ((m = jp.exec(text))) push("JP", `${m[1]}-${m[2].padStart(6, "0")}`);

  // 한국: 10-2023-0123456 (해외 사건 서류에도 우선권 번호로 자주 나오므로 채택 우선순위는 가장 낮다)
  const kr = /(?<![\d-])((?:10|20)-\d{4}-\d{7})(?![\d-])/g;
  while ((m = kr.exec(text))) push("KR", m[1]);

  return found;
}

function findPct(text) {
  const m = text.match(/PCT\s?\/\s?([A-Z]{2})\s?(\d{4})\s?\/?\s?(\d{6})(?!\d)/);
  return m ? `PCT/${m[1]}${m[2]}/${m[3]}` : "";
}

function mostCommon(values) {
  const counts = new Map();
  for (const v of values) counts.set(v, (counts.get(v) || 0) + 1);
  let best = "";
  let bestCount = 0;
  for (const [v, c] of counts) {
    if (c > bestCount) {
      best = v;
      bestCount = c;
    }
  }
  return best;
}

// 같은 서류가 원문과 영문 번역본으로 같이 오는 경우(일본 등)를 구분하기 위한 언어 추정
function detectLanguage(text) {
  const count = (re) => (text.match(re) || []).length;
  const kana = count(/[\u3040-\u30ff]/g);
  const hangul = count(/[\uac00-\ud7a3]/g);
  const han = count(/[\u4e00-\u9fff]/g);
  const latin = count(/[A-Za-z]/g);
  if (kana + hangul + han + latin < 20) return "";
  if (latin > (kana + hangul + han) * 2) return "EN";
  if (hangul >= kana && hangul >= han * 0.3) return "KO";
  if (kana > 0) return "JA";
  return "ZH";
}

// ---------------------------------------------------------------------------
// 서류명 / 분류

// "20240315_Request for substantive examination ORIGINAL.pdf" -> { date, title, version }
function parseDocFileName(name) {
  const base = path.basename(name).replace(/\.pdf$/i, "");
  const m = base.match(/^(\d{4})(\d{2})(\d{2})_(.*)$/);
  if (!m) return { date: "", title: base.trim(), version: "" };
  // 같은 날짜에 같은 서류명이 여러 개면 KIPRIS가 "서류명  (2).pdf"처럼 번호를 붙인다
  let title = m[4].replace(/\s+/g, " ").replace(/\s\(\d+\)$/, "").trim();
  let version = "";
  const v = title.match(/\s(ORIGINAL|TRANSLATION|TRANSLATED|MACHINE TRANSLATION)$/i);
  if (v) {
    version = v[1].toUpperCase();
    title = title.slice(0, v.index).trim();
  }
  return { date: `${m[1]}-${m[2]}-${m[3]}`, title, version };
}

function filingDate(docs) {
  const isFiling = (d) => FILING_KEYWORDS.some((k) => d.title.toLowerCase().includes(k));
  const dates = docs.filter(isFiling).map((d) => d.date).filter(Boolean);
  const all = dates.length ? dates : docs.map((d) => d.date).filter(Boolean);
  return all.length ? all.reduce((a, b) => (a < b ? a : b)) : "";
}

// 출원일 서류면 "출원", 등록 키워드가 있으면 "등록", 나머지는 "중간".
function classify(doc, filedOn, grantKeywords) {
  if (doc.date && doc.date === filedOn) return "출원";
  const t = doc.title.toLowerCase();
  if (grantKeywords.some((k) => t.includes(k))) return "등록";
  return "중간";
}

function safeFileName(s) {
  return s.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim().slice(0, 150);
}

// ---------------------------------------------------------------------------
// zip 1개 처리

// 텍스트에서 해외 번호를 못 찾았을 때만 OCR 한다. 두 서류에서 같은 번호가 나오면 멈춘다.
async function ocrNumbers(docs) {
  const rank = (d) => {
    const t = d.title.toLowerCase();
    const i = OCR_PRIORITY.findIndex((k) => t.includes(k));
    return i === -1 ? OCR_PRIORITY.length : i;
  };
  const candidates = docs
    .filter((d) => /\.pdf$/i.test(d.originalName))
    .sort((a, b) => rank(a) - rank(b) || b.date.localeCompare(a.date))
    .slice(0, OCR_MAX_DOCS);

  const numbers = [];
  let pct = "";
  for (const doc of candidates) {
    const text = normalizeText(await ocrFirstPage(doc.data));
    numbers.push(...findAppNumbers(text).filter((n) => n.country !== "KR"));
    pct = pct || findPct(text);
    const counts = new Map();
    for (const n of numbers) counts.set(n.appNo, (counts.get(n.appNo) || 0) + 1);
    if ([...counts.values()].some((c) => c >= 2)) break;
  }
  return { numbers, pct };
}

async function readZip(zipPath, cfg, ocrCache) {
  const zip = new AdmZip(zipPath);
  const docs = [];
  const numbers = [];
  const refs = [];
  const labeledRefs = [];
  const agentRefs = [];
  let pct = "";

  for (const entry of zip.getEntries()) {
    if (entry.isDirectory) continue;
    const name = entryName(entry);
    const data = entry.getData();
    const doc = { ...parseDocFileName(name), originalName: path.basename(name), data, pages: null, lang: "" };

    if (/\.pdf$/i.test(name)) {
      const extracted = await pdfText(data);
      const text = normalizeText(extracted.text);
      doc.pages = extracted.pages;
      doc.lang = detectLanguage(text);
      numbers.push(...findAppNumbers(text));
      pct = pct || findPct(text);
      for (const m of text.matchAll(cfg.refLabelPattern)) labeledRefs.push(m[1]);
      for (const m of text.matchAll(cfg.agentRefLabelPattern)) agentRefs.push(m[1]);
      refs.push(...(text.match(cfg.refPattern) || []));
    }
    docs.push(doc);
  }

  // 해외 번호가 텍스트에 없으면 (스캔 PDF) OCR 로 보완한다
  let foreign = numbers.filter((n) => n.country !== "KR");
  let ocr = false;
  // KIPRIS 한국 사건은 서류명이 한글이다. 그 경우 한국 번호로 충분하므로 OCR 하지 않는다.
  const koreanCase = docs.some((d) => /[\uac00-\ud7a3]/.test(d.title)) && numbers.length > 0;
  if (foreign.length === 0 && docs.length > 0 && !koreanCase) {
    const cacheKey = `${path.basename(zipPath)}|${fs.statSync(zipPath).size}`;
    if (!ocrCache[cacheKey]) {
      process.stdout.write("OCR 중... ");
      ocrCache[cacheKey] = await ocrNumbers(docs);
    }
    foreign = ocrCache[cacheKey].numbers;
    pct = pct || ocrCache[cacheKey].pct;
    ocr = foreign.length > 0;
    numbers.push(...foreign);
  }

  // 가장 많이 나온 번호를 채택 (서류 여러 개에서 같은 번호가 반복되므로).
  // 한국 번호는 해외 사건 서류에 우선권 번호로도 나오므로 다른 국가 번호가 하나도 없을 때만 쓴다.
  const best = mostCommon((foreign.length ? foreign : numbers).map((n) => `${n.country}|${n.appNo}`));
  const [country, appNo] = best ? best.split("|") : ["", ""];
  const agentRef = mostCommon(agentRefs);
  const ref = mostCommon(labeledRefs) || mostCommon(refs.filter((r) => r !== agentRef));
  return { docs, country, appNo, pct, ref, agentRef, ocr };
}

// ---------------------------------------------------------------------------
// 메인

// input/ 아래 하위 폴더(국가별, 고객사별 등)에 있는 zip까지 모두 찾는다. 결과는 input/ 기준 상대경로.
function listZips(dir, rel = "") {
  const out = [];
  for (const e of fs.readdirSync(path.join(dir, rel), { withFileTypes: true })) {
    const r = rel ? path.join(rel, e.name) : e.name;
    if (e.isDirectory()) out.push(...listZips(dir, r));
    else if (/\.zip$/i.test(e.name)) out.push(r);
  }
  return out;
}

async function main() {
  const inputDir = path.resolve(process.argv[2] || path.join(ROOT, "input"));
  const outputDir = path.resolve(process.argv[3] || path.join(ROOT, "output"));
  if (!fs.existsSync(inputDir)) {
    console.error(`입력 폴더가 없습니다: ${inputDir}`);
    process.exit(1);
  }
  const cfg = loadConfig();
  const mapping = loadMapping(inputDir);
  const zipFiles = listZips(inputDir).sort();
  if (zipFiles.length === 0) {
    console.error(`zip 파일이 없습니다: ${inputDir}`);
    process.exit(1);
  }

  // 매번 zip 전체로부터 다시 만든다 (결과물을 직접 수정하지 말 것)
  const casesDir = path.join(outputDir, "cases");
  fs.rmSync(casesDir, { recursive: true, force: true });
  fs.mkdirSync(casesDir, { recursive: true });

  const ocrCachePath = path.join(outputDir, "ocr-cache.json");
  const ocrCache = fs.existsSync(ocrCachePath) ? JSON.parse(fs.readFileSync(ocrCachePath, "utf8")) : {};

  const cases = new Map(); // key -> { country, appNo, pct, ref, zips:Set, docs:[] }

  for (const zipFile of zipFiles) {
    process.stdout.write(`- ${zipFile} ... `);
    let info;
    try {
      info = await readZip(path.join(inputDir, zipFile), cfg, ocrCache);
    } catch (err) {
      console.log(`읽기 실패 (${err.message})`);
      continue;
    }

    const manual = mapping.get(path.basename(zipFile));
    if (manual) {
      info.country = manual.country || info.country;
      info.appNo = manual.appNo || info.appNo;
      info.ref = manual.ref || info.ref;
    }

    const identified = Boolean(info.country && info.appNo);
    const key = identified
      ? `${info.country}_${info.appNo.replace(/[\/,]/g, "")}`
      : `${UNIDENTIFIED_DIR}/${path.basename(zipFile).replace(/\.zip$/i, "")}`;

    if (!cases.has(key)) {
      cases.set(key, {
        key, identified, country: info.country, appNo: info.appNo,
        pct: info.pct, ref: info.ref, agentRef: info.agentRef, ocr: info.ocr && !manual, zips: new Set(), docs: [],
      });
    }
    const c = cases.get(key);
    c.pct = c.pct || info.pct;
    c.ref = c.ref || info.ref;
    c.agentRef = c.agentRef || info.agentRef;
    c.zips.add(zipFile);

    // 같은 사건을 나중에 다시 받은 경우 같은 서류는 한 번만 넣는다
    for (const doc of info.docs) {
      const dup = c.docs.find((d) => d.originalName === doc.originalName && d.data.length === doc.data.length);
      if (!dup) c.docs.push({ ...doc, zip: zipFile });
    }
    const how = manual ? " (mapping.csv)" : info.ocr ? " (OCR)" : "";
    console.log(identified ? `${info.country} ${info.appNo}${how}` : "출원번호 못 찾음 -> _미식별");
  }
  await closeOcr();
  fs.writeFileSync(ocrCachePath, JSON.stringify(ocrCache, null, 2));

  // 파일 쓰기 + 분류
  const docRows = [];
  const caseRows = [];
  for (const c of [...cases.values()].sort((a, b) => a.key.localeCompare(b.key))) {
    const dir = path.join(casesDir, c.key);
    fs.mkdirSync(dir, { recursive: true });
    c.docs.sort((a, b) => (a.date + a.title).localeCompare(b.date + b.title));
    const dates = c.docs.map((d) => d.date).filter(Boolean);
    const firstDate = dates.length ? dates.reduce((a, b) => (a < b ? a : b)) : "";
    const filedOn = filingDate(c.docs);
    // 같은 날짜·같은 서류명이 여러 개(원문 + 번역본 등)면 파일명에 언어를 붙여 구분한다
    const sameName = new Map();
    for (const d of c.docs) sameName.set(d.date + d.title, (sameName.get(d.date + d.title) || 0) + 1);

    const used = new Set();
    for (const doc of c.docs) {
      const ext = path.extname(doc.originalName) || ".pdf";
      const lang = sameName.get(doc.date + doc.title) > 1 ? doc.lang : "";
      const stem = safeFileName([doc.date, doc.title, doc.version && doc.version !== "ORIGINAL" ? doc.version : "", lang]
        .filter(Boolean).join("_"));
      let fileName = `${stem}${ext}`;
      for (let i = 2; used.has(fileName.toLowerCase()); i++) fileName = `${stem} (${i})${ext}`;
      used.add(fileName.toLowerCase());
      fs.writeFileSync(path.join(dir, fileName), doc.data);

      docRows.push({
        ref: c.ref, country: c.country, appNo: c.appNo, date: doc.date,
        category: c.identified ? classify(doc, filedOn, cfg.grantKeywords) : "",
        title: doc.title, lang: doc.lang, version: doc.version, pages: doc.pages,
        file: path.posix.join("cases", c.key, fileName), zip: doc.zip,
      });
    }

    const last = c.docs[c.docs.length - 1];
    caseRows.push({
      ref: c.ref, agentRef: c.agentRef, country: c.country, appNo: c.appNo, pct: c.pct,
      docCount: c.docs.length, firstDate, lastDate: last ? last.date : "", lastDoc: last ? last.title : "",
      folder: path.posix.join("cases", c.key), zips: [...c.zips].join(", "),
      status: !c.identified ? "출원번호 확인 필요 (mapping.csv에 입력 후 재실행)" : c.ocr ? "OCR로 인식 - 번호 확인 권장" : "",
    });
  }

  const xlsxPath = path.join(outputDir, "filewrapper_db.xlsx");
  await writeWorkbook(xlsxPath, caseRows, docRows);

  const unidentified = caseRows.filter((r) => !r.appNo).length;
  const byOcr = caseRows.filter((r) => r.appNo && r.status).length;
  console.log("");
  console.log(`사건 ${caseRows.length}건 / 서류 ${docRows.length}개 정리 완료`);
  console.log(`엑셀: ${xlsxPath}`);
  if (byOcr) console.log(`[참고] OCR로 출원번호를 읽은 사건 ${byOcr}건 -> 엑셀 "사건목록"에서 번호 한 번 확인 권장`);
  if (unidentified) console.log(`[확인 필요] 출원번호를 못 찾은 zip ${unidentified}개 -> 엑셀 "사건목록"의 상태 열 참고`);
}

async function writeWorkbook(xlsxPath, caseRows, docRows) {
  const wb = new ExcelJS.Workbook();

  const cs = wb.addWorksheet("사건목록", { views: [{ state: "frozen", ySplit: 1 }] });
  cs.columns = [
    { header: "관리번호(추정)", key: "ref", width: 18 },
    { header: "국가", key: "country", width: 6 },
    { header: "출원번호", key: "appNo", width: 18 },
    { header: "PCT번호", key: "pct", width: 20 },
    { header: "현지대리인번호", key: "agentRef", width: 16 },
    { header: "서류수", key: "docCount", width: 8 },
    { header: "최초서류일", key: "firstDate", width: 12 },
    { header: "최근서류일", key: "lastDate", width: 12 },
    { header: "최근서류", key: "lastDoc", width: 50 },
    { header: "폴더", key: "folder", width: 30 },
    { header: "원본zip", key: "zips", width: 30 },
    { header: "상태", key: "status", width: 40 },
  ];
  for (const r of caseRows) {
    const row = cs.addRow(r);
    row.getCell("folder").value = { text: r.folder, hyperlink: r.folder };
  }

  const ds = wb.addWorksheet("서류목록", { views: [{ state: "frozen", ySplit: 1 }] });
  ds.columns = [
    { header: "관리번호(추정)", key: "ref", width: 18 },
    { header: "국가", key: "country", width: 6 },
    { header: "출원번호", key: "appNo", width: 18 },
    { header: "일자", key: "date", width: 12 },
    { header: "서류구분", key: "category", width: 9 },
    { header: "서류명", key: "title", width: 60 },
    { header: "언어", key: "lang", width: 6 },
    { header: "버전", key: "version", width: 12 },
    { header: "페이지", key: "pages", width: 8 },
    { header: "파일", key: "file", width: 50 },
    { header: "원본zip", key: "zip", width: 30 },
  ];
  for (const r of docRows) {
    const row = ds.addRow(r);
    row.getCell("file").value = { text: r.file, hyperlink: r.file };
  }

  for (const ws of [cs, ds]) {
    ws.getRow(1).font = { bold: true };
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columns.length } };
  }
  await wb.xlsx.writeFile(xlsxPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
