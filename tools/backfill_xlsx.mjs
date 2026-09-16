/**
 * 하루치씩 시트로 나뉜 xlsx 를 읽어 날짜 · 유입경로별 건수만 store.json 에 넣는다.
 *
 * 과거분 백필용이다. 지금 다시 통합고객목록을 뽑으면 그날이 아니라 오늘 기준으로
 * 갱신된 상태가 나오므로, 그때 뽑아둔 파일을 그대로 세는 이 경로가 유일하게 기준이 맞는다.
 *
 * 이름과 전화번호는 여기서 버려진다. store.json 에는 건수만 남는다.
 *
 * 사용:
 *   node backfill_xlsx.mjs --xlsx <파일.xlsx> --store <store.json> [--year 2026] [--dry]
 *   node backfill_xlsx.mjs --xlsx <파일.xlsx> --audit    원본 값이 어떻게 처리되는지만 본다
 */

import fs from "node:fs";
import zlib from "node:zlib";

import { findAgentColumn, makeSplitter } from "./roster.mjs";

// aggregate.mjs 와 같은 규칙이어야 한다. 한쪽만 바꾸면 과거분과 앞으로분의 기준이 어긋난다.
const EXCLUDE = ["협력점해피콜"];       // 신규 인입이 아니라 기존 고객 확인 전화다
const UNSET = "미지정";                 // 유입경로가 비어 있거나 엉뚱한 값이 들어간 건

// 화면은 이 문자열을 그대로 찍는다. UTC 로 적으면 새벽 6시 수집이 전날 21시로 보여서,
// 수집이 안 된 날인지 아닌지를 갱신 시각으로 가릴 수 없게 된다. 한국 시각으로 적는다.
const nowSeoul = () => new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 19);


// ── 인자 ───────────────────────────────────────────────

function parseArgs(argv) {
  const out = { dry: false, audit: false, year: "2026" };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--dry") out.dry = true;
    else if (a === "--audit") out.audit = true;
    else if (a.startsWith("--")) out[a.slice(2)] = argv[++i];
  }
  return out;
}

// ── zip ────────────────────────────────────────────────

// xlsx 는 zip 이다. 의존성을 늘리지 않으려고 중앙 디렉터리만 직접 읽는다.
function unzip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 66000; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("zip 끝을 찾지 못했습니다.");

  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const files = new Map();

  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error("중앙 디렉터리가 깨졌습니다.");
    const method = buf.readUInt16LE(p + 10);
    const size = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const cmtLen = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf8", p + 46, p + 46 + nameLen);

    // 로컬 헤더의 이름/추가필드 길이는 중앙 디렉터리와 다를 수 있어 다시 읽는다.
    const lnLen = buf.readUInt16LE(local + 26);
    const leLen = buf.readUInt16LE(local + 28);
    const start = local + 30 + lnLen + leLen;
    const raw = buf.subarray(start, start + size);

    files.set(name, method === 0 ? raw : zlib.inflateRawSync(raw));
    p += 46 + nameLen + extraLen + cmtLen;
  }
  return files;
}

// ── xml ────────────────────────────────────────────────

const unescapeXml = (s) => s
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&amp;/g, "&");

const stripTags = (s) => unescapeXml(s.replace(/<[^>]*>/g, ""));

function sharedStrings(xml) {
  if (!xml) return [];
  // <si> 안에 <r> 조각이 여러 개일 수 있어 태그를 걷어내고 이어 붙인다.
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => stripTags(m[1]));
}

const colIndex = (ref) => {
  let n = 0;
  for (const ch of ref) {
    const c = ch.charCodeAt(0);
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
};

/** 빈 셀은 아예 빠져 있으므로 r="D7" 의 열 문자로 자리를 잡아야 한다. */
function sheetRows(xml, strs) {
  const rows = [];
  for (const rm of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
    const row = [];
    for (const cm of rm[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = cm[1], body = cm[2] ?? "";
      const ref = /r="([A-Z]+)/.exec(attrs);
      const type = /t="([^"]+)"/.exec(attrs)?.[1];
      let text = "";
      if (type === "inlineStr") {
        const is = /<is>([\s\S]*?)<\/is>/.exec(body);
        if (is) text = stripTags(is[1]);
      } else {
        const v = /<v>([\s\S]*?)<\/v>/.exec(body);
        if (v) text = type === "s" ? (strs[+v[1]] ?? "") : unescapeXml(v[1]);
      }
      row[ref ? colIndex(ref[1]) : row.length] = text;
    }
    rows.push(row);
  }
  return rows;
}

// ── 날짜 ───────────────────────────────────────────────

const EPOCH = Date.UTC(1899, 11, 30);          // 엑셀 일련번호 0
const SERIAL_RE = /^\d{5}(\.\d+)?$/;

// 일련번호는 시간대 없는 값이라 UTC 로 계산해야 하루가 밀리지 않는다.
const serialToIso = (n) => new Date(EPOCH + Math.floor(n) * 86400000).toISOString().slice(0, 10);

/** 등록일 열. 거의 모든 행이 일련번호인 열 중 가장 왼쪽을 쓴다. */
function findDateColumn(rows, skip) {
  const width = Math.max(0, ...rows.map(r => r.length));
  for (let i = 0; i < width; i++) {
    if (i === skip) continue;
    const vals = rows.map(r => (r[i] ?? "").trim());
    if (vals.filter(v => SERIAL_RE.test(v)).length >= rows.length * 0.95) return i;
  }
  return -1;
}

/**
 * 시트 이름이 '07-01' 이라 연도가 없다. 행 안의 등록일 일련번호에서 연도를 얻고,
 * 월·일이 시트 이름과 맞는지 확인한다. 맞지 않으면 --year 로 떨어진다.
 */
function resolveDate(rows, sheetName, fallbackYear) {
  const md = /^(\d{2})-(\d{2})$/.exec(sheetName.trim());
  if (!md) return { date: null, match: 0 };

  const width = Math.max(0, ...rows.map(r => r.length));
  let best = { date: null, match: 0 };

  for (let i = 0; i < width; i++) {
    const vals = rows.map(r => (r[i] ?? "").trim()).filter(v => SERIAL_RE.test(v));
    if (vals.length < rows.length * 0.8) continue;
    const tally = new Map();
    for (const v of vals) {
      const iso = serialToIso(+v);
      tally.set(iso, (tally.get(iso) || 0) + 1);
    }
    const [iso, hit] = [...tally].sort((a, b) => b[1] - a[1])[0];
    const rate = hit / vals.length;
    if (iso.slice(5) === `${md[1]}-${md[2]}` && rate > best.match) best = { date: iso, match: rate };
  }

  if (best.date) return best;
  return { date: `${fallbackYear}-${md[1]}-${md[2]}`, match: 0 };
}

// ── 유입경로 ───────────────────────────────────────────

const CODE_RE = /\(\d{3,5}\)/;
// 유입경로 칸에 전화번호가 들어간 건. 그대로 두면 공개 화면에 고객 번호가 실린다.
const PHONE_RE = /^\d{2,4}-\d{3,4}-\d{4}$/;

// 유입경로 자리에 전화번호나 날짜가 오는 행이 있다. 그런 행은 한 칸 옆으로 밀려 있다.
const looksLikeChannel = (v) => !!v && !SERIAL_RE.test(v) && !PHONE_RE.test(v);

/** 그 행의 유입경로. 기본 열이 유입경로가 아니면 바로 옆 칸을 본다. */
function channelAt(row, col) {
  const here = (row[col] ?? "").trim();
  if (looksLikeChannel(here)) return here;
  const next = (row[col + 1] ?? "").trim();
  return looksLikeChannel(next) ? next : here;
}

/** 헤더가 없으므로 괄호 코드가 붙는 비율이 가장 높은 열을 유입경로로 본다. */
function findChannelColumn(rows) {
  const width = Math.max(0, ...rows.map(r => r.length));
  let col = -1, best = 0;
  for (let i = 0; i < width; i++) {
    const vals = rows.map(r => (r[i] ?? "").trim()).filter(Boolean);
    if (!vals.length) continue;
    const rate = vals.filter(v => CODE_RE.test(v)).length / vals.length;
    if (rate > best && rate > 0.4) { best = rate; col = i; }
  }
  return { col, rate: best };
}

// ── 누적 저장 ──────────────────────────────────────────

function loadStore(file) {
  try {
    const j = JSON.parse(fs.readFileSync(file, "utf8"));
    if (Array.isArray(j.rows)) return j;
  } catch {}
  return { generated: null, excluded: EXCLUDE, rows: [] };
}

function mergeStore(store, counts) {
  const key = (r) => r.date + "\t" + r.channel + "\t" + (r.branch ?? "");
  // 같은 날을 다시 넣으면 그 날 것은 통째로 갈아끼운다. 지점이 붙기 전 줄이 섞여
  // 합계가 두 배로 보이는 일을 막는다.
  const days = new Set(counts.map(c => c.date));
  const kept = store.rows.filter(r => !days.has(r.date));
  const removed = store.rows.length - kept.length;

  const map = new Map(kept.map(r => [key(r), r]));
  let added = 0;
  for (const c of counts) {
    if (!map.has(key(c))) { map.set(key(c), c); added++; }
    else map.get(key(c)).count += c.count;
  }
  store.rows = [...map.values()].sort((a, b) =>
    a.date === b.date ? a.channel.localeCompare(b.channel, "ko") : a.date.localeCompare(b.date));
  return { added, updated: removed };
}

// ── 실행 ───────────────────────────────────────────────

const args = parseArgs(process.argv.slice(2));
for (const need of args.audit ? ["xlsx"] : ["xlsx", "store"]) {
  if (!args[need]) { console.error(`--${need} 를 지정하세요.`); process.exit(1); }
}

const zip = unzip(fs.readFileSync(args.xlsx));
const read = (name) => zip.has(name) ? zip.get(name).toString("utf8") : null;

const strs = sharedStrings(read("xl/sharedStrings.xml"));
const rels = new Map([...read("xl/_rels/workbook.xml.rels")
  .matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)].map(m => [m[1], m[2].replace(/^\/?(xl\/)?/, "")]));
const sheets = [...read("xl/workbook.xml")
  .matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"/g)].map(m => ({ name: unescapeXml(m[1]), rel: m[2] }));

console.log(`시트 ${sheets.length}개`);

const AUDIT = new Map();
const store = args.audit ? { rows: [] } : loadStore(args.store);
const all = [];
let warned = 0;

for (const sh of sheets) {
  const xml = read("xl/" + rels.get(sh.rel));
  if (!xml) { console.error(`  건너뜀 ${sh.name} — 시트를 찾지 못함`); warned++; continue; }

  const rows = sheetRows(xml, strs).filter(r => r.some(v => (v ?? "").trim()));
  if (!rows.length) { console.error(`  건너뜀 ${sh.name} — 빈 시트`); warned++; continue; }

  const { col, rate } = findChannelColumn(rows);
  if (col < 0) { console.error(`  건너뜀 ${sh.name} — 유입경로 열을 찾지 못함`); warned++; continue; }

  // 시트 하나가 하루면 시트 이름이 날짜다. 여러 날이 한 시트에 있으면 행마다 읽는다.
  const { date, match } = resolveDate(rows, sh.name, args.year);
  const dateCol = date ? -1 : findDateColumn(rows, col);
  if (!date && dateCol < 0) { console.error(`  건너뜀 ${sh.name} — 날짜를 읽지 못함`); warned++; continue; }

  // 담당자 열은 내보낸 형식마다 자리가 달라서 명단과 맞춰 찾는다.
  const agentCol = findAgentColumn(rows, null);
  if (agentCol < 0) { console.error(`  건너뜀 ${sh.name} — 담당자 열을 찾지 못함`); warned++; continue; }
  const branchOf = makeSplitter();

  const tally = new Map();
  const seen = new Map();
  let dropped = 0;
  for (const r of rows) {
    const day = date || serialToIso(+(r[dateCol] ?? "").trim());
    const raw = channelAt(r, col);
    const branch = branchOf(r[agentCol]);
    if (args.audit) {
      const why = !raw ? "빈값"
                : PHONE_RE.test(raw) ? "전화번호"
                : (EXCLUDE.find(x => raw.startsWith(x)) ? "제외:" + EXCLUDE.find(x => raw.startsWith(x)) : "집계");
      const k = why + "\t" + (raw || "(비어 있음)") + "\t" + day;
      AUDIT.set(k, (AUDIT.get(k) || 0) + 1);
    }
    if (EXCLUDE.some(x => raw.startsWith(x))) { dropped++; continue; }
    const ch = (!raw || PHONE_RE.test(raw)) ? UNSET : raw;
    const key = day + "\t" + ch + "\t" + branch;
    tally.set(key, (tally.get(key) || 0) + 1);
    seen.set(day, (seen.get(day) || 0) + 1);
  }

  const counts = [...tally].map(([k, count]) => {
    const [date, channel, branch] = k.split("\t");
    return { date, channel, branch, count };
  });
  all.push(...counts);

  const kept = counts.reduce((s, c) => s + c.count, 0);
  const span = [...seen.keys()].sort();
  const label = span.length === 1 ? span[0] : `${span[0]} ~ ${span[span.length - 1]} (${span.length}일)`;
  const flag = (date && match < 0.8) ? "  ← 등록일 확인 필요" : "";
  if (date && match < 0.8) warned++;
  console.log(`  ${label}  ${String(kept).padStart(4)}건 (원본 ${rows.length} · 제외 ${dropped}) 열${col} ${Math.round(rate * 100)}%${flag}`);
  if (span.length > 1) for (const d of span) console.log(`      ${d}  ${String(seen.get(d)).padStart(4)}행`);
}

if (args.audit) {
  const rows = [...AUDIT].map(([k, n]) => { const [why, raw, date] = k.split("\t"); return { why, raw, date, n }; });
  const total = rows.reduce((s, r) => s + r.n, 0);
  const byWhy = new Map();
  for (const r of rows) byWhy.set(r.why, (byWhy.get(r.why) || 0) + r.n);

  console.log(`\n원본 ${total}건`);
  for (const [why, n] of [...byWhy].sort((a, b) => b[1] - a[1]))
    console.log(`  ${String(n).padStart(5)}  ${(n / total * 100).toFixed(1).padStart(5)}%  ${why}`);

  for (const [why] of [...byWhy].sort((a, b) => b[1] - a[1])) {
    if (why === "집계") continue;
    console.log(`\n[${why}]`);
    const byRaw = new Map();
    for (const r of rows.filter(r => r.why === why)) byRaw.set(r.raw, (byRaw.get(r.raw) || 0) + r.n);
    for (const [raw, n] of [...byRaw].sort((a, b) => b[1] - a[1]))
      console.log(`  ${String(n).padStart(5)}  ${raw}`);
    // 특정 시기에 몰리는지 보이면 입력 방식이 바뀐 것을 알 수 있다.
    const byDate = new Map();
    for (const r of rows.filter(r => r.why === why)) byDate.set(r.date, (byDate.get(r.date) || 0) + r.n);
    const ds = [...byDate].sort();
    if (ds.length > 1) console.log("  날짜별: " + ds.map(([d, n]) => `${d.slice(5)}:${n}`).join(" "));
  }
  process.exit(0);
}

if (!all.length) { console.error("집계된 데이터가 없습니다."); process.exit(1); }

const { added, updated } = mergeStore(store, all);
store.generated = nowSeoul();
store.excluded = EXCLUDE;

const days = [...new Set(store.rows.map(r => r.date))].sort();
console.log(`\n신규 ${added}줄 · 다시 넣으며 지운 줄 ${updated}`);
console.log(`누적 ${store.rows.length}줄 · ${days.length}일 (${days[0]} ~ ${days[days.length - 1]})`);
if (warned) console.log(`확인이 필요한 시트 ${warned}개`);

if (args.dry) { console.log("--dry 라 저장하지 않았습니다."); process.exit(0); }
fs.writeFileSync(args.store, JSON.stringify(store));
console.log(`저장: ${args.store}`);
