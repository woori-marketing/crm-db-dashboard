/**
 * 내보낸 고객목록 CSV 를 읽어 날짜 · 유입경로별 건수만 남기고, 대시보드 HTML 을 만든다.
 *
 * 이름과 전화번호는 이 단계에서 버려진다. 이후 단계(store.json, HTML, 배포)에는
 * 건수만 흐르므로 개인정보가 따라가지 않는다.
 *
 * 사용:
 *   node aggregate.mjs --raw <원본폴더> --store <store.json> --template <템플릿.html> --out <결과.html>
 *   node aggregate.mjs ... --delete      집계에 성공한 CSV 를 지운다
 */

import fs from "node:fs";
import path from "node:path";

import { findAgentColumn, makeSplitter } from "./roster.mjs";

// 집계에서 빼는 유입경로. 앞부분이 맞으면 제외한다.
const EXCLUDE = ["협력점해피콜"];       // 신규 인입이 아니라 기존 고객 확인 전화다
const UNSET = "미지정";                 // 유입경로가 비어 있거나 엉뚱한 값이 들어간 건

// 화면은 이 문자열을 그대로 찍는다. UTC 로 적으면 새벽 6시 수집이 전날 21시로 보여서,
// 수집이 안 된 날인지 아닌지를 갱신 시각으로 가릴 수 없게 된다. 한국 시각으로 적는다.
const nowSeoul = () => new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 19);
// 유입경로가 비어 있는 건은 세지 않는다. 어디서 들어왔는지 알 수 없어 비교에 쓸 수 없다.

// 원본 유입경로 값을 그대로 저장한다. 이름을 묶어 보여주는 일은 화면(template.html)이 한다.
// 묶는 규칙을 바꿔도 다시 수집할 필요가 없도록 하기 위해서다.

// ── 인자 ───────────────────────────────────────────────

function parseArgs(argv) {
  const out = { delete: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--delete") out.delete = true;
    else if (a.startsWith("--")) out[a.slice(2)] = argv[++i];
  }
  return out;
}

// ── CSV ────────────────────────────────────────────────

// 따옴표 안의 쉼표와 줄바꿈을 지켜야 하므로 split 으로는 안 된다.
function parseCsv(text) {
  const rows = [];
  let row = [], field = "", quoted = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else field += c;
      continue;
    }
    if (c === '"') quoted = true;
    else if (c === ",") { row.push(field); field = ""; }
    else if (c === "\r") continue;
    else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
    else field += c;
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }

  return rows.filter(r => r.some(v => v.trim()));
}

// ── 컬럼 찾기 ──────────────────────────────────────────

const CODE_RE = /\(\d{3,5}\)/;
// 유입경로 칸에 전화번호가 들어간 건. 그대로 두면 공개 화면에 고객 번호가 실린다.
const PHONE_RE = /^\d{2,4}-\d{3,4}-\d{4}$/;
const REG_DATE_RE = /^\d{4}-\d{2}-\d{2}\s+오[전후]\s+\d{1,2}:\d{2}:\d{2}$/;
const PLAIN_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})/;

const isHeader = (row) => row.some(v => v.trim() === "유입경로" || v.trim() === "등록일");

/**
 * 컬럼 순서는 내보낼 때마다 달라질 수 있고 헤더가 없을 수도 있다.
 * 헤더가 있으면 이름으로 잡고, 없으면 값의 생김새로 고른다.
 *  - 등록일: '2026-08-20 오전 9:34:27' 형식은 이 열에만 나온다
 *  - 유입경로: '구글-키워드(4869)' 처럼 괄호 코드가 붙는 비율이 다른 열보다 압도적이다
 */
function locateColumns(rows) {
  if (isHeader(rows[0])) {
    const h = rows[0].map(v => v.trim());
    return {
      header: h,
      date: h.indexOf("등록일"),
      channel: h.indexOf("유입경로"),
      body: rows.slice(1),
    };
  }

  const body = rows;
  const width = Math.max(...body.map(r => r.length));
  let date = -1, channel = -1, bestDate = 0, bestChannel = 0;

  for (let i = 0; i < width; i++) {
    const vals = body.map(r => (r[i] ?? "").trim()).filter(Boolean);
    if (!vals.length) continue;
    const d = vals.filter(v => REG_DATE_RE.test(v)).length / vals.length;
    const c = vals.filter(v => CODE_RE.test(v)).length / vals.length;
    if (d > bestDate && d > 0.8) { bestDate = d; date = i; }
    if (c > bestChannel && c > 0.4) { bestChannel = c; channel = i; }
  }
  return { header: null, date, channel, body };
}

// ── 집계 ───────────────────────────────────────────────

const fileDate = (name) => {
  const m = name.match(/(\d{4})[-_]?(\d{2})[-_]?(\d{2})/);
  return m ? `${m[1]}-${m[2]}-${m[3]}` : null;
};

function readFile(file) {
  const text = new TextDecoder("euc-kr").decode(fs.readFileSync(file));
  const rows = parseCsv(text);
  if (!rows.length) return { counts: [], skipped: 0, reason: "빈 파일" };

  const { date, channel, body, header } = locateColumns(rows);
  if (channel < 0) return { counts: [], skipped: body.length, reason: "유입경로 열을 찾지 못함" };

  // 담당자 열로 시흥·천안을 가른다. 자리는 내보낼 때마다 달라져서 명단과 맞춰 찾는다.
  const agent = findAgentColumn(body, header);
  if (agent < 0) return { counts: [], skipped: body.length, reason: "담당자 열을 찾지 못함" };
  const branchOf = makeSplitter();

  const fallback = fileDate(path.basename(file));
  if (date < 0 && !fallback) {
    return { counts: [], skipped: body.length, reason: "등록일 열도 파일명 날짜도 없음" };
  }

  const tally = new Map();
  let skipped = 0;

  for (const r of body) {
    const raw = (r[channel] ?? "").trim();
    if (EXCLUDE.some(x => raw.startsWith(x))) continue;

    let day = fallback;
    if (date >= 0) {
      const m = (r[date] ?? "").match(PLAIN_DATE_RE);
      if (m) day = `${m[1]}-${m[2]}-${m[3]}`;
    }
    if (!day) { skipped++; continue; }

    const key = day + "\t" + ((!raw || PHONE_RE.test(raw)) ? UNSET : raw) + "\t" + branchOf(r[agent]);
    tally.set(key, (tally.get(key) || 0) + 1);
  }

  const counts = [...tally].map(([k, count]) => {
    const [date, channel, branch] = k.split("\t");
    return { date, channel, branch, count };
  });
  return { counts, skipped, reason: null };
}

// ── 누적 저장 ──────────────────────────────────────────

// 원본 CSV 는 지워지므로, 지금까지 센 결과는 여기에 남는다. 건수뿐이라 개인정보가 없다.
function loadStore(file) {
  try {
    const j = JSON.parse(fs.readFileSync(file, "utf8"));
    if (Array.isArray(j.rows)) return j;
  } catch {}
  return { generated: null, excluded: [], rows: [] };
}

function mergeStore(store, counts) {
  const key = (r) => r.date + "\t" + r.channel + "\t" + (r.branch ?? "");
  // 같은 날을 다시 세면 그 날 줄은 통째로 갈아끼운다. 지점이 붙기 전 줄과 섞여
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
for (const need of ["raw", "store", "template", "out"]) {
  if (!args[need]) { console.error(`--${need} 를 지정하세요.`); process.exit(1); }
}

const files = fs.readdirSync(args.raw)
  .filter(f => /\.csv$/i.test(f))
  .map(f => path.join(args.raw, f))
  .sort();

console.log(`CSV ${files.length} 개`);

const store = loadStore(args.store);
const done = [];
let totalAdded = 0, totalUpdated = 0;

for (const file of files) {
  let res;
  try { res = readFile(file); }
  catch (e) { console.error(`  실패  ${path.basename(file)} — ${e.message}`); continue; }

  if (res.reason) { console.error(`  건너뜀 ${path.basename(file)} — ${res.reason}`); continue; }

  const { added, updated } = mergeStore(store, res.counts);
  totalAdded += added; totalUpdated += updated;
  const n = res.counts.reduce((s, c) => s + c.count, 0);
  console.log(`  읽음  ${path.basename(file)} — ${n}건${res.skipped ? ` (날짜 없어 제외 ${res.skipped})` : ""}`);
  done.push(file);
}

if (!store.rows.length) {
  console.error("집계된 데이터가 없습니다. 결과를 만들지 않습니다.");
  process.exit(1);
}

store.generated = nowSeoul();
store.excluded = EXCLUDE;

// HTML 은 템플릿의 표시 구간만 갈아끼워 만든다.
const tpl = fs.readFileSync(args.template, "utf8");
const marker = /\/\*__DATA__\*\/[\s\S]*?\/\*__END__\*\//;
if (!marker.test(tpl)) {
  console.error("템플릿에서 데이터 자리를 찾지 못했습니다.");
  process.exit(1);
}
const html = tpl.replace(marker, () => `/*__DATA__*/${JSON.stringify(store)}/*__END__*/`);

fs.mkdirSync(path.dirname(path.resolve(args.out)), { recursive: true });
// 누적본을 먼저 쓰고 화면을 나중에 쓴다. 순서가 반대면 화면이 늘 '오래된 것' 으로 보여서
// 받을 것이 없는 날에도 두 시간마다 같은 내용을 다시 만들어 올리게 된다.
fs.writeFileSync(args.store, JSON.stringify(store));
fs.writeFileSync(args.out, html);

const days = [...new Set(store.rows.map(r => r.date))].sort();
console.log(`\n누적 ${store.rows.length}줄 · ${days.length}일 (${days[0]} ~ ${days[days.length - 1]})`);
console.log(`이번 실행: 신규 ${totalAdded}줄 · 다시 세며 지운 줄 ${totalUpdated}`);
console.log(`결과: ${args.out}`);

// 결과 파일을 확실히 쓴 뒤에만 원본을 지운다. 순서가 바뀌면 집계 실패한 날의 데이터를 잃는다.
if (args.delete) {
  for (const f of done) fs.unlinkSync(f);
  console.log(`원본 CSV ${done.length}개 삭제`);
}
