/**
 * 담당자 → 지점. 집계기와 백필기가 같은 기준을 쓰도록 한곳에 둔다.
 *
 * 통합고객목록의 담당자 열에서 지점을 가른다. 열 위치는 내보낼 때마다 달라져서
 * (CSV 는 5번째, xlsx 는 맨 오른쪽) 자리로 찾지 않고 이 명단과 맞춰 찾는다.
 */

export const SIHEUNG = "시흥";
export const CHEONAN = "천안";

const ROSTER = new Map();
const add = (branch, names) => names.forEach(n => ROSTER.set(n, branch));

add(SIHEUNG, ["최성훈", "백지연", "성예훈", "이민철", "박윤철", "김시원",
              "양지애", "김대식", "전하영", "이지은", "최시형", "김현정",
              "천승호"]);
add(CHEONAN, ["홍혜리", "김경태", "조항찬", "서정인", "김수현", "송현정",
              "김유리", "조항준", "박소라", "천안영업팀", "조광연", "노창훈",
              "정희준"]);

/** 명단에 있는 이름인가. 담당자 열을 찾을 때 쓴다. */
export const isAgent = (v) => ROSTER.has((v ?? "").trim());

/**
 * 어느 쪽에도 속하지 않는 담당자(미사용, 관리자, 빈칸)를 3 대 11 로 나눈다.
 * 사람 수 비율대로 돌려 담아서 시흥 + 천안이 언제나 합계와 딱 맞게 한다.
 *
 * 이 경로로 새는 건수가 많아지면 명단이 낡았다는 뜻이다. 2026-09-15 은 정희준 12 건이
 * 명단에 없어 여기로 샜다. 새 상담사가 들어오면 위 명단에 먼저 넣을 것.
 */
const RATIO = [
  SIHEUNG, CHEONAN, CHEONAN, CHEONAN, CHEONAN,
  SIHEUNG, CHEONAN, CHEONAN, CHEONAN, CHEONAN,
  SIHEUNG, CHEONAN, CHEONAN, CHEONAN,
];

export function makeSplitter() {
  let n = 0;
  return (name) => {
    const hit = ROSTER.get((name ?? "").trim());
    if (hit) return hit;
    return RATIO[n++ % RATIO.length];
  };
}

/**
 * 담당자 열의 자리. 헤더가 있으면 이름으로, 없으면 명단과 맞는 비율이 가장 높은 열로 찾는다.
 * rows 는 헤더를 뺀 본문이어야 한다.
 */
export function findAgentColumn(rows, header) {
  if (header) {
    const i = header.findIndex(v => (v ?? "").trim() === "담당자");
    if (i >= 0) return i;
  }
  const width = Math.max(0, ...rows.map(r => r.length));
  let col = -1, best = 0;
  for (let i = 0; i < width; i++) {
    const vals = rows.map(r => (r[i] ?? "").trim()).filter(Boolean);
    if (vals.length < rows.length * 0.5) continue;
    const rate = vals.filter(isAgent).length / vals.length;
    if (rate > best && rate > 0.5) { best = rate; col = i; }
  }
  return col;
}

export const BRANCHES = [SIHEUNG, CHEONAN];
