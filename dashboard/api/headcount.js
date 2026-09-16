/**
 * 지점별 상담사 인원을 한 곳에 두고 모두가 같은 값을 보게 하는 자리.
 *
 *   GET  /api/headcount   지금 값
 *   POST /api/headcount   { "시흥": 6, "천안": 12 } 로 바꾼다
 *
 * 값은 버셀 Blob 에 파일 하나로 남는다. 배포를 다시 해도 지워지지 않는다.
 * 저장소는 Private 이라 이 함수만 읽고 쓴다 - 바깥에서 파일 주소로 직접 열 수 없다.
 *
 * 저장소를 아직 연결하지 않았으면 GET 은 기본값을, POST 는 왜 못 쓰는지를 돌려준다.
 * 그래야 화면이 조용히 틀린 숫자를 보여주는 일이 없다.
 */

import { get, put } from "@vercel/blob";

const DEFAULT = { "시흥": 3, "천안": 11 };
const BRANCHES = Object.keys(DEFAULT);
const FILE = "headcount.json";

const hasStore = () => !!process.env.BLOB_READ_WRITE_TOKEN;

/** 1 미만이나 999 초과, 숫자가 아닌 값은 받지 않는다. 잘못 눌러 화면이 망가지지 않게. */
function clean(body) {
  const out = {};
  for (const b of BRANCHES) {
    const n = Math.round(Number(body?.[b]));
    if (!Number.isFinite(n) || n < 1 || n > 999) return null;
    out[b] = n;
  }
  return out;
}

async function read() {
  if (!hasStore()) return null;
  // useCache:false - 방금 다른 PC 가 쓴 값이 아니라 CDN 에 남은 옛 값이 오면 안 된다.
  const res = await get(FILE, { access: "private", useCache: false });
  if (!res || !res.stream) return null;                    // 아직 한 번도 저장한 적 없음
  const saved = clean(await new Response(res.stream).json());
  return saved ? { ...saved, updated: res.blob?.uploadedAt ?? null } : null;
}

export default async function handler(req, res) {
  res.setHeader("cache-control", "no-store");

  try {
    if (req.method === "GET") {
      const saved = await read();
      return res.status(200).json(saved ?? { ...DEFAULT, updated: null, fallback: true });
    }

    if (req.method === "POST") {
      if (!hasStore()) {
        return res.status(503).json({ error: "저장소가 연결되지 않았습니다." });
      }
      const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
      const value = clean(body);
      if (!value) {
        return res.status(400).json({ error: "인원은 1 이상 999 이하의 숫자여야 합니다." });
      }
      await put(FILE, JSON.stringify(value), {
        access: "private",
        contentType: "application/json",
        addRandomSuffix: false,
        allowOverwrite: true,       // 늘 같은 파일 하나를 덮어쓴다
        cacheControlMaxAge: 0,
      });
      return res.status(200).json({ ...value, updated: new Date().toISOString() });
    }

    res.setHeader("allow", "GET, POST");
    return res.status(405).json({ error: "GET 과 POST 만 받습니다." });
  } catch (e) {
    // 저장에 실패했는데 성공한 척하면 다른 PC 에는 옛 값이 남는다. 실패는 실패로 알린다.
    return res.status(500).json({ error: String(e?.message || e) });
  }
}
