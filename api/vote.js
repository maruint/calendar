// 인기상 투표 중계: 브라우저 → /api/vote → Apps Script (Vote.gs)
// - 목록·이미지는 Vercel CDN에 캐시해 Apps Script 호출을 줄인다
// - 인증·투표 요청에는 접속 IP를 붙여 넘긴다 (속도 제한·사후 점검용)
const GAS = 'https://script.google.com/macros/s/AKfycby7pDCQ7fmFuBBhCuZ9Kjrh2dI6wEg5-8SVtdImbGhZKQx4OinnafD2otwZePF6NgVt/exec';

async function readJson(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string') return JSON.parse(req.body || '{}');
  const chunks = [];
  for await (const c of req) chunks.push(c);
  return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
}

module.exports = async (req, res) => {
  try {
    if (req.method === 'GET') {
      const { action, id, w } = req.query || {};
      if (action === 'list') {
        const r = await fetch(GAS + '?action=list', { redirect: 'follow' });
        const j = await r.json();
        res.setHeader('Cache-Control', 'public, s-maxage=60, stale-while-revalidate=300');
        return res.status(200).json(j);
      }
      if (action === 'img' && id) {
        const size = Math.min(Math.max(parseInt(w, 10) || 800, 200), 2000);
        const r = await fetch(GAS + '?action=img&id=' + encodeURIComponent(id) + '&w=' + size, { redirect: 'follow' });
        const j = await r.json();
        if (!j.ok) { res.setHeader('Cache-Control', 'no-store'); return res.status(404).send('not found'); }
        res.setHeader('Content-Type', j.type || 'image/jpeg');
        res.setHeader('Cache-Control', 'public, max-age=86400, s-maxage=31536000, immutable');
        return res.status(200).send(Buffer.from(j.data, 'base64'));
      }
      return res.status(400).json({ ok: false, error: 'bad request' });
    }

    if (req.method === 'POST') {
      const p = await readJson(req);
      if (p.action !== 'code' && p.action !== 'vote') return res.status(400).json({ ok: false, error: 'bad request' });
      const ip = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.headers['x-real-ip'] || '';
      const body = {
        action: p.action, email: p.email, code: p.code, id: p.id,
        device: p.device, preview: !!p.preview, ip
      };
      const r = await fetch(GAS, {
        method: 'POST', redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body)
      });
      const j = await r.json();
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json(j);
    }

    res.status(405).end();
  } catch (err) {
    console.error(err);
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ ok: false, error: '서버와 연결하지 못했습니다. 잠시 후 다시 시도해 주십시오.' });
  }
};
