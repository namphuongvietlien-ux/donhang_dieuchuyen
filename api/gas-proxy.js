export default async function handler(req, res) {
  const GAS_URL = 'https://script.google.com/macros/s/AKfycbwhqeAzzNrPTm1cH7KMmmj44btXb2OL835xxaItHByohT11sLDrdgfw7BrVlI5txqXonw/exec';

  async function parseRequestBody(request) {
    let body = request.body;

    if (typeof body === 'string') {
      try { return JSON.parse(body); } catch (_) { return { raw: body }; }
    }
    if (body && typeof body === 'object') {
      return body;
    }

    try {
      const chunks = [];
      for await (const chunk of request) chunks.push(Buffer.from(chunk));
      const raw = Buffer.concat(chunks).toString('utf8').trim();
      if (!raw) return {};
      try { return JSON.parse(raw); } catch (_) { return { raw }; }
    } catch (_) {
      return {};
    }
  }

  async function fetchGas(url, init, timeoutMs) {
    const controller = new AbortController();
    const timer = setTimeout(function() { controller.abort(); }, timeoutMs || 50000);
    try {
      const r = await fetch(url, Object.assign({}, init, { signal: controller.signal, redirect: 'follow' }));
      const text = await r.text();
      return { status: r.status, text: text };
    } finally {
      clearTimeout(timer);
    }
  }

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.status(204).end();
    return;
  }

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  try {
    if (req.method === 'GET') {
      const qs = new URLSearchParams(req.query || {}).toString();
      const target = qs ? `${GAS_URL}?${qs}` : GAS_URL;
      const r = await fetchGas(target, { method: 'GET', headers: { Accept: 'application/json' } }, 50000);
      res.status(r.status).send(r.text);
      return;
    }

    if (req.method === 'POST') {
      const body = await parseRequestBody(req);
      if (!body || typeof body !== 'object' || !body.action) {
        res.status(400).json({ error: 'Invalid payload: missing action' });
        return;
      }

      // text/plain giúp GAS nhận body ổn định hơn qua redirect
      const r = await fetchGas(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8', 'Accept': 'application/json' },
        body: JSON.stringify(body)
      }, 50000);
      const looksHtml = /^\s*</.test(r.text || '') || /<!doctype html>/i.test(r.text || '');
      if (looksHtml) {
        res.status(502).json({
          error: 'GAS returned HTML instead of JSON (redirect/doGet). Try direct client POST.',
          action: body.action
        });
        return;
      }
      res.status(r.status).send(r.text);
      return;
    }

    res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    const msg = err && err.name === 'AbortError' ? 'GAS proxy timeout' : (err.message || String(err));
    res.status(504).json({ error: msg });
  }
}
