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

    // Fallback: attempt to read raw stream when body parser doesn't provide req.body.
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
      const r = await fetch(target, { method: 'GET', headers: { Accept: 'application/json' } });
      const text = await r.text();
      res.status(r.status).send(text);
      return;
    }

    if (req.method === 'POST') {
      const body = await parseRequestBody(req);
      if (!body || typeof body !== 'object' || !body.action) {
        res.status(400).json({ error: 'Invalid payload: missing action' });
        return;
      }

      const r = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
        body: JSON.stringify(body)
      });
      const text = await r.text();
      res.status(r.status).send(text);
      return;
    }

    res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
