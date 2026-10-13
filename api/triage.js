// POST /api/triage  { items: [{n, from, subject, importance, preview}] }
// Optional AI triage via the Anthropic API. If ANTHROPIC_API_KEY is not set,
// returns { enabled:false } and the page falls back to keyword rules.

const PROMPT = "You triage a pharma contract-manufacturer Customer Services inbox. Classify each email's priority:\n" +
  "- critical: needs same-day action or risk to patient safety, product quality, recalls, adverse events, formal complaints, regulatory bodies, line stoppages, angry key customers, urgent deadlines today.\n" +
  "- high: time-sensitive customer/QA request, approaching deadline, approval blocking work, delivery/stock problem.\n" +
  "- routine: normal business correspondence, FYI threads, ongoing discussions with no urgency.\n" +
  "- low: automated notifications, newsletters, planner/task bots, pure CC noise.\n" +
  "Respond with ONLY a JSON array like [{\"n\":0,\"p\":\"critical\",\"r\":\"one short reason\"}] covering every item. No other text.\n\nEMAILS:\n";

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return res.status(200).json({ enabled: false, results: [] });
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST only' });

  const items = Array.isArray(req.body?.items) ? req.body.items.slice(0, 25) : [];
  if (!items.length) return res.status(200).json({ enabled: true, results: [] });
  const clean = items.map(i => ({
    n: Number(i.n), from: String(i.from || '').slice(0, 200), subject: String(i.subject || '').slice(0, 300),
    importance: String(i.importance || ''), preview: String(i.preview || '').slice(0, 300)
  }));

  try {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      body: JSON.stringify({
        model: process.env.ANTHROPIC_MODEL || 'claude-haiku-4-5',
        max_tokens: 2000,
        messages: [{ role: 'user', content: PROMPT + JSON.stringify(clean) }]
      })
    });
    const j = await r.json();
    if (!r.ok) return res.status(502).json({ enabled: true, error: j.error?.message || String(r.status), results: [] });
    const text = (j.content || []).map(c => c.text || '').join('');
    const match = text.match(/\[[\s\S]*\]/);
    const arr = match ? JSON.parse(match[0]) : [];
    const results = arr.filter(a => ['critical', 'high', 'routine', 'low'].includes(a.p))
      .map(a => ({ n: Number(a.n), p: a.p, r: String(a.r || '').slice(0, 200) }));
    return res.status(200).json({ enabled: true, results });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ enabled: true, error: e.message, results: [] });
  }
}
