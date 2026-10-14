// GET /api/messages?from=<ISO>&to=<ISO>
// Reads the CS shared mailbox Inbox via Microsoft Graph (app-only / client credentials),
// plus Sent Items to detect replies.
//   assigned  = internal (Central Pharma) people the email was addressed To (excluding the mailbox itself)
//   picked up = a reply exists in the same conversation, sent from the mailbox after the message arrived
// Env: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MAILBOX

let tokenCache = { token: null, exp: 0 };

async function getToken() {
  if (tokenCache.token && Date.now() < tokenCache.exp - 60_000) return tokenCache.token;
  const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = process.env;
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
    const e = new Error('Graph credentials not configured (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).');
    e.status = 503; throw e;
  }
  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID, client_secret: MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
  });
  const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(MS_TENANT_ID)}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body
  });
  const j = await r.json();
  if (!r.ok) { const e = new Error('Token request failed: ' + (j.error_description || j.error || r.status)); e.status = 502; throw e; }
  tokenCache = { token: j.access_token, exp: Date.now() + j.expires_in * 1000 };
  return j.access_token;
}

const INTERNAL_DOMAINS = ['centralpharma.com', 'central-pharma.com'];
const INBOX_SELECT = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,importance,bodyPreview,categories,flag,webLink,conversationId';
const SENT_SELECT = 'conversationId,sentDateTime,from,sender';
const REPLY_SELECT = 'conversationId,receivedDateTime,from';

async function pageAll(startUrl, token, maxPages) {
  const items = [];
  let url = startUrl, page = 0, retries = 0;
  while (url && page < maxPages) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.body-content-type="text"' } });
    if (r.status === 429 && retries < 3) {
      retries++;
      await new Promise(s => setTimeout(s, 1000 * (Number(r.headers.get('retry-after')) || 2)));
      continue;
    }
    const j = await r.json();
    if (!r.ok) { const e = new Error('Graph: ' + (j.error?.message || r.status)); e.status = 502; throw e; }
    retries = 0;
    items.push(...(j.value || []));
    url = j['@odata.nextLink'] || null;
    page++;
  }
  return { items, truncated: !!url };
}

const isInternal = addr => {
  const dom = String(addr || '').toLowerCase().split('@')[1] || '';
  return INTERNAL_DOMAINS.some(d => dom === d || dom.endsWith('.' + d));
};

function firstName(a) {
  return (a.name || String(a.address || '').split('@')[0] || '').trim().split(/\s+/)[0];
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  try {
    const mailbox = String(process.env.MAILBOX || '').toLowerCase();
    if (!mailbox) return res.status(503).json({ error: 'MAILBOX env var not set.' });

    const from = new Date(String(req.query.from || ''));
    const to = new Date(String(req.query.to || ''));
    if (isNaN(from) || isNaN(to) || to <= from) return res.status(400).json({ error: 'Invalid from/to range.' });
    if (to - from > 32 * 864e5) return res.status(400).json({ error: 'Range limited to 31 days.' });

    const token = await getToken();
    const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}`;

    const inboxFilter = `receivedDateTime ge ${from.toISOString()} and receivedDateTime lt ${to.toISOString()}`;
    const inboxUrl = `${base}/mailFolders/inbox/messages?$filter=${encodeURIComponent(inboxFilter)}` +
      `&$orderby=receivedDateTime desc&$select=${INBOX_SELECT}&$top=100`;

    // Replies may be sent after the selected window ends, so only bound the lower end.
    const sentFilter = `sentDateTime ge ${from.toISOString()}`;
    const sentUrl = `${base}/mailFolders/sentitems/messages?$filter=${encodeURIComponent(sentFilter)}` +
      `&$orderby=sentDateTime desc&$select=${SENT_SELECT}&$top=100`;

    // Internal replies can land AFTER the selected window ends (e.g. viewing "Yesterday" when the
    // rep replied this morning), so also scan Inbox arrivals from the window end onwards.
    const afterUrl = `${base}/mailFolders/inbox/messages?$filter=${encodeURIComponent(`receivedDateTime ge ${to.toISOString()}`)}` +
      `&$orderby=receivedDateTime desc&$select=${REPLY_SELECT}&$top=100`;
    const afterP = to < new Date() ? pageAll(afterUrl, token, 5) : Promise.resolve({ items: [] });

    const [inbox, sent, after] = await Promise.all([pageAll(inboxUrl, token, 20), pageAll(sentUrl, token, 10), afterP]);

    // conversationId -> list of replies { at, by }  (by = '' when sent-as the shared mailbox)
    const replies = new Map();
    const addReply = (cid, at, by) => {
      if (!cid || !at) return;
      const list = replies.get(cid) || [];
      list.push({ at, by });
      replies.set(cid, list);
    };
    for (const s of sent.items) {
      const senderAddr = String(s.sender?.emailAddress?.address || '').toLowerCase();
      const by = senderAddr && senderAddr !== mailbox ? firstName(s.sender.emailAddress) : '';
      addReply(s.conversationId, s.sentDateTime, by);
    }
    // Reps usually reply from their OWN mailboxes with the CS mailbox copied in, so the
    // reply arrives in the CS Inbox. Count Inbox messages from internal senders as replies
    // to earlier messages in the same conversation.
    for (const m of [...inbox.items, ...after.items]) {
      const addr = String(m.from?.emailAddress?.address || '').toLowerCase();
      const dom = addr.split('@')[1] || '';
      if (!addr || addr === mailbox) continue;
      if (!INTERNAL_DOMAINS.some(d => dom === d || dom.endsWith('.' + d))) continue;
      addReply(m.conversationId, m.receivedDateTime, firstName(m.from.emailAddress));
    }

    const out = inbox.items.map(m => {
      const received = new Date(m.receivedDateTime);
      // Outbound = one of our people writing to an external party with the CS mailbox copied in.
      // It IS our reply, so it never needs a reply itself. Internal-only mail (all recipients
      // internal) stays actionable, e.g. "please print page 9 x 5".
      const fromAddr = String(m.from?.emailAddress?.address || '').toLowerCase();
      const outbound = fromAddr !== mailbox && isInternal(fromAddr) &&
        [...(m.toRecipients || []), ...(m.ccRecipients || [])]
          .some(r => { const a = String(r.emailAddress?.address || ''); return a && !isInternal(a); });
      const reply = outbound
        ? { at: m.receivedDateTime, by: firstName(m.from.emailAddress) }
        : (replies.get(m.conversationId) || [])
          .filter(x => new Date(x.at) > received)
          .sort((a, b) => new Date(a.at) - new Date(b.at))[0] || null;
      const assignedTo = [...new Set((m.toRecipients || [])
        .map(t => t.emailAddress || {})
        .filter(a => {
          const addr = String(a.address || '').toLowerCase();
          const dom = addr.split('@')[1] || '';
          return addr && addr !== mailbox && INTERNAL_DOMAINS.some(d => dom === d || dom.endsWith('.' + d));
        })
        .map(firstName)
        .filter(Boolean))];
      return {
        id: m.id,
        subject: m.subject || '(no subject)',
        sender: m.from?.emailAddress?.address || 'unknown',
        senderName: m.from?.emailAddress?.name || '',
        received: m.receivedDateTime,
        importance: m.importance || 'normal',
        preview: m.bodyPreview || '',
        categories: (m.categories || []).filter(Boolean),
        flagged: !!(m.flag && m.flag.flagStatus && m.flag.flagStatus !== 'notFlagged'),
        webLink: m.webLink || '#',
        assignedTo,
        outbound,
        replied: !!reply,
        repliedAt: reply ? reply.at : null,
        repliedBy: reply ? reply.by : ''
      };
    });
    return res.status(200).json({ mailbox, count: out.length, truncated: inbox.truncated, messages: out });
  } catch (e) {
    console.error(e);
    return res.status(e.status || 500).json({ error: e.message });
  }
}
