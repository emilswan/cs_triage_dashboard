// Basic-auth gate for the whole site (page + API).
// Fails CLOSED: if DASHBOARD_PASSWORD is not set, nobody gets in.
import { next } from '@vercel/functions';

export const config = { matcher: '/:path*' };

function safeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

export default function middleware(request) {
  const user = process.env.DASHBOARD_USER || 'cs';
  const pass = process.env.DASHBOARD_PASSWORD;
  if (!pass) {
    return new Response('Dashboard not configured: set DASHBOARD_PASSWORD in Vercel project settings.', {
      status: 503, headers: { 'content-type': 'text/plain' }
    });
  }
  const hdr = request.headers.get('authorization') || '';
  if (hdr.startsWith('Basic ')) {
    try {
      const decoded = atob(hdr.slice(6));
      const i = decoded.indexOf(':');
      if (i > -1 && safeEqual(decoded.slice(0, i), user) && safeEqual(decoded.slice(i + 1), pass)) return next();
    } catch (e) { /* fall through */ }
  }
  return new Response('Authentication required', {
    status: 401,
    headers: { 'WWW-Authenticate': 'Basic realm="CS Triage Dashboard", charset="UTF-8"' }
  });
}
