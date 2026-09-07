/**
 * /api/admin-action — Acciones del panel admin
 * POST { key, action, sub, ...params }
 * sub: checkin | update_correo | registrar_acompanante | enviar_bienvenida
 *    | reporte_completo | set_survey_status | analytics
 *
 * sub=analytics: llama Vercel Analytics REST API (no va a GAS).
 *   Requiere env var VERCEL_TOKEN (Personal Access Token de vercel.com/account/tokens).
 *   VERCEL_PROJECT_ID se inyecta automáticamente por Vercel en runtime.
 *
 * reporte_completo usa timeout extendido (28s) porque lee 7 hojas en GAS.
 * El resto usa timeout corto (12s).
 */
const SHEETS_ENDPOINT   = 'https://script.google.com/macros/s/AKfycbxY1AuqgR3sonF2MhxsphCVWHQr5pTJg-Qs_xmEHEFnTaK4Q6y_ivFXrhfHUW69or7ymA/exec';
const DEFAULT_KEY       = 'reinventa2026';
const TIMEOUT_LONG_MS   = 28000;
const TIMEOUT_SHORT_MS  = 12000;
const VERCEL_API        = 'https://vercel.com/api/web/insights';

// ── sub=analytics: datos de Vercel Web Analytics ──────────────────────────
async function handleAnalytics(res) {
  const token     = process.env.VERCEL_TOKEN;
  const projectId = process.env.VERCEL_PROJECT_ID;
  if (!token)     return res.status(500).json({ ok: false, error: 'VERCEL_TOKEN no configurado en variables de entorno de Vercel' });
  if (!projectId) return res.status(500).json({ ok: false, error: 'VERCEL_PROJECT_ID no disponible' });

  const to   = Date.now();
  const from = to - 90 * 24 * 60 * 60 * 1000; // últimos 90 días

  const headers = { Authorization: `Bearer ${token}` };
  const sig     = AbortSignal.timeout(8000);

  try {
    const [pvRes, pagesRes] = await Promise.all([
      fetch(`${VERCEL_API}/views?projectId=${projectId}&from=${from}&to=${to}&environment=production`, { headers, signal: sig }),
      fetch(`${VERCEL_API}/pages?projectId=${projectId}&from=${from}&to=${to}&environment=production&limit=20`, { headers, signal: sig })
    ]);

    const pvData    = await pvRes.json();
    const pagesData = await pagesRes.json();

    const totalViews    = pvData?.data?.totalPageViews ?? pvData?.totalPageViews ?? null;
    const totalVisitors = pvData?.data?.uniqueVisitors  ?? pvData?.uniqueVisitors  ?? null;

    const pages    = pagesData?.data?.rows ?? pagesData?.rows ?? [];
    const reservar = pages.find(p => (p.path || p.page || '').includes('reservar'));
    const hub      = pages.find(p => (p.path || p.page || '').includes('hub'));

    return res.status(200).json({
      ok: true,
      totalViews,
      totalVisitors,
      topPages: pages.slice(0, 10).map(p => ({
        path:     p.path || p.page || '?',
        views:    p.pageViews ?? p.views ?? 0,
        visitors: p.uniqueVisitors ?? p.visitors ?? 0
      })),
      reservar: reservar ? { views: reservar.pageViews ?? reservar.views ?? 0, visitors: reservar.uniqueVisitors ?? reservar.visitors ?? 0 } : null,
      hub:      hub      ? { views: hub.pageViews      ?? hub.views      ?? 0, visitors: hub.uniqueVisitors      ?? hub.visitors      ?? 0 } : null,
      from: new Date(from).toISOString().slice(0, 10),
      to:   new Date(to).toISOString().slice(0, 10)
    });
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
}

// ── handler principal ──────────────────────────────────────────────────────
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Cache-Control', 'no-store');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const body     = req.body || {};
  const adminKey = process.env.ADMIN_KEY || DEFAULT_KEY;
  if (body.key !== adminKey) return res.status(401).json({ error: 'No autorizado' });

  const isLong     = body.sub === 'reporte_completo' || body.sub === 'set_survey_status';
  const timeoutMs  = isLong ? TIMEOUT_LONG_MS : TIMEOUT_SHORT_MS;
  const controller = new AbortController();
  const timeoutId  = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const r = await fetch(SHEETS_ENDPOINT, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body:    JSON.stringify({ action: 'admin_action', ...body }),
      signal:  controller.signal
    });
    clearTimeout(timeoutId);
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    clearTimeout(timeoutId);
    const isTimeout = err.name === 'AbortError';
    return res.status(isTimeout ? 503 : 500).json({
      error: isTimeout
        ? 'El servidor tardó demasiado — intenta de nuevo en unos segundos'
        : err.message,
      retryable: isTimeout
    });
  }
};
