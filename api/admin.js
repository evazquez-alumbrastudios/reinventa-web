/**
 * /api/admin — Panel de administración de REINVENTA
 * GET  (header x-admin-key)         → valida la clave y consulta el Apps Script
 * GET  (header x-admin-key) ?ping=1 → solo valida la clave
 * La clave vive solo en la variable ADMIN_KEY. Sin ella, la ruta queda cerrada.
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx9fcL9RYunwwtYPlCCMu4DnnlGxuQn7DR-y8h-TwSkj0Zk8uzb8YpsBLCpJYT9lwJBIQ/exec';
const GAS_TIMEOUT_MS = 28000; // Vercel hobby limit is 30s; leave 2s margin

const { checkSecret } = require('./_auth');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'x-admin-key');
  res.setHeader('Cache-Control', 'no-store');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { ping } = req.query;
  const auth = checkSecret(process.env.ADMIN_KEY, req.headers['x-admin-key']);

  // Auth check is always instant (no GAS needed)
  if (auth === 'unset') return res.status(503).json({ error: 'Panel deshabilitado' });
  if (auth !== 'ok') return res.status(401).json({ error: 'Acceso no autorizado' });

  // Ping mode: just validate the key, don't call GAS
  if (ping === '1') return res.status(200).json({ ok: true });

  // Full data fetch with explicit timeout so Vercel never kills it abruptly
  const controller = new AbortController();
  const timeoutId  = setTimeout(() => controller.abort(), GAS_TIMEOUT_MS);

  try {
    const url = `${SHEETS_ENDPOINT}?action=admin`;
    const r = await fetch(url, { signal: controller.signal });
    clearTimeout(timeoutId);
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    clearTimeout(timeoutId);
    const isTimeout = err.name === 'AbortError';
    return res.status(503).json({
      error: isTimeout
        ? 'El servidor tardó demasiado — intenta de nuevo en unos segundos'
        : 'Error de conexión con el sheet',
      retryable: true
    });
  }
};
