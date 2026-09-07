/**
 * /api/admin — Panel de administración de REINVENTA
 * GET  ?key=PASSWORD          → valida clave y devuelve datos completos
 * GET  ?key=PASSWORD&ping=1   → solo valida la clave (rápido, sin llamar GAS)
 * Contraseña: variable ADMIN_KEY o por defecto "reinventa2026"
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxY1AuqgR3sonF2MhxsphCVWHQr5pTJg-Qs_xmEHEFnTaK4Q6y_ivFXrhfHUW69or7ymA/exec';
const DEFAULT_KEY = 'reinventa2026';
const GAS_TIMEOUT_MS = 28000; // Vercel hobby limit is 30s; leave 2s margin

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store');

  const { key, ping } = req.query;
  const adminKey = process.env.ADMIN_KEY || DEFAULT_KEY;

  // Auth check is always instant (no GAS needed)
  if (key !== adminKey) return res.status(401).json({ error: 'Acceso no autorizado' });

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
