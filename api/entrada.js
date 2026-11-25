/**
 * /api/entrada — Verifica y registra asistencia por ID único
 * GET ?id=X&preview=1 → devuelve nombre y fase SIN marcar asistencia
 * POST {id, pin}      → valida PIN del staff, luego marca asistencia
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx9fcL9RYunwwtYPlCCMu4DnnlGxuQn7DR-y8h-TwSkj0Zk8uzb8YpsBLCpJYT9lwJBIQ/exec';
const { checkSecret } = require('./_auth');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  // GET preview mode
  if (req.method === 'GET') {
    const { id } = req.query;
    if (!id) return res.status(400).json({ error: 'ID requerido' });
    try {
      const url = `${SHEETS_ENDPOINT}?action=hub&id=${encodeURIComponent(id)}`;
      const r = await fetch(url);
      const data = await r.json();
      if (data.error) return res.status(200).json({ error: data.error });
      return res.status(200).json({ nombre: data.nombre, fase: data.fase, preview: true });
    } catch (err) {
      return res.status(500).json({ error: 'Error de conexión con el sheet' });
    }
  }

  // POST confirm mode
  if (req.method === 'POST') {
    const { id, pin } = req.body || {};
    if (!id) return res.status(400).json({ error: 'ID requerido' });
    const auth = checkSecret(process.env.STAFF_PIN, (pin || '').toString().trim());
    if (auth === 'unset') return res.status(503).json({ error: 'Check-in deshabilitado' });
    if (auth !== 'ok') {
      return res.status(200).json({ error: 'pin_incorrecto' });
    }
    try {
      const url = `${SHEETS_ENDPOINT}?action=entrada&id=${encodeURIComponent(id)}`;
      const r = await fetch(url);
      const data = await r.json();
      return res.status(200).json(data);
    } catch (err) {
      return res.status(500).json({ error: 'Error de conexión con el sheet' });
    }
  }

  return res.status(405).json({ error: 'Método no permitido' });
};
