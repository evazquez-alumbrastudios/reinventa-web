/**
 * /api/hub-medidas — Guarda medidas de cuerpo de una asistente en Google Sheets
 * POST { id, hombros, busto, cintura, cadera, forma }
 * No requiere autenticación — el ID es validación suficiente para datos no sensibles.
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx9fcL9RYunwwtYPlCCMu4DnnlGxuQn7DR-y8h-TwSkj0Zk8uzb8YpsBLCpJYT9lwJBIQ/exec';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const body = req.body || {};
  if (!body.id) return res.status(400).json({ error: 'ID requerido' });

  // Validación básica de medidas (valores numéricos en rango humano)
  const campos = ['hombros', 'busto', 'cintura', 'cadera'];
  for (const c of campos) {
    if (body[c] !== undefined && body[c] !== '') {
      const v = parseFloat(body[c]);
      if (isNaN(v) || v < 30 || v > 250) {
        return res.status(400).json({ error: 'Medida inválida: ' + c });
      }
    }
  }

  try {
    const r = await fetch(SHEETS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'guardar_medidas', ...body })
    });
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error al guardar medidas' });
  }
};
