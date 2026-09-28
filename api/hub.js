/**
 * /api/hub — Devuelve el perfil del asistente (check-in, encuesta, fase)
 * GET ?id=RNV-001
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx9fcL9RYunwwtYPlCCMu4DnnlGxuQn7DR-y8h-TwSkj0Zk8uzb8YpsBLCpJYT9lwJBIQ/exec';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store');

  const { id } = req.query;
  if (!id) return res.status(400).json({ error: 'ID requerido' });

  try {
    const url = `${SHEETS_ENDPOINT}?action=hub&id=${encodeURIComponent(id)}`;
    const r = await fetch(url);
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error de conexión' });
  }
};
