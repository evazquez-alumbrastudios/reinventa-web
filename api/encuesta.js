/**
 * /api/encuesta — Guarda respuestas de encuesta post-evento
 * POST { id, calificacion, comentario, expectativas, aplicabilidad,
 *        contenido_valioso, mary, org_previa, org_comunicacion,
 *        org_registro, org_checkin, org_materiales, org_lugar,
 *        impacto, nps }
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

  // Detectar versión antigua del hub (solo enviaba calificacion + comentario).
  // Responder con mensaje legible para que la pantalla de error le diga al asistente que recargue.
  const esVersionVieja = body.expectativas === undefined && body.aplicabilidad === undefined;
  if (esVersionVieja) {
    return res.status(409).json({
      error: '⚠️ Tienes una versión antigua. Recarga la página (toca los tres puntos del navegador → Recargar, o presiona F5) y vuelve a contestar la encuesta completa.'
    });
  }

  try {
    const r = await fetch(SHEETS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action:           'encuesta',
        id:               body.id,
        calificacion:     body.calificacion,
        comentario:       body.comentario       || '',
        expectativas:     body.expectativas     || '',
        aplicabilidad:    body.aplicabilidad    || '',
        contenido_valioso:body.contenido_valioso|| '',
        mary:             body.mary             || '',
        org_previa:       body.org_previa       || '',
        org_comunicacion: body.org_comunicacion || '',
        org_registro:     body.org_registro     || '',
        org_checkin:      body.org_checkin      || '',
        org_materiales:   body.org_materiales   || '',
        org_lugar:        body.org_lugar        || '',
        impacto:          body.impacto          || '',
        nps:              body.nps !== undefined ? body.nps : -1
      })
    });
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error de conexión' });
  }
};
