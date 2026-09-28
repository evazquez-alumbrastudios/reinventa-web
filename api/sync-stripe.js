/**
 * /api/sync-stripe — Sincronización única de pagos históricos de Stripe → Google Sheets
 * Visita esta URL una vez en el navegador para importar los pagos anteriores al webhook.
 * Después de usarla puedes avisarme y la eliminamos.
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx9fcL9RYunwwtYPlCCMu4DnnlGxuQn7DR-y8h-TwSkj0Zk8uzb8YpsBLCpJYT9lwJBIQ/exec';
const MONTOS_TALLER = [130000, 150000, 170000, 260000, 300000, 340000, 390000, 450000, 510000]; // incluye x2 y x3 lugares

module.exports = async function handler(req, res) {
  const key = process.env.STRIPE_SECRET_KEY;
  if (!key) return res.status(500).json({ error: 'Sin clave de Stripe' });

  try {
    /* Traer todas las checkout sessions completadas */
    let sessions = [];
    let hasMore = true;
    let startingAfter = null;

    while (hasMore) {
      let url = 'https://api.stripe.com/v1/checkout/sessions?status=complete&limit=100';
      if (startingAfter) url += `&starting_after=${startingAfter}`;

      const resp = await fetch(url, {
        headers: { Authorization: `Bearer ${key}` }
      });
      const data = await resp.json();

      for (const s of data.data) {
        if (MONTOS_TALLER.includes(s.amount_total)) sessions.push(s);
      }

      hasMore = data.has_more;
      if (hasMore) startingAfter = data.data[data.data.length - 1].id;
    }

    /* Enviar cada pago al Apps Script como si fuera un webhook de Stripe */
    const resultados = [];
    for (const session of sessions) {
      const payload = {
        type: 'checkout.session.completed',
        data: { object: session }
      };

      const r = await fetch(SHEETS_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });

      resultados.push({
        correo: session.customer_details?.email || '—',
        nombre: session.customer_details?.name || '—',
        monto: `$${(session.amount_total / 100).toFixed(2)}`,
        status: r.ok ? '✓ enviado' : '✗ error'
      });
    }

    return res.status(200).json({
      total: sessions.length,
      pagos: resultados
    });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
