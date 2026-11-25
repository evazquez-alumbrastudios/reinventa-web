const crypto = require('crypto');

// 'unset' (sin variable: ruta cerrada) · 'ok' · 'denied'
function checkSecret(expected, given) {
  if (!expected) return 'unset';
  if (typeof given !== 'string' || !given) return 'denied';
  const h = (v) => crypto.createHash('sha256').update(v).digest();
  return crypto.timingSafeEqual(h(expected), h(given)) ? 'ok' : 'denied';
}

module.exports = { checkSecret };
