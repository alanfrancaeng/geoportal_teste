// api/sync.js - Vercel Serverless Function
// Rota de sincronização para responder quando chamada na Vercel

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  return res.status(200).json({
    status: 'ok',
    environment: 'vercel_cloud',
    timestamp: new Date().toISOString(),
    message: 'GEOportal & BI-Obras online na Vercel'
  });
};
