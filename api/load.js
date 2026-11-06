// api/load.js - Vercel Serverless Function
// Carrega a versão mais recente dos dados dinâmicos na nuvem

module.exports = async (req, res) => {
  // Configurar CORS e Cache
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const { file } = req.query;
  const allowedFiles = ['plano_acao.json', 'estoque.json', 'dados_obras.json'];

  if (!file || !allowedFiles.includes(file)) {
    return res.status(400).json({ error: 'Arquivo inválido ou não permitido.' });
  }

  // 1. Tentar ler do Upstash Redis
  const redisUrl = process.env.UPSTASH_REDIS_REST_URL;
  const redisToken = process.env.UPSTASH_REDIS_REST_TOKEN;
  if (redisUrl && redisToken) {
    try {
      const rRes = await fetch(`${redisUrl}/get/${file}`, {
        headers: { 'Authorization': `Bearer ${redisToken}` }
      });
      if (rRes.ok) {
        const rJson = await rRes.json();
        if (rJson && rJson.result) {
          const parsed = typeof rJson.result === 'string' ? JSON.parse(rJson.result) : rJson.result;
          return res.status(200).json({ source: 'redis', data: parsed });
        }
      }
    } catch (e) {}
  }

  // 2. Tentar ler do GitHub API (versão mais recente da branch)
  const githubToken = process.env.GITHUB_TOKEN;
  const githubRepo = process.env.GITHUB_REPO;
  const githubBranch = process.env.GITHUB_BRANCH || 'main';
  if (githubToken && githubRepo) {
    try {
      const gRes = await fetch(`https://api.github.com/repos/${githubRepo}/contents/${file}?ref=${githubBranch}`, {
        headers: {
          'Authorization': `Bearer ${githubToken}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'BATTRE-GEOportal'
        }
      });
      if (gRes.ok) {
        const fInfo = await gRes.json();
        if (fInfo.content) {
          const text = Buffer.from(fInfo.content, 'base64').toString('utf-8');
          return res.status(200).json({ source: 'github', data: JSON.parse(text) });
        }
      }
    } catch (e) {}
  }

  // 3. Tentar ler do disco local do repositório
  try {
    const fs = require('fs');
    const path = require('path');
    const localPath = path.join(process.cwd(), file);
    if (fs.existsSync(localPath)) {
      const raw = fs.readFileSync(localPath, 'utf-8');
      return res.status(200).json({ source: 'disk', data: JSON.parse(raw) });
    }
  } catch (e) {}

  return res.status(404).json({ error: 'Arquivo não encontrado' });
};
