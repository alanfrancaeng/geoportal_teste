// api/save.js - Vercel Serverless Function
// Persistência em nuvem para GEOportal & BI-Obras BATTRE

module.exports = async (req, res) => {
  // Configurar CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido. Utilize POST.' });
  }

  const { file } = req.query;
  const allowedFiles = ['plano_acao.json', 'estoque.json', 'dados_obras.json'];

  if (!file || !allowedFiles.includes(file)) {
    return res.status(400).json({ error: 'Arquivo inválido ou não permitido.' });
  }

  // Obter o corpo da requisição
  let bodyData = req.body;
  if (typeof bodyData === 'string') {
    try {
      bodyData = JSON.parse(bodyData);
    } catch (e) {
      return res.status(400).json({ error: 'JSON inválido no corpo da requisição.' });
    }
  }

  if (!bodyData) {
    return res.status(400).json({ error: 'Corpo da requisição vazio.' });
  }

  const jsonString = JSON.stringify(bodyData, null, 2);
  const now = new Date().toISOString();

  // ──────────────────────────────────────────────────────────────────────────
  // 1. ESTRATÉGIA A: GitHub API (Commit Direto no Repositório)
  // Se configurado GITHUB_TOKEN e GITHUB_REPO no painel da Vercel
  // ──────────────────────────────────────────────────────────────────────────
  const githubToken = process.env.GITHUB_TOKEN;
  const githubRepo = process.env.GITHUB_REPO; // Ex: 'usuario/protal-de-obras'
  const githubBranch = process.env.GITHUB_BRANCH || 'main';

  if (githubToken && githubRepo) {
    try {
      const getUrl = `https://api.github.com/repos/${githubRepo}/contents/${file}?ref=${githubBranch}`;
      const getRes = await fetch(getUrl, {
        headers: {
          'Authorization': `Bearer ${githubToken}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'BATTRE-GEOportal'
        }
      });

      let sha = null;
      if (getRes.ok) {
        const fileInfo = await getRes.json();
        sha = fileInfo.sha;
      }

      const putUrl = `https://api.github.com/repos/${githubRepo}/contents/${file}`;
      const putBody = {
        message: `chore(dados): atualizar ${file} via portal web [${now}]`,
        content: Buffer.from(jsonString).toString('base64'),
        branch: githubBranch
      };
      if (sha) putBody.sha = sha;

      const putRes = await fetch(putUrl, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${githubToken}`,
          'Accept': 'application/vnd.github.v3+json',
          'Content-Type': 'application/json',
          'User-Agent': 'BATTRE-GEOportal'
        },
        body: JSON.stringify(putBody)
      });

      if (putRes.ok) {
        return res.status(200).json({
          status: 'ok',
          storage: 'github',
          file,
          updated_at: now,
          message: `Salvo e commitado no GitHub (${githubRepo} @ ${githubBranch}) com sucesso!`
        });
      } else {
        const errText = await putRes.text();
        console.warn('[GITHUB-API-ERROR]', errText);
      }
    } catch (errGithub) {
      console.warn('[GITHUB-SAVE-ERROR]', errGithub.message);
    }
  }

  // ──────────────────────────────────────────────────────────────────────────
  // 2. ESTRATÉGIA B: Upstash Redis (Vercel Marketplace)
  // Se configurado UPSTASH_REDIS_REST_URL e UPSTASH_REDIS_REST_TOKEN
  // ──────────────────────────────────────────────────────────────────────────
  const redisUrl = process.env.UPSTASH_REDIS_REST_URL;
  const redisToken = process.env.UPSTASH_REDIS_REST_TOKEN;

  if (redisUrl && redisToken) {
    try {
      const rRes = await fetch(`${redisUrl}/set/${file}`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${redisToken}`,
          'Content-Type': 'text/plain'
        },
        body: jsonString
      });
      if (rRes.ok) {
        return res.status(200).json({
          status: 'ok',
          storage: 'redis',
          file,
          updated_at: now,
          message: 'Salvo no banco de dados na nuvem (Redis) com sucesso!'
        });
      }
    } catch (errRedis) {
      console.warn('[REDIS-SAVE-ERROR]', errRedis.message);
    }
  }

  // ──────────────────────────────────────────────────────────────────────────
  // 3. ESTRATÉGIA C: Vercel Blob Storage
  // Se configurado BLOB_READ_WRITE_TOKEN
  // ──────────────────────────────────────────────────────────────────────────
  const blobToken = process.env.BLOB_READ_WRITE_TOKEN;
  if (blobToken) {
    try {
      const { put } = require('@vercel/blob');
      const blob = await put(file, jsonString, {
        access: 'public',
        addRandomSuffix: false,
        token: blobToken
      });
      return res.status(200).json({
        status: 'ok',
        storage: 'blob',
        file,
        url: blob.url,
        updated_at: now,
        message: 'Salvo no Vercel Blob Storage com sucesso!'
      });
    } catch (errBlob) {
      console.warn('[BLOB-SAVE-ERROR]', errBlob.message);
    }
  }

  // ──────────────────────────────────────────────────────────────────────────
  // 4. ESTRATÉGIA D: Local Disk (se rodando localmente com Node ou server.py)
  // ──────────────────────────────────────────────────────────────────────────
  try {
    const fs = require('fs');
    const path = require('path');
    const localPath = path.join(process.cwd(), file);
    fs.writeFileSync(localPath, jsonString, 'utf-8');
    return res.status(200).json({
      status: 'ok',
      storage: 'disk',
      file,
      updated_at: now,
      message: 'Salvo no disco com sucesso!'
    });
  } catch (errDisk) {
    // Na Vercel sem token configurado, o disco é read-only
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Fallback: Retorna OK informando que o front-end deve persistir no localStorage
  // ──────────────────────────────────────────────────────────────────────────
  return res.status(200).json({
    status: 'ok',
    storage: 'local_browser',
    file,
    updated_at: now,
    message: 'Salvo localmente no navegador. (Para persistir na nuvem para todos os usuários, adicione GITHUB_TOKEN ou Vercel Blob no painel da Vercel).'
  });
};
