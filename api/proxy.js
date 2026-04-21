export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  try {
    const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: process.env.GOOGLE_CLIENT_ID,
        client_secret: process.env.GOOGLE_CLIENT_SECRET,
        refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
        grant_type: 'refresh_token'
      })
    });
    const tokenData = await tokenResp.json();
    const accessToken = tokenData.access_token;
    if (!accessToken) return res.status(500).json({ error: 'Failed to get Google access token', details: tokenData });
    const { action, folderId, fileId, fileName, content, base64, prompt } = req.body;
    if (action === 'list') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files?q='" + folderId + "'+in+parents+and+trashed=false&fields=files(id,name,mimeType)&pageSize=100", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      return res.json(await r.json());
    }
    if (action === 'get') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?alt=media", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      return res.json({ content: await r.text() });
    }
    if (action === 'getExcel') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?alt=media", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      return res.json({ base64: Buffer.from(await r.arrayBuffer()).toString('base64') });
    }
    if (action === 'parseExcel') {
      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 4000, messages: [{ role: 'user', content: [{ type: 'document', source: { type: 'base64', media_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data: base64 } }, { type: 'text', text: prompt }] }] })
      });
      return res.json(await r.json());
    }
    if (action === 'save') {
      const listR = await fetch("https://www.googleapis.com/drive/v3/files?q='" + folderId + "'+in+parents+and+name='" + fileName + "'+and+trashed=false&fields=files(id)", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      const existingId = (await listR.json()).files?.[0]?.id;
      const uploadId = existingId || (await (await fetch('https://www.googleapis.com/drive/v3/files', { method: 'POST', headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' }, body: JSON.stringify({ name: fileName, parents: [folderId] }) })).json()).id;
      await fetch("https://www.googleapis.com/upload/drive/v3/files/" + uploadId + "?uploadType=media", { method: 'PATCH', headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' }, body: content });
      return res.json({ ok: true });
    }
    return res.status(400).json({ error: 'Unknown action: ' + action });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}