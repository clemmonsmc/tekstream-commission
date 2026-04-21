export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    // Get fresh Google access token using refresh token
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
    if (!accessToken) return res.status(401).json({ error: 'Failed to get Google access token', details: tokenData });

    const { action, folderId, fileId, fileName } = req.body;

    // LIST files in a folder
    if (action === 'list') {
      const r = await fetch(`https://www.googleapis.com/drive/v3/files?q='${folderId}'+in+parents&fields=files(id,name,mimeType,size)&pageSize=100`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      const d = await r.json();
      return res.json(d);
    }

    // DOWNLOAD a file by ID (CSV/JSON = export as text, XLSX = download as binary base64)
    if (action === 'download') {
      const metaResp = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?fields=name,mimeType`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      const meta = await metaResp.json();
      const mimeType = meta.mimeType || '';

      let downloadUrl;
      if (mimeType.includes('spreadsheet') || mimeType.includes('google-apps')) {
        downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/csv`;
      } else {
        downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
      }

      const fileResp = await fetch(downloadUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });

      // For XLSX files return base64, otherwise return text
      if (mimeType.includes('spreadsheetml') || (fileName && fileName.endsWith('.xlsx'))) {
        const buffer = await fileResp.arrayBuffer();
        const base64 = Buffer.from(buffer).toString('base64');
        return res.json({ type: 'xlsx', base64, name: meta.name });
      } else {
        const text = await fileResp.text();
        return res.json({ type: 'text', content: text, name: meta.name });
      }
    }

    // ANTHROPIC call (for XLSX parsing)
    if (action === 'anthropic') {
      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': process.env.ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify(req.body.payload)
      });
      const d = await r.json();
      return res.json(d);
    }

    res.status(400).json({ error: 'Unknown action' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}