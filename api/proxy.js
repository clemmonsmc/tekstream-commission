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
    const { access_token: accessToken, error: tokenError } = await tokenResp.json();
    if (!accessToken) return res.status(500).json({ error: 'Token error: ' + tokenError });

    const { action, folderId, fileId, fileName, content, sheetName } = req.body;
    const auth = { 'Authorization': 'Bearer ' + accessToken };

    if (action === 'list') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files?q='" + folderId + "'+in+parents+and+trashed=false&fields=files(id,name,mimeType)&pageSize=100", { headers: auth });
      return res.json(await r.json());
    }

    if (action === 'get') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?alt=media", { headers: auth });
      return res.json({ content: await r.text() });
    }

    if (action === 'getSheetCSV') {
      // Copy the xlsx as a native Google Sheet, read CSV, delete copy
      const copyResp = await fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy', {
        method: 'POST',
        headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: '_tmp_stmt_' + Date.now(), mimeType: 'application/vnd.google-apps.spreadsheet' })
      });
      const copy = await copyResp.json();
      if (copy.error) return res.status(500).json({ error: 'Copy failed: ' + copy.error.message });

      const copyId = copy.id;
      try {
        // Get sheet list
        const metaR = await fetch('https://sheets.googleapis.com/v4/spreadsheets/' + copyId + '?fields=sheets.properties', { headers: auth });
        const meta = await metaR.json();
        const sheets = meta.sheets?.map(s => ({ name: s.properties.title, gid: s.properties.sheetId })) || [];

        // Find requested sheet
        const target = sheets.find(s => s.name === sheetName) || sheets.find(s => s.name.includes(String(sheetName))) || sheets[0];
        if (!target) return res.json({ csv: '', sheets: sheets.map(s => s.name) });

        // Export as CSV
        const csvR = await fetch('https://docs.google.com/spreadsheets/d/' + copyId + '/export?format=csv&gid=' + target.gid, { headers: auth });
        const csv = await csvR.text();

        return res.json({ csv, sheets: sheets.map(s => s.name), sheetName: target.name });
      } finally {
        // Always delete the temp copy
        await fetch('https://www.googleapis.com/drive/v3/files/' + copyId, { method: 'DELETE', headers: auth });
      }
    }

    if (action === 'save') {
      const listR = await fetch("https://www.googleapis.com/drive/v3/files?q='" + folderId + "'+in+parents+and+name='" + fileName + "'+and+trashed=false&fields=files(id)", { headers: auth });
      const existingId = (await listR.json()).files?.[0]?.id;
      const uploadId = existingId || (await (await fetch('https://www.googleapis.com/drive/v3/files', { method: 'POST', headers: { ...auth, 'Content-Type': 'application/json' }, body: JSON.stringify({ name: fileName, parents: [folderId] }) })).json()).id;
      await fetch('https://www.googleapis.com/upload/drive/v3/files/' + uploadId + '?uploadType=media', { method: 'PATCH', headers: { ...auth, 'Content-Type': 'application/json' }, body: content });
      return res.json({ ok: true });
    }

    return res.status(400).json({ error: 'Unknown action: ' + action });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}