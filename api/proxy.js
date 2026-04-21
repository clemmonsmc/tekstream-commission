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

    const { action, folderId, fileId, fileName, content, sheetName } = req.body;

    if (action === 'list') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files?q='" + folderId + "'+in+parents+and+trashed=false&fields=files(id,name,mimeType)&pageSize=100", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      return res.json(await r.json());
    }

    if (action === 'get') {
      const r = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?alt=media", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      return res.json({ content: await r.text() });
    }

    if (action === 'getSheetCSV') {
      // Export a specific sheet from an Excel/Sheets file as CSV using Drive export
      // First get the file metadata to find sheet IDs
      const metaR = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "?fields=id,name,mimeType", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      const meta = await metaR.json();
      
      // If it's an xlsx, we need to use the Sheets API - first check if it has a Sheets version
      // Export as CSV for a specific sheet by name using the export URL
      // Drive allows exporting XLSX as CSV but only for the first sheet
      // For multi-sheet, we use the Sheets API gid parameter
      
      // Get sheet list via Sheets API
      const sheetsMetaR = await fetch("https://sheets.googleapis.com/v4/spreadsheets/" + fileId + "?fields=sheets.properties", { headers: { 'Authorization': 'Bearer ' + accessToken } });
      const sheetsMeta = await sheetsMetaR.json();
      
      if (sheetsMeta.error) {
        // File might be xlsx not native Sheets - try direct export
        const csvR = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "/export?mimeType=text%2Fcsv", { headers: { 'Authorization': 'Bearer ' + accessToken } });
        return res.json({ csv: await csvR.text(), sheets: [] });
      }

      const sheets = sheetsMeta.sheets?.map(s => ({ name: s.properties.title, gid: s.properties.sheetId })) || [];
      
      // Find requested sheet
      const targetSheet = sheets.find(s => s.name === sheetName) || sheets[0];
      if (!targetSheet) return res.json({ csv: '', sheets: sheets.map(s => s.name) });

      // Export that specific sheet as CSV
      const csvR = await fetch("https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=csv&gid=" + targetSheet.gid, { headers: { 'Authorization': 'Bearer ' + accessToken } });
      const csv = await csvR.text();
      return res.json({ csv, sheets: sheets.map(s => s.name), sheetName: targetSheet.name });
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