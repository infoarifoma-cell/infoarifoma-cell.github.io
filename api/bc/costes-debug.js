// GET /api/bc/costes-debug?token=XXX
// Temporal: explorar estructura de G/L Entries y dimensiones en BC

export default async function handler(req, res) {
  const token = req.query.token || req.headers['x-bc-token'];
  if (!token) return res.status(401).json({ error: 'Token requerido' });

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  try {
    // 1. Get company ID
    const cRes = await fetch(base, { headers });
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada');
    const cid = company.id;

    const results = {};

    // 2. List dimensions
    try {
      const dimRes = await fetch(`${base}(${cid})/dimensions?$top=50`, { headers });
      results.dimensions = dimRes.ok ? await dimRes.json() : { status: dimRes.status, text: await dimRes.text() };
    } catch(e) { results.dimensions = { error: e.message }; }

    // 3. Try generalLedgerEntries (top 5, accounts 600-700 range)
    try {
      const glRes = await fetch(`${base}(${cid})/generalLedgerEntries?$top=5&$filter=accountNumber ge '600' and accountNumber lt '800'&$orderby=postingDate desc`, { headers });
      results.glEntries = glRes.ok ? await glRes.json() : { status: glRes.status, text: await glRes.text() };
    } catch(e) { results.glEntries = { error: e.message }; }

    // 4. Try G/L entries with dimension set lines (expand)
    try {
      const glDimRes = await fetch(`${base}(${cid})/generalLedgerEntries?$top=3&$filter=accountNumber ge '600' and accountNumber lt '800'&$expand=dimensionSetLines&$orderby=postingDate desc`, { headers });
      results.glEntriesWithDimensions = glDimRes.ok ? await glDimRes.json() : { status: glDimRes.status, text: await glDimRes.text() };
    } catch(e) { results.glEntriesWithDimensions = { error: e.message }; }

    // 5. Try projects endpoint
    try {
      const projRes = await fetch(`${base}(${cid})/projects?$top=20`, { headers });
      results.projects = projRes.ok ? await projRes.json() : { status: projRes.status, text: await projRes.text() };
    } catch(e) { results.projects = { error: e.message }; }

    return res.status(200).json({ ok: true, companyId: cid, results });
  } catch (error) {
    return res.status(500).json({ ok: false, error: error.message });
  }
}
