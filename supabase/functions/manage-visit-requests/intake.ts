export const allowedStaffRoles = new Set(['owner', 'admin', 'manager', 'office']);
export function bearerToken(header: string | null) {
  return /^Bearer\s+(.+)$/i.exec(header || '')?.[1]?.trim() || '';
}
export function mapWebsiteRequest(row: {id: string; updated_at: string; payload: Record<string, unknown>}) {
  const p = row.payload || {};
  const submitted = String(p.submittedAt || row.updated_at);
  if (!/^[0-9a-f-]{36}$/i.test(row.id) || !Number.isFinite(Date.parse(submitted))) {
    throw new Error('A website request has an invalid identifier or date. No records were removed.');
  }
  const date = String(p.preferredDate || '');
  return {
    source_record_id: row.id,
    submitted_at: submitted,
    full_name: String(p.name || 'Unnamed request'), email: String(p.email || ''),
    phone: String(p.phone || ''), address: String(p.address || ''), city: String(p.city || ''),
    preferred_date: /^\d{4}-\d{2}-\d{2}$/.test(date) && Number.isFinite(Date.parse(date)) ? date : null,
    project_type: ['Ceiling', 'Crown Molding', 'Coffered Ceiling'].includes(String(p.projectType)) ? 'Crown Molding · Ceiling · Coffered Ceiling' : String(p.projectType || ''), message: String(p.details || ''),
    source_page_url: String(p.pageUrl || '')
  };
}
