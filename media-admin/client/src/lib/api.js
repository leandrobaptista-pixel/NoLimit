const API_BASE = import.meta.env.VITE_API_BASE_URL || '';

async function parseResponse(response) {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.message || 'Request failed.');
  }
  return data;
}

function buildUrl(path, params) {
  const url = new URL(`${API_BASE}${path}`, window.location.origin);
  Object.entries(params || {}).forEach(([key, value]) => {
    if (value === undefined || value === null || value === '' || value === 'all') return;
    url.searchParams.set(key, value);
  });
  return url.toString();
}

export async function listCategories() {
  const response = await fetch(buildUrl('/api/categories'));
  return parseResponse(response);
}

export async function listMedia(params) {
  const response = await fetch(buildUrl('/api/media', params));
  return parseResponse(response);
}

export async function uploadMedia(formData) {
  const response = await fetch(buildUrl('/api/media/upload'), {
    method: 'POST',
    body: formData
  });
  return parseResponse(response);
}

export async function generateArt(mediaId) {
  const response = await fetch(buildUrl(`/api/media/${mediaId}/generate-art`), {
    method: 'POST'
  });
  return parseResponse(response);
}

export async function updateMediaStatus(mediaId, payload) {
  const response = await fetch(buildUrl(`/api/media/${mediaId}/status`), {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });
  return parseResponse(response);
}
