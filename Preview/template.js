// Lightweight template loader: fetch Master PPT.pptx as an ArrayBuffer (no-cache)
// and (optionally) return a JSZip if you want to work with the zip directly.
export async function loadTemplateArrayBuffer() {
  const res = await fetch('./Master PPT.pptx', { cache: 'no-store' });
  if (!res.ok) throw new Error(`Template fetch failed: ${res.status}`);
  return await res.arrayBuffer();
}

export async function loadTemplateZip() {
  // dynamic import of JSZip ESM build from CDN
  const { default: JSZip } = await import('https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm');
  const ab = await loadTemplateArrayBuffer();
  return await JSZip.loadAsync(ab);
}