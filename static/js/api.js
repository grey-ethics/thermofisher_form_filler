/**
 * api.js
 * ----------------------------------------
 * Fetch wrappers for server endpoints.
 *
 * Exposes:
 * - fetchOverlayMap()
 * - exportSingle(stateSnapshot, companyId)
 * - processCsv(file)
 */

export async function fetchOverlayMap() {
  const res = await fetch("/overlay-map");
  if (!res.ok) throw new Error("Failed to load overlay map");
  return res.json();
}

export async function exportSingle(stateSnapshot, companyId = "company") {
  const body = {
    company_id: companyId,
    projectLevel: stateSnapshot.projectLevel,
    ticks: stateSnapshot.ticks
  };
  const res = await fetch("/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error("Export failed");
  return res.json();
}

export async function processCsv(file) {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch("/batch", { method: "POST", body: form });
  if (!res.ok) throw new Error("Batch failed");
  return res.json();
}
