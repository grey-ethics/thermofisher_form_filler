/**
 * api.js
 * ----------------------------------------
 * Fetch wrappers for server endpoints.
 *
 * Exposes:
 * - fetchOverlayMap()
 * - extractReference(file)           // POST /extract
 * - exportWithTemplate(snapshot, file)  // POST /api/export
 */

export async function fetchOverlayMap() {
  const res = await fetch("/overlay-map");
  if (!res.ok) throw new Error("Failed to load overlay map");
  return res.json();
}

export async function extractReference(file) {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch("/extract", { method: "POST", body: form });
  if (!res.ok) {
    let msg = "Extract failed";
    try { const j = await res.json(); msg = j.error || msg; } catch {}
    throw new Error(msg);
  }
  return res.json();
}

/**
 * Builds the "content lines" that the server will inject as Page 3 text.
 * (Mirrors what we render in the UI.)
 */
export function snapshotToLines(snapshot) {
  const lines = [];
  lines.push(`Project Level: ${snapshot.projectLevel || "<Choose a Project Level.>"}`);

  const regionNames = { 2: "N. America", 3: "EMEA", 4: "LATAM", 5: "APAC" };
  const rowNames = {
    16: "General Purpose (GP)",
    17: "Medical (MD)",
    18: "In Vitro Diagnostics (IVD)",
    19: "Gen Purpose + Cell Gene Therapy (GP + CGT)",
    20: "Accessories in Scope",
  };

  const picked = {};
  Object.entries(snapshot.ticks || {}).forEach(([id, val]) => {
    if (!val) return;
    const m = id.match(/^glyph_r(\d+)_c(\d+)$/);
    if (!m) return;
    const r = Number(m[1]), c = Number(m[2]);
    picked[r] = picked[r] || [];
    picked[r].push(regionNames[c] || `C${c}`);
  });

  Object.keys(picked).map(Number).sort((a,b)=>a-b).forEach((r) => {
    const name = rowNames[r] || `Row ${r}`;
    const vals = picked[r];
    lines.push(`${name}: ${vals.length ? vals.join(", ") : "—"}`);
  });

  return lines;
}

export async function exportWithTemplate(snapshot, templateFile) {
  if (!templateFile) throw new Error("Please upload a Template Document first.");
  const lines = snapshotToLines(snapshot);
  const form = new FormData();
  form.append("template_file", templateFile);
  form.append("snapshot", JSON.stringify({ content: lines }));

  const res = await fetch("/api/export", { method: "POST", body: form });
  if (!res.ok) {
    let msg = "Export failed";
    try { const j = await res.json(); msg = j.error || msg; } catch {}
    throw new Error(msg);
  }
  // Return blob + filename for download
  const blob = await res.blob();
  const cd = res.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename="?([^"]+)"?/i);
  const filename = m ? m[1] : "output.bin";
  return { blob, filename };
}
