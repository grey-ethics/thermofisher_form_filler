/**
 * main.js
 * ----------------------------------------
 * App bootstrap: load overlay map, init PDF viewer + overlay, wire UI events.
 */

"use strict";

import { AppState } from "./state.js";
import { PdfView } from "./pdfView.js";
import { OverlayRenderer } from "./overlay.js";
import {
  buildDeviceGrid,
  bindDropdown,
  renderBatchResults,
} from "./components.js";
import {
  fetchOverlayMap,
  exportSingle,
  processCsv,
  extractFromInput,
} from "./api.js";

const BASE_PDF_URL = "/static/pdf/reference_template.pdf";
const ZOOM_MIN = 0.5;
const ZOOM_MAX = 3.0;
const ZOOM_STEP = 0.1;

/* =========================
   THEME
   ========================= */
function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem("theme", theme);
  const btn = document.getElementById("themeToggle");
  if (btn) {
    const isLight = theme === "light";
    // The label is visually hidden via CSS, but useful for a11y/AT.
    btn.textContent = isLight ? "Dark" : "Light";
    btn.setAttribute("aria-pressed", String(isLight));
  }
}

// initial theme: saved → system preference → dark
(() => {
  const saved = localStorage.getItem("theme");
  const initial =
    saved ||
    (window.matchMedia &&
    window.matchMedia("(prefers-color-scheme: light)").matches
      ? "light"
      : "dark");
  applyTheme(initial);

  window.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("themeToggle");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const current =
        document.documentElement.getAttribute("data-theme") || "dark";
      const next = current === "light" ? "dark" : "light";
      applyTheme(next);
    });
  });
})();

/* =========================
   UTIL: Busy overlay
   ========================= */
function makeBusy() {
  const busyOverlay = document.getElementById("busyOverlay");

  function showBusy(msg = "Processing…") {
    if (!busyOverlay) return;
    const label = busyOverlay.querySelector(".busy-text");
    if (label) label.textContent = msg;
    busyOverlay.classList.remove("hidden");
    busyOverlay.setAttribute("aria-hidden", "false");
    document.body.setAttribute("aria-busy", "true");
    // Temporarily disable interactions behind the overlay
    document.body.style.pointerEvents = "none";
  }

  function hideBusy() {
    if (!busyOverlay) return;
    busyOverlay.classList.add("hidden");
    busyOverlay.setAttribute("aria-hidden", "true");
    document.body.removeAttribute("aria-busy");
    document.body.style.pointerEvents = "";
  }

  return { showBusy, hideBusy };
}

/* =========================
   HELPERS
   ========================= */
function getPdfUrlFromQuery() {
  try {
    const url = new URL(window.location.href);
    const q = (url.searchParams.get("pdf") || "").trim();
    return q || BASE_PDF_URL;
  } catch {
    return BASE_PDF_URL;
  }
}

/* =========================
   APP INIT
   ========================= */
(async function init() {
  // State
  const state = new AppState();

  // Busy UI helpers
  const { showBusy, hideBusy } = makeBusy();

  // Overlay map
  let overlayMap = null;
  try {
    overlayMap = await fetchOverlayMap();
  } catch (err) {
    console.error("Failed to load overlay map:", err);
    alert("Could not load overlay definition. Please reload the page.");
    return;
  }

  // Controls: project level + default project level
  const projectLevelSelect = document.getElementById("projectLevel");
  bindDropdown(projectLevelSelect, (v) => state.setProjectLevel(v));

  const defaultPLSelect = document.getElementById("defaultProjectLevel");
  const btnSetDefault = document.getElementById("btnSetDefault");
  const allowedPL = new Set(["L1", "L2L", "L2", "L3L"]);

  // Load saved default from localStorage and apply to left select + state
  const savedDefaultPL = (
    localStorage.getItem("defaultProjectLevel") || ""
  ).trim();
  if (savedDefaultPL && allowedPL.has(savedDefaultPL)) {
    if (projectLevelSelect) projectLevelSelect.value = savedDefaultPL;
    state.setProjectLevel(savedDefaultPL);
  }
  if (defaultPLSelect) {
    defaultPLSelect.value = savedDefaultPL || "";
  }

  if (btnSetDefault && defaultPLSelect) {
    btnSetDefault.addEventListener("click", () => {
      const val = (defaultPLSelect.value || "").trim();
      if (val && !allowedPL.has(val)) return; // ignore invalid

      if (val) localStorage.setItem("defaultProjectLevel", val);
      else localStorage.removeItem("defaultProjectLevel");

      if (projectLevelSelect) projectLevelSelect.value = val;
      state.setProjectLevel(val);

      const original = btnSetDefault.textContent;
      btnSetDefault.textContent = "Saved";
      setTimeout(() => (btnSetDefault.textContent = original), 900);
    });
  }

  // Build device grid (checkbox matrix) and wire to state
  const grid = document.getElementById("deviceGrid");
  buildDeviceGrid(grid, (id, val) => state.setTick(id, val));

  // PDF viewer
  const pdfContainer = document.getElementById("pdfContainer");
  const pdfCanvas = document.getElementById("pdfCanvas");
  const overlaySvg = document.getElementById("overlaySvg");

  if (!pdfContainer || !pdfCanvas || !overlaySvg) {
    console.error("PDF container/canvas/overlay not found.");
    return;
  }

  const pdfUrl = getPdfUrlFromQuery();
  const pdfView = new PdfView(pdfContainer, pdfCanvas, overlaySvg, pdfUrl);
  try {
    await pdfView.load();
  } catch (err) {
    console.error("PDF failed to load:", err);
    alert("Could not load the reference PDF. Please check the server.");
    return;
  }

  // Overlay renderer
  const overlay = new OverlayRenderer(overlaySvg, overlayMap);

  // Render overlay when the PDF page renders or when state changes
  pdfView.onRendered((info) => overlay.render(state.getSnapshot(), info));
  state.subscribe(() =>
    overlay.render(
      state.getSnapshot(),
      pdfView.viewport || { width: 0, height: 0, scale: pdfView.scale || 1 }
    )
  );

  // Zoom controls (ONLY +/- from 50% to 300%)
  const zIn = document.getElementById("zoomIn");
  const zOut = document.getElementById("zoomOut");
  const zLabel = document.getElementById("zoomLabel");

  function setZoom(next) {
    const clamped = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, next));
    if (Math.abs(clamped - pdfView.scale) < 0.0001) return;
    pdfView.setScale(parseFloat(clamped.toFixed(2)));
    updateZoomLabel();
  }
  function updateZoomLabel() {
    if (zLabel) zLabel.textContent = `${Math.round(pdfView.scale * 100)}%`;
  }
  if (zIn)
    zIn.addEventListener("click", () => setZoom(pdfView.scale + ZOOM_STEP));
  if (zOut)
    zOut.addEventListener("click", () => setZoom(pdfView.scale - ZOOM_STEP));
  updateZoomLabel();

  // Also allow +/- keyboard for zoom when viewer is focused
  pdfContainer.addEventListener("keydown", (e) => {
    if (e.key === "+" || e.key === "=") {
      e.preventDefault();
      setZoom(pdfView.scale + ZOOM_STEP);
    } else if (e.key === "-" || e.key === "_") {
      e.preventDefault();
      setZoom(pdfView.scale - ZOOM_STEP);
    }
  });

  // Optional: Ctrl/Cmd + wheel zoom (prevent default scroll)
  pdfContainer.addEventListener(
    "wheel",
    (e) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const delta = Math.sign(e.deltaY);
        if (delta < 0) setZoom(pdfView.scale + ZOOM_STEP);
        else if (delta > 0) setZoom(pdfView.scale - ZOOM_STEP);
      }
    },
    { passive: false }
  );

  // Download buttons (server-side templates)
  const btnPdf = document.getElementById("btnDownloadPdf");
  const btnDocx = document.getElementById("btnDownloadDocx");

  async function exportAndOpen(which = "pdf") {
    try {
      const snapshot = state.getSnapshot();
      const resp = await exportSingle(snapshot, "interactive");
      if (which === "pdf") {
        if (resp && resp.pdf_url) {
          window.open(resp.pdf_url, "_blank", "noopener,noreferrer");
        } else {
          alert("Server did not return a PDF URL.");
        }
      } else if (which === "docx") {
        if (resp && resp.docx_url) {
          window.open(resp.docx_url, "_blank", "noopener,noreferrer");
        } else {
          alert("DOCX export is disabled or not available.");
        }
      }
    } catch (err) {
      console.error(err);
      alert("Export failed: " + (err?.message || String(err)));
    }
  }

  if (btnPdf) btnPdf.addEventListener("click", () => exportAndOpen("pdf"));
  if (btnDocx) btnDocx.addEventListener("click", () => exportAndOpen("docx"));

  /* =========================
     NEW: EXPORT WITH USER TEMPLATE
     ========================= */
  const templateForm = document.getElementById("templateExportForm");
  const templateFile = document.getElementById("templateFile");

  function snapshotToLines(snap) {
    const lines = [];
    lines.push(
      `Project Level: ${snap.projectLevel || "<Choose a Project Level.>"}`
    );

    const regionNames = { 2: "N. America", 3: "EMEA", 4: "LATAM", 5: "APAC" };
    const rowNames = {
      16: "General Purpose (GP)",
      17: "Medical (MD)",
      18: "In Vitro Diagnostics (IVD)",
      19: "Gen Purpose + Cell Gene Therapy (GP + CGT)",
      20: "Accessories in Scope",
    };

    const picked = {};
    Object.entries(snap.ticks || {}).forEach(([id, val]) => {
      if (!val) return;
      const m = id.match(/^glyph_r(\d+)_c(\d+)$/);
      if (!m) return;
      const r = Number(m[1]),
        c = Number(m[2]);
      picked[r] = picked[r] || [];
      picked[r].push(regionNames[c] || `C${c}`);
    });

    Object.keys(picked)
      .map(Number)
      .sort((a, b) => a - b)
      .forEach((r) => {
        const name = rowNames[r] || `Row ${r}`;
        const vals = picked[r];
        lines.push(`${name}: ${vals.length ? vals.join(", ") : "—"}`);
      });

    return lines;
  }

  if (templateForm && templateFile) {
    templateForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const file = templateFile.files && templateFile.files[0];
      if (!file) {
        alert("Choose a .docx or .pdf template.");
        return;
      }
      try {
        showBusy("Exporting with your template…");
        const snap = state.getSnapshot();
        const lines = snapshotToLines(snap);

        const form = new FormData();
        form.append("template_file", file);
        form.append("snapshot", JSON.stringify({ content: lines }));

        const res = await fetch("/api/export", { method: "POST", body: form });
        if (!res.ok) {
          let msg = "Export failed";
          try {
            const j = await res.json();
            msg = j.error || msg;
          } catch {}
          throw new Error(msg);
        }
        const blob = await res.blob();
        const cd = res.headers.get("Content-Disposition") || "";
        const m = cd.match(/filename="?([^"]+)"?/i);
        const fallback = file.name.toLowerCase().endsWith(".pdf")
          ? "output.pdf"
          : "output.docx";
        const filename = m ? m[1] : fallback;

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      } catch (err) {
        console.error(err);
        alert(err?.message || "Export failed");
      } finally {
        hideBusy();
        templateFile.value = "";
      }
    });
  }

  /* =========================
     BATCH PROCESSING (CSV)
     ========================= */
  const csvForm = document.getElementById("csvForm");
  const csvFile = document.getElementById("csvFile");
  const batchInitial = document.getElementById("batchInitial");
  const batchResults = document.getElementById("batchResults");
  const processedCount = document.getElementById("processedCount");
  const resultsList = document.getElementById("resultsList");
  const downloadAllZip = document.getElementById("downloadAllZip");

  if (csvForm && csvFile) {
    csvForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const file = csvFile.files && csvFile.files[0];
      if (!file) {
        alert("Choose a CSV file.");
        return;
      }

      // simple disable while processing
      const submitBtn = csvForm.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.disabled = true;

      try {
        showBusy("Processing CSV…");
        const resp = await processCsv(file);

        if (batchInitial) batchInitial.classList.add("hidden");
        if (batchResults) batchResults.classList.remove("hidden");

        if (processedCount)
          processedCount.textContent = `Processed: ${resp?.processed ?? 0}`;

        if (downloadAllZip && resp?.zip_url) {
          downloadAllZip.href = resp.zip_url;
          downloadAllZip.setAttribute("download", "");
        }

        if (resultsList) {
          renderBatchResults(
            resultsList,
            Array.isArray(resp?.items) ? resp.items : []
          );
        }
      } catch (err) {
        console.error(err);
        alert("Batch failed: " + (err?.message || String(err)));
      } finally {
        hideBusy();
        if (submitBtn) submitBtn.disabled = false;
        // Optional: reset file input so the same file can be reselected
        if (csvFile) csvFile.value = "";
      }
    });
  }

  /* =========================
     EXTRACT FROM INPUT (DOCX)
     ========================= */
  const extractForm = document.getElementById("extractForm");
  const extractFile = document.getElementById("extractFile");
  const extractLines = document.getElementById("extractLines");
  const extractMedical = document.getElementById("extractMedical");

  function applyGpTicksToUi(ticksObj) {
    // Only the GP row ids (e.g., glyph_r16_c2 .. glyph_r16_c5) are provided
    if (!ticksObj || typeof ticksObj !== "object") return;
    Object.entries(ticksObj).forEach(([id, val]) => {
      const cb = document.getElementById(id);
      if (cb && "checked" in cb) cb.checked = !!val; // reflect in the visible checkboxes
      state.setTick(id, !!val); // update overlay + state
    });
  }

  if (extractForm && extractFile) {
    extractForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const file = extractFile.files && extractFile.files[0];
      if (!file) {
        alert("Choose the standard .docx document.");
        return;
      }
      try {
        showBusy("Extracting from document…");
        const out = await extractFromInput(file);

        // Show summary
        if (extractMedical) extractMedical.textContent = out?.medical || "—";
        if (extractLines) {
          const lines = Array.isArray(out?.lines) ? out.lines : [];
          extractLines.textContent = lines.length
            ? lines.join("\n")
            : "No lines returned.";
        }

        // Prefill only the GP row
        if (out?.ticks) applyGpTicksToUi(out.ticks);
      } catch (err) {
        console.error(err);
        alert("Extract failed: " + (err?.message || String(err)));
      } finally {
        hideBusy();
        if (extractFile) extractFile.value = "";
      }
    });
  }

  /* =========================
     HELP MODAL
     ========================= */
  const helpLink = document.getElementById("helpLink");
  const helpModal = document.getElementById("helpModal");
  const helpClose = document.getElementById("helpClose");
  const helpOk = document.getElementById("helpOk");

  function openHelp() {
    if (!helpModal) return;
    helpModal.classList.remove("hidden");
    helpModal.setAttribute("aria-hidden", "false");
    // focus the close button for keyboard users
    helpClose?.focus();
  }

  function closeHelp() {
    if (!helpModal) return;
    helpModal.classList.add("hidden");
    helpModal.setAttribute("aria-hidden", "true");
  }

  if (helpLink) {
    helpLink.addEventListener("click", (e) => {
      e.preventDefault();
      openHelp();
    });
  }
  if (helpClose) helpClose.addEventListener("click", closeHelp);
  if (helpOk) helpOk.addEventListener("click", closeHelp);

  if (helpModal) {
    // click backdrop to close
    helpModal.addEventListener("click", (e) => {
      if (e.target === helpModal) closeHelp();
    });
    // ESC to close
    document.addEventListener("keydown", (e) => {
      if (
        helpModal.getAttribute("aria-hidden") === "false" &&
        e.key === "Escape"
      ) {
        closeHelp();
      }
    });
  }
})();
