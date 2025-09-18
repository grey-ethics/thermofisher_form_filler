/**
 * main.js (single Download button version)
 * ---------------------------------------------------------
 * - Loads overlay map
 * - Renders fixed reference PDF with overlay
 * - Lets user upload Regulatory (.docx) + Template (.docx/.pdf)
 * - Extracts prefill -> paints overlay
 * - One Download button calls /download which replaces ONLY page 3
 */

"use strict";

import { AppState } from "./state.js";
import { PdfView } from "./pdfView.js";
import { OverlayRenderer } from "./overlay.js";
import { buildDeviceGrid, bindDropdown } from "./components.js";

const BASE_PDF_URL = "/static/pdf/reference_template.pdf";
const ZOOM_MIN = 0.5, ZOOM_MAX = 3.0, ZOOM_STEP = 0.1;

/* Helpers that do their own fetches (no api.js dependency) */
async function fetchOverlayMap() {
  const res = await fetch("/overlay-map");
  if (!res.ok) throw new Error("Overlay map failed to load");
  return await res.json();
}
async function extractReference(regulatoryFile) {
  const form = new FormData();
  form.append("file", regulatoryFile);
  const res = await fetch("/extract", { method: "POST", body: form });
  if (!res.ok) {
    const j = await res.json().catch(() => ({}));
    throw new Error(j?.error || "Extract failed");
  }
  return await res.json();
}

/* THEME */
function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem("theme", theme);
  const btn = document.getElementById("themeToggle");
  if (btn) {
    const isLight = theme === "light";
    btn.textContent = isLight ? "Dark" : "Light";
    btn.setAttribute("aria-pressed", String(isLight));
  }
}
(() => {
  const saved = localStorage.getItem("theme");
  const initial = saved || (window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark");
  applyTheme(initial);
  window.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("themeToggle");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const current = document.documentElement.getAttribute("data-theme") || "dark";
      applyTheme(current === "light" ? "dark" : "light");
    });
  });
})();

/* Busy overlay */
function makeBusy() {
  const busyOverlay = document.getElementById("busyOverlay");
  const label = busyOverlay?.querySelector(".busy-text");
  return {
    show(msg = "Processing…") {
      if (!busyOverlay) return;
      if (label) label.textContent = msg;
      busyOverlay.classList.remove("hidden");
      busyOverlay.setAttribute("aria-hidden", "false");
    },
    hide() {
      if (!busyOverlay) return;
      busyOverlay.classList.add("hidden");
      busyOverlay.setAttribute("aria-hidden", "true");
    }
  };
}

/* APP INIT */
(async function init() {
  const state = new AppState();
  const { show: showBusy, hide: hideBusy } = makeBusy();

  // Load overlay map
  let overlayMap = null;
  try {
    overlayMap = await fetchOverlayMap();
  } catch (err) {
    console.error(err);
    alert("Could not load overlay definition.");
    return;
  }

  // Controls: project level dropdown -> state
  const projectLevelSelect = document.getElementById("projectLevel");
  bindDropdown(projectLevelSelect, (v) => state.setProjectLevel(v));

  // Build left-side device grid (checkbox matrix)
  const grid = document.getElementById("deviceGrid");
  buildDeviceGrid(grid, (id, val) => state.setTick(id, val));

  // PDF viewer setup (fixed reference PDF)
  const pdfContainer = document.getElementById("pdfContainer");
  const pdfCanvas = document.getElementById("pdfCanvas");
  const overlaySvg = document.getElementById("overlaySvg");

  const pdfView = new PdfView(pdfContainer, pdfCanvas, overlaySvg, BASE_PDF_URL);
  try {
    await pdfView.load();
  } catch (err) {
    console.error(err);
    alert("Could not load the reference PDF.");
    return;
  }

  const overlay = new OverlayRenderer(overlaySvg, overlayMap);
  const renderOverlay = () =>
    overlay.render(
      state.getSnapshot(),
      pdfView.viewport || { width: 0, height: 0, scale: pdfView.scale || 1 }
    );

  pdfView.onRendered(() => renderOverlay());
  state.subscribe(() => renderOverlay());

  // Zoom controls
  const zIn = document.getElementById("zoomIn");
  const zOut = document.getElementById("zoomOut");
  const zLabel = document.getElementById("zoomLabel");

  function setZoom(next) {
    const clamped = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, next));
    if (Math.abs(clamped - pdfView.scale) < 1e-3) return;
    pdfView.setScale(parseFloat(clamped.toFixed(2)));
    if (zLabel) zLabel.textContent = `${Math.round(pdfView.scale * 100)}%`;
  }
  if (zIn) zIn.addEventListener("click", () => setZoom(pdfView.scale + ZOOM_STEP));
  if (zOut) zOut.addEventListener("click", () => setZoom(pdfView.scale - ZOOM_STEP));
  if (zLabel) zLabel.textContent = `${Math.round(pdfView.scale * 100)}%`;

  // Keyboard zoom when the container is focused
  pdfContainer.addEventListener("keydown", (e) => {
    if (e.key === "+" || e.key === "=") { e.preventDefault(); setZoom(pdfView.scale + ZOOM_STEP); }
    else if (e.key === "-" || e.key === "_") { e.preventDefault(); setZoom(pdfView.scale - ZOOM_STEP); }
  });

  // --- Blade 1: Upload & Process
  const regulatoryFile = document.getElementById("regulatoryFile");
  const templateFile = document.getElementById("templateFile");
  const btnProcess = document.getElementById("btnProcess");

  const extractMedical = document.getElementById("extractMedical");
  const extractLines = document.getElementById("extractLines");

  function applyTicksToUi(ticksObj) {
    if (!ticksObj || typeof ticksObj !== "object") return;
    Object.entries(ticksObj).forEach(([id, val]) => {
      const cb = document.getElementById(id);
      if (cb && "checked" in cb) cb.checked = !!val;
      state.setTick(id, !!val);
    });
  }

  if (btnProcess) {
    btnProcess.addEventListener("click", async () => {
      const reg = regulatoryFile?.files?.[0] || null;
      const tmpl = templateFile?.files?.[0] || null;

      if (!reg) { alert("Please choose the Regulatory Document (.docx)."); return; }
      if (!tmpl) { alert("Please choose the Template Document (.docx or .pdf)."); return; }

      try {
        showBusy("Processing documents…");
        // Save template in state for the later download
        state.setTemplate(tmpl);

        const out = await extractReference(reg);
        // Show summary
        if (extractMedical) extractMedical.textContent = out?.medical || "—";
        if (extractLines) {
          const lines = Array.isArray(out?.lines) ? out.lines : [];
          extractLines.textContent = lines.length ? lines.join("\n") : "No lines returned.";
        }
        // Prefill returned ticks into UI and overlay
        if (out?.ticks) applyTicksToUi(out.ticks);
      } catch (err) {
        console.error(err);
        alert(err?.message || "Extract failed.");
      } finally {
        hideBusy();
      }
    });
  }

  // --- Blade 2: Single Download button
  // Expect one button with id="btnDownload"
  const btnDownload = document.getElementById("btnDownload");

  // (Safety) If old buttons still exist, wire them to the same handler:
  const btnPdf = document.getElementById("btnDownloadPdf");
  const btnDocx = document.getElementById("btnDownloadDocx");

  async function doDownload() {
    const tmpl = state.templateFile;
    if (!tmpl) {
      alert("Upload a Template Document in the first blade, then try again.");
      return;
    }
    const snapshot = state.getSnapshot();

    const form = new FormData();
    form.append("template_file", tmpl);
    form.append("snapshot", JSON.stringify(snapshot));

    try {
      showBusy("Exporting…");
      const res = await fetch("/download", { method: "POST", body: form });
      if (!res.ok) {
        const j = await res.json().catch(() => ({}));
        throw new Error(j?.error || "Download failed");
      }
      const blob = await res.blob();
      const cd = res.headers.get("Content-Disposition") || "";
      const m = cd.match(/filename="?([^"]+)"?/i);
      const fallback = tmpl.name.toLowerCase().endsWith(".pdf") ? "output.pdf" : "output.docx";
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
      alert(err?.message || "Download failed.");
    } finally {
      hideBusy();
    }
  }

  if (btnDownload) btnDownload.addEventListener("click", doDownload);
  if (btnPdf) btnPdf.addEventListener("click", doDownload);
  if (btnDocx) btnDocx.addEventListener("click", doDownload);

  // Help modal
  const helpLink = document.getElementById("helpLink");
  const helpModal = document.getElementById("helpModal");
  const helpClose = document.getElementById("helpClose");
  const helpOk = document.getElementById("helpOk");

  function openHelp() { if (!helpModal) return; helpModal.classList.remove("hidden"); helpModal.setAttribute("aria-hidden", "false"); helpClose?.focus(); }
  function closeHelp() { if (!helpModal) return; helpModal.classList.add("hidden"); helpModal.setAttribute("aria-hidden", "true"); }

  if (helpLink) helpLink.addEventListener("click", (e) => { e.preventDefault(); openHelp(); });
  if (helpClose) helpClose.addEventListener("click", closeHelp);
  if (helpOk) helpOk.addEventListener("click", closeHelp);
  if (helpModal) {
    helpModal.addEventListener("click", (e) => { if (e.target === helpModal) closeHelp(); });
    document.addEventListener("keydown", (e) => {
      if (helpModal.getAttribute("aria-hidden") === "false" && e.key === "Escape") closeHelp();
    });
  }
})();
