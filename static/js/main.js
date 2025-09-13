/**
 * main.js
 * ----------------------------------------
 * App bootstrap: load overlay map, init PDF viewer + overlay, wire UI events.
 *
 * Files used:
 * - state.js: AppState
 * - pdfView.js: PdfView
 * - overlay.js: OverlayRenderer
 * - components.js: UI builders
 * - api.js: server calls
 */

import { AppState } from "./state.js";
import { PdfView } from "./pdfView.js";
import { OverlayRenderer } from "./overlay.js";
import { buildDeviceGrid, bindDropdown, renderBatchResults } from "./components.js";
import { fetchOverlayMap, exportSingle, processCsv } from "./api.js";

const BASE_PDF_URL = "/static/pdf/reference_template.pdf";

(async function init() {
  // State
  const state = new AppState();

  // Overlay map
  const overlayMap = await fetchOverlayMap();

  // Viewer
  const pdfContainer = document.getElementById("pdfContainer");
  const pdfCanvas = document.getElementById("pdfCanvas");
  const overlaySvg = document.getElementById("overlaySvg");
  const pdfView = new PdfView(pdfContainer, pdfCanvas, overlaySvg, BASE_PDF_URL);

  await pdfView.load();

  // Overlay Renderer
  const overlay = new OverlayRenderer(overlaySvg, overlayMap);

  // Re-render overlay on:
  // - PDF render completion (for new viewport size)
  // - state changes (user interactions)
  pdfView.onRendered(info => overlay.render(state.getSnapshot(), info));
  state.subscribe(() => overlay.render(state.getSnapshot(), pdfView.viewport));

  // Zoom controls
  const zIn = document.getElementById("zoomIn");
  const zOut = document.getElementById("zoomOut");
  const fitW = document.getElementById("fitWidth");
  const fitP = document.getElementById("fitPage");
  const zLabel = document.getElementById("zoomLabel");
  function updateZoomLabel() {
    zLabel.textContent = `${Math.round(pdfView.scale * 100)}%`;
  }
  zIn.addEventListener("click", () => { pdfView.setScale(pdfView.scale + 0.1); updateZoomLabel(); });
  zOut.addEventListener("click", () => { pdfView.setScale(pdfView.scale - 0.1); updateZoomLabel(); });
  fitW.addEventListener("click", () => { pdfView.fitWidth(); updateZoomLabel(); });
  fitP.addEventListener("click", () => { pdfView.fitPage(); updateZoomLabel(); });
  updateZoomLabel();

  // Left controls
  const select = document.getElementById("projectLevel");
  bindDropdown(select, v => state.setProjectLevel(v));

  const grid = document.getElementById("deviceGrid");
  buildDeviceGrid(grid, (id, val) => state.setTick(id, val));

  // Download buttons
  const btnPdf = document.getElementById("btnDownloadPdf");
  const btnDocx = document.getElementById("btnDownloadDocx");

  async function performExport(docxToo = false) {
    try {
      const resp = await exportSingle(state.getSnapshot(), "interactive");
      // Trigger downloads
      if (resp.pdf_url) window.open(resp.pdf_url, "_blank");
      if (docxToo && resp.docx_url) window.open(resp.docx_url, "_blank");
    } catch (err) {
      alert("Export failed: " + err.message);
    }
  }
  btnPdf.addEventListener("click", () => performExport(false));
  btnDocx.addEventListener("click", () => performExport(true));

  // Batch handling
  const csvForm = document.getElementById("csvForm");
  const csvFile = document.getElementById("csvFile");
  const batchInitial = document.getElementById("batchInitial");
  const batchResults = document.getElementById("batchResults");
  const processedCount = document.getElementById("processedCount");
  const resultsList = document.getElementById("resultsList");
  const downloadAllZip = document.getElementById("downloadAllZip");

  csvForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const file = csvFile.files[0];
    if (!file) { alert("Choose a CSV file"); return; }
    try {
      const resp = await processCsv(file);
      batchInitial.classList.add("hidden");
      batchResults.classList.remove("hidden");
      processedCount.textContent = `Processed: ${resp.processed}`;
      downloadAllZip.href = resp.zip_url;
      renderBatchResults(resultsList, resp.items);
    } catch (err) {
      alert("Batch failed: " + err.message);
    }
  });
})();
