/**
 * pdfView.js
 * ----------------------------------------
 * Minimal PDF.js viewer with +/- zoom only (50%–300%).
 *
 * Exposes:
 * - PdfView: constructor(container, canvas, overlay, pdfUrl)
 *   Methods:
 *     load() -> Promise
 *     setScale(newScale)
 *   Events:
 *     onRendered(cb): called after each render with {width,height,scale}
 */

export class PdfView {
  constructor(containerEl, canvasEl, overlayEl, pdfUrl) {
    this.container = containerEl;
    this.canvas = canvasEl;
    this.overlay = overlayEl;
    this.pdfUrl = pdfUrl;
    this.pdfDoc = null;
    this.page = null;
    this.scale = 1.0;
    this.viewport = null;
    this._renderedSubs = new Set();
  }

  onRendered(cb) { this._renderedSubs.add(cb); }
  _emitRendered() {
    const info = { width: this.viewport.width, height: this.viewport.height, scale: this.scale };
    for (const cb of this._renderedSubs) cb(info);
  }

  async load() {
    this.pdfDoc = await pdfjsLib.getDocument(this.pdfUrl).promise;
    this.page = await this.pdfDoc.getPage(1);
    this.setScale(1.0); // initial 100%
  }

  async _render() {
    const context = this.canvas.getContext("2d");
    const viewport = this.page.getViewport({ scale: this.scale });
    this.viewport = viewport;

    // Set internal pixel buffer
    this.canvas.width = Math.floor(viewport.width);
    this.canvas.height = Math.floor(viewport.height);

    // IMPORTANT: set CSS size to match pixel size (prevents blank space / mismatch)
    this.canvas.style.width = `${viewport.width}px`;
    this.canvas.style.height = `${viewport.height}px`;

    // Match overlay to canvas size
    this.overlay.setAttribute("viewBox", `0 0 ${viewport.width} ${viewport.height}`);
    this.overlay.style.width = `${viewport.width}px`;
    this.overlay.style.height = `${viewport.height}px`;

    const renderContext = { canvasContext: context, viewport };
    await this.page.render(renderContext).promise;
    this._emitRendered();
  }

  setScale(newScale) {
    const minScale = 0.5, maxScale = 3.0;
    this.scale = Math.max(minScale, Math.min(maxScale, newScale));
    this._render();
  }
}
