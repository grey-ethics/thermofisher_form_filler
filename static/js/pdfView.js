/**
 * pdfView.js
 * ----------------------------------------
 * PDF.js viewer bootstrap & simple zoom controls.
 *
 * Exposes:
 * - PdfView: constructor(container, canvas, overlay, pdfUrl)
 *   Methods:
 *     load() -> Promise
 *     setScale(newScale)
 *     fitWidth(), fitPage()
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
    this.fitWidth(); // initial render
  }

  async _render() {
    const context = this.canvas.getContext("2d");
    const viewport = this.page.getViewport({ scale: this.scale });
    this.viewport = viewport;

    this.canvas.width = Math.floor(viewport.width);
    this.canvas.height = Math.floor(viewport.height);

    // size overlay to match canvas
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

  fitWidth() {
    const viewW = this.container.clientWidth - 10; // padding
    this.page.getViewport({ scale: 1.0 });
    const initial = 1.0;
    const vw = this.page.getViewport({ scale: initial }).width;
    const s = viewW / vw;
    this.setScale(s);
  }

  fitPage() {
    const viewW = this.container.clientWidth - 10;
    const viewH = this.container.clientHeight - 10;
    const vw = this.page.getViewport({ scale: 1.0 }).width;
    const vh = this.page.getViewport({ scale: 1.0 }).height;
    const s = Math.min(viewW / vw, viewH / vh);
    this.setScale(s);
  }
}
