/**
 * overlay.js
 * ----------------------------------------
 * Draws SVG overlay elements (dropdown text + ticks) on top of the PDF canvas.
 *
 * Exposes:
 * - OverlayRenderer(svgEl, overlayMap)
 *   Methods:
 *     render(stateSnapshot, viewportInfo)
 */

export class OverlayRenderer {
  constructor(svgEl, overlayMap) {
    this.svg = svgEl;
    this.map = overlayMap;
    this.pageMap = overlayMap.pages["1"]; // single-page template
  }

  _px(norm, total) { return norm * total; }

  _clear() { while (this.svg.firstChild) this.svg.removeChild(this.svg.firstChild); }

  _drawText(x, y, w, h, text, style, scale) {
    const basePx = (style.fontSizePt || 10) * 1.3333; // pt -> px
    const fontPx = basePx * (scale || 1);
    const tx = document.createElementNS("http://www.w3.org/2000/svg", "text");
    tx.setAttribute("x", x.toFixed(2));
    tx.setAttribute("y", (y + h * 0.8).toFixed(2));
    tx.setAttribute("fill", style.color || "#000");
    tx.setAttribute("font-size", fontPx.toFixed(2));
    tx.setAttribute("font-family", style.fontFamily || "sans-serif");
    if (style.italic) tx.setAttribute("font-style", "italic");
    tx.textContent = text;
    this.svg.appendChild(tx);
  }

  _drawTick(x, y, scale) {
    const size = 14 * (scale || 1);
    const tx = document.createElementNS("http://www.w3.org/2000/svg", "text");
    tx.setAttribute("x", x.toFixed(2));
    tx.setAttribute("y", y.toFixed(2));
    tx.setAttribute("class", "tick");
    tx.setAttribute("font-size", size.toFixed(2));
    tx.setAttribute("font-weight", "700");
    tx.textContent = "✓";
    this.svg.appendChild(tx);
    return tx;
  }

  render(stateSnapshot, viewportInfo) {
    const vw = viewportInfo.width;
    const vh = viewportInfo.height;
    const scale = viewportInfo.scale || 1;

    this._clear();

    // Dropdown text
    const dd = this.pageMap.dropdown;
    if (dd) {
      const isPlaceholder = !stateSnapshot.projectLevel;
      const text = isPlaceholder ? dd.values[0] : stateSnapshot.projectLevel;
      const style = isPlaceholder ? dd.styles.placeholder : dd.styles.selected;
      const x = this._px(dd.x, vw), y = this._px(dd.y, vh);
      const w = this._px(dd.w, vw), h = this._px(dd.h, vh);
      this._drawText(x, y, w, h, text, style, scale);
    }

    // Ticks
    const ticks = this.pageMap.ticks || [];
    for (const t of ticks) {
      if (stateSnapshot.ticks[t.id]) {
        const x = this._px(t.x, vw);
        const y = this._px(t.y, vh);
        this._drawTick(x, y, scale);
      }
    }
  }
}
