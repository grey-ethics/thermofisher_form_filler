/**
 * calibrate.js
 * ----------------------------------------
 * Optional helper: click on the PDF to log normalized coords.
 */
(async function() {
  const pdfUrl = "/static/pdf/reference_template.pdf";
  const canvas = document.getElementById("pdfCanvas");
  const overlay = document.getElementById("overlaySvg");
  const log = document.getElementById("log");

  const pdf = await pdfjsLib.getDocument(pdfUrl).promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({ scale: 1.2 });

  canvas.width = viewport.width;
  canvas.height = viewport.height;
  overlay.setAttribute("viewBox", `0 0 ${viewport.width} ${viewport.height}`);
  overlay.style.width = `${viewport.width}px`;
  overlay.style.height = `${viewport.height}px`;

  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;

  overlay.style.pointerEvents = "auto";
  overlay.addEventListener("click", (e) => {
    const rect = overlay.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;
    const nx = xPx / rect.width;
    const ny = yPx / rect.height;
    const text = `{ "x": ${nx.toFixed(3)}, "y": ${ny.toFixed(3)} },`;
    log.textContent += text + "\n";
  });
})();
