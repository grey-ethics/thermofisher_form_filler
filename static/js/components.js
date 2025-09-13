/**
 * components.js
 * ----------------------------------------
 * Small UI helpers to build the left-pane controls & batch cards.
 *
 * Exposes:
 * - buildDeviceGrid(containerEl, onChange): renders headers + 5x4 checkboxes
 * - bindDropdown(selectEl, onChange): attach change -> call onChange(value)
 * - renderBatchResults(listEl, items): render result cards with download links
 */

const REGIONS = ["N. America","EMEA","LATAM","APAC"];
const ROWS = [
  { label: "Row 16", r: 16 },
  { label: "Row 17", r: 17 },
  { label: "Row 18", r: 18 },
  { label: "Row 19", r: 19 },
  { label: "Row 20", r: 20 }
];

export function buildDeviceGrid(container, onChange) {
  container.innerHTML = "";

  // header row: blank cell + 4 regions
  const head0 = document.createElement("div");
  head0.className = "cell head";
  head0.textContent = "";
  container.appendChild(head0);

  for (const reg of REGIONS) {
    const h = document.createElement("div");
    h.className = "cell head";
    h.textContent = reg;
    container.appendChild(h);
  }

  // body rows
  for (const row of ROWS) {
    // left label
    const lbl = document.createElement("div");
    lbl.className = "cell";
    lbl.textContent = row.label;
    container.appendChild(lbl);

    // 4 region checkboxes
    for (let c = 2; c <= 5; c++) {
      const id = `glyph_r${row.r}_c${c}`;
      const cell = document.createElement("div");
      cell.className = "cell";

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = id;

      const lab = document.createElement("label");
      lab.setAttribute("for", id);
      lab.textContent = "Tick";

      cb.addEventListener("change", () => onChange(id, cb.checked));

      cell.appendChild(cb);
      cell.appendChild(lab);
      container.appendChild(cell);
    }
  }
}

export function bindDropdown(selectEl, onChange) {
  selectEl.addEventListener("change", () => {
    const v = selectEl.value || "";
    onChange(v);
  });
}

export function renderBatchResults(listEl, items) {
  listEl.innerHTML = "";
  items.forEach((it, idx) => {
    const card = document.createElement("div");
    card.className = "result-card";

    const serial = document.createElement("div");
    serial.className = "serial";
    serial.textContent = String(idx + 1).padStart(2, "0");

    const info = document.createElement("div");
    info.innerHTML = `<strong>${it.company_id}</strong><br/><small>Completed</small>`;

    const actions = document.createElement("div");
    actions.className = "actions";
    const aPdf = document.createElement("a");
    aPdf.href = it.pdf_url;
    aPdf.className = "btn btn-primary";
    aPdf.textContent = "Download PDF";
    actions.appendChild(aPdf);

    if (it.docx_url) {
      const aDocx = document.createElement("a");
      aDocx.href = it.docx_url;
      aDocx.className = "btn";
      aDocx.textContent = "Download DOCX";
      actions.appendChild(aDocx);
    }

    card.appendChild(serial);
    card.appendChild(info);
    card.appendChild(actions);
    listEl.appendChild(card);
  });
}
