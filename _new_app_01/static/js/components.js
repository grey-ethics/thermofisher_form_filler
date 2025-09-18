/**
 * components.js
 * ----------------------------------------
 * Small UI helpers to build the left-pane controls.
 *
 * Exposes:
 * - buildDeviceGrid(containerEl, onChange): renders headers + 5x4 checkboxes
 * - bindDropdown(selectEl, onChange): attach change -> call onChange(value)
 */

const REGIONS = ["N. America", "EMEA", "LATAM", "APAC"];

const ROWS = [
  { label: "General Purpose (GP)", r: 16 },
  { label: "Medical (MD)", r: 17 },
  { label: "In Vitro Diagnostics (IVD)", r: 18 },
  { label: "Gen Purpose + Cell Gene Therapy (GP + CGT)", r: 19 },
  { label: "Accessories in Scope (GP / MD / IVD / GP + CGT)", r: 20 },
];

export function buildDeviceGrid(container, onChange) {
  container.innerHTML = "";

  const head0 = document.createElement("div");
  head0.className = "cell head";
  head0.textContent = "Category";
  container.appendChild(head0);

  for (const reg of REGIONS) {
    const h = document.createElement("div");
    h.className = "cell head";
    h.textContent = reg;
    container.appendChild(h);
  }

  for (const row of ROWS) {
    const lbl = document.createElement("div");
    lbl.className = "cell";
    lbl.textContent = row.label;
    container.appendChild(lbl);

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
