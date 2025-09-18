/* Base styles, variables, resets — LIGHT THEME */

:root{
  --bg: #f5f7fb;
  --panel: #ffffff;
  --panel-2: #f9fafb;
  --card: #ffffff;

  --text: #0f172a;
  --muted: #475569;

  --brand: #2563eb;   /* primary */
  --brand-2:#06b6d4;  /* accent */

  --positive:#16a34a;
  --warning:#d97706;
  --danger:#dc2626;

  --ring: rgba(37,99,235,.25);
  --radius: 14px;

  --shadow: 0 14px 40px rgba(2,6,23,.08);
  --shadow-soft: 0 8px 24px rgba(2,6,23,.06);
  --shadow-tiny: 0 3px 12px rgba(2,6,23,.06);

  --border: #e5e7eb;
}

* { box-sizing: border-box; }
html, body { height: 100%; }
html { scroll-behavior: smooth; }

/* Page vertical layout so footer never overlaps */
body {
  margin:0;
  display:flex; flex-direction:column;
  min-height:100vh;
  font-family: "Inter", ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, "Helvetica Neue", Arial, "Noto Sans", "Apple Color Emoji","Segoe UI Emoji";
  color: var(--text);
  background: var(--bg);
}
main { flex: 1; }

a { color: var(--brand); text-decoration: none; }
a:hover { text-decoration: underline; }

button, input, select, textarea { font: inherit; color: inherit; }

.container { max-width: 1200px; margin: 0 auto; padding: 24px; }

h1,h2,h3 { margin: 0 0 10px }
h1 { font-size: 2rem; }
h2 { font-size: 1.25rem; color: #111827; }
p { color: var(--muted); line-height:1.65 }

.card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow-soft);
}

/* Inputs */
.input, .select, .textarea {
  width: 100%;
  background: #fff;
  border: 1px solid #e2e8f0;
  color: var(--text);
  padding: 12px 14px;
  border-radius: 12px;
  outline: none;
  transition: border .15s ease, box-shadow .15s ease;
}
.input:focus, .select:focus, .textarea:focus {
  border-color: var(--brand);
  box-shadow: 0 0 0 4px var(--ring);
}
.label { color:#334155; font-size:.92rem; margin-bottom:6px; display:block; font-weight:600; }

/* Buttons */
.btn {
  display:inline-flex; align-items:center; justify-content:center;
  gap:8px; padding: 10px 16px; border-radius: 12px;
  border: 1px solid transparent;
  background: var(--brand);
  color: #fff; cursor: pointer; transition: transform .05s ease, filter .2s ease, box-shadow .15s ease;
  box-shadow: 0 6px 18px rgba(37,99,235,.25);
}
.btn:hover { filter: brightness(1.04); }
.btn:active { transform: translateY(1px) }
.btn.secondary { background: #e2e8f0; color:#111827; box-shadow:none }
.btn.ghost { background: transparent; border-color: #e2e8f0; color:#111827; box-shadow:none }
.btn.success { background: var(--positive); box-shadow: 0 6px 18px rgba(22,163,74,.2) }
.btn.warn { background: var(--warning); box-shadow: 0 6px 18px rgba(217,119,6,.2) }
.btn.danger { background: var(--danger); box-shadow: 0 6px 18px rgba(220,38,38,.2) }

/* Badges */
.badge {
  display:inline-flex; align-items:center; gap:6px;
  padding: 6px 10px; border-radius: 999px;
  background: #eef2ff;
  border: 1px solid #dbeafe;
  font-size: .83rem; color:#1e3a8a
}
.badge.success { background: #ecfdf5; border-color:#d1fae5; color:#065f46 }
.badge.warn { background: #fffbeb; border-color:#fde68a; color:#92400e }
.badge.danger { background: #fef2f2; border-color:#fecaca; color:#991b1b }
.badge.info { background: #eff6ff; border-color:#bfdbfe; color:#1e40af }

/* Tables */
.table { width: 100%; border-collapse: collapse; font-size: .95rem; }
.table th, .table td { padding: 12px 10px; border-bottom: 1px solid var(--border); }
.table th { text-align: left; color:#334155; font-weight:700 }
.table tr:hover td { background: #f8fafc }

/* Kbd */
.kbd { padding:2px 6px; border-radius:6px; border:1px solid #e2e8f0; background:#fff; color:#0f172a; font-size:.82rem }