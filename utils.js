// ── Utilities: constants, esc(), setStatus() ───────────────────────────────
// ══════════════════════════════════════════════════════════════════════════
//  IHC PO Matcher
//
//  Match-richting (nieuw):
//    1. Moederlijst → ParseItemCodes (VBA-logica) → Expanded Output
//       • Kolom A = IHC PO
//       • Kolom B = Item (expanded code)
//       • Kolom C = Combined  (PO + "-" + Item)   ← dit is de lookup waarde
//
//    2. Expediting Lijst Kolom A  ← zowel lookup array als return array
//
//    3. XLOOKUP:
//       =XLOOKUP( [Expanded Output Kol C], [Expediting Kol A], [Expediting Kol A] )
//
//    4. Resultaat = waarden die in BEIDE bestanden voorkomen
// ══════════════════════════════════════════════════════════════════════════


// ── Helpers ───────────────────────────────────────────────────────────────
function esc(s) {
  return String(s || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function setStatus(msg, isError = false) {
  const el = document.getElementById('status-msg');
  el.textContent = msg;
  el.style.color = isError ? 'var(--red)' : 'var(--muted)';
}
