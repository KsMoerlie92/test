// ── Container definitions, pickContainer, expandToUnits ─────────────────────
// ── Shared JS (legplan + CIPL functions) ─────────────────────────────────
const CONTAINERS = {
  '20ft':    { label:"20' Std",  len:589,  w:234, h:239, cbm:33, maxKg:28000 },
  '40ft':    { label:"40' Std",  len:1203, w:234, h:239, cbm:67, maxKg:26500 },
  '40ft HC': { label:"40' HC",   len:1203, w:234, h:269, cbm:76, maxKg:26330 },
};
const COLLO_COLORS = ['#2196F3','#E53935','#43A047','#FB8C00','#8E24AA',
  '#00ACC1','#D81B60','#6D4C41','#3949AB','#00897B','#F4511E','#7CB342',
  '#039BE5','#EF9A9A','#80DEEA'];

function pickContainer(totVol, totKg, maxL) {
  for (const key of ['20ft','40ft','40ft HC']) {
    const c = CONTAINERS[key];
    if (totVol <= c.cbm*0.92 && totKg <= c.maxKg && maxL <= c.len) return key;
  }
  return '40ft HC';  // max beschikbare container
}

function expandToUnits(rows) {
  // Per row with L+B+H filled: compute vol m3 = L*B*H / 1 000 000
  const units = [];
  rows.forEach((r, ri) => {
    const l = parseFloat(r.colLength) || 0;
    const w = parseFloat(r.colWidth)  || 0;
    const h = parseFloat(r.colHeight) || 0;
    if (l > 0 && w > 0 && h > 0) {
      units.push({ ...r, _l:l, _w:w, _h:h,
        _vol: +(l * w * h / 1e6).toFixed(4), _ri:ri });
    }
  });
  return units;
}

