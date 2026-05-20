// ── 3D bin packing: isStackable, packBin, packBinMulti ─────────────────────
// ── 3D stacking helpers ───────────────────────────────────────────────────
function isStackable(item) {
  const pkg = (item.colPackaging || '').toLowerCase();
  return pkg.includes('case') && !pkg.includes('pallet');
}

function hasOverlap3D(cx, cy, cz, ol, ow, oh, placed) {
  for (const p of placed) {
    if (cx < p.px+p.pl-0.1 && cx+ol > p.px+0.1 &&
        cy < p.py+p.pw-0.1 && cy+ow > p.py+0.1 &&
        cz < p.pz+p.ph-0.1 && cz+oh > p.pz+0.1) return true;
  }
  return false;
}

function canStack(cx, cy, cz, ol, ow, itemWeight, placed) {
  if (cz < 0.5) return true; // floor — always OK
  const supports = placed.filter(p => {
    const tz = p.pz + p.ph;
    return Math.abs(tz - cz) < 0.5 &&
      p.px < cx+ol-0.1 && p.px+p.pl > cx+0.1 &&
      p.py < cy+ow-0.1 && p.py+p.pw > cy+0.1;
  });
  if (!supports.length) return false;
  if (!supports.every(s => isStackable(s))) return false;
  const minW = Math.min(...supports.map(s => parseFloat(s.colGrossWeight)||0));
  return itemWeight <= minW;
}

// ── 3D Extreme-Point bin packer with stacking ─────────────────────────────
function packBin(units, containerKey) {
  const C = CONTAINERS[containerKey];

  // Heaviest first → they go to the floor; stacking only allows lighter on top
  const sorted = [...units].sort((a, b) => {
    const wa = parseFloat(a.colGrossWeight)||0;
    const wb = parseFloat(b.colGrossWeight)||0;
    if (Math.abs(wb - wa) > 0.01) return wb - wa;
    return (b._l*b._w*b._h) - (a._l*a._w*a._h);
  });

  const placed = [];
  let eps = [{x:0, y:0, z:0}];

  const addEP = (x, y, z) => {
    if (x > C.len+1 || y > C.w+1 || z > C.h+1) return;
    eps.push({x, y, z});
  };

  for (const item of sorted) {
    let best = null, bestScore = Infinity;
    const iw = parseFloat(item.colGrossWeight)||0;

    for (const [ol, ow] of [[item._l, item._w], [item._w, item._l]]) {
      if (ol > C.len+0.5 || ow > C.w+0.5 || item._h > C.h+0.5) continue;

      for (const {x:cx, y:cy, z:cz} of eps) {
        if (cx+ol > C.len+0.5 || cy+ow > C.w+0.5 || cz+item._h > C.h+0.5) continue;
        if (hasOverlap3D(cx, cy, cz, ol, ow, item._h, placed)) continue;
        if (!canStack(cx, cy, cz, ol, ow, iw, placed)) continue;
        // Prefer: lowest Z → lowest Y → lowest X
        const score = cz*1e9 + cy*1e5 + cx;
        if (score < bestScore) { bestScore = score; best = {cx, cy, cz, ol, ow}; }
      }
    }

    if (best) {
      const {cx, cy, cz, ol, ow} = best;
      placed.push({...item, px:cx, py:cy, pz:cz, pl:ol, pw:ow, ph:item._h});
      addEP(cx+ol, cy,   cz);
      addEP(cx,   cy+ow, cz);
      addEP(cx,   cy,    cz+item._h);
      // Deduplicate EPs
      eps = [...new Map(eps.map(e=>[`${e.x},${e.y},${e.z}`,e])).values()];
    }
  }
  return placed;
}

// ── Multi-container packer: keeps filling containers until all units placed ─
function packBinMulti(units, containerKey) {
  const plan = [];
  let remaining = [...units];
  const MAX_CONTAINERS = 20;

  while (remaining.length > 0 && plan.length < MAX_CONTAINERS) {
    const placed = packBin(remaining, containerKey);
    if (placed.length === 0) break; // nothing fits at all — stop

    plan.push({ placed, containerKey });

    const placedRis = new Set(placed.map(p => p._ri));
    const next = remaining.filter(u => !placedRis.has(u._ri));
    if (next.length === remaining.length) break; // no progress guard
    remaining = next;
  }
  return plan; // [{placed:[...], containerKey:'40ft HC'}, ...]
}

