// ── Canvas drawing: drawSideView, drawTopView, drawVisualization ────────────
// ── Side view — length × height with correct Z stacking ──────────────────
function drawSideView(placed, container) {
  const C  = CONTAINERS[container];
  const cv = document.getElementById('lp-canvas-side');
  if (!cv) return;
  const PAD = 14, W = cv.width;
  const scX = (W - PAD*2) / C.len, scZ = scX;
  cv.height = Math.max(Math.ceil(C.h * scZ + PAD*2 + 16), 80);
  const ctx = cv.getContext('2d');
  ctx.clearRect(0, 0, W, cv.height);
  const bx0 = PAD, by0 = PAD, cW = C.len*scX, cH = C.h*scZ;

  ctx.strokeStyle = '#00B4D8'; ctx.lineWidth = 2;
  ctx.strokeRect(bx0, by0, cW, cH);
  ctx.strokeStyle = 'rgba(0,180,216,0.25)'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(bx0, by0+cH); ctx.lineTo(bx0+cW, by0+cH); ctx.stroke();

  // Draw floor items first, stacked items on top
  [...placed].sort((a,b) => a.pz - b.pz).forEach(item => {
    const col = COLLO_COLORS[item._ri % COLLO_COLORS.length];
    const bx = bx0 + item.px * scX;
    const bw = Math.max(item.pl * scX - 1, 2);
    const bh = Math.max(item.ph * scZ - 1, 2);
    const by = by0 + cH - (item.pz + item.ph) * scZ; // correct Z offset

    ctx.globalAlpha = 0.88; ctx.fillStyle = col;
    ctx.fillRect(bx, by, bw, bh);
    ctx.globalAlpha = 1;
    ctx.strokeStyle = 'rgba(0,0,0,0.4)'; ctx.lineWidth = 0.8;
    ctx.strokeRect(bx, by, bw, bh);

    // Dashed top line indicates stacked item
    if (item.pz > 0.5) {
      ctx.strokeStyle = 'rgba(255,255,255,0.55)'; ctx.lineWidth = 1.2;
      ctx.setLineDash([3,2]);
      ctx.beginPath(); ctx.moveTo(bx, by); ctx.lineTo(bx+bw, by); ctx.stroke();
      ctx.setLineDash([]);
    }
    if (bw > 10) {
      ctx.fillStyle = '#fff';
      ctx.font = `bold ${Math.min(9, bw/2.5)}px monospace`;
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.globalAlpha = 0.92;
      ctx.fillText(item.colCollo || String(item._ri+1), bx+bw/2, by+bh/2);
      ctx.globalAlpha = 1;
    }
  });

  ctx.fillStyle = 'rgba(0,180,216,0.75)'; ctx.font = '9px monospace';
  ctx.textAlign = 'left';  ctx.textBaseline = 'top';
  ctx.fillText(`L ${C.len} cm`, bx0+2, by0+2);
  ctx.textAlign = 'right';
  ctx.fillText(`H ${C.h} cm`, bx0+cW-2, by0+2);
}

// ── Top view — floor plan with stacked items drawn on top ─────────────────
function drawTopView(placed, container) {
  const C  = CONTAINERS[container];
  const cv = document.getElementById('lp-canvas-top');
  if (!cv) return;
  const PAD = 14, W = cv.width;
  const scX = (W - PAD*2) / C.len, scY = scX;
  cv.height = Math.max(Math.ceil(C.w * scY + PAD*2 + 16), 60);
  const ctx = cv.getContext('2d');
  ctx.clearRect(0, 0, W, cv.height);
  const bx0 = PAD, by0 = PAD, cW = C.len*scX, cH = C.w*scY;

  ctx.strokeStyle = '#00B4D8'; ctx.lineWidth = 2;
  ctx.strokeRect(bx0, by0, cW, cH);
  ctx.strokeStyle = 'rgba(0,180,216,0.10)'; ctx.lineWidth = 0.5;
  for (let x = 200; x < C.len; x += 200) {
    ctx.beginPath(); ctx.moveTo(bx0+x*scX, by0); ctx.lineTo(bx0+x*scX, by0+cH); ctx.stroke();
  }

  // Sort by Z: floor items first → stacked items on top visually
  [...placed].sort((a,b) => a.pz - b.pz).forEach(item => {
    const col = COLLO_COLORS[item._ri % COLLO_COLORS.length];
    const bx = bx0 + item.px * scX;
    const by = by0 + item.py * scY;
    const bw = Math.max(item.pl * scX - 1, 2);
    const bh = Math.max(item.pw * scY - 1, 2);
    const stacked = item.pz > 0.5;

    ctx.globalAlpha = stacked ? 0.72 : 0.88;
    ctx.fillStyle = col;
    ctx.fillRect(bx, by, bw, bh);
    ctx.globalAlpha = 1;
    ctx.strokeStyle = stacked ? 'rgba(255,255,255,0.55)' : 'rgba(0,0,0,0.3)';
    ctx.lineWidth   = stacked ? 1.2 : 0.6;
    ctx.strokeRect(bx, by, bw, bh);

    if (bw > 10) {
      ctx.fillStyle = '#fff';
      ctx.font = `bold ${Math.min(9, bw/2.5)}px monospace`;
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.globalAlpha = 0.92;
      ctx.fillText(item.colCollo || String(item._ri+1), bx+bw/2, by+bh/2);
      ctx.globalAlpha = 1;
    }
  });

  ctx.fillStyle = 'rgba(0,180,216,0.75)'; ctx.font = '9px monospace';
  ctx.textAlign = 'left';  ctx.textBaseline = 'top';
  ctx.fillText(`L ${C.len} cm`, bx0+2, by0+2);
  ctx.textAlign = 'right';
  ctx.fillText(`B ${C.w} cm`, bx0+cW-2, by0+2);
}

function drawVisualization(placed, container) {
  drawSideView(placed, container);
  drawTopView(placed, container);
}


function buildLegend(placed) {
  const el = document.getElementById('lp-legend');
  if (!el) return;
  el.innerHTML = placed.map((item, i) => {
    const col = COLLO_COLORS[item._ri % COLLO_COLORS.length];
    return `<span class="lp-legend-item">
      <span class="lp-legend-swatch" style="background:${col}"></span>
      ${esc(item.colCollo||String(i+1))}
    </span>`;
  }).join('');
}

function buildCargoTable(placed) {
  const tbody = document.getElementById('lp-cargo-tbody');
  if (!tbody) return;
  tbody.innerHTML = placed.map((item, i) => {
    const col = COLLO_COLORS[item._ri % COLLO_COLORS.length];
    const sup6 = (item.colSupplier||'').slice(0,6) || '—';
    const vol = (item._l>0&&item._w>0&&item._h>0)
      ? (item._l*item._w*item._h/1000000).toFixed(3)
      : parseFloat(item.colVolume||0).toFixed(3);
    return `<tr>
      <td><span class="lp-cargo-swatch" style="background:${col}"></span></td>
      <td>${esc(item.colCollo||String(i+1))}</td>
      <td style="font-size:.6rem;color:var(--muted)">${esc(sup6)}</td>
      <td>${esc(item.colPackaging||'—')}</td>
      <td class="num">${item.pl.toFixed(0)}</td>
      <td class="num">${item.pw.toFixed(0)}</td>
      <td class="num">${item.ph.toFixed(0)}</td>
      <td class="num">${esc(item.colGrossWeight||'—')}</td>
    </tr>`;
  }).join('');
}

function updateLegplanStats(placed, containerKey) {
  const c = CONTAINERS[containerKey];
  const totVol = placed.reduce((s,r) => s + (r._l>0&&r._w>0&&r._h>0 ? r._l*r._w*r._h/1e6 : parseFloat(r.colVolume||0)), 0);
  const totKg  = placed.reduce((s,r) => s + (parseFloat(r.colGrossWeight||0)||0), 0);
  const usedLen = placed.length ? Math.max(...placed.map(p => p.px + p.pl)) : 0;
  const volP = Math.min(100, totVol / c.cbm * 100);
  const wgtP = Math.min(100, totKg  / c.maxKg * 100);
  const lenP = Math.min(100, usedLen / c.len * 100);
  const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
  const sw=(id,w,bg)=>{const el=document.getElementById(id);if(el){el.style.width=w;if(bg)el.style.background=bg;}};
  sw('lp-bar-vol', volP.toFixed(1)+'%', volP>90?'#ef4444':volP>70?'#f59e0b':'#00B4D8');
  sw('lp-bar-wgt', wgtP.toFixed(1)+'%', wgtP>90?'#ef4444':wgtP>70?'#f59e0b':'#D91F2C');
  sw('lp-bar-len', lenP.toFixed(1)+'%', '#f59e0b');
  set('lp-vol-used', totVol.toFixed(2)); set('lp-vol-max', c.cbm); set('lp-vol-pct', volP.toFixed(1)+'%');
  set('lp-wgt-used', totKg.toLocaleString('nl-NL')); set('lp-wgt-max', c.maxKg.toLocaleString('nl-NL')); set('lp-wgt-pct', wgtP.toFixed(1)+'%');
  set('lp-len-used', usedLen.toFixed(0)); set('lp-len-max', c.len); set('lp-len-pct', lenP.toFixed(1)+'%');
}

