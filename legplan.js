// ── Legplan: containers, bin-packing, canvas, AI advice, PDF export ─────────
// ── Legplan containers + visualization ───────────────────────────────────
const CONTAINERS = {
  '20ft':    { label:"20' Std",  len:589,  w:234, h:239, cbm:33, maxKg:28000 },
  '40ft':    { label:"40' Std",  len:1203, w:234, h:239, cbm:67, maxKg:26500 },
  '40ft HC': { label:"40' HC",   len:1203, w:234, h:269, cbm:76, maxKg:26330 },
  '45ft HC': { label:"45' HC",   len:1355, w:234, h:269, cbm:86, maxKg:27000 },
};
const COLLO_COLORS = ['#2196F3','#E53935','#43A047','#FB8C00','#8E24AA',
  '#00ACC1','#D81B60','#6D4C41','#3949AB','#00897B','#F4511E','#7CB342',
  '#039BE5','#EF9A9A','#80DEEA'];

function pickContainer(totVol, totKg, maxL) {
  for (const key of ['20ft','40ft','40ft HC']) {  // 45ft HC excluded per policy
    const c = CONTAINERS[key];
    if (totVol <= c.cbm*0.92 && totKg <= c.maxKg && maxL <= c.len) return key;
  }
  return '45ft HC';
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

// ── Greedy row packer — tries both orientations, skips (never breaks) ─────
function packBin(units, container) {
  const C = CONTAINERS[container];
  // Sort: heaviest first, then largest footprint
  const sorted = [...units].sort((a, b) =>
    (parseFloat(b.colGrossWeight)||0) - (parseFloat(a.colGrossWeight)||0) ||
    (b._l * b._w) - (a._l * a._w)
  );

  const placed = [];
  // Rows along container width (Y); within each row, items placed along length (X)
  let rows = [{ y: 0, x: 0, rowW: 0 }];  // rowW = deepest width used in this row

  for (const item of sorted) {
    let bestRow = -1, usedL = 0, usedW = 0;

    // Try both orientations: normal (l, w) and rotated (w, l)
    for (const [ol, ow] of [[item._l, item._w], [item._w, item._l]]) {
      if (ol > C.len + 0.5 || ow > C.w + 0.5) continue; // too big in any orientation

      for (let ri = 0; ri < rows.length; ri++) {
        const row = rows[ri];
        if (row.x + ol <= C.len + 0.5 && row.y + ow <= C.w + 0.5) {
          bestRow = ri; usedL = ol; usedW = ow;
          break;
        }
      }
      if (bestRow >= 0) break;
    }

    if (bestRow < 0) {
      // Try opening a new row below existing rows
      const nextY = rows.reduce((mx, r) => Math.max(mx, r.y + r.rowW), 0);
      for (const [ol, ow] of [[item._l, item._w], [item._w, item._l]]) {
        if (ol > C.len + 0.5 || nextY + ow <= C.w + 0.5) {
          if (ol <= C.len + 0.5) {
            rows.push({ y: nextY, x: 0, rowW: 0 });
            bestRow = rows.length - 1;
            usedL = ol; usedW = ow;
            break;
          }
        }
      }
    }

    if (bestRow < 0) continue; // skip — doesn't fit in any orientation

    const row = rows[bestRow];
    placed.push({
      ...item,
      px: row.x,
      py: row.y,
      pz: 0,        // floor level
      pl: usedL,
      pw: usedW,
      ph: item._h,
    });
    row.x += usedL;
    if (usedW > row.rowW) row.rowW = usedW;
  }
  return placed;
}

// ── 2D Side view (length × height, floor-anchored) ───────────────────────
function drawSideView(placed, container) {
  const C  = CONTAINERS[container];
  const cv = document.getElementById('lp-canvas-side');
  if (!cv) return;
  const PAD = 14, W = cv.width;

  const scX  = (W - PAD * 2) / C.len;
  const scZ  = scX;
  cv.height  = Math.max(Math.ceil(C.h * scZ + PAD * 2 + 16), 80);
  const H    = cv.height;

  const ctx  = cv.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  const bx0 = PAD, by0 = PAD;
  const cW  = C.len * scX, cH = C.h * scZ;

  // Container outline
  ctx.strokeStyle = '#00B4D8'; ctx.lineWidth = 2;
  ctx.strokeRect(bx0, by0, cW, cH);
  // Floor line
  ctx.strokeStyle = 'rgba(0,180,216,0.25)'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(bx0, by0 + cH); ctx.lineTo(bx0 + cW, by0 + cH); ctx.stroke();

  // Boxes — floor-anchored (by = top of container area minus box height)
  placed.forEach((item, i) => {
    const col = COLLO_COLORS[item._ri % COLLO_COLORS.length];
    const bx  = bx0 + item.px * scX;
    const bw  = Math.max(item.pl * scX - 1, 2);
    const bh  = Math.max(item.ph * scZ - 1, 2);
    const by  = by0 + cH - bh;

    ctx.globalAlpha = 0.85; ctx.fillStyle = col;
    ctx.fillRect(bx, by, bw, bh);
    ctx.globalAlpha = 1;
    ctx.strokeStyle = 'rgba(0,0,0,0.35)'; ctx.lineWidth = 0.7;
    ctx.strokeRect(bx, by, bw, bh);

    if (bw > 10) {
      ctx.fillStyle = '#fff';
      ctx.font = `bold ${Math.min(9, bw / 2.5)}px monospace`;
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.globalAlpha = 0.92;
      ctx.fillText(item.colCollo || String(i + 1), bx + bw / 2, by + bh / 2);
      ctx.globalAlpha = 1;
    }
  });

  // Labels
  ctx.fillStyle = 'rgba(0,180,216,0.75)'; ctx.font = '9px monospace';
  ctx.textAlign = 'left';  ctx.textBaseline = 'top';
  ctx.fillText(`L ${C.len} cm`, bx0 + 2, by0 + 2);
  ctx.textAlign = 'right';
  ctx.fillText(`H ${C.h} cm`, bx0 + cW - 2, by0 + 2);
}

// ── 2D Top view (length × width, plan view) ───────────────────────────────
function drawTopView(placed, container) {
  const C  = CONTAINERS[container];
  const cv = document.getElementById('lp-canvas-top');
  if (!cv) return;
  const PAD = 14, W = cv.width;

  const scX  = (W - PAD * 2) / C.len;
  const scY  = scX;
  cv.height  = Math.max(Math.ceil(C.w * scY + PAD * 2 + 16), 60);
  const H    = cv.height;

  const ctx  = cv.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  const bx0 = PAD, by0 = PAD;
  const cW  = C.len * scX, cH = C.w * scY;

  // Container outline
  ctx.strokeStyle = '#00B4D8'; ctx.lineWidth = 2;
  ctx.strokeRect(bx0, by0, cW, cH);

  // Grid lines every 200 cm
  ctx.strokeStyle = 'rgba(0,180,216,0.10)'; ctx.lineWidth = 0.5;
  for (let x = 200; x < C.len; x += 200) {
    ctx.beginPath();
    ctx.moveTo(bx0 + x * scX, by0);
    ctx.lineTo(bx0 + x * scX, by0 + cH);
    ctx.stroke();
  }

  // Boxes
  placed.forEach((item, i) => {
    const col = COLLO_COLORS[item._ri % COLLO_COLORS.length];
    const bx  = bx0 + item.px * scX;
    const by  = by0 + item.py * scY;
    const bw  = Math.max(item.pl * scX - 1, 2);
    const bh  = Math.max(item.pw * scY - 1, 2);

    ctx.globalAlpha = 0.85; ctx.fillStyle = col;
    ctx.fillRect(bx, by, bw, bh);
    ctx.globalAlpha = 1;
    ctx.strokeStyle = 'rgba(0,0,0,0.3)'; ctx.lineWidth = 0.6;
    ctx.strokeRect(bx, by, bw, bh);

    if (bw > 10) {
      ctx.fillStyle = '#fff';
      ctx.font = `bold ${Math.min(9, bw / 2.5)}px monospace`;
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.globalAlpha = 0.92;
      ctx.fillText(item.colCollo || String(i + 1), bx + bw / 2, by + bh / 2);
      ctx.globalAlpha = 1;
    }
  });

  // Labels
  ctx.fillStyle = 'rgba(0,180,216,0.75)'; ctx.font = '9px monospace';
  ctx.textAlign = 'left';  ctx.textBaseline = 'top';
  ctx.fillText(`L ${C.len} cm`, bx0 + 2, by0 + 2);
  ctx.textAlign = 'right';
  ctx.fillText(`B ${C.w} cm`, bx0 + cW - 2, by0 + 2);
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

function updateLegplanStats(unique, container, totVol, totKg) {
  const c=CONTAINERS[container];
  const volP=Math.min(100,totVol/c.cbm*100);
  const wgtP=Math.min(100,totKg/c.maxKg*100);
  const usedLen=unique.reduce((s,r)=>s+(parseFloat(r.colLength)||0),0);
  const lenP=Math.min(100,usedLen/c.len*100);
  const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
  const sw=(id,w,bg)=>{const el=document.getElementById(id);if(el){el.style.width=w;if(bg)el.style.background=bg;}};
  const volBg = volP>90?'#ef4444':volP>70?'#f59e0b':'#00B4D8';
  const wgtBg = wgtP>90?'#ef4444':wgtP>70?'#f59e0b':'#D91F2C';
  sw('lp-bar-vol',volP.toFixed(1)+'%',volBg);
  sw('lp-bar-wgt',wgtP.toFixed(1)+'%',wgtBg);
  sw('lp-bar-len',lenP.toFixed(1)+'%','#f59e0b');
  set('lp-vol-used',totVol.toFixed(2)); set('lp-vol-max',c.cbm); set('lp-vol-pct',volP.toFixed(1)+'%');
  set('lp-wgt-used',totKg.toLocaleString('nl-NL')); set('lp-wgt-max',c.maxKg.toLocaleString('nl-NL')); set('lp-wgt-pct',wgtP.toFixed(1)+'%');
  set('lp-len-used',usedLen.toFixed(0)); set('lp-len-max',c.len); set('lp-len-pct',lenP.toFixed(1)+'%');
}

function getLegplanRows() {
  const inclNRY = document.getElementById('toggle-nry')?.checked ?? false;
  return allRows.filter(r => {
    if (!r.noMatch) return false;
    const alBlank = !r.colAL || r.colAL.trim() === '';
    if (!alBlank) return false;
    // Kol Z (colZ) = received indicator; "not received yet" or empty = nog niet ontvangen
    const zVal = (r.colZ||'').trim().toLowerCase();
    const isNRY = zVal === 'not received yet' || zVal === '';
    if (isNRY && !inclNRY) return false;   // uitsluiten tenzij toggle aan
    return true;
  });
}

function renderLegplan() {
  const rows = getLegplanRows();
  const tbody = document.getElementById('legplan-tbody');
  const statsEl = document.getElementById('legplan-stats');
  if(tbody) tbody.innerHTML='';
  if(statsEl) statsEl.textContent='';
  if(!rows.length){
    if(tbody) tbody.innerHTML='<tr><td colspan="12" style="text-align:center;color:var(--muted);padding:1.5rem">Geen regels gevonden</td></tr>';
    return;
  }
  const seenC=new Set();
  const unique=rows.filter(r=>{
    const k=(r.colCollo||r.combined||'').toLowerCase();
    if(seenC.has(k))return false;seenC.add(k);return true;
  });
  const frag=document.createDocumentFragment();
  unique.forEach((r,i)=>{
    const zVal = (r.colZ||'').trim().toLowerCase();
    const isNRY = zVal === 'not received yet' || zVal === '';
    const badge = isNRY
      ? '<span class="badge-nry">⏳ Verwacht</span>'
      : '<span class="badge-confirmed">✓ Bevestigd</span>';
    const tr=document.createElement('tr');
    if(isNRY) tr.className='lp-row-nry';
    tr.innerHTML=`<td>${i+1}</td><td class="cell-supplier">${esc(r.colSupplier||'—')}</td>
      <td>${esc(r.colE||'—')}</td><td class="cell-collo">${esc(r.colCollo||'—')}</td>
      <td>${esc(r.colPackaging||'—')}</td>
      <td style="text-align:right">${esc(r.colLength||'—')}</td>
      <td style="text-align:right">${esc(r.colWidth||'—')}</td>
      <td style="text-align:right">${esc(r.colHeight||'—')}</td>
      <td style="text-align:right">${parseFloat(r.colVolume||0).toFixed(2)}</td>
      <td style="text-align:right;font-weight:600">${esc(r.colGrossWeight||'—')}</td>
      <td style="text-align:right">${esc(r.colNettWeight||'—')}</td>
      <td>${badge}</td>`;
    frag.appendChild(tr);
  });
  if(tbody) tbody.appendChild(frag);
  const nConfirmed = unique.filter(r=>{ const z=(r.colZ||'').trim().toLowerCase(); return z!=='not received yet'&&z!==''; }).length;
  const nExpected  = unique.length - nConfirmed;
  let totVol=0,totGross=0,totNett=0;
  unique.forEach(r=>{
    const l=parseFloat(r.colLength)||0,w=parseFloat(r.colWidth)||0,h=parseFloat(r.colHeight)||0;
    totVol+=(l>0&&w>0&&h>0)?l*w*h/1000000:(parseFloat(r.colVolume)||0);
    totGross+=parseFloat(r.colGrossWeight||0)||0;
    totNett+=parseFloat(r.colNettWeight||0)||0;
  });
  if(statsEl) statsEl.innerHTML=
    `<span>${unique.length} collo's</span>`+
    (nConfirmed ? `<span style="color:var(--green)">✓ ${nConfirmed} bevestigd</span>` : '')+
    (nExpected  ? `<span style="color:var(--amber)">⏳ ${nExpected} verwacht</span>`  : '')+
    `<span>Vol: <strong>${totVol.toFixed(2)} m³</strong></span>`+
    `<span>Bruto: <strong>${totGross.toLocaleString('nl-NL')} kg</strong></span>`+
    `<span>Netto: <strong>${totNett.toLocaleString('nl-NL')} kg</strong></span>`;
}

async function runLegplanAI() {
  const rows=getLegplanRows();
  if(!rows.length){alert('Geen legplan regels beschikbaar.');return;}
  const btn=document.getElementById('btn-legplan-ai');
  const resultEl=document.getElementById('legplan-ai-result');
  btn.disabled=true; btn.textContent='⏳ Berekenen…';

  const seenC=new Set();
  const unique=rows.filter(r=>{
    const k=(r.colCollo||r.combined||'').toLowerCase();
    if(seenC.has(k))return false;seenC.add(k);return true;
  });

  // Expand to individual units with dimensions
  const units = expandToUnits(unique);
  let totVol=0,totKg=0,maxDim=0;
  unique.forEach(r=>{
    const l=parseFloat(r.colLength)||0, w=parseFloat(r.colWidth)||0, h=parseFloat(r.colHeight)||0;
    totVol += (l>0&&w>0&&h>0) ? l*w*h/1000000 : (parseFloat(r.colVolume)||0);
    totKg+=parseFloat(r.colGrossWeight||0)||0;
    [l,w,h].forEach(v=>{if(v>maxDim)maxDim=v;});
  });

  // Pick container from units with actual dimensions
  const container = pickContainer(totVol, totKg, maxDim);

  // 3-D bin-pack using individual units
  const placed = packBin(units, container);

  // Show result panel
  resultEl.classList.add('visible');
  const ctEl=document.getElementById('lp-container-type');
  if(ctEl)ctEl.textContent=CONTAINERS[container].label + ` · ${placed.length}/${units.length} geplaatst`;

  // Canvas is sized inside drawVisualization based on exact projected dimensions
  setTimeout(()=>drawVisualization(placed, container), 60);
  buildLegend(placed);
  buildCargoTable(placed);
  updateLegplanStats(unique, container, totVol, totKg);

  // Update stats badge
  if(placed.length < units.length){
    const warn=document.getElementById('lp-container-type');
    if(warn){ warn.style.background='#f59e0b'; warn.style.color='#000'; }
  }

  // ── Claude API text advice ──────────────────────────────────────────────
  const apiKey=document.getElementById('legplan-apikey').value.trim();
  const adviceEl=document.getElementById('lp-advice-text');
  if(!apiKey){
    if(adviceEl)adviceEl.innerHTML='<em style="color:var(--muted)">Vul een API key in voor AI-advies.</em>';
    btn.disabled=false; btn.textContent='🤖 Container Advies'; return;
  }
  if(adviceEl)adviceEl.innerHTML='<em style="color:var(--muted)">AI analyseert…</em>';
  btn.textContent='⏳ AI…';
  const c=CONTAINERS[container];

  // Build placement summary for AI
  const placedLines = placed.map((item,i) =>
    `Collo ${item.colCollo||i+1}: ${item.pl.toFixed(0)}×${item.pw.toFixed(0)}×${item.ph.toFixed(0)}cm → positie X${item.px.toFixed(0)}/Y${item.py.toFixed(0)}/Z${item.pz.toFixed(0)} · ${item.colGrossWeight||'?'}kg`
  ).join('\n');
  const unplacedLines = units.filter(u=>!placed.includes(u)).map(u=>
    `Collo ${u.colCollo||'?'}: ${u._l}×${u._w}×${u._h}cm · ${u.colGrossWeight||'?'}kg`
  ).join('\n');

  const prompt=`Logistiek expert. Beknopt legplan advies (max 150 woorden):\n\nContainer: ${CONTAINERS[container].label} (${c.len}×${c.w}×${c.h}cm, ${c.cbm}CBM, ${c.maxKg}kg max)\nBezetting: ${(totVol/c.cbm*100).toFixed(1)}% vol · ${(totKg/c.maxKg*100).toFixed(1)}% gewicht\nGeplaatst: ${placed.length} van ${units.length} collo's\n\nPlaatsing (X=lengterichting, Y=breedterichting, Z=hoogte):\n${placedLines}${unplacedLines?'\n\nNiet geplaatst:\n'+unplacedLines:''}\n\nFocus: laadvolgorde, gewichtsverdeling, aandachtspunten (hijsplan bij >5t, ADR etc.). Antwoord in het Nederlands.`;

  try{
    const resp=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},body:JSON.stringify({model:'claude-sonnet-4-20250514',max_tokens:1000,messages:[{role:'user',content:prompt}]})});
    if(!resp.ok){const eb=await resp.json().catch(()=>({}));throw new Error(`HTTP ${resp.status}: ${eb?.error?.message||resp.statusText}`);}
    const data=await resp.json();
    const text=(data.content||[]).filter(b=>b.type==='text').map(b=>b.text).join('\n');
    if(adviceEl)adviceEl.innerHTML=text.split('\n').filter(l=>l.trim()).map(l=>`<p>${esc(l)}</p>`).join('');
  }catch(err){if(adviceEl)adviceEl.innerHTML=`<em style="color:var(--ihc-red)">Fout: ${esc(String(err))}</em>`;}
  btn.disabled=false; btn.textContent='🤖 Container Advies';
}


// ── PDF Export: legplan with container drawings ──────────────────────────
async function exportLegplanPDF() {
  const resultEl = document.getElementById('legplan-ai-result');
  if (!resultEl || !resultEl.classList.contains('visible')) {
    alert('Genereer eerst een Container Advies voordat je exporteert.');
    return;
  }
  const btn = document.getElementById('btn-legplan-pdf');
  btn.disabled = true; btn.textContent = '⏳ PDF maken…';

  try {
    // ── Build PDF as a printable HTML page, then trigger print-to-PDF ──
    const containerType = document.getElementById('lp-container-type')?.textContent || '—';

    // Capture both canvases as PNG data-URLs
    const cvS = document.getElementById('lp-canvas-side');
    const cvT = document.getElementById('lp-canvas-top');
    const imgS = cvS ? cvS.toDataURL('image/png') : '';
    const imgT = cvT ? cvT.toDataURL('image/png') : '';

    // Capture legend HTML
    const legendHTML = document.getElementById('lp-legend')?.innerHTML || '';

    // Capture stats
    const stats = ['lp-vol-used','lp-vol-max','lp-vol-pct',
                   'lp-wgt-used','lp-wgt-max','lp-wgt-pct',
                   'lp-len-used','lp-len-max','lp-len-pct']
      .reduce((o,id)=>{o[id]=document.getElementById(id)?.textContent||'—';return o;},{});

    // Capture cargo table rows
    const cargoRows = document.getElementById('lp-cargo-tbody')?.innerHTML || '';

    // Capture AI advice
    const adviceHTML = document.getElementById('lp-advice-text')?.innerHTML || '';

    // Capture collo tabel from main legplan table
    const legplanRows = getLegplanRows();
    const seenC = new Set();
    const unique = legplanRows.filter(r=>{
      const k=(r.colCollo||r.combined||'').toLowerCase();
      if(seenC.has(k))return false;seenC.add(k);return true;
    });

    const colloTableRows = unique.map((r,i) => {
      const col = COLLO_COLORS[i % COLLO_COLORS.length];
      return `<tr>
        <td><span style="display:inline-block;width:12px;height:12px;background:${col};border-radius:2px;vertical-align:middle"></span></td>
        <td>${esc(r.colCollo||String(i+1))}</td>
        <td>${esc(r.colSupplier||'—')}</td>
        <td>${esc(r.colE||'—')}</td>
        <td>${esc(r.colPackaging||'—')}</td>
        <td style="text-align:right">${esc(r.colLength||'—')}</td>
        <td style="text-align:right">${esc(r.colWidth||'—')}</td>
        <td style="text-align:right">${esc(r.colHeight||'—')}</td>
        <td style="text-align:right">${parseFloat(r.colVolume||0).toFixed(2)}</td>
        <td style="text-align:right">${esc(r.colGrossWeight||'—')}</td>
        <td style="text-align:right">${esc(r.colNettWeight||'—')}</td>
      </tr>`;
    }).join('');

    const now = new Date().toLocaleDateString('nl-NL', {day:'2-digit',month:'2-digit',year:'numeric'});

    const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Legplan — ${containerType}</title>
<style>
  @page { size: A4 landscape; margin: 14mm 16mm; }
  * { box-sizing:border-box; margin:0; padding:0; font-family: Arial, sans-serif; }
  body { background:#fff; color:#111; font-size:10pt; }

  /* Header */
  .pdf-header { display:flex; align-items:center; justify-content:space-between;
    border-bottom:3px solid #D91F2C; padding-bottom:8px; margin-bottom:12px; }
  .pdf-title { font-size:18pt; font-weight:800; letter-spacing:.03em; }
  .pdf-title span { color:#D91F2C; }
  .pdf-meta { font-size:8pt; color:#666; text-align:right; line-height:1.6; }
  .container-badge { display:inline-block; background:#D91F2C; color:#fff;
    font-size:11pt; font-weight:800; padding:3px 12px; border-radius:3px; margin-top:4px; }

  /* Two-column layout */
  .pdf-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-bottom:12px; }

  /* Canvas images */
  .view-section { }
  .view-label { font-size:7pt; font-weight:700; letter-spacing:.1em; text-transform:uppercase;
    color:#888; margin-bottom:4px; }
  .view-img { width:100%; border:1px solid #ccc; border-radius:3px; display:block; }

  /* Stats */
  .stats-row { display:flex; gap:10px; margin-bottom:12px; }
  .stat-box { flex:1; border:1px solid #e0e0e0; border-radius:3px; padding:6px 10px; }
  .stat-box-label { font-size:7pt; text-transform:uppercase; letter-spacing:.08em; color:#888; }
  .stat-box-val { font-size:13pt; font-weight:800; margin:2px 0; }
  .stat-box-sub { font-size:8pt; color:#555; }
  .bar-wrap { height:6px; background:#e0e0e0; border-radius:3px; overflow:hidden; margin-top:4px; }
  .bar-fill-v { height:100%; background:#00B4D8; border-radius:3px; width:${stats['lp-vol-pct']}; }
  .bar-fill-w { height:100%; background:#D91F2C; border-radius:3px; width:${stats['lp-wgt-pct']}; }
  .bar-fill-l { height:100%; background:#f59e0b; border-radius:3px; width:${stats['lp-len-pct']}; }

  /* Legend */
  .legend-section { margin-bottom:10px; }
  .legend-title { font-size:8pt; font-weight:700; text-transform:uppercase; letter-spacing:.08em;
    color:#555; margin-bottom:6px; border-bottom:1px solid #e0e0e0; padding-bottom:3px; }
  .legend-grid { display:flex; flex-wrap:wrap; gap:4px 12px; }
  .legend-item { display:flex; align-items:center; gap:5px; font-size:8pt; font-family:monospace; }
  .legend-swatch { width:12px; height:12px; border-radius:2px; flex-shrink:0; }

  /* Collo table */
  .section-title { font-size:9pt; font-weight:700; text-transform:uppercase; letter-spacing:.08em;
    color:#333; border-bottom:2px solid #D91F2C; padding-bottom:3px; margin-bottom:6px; }
  table { width:100%; border-collapse:collapse; font-size:8pt; }
  th { background:#1a2535; color:#fff; padding:4px 6px; text-align:left;
    font-size:7pt; letter-spacing:.06em; text-transform:uppercase; }
  th.r, td.r { text-align:right; }
  td { padding:3px 6px; border-bottom:1px solid #f0f0f0; }
  tr:nth-child(even) td { background:#f8f9fa; }

  /* AI advice */
  .advice-box { background:#f8f9fa; border-left:3px solid #D91F2C; padding:8px 12px;
    font-size:9pt; line-height:1.6; margin-top:10px; }
  .advice-title { font-size:8pt; font-weight:700; text-transform:uppercase;
    letter-spacing:.08em; color:#D91F2C; margin-bottom:6px; }
  .advice-box p { margin-bottom:4px; }

  /* Footer */
  .pdf-footer { border-top:1px solid #e0e0e0; padding-top:6px; margin-top:12px;
    font-size:7.5pt; color:#888; display:flex; justify-content:space-between; }
</style>
</head>
<body>

<div class="pdf-header">
  <div>
    <div class="pdf-title">LEGPLAN <span>ADVIES</span></div>
    <div class="container-badge">${containerType}</div>
  </div>
  <div class="pdf-meta">
    Royal IHC — Expediting &amp; Logistics<br>
    Gegenereerd: ${now}<br>
    Collo's: ${unique.length}
  </div>
</div>

<!-- Stats strip -->
<div class="stats-row">
  <div class="stat-box">
    <div class="stat-box-label">Volume bezetting</div>
    <div class="stat-box-val" style="color:#00B4D8">${stats['lp-vol-pct']}</div>
    <div class="stat-box-sub">${stats['lp-vol-used']} m³ / ${stats['lp-vol-max']} m³</div>
    <div class="bar-wrap"><div class="bar-fill-v"></div></div>
  </div>
  <div class="stat-box">
    <div class="stat-box-label">Gewicht bezetting</div>
    <div class="stat-box-val" style="color:#D91F2C">${stats['lp-wgt-pct']}</div>
    <div class="stat-box-sub">${stats['lp-wgt-used']} kg / ${stats['lp-wgt-max']} kg</div>
    <div class="bar-wrap"><div class="bar-fill-w"></div></div>
  </div>
  <div class="stat-box">
    <div class="stat-box-label">Lengte bezetting</div>
    <div class="stat-box-val" style="color:#f59e0b">${stats['lp-len-pct']}</div>
    <div class="stat-box-sub">${stats['lp-len-used']} cm / ${stats['lp-len-max']} cm</div>
    <div class="bar-wrap"><div class="bar-fill-l"></div></div>
  </div>
</div>

<!-- Container drawings -->
<div class="pdf-grid">
  <div class="view-section">
    <div class="view-label">Isometrisch aanzicht</div>
    ${imgS ? `<img class="view-img" src="${imgS}">` : '<p style="color:#999;font-size:8pt">Niet beschikbaar</p>'}
  </div>
  <div class="view-section">
    <div class="view-label">Bovenaanzicht (plattegrond)</div>
    ${imgT ? `<img class="view-img" src="${imgT}">` : '<p style="color:#999;font-size:8pt">Niet beschikbaar</p>'}
    <!-- Legend -->
    <div class="legend-section" style="margin-top:8px">
      <div class="legend-title">Collo kleurlegenda</div>
      <div class="legend-grid">${unique.map((r,i)=>{
        const col=COLLO_COLORS[i%COLLO_COLORS.length];
        return `<span class="legend-item"><span class="legend-swatch" style="background:${col}"></span>${esc(r.colCollo||String(i+1))}</span>`;
      }).join('')}</div>
    </div>
  </div>
</div>

<!-- Collo tabel -->
<div class="section-title">Collo Overzicht</div>
<table>
  <thead><tr>
    <th></th><th>Collo #</th><th>Leverancier</th><th>Moeder Kol E</th><th>Type verpakking</th>
    <th class="r">L cm</th><th class="r">B cm</th><th class="r">H cm</th>
    <th class="r">Vol m³</th><th class="r">Bruto kg</th><th class="r">Netto kg</th>
  </tr></thead>
  <tbody>${colloTableRows}</tbody>
</table>

${adviceHTML && adviceHTML.trim() && !adviceHTML.includes('Vul een API') ? `
<div class="advice-box">
  <div class="advice-title">🤖 AI Container Advies</div>
  ${adviceHTML}
</div>` : ''}

<div class="pdf-footer">
  <span>Royal IHC — Expediting Legplan Tool</span>
  <span>${containerType} · ${now}</span>
</div>

</body></html>`;

    // Open in new window and trigger print
    const win = window.open('', '_blank', 'width=1100,height=800');
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); }, 600);

  } catch(err) {
    alert('PDF export mislukt: ' + err.message);
  }

  btn.disabled = false; btn.textContent = '⬇ PDF Export';
}


