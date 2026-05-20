// ── CIPL tab: parseCIPLSheet, initCIPL, renderCIPL, exportCIPLPDF ──────────
// ── CIPL — Commercial Invoice & Packing List ─────────────────────────────
// Columns from Sheet2 used in CIPL (0-based indices):
//   7  = Mark/Label        17 = Collo
//   2  = IHC PO            18 = Type of packaging
//   3  = Item #            23 = Gross Weight
//   4  = Item description  24 = Nett Weight
//   5  = Quantity          29 = Shipment
//   6  = UoM
//  13  = Country of Origin
//  14  = HS Code
//  15  = Value pc (EUR)
//  16  = Value total (EUR)

// ── CIPL — Commercial Invoice & Packing List ─────────────────────────────
// parseCIPLSheet: reads "CIPL format" sheet from XLSX workbook object
// Returns [dataRows, headers, metaObj]
function parseCIPLSheet(ws) {
  if (!ws) return [null, null, null];

  // Helper: read single cell (1-based row/col)
  const cv = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({r: r-1, c: c-1})];
    return cell ? String(cell.v == null ? '' : cell.v).trim() : '';
  };

  // Meta from header rows (fast — only ~15 cells)
  const meta = {
    docNo:     cv(7,10),
    shipper:   [cv(12,2),cv(13,2),cv(14,2),cv(15,2)].filter(Boolean).join(', '),
    consignee: [cv(29,2),cv(30,2),cv(31,2),cv(32,2)].filter(Boolean).join(', '),
    project:   cv(22,2),
    shipmentNr:cv(23,2),
    date:      cv(25,2),
    packages:  cv(40,5),
    nettWgt:   cv(41,5),
    grossWgt:  cv(42,5),
    vessel:    cv(45,5),
    delivery:  cv(48,5),
    c51:       cv(51,3),
  };

  // Use sheet_to_json with range starting at row 54 (0-based = 53)
  // This is MUCH faster than cell-by-cell iteration
  const allRows = XLSX.utils.sheet_to_json(ws, {
    defval: '', raw: false, range: 53  // row index 53 = Excel row 54 = headers
  });

  if (!allRows.length) return [[], [], meta];

  // Filter out non-data rows (end totals, empty delivery refs)
  const dataRows = allRows.filter(row => {
    const ref = String(row['Delivery reference'] || '').trim();
    return ref && !ref.toLowerCase().includes('endtotal');
  });

  const headers = allRows.length ? Object.keys(allRows[0]) : [];
  return [dataRows, headers, meta];
}

// buildShipmentMap: delivery ref → shipment number, from Sheet2
function buildShipmentMap(mData) {
  const headers = mData.length ? Object.keys(mData[0]) : [];
  const shipCol = headers.find(h => h && h.trim().toLowerCase().startsWith('shipment'));
  const refCol  = headers[0];
  if (!shipCol || !refCol) return {};
  const map = {};
  mData.forEach(r => {
    const ref  = String(r[refCol]  || '').trim();
    const ship = String(r[shipCol] || '').trim();
    if (ref && ship) map[ref] = ship;
  });
  return map;
}

let _ciplRows = [];

function initCIPL() {
  const moeder  = fileData.moeder;
  const statsEl = document.getElementById('cipl-stats');
  if (!moeder) {
    if (statsEl) statsEl.textContent = 'Laad de Moederlijst om CIPL te tonen.';
    return;
  }

  // Lazy parse: only run parseCIPLSheet the first time the tab is opened
  if (!moeder.ciplData && moeder._ciplWs) {
    if (statsEl) statsEl.textContent = 'CIPL sheet verwerken…';
    setTimeout(() => {
      try {
        const [d, h, m] = parseCIPLSheet(moeder._ciplWs);
        moeder.ciplData    = d;
        moeder.ciplHeaders = h;
        moeder.ciplMeta    = m;
      } catch(e) {
        if (statsEl) statsEl.textContent = 'Fout bij laden CIPL: ' + e.message;
        return;
      }
      _initCIPLFromData();
    }, 30);
    return;
  }
  _initCIPLFromData();
}

function _initCIPLFromData() {
  const moeder  = fileData.moeder;
  const statsEl = document.getElementById('cipl-stats');
  const ciplData = moeder.ciplData;
  const mData    = moeder.data;
  if (!ciplData || !ciplData.length) {
    if (statsEl) statsEl.textContent = 'Geen "CIPL format" sheet gevonden in de Moederlijst.';
    return;
  }
  const shipMap = buildShipmentMap(mData || []);
  ciplData.forEach(row => {
    const ref = (row['Delivery reference'] || '').trim();
    row._shipment = shipMap[ref] || '';
  });
  const shipVals = [...new Set(ciplData.map(r => r._shipment).filter(Boolean))]
    .sort((a,b) => (parseFloat(a)||999) - (parseFloat(b)||999) || a.localeCompare(b));
  const sel = document.getElementById('cipl-shipment-sel');
  const prev = sel.value;
  sel.innerHTML = '<option value="">— kies shipment —</option>' +
    shipVals.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if (prev && shipVals.includes(prev)) sel.value = prev;
  renderCIPL();
}

// Close _initCIPLFromData

function renderCIPL() {
  const moeder = fileData.moeder;
  if (!moeder?.ciplData) return;
  const sel      = document.getElementById('cipl-shipment-sel');
  const shipment = sel.value;
  const tbody    = document.getElementById('cipl-tbody');
  const statsEl  = document.getElementById('cipl-stats');
  const totEl    = document.getElementById('cipl-totals');
  tbody.innerHTML = '';
  if (statsEl) statsEl.textContent = '';
  if (totEl)   totEl.innerHTML = '';
  if (!shipment) {
    if (statsEl) statsEl.textContent = 'Kies een Shipment # om de CIPL te tonen.';
    return;
  }
  const rows = moeder.ciplData.filter(r => r._shipment === shipment);
  _ciplRows = rows;
  if (!rows.length) {
    if (statsEl) statsEl.textContent = `Geen regels voor Shipment #${shipment}.`;
    return;
  }
  const g = (row, key) => String(row[key] || '').trim().replace('(leeg)', '');
  let totQty=0, totVal=0, totGross=0, totNett=0;
  const frag = document.createDocumentFragment();
  rows.forEach((row, i) => {
    const qty   = parseFloat(g(row,'Totaal')) || 0;
    const vpu   = parseFloat(g(row,'Value pc (EUR)')) || 0;
    const gross = parseFloat(g(row,'Gross Weight (kg)')) || 0;
    const nett  = parseFloat(g(row,'Nett Weight (kg)')) || 0;
    totQty += qty; totVal += vpu*qty; totGross += gross; totNett += nett;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${i+1}</td>
      <td style="font-family:var(--mono);font-size:.7rem">${esc(g(row,'Mark/Label'))}</td>
      <td style="color:var(--teal);font-weight:700;font-family:var(--mono)">${esc(g(row,'Delivery reference'))}</td>
      <td class="num">${esc(g(row,'Collo'))}</td>
      <td style="max-width:200px;white-space:normal;font-size:.74rem">${esc(g(row,'Item description'))}</td>
      <td class="num">${esc(g(row,'Totaal'))}</td>
      <td>${esc(g(row,'UoM'))}</td>
      <td>${esc(g(row,'Country of Origin'))}</td>
      <td style="font-family:var(--mono);font-size:.7rem" class="hs-cell" data-hs="${esc(g(row,'HS Code'))}" onclick="toggleHSMeasures('${esc(g(row,'HS Code'))}',this.closest('tr'))" title="Klik voor douanemaatregelen">${esc(g(row,'HS Code'))}</td>
      <td class="num">${esc(g(row,'Value pc (EUR)'))}</td>
      <td class="num" style="font-weight:600">${(vpu*qty||0).toFixed(2)}</td>
      <td>${esc(g(row,'Type of packaging'))}</td>
      <td class="num">${esc(g(row,'Gross Weight (kg)'))}</td>
      <td class="num">${esc(g(row,'Nett Weight (kg)'))}</td>`;
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
  if (statsEl) statsEl.innerHTML =
    `<span>Shipment <strong>#${shipment}</strong></span><span>${rows.length} regels</span>`;
  if (totEl) totEl.innerHTML =
    `<span>Qty: <strong>${totQty.toLocaleString('nl-NL')}</strong></span>` +
    `<span>Waarde: <strong>€ ${totVal.toLocaleString('nl-NL',{minimumFractionDigits:2})}</strong></span>` +
    `<span>Bruto: <strong>${totGross.toLocaleString('nl-NL')} kg</strong></span>` +
    `<span>Netto: <strong>${totNett.toLocaleString('nl-NL')} kg</strong></span>`;
}

function exportCIPLPDF() {
  if (!_ciplRows || !_ciplRows.length) {
    alert('Kies eerst een Shipment # voordat je exporteert.'); return;
  }
  const sel      = document.getElementById('cipl-shipment-sel');
  const shipment = sel.value;
  const vessel   = (document.getElementById('cipl-vessel')?.value || '').trim();
  const meta     = (typeof moederFile !== 'undefined' ? moederFile?.ciplMeta : null)
                || (typeof fileData   !== 'undefined' ? fileData.moeder?.ciplMeta : null) || {};
  const now = new Date().toLocaleDateString('nl-NL',{day:'2-digit',month:'2-digit',year:'numeric'});

  const _r0   = _ciplRows[0] || {};
  const _keys = Object.keys(_r0);
  const _fc   = (...names) => _keys.find(k => names.some(n => k.trim().toLowerCase().includes(n))) || null;
  const _cQty   = _fc('totaal','quantity');
  const _cVpc   = _fc('value pc');
  const _cVtot  = _fc('value total','value tot');
  const _cGross = _fc('gross weight');
  const _cNett  = _fc('nett weight','net weight');
  const _cMark  = _fc('mark/label','mark');
  const _cRef   = _fc('delivery reference','delivery ref');
  const _cCollo = _fc('collo');
  const _cDesc  = _fc('item description','description');
  const _cUoM   = _fc('uom');
  const _cCOO   = _fc('country of origin');
  const _cHS    = _fc('hs code','hs-code');
  const _cPkg   = _fc('type of packaging','packaging');
  const _cMake  = _fc('make');
  const _cMat   = _fc('material');
  const _cCont  = _fc('container');
  const _cVpc2  = _fc('value p.p','value pc');
  const g   = (row, key) => key ? String(row[key] == null ? '' : row[key]).trim().replace('(leeg)', '') : '';
  const num = (row, key) => { if (!key) return 0; const v = row[key]; return typeof v === 'number' ? v : parseFloat(String(v||'').replace(',','.').trim()) || 0; };

  let totPkg=0, totGross=0, totNett=0, totVal=0;
  const seenCollos = new Set();
  _ciplRows.forEach(r => {
    const collo = g(r, _cCollo);
    if (collo && !seenCollos.has(collo)) { seenCollos.add(collo); totPkg++; }
    const vtot = _cVtot ? num(r,_cVtot) : num(r,_cVpc)*num(r,_cQty);
    const vtot2 = _cVpc2 ? num(r,_cVpc2)*num(r,_cQty) : vtot;
    totVal   += vtot2 || vtot;
    totGross += num(r,_cGross);
    totNett  += num(r,_cNett);
  });

  // Build consignee block
  const consigneeLines = (meta.consignee || 'PACIFIC SHIPBUILDING JSC').split(',').map(l=>l.trim()).join('<br>');
  const projectStr  = meta.project  || 'PROJECT : YN1320';
  const deliveryStr = meta.delivery || 'CIF Haiphong';
  const vesselStr   = vessel || meta.vessel || '';
  // Strip any label prefix the Excel cell may contain (e.g. "SHIPMENT NR: SH02" → "SH02")
  const _stripLabel = v => String(v||'').replace(/^[A-Z /]+[:#]\s*/i,'').trim();
  const shipNr  = _stripLabel(meta.shipmentNr) || shipment;
  const dateStr = now; // always today — never the date stored in the file
  const docNo   = _stripLabel(meta.docNo) || '';

  // Rows — grouped: first row of a collo gets dimensions/weights; subsequent rows get blank dims
  const seenInRows = new Set();
  const rows_html = _ciplRows.map((row, i) => {
    const collo  = g(row, _cCollo);
    const isPrimary = collo && !seenInRows.has(collo);
    if (collo) seenInRows.add(collo);
    const vtot = _cVtot ? num(row,_cVtot) : (_cVpc2 ? num(row,_cVpc2)*num(row,_cQty) : num(row,_cVpc)*num(row,_cQty));
    const grossDisp = isPrimary ? (g(row,_cGross)||'') : '';
    const nettDisp  = isPrimary ? (g(row,_cNett)||'')  : '';
    return `<tr class="${i%2===0?'even':'odd'}">
      <td class="c">${g(row,_cRef)}</td>
      <td class="c">${collo}</td>
      <td class="c">${isPrimary ? (g(row,_cPkg)||'') : ''}</td>
      <td class="r">${isPrimary ? (g(row,'Length cm')||g(row,'length')) : ''}</td>
      <td class="r">${isPrimary ? (g(row,'Width cm') ||g(row,'width'))  : ''}</td>
      <td class="r">${isPrimary ? (g(row,'Heigth cm')||g(row,'height')) : ''}</td>
      <td class="r">${grossDisp}</td>
      <td class="r">${nettDisp}</td>
      <td>${g(row,_cMark)}</td>
      <td style="max-width:150px">${g(row,_cDesc)}</td>
      <td>${g(row, _fc('code supplier') || '')}</td>
      <td>${g(row,_cHS)}</td>
      <td>${g(row,_cMake)}</td>
      <td>${g(row,_cMat)}</td>
      <td>${g(row,_cCOO)}</td>
      <td>${g(row,_cCont)}</td>
      <td>${g(row,_cUoM)}</td>
      <td class="r">${num(row,_cQty)||''}</td>
      <td class="r">${vtot>0?vtot.toFixed(2):''}</td>
      <td class="r sig-col">${vtot>0?vtot.toFixed(2):''}</td>
    </tr>`;
  }).join('');

  const LOGO_B64 = "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";
  const SIG_B64  = "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";

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>CIPL — Shipment ${shipment}</title>
<style>
@page { size: A4 landscape; margin: 10mm 10mm 12mm 10mm; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; font-size: 7pt; color: #000; background: #fff; }

/* ── Top banner ── */
.top-banner { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 4px; }
.logo-wrap img { height: 38px; }
.doc-title-wrap { text-align: center; flex: 1; padding: 0 12px; }
.doc-title { font-size: 12pt; font-weight: 900; letter-spacing: .08em; text-transform: uppercase; border-bottom: 2px solid #D91F2C; padding-bottom: 3px; margin-bottom: 3px; }
.doc-subtitle { font-size: 8pt; font-weight: 700; color: #444; }
.doc-no { font-size: 7pt; color: #666; margin-top: 2px; }

/* ── Header grid ── */
.hdr-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0; border: 1px solid #aaa; margin-bottom: 4px; }
.hdr-cell { padding: 4px 6px; border-right: 1px solid #aaa; font-size: 7pt; line-height: 1.55; }
.hdr-cell:last-child { border-right: none; }
.hdr-label { font-size: 6pt; font-weight: 800; text-transform: uppercase; letter-spacing: .07em; color: #555; display: block; margin-bottom: 2px; border-bottom: 1px solid #ddd; padding-bottom: 1px; }
.hdr-value { font-size: 7.5pt; }
.hdr-value strong { font-weight: 800; font-size: 8pt; }

/* ── Meta strip ── */
.meta-strip { display: flex; gap: 0; border: 1px solid #aaa; border-top: none; margin-bottom: 5px; }
.meta-cell { flex: 1; padding: 3px 6px; border-right: 1px solid #aaa; font-size: 6.5pt; }
.meta-cell:last-child { border-right: none; }
.meta-cell .ml { font-size: 5.5pt; font-weight: 800; text-transform: uppercase; letter-spacing: .07em; color: #777; display: block; }
.meta-cell .mv { font-weight: 700; font-size: 7.5pt; }

/* ── Table ── */
table { width: 100%; border-collapse: collapse; font-size: 6.5pt; }
thead tr th { background: #1a2535; color: #fff; padding: 3px 3px; text-align: left; font-size: 5.5pt; font-weight: 700; letter-spacing: .04em; text-transform: uppercase; white-space: nowrap; border: 1px solid #0e1621; }
th.r, td.r { text-align: right; }
th.c, td.c { text-align: center; }
td { padding: 2px 3px; border-bottom: 1px solid #e8e8e8; border-right: 1px solid #eee; vertical-align: top; }
tr.even td { background: #f7f8fa; }
tr.odd  td { background: #fff; }
tr:hover td { background: #eef4ff; }

/* ── Totals row ── */
.tot-row td { background: #1a2535 !important; color: #fff !important; font-weight: 800; border-top: 2px solid #D91F2C; font-size: 6.5pt; }

/* ── Signature ── */
.sig-col { position: relative; }
.sig-block { margin-top: 6px; text-align: center; }
.sig-block img { max-height: 75px; max-width: 180px; object-fit: contain; }

/* ── Footer ── */
.footer { display: flex; justify-content: space-between; margin-top: 4px; font-size: 6pt; color: #888; border-top: 1px solid #ddd; padding-top: 3px; }
</style>
</head><body>

<!-- Top banner -->
<div class="top-banner">
  <div class="logo-wrap"><img src="data:image/jpeg;base64,${LOGO_B64}" alt="Royal IHC"></div>
  <div class="doc-title-wrap">
    <div class="doc-title">Commercial Invoice &amp; Packing List</div>
    <div class="doc-subtitle">NO COMMERCIAL VALUE / PACKING LIST &mdash; Value for customs purpose only</div>
    <div class="doc-no">Number: ${docNo}</div>
  </div>
  <div style="text-align:right;font-size:7pt;line-height:1.7;color:#444">
    <div><strong>Date:</strong> ${dateStr}</div>
    <div><strong>Shipment NR:</strong> ${shipNr||shipment}</div>
  </div>
</div>

<!-- Header grid: Shipper | Consignee | Notify -->
<div class="hdr-grid">
  <div class="hdr-cell">
    <span class="hdr-label">Shipper's Name and Address</span>
    <div class="hdr-value">
      <strong>IHC HOLLAND B.V.</strong><br>
      SMITWEG 6<br>2961AW KINDERDIJK<br>THE NETHERLANDS<br>
      NL813735646B 01<br>
      T +31 88 015 2871<br>
      rk.balgobind@royalihc.com<br>
      <br>
      <strong>${projectStr}</strong><br>
      SHIPMENT NR: ${shipNr||shipment}<br>
      DATE: ${dateStr}
    </div>
  </div>
  <div class="hdr-cell">
    <span class="hdr-label">Consignee</span>
    <div class="hdr-value">
      <strong>PACIFIC SHIPBUILDING JSC</strong><br>
      Lot 6 Tien Lang, Commune<br>
      Hai Phong City, Vietnam<br>
      <br>
      <strong>${projectStr}</strong><br>
      ${deliveryStr}<br>
      ${vesselStr ? '<strong>Vessel: '+vesselStr+'</strong>' : ''}
    </div>
  </div>
  <div class="hdr-cell" style="border-right:none">
    <span class="hdr-label">Notify</span>
    <div class="hdr-value">
      <strong>PACIFIC SHIPBUILDING JSC</strong><br>
      Lot 6 Tien Lang, Commune<br>
      Hai Phong City, Vietnam<br>
      <br>
      Mrs. Kristyle Le<br>
      Ph # +84-225-3943273<br>
      business@daiduongship.vn<br>
      TAX Code: 0200751035
    </div>
  </div>
</div>

<!-- Meta strip -->
<div class="meta-strip">
  <div class="meta-cell"><span class="ml">Packages</span><span class="mv">${totPkg||seenCollos.size}</span></div>
  <div class="meta-cell"><span class="ml">Gross Weight (kg)</span><span class="mv">${totGross.toLocaleString('nl-NL',{minimumFractionDigits:2})}</span></div>
  <div class="meta-cell"><span class="ml">Nett Weight (kg)</span><span class="mv">${totNett.toLocaleString('nl-NL',{minimumFractionDigits:2})}</span></div>
  <div class="meta-cell"><span class="ml">Total Value USD</span><span class="mv">${totVal.toLocaleString('nl-NL',{minimumFractionDigits:2})}</span></div>
  <div class="meta-cell"><span class="ml">Vessel</span><span class="mv">${vesselStr||'—'}</span></div>
  <div class="meta-cell"><span class="ml">Delivery Terms</span><span class="mv">${deliveryStr}</span></div>
</div>

<!-- Main table -->
<table>
<thead><tr>
  <th>Delivery Ref</th><th class="c">Collo</th><th>Type of Pkg</th>
  <th class="r">L cm</th><th class="r">W cm</th><th class="r">H cm</th>
  <th class="r">Gross (kg)</th><th class="r">Nett (kg)</th>
  <th>Mark / Label</th><th>Item Description</th><th>Code Supplier</th>
  <th>HS Code</th><th>Make</th><th>Material</th>
  <th>COO</th><th>Container</th><th>UoM</th>
  <th class="r">Qty</th><th class="r">Value p.p USD</th><th class="r">Value Tot. USD</th>
</tr></thead>
<tbody>
${rows_html}
<tr class="tot-row">
  <td colspan="3"><strong>TOTAL</strong></td>
  <td colspan="2"></td>
  <td></td>
  <td class="r"><strong>${totGross.toLocaleString('nl-NL',{minimumFractionDigits:2})}</strong></td>
  <td class="r"><strong>${totNett.toLocaleString('nl-NL',{minimumFractionDigits:2})}</strong></td>
  <td colspan="9"></td>
  <td></td>
  <td class="r"><strong>${totVal.toLocaleString('nl-NL',{minimumFractionDigits:2})}</strong></td>
  <td class="r sig-col">
    <strong>${totVal.toLocaleString('nl-NL',{minimumFractionDigits:2})}</strong>
    <div class="sig-block">
      <img src="data:image/jpeg;base64,${SIG_B64}" alt="Handtekening R. Balgobind">
    </div>
  </td>
</tr>
</tbody>
</table>

<div class="footer">
  <span>Royal IHC — Commercial Invoice &amp; Packing List</span>
  <span>${projectStr} · Shipment ${shipNr||shipment} · ${vesselStr ? 'Vessel: '+vesselStr+' · ' : ''}Generated: ${now}</span>
</div>
</body></html>`;

  const win = window.open('','_blank','width=1200,height=820');
  if (!win) { alert('Pop-up geblokkeerd — sta pop-ups toe voor deze pagina.'); return; }
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 500);
}

// ── CIPL overrides: standalone page (must come AFTER js_block to win) ───────
// The js_block contains initCIPL/renderCIPL that reference fileData.moeder.
// These definitions come after the js_block, so they override those versions.

function initCIPL() {
  const statsEl = document.getElementById('cipl-stats');
  if (!moederFile) {
    if (statsEl) statsEl.textContent = 'Laad de Moederlijst om CIPL te tonen.';
    return;
  }
  // Read CIPL meta from CIPL format sheet (for PDF header info only)
  if (!moederFile.ciplMeta && moederFile._ciplWs) {
    try {
      const [,, m] = parseCIPLSheet(moederFile._ciplWs);
      moederFile.ciplMeta = m;
    } catch(e) { moederFile.ciplMeta = {}; }
  }
  // Auto-fill vessel from CIPL format sheet if the input is still empty
  const vesselInput = document.getElementById('cipl-vessel');
  if (vesselInput && !vesselInput.value.trim() && moederFile.ciplMeta?.vessel) {
    vesselInput.value = moederFile.ciplMeta.vessel;
  }
  // Always use Sheet2 for shipment data
  _initCIPLFromSheet2(moederFile.data || []);
}

function renderCIPL() {
  _renderCIPLFromSheet2();
}

function _renderCIPLFromSheet2() {
  const sel      = document.getElementById('cipl-shipment-sel');
  const shipment = sel.value;
  const shipCol  = sel.dataset.shipCol;
  const mData    = moederFile?.data || [];
  const tbody    = document.getElementById('cipl-tbody');
  const statsEl  = document.getElementById('cipl-stats');
  const totEl    = document.getElementById('cipl-totals');
  tbody.innerHTML = ''; if(statsEl)statsEl.textContent=''; if(totEl)totEl.innerHTML='';
  if (!shipment) { if(statsEl)statsEl.textContent='Kies een Shipment #.'; return; }
  const rows = mData.filter(r => String(r[shipCol]||'').trim() === shipment);
  _ciplRows = rows;
  if (!rows.length) { if(statsEl)statsEl.textContent='Geen regels voor Shipment #'+shipment; return; }
  const headers = Object.keys(mData[0]);
  const fC = n => headers.find(h=>h&&h.trim().toLowerCase().includes(n));
  // Support both Moederlijst (IHC PO) and CIPL-format (Delivery reference) column names
  const cRef   = fC('delivery reference') || fC('ihc po');
  const cItem  = fC('item #')||fC('item#')||fC('mark/label')||fC('mark');
  const cDesc  = fC('item description')||fC('description');
  const cQty   = fC('quantity')||fC('totaal');
  const cUoM   = fC('uom');
  const cCOO   = fC('country of origin');
  const cHS    = fC('hs code');
  const cVpc   = fC('value p.p')||fC('value pc');
  const cVtot  = fC('value tot')||fC('value total');
  const cPkg   = fC('type of packaging')||fC('packaging');
  const cCollo = fC('collo');
  const cGross = fC('gross weight');
  const cNett  = fC('nett weight')||fC('net weight');
  const cMake  = fC('make');
  const cMat   = fC('material');
  const cCont  = fC('container');
  const g=(r,c)=>c?String(r[c]||'').trim():'—';
  const num=(r,c)=>{ const v=r[c]; return typeof v==='number'?v:parseFloat(String(v||'').replace(',','.').trim())||0; };
  let tQ=0,tV=0,tG=0,tN=0;
  const html = rows.map((row,i)=>{
    const qty=num(row,cQty);
    const vtot=cVtot?num(row,cVtot):(cVpc?num(row,cVpc)*qty:0);
    tQ+=qty; tV+=vtot; tG+=num(row,cGross); tN+=num(row,cNett);
    return `<tr><td>${i+1}</td>
      <td style="font-family:var(--mono);font-size:.7rem">${esc(g(row,cItem))}</td>
      <td style="color:var(--teal);font-weight:700">${esc(g(row,cRef))}</td>
      <td class="num">${esc(g(row,cCollo))}</td>
      <td style="max-width:200px;white-space:normal;font-size:.74rem">${esc(g(row,cDesc))}</td>
      <td class="num">${esc(g(row,cQty))}</td><td>${esc(g(row,cUoM))}</td>
      <td>${esc(g(row,cCOO))}</td>
      <td style="font-family:var(--mono)" class="hs-cell" data-hs="${esc(g(row,cHS))}" onclick="toggleHSMeasures('${esc(g(row,cHS))}',this.closest('tr'))" title="Klik voor douanemaatregelen">${esc(g(row,cHS))}</td>
      <td class="hs-status" data-hs="${esc(g(row,cHS))}">—</td>
      <td class="num">${vtot>0?vtot.toFixed(2):esc(g(row,cVpc))}</td>
      <td class="num" style="font-weight:600">${vtot>0?vtot.toFixed(2):''}</td>
      <td>${esc(g(row,cPkg))}</td><td class="num">${esc(g(row,cCollo))}</td>
      <td class="num">${esc(g(row,cGross))}</td><td class="num">${esc(g(row,cNett))}</td></tr>`;
  }).join('');
  tbody.innerHTML = html;
  if(statsEl)statsEl.innerHTML=`<span>Shipment <strong>#${shipment}</strong></span><span>${rows.length} regels</span>`;
  if(totEl)totEl.innerHTML=
    `<span>Qty: <strong>${tQ.toLocaleString('nl-NL')}</strong></span>`+
    `<span>Waarde: <strong>$ ${tV.toLocaleString('nl-NL',{minimumFractionDigits:2})}</strong></span>`+
    `<span>Bruto: <strong>${tG.toLocaleString('nl-NL')} kg</strong></span>`+
    `<span>Netto: <strong>${tN.toLocaleString('nl-NL')} kg</strong></span>`;
}

