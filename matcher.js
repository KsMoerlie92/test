// ── Matcher: buildExpColAMap(), runMatcher(), _doMatch() ───────────────────
// ══════════════════════════════════════════════════════════════════════════
//  Stap 2: bouw Expediting Kolom A lookup/return map
//
//  XLOOKUP semantiek:
//    =XLOOKUP( lookup_value, lookup_array, return_array )
//    waarbij lookup_array = Expediting Kol A
//    en      return_array = Expediting Kol A  (zelfde kolom)
//
//  Implementatie: Map<normalised_value → original_value>
// ══════════════════════════════════════════════════════════════════════════
function buildExpColAMap(expData) {
  const map    = new Map();  // key → original col A value
  const mapQty = new Map();  // key → col O (Qty) value
  if (!expData.length) return { map, mapQty };

  const headers  = Object.keys(expData[0]);
  const colAKey  = headers[0];  // Col A = Purchase Order No (combined key)
  // Col O = index 14, but find by name too for robustness
  const colOKey  = headers.find(h => h && ['qty','quantity'].some(n => String(h).trim().toLowerCase() === n))
                || headers[14];

  for (const row of expData) {
    const val = String(row[colAKey] || '').trim();
    if (val) {
      map.set(val.toLowerCase(), val);
      mapQty.set(val.toLowerCase(), String(row[colOKey] || '').trim());
    }
  }
  return { map, mapQty };
}

// ── Kolom-detectie hulpfunctie ─────────────────────────────────────────────
function detectColumns(data) {
  if (!data.length) return null;
  const headers = Object.keys(data[0]);
  let poCol = null, itemCol = null;

  for (const h of headers) {
    const norm = h.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!poCol   && (norm === 'ihc po'  || norm.includes('ihc po')))  poCol   = h;
    if (!itemCol && (norm === 'item #'  || norm === 'item#' ||
                     norm.includes('item #') || norm.includes('item#'))) itemCol = h;
  }
  return { poCol, itemCol };
}

// ══════════════════════════════════════════════════════════════════════════
//  Hoofd match-routine
// ══════════════════════════════════════════════════════════════════════════
function runMatcher() {
  const btn    = document.getElementById('btn-run');
  const loader = document.getElementById('loader');
  btn.disabled = true;
  loader.classList.add('active');
  document.getElementById('btn-label').textContent = 'Verwerken…';

  setTimeout(() => {
    try { _doMatch(); }
    catch(err) {
      setStatus('Fout: ' + err.message, true);
      console.error(err);
    }
    btn.disabled = false;
    loader.classList.remove('active');
    document.getElementById('btn-label').textContent = '▶ Verwerken';
  }, 80);
}

function _doMatch() {
  const expData  = fileData.expediting.data;
  const moedData = fileData.moeder.data;

  if (!expData.length || !moedData.length) {
    setStatus('Lege dataset gedetecteerd.', true); return;
  }

  // ── Stap 1: Moederlijst → Expanded Output ──
  const { rows: expandedRows, error } = buildExpandedOutput(moedData);
  if (error) { setStatus(error, true); return; }
  if (!expandedRows.length) {
    setStatus('Geen expanded rows gegenereerd uit Moederlijst — controleer kolommen "IHC PO" en "Item #".', true);
    return;
  }

  // ── Stap 2: Expediting Kolom A als lookup + return array ──
  //
  //  XLOOKUP(
  //    lookup_value  = Expanded Output Kolom C  (combined key)
  //    lookup_array  = Expediting Kolom A
  //    return_array  = Expediting Kolom A   ← zelfde kolom
  //  )
  //
  const { map: expColAMap, mapQty: expColOMap } = buildExpColAMap(expData);
  const expColASize = expColAMap.size;

  // ── Stap 3: voer XLOOKUP uit voor elke Expanded Output rij ──
  const results = [];

  for (const row of expandedRows) {
    const lookupKey    = row.combined.toLowerCase();   // normaliseer voor vergelijking
    const xlookupValue = expColAMap.get(lookupKey) ?? null;

    results.push({
      po:       row.po,
      item:     row.item,
      combined: row.combined,          // Expanded Output Kolom C = lookup waarde
      colE:        row.colE,        // Moeder Kol E
      colLocation: row.colLocation, // Kol AA — detail panel
      colZ:        row.colZ,        // Kol AB — detail panel
      colSupplier: row.colSupplier,  // Name of Supplier (K) — Moederlijst
      colF:          row.colF,
      colMaterial:   row.colMaterial,
      colCollo:      row.colCollo,
      colPackaging:  row.colPackaging,
      colLength:     row.colLength,
      colWidth:      row.colWidth,
      colHeight:     row.colHeight,
      colVolume:     row.colVolume,
      colGrossWeight:row.colGrossWeight,
      colNettWeight: row.colNettWeight,
      colMRID:         row.colMRID,
      colChecked:      row.colChecked,
      colLogisticInsp: row.colLogisticInsp,
      colInspResult:   row.colInspResult,
      colShipment:     row.colShipment,
      colAL:           row.colAL,
      expColO:  xlookupValue ? (expColOMap.get(row.combined.toLowerCase()) || '') : '',  // Qty col O — Expediting
      xlookup:  xlookupValue,          // Expediting Kolom A waarde (return array)
      noMatch:  xlookupValue === null,
    });
  }

  allRows = results;

  renderTable(results);
  updateStats(results, expColASize);

  const nMatch   = results.filter(r => !r.noMatch).length;
  const nNoMatch = results.filter(r =>  r.noMatch).length;
  setStatus(
    `Verwerkt: ${expandedRows.length} Expanded Output rijen → ` +
    `${nMatch} matches · ${nNoMatch} niet gevonden in Expediting`
  );

  document.getElementById('btn-export').classList.add('visible');
}

