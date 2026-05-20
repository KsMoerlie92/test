// ── PO/Item parser: expandRange(), parseItemCodes(), buildExpandedOutput() ─
// ══════════════════════════════════════════════════════════════════════════
//  VBA-logica gport naar JS
//  (exacte vertaling van ParseItemCodes + ExpandRange uit de .bas)
// ══════════════════════════════════════════════════════════════════════════

/**
 * VBA ExpandRange:
 *   startCode = bijv. "29-1"
 *   endCode   = bijv. "29-4" of korthand "4"
 *   Returns   = ["29-1","29-2","29-3","29-4"]
 */
function expandRange(startCode, endCode) {
  // Verwijder optioneel voorloopstreepje
  if (startCode.startsWith('-')) startCode = startCode.slice(1);
  if (endCode.startsWith('-'))   endCode   = endCode.slice(1);

  // Splits op LAATSTE koppelteken (zodat "10-1" correct werkt)
  const ld1 = startCode.lastIndexOf('-');
  if (ld1 < 0) return [];
  const p1 = startCode.slice(0, ld1);
  const s1 = parseInt(startCode.slice(ld1 + 1), 10);

  let p2, s2;
  if (endCode.includes('-')) {
    const ld2 = endCode.lastIndexOf('-');
    p2 = endCode.slice(0, ld2);
    s2 = parseInt(endCode.slice(ld2 + 1), 10);
  } else {
    p2 = p1;
    s2 = parseInt(endCode, 10);
  }

  const result = [];
  if (p1 === p2 && !isNaN(s1) && !isNaN(s2) && s2 >= s1) {
    for (let i = s1; i <= s2; i++) result.push(p1 + '-' + i);
  }
  return result;
}

/**
 * parseItemCodes — vertaalt een ruwe celwaarde naar een array van losse codes.
 *
 * Ondersteunde notaties:
 *   Enkel:           -3-1   of   3-1
 *   Bare getal:      -3-1 & -2       →  3-1, 3-2   (prefix geerfed van vorig item)
 *   Reeks (t/m/÷):  -4-1 t/m -7     →  4-1 t/m 4-7
 *                   -4-1÷-7          →  zelfde
 *   Slash-lijst:    114/-115/-116    →  114, 115, 116
 *   Scheidingstekens: & + , ; spatie
 */
function parseItemCodes(rawValue) {
  if (!rawValue || !String(rawValue).trim()) return [];

  let cleaned = String(rawValue);

  // ── Normaliseer ────────────────────────────────────────────────────────────
  // ÷ (charcode 247) → t/m
  cleaned = cleaned.replace(/\xf7/g,          't/m');
  cleaned = cleaned.replace(/\u00f7/g,         't/m');
  cleaned = cleaned.replace(/T\/M/g,           't/m');
  cleaned = cleaned.replace(/T\/m/g,           't/m');
  cleaned = cleaned.replace(/t\/M/g,           't/m');
  // Scheidingstekens → spatie  (/ wordt per token apart afgehandeld)
  cleaned = cleaned.replace(/&/g,   ' ');
  cleaned = cleaned.replace(/\+/g,  ' ');
  cleaned = cleaned.replace(/,/g,   ' ');
  cleaned = cleaned.replace(/;/g,   ' ');
  // Meervoudige spaties samenvoegen (belangrijk: NIET /g met replace want split regelt dit)
  const tokens = cleaned.split(/\s+/).filter(t => t.length > 0);

  const output = [];
  let lastPrefix = '';
  let i = 0;

  while (i < tokens.length) {
    let t = tokens[i].replace(/^'+/, '');  // verwijder apostrof-prefix

    if (!t) { i++; continue; }

    // ── GEVAL 1: inline t/m  e.g. "4-1t/m4-7" ───────────────────────────
    const tmIdx = t.toLowerCase().indexOf('t/m');
    if (tmIdx > 0 && t.toLowerCase() !== 't/m') {
      const start = t.slice(0, tmIdx).trim();
      const end   = t.slice(tmIdx + 3).trim();
      const exp = expandRange(start, end);
      exp.forEach(v => { output.push(v); lastPrefix = v.slice(0, v.lastIndexOf('-')); });
      i++; continue;
    }

    // ── GEVAL 2: standalone "t/m" token ─────────────────────────────────
    if (t.toLowerCase() === 't/m') {
      // Vorig item uit output is de startCode; volgend niet-leeg token = endCode
      const startCode = output.length ? output.pop() : '';
      let endCode = '';
      let j = i + 1;
      while (j < tokens.length) {
        const candidate = tokens[j].replace(/^'+/, '');
        if (candidate) { endCode = candidate; tokens[j] = ''; break; }
        j++;
      }
      if (startCode && endCode) {
        const exp = expandRange(startCode, endCode);
        exp.forEach(v => { output.push(v); lastPrefix = v.slice(0, v.lastIndexOf('-')); });
      }
      i = j + 1; continue;
    }

    // ── GEVAL 3: slash-notatie  e.g. "114/-115/-116" ────────────────────
    if (t.indexOf('/', 1) > 0) {
      const parts = t.split('/');
      let known = lastPrefix;
      for (const p of parts) {
        const s = p.replace(/^-/, '');
        if (!s) continue;
        if (s.includes('-')) {
          output.push(s);
          known = s.slice(0, s.lastIndexOf('-'));
          lastPrefix = known;
        } else if (/^\d+$/.test(s) && known) {
          output.push(known + '-' + s);
        } else if (/^\d+$/.test(s)) {
          output.push(s);
        }
      }
      i++; continue;
    }

    // ── GEVAL 4: volledige code  "-4-1" of "4-1" ────────────────────────
    const tc = t.startsWith('-') ? t.slice(1) : t;
    if (tc.includes('-')) {
      const dp = tc.lastIndexOf('-');
      if (dp > 0 && /^\d+$/.test(tc.slice(dp + 1))) {
        output.push(tc);
        lastPrefix = tc.slice(0, dp);
        i++; continue;
      }
    }

    // ── GEVAL 5: bare getal  "-2" of "2"  →  lastPrefix + "-" + getal ──
    const bare = t.startsWith('-') ? t.slice(1) : t;
    if (/^\d+$/.test(bare) && lastPrefix) {
      output.push(lastPrefix + '-' + bare);
    }

    i++;
  }

  return output;
}


function buildExpandedOutput(moedData) {
  if (!moedData.length) return { rows: [], error: 'Moederlijst is leeg.' };

  // IHC PO  = Kolom C (index 2) — altijd, ten alle tijden
  // Item #  = Kolom D (index 3) — altijd, ten alle tijden
  // Kolom AA = index 26 — rijen met "not received yet" worden uitgesloten
  const headers = Object.keys(moedData[0]);

  if (headers.length < 4) {
    return { rows: [], error: `Moederlijst heeft slechts ${headers.length} kolom(men). Verwacht minimaal 4 (A t/m D).` };
  }

  const poCol   = headers[2];    // Kolom C = index 2
  const itemCol = headers[3];    // Kolom D = index 3
  // Zoek kolommen op naam — robuust voor elk bestand met afwijkende kolomvolgorde
  const findCol = (names) => headers.find(h => h && names.some(n => String(h).trim().toLowerCase().includes(n)));
  const colE        = headers[4];                 // Kolom E — index 4
  const colLocation = findCol(['location']);       // Kol AA — detail panel
  const colZ        = findCol(['received']);       // Kol AB — detail panel
  const colSupplier = findCol(['name of supplier','supplier name','leverancier']); // Col K
  const colF           = findCol(['quantity']);       // Col F — Quantity
  const colMaterial     = findCol(['material']);       // Col M
  const colCollo        = findCol(['collo']);          // Col R
  const colPackaging    = findCol(['type of packaging','packaging type']); // Col S
  const colLength       = findCol(['length']);         // Col T — Length cm
  const colWidth        = findCol(['width']);          // Col U — Width cm
  const colHeight       = findCol(['heigth','height']);// Col V — Height cm
  const colVolume       = findCol(['volume']);         // Col W — Volume M3
  const colGrossWeight  = findCol(['gross weight']);   // Col X
  const colNettWeight   = findCol(['nett weight','net weight']); // Col Y
  const colMRID            = findCol(['mr id','mrid']);
  const colChecked          = findCol(['checked']);
  const colLogisticInsp     = findCol(['logistic inspection','logistic']);
  const colInspResult       = findCol(['inspection result']);
  const colShipment         = findCol(['shipment']);
  const colAL               = findCol(['bl #','bl#','bill of lading']);

  console.log('[PO Matcher] colE=', colE);

  const rows = [];

  for (const row of moedData) {

    const po      = String(row[poCol]   || '').trim();
    const rawItem = String(row[itemCol] || '').trim();

    if (!po || !rawItem) continue;

    const codes = parseItemCodes(rawItem);

    for (const code of codes) {
      rows.push({
        po,                          // Expanded Output Kolom A
        item:     code,              // Expanded Output Kolom B
        combined: po + '-' + code,   // Expanded Output Kolom C  ← lookup waarde
        colE:        colE        ? String(row[colE]        || '').trim() : '',  // Moeder Kol E
        colLocation: colLocation  ? String(row[colLocation]  || '').trim() : '',  // Kol AA — detail
        colZ:        colZ        ? String(row[colZ]        || '').trim() : '',  // Kol AB — detail
        colSupplier: colSupplier  ? String(row[colSupplier]  || '').trim() : '',  // Name of Supplier (K)
        colF:          colF          ? String(row[colF]          || '').trim() : '',
        colMaterial:   colMaterial   ? String(row[colMaterial]   || '').trim() : '',
        colCollo:      colCollo      ? String(row[colCollo]      || '').trim() : '',
        colPackaging:  colPackaging  ? String(row[colPackaging]  || '').trim() : '',
        colLength:     colLength     ? String(row[colLength]     || '').trim() : '',
        colWidth:      colWidth      ? String(row[colWidth]      || '').trim() : '',
        colHeight:     colHeight     ? String(row[colHeight]     || '').trim() : '',
        colVolume:     colVolume     ? String(row[colVolume]     || '').trim() : '',
        colGrossWeight:colGrossWeight? String(row[colGrossWeight]|| '').trim() : '',
        colNettWeight: colNettWeight ? String(row[colNettWeight] || '').trim() : '',
        colMRID:          colMRID         ? String(row[colMRID]          || '').trim() : '',
        colChecked:       colChecked      ? String(row[colChecked]       || '').trim() : '',
        colLogisticInsp:  colLogisticInsp ? String(row[colLogisticInsp]  || '').trim() : '',
        colInspResult:    colInspResult   ? String(row[colInspResult]    || '').trim() : '',
        colShipment:      colShipment     ? String(row[colShipment]      || '').trim() : '',
        colAL:            colAL           ? String(row[colAL]            || '').trim() : '',
      });
    }
  }

  return { rows, error: null };
}

