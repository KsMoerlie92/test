// ── Excel export: exportExcel() ────────────────────────────────────────────
function exportExcel() {
  if (!allRows.length) return;

  // Sheet 1: volledige match-resultaten
  const wsData = [[
    '#',
    'XLOOKUP Resultaat (Expediting Kol A)',
    'Leverancier (Moeder Kol K)',
    'Qty (Moeder F / Exp O)',
    'Moeder Kol E',
    'Status'
  ]];

  allRows.forEach((r, i) => {
    wsData.push([
      i + 1,
      r.xlookup     || '',
      r.colSupplier || '',
      (r.colF ? r.colF + (r.expColO ? ' / ' + r.expColO : '') : ''),
      r.colE        || '',
      r.noMatch ? 'ALLEEN IN MOEDERLIJST' : 'IN BEIDE LIJSTEN'
    ]);
  });

  // Sheet 2: alleen matches (in beide lijsten)
  const matchData = [[
    '#',
    'XLOOKUP (Expediting Kol A)',
    'Leverancier (Moeder Kol K)',
    'Qty (Moeder F / Exp O)',
    'Moeder Kol E',
  ]];
  allRows
    .filter(r => !r.noMatch)
    .forEach((r, i) => {
      matchData.push([i + 1, r.xlookup || '', r.colSupplier || '', (r.colF ? r.colF + (r.expColO ? ' / ' + r.expColO : '') : ''), r.colE || '']);
    });

  const wb  = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet(wsData);
  const ws2 = XLSX.utils.aoa_to_sheet(matchData);

  ws1['!cols'] = [{wch:4},{wch:28},{wch:22},{wch:14},{wch:18},{wch:22}];
  ws2['!cols'] = [{wch:4},{wch:28},{wch:22},{wch:14},{wch:18}];

  XLSX.utils.book_append_sheet(wb, ws1, 'Alle Resultaten');
  XLSX.utils.book_append_sheet(wb, ws2, 'In Beide Lijsten');
  XLSX.writeFile(wb, 'IHC_PO_Match_Resultaten.xlsx');
}

