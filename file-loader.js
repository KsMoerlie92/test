// ── File loading: handleFile(), checkReady() ───────────────────────────────
// ── File handling ──────────────────────────────────────────────────────────
function handleFile(evt, role) {
  const file = evt.target.files[0];
  if (!file) return;

  const dzId = role === 'moeder' ? 'dz-moeder' : 'dz-expediting';
  const fnId = role === 'moeder' ? 'fn-moeder' : 'fn-expediting';
  const dz   = document.getElementById(dzId);
  const fn   = document.getElementById(fnId);

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });

      // Moederlijst: zoek het werkblad dat 'IHC PO' in rij 1 bevat (robuust voor elk bestand)
      // Expediting : gebruik altijd het eerste werkblad (index 0)
      let sheetName;
      if (role === 'moeder') {
        sheetName = null;
        for (const name of wb.SheetNames) {
          const ws_check = wb.Sheets[name];
          const range_check = ws_check['!ref'];
          if (!range_check) continue;
          const rng = XLSX.utils.decode_range(range_check);
          // Scan first 5 rows for 'IHC PO'
          outer: for (let R = rng.s.r; R <= Math.min(rng.s.r + 4, rng.e.r); R++) {
            for (let C = rng.s.c; C <= Math.min(rng.s.c + 10, rng.e.c); C++) {
              const cell = ws_check[XLSX.utils.encode_cell({r: R, c: C})];
              if (cell && String(cell.v||'').trim().toLowerCase() === 'ihc po') {
                sheetName = name;
                break outer;
              }
            }
          }
          if (sheetName) break;
        }
        if (!sheetName) {
          setStatus(`Fout: geen werkblad gevonden met kolom "IHC PO" in "${file.name}". Beschikbare bladen: ${wb.SheetNames.join(', ')}.`, true);
          return;
        }
      } else {
        sheetName = wb.SheetNames[0];
      }

      const ws   = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      fileData[role] = { data, name: file.name, sheet: sheetName };
      dz.classList.add('loaded');
      fn.textContent = `${file.name} — blad: "${sheetName}"`;
      checkReady();
    } catch(err) {
      setStatus('Fout bij lezen: ' + err.message, true);
    }
  };
  reader.readAsArrayBuffer(file);
}

function checkReady() {
  const ready = fileData.expediting && fileData.moeder;
  document.getElementById('btn-run').disabled = !ready;
  setStatus(ready
    ? 'Beide bestanden geladen — klik Verwerken'
    : 'Laad beide bestanden om te beginnen');
}

// Drag-and-drop
document.querySelectorAll('.drop-zone').forEach(dz => {
  dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('dragover');
    const input = dz.querySelector('input[type=file]');
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event('change'));
    }
  });
});

