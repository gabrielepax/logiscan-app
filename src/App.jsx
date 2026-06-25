import { useState, useEffect, useRef } from 'react';
import { supabase } from './supabaseClient';
import * as XLSX from 'xlsx';

// Web Audio API — crea il contesto la prima volta che l'utente interagisce
let audioCtx = null;
function getAudioCtx() {
  if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  return audioCtx;
}

function playBeep({ frequency = 880, duration = 0.12, type = 'sine', volume = 0.4, delay = 0 } = {}) {
  try {
    const ctx = getAudioCtx();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.type = type;
    osc.frequency.setValueAtTime(frequency, ctx.currentTime + delay);
    gain.gain.setValueAtTime(volume, ctx.currentTime + delay);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + delay + duration);
    osc.start(ctx.currentTime + delay);
    osc.stop(ctx.currentTime + delay + duration + 0.05);
  } catch { /* audio non disponibile */ }
}

const sounds = {
  // Scansione singola OK — bip acuto breve
  ok() {
    playBeep({ frequency: 1320, duration: 0.1, type: 'sine', volume: 0.35 });
  },
  // Cartone QR OK — doppio bip ascendente
  carton() {
    playBeep({ frequency: 880, duration: 0.09, type: 'sine', volume: 0.35, delay: 0 });
    playBeep({ frequency: 1320, duration: 0.12, type: 'sine', volume: 0.4, delay: 0.12 });
  },
  // Scansione completata — tre note ascendenti
  complete() {
    playBeep({ frequency: 880,  duration: 0.1, volume: 0.35, delay: 0 });
    playBeep({ frequency: 1100, duration: 0.1, volume: 0.38, delay: 0.13 });
    playBeep({ frequency: 1320, duration: 0.18, volume: 0.42, delay: 0.26 });
  },
  // Errore — bip basso lungo
  error() {
    playBeep({ frequency: 220, duration: 0.35, type: 'square', volume: 0.25 });
  },
  // Duplicato — doppio bip medio
  duplicate() {
    playBeep({ frequency: 520, duration: 0.1, type: 'sine', volume: 0.3, delay: 0 });
    playBeep({ frequency: 520, duration: 0.1, type: 'sine', volume: 0.3, delay: 0.15 });
  },
};

export default function App() {
  const [activeModule, setActiveModule] = useState('arrivi');
  const [menuOpen, setMenuOpen] = useState(false);
  const [currentView, setCurrentView] = useState('dashboard');
  const [poLines, setPoLines] = useState([]);
  const [activeLineKey, setActiveLineKey] = useState(null);
  const [loading, setLoading] = useState(false);

  const [activeLine, setActiveLine] = useState(null);
  const [expectedSerials, setExpectedSerials] = useState({});
  const [scannedSerials, setScannedSerials] = useState([]);
  const [cartonsScanned, setCartonsScanned] = useState(0);
  const [scannerValue, setScannerValue] = useState('');
  const [feedback, setFeedback] = useState({ text: '', type: '' });

  const [filterInvoice, setFilterInvoice] = useState('');
  const [filterItem, setFilterItem] = useState('');
  const [filterSNYes, setFilterSNYes] = useState(true);
  const [filterSNNo, setFilterSNNo] = useState(false);
  const [downloadedKeys, setDownloadedKeys] = useState(new Set());

  const scannerInputRef = useRef(null);

  useEffect(() => {
    fetchPOLines();
  }, []);

  useEffect(() => {
    if (currentView === 'scan' && scannerInputRef.current) {
      scannerInputRef.current.focus();
    }
    const handleGlobalClick = () => {
      if (currentView === 'scan' && scannerInputRef.current) {
        scannerInputRef.current.focus();
      }
    };
    window.addEventListener('click', handleGlobalClick);
    return () => window.removeEventListener('click', handleGlobalClick);
  }, [currentView]);

  async function fetchPOLines() {
    setLoading(true);
    const [{ data, error }, { data: scannedKeys }] = await Promise.all([
      supabase.from('po_lines').select('*').order('arrival_date', { ascending: true }),
      supabase.from('scanned_serials').select('po_line_key')
    ]);

    if (error) {
      alert("Errore nel caricamento dei dati: " + error.message);
    } else {
      const countMap = {};
      (scannedKeys || []).forEach(r => {
        countMap[r.po_line_key] = (countMap[r.po_line_key] || 0) + 1;
      });
      setPoLines((data || []).map(l => ({ ...l, scanned_count: countMap[l.unique_key] || 0 })));
    }
    setLoading(false);
  }

  function parseCSV(text, delimiter) {
    const lines = text.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);
    if (lines.length === 0) return [];
    const headers = lines[0].split(delimiter).map(h => h.replace(/^["']|["']$/g, '').trim());
    const result = [];
    for (let i = 1; i < lines.length; i++) {
      const currentLine = lines[i].split(delimiter).map(v => v.replace(/^["']|["']$/g, '').trim());
      if (currentLine.length < headers.length) continue;
      const obj = {};
      headers.forEach((header, index) => { obj[header] = currentLine[index] || ""; });
      result.push(obj);
    }
    return result;
  }

  async function handlePOLinesUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    // Reset input so same file can be re-selected
    event.target.value = '';

    setLoading(true);
    const reader = new FileReader();
    reader.onload = async function(e) {
      const records = parseCSV(e.target.result, ',');

      if (records.length === 0 || !records[0]["PO INTERNAL ID"] || !records[0]["Line ID"]) {
        alert("Formato file non valido. Controlla la presenza delle colonne 'PO INTERNAL ID' e 'Line ID'.");
        setLoading(false);
        return;
      }

      const newKeys = new Set(records.map(row => `${row["PO INTERNAL ID"]}_${row["Line ID"]}`));

      const { data: existingLines } = await supabase.from('po_lines').select('unique_key, is_user_confirmed, item_code, line_id');
      const existing = existingLines || [];

      // Lines not present in the new CSV that must be removed
      const toRemove = existing.filter(l => !newKeys.has(l.unique_key));
      // Among those, ones still in progress (not confirmed) block the entire update
      const blocked = toRemove.filter(l => l.is_user_confirmed === false);

      if (blocked.length > 0) {
        const names = blocked.map(l => `${l.item_code} (Linea ${l.line_id})`).join('\n');
        alert(`Impossibile aggiornare: le seguenti righe sono in corso di verifica e non possono essere rimosse:\n\n${names}\n\nConcludi o rimuovi manualmente queste righe prima di ricaricare il file.`);
        setLoading(false);
        return;
      }

      // Delete confirmed (scaricato) lines not in the new CSV, with all their references
      for (const line of toRemove) {
        await supabase.from('expected_serials').delete().eq('po_line_key', line.unique_key);
        await supabase.from('scanned_serials').delete().eq('po_line_key', line.unique_key);
        await supabase.from('po_lines').delete().eq('unique_key', line.unique_key);
      }

      const existingKeys = new Set(existing.map(l => l.unique_key));

      const rowsToUpsert = records.map(row => {
        const poInternalId = row["PO INTERNAL ID"];
        const lineId = row["Line ID"];
        const key = `${poInternalId}_${lineId}`;
        const chinaInvoice = row["[PAX] CHINA INVOICE"] ? row["[PAX] CHINA INVOICE"].trim() : "SENZA FATTURA (N/D)";
        const itemCode = row["Items - Item"] ? row["Items - Item"].trim() : "N/D";

        const base = {
          unique_key: key,
          po_internal_id: poInternalId,
          line_id: lineId,
          po_name: row["Items - PO"] || "N/D",
          description: row["Items - Description"] || "N/D",
          qty_expected: parseInt(row["Items - Quantity Expected"]) || 0,
          china_invoice: chinaInvoice,
          item_code: itemCode,
          arrival_date: row["DATA DI ARRIVO"] || "N/D",
          part_number: row["[PAX] Vendor Part Number"] || itemCode,
        };
        // sn_required comes from the "SN" column in the CSV (Yes/Si = true, anything else = false)
        const snValue = (row["SN"] || "").trim().toLowerCase();
        base.sn_required = snValue === 'yes' || snValue === 'si' || snValue === 'sì';

        // Only set scan-state defaults for brand-new lines
        if (!existingKeys.has(key)) {
          base.cartons_scanned = 0;
          base.is_user_confirmed = false;
        }
        return base;
      });

      const { error } = await supabase.from('po_lines').upsert(rowsToUpsert);
      if (error) {
        alert("Errore durante il salvataggio su database: " + error.message);
      } else {
        const removedCount = toRemove.length;
        await fetchPOLines();
        if (removedCount > 0) {
          alert(`Aggiornamento completato. ${removedCount} riga/righe non più presenti nel file sono state rimosse.`);
        }
      }
    };
    reader.readAsText(file);
  }

  async function handleSNUpload(event, line) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    setLoading(true);

    const reader = new FileReader();
    reader.onload = async function(e) {
      let snRecords;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        snRecords = XLSX.utils.sheet_to_json(ws, { defval: '' });
      } catch {
        alert("Errore: impossibile leggere il file. Assicurati che sia un file Excel (.xls/.xlsx).");
        setLoading(false);
        return;
      }

      if (snRecords.length === 0 || !Object.hasOwn(snRecords[0], 'SN')) {
        alert("Errore: impossibile trovare la colonna 'SN' nel file delle matricole.");
        setLoading(false);
        return;
      }

      const validRows = snRecords.filter(row => String(row['SN']).trim() !== '');

      const qtyProvided = validRows.length;
      const qtyExpected = line.qty_expected;
      let fullyMatched = true;

      // Validazione quantità
      if (qtyProvided !== qtyExpected) {
        fullyMatched = false;
        const proceed = window.confirm(
          `Attenzione: quantità non corrispondente.\n\n` +
          `Attese: ${qtyExpected} pz\n` +
          `Fornite nel file: ${qtyProvided} pz\n\n` +
          `Vuoi procedere comunque?`
        );
        if (!proceed) { setLoading(false); return; }
      }

      // Validazione PN
      const expectedPN = (line.part_number || '').trim();
      if (expectedPN && expectedPN !== 'N/D') {
        const mismatchedPNs = [...new Set(
          validRows
            .map(row => String(row['PN'] || '').trim())
            .filter(pn => pn !== '' && pn !== expectedPN)
        )];
        if (mismatchedPNs.length > 0) {
          fullyMatched = false;
          const proceed = window.confirm(
            `Attenzione: il file contiene Part Number non corrispondenti.\n\n` +
            `PN atteso dalla riga: ${expectedPN}\n` +
            `PN trovati nel file: ${mismatchedPNs.join(', ')}\n\n` +
            `Vuoi procedere comunque?`
          );
          if (!proceed) { setLoading(false); return; }
        }
      }

      await supabase.from('expected_serials').delete().eq('po_line_key', line.unique_key);
      await supabase.from('scanned_serials').delete().eq('po_line_key', line.unique_key);
      await supabase.from('po_lines').update({
        cartons_scanned: 0,
        is_user_confirmed: false,
        sn_loaded: fullyMatched
      }).eq('unique_key', line.unique_key);

      const serialsToInsert = validRows.map(row => ({
        po_line_key: line.unique_key,
        serial: String(row['SN']).trim(),
        model: String(row['Model'] || 'N/D').trim(),
        pn: String(row['PN'] || 'N/D').trim()
      }));

      const { error } = await supabase.from('expected_serials').insert(serialsToInsert);
      if (error) {
        alert("Errore nel caricamento delle matricole: " + error.message);
      } else {
        await fetchPOLines();
        alert(`Matricole caricate: ${serialsToInsert.length} SN per la riga ${line.item_code}.`);
      }
    };
    reader.readAsArrayBuffer(file);
  }


  async function startScanningSession(line) {
    setLoading(true);
    setActiveLineKey(line.unique_key);
    setActiveLine(line);
    setCartonsScanned(line.cartons_scanned || 0);

    const { data: expectedData } = await supabase
      .from('expected_serials')
      .select('serial, model, pn')
      .eq('po_line_key', line.unique_key);

    const { data: scannedData } = await supabase
      .from('scanned_serials')
      .select('serial, model, pn, scanned_at')
      .eq('po_line_key', line.unique_key)
      .order('scanned_at', { ascending: false });

    const expectedMap = {};
    if (expectedData) {
      expectedData.forEach(d => { expectedMap[d.serial] = { model: d.model, pn: d.pn }; });
    }
    const formattedScanned = (scannedData || []).map(s => ({
      serial: s.serial,
      model: s.model,
      pn: s.pn,
      time: new Date(s.scanned_at).toLocaleTimeString('it-IT')
    }));

    setExpectedSerials(expectedMap);
    setScannedSerials(formattedScanned);

    if (line.is_user_confirmed) {
      openReviewSession();
    } else {
      setFeedback({ text: 'Scanner pronto, punta il laser sul codice.', type: 'success' });
      setCurrentView('scan');
    }
    setLoading(false);
  }

  async function reopenScanningSession(line) {
    if (!window.confirm(`Riaprire la riga "${line.item_code}" per modificare le rilevazioni?\n\nL'arrivo verrà riportato in stato "In Corso".`)) return;
    await supabase.from('po_lines').update({ is_user_confirmed: false }).eq('unique_key', line.unique_key);
    await fetchPOLines();
    await startScanningSession({ ...line, is_user_confirmed: false });
  }

  async function downloadArrivoCSV(line) {
    setLoading(true);
    setDownloadedKeys(prev => new Set(prev).add(line.unique_key));
    const { data: scannedData, error } = await supabase
      .from('scanned_serials')
      .select('serial, model, pn, scanned_at')
      .eq('po_line_key', line.unique_key)
      .order('scanned_at', { ascending: true });
    setLoading(false);
    if (error) { alert("Errore nel recupero dati: " + error.message); return; }

    buildAndDownloadCSV(line, scannedData || []);
  }

  function buildAndDownloadCSV(line, serials) {
    const header = "Internal ID,Date,Document Number,Subsidiary,Item,CODICE CINESE,Fornitore,Memo,[PAX] China Invoice,Quantity Riga,Quantity Serial,Seriale,Line ID,Surrogate ID,Vendor DDT,Vendor DDT Date";
    const rows = serials.map(s => [
      line.po_internal_id,
      line.arrival_date,
      line.po_name,
      "",
      line.item_code,
      s.model,
      "",
      line.description,
      line.china_invoice,
      line.qty_expected,
      1,
      s.serial,
      line.line_id,
      line.line_id,
      "",
      ""
    ]);
    const csv = [header, ...rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(','))].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${line.china_invoice}-${line.item_code}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function deleteScannedSerial(serial) {
    if (!window.confirm(`Eliminare la matricola ${serial} dalle rilevazioni?`)) return;
    const { error } = await supabase
      .from('scanned_serials')
      .delete()
      .eq('po_line_key', activeLineKey)
      .eq('serial', serial);
    if (error) { alert("Errore: " + error.message); return; }
    setScannedSerials(prev => prev.filter(s => s.serial !== serial));
    await fetchPOLines();
  }

  async function deleteAllScannedSerials() {
    if (!window.confirm(`Eliminare TUTTE le ${scannedSerials.length} matricole rilevate per questa riga?\n\nL'operazione non è reversibile.`)) return;
    const { error } = await supabase
      .from('scanned_serials')
      .delete()
      .eq('po_line_key', activeLineKey);
    if (error) { alert("Errore: " + error.message); return; }
    await supabase.from('po_lines').update({ cartons_scanned: 0 }).eq('unique_key', activeLineKey);
    setScannedSerials([]);
    setCartonsScanned(0);
    setFeedback({ text: 'Rilevazioni azzerate. Scanner pronto.', type: 'success' });
    await fetchPOLines();
  }

  function triggerVibration(pattern) {
    if (navigator.vibrate) navigator.vibrate(pattern);
  }

  function isMasterQRCode(str) {
    return str.startsWith("[)>") || str.includes("{GS}") || str.includes("\x1d");
  }

  function parseMasterQRCode(str) {
    // eslint-disable-next-line no-control-regex
    let normalized = str.replace(/\x1e/g, "{RS}").replace(/\x1d/g, "{GS}").replace(/\x04/g, "{Eot}");
    if (normalized.startsWith("[)>")) normalized = normalized.substring(3);
    if (normalized.startsWith("{RS}")) normalized = normalized.substring(4);
    let parts = normalized.split("{GS}").map(p => p.trim());
    parts = parts.map(p => p.replace(/{RS}/g, "").replace(/{Eot}/g, "")).filter(p => p.length > 0);
    if (parts.length < 3) return null;
    return { qty: parseInt(parts[0], 10), model: parts[1], serials: parts.slice(2) };
  }

  async function handleScanSubmit(e) {
    e.preventDefault();
    const rawInput = (scannerInputRef.current?.value || scannerValue).trim();
    setScannerValue('');
    if (scannerInputRef.current) scannerInputRef.current.value = '';
    if (!rawInput || !activeLineKey) return;

    const totalExpected = activeLine?.qty_expected || Object.keys(expectedSerials).length;
    if (scannedSerials.length === totalExpected) {
      openReviewSession();
      return;
    }

    if (isMasterQRCode(rawInput)) {
      const qrData = parseMasterQRCode(rawInput);
      if (!qrData) {
        triggerVibration([350]); sounds.error();
        setFeedback({ text: 'Errore: Formato QR Master non interpretabile.', type: 'error' });
        return;
      }

      let aggiunti = [];
      let duplicati = 0;
      let errati = 0;

      qrData.serials.forEach(serial => {
        const giaLetto = scannedSerials.some(s => s.serial === serial);
        if (giaLetto) { duplicati++; return; }
        if (!Object.hasOwn(expectedSerials, serial)) { errati++; return; }
        aggiunti.push({ po_line_key: activeLineKey, serial, model: expectedSerials[serial].model, pn: expectedSerials[serial].pn });
      });

      if (aggiunti.length > 0) {
        const newCartonsCount = cartonsScanned + 1;
        await supabase.from('scanned_serials').insert(aggiunti);
        await supabase.from('po_lines').update({ cartons_scanned: newCartonsCount }).eq('unique_key', activeLineKey);

        const updatedScanned = [
          ...aggiunti.map(a => ({ serial: a.serial, model: a.model, pn: a.pn, time: new Date().toLocaleTimeString('it-IT') })),
          ...scannedSerials
        ];
        setScannedSerials(updatedScanned);
        setCartonsScanned(newCartonsCount);
        if (updatedScanned.length >= totalExpected) {
          triggerVibration([150, 100, 150, 100, 200]); sounds.complete();
          setTimeout(() => openReviewSession(), 1200);
        } else {
          triggerVibration([150, 100, 150]); sounds.carton();
        }
        setFeedback({ text: `Cartone Rilevato! +${aggiunti.length} matricole acquisite.`, type: 'success' });
      } else {
        triggerVibration([300]); sounds.error();
        const details = [duplicati > 0 && `${duplicati} già lette`, errati > 0 && `${errati} non appartenenti`].filter(Boolean).join(', ');
        setFeedback({ text: `Cartone scartato (${details || 'nessuna matricola valida'}).`, type: 'error' });
      }

    } else {
      const serial = rawInput;
      const giaLetto = scannedSerials.some(s => s.serial === serial);
      if (giaLetto) {
        triggerVibration([100, 50, 100]); sounds.duplicate();
        setFeedback({ text: `Duplicato! Matricola ${serial} già letta.`, type: 'error' });
        return;
      }
      if (!Object.hasOwn(expectedSerials, serial)) {
        triggerVibration([300]); sounds.error();
        setFeedback({ text: `Errore! La matricola ${serial} non appartiene a questa riga.`, type: 'error' });
        return;
      }
      const meta = expectedSerials[serial];
      await supabase.from('scanned_serials').insert({ po_line_key: activeLineKey, serial, model: meta.model, pn: meta.pn });

      const updatedScanned = [
        { serial, model: meta.model, pn: meta.pn, time: new Date().toLocaleTimeString('it-IT') },
        ...scannedSerials
      ];
      setScannedSerials(updatedScanned);

      if (updatedScanned.length >= totalExpected) {
        triggerVibration([150, 100, 150, 100, 200]); sounds.complete();
        setFeedback({ text: `Completato! Ultima matricola: ${meta.model}`, type: 'success' });
        setTimeout(() => openReviewSession(), 1200);
      } else {
        triggerVibration([150]); sounds.ok();
        setFeedback({ text: `OK: Rilevato modello ${meta.model}`, type: 'success' });
      }
    }
  }

  function openReviewSession() {
    setCurrentView('review');
  }

  async function confirmAndFinalizeVerification() {
    setLoading(true);
    const { error } = await supabase
      .from('po_lines')
      .update({ is_user_confirmed: true })
      .eq('unique_key', activeLineKey);
    if (error) {
      alert("Errore nel salvataggio finale: " + error.message);
    } else {
      triggerVibration([100, 50, 100, 50, 200]);
      await fetchPOLines();
      setActiveLineKey(null);
      setActiveLine(null);
      setCurrentView('dashboard');
    }
    setLoading(false);
  }

  async function resetToDashboard() {
    setActiveLineKey(null);
    setActiveLine(null);
    setCurrentView('dashboard');
    await fetchPOLines();
  }

  const uniqueInvoices = [...new Set(poLines.map(item => item.china_invoice))].sort();
  const uniqueItems = [...new Set(poLines.map(item => item.item_code))].sort();

  const filteredLines = poLines.filter(line => {
    const matchInvoice = !filterInvoice || line.china_invoice === filterInvoice;
    const matchItem = !filterItem || line.item_code === filterItem;
    const snRequired = line.sn_required == null ? true : line.sn_required;
    const matchSN = (snRequired && filterSNYes) || (!snRequired && filterSNNo);
    return matchInvoice && matchItem && matchSN;
  });

  // Raggruppa per invoice + sn_required
  const invoiceGroups = {};
  filteredLines.forEach(line => {
    const snRequired = line.sn_required == null ? true : line.sn_required;
    const groupKey = `${line.china_invoice}__${snRequired ? 'yes' : 'no'}`;
    if (!invoiceGroups[groupKey]) invoiceGroups[groupKey] = { invoice: line.china_invoice, snRequired, lines: [] };
    invoiceGroups[groupKey].lines.push(line);
  });

  return (
    <div className="bg-gray-50 font-sans min-h-screen text-gray-800 flex flex-col justify-between">

      {/* Overlay menu */}
      {menuOpen && (
        <div className="fixed inset-0 z-50 flex">
          <div className="fixed inset-0 bg-black/40" onClick={() => setMenuOpen(false)} />
          <div className="relative z-10 w-72 bg-white h-full shadow-2xl flex flex-col">
            <div className="p-5 border-b border-gray-100 flex items-center justify-between">
              <img src="/logo.png" alt="Logo" className="h-10 w-auto object-contain" />
              <button onClick={() => setMenuOpen(false)} className="text-gray-400 hover:text-gray-700 text-2xl leading-none cursor-pointer">✕</button>
            </div>
            <nav className="flex-grow p-4 space-y-1">
              {[
                { id: 'arrivi', label: 'Piano Arrivi', icon: '📦' },
                { id: 'spare-parts', label: 'DB Spare Parts', icon: '🔧' },
                { id: 'stock', label: 'Stock', icon: '🗄️' },
              ].map(mod => (
                <button
                  key={mod.id}
                  onClick={() => { setActiveModule(mod.id); setMenuOpen(false); setCurrentView('dashboard'); }}
                  className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold transition cursor-pointer text-left ${activeModule === mod.id ? 'bg-blue-600 text-white' : 'text-gray-700 hover:bg-gray-100'}`}
                >
                  <span className="text-lg">{mod.icon}</span>
                  {mod.label}
                </button>
              ))}
            </nav>
          </div>
        </div>
      )}

      <header className="bg-white border-b border-gray-200 sticky top-0 z-40 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex justify-between items-center">
          <div className="flex items-center gap-3">
            <button onClick={() => setMenuOpen(true)} className="p-2 rounded-xl hover:bg-gray-100 transition cursor-pointer" aria-label="Menu">
              <div className="space-y-1.5">
                <span className="block w-6 h-0.5 bg-gray-700"></span>
                <span className="block w-6 h-0.5 bg-gray-700"></span>
                <span className="block w-6 h-0.5 bg-gray-700"></span>
              </div>
            </button>
            <img src="/logo.png" alt="Logo" className="h-12 sm:h-16 w-auto object-contain" />
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-gray-800 uppercase tracking-widest">
              {activeModule === 'arrivi' && 'Piano Arrivi'}
              {activeModule === 'spare-parts' && 'DB Spare Parts'}
              {activeModule === 'stock' && 'Stock'}
            </h1>
          </div>
          {activeModule === 'arrivi' && currentView === 'scan' && (
            <button className="bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs sm:text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer transition" onClick={resetToDashboard}>
              ← Torna alla Dashboard
            </button>
          )}
        </div>
      </header>

      <main className="w-full max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-grow">

        {loading && (
          <div className="text-center py-4 text-xs font-bold text-amber-600 animate-pulse bg-amber-50 border border-amber-100 rounded-xl mb-4">
            Sincronizzazione in corso con il Database Supabase...
          </div>
        )}

        {/* ==================== MODULI SPARE PARTS / STOCK ==================== */}
        {activeModule === 'spare-parts' && (
          <div className="flex flex-col items-center justify-center py-24 text-center space-y-3">
            <span className="text-5xl">🔧</span>
            <h2 className="text-xl font-black text-gray-700">DB Spare Parts</h2>
            <p className="text-sm text-gray-400">Modulo in sviluppo.</p>
          </div>
        )}
        {activeModule === 'stock' && (
          <div className="flex flex-col items-center justify-center py-24 text-center space-y-3">
            <span className="text-5xl">🗄️</span>
            <h2 className="text-xl font-black text-gray-700">Stock</h2>
            <p className="text-sm text-gray-400">Modulo in sviluppo.</p>
          </div>
        )}

        {/* ==================== MODULO PIANO ARRIVI ==================== */}
        {activeModule === 'arrivi' && (
        <>{/* ==================== VISTA 1: DASHBOARD ==================== */}
        {currentView === 'dashboard' && (
          <div className="space-y-6">

            {/* Box Caricamento File */}
            <div className={`bg-white p-4 sm:p-5 rounded-2xl border border-gray-200 shadow-xs ${poLines.length === 0 ? 'max-w-xl mx-auto my-8 text-center space-y-4' : ''}`}>
              {poLines.length === 0 && (
                <div className="space-y-1">
                  <h2 className="text-lg font-black text-gray-800">1. Carica Piano Arrivi</h2>
                  <p className="text-sm text-gray-500">Importa il file CSV generale contenente le righe di pianificazione merci.</p>
                </div>
              )}
              <div className={`border-2 border-dashed border-gray-300 rounded-xl p-5 bg-gray-50 hover:bg-gray-100 transition relative cursor-pointer group ${poLines.length > 0 ? 'flex items-center gap-3' : ''}`}>
                <input type="file" accept=".csv" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handlePOLinesUpload} />
                <div className={`pointer-events-none ${poLines.length > 0 ? 'flex items-center gap-3' : 'space-y-2 text-center'}`}>
                  <span className={`block group-hover:scale-110 transition-transform ${poLines.length > 0 ? 'text-xl' : 'text-3xl'}`}>📥</span>
                  <p className={`font-bold text-blue-600 ${poLines.length > 0 ? 'text-xs' : 'text-sm'}`}>
                    {poLines.length > 0 ? 'Ricarica / Aggiorna Piano Arrivi' : 'Seleziona il file del Piano Arrivi'}
                  </p>
                  {poLines.length > 0 && (
                    <span className="text-[10px] text-gray-400 font-medium">Le righe non più presenti verranno rimosse, quelle in corso bloccheranno l&apos;aggiornamento.</span>
                  )}
                </div>
              </div>
            </div>

            {poLines.length > 0 && (
              <>
                {/* Filtri */}
                <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Filtri di Visualizzazione</span>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div className="space-y-1">
                      <label className="block text-xs font-bold text-gray-500">Fattura Doganale (China Invoice):</label>
                      <select value={filterInvoice} onChange={e => setFilterInvoice(e.target.value)} className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs sm:text-sm focus:outline-hidden">
                        <option value="">Tutte le fatture ({uniqueInvoices.length})</option>
                        {uniqueInvoices.map(inv => <option key={inv} value={inv}>{inv}</option>)}
                      </select>
                    </div>
                    <div className="space-y-1">
                      <label className="block text-xs font-bold text-gray-500">Codice Prodotto (Item):</label>
                      <select value={filterItem} onChange={e => setFilterItem(e.target.value)} className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs sm:text-sm focus:outline-hidden">
                        <option value="">Tutti gli articoli ({uniqueItems.length})</option>
                        {uniqueItems.map(it => <option key={it} value={it}>{it}</option>)}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="block text-xs font-bold text-gray-500">Tracciamento Matricole (SN):</label>
                      <div className="flex gap-4">
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={filterSNYes} onChange={e => setFilterSNYes(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                          <span className="text-sm font-semibold text-gray-700">Serializzato</span>
                        </label>
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={filterSNNo} onChange={e => setFilterSNNo(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                          <span className="text-sm font-semibold text-gray-700">Solo quantità</span>
                        </label>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="flex justify-between items-center border-b border-gray-200 pb-2">
                  <div>
                    <h3 className="text-sm font-black text-gray-400 uppercase tracking-wider">Piani di Carico Rilevati</h3>
                    <p className="text-xs text-gray-500">I dati sono salvati in tempo reale nel cloud persistente.</p>
                  </div>
                  <span className="text-sm bg-blue-50 text-blue-600 font-black px-3 py-1 rounded-full border border-blue-100">
                    {filteredLines.length} Righe
                  </span>
                </div>

                <div className="space-y-8">
                  {Object.values(invoiceGroups).map(group => (
                    <div key={`${group.invoice}_${group.snRequired}`} className="bg-gray-100/60 p-4 sm:p-5 rounded-2xl border border-gray-200/80 space-y-3">
                      <div className="flex justify-between items-center border-b border-gray-200 pb-2">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-sm sm:text-base font-black text-gray-800 tracking-tight">📄 {group.invoice}</span>
                          {group.snRequired
                            ? <span className="text-[10px] bg-indigo-50 text-indigo-700 border border-indigo-200 font-bold px-2 py-0.5 rounded-md">Serializzato</span>
                            : <span className="text-[10px] bg-gray-100 text-gray-500 border border-gray-200 font-bold px-2 py-0.5 rounded-md">Solo quantità</span>
                          }
                        </div>
                      </div>
                      <div className="flex flex-col space-y-3">
                        {group.lines.map(item => {
                          const snRequired = item.sn_required == null ? true : item.sn_required;
                          const isConfirmed = item.is_user_confirmed === true;
                          return (
                            <div key={item.unique_key} className="bg-white p-4 rounded-xl border border-gray-200 hover:border-gray-300 transition shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">

                              {/* Data + stato */}
                              <div className="flex flex-row md:flex-col items-center md:items-start justify-between md:justify-center gap-2 md:w-48 shrink-0 border-b md:border-b-0 pb-2 md:pb-0 border-gray-100">
                                <span className="text-xs font-bold flex items-center gap-1 bg-amber-50 text-amber-800 border border-amber-200 px-2.5 py-1 rounded-lg">
                                  📅 {item.arrival_date}
                                </span>
                                {isConfirmed ? (
                                  <span className="text-xs bg-green-600 text-white px-3 py-1.5 rounded-lg font-black border border-green-700 shadow-xs">✓ Concluso</span>
                                ) : item.scanned_count > 0 ? (
                                  <span className="text-xs bg-amber-500 text-white px-3 py-1.5 rounded-lg font-black border border-amber-600 shadow-xs">⏳ In Corso</span>
                                ) : (
                                  <span className="text-xs bg-gray-200 text-gray-600 px-3 py-1.5 rounded-lg font-black border border-gray-300">◌ In Attesa</span>
                                )}
                              </div>

                              {/* Descrizione */}
                              <div className="flex-grow min-w-0 space-y-1">
                                <div className="flex flex-wrap items-center gap-2">
                                  <h4 className="text-xl font-black text-blue-700 font-mono tracking-tight break-all">{item.item_code}</h4>
                                  <span className="text-xs text-gray-400 font-mono hidden md:inline">(L: {item.line_id})</span>
                                </div>
                                <p className="text-sm text-gray-500 font-medium line-clamp-1">{item.description}</p>
                                <div className="flex flex-wrap gap-x-4 text-xs text-gray-400 font-semibold uppercase tracking-wider">
                                  <span>Rif. Ordine: <span className="text-gray-600 font-mono">{item.po_name}</span></span>
                                  {item.part_number && item.part_number !== 'N/D' && (
                                    <span>PN: <span className="text-indigo-600 font-mono font-bold">{item.part_number}</span></span>
                                  )}
                                </div>
                              </div>

                              {/* Livello 1+2+3 — larghezza fissa uniforme su desktop, full-width su mobile */}
                              <div className="w-full md:w-36 flex flex-col gap-2 pt-2 md:pt-0 border-t md:border-t-0 border-gray-100 shrink-0">

                                {/* Livello 1: Matricole attese */}
                                <div className="w-full bg-gray-900 text-white rounded-xl px-4 py-2 flex items-center justify-between md:flex-col md:items-center shadow-xs">
                                  <span className="text-[9px] uppercase font-bold tracking-wider text-gray-400 md:leading-tight">Attesi</span>
                                  <span className="text-2xl font-black font-mono text-amber-400 md:leading-tight">{item.qty_expected}</span>
                                  <span className="text-[9px] uppercase font-bold tracking-wider text-gray-500 md:leading-tight">pezzi</span>
                                </div>

                                {/* Livello 2: Matricole rilevate — solo se scansione avviata */}
                                {snRequired && item.sn_loaded && item.scanned_count > 0 && (
                                  <div className={`w-full rounded-xl px-4 py-2 flex items-center justify-between md:flex-col md:items-center shadow-xs border ${item.scanned_count >= item.qty_expected ? 'bg-green-50 border-green-200' : 'bg-blue-50 border-blue-200'}`}>
                                    <span className="text-[9px] uppercase font-bold tracking-wider text-gray-400 md:leading-tight">Rilevate</span>
                                    <span className={`text-2xl font-black font-mono md:leading-tight ${item.scanned_count >= item.qty_expected ? 'text-green-600' : 'text-blue-600'}`}>
                                      {item.scanned_count}
                                    </span>
                                    <span className="text-[9px] uppercase font-bold tracking-wider text-gray-400 md:leading-tight">/ {item.qty_expected}</span>
                                  </div>
                                )}

                                {/* Livello 3: Pulsanti azione */}
                                {snRequired && (
                                  <div className="flex flex-col gap-2">
                                    {!isConfirmed && (
                                      <div className="relative w-full">
                                        {item.sn_loaded ? (
                                          <>
                                            <button className="w-full bg-gray-200 hover:bg-gray-300 text-gray-500 text-xs font-medium py-2.5 rounded-xl transition cursor-pointer border border-gray-300 text-center">
                                              Sovrascrivi SN
                                            </button>
                                            <input type="file" accept=".xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={(e) => handleSNUpload(e, item)} />
                                          </>
                                        ) : (
                                          <>
                                            <button className="w-full bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold py-2.5 rounded-xl transition cursor-pointer shadow-xs text-center">
                                              Carica SN
                                            </button>
                                            <input type="file" accept=".xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={(e) => handleSNUpload(e, item)} />
                                          </>
                                        )}
                                      </div>
                                    )}
                                    <div className="w-full flex flex-col gap-2">
                                      {isConfirmed ? (
                                        <>
                                          <button onClick={() => downloadArrivoCSV(item)} className={`w-full text-xs font-bold py-2.5 rounded-xl transition cursor-pointer text-center ${downloadedKeys.has(item.unique_key) ? 'bg-gray-100 hover:bg-gray-200 text-gray-400 border border-gray-200' : 'bg-blue-600 hover:bg-blue-700 text-white shadow-xs'}`}>
                                            📥 Scarica Arrivo
                                          </button>
                                          <button onClick={() => reopenScanningSession(item)} className="w-full bg-gray-200 hover:bg-gray-300 text-gray-500 text-xs font-medium py-2.5 rounded-xl transition cursor-pointer border border-gray-300 text-center">
                                            Modifica / Rivedi
                                          </button>
                                        </>
                                      ) : (
                                        <button onClick={() => startScanningSession(item)} className="w-full bg-green-600 hover:bg-green-700 text-white text-xs font-bold py-2.5 rounded-xl transition cursor-pointer shadow-xs text-center">
                                          {item.scanned_count >= item.qty_expected && item.scanned_count > 0 ? 'Modifica / Rivedi' : 'Avvia Controllo'}
                                        </button>
                                      )}
                                    </div>
                                  </div>
                                )}

                              </div>

                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        )}

        {/* ==================== VISTA 2: SCANSIONE ==================== */}
        {currentView === 'scan' && activeLine && (
          <div className="grid grid-cols-1 md:grid-cols-12 gap-6 items-start">
            <div className="md:col-span-5 space-y-6">
              <div className="bg-white p-5 rounded-2xl shadow-xs border border-gray-200 space-y-4">
                <div>
                  <span className="text-[10px] sm:text-xs bg-blue-50 text-blue-700 font-bold px-2.5 py-1 rounded-md uppercase border border-blue-100">
                    China Invoice: {activeLine.china_invoice}
                  </span>
                  <h2 className="text-lg sm:text-xl font-black text-gray-800 mt-2.5">{activeLine.item_code} - Linea {activeLine.line_id}</h2>
                  <p className="text-xs sm:text-sm text-gray-500 font-medium mt-1">{activeLine.description}</p>
                </div>
                <div className="pt-3 border-t border-gray-100 space-y-3">
                  <div className="flex justify-between items-center">
                    <span className="text-xs sm:text-sm text-gray-400 font-bold">Progresso Scansione:</span>
                    <span className="text-xl sm:text-2xl font-black text-blue-600 font-mono">
                      {scannedSerials.length}/{activeLine.qty_expected}
                    </span>
                  </div>
                  <div className="w-full bg-gray-100 rounded-full h-3 overflow-hidden">
                    <div className="bg-blue-600 h-3 rounded-full transition-all duration-300"
                      style={{ width: `${activeLine.qty_expected > 0 ? Math.min((scannedSerials.length / activeLine.qty_expected) * 100, 100) : 0}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center bg-amber-50/70 px-3 py-2 rounded-xl border border-amber-100 mt-1">
                    <span className="text-xs text-amber-800 font-bold flex items-center gap-1.5">📦 Cartoni Caricati (QR Master):</span>
                    <span className="text-base font-black text-amber-700 font-mono bg-white px-2 py-0.5 rounded-md border border-amber-200/60 shadow-xs">{cartonsScanned}</span>
                  </div>
                </div>
              </div>

              <form onSubmit={handleScanSubmit} className="space-y-3">
                <label className="block text-xs sm:text-sm font-bold text-gray-500 uppercase tracking-wider">Input Terminale / Scanner:</label>
                <input
                  type="text"
                  ref={scannerInputRef}
                  value={scannerValue}
                  onChange={e => setScannerValue(e.target.value)}
                  autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
                  placeholder="Spara al seriale singolo o al QR Master del cartone..."
                  className="w-full bg-white border-2 border-blue-500 text-gray-800 font-mono text-base sm:text-xl p-4 rounded-xl shadow-inner focus:outline-hidden text-center"
                />
              </form>

              {feedback.text && (
                <div className={`p-4 rounded-xl text-center text-sm font-bold border ${feedback.type === 'success' ? 'bg-green-50 text-green-800 border-green-200' : 'bg-red-50 text-red-800 border-red-200'}`}>
                  {feedback.text}
                </div>
              )}

              {scannedSerials.length > 0 && scannedSerials.length >= activeLine.qty_expected && (
                <button
                  onClick={openReviewSession}
                  className="w-full bg-green-600 hover:bg-green-700 text-white font-black p-4 rounded-xl text-base shadow-md transition cursor-pointer flex items-center justify-center gap-2"
                >
                  ✓ Scansione completata — Procedi alla revisione
                </button>
              )}
            </div>

            <div className="md:col-span-7 bg-white p-5 rounded-2xl shadow-xs border border-gray-200 space-y-4">
              <div className="flex justify-between items-center border-b border-gray-100 pb-2">
                <h3 className="text-xs sm:text-sm font-bold text-gray-400 uppercase tracking-wider">Matricole Rilevate Con Conformità</h3>
                {scannedSerials.length > 0 && (
                  <button
                    onClick={deleteAllScannedSerials}
                    className="text-[11px] font-bold text-red-500 hover:text-red-700 hover:bg-red-50 border border-red-200 px-2.5 py-1 rounded-lg transition cursor-pointer"
                  >
                    🗑 Azzera tutte
                  </button>
                )}
              </div>
              <ul className="divide-y divide-gray-100 max-h-64 md:max-h-[500px] overflow-y-auto font-mono text-sm pr-1">
                {scannedSerials.map((s, index) => (
                  <li key={index} className="py-2.5 flex items-center gap-3 border-b border-gray-100 last:border-none group">
                    <div className="flex-grow min-w-0">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-gray-900 text-base">🟢 {s.serial}</span>
                        <span className="text-xs text-green-700 bg-green-50 font-bold px-2.5 py-0.5 rounded-full border border-green-100">OK</span>
                      </div>
                      <div className="text-xs text-gray-400 mt-1 flex gap-x-4">
                        <span><strong>Modello:</strong> {s.model}</span>
                        <span><strong>P/N:</strong> {s.pn}</span>
                        <span className="ml-auto text-[11px] font-medium text-gray-300">{s.time}</span>
                      </div>
                    </div>
                    <button
                      onClick={() => deleteScannedSerial(s.serial)}
                      className="shrink-0 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-lg p-1.5 transition cursor-pointer md:opacity-0 md:group-hover:opacity-100"
                      title="Elimina questa matricola"
                    >
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
            </div>
          </div>
        )}

        {/* ==================== VISTA 3: REVISIONE FINALE ==================== */}
        {currentView === 'review' && activeLine && (
          <div className="max-w-3xl mx-auto bg-white p-6 sm:p-8 rounded-2xl shadow-md border border-gray-200 space-y-6 my-4">
            <div className="border-b border-gray-200 pb-4 text-center sm:text-left">
              <span className="text-[10px] bg-amber-100 text-amber-800 border border-amber-200 font-black px-3 py-1 rounded-md uppercase tracking-wider">Fase di Controllo Finale</span>
              <h2 className="text-xl sm:text-2xl font-black text-gray-900 mt-2">{activeLine.item_code} - Linea {activeLine.line_id}</h2>
              <p className="text-sm text-gray-400 font-bold uppercase mt-0.5">China Invoice: {activeLine.china_invoice} | Arrivo: {activeLine.arrival_date}</p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div className="bg-blue-50 border border-blue-100 p-4 rounded-xl text-center">
                <span className="block text-xs font-bold text-blue-500 uppercase tracking-wider">Matricole Verificate</span>
                <span className="text-3xl font-black text-blue-700 font-mono">{scannedSerials.length} / {activeLine.qty_expected}</span>
              </div>
              <div className="bg-amber-50 border border-amber-100 p-4 rounded-xl text-center">
                <span className="block text-xs font-bold text-amber-600 uppercase tracking-wider">Cartoni Totali Caricati (QR)</span>
                <span className="text-3xl font-black text-amber-700 font-mono">{cartonsScanned}</span>
              </div>
            </div>

            <div className="space-y-2">
              <label className="block text-xs font-bold text-gray-400 uppercase tracking-wider">Anteprima Registro Scansioni:</label>
              <div className="border border-gray-200 rounded-xl overflow-hidden max-h-60 overflow-y-auto shadow-inner bg-gray-50">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-gray-100 text-[11px] font-bold text-gray-500 sticky top-0 border-b border-gray-200">
                    <tr>
                      <th className="p-2.5 pl-4">Matricola (SN)</th>
                      <th className="p-2.5">Modello</th>
                      <th className="p-2.5">Part Number</th>
                      <th className="p-2.5 pr-4 text-right">Ora</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 bg-white">
                    {scannedSerials.map((s, idx) => (
                      <tr key={idx} className="border-b border-gray-100 text-xs font-mono hover:bg-gray-50/80 transition">
                        <td className="p-2.5 pl-4 font-bold text-gray-900">🟢 {s.serial}</td>
                        <td className="p-2.5 text-gray-600">{s.model}</td>
                        <td className="p-2.5 text-gray-500">{s.pn}</td>
                        <td className="p-2.5 pr-4 text-right text-gray-400">{s.time}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-3 pt-2">
              <button onClick={confirmAndFinalizeVerification} className="w-full bg-green-600 hover:bg-green-700 text-white font-black p-4 rounded-xl text-base shadow-md transition cursor-pointer flex items-center justify-center gap-2">
                ✓ Approva e Registra Carico su Cloud
              </button>
            </div>
          </div>
        )}

        {/* ==================== VISTA 4: COMPLETAMENTO ==================== */}
        {currentView === 'complete' && activeLine && (
          <div className="max-w-xl mx-auto bg-white p-6 sm:p-8 rounded-2xl shadow-xs border border-gray-200 text-center space-y-6 my-12">
            <div className="w-14 h-14 bg-green-100 text-green-600 rounded-full flex items-center justify-center mx-auto text-3xl font-bold">✓</div>
            <div className="space-y-1">
              <h2 className="text-2xl font-black text-gray-800">Verifica Completata!</h2>
              <p className="text-sm text-gray-500">I dati sono salvati sul database centrale di Supabase e archiviati.</p>
            </div>
            <div className="flex flex-col gap-3">
              <button onClick={() => buildAndDownloadCSV(activeLine, scannedSerials)} className="w-full bg-green-600 hover:bg-green-700 text-white font-black p-4 rounded-xl shadow-md transition cursor-pointer flex items-center justify-center gap-2 text-base">
                📥 Scarica File CSV Finalizzato
              </button>
              <button onClick={resetToDashboard} className="w-full bg-gray-200 hover:bg-gray-300 text-gray-700 font-bold p-4 rounded-xl transition cursor-pointer text-sm">
                Torna alla Dashboard Generale
              </button>
            </div>
          </div>
        )}
        </>)}

      </main>

      <footer className="w-full text-center py-4 text-xs text-gray-400 border-t border-gray-200 bg-white">
        <div className="max-w-7xl mx-auto px-4">LogiScan Enterprise &copy; 2026 - Connessione Cloud Attiva</div>
      </footer>
    </div>
  );
}
