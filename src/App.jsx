import { useState, useEffect, useRef } from 'react';
import { supabase } from './supabaseClient';
import * as XLSX from 'xlsx';
import { useColumnWidths } from './useColumnWidths';
import JsBarcode from 'jsbarcode';

const SP_DEFAULT_WIDTHS = {
  pn: 130, terminal_pn: 120, modello: 60, pnit: 80, english_name: 110,
  descrizione: 175, type: 120, eol: 36, ref: 34, rplus: 34,
  to_order: 36, price: 50, locked: 34, edit: 56,
};

const STOCK_DEFAULT_WIDTHS = {
  locazione: 110, numero_bancale: 110, magazzino: 110,
  codice: 140, modello: 80, eol: 40, english_name: 155, descrizione: 190, stock: 68, edit: 60,
};

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
  const [currentUser, setCurrentUser] = useState(() => localStorage.getItem('logiscan_username') || '');

  const { widths: spWidths, startResize: spStartResize, resetWidths: spResetWidths } = useColumnWidths('logiscan_sp_cols', SP_DEFAULT_WIDTHS);
  const { widths: stWidths, startResize: stStartResize, resetWidths: stResetWidths } = useColumnWidths('logiscan_stock_cols', STOCK_DEFAULT_WIDTHS);
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

  // Filtri Piano Arrivi
  const [filterInvoice, setFilterInvoice] = useState('');
  const [filterItem, setFilterItem] = useState('');
  const [filterSNYes, setFilterSNYes] = useState(true);
  const [filterSNNo, setFilterSNNo] = useState(false);

  // DB Spare Parts
  const [spareParts, setSpareParts] = useState([]);
  const [spLoading, setSpLoading] = useState(false);
  const [spSearch, setSpSearch] = useState('');
  const [spFilterEOL, setSpFilterEOL] = useState('');
  const [spFilterType, setSpFilterType] = useState('');
  const [spFilterToOrder, setSpFilterToOrder] = useState('');
  const [spFilterTerminalPN, setSpFilterTerminalPN] = useState('');
  const [spFilterRef, setSpFilterRef] = useState('');
  const [spFilterRplus, setSpFilterRplus] = useState('');

  // Stock
  const [stockItems, setStockItems] = useState([]);
  const [stockLoading, setStockLoading] = useState(false);
  const [stockSearch, setStockSearch] = useState('');
  const [stockFilterMagazzino, setStockFilterMagazzino] = useState('');
  const [stockFilterLocazione, setStockFilterLocazione] = useState('');
  const [stockPage, setStockPage] = useState(0);
  const [stockSortCol, setStockSortCol] = useState('codice');
  const [stockSortDir, setStockSortDir] = useState('asc');
  const [stockFilterNoMatch, setStockFilterNoMatch] = useState(false);
  const [stockEditMode, setStockEditMode] = useState(false);
  const [stockSearch2, setStockSearch2] = useState('');
  const [stockPendingChanges, setStockPendingChanges] = useState({});
  const [moveBancaleOpen, setMoveBancaleOpen] = useState(false);

  // Arrivo Quantità
  const [, setArrivoQtyActive] = useState(false);
  const [arrivoQtyInvoice, setArrivoQtyInvoice] = useState('');
  const [arrivoQtyBancale, setArrivoQtyBancale] = useState('');
  const [arrivoQtyMagazzino, setArrivoQtyMagazzino] = useState('GESSATE');
  const [arrivoQtyCartoni, setArrivoQtyCartoni] = useState([]);
  const [arrivoQtyScanner, setArrivoQtyScanner] = useState('');
  const [arrivoQtyManuale, setArrivoQtyManuale] = useState({ codice: '', quantita: '' });
  const [arrivoQtyFeedback, setArrivoQtyFeedback] = useState({ text: '', type: '' });
  const arrivoQtyScannerRef = useRef(null);
  const [barcodeToPrint] = useState(null); // reserved for future print implementation
  const [printedCartons, setPrintedCartons] = useState(new Set());
  const [moveBancaleSrc, setMoveBancaleSrc] = useState('');
  const [moveBancaleDest, setMoveBancaleDest] = useState('GESSATE');
  const [moveBancaleLocazione, setMoveBancaleLocazione] = useState('');
  const [stockFilterBancale, setStockFilterBancale] = useState('');
  const [spSearch2, setSpSearch2] = useState('');
  const [spSortCol, setSpSortCol] = useState('pn');
  const [spSortDir, setSpSortDir] = useState('asc');
  const [spPage, setSpPage] = useState(0);
  const [spEditMode, setSpEditMode] = useState(false);
  const [spPendingChanges, setSpPendingChanges] = useState({});
  const [downloadedKeys, setDownloadedKeys] = useState(new Set());

  // Prelievi
  const [prelievoView, setPrelievoView] = useState('list'); // 'list' | 'new' | 'detail'
  const [prelieviList, setPrelieviList] = useState([]);
  const [prelievoDetail, setPrelievoDetail] = useState(null); // { testata, righe }
  const [prelieviLoading, setPrelieviLoading] = useState(false);
  const [prelievoUtente, setPrelievoUtente] = useState('');
  const [prelievoDest, setPrelievoDest] = useState('');
  const [prelievoRighe, setPrelievoRighe] = useState([]); // { stockId, idCartone, codice, numero_bancale, magazzino, quantita, qtaDisponibile }
  const [prelievoScanner, setPrelievoScanner] = useState('');
  const [prelievoFeedback, setPrelievoFeedback] = useState({ text: '', type: '' });
  const [prelievoManuale, setPrelievoManuale] = useState({ codice: '', stockId: '', quantita: '' });
  const [prelievoShowEsprinet, setPrelievoShowEsprinet] = useState(false);
  const prelievoScannerRef = useRef(null);

  const scannerInputRef = useRef(null);

  useEffect(() => {
    fetchPOLines();
    fetchSpareParts();
    fetchStock();
    fetchPrelievi();
  }, []);

  useEffect(() => {
    if (!barcodeToPrint) return;
    const el = document.getElementById('barcode-svg');
    if (!el) return;
    try {
      JsBarcode(el, barcodeToPrint.id, {
        format: 'CODE128', width: 2, height: 60,
        displayValue: true, fontSize: 14, margin: 8
      });
    } catch { /* ignore */ }
  }, [barcodeToPrint]);


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
    const [{ data, error }, { data: scannedKeys }, { data: cartonData }] = await Promise.all([
      supabase.from('po_lines').select('*').order('arrival_date', { ascending: true }),
      supabase.from('scanned_serials').select('po_line_key'),
      supabase.from('carton_arrivals').select('po_line_key, quantita, invoice, codice')
    ]);

    if (error) {
      alert("Errore nel caricamento dei dati: " + error.message);
    } else {
      const countMap = {};
      (scannedKeys || []).forEach(r => {
        countMap[r.po_line_key] = (countMap[r.po_line_key] || 0) + 1;
      });
      // Mappa quantità caricate: usa po_line_key se presente, altrimenti fallback su invoice+codice
      const cartonMap = {};
      const cartonByCode = {};
      (cartonData || []).forEach(r => {
        if (r.po_line_key) {
          cartonMap[r.po_line_key] = (cartonMap[r.po_line_key] || 0) + (r.quantita || 0);
        } else if (r.invoice && r.codice) {
          const k = `${r.invoice}__${r.codice}`;
          cartonByCode[k] = (cartonByCode[k] || 0) + (r.quantita || 0);
        }
      });
      setPoLines((data || []).map(l => ({
        ...l,
        scanned_count: countMap[l.unique_key] || 0,
        qty_loaded: (cartonMap[l.unique_key] || 0) +
          (cartonByCode[`${l.china_invoice}__${l.item_code}`] || 0) +
          (cartonByCode[`${l.china_invoice}__${l.part_number}`] || 0)
      })));
    }
    setLoading(false);
  }

  async function fetchSpareParts() {
    setSpLoading(true);
    const pageSize = 1000;
    let all = [];
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from('spare_parts')
        .select('*')
        .order('pn', { ascending: true })
        .range(from, from + pageSize - 1);
      if (error) { alert("Errore caricamento spare parts: " + error.message); break; }
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    setSpareParts(all);
    setSpLoading(false);
  }

  async function handleSparePartsUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    setSpLoading(true);

    const reader = new FileReader();
    reader.onload = async function(e) {
      let rows;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      } catch {
        alert("Errore: impossibile leggere il file Excel.");
        setSpLoading(false);
        return;
      }

      if (rows.length === 0 || !Object.hasOwn(rows[0], 'PN')) {
        alert("Formato non valido: colonna 'PN' non trovata.");
        setSpLoading(false);
        return;
      }

      const seen = new Set();
      const toUpsert = rows
        .filter(r => {
          const key = `${String(r['PN']).trim()}__${String(r['Terminal PN']).trim()}`;
          if (!String(r['PN']).trim() || seen.has(key)) return false;
          seen.add(key);
          return true;
        })
        .map(r => ({
          pn:          String(r['PN'] || '').trim(),
          terminal_pn:  String(r['Terminal PN'] || '').trim(),
          eol:          String(r['EOL'] || '').trim(),
          english_name: String(r['English name'] || '').trim(),
          type:         String(r['TYPE'] || '').trim(),
          to_order:     String(r['TO ORDER'] || '').trim(),
          qty:          parseFloat(r['Qty.']) || 0,
          descrizione:  String(r['DESCRIZIONE'] || '').trim(),
          pnit:         String(r['PNIT'] || '').trim(),
          ref:          String(r['REF'] || '').trim(),
          rplus:        String(r['RPLUS'] || '').trim(),
          price:        parseFloat(r['$']) || 0,
          locked: (['EOL','CLI'].includes(String(r['EOL']||'').trim()) || String(r['TO ORDER']||'').trim() === 'NO' || String(r['TYPE']||'').trim() === 'NO') ? 'Y' : '',
          modified_by:  currentUser || 'import',
        }));

      const { error } = await supabase.from('spare_parts').upsert(toUpsert, { onConflict: 'pn,terminal_pn' });
      if (error) { alert("Errore salvataggio: " + error.message); }
      else { await fetchSpareParts(); alert(`${toUpsert.length} ricambi caricati.`); }
    };
    reader.readAsArrayBuffer(file);
  }

  function setSpFieldChange(rowKey, field, value) {
    setSpPendingChanges(prev => ({ ...prev, [rowKey]: { ...(prev[rowKey] || {}), [field]: value } }));
  }

  async function saveAllSpChanges() {
    let user = currentUser;
    if (!user) {
      const name = window.prompt('Inserisci il tuo nome utente (verrà salvato):');
      if (!name) return;
      localStorage.setItem('logiscan_username', name);
      setCurrentUser(name);
      user = name;
    }
    const entries = Object.entries(spPendingChanges);
    if (!entries.length) return;
    setSpLoading(true);
    const errors = [];
    for (const [rowKey, changes] of entries) {
      const original = spareParts.find(p => `${p.pn}_${p.terminal_pn}` === rowKey);
      const merged = { ...original, ...changes };
      const pnit = (merged.pnit || '').trim();
      const type = (merged.type || '').trim();
      const eol = (merged.eol || '').trim();
      const to_order = (merged.to_order || '').trim();
      const codice = [pnit, type].filter(Boolean).join('');
      const locked = (eol === 'EOL' || eol === 'CLI' || to_order === 'NO' || type === 'NO') ? 'Y' : '';
      const { error } = await supabase.from('spare_parts').update({
        eol, english_name: merged.english_name, type, to_order,
        qty: parseFloat(merged.qty) || 0, descrizione: merged.descrizione,
        pnit, ref: merged.ref, rplus: merged.rplus,
        price: parseFloat(merged.price) || 0, locked, codice, modified_by: user,
      }).eq('pn', original.pn).eq('terminal_pn', original.terminal_pn);
      if (error) errors.push(error.message);
      else setSpareParts(prev => prev.map(p => `${p.pn}_${p.terminal_pn}` === rowKey ? { ...p, ...merged, codice, locked, modified_by: user } : p));
    }
    if (errors.length) alert('Errori:\n' + errors.join('\n'));
    setSpPendingChanges({});
    setSpEditMode(false);
    setSpLoading(false);
  }

  async function fetchStock() {
    setStockLoading(true);
    const pageSize = 1000;
    let all = [];
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from('stock_inventory')
        .select('*')
        .order('codice', { ascending: true })
        .range(from, from + pageSize - 1);
      if (error) { alert("Errore caricamento stock: " + error.message); break; }
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    setStockItems(all);
    setStockLoading(false);
  }

  async function handleStockUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    setStockLoading(true);

    const reader = new FileReader();
    reader.onload = async function(e) {
      let rows;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        // Foglio 2
        const sheetName = wb.SheetNames[1] || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      } catch {
        alert("Errore: impossibile leggere il file Excel.");
        setStockLoading(false);
        return;
      }

      if (rows.length === 0) {
        alert("Nessun dato trovato nel foglio.");
        setStockLoading(false);
        return;
      }

      const seen = new Set();
      const allRows = rows
        .filter(r => {
          const codice = String(r['CODICE'] || '').trim();
          if (!codice) return false;
          const key = `${codice}__${String(r['MAGAZZINO']||'').trim()}__${String(r['NUMERO BANCALE']||'').trim()}`;
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        })
        .map(r => ({
          locazione:      String(r['LOCAZIONE'] || '').trim(),
          numero_bancale: String(r['NUMERO BANCALE'] || '').trim(),
          magazzino:      String(r['MAGAZZINO'] || '').trim(),
          codice:         String(r['CODICE'] || '').trim(),
          stock:          parseFloat(r['STOCK']) || 0,
        }));

      // Righe con stock > 0: upsert; righe con stock = 0: elimina dal DB
      const toUpsert = allRows.filter(r => r.stock !== 0);
      const toDelete = allRows.filter(r => r.stock === 0);

      const { error } = await supabase
        .from('stock_inventory')
        .upsert(toUpsert, { onConflict: 'codice,magazzino,numero_bancale' });

      if (toDelete.length > 0) {
        for (const r of toDelete) {
          await supabase.from('stock_inventory').delete()
            .eq('codice', r.codice).eq('magazzino', r.magazzino).eq('numero_bancale', r.numero_bancale);
        }
      }

      if (error) { alert("Errore salvataggio stock: " + error.message); }
      else { await fetchStock(); alert(`${toUpsert.length} record caricati, ${toDelete.length} a zero rimossi.`); }
    };
    reader.readAsArrayBuffer(file);
  }


  function setStockFieldChange(id, field, value) {
    setStockPendingChanges(prev => ({
      ...prev,
      [id]: { ...(prev[id] || {}), [field]: value }
    }));
  }

  async function moveBancale() {
    if (!moveBancaleSrc) { alert('Seleziona il bancale da spostare.'); return; }
    const toMove = stockItems.filter(s => s.numero_bancale === moveBancaleSrc);
    if (toMove.length === 0) { alert('Nessuna riga trovata per questo bancale.'); return; }
    const updates = {};
    updates.magazzino = moveBancaleDest;
    if (moveBancaleLocazione.trim()) updates.locazione = moveBancaleLocazione.trim();
    const lines = [`Bancale: ${moveBancaleSrc} (${toMove.length} righe)`, `Magazzino: ${toMove[0]?.magazzino || '?'} → ${moveBancaleDest}`];
    if (updates.locazione) lines.push(`Locazione: ${toMove[0]?.locazione || '—'} → ${updates.locazione}`);
    if (!window.confirm(`Spostare bancale?\n\n${lines.join('\n')}`)) return;
    setStockLoading(true);
    const { error } = await supabase.from('stock_inventory').update(updates).eq('numero_bancale', moveBancaleSrc);
    if (error) { alert('Errore: ' + error.message); }
    else {
      setStockItems(prev => prev.map(s => s.numero_bancale === moveBancaleSrc ? { ...s, ...updates } : s));
      setMoveBancaleOpen(false);
      setMoveBancaleSrc('');
      setMoveBancaleLocazione('');
    }
    setStockLoading(false);
  }

  function parseCartonQR(raw) {
    // Format: >{RS}06{GS}CODICE{GS}..{GS}QUANTITA{GS}..{RS}{Eot}YZ
    // eslint-disable-next-line no-control-regex
    let s = raw.replace(/\x1e/g, '{RS}').replace(/\x1d/g, '{GS}').replace(/\x04/g, '{Eot}');
    s = s.replace(/^[>{].*?{RS}/, '').replace(/{RS}.*$/, '').replace(/{Eot}.*$/, '').trim();
    const parts = s.split('{GS}').map(p => p.trim()).filter(Boolean);
    if (parts.length < 5) return null;
    const codice = parts[1];
    const quantita = parseFloat(parts[4]);
    if (!codice || isNaN(quantita)) return null;
    const idCartone = parts.join('-');
    return { codice, quantita, idCartone, qrRaw: raw };
  }

  function checkCodiceInPianoArrivi(codice) {
    return poLines.find(l =>
      l.china_invoice === arrivoQtyInvoice &&
      l.sn_required === false &&
      (l.item_code === codice || l.part_number === codice)
    ) || null;
  }

  async function handleArrivoQtySubmit(e) {
    e.preventDefault();
    const raw = (arrivoQtyScannerRef.current?.value || arrivoQtyScanner).trim();
    setArrivoQtyScanner('');
    if (arrivoQtyScannerRef.current) arrivoQtyScannerRef.current.value = '';
    if (!raw) return;

    if (!arrivoQtyBancale.trim()) {
      setArrivoQtyFeedback({ text: 'Inserisci il nome del bancale prima di scansionare.', type: 'error' });
      return;
    }
    const parsed = parseCartonQR(raw);
    if (!parsed) {
      setArrivoQtyFeedback({ text: 'QR non riconosciuto. Usa inserimento manuale.', type: 'error' });
      sounds.error(); triggerVibration([300]);
      return;
    }
    if (!checkCodiceInPianoArrivi(parsed.codice)) {
      setArrivoQtyFeedback({ text: `Codice ${parsed.codice} non previsto in arrivo per invoice ${arrivoQtyInvoice}. Bloccato.`, type: 'error' });
      sounds.error(); triggerVibration([300]);
      return;
    }
    const duplicate = arrivoQtyCartoni.find(c => c.idCartone === raw);
    if (duplicate) {
      setArrivoQtyFeedback({ text: 'Duplicato! Cartone già rilevato.', type: 'error' });
      sounds.duplicate(); triggerVibration([100, 50, 100]);
      return;
    }
    const matchedLine = checkCodiceInPianoArrivi(parsed.codice);
    const record = { id_cartone: raw, invoice: arrivoQtyInvoice, bancale: arrivoQtyBancale, magazzino: arrivoQtyMagazzino, codice: parsed.codice, quantita: parsed.quantita, qr_raw: raw, stato: 'pending', po_line_key: matchedLine?.unique_key || null };
    const { error: insErr } = await supabase.from('carton_arrivals').insert(record);
    if (insErr) { setArrivoQtyFeedback({ text: 'Errore DB: ' + insErr.message, type: 'error' }); return; }
    setArrivoQtyCartoni(prev => [{ ...parsed, idCartone: raw, bancale: arrivoQtyBancale, manuale: false }, ...prev]);
    // Aggiorna qty_loaded direttamente in memoria
    if (matchedLine) {
      setPoLines(prev => prev.map(l => l.unique_key === matchedLine.unique_key ? { ...l, qty_loaded: (l.qty_loaded || 0) + parsed.quantita } : l));
    }
    setArrivoQtyFeedback({ text: `OK: ${parsed.codice} — qtà ${parsed.quantita}`, type: 'success' });
    sounds.ok(); triggerVibration([150]);
  }

  async function addCartonManuale() {
    const { codice, quantita } = arrivoQtyManuale;
    if (!arrivoQtyBancale.trim()) { alert('Inserisci il nome del bancale prima di aggiungere cartoni.'); return; }
    if (!codice.trim() || !quantita) { alert('Inserisci codice e quantità.'); return; }
    if (!checkCodiceInPianoArrivi(codice.trim())) {
      alert(`Il codice "${codice.trim()}" non è previsto in arrivo per invoice ${arrivoQtyInvoice}.\n\nOperazione bloccata.`);
      return;
    }
    const matchedLine2 = checkCodiceInPianoArrivi(codice.trim());
    const qty = parseFloat(quantita);

    // Genera id_cartone formato AAMMGG-X
    const now = new Date();
    const datePrefix = String(now.getFullYear()).slice(2).padStart(2,'0')
      + String(now.getMonth() + 1).padStart(2,'0')
      + String(now.getDate()).padStart(2,'0');
    const { count } = await supabase.from('carton_arrivals')
      .select('*', { count: 'exact', head: true })
      .like('id_cartone', `${datePrefix}-%`);
    const manualId = `${datePrefix}-${(count || 0) + 1}`;

    const record = { id_cartone: manualId, invoice: arrivoQtyInvoice, bancale: arrivoQtyBancale, magazzino: arrivoQtyMagazzino, codice: codice.trim(), quantita: qty, qr_raw: null, stato: 'pending', po_line_key: matchedLine2?.unique_key || null };
    const { error: insErr2 } = await supabase.from('carton_arrivals').insert(record);
    if (insErr2) { alert('Errore DB: ' + insErr2.message); return; }
    const dbId = manualId;
    setArrivoQtyCartoni(prev => [{ codice: codice.trim(), quantita: qty, idCartone: dbId, qrRaw: null, bancale: arrivoQtyBancale, manuale: true }, ...prev]);
    if (matchedLine2) {
      setPoLines(prev => prev.map(l => l.unique_key === matchedLine2.unique_key ? { ...l, qty_loaded: (l.qty_loaded || 0) + qty } : l));
    }
    setArrivoQtyManuale({ codice: '', quantita: '' });
    setArrivoQtyFeedback({ text: `Aggiunto: ${codice.trim()} — qtà ${quantita}`, type: 'success' });
  }

  async function removeCarton(idCartone) {
    await supabase.from('carton_arrivals').delete().eq('id_cartone', idCartone).eq('stato', 'pending');
    setArrivoQtyCartoni(prev => prev.filter(c => c.idCartone !== idCartone));
    await fetchPOLines();
  }

  function printCartonLabel(carton) {
    setPrintedCartons(prev => new Set(prev).add(carton.idCartone));
    // Genera SVG barcode in memoria
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    JsBarcode(svg, carton.idCartone, {
      format: 'CODE128', width: 1.5, height: 28,
      displayValue: true, fontSize: 8, margin: 2,
      textMargin: 1
    });
    const svgData = new XMLSerializer().serializeToString(svg);

    const win = window.open('', '_blank', 'width=300,height=200');
    win.document.write(`<!DOCTYPE html><html><head><title>Etichetta</title>
      <style>
        @page { size: 51mm 19mm; margin: 0; }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { width: 51mm; height: 19mm; display: flex; flex-direction: column; align-items: center; justify-content: center; font-family: monospace; overflow: hidden; }
        svg { width: 47mm; height: 14mm; }
        .info { font-size: 5pt; text-align: center; line-height: 1.2; width: 47mm; }
      </style>
    </head><body>
      ${svgData}
      <div class="info">${carton.codice} — Qtà: ${carton.quantita}</div>
    </body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 300);
  }

  async function annullaTuttoArrivo() {
    const total = arrivoQtyCartoni.length;
    if (!window.confirm(`Annullare tutti i ${total} cartoni per invoice ${arrivoQtyInvoice}?\n\nVerranno rimossi dall'inventario e dal piano arrivi.`)) return;

    // 1. Raccoglie gli id_cartone per pulire stock_inventory
    const cartoneIds = arrivoQtyCartoni.map(c => c.idCartone).filter(Boolean);

    // 2. Elimina tutti i carton_arrivals dell'invoice (pending + caricato)
    await supabase.from('carton_arrivals').delete().eq('invoice', arrivoQtyInvoice);

    // 3. Rimuove le righe stock_inventory create da questi cartoni
    for (const cid of cartoneIds) {
      await supabase.from('stock_inventory').delete().contains('carton_ids', [cid]);
    }

    // 4. Resetta is_user_confirmed sulle righe non serializzate
    await supabase.from('po_lines')
      .update({ is_user_confirmed: false })
      .eq('china_invoice', arrivoQtyInvoice)
      .eq('sn_required', false);

    setArrivoQtyCartoni([]);
    setArrivoQtyFeedback({ text: '', type: '' });
    setCurrentView('dashboard');
    setArrivoQtyActive(false);
    await fetchPOLines();
    await fetchStock();
  }

  async function caricaArrivoSuInventario() {
    if (arrivoQtyCartoni.length === 0) { alert('Nessun cartone da caricare.'); return; }
    if (!window.confirm(`Caricare ${arrivoQtyCartoni.length} cartoni sull'Inventario Spare Parts?\n\nInvoice: ${arrivoQtyInvoice}\nMagazzino: ${arrivoQtyMagazzino}`)) return;

    setLoading(true);
    const errors = [];

    // 1. Aggiorna pending → caricato
    const { error: errArr } = await supabase.from('carton_arrivals')
      .update({ stato: 'caricato' })
      .eq('invoice', arrivoQtyInvoice)
      .eq('stato', 'pending');
    if (errArr) { errors.push('carton_arrivals: ' + errArr.message); }

    // Marca le righe non serializzate dell'invoice come confermate
    await supabase.from('po_lines')
      .update({ is_user_confirmed: true })
      .eq('china_invoice', arrivoQtyInvoice)
      .eq('sn_required', false);

    // 2. Inserisci una riga per cartone in stock_inventory
    for (const c of arrivoQtyCartoni) {
      const { error } = await supabase.from('stock_inventory').insert({
        codice:         c.codice,
        stock:          c.quantita,
        numero_bancale: c.bancale || arrivoQtyBancale,
        magazzino:      arrivoQtyMagazzino,
        carton_ids:     c.idCartone ? [c.idCartone] : [],
        locazione:      '',
      });
      if (error) errors.push(`${c.codice}: ${error.message}`);
    }

    if (errors.length) {
      alert('Completato con errori:\n' + errors.join('\n'));
    } else {
      alert(`${arrivoQtyCartoni.length} cartoni caricati con successo!`);
      setArrivoQtyCartoni([]);
      setArrivoQtyActive(false);
      setArrivoQtyInvoice('');
      setArrivoQtyBancale('');
      await fetchStock();
    }
    setLoading(false);
  }

  function handleSpKeyNav(e, rowIndex, field) {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    e.preventDefault();
    const next = e.key === 'ArrowDown' ? rowIndex + 1 : rowIndex - 1;
    const target = document.querySelector(`[data-sp-row="${next}"][data-sp-field="${field}"]`);
    if (target) target.focus();
  }

  // ==================== PRELIEVI ====================
  async function fetchPrelievi() {
    setPrelieviLoading(true);
    const [{ data: testate, error }, { data: righe }] = await Promise.all([
      supabase.from('prelievi').select('*').order('data_prelievo', { ascending: false }),
      supabase.from('prelievi_righe').select('prelievo_id, quantita')
    ]);
    if (error) { alert('Errore caricamento prelievi: ' + error.message); setPrelieviLoading(false); return; }
    const agg = {};
    (righe || []).forEach(r => {
      if (!agg[r.prelievo_id]) agg[r.prelievo_id] = { righe: 0, pezzi: 0 };
      agg[r.prelievo_id].righe += 1;
      agg[r.prelievo_id].pezzi += (r.quantita || 0);
    });
    setPrelieviList((testate || []).map(t => ({ ...t, n_righe: agg[t.id]?.righe || 0, n_pezzi: agg[t.id]?.pezzi || 0 })));
    setPrelieviLoading(false);
  }

  async function openPrelievoDetail(prelievo) {
    setPrelieviLoading(true);
    const { data: righe, error } = await supabase.from('prelievi_righe')
      .select('*').eq('prelievo_id', prelievo.id).order('id', { ascending: true });
    if (error) { alert('Errore: ' + error.message); setPrelieviLoading(false); return; }
    setPrelievoDetail({ testata: prelievo, righe: righe || [] });
    setPrelievoView('detail');
    setPrelieviLoading(false);
  }

  async function deletePrelievo(prelievo) {
    if (!window.confirm(`Eliminare il prelievo ${prelievo.id_prelievo}?\n\nLe ${prelievo.n_righe} righe verranno ripristinate nell'inventario.`)) return;
    setPrelieviLoading(true);
    const errors = [];

    // 1. Recupera le righe del prelievo
    const { data: righe, error: errR } = await supabase.from('prelievi_righe').select('*').eq('prelievo_id', prelievo.id);
    if (errR) { alert('Errore: ' + errR.message); setPrelieviLoading(false); return; }

    // 2. Ripristina lo stock per ogni riga
    for (const r of (righe || [])) {
      const { data: existing } = await supabase.from('stock_inventory').select('id, stock').eq('id', r.stock_id).maybeSingle();
      if (existing) {
        const { error } = await supabase.from('stock_inventory').update({ stock: (existing.stock || 0) + (r.quantita || 0) }).eq('id', r.stock_id);
        if (error) errors.push(error.message);
      } else {
        const { error } = await supabase.from('stock_inventory').insert({
          codice: r.codice, stock: r.quantita, numero_bancale: r.numero_bancale,
          magazzino: r.magazzino, carton_ids: r.id_cartone ? [r.id_cartone] : [], locazione: '',
        });
        if (error) errors.push(error.message);
      }
    }

    // 3. Elimina righe e testata
    await supabase.from('prelievi_righe').delete().eq('prelievo_id', prelievo.id);
    await supabase.from('prelievi').delete().eq('id', prelievo.id);

    if (errors.length) alert('Completato con errori:\n' + errors.join('\n'));
    await fetchPrelievi();
    await fetchStock();
    setPrelieviLoading(false);
  }

  function addPrelievoRiga(stockRow, quantita) {
    if (prelievoRighe.some(r => r.stockId === stockRow.id)) {
      setPrelievoFeedback({ text: 'Questa riga di stock è già nel prelievo.', type: 'error' });
      sounds.duplicate(); triggerVibration([100, 50, 100]);
      return false;
    }
    const qta = parseFloat(quantita);
    if (!qta || qta <= 0) { setPrelievoFeedback({ text: 'Quantità non valida.', type: 'error' }); return false; }
    if (qta > stockRow.stock) {
      setPrelievoFeedback({ text: `Quantità ${qta} superiore alla disponibilità (${stockRow.stock}).`, type: 'error' });
      sounds.error(); triggerVibration([300]);
      return false;
    }
    setPrelievoRighe(prev => [{
      stockId: stockRow.id,
      idCartone: (stockRow.carton_ids && stockRow.carton_ids[0]) || null,
      codice: stockRow.codice,
      numero_bancale: stockRow.numero_bancale,
      magazzino: stockRow.magazzino,
      quantita: qta,
      qtaDisponibile: stockRow.stock,
    }, ...prev]);
    sounds.ok(); triggerVibration([150]);
    return true;
  }

  function handlePrelievoScan(e) {
    e.preventDefault();
    const raw = (prelievoScannerRef.current?.value || prelievoScanner).trim();
    setPrelievoScanner('');
    if (prelievoScannerRef.current) prelievoScannerRef.current.value = '';
    if (!raw) return;

    // Cerca riga stock per id_cartone (sia QR grezzo che id manuale)
    const stockRow = stockItems.find(s => (s.carton_ids || []).includes(raw));
    if (!stockRow) {
      setPrelievoFeedback({ text: `Nessun cartone trovato in inventario con ID: ${raw}`, type: 'error' });
      sounds.error(); triggerVibration([300]);
      return;
    }
    if (addPrelievoRiga(stockRow, stockRow.stock)) {
      setPrelievoFeedback({ text: `OK: ${stockRow.codice} — disp. ${stockRow.stock} (modificabile)`, type: 'success' });
    }
  }

  function addPrelievoManuale() {
    const { stockId, quantita } = prelievoManuale;
    if (!stockId) { alert('Seleziona la riga di stock (bancale).'); return; }
    const stockRow = stockItems.find(s => String(s.id) === String(stockId));
    if (!stockRow) { alert('Riga di stock non trovata.'); return; }
    if (addPrelievoRiga(stockRow, quantita || stockRow.stock)) {
      setPrelievoManuale({ codice: '', stockId: '', quantita: '' });
      setPrelievoFeedback({ text: `Aggiunto: ${stockRow.codice} — qtà ${quantita || stockRow.stock}`, type: 'success' });
    }
  }

  function removePrelievoRiga(stockId) {
    setPrelievoRighe(prev => prev.filter(r => r.stockId !== stockId));
  }

  function updatePrelievoQta(stockId, value) {
    setPrelievoRighe(prev => prev.map(r => r.stockId === stockId ? { ...r, quantita: value } : r));
  }

  async function registraPrelievo() {
    if (prelievoRighe.length === 0) { alert('Nessuna riga da prelevare.'); return; }
    if (!prelievoDest) { alert('Seleziona la destinazione (Secure Room o Repair).'); return; }
    // Valida quantità
    for (const r of prelievoRighe) {
      const q = parseFloat(r.quantita);
      if (!q || q <= 0 || q > r.qtaDisponibile) {
        alert(`Quantità non valida per ${r.codice} (max ${r.qtaDisponibile}).`); return;
      }
    }
    let user = prelievoUtente.trim() || currentUser;
    if (!user) {
      const name = window.prompt('Inserisci il tuo nome utente:');
      if (!name) return;
      localStorage.setItem('logiscan_username', name);
      setCurrentUser(name);
      user = name;
    }
    if (!window.confirm(`Registrare il prelievo di ${prelievoRighe.length} righe?\n\nLe giacenze verranno decurtate dall'inventario.`)) return;

    setLoading(true);
    const errors = [];

    // Genera id_prelievo AAMMGG-X
    const now = new Date();
    const datePrefix = String(now.getFullYear()).slice(2).padStart(2,'0')
      + String(now.getMonth() + 1).padStart(2,'0')
      + String(now.getDate()).padStart(2,'0');
    const { count } = await supabase.from('prelievi')
      .select('*', { count: 'exact', head: true })
      .like('id_prelievo', `${datePrefix}-%`);
    const idPrelievo = `${datePrefix}-${(count || 0) + 1}`;

    // 1. Crea testata prelievo
    const { data: testata, error: errT } = await supabase.from('prelievi')
      .insert({ id_prelievo: idPrelievo, utente: user, destinazione: prelievoDest.trim() || null })
      .select('id').single();
    if (errT) { alert('Errore creazione prelievo: ' + errT.message); setLoading(false); return; }

    // 2. Righe + decurtazione stock
    for (const r of prelievoRighe) {
      const q = parseFloat(r.quantita);
      await supabase.from('prelievi_righe').insert({
        prelievo_id: testata.id, stock_id: r.stockId, id_cartone: r.idCartone,
        codice: r.codice, numero_bancale: r.numero_bancale, magazzino: r.magazzino, quantita: q,
      });
      const nuovoStock = r.qtaDisponibile - q;
      if (nuovoStock <= 0) {
        const { error } = await supabase.from('stock_inventory').delete().eq('id', r.stockId);
        if (error) errors.push(`${r.codice}: ${error.message}`);
      } else {
        const { error } = await supabase.from('stock_inventory').update({ stock: nuovoStock }).eq('id', r.stockId);
        if (error) errors.push(`${r.codice}: ${error.message}`);
      }
    }

    if (errors.length) alert('Completato con errori:\n' + errors.join('\n'));
    else alert(`Prelievo ${idPrelievo} registrato (${prelievoRighe.length} righe).`);
    setPrelievoRighe([]);
    setPrelievoDest('');
    setPrelievoFeedback({ text: '', type: '' });
    setPrelievoView('list');
    await fetchStock();
    await fetchPrelievi();
    setLoading(false);
  }

  function handleStockKeyNav(e, rowIndex, field) {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    e.preventDefault();
    const next = e.key === 'ArrowDown' ? rowIndex + 1 : rowIndex - 1;
    const target = document.querySelector(`[data-rowindex="${next}"][data-field="${field}"]`);
    if (target) target.focus();
  }

  async function saveAllStockChanges() {
    const entries = Object.entries(stockPendingChanges);
    if (entries.length === 0) return;
    setStockLoading(true);
    let errors = [];
    for (const [id, changes] of entries) {
      const original = stockItems.find(s => String(s.id) === String(id));
      const merged = { ...original, ...changes, stock: parseFloat(changes.stock ?? original?.stock) || 0 };
      if (!merged.magazzino) { errors.push(`Riga ${merged.codice}: Magazzino obbligatorio.`); continue; }
      const { error } = await supabase
        .from('stock_inventory')
        .update({ locazione: merged.locazione, numero_bancale: merged.numero_bancale, magazzino: merged.magazzino, codice: merged.codice, stock: merged.stock })
        .eq('id', Number(id));
      if (error) errors.push(error.message);
    }
    if (errors.length > 0) { alert("Errori: " + errors.join('\n')); }
    else {
      setStockItems(prev => prev.map(s => {
        const changes = stockPendingChanges[s.id];
        if (!changes) return s;
        return { ...s, ...changes, stock: parseFloat(changes.stock ?? s.stock) || 0 };
      }));
      setStockPendingChanges({});
      setStockEditMode(false);
    }
    setStockLoading(false);
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

      const { data: existingLines } = await supabase.from('po_lines').select('unique_key, is_user_confirmed, item_code, line_id, sn_required');
      const existing = existingLines || [];

      // Lines not present in the new CSV that must be removed
      const toRemove = existing.filter(l => !newKeys.has(l.unique_key));

      // Una riga blocca la rimozione SOLO se ha lavoro effettivo in corso:
      // serializzati con matricole scansionate, oppure non serializzati con cartoni rilevati
      const blocked = [];
      for (const l of toRemove) {
        if (l.is_user_confirmed) continue; // già conclusa → si può rimuovere
        const table = l.sn_required === false ? 'carton_arrivals' : 'scanned_serials';
        const { count } = await supabase.from(table)
          .select('*', { count: 'exact', head: true }).eq('po_line_key', l.unique_key);
        if ((count || 0) > 0) blocked.push(l);
      }

      if (blocked.length > 0) {
        const names = blocked.map(l => `${l.item_code} (Linea ${l.line_id})`).join('\n');
        alert(`Impossibile aggiornare: le seguenti righe hanno rilevazioni in corso e non possono essere rimosse:\n\n${names}\n\nConcludi o annulla le rilevazioni prima di ricaricare il file.`);
        setLoading(false);
        return;
      }

      // Delete confirmed lines not in the new CSV, with all their references
      for (const line of toRemove) {
        await supabase.from('expected_serials').delete().eq('po_line_key', line.unique_key);
        await supabase.from('scanned_serials').delete().eq('po_line_key', line.unique_key);
        // Per i non serializzati: pulisci anche carton_arrivals per quella invoice
        if (line.sn_required === false) {
          await supabase.from('carton_arrivals').delete().eq('po_line_key', line.unique_key);
        }
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
        snRecords = XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false, raw: false });
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

      // Rileva serial duplicati nel file
      const seenSerials = new Set();
      const duplicateSerials = [];
      validRows.forEach(row => {
        const sn = String(row['SN']).trim();
        if (seenSerials.has(sn)) duplicateSerials.push(sn);
        else seenSerials.add(sn);
      });
      if (duplicateSerials.length > 0) {
        const proceed = window.confirm(
          `Attenzione: il file contiene ${duplicateSerials.length} matricole DUPLICATE:\n\n` +
          `${[...new Set(duplicateSerials)].slice(0, 10).join(', ')}${duplicateSerials.length > 10 ? '...' : ''}\n\n` +
          `Verranno caricate solo le matricole univoche (${seenSerials.size}). Procedere?`
        );
        if (!proceed) { setLoading(false); return; }
      }

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

      // Deduplica per serial (la chiave del controllo è il serial univoco)
      const dedup = new Map();
      validRows.forEach(row => {
        const sn = String(row['SN']).trim();
        if (!dedup.has(sn)) dedup.set(sn, {
          po_line_key: line.unique_key,
          serial: sn,
          model: String(row['Model'] || 'N/D').trim(),
          pn: String(row['PN'] || 'N/D').trim()
        });
      });
      const serialsToInsert = [...dedup.values()];

      // Inserimento a blocchi di 500 per evitare limiti di payload
      const CHUNK = 500;
      let insertErr = null;
      for (let i = 0; i < serialsToInsert.length; i += CHUNK) {
        const { error } = await supabase.from('expected_serials').insert(serialsToInsert.slice(i, i + CHUNK));
        if (error) { insertErr = error; break; }
      }

      if (insertErr) {
        alert("Errore nel caricamento delle matricole: " + insertErr.message);
      } else {
        // Conteggio reale dal DB
        const { count } = await supabase.from('expected_serials')
          .select('*', { count: 'exact', head: true }).eq('po_line_key', line.unique_key);
        await fetchPOLines();
        alert(`Matricole caricate: ${count} SN salvate (file: ${qtyProvided} righe, univoche: ${serialsToInsert.length}).`);
      }
    };
    reader.readAsArrayBuffer(file);
  }


  async function startScanningSession(line) {
    setLoading(true);
    setActiveLineKey(line.unique_key);
    setActiveLine(line);
    setCartonsScanned(line.cartons_scanned || 0);

    // Fetch paginato per superare il limite di 1000 righe di Supabase
    const fetchAllByLine = async (table, cols, orderCol) => {
      const pageSize = 1000;
      let all = [];
      let from = 0;
      while (true) {
        let qy = supabase.from(table).select(cols).eq('po_line_key', line.unique_key).range(from, from + pageSize - 1);
        if (orderCol) qy = qy.order(orderCol, { ascending: false });
        const { data, error } = await qy;
        if (error) { alert(`Errore caricamento ${table}: ${error.message}`); break; }
        all = [...all, ...(data || [])];
        if (!data || data.length < pageSize) break;
        from += pageSize;
      }
      return all;
    };

    const expectedData = await fetchAllByLine('expected_serials', 'serial, model, pn', null);
    const scannedData = await fetchAllByLine('scanned_serials', 'serial, model, pn, scanned_at', 'scanned_at');

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
                { id: 'stock', label: 'Inventario Spare Parts', icon: '🗄️' },
                { id: 'prelievi', label: 'Prelievi', icon: '📤' },
              ].map(mod => (
                <button
                  key={mod.id}
                  onClick={() => { setActiveModule(mod.id); setMenuOpen(false); setCurrentView('dashboard'); if (mod.id === 'prelievi') { setPrelievoView('list'); fetchPrelievi(); } }}
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
              {activeModule === 'stock' && 'Inventario Spare Parts'}
              {activeModule === 'prelievi' && 'Prelievi'}
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

        {/* ==================== MODULO DB SPARE PARTS ==================== */}
        {activeModule === 'spare-parts' && (() => {
          const uniqueTypes       = [...new Set(spareParts.map(p => p.type).filter(Boolean))].sort();
          const uniqueEOL         = [...new Set(spareParts.map(p => p.eol).filter(Boolean))].sort();
          const uniqueTerminalPNs = [...new Set(spareParts.map(p => p.terminal_pn).filter(Boolean))].sort();
          const uniqueRefs        = [...new Set(spareParts.map(p => p.ref).filter(Boolean))].sort();
          const uniqueRplus       = [...new Set(spareParts.map(p => p.rplus).filter(Boolean))].sort();
          const toggleSort = (col) => {
            if (spSortCol === col) setSpSortDir(d => d === 'asc' ? 'desc' : 'asc');
            else { setSpSortCol(col); setSpSortDir('asc'); }
            setSpPage(0);
          };
          const sortIcon = (col) => spSortCol === col ? (spSortDir === 'asc' ? ' ↑' : ' ↓') : '';

          const filtered = spareParts.filter(p => {
            const q = spSearch.toLowerCase();
            const q2 = spSearch2.toLowerCase();
            const matchField = (val) => (val || '').toLowerCase().includes;
            const matchAny = (q) => !q ||
              (p.pn || '').toLowerCase().includes(q) ||
              (p.terminal_pn || '').toLowerCase().includes(q) ||
              (p.english_name || '').toLowerCase().includes(q) ||
              (p.descrizione || '').toLowerCase().includes(q) ||
              (p.pnit || '').toLowerCase().includes(q) ||
              ((p.terminal_pn || '').split('-')[0]).toLowerCase().includes(q) ||
              (p.type || '').toLowerCase().includes(q) ||
              (p.ref || '').toLowerCase().includes(q) ||
              (p.rplus || '').toLowerCase().includes(q);
            void matchField;
            const matchSearch = matchAny(q);
            const matchSearch2 = matchAny(q2);
            const matchEOL        = !spFilterEOL        || p.eol === spFilterEOL;
            const matchType       = !spFilterType       || p.type === spFilterType;
            const matchToOrder    = !spFilterToOrder    || p.to_order === spFilterToOrder;
            const matchTerminalPN = !spFilterTerminalPN || p.terminal_pn === spFilterTerminalPN;
            const matchRef        = !spFilterRef        || p.ref === spFilterRef;
            const matchRplus      = !spFilterRplus      || p.rplus === spFilterRplus;
            return matchSearch && matchSearch2 && matchEOL && matchType && matchToOrder && matchTerminalPN && matchRef && matchRplus;
          }).sort((a, b) => {
            const getVal = (row) => spSortCol === 'modello' ? (row.terminal_pn || '').split('-')[0].toLowerCase() : String(row[spSortCol] ?? '').toLowerCase();
            const va = getVal(a);
            const vb = getVal(b);
            const num_a = parseFloat(a[spSortCol]);
            const num_b = parseFloat(b[spSortCol]);
            const isNum = !isNaN(num_a) && !isNaN(num_b);
            const cmp = isNum ? num_a - num_b : va.localeCompare(vb);
            return spSortDir === 'asc' ? cmp : -cmp;
          });

          return (
            <div className="space-y-5">

              {/* Upload + contatore */}
              <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs flex flex-col sm:flex-row sm:items-center gap-4">
                <div className="relative flex-shrink-0">
                  <button className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer shadow-xs transition">
                    📥 Importa DB Excel
                  </button>
                  <input type="file" accept=".xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handleSparePartsUpload} />
                </div>
                <p className="text-xs text-gray-400 flex-grow">Importa il file Excel del DB Spare Parts (Foglio 1). L&apos;import aggiorna i record esistenti per PN + Terminal PN.</p>
                {spareParts.length > 0 && (
                  <span className="text-sm bg-blue-50 text-blue-600 font-black px-3 py-1 rounded-full border border-blue-100 shrink-0">
                    {spareParts.length} ricambi
                  </span>
                )}
              </div>

              {/* Filtri */}
              {spareParts.length > 0 && (
                <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Filtri</span>
                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                    <input
                      type="text" value={spSearch} onChange={e => { setSpSearch(e.target.value); setSpPage(0); }}
                      placeholder="Cerca per PN, Terminal PN, nome, tipo, descrizione, REF, R+..."
                      className="sm:col-span-2 bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden"
                    />
                    <input
                      type="text" value={spSearch2} onChange={e => { setSpSearch2(e.target.value); setSpPage(0); }}
                      placeholder="Secondo filtro (AND)..."
                      className="sm:col-span-2 bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden"
                    />
                    <select value={spFilterTerminalPN} onChange={e => { setSpFilterTerminalPN(e.target.value); setSpPage(0); }} className="sm:col-span-2 bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">Tutti i Terminal PN ({uniqueTerminalPNs.length})</option>
                      {uniqueTerminalPNs.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                    <select value={spFilterType} onChange={e => { setSpFilterType(e.target.value); setSpPage(0); }} className="sm:col-span-2 bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">Tutti i TYPE</option>
                      {uniqueTypes.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                    <select value={spFilterEOL} onChange={e => { setSpFilterEOL(e.target.value); setSpPage(0); }} className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">Tutti gli stati</option>
                      {uniqueEOL.map(eol => <option key={eol} value={eol}>{eol}</option>)}
                    </select>
                    <select value={spFilterToOrder} onChange={e => { setSpFilterToOrder(e.target.value); setSpPage(0); }} className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">To Order: tutti</option>
                      <option value="SI">To Order: SI</option>
                      <option value="NO">To Order: NO</option>
                    </select>
                    <select value={spFilterRef} onChange={e => { setSpFilterRef(e.target.value); setSpPage(0); }} className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">REF: tutti ({uniqueRefs.length})</option>
                      {uniqueRefs.map(r => <option key={r} value={r}>{r}</option>)}
                    </select>
                    <select value={spFilterRplus} onChange={e => { setSpFilterRplus(e.target.value); setSpPage(0); }} className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">R+: tutti ({uniqueRplus.length})</option>
                      {uniqueRplus.map(r => <option key={r} value={r}>{r}</option>)}
                    </select>
                  </div>
                </div>
              )}

              {/* Barra strumenti — separata dai filtri */}
              {spareParts.length > 0 && (
                <div className="bg-white px-4 py-2.5 rounded-2xl border border-gray-200 shadow-xs flex items-center justify-between gap-3 flex-wrap">
                  <div className="flex items-center gap-3">
                    <p className="text-[10px] text-gray-400">
                      {filtered.length} risultati
                      {filtered.length > 100 && ` — pag. ${spPage + 1}/${Math.ceil(filtered.length / 100)}`}
                    </p>
                    {filtered.length > 100 && (
                      <div className="flex gap-1">
                        <button onClick={() => setSpPage(p => Math.max(0, p - 1))} disabled={spPage === 0}
                          className="text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer disabled:opacity-30 disabled:cursor-default bg-white hover:bg-gray-100 border-gray-200">← Prec.</button>
                        <button onClick={() => setSpPage(p => Math.min(Math.ceil(filtered.length / 100) - 1, p + 1))} disabled={(spPage + 1) * 100 >= filtered.length}
                          className="text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer disabled:opacity-30 disabled:cursor-default bg-white hover:bg-gray-100 border-gray-200">Succ. →</button>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    {filtered.length > 0 && (
                      <button onClick={() => {
                        const rows = filtered.map(p => ({
                          'PN': p.pn, 'Terminal PN': p.terminal_pn, 'Modello': (p.terminal_pn || '').split('-')[0], 'PNIT': p.pnit,
                          'English Name': p.english_name, 'Descrizione': p.descrizione,
                          'TYPE': p.type, 'Stato': p.eol, 'REF': p.ref, 'R+': p.rplus,
                          'To Order': p.to_order, '$': p.price, 'Locked': p.locked, 'CODICE': p.codice
                        }));
                        const ws = XLSX.utils.json_to_sheet(rows);
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, 'Spare Parts');
                        XLSX.writeFile(wb, `spare_parts_${filtered.length}.xlsx`);
                      }} className="text-[10px] font-bold text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 border border-blue-200 px-2.5 py-1 rounded-lg transition cursor-pointer">
                        📥 Esporta XLS ({filtered.length})
                      </button>
                    )}
                    <button onClick={() => { setSpEditMode(v => !v); if (Object.keys(spPendingChanges).length === 0) setSpPendingChanges({}); }}
                      className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer ${spEditMode ? 'bg-amber-500 text-white border-amber-600' : 'bg-gray-50 text-gray-500 hover:bg-gray-100 border-gray-200'}`}>
                      {spEditMode ? '✏️ Modifica ON' : '✏️ Modifica'}
                    </button>
                    <button onClick={() => { setSpSearch(''); setSpSearch2(''); setSpFilterTerminalPN(''); setSpFilterType(''); setSpFilterEOL(''); setSpFilterToOrder(''); setSpFilterRef(''); setSpFilterRplus(''); setSpPage(0); }}
                      className="text-[10px] font-bold text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-200 px-2.5 py-1 rounded-lg transition cursor-pointer">
                      ✕ Reset filtri
                    </button>
                    <button onClick={spResetWidths} className="text-[10px] font-bold text-gray-400 hover:text-gray-600 bg-gray-50 hover:bg-gray-100 border border-gray-200 px-2.5 py-1 rounded-lg transition cursor-pointer">
                      ↺ Reset colonne
                    </button>
                  </div>
                </div>
              )}

              {/* Tabella */}
              {spLoading && <div className="text-center py-4 text-xs font-bold text-amber-600 animate-pulse">Caricamento...</div>}
              {!spLoading && spareParts.length > 0 && (
                <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
                  <table className="text-left border-collapse text-[11px]" style={{ tableLayout: 'fixed', width: Object.values(spWidths).reduce((a,b) => a+b, 0) + 'px' }}>
                    <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                      <tr>
                        {[
                          ['pn','PN'], ['terminal_pn','Terminal PN'], ['modello','Modello'], ['pnit','PNIT'],
                          ['english_name','English Name', 'hidden md:table-cell'], ['descrizione','Descrizione', 'hidden md:table-cell'],
                          ['type','TYPE'], ['eol','ST.'], ['ref','REF'], ['rplus','R+'],
                          ['to_order','ORD'], ['price','$']
                        ].map(([col, label, extra = '']) => (
                          <th key={col}
                            style={{ width: spWidths[col] }}
                            onClick={() => toggleSort(col)}
                            className={`relative px-2 py-3 truncate cursor-pointer select-none hover:bg-gray-100 transition ${col === 'price' ? 'text-right' : ''} ${spSortCol === col ? 'text-blue-600' : ''} ${extra}`}>
                            {label}{sortIcon(col)}
                            <div onMouseDown={e => spStartResize(col, e)}
                              className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/40 z-10" />
                          </th>
                        ))}
                        <th style={{ width: spWidths.locked }} className="relative px-2 py-3 text-center truncate">
                          🔒
                          <div onMouseDown={e => spStartResize('locked', e)} className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/40 z-10" />
                        </th>
                        <th style={{ width: spWidths.edit }} className="px-2 py-3"></th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {filtered.slice(spPage * 100, spPage * 100 + 100).map((p, rowIndex) => {
                        const rowKey = `${p.pn}_${p.terminal_pn}`;
                        const pending = spPendingChanges[rowKey];
                        const d = pending ? { ...p, ...pending } : p;
                        const iCls = `w-full rounded px-1 py-0.5 text-xs focus:outline-none border ${pending ? 'bg-amber-50 border-amber-300' : 'bg-transparent border-transparent hover:border-gray-300 focus:border-blue-400 focus:bg-white'}`;
                        const nav = (field) => ({ 'data-sp-row': rowIndex, 'data-sp-field': field, onKeyDown: (e) => handleSpKeyNav(e, rowIndex, field) });
                        const computedEol = (d.eol || '').trim();
                        const computedType = (d.type || '').trim();
                        const computedToOrder = (d.to_order || '').trim();
                        const computedLocked = (computedEol === 'EOL' || computedEol === 'CLI' || computedToOrder === 'NO' || computedType === 'NO') ? 'Y' : '';
                        return (
                          <tr key={rowKey} className={`transition ${pending ? 'bg-amber-50/40' : 'hover:bg-gray-50/80'}`}>
                            <td className="px-2 py-2 font-mono font-bold text-blue-700 truncate" title={p.pn}>{p.pn}</td>
                            <td className="px-2 py-2 font-mono text-gray-600 text-[10px] truncate" title={p.terminal_pn}>{p.terminal_pn}</td>
                            <td className="px-2 py-2 font-mono font-bold text-gray-700 text-[10px] truncate">{(p.terminal_pn || '').split('-')[0]}</td>
                            <td className="px-2 py-1 font-mono text-gray-500">
                              {spEditMode ? <input className={iCls + ' font-mono'} value={d.pnit || ''} onChange={e => setSpFieldChange(rowKey, 'pnit', e.target.value)} {...nav('pnit')} /> : p.pnit}
                            </td>
                            <td className="px-2 py-1 text-[10px] hidden md:table-cell">
                              {spEditMode ? <input className={iCls} value={d.english_name || ''} onChange={e => setSpFieldChange(rowKey, 'english_name', e.target.value)} {...nav('english_name')} /> : p.english_name}
                            </td>
                            <td className="px-2 py-1 text-[10px] hidden md:table-cell">
                              {spEditMode ? <input className={iCls} value={d.descrizione || ''} onChange={e => setSpFieldChange(rowKey, 'descrizione', e.target.value)} {...nav('descrizione')} /> : p.descrizione}
                            </td>
                            <td className="px-2 py-1 text-[10px]">
                              {spEditMode ? <input className={iCls} value={d.type || ''} onChange={e => setSpFieldChange(rowKey, 'type', e.target.value)} {...nav('type')} /> : p.type}
                            </td>
                            <td className="px-1 py-1 text-center">
                              {spEditMode
                                ? <select className={iCls} value={d.eol || ''} onChange={e => setSpFieldChange(rowKey, 'eol', e.target.value)} {...nav('eol')}>
                                    <option value="">—</option><option value="EOL">EOL</option><option value="ALT">ALT</option><option value="CLI">CLI</option>
                                  </select>
                                : p.eol ? <span className={`px-1 py-px rounded font-black text-[9px] border ${p.eol === 'EOL' ? 'bg-red-50 text-red-600 border-red-100' : p.eol === 'ALT' ? 'bg-amber-50 text-amber-700 border-amber-100' : 'bg-blue-50 text-blue-700 border-blue-100'}`}>{p.eol}</span> : ''}
                            </td>
                            <td className="px-1 py-1 text-center">
                              {spEditMode
                                ? <select className={iCls} value={d.ref || ''} onChange={e => setSpFieldChange(rowKey, 'ref', e.target.value)} {...nav('ref')}>
                                    <option value="">—</option><option value="X">X</option>
                                  </select>
                                : p.ref === 'X' ? <span className="font-black text-gray-600 text-[11px]">✓</span> : ''}
                            </td>
                            <td className="px-1 py-1 text-center">
                              {spEditMode
                                ? <select className={iCls} value={d.rplus || ''} onChange={e => setSpFieldChange(rowKey, 'rplus', e.target.value)} {...nav('rplus')}>
                                    <option value="">—</option><option value="X">X</option>
                                  </select>
                                : p.rplus === 'X' ? <span className="font-black text-gray-600 text-[11px]">✓</span> : ''}
                            </td>
                            <td className="px-1 py-1 text-center">
                              {spEditMode
                                ? <select className={iCls} value={d.to_order || ''} onChange={e => setSpFieldChange(rowKey, 'to_order', e.target.value)} {...nav('to_order')}>
                                    <option value="SI">SI</option><option value="NO">NO</option>
                                  </select>
                                : p.to_order === 'SI'
                                  ? <span className="bg-green-50 text-green-700 border border-green-100 px-1 py-px rounded font-black text-[9px]">SI</span>
                                  : <span className="text-gray-300 text-[9px]">NO</span>}
                            </td>
                            <td className="px-2 py-1 text-right font-mono text-gray-700">
                              {spEditMode ? <input className={iCls + ' text-right'} value={d.price ?? ''} onChange={e => setSpFieldChange(rowKey, 'price', e.target.value)} {...nav('price')} /> : (p.price > 0 ? `$${p.price.toFixed(2)}` : '—')}
                            </td>
                            <td className="px-2 py-2 text-center text-[11px]">{computedLocked === 'Y' ? '🔒' : '—'}</td>
                            <td className="px-2 py-2 text-center">
                              {pending && (
                                <button onClick={() => setSpPendingChanges(prev => { const n = {...prev}; delete n[rowKey]; return n; })}
                                  className="text-[10px] text-gray-400 hover:text-red-500 cursor-pointer" title="Annulla modifiche riga">↺</button>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}

              {!spLoading && spareParts.length === 0 && (
                <div className="text-center py-20 text-gray-400 text-sm">
                  Nessun ricambio caricato. Importa il file Excel per iniziare.
                </div>
              )}
            </div>
          );
        })()}
        {activeModule === 'stock' && (() => {
          // Join client-side: codice -> spare_parts.pn
          const spMap = {};
          spareParts.forEach(p => {
            const mod = (p.terminal_pn || '').split('-')[0];
            if (!spMap[p.pn]) spMap[p.pn] = { english_name: p.english_name, descrizione: p.descrizione, eol: p.eol, modelli: new Set() };
            if (mod) spMap[p.pn].modelli.add(mod);
          });
          Object.values(spMap).forEach(v => { v.modello = [...v.modelli].join(', '); delete v.modelli; });

          const noMatchCount = stockItems.filter(s => !spMap[s.codice]).length;
          const uniqueMagazzini  = [...new Set(stockItems.map(s => s.magazzino).filter(Boolean))].sort();
          const uniqueLocazioni  = [...new Set(stockItems.map(s => s.locazione).filter(Boolean))].sort();

          const toggleStockSort = (col) => {
            if (stockSortCol === col) setStockSortDir(d => d === 'asc' ? 'desc' : 'asc');
            else { setStockSortCol(col); setStockSortDir('asc'); }
            setStockPage(0);
          };
          const stockSortIcon = (col) => stockSortCol === col ? (stockSortDir === 'asc' ? ' ↑' : ' ↓') : '';

          const filtered = stockItems
            .map(s => ({ ...s, ...(spMap[s.codice] || { english_name: '', descrizione: '' }) }))
            .filter(s => {
              const matchAny = (q) => !q ||
                (s.codice || '').toLowerCase().includes(q) ||
                (s.english_name || '').toLowerCase().includes(q) ||
                (s.descrizione || '').toLowerCase().includes(q) ||
                (s.numero_bancale || '').toLowerCase().includes(q) ||
                (s.modello || '').toLowerCase().includes(q) ||
                (s.magazzino || '').toLowerCase().includes(q) ||
                (s.locazione || '').toLowerCase().includes(q);
              const matchSearch = matchAny(stockSearch.toLowerCase());
              const matchSearch2 = matchAny(stockSearch2.toLowerCase());
              const matchMag = !stockFilterMagazzino || s.magazzino === stockFilterMagazzino;
              const matchLoc = !stockFilterLocazione || s.locazione === stockFilterLocazione;
              const matchBancale = !stockFilterBancale || s.numero_bancale === stockFilterBancale;
              const matchNoMatch = !stockFilterNoMatch || !spMap[s.codice];
              return matchSearch && matchSearch2 && matchMag && matchLoc && matchBancale && matchNoMatch && s.stock && s.stock !== 0;
            })
            .sort((a, b) => {
              const va = String(a[stockSortCol] ?? '').toLowerCase();
              const vb = String(b[stockSortCol] ?? '').toLowerCase();
              const na = parseFloat(a[stockSortCol]), nb = parseFloat(b[stockSortCol]);
              const isNum = !isNaN(na) && !isNaN(nb);
              const cmp = isNum ? na - nb : va.localeCompare(vb);
              return stockSortDir === 'asc' ? cmp : -cmp;
            });

          const cols = [
            { key: 'locazione',      label: 'Locazione' },
            { key: 'numero_bancale', label: 'Bancale' },
            { key: 'magazzino',      label: 'Magazzino' },
            { key: 'codice',         label: 'Codice' },
            { key: 'modello',        label: 'Modello' },
            { key: 'eol',            label: 'ST.' },
            { key: 'english_name',   label: 'English Name' },
            { key: 'descrizione',    label: 'Descrizione' },
            { key: 'stock',          label: 'Stock', right: true },
            { key: 'edit',           label: '', noSort: true },
          ];

          return (
            <div className="space-y-5">

              {/* Upload + contatore */}
              <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs flex flex-col sm:flex-row sm:items-center gap-4">
                <div className="relative flex-shrink-0">
                  <button className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer shadow-xs transition">
                    📥 Importa Stock Excel
                  </button>
                  <input type="file" accept=".xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handleStockUpload} />
                </div>
                <p className="text-xs text-gray-400 flex-grow">Importa il file Excel dello Stock (Foglio 2). Chiave: Codice + Magazzino + Bancale.</p>
                {stockItems.length > 0 && (
                  <span className="text-sm bg-blue-50 text-blue-600 font-black px-3 py-1 rounded-full border border-blue-100 shrink-0">
                    {filtered.length} record
                  </span>
                )}
              </div>

              {/* Sposta bancale */}
              {stockItems.length > 0 && (
                <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs">
                  {!moveBancaleOpen ? (
                    <button onClick={() => setMoveBancaleOpen(true)}
                      className="text-sm font-bold text-blue-600 hover:text-blue-800 cursor-pointer flex items-center gap-2">
                      🏭 Sposta bancale tra magazzini
                    </button>
                  ) : (
                    <div className="flex flex-wrap items-end gap-3">
                      <div className="space-y-1">
                        <label className="block text-[10px] font-bold text-gray-500 uppercase">Bancale</label>
                        <select value={moveBancaleSrc} onChange={e => setMoveBancaleSrc(e.target.value)}
                          className="bg-gray-50 border border-gray-300 rounded-xl p-2 text-xs focus:outline-hidden">
                          <option value="">Seleziona bancale...</option>
                          {[...new Set(stockItems.map(s => s.numero_bancale).filter(Boolean))].sort().map(b => (
                            <option key={b} value={b}>{b} ({stockItems.filter(s => s.numero_bancale === b)[0]?.magazzino})</option>
                          ))}
                        </select>
                      </div>
                      <div className="space-y-1">
                        <label className="block text-[10px] font-bold text-gray-500 uppercase">Magazzino dest.</label>
                        <select value={moveBancaleDest} onChange={e => setMoveBancaleDest(e.target.value)}
                          className="bg-gray-50 border border-gray-300 rounded-xl p-2 text-xs focus:outline-hidden">
                          <option value="GESSATE">GESSATE</option>
                          <option value="ESPRINET">ESPRINET</option>
                        </select>
                      </div>
                      <div className="space-y-1">
                        <label className="block text-[10px] font-bold text-gray-500 uppercase">Locazione dest. <span className="text-gray-400 normal-case font-normal">(opz.)</span></label>
                        <input value={moveBancaleLocazione} onChange={e => setMoveBancaleLocazione(e.target.value)}
                          placeholder="Es. SCAFFALE-A3"
                          className="bg-gray-50 border border-gray-300 rounded-xl p-2 text-xs focus:outline-hidden w-full" />
                      </div>
                      <button onClick={moveBancale}
                        className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold px-4 py-2 rounded-xl cursor-pointer transition shadow-xs">
                        ↗ Sposta
                      </button>
                      <button onClick={() => { setMoveBancaleOpen(false); setMoveBancaleSrc(''); }}
                        className="text-xs text-gray-400 hover:text-gray-600 cursor-pointer px-2 py-2">
                        Annulla
                      </button>
                    </div>
                  )}
                </div>
              )}

              {/* Filtri */}
              {stockItems.length > 0 && (
                <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Filtri</span>
                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                    <input type="text" value={stockSearch}
                      onChange={e => { setStockSearch(e.target.value); setStockPage(0); }}
                      placeholder="Cerca per codice, modello, nome, descrizione..."
                      className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
                    <input type="text" value={stockSearch2}
                      onChange={e => { setStockSearch2(e.target.value); setStockPage(0); }}
                      placeholder="Secondo filtro (AND)..."
                      className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
                    <select value={stockFilterMagazzino}
                      onChange={e => { setStockFilterMagazzino(e.target.value); setStockPage(0); }}
                      className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">Tutti i magazzini ({uniqueMagazzini.length})</option>
                      {uniqueMagazzini.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                    <select value={stockFilterLocazione}
                      onChange={e => { setStockFilterLocazione(e.target.value); setStockPage(0); }}
                      className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">Tutte le locazioni ({uniqueLocazioni.length})</option>
                      {uniqueLocazioni.map(l => <option key={l} value={l}>{l}</option>)}
                    </select>
                    <select value={stockFilterBancale}
                      onChange={e => { setStockFilterBancale(e.target.value); setStockPage(0); }}
                      className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                      <option value="">Tutti i bancali ({[...new Set(stockItems.map(s => s.numero_bancale).filter(Boolean))].length})</option>
                      {[...new Set(stockItems.map(s => s.numero_bancale).filter(Boolean))].sort().map(b => (
                        <option key={b} value={b}>{b}</option>
                      ))}
                    </select>
                  </div>
                  {noMatchCount > 0 && (
                    <button
                      onClick={() => { setStockFilterNoMatch(v => !v); setStockPage(0); }}
                      className={`w-full text-left text-xs font-bold px-3 py-2 rounded-xl border transition cursor-pointer ${stockFilterNoMatch ? 'bg-red-600 text-white border-red-700' : 'bg-red-50 text-red-600 border-red-200 hover:bg-red-100'}`}>
                      ⚠ {noMatchCount} codici senza riscontro nel DB Spare Parts
                      {stockFilterNoMatch ? ' — clicca per mostrare tutti' : ' — clicca per filtrare'}
                    </button>
                  )}
                </div>
              )}

              {/* Barra risultati + azioni — separata dai filtri */}
              {stockItems.length > 0 && (
                <div className="bg-white px-4 py-2.5 rounded-2xl border border-gray-200 shadow-xs flex items-center justify-between gap-3 flex-wrap">
                  <div className="flex items-center gap-3">
                    <p className="text-[10px] text-gray-400">
                      {filtered.length} risultati
                      {filtered.length > 100 && ` — pag. ${stockPage + 1}/${Math.ceil(filtered.length / 100)}`}
                    </p>
                    {filtered.length > 100 && (
                      <div className="flex gap-1">
                        <button onClick={() => setStockPage(p => Math.max(0, p - 1))} disabled={stockPage === 0}
                          className="text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer disabled:opacity-30 bg-white hover:bg-gray-100 border-gray-200">← Prec.</button>
                        <button onClick={() => setStockPage(p => Math.min(Math.ceil(filtered.length / 100) - 1, p + 1))} disabled={(stockPage + 1) * 100 >= filtered.length}
                          className="text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer disabled:opacity-30 bg-white hover:bg-gray-100 border-gray-200">Succ. →</button>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    {filtered.length > 0 && (
                      <button onClick={() => {
                        const rows = filtered.map(s => ({
                          'Locazione': s.locazione, 'Bancale': s.numero_bancale,
                          'Magazzino': s.magazzino, 'Codice': s.codice,
                          'English Name': s.english_name, 'Descrizione': s.descrizione,
                          'Stock': s.stock
                        }));
                        const ws = XLSX.utils.json_to_sheet(rows);
                        const wb2 = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb2, ws, 'Stock');
                        XLSX.writeFile(wb2, `stock_${filtered.length}.xlsx`);
                      }} className="text-[10px] font-bold text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 border border-blue-200 px-2.5 py-1 rounded-lg transition cursor-pointer">
                        📥 Esporta XLS ({filtered.length})
                      </button>
                    )}
                    <button onClick={() => setStockEditMode(v => !v)}
                      className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer ${stockEditMode ? 'bg-amber-500 text-white border-amber-600' : 'bg-gray-50 text-gray-500 hover:bg-gray-100 border-gray-200'}`}>
                      {stockEditMode ? '✏️ Modifica ON' : '✏️ Modifica'}
                    </button>
                    <button onClick={() => { setStockSearch(''); setStockSearch2(''); setStockFilterMagazzino(''); setStockFilterLocazione(''); setStockFilterBancale(''); setStockFilterNoMatch(false); setStockPage(0); }}
                      className="text-[10px] font-bold text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-200 px-2.5 py-1 rounded-lg transition cursor-pointer">
                      ✕ Reset filtri
                    </button>
                    <button onClick={stResetWidths} className="text-[10px] font-bold text-gray-400 hover:text-gray-600 bg-gray-50 hover:bg-gray-100 border border-gray-200 px-2.5 py-1 rounded-lg transition cursor-pointer">
                      ↺ Reset colonne
                    </button>
                  </div>
                </div>
              )}

              {/* Tabella */}
              {stockLoading && <div className="text-center py-4 text-xs font-bold text-amber-600 animate-pulse">Caricamento...</div>}
              {!stockLoading && stockItems.length > 0 && (
                <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
                  <table className="text-left border-collapse text-[11px]" style={{ tableLayout: 'fixed', width: Object.values(stWidths).reduce((a,b) => a+b, 0) + 'px' }}>
                    <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                      <tr>
                        {cols.map(({ key, label, extra = '', right = false, noSort = false }) => (
                          <th key={key}
                            style={{ width: stWidths[key] }}
                            onClick={() => !noSort && toggleStockSort(key)}
                            className={`relative px-2 py-3 truncate bg-gray-50 ${noSort ? '' : 'cursor-pointer select-none hover:bg-gray-100'} transition ${right ? 'text-right' : ''} ${!noSort && stockSortCol === key ? 'text-blue-600' : ''} ${extra}`}>
                            {label}{!noSort && stockSortIcon(key)}
                            {!noSort && <div onMouseDown={e => stStartResize(key, e)} className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/40 z-10" />}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {filtered.slice(stockPage * 100, stockPage * 100 + 100).map((s, rowIndex) => {
                          const pending = stockPendingChanges[s.id];
                          const d = pending ? { ...s, ...pending } : s;
                          const iCls = `w-full rounded px-1 py-0.5 text-xs focus:outline-none border ${pending ? 'bg-amber-50 border-amber-300' : 'bg-transparent border-transparent hover:border-gray-300 focus:border-blue-400 focus:bg-white'}`;
                          const nav = (field) => ({ 'data-rowindex': rowIndex, 'data-field': field, onKeyDown: (e) => handleStockKeyNav(e, rowIndex, field) });
                          const eolBadge = (eol) => eol ? <span className={`px-1 py-px rounded font-black text-[9px] border ${eol === 'EOL' ? 'bg-red-50 text-red-600 border-red-100' : eol === 'ALT' ? 'bg-amber-50 text-amber-700 border-amber-100' : 'bg-blue-50 text-blue-700 border-blue-100'}`}>{eol}</span> : '';
                          return (
                            <tr key={s.id} className={`transition ${pending ? 'bg-amber-50/40' : 'hover:bg-gray-50/80'}`}>
                              <td className="px-2 py-2 truncate">
                                {stockEditMode ? <input className={iCls} value={d.locazione || ''} onChange={e => setStockFieldChange(s.id, 'locazione', e.target.value)} {...nav('locazione')} /> : s.locazione || '—'}
                              </td>
                              <td className="px-2 py-2 truncate">
                                {stockEditMode ? <input className={iCls} value={d.numero_bancale || ''} onChange={e => setStockFieldChange(s.id, 'numero_bancale', e.target.value)} {...nav('numero_bancale')} /> : s.numero_bancale || '—'}
                              </td>
                              <td className="px-2 py-2 truncate">
                                {stockEditMode ? (
                                  <select className={iCls} value={d.magazzino || 'GESSATE'} onChange={e => setStockFieldChange(s.id, 'magazzino', e.target.value)} {...nav('magazzino')}>
                                    <option value="GESSATE">GESSATE</option>
                                    <option value="ESPRINET">ESPRINET</option>
                                  </select>
                                ) : s.magazzino || '—'}
                              </td>
                              <td className="px-2 py-2 font-mono font-bold text-blue-700 truncate">
                                {stockEditMode ? <input className={iCls + ' font-mono font-bold text-blue-700'} value={d.codice || ''} onChange={e => setStockFieldChange(s.id, 'codice', e.target.value)} {...nav('codice')} /> : s.codice}
                              </td>
                              <td className="px-2 py-2 font-mono font-bold text-gray-700 text-[10px] truncate">{s.modello || '—'}</td>
                              <td className="px-1 py-2 text-center">{eolBadge(s.eol)}</td>
                              <td className="px-2 py-2 text-[10px] leading-snug">{s.english_name || '—'}</td>
                              <td className="px-2 py-2 text-[10px] leading-snug">{s.descrizione || '—'}</td>
                              <td className="px-2 py-2 text-right font-mono font-black">
                                {stockEditMode ? <input className={iCls + ' text-right font-mono font-black'} value={d.stock ?? ''} onChange={e => setStockFieldChange(s.id, 'stock', e.target.value)} {...nav('stock')} /> : (s.stock ?? '—')}
                              </td>
                              <td className="px-2 py-2 text-center">
                                {pending && (
                                  <button onClick={() => setStockPendingChanges(prev => { const n = {...prev}; delete n[s.id]; return n; })}
                                    className="text-[10px] text-gray-400 hover:text-red-500 cursor-pointer" title="Annulla modifiche riga">↺</button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              )}

              {!stockLoading && stockItems.length === 0 && (
                <div className="text-center py-20 text-gray-400 text-sm">
                  Nessun record stock. Importa il file Excel per iniziare.
                </div>
              )}
            </div>
          );
        })()}

        {/* ==================== MODULO PRELIEVI — LISTA ==================== */}
        {activeModule === 'prelievi' && prelievoView === 'list' && (
          <div className="space-y-5">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-lg font-black text-gray-800">📤 Prelievi effettuati</h2>
                <p className="text-xs text-gray-500">{prelieviList.length} prelievi registrati</p>
              </div>
              <div className="flex gap-2">
                <button onClick={fetchPrelievi}
                  className="bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold px-3 py-2.5 rounded-xl cursor-pointer transition" title="Aggiorna">
                  ↻
                </button>
                <button onClick={() => { setPrelievoView('new'); setPrelievoRighe([]); setPrelievoFeedback({ text: '', type: '' }); setTimeout(() => prelievoScannerRef.current?.focus(), 100); }}
                  className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer transition shadow-xs">
                  + Nuovo prelievo
                </button>
              </div>
            </div>

            {prelieviLoading && <div className="text-center py-4 text-xs font-bold text-amber-600 animate-pulse">Caricamento...</div>}

            {!prelieviLoading && prelieviList.length > 0 && (
              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                    <tr>
                      <th className="px-3 py-3">ID Prelievo</th>
                      <th className="px-3 py-3">Data</th>
                      <th className="px-3 py-3">Operatore</th>
                      <th className="px-3 py-3">Destinazione</th>
                      <th className="px-3 py-3 text-right">Righe</th>
                      <th className="px-3 py-3 text-right">Pezzi</th>
                      <th className="px-3 py-3"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {prelieviList.map(p => (
                      <tr key={p.id} className="hover:bg-blue-50/50 transition cursor-pointer" onClick={() => openPrelievoDetail(p)}>
                        <td className="px-3 py-2.5 font-mono font-bold text-blue-700 underline">{p.id_prelievo}</td>
                        <td className="px-3 py-2.5 text-gray-600">{p.data_prelievo ? new Date(p.data_prelievo).toLocaleString('it-IT') : '—'}</td>
                        <td className="px-3 py-2.5 text-gray-700">{p.utente || '—'}</td>
                        <td className="px-3 py-2.5 text-gray-600">{p.destinazione || '—'}</td>
                        <td className="px-3 py-2.5 text-right font-mono">{p.n_righe}</td>
                        <td className="px-3 py-2.5 text-right font-mono font-black">{p.n_pezzi}</td>
                        <td className="px-3 py-2.5 text-center">
                          <button onClick={(e) => { e.stopPropagation(); deletePrelievo(p); }}
                            className="text-[10px] text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-300 px-2 py-1 rounded-lg cursor-pointer transition" title="Elimina e ripristina inventario">🗑</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {!prelieviLoading && prelieviList.length === 0 && (
              <div className="text-center py-16 text-gray-400 text-sm">
                Nessun prelievo registrato. Clicca &quot;Nuovo prelievo&quot; per iniziare.
              </div>
            )}
          </div>
        )}

        {/* ==================== MODULO PRELIEVI — DETTAGLIO ==================== */}
        {activeModule === 'prelievi' && prelievoView === 'detail' && prelievoDetail && (
          <div className="space-y-5 max-w-4xl mx-auto">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-lg font-black text-gray-800">📤 Prelievo {prelievoDetail.testata.id_prelievo}</h2>
                <p className="text-xs text-gray-500">
                  {prelievoDetail.testata.data_prelievo ? new Date(prelievoDetail.testata.data_prelievo).toLocaleString('it-IT') : '—'}
                  {' · '}Operatore: <strong>{prelievoDetail.testata.utente || '—'}</strong>
                  {prelievoDetail.testata.destinazione ? <> · Dest: <strong>{prelievoDetail.testata.destinazione}</strong></> : null}
                </p>
              </div>
              <button onClick={() => { setPrelievoView('list'); setPrelievoDetail(null); }}
                className="bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition">
                ← Torna alla lista
              </button>
            </div>

            <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
              <div className="flex items-center justify-between px-4 py-3 border-b border-gray-100">
                <span className="text-xs font-black text-gray-500 uppercase tracking-wider">{prelievoDetail.righe.length} righe</span>
                <span className="text-sm font-black text-blue-600">Tot. pezzi: {prelievoDetail.righe.reduce((s, r) => s + (r.quantita || 0), 0)}</span>
              </div>
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                  <tr>
                    <th className="px-3 py-3">Codice</th>
                    <th className="px-3 py-3">ID Cartone</th>
                    <th className="px-3 py-3">Magazzino</th>
                    <th className="px-3 py-3">Bancale</th>
                    <th className="px-3 py-3 text-right">Quantità</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {prelievoDetail.righe.map(r => (
                    <tr key={r.id} className="hover:bg-gray-50/80">
                      <td className="px-3 py-2.5 font-mono font-bold text-blue-700">{r.codice}</td>
                      <td className="px-3 py-2.5 font-mono text-[10px] text-gray-500">{r.id_cartone || '—'}</td>
                      <td className="px-3 py-2.5 text-gray-600">{r.magazzino}</td>
                      <td className="px-3 py-2.5 text-gray-600">{r.numero_bancale || '—'}</td>
                      <td className="px-3 py-2.5 text-right font-mono font-black">{r.quantita}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ==================== MODULO PRELIEVI — NUOVO ==================== */}
        {activeModule === 'prelievi' && prelievoView === 'new' && (() => {
          // Mappa pn -> descrizione (precalcolata una volta)
          const spDescMap = {};
          for (const p of spareParts) { if (!spDescMap[p.pn]) spDescMap[p.pn] = p.descrizione || p.english_name || ''; }
          // Codici in stock, filtrati sul testo digitato (max 50 suggerimenti)
          const allStockCodici = [...new Set(stockItems.filter(s => s.stock > 0).map(s => s.codice).filter(Boolean))].sort();
          const q = prelievoManuale.codice.trim().toLowerCase();
          const stockCodici = q.length >= 2
            ? allStockCodici.filter(c => c.toLowerCase().includes(q) || (spDescMap[c] || '').toLowerCase().includes(q)).slice(0, 50)
            : [];
          // Righe stock per il codice selezionato
          const righeCodice = prelievoManuale.codice
            ? stockItems.filter(s => s.codice === prelievoManuale.codice.trim() && s.stock > 0
                && (prelievoShowEsprinet || s.magazzino === 'GESSATE'))
            : [];
          const totalePezzi = prelievoRighe.reduce((sum, r) => sum + (parseFloat(r.quantita) || 0), 0);

          return (
            <div className="space-y-5 max-w-4xl mx-auto">

              <div className="flex items-center justify-between">
                <h2 className="text-lg font-black text-gray-800">📤 Nuovo prelievo</h2>
                <button onClick={() => { if (prelievoRighe.length === 0 || window.confirm('Uscire? Le righe non registrate andranno perse.')) { setPrelievoView('list'); setPrelievoRighe([]); } }}
                  className="bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition">
                  ← Torna alla lista
                </button>
              </div>


              {/* Testata prelievo */}
              <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Dati prelievo</span>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Operatore</label>
                    <input value={prelievoUtente} onChange={e => setPrelievoUtente(e.target.value)}
                      placeholder={currentUser || 'Nome operatore'}
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-sm focus:outline-hidden" />
                  </div>
                  <div className="space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Destinazione</label>
                    <select value={prelievoDest} onChange={e => setPrelievoDest(e.target.value)}
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-sm focus:outline-hidden">
                      <option value="">Seleziona destinazione...</option>
                      <option value="Secure Room">Secure Room</option>
                      <option value="Repair">Repair</option>
                    </select>
                  </div>
                </div>
                <label className="flex items-center gap-2 cursor-pointer pt-1">
                  <input type="checkbox" checked={prelievoShowEsprinet} onChange={e => { setPrelievoShowEsprinet(e.target.checked); setPrelievoManuale(v => ({ ...v, stockId: '' })); }} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                  <span className="text-xs font-semibold text-gray-700">Mostra anche ubicazioni ESPRINET (default: solo GESSATE)</span>
                </label>
              </div>

              {/* Scanner ID cartone */}
              <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Scansione ID Cartone / QR</span>
                <form onSubmit={handlePrelievoScan}>
                  <input type="text" ref={prelievoScannerRef} value={prelievoScanner}
                    onChange={e => setPrelievoScanner(e.target.value)}
                    autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
                    placeholder="Spara l'etichetta o il QR del cartone..."
                    className="w-full bg-white border-2 border-blue-500 text-gray-800 font-mono text-base p-4 rounded-xl shadow-inner focus:outline-hidden text-center" />
                </form>
                {prelievoFeedback.text && (
                  <div className={`p-3 rounded-xl text-center text-sm font-bold border ${prelievoFeedback.type === 'success' ? 'bg-green-50 text-green-800 border-green-200' : 'bg-red-50 text-red-800 border-red-200'}`}>
                    {prelievoFeedback.text}
                  </div>
                )}
              </div>

              {/* Inserimento manuale */}
              <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Prelievo manuale (codice + ubicazione)</span>
                <div className="grid grid-cols-1 sm:grid-cols-12 gap-3 items-end">
                  <div className="sm:col-span-4 space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Codice articolo</label>
                    <input list="prelievo-codici" value={prelievoManuale.codice}
                      onChange={e => setPrelievoManuale(v => ({ ...v, codice: e.target.value, stockId: '' }))}
                      placeholder="Cerca codice..."
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs font-mono focus:outline-hidden" />
                    <datalist id="prelievo-codici">
                      {stockCodici.map(c => <option key={c} value={c}>{spDescMap[c] || ''}</option>)}
                    </datalist>
                  </div>
                  <div className="sm:col-span-5 space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Ubicazione (bancale / riga)</label>
                    <select value={prelievoManuale.stockId}
                      onChange={e => { const sid = e.target.value; const row = stockItems.find(s => String(s.id) === sid); setPrelievoManuale(v => ({ ...v, stockId: sid, quantita: row ? row.stock : '' })); }}
                      disabled={!prelievoManuale.codice}
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden disabled:opacity-50">
                      <option value="">{prelievoManuale.codice ? `Seleziona (${righeCodice.length})` : 'Prima scegli un codice'}</option>
                      {righeCodice.map(s => (
                        <option key={s.id} value={s.id}>
                          {s.magazzino} / {s.numero_bancale || '—'}{s.locazione ? ` / ${s.locazione}` : ''} — disp. {s.stock}{s.carton_ids?.[0] ? ` — ${s.carton_ids[0]}` : ''}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="sm:col-span-2 space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Qtà</label>
                    <input type="number" value={prelievoManuale.quantita}
                      onChange={e => setPrelievoManuale(v => ({ ...v, quantita: e.target.value }))}
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
                  </div>
                  <div className="sm:col-span-1">
                    <button onClick={addPrelievoManuale}
                      className="w-full bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-3 py-2.5 rounded-xl cursor-pointer transition shadow-xs">+</button>
                  </div>
                </div>
              </div>

              {/* Righe prelievo */}
              {prelievoRighe.length > 0 && (
                <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
                  <div className="flex items-center justify-between px-4 py-3 border-b border-gray-100">
                    <span className="text-xs font-black text-gray-500 uppercase tracking-wider">{prelievoRighe.length} righe da prelevare</span>
                    <span className="text-sm font-black text-blue-600">Tot. pezzi: {totalePezzi}</span>
                  </div>
                  <table className="w-full text-xs">
                    <thead className="bg-gray-50 border-b border-gray-100 text-[10px] font-black text-gray-500 uppercase">
                      <tr>
                        <th className="px-3 py-2 text-left">Codice</th>
                        <th className="px-3 py-2 text-left">ID Cartone</th>
                        <th className="px-3 py-2 text-left">Magazzino</th>
                        <th className="px-3 py-2 text-left">Bancale</th>
                        <th className="px-3 py-2 text-right">Disp.</th>
                        <th className="px-3 py-2 text-right">Preleva</th>
                        <th className="px-3 py-2"></th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {prelievoRighe.map(r => (
                        <tr key={r.stockId} className="hover:bg-gray-50/80">
                          <td className="px-3 py-2 font-mono font-bold text-blue-700">{r.codice}</td>
                          <td className="px-3 py-2 font-mono text-[10px] text-gray-500 truncate max-w-[140px]" title={r.idCartone}>{r.idCartone || '—'}</td>
                          <td className="px-3 py-2 text-gray-600">{r.magazzino}</td>
                          <td className="px-3 py-2 text-gray-600">{r.numero_bancale || '—'}</td>
                          <td className="px-3 py-2 text-right text-gray-400">{r.qtaDisponibile}</td>
                          <td className="px-3 py-2 text-right">
                            <input type="number" value={r.quantita}
                              onChange={e => updatePrelievoQta(r.stockId, e.target.value)}
                              className={`w-20 text-right border rounded px-1 py-0.5 text-xs font-black focus:outline-none ${parseFloat(r.quantita) > r.qtaDisponibile || !parseFloat(r.quantita) ? 'border-red-400 bg-red-50 text-red-700' : 'border-gray-300'}`} />
                          </td>
                          <td className="px-3 py-2 text-center">
                            <button onClick={() => removePrelievoRiga(r.stockId)}
                              className="text-[10px] text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-300 px-2 py-1 rounded-lg cursor-pointer transition">✕</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <div className="p-4 border-t border-gray-100">
                    <button onClick={registraPrelievo}
                      className="w-full bg-green-600 hover:bg-green-700 text-white font-black p-4 rounded-xl text-base shadow-md transition cursor-pointer flex items-center justify-center gap-2">
                      ✓ Registra prelievo ({prelievoRighe.length} righe — {totalePezzi} pz)
                    </button>
                  </div>
                </div>
              )}

              {prelievoRighe.length === 0 && (
                <div className="text-center py-16 text-gray-400 text-sm">
                  Spara un cartone o aggiungi manualmente per iniziare il prelievo.
                </div>
              )}
            </div>
          );
        })()}

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
                          <span className="text-sm font-semibold text-gray-700">Terminali</span>
                        </label>
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={filterSNNo} onChange={e => setFilterSNNo(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                          <span className="text-sm font-semibold text-gray-700">Spare Parts/Accessori</span>
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
                  {Object.values(invoiceGroups).sort((a, b) => {
                    const toISO = d => { if (!d || d === 'N/D') return ''; const p = d.split('/'); return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : d; };
                    return toISO(a.lines[0]?.arrival_date).localeCompare(toISO(b.lines[0]?.arrival_date));
                  }).map(group => (
                    <div key={`${group.invoice}_${group.snRequired}`} className="bg-gray-100/60 p-4 sm:p-5 rounded-2xl border border-gray-200/80 space-y-3">
                      <div className="flex justify-between items-center border-b border-gray-200 pb-2">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-sm sm:text-base font-black text-gray-800 tracking-tight">
                            📄 Arrivo del: {group.lines[0]?.arrival_date || '—'} — {group.invoice}
                          </span>
                        </div>
                        {!group.snRequired && (
                          <button
                            onClick={async () => {
                              setArrivoQtyInvoice(group.invoice);
                              setArrivoQtyActive(true);
                              setArrivoQtyFeedback({ text: '', type: '' });
                              // Carica cartoni pending E caricato (per "Modifica Arrivo")
                              const { data: existing } = await supabase.from('carton_arrivals')
                                .select('*').eq('invoice', group.invoice)
                                .order('data_arrivo', { ascending: true });
                              if (existing && existing.length > 0) {
                                setArrivoQtyBancale(existing[0].bancale || '');
                                setArrivoQtyMagazzino(existing[0].magazzino || 'GESSATE');
                                setArrivoQtyCartoni(existing.map(c => ({ codice: c.codice, quantita: c.quantita, idCartone: c.id_cartone, qrRaw: c.qr_raw, bancale: c.bancale, manuale: !c.qr_raw, stato: c.stato })));
                                const qtyMap = {};
                                existing.forEach(c => {
                                  if (c.po_line_key) qtyMap[c.po_line_key] = (qtyMap[c.po_line_key] || 0) + (c.quantita || 0);
                                });
                                setPoLines(prev => prev.map(l => ({ ...l, qty_loaded: qtyMap[l.unique_key] || l.qty_loaded || 0 })));
                              } else {
                                setArrivoQtyCartoni([]);
                                setArrivoQtyBancale('');
                              }
                              setCurrentView('arrivo_qty');
                            }}
                            className={`text-white text-xs font-bold px-3.5 py-2 rounded-xl transition cursor-pointer shadow-xs ${group.lines.every(l => l.is_user_confirmed) ? 'bg-gray-500 hover:bg-gray-600' : group.lines.some(l => (l.qty_loaded || 0) > 0) ? 'bg-amber-500 hover:bg-amber-600' : 'bg-blue-600 hover:bg-blue-700'}`}>
                            {group.lines.every(l => l.is_user_confirmed) ? '✎ Modifica Arrivo' : group.lines.some(l => (l.qty_loaded || 0) > 0) ? '▶ Riprendi Arrivo' : '📦 Processa Arrivo'}
                          </button>
                        )}
                      </div>
                      <div className="flex flex-col space-y-3">
                        {group.lines.map(item => {
                          const snRequired = item.sn_required == null ? true : item.sn_required;
                          const isConfirmed = item.is_user_confirmed === true;
                          return (
                            <div key={item.unique_key} className="bg-white px-3 py-2.5 rounded-xl border border-gray-200 hover:border-gray-300 transition shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-2">

                              {/* Stato */}
                              <div className="flex flex-row md:flex-col items-center md:items-start justify-between md:justify-center gap-1.5 shrink-0 border-b md:border-b-0 pb-1.5 md:pb-0 border-gray-100">
                                {isConfirmed ? (
                                  <span className="text-[10px] bg-green-600 text-white px-2 py-0.5 rounded-md font-black border border-green-700">✓ Concluso</span>
                                ) : (item.scanned_count > 0 || item.qty_loaded > 0) ? (
                                  <span className="text-[10px] bg-amber-500 text-white px-2 py-0.5 rounded-md font-black border border-amber-600">⏳ In Corso</span>
                                ) : (
                                  <span className="text-[10px] bg-gray-200 text-gray-600 px-2 py-0.5 rounded-md font-black border border-gray-300">◌ In Attesa</span>
                                )}
                              </div>

                              {/* Descrizione */}
                              <div className="flex-grow min-w-0 space-y-0.5">
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <h4 className="text-base font-black text-blue-700 font-mono tracking-tight break-all">{item.item_code}</h4>
                                  <span className="text-[10px] text-gray-400 font-mono hidden md:inline">(L: {item.line_id})</span>
                                  {!snRequired && !spareParts.some(p => p.pn === item.item_code) && (
                                    <span className="text-[9px] bg-red-50 text-red-600 border border-red-200 font-bold px-1.5 py-px rounded" title="Codice non presente nel DB Spare Parts">⚠ NON IN DB</span>
                                  )}
                                </div>
                                <p className="text-xs text-gray-500 font-medium line-clamp-1">{item.description}</p>
                                <div className="flex flex-wrap gap-x-3 text-[10px] text-gray-400 font-semibold uppercase tracking-wider">
                                  <span>Rif: <span className="text-gray-600 font-mono">{item.po_name}</span></span>
                                  {item.part_number && item.part_number !== 'N/D' && (
                                    <span>PN: <span className="text-indigo-600 font-mono font-bold">{item.part_number}</span></span>
                                  )}
                                </div>
                              </div>

                              {/* Quantità — inline, subito dopo la descrizione */}
                              <div className="flex flex-col justify-center shrink-0 text-right min-w-[80px]">
                                {/* Solo attesi: nessuna rilevazione avviata */}
                                {snRequired && !item.sn_loaded && (
                                  <span className="text-2xl font-black font-mono text-gray-900">{item.qty_expected}<span className="text-xs font-bold text-gray-400 ml-1">pz</span></span>
                                )}
                                {/* Serializzati con rilevazione */}
                                {snRequired && item.sn_loaded && (
                                  <span className="text-2xl font-black font-mono text-gray-900">
                                    {item.scanned_count > 0
                                      ? <><span className={item.scanned_count >= item.qty_expected ? 'text-green-600' : 'text-blue-600'}>{item.scanned_count}</span><span className="text-sm font-bold text-gray-400">/{item.qty_expected}</span></>
                                      : <>{item.qty_expected}<span className="text-xs font-bold text-gray-400 ml-1">pz</span></>
                                    }
                                  </span>
                                )}
                                {/* Non serializzati */}
                                {!snRequired && (
                                  <span className="text-2xl font-black font-mono">
                                    {item.qty_loaded > 0
                                      ? <><span className={item.qty_loaded > item.qty_expected ? 'text-red-600' : item.qty_loaded === item.qty_expected ? 'text-green-600' : 'text-yellow-600'}>{item.qty_loaded}</span><span className="text-sm font-bold text-gray-400">/{item.qty_expected}</span></>
                                      : <span className="text-gray-900">{item.qty_expected}<span className="text-xs font-bold text-gray-400 ml-1">pz</span></span>
                                    }
                                  </span>
                                )}
                              </div>

                              {/* Pulsanti — destra, affiancati */}
                              <div className="flex flex-row flex-wrap gap-2 justify-end shrink-0">
                                {snRequired && !isConfirmed && (
                                  <div className="relative">
                                    {item.sn_loaded ? (
                                      <>
                                        <button className="bg-gray-200 hover:bg-gray-300 text-gray-500 text-xs font-medium px-3 py-2 rounded-xl transition cursor-pointer border border-gray-300 whitespace-nowrap">Sovrascrivi SN</button>
                                        <input type="file" accept=".xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={(e) => handleSNUpload(e, item)} />
                                      </>
                                    ) : (
                                      <>
                                        <button className="bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-3 py-2 rounded-xl transition cursor-pointer shadow-xs whitespace-nowrap">Carica SN</button>
                                        <input type="file" accept=".xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={(e) => handleSNUpload(e, item)} />
                                      </>
                                    )}
                                  </div>
                                )}
                                {snRequired && isConfirmed && (
                                  <>
                                    <button onClick={() => downloadArrivoCSV(item)} className={`text-xs font-bold px-3 py-2 rounded-xl transition cursor-pointer whitespace-nowrap ${downloadedKeys.has(item.unique_key) ? 'bg-gray-100 hover:bg-gray-200 text-gray-400 border border-gray-200' : 'bg-blue-600 hover:bg-blue-700 text-white shadow-xs'}`}>📥 Scarica</button>
                                    <button onClick={() => reopenScanningSession(item)} className="bg-gray-200 hover:bg-gray-300 text-gray-500 text-xs font-medium px-3 py-2 rounded-xl transition cursor-pointer border border-gray-300 whitespace-nowrap">Modifica</button>
                                  </>
                                )}
                                {snRequired && !isConfirmed && (
                                  <button onClick={() => startScanningSession(item)} className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold px-3 py-2 rounded-xl transition cursor-pointer shadow-xs whitespace-nowrap">
                                    {item.scanned_count >= item.qty_expected && item.scanned_count > 0 ? 'Modifica' : 'Avvia'}
                                  </button>
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

        {/* ==================== VISTA ARRIVO QUANTITÀ ==================== */}
        {currentView === 'arrivo_qty' && activeModule === 'arrivi' && (
          <div className="space-y-5 max-w-3xl mx-auto">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-lg font-black text-gray-800">📦 Arrivo Quantità</h2>
                <p className="text-xs text-gray-500">Invoice: <span className="font-bold text-blue-600">{arrivoQtyInvoice}</span></p>
              </div>
              <div className="flex gap-2">
                {arrivoQtyCartoni.length > 0 && (
                  <button onClick={annullaTuttoArrivo}
                    className="bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition">
                    🗑 Annulla tutto
                  </button>
                )}
                <button onClick={async () => { setCurrentView('dashboard'); setArrivoQtyActive(false); await fetchPOLines(); }}
                  className="bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition">
                  ← Sospendi
                </button>
              </div>
            </div>

            {/* Configurazione bancale */}
            <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
              <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Bancale attivo</span>
              <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1">
                  <label className="block text-xs font-bold text-gray-500">Nome bancale <span className="text-red-500">*</span></label>
                  <input value={arrivoQtyBancale} onChange={e => setArrivoQtyBancale(e.target.value)}
                    placeholder="Es. BANCALE-001"
                    className={`w-full bg-gray-50 rounded-xl p-2.5 text-sm focus:outline-hidden font-mono border ${!arrivoQtyBancale.trim() ? 'border-red-300 bg-red-50' : 'border-gray-300'}`} />
                </div>
                <div className="space-y-1">
                  <label className="block text-xs font-bold text-gray-500">Magazzino</label>
                  <select value={arrivoQtyMagazzino} onChange={e => setArrivoQtyMagazzino(e.target.value)}
                    className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-sm focus:outline-hidden">
                    <option value="GESSATE">GESSATE</option>
                    <option value="ESPRINET">ESPRINET</option>
                  </select>
                </div>
              </div>
            </div>

            {/* Scanner QR */}
            <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
              <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Scansione QR cartone</span>
              <form onSubmit={handleArrivoQtySubmit}>
                <input type="text" ref={arrivoQtyScannerRef} value={arrivoQtyScanner}
                  onChange={e => setArrivoQtyScanner(e.target.value)}
                  autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
                  placeholder="Spara il QR del cartone..."
                  className="w-full bg-white border-2 border-blue-500 text-gray-800 font-mono text-base p-4 rounded-xl shadow-inner focus:outline-hidden text-center" />
              </form>
              {arrivoQtyFeedback.text && (
                <div className={`p-3 rounded-xl text-center text-sm font-bold border ${
                  arrivoQtyFeedback.type === 'success' ? 'bg-green-50 text-green-800 border-green-200' :
                  arrivoQtyFeedback.type === 'warning' ? 'bg-amber-50 text-amber-800 border-amber-200' :
                  'bg-red-50 text-red-800 border-red-200'}`}>
                  {arrivoQtyFeedback.text}
                </div>
              )}
            </div>


            {/* Inserimento manuale */}
            <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
              <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Inserimento manuale (cartone senza QR)</span>
              <div className="flex gap-3 items-end">
                <div className="flex-grow space-y-1">
                  <label className="block text-xs font-bold text-gray-500">Codice</label>
                  <input list="arrivo-codici-attesi" value={arrivoQtyManuale.codice} onChange={e => setArrivoQtyManuale(v => ({...v, codice: e.target.value}))}
                    placeholder="Cerca tra i codici attesi..."
                    className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs font-mono focus:outline-hidden" />
                  <datalist id="arrivo-codici-attesi">
                    {poLines.filter(l => l.china_invoice === arrivoQtyInvoice && l.sn_required === false)
                      .map(l => <option key={l.unique_key} value={l.item_code}>{l.description} — attesi {l.qty_expected}</option>)}
                  </datalist>
                </div>
                <div className="w-28 space-y-1">
                  <label className="block text-xs font-bold text-gray-500">Quantità</label>
                  <input type="number" value={arrivoQtyManuale.quantita} onChange={e => setArrivoQtyManuale(v => ({...v, quantita: e.target.value}))}
                    placeholder="96"
                    className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
                </div>
                <button onClick={addCartonManuale}
                  className="bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition shadow-xs">
                  + Aggiungi
                </button>
              </div>
            </div>

            {/* Lista cartoni */}
            {arrivoQtyCartoni.length > 0 && (() => {
              // Riepilogo per codice
              const summary = {};
              arrivoQtyCartoni.forEach(c => {
                if (!summary[c.codice]) {
                  const line = checkCodiceInPianoArrivi(c.codice);
                  summary[c.codice] = { caricata: 0, attesa: line?.qty_expected || 0 };
                }
                summary[c.codice].caricata += c.quantita || 0;
              });
              // Righe dell'invoice non ancora toccate
              const invoiceLines = poLines.filter(l => l.china_invoice === arrivoQtyInvoice && l.sn_required === false);
              invoiceLines.forEach(l => {
                const codice = l.item_code || l.part_number;
                if (!summary[codice]) summary[codice] = { caricata: 0, attesa: l.qty_expected || 0 };
              });
              const invoiceComplete = Object.values(summary).every(({ caricata, attesa }) => attesa > 0 && caricata >= attesa);
              return (
              <>
              {/* Riepilogo per codice */}
              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
                <div className="px-4 py-3 border-b border-gray-100">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider">Riepilogo per codice</span>
                </div>
                <div className="divide-y divide-gray-100">
                  {Object.entries(summary).map(([codice, { caricata, attesa }]) => {
                    const status = caricata === attesa ? 'ok' : caricata < attesa ? 'under' : 'over';
                    return (
                      <div key={codice} className={`flex items-center justify-between px-4 py-2.5 ${status === 'under' ? 'bg-yellow-50' : status === 'over' ? 'bg-red-50' : 'bg-green-50'}`}>
                        <span className="font-mono font-bold text-sm text-gray-800">{codice}</span>
                        <div className="flex items-center gap-3">
                          <span className="text-xs text-gray-500">Attesi: <strong>{attesa}</strong></span>
                          <span className={`text-sm font-black ${status === 'under' ? 'text-yellow-700' : status === 'over' ? 'text-red-700' : 'text-green-700'}`}>
                            Caricati: {caricata}
                          </span>
                          <span className="text-base">{status === 'ok' ? '✅' : status === 'under' ? '🟡' : '🔴'}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
                <div className="flex items-center justify-between px-4 py-3 border-b border-gray-100">
                  <span className="text-xs font-black text-gray-500 uppercase tracking-wider">{arrivoQtyCartoni.length} cartoni rilevati</span>
                  <span className="text-sm font-black text-blue-600">
                    Tot. pezzi: {arrivoQtyCartoni.reduce((s, c) => s + (c.quantita || 0), 0)}
                  </span>
                </div>
                <table className="w-full text-xs">
                  <thead className="bg-gray-50 border-b border-gray-100 text-[10px] font-black text-gray-500 uppercase">
                    <tr>
                      <th className="px-3 py-2 text-left">ID Cartone</th>
                      <th className="px-3 py-2 text-left">Codice</th>
                      <th className="px-3 py-2 text-left">Bancale</th>
                      <th className="px-3 py-2 text-right">Qtà</th>
                      <th className="px-3 py-2 text-center">Tipo</th>
                      <th className="px-3 py-2"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {arrivoQtyCartoni.map(c => (
                      <tr key={c.idCartone} className="hover:bg-gray-50/80">
                        <td className="px-3 py-2 font-mono text-[10px] text-gray-500 truncate max-w-[180px]" title={c.idCartone}>{c.idCartone}</td>
                        <td className="px-3 py-2 font-mono font-bold text-blue-700">{c.codice}</td>
                        <td className="px-3 py-2 text-gray-600">{c.bancale}</td>
                        <td className="px-3 py-2 text-right font-black">{c.quantita}</td>
                        <td className="px-3 py-2 text-center">
                          <div className="flex flex-col gap-0.5 items-center">
                            {c.manuale
                              ? <span className="text-[9px] bg-amber-50 text-amber-700 border border-amber-200 px-1.5 py-px rounded font-bold">MAN</span>
                              : <span className="text-[9px] bg-green-50 text-green-700 border border-green-200 px-1.5 py-px rounded font-bold">QR</span>}
                            {c.stato === 'caricato' && <span className="text-[9px] bg-green-50 text-green-700 border border-green-200 px-1.5 py-px rounded font-bold">✓</span>}
                            {c.warning && <span className="text-[9px] bg-red-50 text-red-600 border border-red-200 px-1.5 py-px rounded font-bold">⚠</span>}
                          </div>
                        </td>
                        <td className="px-3 py-2">
                          <div className="flex items-center justify-end gap-3">
                            {c.manuale && (
                              <button onClick={() => printCartonLabel(c)} className={`text-[10px] font-bold px-2.5 py-1 rounded-lg cursor-pointer transition shadow-xs whitespace-nowrap ${printedCartons.has(c.idCartone) ? 'bg-gray-200 text-gray-500 border border-gray-300' : 'bg-blue-600 hover:bg-blue-700 text-white'}`}>
                                🖨️ Stampa
                              </button>
                            )}
                            <button onClick={() => removeCarton(c.idCartone)} className="text-[10px] text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-300 px-2 py-1 rounded-lg cursor-pointer transition">✕</button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="p-4 border-t border-gray-100 space-y-2">
                  {!invoiceComplete && (
                    <p className="text-xs text-center font-bold text-amber-600">
                      ⚠ Invoice incompleta — completa tutti i codici prima di caricare
                    </p>
                  )}
                  <button
                    onClick={caricaArrivoSuInventario}
                    disabled={!invoiceComplete}
                    className={`w-full font-black p-4 rounded-xl text-base shadow-md transition flex items-center justify-center gap-2 ${!invoiceComplete ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700 text-white cursor-pointer'}`}>
                    ✓ Carica {arrivoQtyCartoni.length} cartoni su Inventario Spare Parts
                  </button>
                </div>
              </div>
              </>);
            })()}
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


      {activeModule === 'spare-parts' && Object.keys(spPendingChanges).length > 0 && (
        <div className="fixed bottom-0 left-0 right-0 z-50 bg-white border-t border-amber-200 shadow-lg px-4 py-3 flex items-center justify-between gap-4">
          <span className="text-sm font-bold text-amber-700">✏️ {Object.keys(spPendingChanges).length} riga/righe modificate non salvate</span>
          <div className="flex gap-2">
            <button onClick={() => { setSpPendingChanges({}); setSpEditMode(false); }}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-gray-100 hover:bg-gray-200 text-gray-600 cursor-pointer transition">Annulla tutto</button>
            <button onClick={saveAllSpChanges}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-green-600 hover:bg-green-700 text-white cursor-pointer transition shadow-xs">
              ✓ Salva {Object.keys(spPendingChanges).length} modifiche
            </button>
          </div>
        </div>
      )}

      {activeModule === 'stock' && Object.keys(stockPendingChanges).length > 0 && (
        <div className="fixed bottom-0 left-0 right-0 z-50 bg-white border-t border-amber-200 shadow-lg px-4 py-3 flex items-center justify-between gap-4">
          <span className="text-sm font-bold text-amber-700">
            ✏️ {Object.keys(stockPendingChanges).length} riga/righe modificate non salvate
          </span>
          <div className="flex gap-2">
            <button onClick={() => setStockPendingChanges({})}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-gray-100 hover:bg-gray-200 text-gray-600 cursor-pointer transition">
              Annulla tutto
            </button>
            <button onClick={saveAllStockChanges}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-green-600 hover:bg-green-700 text-white cursor-pointer transition shadow-xs">
              ✓ Salva {Object.keys(stockPendingChanges).length} modifiche
            </button>
          </div>
        </div>
      )}

      <footer className="w-full text-center py-4 text-xs text-gray-400 border-t border-gray-200 bg-white">
        <div className="max-w-7xl mx-auto px-4">LogiScan Enterprise &copy; 2026 - Connessione Cloud Attiva</div>
      </footer>
    </div>
  );
}
