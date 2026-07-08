import { useState, useEffect, useRef, Fragment } from 'react';
import { supabase } from './supabaseClient';
import * as XLSX from 'xlsx';
import { useColumnWidths } from './useColumnWidths';

// Moduli work-in-progress da nascondere dal menu (impostare a true per riattivarli)
// Moduli WIP visibili solo in sviluppo locale (npm run dev), nascosti nella build di produzione (Vercel)
const SHOW_WIP_MODULES = import.meta.env.DEV;
const WIP_MODULE_IDS = ['riepilogo', 'anagrafica', 'matrice'];

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
  // OK matricola SINGOLA — bip singolo acuto
  ok() {
    playBeep({ frequency: 1320, duration: 0.1, type: 'sine', volume: 0.35 });
  },
  // OK QR leggibile (ID cartone accessori / QR scatola matricole) — doppio bip
  carton() {
    playBeep({ frequency: 1180, duration: 0.09, type: 'sine', volume: 0.35, delay: 0 });
    playBeep({ frequency: 1180, duration: 0.11, type: 'sine', volume: 0.4, delay: 0.13 });
  },
  // ERRORE unificato — segnale grave e marcato (doppio tono basso discendente)
  error() {
    playBeep({ frequency: 200, duration: 0.22, type: 'square', volume: 0.3, delay: 0 });
    playBeep({ frequency: 140, duration: 0.32, type: 'square', volume: 0.3, delay: 0.2 });
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
  const scannedSetRef = useRef(new Set()); // guardia sincrona anti-duplicati
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
  const [stockNewRows, setStockNewRows] = useState([]); // righe nuove da inserire (in modalità modifica)
  const [moveBancaleOpen, setMoveBancaleOpen] = useState(false);

  // Arrivo Quantità
  const [, setArrivoQtyActive] = useState(false);
  const [arrivoQtyInvoice, setArrivoQtyInvoice] = useState('');
  const [arrivoQtyBancale, setArrivoQtyBancale] = useState('');
  const [arrivoQtyMagazzino, setArrivoQtyMagazzino] = useState('GESSATE');
  const [arrivoQtyCartoni, setArrivoQtyCartoni] = useState([]);
  const [arrivoQtyScanner, setArrivoQtyScanner] = useState('');
  const [arrivoQtyManuale, setArrivoQtyManuale] = useState({ codice: '', quantita: '' });
  const [arrivoCodiceOpen, setArrivoCodiceOpen] = useState(false);
  const [arrivoBancaleForm, setArrivoBancaleForm] = useState({ open: false, nome: '', magazzino: 'GESSATE' });
  const [riepCodiceOpen, setRiepCodiceOpen] = useState(true);
  const [arrivoQtyFeedback, setArrivoQtyFeedback] = useState({ text: '', type: '' });
  const arrivoQtyScannerRef = useRef(null);
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
  const [spNewRows, setSpNewRows] = useState([]); // righe nuove da inserire (in modalità modifica)
  const [downloadedKeys, setDownloadedKeys] = useState(new Set());

  // Prelievi
  const [prelievoView, setPrelievoView] = useState('list'); // 'list' | 'new' | 'detail'
  const [prelieviList, setPrelieviList] = useState([]);
  const [prelievoDetail, setPrelievoDetail] = useState(null); // { testata, righe }
  const [prelievoTab, setPrelievoTab] = useState('attivi'); // 'attivi' | 'registrati'
  const [prelieviLoading, setPrelieviLoading] = useState(false);
  const [prelievoUtente, setPrelievoUtente] = useState('');
  const [prelievoDest, setPrelievoDest] = useState('');
  const [prelievoRighe, setPrelievoRighe] = useState([]); // { stockId, idCartone, codice, numero_bancale, magazzino, quantita, qtaDisponibile }
  const [prelievoScanner, setPrelievoScanner] = useState('');
  const [prelievoFeedback, setPrelievoFeedback] = useState({ text: '', type: '' });
  const [prelievoManuale, setPrelievoManuale] = useState({ codice: '', stockId: '', quantita: '' });
  const [prelievoShowEsprinet, setPrelievoShowEsprinet] = useState(false);
  const prelievoScannerRef = useRef(null);

  // Anagrafica Terminali (tipoterminale.xlsx → foglio "famiglie", chiave CODICE = PNIT del DB spare parts)
  const [anagrafica, setAnagrafica] = useState([]);
  const [anagLoading, setAnagLoading] = useState(false);
  const [anagSearch, setAnagSearch] = useState('');

  // Riepilogo Stock
  const [riepSearch, setRiepSearch] = useState('');
  const [riepFilterModello, setRiepFilterModello] = useState('');
  const [riepFilterPnit, setRiepFilterPnit] = useState('');
  const [riepFilterRef, setRiepFilterRef] = useState(false);
  const [riepFilterRplus, setRiepFilterRplus] = useState(false);
  const [riepExpanded, setRiepExpanded] = useState(new Set());
  const [riepGroupMode, setRiepGroupMode] = useState('type'); // 'type' = TYPE→Gruppo, 'gruppo' = Gruppo→TYPE
  const [ordini, setOrdini] = useState({}); // codice -> { in_arrivo, in_ordine }
  const [ordiniLoading, setOrdiniLoading] = useState(false);
  const [importMeta, setImportMeta] = useState({}); // chiave -> updated_at ISO
  const [matriceSearch, setMatriceSearch] = useState('');
  const [matriceSort, setMatriceSort] = useState({ col: 'pnit', dir: 'asc' });
  const [consumi, setConsumi] = useState({}); // id (pnit+type) -> { media, nomaterial }
  const [consumiLoading, setConsumiLoading] = useState(false);
  const [mesiCopertura, setMesiCopertura] = useState(3);
  const [matriceSoloStima, setMatriceSoloStima] = useState(false);
  const [matriceSoloDaOrdinare, setMatriceSoloDaOrdinare] = useState(false);

  const scannerInputRef = useRef(null);

  useEffect(() => {
    fetchPOLines();
    fetchSpareParts();
    fetchStock();
    fetchPrelievi();
    fetchAnagrafica();
    fetchOrdini();
    fetchConsumi();
    fetchImportMeta();
    // eslint-disable-next-line react-hooks/exhaustive-deps
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

  // Helper: carica TUTTE le righe di una tabella superando il limite di 1000
  async function fetchAllRows(table, cols) {
    const pageSize = 1000;
    let all = [];
    let from = 0;
    while (true) {
      // Ordinamento per 'id' per una paginazione stabile (evita duplicati/salti sui confini di pagina)
      const { data, error } = await supabase.from(table).select(cols).order('id', { ascending: true }).range(from, from + pageSize - 1);
      if (error) break;
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    return all;
  }

  async function fetchPOLines() {
    setLoading(true);
    const [{ data, error }, scannedKeys, cartonData] = await Promise.all([
      supabase.from('po_lines').select('*').order('arrival_date', { ascending: true }),
      fetchAllRows('scanned_serials', 'po_line_key'),
      fetchAllRows('carton_arrivals', 'po_line_key, quantita, invoice, codice')
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
        .order('id', { ascending: true })
        .range(from, from + pageSize - 1);
      if (error) { alert("Errore caricamento spare parts: " + error.message); break; }
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    setSpareParts(all);
    setSpLoading(false);
  }

  async function fetchAnagrafica() {
    setAnagLoading(true);
    const pageSize = 1000;
    let all = [];
    let from = 0;
    while (true) {
      const { data, error } = await supabase.from('anagrafica_terminali').select('*').order('codice', { ascending: true }).range(from, from + pageSize - 1);
      if (error) break;
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    setAnagrafica(all);
    setAnagLoading(false);
  }

  async function handleAnagraficaUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    setAnagLoading(true);

    const reader = new FileReader();
    reader.onload = async function(e) {
      let rows;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        // Preferisci il foglio "famiglie"; altrimenti il primo
        const shName = wb.SheetNames.find(n => n.toLowerCase() === 'famiglie') || wb.SheetNames[0];
        const ws = wb.Sheets[shName];
        rows = XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false, raw: false });
      } catch {
        alert("Errore: impossibile leggere il file.");
        setAnagLoading(false);
        return;
      }
      if (rows.length === 0) { alert("File vuoto."); setAnagLoading(false); return; }

      // Ricerca flessibile delle colonne
      const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const cols = Object.keys(rows[0]);
      const colFor = (...cands) => cols.find(k => cands.includes(norm(k)));
      const cCodice = colFor('codice');
      if (!cCodice) {
        alert("Colonna obbligatoria mancante: serve 'CODICE' (foglio 'famiglie').");
        setAnagLoading(false);
        return;
      }
      const c = {
        tipo: colFor('tipo'), modello: colFor('modello'), blocco: colFor('blocco'),
        famiglia: colFor('famiglia'), famigliakey: colFor('famigliakey'),
        descrizione: colFor('descrizione'), gruppo: colFor('gruppo'), gruppo2a: colFor('gruppo2a'),
        family: colFor('family'), microfamily: colFor('microfamily'),
        note: colFor('note'), versione_pci: colFor('versionepci'),
      };
      const g = (r, col) => col ? String(r[col] || '').trim() : '';

      const seen = new Set();
      const toUpsert = rows.filter(r => {
        const cod = String(r[cCodice] || '').trim();
        if (!cod || seen.has(cod)) return false;
        seen.add(cod);
        return true;
      }).map(r => ({
        codice: String(r[cCodice] || '').trim(),
        tipo: g(r, c.tipo), modello: g(r, c.modello), blocco: g(r, c.blocco),
        famiglia: g(r, c.famiglia), famigliakey: g(r, c.famigliakey),
        descrizione: g(r, c.descrizione), gruppo: g(r, c.gruppo), gruppo2a: g(r, c.gruppo2a),
        family: g(r, c.family), microfamily: g(r, c.microfamily),
        note: g(r, c.note), versione_pci: g(r, c.versione_pci),
      }));

      // Upsert a blocchi
      const CHUNK = 500;
      let err = null;
      for (let i = 0; i < toUpsert.length; i += CHUNK) {
        const { error } = await supabase.from('anagrafica_terminali').upsert(toUpsert.slice(i, i + CHUNK), { onConflict: 'codice' });
        if (error) { err = error; break; }
      }
      if (err) alert("Errore salvataggio: " + err.message);
      else { await fetchAnagrafica(); await recordImportMeta('anagrafica'); alert(`${toUpsert.length} terminali caricati in anagrafica.`); }
      setAnagLoading(false);
    };
    reader.readAsArrayBuffer(file);
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
      else { await fetchSpareParts(); await recordImportMeta('spare_parts'); alert(`${toUpsert.length} ricambi caricati.`); }
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
    const validNewRows = spNewRows.filter(r => (r.pn || '').trim() && (r.terminal_pn || '').trim());
    if (!entries.length && validNewRows.length === 0) return;
    setSpLoading(true);
    const errors = [];
    const calc = (m) => {
      const pnit = (m.pnit || '').trim();
      const type = (m.type || '').trim();
      const eol = (m.eol || '').trim();
      const to_order = (m.to_order || '').trim();
      return {
        pnit, type, eol, to_order,
        codice: [pnit, type].filter(Boolean).join(''),
        locked: (eol === 'EOL' || eol === 'CLI' || to_order === 'NO' || type === 'NO') ? 'Y' : '',
      };
    };

    // 1. Aggiornamenti righe esistenti
    for (const [rowKey, changes] of entries) {
      const original = spareParts.find(p => `${p.pn}_${p.terminal_pn}` === rowKey);
      const merged = { ...original, ...changes };
      const c = calc(merged);
      const { error } = await supabase.from('spare_parts').update({
        eol: c.eol, english_name: merged.english_name, type: c.type, to_order: c.to_order,
        qty: parseFloat(merged.qty) || 0, descrizione: merged.descrizione,
        pnit: c.pnit, ref: merged.ref, rplus: merged.rplus,
        price: parseFloat(merged.price) || 0, locked: c.locked, codice: c.codice, modified_by: user,
      }).eq('pn', original.pn).eq('terminal_pn', original.terminal_pn);
      if (error) errors.push(error.message);
    }

    // 2. Inserimento righe nuove
    const toInsert = validNewRows.map(r => {
      const c = calc(r);
      return {
        pn: r.pn.trim(), terminal_pn: r.terminal_pn.trim(),
        eol: c.eol, english_name: (r.english_name || '').trim(), type: c.type, to_order: c.to_order,
        qty: 0, descrizione: (r.descrizione || '').trim(), pnit: c.pnit,
        ref: r.ref || '', rplus: r.rplus || '', price: parseFloat(r.price) || 0,
        locked: c.locked, codice: c.codice, modified_by: user,
      };
    });
    if (toInsert.length > 0) {
      const { error } = await supabase.from('spare_parts').upsert(toInsert, { onConflict: 'pn,terminal_pn' });
      if (error) errors.push('Nuove righe: ' + error.message);
    }

    if (errors.length) alert('Errori:\n' + errors.join('\n'));
    setSpPendingChanges({});
    setSpNewRows([]);
    setSpEditMode(false);
    await fetchSpareParts();
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
        .order('id', { ascending: true })
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
      const norm = s => String(s).trim().toUpperCase();
      let aoa = null;
      let headerIdx = -1;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        // Cerca in TUTTI i fogli quello che contiene una riga di intestazione con 'CODICE'
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false, raw: false });
          const hIdx = rows.findIndex(row => row.some(c => norm(c) === 'CODICE'));
          if (hIdx !== -1) { aoa = rows; headerIdx = hIdx; break; }
        }
      } catch {
        alert("Errore: impossibile leggere il file Excel.");
        setStockLoading(false);
        return;
      }

      if (!aoa || headerIdx === -1) {
        alert("Errore: impossibile trovare la colonna 'CODICE' in nessun foglio del file.\nControlla che il file contenga le intestazioni (LOCAZIONE, NUMERO BANCALE, MAGAZZINO, CODICE, STOCK).");
        setStockLoading(false);
        return;
      }

      const headers = aoa[headerIdx].map(h => norm(h));
      const idx = (name) => headers.indexOf(name);
      const iCod = idx('CODICE');
      const iStock = idx('STOCK');
      const iMag = idx('MAGAZZINO');
      const iBanc = headers.findIndex(h => h.includes('BANCALE'));
      const iLoc = idx('LOCAZIONE');

      // Aggregazione per (codice, magazzino, bancale): somma le quantità di stock
      const aggMap = new Map();
      aoa.slice(headerIdx + 1).forEach(row => {
        const codice = String(row[iCod] || '').trim();
        if (!codice) return;
        const magazzino = iMag >= 0 ? String(row[iMag] || '').trim() : '';
        const numero_bancale = iBanc >= 0 ? String(row[iBanc] || '').trim() : '';
        const locazione = iLoc >= 0 ? String(row[iLoc] || '').trim() : '';
        const stock = parseFloat(row[iStock]) || 0;
        const key = `${codice}__${magazzino}__${numero_bancale}`;
        if (aggMap.has(key)) {
          aggMap.get(key).stock += stock;
        } else {
          aggMap.set(key, { locazione, numero_bancale, magazzino, codice, stock });
        }
      });

      // Solo righe con stock > 0
      const toInsert = [...aggMap.values()].filter(r => r.stock !== 0);

      // GUARDIA: non svuotare l'inventario se il file non produce righe valide
      if (toInsert.length === 0) {
        alert("Il file non contiene righe valide con stock > 0. Import annullato (l'inventario attuale NON è stato modificato).");
        setStockLoading(false);
        return;
      }

      // L'import Excel è uno snapshot completo: SOSTITUISCE l'intero inventario
      if (!window.confirm(`ATTENZIONE: l'import Excel SOSTITUISCE l'intero Inventario Spare Parts.\n\nTutte le righe attuali (incluse quelle create dai processi di arrivo) verranno eliminate e rimpiazzate con le ${toInsert.length} righe del file.\n\nProcedere?`)) {
        setStockLoading(false);
        return;
      }

      // 1. Svuota la tabella
      const { error: delErr } = await supabase.from('stock_inventory').delete().gte('id', 0);
      if (delErr) { alert("Errore durante lo svuotamento: " + delErr.message); setStockLoading(false); return; }

      // 2. Inserimento a blocchi
      let insErr = null;
      for (let i = 0; i < toInsert.length; i += 500) {
        const { error } = await supabase.from('stock_inventory').insert(toInsert.slice(i, i + 500));
        if (error) { insErr = error; break; }
      }

      if (insErr) { alert("Errore salvataggio stock: " + insErr.message); }
      else { await fetchStock(); await recordImportMeta('stock'); alert(`Inventario sostituito: ${toInsert.length} record caricati.`); }
    };
    reader.readAsArrayBuffer(file);
  }

  // ===== Meta "ultimo aggiornamento" per tutti gli import CSV/Excel =====
  async function fetchImportMeta() {
    const { data } = await supabase.from('import_meta').select('*');
    const map = {};
    (data || []).forEach(r => { map[r.chiave] = r.updated_at; });
    setImportMeta(map);
  }
  async function recordImportMeta(chiave) {
    const now = new Date().toISOString();
    await supabase.from('import_meta').upsert({ chiave, updated_at: now, updated_by: currentUser || 'import' }, { onConflict: 'chiave' });
    setImportMeta(prev => ({ ...prev, [chiave]: now }));
  }

  // ===== Quantità IN ORDINE / IN ARRIVO (stock_ordini) =====
  async function fetchOrdini() {
    const pageSize = 1000;
    let all = [];
    let from = 0;
    while (true) {
      const { data, error } = await supabase.from('stock_ordini').select('*').order('codice', { ascending: true }).range(from, from + pageSize - 1);
      if (error) break;
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    const map = {};
    all.forEach(r => { map[r.codice] = { in_arrivo: r.in_arrivo || 0, in_ordine: r.in_ordine || 0 }; });
    setOrdini(map);
  }

  async function handleOrdiniUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    setOrdiniLoading(true);

    const reader = new FileReader();
    reader.onload = async function(e) {
      let rows;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false, raw: false });
      } catch {
        alert("Errore: impossibile leggere il file.");
        setOrdiniLoading(false);
        return;
      }
      if (rows.length === 0) { alert("File vuoto."); setOrdiniLoading(false); return; }

      const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const cols = Object.keys(rows[0]);
      const colFor = (...cands) => cols.find(k => cands.includes(norm(k)));
      const cPnit = colFor('pnit');
      const cPend = colFor('pendingqty');
      if (!cPnit || !cPend) {
        alert("Colonne mancanti: servono 'PNIT' e 'Pending Qty' (in ordine).");
        setOrdiniLoading(false);
        return;
      }
      const toNum = v => parseFloat(String(v || '').replace(/[^0-9.-]/g, '')) || 0;

      // Aggregazione per PNIT: somma Pending Qty (in ordine). L'in arrivo NON viene dal CSV
      // ma dal Piano Arrivi (po_lines) aggregato per codice.
      const agg = new Map();
      rows.forEach(r => {
        const cod = String(r[cPnit] || '').trim();
        if (!cod) return;
        if (cod.toUpperCase().startsWith('BULK')) return; // codice BULK non considerato
        if (!agg.has(cod)) agg.set(cod, { codice: cod, in_ordine: 0 });
        agg.get(cod).in_ordine += toNum(r[cPend]);
      });
      const nowIso = new Date().toISOString();
      const toInsert = [...agg.values()].map(a => ({
        codice: a.codice, in_arrivo: 0, in_ordine: Math.round(a.in_ordine), updated_at: nowIso,
      }));
      if (toInsert.length === 0) { alert("Nessuna riga valida (PNIT mancante)."); setOrdiniLoading(false); return; }

      // Snapshot completo: sostituisce l'intera tabella ordini/arrivi
      const { error: delErr } = await supabase.from('stock_ordini').delete().gte('codice', '');
      if (delErr) { alert("Errore svuotamento: " + delErr.message); setOrdiniLoading(false); return; }
      let insErr = null;
      for (let i = 0; i < toInsert.length; i += 500) {
        const { error } = await supabase.from('stock_ordini').insert(toInsert.slice(i, i + 500));
        if (error) { insErr = error; break; }
      }
      if (insErr) { alert("Errore salvataggio: " + insErr.message); }
      else { await fetchOrdini(); await recordImportMeta('ordini'); alert(`${toInsert.length} codici aggiornati (in ordine / in arrivo).`); }
      setOrdiniLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }

  // ===== Consumi medi mensili + NoMaterial (matrice_consumi), chiave ID = PNIT+TYPE =====
  async function fetchConsumi() {
    const pageSize = 1000;
    let all = [];
    let from = 0;
    while (true) {
      const { data, error } = await supabase.from('matrice_consumi').select('*').order('id', { ascending: true }).range(from, from + pageSize - 1);
      if (error) break;
      all = [...all, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    const map = {};
    all.forEach(r => { map[r.id] = { media: r.media || 0, nomaterial: r.nomaterial || 0 }; });
    setConsumi(map);
  }

  async function handleConsumiUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    event.target.value = '';
    setConsumiLoading(true);

    const reader = new FileReader();
    reader.onload = async function(e) {
      let rows;
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false, raw: false });
      } catch {
        alert("Errore: impossibile leggere il file.");
        setConsumiLoading(false);
        return;
      }
      if (rows.length === 0) { alert("File vuoto."); setConsumiLoading(false); return; }

      const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const cols = Object.keys(rows[0]);
      const colFor = (...cands) => cols.find(k => cands.includes(norm(k)));
      const cId = colFor('id');
      const cMedia = colFor('media');
      const cNoMat = colFor('nomaterial');
      if (!cId || !cMedia) {
        alert("Colonne mancanti: servono 'ID' (PNIT+TYPE) e 'MEDIA'.");
        setConsumiLoading(false);
        return;
      }
      const toNum = v => parseFloat(String(v || '').replace(',', '.').replace(/[^0-9.-]/g, '')) || 0;

      const seen = new Set();
      const toInsert = [];
      rows.forEach(r => {
        const id = String(r[cId] || '').trim();
        if (!id || seen.has(id)) return;
        seen.add(id);
        toInsert.push({ id, media: toNum(r[cMedia]), nomaterial: cNoMat ? toNum(r[cNoMat]) : 0, updated_at: new Date().toISOString() });
      });
      if (toInsert.length === 0) { alert("Nessuna riga valida."); setConsumiLoading(false); return; }

      // Snapshot completo
      const { error: delErr } = await supabase.from('matrice_consumi').delete().gte('id', '');
      if (delErr) { alert("Errore svuotamento: " + delErr.message); setConsumiLoading(false); return; }
      let insErr = null;
      for (let i = 0; i < toInsert.length; i += 500) {
        const { error } = await supabase.from('matrice_consumi').insert(toInsert.slice(i, i + 500));
        if (error) { insErr = error; break; }
      }
      if (insErr) { alert("Errore salvataggio: " + insErr.message); }
      else { await fetchConsumi(); await recordImportMeta('consumi'); alert(`${toInsert.length} righe consumi/NoMaterial caricate.`); }
      setConsumiLoading(false);
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

  // Aggiunge/incrementa una quantità per codice alle coordinate correnti (magazzino + bancale)
  async function addArrivoQty(codice, qty, matchedLine) {
    const bancale = arrivoQtyBancale.trim();
    const mag = arrivoQtyMagazzino;
    // Cerca una riga pending esistente per queste coordinate
    const { data: existing } = await supabase.from('carton_arrivals')
      .select('id, quantita, rilievi').eq('invoice', arrivoQtyInvoice).eq('codice', codice)
      .eq('magazzino', mag).eq('bancale', bancale).eq('stato', 'pending').limit(1);
    if (existing && existing.length > 0) {
      const { error } = await supabase.from('carton_arrivals')
        .update({ quantita: (existing[0].quantita || 0) + qty, rilievi: (existing[0].rilievi || 1) + 1 }).eq('id', existing[0].id);
      if (error) { setArrivoQtyFeedback({ text: 'Errore DB: ' + error.message, type: 'error' }); return false; }
    } else {
      const synthId = `${arrivoQtyInvoice}__${codice}__${mag}__${bancale}`;
      const { error } = await supabase.from('carton_arrivals').insert({
        id_cartone: synthId, invoice: arrivoQtyInvoice, bancale, magazzino: mag,
        codice, quantita: qty, rilievi: 1, qr_raw: null, stato: 'pending', po_line_key: matchedLine?.unique_key || null
      });
      if (error) { setArrivoQtyFeedback({ text: 'Errore DB: ' + error.message, type: 'error' }); return false; }
    }
    // Stato locale: aggrega per codice+bancale+magazzino, contando i rilievi (scansioni/inserimenti)
    setArrivoQtyCartoni(prev => {
      const i = prev.findIndex(c => c.codice === codice && c.bancale === bancale && c.magazzino === mag);
      if (i >= 0) { const copy = [...prev]; copy[i] = { ...copy[i], quantita: copy[i].quantita + qty, rilievi: (copy[i].rilievi || 1) + 1 }; return copy; }
      return [{ codice, quantita: qty, bancale, magazzino: mag, rilievi: 1 }, ...prev];
    });
    if (matchedLine) {
      setPoLines(prev => prev.map(l => l.unique_key === matchedLine.unique_key ? { ...l, qty_loaded: (l.qty_loaded || 0) + qty } : l));
    }
    return true;
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
    const matchedLine = checkCodiceInPianoArrivi(parsed.codice);
    if (!matchedLine) {
      setArrivoQtyFeedback({ text: `Codice ${parsed.codice} non previsto in arrivo per invoice ${arrivoQtyInvoice}. Bloccato.`, type: 'error' });
      sounds.error(); triggerVibration([300]);
      return;
    }
    if (await addArrivoQty(parsed.codice, parsed.quantita, matchedLine)) {
      setArrivoQtyFeedback({ text: `OK: ${parsed.codice} +${parsed.quantita} pz`, type: 'success' });
      sounds.carton(); triggerVibration([150, 100, 150]);
    }
  }

  async function addCartonManuale() {
    const { codice, quantita } = arrivoQtyManuale;
    if (!arrivoQtyBancale.trim()) { alert('Inserisci il nome del bancale prima di aggiungere.'); return; }
    if (!codice.trim() || !quantita) { alert('Inserisci codice e quantità.'); return; }
    const matchedLine = checkCodiceInPianoArrivi(codice.trim());
    if (!matchedLine) {
      alert(`Il codice "${codice.trim()}" non è previsto in arrivo per invoice ${arrivoQtyInvoice}.\n\nOperazione bloccata.`);
      return;
    }
    const qty = parseFloat(quantita);
    if (await addArrivoQty(codice.trim(), qty, matchedLine)) {
      setArrivoQtyManuale({ codice: '', quantita: '' });
      setArrivoQtyFeedback({ text: `Aggiunto: ${codice.trim()} +${qty} pz`, type: 'success' });
      sounds.ok(); triggerVibration([150]);
    }
  }

  async function removeCarton(item) {
    await supabase.from('carton_arrivals').delete()
      .eq('invoice', arrivoQtyInvoice).eq('codice', item.codice)
      .eq('magazzino', item.magazzino).eq('bancale', item.bancale).eq('stato', 'pending');
    setArrivoQtyCartoni(prev => prev.filter(c => !(c.codice === item.codice && c.bancale === item.bancale && c.magazzino === item.magazzino)));
    await fetchPOLines();
  }

  async function annullaTuttoArrivo() {
    const total = arrivoQtyCartoni.length;
    if (!window.confirm(`Annullare le ${total} righe in corso per invoice ${arrivoQtyInvoice}?\n\nVerranno rimosse (lo stock non ancora registrato non viene toccato).`)) return;

    // Elimina i pending dell'invoice e resetta lo stato riga
    await supabase.from('carton_arrivals').delete().eq('invoice', arrivoQtyInvoice).eq('stato', 'pending');
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
    if (arrivoQtyCartoni.length === 0) { alert('Nessuna riga da caricare.'); return; }
    if (!window.confirm(`Caricare ${arrivoQtyCartoni.length} righe sull'Inventario Spare Parts?\n\nInvoice: ${arrivoQtyInvoice}`)) return;

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

    // 2. Applica allo stock secondo le coordinate (codice + magazzino + bancale): incrementa o crea la riga
    for (const c of arrivoQtyCartoni) {
      const { data: existing } = await supabase.from('stock_inventory')
        .select('id, stock').eq('codice', c.codice).eq('magazzino', c.magazzino).eq('numero_bancale', c.bancale).limit(1);
      if (existing && existing.length > 0) {
        const { error } = await supabase.from('stock_inventory')
          .update({ stock: (existing[0].stock || 0) + c.quantita }).eq('id', existing[0].id);
        if (error) errors.push(`${c.codice}: ${error.message}`);
      } else {
        const { error } = await supabase.from('stock_inventory').insert({
          codice: c.codice, stock: c.quantita, numero_bancale: c.bancale, magazzino: c.magazzino, locazione: '',
        });
        if (error) errors.push(`${c.codice}: ${error.message}`);
      }
    }

    if (errors.length) {
      alert('Completato con errori:\n' + errors.join('\n'));
    } else {
      alert(`${arrivoQtyCartoni.length} righe caricate sull'inventario.`);
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

  async function exportPrelieviRettifica() {
    // Esporta solo i prelievi ATTIVI (non ancora registrati)
    const attivi = prelieviList.filter(p => p.stato !== 'registrato');
    if (attivi.length === 0) { alert('Nessun prelievo attivo da esportare.'); return; }
    if (!window.confirm(`Esportare ${attivi.length} prelievi?\n\nDopo il download verranno marcati come REGISTRATI e non saranno più eliminabili né modificabili.`)) return;

    setPrelieviLoading(true);
    const attiviIds = new Set(attivi.map(p => String(p.id)));
    const [{ data: testate }, { data: righe }] = await Promise.all([
      supabase.from('prelievi').select('*'),
      supabase.from('prelievi_righe').select('*')
    ]);
    const testataById = {};
    (testate || []).forEach(t => { if (attiviIds.has(String(t.id))) testataById[t.id] = t; });

    // Raggruppa per prelievo → (codice + magazzino) sommando le quantità (solo prelievi attivi)
    const byPrelievo = {};
    (righe || []).forEach(r => {
      if (!attiviIds.has(String(r.prelievo_id))) return;
      if (!byPrelievo[r.prelievo_id]) byPrelievo[r.prelievo_id] = {};
      const key = `${r.codice}__${r.magazzino}`;
      if (!byPrelievo[r.prelievo_id][key]) byPrelievo[r.prelievo_id][key] = { codice: r.codice, magazzino: r.magazzino, qty: 0 };
      byPrelievo[r.prelievo_id][key].qty += (r.quantita || 0);
    });

    const locName = (mag) => mag === 'ESPRINET' ? 'ESVAL - PAX NUOVO CAVENAGO' : '0MAG03 - PAX NUOVO GESSATE';
    const fmtDate = (iso) => { if (!iso) return ''; const d = new Date(iso); return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; };

    const header = 'data;line;item;Inventory Location: Name;Adjust Qty. By;extid;[PAX] Document Type';
    const rows = [];
    // ordina per id_prelievo per raggruppare le righe dello stesso extid
    Object.keys(byPrelievo)
      .sort((a, b) => (testataById[a]?.id_prelievo || '').localeCompare(testataById[b]?.id_prelievo || ''))
      .forEach(pid => {
        const t = testataById[pid];
        const data = fmtDate(t?.data_prelievo);
        const extid = t?.id_prelievo || '';
        const docType = t?.destinazione || '';
        let line = 0;
        Object.values(byPrelievo[pid]).forEach(g => {
          line += 1;
          rows.push([data, line, g.codice, locName(g.magazzino), -Math.abs(g.qty), extid, docType].join(';'));
        });
      });

    if (rows.length === 0) { alert('Nessun prelievo da esportare.'); setPrelieviLoading(false); return; }
    const csv = [header, ...rows].join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url;
    a.download = `rettifica_prelievi_${new Date().toISOString().slice(0,10)}.csv`; a.click();
    URL.revokeObjectURL(url);

    // Marca i prelievi esportati come REGISTRATI (non più eliminabili/modificabili)
    const idsToMark = [...attiviIds];
    const { error: markErr } = await supabase
      .from('prelievi')
      .update({ stato: 'registrato' })
      .in('id', idsToMark);
    if (markErr) {
      alert('File scaricato, ma errore nel marcare i prelievi come registrati: ' + markErr.message);
    } else {
      setPrelieviList(prev => prev.map(p => attiviIds.has(String(p.id)) ? { ...p, stato: 'registrato' } : p));
      setPrelievoTab('registrati');
    }
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
    if (prelievo.stato === 'registrato') { alert('Prelievo registrato: non può essere eliminato.'); return; }
    if (!window.confirm(`Eliminare il prelievo ${prelievo.id_prelievo}?\n\nLe ${prelievo.n_righe} righe verranno ripristinate nell'inventario.`)) return;
    setPrelieviLoading(true);
    const errors = [];

    // 1. Recupera le righe del prelievo
    const { data: righe, error: errR } = await supabase.from('prelievi_righe').select('*').eq('prelievo_id', prelievo.id);
    if (errR) { alert('Errore: ' + errR.message); setPrelieviLoading(false); return; }

    // 2. Ripristina lo stock per ogni riga: incrementa alle coordinate (codice, magazzino, bancale)
    for (const r of (righe || [])) {
      const { data: existing } = await supabase.from('stock_inventory')
        .select('id, stock').eq('codice', r.codice).eq('magazzino', r.magazzino).eq('numero_bancale', r.numero_bancale).limit(1);
      if (existing && existing.length > 0) {
        const { error } = await supabase.from('stock_inventory').update({ stock: (existing[0].stock || 0) + (r.quantita || 0) }).eq('id', existing[0].id);
        if (error) errors.push(error.message);
      } else {
        const { error } = await supabase.from('stock_inventory').insert({
          codice: r.codice, stock: r.quantita, numero_bancale: r.numero_bancale, magazzino: r.magazzino, locazione: '',
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
      sounds.error(); triggerVibration([300]);
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
    return true;
  }

  function handlePrelievoScan(e) {
    e.preventDefault();
    const raw = (prelievoScannerRef.current?.value || prelievoScanner).trim();
    setPrelievoScanner('');
    if (prelievoScannerRef.current) prelievoScannerRef.current.value = '';
    if (!raw) return;

    // Il QR fornisce codice + quantità: si precompila il form, poi l'utente sceglie l'ubicazione e conferma con +
    const parsed = parseCartonQR(raw);
    if (parsed) {
      setPrelievoManuale({ codice: parsed.codice, stockId: '', quantita: parsed.quantita });
      setPrelievoFeedback({ text: `Codice ${parsed.codice} (qtà ${parsed.quantita}) precompilato: seleziona l'ubicazione (bancale) e premi +.`, type: 'success' });
      sounds.carton(); triggerVibration([150, 100, 150]);
      return;
    }

    setPrelievoFeedback({ text: `QR non riconosciuto: ${raw}`, type: 'error' });
    sounds.error(); triggerVibration([300]);
  }

  function addPrelievoManuale() {
    const { stockId, quantita } = prelievoManuale;
    if (!stockId) { alert('Seleziona la riga di stock (bancale).'); return; }
    const stockRow = stockItems.find(s => String(s.id) === String(stockId));
    if (!stockRow) { alert('Riga di stock non trovata.'); return; }
    if (addPrelievoRiga(stockRow, quantita || stockRow.stock)) {
      setPrelievoManuale({ codice: '', stockId: '', quantita: '' });
      setPrelievoFeedback({ text: `Aggiunto: ${stockRow.codice} — qtà ${quantita || stockRow.stock}`, type: 'success' });
      sounds.ok(); triggerVibration([150]);
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
    const validNewRows = stockNewRows.filter(r => (r.codice || '').trim());
    if (entries.length === 0 && validNewRows.length === 0) return;
    setStockLoading(true);
    let errors = [];

    // 1. Aggiorna righe esistenti modificate
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

    // 2. Inserisci righe nuove
    const toInsert = validNewRows.map(r => ({
      locazione: (r.locazione || '').trim(),
      numero_bancale: (r.numero_bancale || '').trim(),
      magazzino: r.magazzino || 'GESSATE',
      codice: r.codice.trim(),
      stock: parseFloat(r.stock) || 0,
    }));
    if (toInsert.length > 0) {
      const { error } = await supabase.from('stock_inventory').insert(toInsert);
      if (error) errors.push('Nuove righe: ' + error.message);
    }

    if (errors.length > 0) { alert("Errori: " + errors.join('\n')); }
    else {
      setStockPendingChanges({});
      setStockNewRows([]);
      setStockEditMode(false);
      await fetchStock();
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

      // Trova in modo flessibile la colonna del vendor part number (ignora maiuscole/spazi/parentesi)
      const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const findVendorPN = (row) => {
        const key = Object.keys(row).find(k => {
          const n = norm(k);
          return n.includes('vendorpartnumber') || n.includes('vendorpn') || n === 'vpn';
        });
        return key ? String(row[key] || '').trim() : '';
      };

      // Controllo righe duplicate nel file (stessa PO INTERNAL ID + Line ID): il file NON deve contenerle
      const keyCount = new Map();
      records.forEach(row => {
        const key = `${row["PO INTERNAL ID"]}_${row["Line ID"]}`;
        keyCount.set(key, (keyCount.get(key) || 0) + 1);
      });
      const dupKeys = [...keyCount.entries()].filter(([, n]) => n > 1);
      if (dupKeys.length > 0) {
        const dettaglio = dupKeys.slice(0, 15).map(([k, n]) => {
          const [po, line] = k.split('_');
          return `• PO ${po} / Linea ${line} — ${n} volte`;
        }).join('\n');
        alert(`CARICAMENTO BLOCCATO — il file contiene ${dupKeys.length} righe duplicate (stessa PO INTERNAL ID + Line ID):\n\n${dettaglio}${dupKeys.length > 15 ? '\n...' : ''}\n\nCorreggi il file di origine rimuovendo i duplicati e ricaricalo.`);
        setLoading(false);
        return;
      }

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
          part_number: findVendorPN(row),
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
        await recordImportMeta('po_lines');
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

      if (snRecords.length === 0) {
        alert("Errore: il file è vuoto o non leggibile.");
        setLoading(false);
        return;
      }

      const cols = Object.keys(snRecords[0]);
      const hasSN = cols.includes('SN');
      const hasPN = cols.includes('PN');

      const errors = [];

      // 1. Presenza colonne obbligatorie
      if (!hasSN) errors.push(`• Colonna "SN" mancante nel file.`);
      if (!hasPN) errors.push(`• Colonna "PN" mancante nel file.`);

      const validRows = hasSN ? snRecords.filter(row => String(row['SN']).trim() !== '') : [];
      const qtyProvided = validRows.length;
      const qtyExpected = line.qty_expected;

      // 2. Corrispondenza quantità
      if (hasSN && qtyProvided !== qtyExpected) {
        errors.push(`• QUANTITÀ non corrispondente:\n   Attese: ${qtyExpected} pz — Fornite: ${qtyProvided} pz`);
      }

      // 3. Corrispondenza VPN / PN (case/spazi-insensitive)
      const normPN = s => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
      const expectedPN = (line.part_number || '').trim();
      if (!expectedPN || expectedPN === 'N/D') {
        errors.push(`• VPN non presente sulla riga del piano arrivi: impossibile verificare la corrispondenza Part Number.`);
      } else if (hasPN) {
        const expN = normPN(expectedPN);
        const mismatchedPNs = [...new Set(
          validRows.map(row => String(row['PN'] || '').trim()).filter(pn => pn !== '' && normPN(pn) !== expN)
        )];
        if (mismatchedPNs.length > 0) {
          errors.push(`• PART NUMBER non corrispondente:\n   VPN atteso: ${expectedPN}\n   PN nel file: ${mismatchedPNs.slice(0, 5).join(', ')}${mismatchedPNs.length > 5 ? '...' : ''}`);
        }
      }

      // Controlli BLOCCANTI: se anche solo uno fallisce, il caricamento è impedito
      if (errors.length > 0) {
        alert(`CARICAMENTO BLOCCATO — il file non rispetta i requisiti:\n\n${errors.join('\n\n')}\n\nCorreggi il file e riprova.`);
        setLoading(false);
        return;
      }

      // Avviso (non bloccante) sui serial duplicati: vengono caricati solo gli univoci
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

      const fullyMatched = true;

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
    const fetchAllByLine = async (table, cols) => {
      const pageSize = 1000;
      let all = [];
      let from = 0;
      while (true) {
        // Ordinamento per 'id' (chiave univoca) per una paginazione stabile: evita duplicati/salti sui confini di pagina
        const { data, error } = await supabase.from(table).select(cols).eq('po_line_key', line.unique_key)
          .order('id', { ascending: true }).range(from, from + pageSize - 1);
        if (error) { alert(`Errore caricamento ${table}: ${error.message}`); break; }
        all = [...all, ...(data || [])];
        if (!data || data.length < pageSize) break;
        from += pageSize;
      }
      return all;
    };

    const expectedData = await fetchAllByLine('expected_serials', 'serial, model, pn');
    const scannedData = await fetchAllByLine('scanned_serials', 'serial, model, pn, scanned_at');

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
    scannedSetRef.current = new Set(formattedScanned.map(s => s.serial));

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
    // Fetch paginato per superare il limite di 1000 righe di Supabase
    const pageSize = 1000;
    let scannedData = [];
    let from = 0;
    let fetchErr = null;
    while (true) {
      const { data, error } = await supabase
        .from('scanned_serials')
        .select('serial, model, pn, scanned_at')
        .eq('po_line_key', line.unique_key)
        .order('id', { ascending: true })
        .range(from, from + pageSize - 1);
      if (error) { fetchErr = error; break; }
      scannedData = [...scannedData, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    setLoading(false);
    if (fetchErr) { alert("Errore nel recupero dati: " + fetchErr.message); return; }

    buildAndDownloadCSV(line, scannedData);
  }

  function buildAndDownloadCSV(line, serials) {
    // Nessun dedup: il file riporta TUTTE le matricole registrate (anche i doppioni) per piena tracciabilità
    serials = serials || [];
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
    scannedSetRef.current.delete(serial);
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
    scannedSetRef.current = new Set();
    setScannedSerials([]);
    setCartonsScanned(0);
    setFeedback({ text: 'Rilevazioni azzerate. Scanner pronto.', type: 'success' });
    await fetchPOLines();
  }

  async function removeDuplicateSerials() {
    // Rimuove le occorrenze extra dei seriali duplicati, mantenendone UNA per seriale
    setLoading(true);
    // Recupera tutte le righe con id (paginato)
    const pageSize = 1000;
    let allRows = [];
    let from = 0;
    while (true) {
      const { data, error } = await supabase.from('scanned_serials')
        .select('id, serial').eq('po_line_key', activeLineKey)
        .order('id', { ascending: true }).range(from, from + pageSize - 1);
      if (error) { alert("Errore: " + error.message); setLoading(false); return; }
      allRows = [...allRows, ...(data || [])];
      if (!data || data.length < pageSize) break;
      from += pageSize;
    }
    const seen = new Set();
    const extraIds = [];
    allRows.forEach(r => {
      if (seen.has(r.serial)) extraIds.push(r.id);
      else seen.add(r.serial);
    });
    if (extraIds.length === 0) { setLoading(false); return; }
    // Elimina a blocchi
    for (let i = 0; i < extraIds.length; i += 500) {
      await supabase.from('scanned_serials').delete().in('id', extraIds.slice(i, i + 500));
    }
    // Ricarica lo stato locale
    setScannedSerials(prev => {
      const s = new Set();
      return prev.filter(x => { if (s.has(x.serial)) return false; s.add(x.serial); return true; });
    });
    scannedSetRef.current = seen;
    setFeedback({ text: `${extraIds.length} doppioni eliminati.`, type: 'success' });
    await fetchPOLines();
    setLoading(false);
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

    // Il totale da rilevare è il numero di matricole ATTESE (tutte devono essere lette per completare)
    const totalExpected = Object.keys(expectedSerials).length || activeLine?.qty_expected || 0;
    if (scannedSerials.length >= totalExpected && totalExpected > 0) {
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
        if (scannedSetRef.current.has(serial)) { duplicati++; return; }
        if (!Object.hasOwn(expectedSerials, serial)) { errati++; return; }
        scannedSetRef.current.add(serial); // guardia sincrona: previene doppi da scansioni rapide o QR con seriali ripetuti
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
        triggerVibration([150, 100, 150]); sounds.carton();
        if (updatedScanned.length >= totalExpected) {
          setTimeout(() => openReviewSession(), 1200);
        }
        setFeedback({ text: `Cartone Rilevato! +${aggiunti.length} matricole acquisite.`, type: 'success' });
      } else {
        triggerVibration([300]); sounds.error();
        const details = [duplicati > 0 && `${duplicati} già lette`, errati > 0 && `${errati} non appartenenti`].filter(Boolean).join(', ');
        setFeedback({ text: `Cartone scartato (${details || 'nessuna matricola valida'}).`, type: 'error' });
      }

    } else {
      const serial = rawInput;
      if (scannedSetRef.current.has(serial)) {
        triggerVibration([300]); sounds.error();
        setFeedback({ text: `Duplicato! Matricola ${serial} già letta.`, type: 'error' });
        return;
      }
      if (!Object.hasOwn(expectedSerials, serial)) {
        triggerVibration([300]); sounds.error();
        setFeedback({ text: `Errore! La matricola ${serial} non appartiene a questa riga.`, type: 'error' });
        return;
      }
      scannedSetRef.current.add(serial); // guardia sincrona anti-duplicati
      const meta = expectedSerials[serial];
      await supabase.from('scanned_serials').insert({ po_line_key: activeLineKey, serial, model: meta.model, pn: meta.pn });

      const updatedScanned = [
        { serial, model: meta.model, pn: meta.pn, time: new Date().toLocaleTimeString('it-IT') },
        ...scannedSerials
      ];
      setScannedSerials(updatedScanned);

      triggerVibration([150]); sounds.ok();
      if (updatedScanned.length >= totalExpected) {
        setFeedback({ text: `Completato! Ultima matricola: ${meta.model}`, type: 'success' });
        setTimeout(() => openReviewSession(), 1200);
      } else {
        setFeedback({ text: `OK: Rilevato modello ${meta.model}`, type: 'success' });
      }
    }
  }

  async function openReviewSession() {
    // Rilegge TUTTE le matricole dal DB (paginato) così la revisione riflette esattamente ciò che c'è nel database
    if (activeLineKey) {
      setLoading(true);
      const pageSize = 1000;
      let all = [];
      let from = 0;
      while (true) {
        const { data, error } = await supabase.from('scanned_serials')
          .select('serial, model, pn, scanned_at').eq('po_line_key', activeLineKey)
          .order('id', { ascending: true }).range(from, from + pageSize - 1);
        if (error) break;
        all = [...all, ...(data || [])];
        if (!data || data.length < pageSize) break;
        from += pageSize;
      }
      all.reverse(); // più recenti in cima
      setScannedSerials(all.map(s => ({ serial: s.serial, model: s.model, pn: s.pn, time: new Date(s.scanned_at).toLocaleTimeString('it-IT') })));
      setLoading(false);
    }
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
                { id: 'riepilogo', label: 'Riepilogo Stock', icon: '📊' },
                { id: 'matrice', label: 'Matrice PNIT × TYPE', icon: '🧮' },
                { id: 'prelievi', label: 'Prelievi', icon: '📤' },
                { id: 'anagrafica', label: 'Anagrafica', icon: '📇' },
              ].filter(mod => SHOW_WIP_MODULES || !WIP_MODULE_IDS.includes(mod.id)).map(mod => (
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
              {activeModule === 'riepilogo' && 'Riepilogo Stock'}
              {activeModule === 'matrice' && 'Matrice PNIT × TYPE'}
              {activeModule === 'prelievi' && 'Prelievi'}
              {activeModule === 'anagrafica' && 'Anagrafica'}
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
                <div className="flex-grow">
                  <p className="text-xs text-gray-400">Importa il file Excel del DB Spare Parts (Foglio 1). L&apos;import aggiorna i record esistenti per PN + Terminal PN.</p>
                  {importMeta.spare_parts && <p className="text-[11px] text-gray-400">Ultimo aggiornamento: {new Date(importMeta.spare_parts).toLocaleString('it-IT')}</p>}
                </div>
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
                    <button onClick={() => { setSpEditMode(v => !v); if (spEditMode) setSpNewRows([]); }}
                      className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer ${spEditMode ? 'bg-amber-500 text-white border-amber-600' : 'bg-gray-50 text-gray-500 hover:bg-gray-100 border-gray-200'}`}>
                      {spEditMode ? '✏️ Modifica ON' : '✏️ Modifica'}
                    </button>
                    {spEditMode && (
                      <button onClick={() => setSpNewRows(prev => [{ tempId: Date.now() + Math.random(), pn: '', terminal_pn: '', pnit: '', english_name: '', descrizione: '', type: '', eol: '', ref: '', rplus: '', to_order: 'SI', price: '' }, ...prev])}
                        className="text-[10px] font-bold text-white bg-green-600 hover:bg-green-700 px-2.5 py-1 rounded-lg transition cursor-pointer">
                        + Aggiungi riga
                      </button>
                    )}
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
                      {/* Righe nuove in inserimento (modalità modifica) */}
                      {spEditMode && spNewRows.map(r => {
                        const nCls = "w-full bg-green-50 border border-green-300 rounded px-1 py-0.5 text-xs focus:outline-none";
                        const upd = (field, val) => setSpNewRows(prev => prev.map(x => x.tempId === r.tempId ? { ...x, [field]: val } : x));
                        const locked = (r.eol === 'EOL' || r.eol === 'CLI' || r.to_order === 'NO' || r.type === 'NO') ? 'Y' : '';
                        return (
                          <tr key={r.tempId} className="bg-green-50/40">
                            <td className="px-2 py-1"><input className={nCls + ' font-mono font-bold text-blue-700'} value={r.pn} onChange={e => upd('pn', e.target.value)} placeholder="PN *" /></td>
                            <td className="px-2 py-1"><input className={nCls + ' font-mono text-[10px]'} value={r.terminal_pn} onChange={e => upd('terminal_pn', e.target.value)} placeholder="Terminal PN *" /></td>
                            <td className="px-2 py-2 font-mono text-[10px] text-gray-400">{(r.terminal_pn || '').split('-')[0]}</td>
                            <td className="px-2 py-1"><input className={nCls + ' font-mono'} value={r.pnit} onChange={e => upd('pnit', e.target.value)} placeholder="PNIT" /></td>
                            <td className="px-2 py-1 hidden md:table-cell"><input className={nCls} value={r.english_name} onChange={e => upd('english_name', e.target.value)} placeholder="English Name" /></td>
                            <td className="px-2 py-1 hidden md:table-cell"><input className={nCls} value={r.descrizione} onChange={e => upd('descrizione', e.target.value)} placeholder="Descrizione" /></td>
                            <td className="px-2 py-1"><input className={nCls} value={r.type} onChange={e => upd('type', e.target.value)} placeholder="TYPE" /></td>
                            <td className="px-1 py-1 text-center">
                              <select className={nCls} value={r.eol} onChange={e => upd('eol', e.target.value)}>
                                <option value="">—</option><option value="EOL">EOL</option><option value="ALT">ALT</option><option value="CLI">CLI</option>
                              </select>
                            </td>
                            <td className="px-1 py-1 text-center">
                              <select className={nCls} value={r.ref} onChange={e => upd('ref', e.target.value)}><option value="">—</option><option value="X">X</option></select>
                            </td>
                            <td className="px-1 py-1 text-center">
                              <select className={nCls} value={r.rplus} onChange={e => upd('rplus', e.target.value)}><option value="">—</option><option value="X">X</option></select>
                            </td>
                            <td className="px-1 py-1 text-center">
                              <select className={nCls} value={r.to_order} onChange={e => upd('to_order', e.target.value)}><option value="SI">SI</option><option value="NO">NO</option></select>
                            </td>
                            <td className="px-2 py-1 text-right"><input className={nCls + ' text-right'} value={r.price} onChange={e => upd('price', e.target.value)} placeholder="0" /></td>
                            <td className="px-2 py-2 text-center text-[11px]">{locked === 'Y' ? '🔒' : '—'}</td>
                            <td className="px-2 py-2 text-center">
                              <button onClick={() => setSpNewRows(prev => prev.filter(x => x.tempId !== r.tempId))}
                                className="text-[10px] text-gray-400 hover:text-red-500 cursor-pointer" title="Rimuovi riga">✕</button>
                            </td>
                          </tr>
                        );
                      })}
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
                <div className="flex-grow">
                  <p className="text-xs text-gray-400">Importa il file Excel dello Stock (Foglio 2). Chiave: Codice + Magazzino + Bancale.</p>
                  {importMeta.stock && <p className="text-[11px] text-gray-400">Ultimo aggiornamento: {new Date(importMeta.stock).toLocaleString('it-IT')}</p>}
                </div>
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
                    <button onClick={() => { setStockEditMode(v => !v); if (stockEditMode) setStockNewRows([]); }}
                      className={`text-[10px] font-bold px-2.5 py-1 rounded-lg border transition cursor-pointer ${stockEditMode ? 'bg-amber-500 text-white border-amber-600' : 'bg-gray-50 text-gray-500 hover:bg-gray-100 border-gray-200'}`}>
                      {stockEditMode ? '✏️ Modifica ON' : '✏️ Modifica'}
                    </button>
                    {stockEditMode && (
                      <button onClick={() => setStockNewRows(prev => [{ tempId: Date.now() + Math.random(), locazione: '', numero_bancale: '', magazzino: 'GESSATE', codice: '', stock: '' }, ...prev])}
                        className="text-[10px] font-bold text-white bg-green-600 hover:bg-green-700 px-2.5 py-1 rounded-lg transition cursor-pointer">
                        + Aggiungi riga
                      </button>
                    )}
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
                      {/* Righe nuove in inserimento (modalità modifica) */}
                      {stockEditMode && stockNewRows.map(r => {
                        const nCls = "w-full bg-green-50 border border-green-300 rounded px-1 py-0.5 text-xs focus:outline-none";
                        const upd = (field, val) => setStockNewRows(prev => prev.map(x => x.tempId === r.tempId ? { ...x, [field]: val } : x));
                        return (
                          <tr key={r.tempId} className="bg-green-50/40">
                            <td className="px-2 py-1"><input className={nCls} value={r.locazione} onChange={e => upd('locazione', e.target.value)} placeholder="Locazione" /></td>
                            <td className="px-2 py-1"><input className={nCls} value={r.numero_bancale} onChange={e => upd('numero_bancale', e.target.value)} placeholder="Bancale" /></td>
                            <td className="px-2 py-1">
                              <select className={nCls} value={r.magazzino} onChange={e => upd('magazzino', e.target.value)}>
                                <option value="GESSATE">GESSATE</option><option value="ESPRINET">ESPRINET</option>
                              </select>
                            </td>
                            <td className="px-2 py-1"><input className={nCls + ' font-mono font-bold text-blue-700'} value={r.codice} onChange={e => upd('codice', e.target.value)} placeholder="Codice *" /></td>
                            <td className="px-2 py-2 text-gray-300 text-[10px]">nuovo</td>
                            <td className="px-1 py-2"></td>
                            <td className="px-2 py-2"></td>
                            <td className="px-2 py-2"></td>
                            <td className="px-2 py-1 text-right"><input className={nCls + ' text-right font-mono font-black'} value={r.stock} onChange={e => upd('stock', e.target.value)} placeholder="0" /></td>
                            <td className="px-2 py-2 text-center">
                              <button onClick={() => setStockNewRows(prev => prev.filter(x => x.tempId !== r.tempId))}
                                className="text-[10px] text-gray-400 hover:text-red-500 cursor-pointer" title="Rimuovi riga">✕</button>
                            </td>
                          </tr>
                        );
                      })}
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

        {/* ==================== MODULO RIEPILOGO STOCK ==================== */}
        {activeModule === 'riepilogo' && (() => {
          // Stock totale per codice
          const stockByCodice = {};
          stockItems.forEach(s => { if (s.stock > 0) stockByCodice[s.codice] = (stockByCodice[s.codice] || 0) + s.stock; });

          // IN ARRIVO per codice: dal Piano Arrivi (po_lines), somma qty_expected per item_code
          const arrivoByCodice = {};
          poLines.forEach(l => {
            const cod = String(l.item_code || '').trim();
            if (!cod || cod === 'N/D') return;
            arrivoByCodice[cod] = (arrivoByCodice[cod] || 0) + (l.qty_expected || 0);
          });

          // Anagrafica terminali per codice (= PNIT): gruppo = MODELLO - VERSIONE PCI - NOTE
          const anagByCodice = {};
          anagrafica.forEach(a => { anagByCodice[String(a.codice || '').trim()] = a; });
          const gruppoDesc = (a) => a ? [a.modello, a.versione_pci, a.note].map(x => String(x || '').trim()).filter(Boolean).join(' - ') : '';

          // Info per pn dal DB spare parts
          const pnInfo = {};
          spareParts.forEach(p => {
            const mod = (p.terminal_pn || '').split('-')[0];
            if (!pnInfo[p.pn]) pnInfo[p.pn] = { models: new Set(), gruppi: new Set(), type: '', ref: '', rplus: '', descrizione: '', eol: '', pnits: new Set() };
            const inf = pnInfo[p.pn];
            if (mod) inf.models.add(mod);
            if (!inf.type && p.type) inf.type = p.type;
            if (!inf.ref && p.ref) inf.ref = p.ref;
            if (!inf.rplus && p.rplus) inf.rplus = p.rplus;
            if (!inf.eol && p.eol) inf.eol = p.eol;
            if (!inf.descrizione && (p.descrizione || p.english_name)) inf.descrizione = p.descrizione || p.english_name;
            const pnit = (p.pnit || '').trim();
            if (pnit) {
              inf.pnits.add(pnit);
              const g = gruppoDesc(anagByCodice[pnit]);
              if (g) inf.gruppi.add(g);
            }
          });

          // Liste filtri CONNESSE: la selezione su un filtro restringe le opzioni dell'altro
          const tuttiModelli = new Set();
          const tuttiPnit = new Set();
          Object.values(pnInfo).forEach(inf => {
            if (!riepFilterPnit || inf.pnits.has(riepFilterPnit)) {
              inf.models.forEach(m => tuttiModelli.add(m));
            }
            if (!riepFilterModello || inf.models.has(riepFilterModello)) {
              inf.pnits.forEach(p => tuttiPnit.add(p));
            }
          });
          if (riepFilterModello) tuttiModelli.add(riepFilterModello);
          if (riepFilterPnit) tuttiPnit.add(riepFilterPnit);
          const modelliList = [...tuttiModelli].sort();
          const pnitList = [...tuttiPnit].sort();

          const q = riepSearch.trim().toLowerCase();
          // Costruisci elenco codici con stock, applicando filtri
          const codici = Object.keys(stockByCodice).filter(c => {
            const inf = pnInfo[c] || { models: new Set(), gruppi: new Set(), type: '', ref: '', rplus: '', descrizione: '', pnits: new Set() };
            if (riepFilterRef && inf.ref !== 'X') return false;
            if (riepFilterRplus && inf.rplus !== 'X') return false;
            if (riepFilterModello && !inf.models.has(riepFilterModello)) return false;
            if (riepFilterPnit && !inf.pnits.has(riepFilterPnit)) return false;
            if (q) {
              const hay = `${c} ${inf.type} ${inf.descrizione} ${[...inf.models].join(' ')} ${[...inf.gruppi].join(' ')} ${[...inf.pnits].join(' ')}`.toLowerCase();
              if (!hay.includes(q)) return false;
            }
            return true;
          });

          // Raggruppamento a due livelli, ordine dei livelli in base a riepGroupMode:
          //  'type'   → livello1 = TYPE,   livello2 = Gruppo
          //  'gruppo' → livello1 = Gruppo, livello2 = TYPE
          const groups = {};
          codici.forEach(c => {
            const inf = pnInfo[c] || { models: new Set(), gruppi: new Set(), pnits: new Set(), type: '', ref: '', rplus: '', descrizione: '' };
            const stock = stockByCodice[c];
            const type = inf.type || '—';
            const pnitsList = inf.pnits.size > 0 ? [...inf.pnits] : ['—'];
            const pnitKey = pnitsList.join(', ');
            // In vista 'gruppo' (PNIT → TYPE) il codice viene duplicato su ogni singolo PNIT.
            // In vista 'type' (TYPE → PNIT) resta un'unica riga con i PNIT concatenati.
            const branches = riepGroupMode === 'gruppo'
              ? pnitsList.map(g => ({ lvl1: g, lvl2: type, gruppoLabel: g }))
              : [{ lvl1: type, lvl2: pnitKey, gruppoLabel: pnitKey }];
            const inArrivo = arrivoByCodice[c] || 0;
            const inOrdine = ordini[c]?.in_ordine || 0;
            branches.forEach(({ lvl1, lvl2, gruppoLabel }) => {
              if (!groups[lvl1]) groups[lvl1] = { key: lvl1, type: lvl1, totale: 0, inArrivo: 0, inOrdine: 0, nCodici: 0, models: {} };
              groups[lvl1].totale += stock;
              groups[lvl1].inArrivo += inArrivo;
              groups[lvl1].inOrdine += inOrdine;
              groups[lvl1].nCodici += 1;
              if (!groups[lvl1].models[lvl2]) groups[lvl1].models[lvl2] = { key: `${lvl1}||${lvl2}`, modello: lvl2, totale: 0, inArrivo: 0, inOrdine: 0, codici: [] };
              const mg = groups[lvl1].models[lvl2];
              mg.totale += stock;
              mg.inArrivo += inArrivo;
              mg.inOrdine += inOrdine;
              mg.codici.push({ codice: c, descrizione: inf.descrizione, ref: inf.ref, rplus: inf.rplus, eol: inf.eol, modello: gruppoLabel, stock, inArrivo, inOrdine });
            });
          });
          const groupList = Object.values(groups)
            .map(g => ({ ...g, modelList: Object.values(g.models).sort((a, b) => a.modello.localeCompare(b.modello)) }))
            .sort((a, b) => a.type.localeCompare(b.type));
          // Stock totale reale (unico per codice) — non conta le duplicazioni tra gruppi
          const totaleGenerale = codici.reduce((s, c) => s + (stockByCodice[c] || 0), 0);
          const totArrivo = codici.reduce((s, c) => s + (arrivoByCodice[c] || 0), 0);
          const totOrdine = codici.reduce((s, c) => s + (ordini[c]?.in_ordine || 0), 0);

          const toggleGroup = (key) => setRiepExpanded(prev => { const n = new Set(prev); if (n.has(key)) n.delete(key); else n.add(key); return n; });

          return (
            <div className="space-y-5">
              {/* Controlli */}
              <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
                <div className="flex flex-wrap items-center gap-3">
                  <select value={riepFilterModello} onChange={e => { setRiepFilterModello(e.target.value); setRiepExpanded(new Set()); }}
                    className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                    <option value="">Tutti i modelli ({modelliList.length})</option>
                    {modelliList.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                  <select value={riepFilterPnit} onChange={e => { setRiepFilterPnit(e.target.value); setRiepExpanded(new Set()); }}
                    className="bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden">
                    <option value="">Tutti i PNIT ({pnitList.length})</option>
                    {pnitList.map(p => <option key={p} value={p}>{p}</option>)}
                  </select>
                  <input value={riepSearch} onChange={e => setRiepSearch(e.target.value)}
                    placeholder="Cerca per type, codice, descrizione..."
                    className="flex-grow min-w-[200px] bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={riepFilterRef} onChange={e => setRiepFilterRef(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                    <span className="text-xs font-semibold text-gray-700">Solo REF</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={riepFilterRplus} onChange={e => setRiepFilterRplus(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                    <span className="text-xs font-semibold text-gray-700">Solo R+</span>
                  </label>
                </div>
                <div className="flex items-center justify-between text-[11px] text-gray-500 flex-wrap gap-2">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-gray-600">Vista:</span>
                    <div className="flex gap-1 bg-gray-100 p-1 rounded-lg">
                      {[
                        { id: 'type', label: 'TYPE → PNIT' },
                        { id: 'gruppo', label: 'PNIT → TYPE' },
                      ].map(v => (
                        <button key={v.id} onClick={() => { setRiepGroupMode(v.id); setRiepExpanded(new Set()); }}
                          className={`text-[11px] font-bold px-3 py-1 rounded-md cursor-pointer transition ${riepGroupMode === v.id ? 'bg-white text-gray-800 shadow-xs' : 'text-gray-500 hover:text-gray-700'}`}>
                          {v.label}
                        </button>
                      ))}
                    </div>
                    <span>· {groupList.length} {riepGroupMode === 'gruppo' ? 'PNIT' : 'type'} · {codici.length} codici</span>
                  </div>
                  <span className="font-black text-blue-600">Stock: {totaleGenerale} · <span className="text-emerald-600">In arrivo: {totArrivo}</span> · <span className="text-amber-600">In ordine: {totOrdine}</span></span>
                </div>
              </div>

              {/* Tabella gruppi */}
              {groupList.length > 0 ? (
                <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                      <tr>
                        <th className="px-3 py-3 w-8"></th>
                        <th className="px-3 py-3">{riepGroupMode === 'gruppo' ? 'PNIT' : 'Type'}</th>
                        <th className="px-3 py-3 text-right">N. Codici</th>
                        <th className="px-3 py-3 text-right">Stock Totale</th>
                        <th className="px-3 py-3 text-right text-emerald-600">In arrivo</th>
                        <th className="px-3 py-3 text-right text-amber-600">In ordine</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {groupList.map(g => (
                        <Fragment key={g.key}>
                          <tr className="hover:bg-blue-50/50 transition cursor-pointer font-bold" onClick={() => toggleGroup(g.key)}>
                            <td className="px-3 py-2.5 text-gray-400">{riepExpanded.has(g.key) ? '▾' : '▸'}</td>
                            <td className="px-3 py-2.5 text-gray-700">{g.type}</td>
                            <td className="px-3 py-2.5 text-right font-mono">{g.nCodici}</td>
                            <td className="px-3 py-2.5 text-right font-mono font-black text-blue-700">{g.totale}</td>
                            <td className="px-3 py-2.5 text-right font-mono font-black text-emerald-600">{g.inArrivo || ''}</td>
                            <td className="px-3 py-2.5 text-right font-mono font-black text-amber-600">{g.inOrdine || ''}</td>
                          </tr>
                          {riepExpanded.has(g.key) && g.modelList.map(m => (
                            <Fragment key={m.key}>
                              <tr className="bg-indigo-50/40 hover:bg-indigo-50 transition cursor-pointer font-bold border-t border-indigo-100" onClick={() => toggleGroup(m.key)}>
                                <td className="px-3 py-2 text-indigo-400 pl-8">{riepExpanded.has(m.key) ? '▾' : '▸'}</td>
                                <td className="px-3 py-2 font-mono font-black text-indigo-800">{m.modello}</td>
                                <td className="px-3 py-2 text-right font-mono text-gray-500">{m.codici.length}</td>
                                <td className="px-3 py-2 text-right font-mono font-black text-indigo-700">{m.totale}</td>
                                <td className="px-3 py-2 text-right font-mono font-black text-emerald-600">{m.inArrivo || ''}</td>
                                <td className="px-3 py-2 text-right font-mono font-black text-amber-600">{m.inOrdine || ''}</td>
                              </tr>
                              {riepExpanded.has(m.key) && (
                                <tr>
                                  <td colSpan={6} className="p-0">
                                    <table className="w-full text-[11px] bg-gray-50/60">
                                      <thead className="text-[9px] font-black text-gray-400 uppercase">
                                        <tr>
                                          <th className="px-3 py-1.5 text-left pl-14">PNIT</th>
                                          <th className="px-3 py-1.5 text-left">Codice</th>
                                          <th className="px-3 py-1.5 text-left">Descrizione</th>
                                          <th className="px-3 py-1.5 text-center">ST.</th>
                                          <th className="px-3 py-1.5 text-center">REF</th>
                                          <th className="px-3 py-1.5 text-center">R+</th>
                                          <th className="px-3 py-1.5 text-right">Stock</th>
                                          <th className="px-3 py-1.5 text-right text-emerald-600">In arrivo</th>
                                          <th className="px-3 py-1.5 text-right text-amber-600">In ordine</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {m.codici.sort((a, b) => b.stock - a.stock).map(c => (
                                          <tr key={c.codice} className="border-t border-gray-100">
                                            <td className="px-3 py-1.5 pl-14 font-mono font-black text-indigo-700">{c.modello}</td>
                                            <td className="px-3 py-1.5 font-mono font-bold text-blue-700">{c.codice}</td>
                                            <td className="px-3 py-1.5 text-gray-600">{c.descrizione || '—'}</td>
                                            <td className="px-3 py-1.5 text-center">
                                              {c.eol ? <span className={`px-1 py-px rounded font-black text-[9px] border ${c.eol === 'EOL' ? 'bg-red-50 text-red-600 border-red-100' : c.eol === 'ALT' ? 'bg-amber-50 text-amber-700 border-amber-100' : 'bg-blue-50 text-blue-700 border-blue-100'}`}>{c.eol}</span> : ''}
                                            </td>
                                            <td className="px-3 py-1.5 text-center">{c.ref === 'X' ? '✓' : ''}</td>
                                            <td className="px-3 py-1.5 text-center">{c.rplus === 'X' ? '✓' : ''}</td>
                                            <td className="px-3 py-1.5 text-right font-mono font-black">{c.stock}</td>
                                            <td className="px-3 py-1.5 text-right font-mono text-emerald-600">{c.inArrivo || ''}</td>
                                            <td className="px-3 py-1.5 text-right font-mono text-amber-600">{c.inOrdine || ''}</td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              )}
                            </Fragment>
                          ))}
                        </Fragment>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="text-center py-16 text-gray-400 text-sm">Nessuno stock da visualizzare.</div>
              )}
            </div>
          );
        })()}

        {/* ==================== MODULO MATRICE PNIT × TYPE ==================== */}
        {activeModule === 'matrice' && (() => {
          // Stock / in arrivo / in ordine per codice
          const stockByCodice = {};
          stockItems.forEach(s => { if (s.stock > 0) stockByCodice[s.codice] = (stockByCodice[s.codice] || 0) + s.stock; });
          const arrivoByCodice = {};
          poLines.forEach(l => {
            const cod = String(l.item_code || '').trim();
            if (!cod || cod === 'N/D') return;
            arrivoByCodice[cod] = (arrivoByCodice[cod] || 0) + (l.qty_expected || 0);
          });

          // Prezzo unitario e stato locked per pn (dal DB spare parts)
          const priceByPn = {};
          const lockedByPn = {};
          spareParts.forEach(p => {
            const pn = (p.pn || '').trim();
            if (!pn) return;
            if (!priceByPn[pn]) priceByPn[pn] = p.price || 0;
            // un pn è ordinabile se almeno una sua riga non è locked
            if ((p.locked || '').trim().toUpperCase() !== 'Y') lockedByPn[pn] = false;
            else if (lockedByPn[pn] === undefined) lockedByPn[pn] = true;
          });

          // Tutte le combinazioni (PNIT, TYPE) presenti nel DB spare parts.
          // Ogni combinazione raccoglie i propri codici (pn) su cui sommare le quantità.
          const comboMap = {};
          spareParts.forEach(p => {
            const pnit = (p.pnit || '').trim() || '—';
            const type = (p.type || '').trim() || '—';
            const pn = (p.pn || '').trim();
            if (!pn) return;
            const key = `${pnit}||${type}`;
            if (!comboMap[key]) comboMap[key] = { pnit, type, codici: new Set(), orderable: new Set() };
            comboMap[key].codici.add(pn);
            // ordinabile = non locked (esclude EOL/CLI, TO ORDER=NO, TYPE=NO)
            if ((p.locked || '').trim().toUpperCase() !== 'Y') comboMap[key].orderable.add(pn);
          });

          const mesi = parseFloat(mesiCopertura) || 0;
          let combos = Object.values(comboMap).map(c => {
            const codici = [...c.codici];
            // Codice da ordinare: il più idoneo (ordinabile); se più d'uno equivalente, il primo per codice
            const orderable = [...c.orderable].sort();
            const codiceOrdine = (orderable.length > 0 ? orderable : codici.slice().sort())[0] || '';
            const nOrderable = orderable.length;
            const stock = codici.reduce((s, k) => s + (stockByCodice[k] || 0), 0);
            const inArrivo = codici.reduce((s, k) => s + (arrivoByCodice[k] || 0), 0);
            const inOrdine = codici.reduce((s, k) => s + (ordini[k]?.in_ordine || 0), 0);
            const cons = consumi[`${c.pnit}${c.type}`] || { media: 0, nomaterial: 0 };
            const media = cons.media || 0;
            const nomaterial = cons.nomaterial || 0;
            const stima = Math.round(media * mesi + nomaterial);
            // Se il codice da ordinare è locked, la quantità da ordinare è 0
            const locked = !!lockedByPn[codiceOrdine];
            const fabbisogno = stima - stock - inOrdine;
            // Arrotonda per eccesso a multipli di 100
            const daOrdinare = (locked || fabbisogno <= 0) ? 0 : Math.ceil(fabbisogno / 100) * 100;
            const prezzo = priceByPn[codiceOrdine] || 0;
            const ptot = daOrdinare > 0 ? prezzo * daOrdinare : 0;
            return { pnit: c.pnit, type: c.type, nCodici: codici.length, codiceOrdine, nOrderable, locked, media, nomaterial, stock, inArrivo, inOrdine, stima, daOrdinare, prezzo, ptot };
          });

          const q = matriceSearch.trim().toLowerCase();
          if (q) combos = combos.filter(c => `${c.pnit} ${c.type}`.toLowerCase().includes(q));
          if (matriceSoloStima) combos = combos.filter(c => c.stima > 0);
          if (matriceSoloDaOrdinare) combos = combos.filter(c => c.daOrdinare > 0);

          const { col, dir } = matriceSort;
          const mul = dir === 'asc' ? 1 : -1;
          combos.sort((a, b) => {
            const va = a[col], vb = b[col];
            if (typeof va === 'number') return (va - vb) * mul;
            return String(va).localeCompare(String(vb)) * mul;
          });

          const tot = combos.reduce((acc, c) => { acc.stock += c.stock; acc.inArrivo += c.inArrivo; acc.inOrdine += c.inOrdine; acc.ptot += c.ptot; return acc; }, { stock: 0, inArrivo: 0, inOrdine: 0, ptot: 0 });
          const fmtEuro = (n) => (n || 0).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
          const toggleSort = (c) => setMatriceSort(prev => prev.col === c ? { col: c, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { col: c, dir: 'asc' });
          const arrow = (c) => matriceSort.col === c ? (matriceSort.dir === 'asc' ? ' ▲' : ' ▼') : '';

          return (
            <div className="space-y-5">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <h2 className="text-lg font-black text-gray-800">🧮 Matrice PNIT × TYPE</h2>
                  <p className="text-xs text-gray-500">Tutte le combinazioni PNIT + TYPE del DB spare parts, con stock, in arrivo (Piano Arrivi) e in ordine.</p>
                </div>
                <div className="flex items-center gap-3 flex-wrap">
                  <input value={matriceSearch} onChange={e => setMatriceSearch(e.target.value)}
                    placeholder="Cerca per PNIT o TYPE..."
                    className="min-w-[220px] bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={matriceSoloStima} onChange={e => setMatriceSoloStima(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                    <span className="text-xs font-semibold text-gray-700">Solo con stima necessaria</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={matriceSoloDaOrdinare} onChange={e => setMatriceSoloDaOrdinare(e.target.checked)} className="w-4 h-4 accent-blue-600 cursor-pointer" />
                    <span className="text-xs font-semibold text-gray-700">Solo da ordinare</span>
                  </label>
                </div>
              </div>

              {/* Upload file quantità in ordine */}
              <div className="bg-white p-3 rounded-2xl border border-gray-200 shadow-xs flex items-center justify-between flex-wrap gap-2">
                <label className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 text-xs font-bold px-3 py-2 rounded-xl cursor-pointer transition relative">
                  📥 Aggiorna quantità in ordine (CSV)
                  <input type="file" accept=".csv,.xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handleOrdiniUpload} />
                </label>
                <label className="bg-purple-50 hover:bg-purple-100 text-purple-700 border border-purple-200 text-xs font-bold px-3 py-2 rounded-xl cursor-pointer transition relative">
                  📥 Aggiorna consumi / NoMaterial
                  <input type="file" accept=".csv,.xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handleConsumiUpload} />
                </label>
                <span className="text-[11px] text-gray-400">In arrivo dal Piano Arrivi</span>
                {(ordiniLoading || consumiLoading) && <span className="text-[11px] font-bold text-amber-600 animate-pulse">Caricamento...</span>}
                {importMeta.ordini && <span className="text-[11px] text-gray-400">In ordine: {new Date(importMeta.ordini).toLocaleString('it-IT')}</span>}
                {importMeta.consumi && <span className="text-[11px] text-gray-400">Consumi: {new Date(importMeta.consumi).toLocaleString('it-IT')}</span>}
              </div>

              <div className="flex items-center justify-between text-[11px] text-gray-500 flex-wrap gap-2">
                <div className="flex items-center gap-2">
                  <span className="font-semibold text-gray-600">Mesi copertura:</span>
                  <input type="number" min="0" step="1" value={mesiCopertura}
                    onChange={e => setMesiCopertura(e.target.value)}
                    className="w-20 bg-gray-50 border border-gray-300 rounded-lg p-1.5 text-xs text-right focus:outline-hidden" />
                  <span className="text-gray-400">· Stima = consumo medio × mesi + NoMaterial</span>
                  <span>· {combos.length} combinazioni</span>
                </div>
                <span className="font-black text-blue-600">Stock: {tot.stock} · <span className="text-emerald-600">In arrivo: {tot.inArrivo}</span> · <span className="text-amber-600">In ordine: {tot.inOrdine}</span> · <span className="text-gray-800">Valore d&apos;acquisto: {fmtEuro(tot.ptot)} $</span></span>
              </div>

              {combos.length > 0 ? (
                <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
                  <table className="w-full min-w-[1300px] text-left border-collapse text-xs">
                    <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                      <tr>
                        <th className="px-3 py-3 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('pnit')}>PNIT{arrow('pnit')}</th>
                        <th className="px-3 py-3 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('type')}>TYPE{arrow('type')}</th>
                        <th className="px-3 py-3 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('codiceOrdine')}>Codice da ordinare{arrow('codiceOrdine')}</th>
                        <th className="px-3 py-3 text-right cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('nCodici')}>N. Codici{arrow('nCodici')}</th>
                        <th className="px-3 py-3 text-right text-purple-600 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('media')}>Cons. medio/mese{arrow('media')}</th>
                        <th className="px-3 py-3 text-right text-rose-600 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('nomaterial')}>NoMaterial{arrow('nomaterial')}</th>
                        <th className="px-3 py-3 text-right text-blue-700 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('stima')}>Stima necessaria{arrow('stima')}</th>
                        <th className="px-3 py-3 text-right cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('stock')}>Stock{arrow('stock')}</th>
                        <th className="px-3 py-3 text-right text-emerald-600 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('inArrivo')}>In arrivo{arrow('inArrivo')}</th>
                        <th className="px-3 py-3 text-right text-amber-600 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('inOrdine')}>In ordine{arrow('inOrdine')}</th>
                        <th className="px-3 py-3 text-right text-red-700 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('daOrdinare')}>Da ordinare{arrow('daOrdinare')}</th>
                        <th className="px-3 py-3 text-center cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('locked')}>Locked{arrow('locked')}</th>
                        <th className="px-3 py-3 text-right cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('prezzo')}>Prezzo unit.{arrow('prezzo')}</th>
                        <th className="px-3 py-3 text-right text-gray-800 cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition" onClick={() => toggleSort('ptot')}>P.tot{arrow('ptot')}</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {combos.map(c => (
                        <tr key={`${c.pnit}||${c.type}`} className="hover:bg-blue-50/50 transition">
                          <td className="px-3 py-2.5 font-mono font-bold text-indigo-800">{c.pnit}</td>
                          <td className="px-3 py-2.5 text-gray-700">{c.type}</td>
                          <td className="px-3 py-2.5 font-mono font-bold text-blue-700">{c.codiceOrdine || '—'}{c.nOrderable > 1 && <span className="ml-1 text-[9px] font-black text-amber-600" title={`${c.nOrderable} codici equivalenti`}>⚑{c.nOrderable}</span>}</td>
                          <td className="px-3 py-2.5 text-right font-mono text-gray-500">{c.nCodici}</td>
                          <td className="px-3 py-2.5 text-right font-mono text-purple-600">{c.media ? c.media.toFixed(1) : ''}</td>
                          <td className="px-3 py-2.5 text-right font-mono text-rose-600">{c.nomaterial || ''}</td>
                          <td className="px-3 py-2.5 text-right font-mono font-black text-blue-800">{c.stima || ''}</td>
                          <td className="px-3 py-2.5 text-right font-mono font-black text-blue-700">{c.stock}</td>
                          <td className="px-3 py-2.5 text-right font-mono font-black text-emerald-600">{c.inArrivo || ''}</td>
                          <td className="px-3 py-2.5 text-right font-mono font-black text-amber-600">{c.inOrdine || ''}</td>
                          <td className={`px-3 py-2.5 text-right font-mono font-black ${c.daOrdinare > 0 ? 'text-red-700' : 'text-gray-300'}`}>{c.daOrdinare > 0 ? c.daOrdinare : ''}</td>
                          <td className="px-3 py-2.5 text-center">{c.locked ? <span className="text-[9px] font-black text-red-600 bg-red-50 border border-red-100 px-1.5 py-0.5 rounded">🔒 LOCKED</span> : ''}</td>
                          <td className="px-3 py-2.5 text-right font-mono text-gray-600">{c.prezzo ? fmtEuro(c.prezzo) : ''}</td>
                          <td className="px-3 py-2.5 text-right font-mono font-black text-gray-800">{c.ptot ? fmtEuro(c.ptot) : ''}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="text-center py-16 text-gray-400 text-sm">Nessuna combinazione da visualizzare.</div>
              )}
            </div>
          );
        })()}

        {/* ==================== MODULO ANAGRAFICA TERMINALI ==================== */}
        {activeModule === 'anagrafica' && (() => {
          const gruppoDesc = (a) => [a.modello, a.versione_pci, a.note].map(x => String(x || '').trim()).filter(Boolean).join(' - ');
          const q = anagSearch.trim().toLowerCase();
          const list = anagrafica.filter(a => {
            if (!q) return true;
            return `${a.codice} ${a.tipo} ${a.modello} ${a.famiglia} ${a.descrizione} ${gruppoDesc(a)}`.toLowerCase().includes(q);
          });
          return (
          <div className="space-y-5">
            <div className="flex items-center justify-between flex-wrap gap-3">
              <div>
                <h2 className="text-lg font-black text-gray-800">📇 Anagrafica Terminali</h2>
                <p className="text-xs text-gray-500">Chiave <strong>CODICE = PNIT</strong> del DB spare parts. Import da <em>tipoterminale.xlsx</em> (foglio &quot;famiglie&quot;).</p>
                {importMeta.anagrafica && <p className="text-[11px] text-gray-400">Ultimo aggiornamento: {new Date(importMeta.anagrafica).toLocaleString('it-IT')}</p>}
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={fetchAnagrafica} className="bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold px-3 py-2.5 rounded-xl cursor-pointer transition" title="Aggiorna">↻</button>
                <label className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer transition shadow-xs relative">
                  📥 Importa file
                  <input type="file" accept=".csv,.xls,.xlsx" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onChange={handleAnagraficaUpload} />
                </label>
              </div>
            </div>

            <div className="flex items-center gap-3 flex-wrap">
              <input value={anagSearch} onChange={e => setAnagSearch(e.target.value)}
                placeholder="Cerca per codice, modello, famiglia, descrizione..."
                className="flex-grow min-w-[200px] bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs focus:outline-hidden" />
              <span className="text-[11px] text-gray-500">{list.length} / {anagrafica.length} terminali</span>
            </div>

            {anagLoading && <div className="text-center py-4 text-xs font-bold text-amber-600 animate-pulse">Caricamento...</div>}

            {!anagLoading && anagrafica.length === 0 && (
              <div className="text-center py-16 text-gray-400 text-sm">Nessun terminale in anagrafica. Importa il file tipoterminale.xlsx.</div>
            )}

            {!anagLoading && list.length > 0 && (
              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
                <table className="w-full min-w-[720px] text-left border-collapse text-xs">
                  <thead className="bg-gray-50 border-b border-gray-200 text-[10px] font-black text-gray-500 uppercase tracking-wider">
                    <tr>
                      <th className="px-3 py-3">Codice (PNIT)</th>
                      <th className="px-3 py-3">Gruppo</th>
                      <th className="px-3 py-3">Tipo</th>
                      <th className="px-3 py-3">Famiglia</th>
                      <th className="px-3 py-3">Descrizione</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {list.map(a => (
                      <tr key={a.codice} className="hover:bg-blue-50/50 transition">
                        <td className="px-3 py-2.5 font-mono font-bold text-blue-700">{a.codice}</td>
                        <td className="px-3 py-2.5 font-black text-indigo-800">{gruppoDesc(a) || '—'}</td>
                        <td className="px-3 py-2.5 text-gray-600">{a.tipo || '—'}</td>
                        <td className="px-3 py-2.5 text-gray-600">{a.famiglia || '—'}</td>
                        <td className="px-3 py-2.5 text-gray-600">{a.descrizione || '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
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
                <p className="text-xs text-gray-500">{prelieviList.length} prelievi totali</p>
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={fetchPrelievi}
                  className="bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold px-3 py-2.5 rounded-xl cursor-pointer transition" title="Aggiorna">
                  ↻
                </button>
                {prelieviList.some(p => p.stato !== 'registrato') && (
                  <button onClick={exportPrelieviRettifica}
                    className="bg-green-50 hover:bg-green-100 text-green-700 border border-green-200 text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer transition">
                    📥 Esporta rettifica
                  </button>
                )}
                <button onClick={() => { setPrelievoView('new'); setPrelievoRighe([]); setPrelievoFeedback({ text: '', type: '' }); setTimeout(() => prelievoScannerRef.current?.focus(), 100); }}
                  className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer transition shadow-xs">
                  + Nuovo prelievo
                </button>
              </div>
            </div>

            {/* Sotto-schede: Attivi / Registrati */}
            <div className="flex gap-1 bg-gray-100 p-1 rounded-xl w-fit">
              {[
                { id: 'attivi', label: 'Attivi', n: prelieviList.filter(p => p.stato !== 'registrato').length },
                { id: 'registrati', label: 'Registrati', n: prelieviList.filter(p => p.stato === 'registrato').length },
              ].map(t => (
                <button key={t.id} onClick={() => setPrelievoTab(t.id)}
                  className={`text-xs font-bold px-4 py-2 rounded-lg cursor-pointer transition ${prelievoTab === t.id ? 'bg-white text-gray-800 shadow-xs' : 'text-gray-500 hover:text-gray-700'}`}>
                  {t.label} <span className="ml-1 text-[10px] opacity-70">({t.n})</span>
                </button>
              ))}
            </div>

            {prelieviLoading && <div className="text-center py-4 text-xs font-bold text-amber-600 animate-pulse">Caricamento...</div>}

            {(() => {
              const filtered = prelieviList.filter(p => prelievoTab === 'registrati' ? p.stato === 'registrato' : p.stato !== 'registrato');
              if (prelieviLoading) return null;
              if (filtered.length === 0) return (
                <div className="text-center py-16 text-gray-400 text-sm">
                  {prelievoTab === 'registrati'
                    ? 'Nessun prelievo registrato. Esporta la rettifica dagli attivi per registrarli.'
                    : 'Nessun prelievo attivo. Clicca "Nuovo prelievo" per iniziare.'}
                </div>
              );
              return (
              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-x-auto">
                <table className="w-full min-w-[680px] text-left border-collapse text-xs">
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
                    {filtered.map(p => (
                      <tr key={p.id} className="hover:bg-blue-50/50 transition cursor-pointer" onClick={() => openPrelievoDetail(p)}>
                        <td className="px-3 py-2.5 font-mono font-bold text-blue-700 underline">
                          {p.id_prelievo}
                          {p.stato === 'registrato' && <span className="ml-2 text-[9px] font-black text-emerald-700 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded-md uppercase tracking-wide">Registrato</span>}
                        </td>
                        <td className="px-3 py-2.5 text-gray-600">{p.data_prelievo ? new Date(p.data_prelievo).toLocaleString('it-IT') : '—'}</td>
                        <td className="px-3 py-2.5 text-gray-700">{p.utente || '—'}</td>
                        <td className="px-3 py-2.5 text-gray-600">{p.destinazione || '—'}</td>
                        <td className="px-3 py-2.5 text-right font-mono">{p.n_righe}</td>
                        <td className="px-3 py-2.5 text-right font-mono font-black">{p.n_pezzi}</td>
                        <td className="px-3 py-2.5 text-center">
                          {p.stato === 'registrato'
                            ? <span className="text-[10px] text-gray-300" title="Prelievo registrato: non eliminabile">🔒</span>
                            : <button onClick={(e) => { e.stopPropagation(); deletePrelievo(p); }}
                                className="text-[10px] text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-300 px-2 py-1 rounded-lg cursor-pointer transition" title="Elimina e ripristina inventario">🗑</button>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              );
            })()}
          </div>
        )}

        {/* ==================== MODULO PRELIEVI — DETTAGLIO ==================== */}
        {activeModule === 'prelievi' && prelievoView === 'detail' && prelievoDetail && (
          <div className="space-y-5 max-w-4xl mx-auto">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-lg font-black text-gray-800">📤 Prelievo {prelievoDetail.testata.id_prelievo}
                  {prelievoDetail.testata.stato === 'registrato' && <span className="ml-2 text-[10px] align-middle font-black text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md uppercase tracking-wide">🔒 Registrato</span>}
                </h2>
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

            <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
              <div className="flex items-center justify-between px-4 py-3 border-b border-gray-100">
                <span className="text-xs font-black text-gray-500 uppercase tracking-wider">{prelievoDetail.righe.length} righe</span>
                <span className="text-sm font-black text-blue-600">Tot. pezzi: {prelievoDetail.righe.reduce((s, r) => s + (r.quantita || 0), 0)}</span>
              </div>
              <div className="overflow-x-auto">
              <table className="w-full min-w-[560px] text-left border-collapse text-xs">
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
                  <div className={`p-3 rounded-xl text-center text-sm font-bold border ${prelievoFeedback.type === 'success' ? 'bg-green-50 text-green-800 border-green-200' : prelievoFeedback.type === 'warning' ? 'bg-amber-50 text-amber-800 border-amber-200' : 'bg-red-50 text-red-800 border-red-200'}`}>
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
                  <div className="overflow-x-auto">
                  <table className="w-full min-w-[620px] text-xs">
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
                  </div>
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
              {importMeta.po_lines && <p className="text-[11px] text-gray-400 mt-2 text-center sm:text-left">Ultimo aggiornamento piano arrivi: {new Date(importMeta.po_lines).toLocaleString('it-IT')}</p>}
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
                              // Carica le righe pending, aggregate per codice + magazzino + bancale
                              const { data: existing } = await supabase.from('carton_arrivals')
                                .select('*').eq('invoice', group.invoice).eq('stato', 'pending')
                                .order('id', { ascending: true });
                              if (existing && existing.length > 0) {
                                setArrivoQtyBancale(existing[existing.length - 1].bancale || '');
                                setArrivoQtyMagazzino(existing[existing.length - 1].magazzino || 'GESSATE');
                                const agg = new Map();
                                existing.forEach(c => {
                                  const key = `${c.codice}__${c.magazzino}__${c.bancale}`;
                                  if (agg.has(key)) { agg.get(key).quantita += (c.quantita || 0); agg.get(key).rilievi += (c.rilievi || 1); }
                                  else agg.set(key, { codice: c.codice, quantita: c.quantita || 0, bancale: c.bancale, magazzino: c.magazzino, rilievi: c.rilievi || 1 });
                                });
                                setArrivoQtyCartoni([...agg.values()]);
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
                                  {item.part_number && item.part_number !== 'N/D' && item.part_number !== item.item_code && (
                                    <span>VPN: <span className="text-indigo-600 font-mono font-bold">{item.part_number}</span></span>
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
                                    {item.scanned_count > 0 ? 'Modifica' : 'Avvia'}
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

            {/* Form nuovo bancale */}
            {arrivoBancaleForm.open && (
              <div className="bg-white p-4 rounded-2xl border-2 border-blue-300 shadow-xs space-y-3">
                <span className="text-[10px] font-black text-blue-700 uppercase tracking-wider block">Nuovo bancale</span>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Nome bancale</label>
                    <input value={arrivoBancaleForm.nome} onChange={e => setArrivoBancaleForm(f => ({ ...f, nome: e.target.value }))}
                      placeholder="Es. BANCALE-001" autoFocus
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-sm focus:outline-hidden font-mono" />
                  </div>
                  <div className="space-y-1">
                    <label className="block text-xs font-bold text-gray-500">Magazzino</label>
                    <select value={arrivoBancaleForm.magazzino} onChange={e => setArrivoBancaleForm(f => ({ ...f, magazzino: e.target.value }))}
                      className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-sm focus:outline-hidden">
                      <option value="GESSATE">GESSATE</option>
                      <option value="ESPRINET">ESPRINET</option>
                    </select>
                  </div>
                </div>
                <div className="flex gap-2">
                  <button onClick={() => {
                      if (!arrivoBancaleForm.nome.trim()) { alert('Inserisci il nome del bancale.'); return; }
                      setArrivoQtyBancale(arrivoBancaleForm.nome.trim());
                      setArrivoQtyMagazzino(arrivoBancaleForm.magazzino);
                      setArrivoBancaleForm({ open: false, nome: '', magazzino: arrivoBancaleForm.magazzino });
                    }}
                    className="flex-1 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold py-2.5 rounded-xl cursor-pointer transition">Conferma bancale</button>
                  <button onClick={() => setArrivoBancaleForm(f => ({ ...f, open: false }))}
                    className="bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold px-4 py-2.5 rounded-xl cursor-pointer transition">Annulla</button>
                </div>
              </div>
            )}

            {/* Scanner QR */}
            <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-xs space-y-3">
              {/* Bancale in lavorazione */}
              {arrivoQtyBancale.trim() ? (
                <div className="flex items-center justify-between bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 gap-3">
                  <div className="min-w-0">
                    <span className="text-[9px] font-black text-blue-400 uppercase tracking-wider block">Bancale in lavorazione</span>
                    <span className="text-lg sm:text-xl font-black text-blue-800 font-mono truncate block">📦 {arrivoQtyBancale}</span>
                    <span className="text-[10px] font-bold text-blue-400 uppercase">{arrivoQtyMagazzino}</span>
                  </div>
                  <button onClick={() => setArrivoBancaleForm({ open: true, nome: '', magazzino: arrivoQtyMagazzino })}
                    className="shrink-0 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold px-3 py-2 rounded-xl cursor-pointer transition shadow-xs whitespace-nowrap">+ Nuovo bancale</button>
                </div>
              ) : (
                <button onClick={() => setArrivoBancaleForm({ open: true, nome: '', magazzino: arrivoQtyMagazzino || 'GESSATE' })}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold py-3 rounded-xl cursor-pointer transition shadow-xs">
                  + Nuovo bancale
                </button>
              )}
              <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Scansione QR (codice + quantità)</span>
              <form onSubmit={handleArrivoQtySubmit}>
                <input type="text" ref={arrivoQtyScannerRef} value={arrivoQtyScanner}
                  onChange={e => setArrivoQtyScanner(e.target.value)}
                  autoComplete="off" autoCorrect="off" autoCapitalize="off" spellCheck="false"
                  placeholder="Spara il QR (stesso codice incrementa la quantità)..."
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
              <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider block">Inserimento manuale (codice + quantità)</span>
              <div className="flex gap-3 items-end">
                <div className="flex-grow space-y-1 relative">
                  <label className="block text-xs font-bold text-gray-500">Codice</label>
                  <input value={arrivoQtyManuale.codice}
                    onChange={e => { setArrivoQtyManuale(v => ({...v, codice: e.target.value})); setArrivoCodiceOpen(true); }}
                    onFocus={() => setArrivoCodiceOpen(true)}
                    onBlur={() => setTimeout(() => setArrivoCodiceOpen(false), 200)}
                    autoComplete="off"
                    placeholder="Cerca tra i codici attesi..."
                    className="w-full bg-gray-50 border border-gray-300 rounded-xl p-2.5 text-xs font-mono focus:outline-hidden" />
                  {arrivoCodiceOpen && (() => {
                    const q = arrivoQtyManuale.codice.trim().toLowerCase();
                    const opts = poLines
                      .filter(l => l.china_invoice === arrivoQtyInvoice && l.sn_required === false)
                      .filter(l => !q || (l.item_code || '').toLowerCase().includes(q) || (l.description || '').toLowerCase().includes(q))
                      .slice(0, 30);
                    if (opts.length === 0) return null;
                    return (
                      <div className="absolute z-30 left-0 right-0 mt-1 bg-white border border-gray-200 rounded-xl shadow-lg max-h-56 overflow-y-auto">
                        {opts.map(l => (
                          <button key={l.unique_key} type="button"
                            onMouseDown={(ev) => { ev.preventDefault(); setArrivoQtyManuale(v => ({...v, codice: l.item_code})); setArrivoCodiceOpen(false); }}
                            className="w-full text-left px-3 py-2 hover:bg-blue-50 border-b border-gray-100 last:border-none">
                            <div className="font-mono font-bold text-xs text-blue-700">{l.item_code}</div>
                            <div className="text-[10px] text-gray-500 truncate">{l.description} — attesi {l.qty_expected}</div>
                          </button>
                        ))}
                      </div>
                    );
                  })()}
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
              {/* Riepilogo per codice — collassabile */}
              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
                <button onClick={() => setRiepCodiceOpen(o => !o)}
                  className="w-full flex items-center justify-between px-4 py-3 border-b border-gray-100 cursor-pointer hover:bg-gray-50">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-wider">Riepilogo per codice ({Object.keys(summary).length})</span>
                  <span className="text-gray-400 text-sm">{riepCodiceOpen ? '▾' : '▸'}</span>
                </button>
                {riepCodiceOpen && (
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
                )}
              </div>

              {/* Rilievi divisi per bancale */}
              {(() => {
                const perBancale = {};
                arrivoQtyCartoni.forEach(c => {
                  const key = `${c.magazzino}__${c.bancale}`;
                  if (!perBancale[key]) perBancale[key] = { magazzino: c.magazzino, bancale: c.bancale, righe: [], rilievi: 0, pezzi: 0 };
                  perBancale[key].righe.push(c);
                  perBancale[key].rilievi += (c.rilievi || 1);
                  perBancale[key].pezzi += (c.quantita || 0);
                });
                const bancali = Object.values(perBancale);
                return (
                  <div className="space-y-3">
                    {bancali.map(b => (
                      <div key={`${b.magazzino}__${b.bancale}`} className="bg-white rounded-2xl border border-gray-200 shadow-xs overflow-hidden">
                        <div className="flex items-center justify-between px-4 py-2.5 bg-blue-50 border-b border-blue-100 gap-2">
                          <span className="text-xs font-black text-blue-800 min-w-0 truncate">📦 {b.bancale} <span className="text-blue-400 font-bold">· {b.magazzino}</span></span>
                          <div className="flex items-center gap-2 shrink-0">
                            <span className="text-[11px] font-bold text-blue-600">{b.rilievi} rilievi · {b.pezzi} pz</span>
                            {(arrivoQtyBancale !== b.bancale || arrivoQtyMagazzino !== b.magazzino) && (
                              <button onClick={() => { setArrivoQtyBancale(b.bancale); setArrivoQtyMagazzino(b.magazzino); window.scrollTo({ top: 0, behavior: 'smooth' }); }}
                                className="text-[10px] font-bold text-white bg-blue-600 hover:bg-blue-700 px-2 py-1 rounded-lg cursor-pointer transition whitespace-nowrap">✎ Modifica</button>
                            )}
                            {arrivoQtyBancale === b.bancale && arrivoQtyMagazzino === b.magazzino && (
                              <span className="text-[9px] font-black text-green-700 bg-green-100 border border-green-200 px-2 py-1 rounded-lg">IN LAVORAZIONE</span>
                            )}
                          </div>
                        </div>
                        <table className="w-full text-xs">
                          <thead className="bg-gray-50 border-b border-gray-100 text-[10px] font-black text-gray-500 uppercase">
                            <tr>
                              <th className="px-3 py-2 text-left">Codice</th>
                              <th className="px-3 py-2 text-center">Rilievi</th>
                              <th className="px-3 py-2 text-right">Qtà</th>
                              <th className="px-3 py-2"></th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {b.righe.map(c => (
                              <tr key={`${c.codice}__${c.magazzino}__${c.bancale}`} className="hover:bg-gray-50/80">
                                <td className="px-3 py-2 font-mono font-bold text-blue-700">{c.codice}</td>
                                <td className="px-3 py-2 text-center text-gray-500">{c.rilievi || 1}</td>
                                <td className="px-3 py-2 text-right font-black">{c.quantita}</td>
                                <td className="px-3 py-2 text-right">
                                  <button onClick={() => removeCarton(c)} className="text-[10px] text-gray-400 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 hover:border-red-300 px-2 py-1 rounded-lg cursor-pointer transition">✕</button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ))}
                  </div>
                );
              })()}

              {/* Pulsante carico */}
              <div className="bg-white rounded-2xl border border-gray-200 shadow-xs p-4 space-y-2">
                {!invoiceComplete && (
                  <p className="text-xs text-center font-bold text-amber-600">
                    ⚠ Invoice incompleta — completa tutti i codici prima di caricare
                  </p>
                )}
                <button
                  onClick={caricaArrivoSuInventario}
                  disabled={!invoiceComplete}
                  className={`w-full font-black p-4 rounded-xl text-base shadow-md transition flex items-center justify-center gap-2 ${!invoiceComplete ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700 text-white cursor-pointer'}`}>
                  ✓ Carica su Inventario Spare Parts
                </button>
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
                      {scannedSerials.length}/{Object.keys(expectedSerials).length || activeLine.qty_expected}
                    </span>
                  </div>
                  <div className="w-full bg-gray-100 rounded-full h-3 overflow-hidden">
                    <div className="bg-blue-600 h-3 rounded-full transition-all duration-300"
                      style={{ width: `${(Object.keys(expectedSerials).length || activeLine.qty_expected) > 0 ? Math.min((scannedSerials.length / (Object.keys(expectedSerials).length || activeLine.qty_expected)) * 100, 100) : 0}%` }}></div>
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

              {scannedSerials.length > 0 && scannedSerials.length >= Object.keys(expectedSerials).length && (
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
              {scannedSerials.length > 50 && (
                <p className="text-[11px] text-gray-400 italic">Mostrate le ultime 50 di {scannedSerials.length} matricole rilevate.</p>
              )}
              <ul className="divide-y divide-gray-100 max-h-64 md:max-h-[500px] overflow-y-auto font-mono text-sm pr-1">
                {scannedSerials.slice(0, 50).map((s) => (
                  <li key={s.serial} className="py-2.5 flex items-center gap-3 border-b border-gray-100 last:border-none group">
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
        {currentView === 'review' && activeLine && (() => {
          // Rileva i doppioni: seriali presenti più di una volta
          const counts = {};
          scannedSerials.forEach(s => { counts[s.serial] = (counts[s.serial] || 0) + 1; });
          const extraCount = scannedSerials.length - Object.keys(counts).length; // righe in eccesso da rimuovere
          const hasDuplicates = extraCount > 0;
          // Ordina mettendo i doppioni in cima
          const ordered = [...scannedSerials].sort((a, b) => (counts[b.serial] > 1 ? 1 : 0) - (counts[a.serial] > 1 ? 1 : 0));

          return (
          <div className="max-w-3xl mx-auto bg-white p-6 sm:p-8 rounded-2xl shadow-md border border-gray-200 space-y-6 my-4">
            <div className="border-b border-gray-200 pb-4 text-center sm:text-left">
              <span className="text-[10px] bg-amber-100 text-amber-800 border border-amber-200 font-black px-3 py-1 rounded-md uppercase tracking-wider">Fase di Controllo Finale</span>
              <h2 className="text-xl sm:text-2xl font-black text-gray-900 mt-2">{activeLine.item_code} - Linea {activeLine.line_id}</h2>
              <p className="text-sm text-gray-400 font-bold uppercase mt-0.5">China Invoice: {activeLine.china_invoice} | Arrivo: {activeLine.arrival_date}</p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="bg-blue-50 border border-blue-100 p-4 rounded-xl text-center">
                <span className="block text-xs font-bold text-blue-500 uppercase tracking-wider">Matricole Registrate</span>
                <span className="text-3xl font-black text-blue-700 font-mono">{scannedSerials.length}</span>
              </div>
              <div className={`p-4 rounded-xl text-center border ${hasDuplicates ? 'bg-red-50 border-red-200' : 'bg-green-50 border-green-100'}`}>
                <span className={`block text-xs font-bold uppercase tracking-wider ${hasDuplicates ? 'text-red-600' : 'text-green-600'}`}>Doppioni</span>
                <span className={`text-3xl font-black font-mono ${hasDuplicates ? 'text-red-700' : 'text-green-700'}`}>{extraCount}</span>
              </div>
              <div className="bg-amber-50 border border-amber-100 p-4 rounded-xl text-center">
                <span className="block text-xs font-bold text-amber-600 uppercase tracking-wider">Cartoni (QR)</span>
                <span className="text-3xl font-black text-amber-700 font-mono">{cartonsScanned}</span>
              </div>
            </div>

            {hasDuplicates && (
              <div className="bg-red-50 border border-red-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <span className="text-sm font-bold text-red-700">
                  ⚠ Sono presenti {extraCount} matricole doppie. Rimuovile prima di registrare il carico.
                </span>
                <button onClick={removeDuplicateSerials}
                  className="bg-red-600 hover:bg-red-700 text-white text-xs font-bold px-4 py-2 rounded-xl cursor-pointer transition shadow-xs whitespace-nowrap">
                  🗑 Elimina {extraCount} doppioni
                </button>
              </div>
            )}

            <div className="space-y-2">
              <label className="block text-xs font-bold text-gray-400 uppercase tracking-wider">Registro Scansioni (doppioni in evidenza):</label>
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
                    {ordered.map((s, idx) => {
                      const dup = counts[s.serial] > 1;
                      return (
                        <tr key={idx} className={`border-b border-gray-100 text-xs font-mono transition ${dup ? 'bg-red-50 hover:bg-red-100' : 'hover:bg-gray-50/80'}`}>
                          <td className={`p-2.5 pl-4 font-bold ${dup ? 'text-red-700' : 'text-gray-900'}`}>
                            {dup ? '🔴' : '🟢'} {s.serial}{dup ? ` (×${counts[s.serial]})` : ''}
                          </td>
                          <td className="p-2.5 text-gray-600">{s.model}</td>
                          <td className="p-2.5 text-gray-500">{s.pn}</td>
                          <td className="p-2.5 pr-4 text-right text-gray-400">{s.time}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-3 pt-2">
              <button onClick={confirmAndFinalizeVerification} disabled={hasDuplicates}
                className={`w-full font-black p-4 rounded-xl text-base shadow-md transition flex items-center justify-center gap-2 ${hasDuplicates ? 'bg-gray-200 text-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700 text-white cursor-pointer'}`}>
                {hasDuplicates ? '⚠ Rimuovi i doppioni per registrare' : '✓ Approva e Registra Carico su Cloud'}
              </button>
            </div>
          </div>
          );
        })()}

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


      {activeModule === 'spare-parts' && (Object.keys(spPendingChanges).length > 0 || spNewRows.length > 0) && (
        <div className="fixed bottom-0 left-0 right-0 z-50 bg-white border-t border-amber-200 shadow-lg px-4 py-3 flex items-center justify-between gap-4">
          <span className="text-sm font-bold text-amber-700">✏️ {Object.keys(spPendingChanges).length} modificate{spNewRows.length > 0 ? ` · ${spNewRows.length} nuove` : ''}</span>
          <div className="flex gap-2">
            <button onClick={() => { setSpPendingChanges({}); setSpNewRows([]); setSpEditMode(false); }}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-gray-100 hover:bg-gray-200 text-gray-600 cursor-pointer transition">Annulla tutto</button>
            <button onClick={saveAllSpChanges}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-green-600 hover:bg-green-700 text-white cursor-pointer transition shadow-xs">
              ✓ Salva
            </button>
          </div>
        </div>
      )}

      {activeModule === 'stock' && (Object.keys(stockPendingChanges).length > 0 || stockNewRows.length > 0) && (
        <div className="fixed bottom-0 left-0 right-0 z-50 bg-white border-t border-amber-200 shadow-lg px-4 py-3 flex items-center justify-between gap-4">
          <span className="text-sm font-bold text-amber-700">
            ✏️ {Object.keys(stockPendingChanges).length} modificate{stockNewRows.length > 0 ? ` · ${stockNewRows.length} nuove` : ''}
          </span>
          <div className="flex gap-2">
            <button onClick={() => { setStockPendingChanges({}); setStockNewRows([]); }}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-gray-100 hover:bg-gray-200 text-gray-600 cursor-pointer transition">
              Annulla tutto
            </button>
            <button onClick={saveAllStockChanges}
              className="text-sm font-bold px-4 py-2 rounded-xl bg-green-600 hover:bg-green-700 text-white cursor-pointer transition shadow-xs">
              ✓ Salva
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
