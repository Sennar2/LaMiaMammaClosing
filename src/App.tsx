/* ================================================
   src/App.tsx  — PART 1/3 (UPDATED)
   ================================================ */
import React, { useEffect, useMemo, useState } from 'react';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

// Sales PDF helpers (NOW also import parseTotalsFromText)
import { extractPdfText, parseSalesFromText, parseTotalsFromText } from './utils/parseSalesPdf';

/* ===================== Config ===================== */
const APPS_SCRIPT_URL =
  'https://script.google.com/macros/s/AKfycbymfW-EJlPgFgW9UEpdAb6GqZm5wLc_TNCqd0lgF8b58E1IDwtDwe29kcNBZJJc1vovRQ/exec';

const RECIPIENT_CSV_URL =
  import.meta.env.VITE_RECIPIENT_CSV_URL ||
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vS2uBaXbEQf62AsIqXTADN3tTBdf87EB2Ab4kCNS_tmse5coNKWOIf4hMakDD0GamKW7qMW1UF2hwc-/pub?output=csv';

const SENDER_NAME = 'Closing';

/* ===================== Types ===================== */
type Location = { id: string; name: string; email?: string };

type TipRow = { id: string; name: string; amount: string };

type VoidDiscountRow = {
  id: string;
  type: 'Void' | 'Discount' | string;
  item: string;
  amount: string;
  reason: string;
  employee: string;
  user?: string;
  confirmed?: boolean;
  // NEW: manager must comment if they select "No"
  managerComment?: string;
};

type Issues = {
  bar86: boolean;
  bar86Impact: string;
  kitchen86: boolean;
  kitchen86Impact: string;
  guestFeedback: string;
  issuesChallenges: string;
  teamPerformance: string;
  weatherImpact: string;
  dishTasted: string;
  fireDrillDate: string;
  emergencyLightDate: string;
  weeklyFireAlarmDate: string;
  flowTrainingMissing: string;
  uniformOk: boolean;
  uniformMissing: string;
  tabletsCount: string;
  cableChargersCount: string;
  chargerBricksCount: string;
  pdqsCount: string;
  foodBiblesCount: string;
  allergensFoldersCount: string;
  actionPoints: string;
};

/* ===================== Utils ===================== */
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(
    d.getDate()
  ).padStart(2, '0')}`;
}
// NEW: if managers submit after midnight, default to yesterday until 6am
function serviceDateISO(cutoffHour = 6) {
  const d = new Date();
  if (d.getHours() < cutoffHour) d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(
    d.getDate()
  ).padStart(2, '0')}`;
}
function currency(n: unknown) {
  if (n === null || n === undefined || n === '') return '£0.00';
  const num =
    typeof n === 'number' ? n : parseFloat(String(n).replace(/[^0-9.-]/g, '')) || 0;
  return num.toLocaleString(undefined, { style: 'currency', currency: 'GBP' });
}
function uuid() {
  // @ts-ignore
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}
function useLocalStorage<T>(key: string, initial: T) {
  const [state, setState] = useState<T>(() => {
    try {
      const raw = localStorage.getItem(key);
      return raw ? (JSON.parse(raw) as T) : initial;
    } catch {
      return initial;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem(key, JSON.stringify(state));
    } catch {}
  }, [key, state]);
  return [state, setState] as const;
}
async function loadImageAsDataURL(url: string): Promise<string | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

/* ===== Excel helpers ===== */
const norm = (x: unknown) =>
  String(x ?? '')
    .replace(/\u00A0/g, ' ')
    .toLowerCase()
    .trim()
    .replace(/[\s._-]+/g, '')
    .replace(/[^a-z0-9]/g, '');

function parseAmountCell(v: unknown): number {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  let t = String(v).trim();
  if (/^\(.*\)$/.test(t)) t = '-' + t.slice(1, -1);
  t = t.replace(/[^0-9.,-]/g, '').replace(/,/g, '');
  const n = parseFloat(t);
  return Number.isFinite(n) ? n : 0;
}

/* ===================== Excel Parsers ===================== */
async function parseVoidsExcel(file: File): Promise<VoidDiscountRow[]> {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: 'array' });

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const A = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) as any[][];
    if (!A || !A.length) continue;

    let headerRow = -1;
    let header: string[] = [];
    for (let i = 0; i < Math.min(25, A.length); i++) {
      const row = A[i] || [];
      const n = row.map(norm);
      const ok =
        n.some((c) => c.includes('item')) &&
        n.some((c) => ['gross', 'amount', 'value', 'total'].some((k) => c.includes(k)));
      if (ok) {
        headerRow = i;
        header = (row as any[]).map((c) => String(c ?? ''));
        break;
      }
    }
    if (headerRow === -1) continue;

    const headNorm = header.map(norm);
    const findIdx = (aliases: string[]) =>
      headNorm.findIndex((h) => aliases.some((a) => h === a || h.includes(a)));

    const idxItem = findIdx(['itemname', 'item']);
    const idxAmount = findIdx(['gross', 'amount', 'value', 'total']);
    const idxReason = findIdx(['reasontext', 'reason']);
    const idxUser = findIdx(['user', 'employee', 'staff', 'name']);

    const out: VoidDiscountRow[] = [];
    for (let r = headerRow + 1; r < A.length; r++) {
      const row = A[r] || [];
      if (!row || row.every((c: any) => c == null || String(c).trim() === '')) continue;

      const item = idxItem >= 0 ? String(row[idxItem] ?? '').trim() : '';
      const amt = idxAmount >= 0 ? parseAmountCell(row[idxAmount]) : 0;
      const reason = idxReason >= 0 ? String(row[idxReason] ?? '').trim() : '';
      const user = idxUser >= 0 ? String(row[idxUser] ?? '').trim() : '';

      if (!item && !amt && !reason && !user) continue;

      out.push({
        id: `v-${sheetName}-${r}`,
        type: 'Void',
        item,
        amount: String(amt),
        reason,
        employee: user,
        user,
        confirmed: false,
        managerComment: '',
      });
    }
    if (out.length) return out;
  }
  return [];
}

async function parseDiscountsExcel(file: File): Promise<VoidDiscountRow[]> {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: 'array' });

  const looksLikeTitle = (product: string) => {
    const n = norm(product);
    return !n || ['product', 'discount', 'discounts', 'reason', 'user', 'name'].includes(n);
  };

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const A = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) as any[][];
    if (!A || !A.length) continue;

    let headerRow = -1;
    let header: string[] = [];
    for (let i = 0; i < Math.min(40, A.length); i++) {
      const row = A[i] || [];
      const n = row.map(norm);
      const hasProduct = n.includes('product');
      const hasDiscounts = n.includes('discounts') || n.includes('discount');
      if (hasProduct && hasDiscounts) {
        headerRow = i;
        header = (row as any[]).map((c) => String(c ?? ''));
        break;
      }
    }
    if (headerRow === -1) continue;

    const headNorm = header.map(norm);
    const findExact = (name: string) => headNorm.findIndex((h) => h === name);

    const idxProduct = findExact('product');
    const idxDiscounts = headNorm.findIndex((h) => h === 'discounts' || h === 'discount');
    const idxReason = findExact('reason');
    const idxUser = findExact('user');
    const idxName = findExact('name');

    const rows: VoidDiscountRow[] = [];
    for (let r = headerRow + 1; r < A.length; r++) {
      const row = A[r] || [];
      if (!row || row.every((c: any) => c == null || String(c).trim() === '')) continue;

      const product = idxProduct >= 0 ? String(row[idxProduct] ?? '').trim() : '';
      if (!product || looksLikeTitle(product)) continue;

      const amountN = idxDiscounts >= 0 ? parseAmountCell(row[idxDiscounts]) : 0;
      const reason = idxReason >= 0 ? String(row[idxReason] ?? '').trim() : '';
      const user = idxUser >= 0 ? String(row[idxUser] ?? '').trim() : '';
      const name = idxName >= 0 ? String(row[idxName] ?? '').trim() : '';

      rows.push({
        id: `d-${sheetName}-${r}`,
        type: 'Discount',
        item: product,
        amount: String(amountN),
        reason,
        employee: name,
        user,
        confirmed: false,
        managerComment: '',
      });
    }

    if (rows.length) return rows;
  }
  return [];
}

/* ===================== Component ===================== */
export default function App() {
  // Recipients & locations
  const [config, setConfig] = useState<{
    locations: Location[];
    defaults: { to: string[]; cc: string[]; bcc: string[] };
  }>({
    locations: [],
    defaults: { to: [], cc: [], bcc: [] },
  });

  const [selectedLocId, setSelectedLocId] = useLocalStorage<string>('selectedLocId', '');

  const [store, setStore] = useLocalStorage('store', {
    manager: '',
    recipients: '',
    // NEW: service date default (yesterday before 6am)
    date: serviceDateISO(),
  });

  // Sales
  const [sales, setSales] = useLocalStorage('sales', {
    grossSales: '',
    netServ: '',
    cash: '',
    card: '',
    covers: '',
    deliveroo: '',
    notes: '',
    serviceCharge: '',
    tipsTotal: '',
    tipsRows: [] as TipRow[],
  });
  const [salesFile, setSalesFile] = useState<File | null>(null);

  // Track manual edits so auto-fill doesn't overwrite unless asked
  const [salesEdited, setSalesEdited] = useState<Record<string, boolean>>({});
  const markEdited = (k: string) => setSalesEdited((p) => ({ ...p, [k]: true }));

  // NEW: parsed values (Today row + Week totals), and UI state for the preview tabs
  const [parsedSales, setParsedSales] = useState<Partial<{
    grossSales: string;
    netSales: string;
    serviceCharge: string;
    cash: string;
    card: string;
    deliveroo: string;
    covers: string;
  }> | null>(null);
  const [parsedWeek, setParsedWeek] = useState<typeof parsedSales>(null);
  const [parseStatus, setParseStatus] = useState<'' | 'ready' | 'applied' | 'error'>('');
  const [activeSalesTab, setActiveSalesTab] = useState<'today' | 'week'>('today');

  // Voids + Discounts combined
  const [voidsDiscounts, setVoidsDiscounts] = useLocalStorage<VoidDiscountRow[]>(
    'voidsDiscounts',
    []
  );

  // Issues
  const [issues, setIssues] = useLocalStorage<Issues>('issues', {
    bar86: false,
    bar86Impact: '',
    kitchen86: false,
    kitchen86Impact: '',
    guestFeedback: '',
    issuesChallenges: '',
    teamPerformance: '',
    weatherImpact: '',
    dishTasted: '',
    fireDrillDate: '',
    emergencyLightDate: '',
    weeklyFireAlarmDate: '',
    flowTrainingMissing: '',
    uniformOk: true,
    uniformMissing: '',
    tabletsCount: '',
    cableChargersCount: '',
    chargerBricksCount: '',
    pdqsCount: '',
    foodBiblesCount: '',
    allergensFoldersCount: '',
    actionPoints: '',
  });

  // UI
  const [activeTab, _setActiveTab] =
    useState<'store' | 'sales' | 'voids' | 'issues' | 'send'>('store');
  const [sendMsg, setSendMsg] = useState('');
  const [sending, setSending] = useState(false);

  /* ===== Load locations/recipients from CSV ===== */
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch(RECIPIENT_CSV_URL);
        if (!res.ok) throw new Error('CSV fetch failed');
        const text = await res.text();
        const parsed = Papa.parse(text, { header: true });
        const rows = (parsed.data as unknown[]).filter(Boolean) as Record<string, any>[];

        const locs: Location[] = [];
        let to: string[] = [];
        let cc: string[] = [];
        let bcc: string[] = [];

        const splitEmails = (s: unknown) =>
          String(s || '')
            .split(/[;,\n\t ]+/)
            .map((x) => x.trim())
            .filter(Boolean);

        for (const row of rows) {
          const emailCol =
            row['Eamil to send to'] || row['Email to send to'] || row['Emails'] || '';
          const ccCol = row['CC'] || row['Cc'] || row['Cc emails'] || row['CC emails'] || '';
          const bccCol =
            row['BCC'] || row['Bcc'] || row['Bcc emails'] || row['BCC emails'] || '';
          const locCol = row['Location List'] || row['Locations'] || '';

          to = to.concat(splitEmails(emailCol));
          cc = cc.concat(splitEmails(ccCol));
          bcc = bcc.concat(splitEmails(bccCol));

          if (locCol) {
            const name = String(locCol).trim();
            const id = name;
            locs.push({ id, name });
          }
        }

        const dedupe = (arr: string[]) => Array.from(new Set(arr.filter(Boolean)));
        to = dedupe(to);
        cc = dedupe(cc);
        bcc = dedupe(bcc);

        const map = new Map<string, Location>();
        for (const l of locs) if (!map.has(l.id)) map.set(l.id, l);
        const dedupLocs = Array.from(map.values());

        setConfig({ locations: dedupLocs, defaults: { to, cc, bcc } });
        if (dedupLocs.length && !selectedLocId) setSelectedLocId(dedupLocs[0].id);
      } catch (e) {
        console.error(e);
        setSendMsg('Could not load recipients from CSV.');
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const selectedLocation = useMemo(
    () => config.locations.find((l) => l.id === selectedLocId) || null,
    [config.locations, selectedLocId]
  );
  const locationName = selectedLocation?.name || 'Store';

  const finalRecipients = useMemo(() => {
    const baseTo = config.defaults.to || [];
    const extra = (store.recipients || '')
      .split(/[,;\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    return {
      to: Array.from(new Set([...baseTo, ...extra])),
      cc: config.defaults.cc || [],
      bcc: config.defaults.bcc || [],
    };
  }, [config.defaults, store.recipients]);

  /* ===== Helper to get weekday label for parser ===== */
  const dayNameFromStore = () => {
    if (!store?.date) return undefined;
    const d = new Date(`${store.date}T00:00:00`);
    const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    return days[d.getDay()];
  };

  /* ======= CONTINUES IN PART 2/3... ======= */

/* ================================================
   src/App.tsx  — PART 2/3 (UPDATED)
   ================================================ */

  /* ===== Derived totals ===== */
  const manualTipsCalc = useMemo(() => {
    const rowsSum = (sales.tipsRows || []).reduce(
      (s, r) => s + (parseFloat(r.amount) || 0),
      0
    );
    const manual = parseFloat(String(sales.tipsTotal).replace(/[^0-9.-]/g, '')) || 0;
    return manual || rowsSum;
  }, [sales.tipsRows, sales.tipsTotal]);

  const serviceChargeCalc =
    parseFloat(String(sales.serviceCharge || '').replace(/[^0-9.-]/g, '')) || 0;
  const tipsTotalCalc = serviceChargeCalc + manualTipsCalc;

  const voidsTotal = useMemo(
    () => (voidsDiscounts || []).reduce((sum, r) => sum + (parseFloat(r.amount) || 0), 0),
    [voidsDiscounts]
  );

  // Gate: must confirm YES or leave comment when NO
  const canProceedVoids = useMemo(() => {
    if ((voidsDiscounts || []).length === 0) return true;
    return (voidsDiscounts || []).every((r) =>
      r.confirmed ? true : Boolean((r.managerComment || '').trim())
    );
  }, [voidsDiscounts]);

  /* ===== Nav Guard ===== */
  function setActiveTabGuard(next: typeof activeTab) {
    if ((next === 'issues' || next === 'send') && !canProceedVoids) {
      setSendMsg('Please add a manager comment for every item marked "No" to proceed.');
      _setActiveTab('voids');
      return;
    }
    _setActiveTab(next);
  }

  /* ===================== SALES PDF AUTO-FILL ===================== */
  // Apply parsed values (Today or Week preview → form)
  function applyParsed(source: 'today' | 'week', mode: 'empty' | 'overwrite') {
    const payload = source === 'today' ? parsedSales : parsedWeek;
    if (!payload) return;
    setSales((prev) => {
      const next = { ...prev };
      const pairs: Array<[keyof typeof payload, keyof typeof prev]> = [
        ['grossSales', 'grossSales'],
        ['netSales', 'netServ'],
        ['cash', 'cash'],
        ['card', 'card'],
        ['serviceCharge', 'serviceCharge'],
        ['deliveroo', 'deliveroo'],
        ['covers', 'covers'],
      ];
      for (const [from, to] of pairs) {
        const incoming = (payload as any)?.[from];
        if (!incoming) continue;
        const wasEdited = salesEdited[to as string];
        if (mode === 'overwrite' || !prev[to] || !wasEdited) {
          (next as any)[to] = String(incoming);
        }
      }

      // Migrate drafts created by the previous version, which stored an
      // auto-filled service charge in the manual tips field.
      const incomingServiceCharge = String(payload.serviceCharge || '');
      const previousServiceCharge = String(prev.serviceCharge || '');
      const previousTips = String(prev.tipsTotal || '');
      const normaliseAmount = (value: string) =>
        Number.parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
      if (
        incomingServiceCharge &&
        !previousServiceCharge &&
        !salesEdited.tipsTotal &&
        normaliseAmount(previousTips) === normaliseAmount(incomingServiceCharge)
      ) {
        next.tipsTotal = '';
      }
      return next;
    });
    setParseStatus('applied');
  }

  // When a sales PDF is selected, parse both the "today" row and the TOTALS row
  async function onSalesPdfSelected(file: File | null) {
    setSalesFile(file || null);
    setParsedSales(null);
    setParsedWeek(null);
    setParseStatus('');
    if (!file) return;
    try {
      const text = await extractPdfText(file);

      const todayParsed = parseSalesFromText(text, { day: dayNameFromStore() });
      const weekTotals  = parseTotalsFromText(text);

      const keep = (p: any) => ({
        grossSales: p?.grossSales,
        netSales:   p?.netSales,
        serviceCharge: p?.serviceCharge,
        cash:       p?.cash,
        card:       p?.card,
        deliveroo:  p?.deliveroo,
        covers:     p?.covers,
      });

      setParsedSales(keep(todayParsed));
      setParsedWeek(keep(weekTotals));
      setParseStatus('ready');
    } catch (e) {
      console.error(e);
      setParsedSales(null);
      setParsedWeek(null);
      setParseStatus('error');
      setSendMsg('Could not read that PDF for auto-fill.');
    }
  }

  /* ===================== PDF generation (everything included) ===================== */
  async function generatePDFBlob() {
    const doc = new jsPDF({ unit: 'pt', format: 'a4' });

    const margin = 46;
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    let y = margin;

    const FONT = { normal: 'helvetica', bold: 'helvetica' };
    function ensure(h = 24) {
      if (y + h > pageH - margin) {
        doc.addPage();
        y = margin;
      }
    }
    function divider() {
      doc.setDrawColor(210);
      doc.setLineWidth(0.8);
      doc.line(margin, y, pageW - margin, y);
      y += 12;
    }
    function section(title: string) {
      ensure(84);
      y += 8;
      doc.setFont(FONT.bold, 'bold');
      doc.setFontSize(16);
      doc.setTextColor(0, 0, 0);
      doc.text(title, margin, y);
      y += 20;
      divider();
      y += 8;
    }
    function kv(label: string, value: string) {
      ensure(28);
      doc.setFont(FONT.bold, 'bold');
      doc.setFontSize(11);
      doc.setTextColor(70, 74, 79);
      doc.text(label, margin, y);

      doc.setFont(FONT.normal, 'normal');
      doc.setTextColor(25, 25, 25);
      doc.text(value || '—', pageW - margin, y, { align: 'right' } as any);
      y += 22;
    }
    function para(label: string, text: string) {
      ensure(36);
      doc.setFont(FONT.bold, 'bold');
      doc.setFontSize(11);
      doc.setTextColor(70, 74, 79);
      doc.text(label, margin, y);
      y += 14;

      doc.setFont(FONT.normal, 'normal');
      doc.setTextColor(25, 25, 25);
      const lines = doc.splitTextToSize((text || '—').trim() || '—', pageW - margin * 2) as string[];
      for (const ln of lines) {
        ensure(18);
        doc.text(ln, margin, y);
        y += 16;
      }
      y += 12;
    }
    function table(headers: string[], rows: (string | number)[][]) {
      const baseRowH = 28;
      const colW = (pageW - margin * 2) / headers.length;

      ensure(baseRowH + 16);
      doc.setFillColor(242, 244, 247);
      doc.rect(margin, y, colW * headers.length, baseRowH, 'F');
      doc.setFont(FONT.bold, 'bold');
      doc.setFontSize(11);
      doc.setTextColor(0, 0, 0);
      headers.forEach((h, i) => doc.text(String(h), margin + i * colW + 6, y + 18));
      y += baseRowH;

      doc.setFont(FONT.normal, 'normal');
      doc.setTextColor(25, 25, 25);
      rows.forEach((r, rIdx) => {
        const wraps: string[][] = r.map((cell) =>
          doc.splitTextToSize(String(cell ?? ''), colW - 12) as string[]
        );
        const maxLines = Math.max(1, ...wraps.map((w) => w.length));
        const rowH = Math.max(baseRowH, maxLines * 14 + 10);
        ensure(rowH);

        if (rIdx % 2 === 0) {
          doc.setFillColor(252, 253, 255);
          doc.rect(margin, y - 2, colW * headers.length, rowH + 2, 'F');
        }

        wraps.forEach((lines: string[], i: number) => {
          lines.forEach((ln: string, idx: number) => {
            doc.text(ln, margin + i * colW + 6, y + 18 + idx * 14);
          });
        });

        y += rowH;
      });

      y += 10;
    }

    try {
      const dataUrl = await loadImageAsDataURL('/logo.png');
      if (dataUrl) {
        const imgW = 120;
        const imgH = 40;
        const imgX = (pageW - imgW) / 2;
        doc.addImage(dataUrl, 'PNG', imgX, y, imgW, imgH, undefined, 'FAST');
        y += imgH + 18;
      }
    } catch {}

    doc.setFont(FONT.bold, 'bold');
    doc.setFontSize(18);
    doc.setTextColor(0, 0, 0);
    doc.text(`${selectedLocation?.name || 'Store'} — Closing Report`, margin, y);
    y += 20;

    doc.setFont(FONT.normal, 'normal');
    doc.setFontSize(11);
    doc.setTextColor(90, 90, 90);
    doc.text(`Generated: ${new Date().toLocaleString()}`, margin, y);
    y += 26;

    section('Overview');
    kv('Date', store.date);
    kv('Manager', store.manager || '—');
    kv('Location', selectedLocation?.name || '—');

    section('Sales Summary');
    kv('Gross Sales', currency(sales.grossSales));
    kv('Net+Serv', currency(sales.netServ));
    kv('Cash', currency(sales.cash));
    kv('Card', currency(sales.card));
    kv('Covers', String(sales.covers || 0));
    kv('Deliveroo', currency(sales.deliveroo));
    kv('Service Charge', currency(serviceChargeCalc));
    kv('Additional Tips', currency(manualTipsCalc));
    kv('Service Charge + Tips', currency(tipsTotalCalc));
    if ((sales.tipsRows || []).length) {
      para('Tips Breakdown', '');
      table(
        ['Waiter', 'Amount'],
        sales.tipsRows.map((r) => [r.name || '—', currency(r.amount || 0)])
      );
    }
    if ((sales.notes || '').trim()) para('Sales Notes', sales.notes);

    section('Voids & Discounts');
    if ((voidsDiscounts || []).length) {
      table(
        ['#', 'Type', 'Item', 'Amount', 'Reason', 'Employee', 'User', 'Confirmed', 'Manager Comment'],
        voidsDiscounts.map((r, i) => [
          String(i + 1),
          r.type || '—',
          r.item || '—',
          currency(r.amount || 0),
          r.reason || '—',
          r.employee || '—',
          r.user || '—',
          r.confirmed ? 'Yes' : 'No',
          (r.managerComment || '—'),
        ])
      );
      kv('Total Voids+Discounts', currency(
        (voidsDiscounts || []).reduce((sum, r) => sum + (parseFloat(r.amount) || 0), 0)
      ));
    } else {
      para('', 'No voids or discounts recorded.');
    }

    section('Daily Recap & Issues');
    kv('86 items (Bar)', (issues.bar86 ? 'Yes' : 'No') + (issues.bar86Impact ? ` — ${issues.bar86Impact}` : ''));
    kv('86 items (Kitchen)', (issues.kitchen86 ? 'Yes' : 'No') + (issues.kitchen86Impact ? ` — ${issues.kitchen86Impact}` : ''));
    para('Guest Feedback', issues.guestFeedback);
    para('Issues & Challenges', issues.issuesChallenges);
    para('Team Performance', issues.teamPerformance);
    para('Weather Impact', issues.weatherImpact);
    para('Team Tasting — Dish & Feedback', issues.dishTasted);

    section('Compliance — Last Done Dates');
    kv('Fire Drill', issues.fireDrillDate || '—');
    kv('Monthly Emergency Light Checks', issues.emergencyLightDate || '—');
    kv('Weekly Fire Alarm Check', issues.weeklyFireAlarmDate || '—');

    section('Training & Uniform');
    para('FLOW Training — Missing Team Members', issues.flowTrainingMissing);
    kv('Full team with correct uniform?', issues.uniformOk ? 'Yes' : 'No');
    if (!issues.uniformOk) para('Uniform issues', issues.uniformMissing);

    section('Equipment / Collateral Counts');
    ([
      ['Tablets', issues.tabletsCount],
      ['Cable Chargers', issues.cableChargersCount],
      ['Charger Bricks', issues.chargerBricksCount],
      ['PDQs', issues.pdqsCount],
      ['Food Bibles', issues.foodBiblesCount],
      ['Allergens Folders', issues.allergensFoldersCount],
    ] as const).forEach(([label, val]) => kv(String(label), String(val || '0')));

    section('Action Points (follow-ups)');
    para('', issues.actionPoints);

    ensure(34);
    const footerY = doc.internal.pageSize.getHeight() - margin + 12;
    doc.setFont(FONT.normal, 'normal');
    doc.setFontSize(9);
    doc.setTextColor(140, 140, 140);
    doc.text('Generated by Closing Report Wizard', margin, footerY);

    const blob = doc.output('blob') as Blob & { name?: string };
    blob.name = `Closing_Report_${selectedLocation?.name || 'Store'}_${store.date}.pdf`;
    return blob;
  }

  /* ===================== Attachments helpers ===================== */
  async function readBlobAsBase64(blob: Blob): Promise<string> {
    const buf = await blob.arrayBuffer();
    let binary = '';
    const bytes = new Uint8Array(buf);
    const step = 0x8000;
    for (let i = 0; i < bytes.length; i += step) {
      // eslint-disable-next-line @typescript-eslint/no-implied-eval
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + step) as any);
    }
    return btoa(binary);
  }
  async function fileToAttachment(f: File) {
    return {
      filename: f.name,
      mimeType: f.type || 'application/octet-stream',
      data: await readBlobAsBase64(f),
    };
  }

  /* ===================== Excel export (itemised) ===================== */
  function generateExcelBlob() {
    const rows = (voidsDiscounts || []).map((r, i) => ({
      No: i + 1,
      Type: r.type,
      Item: r.item,
      Amount: Number(r.amount) || 0,
      Reason: r.reason,
      Employee: r.employee,
      User: r.user || '',
      Confirmed: r.confirmed ? 'Yes' : 'No',
      ManagerComment: r.managerComment || '',
      Date: store.date,
      Store: locationName,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Voids_Discounts');
    const array = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([array], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }) as Blob & { name?: string };
    blob.name = `Voids_Discounts_${locationName}_${store.date}.xlsx`;
    return blob;
  }

  /* ===================== Send Email ===================== */
  async function sendEmailNow() {
    try {
      setSending(true);
      setSendMsg('');

      const pdf = await generatePDFBlob();
      const xlsx = generateExcelBlob();
      const attachments: Array<{ filename: string; mimeType: string; data: string }> = [
        {
          filename: (pdf as any).name || 'Closing_Report.pdf',
          mimeType: 'application/pdf',
          data: await readBlobAsBase64(pdf),
        },
        {
          filename: (xlsx as any).name || 'Voids_Discounts.xlsx',
          mimeType:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          data: await readBlobAsBase64(xlsx),
        },
      ];
      if (salesFile) attachments.push(await fileToAttachment(salesFile));

      const subject = `${locationName} close`;
      const body = [
        'Hi team,',
        '',
        'Closing report attached.',
        '',
        'Sales Recap:',
        `• Gross: ${currency(sales.grossSales)} | Net+Serv: ${currency(sales.netServ)}`,
        `• Cash: ${currency(sales.cash)} | Card: ${currency(sales.card)} | Covers: ${sales.covers || 0} | Deliveroo: ${currency(sales.deliveroo)}`,
        `• Service Charge: ${currency(serviceChargeCalc)} | Additional Tips: ${currency(manualTipsCalc)} | Total: ${currency(tipsTotalCalc)}`,
        '',
        'Voids & Discounts:',
        `• Items: ${(voidsDiscounts || []).length} | Total: ${currency(voidsTotal)}`,
        '',
        'Regards,',
        store.manager || locationName,
      ].join('\n');

      const res = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          subject,
          textBody: body,
          to: finalRecipients.to,
          cc: finalRecipients.cc,
          bcc: finalRecipients.bcc,
          attachments,
          senderName: SENDER_NAME,
        }),
      });

      if (!res.ok) throw new Error('Send failed');
      const j = (await res.json()) as { status?: string; message?: string };
      if (j.status !== 'ok') throw new Error(j.message || 'Unknown error');
      setSendMsg('Email sent successfully.');
    } catch (e: unknown) {
      setSendMsg('Error: ' + (e as Error).message);
      console.error(e);
    } finally {
      setSending(false);
    }
  }

/* ======= CONTINUES IN PART 3/3... ======= */

/* ================================================
   src/App.tsx  — PART 3/3 (UPDATED)
   ================================================ */

  /* ===================== UI ===================== */
  return (
    <div>
      {/* Top bar */}
      <header
        style={{
          position: 'sticky',
          top: 0,
          background: '#fff',
          zIndex: 5,
          borderBottom: '1px solid #e5e7eb',
        }}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            padding: '12px 24px',
            alignItems: 'center',
          }}
        >
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <img src="/logo.png" alt="logo" style={{ height: 36 }} />
            <div style={{ fontWeight: 700 }}>Closing Report Wizard</div>
          </div>
          <span style={{ fontSize: 12, color: '#6b7280' }}>Light UI</span>
        </div>
      </header>

      <div style={{ maxWidth: 1100, margin: '0 auto', padding: 24 }}>
        {/* Tabs */}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
          {(['store', 'sales', 'voids', 'issues', 'send'] as const).map((k) => {
            const disabled = (k === 'issues' || k === 'send') && !canProceedVoids;
            return (
              <button
                key={k}
                onClick={() => setActiveTabGuard(k)}
                disabled={disabled}
                style={{
                  padding: '10px 14px',
                  borderRadius: 10,
                  border: '1px solid #e5e7eb',
                  background: activeTab === k ? '#111827' : disabled ? '#f3f4f6' : '#fff',
                  color: activeTab === k ? '#fff' : '#111827',
                  cursor: disabled ? 'not-allowed' : 'pointer',
                }}
              >
                {k[0].toUpperCase() + k.slice(1)}
              </button>
            );
          })}
        </div>

        {/* STORE */}
        {activeTab === 'store' && (
          <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
            <div
              style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}
            >
              Location & Recipients
            </div>
            <div style={{ padding: '16px 20px' }}>
              <div style={{ display: 'grid', gap: 14 }}>
                <div>
                  <label>Location</label>
                  <select
                    value={selectedLocId}
                    onChange={(e) => setSelectedLocId(e.target.value)}
                    style={{ width: '100%', height: 42 }}
                  >
                    {config.locations.length === 0 && <option>Loading…</option>}
                    {config.locations.map((l) => (
                      <option key={l.id} value={l.id}>
                        {l.name}
                      </option>
                    ))}
                  </select>
                  <div style={{ fontSize: 12, color: '#6b7280', marginTop: 6 }}>
                    Locations & recipients are loaded from your Google Sheet CSV.
                  </div>
                </div>
              </div>

              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(2, 1fr)',
                  gap: 14,
                  marginTop: 14,
                }}
              >
                <div>
                  <label>Manager on Duty</label>
                  <input
                    value={store.manager}
                    onChange={(e) => setStore({ ...store, manager: e.target.value })}
                    placeholder="Full name"
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
                <div>
                  <label>Date</label>
                  <input
                    type="date"
                    value={store.date}
                    onChange={(e) => setStore({ ...store, date: e.target.value })}
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
              </div>

              <div style={{ display: 'grid', gap: 14, marginTop: 14 }}>
                <div>
                  <label>Additional recipients (optional)</label>
                  <input
                    value={store.recipients}
                    onChange={(e) => setStore({ ...store, recipients: e.target.value })}
                    placeholder="email1@company.com, email2@company.com"
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* SALES (with Today / Week tabs) */}
        {activeTab === 'sales' && (
          <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
            <div
              style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}
            >
              Sales Recap
            </div>
            <div style={{ padding: '16px 20px' }}>
              <div style={{ display: 'grid', gap: 8, marginBottom: 12 }}>
                <div>
                  <label>Upload sales PDF (auto-fill & attach)</label>
                  <input
                    type="file"
                    accept="application/pdf"
                    onChange={async (e) => {
                      const f = e.currentTarget.files?.[0] || null;
                      await onSalesPdfSelected(f);
                    }}
                  />
                  <div style={{ fontSize: 12, color: '#6b7280', marginTop: 4 }}>
                    This file will be used to auto-fill and also sent as an attachment.
                  </div>
                </div>

                {/* === Auto-fill previews (Today / Week TOTAL) === */}
                {(parsedSales || parsedWeek || parseStatus === 'error') && (
                  <div style={{ marginTop: 6, border: '1px solid #e5e7eb', borderRadius: 10 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #e5e7eb' }}>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button
                          onClick={() => setActiveSalesTab('today')}
                          style={{
                            padding: '6px 10px',
                            border: '1px solid #e5e7eb',
                            borderRadius: 8,
                            background: activeSalesTab === 'today' ? '#111827' : '#fff',
                            color: activeSalesTab === 'today' ? '#fff' : '#111827',
                          }}
                        >
                          Today ({dayNameFromStore() || '—'})
                        </button>
                        <button
                          onClick={() => setActiveSalesTab('week')}
                          style={{
                            padding: '6px 10px',
                            border: '1px solid #e5e7eb',
                            borderRadius: 8,
                            background: activeSalesTab === 'week' ? '#111827' : '#fff',
                            color: activeSalesTab === 'week' ? '#fff' : '#111827',
                          }}
                        >
                          Week (TOTAL)
                        </button>
                      </div>

                      <div style={{ marginLeft: 'auto', display: 'flex', gap: 8 }}>
                        {parseStatus === 'ready' && (
                          <>
                            <button onClick={() => applyParsed(activeSalesTab, 'empty')}>Apply to empty</button>
                            <button onClick={() => applyParsed(activeSalesTab, 'overwrite')}>Overwrite all</button>
                          </>
                        )}
                        {parseStatus === 'applied' && (
                          <span style={{ padding: '4px 8px', border: '1px dashed #22c55e', borderRadius: 6, color: '#166534' }}>
                            Values auto-filled. You can edit below.
                          </span>
                        )}
                        {parseStatus === 'error' && (
                          <span style={{ padding: '4px 8px', border: '1px dashed #ef4444', borderRadius: 6, color: '#991b1b' }}>
                            Couldn’t read that PDF. Try another file.
                          </span>
                        )}
                      </div>
                    </div>

                    <div style={{ padding: 12 }}>
                      {activeSalesTab === 'today' ? (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
                          <div><b>Gross</b><div>{parsedSales?.grossSales ?? '—'}</div></div>
                          <div><b>Net + Serv</b><div>{parsedSales?.netSales ?? '—'}</div></div>
                          <div><b>Service Charge</b><div>{parsedSales?.serviceCharge ?? '—'}</div></div>
                          <div><b>Cash</b><div>{parsedSales?.cash ?? '—'}</div></div>
                          <div><b>Card</b><div>{parsedSales?.card ?? '—'}</div></div>
                          <div><b>Deliveroo</b><div>{parsedSales?.deliveroo ?? '—'}</div></div>
                          <div><b>Covers</b><div>{parsedSales?.covers ?? '—'}</div></div>
                        </div>
                      ) : (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
                          <div><b>Gross (Wk)</b><div>{parsedWeek?.grossSales ?? '—'}</div></div>
                          <div><b>Net + Serv (Wk)</b><div>{parsedWeek?.netSales ?? '—'}</div></div>
                          <div><b>Service Charge (Wk)</b><div>{parsedWeek?.serviceCharge ?? '—'}</div></div>
                          <div><b>Cash (Wk)</b><div>{parsedWeek?.cash ?? '—'}</div></div>
                          <div><b>Card (Wk)</b><div>{parsedWeek?.card ?? '—'}</div></div>
                          <div><b>Deliveroo (Wk)</b><div>{parsedWeek?.deliveroo ?? '—'}</div></div>
                          <div><b>Covers (Wk)</b><div>{parsedWeek?.covers ?? '—'}</div></div>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 14 }}>
                <div>
                  <label>Gross Sales</label>
                  <input
                    inputMode="decimal"
                    value={sales.grossSales}
                    onChange={(e) => {
                      setSales({ ...sales, grossSales: e.target.value });
                      markEdited('grossSales');
                    }}
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
                <div>
                  <label>Net+Serv</label>
                  <input
                    inputMode="decimal"
                    value={sales.netServ}
                    onChange={(e) => {
                      setSales({ ...sales, netServ: e.target.value });
                      markEdited('netServ');
                    }}
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
                <div>
                  <label>Cash</label>
                  <input
                    inputMode="decimal"
                    value={sales.cash}
                    onChange={(e) => {
                      setSales({ ...sales, cash: e.target.value });
                      markEdited('cash');
                    }}
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
                <div>
                  <label>Card</label>
                  <input
                    inputMode="decimal"
                    value={sales.card}
                    onChange={(e) => {
                      setSales({ ...sales, card: e.target.value });
                      markEdited('card');
                    }}
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
                <div>
                  <label>Covers</label>
                  <input
                    inputMode="numeric"
                    value={sales.covers}
                    onChange={(e) => setSales({ ...sales, covers: e.target.value })}
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
                <div>
                  <label>Deliveroo</label>
                  <input
                    inputMode="decimal"
                    value={sales.deliveroo}
                    onChange={(e) => setSales({ ...sales, deliveroo: e.target.value })}
                    placeholder="Deliveroo £"
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14, marginTop: 14 }}>
                <div>
                  <label>Service Charge (from sales report)</label>
                  <input
                    inputMode="decimal"
                    value={sales.serviceCharge || ''}
                    readOnly
                    style={{ width: '100%', height: 42, background: '#f3f4f6' }}
                  />
                </div>
                <div>
                  <label>Additional Tips — optional</label>
                  <input
                    inputMode="decimal"
                    value={sales.tipsTotal}
                    onChange={(e) => {
                      setSales({ ...sales, tipsTotal: e.target.value });
                      markEdited('tipsTotal');
                    }}
                    placeholder="If blank, we sum the waiter breakdown"
                    style={{ width: '100%', height: 42 }}
                  />
                </div>
              </div>

              <div style={{ display: 'grid', gap: 10, marginTop: 10 }}>
                <div
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                  }}
                >
                  <label style={{ margin: 0 }}>Tips Breakdown</label>
                  <button
                    onClick={() =>
                      setSales({
                        ...sales,
                        tipsRows: [
                          ...(sales.tipsRows || []),
                          { id: uuid(), name: '', amount: '' },
                        ],
                      })
                    }
                  >
                    Add waiter
                  </button>
                </div>

                {(sales.tipsRows || []).length === 0 && (
                  <div style={{ fontSize: 12, color: '#6b7280' }}>
                    No rows — add your first waiter & amount.
                  </div>
                )}

                {(sales.tipsRows || []).map((r) => (
                  <div
                    key={r.id}
                    style={{
                      display: 'grid',
                      gridTemplateColumns: '7fr 3fr 2fr',
                      gap: 8,
                      alignItems: 'center',
                    }}
                  >
                    <input
                      placeholder="Waiter name"
                      value={r.name}
                      onChange={(e) => {
                        setSales({
                          ...sales,
                          tipsRows: (sales.tipsRows || []).map((x) =>
                            x.id === r.id ? { ...x, name: e.target.value } : x
                          ),
                        });
                      }}
                      style={{ height: 42 }}
                    />
                    <input
                      inputMode="decimal"
                      placeholder="0.00"
                      value={r.amount}
                      onChange={(e) => {
                        setSales({
                          ...sales,
                          tipsRows: (sales.tipsRows || []).map((x) =>
                            x.id === r.id ? { ...x, amount: e.target.value } : x
                          ),
                        });
                      }}
                      style={{ height: 42 }}
                    />
                    <div style={{ textAlign: 'right' }}>
                      <button
                        onClick={() =>
                          setSales({
                            ...sales,
                            tipsRows: (sales.tipsRows || []).filter((x) => x.id !== r.id),
                          })
                        }
                      >
                        Remove
                      </button>
                    </div>
                  </div>
                ))}

                <div style={{ fontSize: 12, color: '#6b7280' }}>
                  Additional Tips: <b>{currency(manualTipsCalc)}</b>
                  {' · '}Service Charge + Tips: <b>{currency(tipsTotalCalc)}</b>
                </div>
              </div>

              <div style={{ display: 'grid', gap: 14, marginTop: 14 }}>
                <div>
                  <label>Sales Notes</label>
                  <textarea
                    value={sales.notes}
                    onChange={(e) => setSales({ ...sales, notes: e.target.value })}
                    placeholder="Anything notable about sales today…"
                    style={{ width: '100%', minHeight: 110 }}
                  />
                </div>
              </div>
            </div>

            <div
              style={{
                padding: '12px 20px',
                borderTop: '1px solid #e5e7eb',
                display: 'flex',
                justifyContent: 'space-between',
              }}
            >
              <div style={{ fontSize: 12, color: '#6b7280' }}>
                The uploaded PDF will be sent with the email.
              </div>
              <button
                onClick={async () => {
                  const b = await generatePDFBlob();
                  const a = document.createElement('a');
                  a.href = URL.createObjectURL(b);
                  a.download = (b as any).name || 'Closing_Report.pdf';
                  a.click();
                }}
              >
                Generate PDF
              </button>
            </div>
          </div>
        )}

        {/* VOIDS / DISCOUNTS (with "No" → comment pop-up) */}
        {activeTab === 'voids' && (
          <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
            <div
              style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}
            >
              Voids & Discounts (Itemised)
            </div>
            <div style={{ padding: '16px 20px' }}>
              <p
                style={{
                  background: '#fff7d6',
                  padding: 10,
                  borderRadius: 8,
                  fontSize: 14,
                  marginBottom: 12,
                }}
              >
                <b>Manager Declaration:</b> For each row select “Yes” to confirm approval, or “No”
                and leave a brief comment to proceed to the next section.
              </p>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14 }}>
                <div>
                  <label>Upload Error Correct (Voids) Excel</label>
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={async (e) => {
                      const f = e.currentTarget.files?.[0];
                      if (!f) return;
                      const rows = await parseVoidsExcel(f);
                      setVoidsDiscounts((prev) => [
                        ...prev.filter((r) => r.type !== 'Void'),
                        ...rows,
                      ]);
                    }}
                  />
                </div>
                <div>
                  <label>Upload Product Discount (Discounts) Excel</label>
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={async (e) => {
                      const f = e.currentTarget.files?.[0];
                      if (!f) return;
                      const rows = await parseDiscountsExcel(f);
                      setVoidsDiscounts((prev) => [
                        ...prev.filter((r) => r.type !== 'Discount'),
                        ...rows,
                      ]);
                    }}
                  />
                </div>
              </div>

              <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12 }}>
                <thead style={{ background: '#f3f4f6' }}>
                  <tr>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      Type
                    </th>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      Item
                    </th>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      Amount
                    </th>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      Reason
                    </th>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      Employee
                    </th>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      User
                    </th>
                    <th style={{ textAlign: 'left', padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                      Confirm
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {(voidsDiscounts || []).length === 0 && (
                    <tr>
                      <td colSpan={7} style={{ padding: 10 }}>
                        <div style={{ fontSize: 12, color: '#6b7280' }}>
                          Upload the two Excel files to populate rows.
                        </div>
                      </td>
                    </tr>
                  )}
                  {(voidsDiscounts || []).map((r) => (
                    <React.Fragment key={r.id}>
                      <tr>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>{r.type}</td>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>{r.item}</td>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                          {currency(r.amount)}
                        </td>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                          {r.reason || '—'}
                        </td>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                          {r.employee || '—'}
                        </td>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                          {r.user || '—'}
                        </td>
                        <td style={{ padding: 10, borderBottom: '1px solid #e5e7eb' }}>
                          <select
                            value={r.confirmed ? 'yes' : 'no'}
                            onChange={(e) => {
                              const yes = e.currentTarget.value === 'yes';
                              setVoidsDiscounts((rows) =>
                                rows.map((v) =>
                                  v.id === r.id
                                    ? { ...v, confirmed: yes, managerComment: yes ? '' : (v.managerComment || '') }
                                    : v
                                )
                              );
                            }}
                          >
                            <option value="no">No</option>
                            <option value="yes">Yes</option>
                          </select>
                        </td>
                      </tr>
                      {!r.confirmed && (
                        <tr>
                          <td colSpan={7} style={{ padding: 10, background: '#fff7f7' }}>
                            <div style={{ display: 'grid', gap: 6 }}>
                              <label style={{ fontWeight: 600 }}>
                                Manager Comment (required because you selected “No”)
                              </label>
                              <textarea
                                value={r.managerComment || ''}
                                onChange={(e) =>
                                  setVoidsDiscounts((rows) =>
                                    rows.map((v) =>
                                      v.id === r.id ? { ...v, managerComment: e.target.value } : v
                                    )
                                  )
                                }
                                placeholder="Brief reason…"
                                rows={2}
                                style={{ width: '100%' }}
                              />
                              {!(r.managerComment || '').trim() && (
                                <div style={{ color: '#ef4444', fontSize: 12 }}>
                                  Add a comment to proceed to the next section.
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>

              <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 8 }}>
                <div style={{ fontSize: 12, color: '#6b7280' }}>
                  Rows: {(voidsDiscounts || []).length} | Total: <b>{currency(voidsTotal)}</b>
                </div>
                <button
                  onClick={() => {
                    const b = generateExcelBlob();
                    const a = document.createElement('a');
                    a.href = URL.createObjectURL(b);
                    a.download = (b as any).name || 'Voids_Discounts.xlsx';
                    a.click();
                  }}
                >
                  Export Excel
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ISSUES */}
        {activeTab === 'issues' && (
          <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
            <div
              style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}
            >
              Daily Recap & Issues
            </div>
            <div style={{ padding: '16px 20px' }}>
              {/* ISSUES */}
{activeTab === 'issues' && (
  <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
    <div
      style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}
    >
      Daily Recap & Issues
    </div>

    <div style={{ padding: '16px 20px' }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14 }}>
        <div>
          <label>86 items from the bar?</label>
          <select
            value={issues.bar86 ? 'Yes' : 'No'}
            onChange={(e) => setIssues({ ...issues, bar86: e.target.value === 'Yes' })}
            style={{ width: '100%', height: 42 }}
          >
            <option>No</option>
            <option>Yes</option>
          </select>
          <div style={{ marginTop: 8 }}>
            <input
              placeholder="Has this affected sales?"
              value={issues.bar86Impact}
              onChange={(e) => setIssues({ ...issues, bar86Impact: e.target.value })}
              style={{ width: '100%', height: 42 }}
            />
          </div>
        </div>

        <div>
          <label>86 items from the kitchen?</label>
          <select
            value={issues.kitchen86 ? 'Yes' : 'No'}
            onChange={(e) => setIssues({ ...issues, kitchen86: e.target.value === 'Yes' })}
            style={{ width: '100%', height: 42 }}
          >
            <option>No</option>
            <option>Yes</option>
          </select>
          <div style={{ marginTop: 8 }}>
            <input
              placeholder="Has this affected sales?"
              value={issues.kitchen86Impact}
              onChange={(e) => setIssues({ ...issues, kitchen86Impact: e.target.value })}
              style={{ width: '100%', height: 42 }}
            />
          </div>
        </div>
      </div>

      <div style={{ display: 'grid', gap: 14, marginTop: 14 }}>
        <div>
          <label>Guest Feedback</label>
          <textarea
            value={issues.guestFeedback}
            onChange={(e) => setIssues({ ...issues, guestFeedback: e.target.value })}
            style={{ width: '100%', minHeight: 90 }}
          />
        </div>
        <div>
          <label>Issues & Challenges</label>
          <textarea
            value={issues.issuesChallenges}
            onChange={(e) => setIssues({ ...issues, issuesChallenges: e.target.value })}
            style={{ width: '100%', minHeight: 90 }}
          />
        </div>
        <div>
          <label>Team Performance</label>
          <textarea
            value={issues.teamPerformance}
            onChange={(e) => setIssues({ ...issues, teamPerformance: e.target.value })}
            style={{ width: '100%', minHeight: 90 }}
          />
        </div>
        <div>
          <label>Weather Impact</label>
          <textarea
            value={issues.weatherImpact}
            onChange={(e) => setIssues({ ...issues, weatherImpact: e.target.value })}
            style={{ width: '100%', minHeight: 90 }}
          />
        </div>
        <div>
          <label>Team Tasting — Dish & Feedback</label>
          <textarea
            value={issues.dishTasted}
            onChange={(e) => setIssues({ ...issues, dishTasted: e.target.value })}
            style={{ width: '100%', minHeight: 90 }}
          />
        </div>
      </div>

      <div
        style={{
          display: 'grid',
          gap: 14,
          gridTemplateColumns: 'repeat(3,1fr)',
          marginTop: 14,
        }}
      >
        <div>
          <label>Fire Drill (date)</label>
          <input
            type="date"
            value={issues.fireDrillDate}
            onChange={(e) => setIssues({ ...issues, fireDrillDate: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>Monthly Emergency Light Checks (date)</label>
          <input
            type="date"
            value={issues.emergencyLightDate}
            onChange={(e) => setIssues({ ...issues, emergencyLightDate: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>Weekly Fire Alarm Check (date)</label>
          <input
            type="date"
            value={issues.weeklyFireAlarmDate}
            onChange={(e) => setIssues({ ...issues, weeklyFireAlarmDate: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
      </div>

      <div style={{ display: 'grid', gap: 14, marginTop: 14 }}>
        <div>
          <label>FLOW Training — Missing Team Members</label>
          <textarea
            value={issues.flowTrainingMissing}
            onChange={(e) => setIssues({ ...issues, flowTrainingMissing: e.target.value })}
            style={{ width: '100%', minHeight: 80 }}
          />
        </div>
        <div>
          <label>Full team with correct uniform?</label>
          <select
            value={issues.uniformOk ? 'Yes' : 'No'}
            onChange={(e) => setIssues({ ...issues, uniformOk: e.target.value === 'Yes' })}
            style={{ width: '100%', height: 42 }}
          >
            <option>Yes</option>
            <option>No</option>
          </select>
        </div>
        {!issues.uniformOk && (
          <div>
            <label>Uniform issues</label>
            <textarea
              value={issues.uniformMissing}
              onChange={(e) => setIssues({ ...issues, uniformMissing: e.target.value })}
              style={{ width: '100%', minHeight: 80 }}
            />
          </div>
        )}
      </div>

      <div
        style={{
          display: 'grid',
          gap: 14,
          gridTemplateColumns: 'repeat(3,1fr)',
          marginTop: 14,
        }}
      >
        <div>
          <label>Tablets</label>
          <input
            inputMode="numeric"
            value={issues.tabletsCount}
            onChange={(e) => setIssues({ ...issues, tabletsCount: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>Cable Chargers</label>
          <input
            inputMode="numeric"
            value={issues.cableChargersCount}
            onChange={(e) => setIssues({ ...issues, cableChargersCount: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>Charger Bricks</label>
          <input
            inputMode="numeric"
            value={issues.chargerBricksCount}
            onChange={(e) => setIssues({ ...issues, chargerBricksCount: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>PDQs</label>
          <input
            inputMode="numeric"
            value={issues.pdqsCount}
            onChange={(e) => setIssues({ ...issues, pdqsCount: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>Food Bibles</label>
          <input
            inputMode="numeric"
            value={issues.foodBiblesCount}
            onChange={(e) => setIssues({ ...issues, foodBiblesCount: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
        <div>
          <label>Allergens Folders</label>
          <input
            inputMode="numeric"
            value={issues.allergensFoldersCount}
            onChange={(e) => setIssues({ ...issues, allergensFoldersCount: e.target.value })}
            style={{ width: '100%', height: 42 }}
          />
        </div>
      </div>

      <div style={{ display: 'grid', gap: 14, marginTop: 14 }}>
        <div>
          <label>Action Points (follow-ups)</label>
          <textarea
            value={issues.actionPoints}
            onChange={(e) => setIssues({ ...issues, actionPoints: e.target.value })}
            style={{ width: '100%', minHeight: 110 }}
          />
        </div>
      </div>
    </div>
  </div>
)}

            </div>
          </div>
        )}

        {/* SEND (unchanged) */}
        {activeTab === 'send' && (
          <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
            {/* SEND */}
{activeTab === 'send' && (
  <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 14 }}>
    <div
      style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}
    >
      Review & Send
    </div>

    <div style={{ padding: '16px 20px' }}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14 }}>
        <div>
          <label>From</label>
          <div style={{ border: '1px solid #e5e7eb', padding: 10, borderRadius: 10 }}>
            {SENDER_NAME} &lt;Apps Script owner account&gt;
          </div>
        </div>
        <div>
          <label>To</label>
          <div style={{ border: '1px solid #e5e7eb', padding: 10, borderRadius: 10 }}>
            {finalRecipients.to.join(', ') || '—'}
          </div>
        </div>
      </div>

      {config.defaults.cc.length > 0 && (
        <div style={{ marginTop: 10 }}>
          <label>CC</label>
          <div style={{ border: '1px solid #e5e7eb', padding: 10, borderRadius: 10 }}>
            {config.defaults.cc.join(', ')}
          </div>
        </div>
      )}
      {config.defaults.bcc.length > 0 && (
        <div style={{ marginTop: 10 }}>
          <label>BCC</label>
          <div style={{ border: '1px solid #e5e7eb', padding: 10, borderRadius: 10 }}>
            {config.defaults.bcc.join(', ')}
          </div>
        </div>
      )}
      {sendMsg && (
        <div style={{ marginTop: 10 }}>
          <div style={{ border: '1px solid #e5e7eb', padding: 10, borderRadius: 10 }}>
            {sendMsg}
          </div>
        </div>
      )}

      <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 12 }}>
        <button
          onClick={async () => {
            const subject = `${locationName} close`;
            const text = [
              'Hi team,',
              '',
              'Closing report attached.',
              '',
              'Sales Recap:',
              `• Gross: ${currency(sales.grossSales)} | Net+Serv: ${currency(sales.netServ)}`,
              `• Cash: ${currency(sales.cash)} | Card: ${currency(sales.card)} | Covers: ${sales.covers || 0} | Deliveroo: ${currency(sales.deliveroo)}`,
              `• Service Charge: ${currency(serviceChargeCalc)} | Additional Tips: ${currency(manualTipsCalc)} | Total: ${currency(tipsTotalCalc)}`,
              '',
              'Voids & Discounts:',
              `• Items: ${(voidsDiscounts || []).length} | Total: ${currency(voidsTotal)}`,
            ].join('\n');
            await navigator.clipboard.writeText(`Subject: ${subject}\n\n${text}`);
            setSendMsg('Email draft copied to clipboard.');
          }}
        >
          Copy email draft
        </button>
        <button onClick={sendEmailNow} disabled={sending || !canProceedVoids}>
          {sending ? 'Sending…' : 'Send email now'}
        </button>
      </div>

      <div style={{ fontSize: 12, color: '#6b7280', marginTop: 8 }}>
        Make sure your Apps Script is deployed as Web App (Execute as Me, Anyone with link).
      </div>
    </div>
  </div>
)}

          </div>
        )}

        <footer
          style={{
            fontSize: 12,
            color: '#6b7280',
            padding: '16px 0 40px',
            textAlign: 'center',
          }}
        >
          App by Honeyscuklesdesign for La Mia Mamma LTD.
        </footer>
      </div>
    </div>
  );
}

