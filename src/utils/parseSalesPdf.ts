// src/utils/parseSalesPdf.ts
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";

try {
  GlobalWorkerOptions.workerSrc = new URL(
    "pdfjs-dist/build/pdf.worker.mjs",
    import.meta.url
  ).toString();
} catch {}

export type SalesFormValues = {
  grossSales?: string; // Takings
  netSales?: string;   // Net + Service Charge
  tax?: string;        // VAT
  tips?: string;       // Service Charge
  cash?: string;       // Cash
  card?: string;       // Card
  deliveroo?: string;  // Deliveroo
  covers?: string;     // Guests
  discounts?: string;  // Discounts
  voids?: string;      // Voids
};

// numbers like "1,388.42 £", "(123.45)", "1.234,56", "-" (zero)
const NUM_TOKEN = /-?\d[\d.,]*(?:\s*£)?|-\s*(?:£)?/g;

function cleanNumberToken(raw?: string): string | undefined {
  if (!raw) return undefined;
  let t = raw.trim();
  if (t === "-" || t === "–" || t === "—") return "0";
  const neg = /^\(.*\)$/.test(t);
  t = t.replace(/[£$,()]/g, "").trim();
  if (/,\d{1,2}$/.test(t)) {
    t = t.replace(/\./g, "").replace(",", ".");
  } else {
    t = t.replace(/,/g, "");
  }
  if (t === "") t = "0";
  if (neg) t = "-" + t;
  return t;
}

const DAY_ALIASES: Record<string, RegExp[]> = {
  monday:    [/monday/i, /lunedi\b/i, /luned\u00EC/i],
  tuesday:   [/tuesday/i, /martedi\b/i, /marted\u00EC/i],
  wednesday: [/wednesday/i, /mercoledi\b/i, /mercoled\u00EC/i],
  thursday:  [/thursday/i, /giovedi\b/i, /gioved\u00EC/i],
  friday:    [/friday/i, /venerdi\b/i, /venerd\u00EC/i],
  saturday:  [/saturday/i, /sabato/i],
  sunday:    [/sunday/i, /domenica/i],
};

// Map a row’s 14 numeric columns to fields
function mapTokensToFields(tokens: string[]): Partial<SalesFormValues> | null {
  // Order after each weekday in your PDF:
  // 1 Takings | 2 Service Charge | 3 VAT | 4 Net Takings | 5 Net+Service |
  // 6 Cash | 7 Card | 8 Deliveroo | 9 Voucher/Gift | 10 Paid out |
  // 11 Guests | 12 Avg Spending | 13 Discounts | 14 Voids
  if (tokens.length < 14) return null;

  const [
    TAKINGS,
    SERVICE_CHG,
    VAT,
    /* NET_TAKINGS */,
    NET_PLUS_SERVICE,
    CASH,
    CARD,
    DELIVEROO,
    /* VOUCHER */,
    /* PAID_OUT */,
    GUESTS,
    /* AVG_SPEND */,
    DISCOUNTS,
    VOIDS,
  ] = tokens;

  return {
    grossSales: TAKINGS,
    tips: SERVICE_CHG,
    tax: VAT,
    netSales: NET_PLUS_SERVICE,
    cash: CASH,
    card: CARD,
    deliveroo: DELIVEROO,
    covers: GUESTS,      // guests == covers
    discounts: DISCOUNTS,
    voids: VOIDS,
  };
}

function parseFromDayRow(T: string, day?: string): Partial<SalesFormValues> | null {
  if (!day) return null;
  const regs = DAY_ALIASES[day.toLowerCase()];
  if (!regs) return null;

  for (const r of regs) {
    const m = T.match(r);
    if (!m || m.index == null) continue;
    // scan right of the weekday for the 14 numeric cells
    const slice = T.slice(m.index + m[0].length, m.index + m[0].length + 450);
    const tokens: string[] = [];
    for (const hit of slice.matchAll(NUM_TOKEN)) {
      const n = cleanNumberToken(hit[0]);
      if (n != null) tokens.push(n);
      if (tokens.length >= 14) break;
    }
    const out = mapTokensToFields(tokens);
    if (out) return out;
  }
  return null;
}

function parseFromTotalsRow(T: string): Partial<SalesFormValues> | null {
  const m = T.match(/\bTOTALS?\b/i);
  if (!m || m.index == null) return null;
  const tail = T.slice(m.index + m[0].length);
  const tokens: string[] = [];
  for (const hit of tail.matchAll(NUM_TOKEN)) {
    const n = cleanNumberToken(hit[0]);
    if (n != null) tokens.push(n);
    if (tokens.length >= 14) break;
  }
  return mapTokensToFields(tokens);
}

export async function extractPdfText(file: File): Promise<string> {
  const data = new Uint8Array(await file.arrayBuffer());
  const doc = await getDocument({ data }).promise;
  let text = "";
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    text += " " + content.items.map((it: any) => it.str).join(" ");
  }
  return text;
}

/** Prefer the requested weekday; fall back to TOTALS. */
export function parseSalesFromText(
  text: string,
  opts?: { day?: string }
): Partial<SalesFormValues> {
  const T = text.replace(/\s+/g, " ").replace(/[–—]/g, "-");

  const byDay = parseFromDayRow(T, opts?.day);
  if (byDay) {
    try { console.log("[Sales PDF] day mapping:", opts?.day, byDay); } catch {}
    return byDay;
  }

  const totals = parseFromTotalsRow(T);
  if (totals) {
    try { console.log("[Sales PDF] totals fallback:", totals); } catch {}
    return totals;
  }

  try { console.warn("[Sales PDF] no matches for day/totals."); } catch {}
  return {};
}

export function parseTotalsFromText(text: string) {
  // This assumes the column order used earlier in our parser:
  // Takings, Service Charge, VAT, Net Takings, Net+Service,
  // Cash, Card, Deliveroo, Voucher, Paid out, Guests, Avg Spending, Discounts, Voids
  const NUM = /-?\d[\d.,]*(?:\s*£)?|-\s*(?:£)?/g;
  const T = text.replace(/\s+/g, ' ').replace(/[–—]/g, '-');
  const m = T.match(/\bTOTALS?\b/i);
  if (!m || m.index == null) return {};
  const tail = T.slice(m.index + m[0].length);

  const clean = (s: string) => {
    let t = s.trim();
    if (t === '-' || t === '–' || t === '—') t = '0';
    const neg = /^\(.*\)$/.test(t);
    t = t.replace(/[£$,()]/g, '');
    if (/,\d{1,2}$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
    else t = t.replace(/,/g, '');
    if (!t) t = '0';
    if (neg) t = '-' + t;
    return t;
  };

  const toks: string[] = [];
  for (const hit of tail.matchAll(NUM)) {
    toks.push(clean(hit[0]));
    if (toks.length >= 14) break;
  }

  const [
    TAKINGS,
    SERVICE_CHG,
    /* VAT */, /* NET_TAKINGS */, NET_PLUS_SERVICE,
    CASH, CARD, DELIVEROO,
    /* VOUCHER */, /* PAID_OUT */, GUESTS,
    /* AVG_SPEND */, /* DISCOUNTS */, /* VOIDS */,
  ] = toks;

  return {
    grossSales: TAKINGS,
    tips: SERVICE_CHG,
    netSales: NET_PLUS_SERVICE,
    cash: CASH,
    card: CARD,
    deliveroo: DELIVEROO,
    covers: GUESTS,
  };
}
