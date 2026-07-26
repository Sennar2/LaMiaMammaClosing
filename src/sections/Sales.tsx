// src/sections/Sales.tsx
import React, { useMemo, useState } from "react";
import {
  extractPdfText,
  parseSalesFromText,
  SalesFormValues,
} from "../utils/parseSalesPdf";

type Flags = { [K in keyof SalesFormValues]?: boolean };

const FIELD_LIST: Array<keyof SalesFormValues> = [
  "grossSales",
  "netSales",
  "tax",
  "serviceCharge",
  "cash",
  "card",
  "discounts",
  "voids",
];

export default function SalesSection() {
  const [values, setValues] = useState<SalesFormValues>({});
  const [manualEdits, setManualEdits] = useState<Flags>({});
  const [autoExtract, setAutoExtract] = useState<Partial<SalesFormValues> | null>(
    null
  );
  const [status, setStatus] = useState<"" | "ready" | "applied" | "error">("");

  function setField(k: keyof SalesFormValues, v: string) {
    setValues((prev) => ({ ...prev, [k]: v }));
    setManualEdits((prev) => ({ ...prev, [k]: true }));
  }

  async function onPdfSelected(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await extractPdfText(file);
      const parsed = parseSalesFromText(text);
      setAutoExtract(parsed);
      setStatus("ready");
    } catch (err) {
      console.error(err);
      setAutoExtract(null);
      setStatus("error");
      alert(
        "Could not read that PDF. Please try a different file or adjust patterns in src/utils/parseSalesPdf.ts"
      );
    }
  }

  function applyParsed(mode: "emptyOnly" | "overwriteAll") {
    if (!autoExtract) return;
    setValues((prev) => {
      const next = { ...prev };
      FIELD_LIST.forEach((f) => {
        const incoming = autoExtract[f];
        if (!incoming) return;
        if (mode === "overwriteAll") {
          next[f] = incoming;
        } else {
          if (!prev[f] || !manualEdits[f]) next[f] = incoming;
        }
      });
      return next;
    });
    setStatus("applied");
  }

  const hasAnyValue = useMemo(
    () => FIELD_LIST.some((f) => (values[f] ?? "") !== ""),
    [values]
  );

  return (
    <section className="card">
      <header className="card-header">
        <h2>Sales</h2>
        <div className="pdf-upload">
          <input
            type="file"
            accept="application/pdf"
            onChange={onPdfSelected}
            aria-label="Upload PDF"
          />
        </div>
      </header>

      {status === "ready" && autoExtract && (
        <div className="notice">
          <b>PDF parsed.</b>
          <button onClick={() => applyParsed("emptyOnly")}>
            Apply to empty fields
          </button>
          <button onClick={() => applyParsed("overwriteAll")}>
            Overwrite all
          </button>
        </div>
      )}

      {status === "applied" && (
        <div className="notice success">Values auto-filled from PDF. You can edit any field.</div>
      )}

      <div className="grid">
        {FIELD_LIST.map((f) => (
          <label key={f} className="row">
            <span>{labelFor(f)}</span>
            <input
              inputMode="decimal"
              value={values[f] ?? ""}
              onChange={(e) => setField(f, e.target.value)}
              placeholder="0.00"
            />
            {manualEdits[f] && <small className="chip">edited</small>}
          </label>
        ))}
      </div>

      {!hasAnyValue && (
        <p className="muted">
          Upload a PDF to auto-populate, or type values manually.
        </p>
      )}
    </section>
  );
}

function labelFor(k: keyof SalesFormValues) {
  switch (k) {
    case "grossSales":
      return "Gross Sales";
    case "netSales":
      return "Net Sales";
    case "tax":
      return "VAT/Tax";
    case "serviceCharge":
      return "Service Charge";
    case "cash":
      return "Cash";
    case "card":
      return "Card";
    case "discounts":
      return "Discounts";
    case "voids":
      return "Voids";
    default:
      return String(k);
  }
}
