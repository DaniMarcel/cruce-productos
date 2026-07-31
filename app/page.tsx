"use client";

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type ExcelRow = Record<string, string>;
type ParsedSheet = { name: string; headers: string[]; rows: ExcelRow[] };
type ParsedBook = { fileName: string; sheets: ParsedSheet[] };
type MatchRow = { key: string; originalName: string; matched: boolean };
type Results = {
  rows: ExcelRow[];
  preview: MatchRow[];
  matched: number;
  unmatched: number;
  duplicates: number;
};

const fileAccept = ".xlsx,.xls,.csv";

function normalizeKey(value: unknown) {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLocaleLowerCase("es");
}

function parseFile(file: File): Promise<ParsedBook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("No se pudo leer el archivo."));
    reader.onload = () => {
      try {
        const workbook = XLSX.read(reader.result, { type: "array", cellDates: true });
        const sheets = workbook.SheetNames.map((name) => {
          const sheet = workbook.Sheets[name];
          const matrix = XLSX.utils.sheet_to_json<string[]>(sheet, {
            header: 1,
            defval: "",
            raw: false,
          });
          const firstNonEmpty = matrix.findIndex((row) => row.some((cell) => String(cell).trim()));
          if (firstNonEmpty < 0) return { name, headers: [], rows: [] };

          const used = new Map<string, number>();
          const headers = matrix[firstNonEmpty].map((cell, index) => {
            const base = String(cell).trim() || `Columna ${index + 1}`;
            const seen = used.get(base) ?? 0;
            used.set(base, seen + 1);
            return seen ? `${base} (${seen + 1})` : base;
          });
          const rows = matrix.slice(firstNonEmpty + 1)
            .filter((row) => row.some((cell) => String(cell).trim()))
            .map((row) => Object.fromEntries(headers.map((header, index) => [header, String(row[index] ?? "")])));
          return { name, headers, rows };
        });
        resolve({ fileName: file.name, sheets });
      } catch {
        reject(new Error("El archivo no parece ser un Excel o CSV válido."));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function findNameColumn(headers: string[]) {
  const priorities = ["nombre original", "nombre producto", "producto", "nombre", "descripcion", "descripción"];
  return priorities.map((candidate) => headers.find((h) => normalizeKey(h).includes(candidate))).find(Boolean) ?? "";
}

function UploadCard({
  eyebrow,
  title,
  hint,
  book,
  selectedSheet,
  onSheetChange,
  onFile,
}: {
  eyebrow: string;
  title: string;
  hint: string;
  book: ParsedBook | null;
  selectedSheet: number;
  onSheetChange: (index: number) => void;
  onFile: (file: File) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const sheet = book?.sheets[selectedSheet];

  const receive = (file?: File) => {
    if (file) onFile(file);
  };

  return (
    <section className={`upload-card ${book ? "is-ready" : ""}`}>
      <div className="step-line">
        <span className="step-dot" aria-hidden="true">{book ? "✓" : eyebrow}</span>
        <span>{book ? "Archivo listo" : `Paso ${eyebrow}`}</span>
      </div>
      <h2>{title}</h2>
      <p>{hint}</p>

      <div
        className={`dropzone ${dragging ? "is-dragging" : ""}`}
        onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(event: DragEvent<HTMLDivElement>) => {
          event.preventDefault();
          setDragging(false);
          receive(event.dataTransfer.files[0]);
        }}
      >
        <input
          ref={inputRef}
          type="file"
          accept={fileAccept}
          onChange={(event: ChangeEvent<HTMLInputElement>) => receive(event.target.files?.[0])}
          aria-label={`Cargar ${title}`}
        />
        <span className="file-icon" aria-hidden="true">▦</span>
        {book ? (
          <>
            <strong className="file-name">{book.fileName}</strong>
            <span>{sheet?.rows.length.toLocaleString("es-CL") ?? 0} filas · {sheet?.headers.length ?? 0} columnas</span>
            <button className="text-button" type="button" onClick={() => inputRef.current?.click()}>Cambiar archivo</button>
          </>
        ) : (
          <>
            <strong>Arrastra tu archivo aquí</strong>
            <span>o selecciónalo desde tu computador</span>
            <button className="secondary-button" type="button" onClick={() => inputRef.current?.click()}>Elegir archivo</button>
            <small>Excel (.xlsx, .xls) o CSV</small>
          </>
        )}
      </div>

      {book && book.sheets.length > 1 && (
        <label className="sheet-picker">
          Hoja a usar
          <select value={selectedSheet} onChange={(event) => onSheetChange(Number(event.target.value))}>
            {book.sheets.map((item, index) => <option value={index} key={item.name}>{item.name}</option>)}
          </select>
        </label>
      )}
    </section>
  );
}

function FieldSelect({ label, value, onChange, options, help }: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: string[];
  help?: string;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="">Selecciona una columna</option>
        {options.map((option) => <option key={option} value={option}>{option}</option>)}
      </select>
      {help && <small>{help}</small>}
    </label>
  );
}

export default function Home() {
  const [masterBook, setMasterBook] = useState<ParsedBook | null>(null);
  const [salesBook, setSalesBook] = useState<ParsedBook | null>(null);
  const [masterSheetIndex, setMasterSheetIndex] = useState(0);
  const [salesSheetIndex, setSalesSheetIndex] = useState(0);
  const [masterKey, setMasterKey] = useState("");
  const [salesKey, setSalesKey] = useState("");
  const [nameColumn, setNameColumn] = useState("");
  const [outputColumn, setOutputColumn] = useState("Nombre original");
  const [results, setResults] = useState<Results | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState<"master" | "sales" | "">("");

  const masterSheet = masterBook?.sheets[masterSheetIndex];
  const salesSheet = salesBook?.sheets[salesSheetIndex];

  useEffect(() => {
    if (!masterSheet || !salesSheet) return;
    const salesByNormalized = new Map(salesSheet.headers.map((header) => [normalizeKey(header), header]));
    const sharedMaster = masterSheet.headers.find((header) => salesByNormalized.has(normalizeKey(header)));
    const likelyKey = masterSheet.headers.find((header) => /(^|\s)(sku|key|llave|codigo|código|id)(\s|$)/i.test(header));
    const chosenMaster = likelyKey ?? sharedMaster ?? "";
    const chosenSales = chosenMaster
      ? salesByNormalized.get(normalizeKey(chosenMaster)) ?? ""
      : sharedMaster ? salesByNormalized.get(normalizeKey(sharedMaster)) ?? "" : "";
    setMasterKey((current) => masterSheet.headers.includes(current) ? current : chosenMaster);
    setSalesKey((current) => salesSheet.headers.includes(current) ? current : chosenSales);
    setNameColumn((current) => masterSheet.headers.includes(current) ? current : findNameColumn(masterSheet.headers));
    setResults(null);
  }, [masterSheet, salesSheet]);

  const canMatch = Boolean(masterSheet?.rows.length && salesSheet?.rows.length && masterKey && salesKey && nameColumn && outputColumn.trim());

  const load = async (kind: "master" | "sales", file: File) => {
    setBusy(kind);
    setError("");
    setResults(null);
    try {
      const parsed = await parseFile(file);
      if (!parsed.sheets.some((sheet) => sheet.headers.length)) throw new Error("No encontramos datos en ese archivo.");
      if (kind === "master") { setMasterBook(parsed); setMasterSheetIndex(0); }
      else { setSalesBook(parsed); setSalesSheetIndex(0); }
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "No se pudo cargar el archivo.");
    } finally {
      setBusy("");
    }
  };

  const crossFiles = () => {
    if (!masterSheet || !salesSheet || !canMatch) return;
    const lookup = new Map<string, string>();
    let duplicates = 0;
    masterSheet.rows.forEach((row) => {
      const key = normalizeKey(row[masterKey]);
      if (!key) return;
      if (lookup.has(key)) duplicates += 1;
      else lookup.set(key, row[nameColumn] ?? "");
    });

    let matched = 0;
    let unmatched = 0;
    const preview: MatchRow[] = [];
    const rows = salesSheet.rows.map((row) => {
      const displayedKey = row[salesKey] ?? "";
      const key = normalizeKey(displayedKey);
      const found = Boolean(key) && lookup.has(key);
      const originalName = found ? lookup.get(key) ?? "" : "";
      if (found) matched += 1;
      else unmatched += 1;
      if (preview.length < 8) preview.push({ key: displayedKey, originalName, matched: found });
      return { ...row, [outputColumn.trim()]: originalName };
    });
    setResults({ rows, preview, matched, unmatched, duplicates });
    requestAnimationFrame(() => document.getElementById("resultado")?.scrollIntoView({ behavior: "smooth", block: "start" }));
  };

  const download = () => {
    if (!results || !salesBook) return;
    const sheet = XLSX.utils.json_to_sheet(results.rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Ventas cruzadas");
    const base = salesBook.fileName.replace(/\.(xlsx|xls|csv)$/i, "");
    XLSX.writeFile(workbook, `${base}_con_nombres.xlsx`);
  };

  const matchRate = results ? Math.round((results.matched / Math.max(results.rows.length, 1)) * 100) : 0;

  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#inicio" aria-label="Cruce Fácil, inicio">
          <span className="brand-mark">C</span>
          <span>Cruce Fácil</span>
        </a>
        <span className="privacy-pill"><span>●</span> Tus archivos se procesan solo en este navegador</span>
      </header>

      <section className="hero" id="inicio">
        <div className="hero-copy">
          <span className="kicker">Cruce de productos</span>
          <h1>Recupera el nombre original de cada producto.</h1>
          <p>Carga tu maestro y tu archivo de compras, conecta la columna llave y descarga un Excel listo para usar.</p>
        </div>
        <div className="mini-flow" aria-label="Proceso en tres pasos">
          <div><b>1</b><span>Maestro</span></div><i>→</i><div><b>2</b><span>Compras</span></div><i>→</i><div><b>3</b><span>Resultado</span></div>
        </div>
      </section>

      <div className="workspace">
        {error && <div className="error-message" role="alert">{error}</div>}
        {busy && <div className="loading-bar" role="status">Leyendo tu archivo…</div>}

        <div className="upload-grid">
          <UploadCard
            eyebrow="1"
            title="Excel maestro"
            hint="El archivo que conserva el nombre original de tus productos."
            book={masterBook}
            selectedSheet={masterSheetIndex}
            onSheetChange={setMasterSheetIndex}
            onFile={(file) => load("master", file)}
          />
          <UploadCard
            eyebrow="2"
            title="Excel de compras"
            hint="El archivo generado después de que las personas compran."
            book={salesBook}
            selectedSheet={salesSheetIndex}
            onSheetChange={setSalesSheetIndex}
            onFile={(file) => load("sales", file)}
          />
        </div>

        {masterSheet && salesSheet && (
          <section className="config-card">
            <div className="section-heading">
              <div>
                <span className="kicker">Paso 3</span>
                <h2>Indica cómo conectar los archivos</h2>
                <p>La llave es el dato que identifica al mismo producto en ambos Excel, por ejemplo SKU o código.</p>
              </div>
              <span className="auto-badge">Sugerencias automáticas</span>
            </div>

            <div className="mapping-row">
              <FieldSelect label="Llave en el maestro" value={masterKey} onChange={(value) => { setMasterKey(value); setResults(null); }} options={masterSheet.headers} />
              <span className="link-symbol" aria-hidden="true">↔</span>
              <FieldSelect label="Llave en compras" value={salesKey} onChange={(value) => { setSalesKey(value); setResults(null); }} options={salesSheet.headers} />
            </div>
            <div className="mapping-row second-row">
              <FieldSelect label="Nombre original en el maestro" value={nameColumn} onChange={(value) => { setNameColumn(value); setResults(null); }} options={masterSheet.headers} />
              <span className="link-symbol muted" aria-hidden="true">→</span>
              <label className="field">
                <span>Nombre de la nueva columna</span>
                <input value={outputColumn} onChange={(event) => { setOutputColumn(event.target.value); setResults(null); }} placeholder="Nombre original" />
                <small>Se agregará al Excel de compras.</small>
              </label>
            </div>

            <div className="config-footer">
              <p><span aria-hidden="true">✓</span> El cruce ignora mayúsculas y espacios extra.</p>
              <button className="primary-button" type="button" disabled={!canMatch} onClick={crossFiles}>Cruzar archivos <span>→</span></button>
            </div>
          </section>
        )}

        {results && (
          <section className="result-card" id="resultado">
            <div className="result-top">
              <div className="success-icon" aria-hidden="true">✓</div>
              <div>
                <span className="kicker">Cruce completado</span>
                <h2>Tu archivo está listo</h2>
                <p>Agregamos la columna <strong>{outputColumn}</strong> sin modificar las demás columnas de compras.</p>
              </div>
              <button className="download-button" type="button" onClick={download}>↓ Descargar Excel</button>
            </div>

            <div className="stats-grid">
              <div><span>Coincidencias</span><strong>{results.matched.toLocaleString("es-CL")}</strong><small>{matchRate}% del archivo</small></div>
              <div><span>Sin coincidencia</span><strong className={results.unmatched ? "warning-text" : ""}>{results.unmatched.toLocaleString("es-CL")}</strong><small>Quedan con el nombre vacío</small></div>
              <div><span>Total procesado</span><strong>{results.rows.length.toLocaleString("es-CL")}</strong><small>Filas de compras</small></div>
            </div>

            {results.duplicates > 0 && <div className="notice">Encontramos {results.duplicates} llave(s) repetida(s) en el maestro. Se utilizó la primera aparición.</div>}

            <div className="preview-wrap">
              <div className="preview-title"><h3>Vista previa</h3><span>Primeras {results.preview.length} filas</span></div>
              <div className="table-scroll">
                <table>
                  <thead><tr><th>{salesKey}</th><th>{outputColumn}</th><th>Estado</th></tr></thead>
                  <tbody>
                    {results.preview.map((row, index) => (
                      <tr key={`${row.key}-${index}`}>
                        <td>{row.key || <em>Vacío</em>}</td>
                        <td>{row.originalName || <span className="empty-value">Sin coincidencia</span>}</td>
                        <td><span className={`status ${row.matched ? "matched" : "unmatched"}`}>{row.matched ? "Encontrado" : "Revisar"}</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}
      </div>

      <footer><span>Cruce Fácil</span><p>Simple, privado y sin subir información a internet.</p></footer>
    </main>
  );
}
