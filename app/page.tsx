"use client";

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type ExcelRow = Record<string, string>;
type ParsedSheet = { name: string; headers: string[]; rows: ExcelRow[] };
type ParsedBook = { fileName: string; sheets: ParsedSheet[] };
type MasterFile = { updatedAt: string | null; source: string; products: ExcelRow[] };
type MatchRow = { key: string; originalName: string; stock: string; matched: boolean };
type Results = { rows: ExcelRow[]; preview: MatchRow[]; matched: number; unmatched: number; duplicates: number };

const MASTER_KEYS = ["SKU seller", "ShopSku Falabella"];
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
          const matrix = XLSX.utils.sheet_to_json<string[]>(workbook.Sheets[name], { header: 1, defval: "", raw: false });
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

function MasterCard({ master, loading, error }: { master: MasterFile | null; loading: boolean; error: string }) {
  const ready = Boolean(master?.products.length);
  const updated = master?.updatedAt
    ? new Intl.DateTimeFormat("es-CL", { dateStyle: "medium", timeStyle: "short" }).format(new Date(master.updatedAt))
    : "Aún no sincronizado";

  return (
    <section className={`upload-card master-card ${ready ? "is-ready" : ""}`}>
      <div className="step-line">
        <span className="step-dot" aria-hidden="true">{ready ? "✓" : "1"}</span>
        <span>{ready ? "Maestro disponible" : "Paso 1 automático"}</span>
      </div>
      <h2>Maestro Falabella</h2>
      <p>Se actualiza automáticamente desde Seller Center mediante GitHub.</p>
      <div className={`master-status ${ready ? "ready" : "empty"}`}>
        <span className="file-icon" aria-hidden="true">◫</span>
        {loading ? (
          <><strong>Cargando maestro…</strong><span>Un momento, por favor.</span></>
        ) : ready ? (
          <>
            <strong>{master!.products.length.toLocaleString("es-CL")} productos disponibles</strong>
            <span>Última actualización: {updated}</span>
            <div className="master-fields"><span>SKU seller</span><span>ShopSku</span><span>Producto</span><span>Marca</span><span>Estado FACL</span><span>Stock FACL</span></div>
          </>
        ) : (
          <>
            <strong>El maestro todavía está vacío</strong>
            <span>{error || "Ejecuta por primera vez la acción “Actualizar maestro Falabella” en GitHub."}</span>
          </>
        )}
      </div>
    </section>
  );
}

function OrdersCard({ book, selectedSheet, onSheetChange, onFile }: {
  book: ParsedBook | null;
  selectedSheet: number;
  onSheetChange: (index: number) => void;
  onFile: (file: File) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const sheet = book?.sheets[selectedSheet];
  const receive = (file?: File) => file && onFile(file);

  return (
    <section className={`upload-card ${book ? "is-ready" : ""}`}>
      <div className="step-line"><span className="step-dot" aria-hidden="true">{book ? "✓" : "2"}</span><span>{book ? "Pedidos listos" : "Paso 2"}</span></div>
      <h2>Excel de pedidos</h2>
      <p>Carga el archivo generado cuando las personas realizan sus compras.</p>
      <div
        className={`dropzone ${dragging ? "is-dragging" : ""}`}
        onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(event: DragEvent<HTMLDivElement>) => { event.preventDefault(); setDragging(false); receive(event.dataTransfer.files[0]); }}
      >
        <input ref={inputRef} type="file" accept={fileAccept} onChange={(event: ChangeEvent<HTMLInputElement>) => receive(event.target.files?.[0])} aria-label="Cargar Excel de pedidos" />
        <span className="file-icon" aria-hidden="true">▦</span>
        {book ? (
          <><strong className="file-name">{book.fileName}</strong><span>{sheet?.rows.length.toLocaleString("es-CL") ?? 0} filas · {sheet?.headers.length ?? 0} columnas</span><button className="text-button" type="button" onClick={() => inputRef.current?.click()}>Cambiar archivo</button></>
        ) : (
          <><strong>Arrastra tu archivo aquí</strong><span>o selecciónalo desde tu computador</span><button className="secondary-button" type="button" onClick={() => inputRef.current?.click()}>Elegir archivo</button><small>Excel (.xlsx, .xls) o CSV</small></>
        )}
      </div>
      {book && book.sheets.length > 1 && (
        <label className="sheet-picker">Hoja a usar<select value={selectedSheet} onChange={(event) => onSheetChange(Number(event.target.value))}>{book.sheets.map((item, index) => <option value={index} key={item.name}>{item.name}</option>)}</select></label>
      )}
    </section>
  );
}

function FieldSelect({ label, value, onChange, options }: { label: string; value: string; onChange: (value: string) => void; options: string[] }) {
  return (
    <label className="field"><span>{label}</span><select value={value} onChange={(event) => onChange(event.target.value)}><option value="">Selecciona una columna</option>{options.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
  );
}

export default function Home() {
  const [master, setMaster] = useState<MasterFile | null>(null);
  const [masterLoading, setMasterLoading] = useState(true);
  const [masterError, setMasterError] = useState("");
  const [ordersBook, setOrdersBook] = useState<ParsedBook | null>(null);
  const [sheetIndex, setSheetIndex] = useState(0);
  const [masterKey, setMasterKey] = useState("ShopSku Falabella");
  const [ordersKey, setOrdersKey] = useState("");
  const [results, setResults] = useState<Results | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const ordersSheet = ordersBook?.sheets[sheetIndex];

  useEffect(() => {
    fetch(`/maestro.json?v=${Date.now()}`, { cache: "no-store" })
      .then((response) => { if (!response.ok) throw new Error(); return response.json(); })
      .then((data: MasterFile) => setMaster({ updatedAt: data.updatedAt ?? null, source: data.source ?? "Falabella Seller Center", products: Array.isArray(data.products) ? data.products : [] }))
      .catch(() => setMasterError("No se pudo leer el maestro generado por GitHub."))
      .finally(() => setMasterLoading(false));
  }, []);

  useEffect(() => {
    if (!ordersSheet) return;
    const likelyShopSku = ordersSheet.headers.find((header) => /shop.?sku|sku.*falabella|falabella.*sku/i.test(header));
    const likelySellerSku = ordersSheet.headers.find((header) => /seller.?sku|sku.*seller|sku.*vendedor/i.test(header));
    const genericSku = ordersSheet.headers.find((header) => /(^|\s|_)sku($|\s|_)/i.test(header));
    if (likelyShopSku) { setMasterKey("ShopSku Falabella"); setOrdersKey(likelyShopSku); }
    else if (likelySellerSku) { setMasterKey("SKU seller"); setOrdersKey(likelySellerSku); }
    else setOrdersKey(genericSku ?? "");
    setResults(null);
  }, [ordersSheet]);

  const canMatch = Boolean(master?.products.length && ordersSheet?.rows.length && masterKey && ordersKey);

  const loadOrders = async (file: File) => {
    setBusy(true); setError(""); setResults(null);
    try {
      const parsed = await parseFile(file);
      if (!parsed.sheets.some((sheet) => sheet.headers.length)) throw new Error("No encontramos datos en ese archivo.");
      setOrdersBook(parsed); setSheetIndex(0);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "No se pudo cargar el archivo.");
    } finally { setBusy(false); }
  };

  const crossFiles = () => {
    if (!master || !ordersSheet || !canMatch) return;
    const lookup = new Map<string, { name: string; brand: string; status: string; stock: string }>();
    let duplicates = 0;
    master.products.forEach((product) => {
      const key = normalizeKey(product[masterKey]);
      if (!key) return;
      if (lookup.has(key)) duplicates += 1;
      else lookup.set(key, {
        name: product.Producto ?? "",
        brand: product.Marca ?? "",
        status: product["Estado FACL"] ?? "",
        stock: product["Stock FACL"] ?? "",
      });
    });

    let matched = 0;
    let unmatched = 0;
    const preview: MatchRow[] = [];
    const rows = ordersSheet.rows.map((row) => {
      const displayedKey = row[ordersKey] ?? "";
      const key = normalizeKey(displayedKey);
      const found = Boolean(key) && lookup.has(key);
      const product = found ? lookup.get(key) : undefined;
      const originalName = product?.name ?? "";
      found ? matched += 1 : unmatched += 1;
      if (preview.length < 8) preview.push({ key: displayedKey, originalName, stock: product?.stock ?? "", matched: found });
      return {
        ...row,
        "Nombre original": originalName,
        "Marca": product?.brand ?? "",
        "Estado FACL": product?.status ?? "",
        "Stock FACL": product?.stock ?? "",
      };
    });
    setResults({ rows, preview, matched, unmatched, duplicates });
    requestAnimationFrame(() => document.getElementById("resultado")?.scrollIntoView({ behavior: "smooth", block: "start" }));
  };

  const download = () => {
    if (!results || !ordersBook) return;
    const sheet = XLSX.utils.json_to_sheet(results.rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Pedidos cruzados");
    XLSX.writeFile(workbook, `${ordersBook.fileName.replace(/\.(xlsx|xls|csv)$/i, "")}_con_nombres.xlsx`);
  };

  const matchRate = useMemo(() => results ? Math.round((results.matched / Math.max(results.rows.length, 1)) * 100) : 0, [results]);

  return (
    <main>
      <header className="topbar"><a className="brand" href="#inicio" aria-label="Cruce Fácil, inicio"><span className="brand-mark">C</span><span>Cruce Fácil</span></a><span className="privacy-pill"><span>●</span> Tus pedidos se procesan solo en este navegador</span></header>
      <div className="workspace">
        {error && <div className="error-message" role="alert">{error}</div>}
        {busy && <div className="loading-bar" role="status">Leyendo tu archivo…</div>}
        <div className="upload-grid"><MasterCard master={master} loading={masterLoading} error={masterError} /><OrdersCard book={ordersBook} selectedSheet={sheetIndex} onSheetChange={setSheetIndex} onFile={loadOrders} /></div>

        {ordersSheet && master?.products.length ? (
          <section className="config-card">
            <div className="section-heading"><div><span className="kicker">Paso 3</span><h2>Conecta la llave de los pedidos</h2><p>Elige qué identificador del maestro aparece en tu Excel de pedidos.</p></div><span className="auto-badge">Sugerencia automática</span></div>
            <div className="mapping-row"><FieldSelect label="Llave del maestro Falabella" value={masterKey} onChange={(value) => { setMasterKey(value); setResults(null); }} options={MASTER_KEYS} /><span className="link-symbol" aria-hidden="true">↔</span><FieldSelect label="Columna equivalente en pedidos" value={ordersKey} onChange={(value) => { setOrdersKey(value); setResults(null); }} options={ordersSheet.headers} /></div>
            <div className="config-footer"><p><span aria-hidden="true">✓</span> Se agregarán nombre, marca, estado y stock FACL sin alterar las demás columnas.</p><button className="primary-button" type="button" disabled={!canMatch} onClick={crossFiles}>Cruzar pedidos <span>→</span></button></div>
          </section>
        ) : ordersSheet && !masterLoading ? <div className="error-message master-help">Primero debes ejecutar la actualización del maestro en GitHub.</div> : null}

        {results && (
          <section className="result-card" id="resultado">
            <div className="result-top"><div className="success-icon" aria-hidden="true">✓</div><div><span className="kicker">Cruce completado</span><h2>Tu archivo está listo</h2><p>Agregamos <strong>nombre original, marca, estado y stock FACL</strong> sin modificar las demás columnas.</p></div><button className="download-button" type="button" onClick={download}>↓ Descargar Excel</button></div>
            <div className="stats-grid"><div><span>Coincidencias</span><strong>{results.matched.toLocaleString("es-CL")}</strong><small>{matchRate}% del archivo</small></div><div><span>Sin coincidencia</span><strong className={results.unmatched ? "warning-text" : ""}>{results.unmatched.toLocaleString("es-CL")}</strong><small>Quedan con el nombre vacío</small></div><div><span>Total procesado</span><strong>{results.rows.length.toLocaleString("es-CL")}</strong><small>Filas de pedidos</small></div></div>
            {results.duplicates > 0 && <div className="notice">Encontramos {results.duplicates} llave(s) repetida(s) en el maestro. Se utilizó la primera aparición.</div>}
            <div className="preview-wrap"><div className="preview-title"><h3>Vista previa</h3><span>Primeras {results.preview.length} filas</span></div><div className="table-scroll"><table><thead><tr><th>{ordersKey}</th><th>Nombre original</th><th>Stock FACL</th><th>Estado</th></tr></thead><tbody>{results.preview.map((row, index) => <tr key={`${row.key}-${index}`}><td>{row.key || <em>Vacío</em>}</td><td>{row.originalName || <span className="empty-value">Sin coincidencia</span>}</td><td>{row.matched ? row.stock || "0" : "—"}</td><td><span className={`status ${row.matched ? "matched" : "unmatched"}`}>{row.matched ? "Encontrado" : "Revisar"}</span></td></tr>)}</tbody></table></div></div>
          </section>
        )}
      </div>
      <footer><span>Cruce Fácil</span><p>Maestro automático · Pedidos procesados localmente</p></footer>
    </main>
  );
}
