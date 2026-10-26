"use client";

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type ExcelRow = Record<string, string>;
type ParsedSheet = { name: string; headers: string[]; rows: ExcelRow[] };
type ParsedBook = { fileName: string; sheets: ParsedSheet[] };
type MasterFile = { fileName: string; products: ExcelRow[]; flavorColumns: string[]; flavorCatalog: string[] };
type FlavorColumn = { source: string; output: string };
type MatchRow = { key: string; updatedName: string; flavors: ExcelRow; matched: boolean };
type DerivedColumns = { flavor: string; source: string; review: string };
type Results = { rows: ExcelRow[]; preview: MatchRow[]; flavorColumns: FlavorColumn[]; derivedColumns: DerivedColumns; matched: number; unmatched: number; duplicates: number };

const MASTER_KEYS = ["SKU seller", "ShopSku Falabella"];
const fileAccept = ".xlsx,.xls,.csv";
const FLAVOR_SOURCE_HEADERS = new Set([
  "variation",
  "productdata model",
  "productdata sabor",
  "productdata saborvariant",
  "saborvariant",
]);
const NON_SPECIFIC_FLAVORS = new Set([
  "otro sabor",
  "mix de sabores",
  "endulzado con stevia",
  "mix de verduras",
]);

function normalizeKey(value: unknown) {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLocaleLowerCase("es");
}

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("es")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function cleanCell(value: unknown) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function isValidFinalName(value: unknown) {
  const name = cleanCell(value);
  return Boolean(name) && !name.startsWith("#");
}

function isFlavorHeader(value: unknown) {
  const header = normalizeHeader(value);
  return FLAVOR_SOURCE_HEADERS.has(header) || /(sabor|flavou?r|taste)/i.test(header);
}

function isSpecificFlavor(value: unknown) {
  const flavor = normalizeHeader(value);
  return Boolean(flavor) && !NON_SPECIFIC_FLAVORS.has(flavor);
}

function canonicalFlavor(value: unknown) {
  const flavor = normalizeHeader(value);
  const aliases: Record<string, string> = {
    fresa: "frutilla",
    strawberry: "frutilla",
    "neutro sin sabor": "sin sabor",
    "cookies cream": "cookies and cream",
    "cookie and cream": "cookies and cream",
    "cookies n cream": "cookies and cream",
  };
  return aliases[flavor] ?? flavor;
}

function getProductValue(product: ExcelRow, normalizedHeader: string) {
  const entry = Object.entries(product).find(([header]) => normalizeHeader(header) === normalizedHeader);
  return cleanCell(entry?.[1]);
}

function extractKnownFlavor(value: unknown, flavorCatalog: string[]) {
  const text = ` ${normalizeHeader(value)} `;
  if (!text.trim()) return "";
  return flavorCatalog.find((flavor) => text.includes(` ${normalizeHeader(flavor)} `)) ?? "";
}

function resolveFlavor(product: ExcelRow, flavorCatalog: string[]) {
  const candidates = [
    { source: "SaborVariant", value: getProductValue(product, "saborvariant") },
    { source: "Variation", value: getProductValue(product, "variation") },
    { source: "ProductData.Sabor", value: getProductValue(product, "productdata sabor") },
    { source: "ProductData.SaborVariant", value: getProductValue(product, "productdata saborvariant") },
  ];
  const modelFlavor = extractKnownFlavor(getProductValue(product, "productdata model"), flavorCatalog);
  const nameFlavor = extractKnownFlavor(product.Producto, flavorCatalog);
  const lowConfidenceFlavor = candidates[2]?.value;
  const isLowConfidence = canonicalFlavor(lowConfidenceFlavor) === "sin sabor";
  const priority = [
    candidates[0],
    candidates[1],
    !isLowConfidence ? candidates[2] : null,
    candidates[3],
    modelFlavor ? { source: "ProductData.Model", value: modelFlavor } : null,
    nameFlavor ? { source: "Nombre final", value: nameFlavor } : null,
    isLowConfidence ? candidates[2] : null,
  ].filter((candidate): candidate is { source: string; value: string } => Boolean(candidate && isSpecificFlavor(candidate.value)));

  const selected = priority[0] ?? { source: "", value: "" };
  const comparisonValues = [
    ...candidates.filter((candidate) => isSpecificFlavor(candidate.value)).map((candidate) => candidate.value),
    modelFlavor,
    nameFlavor,
  ].filter(isSpecificFlavor);
  const distinctValues = new Set(comparisonValues.map(canonicalFlavor));
  const review = !selected.value
    ? "Sí: sin sabor identificado"
    : distinctValues.size > 1
      ? "Sí: valores distintos"
      : "No";

  return { flavor: selected.value, source: selected.source, review };
}

function createUniqueHeader(header: string, usedHeaders: Set<string>) {
  if (!usedHeaders.has(header)) {
    usedHeaders.add(header);
    return header;
  }

  const prefixed = `Maestro - ${header}`;
  let candidate = prefixed;
  let suffix = 2;
  while (usedHeaders.has(candidate)) {
    candidate = `${prefixed} (${suffix})`;
    suffix += 1;
  }
  usedHeaders.add(candidate);
  return candidate;
}

function suggestOrderKey(headers: string[], currentMasterKey: string) {
  const likelyShopSku = headers.find((header) => /shop.?sku|sku.*falabella|falabella.*sku/i.test(header));
  const likelySellerSku = headers.find((header) => /seller.?sku|sku.*seller|sku.*vendedor/i.test(header));
  const genericSku = headers.find((header) => /(^|\s|_)sku($|\s|_)/i.test(header));
  if (likelyShopSku) return { masterKey: "ShopSku Falabella", ordersKey: likelyShopSku };
  if (likelySellerSku) return { masterKey: "SKU seller", ordersKey: likelySellerSku };
  return { masterKey: currentMasterKey, ordersKey: genericSku ?? "" };
}

function parseMasterWorkbook(workbook: XLSX.WorkBook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
  const headerIndex = matrix.findIndex((row) => {
    const headers = row.map(normalizeHeader);
    const hasSellerSku = headers.some((header) => ["sku seller", "seller sku", "sellersku"].includes(header));
    const hasProductName = headers.some((header) => ["producto", "nombre producto", "name", "columna 1", "columna1"].includes(header));
    return hasSellerSku && hasProductName;
  });
  if (headerIndex < 0) throw new Error("El maestro no contiene SellerSku y una columna de nombre válida.");

  const originalHeaders = matrix[headerIndex].map((header, index) => cleanCell(header) || `Columna ${index + 1}`);
  const headers = originalHeaders.map(normalizeHeader);
  const findColumn = (...aliases: string[]) => headers.findIndex((header) => aliases.includes(header));
  const sellerSkuIndex = findColumn("sku seller", "seller sku", "sellersku");
  const falabellaSkuIndex = findColumn("shopsku falabella", "shop sku falabella", "sku falabella", "sku", "shopsku");
  const productIndex = findColumn("producto", "nombre producto", "name");
  const finalNameIndex = findColumn("columna 1", "columna1");

  if (sellerSkuIndex < 0 || falabellaSkuIndex < 0 || (productIndex < 0 && finalNameIndex < 0)) {
    throw new Error("El maestro no contiene todas las columnas necesarias para el cruce.");
  }

  const flavorColumns = originalHeaders
    .map((header, index) => ({ header, index }))
    .filter(({ header }) => isFlavorHeader(header));

  const products = matrix.slice(headerIndex + 1)
    .filter((row) => row.some((cell) => cleanCell(cell)))
    .map((row) => {
      const currentName = productIndex >= 0 ? cleanCell(row[productIndex]) : "";
      const finalName = finalNameIndex >= 0 ? cleanCell(row[finalNameIndex]) : "";
      const product: ExcelRow = {
        "SKU seller": cleanCell(row[sellerSkuIndex]),
        "ShopSku Falabella": cleanCell(row[falabellaSkuIndex]),
        Producto: isValidFinalName(finalName) ? finalName : currentName,
      };
      flavorColumns.forEach(({ header, index }) => { product[header] = cleanCell(row[index]); });
      return product;
    })
    .filter((row) => row["SKU seller"] || row["ShopSku Falabella"] || row.Producto);

  const flavorCatalogMap = new Map<string, string>();
  products.forEach((product) => {
    flavorColumns.forEach(({ header }) => {
      const value = product[header];
      if (!isSpecificFlavor(value) || normalizeHeader(header) === "productdata model") return;
      const canonical = canonicalFlavor(value);
      if (!flavorCatalogMap.has(canonical)) flavorCatalogMap.set(canonical, value);
    });
  });
  const flavorCatalog = [...flavorCatalogMap.values()]
    .sort((left, right) => normalizeHeader(right).length - normalizeHeader(left).length);

  return { products, flavorColumns: flavorColumns.map(({ header }) => header), flavorCatalog };
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

  return (
    <section className={`upload-card master-card ${ready ? "is-ready" : ""}`}>
      <div className="step-line">
        <span className="step-dot" aria-hidden="true">{ready ? "✓" : "1"}</span>
        <span>{ready ? "Maestro actualizado disponible" : "Paso 1"}</span>
      </div>
      <h2>Maestro Falabella</h2>
      <p>Maestro fijo con el nombre final definido para cada producto.</p>
      <div className={`master-status ${ready ? "ready" : "empty"}`}>
        <span className="file-icon" aria-hidden="true">◫</span>
        {loading ? (
          <><strong>Cargando maestro…</strong><span>Un momento, por favor.</span></>
        ) : ready ? (
          <>
            <strong>{master!.products.length.toLocaleString("es-CL")} productos disponibles</strong>
            <span>Archivo fijo: {master!.fileName}</span>
            <div className="master-fields"><span>SKU seller</span><span>ShopSku Falabella</span><span>Nombre final</span>{master!.flavorColumns.map((column) => <span key={column}>{column}</span>)}</div>
          </>
        ) : (
          <>
            <strong>El maestro todavía está vacío</strong>
            <span>{error || "No se pudo cargar el Excel maestro incluido en la página."}</span>
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
    fetch("/falabella-productos.xlsx")
      .then((response) => { if (!response.ok) throw new Error(); return response.arrayBuffer(); })
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: "array" });
        const parsedMaster = parseMasterWorkbook(workbook);
        setMaster({ fileName: "falabella-productos.xlsx", ...parsedMaster });
      })
      .catch(() => setMasterError("No se pudo leer el Excel maestro incluido en la página."))
      .finally(() => setMasterLoading(false));
  }, []);

  const canMatch = Boolean(master?.products.length && ordersSheet?.rows.length && masterKey && ordersKey);

  const loadOrders = async (file: File) => {
    setBusy(true); setError(""); setResults(null);
    try {
      const parsed = await parseFile(file);
      if (!parsed.sheets.some((sheet) => sheet.headers.length)) throw new Error("No encontramos datos en ese archivo.");
      const suggested = suggestOrderKey(parsed.sheets[0]?.headers ?? [], masterKey);
      setOrdersBook(parsed); setSheetIndex(0); setMasterKey(suggested.masterKey); setOrdersKey(suggested.ordersKey);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "No se pudo cargar el archivo.");
    } finally { setBusy(false); }
  };

  const selectOrdersSheet = (index: number) => {
    const suggested = suggestOrderKey(ordersBook?.sheets[index]?.headers ?? [], masterKey);
    setSheetIndex(index);
    setMasterKey(suggested.masterKey);
    setOrdersKey(suggested.ordersKey);
    setResults(null);
  };

  const crossFiles = () => {
    if (!master || !ordersSheet || !canMatch) return;
    const lookup = new Map<string, ExcelRow>();
    let duplicates = 0;
    master.products.forEach((product) => {
      const key = normalizeKey(product[masterKey]);
      if (!key) return;
      if (lookup.has(key)) duplicates += 1;
      else lookup.set(key, product);
    });

    const usedHeaders = new Set(ordersSheet.headers);
    usedHeaders.add("Nombre actualizado");
    const flavorColumns = master.flavorColumns.map((source) => ({
      source,
      output: createUniqueHeader(source, usedHeaders),
    }));
    const derivedColumns: DerivedColumns = {
      flavor: createUniqueHeader("Sabor consolidado", usedHeaders),
      source: createUniqueHeader("Fuente del sabor", usedHeaders),
      review: createUniqueHeader("Revisar sabor", usedHeaders),
    };

    let matched = 0;
    let unmatched = 0;
    const preview: MatchRow[] = [];
    const rows = ordersSheet.rows.map((row) => {
      const displayedKey = row[ordersKey] ?? "";
      const key = normalizeKey(displayedKey);
      const found = Boolean(key) && lookup.has(key);
      const product = found ? lookup.get(key) : undefined;
      const updatedName = product?.Producto ?? "";
      const flavors = Object.fromEntries(flavorColumns.map(({ source, output }) => [output, product?.[source] ?? ""]));
      const resolved = product
        ? resolveFlavor(product, master.flavorCatalog)
        : { flavor: "", source: "", review: "Sí: sin coincidencia" };
      flavors[derivedColumns.flavor] = resolved.flavor;
      flavors[derivedColumns.source] = resolved.source;
      flavors[derivedColumns.review] = resolved.review;
      if (found) matched += 1;
      else unmatched += 1;
      if (preview.length < 8) preview.push({ key: displayedKey, updatedName, flavors, matched: found });
      return { ...row, "Nombre actualizado": updatedName, ...flavors };
    });
    setResults({ rows, preview, flavorColumns, derivedColumns, matched, unmatched, duplicates });
    requestAnimationFrame(() => document.getElementById("resultado")?.scrollIntoView({ behavior: "smooth", block: "start" }));
  };

  const download = () => {
    if (!results || !ordersBook) return;
    const sheet = XLSX.utils.json_to_sheet(results.rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Pedidos cruzados");
    XLSX.writeFile(workbook, `${ordersBook.fileName.replace(/\.(xlsx|xls|csv)$/i, "")}_con_nombres_y_sabores.xlsx`);
  };

  const matchRate = useMemo(() => results ? Math.round((results.matched / Math.max(results.rows.length, 1)) * 100) : 0, [results]);

  return (
    <main>
      <header className="topbar"><a className="brand" href="#inicio" aria-label="Cruce Fácil, inicio"><span className="brand-mark">C</span><span>Cruce Fácil</span></a><span className="privacy-pill"><span>●</span> Tus pedidos se procesan solo en este navegador</span></header>
      <div className="workspace">
        {error && <div className="error-message" role="alert">{error}</div>}
        {busy && <div className="loading-bar" role="status">Leyendo tu archivo…</div>}
        <div className="upload-grid"><MasterCard master={master} loading={masterLoading} error={masterError} /><OrdersCard book={ordersBook} selectedSheet={sheetIndex} onSheetChange={selectOrdersSheet} onFile={loadOrders} /></div>

        {ordersSheet && master?.products.length ? (
          <section className="config-card">
            <div className="section-heading"><div><span className="kicker">Paso 3</span><h2>Conecta la llave de los pedidos</h2><p>Elige qué identificador del maestro aparece en tu Excel de pedidos.</p></div><span className="auto-badge">Sugerencia automática</span></div>
            <div className="mapping-row"><FieldSelect label="Llave del maestro Falabella" value={masterKey} onChange={(value) => { setMasterKey(value); setResults(null); }} options={MASTER_KEYS} /><span className="link-symbol" aria-hidden="true">↔</span><FieldSelect label="Columna equivalente en pedidos" value={ordersKey} onChange={(value) => { setOrdersKey(value); setResults(null); }} options={ordersSheet.headers} /></div>
            <div className="config-footer"><p><span aria-hidden="true">✓</span> Se conservarán todas las fuentes de sabor y se agregará un sabor consolidado con su origen y control de diferencias.</p><button className="primary-button" type="button" disabled={!canMatch} onClick={crossFiles}>Cruzar pedidos <span>→</span></button></div>
          </section>
        ) : ordersSheet && !masterLoading ? <div className="error-message master-help">No fue posible cargar el maestro actualizado incluido en la página.</div> : null}

        {results && (
          <section className="result-card" id="resultado">
            <div className="result-top"><div className="success-icon" aria-hidden="true">✓</div><div><span className="kicker">Cruce completado</span><h2>Tu archivo está listo</h2><p>Agregamos el <strong>nombre actualizado, {results.flavorColumns.length} fuentes de sabor y el sabor consolidado</strong> sin modificar las demás columnas.</p></div><button className="download-button" type="button" onClick={download}>↓ Descargar Excel</button></div>
            <div className="stats-grid"><div><span>Coincidencias</span><strong>{results.matched.toLocaleString("es-CL")}</strong><small>{matchRate}% del archivo</small></div><div><span>Sin coincidencia</span><strong className={results.unmatched ? "warning-text" : ""}>{results.unmatched.toLocaleString("es-CL")}</strong><small>Quedan con el nombre vacío</small></div><div><span>Total procesado</span><strong>{results.rows.length.toLocaleString("es-CL")}</strong><small>Filas de pedidos</small></div></div>
            {results.duplicates > 0 && <div className="notice">Encontramos {results.duplicates} llave(s) repetida(s) en el maestro. Se utilizó la primera aparición.</div>}
            <div className="preview-wrap"><div className="preview-title"><h3>Vista previa</h3><span>Primeras {results.preview.length} filas</span></div><div className="table-scroll"><table><thead><tr><th>{ordersKey}</th><th>Nombre actualizado</th>{results.flavorColumns.map(({ output }) => <th key={output}>{output}</th>)}<th>{results.derivedColumns.flavor}</th><th>{results.derivedColumns.source}</th><th>{results.derivedColumns.review}</th><th>Estado</th></tr></thead><tbody>{results.preview.map((row, index) => <tr key={`${row.key}-${index}`}><td>{row.key || <em>Vacío</em>}</td><td>{row.updatedName || <span className="empty-value">Sin coincidencia</span>}</td>{results.flavorColumns.map(({ output }) => <td key={output}>{row.flavors[output] || <span className="empty-value">—</span>}</td>)}<td>{row.flavors[results.derivedColumns.flavor] || <span className="empty-value">—</span>}</td><td>{row.flavors[results.derivedColumns.source] || <span className="empty-value">—</span>}</td><td>{row.flavors[results.derivedColumns.review]}</td><td><span className={`status ${row.matched ? "matched" : "unmatched"}`}>{row.matched ? "Encontrado" : "Revisar"}</span></td></tr>)}</tbody></table></div></div>
          </section>
        )}
      </div>
      <footer><span>Cruce Fácil</span><p>Maestro actualizado fijo · Pedidos procesados localmente</p></footer>
    </main>
  );
}
