import { createHmac } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";

const userId = process.env.FALABELLA_USER_ID?.trim();
const apiKey = process.env.FALABELLA_API_KEY?.trim();
const baseUrl = (process.env.FALABELLA_BASE_URL || "https://sellercenter-api.falabella.com").replace(/\/$/, "");

if (!userId || !apiKey) {
  throw new Error("Faltan FALABELLA_USER_ID o FALABELLA_API_KEY en GitHub Secrets.");
}

const encodeRfc3986 = (value) => encodeURIComponent(String(value)).replace(/[!'()*]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`);

function signedQuery(parameters) {
  const canonical = Object.entries(parameters)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${encodeRfc3986(name)}=${encodeRfc3986(value)}`)
    .join("&");
  const signature = createHmac("sha256", apiKey).update(canonical).digest("hex");
  return `${canonical}&Signature=${signature}`;
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}

function locateProducts(payload) {
  const candidates = [
    payload?.SuccessResponse?.Body?.Products?.Product,
    payload?.SuccessResponse?.Body?.Product,
    payload?.Body?.Products?.Product,
    payload?.Products?.Product,
    payload?.Products,
  ];
  return asArray(candidates.find((candidate) => candidate != null));
}

function falabellaUnit(product) {
  const units = asArray(product?.BusinessUnits?.BusinessUnit ?? product?.BusinessUnit);
  return units.find((unit) => String(unit?.OperatorCode ?? "").toLowerCase() === "facl") ?? units[0] ?? {};
}

async function fetchPage(offset, limit) {
  const parameters = {
    Action: "GetProducts",
    Filter: "all",
    Format: "JSON",
    Limit: String(limit),
    Offset: String(offset),
    Timestamp: new Date().toISOString(),
    UserID: userId,
    Version: "1.0",
  };
  const response = await fetch(`${baseUrl}/?${signedQuery(parameters)}`, {
    headers: {
      Accept: "application/json",
    },
  });
  const text = await response.text();
  if (!response.ok) throw new Error(`Falabella respondió HTTP ${response.status}: ${text.slice(0, 400)}`);

  let payload;
  try { payload = JSON.parse(text); }
  catch { throw new Error(`Falabella no devolvió JSON: ${text.slice(0, 400)}`); }

  const apiError = payload?.ErrorResponse ?? payload?.error ?? payload?.Error;
  if (apiError) throw new Error(`Error de Falabella: ${JSON.stringify(apiError).slice(0, 600)}`);
  return locateProducts(payload);
}

const limit = 1000;
const allProducts = [];
for (let offset = 0; ; offset += limit) {
  const page = await fetchPage(offset, limit);
  allProducts.push(...page);
  console.log(`Productos recibidos: ${allProducts.length}`);
  if (page.length < limit) break;
  if (offset >= 999000) throw new Error("Se detuvo la paginación por seguridad.");
}

const products = allProducts.map((product) => {
  const unit = falabellaUnit(product);
  return {
    "SKU seller": String(product?.SellerSku ?? ""),
    "ShopSku Falabella": String(product?.ShopSku ?? ""),
    "Producto": String(product?.Name ?? ""),
    "Marca": String(product?.Brand ?? ""),
    "Estado FACL": String(unit?.Status ?? ""),
    "Stock FACL": String(unit?.Stock ?? ""),
  };
}).filter((product) => product["SKU seller"] || product["ShopSku Falabella"]);

products.sort((a, b) => a["SKU seller"].localeCompare(b["SKU seller"], "es", { numeric: true }));

const masterUrl = new URL("../../public/maestro.json", import.meta.url);
let previousProducts = [];
try {
  const previous = JSON.parse(await readFile(masterUrl, "utf8"));
  previousProducts = Array.isArray(previous?.products) ? previous.products : [];
} catch {
  // La primera sincronización no tiene un maestro anterior.
}

if (JSON.stringify(previousProducts) === JSON.stringify(products)) {
  console.log("El maestro no tuvo cambios; no se generará un despliegue nuevo.");
  process.exit(0);
}

const master = {
  updatedAt: new Date().toISOString(),
  source: "Falabella Seller Center",
  products,
};

await writeFile(masterUrl, `${JSON.stringify(master, null, 2)}\n`, "utf8");
console.log(`Maestro guardado con ${products.length} productos.`);
