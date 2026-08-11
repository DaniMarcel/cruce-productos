# Cruce Fácil

Herramienta interna para agregar a los pedidos el nombre actualizado de los productos vendidos en Falabella.

## Funcionamiento

- La página incluye `public/falabella-productos.xlsx` como maestro fijo.
- El maestro utiliza `Columna1` como nombre final del producto y `Name` como respaldo cuando `Columna1` está vacía o contiene un error.
- El usuario solo carga el Excel de pedidos.
- Puede cruzar usando `SKU seller` o `ShopSku Falabella`.
- El Excel descargado conserva todas las columnas de pedidos y agrega `Nombre actualizado`.
- Todo el cruce ocurre dentro del navegador.

No utiliza la API de Falabella, GitHub Actions, tareas programadas, secretos ni una base de datos.

## Reemplazar el maestro en el futuro

Reemplaza `public/falabella-productos.xlsx` por otro archivo con el mismo nombre y estas columnas:

- `SellerSku`
- `ShopSku`
- `Name`
- `Columna1`

Después confirma el cambio en Git y Vercel publicará la nueva versión.

## Desarrollo local

```bash
npm install
npm run dev
```

La página estará disponible en `http://localhost:3000`.
