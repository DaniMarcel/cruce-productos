# Cruce Fácil

Página para cruzar un Excel de pedidos con el maestro de productos de Falabella. El maestro se actualiza mediante GitHub Actions y Vercel solo aloja la página.

## Configuración inicial en GitHub

En el repositorio, abre **Settings → Secrets and variables → Actions → New repository secret** y crea estos secretos:

- `FALABELLA_USER_ID`: correo del usuario API de Seller Center.
- `FALABELLA_API_KEY`: clave API de ese usuario.
- `FALABELLA_BASE_URL`: `https://sellercenter-api.falabella.com`.

No agregues estos valores a archivos del repositorio ni a variables públicas de Vercel.

## Primera actualización del maestro

1. Abre la pestaña **Actions** del repositorio.
2. Selecciona **Actualizar maestro Falabella**.
3. Presiona **Run workflow**.
4. La acción consulta todos los productos, guarda `public/maestro.json` y publica el cambio en el repositorio.

Después de la primera ejecución, GitHub revisará los productos cada 5 minutos. También puede ejecutarse manualmente en cualquier momento. Si los datos no cambiaron, no se crea un commit ni un despliegue nuevo en Vercel.

El maestro contiene solamente:

- SKU seller
- ShopSku Falabella
- Producto
- Marca
- Estado FACL
- Stock FACL

> En repositorios privados, GitHub contabiliza los minutos de ejecución. Una revisión cada 5 minutos puede superar ampliamente la cuota mensual gratuita según el plan. En repositorios públicos, los runners estándar no consumen minutos facturables.

## Vercel

Importa el repositorio desde Vercel. La plataforma detectará Next.js automáticamente y no necesita las credenciales de Falabella ni una tarea programada.

Cada cambio de `public/maestro.json` provoca un nuevo despliegue con el maestro actualizado.

## Desarrollo local

```bash
npm install
npm run dev
```

La página estará disponible en `http://localhost:3000`.
