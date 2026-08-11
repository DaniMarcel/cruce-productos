import type { Metadata } from "next";
import "./globals.css";

const productionHost = process.env.VERCEL_PROJECT_PRODUCTION_URL;
const siteUrl = productionHost ? `https://${productionHost}` : "http://localhost:3000";

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: "Cruce Fácil | Actualiza nombres de productos",
  description: "Cruza tus pedidos con el maestro de Falabella y agrega automáticamente el nombre actualizado de cada producto.",
  icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
  openGraph: {
    title: "Cruce Fácil",
    description: "Agrega el nombre actualizado de cada producto.",
    type: "website",
    images: [{ url: "/og.png", width: 1536, height: 896, alt: "Cruce Fácil" }],
  },
  twitter: { card: "summary_large_image", images: ["/og.png"] },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="es"><body>{children}</body></html>;
}
