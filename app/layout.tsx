import type { Metadata } from "next";
import "./globals.css";

const productionHost = process.env.VERCEL_PROJECT_PRODUCTION_URL;
const siteUrl = productionHost ? `https://${productionHost}` : "http://localhost:3000";

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: "Cruce Fácil | Recupera nombres de productos",
  description: "Cruza tus pedidos con el maestro actualizado de Falabella y recupera automáticamente el nombre original de cada producto.",
  icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
  openGraph: {
    title: "Cruce Fácil",
    description: "Recupera el nombre original de cada producto.",
    type: "website",
    images: [{ url: "/og.png", width: 1536, height: 896, alt: "Cruce Fácil" }],
  },
  twitter: { card: "summary_large_image", images: ["/og.png"] },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="es"><body>{children}</body></html>;
}
