import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Alertas Procesos",
  description: "Carga de Excel y vista previa de alertas Administrativo/Penal"
};

export default function RootLayout({
  children
}: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
