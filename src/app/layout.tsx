import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Kimche Analyzer",
  description: "Análisis local de observaciones escolares desde CSV",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es" className="h-full antialiased">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
