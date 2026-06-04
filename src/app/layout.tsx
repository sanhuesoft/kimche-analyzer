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
      <body className="min-h-full flex flex-col bg-slate-100 text-slate-900">
        <div className="flex-1 flex flex-col">{children}</div>
        <footer className="w-full py-4 text-center text-xs text-slate-450 border-t border-slate-200/50 bg-slate-50">
          Desarrollado por{" "}
          <a
            href="https://www.fabiansanhueza.cl"
            target="_blank"
            rel="noopener noreferrer"
            className="font-semibold text-indigo-600 hover:text-indigo-800 transition hover:underline"
          >
            Fabián Sanhueza Vásquez
          </a>
        </footer>
      </body>
    </html>
  );
}
