import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Joytree GST Manager",
  description: "Purchase/Sale GST tracking for Joytree Global — Indore, MP",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-slate-50 text-slate-900 antialiased">
        <header className="border-b border-slate-200 bg-white">
          <div className="mx-auto flex max-w-4xl items-center justify-between px-6 py-4">
            <a href="/" className="text-lg font-semibold tracking-tight">
              Joytree <span className="text-amber-600">GST</span> Manager
            </a>
            <nav className="flex gap-5 text-sm font-medium text-slate-600">
              <a href="/" className="hover:text-slate-900">Dashboard</a>
              <a href="/transactions/new" className="hover:text-slate-900">New Transaction</a>
            </nav>
          </div>
        </header>
        <main className="mx-auto max-w-4xl px-6 py-8">{children}</main>
      </body>
    </html>
  );
}
