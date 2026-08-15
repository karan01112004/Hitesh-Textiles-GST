"use client";

/**
 * app/page.tsx
 * Dashboard: pulls the current month's rollup from the `v_gst_summary`
 * view (see db/schema.sql) and gives quick access to transaction entry.
 * Vendor/Client list pages are Day 2-3 of the roadmap — not yet built here.
 */

import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabaseClient";

interface SummaryRow {
  transaction_type: "PURCHASE" | "SALE";
  period: string;
  gst_type: "INTRA_STATE" | "INTER_STATE";
  total_taxable: number;
  total_cgst: number;
  total_sgst: number;
  total_igst: number;
  total_invoice_value: number;
}

export default function DashboardPage() {
  const [rows, setRows] = useState<SummaryRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    const load = async () => {
      const { data, error } = await supabase
        .from("v_gst_summary")
        .select("*")
        .order("period", { ascending: false })
        .limit(12);

      if (error) setError(error.message);
      else setRows((data as SummaryRow[]) || []);
      setLoading(false);
    };
    load();
  }, []);

  const totals = rows.reduce(
    (acc, r) => {
      acc.taxable += Number(r.total_taxable);
      acc.gst += Number(r.total_cgst) + Number(r.total_sgst) + Number(r.total_igst);
      acc.invoiceValue += Number(r.total_invoice_value);
      return acc;
    },
    { taxable: 0, gst: 0, invoiceValue: 0 }
  );

  return (
    <div className="space-y-8">
      <section className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-semibold">Dashboard</h1>
          <p className="text-sm text-slate-500">Last 12 months, from the v_gst_summary view.</p>
        </div>
        <a
          href="/transactions/new"
          className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white"
        >
          + New Transaction
        </a>
      </section>

      {loading && <p className="text-sm text-slate-500">Loading…</p>}

      {error && (
        <p className="rounded-md bg-red-50 px-3 py-2 text-sm text-red-700">
          Could not load summary — check your Supabase env vars. ({error})
        </p>
      )}

      {!loading && !error && (
        <>
          <section className="grid grid-cols-3 gap-4">
            <StatCard label="Taxable Value" value={totals.taxable} />
            <StatCard label="Total GST" value={totals.gst} />
            <StatCard label="Invoice Value" value={totals.invoiceValue} />
          </section>

          <section>
            <h2 className="mb-2 text-sm font-semibold text-slate-700">Monthly breakdown</h2>
            {rows.length === 0 ? (
              <p className="rounded-md border border-dashed border-slate-300 px-4 py-6 text-center text-sm text-slate-500">
                No transactions yet — add your first one to see it here.
              </p>
            ) : (
              <div className="overflow-hidden rounded-md border border-slate-200">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-100 text-xs uppercase text-slate-500">
                    <tr>
                      <th className="px-3 py-2">Period</th>
                      <th className="px-3 py-2">Type</th>
                      <th className="px-3 py-2">GST Type</th>
                      <th className="px-3 py-2 text-right">Taxable</th>
                      <th className="px-3 py-2 text-right">Total GST</th>
                      <th className="px-3 py-2 text-right">Invoice Value</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {rows.map((r, i) => (
                      <tr key={i}>
                        <td className="px-3 py-2">
                          {new Date(r.period).toLocaleDateString("en-IN", { month: "short", year: "numeric" })}
                        </td>
                        <td className="px-3 py-2">{r.transaction_type}</td>
                        <td className="px-3 py-2">{r.gst_type === "INTRA_STATE" ? "CGST+SGST" : "IGST"}</td>
                        <td className="px-3 py-2 text-right">₹{Number(r.total_taxable).toFixed(2)}</td>
                        <td className="px-3 py-2 text-right">
                          ₹{(Number(r.total_cgst) + Number(r.total_sgst) + Number(r.total_igst)).toFixed(2)}
                        </td>
                        <td className="px-3 py-2 text-right font-medium">
                          ₹{Number(r.total_invoice_value).toFixed(2)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </>
      )}
    </div>
  );
}

function StatCard({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-md border border-slate-200 bg-white p-4">
      <div className="text-xs font-medium uppercase text-slate-500">{label}</div>
      <div className="mt-1 text-lg font-semibold text-slate-900">₹{value.toFixed(2)}</div>
    </div>
  );
}
