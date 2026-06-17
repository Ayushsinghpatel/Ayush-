import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, Area, AreaChart, ComposedChart
} from "recharts";
import { mapPORows, mapReceiptRows, mapRateRows } from "./poDataMapper.js";

// ─── MOCK DATA ────────────────────────────────────────────────────────────────
const MOCK_PO = [
  { poDate:"2026-01-05", poNo:"PO-001", supplier:"Tata Steel", item:"HR Coil 2mm", qty:1000, rate:65, deliveryType:"Regular", gstType:"GST18", taxableAmt:65000, totalAmt:76700, preparedBy:"Rahul", requestedBy:"Priya", authorisedBy:"Director", sentToParty:true, receivedQty:900, status:"Partial", remarks:"", internalRemarks:"Shortage reported" },
  { poDate:"2026-01-12", poNo:"PO-002", supplier:"JSW Steel", item:"MS Plate 6mm", qty:500, rate:85, deliveryType:"Urgent Air", gstType:"GST18", taxableAmt:42500, totalAmt:50150, preparedBy:"Ankit", requestedBy:"Suresh", authorisedBy:"Director", sentToParty:true, receivedQty:500, status:"Complete", remarks:"", internalRemarks:"" },
  { poDate:"2026-01-20", poNo:"PO-003", supplier:"Sail India", item:"Angle Iron 50x50", qty:2000, rate:48, deliveryType:"Regular", gstType:"GST18", taxableAmt:96000, totalAmt:113280, preparedBy:"Rahul", requestedBy:"Meena", authorisedBy:"GM", sentToParty:true, receivedQty:2100, status:"Complete", remarks:"", internalRemarks:"Excess 100 pcs" },
  { poDate:"2026-02-03", poNo:"PO-004", supplier:"Tata Steel", item:"CR Coil 1.2mm", qty:800, rate:72, deliveryType:"Urgent Rail", gstType:"GST18", taxableAmt:57600, totalAmt:67968, preparedBy:"Ankit", requestedBy:"Priya", authorisedBy:"Director", sentToParty:true, receivedQty:0, status:"Pending", remarks:"", internalRemarks:"Supplier delay" },
  { poDate:"2026-02-14", poNo:"PO-005", supplier:"Hindalco", item:"Aluminium Sheet 3mm", qty:300, rate:220, deliveryType:"Regular", gstType:"GST18", taxableAmt:66000, totalAmt:77880, preparedBy:"Rahul", requestedBy:"Suresh", authorisedBy:"GM", sentToParty:true, receivedQty:280, status:"Partial", remarks:"", internalRemarks:"" },
  { poDate:"2026-02-22", poNo:"PO-006", supplier:"Jindal Poly", item:"HDPE Pipe 4inch", qty:1500, rate:35, deliveryType:"Regular", gstType:"GST12", taxableAmt:52500, totalAmt:58800, preparedBy:"Meena", requestedBy:"Meena", authorisedBy:"Director", sentToParty:false, receivedQty:0, status:"Pending", remarks:"", internalRemarks:"" },
  { poDate:"2026-03-01", poNo:"PO-007", supplier:"JSW Steel", item:"MS Plate 8mm", qty:600, rate:90, deliveryType:"Regular", gstType:"GST18", taxableAmt:54000, totalAmt:63720, preparedBy:"Ankit", requestedBy:"Priya", authorisedBy:"Director", sentToParty:true, receivedQty:600, status:"Complete", remarks:"", internalRemarks:"" },
  { poDate:"2026-03-10", poNo:"PO-008", supplier:"Sail India", item:"HR Coil 3mm", qty:1200, rate:62, deliveryType:"Urgent Air", gstType:"GST18", taxableAmt:74400, totalAmt:87792, preparedBy:"Rahul", requestedBy:"Suresh", authorisedBy:"GM", sentToParty:true, receivedQty:1100, status:"Partial", remarks:"", internalRemarks:"" },
  { poDate:"2026-03-18", poNo:"PO-009", supplier:"Hindalco", item:"Aluminium Extrusion", qty:200, rate:310, deliveryType:"Regular", gstType:"GST18", taxableAmt:62000, totalAmt:73160, preparedBy:"Meena", requestedBy:"Meena", authorisedBy:"Director", sentToParty:true, receivedQty:0, status:"Pending", remarks:"", internalRemarks:"Long lead time" },
  { poDate:"2026-04-02", poNo:"PO-010", supplier:"Tata Steel", item:"MS Flat 40x6", qty:3000, rate:44, deliveryType:"Regular", gstType:"GST18", taxableAmt:132000, totalAmt:155760, preparedBy:"Ankit", requestedBy:"Priya", authorisedBy:"Director", sentToParty:true, receivedQty:3000, status:"Complete", remarks:"", internalRemarks:"" },
  { poDate:"2026-04-15", poNo:"PO-011", supplier:"JSW Steel", item:"HR Coil 2mm", qty:900, rate:67, deliveryType:"Urgent Rail", gstType:"GST18", taxableAmt:60300, totalAmt:71154, preparedBy:"Rahul", requestedBy:"Suresh", authorisedBy:"GM", sentToParty:true, receivedQty:850, status:"Partial", remarks:"", internalRemarks:"" },
  { poDate:"2026-05-05", poNo:"PO-012", supplier:"Sail India", item:"Angle Iron 50x50", qty:1800, rate:50, deliveryType:"Regular", gstType:"GST18", taxableAmt:90000, totalAmt:106200, preparedBy:"Meena", requestedBy:"Meena", authorisedBy:"Director", sentToParty:false, receivedQty:0, status:"Pending", remarks:"Awaiting approval", internalRemarks:"" },
  { poDate:"2026-05-20", poNo:"PO-013", supplier:"Tata Steel", item:"CR Coil 1.5mm", qty:600, rate:80, deliveryType:"Regular", gstType:"GST18", taxableAmt:48000, totalAmt:56640, preparedBy:"Ankit", requestedBy:"Priya", authorisedBy:"GM", sentToParty:true, receivedQty:600, status:"Complete", remarks:"", internalRemarks:"" },
  { poDate:"2026-06-01", poNo:"PO-014", supplier:"Hindalco", item:"Aluminium Sheet 3mm", qty:450, rate:225, deliveryType:"Regular", gstType:"GST18", taxableAmt:101250, totalAmt:119475, preparedBy:"Rahul", requestedBy:"Suresh", authorisedBy:"Director", sentToParty:true, receivedQty:0, status:"Pending", remarks:"", internalRemarks:"" },
  { poDate:"2026-06-10", poNo:"PO-015", supplier:"Jindal Poly", item:"HDPE Pipe 6inch", qty:800, rate:55, deliveryType:"Urgent Air", gstType:"GST12", taxableAmt:44000, totalAmt:49280, preparedBy:"Meena", requestedBy:"Meena", authorisedBy:"Director", sentToParty:true, receivedQty:800, status:"Complete", remarks:"", internalRemarks:"" },
];

const MOCK_RECEIPTS = [
  { date:"2026-01-18", poNo:"PO-001", voucherNo:"GRN-001", supplier:"Tata Steel", item:"HR Coil 2mm", receivedQty:900, rate:65 },
  { date:"2026-01-25", poNo:"PO-002", voucherNo:"GRN-002", supplier:"JSW Steel", item:"MS Plate 6mm", receivedQty:500, rate:85 },
  { date:"2026-01-28", poNo:"PO-003", voucherNo:"GRN-003", supplier:"Sail India", item:"Angle Iron 50x50", receivedQty:2100, rate:48 },
  { date:"2026-02-20", poNo:"PO-005", voucherNo:"GRN-004", supplier:"Hindalco", item:"Aluminium Sheet 3mm", receivedQty:280, rate:220 },
  { date:"2026-03-12", poNo:"PO-007", voucherNo:"GRN-005", supplier:"JSW Steel", item:"MS Plate 8mm", receivedQty:600, rate:90 },
  { date:"2026-03-22", poNo:"PO-008", voucherNo:"GRN-006", supplier:"Sail India", item:"HR Coil 3mm", receivedQty:1100, rate:62 },
  { date:"2026-04-10", poNo:"PO-010", voucherNo:"GRN-007", supplier:"Tata Steel", item:"MS Flat 40x6", receivedQty:3000, rate:44 },
  { date:"2026-04-28", poNo:"PO-011", voucherNo:"GRN-008", supplier:"JSW Steel", item:"HR Coil 2mm", receivedQty:850, rate:67 },
  { date:"2026-05-28", poNo:"PO-013", voucherNo:"GRN-009", supplier:"Tata Steel", item:"CR Coil 1.5mm", receivedQty:600, rate:80 },
  { date:"2026-06-14", poNo:"PO-015", voucherNo:"GRN-010", supplier:"Jindal Poly", item:"HDPE Pipe 6inch", receivedQty:800, rate:55 },
];

const MOCK_RATES = [
  { date:"2025-10-01", supplier:"Tata Steel", item:"HR Coil 2mm", newRate:60, oldRate:58 },
  { date:"2025-12-01", supplier:"Tata Steel", item:"HR Coil 2mm", newRate:63, oldRate:60 },
  { date:"2026-01-05", supplier:"Tata Steel", item:"HR Coil 2mm", newRate:65, oldRate:63 },
  { date:"2026-04-15", supplier:"Tata Steel", item:"HR Coil 2mm", newRate:67, oldRate:65 },
  { date:"2025-11-01", supplier:"JSW Steel", item:"MS Plate 6mm", newRate:80, oldRate:76 },
  { date:"2026-01-12", supplier:"JSW Steel", item:"MS Plate 6mm", newRate:85, oldRate:80 },
  { date:"2026-03-01", supplier:"JSW Steel", item:"MS Plate 8mm", newRate:90, oldRate:86 },
  { date:"2025-09-01", supplier:"Hindalco", item:"Aluminium Sheet 3mm", newRate:210, oldRate:200 },
  { date:"2026-02-14", supplier:"Hindalco", item:"Aluminium Sheet 3mm", newRate:220, oldRate:210 },
  { date:"2026-06-01", supplier:"Hindalco", item:"Aluminium Sheet 3mm", newRate:225, oldRate:220 },
  { date:"2025-10-15", supplier:"Sail India", item:"Angle Iron 50x50", newRate:45, oldRate:42 },
  { date:"2026-01-20", supplier:"Sail India", item:"Angle Iron 50x50", newRate:48, oldRate:45 },
  { date:"2026-05-12", supplier:"Sail India", item:"Angle Iron 50x50", newRate:50, oldRate:48 },
];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt = (n) => new Intl.NumberFormat("en-IN").format(Math.round(Number(n) || 0));
const fmtC = (n) => "₹" + new Intl.NumberFormat("en-IN").format(Math.round(Number(n) || 0));

function ageDays(dateStr) {
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return 0;
  return Math.floor((new Date() - d) / 86400000);
}

function ageBucket(days) {
  if (days <= 15) return { label: "0–15 Days", color: "#16a34a", badge: "15d" };
  if (days <= 30) return { label: "16–30 Days", color: "#ca8a04", badge: "30d" };
  if (days <= 45) return { label: "31–45 Days", color: "#ea580c", badge: "45d" };
  if (days <= 60) return { label: "46–60 Days", color: "#dc2626", badge: "60d" };
  return { label: "60+ Days", color: "#7f1d1d", badge: "60d+" };
}

const DELIVERY_COLORS = { Regular:"#3b82f6", "Urgent Rail":"#f59e0b", "Urgent Air":"#ef4444" };
const SUPPLIER_COLORS = ["#6366f1","#06b6d4","#10b981","#f59e0b","#ef4444","#8b5cf6"];
const STATUS_COLOR = { Complete:"#16a34a", Partial:"#ca8a04", Pending:"#dc2626", Cancelled:"#6b7280" };

// ─── COMPONENTS ───────────────────────────────────────────────────────────────

function KpiCard({ icon, label, value, sub, color = "#6366f1", dark }) {
  return (
    <div className={`rounded-xl p-4 flex flex-col gap-1 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
      <div className="flex items-center justify-between">
        <span className="text-2xl">{icon}</span>
        <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ background: color + "22", color }}>{sub}</span>
      </div>
      <p className={`text-2xl font-bold mt-1 ${dark ? "text-white" : "text-gray-800"}`}>{value}</p>
      <p className={`text-xs ${dark ? "text-gray-400" : "text-gray-500"}`}>{label}</p>
    </div>
  );
}

function Badge({ status }) {
  if (!status || status === "N/A") {
    return <span className="px-2 py-0.5 rounded-full text-xs font-semibold bg-gray-100 text-gray-400">—</span>;
  }
  return (
    <span className="px-2 py-0.5 rounded-full text-xs font-semibold"
      style={{ background: (STATUS_COLOR[status] || "#6b7280") + "22", color: STATUS_COLOR[status] || "#6b7280" }}>
      {status}
    </span>
  );
}

function AgeBadge({ days }) {
  const b = ageBucket(days);
  return (
    <span className="px-2 py-0.5 rounded-full text-xs font-semibold"
      style={{ background: b.color + "22", color: b.color }}>
      {b.badge}
    </span>
  );
}

function ShortExcess({ ordered, received }) {
  const diff = received - ordered;
  const pct = ordered ? ((diff / ordered) * 100).toFixed(1) : 0;
  if (diff === 0) return <span className="text-green-600 font-medium text-xs">✓ Complete</span>;
  if (diff < 0) return <span className="text-red-600 font-medium text-xs">▼ {Math.abs(diff)} ({pct}%)</span>;
  return <span className="text-blue-600 font-medium text-xs">▲ +{diff} (+{pct}%)</span>;
}

function SectionHeader({ title, sub, dark }) {
  return (
    <div className="mb-4">
      <h2 className={`text-lg font-semibold ${dark ? "text-white" : "text-gray-800"}`}>{title}</h2>
      {sub && <p className={`text-xs mt-0.5 ${dark ? "text-gray-400" : "text-gray-500"}`}>{sub}</p>}
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({ po, receipts, dark }) {
  const totalQty = po.reduce((a, r) => a + r.qty, 0);
  const totalReceived = po.reduce((a, r) => a + r.receivedQty, 0);
  const totalPending = po.filter(r => r.status !== "Complete").reduce((a, r) => a + (r.qty - r.receivedQty), 0);
  const totalShort = po.filter(r => r.receivedQty < r.qty && r.receivedQty > 0).reduce((a, r) => a + (r.qty - r.receivedQty), 0);
  const totalExcess = po.filter(r => r.receivedQty > r.qty).reduce((a, r) => a + (r.receivedQty - r.qty), 0);
  const totalValue = po.reduce((a, r) => a + r.totalAmt, 0);
  const totalGST = po.reduce((a, r) => a + (r.totalAmt - r.taxableAmt), 0);
  const suppliers = [...new Set(po.map(r => r.supplier))].length;

  // Monthly data
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
  const monthlyData = months.map((m, i) => {
    const mpo = po.filter(r => new Date(r.poDate).getMonth() === i);
    return {
      month: m,
      ordered: mpo.reduce((a, r) => a + r.qty, 0),
      received: mpo.reduce((a, r) => a + r.receivedQty, 0),
      pending: mpo.filter(r => r.status !== "Complete").reduce((a, r) => a + Math.max(0, r.qty - r.receivedQty), 0),
      value: Math.round(mpo.reduce((a, r) => a + r.totalAmt, 0) / 1000),
    };
  }).filter(d => d.ordered > 0);

  // Delivery distribution
  const deliveryData = ["Regular", "Urgent Rail", "Urgent Air"].map(t => ({
    name: t,
    count: po.filter(r => r.deliveryType === t).length,
    qty: po.filter(r => r.deliveryType === t).reduce((a, r) => a + r.qty, 0),
    value: Math.round(po.filter(r => r.deliveryType === t).reduce((a, r) => a + r.totalAmt, 0) / 1000),
  }));

  // Supplier data
  const supplierMap = {};
  po.forEach(r => {
    if (!supplierMap[r.supplier]) supplierMap[r.supplier] = { qty: 0, value: 0, pending: 0 };
    supplierMap[r.supplier].qty += r.qty;
    supplierMap[r.supplier].value += r.totalAmt;
    supplierMap[r.supplier].pending += Math.max(0, r.qty - r.receivedQty);
  });
  const supplierData = Object.entries(supplierMap).map(([name, d]) => ({ name, ...d, value: Math.round(d.value / 1000) })).sort((a, b) => b.value - a.value);

  // Ageing
  const pendingPO = po.filter(r => r.status !== "Complete");
  const ageGroups = ["0–15 Days", "16–30 Days", "31–45 Days", "46–60 Days", "60+ Days"];
  const ageData = ageGroups.map(g => ({
    name: g,
    count: pendingPO.filter(r => ageBucket(ageDays(r.poDate)).label === g).length,
  }));

  // Notifications
  const notifs = pendingPO.filter(r => ageDays(r.poDate) > 15);

  return (
    <div className="space-y-6">
      {/* Notifications */}
      {notifs.length > 0 && (
        <div className={`rounded-xl p-3 border-l-4 border-red-500 ${dark ? "bg-red-900/20" : "bg-red-50"}`}>
          <p className={`text-sm font-semibold ${dark ? "text-red-400" : "text-red-700"}`}>⚠ {notifs.length} POs overdue (15+ days pending)</p>
          <div className="flex flex-wrap gap-2 mt-1">
            {notifs.slice(0, 5).map((p, i) => (
              <span key={p._rowId ?? i} className="text-xs bg-red-100 text-red-700 px-2 py-0.5 rounded-full">{p.poNo} — {ageDays(p.poDate)}d</span>
            ))}
          </div>
        </div>
      )}

      {/* KPI Grid */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <KpiCard icon="📦" label="Total PO Quantity" value={fmt(totalQty)} sub="units" color="#6366f1" dark={dark} />
        <KpiCard icon="✅" label="Total Received" value={fmt(totalReceived)} sub="units" color="#16a34a" dark={dark} />
        <KpiCard icon="⏳" label="Total Pending" value={fmt(totalPending)} sub="units" color="#ca8a04" dark={dark} />
        <KpiCard icon="📉" label="Total Short" value={fmt(totalShort)} sub="units" color="#dc2626" dark={dark} />
        <KpiCard icon="📈" label="Total Excess" value={fmt(totalExcess)} sub="units" color="#3b82f6" dark={dark} />
        <KpiCard icon="💰" label="Total PO Value" value={fmtC(totalValue)} sub="incl. GST" color="#8b5cf6" dark={dark} />
        <KpiCard icon="🧾" label="Total GST" value={fmtC(totalGST)} sub="payable" color="#06b6d4" dark={dark} />
        <KpiCard icon="🏭" label="Suppliers" value={suppliers} sub="active" color="#f59e0b" dark={dark} />
      </div>

      {/* Monthly Charts */}
      <div className={`rounded-xl p-4 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
        <SectionHeader title="Monthly Procurement" sub="Quantity ordered vs received vs pending" dark={dark} />
        <ResponsiveContainer width="100%" height={260}>
          <ComposedChart data={monthlyData}>
            <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#374151" : "#f3f4f6"} />
            <XAxis dataKey="month" tick={{ fill: dark ? "#9ca3af" : "#6b7280", fontSize: 12 }} />
            <YAxis tick={{ fill: dark ? "#9ca3af" : "#6b7280", fontSize: 11 }} />
            <Tooltip contentStyle={{ background: dark ? "#1f2937" : "#fff", border: "1px solid #e5e7eb", borderRadius: 8 }} />
            <Legend />
            <Bar dataKey="ordered" name="Ordered" fill="#6366f1" radius={[3, 3, 0, 0]} />
            <Bar dataKey="received" name="Received" fill="#10b981" radius={[3, 3, 0, 0]} />
            <Line type="monotone" dataKey="value" name="Value (₹K)" stroke="#f59e0b" strokeWidth={2} dot={{ r: 4 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* Delivery Type */}
        <div className={`rounded-xl p-4 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
          <SectionHeader title="Delivery Type" sub="Distribution by type" dark={dark} />
          <div className="space-y-3 mt-2">
            {deliveryData.map(d => (
              <div key={d.name}>
                <div className="flex justify-between text-xs mb-1">
                  <span className={dark ? "text-gray-300" : "text-gray-600"}>{d.name}</span>
                  <span className="font-semibold" style={{ color: DELIVERY_COLORS[d.name] }}>{d.count} POs · {fmt(d.qty)} units</span>
                </div>
                <div className={`h-2 rounded-full ${dark ? "bg-gray-700" : "bg-gray-100"}`}>
                  <div className="h-2 rounded-full" style={{ width: `${(d.count / po.length) * 100}%`, background: DELIVERY_COLORS[d.name] }} />
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* PO Ageing */}
        <div className={`rounded-xl p-4 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
          <SectionHeader title="Pending PO Ageing" sub="Overdue analysis" dark={dark} />
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={ageData} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#374151" : "#f3f4f6"} />
              <XAxis type="number" tick={{ fill: dark ? "#9ca3af" : "#6b7280", fontSize: 11 }} />
              <YAxis dataKey="name" type="category" tick={{ fill: dark ? "#9ca3af" : "#6b7280", fontSize: 11 }} width={80} />
              <Tooltip contentStyle={{ background: dark ? "#1f2937" : "#fff", border: "1px solid #e5e7eb", borderRadius: 8 }} />
              <Bar dataKey="count" name="PO Count" fill="#ef4444" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Supplier Dashboard */}
      <div className={`rounded-xl p-4 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
        <SectionHeader title="Supplier Performance" sub="Ranked by procurement value" dark={dark} />
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className={dark ? "border-b border-gray-700" : "border-b border-gray-100"}>
                {["Rank", "Supplier", "Total Qty", "Pending Qty", "Value (₹K)", "Performance"].map(h => (
                  <th key={h} className={`text-left py-2 px-3 text-xs font-semibold ${dark ? "text-gray-400" : "text-gray-500"}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {supplierData.map((s, i) => {
                const perf = Math.round(((s.qty - s.pending) / s.qty) * 100);
                return (
                  <tr key={s.name} className={`border-b ${dark ? "border-gray-700/50 hover:bg-gray-700/30" : "border-gray-50 hover:bg-gray-50"}`}>
                    <td className="py-2 px-3">
                      <span className="w-6 h-6 rounded-full text-xs font-bold flex items-center justify-center"
                        style={{ background: SUPPLIER_COLORS[i % 6] + "22", color: SUPPLIER_COLORS[i % 6] }}>#{i + 1}</span>
                    </td>
                    <td className={`py-2 px-3 font-medium ${dark ? "text-white" : "text-gray-800"}`}>{s.name}</td>
                    <td className={`py-2 px-3 ${dark ? "text-gray-300" : "text-gray-600"}`}>{fmt(s.qty)}</td>
                    <td className="py-2 px-3">
                      <span className={s.pending > 0 ? "text-red-500 font-medium" : "text-green-600"}>{fmt(s.pending)}</span>
                    </td>
                    <td className={`py-2 px-3 font-semibold ${dark ? "text-indigo-400" : "text-indigo-600"}`}>₹{fmt(s.value)}K</td>
                    <td className="py-2 px-3">
                      <div className="flex items-center gap-2">
                        <div className={`flex-1 h-1.5 rounded-full ${dark ? "bg-gray-700" : "bg-gray-100"}`}>
                          <div className="h-1.5 rounded-full" style={{ width: `${perf}%`, background: perf > 80 ? "#16a34a" : perf > 50 ? "#ca8a04" : "#dc2626" }} />
                        </div>
                        <span className="text-xs font-medium" style={{ color: perf > 80 ? "#16a34a" : perf > 50 ? "#ca8a04" : "#dc2626" }}>{perf}%</span>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── PO TABLE ─────────────────────────────────────────────────────────────────
function POTable({ po, dark }) {
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterDelivery, setFilterDelivery] = useState("All");
  const [filterSupplier, setFilterSupplier] = useState("All");
  const [sortKey, setSortKey] = useState("poDate");
  const [sortDir, setSortDir] = useState("desc");
  const [page, setPage] = useState(1);
  const PER_PAGE = 8;

  const suppliers = useMemo(() => ["All", ...new Set(po.map(r => r.supplier))], [po]);

  const filtered = useMemo(() => {
    let d = po.filter(r => {
      const q = search.toLowerCase();
      const matchSearch = !q || r.poNo.toLowerCase().includes(q) || r.supplier.toLowerCase().includes(q) || r.item.toLowerCase().includes(q);
      const matchStatus = filterStatus === "All" || r.status === filterStatus;
      const matchDelivery = filterDelivery === "All" || r.deliveryType === filterDelivery;
      const matchSupplier = filterSupplier === "All" || r.supplier === filterSupplier;
      return matchSearch && matchStatus && matchDelivery && matchSupplier;
    });
    d.sort((a, b) => {
      let av = a[sortKey], bv = b[sortKey];
      if (typeof av === "string") av = av.toLowerCase(), bv = bv.toLowerCase();
      return sortDir === "asc" ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
    });
    return d;
  }, [po, search, filterStatus, filterDelivery, filterSupplier, sortKey, sortDir]);

  const pages = Math.ceil(filtered.length / PER_PAGE);
  const paged = filtered.slice((page - 1) * PER_PAGE, page * PER_PAGE);

  const th = (key, label) => (
    <th key={key} className={`text-left py-2 px-3 text-xs font-semibold cursor-pointer select-none ${dark ? "text-gray-400 hover:text-gray-200" : "text-gray-500 hover:text-gray-700"}`}
      onClick={() => { setSortKey(key); setSortDir(k => k === "asc" ? "desc" : "asc"); }}>
      {label} {sortKey === key ? (sortDir === "asc" ? "↑" : "↓") : ""}
    </th>
  );

  const input = `px-3 py-1.5 rounded-lg text-sm border outline-none focus:ring-2 focus:ring-indigo-300 ${dark ? "bg-gray-700 border-gray-600 text-white" : "bg-white border-gray-200 text-gray-800"}`;

  return (
    <div className="space-y-4">
      {/* Filters */}
      <div className="flex flex-wrap gap-2 items-center">
        <input className={`${input} flex-1 min-w-48`} placeholder="🔍 Search PO, supplier, item..." value={search} onChange={e => { setSearch(e.target.value); setPage(1); }} />
        <select className={input} value={filterStatus} onChange={e => { setFilterStatus(e.target.value); setPage(1); }}>
          {["All", "Complete", "Partial", "Pending"].map(s => <option key={s}>{s}</option>)}
        </select>
        <select className={input} value={filterDelivery} onChange={e => { setFilterDelivery(e.target.value); setPage(1); }}>
          {["All", "Regular", "Urgent Rail", "Urgent Air"].map(s => <option key={s}>{s}</option>)}
        </select>
        <select className={input} value={filterSupplier} onChange={e => { setFilterSupplier(e.target.value); setPage(1); }}>
          {suppliers.map(s => <option key={s}>{s}</option>)}
        </select>
        <span className={`text-xs px-2 py-1 rounded-lg ${dark ? "bg-gray-700 text-gray-300" : "bg-gray-100 text-gray-500"}`}>{filtered.length} records</span>
      </div>

      {/* Table */}
      <div className={`rounded-xl border overflow-hidden ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm min-w-[900px]">
            <thead className={dark ? "bg-gray-700/50" : "bg-gray-50"}>
              <tr className={`border-b ${dark ? "border-gray-700" : "border-gray-100"}`}>
                {th("poDate", "PO Date")}
                {th("poNo", "PO No")}
                {th("supplier", "Supplier")}
                {th("item", "Item")}
                {th("qty", "Ordered")}
                {th("receivedQty", "Received")}
                <th className={`text-left py-2 px-3 text-xs font-semibold ${dark ? "text-gray-400" : "text-gray-500"}`}>Short/Excess</th>
                {th("deliveryType", "Delivery")}
                {th("totalAmt", "Value")}
                {th("status", "Status")}
                <th className={`text-left py-2 px-3 text-xs font-semibold ${dark ? "text-gray-400" : "text-gray-500"}`}>Age</th>
              </tr>
            </thead>
            <tbody>
              {paged.length === 0 && (
                <tr><td colSpan={12} className={`py-10 text-center text-sm ${dark ? "text-gray-500" : "text-gray-400"}`}>No records found. Adjust your filters.</td></tr>
              )}
              {paged.map((r) => (
                <tr key={r._rowId ?? r.poNo} className={`border-b transition-colors ${dark ? "border-gray-700/50 hover:bg-gray-700/30" : "border-gray-50 hover:bg-indigo-50/30"}`}>
                  <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-400" : "text-gray-500"}`}>{r.poDate}</td>
                  <td className={`py-2.5 px-3 font-mono text-xs font-semibold ${dark ? "text-indigo-400" : "text-indigo-600"}`}>{r.poNo}</td>
                  <td className={`py-2.5 px-3 font-medium ${dark ? "text-white" : "text-gray-800"}`}>{r.supplier}</td>
                  <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-300" : "text-gray-600"}`}>{r.item || <span className="text-gray-400 italic">No item (PO not created)</span>}</td>
                  <td className={`py-2.5 px-3 text-right font-medium ${dark ? "text-gray-200" : "text-gray-700"}`}>{fmt(r.qty)}</td>
                  <td className={`py-2.5 px-3 text-right ${dark ? "text-gray-200" : "text-gray-700"}`}>{fmt(r.receivedQty)}</td>
                  <td className="py-2.5 px-3"><ShortExcess ordered={r.qty} received={r.receivedQty} /></td>
                  <td className="py-2.5 px-3">
                    <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{ background: (DELIVERY_COLORS[r.deliveryType] || "#6b7280") + "22", color: DELIVERY_COLORS[r.deliveryType] || "#6b7280" }}>
                      {r.deliveryType}
                    </span>
                  </td>
                  <td className={`py-2.5 px-3 text-right text-xs font-semibold ${dark ? "text-emerald-400" : "text-emerald-600"}`}>{fmtC(r.totalAmt)}</td>
                  <td className="py-2.5 px-3"><Badge status={r.status} /></td>
                  <td className="py-2.5 px-3">{r.status !== "Complete" ? <AgeBadge days={ageDays(r.poDate)} /> : <span className="text-xs text-gray-400">—</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Pagination */}
      {pages > 1 && (
        <div className="flex items-center justify-between">
          <span className={`text-xs ${dark ? "text-gray-400" : "text-gray-500"}`}>Page {page} of {pages}</span>
          <div className="flex gap-1">
            <button onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page === 1}
              className={`px-3 py-1 rounded-lg text-xs border ${dark ? "bg-gray-700 border-gray-600 text-gray-300 disabled:opacity-40" : "bg-white border-gray-200 text-gray-600 disabled:opacity-40"}`}>← Prev</button>
            {Array.from({ length: Math.min(pages, 5) }, (_, i) => i + 1).map(p => (
              <button key={p} onClick={() => setPage(p)}
                className={`px-3 py-1 rounded-lg text-xs border font-medium ${p === page ? "bg-indigo-600 border-indigo-600 text-white" : dark ? "bg-gray-700 border-gray-600 text-gray-300" : "bg-white border-gray-200 text-gray-600"}`}>{p}</button>
            ))}
            <button onClick={() => setPage(p => Math.min(pages, p + 1))} disabled={page === pages}
              className={`px-3 py-1 rounded-lg text-xs border ${dark ? "bg-gray-700 border-gray-600 text-gray-300 disabled:opacity-40" : "bg-white border-gray-200 text-gray-600 disabled:opacity-40"}`}>Next →</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── RECEIPTS TABLE ───────────────────────────────────────────────────────────
function ReceiptsTable({ receipts, dark }) {
  return (
    <div className={`rounded-xl border overflow-hidden ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead className={dark ? "bg-gray-700/50" : "bg-gray-50"}>
            <tr className={`border-b ${dark ? "border-gray-700" : "border-gray-100"}`}>
              {["Date", "PO No", "Voucher No", "Supplier", "Item", "Received Qty", "Rate", "Value"].map(h => (
                <th key={h} className={`text-left py-2 px-3 text-xs font-semibold ${dark ? "text-gray-400" : "text-gray-500"}`}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {receipts.map((r, i) => (
              <tr key={i} className={`border-b ${dark ? "border-gray-700/50 hover:bg-gray-700/30" : "border-gray-50 hover:bg-gray-50"}`}>
                <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-400" : "text-gray-500"}`}>{r.date}</td>
                <td className={`py-2.5 px-3 font-mono text-xs font-semibold ${dark ? "text-indigo-400" : "text-indigo-600"}`}>{r.poNo}</td>
                <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-300" : "text-gray-600"}`}>{r.voucherNo}</td>
                <td className={`py-2.5 px-3 font-medium ${dark ? "text-white" : "text-gray-800"}`}>{r.supplier}</td>
                <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-300" : "text-gray-600"}`}>{r.item}</td>
                <td className={`py-2.5 px-3 text-right font-semibold ${dark ? "text-green-400" : "text-green-600"}`}>{fmt(r.receivedQty)}</td>
                <td className={`py-2.5 px-3 text-right ${dark ? "text-gray-300" : "text-gray-600"}`}>{fmtC(r.rate)}</td>
                <td className={`py-2.5 px-3 text-right font-semibold ${dark ? "text-emerald-400" : "text-emerald-600"}`}>{fmtC(r.receivedQty * r.rate)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── RATE HISTORY ─────────────────────────────────────────────────────────────
function RateHistory({ rates, dark }) {
  const items = [...new Set(rates.map(r => r.item))];
  const suppliers = [...new Set(rates.map(r => r.supplier))];
  const [selItem, setSelItem] = useState(items[0] || "");
  const [selSupplier, setSelSupplier] = useState("All");

  const filtered = rates.filter(r =>
    r.item === selItem && (selSupplier === "All" || r.supplier === selSupplier)
  ).sort((a, b) => a.date.localeCompare(b.date));

  const chartData = filtered.map(r => ({ date: r.date.slice(5), rate: r.newRate }));
  const input = `px-3 py-1.5 rounded-lg text-sm border outline-none ${dark ? "bg-gray-700 border-gray-600 text-white" : "bg-white border-gray-200 text-gray-800"}`;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-2">
        <select className={input} value={selItem} onChange={e => setSelItem(e.target.value)}>
          {items.map(i => <option key={i}>{i}</option>)}
        </select>
        <select className={input} value={selSupplier} onChange={e => setSelSupplier(e.target.value)}>
          {["All", ...suppliers].map(s => <option key={s}>{s}</option>)}
        </select>
      </div>

      {/* Rate Trend Chart */}
      {chartData.length > 0 && (
        <div className={`rounded-xl p-4 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
          <p className={`text-sm font-semibold mb-3 ${dark ? "text-white" : "text-gray-800"}`}>Rate trend — {selItem}</p>
          <ResponsiveContainer width="100%" height={180}>
            <AreaChart data={chartData}>
              <defs>
                <linearGradient id="rateGrad" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#6366f1" stopOpacity={0.2} />
                  <stop offset="95%" stopColor="#6366f1" stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#374151" : "#f3f4f6"} />
              <XAxis dataKey="date" tick={{ fill: dark ? "#9ca3af" : "#6b7280", fontSize: 11 }} />
              <YAxis tick={{ fill: dark ? "#9ca3af" : "#6b7280", fontSize: 11 }} />
              <Tooltip contentStyle={{ background: dark ? "#1f2937" : "#fff", border: "1px solid #e5e7eb", borderRadius: 8 }} />
              <Area type="monotone" dataKey="rate" name="Rate (₹)" stroke="#6366f1" strokeWidth={2} fill="url(#rateGrad)" dot={{ r: 4 }} />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}

      {/* Rate History Table */}
      <div className={`rounded-xl border overflow-hidden ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className={dark ? "bg-gray-700/50" : "bg-gray-50"}>
              <tr className={`border-b ${dark ? "border-gray-700" : "border-gray-100"}`}>
                {["Date", "Supplier", "Item", "Old Rate", "New Rate", "Change", "Change %"].map(h => (
                  <th key={h} className={`text-left py-2 px-3 text-xs font-semibold ${dark ? "text-gray-400" : "text-gray-500"}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr><td colSpan={7} className={`py-8 text-center text-sm ${dark ? "text-gray-500" : "text-gray-400"}`}>No rate history found.</td></tr>
              )}
              {filtered.map((r, i) => {
                const diff = r.newRate - r.oldRate;
                const pct = ((diff / r.oldRate) * 100).toFixed(2);
                return (
                  <tr key={i} className={`border-b ${dark ? "border-gray-700/50 hover:bg-gray-700/30" : "border-gray-50 hover:bg-gray-50"}`}>
                    <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-400" : "text-gray-500"}`}>{r.date}</td>
                    <td className={`py-2.5 px-3 font-medium ${dark ? "text-white" : "text-gray-800"}`}>{r.supplier}</td>
                    <td className={`py-2.5 px-3 text-xs ${dark ? "text-gray-300" : "text-gray-600"}`}>{r.item}</td>
                    <td className={`py-2.5 px-3 text-right ${dark ? "text-gray-300" : "text-gray-600"}`}>₹{r.oldRate}</td>
                    <td className={`py-2.5 px-3 text-right font-semibold ${dark ? "text-white" : "text-gray-800"}`}>₹{r.newRate}</td>
                    <td className={`py-2.5 px-3 text-right font-semibold ${diff >= 0 ? "text-red-500" : "text-green-600"}`}>{diff >= 0 ? "+" : ""}{diff}</td>
                    <td className={`py-2.5 px-3 text-right text-xs font-medium ${diff >= 0 ? "text-red-500" : "text-green-600"}`}>{diff >= 0 ? "+" : ""}{pct}%</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── EXCEL UPLOAD ─────────────────────────────────────────────────────────────
function ExcelUpload({ onUpload, dark }) {
  const [dragging, setDragging] = useState(false);
  const [msg, setMsg] = useState("");

  const parse = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary", cellDates: true });

        // Find the sheet that actually looks like PO data, by header match —
        // rather than assuming sheet order, since real exports may only have
        // a single "PO_Data" sheet (no separate Receipts / Rate History tabs).
        let poRows = [];
        let foundSheet = null;
        for (const name of wb.SheetNames) {
          const json = XLSX.utils.sheet_to_json(wb.Sheets[name]);
          if (json.length === 0) continue;
          const keys = Object.keys(json[0]).map(k => k.trim());
          const looksLikePO = keys.some(k =>
            ["P.O No.", "poNo", "Items", "item", "Supplier Name", "supplier"].includes(k)
          );
          if (looksLikePO) { poRows = json; foundSheet = name; break; }
        }

        if (!foundSheet) {
          setMsg("❌ Couldn't find a PO data sheet. Expected columns like 'P.O No.', 'Supplier Name', 'Items'.");
          return;
        }

        const mapped = mapPORows(poRows);
        if (mapped.length === 0) {
          setMsg("❌ Sheet found but no usable rows after parsing. Check the file isn't empty or malformed.");
          return;
        }

        // Receipts / Rate History sheets — matched by recognisable name
        // first (real export uses "Recvd_Item" and "Rate Update"; a
        // "Report" pivot-summary sheet, if present, is intentionally
        // skipped since it has no usable row-level data). Falls back to
        // positional sheets for older/simple templates.
        const byName = (patterns) => wb.SheetNames.find(n => patterns.some(p => p.test(n)));
        const recSheetName = byName([/recv/i, /receipt/i]);
        const rateSheetName = byName([/rate/i]);
        const otherSheets = wb.SheetNames.filter(n =>
          n !== foundSheet && !/report/i.test(n)
        );

        const recRaw = recSheetName ? XLSX.utils.sheet_to_json(wb.Sheets[recSheetName])
          : (otherSheets[0] ? XLSX.utils.sheet_to_json(wb.Sheets[otherSheets[0]]) : []);
        const ratRaw = rateSheetName ? XLSX.utils.sheet_to_json(wb.Sheets[rateSheetName])
          : (otherSheets[1] ? XLSX.utils.sheet_to_json(wb.Sheets[otherSheets[1]]) : []);

        const rec = mapReceiptRows(recRaw);
        const rat = mapRateRows(ratRaw);

        onUpload({ po: mapped, receipts: rec, rates: rat });
        const skipped = poRows.length - mapped.length;
        setMsg(`✅ Imported ${mapped.length} line items from "${foundSheet}"${rec.length ? `, ${rec.length} receipts` : ""}${rat.length ? `, ${rat.length} rate records` : ""}${skipped ? ` (${skipped} blank/cancelled rows skipped)` : ""}.`);
      } catch (err) {
        setMsg("❌ Error reading file: " + (err?.message || "unknown error"));
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="space-y-4 max-w-xl mx-auto mt-8">
      <div
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={e => { e.preventDefault(); setDragging(false); parse(e.dataTransfer.files[0]); }}
        className={`border-2 border-dashed rounded-2xl p-10 text-center cursor-pointer transition-all ${dragging ? "border-indigo-400 bg-indigo-50" : dark ? "border-gray-600 bg-gray-800 hover:border-gray-500" : "border-gray-200 bg-gray-50 hover:border-indigo-300"}`}
        onClick={() => document.getElementById("xl-input").click()}
      >
        <div className="text-5xl mb-3">📊</div>
        <p className={`font-semibold text-base ${dark ? "text-white" : "text-gray-700"}`}>Drop your Excel file here</p>
        <p className={`text-sm mt-1 ${dark ? "text-gray-400" : "text-gray-400"}`}>or click to browse · .xlsx / .xls</p>
        <p className={`text-xs mt-3 ${dark ? "text-gray-500" : "text-gray-400"}`}>Auto-detects PP Auto PO_Data export format or the dashboard's own template</p>
      </div>
      <input id="xl-input" type="file" accept=".xlsx,.xls" className="hidden" onChange={e => parse(e.target.files[0])} />
      {msg && <p className={`text-sm text-center ${msg.startsWith("✅") ? "text-green-600" : "text-red-500"}`}>{msg}</p>}
      <p className={`text-xs text-center ${dark ? "text-gray-500" : "text-gray-400"}`}>Currently showing sample data. Upload your Excel to replace.</p>
    </div>
  );
}

// ─── ITEM DASHBOARD ───────────────────────────────────────────────────────────
function ItemDashboard({ po, dark }) {
  const [selectedItem, setSelectedItem] = useState(null);

  const itemMap = {};

  po.forEach(r => {
    if (!itemMap[r.item]) {
      itemMap[r.item] = {
        qty: 0,
        received: 0,
        pending: 0,
        spend: 0
      };
    }

    itemMap[r.item].qty += r.qty;
    itemMap[r.item].received += r.receivedQty;
    itemMap[r.item].pending += Math.max(0, r.qty - r.receivedQty);
    itemMap[r.item].spend += r.totalAmt;
  });

  const data = Object.entries(itemMap)
    .map(([name, d]) => ({
      fullName: name,
      name: name.length > 18 ? name.slice(0, 18) + "…" : name,
      ...d,
      spend: Math.round(d.spend / 1000)
    }))
    .sort((a, b) => b.spend - a.spend)
    .slice(0, 10);

  return (
    <div className="space-y-4">

      <div className={`rounded-xl p-4 border ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
        <SectionHeader
          title="Top 10 Items by Spend"
          sub="Procurement value in ₹K"
          dark={dark}
        />

        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={data} layout="vertical">
            <CartesianGrid
              strokeDasharray="3 3"
              stroke={dark ? "#374151" : "#f3f4f6"}
            />
            <XAxis
              type="number"
              tick={{
                fill: dark ? "#9ca3af" : "#6b7280",
                fontSize: 11
              }}
            />
            <YAxis
              dataKey="name"
              type="category"
              tick={{
                fill: dark ? "#9ca3af" : "#6b7280",
                fontSize: 10
              }}
              width={120}
            />
            <Tooltip
              contentStyle={{
                background: dark ? "#1f2937" : "#fff",
                border: "1px solid #e5e7eb",
                borderRadius: 8
              }}
              formatter={v => "₹" + fmt(v) + "K"}
            />
            <Legend />
            <Bar
              dataKey="spend"
              name="Spend (₹K)"
              fill="#6366f1"
              radius={[0, 4, 4, 0]}
            />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className={`rounded-xl border overflow-hidden ${dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"} shadow-sm`}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className={dark ? "bg-gray-700/50" : "bg-gray-50"}>
              <tr className={`border-b ${dark ? "border-gray-700" : "border-gray-100"}`}>
                {["Item", "Ordered Qty", "Received Qty", "Pending Qty", "Spend (₹K)"].map(h => (
                  <th
                    key={h}
                    className={`text-left py-2 px-3 text-xs font-semibold ${dark ? "text-gray-400" : "text-gray-500"}`}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>

            <tbody>
              {data.map((r, i) => (
                <tr
                  key={i}
                  onClick={() => setSelectedItem(r.fullName)}
                  className={`cursor-pointer border-b ${
                    dark
                      ? "border-gray-700/50 hover:bg-gray-700/30"
                      : "border-gray-50 hover:bg-blue-50"
                  }`}
                >
                  <td className={`py-2.5 px-3 font-medium ${dark ? "text-white" : "text-gray-800"}`}>
                    {r.name}
                  </td>

                  <td className={`py-2.5 px-3 text-right ${dark ? "text-gray-300" : "text-gray-600"}`}>
                    {fmt(r.qty)}
                  </td>

                  <td className="py-2.5 px-3 text-right text-green-600 font-medium">
                    {fmt(r.received)}
                  </td>

                  <td className={`py-2.5 px-3 text-right ${
                    r.pending > 0
                      ? "text-red-500 font-medium"
                      : "text-green-600"
                  }`}>
                    {fmt(r.pending)}
                  </td>

                  <td className={`py-2.5 px-3 text-right font-semibold ${
                    dark ? "text-indigo-400" : "text-indigo-600"
                  }`}>
                    ₹{fmt(r.spend)}K
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {selectedItem && (
        <div className={`rounded-xl border overflow-hidden ${
          dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100"
        } shadow-sm`}>

          <div className="p-3 border-b">
            <h3 className="font-semibold">
              Pending PO Details - {selectedItem}
            </h3>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-sm">

              <thead className={dark ? "bg-gray-700/50" : "bg-gray-50"}>
                <tr>
                  <th className="text-left py-2 px-3">PO No</th>
                  <th className="text-left py-2 px-3">Supplier</th>
                  <th className="text-right py-2 px-3">Ordered Qty</th>
                  <th className="text-right py-2 px-3">Received Qty</th>
                  <th className="text-right py-2 px-3">Pending Qty</th>
                </tr>
              </thead>

              <tbody>
                {po
                  .filter(
                    p =>
                      p.item &&
                      p.item === selectedItem
                  )
                  .map((p, idx) => (
                    <tr key={idx}>
                      <td className="py-2 px-3 font-semibold text-indigo-600">
                        {p.poNo}
                      </td>

                      <td className="py-2 px-3">
                        {p.supplier}
                      </td>

                      <td className="py-2 px-3 text-right">
                        {fmt(p.qty)}
                      </td>

                      <td className="py-2 px-3 text-right text-green-600">
                        {fmt(p.receivedQty)}
                      </td>

                      <td className="py-2 px-3 text-right text-red-500 font-semibold">
                        {fmt(Math.max(0, p.qty - p.receivedQty))}
                      </td>
                    </tr>
                  ))}
              </tbody>

            </table>
          </div>
        </div>
      )}

    </div>
  );
}
// ─── MAIN APP ─────────────────────────────────────────────────────────────────
const NAV = [
  { id: "dashboard", label: "Dashboard", icon: "📊" },
  { id: "po", label: "PO Master", icon: "📋" },
  { id: "receipts", label: "Receipts", icon: "📥" },
  { id: "items", label: "Item Analytics", icon: "📦" },
  { id: "rates", label: "Rate History", icon: "📈" },
  { id: "upload", label: "Import Excel", icon: "⬆️" },
];

export default function POManagement() {
  const [dark, setDark] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [sideOpen, setSideOpen] = useState(true);
  const [po, setPO] = useState(() => mapPORows(MOCK_PO));
  const [receipts, setReceipts] = useState(() => mapReceiptRows(MOCK_RECEIPTS));
  const [rates, setRates] = useState(() => mapRateRows(MOCK_RATES));
  const [user] = useState({ name: "Ayush Singh", role: "Admin" });

  const handleUpload = ({ po: p, receipts: r, rates: rt }) => {
    if (p.length) setPO(p);
    if (r.length) setReceipts(r);
    if (rt.length) setRates(rt);
    setTab("dashboard");
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(po), "PO Master");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(receipts), "Receipts");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rates), "Rate History");
    XLSX.writeFile(wb, "PO_Report_" + new Date().toISOString().slice(0, 10) + ".xlsx");
  };

  const bg = dark ? "bg-gray-900" : "bg-gray-50";
  const sidebar = dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100";
  const topbar = dark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-100";

  const pendingCount = po.filter(r => r.status !== "Complete" && ageDays(r.poDate) > 15).length;

  return (
    <div className={`min-h-screen flex ${bg} font-sans`} style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      {/* Sidebar */}
      <aside className={`${sideOpen ? "w-56" : "w-14"} transition-all duration-200 flex-shrink-0 border-r ${sidebar} flex flex-col`}>
        <div className={`flex items-center gap-2 p-4 border-b ${dark ? "border-gray-700" : "border-gray-100"}`}>
          <span className="text-2xl">🏭</span>
          {sideOpen && <div>
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-gray-800"}`}>ProProcure</p>
            <p className={`text-xs ${dark ? "text-gray-400" : "text-gray-400"}`}>PO Management</p>
          </div>}
        </div>
        <nav className="flex-1 p-2 space-y-0.5">
          {NAV.map(n => (
            <button key={n.id} onClick={() => setTab(n.id)}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors ${tab === n.id ? "bg-indigo-600 text-white font-medium" : dark ? "text-gray-300 hover:bg-gray-700" : "text-gray-600 hover:bg-gray-50"}`}>
              <span className="text-base flex-shrink-0">{n.icon}</span>
              {sideOpen && <span className="truncate">{n.label}</span>}
              {sideOpen && n.id === "dashboard" && pendingCount > 0 && (
                <span className="ml-auto bg-red-500 text-white text-xs px-1.5 py-0.5 rounded-full">{pendingCount}</span>
              )}
            </button>
          ))}
        </nav>
        <div className={`p-3 border-t ${dark ? "border-gray-700" : "border-gray-100"}`}>
          {sideOpen && (
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-full bg-indigo-100 flex items-center justify-center text-xs font-bold text-indigo-700">
                {user.name.split(" ").map(n => n[0]).join("")}
              </div>
              <div>
                <p className={`text-xs font-semibold ${dark ? "text-white" : "text-gray-800"}`}>{user.name}</p>
                <p className={`text-xs ${dark ? "text-gray-400" : "text-gray-400"}`}>{user.role}</p>
              </div>
            </div>
          )}
        </div>
      </aside>

      {/* Main */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Topbar */}
        <header className={`flex items-center justify-between px-4 py-3 border-b ${topbar} sticky top-0 z-10`}>
          <div className="flex items-center gap-3">
            <button onClick={() => setSideOpen(o => !o)} className={`p-1.5 rounded-lg ${dark ? "hover:bg-gray-700 text-gray-300" : "hover:bg-gray-100 text-gray-500"}`}>
              ☰
            </button>
            <div>
              <h1 className={`text-base font-semibold ${dark ? "text-white" : "text-gray-800"}`}>
                {NAV.find(n => n.id === tab)?.icon} {NAV.find(n => n.id === tab)?.label}
              </h1>
              <p className={`text-xs ${dark ? "text-gray-400" : "text-gray-400"}`}>PP Auto Innovators · {new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })}</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={exportExcel}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border ${dark ? "border-gray-600 text-gray-300 hover:bg-gray-700" : "border-gray-200 text-gray-600 hover:bg-gray-50"}`}>
              ⬇ Export Excel
            </button>
            <button onClick={() => setDark(d => !d)}
              className={`p-1.5 rounded-lg text-sm ${dark ? "hover:bg-gray-700 text-yellow-300" : "hover:bg-gray-100 text-gray-500"}`}>
              {dark ? "☀️" : "🌙"}
            </button>
          </div>
        </header>

        {/* Content */}
        <main className="flex-1 p-4 md:p-5 overflow-auto">
          {tab === "dashboard" && <Dashboard po={po} receipts={receipts} dark={dark} />}
          {tab === "po" && <POTable po={po} dark={dark} />}
          {tab === "receipts" && <ReceiptsTable receipts={receipts} dark={dark} />}
          {tab === "items" && <ItemDashboard po={po} dark={dark} />}
          {tab === "rates" && <RateHistory rates={rates} dark={dark} />}
          {tab === "upload" && <ExcelUpload onUpload={handleUpload} dark={dark} />}
        </main>
      </div>
    </div>
  );
}
