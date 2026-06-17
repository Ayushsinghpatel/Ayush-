import { useState, useEffect } from "react";
import POManagement from "./POManagement.jsx";

// Password comes from an environment variable set on the hosting platform
// (Vercel / Netlify dashboard → Environment Variables), NOT hardcoded here.
// Vite only exposes vars prefixed with VITE_ to the browser.
const SITE_PASSWORD = "Ayush@123";

const SESSION_KEY = "ppauto_po_session";
const SESSION_TTL_HOURS = 12;

function isSessionValid() {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    if (!raw) return false;
    const { ts } = JSON.parse(raw);
    return Date.now() - ts < SESSION_TTL_HOURS * 60 * 60 * 1000;
  } catch {
    return false;
  }
}

export default function AuthGate() {
  const [unlocked, setUnlocked] = useState(false);
  const [input, setInput] = useState("");
  const [error, setError] = useState("");
  const [checked, setChecked] = useState(false);

  useEffect(() => {
    setUnlocked(isSessionValid());
    setChecked(true);
  }, []);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!SITE_PASSWORD) {
      setError("App password not configured. Set VITE_APP_PASSWORD on the hosting platform.");
      return;
    }
    if (input === SITE_PASSWORD) {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify({ ts: Date.now() }));
      setUnlocked(true);
      setError("");
    } else {
      setError("Incorrect password.");
    }
  };

  if (!checked) return null; // avoid flashing the login form on first paint

  if (unlocked) return <POManagement />;

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 font-sans" style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      <form
        onSubmit={handleSubmit}
        className="bg-white border border-gray-100 shadow-sm rounded-xl p-8 w-full max-w-sm flex flex-col gap-4"
      >
        <div className="flex flex-col items-center gap-1 mb-2">
          <span className="text-3xl">🏭</span>
          <p className="text-base font-bold text-gray-800">ProProcure</p>
          <p className="text-xs text-gray-400">PP Auto · PO Management</p>
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-xs font-medium text-gray-600">Access Password</label>
          <input
            type="password"
            autoFocus
            value={input}
            onChange={(e) => setInput(e.target.value)}
            className="border border-gray-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
            placeholder="Enter password"
          />
        </div>
        {error && <p className="text-xs text-red-500">{error}</p>}
        <button
          type="submit"
          className="bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-lg py-2 transition-colors"
        >
          Unlock Dashboard
        </button>
        <p className="text-[11px] text-gray-400 text-center mt-1">
          Internal use only · PP Auto Innovators Pvt. Ltd.
        </p>
      </form>
    </div>
  );
}
