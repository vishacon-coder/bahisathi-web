import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://neozddoifbzoztgyvvyy.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_Rb7ovT0DeR6fkqmSR-bpqQ_ad3RbKSv";
// The bill-extraction endpoint added to the already-deployed WhatsApp bot
// server, so the Anthropic API key stays safely server-side.
const EXTRACT_API_URL = "https://tivaro-whatsapp-server.onrender.com/api/extract-bill";

const C = {
  paper: "#F1EAD6",
  paperLine: "#D9CBA3",
  ink: "#1E3358",
  inkLight: "#3C557E",
  red: "#A63A3A",
  gold: "#B98A2E",
  text: "#2A2419",
  textMuted: "#6B6350",
  green: "#3F6B4A",
  amber: "#B27A1E",
  white: "#FBF8F0",
  cardBorder: "#E4D7B0",
};
const display = "'Fraunces', serif";
const sans = "'IBM Plex Sans', sans-serif";
const mono = "'IBM Plex Mono', monospace";

// ---------- Auth (plain REST calls to Supabase, real localStorage is fine
// here since this is a normal deployed website, not a sandboxed artifact) ----------
async function authSignUp(email, password) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/signup`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ email, password }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.msg || data.error_description || "Sign up failed");
  return data;
}
async function authSignIn(email, password) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ email, password }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.msg || data.error_description || "Login failed");
  return data;
}
async function authSignOut(accessToken) {
  try {
    await fetch(`${SUPABASE_URL}/auth/v1/logout`, {
      method: "POST",
      headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
    });
  } catch (e) {}
}
async function billsList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bills?select=*&order=created_at.asc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load bills");
  return res.json();
}
async function billsInsert(accessToken, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bills`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify(row),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.message || "Failed to save bill");
  return data[0];
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function extractBillViaApi(base64, mediaType) {
  const res = await fetch(EXTRACT_API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ base64, mediaType }),
  });
  if (!res.ok) throw new Error("Extraction failed");
  return res.json();
}

const CATEGORY_KEYS = ["raw_materials", "utilities", "rent", "salaries", "transport", "office_supplies", "food", "other"];
const CATEGORY_LABELS = {
  raw_materials: "Raw materials",
  utilities: "Utilities",
  rent: "Rent",
  salaries: "Salaries",
  transport: "Transport",
  office_supplies: "Office supplies",
  food: "Food",
  other: "Other",
};

function ConfidenceDot({ score }) {
  const color = score == null ? C.textMuted : score >= 0.8 ? C.green : score >= 0.5 ? C.amber : C.red;
  return <span style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: color, marginRight: 6, flexShrink: 0 }} />;
}

// ---------- Nav / layout ----------
function NavBar({ page, setPage, loggedIn }) {
  const links = [
    ["home", "Home"],
    ["features", "Features"],
    ["about", "About"],
    ["app", loggedIn ? "My Ledger" : "Try it"],
  ];
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 28px", background: C.ink, flexWrap: "wrap", gap: 12 }}>
      <div onClick={() => setPage("home")} style={{ display: "flex", alignItems: "center", gap: 10, fontFamily: display, fontWeight: 600, fontSize: 21, color: C.white, cursor: "pointer" }}>
        <img src="/logo.png" alt="BahiSathi logo" style={{ height: 34, width: "auto" }} />
        BahiSathi
      </div>
      <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
        {links.map(([key, label]) => (
          <button key={key} onClick={() => setPage(key)} style={{ background: page === key ? C.paper : "transparent", color: page === key ? C.ink : "#C7CFDF", border: "none", borderRadius: 20, padding: "7px 15px", fontSize: 13.5, fontFamily: sans, fontWeight: 500, cursor: "pointer" }}>
            {label}
          </button>
        ))}
      </div>
      {!loggedIn && (
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => setPage("app")} style={{ background: C.gold, border: "none", color: C.white, borderRadius: 20, padding: "7px 16px", fontSize: 13.5, fontFamily: sans, fontWeight: 500, cursor: "pointer" }}>
            Log in / Sign up
          </button>
        </div>
      )}
    </div>
  );
}

function Section({ children, style }) {
  return <div style={{ padding: "56px 28px", maxWidth: 920, margin: "0 auto", ...style }}>{children}</div>;
}

function Footer() {
  const socials = [
    { label: "Instagram", href: "#" },
    { label: "LinkedIn", href: "#" },
    { label: "X", href: "#" },
    { label: "WhatsApp", href: "#" },
    { label: "Email", href: "mailto:hello@bahisathi.in" },
  ];
  return (
    <div style={{ background: C.ink, padding: "28px 28px", marginTop: 20 }}>
      <div style={{ maxWidth: 920, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 14 }}>
        <div style={{ color: "#B9C3D6", fontSize: 12.5 }}>© 2026 Tivaro LLP · BahiSathi · Made in India</div>
        <div style={{ display: "flex", gap: 16 }}>
          {socials.map((s) => (
            <a key={s.label} href={s.href} style={{ color: "#C7CFDF", fontSize: 12.5, textDecoration: "none", fontFamily: sans }}>
              {s.label}
            </a>
          ))}
        </div>
      </div>
    </div>
  );
}

// ---------- Marketing pages ----------
const FEATURES = [
  { icon: "🧾", title: "Bill photo to ledger", body: "Snap any bill and BahiSathi reads the vendor, GST split, and total straight into your books." },
  { icon: "💬", title: "WhatsApp-native", body: "Send a photo on WhatsApp. No app to open, no menu to learn - just chat." },
  { icon: "🌐", title: "Hindi, not just English", body: "Confirm bills and ask about your money in Hindi, with more languages on the way." },
  { icon: "✓", title: "Honest about uncertainty", body: "When a handwritten bill is unclear, BahiSathi tells you - instead of guessing quietly." },
  { icon: "📤", title: "Tally and Excel ready", body: "One tap export in a format your CA or Tally already understands." },
  { icon: "🗣️", title: "Zero accounting jargon", body: `Confirm bills the way you'd say it out loud: "₹1,350 at Sharma Traders - correct?"` },
];

function HomePage({ setPage }) {
  return (
    <div>
      <Section style={{ paddingBottom: 20 }}>
        <div style={{ display: "flex", gap: 40, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ flex: "1 1 380px" }}>
            <div style={{ fontFamily: display, fontWeight: 600, fontSize: 40, lineHeight: 1.15, color: C.text }}>
              Your bahi khata, read by AI.
            </div>
            <div style={{ fontSize: 16, color: C.textMuted, marginTop: 14, lineHeight: 1.6 }}>
              Snap a bill, get GST-ready books - in Hindi or English, on WhatsApp or the web. No accounting degree required.
            </div>
            <div style={{ display: "flex", gap: 12, marginTop: 26, flexWrap: "wrap" }}>
              <button onClick={() => setPage("app")} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "12px 22px", fontSize: 14.5, fontFamily: sans, fontWeight: 500, cursor: "pointer" }}>
                Try it free
              </button>
            </div>
          </div>
          <div style={{ flex: "1 1 300px" }}>
            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, position: "relative", paddingLeft: 22, overflow: "hidden" }}>
              <div style={{ position: "absolute", left: 14, top: 0, bottom: 0, width: 1.5, background: C.red, opacity: 0.55 }} />
              <div style={{ padding: "18px 18px 18px 8px" }}>
                <div style={{ background: "#FDF3DC", border: `1px solid ${C.gold}`, borderRadius: 6, padding: "10px 14px", fontSize: 13.5, color: "#5C4413", marginBottom: 14 }}>
                  ₹1,350 spent at Sharma Traders on 05-01-2026 - correct?
                </div>
                {[["Vendor", "Sharma Traders"], ["Date", "05-01-2026"], ["Category", "Raw materials"]].map(([l, v]) => (
                  <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
                    <span style={{ color: C.textMuted }}>{l}</span>
                    <span style={{ color: C.text }}>{v}</span>
                  </div>
                ))}
                <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 0 0" }}>
                  <span style={{ color: C.ink, fontWeight: 500, fontSize: 13.5 }}>Total</span>
                  <span style={{ fontFamily: mono, color: C.ink, fontWeight: 600, fontSize: 17 }}>₹1,350</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </Section>
    </div>
  );
}

function FeaturesPage() {
  return (
    <Section>
      <div style={{ fontFamily: display, fontSize: 28, fontWeight: 600, marginBottom: 8 }}>Built for how you actually work</div>
      <div style={{ fontSize: 14.5, color: C.textMuted, marginBottom: 30 }}>Not another ledger screen - a bookkeeper that reads bills and speaks plainly.</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 16 }}>
        {FEATURES.map((f) => (
          <div key={f.title} style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px" }}>
            <div style={{ fontSize: 24, marginBottom: 10 }}>{f.icon}</div>
            <div style={{ fontWeight: 500, fontSize: 15, marginBottom: 6 }}>{f.title}</div>
            <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.55 }}>{f.body}</div>
          </div>
        ))}
      </div>
    </Section>
  );
}

function AboutPage() {
  return (
    <Section>
      <div style={{ fontFamily: display, fontSize: 28, fontWeight: 600, marginBottom: 18 }}>About BahiSathi</div>
      <div style={{ fontSize: 15, color: C.text, lineHeight: 1.75, maxWidth: 680 }}>
        <p>BahiSathi is built by Tivaro LLP - AI-native accounting for India's small businesses, starting where bookkeeping hurts the most: reading bills.</p>
        <p>Most accounting software was built for people who already think in ledgers and debit-credit. We're building for the shop owner, trader, and service provider who never needed an accounting degree to run their business well, and shouldn't need one to keep their books.</p>
      </div>
    </Section>
  );
}

// ---------- Auth gate ----------
function AuthGate({ onAuthed }) {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [checkEmail, setCheckEmail] = useState(false);

  async function handleSubmit() {
    setLoading(true);
    setError(null);
    try {
      const data = mode === "login" ? await authSignIn(email, password) : await authSignUp(email, password);
      if (data.access_token) {
        localStorage.setItem("bahisathi_session", JSON.stringify({ access_token: data.access_token, refresh_token: data.refresh_token, user: data.user }));
        onAuthed({ access_token: data.access_token, user: data.user });
      } else {
        setCheckEmail(true);
      }
    } catch (err) {
      setError(err.message || "Something went wrong. Try again.");
    } finally {
      setLoading(false);
    }
  }

  if (checkEmail) {
    return (
      <Section style={{ maxWidth: 400, textAlign: "center" }}>
        <div style={{ fontFamily: display, fontSize: 20, fontWeight: 600, marginBottom: 8 }}>Check your email</div>
        <div style={{ fontSize: 13.5, color: C.textMuted }}>We sent a confirmation link to {email}. Click it, then come back and log in.</div>
      </Section>
    );
  }

  return (
    <Section style={{ maxWidth: 380 }}>
      <div style={{ fontFamily: display, fontSize: 24, fontWeight: 600, marginBottom: 4 }}>Welcome to BahiSathi</div>
      <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 22 }}>Log in or create a free account to keep your ledger.</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <input type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} style={{ border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 12px", fontSize: 14, background: C.white, color: C.text, outline: "none" }} />
        <input
          type="password"
          placeholder="Password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && email && password && !loading) handleSubmit(); }}
          style={{ border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 12px", fontSize: 14, background: C.white, color: C.text, outline: "none" }}
        />
        {error && <div style={{ color: C.red, fontSize: 12.5 }}>{error}</div>}
        <button type="button" onClick={handleSubmit} disabled={loading || !email || !password} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "11px 0", fontSize: 14, fontWeight: 500, cursor: loading ? "default" : "pointer", opacity: loading || !email || !password ? 0.7 : 1 }}>
          {mode === "login" ? "Log in" : "Create account"}
        </button>
      </div>
      <button onClick={() => setMode(mode === "login" ? "signup" : "login")} style={{ marginTop: 14, background: "transparent", border: "none", color: C.inkLight, fontSize: 12.5, cursor: "pointer", textDecoration: "underline" }}>
        {mode === "login" ? "New here? Create an account" : "Already have an account? Log in"}
      </button>
    </Section>
  );
}

// ---------- Main product (scan + ledger) ----------
function ProductApp({ session, onLogout }) {
  const [tab, setTab] = useState("scan");
  const [step, setStep] = useState("upload");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [draft, setDraft] = useState(null);
  const [entries, setEntries] = useState([]);
  const [entriesLoaded, setEntriesLoaded] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const data = await billsList(session.access_token);
        setEntries(
          data.map((row) => ({ id: row.id, vendor: row.vendor, date: row.bill_date, gstin: row.gstin, cgst: row.cgst, sgst: row.sgst, igst: row.igst, total: row.total, category: row.category, vendor_source: row.vendor_source, confidence: row.confidence }))
        );
      } catch (e) {
        console.error(e);
      }
      setEntriesLoaded(true);
    })();
  }, [session]);

  async function processFile(file) {
    setError(null);
    setLoading(true);
    setStep("processing");
    try {
      const base64 = await fileToBase64(file);
      const extracted = await extractBillViaApi(base64, file.type || "image/jpeg");
      setDraft({ ...extracted, items: extracted.items || [] });
      setStep("review");
    } catch (e) {
      setError("Could not read this bill. Try a clearer photo.");
      setStep("upload");
    } finally {
      setLoading(false);
    }
  }

  function handleFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (file) processFile(file);
    e.target.value = "";
  }

  function updateDraft(field, value) {
    setDraft((d) => ({ ...d, [field]: value }));
  }

  async function saveDraft() {
    try {
      const saved = await billsInsert(session.access_token, {
        user_id: session.user.id,
        vendor: draft.vendor,
        bill_date: draft.date,
        gstin: draft.gstin,
        items: draft.items,
        cgst: draft.cgst,
        sgst: draft.sgst,
        igst: draft.igst,
        total: draft.total,
        category: draft.category,
        vendor_source: draft.vendor_source,
        confidence: draft.confidence,
      });
      setEntries((prev) => [...prev, { id: saved.id, vendor: saved.vendor, date: saved.bill_date, gstin: saved.gstin, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst, total: saved.total, category: saved.category, vendor_source: saved.vendor_source, confidence: saved.confidence }]);
    } catch (e) {
      setError("Could not save this bill - please try again.");
      return;
    }
    setDraft(null);
    setStep("upload");
  }

  function discardDraft() {
    setDraft(null);
    setStep("upload");
  }

  function exportToExcel() {
    const rows = entries.map((e) => ({
      Date: e.date || "",
      Vendor: e.vendor || "",
      GSTIN: e.gstin || "",
      Category: CATEGORY_LABELS[e.category] || e.category || "",
      CGST: e.cgst ?? "",
      SGST: e.sgst ?? "",
      IGST: e.igst ?? "",
      Total: e.total ?? "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ledger");
    XLSX.writeFile(wb, "bahisathi-ledger.xlsx");
  }

  const monthTotal = entries.reduce((sum, e) => sum + (Number(e.total) || 0), 0);

  return (
    <div>
      <div style={{ display: "flex", background: C.ink, paddingLeft: 24, gap: 4, justifyContent: "space-between", alignItems: "center", paddingRight: 24 }}>
        <div style={{ display: "flex", gap: 4 }}>
          {["scan", "ledger"].map((key) => (
            <button key={key} onClick={() => setTab(key)} style={{ background: tab === key ? C.paper : "transparent", color: tab === key ? C.ink : "#B9C3D6", border: "none", borderRadius: "8px 8px 0 0", padding: "9px 18px", fontSize: 13.5, fontWeight: 500, cursor: "pointer", fontFamily: sans }}>
              {key === "scan" ? "Scan bill" : "Ledger"}
            </button>
          ))}
        </div>
        <button onClick={onLogout} style={{ background: "transparent", border: "1px solid #5C6B87", color: "#B9C3D6", borderRadius: 20, padding: "5px 14px", fontSize: 12.5, cursor: "pointer", fontFamily: sans }}>
          Log out
        </button>
      </div>

      <div style={{ padding: 24, maxWidth: 640, margin: "0 auto" }}>
        {tab === "scan" && step === "upload" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Snap or upload a bill</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18 }}>Photo of any invoice, receipt, or bill</div>
            <div onClick={() => fileInputRef.current && fileInputRef.current.click()} style={{ border: `2px dashed ${C.gold}`, borderRadius: 8, padding: "44px 20px", textAlign: "center", cursor: "pointer", background: C.white }}>
              <div style={{ fontSize: 32, marginBottom: 10 }}>🧾</div>
              <div style={{ fontSize: 14.5, fontWeight: 500, color: C.ink }}>Click to choose a photo</div>
            </div>
            <input ref={fileInputRef} type="file" accept="image/*" onChange={handleFileChange} style={{ display: "none" }} />
            {error && <div style={{ marginTop: 14, color: C.red, fontSize: 13 }}>{error}</div>}
          </div>
        )}

        {tab === "scan" && step === "processing" && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ width: 34, height: 34, border: `3px solid ${C.paperLine}`, borderTopColor: C.gold, borderRadius: "50%", margin: "0 auto 16px", animation: "spin 0.9s linear infinite" }} />
            <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
            <div style={{ fontSize: 14, color: C.textMuted }}>Reading your bill…</div>
          </div>
        )}

        {tab === "scan" && step === "review" && draft && (
          <div>
            <div style={{ background: "#FDF3DC", border: `1px solid ${C.gold}`, borderRadius: 6, padding: "10px 14px", fontSize: 13.5, marginBottom: 16, color: "#5C4413" }}>
              ₹{draft.total ?? "?"} spent at {draft.vendor || "this vendor"} on {draft.date || "this date"} - correct?
            </div>
            {(draft.vendor_source === "handwritten_only" || draft.vendor_source === "none") && (
              <div style={{ fontSize: 12, color: C.red, marginBottom: 10 }}>⚠️ Only handwriting found - please check this vendor name carefully</div>
            )}
            {draft.vendor_source === "printed_unclear" && (
              <div style={{ fontSize: 12, color: C.amber, marginBottom: 10 }}>⚠️ Printed name was unclear - double-check this</div>
            )}
            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, position: "relative", paddingLeft: 22, overflow: "hidden" }}>
              <div style={{ position: "absolute", left: 14, top: 0, bottom: 0, width: 1.5, background: C.red, opacity: 0.55 }} />
              <div style={{ padding: "16px 18px 16px 6px" }}>
                {[
                  ["vendor", "Vendor", draft.confidence?.vendor],
                  ["date", "Date", draft.confidence?.date],
                  ["gstin", "GSTIN", draft.confidence?.gstin],
                ].map(([field, label, conf]) => (
                  <div key={field} style={{ display: "flex", alignItems: "center", padding: "10px 0", borderBottom: `1px solid ${C.paperLine}`, gap: 12 }}>
                    <div style={{ width: 90, flexShrink: 0, fontSize: 12, color: C.textMuted, display: "flex", alignItems: "center" }}>
                      <ConfidenceDot score={conf} />
                      {label}
                    </div>
                    <input value={draft[field] ?? ""} onChange={(e) => updateDraft(field, e.target.value)} style={{ flex: 1, border: "none", background: "transparent", fontSize: 14, color: C.text, outline: "none" }} />
                  </div>
                ))}
                <div style={{ display: "flex", alignItems: "center", padding: "10px 0", gap: 12 }}>
                  <div style={{ width: 90, flexShrink: 0, fontSize: 12.5, color: C.ink, fontWeight: 500 }}>Total</div>
                  <input value={draft.total ?? ""} onChange={(e) => updateDraft("total", e.target.value)} style={{ flex: 1, border: "none", background: "transparent", fontSize: 17, fontWeight: 600, color: C.ink, fontFamily: mono, outline: "none" }} />
                </div>
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
              <button onClick={saveDraft} style={{ flex: 1, background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "11px 0", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>
                Save to ledger
              </button>
              <button onClick={discardDraft} style={{ background: "transparent", color: C.red, border: `1px solid ${C.red}`, borderRadius: 6, padding: "11px 18px", fontSize: 14, cursor: "pointer" }}>
                Discard
              </button>
            </div>
          </div>
        )}

        {tab === "ledger" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 14 }}>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600 }}>This month</div>
              <div style={{ fontFamily: mono, fontSize: 18, color: C.ink }}>₹{monthTotal.toLocaleString("en-IN")}</div>
            </div>
            {entriesLoaded && entries.length === 0 && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>No bills saved yet. Scan your first one.</div>}
            {entries.length > 0 && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, position: "relative", paddingLeft: 22, overflow: "hidden" }}>
                <div style={{ position: "absolute", left: 14, top: 0, bottom: 0, width: 1.5, background: C.red, opacity: 0.55 }} />
                <div style={{ padding: "16px 18px 16px 6px" }}>
                  {entries.slice().reverse().map((e, i) => (
                    <div key={e.id || i} style={{ display: "flex", justifyContent: "space-between", padding: "9px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
                      <div style={{ color: C.textMuted, width: 78, flexShrink: 0 }}>{e.date || "—"}</div>
                      <div style={{ flex: 1 }}>
                        <div style={{ color: C.text }}>{e.vendor || "—"}</div>
                        <div style={{ fontSize: 11, color: C.textMuted }}>{CATEGORY_LABELS[e.category] || e.category}</div>
                      </div>
                      <div style={{ fontFamily: mono, color: C.ink, fontWeight: 500 }}>₹{Number(e.total || 0).toLocaleString("en-IN")}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {entries.length > 0 && (
              <button onClick={exportToExcel} style={{ marginTop: 16, background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "10px 20px", fontSize: 13.5, fontWeight: 500, cursor: "pointer", width: "100%" }}>
                Export to Excel
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [page, setPage] = useState("home");
  const [session, setSession] = useState(undefined);

  useEffect(() => {
    const stored = localStorage.getItem("bahisathi_session");
    setSession(stored ? JSON.parse(stored) : null);
  }, []);

  async function handleLogout() {
    if (session) await authSignOut(session.access_token);
    localStorage.removeItem("bahisathi_session");
    setSession(null);
  }

  return (
    <div style={{ fontFamily: sans, background: C.paper, color: C.text, minHeight: "100vh" }}>
      <NavBar page={page} setPage={setPage} loggedIn={!!session} />
      {page === "home" && <HomePage setPage={setPage} />}
      {page === "features" && <FeaturesPage />}
      {page === "about" && <AboutPage />}
      {page === "app" && (
        session === undefined ? (
          <Section style={{ textAlign: "center", color: C.textMuted }}>Loading…</Section>
        ) : session ? (
          <ProductApp session={session} onLogout={handleLogout} />
        ) : (
          <AuthGate onAuthed={(s) => setSession(s)} />
        )
      )}
      <Footer />
    </div>
  );
}
