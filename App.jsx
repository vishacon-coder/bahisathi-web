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

// Real flag artwork instead of the 🇮🇳 emoji — Windows has no glyph for
// regional-indicator flag emoji and falls back to showing "IN" as text.
function IndiaFlag({ size = 14, style = {} }) {
  return (
    <svg
      width={size}
      height={size * 0.667}
      viewBox="0 0 900 600"
      style={{ display: "inline-block", verticalAlign: "middle", borderRadius: 2, flex: "none", ...style }}
      aria-label="India"
    >
      <rect width="900" height="200" y="0" fill="#FF9933" />
      <rect width="900" height="200" y="200" fill="#FFFFFF" />
      <rect width="900" height="200" y="400" fill="#138808" />
      <circle cx="450" cy="300" r="80" fill="none" stroke="#000080" strokeWidth="4" />
      <circle cx="450" cy="300" r="6" fill="#000080" />
      {Array.from({ length: 24 }).map((_, i) => (
        <line
          key={i}
          x1="450"
          y1="300"
          x2={450 + 80 * Math.cos((i * Math.PI) / 12)}
          y2={300 + 80 * Math.sin((i * Math.PI) / 12)}
          stroke="#000080"
          strokeWidth="2"
        />
      ))}
    </svg>
  );
}

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
async function authRefreshToken(refreshToken) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ refresh_token: refreshToken }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.msg || data.error_description || "Session refresh failed");
  return data;
}
// Reads the "exp" (expiry, unix seconds) claim out of a Supabase access
// token so we know exactly when to refresh it - no need to guess a TTL.
function parseJwtExp(token) {
  try {
    const payload = token.split(".")[1];
    const json = JSON.parse(atob(payload.replace(/-/g, "+").replace(/_/g, "/")));
    return json.exp || null;
  } catch (e) {
    return null;
  }
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
async function billsUpdate(accessToken, id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bills?id=eq.${id}`, {
    method: "PATCH",
    headers: {
      apikey: SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify(patch),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to update bill");
  return data[0];
}
async function billsDelete(accessToken, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bills?id=eq.${id}`, {
    method: "DELETE",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.message || "Failed to delete bill");
  }
}

// Business profile is optional (see Terms/Privacy) - collected after signup,
// not required to use the app. Stored one row per user (user_id is the
// primary key), so a save is just an upsert.
async function businessProfileGet(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/business_profiles?select=*&limit=1`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load business profile");
  const data = await res.json();
  return data[0] || null;
}
async function businessProfileSave(accessToken, userId, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/business_profiles`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Prefer: "resolution=merge-duplicates,return=representation",
    },
    body: JSON.stringify({ user_id: userId, ...patch, updated_at: new Date().toISOString() }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to save business profile");
  return data[0];
}

// Flags a likely-duplicate bill: same vendor (case/space-insensitive) and the
// same total (within a rupee, to allow for rounding), regardless of date -
// this is the common accident (re-scanning the same photo, or a bill sent on
// both WhatsApp and the web before linking).
function findPossibleDuplicate(draft, entries) {
  if (!draft || !draft.vendor || draft.total == null) return null;
  const vendor = String(draft.vendor).trim().toLowerCase();
  const total = Number(draft.total);
  if (!vendor || Number.isNaN(total)) return null;
  return (
    entries.find((e) => {
      if (!e.vendor || e.total == null) return false;
      const eVendor = String(e.vendor).trim().toLowerCase();
      const eTotal = Number(e.total);
      return eVendor === vendor && !Number.isNaN(eTotal) && Math.abs(eTotal - total) < 1;
    }) || null
  );
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function extractBillViaApi(base64, mediaType, accessToken) {
  const res = await fetch(EXTRACT_API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify({ base64, mediaType }),
  });
  if (!res.ok) {
    let payload = null;
    try { payload = await res.json(); } catch (e) {}
    const err = new Error((payload && payload.message) || "Extraction failed");
    err.code = payload && payload.error;
    err.status = res.status;
    throw err;
  }
  return res.json();
}

const API_BASE = EXTRACT_API_URL.replace(/\/api\/extract-bill$/, "");

async function apiCall(path, accessToken, body) {
  const res = await fetch(`${API_BASE}${path}`, {
    method: body ? "POST" : "GET",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Request failed");
  return data;
}

const getSubscriptionStatus = (accessToken) => apiCall("/api/subscription-status", accessToken);
const getSpendSummary = (accessToken, question) => apiCall("/api/spend-summary", accessToken, { question });
const createProOrder = (accessToken, cycle) => apiCall("/api/create-order", accessToken, { cycle });
const verifyProPayment = (accessToken, payload) => apiCall("/api/verify-payment", accessToken, payload);
const startWhatsAppLink = (accessToken, phone) => apiCall("/api/whatsapp/link-start", accessToken, { phone });
const verifyWhatsAppLink = (accessToken, otp) => apiCall("/api/whatsapp/link-verify", accessToken, { otp });
async function unlinkWhatsApp(accessToken, phone) {
  const res = await fetch(`${API_BASE}/api/whatsapp/link`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify({ phone }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Request failed");
  return data;
}

// Single source of truth for Pro pricing, shared by the public Plans page and
// the in-app upgrade flow. Both prices are final and GST-inclusive.
const PRO_PRICING = {
  monthly: { amount: "₹199", per: "/month", note: "Billed every month. Switch to annual anytime and save." },
  annual: { amount: "₹1,990", per: "/year", note: "Works out to ~₹166/month — 2 months free versus paying monthly." },
};

// Loads the Razorpay Checkout script once and reuses it on later calls.
let razorpayScriptPromise = null;
function loadRazorpayScript() {
  if (window.Razorpay) return Promise.resolve();
  if (!razorpayScriptPromise) {
    razorpayScriptPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://checkout.razorpay.com/v1/checkout.js";
      script.onload = resolve;
      script.onerror = () => reject(new Error("Could not load payment gateway. Check your connection and try again."));
      document.body.appendChild(script);
    });
  }
  return razorpayScriptPromise;
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
  mixed: "Mixed",
};

const BUSINESS_TYPE_KEYS = ["retail", "wholesale", "service", "manufacturing", "trading", "other"];
const BUSINESS_TYPE_LABELS = {
  retail: "Retail / shop",
  wholesale: "Wholesale / distributor",
  service: "Service provider",
  manufacturing: "Manufacturer",
  trading: "Trader",
  other: "Other",
};

// A bill's line items can span more than one expense head (e.g. hardware
// plus a delivery charge). This rolls that up to a single label for the
// ledger list and CSV/Excel export - "Mixed" when items disagree, otherwise
// whichever one category everyone agrees on.
function deriveCategory(items) {
  if (!Array.isArray(items) || items.length === 0) return "other";
  const distinct = new Set(items.map((it) => it.category || "other"));
  return distinct.size === 1 ? [...distinct][0] : "mixed";
}

function ItemsEditor({ items, onChange }) {
  function updateItem(i, field, value) {
    const next = items.slice();
    next[i] = { ...next[i], [field]: value };
    onChange(next);
  }
  function removeItem(i) {
    onChange(items.filter((_, idx) => idx !== i));
  }
  function addItem() {
    onChange([...items, { description: "", amount: "", category: "other" }]);
  }
  const sum = items.reduce((s, it) => s + (Number(it.amount) || 0), 0);
  return (
    <div>
      {items.map((it, i) => (
        <div key={i} style={{ display: "flex", gap: 6, marginBottom: 6, alignItems: "center" }}>
          <input value={it.description || ""} onChange={(e) => updateItem(i, "description", e.target.value)} placeholder="Item" style={{ flex: 1, minWidth: 0, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 9px", fontSize: 12.5, background: C.white, color: C.text, outline: "none" }} />
          <input value={it.amount ?? ""} onChange={(e) => updateItem(i, "amount", e.target.value)} placeholder="₹" style={{ width: 70, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 9px", fontSize: 12.5, fontFamily: mono, background: C.white, color: C.text, outline: "none" }} />
          <select value={it.category || "other"} onChange={(e) => updateItem(i, "category", e.target.value)} style={{ width: 118, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 4px", fontSize: 11.5, background: C.white, color: C.text, outline: "none" }}>
            {CATEGORY_KEYS.map((k) => (
              <option key={k} value={k}>{CATEGORY_LABELS[k]}</option>
            ))}
          </select>
          <button onClick={() => removeItem(i)} style={{ background: "transparent", border: "none", color: C.red, cursor: "pointer", fontSize: 16, lineHeight: 1, padding: "0 2px" }}>×</button>
        </div>
      ))}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
        <button onClick={addItem} style={{ background: "transparent", border: "none", color: C.ink, textDecoration: "underline", cursor: "pointer", fontSize: 12.5, fontFamily: sans, padding: 0 }}>
          + Add item
        </button>
        <div style={{ fontSize: 11.5, color: C.textMuted, fontFamily: mono }}>Items total: ₹{sum.toLocaleString("en-IN")}</div>
      </div>
    </div>
  );
}

function ConfidenceDot({ score }) {
  const color = score == null ? C.textMuted : score >= 0.8 ? C.green : score >= 0.5 ? C.amber : C.red;
  return <span style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: color, marginRight: 6, flexShrink: 0 }} />;
}

// ---------- Nav / layout ----------
function NavBar({ page, setPage, loggedIn }) {
  const links = [
    ["home", "Home"],
    ["features", "Features"],
    ["plans", "Plans"],
    ["about", "About"],
    ["app", loggedIn ? "My Ledger" : "Try it"],
  ];
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 28px", background: C.ink, flexWrap: "wrap", gap: 12, position: "sticky", top: 0, zIndex: 50, boxShadow: "0 2px 14px rgba(10,20,40,0.35)" }}>
      <div onClick={() => setPage("home")} style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }}>
        <div style={{ background: C.paper, borderRadius: 12, padding: 5, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 2px 8px rgba(0,0,0,0.25)" }}>
          <img src="/logo.png" alt="BahiSathi logo" style={{ height: 38, width: "auto", display: "block" }} />
        </div>
        <div style={{ fontFamily: display, fontWeight: 700, fontSize: 25, letterSpacing: 0.3, lineHeight: 1 }}>
          <span style={{ color: "#E8B44C" }}>Bahi</span>
          <span style={{ color: C.white }}>Sathi</span>
          <div style={{ fontFamily: sans, fontWeight: 400, fontSize: 10.5, color: "#8FA0BE", letterSpacing: 1.6, marginTop: 3, textTransform: "uppercase" }}>Apna khata, apni bhasha</div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
        {links.map(([key, label]) => (
          <button key={key} className="navlink" onClick={() => setPage(key)} style={{ background: page === key ? C.paper : "transparent", color: page === key ? C.ink : "#C7CFDF", border: "none", borderRadius: 20, padding: "8px 16px", fontSize: 13.5, fontFamily: sans, fontWeight: 500, cursor: "pointer" }}>
            {label}
          </button>
        ))}
      </div>
      {!loggedIn && (
        <div style={{ display: "flex", gap: 8 }}>
          <button className="cta" onClick={() => setPage("app")} style={{ background: "linear-gradient(135deg, #C9962F, #E8B44C)", border: "none", color: "#2A2005", borderRadius: 20, padding: "9px 18px", fontSize: 13.5, fontFamily: sans, fontWeight: 600, cursor: "pointer", boxShadow: "0 2px 10px rgba(232,180,76,0.35)" }}>
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

// Lightweight overlay used to show Terms/Privacy without losing in-progress
// form state (the app has no URL routing, so navigating pages would do that).
function Modal({ onClose, children }) {
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(20,25,40,0.55)", zIndex: 200, display: "flex", justifyContent: "center", padding: "40px 16px", overflowY: "auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.paper, borderRadius: 12, maxWidth: 760, width: "100%", height: "fit-content", position: "relative", boxShadow: "0 20px 60px rgba(0,0,0,0.4)" }}>
        <button onClick={onClose} style={{ position: "absolute", top: 10, right: 12, background: "transparent", border: "none", fontSize: 26, lineHeight: 1, cursor: "pointer", color: C.textMuted }}>×</button>
        {children}
      </div>
    </div>
  );
}

function LegalSection({ title, children }) {
  return (
    <div style={{ marginBottom: 26 }}>
      <div style={{ fontFamily: display, fontSize: 19, fontWeight: 700, marginBottom: 8, color: C.text }}>{title}</div>
      <div style={{ fontSize: 14, color: C.text, lineHeight: 1.7 }}>{children}</div>
    </div>
  );
}

function PrivacyPage() {
  return (
    <Section style={{ maxWidth: 760 }}>
      <div style={{ fontFamily: display, fontSize: 32, fontWeight: 700, marginBottom: 4 }}>Privacy Policy</div>
      <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 28 }}>Effective 11 August 2026 · BahiSathi, a product of Tivaro LLP</div>

      <LegalSection title="What we collect">
        <p style={{ marginTop: 0 }}>When you use BahiSathi we collect: your email address and password (web accounts); your WhatsApp phone number (if you use the WhatsApp bot); the bill details extracted from photos you send us (vendor, date, GSTIN, amounts, category); and basic usage counts (how many bills you scan each month, used to apply free-plan limits).</p>
        <p>We do <b>not</b> store the bill photos themselves. A photo is processed once to read its contents and is not saved on our servers.</p>
      </LegalSection>

      <LegalSection title="How your data is used">
        <p style={{ marginTop: 0 }}>Your data is used only to run BahiSathi: keeping your ledger, showing your entries back to you, exporting them when you ask, applying plan limits, and improving reliability. We do not sell your data, and we do not use your ledger data for advertising.</p>
      </LegalSection>

      <LegalSection title="AI processing">
        <p style={{ marginTop: 0 }}>Bill photos are read using Anthropic's Claude AI service. The image is sent securely to Anthropic's API for one-time processing and the extracted text is returned to us. Anthropic's API does not use customer data sent to it to train its models by default.</p>
      </LegalSection>

      <LegalSection title="Where your data lives">
        <p style={{ marginTop: 0 }}>Ledger entries and account details are stored in our database hosted on Supabase (servers currently located on AWS in the Asia-Pacific region). Passwords are stored only as secure hashes. Each user's entries are isolated with row-level security, so one account can never read another account's data.</p>
      </LegalSection>

      <LegalSection title="Third-party services we rely on">
        <p style={{ marginTop: 0 }}>Supabase (database and login), Anthropic (AI bill reading), Meta WhatsApp Business Platform (WhatsApp messages), and our hosting providers (Hostinger and Render). Each processes only what is needed for its function.</p>
      </LegalSection>

      <LegalSection title="Cookies and local storage">
        <p style={{ marginTop: 0 }}>The website stores your login session in your browser's local storage so you stay signed in. We do not use advertising or tracking cookies.</p>
      </LegalSection>

      <LegalSection title="Your rights">
        <p style={{ marginTop: 0 }}>You can ask us at any time to show you the data we hold about you, correct it, export it, or delete your account and all its entries. Write to <a href="mailto:hello@bahisathi.in" style={{ color: C.red }}>hello@bahisathi.in</a> and we will act within 30 days. These rights align with India's Digital Personal Data Protection Act, 2023.</p>
      </LegalSection>

      <LegalSection title="Grievances and changes">
        <p style={{ marginTop: 0 }}>For any privacy concern, contact <a href="mailto:hello@bahisathi.in" style={{ color: C.red }}>hello@bahisathi.in</a>. If we make material changes to this policy we will post the updated version here with a new effective date.</p>
      </LegalSection>
    </Section>
  );
}

function TermsPage() {
  return (
    <Section style={{ maxWidth: 760 }}>
      <div style={{ fontFamily: display, fontSize: 32, fontWeight: 700, marginBottom: 4 }}>Terms of Service</div>
      <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 28 }}>Effective 11 August 2026 · BahiSathi, a product of Tivaro LLP</div>

      <LegalSection title="The service">
        <p style={{ marginTop: 0 }}>BahiSathi is an AI bookkeeping assistant: you send a photo of a bill on WhatsApp or the web, and BahiSathi extracts its details into a digital ledger you can review, edit, and export. By creating an account or messaging the bot you agree to these terms.</p>
      </LegalSection>

      <LegalSection title="Your account">
        <p style={{ marginTop: 0 }}>You must be at least 18 and provide accurate information. Keep your password safe - you're responsible for activity on your account. One account is for one business/person unless we agree otherwise.</p>
      </LegalSection>

      <LegalSection title="Free plan and limits">
        <p style={{ marginTop: 0 }}>The free plan currently includes 30 bill scans per calendar month per account (and per WhatsApp number), resetting on the 1st. Limits and plans may change; paid plans will be announced before any free feature is restricted. Fair use applies - automated or bulk abuse of the service may lead to suspension.</p>
      </LegalSection>

      <LegalSection title="Subscription & billing (Pro plan)">
        <p style={{ marginTop: 0 }}>Pro costs ₹199/month or ₹1,990/year, billed upfront for the period you choose. See our <b>Plans</b> page for full pricing and features. <b>These prices are inclusive of GST</b> - the amount shown is exactly what you pay, nothing is added at checkout.</p>
        <p>Pro is a fixed-term purchase, not an auto-renewing subscription: we do not store your card and do not charge you again automatically. When your paid period ends, your account simply returns to the Free plan until you choose to renew from the Plan tab.</p>
        <p>Because there's no recurring charge, "cancelling" just means not renewing - there's nothing to switch off. Payments already made are non-refundable, including if you stop using the service before your paid period ends.</p>
        <p>We may change Pro pricing for future billing periods. Any change will be posted on the Plans page before it applies - it will never change the price of a period you've already paid for.</p>
      </LegalSection>

      <LegalSection title="Accuracy - please read">
        <p style={{ marginTop: 0 }}>BahiSathi uses AI to read bills. AI can make mistakes, especially with handwriting or unclear photos - that's why every entry asks for your confirmation. <b>You are responsible for verifying entries before relying on them.</b> BahiSathi is a bookkeeping aid, not an accountant: it does not provide tax, legal, or financial advice, and its output does not replace your CA or your statutory obligations under GST or other laws.</p>
      </LegalSection>

      <LegalSection title="Acceptable use">
        <p style={{ marginTop: 0 }}>Don't use BahiSathi for anything unlawful, don't upload content you have no right to share, don't attempt to break, overload, or reverse-engineer the service, and don't use it to process another person's data without their permission.</p>
      </LegalSection>

      <LegalSection title="Your data">
        <p style={{ marginTop: 0 }}>Your ledger data belongs to you. You can export it at any time and request deletion of your account. How we handle data is described in our Privacy Policy, which is part of these terms.</p>
      </LegalSection>

      <LegalSection title="Liability">
        <p style={{ marginTop: 0 }}>BahiSathi is provided "as is". To the maximum extent permitted by law, Tivaro LLP is not liable for indirect or consequential losses, or for losses arising from unverified AI-extracted entries, service interruptions, or third-party service failures. Our total liability for any claim is limited to the amount you paid us in the 12 months before the claim (or ₹1,000 if you're on the free plan).</p>
      </LegalSection>

      <LegalSection title="Termination and changes">
        <p style={{ marginTop: 0 }}>You can stop using BahiSathi and delete your account at any time. We may suspend accounts that break these terms. We may update these terms; material changes will be posted here with a new effective date, and continued use means acceptance.</p>
      </LegalSection>

      <LegalSection title="Governing law and contact">
        <p style={{ marginTop: 0 }}>These terms are governed by the laws of India, with courts of competent jurisdiction in India. Questions: <a href="mailto:hello@bahisathi.in" style={{ color: C.red }}>hello@bahisathi.in</a>.</p>
      </LegalSection>
    </Section>
  );
}

function Footer({ setPage }) {
  const socials = [
    { label: "Facebook", href: "https://www.facebook.com/profile.php?id=61592587118534" },
    { label: "Instagram", href: "https://www.instagram.com/bahisathi/" },
    { label: "X", href: "https://x.com/bahisathi" },
    { label: "Email", href: "mailto:hello@bahisathi.in" },
  ];
  return (
    <div style={{ background: C.ink, padding: "40px 28px 26px", marginTop: 20 }}>
      <div style={{ maxWidth: 920, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 26, marginBottom: 28 }}>
          <div style={{ maxWidth: 300 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
              <div style={{ background: C.paper, borderRadius: 10, padding: 4 }}>
                <img src="/logo.png" alt="BahiSathi" style={{ height: 30, display: "block" }} />
              </div>
              <span style={{ fontFamily: display, fontWeight: 700, fontSize: 20 }}>
                <span style={{ color: "#E8B44C" }}>Bahi</span><span style={{ color: C.white }}>Sathi</span>
              </span>
            </div>
            <div style={{ color: "#8FA0BE", fontSize: 12.5, lineHeight: 1.6 }}>AI se bill padho, khata khud bane. Built for India's shop owners, traders, and service providers.</div>
          </div>
          <div style={{ display: "flex", gap: 16, alignItems: "flex-start", flexWrap: "wrap" }}>
            {socials.map((s) => (
              <a
                key={s.label}
                href={s.href}
                target={s.href.startsWith("mailto:") ? undefined : "_blank"}
                rel={s.href.startsWith("mailto:") ? undefined : "noopener noreferrer"}
                style={{ color: "#C7CFDF", fontSize: 12.5, textDecoration: "none", fontFamily: sans }}
              >
                {s.label}
              </a>
            ))}
          </div>
        </div>
        <div style={{ borderTop: "1px solid #33476B", paddingTop: 16, color: "#8FA0BE", fontSize: 12, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
          <span>
            © 2026 <a href="https://tivaro.co.in/" target="_blank" rel="noopener noreferrer" style={{ color: "#C7CFDF", textDecoration: "underline" }}>Tivaro LLP</a> · BahiSathi · Made in India <IndiaFlag size={14} style={{ marginLeft: 2, marginBottom: -1 }} />
          </span>
          <span style={{ display: "flex", gap: 14 }}>
            <button onClick={() => setPage("plans")} style={{ background: "transparent", border: "none", color: "#C7CFDF", fontSize: 12, cursor: "pointer", textDecoration: "underline", fontFamily: sans, padding: 0 }}>Plans & Pricing</button>
            <button onClick={() => setPage("privacy")} style={{ background: "transparent", border: "none", color: "#C7CFDF", fontSize: 12, cursor: "pointer", textDecoration: "underline", fontFamily: sans, padding: 0 }}>Privacy Policy</button>
            <button onClick={() => setPage("terms")} style={{ background: "transparent", border: "none", color: "#C7CFDF", fontSize: 12, cursor: "pointer", textDecoration: "underline", fontFamily: sans, padding: 0 }}>Terms of Service</button>
          </span>
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

function LedgerMockCard() {
  return (
    <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 14, position: "relative", paddingLeft: 22, overflow: "hidden", boxShadow: "0 18px 45px rgba(30,51,88,0.18)" }}>
      <div style={{ position: "absolute", left: 14, top: 0, bottom: 0, width: 1.5, background: C.red, opacity: 0.55 }} />
      <div style={{ padding: "18px 18px 18px 8px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
          <img src="/logo.png" alt="" style={{ height: 20 }} />
          <span style={{ fontSize: 11.5, fontWeight: 600, color: C.textMuted, letterSpacing: 1, textTransform: "uppercase" }}>Auto-filled from photo</span>
        </div>
        {[["Vendor", "Sharma Traders"], ["Date", "05-01-2026"], ["GST (CGST+SGST)", "₹103 + ₹103"], ["Category", "Raw materials"]].map(([l, v]) => (
          <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
            <span style={{ color: C.textMuted }}>{l}</span>
            <span style={{ color: C.text, fontWeight: 500 }}>{v}</span>
          </div>
        ))}
        <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 0 2px" }}>
          <span style={{ color: C.ink, fontWeight: 600, fontSize: 13.5 }}>Total</span>
          <span style={{ fontFamily: mono, color: C.ink, fontWeight: 600, fontSize: 19 }}>₹1,350</span>
        </div>
      </div>
    </div>
  );
}

function WhatsAppBubble() {
  return (
    <div style={{ position: "absolute", top: -26, right: -14, zIndex: 2, maxWidth: 250, background: "#E7FFDB", borderRadius: "14px 14px 4px 14px", padding: "11px 14px", fontSize: 12.5, color: "#1F2C1A", boxShadow: "0 10px 30px rgba(20,60,20,0.25)", border: "1px solid #BCE5A8", transform: "rotate(2deg)" }}>
      <div style={{ fontWeight: 600, fontSize: 11, color: "#128C7E", marginBottom: 3 }}>BahiSathi on WhatsApp</div>
      ₹1,350 spent at Sharma Traders on 05-01-2026 — correct?
      <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
        <span style={{ background: "#128C7E", color: "#fff", borderRadius: 12, padding: "3px 12px", fontSize: 11.5, fontWeight: 600 }}>✓ Yes</span>
        <span style={{ background: "#fff", color: "#128C7E", borderRadius: 12, padding: "3px 12px", fontSize: 11.5, fontWeight: 600, border: "1px solid #128C7E" }}>Edit</span>
      </div>
    </div>
  );
}

const STEPS = [
  { n: "1", title: "Snap the bill", body: "Photo on WhatsApp or upload on the web - printed or handwritten, Hindi or English." },
  { n: "2", title: "Confirm in one tap", body: `BahiSathi reads vendor, GST split and total, then asks plainly: "₹1,350 at Sharma Traders - correct?"` },
  { n: "3", title: "Books stay ready", body: "Every bill lands in a GST-ready ledger you can export to Excel or hand to your CA." },
];

function HomePage({ setPage }) {
  return (
    <div>
      {/* HERO */}
      <div style={{ background: `linear-gradient(180deg, ${C.paper} 0%, #F7F2E3 100%)` }}>
        <Section style={{ paddingBottom: 48, paddingTop: 60 }}>
          <div style={{ display: "flex", gap: 48, flexWrap: "wrap", alignItems: "center" }}>
            <div style={{ flex: "1 1 400px" }}>
              <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "#FBF4DF", border: `1px solid ${C.cardBorder}`, borderRadius: 20, padding: "6px 14px", fontSize: 12.5, fontWeight: 600, color: C.gold, marginBottom: 20 }}>
                <IndiaFlag size={15} /> AI bookkeeping for India's small businesses
              </div>
              <h1 style={{ fontFamily: display, fontWeight: 900, fontSize: "clamp(36px, 5.5vw, 56px)", lineHeight: 1.08, color: C.text, margin: 0 }}>
                Your <span style={{ color: C.red }}>bahi khata</span>,<br />read by AI.
              </h1>
              <div style={{ fontSize: 17, color: C.textMuted, marginTop: 18, lineHeight: 1.65, maxWidth: 460 }}>
                Snap a bill, get GST-ready books — in Hindi or English, on WhatsApp or the web. No accounting degree required.
              </div>
              <div style={{ display: "flex", gap: 12, marginTop: 30, flexWrap: "wrap" }}>
                <button className="cta" onClick={() => setPage("app")} style={{ background: C.green, color: C.white, border: "none", borderRadius: 10, padding: "15px 30px", fontSize: 16, fontFamily: sans, fontWeight: 600, cursor: "pointer", boxShadow: "0 6px 20px rgba(63,107,74,0.35)" }}>
                  Try it free →
                </button>
                <button className="cta" onClick={() => setPage("features")} style={{ background: "transparent", color: C.ink, border: `2px solid ${C.ink}`, borderRadius: 10, padding: "13px 26px", fontSize: 16, fontFamily: sans, fontWeight: 600, cursor: "pointer" }}>
                  See features
                </button>
              </div>
              <div style={{ display: "flex", gap: 10, marginTop: 26, flexWrap: "wrap" }}>
                {["✓ GST-ready", "✓ Hindi + English", "✓ WhatsApp native", "✓ Excel export"].map((t) => (
                  <span key={t} style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 16, padding: "5px 13px", fontSize: 12.5, fontWeight: 500, color: C.green }}>{t}</span>
                ))}
              </div>
            </div>
            <div style={{ flex: "1 1 320px", position: "relative", marginTop: 26 }}>
              <WhatsAppBubble />
              <LedgerMockCard />
            </div>
          </div>
        </Section>
      </div>

      {/* VALUE BAND */}
      <div style={{ background: C.ink }}>
        <div style={{ maxWidth: 920, margin: "0 auto", padding: "34px 28px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 22 }}>
          {[
            ["📸 → 📒", "Photo to ledger entry"],
            ["हिंदी + English", "Speaks your language"],
            ["GST", "CGST / SGST auto-split"],
            ["Excel · Tally", "CA-friendly export"],
          ].map(([big, small]) => (
            <div key={small} style={{ textAlign: "center" }}>
              <div style={{ fontFamily: display, fontWeight: 700, fontSize: 21, color: "#E8B44C", marginBottom: 5 }}>{big}</div>
              <div style={{ fontSize: 12.5, color: "#B9C3D6" }}>{small}</div>
            </div>
          ))}
        </div>
      </div>

      {/* HOW IT WORKS */}
      <Section style={{ paddingBottom: 30 }}>
        <div style={{ textAlign: "center", marginBottom: 38 }}>
          <div style={{ fontFamily: display, fontSize: 32, fontWeight: 700, color: C.text }}>Three steps. Zero jargon.</div>
          <div style={{ fontSize: 15, color: C.textMuted, marginTop: 8 }}>Bookkeeping that works the way you already talk.</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: 18 }}>
          {STEPS.map((s) => (
            <div key={s.n} className="card" style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 14, padding: "26px 24px", position: "relative" }}>
              <div style={{ width: 42, height: 42, borderRadius: "50%", background: `linear-gradient(135deg, ${C.red}, #C4574F)`, color: C.white, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: display, fontWeight: 700, fontSize: 20, marginBottom: 16, boxShadow: "0 4px 12px rgba(166,58,58,0.3)" }}>{s.n}</div>
              <div style={{ fontWeight: 600, fontSize: 16.5, marginBottom: 8, color: C.text }}>{s.title}</div>
              <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.6 }}>{s.body}</div>
            </div>
          ))}
        </div>
      </Section>

      {/* FEATURE PREVIEW */}
      <Section style={{ paddingTop: 26 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 16 }}>
          {FEATURES.slice(0, 3).map((f) => (
            <div key={f.title} className="card" style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 14, padding: "20px 22px" }}>
              <div style={{ width: 46, height: 46, borderRadius: 12, background: "#FBF4DF", border: `1px solid ${C.cardBorder}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, marginBottom: 12 }}>{f.icon}</div>
              <div style={{ fontWeight: 600, fontSize: 15.5, marginBottom: 6, color: C.text }}>{f.title}</div>
              <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.55 }}>{f.body}</div>
            </div>
          ))}
        </div>
        <div style={{ textAlign: "center", marginTop: 22 }}>
          <button onClick={() => setPage("features")} style={{ background: "transparent", border: "none", color: C.red, fontSize: 14.5, fontWeight: 600, cursor: "pointer", fontFamily: sans }}>
            All features →
          </button>
        </div>
      </Section>

      {/* CTA BANNER */}
      <Section style={{ paddingTop: 10 }}>
        <div style={{ background: `linear-gradient(135deg, ${C.ink}, #2A4570)`, borderRadius: 18, padding: "44px 34px", textAlign: "center", boxShadow: "0 18px 45px rgba(30,51,88,0.25)" }}>
          <div style={{ fontFamily: display, fontSize: 30, fontWeight: 700, color: C.white }}>Aaj ka hisaab, aaj hi done.</div>
          <div style={{ fontSize: 15, color: "#B9C3D6", marginTop: 10, marginBottom: 26 }}>Start free on the web — your first scanned bill takes under a minute.</div>
          <button className="cta" onClick={() => setPage("app")} style={{ background: "linear-gradient(135deg, #C9962F, #E8B44C)", color: "#2A2005", border: "none", borderRadius: 10, padding: "15px 34px", fontSize: 16, fontFamily: sans, fontWeight: 700, cursor: "pointer", boxShadow: "0 6px 20px rgba(232,180,76,0.4)" }}>
            Start free →
          </button>
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
          <div key={f.title} className="card" style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 14, padding: "20px 22px" }}>
            <div style={{ width: 46, height: 46, borderRadius: 12, background: "#FBF4DF", border: `1px solid ${C.cardBorder}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, marginBottom: 12 }}>{f.icon}</div>
            <div style={{ fontWeight: 600, fontSize: 15.5, marginBottom: 6, color: C.text }}>{f.title}</div>
            <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.55 }}>{f.body}</div>
          </div>
        ))}
      </div>
    </Section>
  );
}

const FREE_PLAN_FEATURES = [
  "30 bill scans per month",
  "Web app + WhatsApp bot",
  "Hindi and English",
  "Automatic GST (CGST/SGST) split",
  "Excel export",
  "Email support",
];
const PRO_PLAN_FEATURES = [
  "Unlimited bill scans, web + WhatsApp",
  "Everything in Free, no monthly cap",
  "No waiting for the 1st of the month",
  "Priority email support",
];

function PlanFeatureList({ items, color }) {
  return (
    <ul style={{ listStyle: "none", padding: 0, margin: "0 0 22px" }}>
      {items.map((f) => (
        <li key={f} style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 13.5, lineHeight: 1.6, marginBottom: 9, color }}>
          <span style={{ flexShrink: 0 }}>✓</span>
          <span>{f}</span>
        </li>
      ))}
    </ul>
  );
}

function PlansPage({ setPage }) {
  const [cycle, setCycle] = useState("monthly");
  const pro = PRO_PRICING[cycle];

  return (
    <Section>
      <div style={{ textAlign: "center", marginBottom: 6 }}>
        <div style={{ fontFamily: display, fontSize: 32, fontWeight: 700, color: C.text }}>Simple, transparent pricing</div>
        <div style={{ fontSize: 14.5, color: C.textMuted, marginTop: 8, maxWidth: 520, marginLeft: "auto", marginRight: "auto" }}>
          No sales calls, no custom quotes - every plan and price is right here. All prices shown are <b>inclusive of GST</b>.
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "center", gap: 6, margin: "28px 0 34px" }}>
        {[["monthly", "Monthly"], ["annual", "Annual · save ~17%"]].map(([key, label]) => (
          <button
            key={key}
            onClick={() => setCycle(key)}
            style={{ background: cycle === key ? C.ink : C.white, color: cycle === key ? C.white : C.text, border: `1px solid ${cycle === key ? C.ink : C.cardBorder}`, borderRadius: 20, padding: "9px 20px", fontSize: 13.5, fontWeight: 600, cursor: "pointer", fontFamily: sans }}
          >
            {label}
          </button>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 20, maxWidth: 700, margin: "0 auto" }}>
        <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 16, padding: "28px 26px" }}>
          <div style={{ fontFamily: display, fontSize: 20, fontWeight: 700, color: C.text }}>Free</div>
          <div style={{ fontSize: 13, color: C.textMuted, marginTop: 4, marginBottom: 18 }}>For trying BahiSathi out</div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
            <span style={{ fontFamily: display, fontSize: 36, fontWeight: 700, color: C.ink }}>₹0</span>
          </div>
          <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 22 }}>forever, no card needed</div>
          <PlanFeatureList items={FREE_PLAN_FEATURES} color={C.text} />
          <button className="cta" onClick={() => setPage("app")} style={{ width: "100%", background: "transparent", color: C.ink, border: `2px solid ${C.ink}`, borderRadius: 8, padding: "12px 0", fontSize: 14.5, fontFamily: sans, fontWeight: 600, cursor: "pointer" }}>
            Start free →
          </button>
        </div>

        <div style={{ background: `linear-gradient(135deg, ${C.ink}, #2A4570)`, borderRadius: 16, padding: "28px 26px", color: C.white, position: "relative", boxShadow: "0 18px 40px rgba(30,51,88,0.3)" }}>
          <div style={{ position: "absolute", top: 20, right: 24, background: C.gold, color: C.white, fontSize: 10.5, fontWeight: 700, borderRadius: 20, padding: "3px 11px", letterSpacing: 0.3 }}>MOST POPULAR</div>
          <div style={{ fontFamily: display, fontSize: 20, fontWeight: 700 }}>Pro</div>
          <div style={{ fontSize: 13, color: "#C7CFDF", marginTop: 4, marginBottom: 18 }}>For shops that scan every day</div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
            <span style={{ fontFamily: display, fontSize: 36, fontWeight: 700 }}>{pro.amount}</span>
            <span style={{ fontSize: 13, color: "#C7CFDF" }}>{pro.per}</span>
          </div>
          <div style={{ fontSize: 12, color: "#C7CFDF", marginBottom: 22, lineHeight: 1.5 }}>{pro.note}</div>
          <PlanFeatureList items={PRO_PLAN_FEATURES} color="#E7ECF7" />
          <button className="cta" onClick={() => setPage("app")} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 8, padding: "12px 0", fontSize: 14.5, fontFamily: sans, fontWeight: 700, cursor: "pointer" }}>
            Get Pro →
          </button>
        </div>
      </div>

      <div style={{ marginTop: 46, maxWidth: 700, marginLeft: "auto", marginRight: "auto" }}>
        <div style={{ fontFamily: display, fontSize: 19, fontWeight: 700, marginBottom: 12, color: C.text }}>Plan terms</div>
        <ul style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.85, paddingLeft: 18, margin: 0 }}>
          <li>All prices shown are final and inclusive of GST - nothing extra is added at checkout.</li>
          <li>Pro is billed upfront for the period you pick (monthly or annual) and does not auto-renew - you're charged again only when you choose to renew.</li>
          <li>You can stop anytime; your Pro access simply continues until the period you already paid for ends.</li>
          <li>Payments are non-refundable once processed.</li>
          <li>Free plan scans reset on the 1st of every month.</li>
          <li>We may change prices for future billing periods; any change will be posted here before it applies.</li>
        </ul>
        <div style={{ fontSize: 12.5, color: C.textMuted, marginTop: 16 }}>
          Full details in our{" "}
          <button onClick={() => setPage("terms")} style={{ background: "transparent", border: "none", padding: 0, color: C.red, textDecoration: "underline", cursor: "pointer", fontSize: 12.5, fontFamily: sans }}>
            Terms of Service
          </button>
          .
        </div>
      </div>
    </Section>
  );
}

const ABOUT_NOW = [
  ["🧾", "Bill photo → ledger entry", "Snap any bill - printed or handwritten - and the vendor, date, GST split, and total are read into your books automatically."],
  ["💬", "WhatsApp bot", "Send a bill photo on WhatsApp, confirm with one tap, done. No new app to learn."],
  ["🌐", "Web ledger", "A simple online khata at bahisathi.in - scan bills, review entries, and see your monthly total anywhere."],
  ["📊", "AI spend summaries", "Ask \"is mahine kitna kharcha hua?\" on the web or WhatsApp and get a straight answer with category breakdowns."],
  ["📤", "Excel export", "One tap gives you a clean spreadsheet your CA or Tally workflow already understands."],
  ["✓", "Honest AI", "When a bill is unclear, BahiSathi flags it and asks - it never quietly guesses with your money."],
];
const ABOUT_COMING = [
  ["🗣️", "More Indian languages", "Beyond Hindi and English - confirm bills and ask questions in the language you think in."],
  ["🔔", "GST filing reminders", "Deadline nudges with your numbers already totalled and ready."],
  ["🤝", "CA sharing", "Give your accountant live, read-only access - no more shoebox of bills at month end."],
];

function AboutPage({ setPage }) {
  return (
    <div>
      <Section style={{ paddingBottom: 24 }}>
        <div style={{ fontFamily: display, fontSize: 34, fontWeight: 700, marginBottom: 6 }}>About BahiSathi</div>
        <div style={{ fontSize: 14, color: C.gold, fontWeight: 600, marginBottom: 22, letterSpacing: 0.4 }}>
          A product of <a href="https://tivaro.co.in/" target="_blank" rel="noopener noreferrer" style={{ color: C.gold, textDecoration: "underline" }}>Tivaro LLP</a>
        </div>
        <div style={{ fontSize: 15, color: C.text, lineHeight: 1.75, maxWidth: 720 }}>
          <p style={{ marginTop: 0 }}>
            <b>Tivaro LLP</b> builds AI-powered digital platforms, mobile applications, and scalable technology solutions for Indian users, businesses, and communities. BahiSathi is our answer to the place where running a small business hurts the most: keeping the books.
          </p>
          <p>
            <b>BahiSathi</b> (bahi = ledger, sathi = companion) is an AI bookkeeping companion for India's shop owners, traders, and service providers. You photograph a bill; BahiSathi reads it, splits the GST, asks you to confirm in plain language, and keeps your khata ready - on WhatsApp or the web, in Hindi or English.
          </p>
        </div>
      </Section>

      <Section style={{ paddingTop: 10, paddingBottom: 24 }}>
        <div style={{ fontFamily: display, fontSize: 24, fontWeight: 700, marginBottom: 18 }}>What BahiSathi does today</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
          {ABOUT_NOW.map(([icon, title, body]) => (
            <div key={title} className="card" style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 14, padding: "18px 20px" }}>
              <div style={{ fontSize: 22, marginBottom: 8 }}>{icon}</div>
              <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 6, color: C.text }}>{title}</div>
              <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.55 }}>{body}</div>
            </div>
          ))}
        </div>
      </Section>

      <Section style={{ paddingTop: 10, paddingBottom: 24 }}>
        <div style={{ fontFamily: display, fontSize: 24, fontWeight: 700, marginBottom: 18 }}>Coming up next</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
          {ABOUT_COMING.map(([icon, title, body]) => (
            <div key={title} className="card" style={{ background: "#FBF4DF", border: `1px dashed ${C.gold}`, borderRadius: 14, padding: "18px 20px" }}>
              <div style={{ fontSize: 22, marginBottom: 8 }}>{icon}</div>
              <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 6, color: C.text }}>{title}</div>
              <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.55 }}>{body}</div>
            </div>
          ))}
        </div>
      </Section>

      <Section style={{ paddingTop: 10, paddingBottom: 24 }}>
        <div style={{ fontFamily: display, fontSize: 24, fontWeight: 700, marginBottom: 18 }}>What makes it different</div>
        <div style={{ fontSize: 15, color: C.text, lineHeight: 1.75, maxWidth: 720 }}>
          <p style={{ marginTop: 0 }}>
            <b>It speaks your language, literally.</b> Most accounting software was built for people who already think in debit-credit. BahiSathi confirms a bill the way you'd say it out loud: "₹1,350 at Sharma Traders - correct?"
          </p>
          <p>
            <b>It reads real Indian bills.</b> Thermal-printer receipts, handwritten parchis, GST invoices - the messy paper that actually crosses your counter, not just neat PDFs.
          </p>
          <p>
            <b>It's honest about uncertainty.</b> When handwriting is unclear, BahiSathi tells you and asks you to check - because wrong books are worse than late books.
          </p>
          <p>
            <b>It lives where you already are.</b> No new app to install or learn. WhatsApp for daily use, the web when you want a bigger screen.
          </p>
        </div>
      </Section>

      <Section style={{ paddingTop: 10 }}>
        <div style={{ fontFamily: display, fontSize: 24, fontWeight: 700, marginBottom: 18 }}>How it helps</div>
        <div style={{ fontSize: 15, color: C.text, lineHeight: 1.75, maxWidth: 720 }}>
          <p style={{ marginTop: 0 }}>
            Every bill captured at the counter means no pile of paper at month end, no forgotten expenses, and no late-night data entry. GST amounts are split as you go, so filing time is totals, not archaeology. Your CA gets a clean Excel sheet instead of a shoebox. And because your khata is always current, you actually know - today - how the month is going.
          </p>
        </div>
        <button className="cta" onClick={() => setPage("app")} style={{ marginTop: 10, background: C.green, color: C.white, border: "none", borderRadius: 10, padding: "13px 26px", fontSize: 15, fontFamily: sans, fontWeight: 600, cursor: "pointer" }}>
          Try BahiSathi free →
        </button>
      </Section>
    </div>
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
  const [agree, setAgree] = useState(false);
  const [showTerms, setShowTerms] = useState(false);
  const [showPrivacy, setShowPrivacy] = useState(false);

  async function handleSubmit() {
    if (mode === "signup" && !agree) return;
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
        {mode === "signup" && (
          <label style={{ display: "flex", gap: 8, alignItems: "flex-start", cursor: "pointer" }}>
            <input type="checkbox" checked={agree} onChange={(e) => setAgree(e.target.checked)} style={{ marginTop: 3, flexShrink: 0 }} />
            <span style={{ fontSize: 12, color: C.textMuted, lineHeight: 1.5 }}>
              I agree to the{" "}
              <button type="button" onClick={() => setShowTerms(true)} style={{ background: "transparent", border: "none", padding: 0, color: C.red, textDecoration: "underline", cursor: "pointer", fontSize: 12, fontFamily: sans }}>
                Terms of Service
              </button>{" "}
              and{" "}
              <button type="button" onClick={() => setShowPrivacy(true)} style={{ background: "transparent", border: "none", padding: 0, color: C.red, textDecoration: "underline", cursor: "pointer", fontSize: 12, fontFamily: sans }}>
                Privacy Policy
              </button>
              , including that Pro plan payments are GST-inclusive and non-refundable.
            </span>
          </label>
        )}
        {error && <div style={{ color: C.red, fontSize: 12.5 }}>{error}</div>}
        <button
          type="button"
          onClick={handleSubmit}
          disabled={loading || !email || !password || (mode === "signup" && !agree)}
          style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "11px 0", fontSize: 14, fontWeight: 500, cursor: loading ? "default" : "pointer", opacity: loading || !email || !password || (mode === "signup" && !agree) ? 0.7 : 1 }}
        >
          {mode === "login" ? "Log in" : "Create account"}
        </button>
      </div>
      <button onClick={() => setMode(mode === "login" ? "signup" : "login")} style={{ marginTop: 14, background: "transparent", border: "none", color: C.inkLight, fontSize: 12.5, cursor: "pointer", textDecoration: "underline" }}>
        {mode === "login" ? "New here? Create an account" : "Already have an account? Log in"}
      </button>
      {showTerms && (
        <Modal onClose={() => setShowTerms(false)}>
          <TermsPage />
        </Modal>
      )}
      {showPrivacy && (
        <Modal onClose={() => setShowPrivacy(false)}>
          <PrivacyPage />
        </Modal>
      )}
    </Section>
  );
}

// ---------- Main product (scan + ledger) ----------
function ProductApp({ session, onLogout, setPage }) {
  const [tab, setTab] = useState("scan");
  const [step, setStep] = useState("upload");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [draft, setDraft] = useState(null);
  const [entries, setEntries] = useState([]);
  const [entriesLoaded, setEntriesLoaded] = useState(false);
  const [sub, setSub] = useState(null);
  const [upgrading, setUpgrading] = useState(false);
  const [upgradeError, setUpgradeError] = useState(null);
  const [upgradeCycle, setUpgradeCycle] = useState("monthly");
  const [linkPhone, setLinkPhone] = useState("");
  const [linkOtp, setLinkOtp] = useState("");
  const [linkStep, setLinkStep] = useState("idle"); // idle | otp_sent
  const [linkLoading, setLinkLoading] = useState(false);
  const [linkError, setLinkError] = useState(null);
  const [linkedNumbers, setLinkedNumbers] = useState([]);
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const [savingEdit, setSavingEdit] = useState(false);
  const [editError, setEditError] = useState(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState(null);
  const [dupConfirmed, setDupConfirmed] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [businessProfile, setBusinessProfile] = useState(null);
  const [savedBusinessProfile, setSavedBusinessProfile] = useState(null);
  const [businessLoaded, setBusinessLoaded] = useState(false);
  const [businessEditing, setBusinessEditing] = useState(false);
  const [businessSaving, setBusinessSaving] = useState(false);
  const [businessError, setBusinessError] = useState(null);
  const [businessSaved, setBusinessSaved] = useState(false);
  const [askQuestion, setAskQuestion] = useState("");
  const [askAnswer, setAskAnswer] = useState(null);
  const [askLoading, setAskLoading] = useState(false);
  const [askError, setAskError] = useState(null);
  const fileInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const data = await billsList(session.access_token);
        setEntries(
          data.map((row) => ({ id: row.id, vendor: row.vendor, date: row.bill_date, gstin: row.gstin, cgst: row.cgst, sgst: row.sgst, igst: row.igst, total: row.total, category: row.category, items: row.items, vendor_source: row.vendor_source, confidence: row.confidence }))
        );
      } catch (e) {
        console.error(e);
      }
      setEntriesLoaded(true);
    })();
  }, [session]);

  async function refreshSubscription() {
    try {
      const status = await getSubscriptionStatus(session.access_token);
      setSub(status);
      setLinkedNumbers(status.linked_whatsapp || []);
    } catch (e) {
      console.error(e);
    }
  }

  useEffect(() => {
    refreshSubscription();
  }, [session]);

  useEffect(() => {
    (async () => {
      try {
        const profile = await businessProfileGet(session.access_token);
        const loaded = {
          business_name: (profile && profile.business_name) || "",
          business_type: (profile && profile.business_type) || "",
          gstin: (profile && profile.gstin) || "",
          address: (profile && profile.address) || "",
        };
        setBusinessProfile(loaded);
        setSavedBusinessProfile(loaded);
        // Nothing saved yet - go straight to the form instead of an empty card.
        setBusinessEditing(!(loaded.business_name || loaded.gstin));
      } catch (e) {
        console.error(e);
        const empty = { business_name: "", business_type: "", gstin: "", address: "" };
        setBusinessProfile(empty);
        setSavedBusinessProfile(empty);
        setBusinessEditing(true);
      }
      setBusinessLoaded(true);
    })();
  }, [session]);

  function updateBusinessField(field, value) {
    setBusinessProfile((p) => ({ ...p, [field]: value }));
    setBusinessSaved(false);
  }

  function cancelEditBusiness() {
    setBusinessProfile(savedBusinessProfile);
    setBusinessError(null);
    setBusinessEditing(false);
  }

  async function saveBusinessProfile() {
    setBusinessSaving(true);
    setBusinessError(null);
    try {
      await businessProfileSave(session.access_token, session.user.id, {
        business_name: businessProfile.business_name || null,
        business_type: businessProfile.business_type || null,
        gstin: businessProfile.gstin || null,
        address: businessProfile.address || null,
      });
      setSavedBusinessProfile(businessProfile);
      setBusinessSaved(true);
      setBusinessEditing(false);
    } catch (e) {
      setBusinessError(e.message || "Could not save - please try again.");
    } finally {
      setBusinessSaving(false);
    }
  }

  async function sendLinkCode() {
    setLinkError(null);
    setLinkLoading(true);
    try {
      await startWhatsAppLink(session.access_token, linkPhone);
      setLinkStep("otp_sent");
    } catch (e) {
      setLinkError(e.message || "Could not send the code - please try again.");
    } finally {
      setLinkLoading(false);
    }
  }

  async function confirmLinkCode() {
    setLinkError(null);
    setLinkLoading(true);
    try {
      await verifyWhatsAppLink(session.access_token, linkOtp);
      setLinkPhone("");
      setLinkOtp("");
      setLinkStep("idle");
      await refreshSubscription();
    } catch (e) {
      setLinkError(e.message || "That code didn't work - please try again.");
    } finally {
      setLinkLoading(false);
    }
  }

  async function removeLink(phone) {
    try {
      await unlinkWhatsApp(session.access_token, phone);
      await refreshSubscription();
    } catch (e) {
      setLinkError(e.message || "Could not unlink - please try again.");
    }
  }

  async function startUpgrade(cycle) {
    setUpgradeError(null);
    setUpgrading(true);
    try {
      await loadRazorpayScript();
      const order = await createProOrder(session.access_token, cycle);
      const rzp = new window.Razorpay({
        key: order.key_id,
        amount: order.amount,
        currency: order.currency,
        order_id: order.order_id,
        name: "BahiSathi",
        description: `Pro plan — ${PRO_PRICING[cycle].amount}${PRO_PRICING[cycle].per}, unlimited scans (GST inclusive)`,
        prefill: { email: session.user?.email || "" },
        theme: { color: "#A63A3A" },
        handler: async (response) => {
          try {
            const result = await verifyProPayment(session.access_token, {
              razorpay_order_id: response.razorpay_order_id,
              razorpay_payment_id: response.razorpay_payment_id,
              razorpay_signature: response.razorpay_signature,
            });
            setSub({ plan: result.plan, status: "active", current_period_end: result.current_period_end, cycle: result.cycle });
          } catch (e) {
            setUpgradeError("Payment went through but activation failed - contact hello@bahisathi.in and we'll sort it out.");
          } finally {
            setUpgrading(false);
          }
        },
        modal: { ondismiss: () => setUpgrading(false) },
      });
      rzp.on("payment.failed", () => {
        setUpgradeError("Payment failed - please try again.");
        setUpgrading(false);
      });
      rzp.open();
    } catch (e) {
      setUpgradeError(e.message || "Could not start payment - please try again.");
      setUpgrading(false);
    }
  }

  async function processFile(file) {
    setError(null);
    setLoading(true);
    setStep("processing");
    try {
      const base64 = await fileToBase64(file);
      const extracted = await extractBillViaApi(base64, file.type || "image/jpeg", session.access_token);
      setDraft({ ...extracted, items: extracted.items || [] });
      setStep("review");
    } catch (e) {
      if (e.status === 429) {
        setError(e.message || "You've used all your free scans for this month. Limit resets on the 1st - paid plans coming soon.");
      } else if (e.status === 401) {
        setError("Your session expired - please log out and log in again.");
      } else {
        setError("Could not read this bill. Try a clearer photo.");
      }
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

  function handleDragOver(e) {
    e.preventDefault();
    setDragOver(true);
  }

  function handleDragLeave(e) {
    e.preventDefault();
    setDragOver(false);
  }

  function handleDrop(e) {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file && file.type.startsWith("image/")) processFile(file);
  }

  function updateDraft(field, value) {
    setDraft((d) => ({ ...d, [field]: value }));
    setDupConfirmed(false);
  }

  async function saveDraft() {
    const dup = findPossibleDuplicate(draft, entries);
    if (dup && !dupConfirmed) {
      setDupConfirmed(true);
      return;
    }
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
      setEntries((prev) => [...prev, { id: saved.id, vendor: saved.vendor, date: saved.bill_date, gstin: saved.gstin, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst, total: saved.total, category: saved.category, items: saved.items, vendor_source: saved.vendor_source, confidence: saved.confidence }]);
    } catch (e) {
      setError("Could not save this bill - please try again.");
      return;
    }
    setDraft(null);
    setDupConfirmed(false);
    setStep("upload");
  }

  function discardDraft() {
    setDraft(null);
    setDupConfirmed(false);
    setStep("upload");
  }

  function startEditEntry(entry) {
    setDeleteConfirmId(null);
    setEditError(null);
    setEditingId(entry.id);
    setEditDraft({ vendor: entry.vendor || "", date: entry.date || "", category: entry.category || "other", total: entry.total ?? "", cgst: entry.cgst ?? "", sgst: entry.sgst ?? "", igst: entry.igst ?? "", items: Array.isArray(entry.items) ? entry.items : [] });
  }

  function cancelEditEntry() {
    setEditingId(null);
    setEditDraft(null);
    setEditError(null);
  }

  function updateEditField(field, value) {
    setEditDraft((d) => ({ ...d, [field]: value }));
  }

  async function saveEditEntry(id) {
    setSavingEdit(true);
    setEditError(null);
    try {
      const hasItems = editDraft.items && editDraft.items.length > 1;
      const patch = {
        vendor: editDraft.vendor,
        bill_date: editDraft.date,
        category: hasItems ? deriveCategory(editDraft.items) : editDraft.category,
        items: editDraft.items,
        total: editDraft.total === "" ? null : Number(editDraft.total),
        cgst: editDraft.cgst === "" ? null : Number(editDraft.cgst),
        sgst: editDraft.sgst === "" ? null : Number(editDraft.sgst),
        igst: editDraft.igst === "" ? null : Number(editDraft.igst),
      };
      const saved = await billsUpdate(session.access_token, id, patch);
      setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, vendor: saved.vendor, date: saved.bill_date, category: saved.category, items: saved.items, total: saved.total, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst } : e)));
      setEditingId(null);
      setEditDraft(null);
    } catch (e) {
      setEditError(e.message || "Could not save changes - please try again.");
    } finally {
      setSavingEdit(false);
    }
  }

  async function deleteEntry(id) {
    if (deleteConfirmId !== id) {
      setDeleteConfirmId(id);
      return;
    }
    try {
      await billsDelete(session.access_token, id);
      setEntries((prev) => prev.filter((e) => e.id !== id));
      if (editingId === id) {
        setEditingId(null);
        setEditDraft(null);
      }
    } catch (e) {
      setEditError(e.message || "Could not delete this entry - please try again.");
    } finally {
      setDeleteConfirmId(null);
    }
  }

  async function askAboutSpending() {
    if (!askQuestion.trim()) return;
    setAskLoading(true);
    setAskError(null);
    setAskAnswer(null);
    try {
      const result = await getSpendSummary(session.access_token, askQuestion.trim());
      setAskAnswer(result.answer);
    } catch (e) {
      setAskError(e.message || "Could not get an answer - please try again.");
    } finally {
      setAskLoading(false);
    }
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
      Items:
        Array.isArray(e.items) && e.items.length > 1
          ? e.items.map((it) => `${it.description || "item"} (₹${it.amount ?? "?"}, ${CATEGORY_LABELS[it.category] || it.category || "other"})`).join("; ")
          : "",
    }));
    const wb = XLSX.utils.book_new();
    let ws;
    if (businessProfile && (businessProfile.business_name || businessProfile.gstin)) {
      const headerRows = [
        [businessProfile.business_name || "BahiSathi Ledger"],
        ...(businessProfile.gstin ? [[`GSTIN: ${businessProfile.gstin}`]] : []),
        ...(businessProfile.address ? [[businessProfile.address]] : []),
        [],
      ];
      ws = XLSX.utils.aoa_to_sheet(headerRows);
      XLSX.utils.sheet_add_json(ws, rows, { origin: -1 });
    } else {
      ws = XLSX.utils.json_to_sheet(rows);
    }
    XLSX.utils.book_append_sheet(wb, ws, "Ledger");
    XLSX.writeFile(wb, "bahisathi-ledger.xlsx");
  }

  const monthTotal = entries.reduce((sum, e) => sum + (Number(e.total) || 0), 0);

  return (
    <div>
      <div style={{ display: "flex", background: C.ink, paddingLeft: 24, gap: 4, justifyContent: "space-between", alignItems: "center", paddingRight: 24 }}>
        <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
          {["scan", "ledger", "plan", "business"].map((key) => (
            <button key={key} onClick={() => setTab(key)} style={{ background: tab === key ? C.paper : "transparent", color: tab === key ? C.ink : "#B9C3D6", border: "none", borderRadius: "8px 8px 0 0", padding: "9px 18px", fontSize: 13.5, fontWeight: 500, cursor: "pointer", fontFamily: sans }}>
              {key === "scan" ? "Scan bill" : key === "ledger" ? "Ledger" : key === "plan" ? "Plan" : "Business"}
            </button>
          ))}
          {sub?.plan === "pro" && (
            <span style={{ background: C.gold, color: C.white, fontSize: 10.5, fontWeight: 600, borderRadius: 20, padding: "3px 10px", letterSpacing: 0.3 }}>PRO</span>
          )}
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
            <div
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              style={{ border: `2px dashed ${C.gold}`, borderRadius: 8, padding: "44px 20px", textAlign: "center", cursor: "pointer", background: dragOver ? "#FDF3DC" : C.white, transition: "background 0.15s" }}
            >
              <div style={{ fontSize: 32, marginBottom: 10 }}>🧾</div>
              <div style={{ fontSize: 14.5, fontWeight: 500, color: C.ink }}>{dragOver ? "Drop to upload" : "Click to choose a photo, or drag and drop it here"}</div>
            </div>
            <input ref={fileInputRef} type="file" accept="image/*" onChange={handleFileChange} style={{ display: "none" }} />
            {error && (
              <div style={{ marginTop: 14, color: C.red, fontSize: 13 }}>
                {error}
                {error.toLowerCase().includes("free scan") && (
                  <div>
                    <button onClick={() => setTab("plan")} style={{ marginTop: 8, background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
                      Upgrade to Pro — ₹199/mo
                    </button>
                  </div>
                )}
              </div>
            )}
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
            {(() => {
              const dup = findPossibleDuplicate(draft, entries);
              if (!dup) return null;
              return (
                <div style={{ background: "#FDECEC", border: `1px solid ${C.red}`, borderRadius: 6, padding: "10px 14px", fontSize: 12.5, marginBottom: 10, color: "#7A2323" }}>
                  ⚠️ Possible duplicate — you already have ₹{Number(dup.total).toLocaleString("en-IN")} at {dup.vendor} {dup.date ? `on ${dup.date}` : ""} in your ledger. Saving again will add a second entry.
                </div>
              );
            })()}
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
            {Array.isArray(draft.items) && draft.items.length > 1 && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "14px 16px", marginTop: 12 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>
                  {draft.items.length} items — this bill spans more than one expense head
                </div>
                <ItemsEditor items={draft.items} onChange={(items) => setDraft((d) => ({ ...d, items, category: deriveCategory(items) }))} />
              </div>
            )}
            <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
              <button onClick={saveDraft} style={{ flex: 1, background: dupConfirmed ? C.red : C.green, color: C.white, border: "none", borderRadius: 6, padding: "11px 0", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>
                {dupConfirmed ? "Yes, save anyway" : "Save to ledger"}
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

            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "14px 16px", marginBottom: 16 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 8 }}>Ask BahiSathi</div>
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  value={askQuestion}
                  onChange={(e) => setAskQuestion(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && askQuestion.trim() && !askLoading) askAboutSpending(); }}
                  placeholder={entries.length > 0 ? "is mahine kitna kharcha hua?" : "how do I export to Excel?"}
                  style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
                />
                <button
                  onClick={askAboutSpending}
                  disabled={askLoading || !askQuestion.trim()}
                  style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 13, fontWeight: 500, cursor: askLoading ? "default" : "pointer", opacity: askLoading || !askQuestion.trim() ? 0.7 : 1, whiteSpace: "nowrap" }}
                >
                  {askLoading ? "Thinking…" : "Ask"}
                </button>
              </div>
              <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 6 }}>About your spending, or how BahiSathi works - ask anything.</div>
              {askError && <div style={{ fontSize: 12.5, color: C.red, marginTop: 8 }}>{askError}</div>}
              {askAnswer && (
                <div style={{ marginTop: 10, background: "#E7FFDB", border: "1px solid #BCE5A8", borderRadius: "10px 10px 10px 2px", padding: "10px 14px", fontSize: 13, color: "#1F2C1A", lineHeight: 1.55 }}>
                  {askAnswer}
                </div>
              )}
            </div>

            {entriesLoaded && entries.length === 0 && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>No bills saved yet. Scan your first one.</div>}
            {entries.length > 0 && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, position: "relative", paddingLeft: 22, overflow: "hidden" }}>
                <div style={{ position: "absolute", left: 14, top: 0, bottom: 0, width: 1.5, background: C.red, opacity: 0.55 }} />
                <div style={{ padding: "16px 18px 16px 6px" }}>
                  {entries.slice().reverse().map((e, i) => {
                    const isEditing = editingId === e.id;
                    if (isEditing) {
                      return (
                        <div key={e.id || i} style={{ padding: "12px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                            <input value={editDraft.vendor} onChange={(ev) => updateEditField("vendor", ev.target.value)} placeholder="Vendor" style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                            <input value={editDraft.date} onChange={(ev) => updateEditField("date", ev.target.value)} placeholder="DD-MM-YYYY" style={{ width: 110, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                          </div>
                          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                            {editDraft.items && editDraft.items.length > 1 ? (
                              <div style={{ flex: 1, fontSize: 12.5, color: C.textMuted, display: "flex", alignItems: "center" }}>Category set per item below</div>
                            ) : (
                              <select value={editDraft.category} onChange={(ev) => updateEditField("category", ev.target.value)} style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }}>
                                {CATEGORY_KEYS.map((k) => (
                                  <option key={k} value={k}>{CATEGORY_LABELS[k]}</option>
                                ))}
                              </select>
                            )}
                            <input value={editDraft.total} onChange={(ev) => updateEditField("total", ev.target.value)} placeholder="Total" style={{ width: 110, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 13, fontFamily: mono, background: C.paper, color: C.text, outline: "none" }} />
                          </div>
                          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                            <input value={editDraft.cgst} onChange={(ev) => updateEditField("cgst", ev.target.value)} placeholder="CGST" style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, fontFamily: mono, background: C.paper, color: C.text, outline: "none" }} />
                            <input value={editDraft.sgst} onChange={(ev) => updateEditField("sgst", ev.target.value)} placeholder="SGST" style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, fontFamily: mono, background: C.paper, color: C.text, outline: "none" }} />
                            <input value={editDraft.igst} onChange={(ev) => updateEditField("igst", ev.target.value)} placeholder="IGST" style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, fontFamily: mono, background: C.paper, color: C.text, outline: "none" }} />
                          </div>
                          {editDraft.items && editDraft.items.length > 1 ? (
                            <div style={{ marginBottom: 10 }}>
                              <ItemsEditor items={editDraft.items} onChange={(items) => updateEditField("items", items)} />
                            </div>
                          ) : (
                            <button
                              onClick={() => updateEditField("items", [{ description: editDraft.vendor || "", amount: editDraft.total || 0, category: editDraft.category }, { description: "", amount: "", category: "other" }])}
                              style={{ background: "transparent", border: "none", color: C.ink, textDecoration: "underline", cursor: "pointer", fontSize: 12, fontFamily: sans, padding: 0, marginBottom: 10, display: "block" }}
                            >
                              Split into items
                            </button>
                          )}
                          {editError && <div style={{ fontSize: 12, color: C.red, marginBottom: 8 }}>{editError}</div>}
                          <div style={{ display: "flex", gap: 8 }}>
                            <button onClick={() => saveEditEntry(e.id)} disabled={savingEdit} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 12.5, fontWeight: 500, cursor: savingEdit ? "default" : "pointer", opacity: savingEdit ? 0.7 : 1 }}>
                              {savingEdit ? "Saving…" : "Save"}
                            </button>
                            <button onClick={cancelEditEntry} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
                              Cancel
                            </button>
                            <button onClick={() => deleteEntry(e.id)} style={{ marginLeft: "auto", background: deleteConfirmId === e.id ? C.red : "transparent", color: deleteConfirmId === e.id ? C.white : C.red, border: `1px solid ${C.red}`, borderRadius: 6, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
                              {deleteConfirmId === e.id ? "Confirm delete?" : "Delete"}
                            </button>
                          </div>
                        </div>
                      );
                    }
                    return (
                      <div key={e.id || i} onClick={() => startEditEntry(e)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5, cursor: "pointer" }}>
                        <div style={{ color: C.textMuted, width: 78, flexShrink: 0 }}>{e.date || "—"}</div>
                        <div style={{ flex: 1 }}>
                          <div style={{ color: C.text }}>{e.vendor || "—"}</div>
                          <div style={{ fontSize: 11, color: C.textMuted }}>
                            {CATEGORY_LABELS[e.category] || e.category}
                            {Array.isArray(e.items) && e.items.length > 1 ? ` · ${e.items.length} items` : ""}
                          </div>
                        </div>
                        <div style={{ fontFamily: mono, color: C.ink, fontWeight: 500, marginRight: 10 }}>₹{Number(e.total || 0).toLocaleString("en-IN")}</div>
                        <div style={{ fontSize: 13, color: C.textMuted }}>✏️</div>
                      </div>
                    );
                  })}
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

        {tab === "plan" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 18 }}>Your plan</div>

            {sub === null && <div style={{ fontSize: 13.5, color: C.textMuted }}>Loading…</div>}

            {sub && sub.plan === "pro" && (
              <div style={{ background: "#FBF4DF", border: `1px solid ${C.gold}`, borderRadius: 8, padding: "20px 22px" }}>
                <div style={{ fontFamily: display, fontSize: 17, fontWeight: 600, color: C.ink, marginBottom: 6 }}>You're on Pro ✨</div>
                <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.6 }}>
                  Unlimited scans, every month.
                  {sub.current_period_end && (
                    <> Renews on {new Date(sub.current_period_end).toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}.</>
                  )}
                </div>
              </div>
            )}

            {sub && sub.plan === "free" && (
              <div>
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginBottom: 18 }}>
                  <div style={{ fontSize: 13.5, color: C.text, fontWeight: 500, marginBottom: 4 }}>Free plan</div>
                  <div style={{ fontSize: 13, color: C.textMuted }}>{sub.scans_used} / {sub.scans_limit} scans used this month</div>
                </div>
                <div style={{ background: "linear-gradient(135deg, #1E3358, #2E4A78)", borderRadius: 8, padding: "22px 24px", color: C.white }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                    <div style={{ fontFamily: display, fontSize: 20, fontWeight: 700 }}>
                      Pro — {PRO_PRICING[upgradeCycle].amount}{PRO_PRICING[upgradeCycle].per}
                    </div>
                    <div style={{ display: "flex", gap: 4, background: "rgba(0,0,0,0.2)", borderRadius: 16, padding: 3 }}>
                      {[["monthly", "Monthly"], ["annual", "Annual"]].map(([key, label]) => (
                        <button
                          key={key}
                          onClick={() => setUpgradeCycle(key)}
                          style={{ background: upgradeCycle === key ? C.gold : "transparent", color: C.white, border: "none", borderRadius: 14, padding: "5px 12px", fontSize: 11.5, fontWeight: 600, cursor: "pointer", fontFamily: sans }}
                        >
                          {label}
                        </button>
                      ))}
                    </div>
                  </div>
                  <div style={{ fontSize: 13.5, color: "#C7CFDF", lineHeight: 1.7, margin: "10px 0 18px" }}>
                    Unlimited bill scans on web and WhatsApp. No caps, no waiting for the 1st of the month. Price shown is inclusive of GST.
                  </div>
                  <button onClick={() => startUpgrade(upgradeCycle)} disabled={upgrading} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "11px 22px", fontSize: 14, fontWeight: 600, cursor: upgrading ? "default" : "pointer", opacity: upgrading ? 0.7 : 1 }}>
                    {upgrading ? "Opening checkout…" : `Upgrade to Pro — ${PRO_PRICING[upgradeCycle].amount}${PRO_PRICING[upgradeCycle].per}`}
                  </button>
                  {upgradeError && <div style={{ marginTop: 12, fontSize: 12.5, color: "#F3B8B8" }}>{upgradeError}</div>}
                  <div style={{ fontSize: 11.5, color: "#9FADC7", marginTop: 12 }}>
                    Non-refundable, does not auto-renew. Full terms on the <button onClick={() => setPage("plans")} style={{ background: "transparent", border: "none", padding: 0, color: "#C7CFDF", textDecoration: "underline", cursor: "pointer", fontSize: 11.5, fontFamily: sans }}>Plans page</button>.
                  </div>
                </div>
              </div>
            )}

            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginTop: 18 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>Link your WhatsApp number</div>
              <div style={{ fontSize: 12.5, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                Scan bills on WhatsApp and the web with one shared free-scan limit and one Pro subscription. Bills you've already sent on WhatsApp will move into this ledger too.
              </div>

              {linkedNumbers.length > 0 && (
                <div style={{ marginBottom: 14 }}>
                  {linkedNumbers.map((num) => (
                    <div key={num} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: C.paper, borderRadius: 6, marginBottom: 6 }}>
                      <span style={{ fontSize: 13, color: C.text, fontFamily: mono }}>+{num}</span>
                      <button onClick={() => removeLink(num)} style={{ background: "transparent", border: "none", color: C.red, fontSize: 12, cursor: "pointer", textDecoration: "underline", fontFamily: sans }}>
                        Unlink
                      </button>
                    </div>
                  ))}
                </div>
              )}

              {linkStep === "idle" && (
                <div style={{ display: "flex", gap: 8 }}>
                  <input
                    type="tel"
                    placeholder="91XXXXXXXXXX"
                    value={linkPhone}
                    onChange={(e) => setLinkPhone(e.target.value)}
                    style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
                  />
                  <button onClick={sendLinkCode} disabled={linkLoading || !linkPhone} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 13, fontWeight: 500, cursor: linkLoading ? "default" : "pointer", opacity: linkLoading || !linkPhone ? 0.7 : 1, whiteSpace: "nowrap" }}>
                    {linkLoading ? "Sending…" : "Send code"}
                  </button>
                </div>
              )}

              {linkStep === "otp_sent" && (
                <div>
                  <div style={{ fontSize: 12.5, color: C.textMuted, marginBottom: 8 }}>Enter the code sent to +{linkPhone} on WhatsApp.</div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <input
                      type="text"
                      placeholder="6-digit code"
                      value={linkOtp}
                      onChange={(e) => setLinkOtp(e.target.value)}
                      style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
                    />
                    <button onClick={confirmLinkCode} disabled={linkLoading || !linkOtp} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 13, fontWeight: 500, cursor: linkLoading ? "default" : "pointer", opacity: linkLoading || !linkOtp ? 0.7 : 1, whiteSpace: "nowrap" }}>
                      {linkLoading ? "Checking…" : "Verify"}
                    </button>
                  </div>
                  <button onClick={() => { setLinkStep("idle"); setLinkOtp(""); setLinkError(null); }} style={{ marginTop: 8, background: "transparent", border: "none", color: C.inkLight, fontSize: 12, cursor: "pointer", textDecoration: "underline", fontFamily: sans, padding: 0 }}>
                    Use a different number
                  </button>
                </div>
              )}

              {linkError && <div style={{ marginTop: 10, fontSize: 12.5, color: C.red }}>{linkError}</div>}
            </div>
          </div>
        )}

        {tab === "business" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Business details</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
              Optional - not needed to scan bills or use your ledger. Fill this in if you want your business name and GSTIN to appear on your Excel exports and, later, on GST-ready reports.
            </div>

            {!businessLoaded && <div style={{ fontSize: 13.5, color: C.textMuted }}>Loading…</div>}

            {businessLoaded && businessProfile && !businessEditing && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 14 }}>
                  <div style={{ fontFamily: display, fontSize: 17, fontWeight: 600, color: C.text }}>
                    {businessProfile.business_name || "No business name set"}
                  </div>
                  <button onClick={() => setBusinessEditing(true)} style={{ flexShrink: 0, background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 6, padding: "6px 14px", fontSize: 12.5, fontWeight: 500, cursor: "pointer" }}>
                    Edit
                  </button>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
                    <span style={{ color: C.textMuted }}>Business type</span>
                    <span style={{ color: C.text }}>{businessProfile.business_type ? BUSINESS_TYPE_LABELS[businessProfile.business_type] : "Not specified"}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
                    <span style={{ color: C.textMuted }}>GSTIN</span>
                    <span style={{ color: C.text, fontFamily: mono }}>{businessProfile.gstin || "Not set"}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", fontSize: 13.5, gap: 16 }}>
                    <span style={{ color: C.textMuted, flexShrink: 0 }}>Address</span>
                    <span style={{ color: C.text, textAlign: "right" }}>{businessProfile.address || "Not set"}</span>
                  </div>
                </div>
              </div>
            )}

            {businessLoaded && businessProfile && businessEditing && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px" }}>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Business name</div>
                  <input
                    value={businessProfile.business_name}
                    onChange={(e) => updateBusinessField("business_name", e.target.value)}
                    placeholder="e.g. Sharma Traders"
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
                  />
                </div>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Business type</div>
                  <select
                    value={businessProfile.business_type}
                    onChange={(e) => updateBusinessField("business_type", e.target.value)}
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
                  >
                    <option value="">Not specified</option>
                    {BUSINESS_TYPE_KEYS.map((k) => (
                      <option key={k} value={k}>{BUSINESS_TYPE_LABELS[k]}</option>
                    ))}
                  </select>
                </div>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>GSTIN</div>
                  <input
                    value={businessProfile.gstin}
                    onChange={(e) => updateBusinessField("gstin", e.target.value.toUpperCase())}
                    placeholder="e.g. 27ABCDE1234F1Z5"
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, fontFamily: mono, background: C.white, color: C.text, outline: "none" }}
                  />
                </div>
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Address</div>
                  <textarea
                    value={businessProfile.address}
                    onChange={(e) => updateBusinessField("address", e.target.value)}
                    placeholder="Shop / office address"
                    rows={3}
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none", fontFamily: sans, resize: "vertical" }}
                  />
                </div>
                {businessError && <div style={{ fontSize: 12.5, color: C.red, marginBottom: 10 }}>{businessError}</div>}
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <button onClick={saveBusinessProfile} disabled={businessSaving} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "10px 20px", fontSize: 13.5, fontWeight: 500, cursor: businessSaving ? "default" : "pointer", opacity: businessSaving ? 0.7 : 1 }}>
                    {businessSaving ? "Saving…" : "Save"}
                  </button>
                  {savedBusinessProfile && (savedBusinessProfile.business_name || savedBusinessProfile.gstin) && (
                    <button onClick={cancelEditBusiness} disabled={businessSaving} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 16px", fontSize: 13.5, cursor: businessSaving ? "default" : "pointer" }}>
                      Cancel
                    </button>
                  )}
                  {businessSaved && !businessSaving && <span style={{ fontSize: 12.5, color: C.green }}>Saved ✓</span>}
                </div>
              </div>
            )}

            <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 14, lineHeight: 1.5 }}>
              No company ID or GST verification is required to use BahiSathi - these details are only used to label your own exports.
            </div>
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

  useEffect(() => {
    window.scrollTo(0, 0);
  }, [page]);

  // Supabase access tokens expire after ~1hr. Without this, anyone actively
  // using the app past that point starts seeing silent 401s ("session
  // expired") mid-task. Schedule a refresh 5 minutes before the token's real
  // expiry (from its own "exp" claim), and re-schedule after every refresh -
  // if the token's already expired (e.g. tab was asleep), this fires almost
  // immediately instead.
  useEffect(() => {
    if (!session || !session.refresh_token || !session.access_token) return;
    const exp = parseJwtExp(session.access_token);
    if (!exp) return;
    const msUntilRefresh = Math.max(exp * 1000 - Date.now() - 5 * 60 * 1000, 5000);
    const timer = setTimeout(async () => {
      try {
        const data = await authRefreshToken(session.refresh_token);
        const refreshed = { access_token: data.access_token, refresh_token: data.refresh_token, user: data.user };
        localStorage.setItem("bahisathi_session", JSON.stringify(refreshed));
        setSession(refreshed);
      } catch (e) {
        console.error("Session refresh failed:", e);
      }
    }, msUntilRefresh);
    return () => clearTimeout(timer);
  }, [session]);

  async function handleLogout() {
    if (session) await authSignOut(session.access_token);
    localStorage.removeItem("bahisathi_session");
    setSession(null);
  }

  return (
    <div style={{ fontFamily: sans, background: C.paper, color: C.text, minHeight: "100vh" }}>
      <style>{`
        *, *::before, *::after { box-sizing: border-box; }
        .cta { transition: transform 0.15s ease, box-shadow 0.15s ease, filter 0.15s ease; }
        .cta:hover { transform: translateY(-2px); filter: brightness(1.06); }
        .navlink { transition: background 0.15s ease, color 0.15s ease; }
        .navlink:hover { background: rgba(241,234,214,0.14); color: #FBF8F0; }
        .card { transition: transform 0.18s ease, box-shadow 0.18s ease; }
        .card:hover { transform: translateY(-4px); box-shadow: 0 14px 34px rgba(30,51,88,0.14); }
      `}</style>
      <NavBar page={page} setPage={setPage} loggedIn={!!session} />
      {page === "home" && <HomePage setPage={setPage} />}
      {page === "features" && <FeaturesPage />}
      {page === "plans" && <PlansPage setPage={setPage} />}
      {page === "about" && <AboutPage setPage={setPage} />}
      {page === "privacy" && <PrivacyPage />}
      {page === "terms" && <TermsPage />}
      {page === "app" && (
        session === undefined ? (
          <Section style={{ textAlign: "center", color: C.textMuted }}>Loading…</Section>
        ) : session ? (
          <ProductApp session={session} onLogout={handleLogout} setPage={setPage} />
        ) : (
          <AuthGate onAuthed={(s) => setSession(s)} />
        )
      )}
      <Footer setPage={setPage} />
    </div>
  );
}
