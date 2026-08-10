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
    { label: "Instagram", href: "#" },
    { label: "LinkedIn", href: "#" },
    { label: "X", href: "#" },
    { label: "WhatsApp", href: "#" },
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
              <a key={s.label} href={s.href} style={{ color: "#C7CFDF", fontSize: 12.5, textDecoration: "none", fontFamily: sans }}>
                {s.label}
              </a>
            ))}
          </div>
        </div>
        <div style={{ borderTop: "1px solid #33476B", paddingTop: 16, color: "#8FA0BE", fontSize: 12, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
          <span>
            © 2026 <a href="https://tivaro.co.in/" target="_blank" rel="noopener noreferrer" style={{ color: "#C7CFDF", textDecoration: "underline" }}>Tivaro LLP</a> · BahiSathi · Made in India 🇮🇳
          </span>
          <span style={{ display: "flex", gap: 14 }}>
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
                🇮🇳 AI bookkeeping for India's small businesses
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

const ABOUT_NOW = [
  ["🧾", "Bill photo → ledger entry", "Snap any bill - printed or handwritten - and the vendor, date, GST split, and total are read into your books automatically."],
  ["💬", "WhatsApp bot", "Send a bill photo on WhatsApp, confirm with one tap, done. No new app to learn."],
  ["🌐", "Web ledger", "A simple online khata at bahisathi.in - scan bills, review entries, and see your monthly total anywhere."],
  ["📤", "Excel export", "One tap gives you a clean spreadsheet your CA or Tally workflow already understands."],
  ["✓", "Honest AI", "When a bill is unclear, BahiSathi flags it and asks - it never quietly guesses with your money."],
];
const ABOUT_COMING = [
  ["🗣️", "More Indian languages", "Beyond Hindi and English - confirm bills and ask questions in the language you think in."],
  ["📊", "Smart summaries", "Ask \"is mahine kitna kharcha hua?\" and get a straight answer with category breakdowns."],
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

  useEffect(() => {
    window.scrollTo(0, 0);
  }, [page]);

  async function handleLogout() {
    if (session) await authSignOut(session.access_token);
    localStorage.removeItem("bahisathi_session");
    setSession(null);
  }

  return (
    <div style={{ fontFamily: sans, background: C.paper, color: C.text, minHeight: "100vh" }}>
      <style>{`
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
      {page === "about" && <AboutPage setPage={setPage} />}
      {page === "privacy" && <PrivacyPage />}
      {page === "terms" && <TermsPage />}
      {page === "app" && (
        session === undefined ? (
          <Section style={{ textAlign: "center", color: C.textMuted }}>Loading…</Section>
        ) : session ? (
          <ProductApp session={session} onLogout={handleLogout} />
        ) : (
          <AuthGate onAuthed={(s) => setSession(s)} />
        )
      )}
      <Footer setPage={setPage} />
    </div>
  );
}
