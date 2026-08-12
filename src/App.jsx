import React, { useState, useEffect, useRef, createContext, useContext } from "react";
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

// --- Interface language toggle (Phase 11) ---
// Separate from the AI's existing ability to read/reply in whatever
// language a bill or WhatsApp message is written in (that's a model
// capability, not a UI concern) - this is a literal English/Hindi switch
// for the app's own labels, buttons, and descriptions. Keyed by the exact
// English string rather than invented keys, so any text not yet in the
// dictionary just falls back to English instead of breaking - the
// dictionary can grow incrementally without ever risking a blank label.
const HI = {
  // Marketing nav
  Home: "होम",
  Features: "फ़ीचर्स",
  Plans: "प्लान्स",
  About: "हमारे बारे में",
  "My Ledger": "मेरी खाता बही",
  "Try it": "इस्तेमाल करें",
  "Log in / Sign up": "लॉग इन / साइन अप",
  // Top nav / tab bar
  "Action Center": "एक्शन सेंटर",
  "Scan bill": "बिल स्कैन करें",
  Ledger: "खाता बही",
  GST: "जीएसटी",
  Invoices: "इनवॉइस",
  Bank: "बैंक",
  Inventory: "इन्वेंटरी",
  Closing: "क्लोजिंग",
  Health: "स्वास्थ्य",
  "P&L": "लाभ-हानि",
  Memory: "मेमोरी",
  "What if?": "क्या हो अगर?",
  Audit: "ऑडिट",
  Plan: "प्लान",
  Business: "बिज़नेस",
  "Log out": "लॉग आउट",
  "👤 CA mode": "👤 सीए मोड",
  "🧑‍💼 Owner mode": "🧑‍💼 मालिक मोड",
  // Auth
  "Log in": "लॉग इन करें",
  "Sign up": "साइन अप करें",
  Email: "ईमेल",
  Password: "पासवर्ड",
  "Create account": "खाता बनाएं",
  "Don't have an account?": "खाता नहीं है?",
  "Already have an account?": "पहले से खाता है?",
  // Scan bill
  "Snap or upload a bill": "बिल की फोटो लें या अपलोड करें",
  "Photo of any invoice, receipt, or bill": "किसी भी इनवॉइस, रसीद, या बिल की फोटो",
  "Reading your bill…": "आपका बिल पढ़ा जा रहा है…",
  Vendor: "विक्रेता",
  Date: "तारीख",
  Category: "श्रेणी",
  Total: "कुल",
  Save: "सेव करें",
  Cancel: "रद्द करें",
  Delete: "हटाएं",
  Edit: "बदलें",
  "+ Add": "+ जोड़ें",
  // Ledger
  "This month": "इस महीने",
  "Ask BahiSathi": "बहीसाथी से पूछें",
  Ask: "पूछें",
  "Export to Excel": "एक्सेल में भेजें",
  "No entries yet - scan your first bill to get started.": "अभी तक कोई एंट्री नहीं - शुरू करने के लिए पहला बिल स्कैन करें।",
  // Action Center
  "Since yesterday": "कल से",
  "No new bills, invoices, or collections logged in the last day.": "पिछले एक दिन में कोई नया बिल, इनवॉइस, या वसूली दर्ज नहीं हुई।",
  "Coming up (next 30 days)": "आने वाला (अगले 30 दिन)",
  "Nothing dated on the calendar right now.": "फ़िलहाल कैलेंडर पर कुछ भी तय नहीं है।",
  "Full health breakdown": "पूरा स्वास्थ्य विवरण",
  "Month-end review": "महीने के अंत की समीक्षा",
  "Not enough of a change vs last month to call out yet.": "पिछले महीने की तुलना में अभी बताने लायक बदलाव नहीं है।",
  "View full P&L statement": "पूरा लाभ-हानि विवरण देखें",
  "All clear - no priority actions right now.": "सब ठीक है - अभी कोई ज़रूरी काम नहीं।",
  "Recent changes": "हाल के बदलाव",
  "Undoing…": "पूर्ववत हो रहा है…",
  "Undo": "पूर्ववत करें",
  "View full audit trail": "पूरा ऑडिट ट्रेल देखें",
  "Your business pulse in one place - Health, GST Guardian, Cash-Flow Guardian, Trust Score, Payment Behaviour, Fix My Books, Recurring Transactions, Price Watch, Leak Detector, Receivable Probability, and the Missing Money Detector, sorted by what's worth doing first.":
    "आपके बिज़नेस की पूरी नब्ज़ एक जगह - स्वास्थ्य, जीएसटी गार्जियन, कैश-फ्लो गार्जियन, ट्रस्ट स्कोर, भुगतान व्यवहार, फिक्स माय बुक्स, रिकरिंग ट्रांज़ैक्शन, प्राइस वॉच, लीक डिटेक्टर, रिसीवेबल प्रोबेबिलिटी, और मिसिंग मनी डिटेक्टर - सबसे ज़रूरी काम पहले।",
  // Scan bill (upload / processing / review)
  "Take a photo or upload a file": "फोटो लें या फाइल अपलोड करें",
  "Choose from gallery": "गैलरी से चुनें",
  "Take photo": "फोटो लें",
  "Upload file": "फाइल अपलोड करें",
  "Add another item": "एक और आइटम जोड़ें",
  "Confirm & save": "पुष्टि करें और सेव करें",
  "Scan another bill": "एक और बिल स्कैन करें",
  // Ledger tab
  "All entries": "सभी एंट्री",
  Amount: "राशि",
  Actions: "कार्रवाई",
  Search: "खोजें",
  "Filter by category": "श्रेणी से छाँटें",
  "All categories": "सभी श्रेणियाँ",
  "No photo? Add a manual entry": "फोटो नहीं है? मैन्युअल एंट्री जोड़ें",
  "Describe the transaction in your own words": "लेनदेन को अपने शब्दों में बताएं",
  "Add entry": "एंट्री जोड़ें",
  "Discard": "मिटाएं",
  "Save to ledger": "खाता बही में सेव करें",
  "Yes, save anyway": "हां, फिर भी सेव करें",
  "Reading…": "पढ़ा जा रहा है…",
  GSTIN: "जीएसटीआईएन",
  "Thinking…": "सोचा जा रहा है…",
  "About your spending, or how BahiSathi works - ask anything.": "आपके खर्च के बारे में, या बहीसाथी कैसे काम करता है - कुछ भी पूछें।",
  "No bills saved yet. Scan your first one.": "अभी तक कोई बिल सेव नहीं हुआ। अपना पहला बिल स्कैन करें।",
  "Split into items": "आइटम में बांटें",
  "Saving…": "सेव हो रहा है…",
  "Confirm delete?": "हटाना पुष्टि करें?",
  "Personal expense, not business": "निजी खर्च, बिज़नेस का नहीं",
  "Check your email": "अपना ईमेल जांचें",
  "We sent a confirmation link to": "हमने एक पुष्टिकरण लिंक भेजा है",
  "Click it, then come back and log in.": "उस पर क्लिक करें, फिर वापस आकर लॉग इन करें।",
  "Welcome to BahiSathi": "बहीसाथी में आपका स्वागत है",
  "Log in or create a free account to keep your ledger.": "अपनी खाता बही रखने के लिए लॉग इन करें या मुफ्त खाता बनाएं।",
  "I agree to the": "मैं सहमत हूं",
  "and": "और",
  "Terms of Service": "सेवा की शर्तें",
  "Privacy Policy": "गोपनीयता नीति",
  ", including that Pro plan payments are GST-inclusive and non-refundable.": ", जिसमें यह भी शामिल है कि प्रो प्लान का भुगतान जीएसटी सहित है और वापस नहीं किया जाएगा।",
  "New here? Create an account": "यहां नए हैं? खाता बनाएं",
  "Already have an account? Log in": "पहले से खाता है? लॉग इन करें",
  // GST tab
  "GST filing reminders": "जीएसटी फाइलिंग रिमाइंडर",
  "Standard monthly filing dates - if you're on the quarterly QRMP scheme, your actual dates differ. Confirm with your CA.": "मानक मासिक फाइलिंग तारीखें - अगर आप तिमाही QRMP योजना में हैं, तो आपकी असल तारीखें अलग होंगी। अपने सीए से पुष्टि करें।",
  "Your recorded purchase GST": "आपका दर्ज खरीद जीएसटी",
  "GST reserve": "जीएसटी रिज़र्व",
  "Estimated GST to reserve": "अनुमानित जीएसटी जो अलग रखना है",
  "Your current cash/bank balance (entered manually - BahiSathi doesn't connect to your bank yet)": "आपका मौजूदा नकद/बैंक बैलेंस (मैन्युअल रूप से दर्ज - बहीसाथी अभी आपके बैंक से नहीं जुड़ा है)",
  "Saved": "सेव हो गया",
  "Available for business": "बिज़नेस के लिए उपलब्ध",
  // Invoices tab
  Customers: "ग्राहक",
  "New invoice": "नया इनवॉइस",
  "No customers yet - add one from a new invoice.": "अभी तक कोई ग्राहक नहीं - नए इनवॉइस से एक जोड़ें।",
  "Customer Profitability": "ग्राहक लाभप्रदता",
  "Paid revenue adjusted for payment delay and at-risk balances - not true margin (BahiSathi doesn't track cost per sale yet), but a sales total alone can be misleading.":
    "भुगतान में देरी और जोखिम वाले बैलेंस के हिसाब से समायोजित आय - असली मार्जिन नहीं (बहीसाथी अभी हर बिक्री की लागत ट्रैक नहीं करता), पर सिर्फ कुल बिक्री भ्रामक हो सकती है।",
  // Plan tab
  "Your plan": "आपका प्लान",
  "Loading…": "लोड हो रहा है…",
  "You're on Pro": "आप प्रो पर हैं",
  "Unlimited scans, every month.": "हर महीने असीमित स्कैन।",
  "Renews on": "नवीनीकरण होगा",
  "Free plan": "फ्री प्लान",
  "scans used this month": "स्कैन इस महीने इस्तेमाल हुए",
  "Unlimited bill scans on web and WhatsApp. No caps, no waiting for the 1st of the month. Price shown is inclusive of GST.": "वेब और व्हाट्सएप पर असीमित बिल स्कैन। कोई सीमा नहीं, महीने की 1 तारीख का इंतज़ार नहीं। दिखाई गई कीमत जीएसटी सहित है।",
  "Opening checkout…": "चेकआउट खुल रहा है…",
  "Upgrade to Pro": "प्रो में अपग्रेड करें",
  "Non-refundable, does not auto-renew. Full terms on the": "वापसी योग्य नहीं है, अपने आप नवीनीकृत नहीं होता। पूरी शर्तें यहां देखें -",
  "Plans page": "प्लान्स पेज",
  "Link your WhatsApp number": "अपना व्हाट्सएप नंबर जोड़ें",
  "Scan bills on WhatsApp and the web with one shared free-scan limit and one Pro subscription. Bills you've already sent on WhatsApp will move into this ledger too.": "व्हाट्सएप और वेब पर एक ही फ्री-स्कैन सीमा और एक प्रो सब्सक्रिप्शन से बिल स्कैन करें। व्हाट्सएप पर पहले भेजे गए बिल भी इस खाता बही में आ जाएंगे।",
  // Business tab
  "Business details": "बिज़नेस विवरण",
  "Optional - not needed to scan bills or use your ledger. Fill this in if you want your business name and GSTIN to appear on your Excel exports and, later, on GST-ready reports.": "वैकल्पिक - बिल स्कैन करने या खाता बही इस्तेमाल करने के लिए ज़रूरी नहीं। अगर आप चाहते हैं कि आपका बिज़नेस नाम और जीएसटीआईएन एक्सेल एक्सपोर्ट और बाद में जीएसटी रिपोर्ट पर दिखे, तो इसे भरें।",
  "AI Business Goals": "एआई बिज़नेस लक्ष्य",
  "Set a target, BahiSathi tracks it against numbers it already computes elsewhere.": "एक लक्ष्य सेट करें, बहीसाथी उसे पहले से मौजूद आंकड़ों के मुकाबले ट्रैक करेगा।",
  "Remove": "हटाएं",
  // What changed? window toggle
  "What changed?": "क्या बदला?",
  "This week": "इस हफ्ते",
  "Nothing new logged in this period.": "इस अवधि में कुछ नया दर्ज नहीं हुआ।",
  // Money Map / Where Did My Profit Go
  "Your business money": "आपका बिज़नेस पैसा",
  "Bank & Cash": "बैंक और नकद",
  "Customers owe you": "ग्राहकों पर बकाया",
  "Enter your cash balance on the GST tab to see your full money map.": "अपना पूरा मनी मैप देखने के लिए जीएसटी टैब पर अपना नकद बैलेंस दर्ज करें।",
  "Inventory value and other assets aren't tracked yet.": "इन्वेंटरी वैल्यू और अन्य संपत्तियां अभी ट्रैक नहीं की जातीं।",
  "Where did my profit go?": "मेरा मुनाफ़ा कहां गया?",
  "Profit this month": "इस महीने का मुनाफ़ा",
  "Sitting with customers (not yet paid)": "ग्राहकों के पास बकाया (अभी भुगतान नहीं हुआ)",
  "Set aside for GST": "जीएसटी के लिए अलग रखा गया",
  "Actually in hand from this month": "इस महीने से असल में हाथ में",
  "Not enough invoiced yet this month to show a breakdown.": "इस महीने अभी इतना इनवॉइस नहीं हुआ कि विवरण दिखाया जा सके।",
  // Customer Churn Radar
  "Customers going quiet": "ग्राहक शांत हो रहे हैं",
  "normally orders every": "आमतौर पर हर",
  "days": "दिन में ऑर्डर करते हैं",
  "Last order": "आखिरी ऑर्डर",
  "days ago": "दिन पहले",
  "Possible revenue at risk": "जोखिम में संभावित राजस्व",
  "annually": "सालाना",
  "Ordering pattern has lapsed vs their own normal cadence - worth a check-in before it becomes lost revenue.": "ऑर्डर पैटर्न उनकी सामान्य लय से हट गया है - राजस्व खोने से पहले संपर्क करना उचित होगा।",
  // Home-page IA
  "To-do": "करने योग्य काम",
  "Books": "बहीखाता",
  "Your business at a glance.": "आपका बिज़नेस एक नज़र में।",
  "Today's priority": "आज की प्राथमिकता",
  "Business pulse": "बिज़नेस पल्स",
  "Cash flow": "कैश फ्लो",
  "Collections": "वसूली",
  "GST readiness": "जीएसटी तैयारी",
  "Safe to spend": "खर्च करने के लिए सुरक्षित",
  "Bank balance minus your estimated GST reserve.": "बैंक बैलेंस में से आपका अनुमानित जीएसटी रिज़र्व घटाकर।",
  "Scan a bill": "बिल स्कैन करें",
  "Cash is at or below your GST reserve - review collections or spending today.": "नकद आपके जीएसटी रिज़र्व के बराबर या उससे कम है - आज वसूली या खर्च की समीक्षा करें।",
  "in receivables looks at risk - worth a follow-up.": "बकाया राशि जोखिम में लग रही है - फॉलो-अप ज़रूरी है।",
  "bill looks overdue.": "बिल का भुगतान बाकी लग रहा है।",
  "has a low trust score - consider requesting advance payment.": "का ट्रस्ट स्कोर कम है - एडवांस भुगतान मांगने पर विचार करें।",
  "Spending is outpacing income lately - worth a look.": "हाल ही में खर्च आय से ज़्यादा हो रहा है - जांच लायक है।",
};
// Home-page IA (Roadmap 3 #49): the owner-mode tab bar collapses from a
// flat, ever-growing row into four sections - Home, Ask BahiSathi, To-do
// (Action Center), and Books. Every tab that isn't one of the first three
// lives inside Books, so new features slot into the Books sub-nav instead
// of adding another top-level tab. CA and staff modes are already narrow
// (5 and 4 tabs) and keep their existing flat bar unchanged.
const BOOKS_TABS = ["scan", "ledger", "gst", "invoices", "bank", "inventory", "closing", "health", "pnl", "memory", "whatif", "audit", "business", "plan"];
const LangContext = createContext({ lang: "en", setLang: () => {}, t: (s) => s });
function useLang() {
  return useContext(LangContext);
}
// EN / हिं pill, same visual language as the existing CA-mode toggle.
function LangToggle({ style }) {
  const { lang, setLang } = useLang();
  return (
    <button
      onClick={() => setLang(lang === "hi" ? "en" : "hi")}
      title="Switch interface language"
      style={{ background: lang === "hi" ? C.gold : "transparent", border: "1px solid #5C6B87", color: lang === "hi" ? C.white : "#B9C3D6", borderRadius: 20, padding: "5px 12px", fontSize: 12.5, cursor: "pointer", fontFamily: sans, ...style }}
    >
      {lang === "hi" ? "हिं / EN" : "EN / हिं"}
    </button>
  );
}

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

// CA sharing - a read-only link (random token, no login) the owner generates
// and sends to their accountant themselves. Only one live link at a time;
// "revoke" just flags it dead rather than deleting the row.
function randomToken() {
  const bytes = new Uint8Array(20);
  (window.crypto || window.msCrypto).getRandomValues(bytes);
  return Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function shareLinkGet(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/share_links?select=*&revoked=eq.false&order=created_at.desc&limit=1`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load share link");
  const data = await res.json();
  return data[0] || null;
}
async function shareLinkCreate(accessToken, userId) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/share_links`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({ token: randomToken(), user_id: userId }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to create share link");
  return data[0];
}
async function shareLinkRevoke(accessToken, token) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/share_links?token=eq.${token}`, {
    method: "PATCH",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({ revoked: true }),
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error((data && data.message) || "Failed to revoke share link");
  }
}
// Public, no-login fetch for whoever has the link (the CA) - hits the
// backend, which uses the service key to read across RLS.
async function fetchSharedLedger(token) {
  const res = await fetch(`${API_BASE}/api/shared/${token}`);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error === "invalid_link" ? "This link is invalid or has been revoked." : "Could not load this ledger - please try again.");
  return data;
}

// --- Invoicing (sales side - separate from the purchase ledger above) ---
async function customersList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/customers?select=*&order=created_at.desc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load customers");
  return res.json();
}
async function customersInsert(accessToken, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/customers`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(row),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to save customer");
  return data[0];
}
async function invoicesList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/invoices?select=*&order=created_at.desc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load invoices");
  return res.json();
}
async function invoicesInsert(accessToken, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/invoices`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(row),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to save invoice");
  return data[0];
}
async function invoicesUpdate(accessToken, id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/invoices?id=eq.${id}`, {
    method: "PATCH",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(patch),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to update invoice");
  return data[0];
}
async function invoicesDelete(accessToken, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/invoices?id=eq.${id}`, {
    method: "DELETE",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error((data && data.message) || "Failed to delete invoice");
  }
}

// --- Inventory (Phase 5) ---
async function inventoryItemsList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/inventory_items?select=*&order=name.asc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load inventory");
  return res.json();
}
async function inventoryItemsInsert(accessToken, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/inventory_items`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(row),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to add item");
  return data[0];
}
async function inventoryItemsUpdate(accessToken, id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/inventory_items?id=eq.${id}`, {
    method: "PATCH",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(patch),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to update item");
  return data[0];
}
async function inventoryItemsDelete(accessToken, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/inventory_items?id=eq.${id}`, {
    method: "DELETE",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error((data && data.message) || "Failed to delete item");
  }
}
async function inventoryMovementsInsert(accessToken, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/inventory_movements`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(row),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to record stock movement");
  return data[0];
}
async function inventoryMovementsList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/inventory_movements?select=*&order=created_at.desc&limit=500`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load stock movements");
  return res.json();
}

// --- Staff access / Owner Vault (Phase 6) ---
// staff_members rows are readable by either side of the relationship (RLS:
// auth.uid() = owner_user_id or auth.uid() = staff_user_id), which is what
// lets both the "my staff list" query and the "am I staff for someone" query
// below use the same table with no separate endpoint.
async function staffMembersList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/staff_members?select=*&order=created_at.desc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load staff list");
  return res.json();
}
async function staffInvite(accessToken, ownerUserId, phone, name) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/staff_members`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({ owner_user_id: ownerUserId, phone, name: name || null, status: "invited" }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to invite staff member");
  return data[0];
}
async function staffRevoke(accessToken, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/staff_members?id=eq.${id}`, {
    method: "PATCH",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({ status: "revoked" }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to revoke access");
  return data[0];
}
// Am I an active staff member for someone else's business? Returns that row
// (with owner_user_id) or null if this account is a normal owner account.
async function staffOfGet(accessToken, myUserId) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/staff_members?select=*&staff_user_id=eq.${myUserId}&status=eq.active&limit=1`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to check staff status");
  const rows = await res.json();
  return rows[0] || null;
}

// --- Bank/UPI reconciliation (Phase 4) ---
// Owner Vault data - no staff RLS policy exists on this table at all, so
// these calls only ever succeed for the account owner.
async function bankTransactionsList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bank_transactions?select=*&order=txn_date.desc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load bank transactions");
  return res.json();
}
async function bankTransactionsInsertMany(accessToken, rows) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bank_transactions`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(rows),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to save statement");
  return data;
}
async function bankTransactionsUpdate(accessToken, id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bank_transactions?id=eq.${id}`, {
    method: "PATCH",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(patch),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to update transaction");
  return data[0];
}
async function bankTransactionsDelete(accessToken, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/bank_transactions?id=eq.${id}`, {
    method: "DELETE",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error((data && data.message) || "Failed to delete transaction");
  }
}

// Bank/UPI exports all name their columns slightly differently, so the
// parser hunts for a header row matching known aliases instead of assuming
// a fixed layout, then reads either a combined debit/credit pair of columns
// or a single signed "amount" column.
const BANK_HEADER_ALIASES = {
  date: ["date", "txn date", "transaction date", "value date", "tran date"],
  description: ["description", "narration", "particulars", "details", "remarks", "transaction details"],
  debit: ["debit", "withdrawal", "withdrawal amt", "withdrawal amt.", "dr", "debit amount"],
  credit: ["credit", "deposit", "deposit amt", "deposit amt.", "cr", "credit amount"],
  amount: ["amount", "amt", "transaction amount"],
};
function findBankHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const cells = (rows[i] || []).map((c) => String(c || "").trim().toLowerCase());
    const hasDate = cells.some((c) => BANK_HEADER_ALIASES.date.includes(c));
    const hasAmount = cells.some((c) => BANK_HEADER_ALIASES.debit.includes(c) || BANK_HEADER_ALIASES.credit.includes(c) || BANK_HEADER_ALIASES.amount.includes(c));
    if (hasDate && hasAmount) return i;
  }
  return -1;
}
function bankColIndex(headerCells, aliases) {
  return headerCells.findIndex((c) => aliases.includes(c));
}
function parseBankAmount(v) {
  if (v == null || v === "") return 0;
  const n = Number(String(v).replace(/[₹,\s]/g, ""));
  return isNaN(n) ? 0 : n;
}
function parseBankDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  const parsed = new Date(s);
  return isNaN(parsed) ? null : parsed.toISOString().slice(0, 10);
}
// Reads the first sheet of a workbook already parsed by XLSX.read (covers
// both .xlsx/.xls and .csv, since XLSX handles CSV the same way) into plain
// { txn_date, description, amount, direction } rows.
function parseBankStatement(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  const headerIdx = findBankHeaderRow(rows);
  if (headerIdx === -1) throw new Error("Could not find date/amount columns in this file - check it's a bank or UPI statement export.");
  const headerCells = rows[headerIdx].map((c) => String(c || "").trim().toLowerCase());
  const dateCol = bankColIndex(headerCells, BANK_HEADER_ALIASES.date);
  const descCol = bankColIndex(headerCells, BANK_HEADER_ALIASES.description);
  const debitCol = bankColIndex(headerCells, BANK_HEADER_ALIASES.debit);
  const creditCol = bankColIndex(headerCells, BANK_HEADER_ALIASES.credit);
  const amountCol = bankColIndex(headerCells, BANK_HEADER_ALIASES.amount);
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every((c) => c === "")) continue;
    const txn_date = dateCol >= 0 ? parseBankDate(row[dateCol]) : null;
    const description = descCol >= 0 ? String(row[descCol] || "").trim() : "";
    let amount = 0, direction = null;
    if (debitCol >= 0 && parseBankAmount(row[debitCol]) > 0) { amount = parseBankAmount(row[debitCol]); direction = "debit"; }
    else if (creditCol >= 0 && parseBankAmount(row[creditCol]) > 0) { amount = parseBankAmount(row[creditCol]); direction = "credit"; }
    else if (amountCol >= 0) { const v = parseBankAmount(row[amountCol]); amount = Math.abs(v); direction = v < 0 ? "debit" : "credit"; }
    if (!txn_date || !amount || !direction) continue;
    out.push({ txn_date, description, amount, direction });
  }
  return out;
}
// Auto-match a freshly parsed row against unpaid bills (for debits) or
// unpaid invoices (for credits) by amount and nearby date - a first pass
// the owner can always correct by hand afterwards.
function autoMatchBankTxn(txn, bills, invoices, matchedBillIds, matchedInvoiceIds) {
  const withinDays = (a, b, days) => a && b && Math.abs((new Date(a) - new Date(b)) / 86400000) <= days;
  if (txn.direction === "debit") {
    const bill = bills.find((b) => !matchedBillIds.has(b.id) && Math.abs((Number(b.total) || 0) - txn.amount) < 1 && withinDays(b.date, txn.txn_date, 5));
    if (bill) return { status: "matched_bill", matched_bill_id: bill.id, matched_invoice_id: null };
  } else {
    const inv = invoices.find((i) => !matchedInvoiceIds.has(i.id) && i.status !== "paid" && Math.abs((Number(i.total) || 0) - txn.amount) < 1 && withinDays(i.invoice_date, txn.txn_date, 30));
    if (inv) return { status: "matched_invoice", matched_bill_id: null, matched_invoice_id: inv.id };
  }
  return { status: "unmatched", matched_bill_id: null, matched_invoice_id: null };
}

// --- Business Memory (Phase 7) ---
// business_notes holds only the owner's own free-text notes - the computed
// pattern cards below (seasonal revenue, top vendors, payment behaviour) are
// derived fresh from entries/invoices each render, same "no new data store"
// philosophy as Cash-Flow Guardian, and are never persisted anywhere.
async function businessNotesList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/business_notes?select=*&order=created_at.desc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load notes");
  return res.json();
}
async function businessNotesInsert(accessToken, userId, text) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/business_notes`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({ user_id: userId, text }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to save note");
  return data[0];
}
async function businessNotesDelete(accessToken, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/business_notes?id=eq.${id}`, {
    method: "DELETE",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error((data && data.message) || "Failed to delete note");
  }
}

// Revenue by calendar month name, averaged across however many distinct
// year-months of invoice history exist - surfaces seasonality (e.g. a
// festive-season spike) once there's enough spread of months to say
// anything meaningful; returns null rather than a misleading 1-2-point read.
function seasonalRevenuePattern(invoices) {
  const byMonth = {};
  for (const inv of invoices) {
    if (!inv.invoice_date) continue;
    const [y, m] = inv.invoice_date.split("-");
    if (!y || !m) continue;
    if (!byMonth[m]) byMonth[m] = { total: 0, yearMonths: new Set() };
    byMonth[m].total += Number(inv.total) || 0;
    byMonth[m].yearMonths.add(`${y}-${m}`);
  }
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const rows = Object.entries(byMonth).map(([m, d]) => ({ month: monthNames[Number(m) - 1], avg: d.total / d.yearMonths.size }));
  if (rows.length < 3) return null;
  rows.sort((a, b) => b.avg - a.avg);
  return rows;
}

// Vendors this business has paid the most, all-time - a longer memory than
// the per-message "usual category" vendor memory used during bill scanning.
function topVendorsAllTime(entries) {
  const byVendor = {};
  for (const e of entries) {
    const vendor = e.vendor && String(e.vendor).trim();
    if (!vendor || e.is_personal) continue;
    byVendor[vendor] = (byVendor[vendor] || 0) + (Number(e.total) || 0);
  }
  return Object.entries(byVendor)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([vendor, total]) => ({ vendor, total }));
}

// Business-wide payment behaviour trend - the aggregate counterpart to
// paymentProfileFor's per-customer view: are customers as a whole paying
// faster or slower lately than they used to.
function paymentBehaviorTrend(invoices) {
  const paidWithDate = invoices
    .filter((inv) => inv.status === "paid" && inv.paid_date && inv.invoice_date)
    .sort((a, b) => new Date(a.invoice_date) - new Date(b.invoice_date));
  const delayOf = (inv) => Math.round((new Date(inv.paid_date) - new Date(inv.invoice_date)) / 86400000);
  const delays = paidWithDate.map(delayOf).filter((d) => Number.isFinite(d) && d >= 0);
  if (delays.length < 4) return null;
  const avgDelay = Math.round(delays.reduce((s, d) => s + d, 0) / delays.length);
  const mid = Math.floor(delays.length / 2);
  const avg = (arr) => arr.reduce((s, d) => s + d, 0) / arr.length;
  const diff = avg(delays.slice(mid)) - avg(delays.slice(0, mid));
  const trend = diff > 3 ? "worse" : diff < -3 ? "better" : "steady";
  return { avgDelay, trend, sampleSize: delays.length };
}

// Dead-stock finding: items with no recorded movement (in or out) in the
// last 60 days, and items sitting at/under their reorder level. Purely
// client-side over already-fetched data, same pattern as fixMyBooksFindings.
function inventoryFindings(items, movements) {
  const findings = [];
  const now = new Date();
  const lastMovementByItem = {};
  for (const m of movements) {
    const t = new Date(m.created_at).getTime();
    if (!lastMovementByItem[m.item_id] || t > lastMovementByItem[m.item_id]) lastMovementByItem[m.item_id] = t;
  }
  for (const item of items) {
    const last = lastMovementByItem[item.id];
    const daysSince = last ? Math.round((now - last) / 86400000) : null;
    if (daysSince == null && item.quantity_on_hand > 0) {
      findings.push({ severity: "medium", label: `${item.name}: no recorded movement yet`, hint: "Log a stock in/out to start tracking turnover for this item." });
    } else if (daysSince != null && daysSince >= 60 && Number(item.quantity_on_hand) > 0) {
      findings.push({ severity: "medium", label: `${item.name}: dead stock`, hint: `No movement in ${daysSince} days, ${item.quantity_on_hand} ${item.unit || "units"} still on hand.` });
    }
    if (item.reorder_level != null && Number(item.quantity_on_hand) <= Number(item.reorder_level)) {
      findings.push({ severity: "high", label: `${item.name}: at/below reorder level`, hint: `${item.quantity_on_hand} ${item.unit || "units"} left, reorder level is ${item.reorder_level}.` });
    }
  }
  return findings;
}

// --- Daily closing (AI Evening Closing / cash reconciliation) ---
async function dailyClosingsList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/daily_closings?select=*&order=closing_date.desc`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load closings");
  return res.json();
}
// One row per user per day - upsert on that unique pair so re-saving the
// same day's closing (e.g. after fixing a typo) updates it instead of erroring.
async function dailyClosingUpsert(accessToken, userId, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/daily_closings?on_conflict=user_id,closing_date`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Prefer: "resolution=merge-duplicates,return=representation",
    },
    body: JSON.stringify({ ...row, user_id: userId }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to save closing");
  return data[0];
}

// --- Audit Trail (Phase 10) ---
// Persistent, cross-session change log - the upgrade from the earlier
// in-memory-only Undo. Needs a one-time table + RLS setup in Supabase (see
// the SQL note above logChange below); until that's run, these calls fail
// quietly and the app behaves exactly as it did before this table existed -
// editing/deleting/toggling still works, it just isn't logged yet.
async function auditLogList(accessToken) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/audit_log?select=*&order=created_at.desc&limit=50`, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to load audit trail");
  return res.json();
}
async function auditLogInsert(accessToken, row) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/audit_log`, {
    method: "POST",
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(row),
  });
  const data = await res.json();
  if (!res.ok) throw new Error((data && data.message) || "Failed to log change");
  return data[0];
}

function computeInvoiceTotals(items, gstMode, gstRate) {
  const subtotal = items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.rate) || 0), 0);
  const gstAmount = (subtotal * (Number(gstRate) || 0)) / 100;
  const cgst = gstMode === "same" ? gstAmount / 2 : 0;
  const sgst = gstMode === "same" ? gstAmount / 2 : 0;
  const igst = gstMode === "different" ? gstAmount : 0;
  return { subtotal, cgst, sgst, igst, total: subtotal + gstAmount };
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
const deleteAccountApi = (accessToken) => apiCall("/api/account/delete", accessToken, {});
const getSpendSummary = (accessToken, question) => apiCall("/api/spend-summary", accessToken, { question });
const parseManualEntry = (accessToken, text) => apiCall("/api/parse-manual-entry", accessToken, { text });
const getCreditAdvice = (accessToken, payload) => apiCall("/api/credit-advice", accessToken, payload);
const getPnlNarration = (accessToken, payload) => apiCall("/api/explain-pnl", accessToken, payload);
const getInventoryAdvice = (accessToken, payload) => apiCall("/api/inventory-advice", accessToken, payload);
const getWhatIfAnswer = (accessToken, payload) => apiCall("/api/whatif", accessToken, payload);
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

// --- AI Business Goals (Phase 10) ---
// Simple owner-set targets, tracked against figures BahiSathi already
// computes elsewhere (cash balance, outstanding receivables, this month's
// revenue and expense ratio) - no new tracking logic, just a target to
// measure the existing numbers against.
const GOAL_TYPES = {
  min_cash: { label: "Minimum cash balance", unit: "₹", prompt: "Keep at least" },
  max_receivables: { label: "Maximum receivables outstanding", unit: "₹", prompt: "Keep receivables below" },
  monthly_revenue: { label: "Monthly revenue target", unit: "₹", prompt: "Reach" },
  max_expense_ratio: { label: "Expense ratio cap", unit: "%", prompt: "Keep expenses under" },
};
function goalProgress(goal, ctx) {
  const fmt = (n) => `₹${Math.round(n).toLocaleString("en-IN")}`;
  if (goal.type === "min_cash") {
    if (ctx.cashBalance == null) return { status: "no_data", label: "No cash balance entered yet - set it on the GST tab." };
    const pct = goal.target > 0 ? Math.min(100, Math.round((ctx.cashBalance / goal.target) * 100)) : 100;
    return { status: ctx.cashBalance >= goal.target ? "met" : "behind", pct, label: `${fmt(ctx.cashBalance)} of ${fmt(goal.target)} target` };
  }
  if (goal.type === "max_receivables") {
    const pct = goal.target > 0 ? Math.min(100, Math.round((ctx.totalOutstanding / goal.target) * 100)) : 100;
    return { status: ctx.totalOutstanding <= goal.target ? "met" : "behind", pct, label: `${fmt(ctx.totalOutstanding)} outstanding, target under ${fmt(goal.target)}` };
  }
  if (goal.type === "monthly_revenue") {
    const pct = goal.target > 0 ? Math.min(100, Math.round((ctx.monthRevenue / goal.target) * 100)) : 0;
    return { status: ctx.monthRevenue >= goal.target ? "met" : "behind", pct, label: `${fmt(ctx.monthRevenue)} of ${fmt(goal.target)} this month` };
  }
  if (goal.type === "max_expense_ratio") {
    if (ctx.monthExpenseRatio == null) return { status: "no_data", label: "Not enough revenue this month to compute a ratio yet." };
    const pct = goal.target > 0 ? Math.min(100, Math.round((ctx.monthExpenseRatio / goal.target) * 100)) : 100;
    return { status: ctx.monthExpenseRatio <= goal.target ? "met" : "behind", pct, label: `Expenses at ${Math.round(ctx.monthExpenseRatio)}% of revenue, target under ${goal.target}%` };
  }
  return { status: "no_data", label: "Unknown goal type." };
}

// A bill's line items can span more than one expense head (e.g. hardware
// plus a delivery charge). This rolls that up to a single label for the
// ledger list and CSV/Excel export - "Mixed" when items disagree, otherwise
// whichever one category everyone agrees on.
function deriveCategory(items) {
  if (!Array.isArray(items) || items.length === 0) return "other";
  const distinct = new Set(items.map((it) => it.category || "other"));
  return distinct.size === 1 ? [...distinct][0] : "mixed";
}

// Confidence + Explanation (Phase 10): a plain-language reason for why a
// bill landed in its category, computed from this vendor's own recorded
// history - the same evidence a human bookkeeper would point to ("we've
// always filed these under office supplies"), surfaced on request instead
// of asking the owner to trust a black box.
function categoryExplanation(entry, allEntries) {
  if (!entry.vendor || !entry.vendor.trim()) {
    return "No vendor name recorded on this bill, so this is a general fallback category rather than a learned one.";
  }
  const label = CATEGORY_LABELS[entry.category] || entry.category || "this category";
  const sameVendor = allEntries.filter((e) => e.id !== entry.id && e.vendor && e.vendor.trim().toLowerCase() === entry.vendor.trim().toLowerCase());
  if (sameVendor.length === 0) {
    return entry.vendor_source === "handwritten_only"
      ? `Best guess from a handwritten bill with no earlier history for ${entry.vendor.trim()} - worth double-checking.`
      : `First bill recorded from ${entry.vendor.trim()} - "${label}" is the AI's read of the bill text, not a learned pattern yet.`;
  }
  const matching = sameVendor.filter((e) => (e.category || "other") === (entry.category || "other"));
  if (matching.length === sameVendor.length) {
    return `Marked as ${label} because all ${sameVendor.length} previous bill${sameVendor.length === 1 ? "" : "s"} from ${entry.vendor.trim()} were too.`;
  }
  if (matching.length > sameVendor.length / 2) {
    return `Marked as ${label} because ${matching.length} of your last ${sameVendor.length} bills from ${entry.vendor.trim()} were also ${label} - a few others were filed differently.`;
  }
  return `${entry.vendor.trim()} has been categorized inconsistently before (only ${matching.length} of ${sameVendor.length} past bills match "${label}") - worth checking this one is right.`;
}

// Field-by-field diff for an audit_log row's before/after snapshots - only
// the fields that actually changed, in plain language rather than a raw
// JSON dump.
const AUDIT_FIELD_LABELS = {
  bill: { vendor: "Vendor", date: "Date", category: "Category", total: "Total", cgst: "CGST", sgst: "SGST", igst: "IGST", is_personal: "Personal" },
  invoice: { status: "Status", paid_date: "Paid date", payment_method: "Payment method" },
};
function auditDiffLines(log) {
  if (!log.before || !log.after) return [];
  const fields = AUDIT_FIELD_LABELS[log.entity_type] || {};
  const lines = [];
  for (const key in fields) {
    const a = log.before[key];
    const b = log.after[key];
    if (JSON.stringify(a ?? null) !== JSON.stringify(b ?? null)) {
      lines.push(`${fields[key]}: ${a ?? "—"} → ${b ?? "—"}`);
    }
  }
  return lines;
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

// Standard monthly GST filing calendar: GSTR-1 (outward supplies) is due the
// 11th, GSTR-3B (summary return + tax payment) is due the 20th, both of the
// month following the filing period. This doesn't know if a business is on
// the quarterly QRMP scheme - it's a general-case reminder, not a filing
// calculation (BahiSathi only tracks purchases/expenses, not sales, so it
// can't compute an actual output-tax liability).
function nextOccurrence(day) {
  const now = new Date();
  let d = new Date(now.getFullYear(), now.getMonth(), day, 0, 0, 0, 0);
  if (d < new Date(now.getFullYear(), now.getMonth(), now.getDate())) {
    d = new Date(now.getFullYear(), now.getMonth() + 1, day);
  }
  return d;
}
function daysUntil(date) {
  const now = new Date();
  const startToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.round((date - startToday) / (24 * 60 * 60 * 1000));
}
function periodLabel(dueDate) {
  const period = new Date(dueDate.getFullYear(), dueDate.getMonth() - 1, 1);
  return period.toLocaleDateString("en-IN", { month: "long", year: "numeric" });
}
function periodKey(dueDate) {
  const period = new Date(dueDate.getFullYear(), dueDate.getMonth() - 1, 1);
  return `${period.getFullYear()}-${String(period.getMonth() + 1).padStart(2, "0")}`;
}
function getGstReminders() {
  const gstr1Due = nextOccurrence(11);
  const gstr3bDue = nextOccurrence(20);
  return {
    gstr1: { due: gstr1Due, daysLeft: daysUntil(gstr1Due), period: periodLabel(gstr1Due), periodKey: periodKey(gstr1Due) },
    gstr3b: { due: gstr3bDue, daysLeft: daysUntil(gstr3bDue), period: periodLabel(gstr3bDue), periodKey: periodKey(gstr3bDue) },
  };
}
// Recorded purchase-side GST (CGST/SGST/IGST) for a given "YYYY-MM" period,
// bucketed by when each bill was scanned (created_at), since bill_date on
// scanned entries can come in inconsistent formats. This is input GST from
// purchases - useful for input tax credit and for a CA's reference, not the
// business's own output/sales GST liability.
function purchaseGstForPeriod(entries, pKey) {
  const rows = entries.filter((e) => e.created_at && e.created_at.slice(0, 7) === pKey);
  const sum = (field) => rows.reduce((s, e) => s + (Number(e[field]) || 0), 0);
  return { cgst: sum("cgst"), sgst: sum("sgst"), igst: sum("igst"), total: sum("total"), count: rows.length };
}
// Output GST (what's owed on sales) for a given "YYYY-MM" period, bucketed by
// invoice date. Used together with purchaseGstForPeriod's input-credit total
// to estimate the net GST an owner should set aside before filing.
function salesGstForPeriod(invoices, pKey) {
  const rows = invoices.filter((inv) => inv.invoice_date && inv.invoice_date.slice(0, 7) === pKey);
  const sum = (field) => rows.reduce((s, inv) => s + (Number(inv[field]) || 0), 0);
  return { cgst: sum("cgst"), sgst: sum("sgst"), igst: sum("igst"), total: sum("cgst") + sum("sgst") + sum("igst"), count: rows.length };
}

// "YYYY-MM" keys for the n completed calendar months before the current one
// (oldest first) - used by the cash-flow forecast and P&L to average recent
// history without counting the still-in-progress current month.
function monthKeysBack(n) {
  const now = new Date();
  const keys = [];
  for (let i = n; i >= 1; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    keys.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  }
  return keys;
}

// AI Cash-Flow Guardian: a predictive balance-risk estimate built from data
// BahiSathi already has - no bank feed. Averages the last 3 completed
// months' business spend (bills, personal-tagged ones excluded) against
// money actually received (invoices marked paid), nets that against the
// current GST reserve obligation, and projects how many months the
// available cash lasts at that burn rate. Deliberately conservative: if
// there isn't enough history or no cash balance entered, it says so rather
// than guessing.
function cashFlowForecast(entries, invoices, businessProfile) {
  const pastMonths = monthKeysBack(3);
  const businessEntries = entries.filter((e) => !e.is_personal);
  const monthlyExpenses = pastMonths.map((mk) => businessEntries.filter((e) => e.created_at && e.created_at.slice(0, 7) === mk).reduce((s, e) => s + (Number(e.total) || 0), 0));
  const monthlyIncome = pastMonths.map((mk) => invoices.filter((inv) => inv.status === "paid" && inv.paid_date && inv.paid_date.slice(0, 7) === mk).reduce((s, inv) => s + (Number(inv.total) || 0), 0));
  const monthsWithData = monthlyExpenses.filter((v, i) => v > 0 || monthlyIncome[i] > 0).length;
  const avgExpense = monthlyExpenses.reduce((s, v) => s + v, 0) / pastMonths.length;
  const avgIncome = monthlyIncome.reduce((s, v) => s + v, 0) / pastMonths.length;
  const netMonthlyBurn = avgExpense - avgIncome;

  const reminders = getGstReminders();
  const purchaseGst = purchaseGstForPeriod(entries, reminders.gstr3b.periodKey);
  const salesGst = salesGstForPeriod(invoices, reminders.gstr3b.periodKey);
  const netReserve = Math.max(salesGst.total - purchaseGst.total, 0);
  const balance = businessProfile && businessProfile.cash_balance != null ? Number(businessProfile.cash_balance) : null;
  const available = balance != null ? balance - netReserve : null;

  const base = { avgExpense, avgIncome, netMonthlyBurn, available, netReserve };
  if (available == null || monthsWithData < 1) return { ...base, status: "insufficient_data" };
  if (netMonthlyBurn <= 0) return { ...base, status: "healthy" };
  if (available <= 0) return { ...base, status: "immediate_risk", monthsRunway: 0 };
  const monthsRunway = available / netMonthlyBurn;
  const riskDate = new Date();
  riskDate.setDate(riskDate.getDate() + Math.round(monthsRunway * 30));
  return { ...base, status: monthsRunway < 2 ? "at_risk" : "watch", monthsRunway, riskDate };
}

// A simple period P&L from data BahiSathi already tracks: revenue from sales
// invoices (subtotal, before GST) minus business expenses from scanned
// bills (personal-tagged ones excluded, GST-inclusive since that's what
// actually left the till). A genuine estimate for the owner, not a
// CA-grade statement.
function pnlForPeriod(entries, invoices, pKey) {
  const businessEntries = entries.filter((e) => !e.is_personal && e.created_at && e.created_at.slice(0, 7) === pKey);
  const revenueRows = invoices.filter((inv) => inv.invoice_date && inv.invoice_date.slice(0, 7) === pKey);
  const revenue = revenueRows.reduce((s, inv) => s + (Number(inv.subtotal) || 0), 0);
  const expenses = businessEntries.reduce((s, e) => s + (Number(e.total) || 0), 0);
  const byCategory = {};
  for (const e of businessEntries) {
    const cat = e.category || "other";
    byCategory[cat] = (byCategory[cat] || 0) + (Number(e.total) || 0);
  }
  const topCategories = Object.entries(byCategory)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 4)
    .map(([category, total]) => ({ category: CATEGORY_LABELS[category] || category, total }));
  return { revenue, expenses, net: revenue - expenses, topCategories, invoiceCount: revenueRows.length, billCount: businessEntries.length };
}

// 6-month P&L trend (Phase 9): the same pnlForPeriod math run once per
// month, oldest first, ending with the current in-progress month - lets an
// owner see at a glance whether margins are improving or slipping instead
// of only ever seeing one month in isolation.
function pnlTrend(entries, invoices, months = 6) {
  const now = new Date();
  const keys = [];
  for (let i = months - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    keys.push({ key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`, label: d.toLocaleDateString("en-IN", { month: "short", year: "2-digit" }) });
  }
  return keys.map(({ key, label }) => ({ key, label, ...pnlForPeriod(entries, invoices, key) }));
}

// Money Map (Roadmap 3 #4): the simplest possible answer to "mera paisa
// kahan hai" from data BahiSathi already has - manually-entered cash/bank
// balance plus what customers currently owe. Inventory value and fixed
// assets aren't included because BahiSathi doesn't capture a cost/value per
// item yet (Inventory tracks quantity only) - shown as an honest gap rather
// than a fabricated number, same policy as everywhere else in the app.
function moneyMapFor(businessProfile, invoices) {
  const cash = businessProfile && businessProfile.cash_balance != null ? Number(businessProfile.cash_balance) : null;
  const outstanding = invoices.filter((inv) => inv.status !== "paid").reduce((s, inv) => s + (Number(inv.total) || 0), 0);
  const total = (cash || 0) + outstanding;
  return { cash, outstanding, total, hasCash: cash != null };
}

// Safe-to-Spend (Roadmap 3 #6, compact Home-tab version): the same
// bank-balance-minus-GST-reserve math the GST tab already shows in full
// detail, factored out so Home can show the one headline number without
// duplicating the whole GST reserve panel.
function safeToSpendFor(businessProfile, entries, invoices) {
  const reminders = getGstReminders();
  const purchaseGst = purchaseGstForPeriod(entries, reminders.gstr3b.periodKey);
  const salesGst = salesGstForPeriod(invoices, reminders.gstr3b.periodKey);
  const netReserve = Math.max(salesGst.total - purchaseGst.total, 0);
  const balance = businessProfile && businessProfile.cash_balance != null ? Number(businessProfile.cash_balance) : null;
  const available = balance != null ? balance - netReserve : null;
  return { available, netReserve, balance };
}

// Where Did My Profit Go (Roadmap 3 #5): bridges this period's accrual
// profit (revenue invoiced minus expenses recorded) against what actually
// landed as cash - the gap is mostly money still sitting with customers who
// haven't paid the invoice yet, plus GST that's earmarked but not yet filed.
// This is a reconciliation over data already collected, not a full cash-flow
// statement (no fixed-asset/loan/drawing tracking yet, see Roadmap 3), so
// it's presented as an approximate bridge, not an exact accounting figure.
function profitToCashBridge(pnl, invoices, pKey, gstReserveThisPeriod) {
  const periodInvoices = invoices.filter((inv) => inv.invoice_date && inv.invoice_date.slice(0, 7) === pKey);
  const collectedThisPeriod = periodInvoices
    .filter((inv) => inv.status === "paid")
    .reduce((s, inv) => s + (Number(inv.subtotal) || 0), 0);
  const uncollected = Math.max(pnl.revenue - collectedThisPeriod, 0);
  const gstReserve = Math.max(gstReserveThisPeriod || 0, 0);
  const actualCashFromThisMonth = pnl.net - uncollected - gstReserve;
  return { profit: pnl.net, uncollected, gstReserve, actualCashFromThisMonth };
}

// Month-End AI Review (Phase 10): a computed (not a new AI call) "what
// changed / what improved / what needs attention / top 3 for next month"
// summary, built from this month's P&L against last month's, plus the same
// prioritized items Action Center already assembled - a mini CFO readout
// without inventing a new data source or backend call.
function monthEndReview(pnl, prevPnl, actionItems) {
  const revenueChangePct = prevPnl.revenue > 0 ? ((pnl.revenue - prevPnl.revenue) / prevPnl.revenue) * 100 : null;
  const expenseChangePct = prevPnl.expenses > 0 ? ((pnl.expenses - prevPnl.expenses) / prevPnl.expenses) * 100 : null;
  const netChange = pnl.net - prevPnl.net;
  const improved = [];
  const attention = [];
  if (revenueChangePct != null && revenueChangePct >= 5) improved.push(`Revenue up ${Math.round(revenueChangePct)}% vs last month`);
  if (revenueChangePct != null && revenueChangePct <= -5) attention.push(`Revenue down ${Math.abs(Math.round(revenueChangePct))}% vs last month`);
  if (expenseChangePct != null && expenseChangePct <= -5) improved.push(`Expenses down ${Math.abs(Math.round(expenseChangePct))}% vs last month`);
  if (expenseChangePct != null && expenseChangePct >= 10) attention.push(`Expenses up ${Math.round(expenseChangePct)}% vs last month`);
  if (netChange > 0) improved.push(`Net profit up ₹${Math.round(netChange).toLocaleString("en-IN")} vs last month`);
  else if (netChange < 0) attention.push(`Net profit down ₹${Math.round(Math.abs(netChange)).toLocaleString("en-IN")} vs last month`);
  const topActions = actionItems.slice(0, 3).map((it) => it.title);
  return { revenueChangePct, expenseChangePct, netChange, improved, attention, topActions };
}

// AI GST Guardian (basic): checks only BahiSathi's own captured data - no
// GST portal integration. Flags bills whose recorded tax doesn't correspond
// to a standard GST slab (likely a scan/entry error), and high-value
// invoices billed to a customer with no GSTIN on file.
function gstGuardianFindings(entries, invoices, customers) {
  const findings = [];
  const standardRates = [0, 5, 12, 18, 28];
  const badRateBills = entries.filter((e) => {
    const total = Number(e.total);
    const taxTotal = (Number(e.cgst) || 0) + (Number(e.sgst) || 0) + (Number(e.igst) || 0);
    if (!total || taxTotal <= 0) return false;
    const base = total - taxTotal;
    if (base <= 0) return false;
    const rate = (taxTotal / base) * 100;
    const closest = standardRates.reduce((a, b) => (Math.abs(b - rate) < Math.abs(a - rate) ? b : a));
    return Math.abs(closest - rate) > 2;
  });
  if (badRateBills.length) {
    findings.push({
      severity: "medium",
      label: `${badRateBills.length} bill${badRateBills.length === 1 ? "" : "s"} with an unusual GST rate`,
      hint: "The recorded tax doesn't match a standard slab (0/5/12/18/28%) - could be a scan misread or a genuinely special rate. Worth a quick check.",
    });
  }

  const bigNoGstinInvoices = invoices.filter((inv) => {
    const cust = customers.find((c) => c.id === inv.customer_id);
    return Number(inv.total) > 10000 && (!cust || !cust.gstin);
  });
  if (bigNoGstinInvoices.length) {
    findings.push({
      severity: "low",
      label: `${bigNoGstinInvoices.length} invoice${bigNoGstinInvoices.length === 1 ? "" : "s"} over ₹10,000 with no customer GSTIN on file`,
      hint: "Not required for every sale, but worth confirming if the customer needs it for their own input tax credit.",
    });
  }

  const badTotalInvoices = invoices.filter((inv) => {
    const computed = (Number(inv.subtotal) || 0) + (Number(inv.cgst) || 0) + (Number(inv.sgst) || 0) + (Number(inv.igst) || 0);
    return inv.total != null && Math.abs(computed - Number(inv.total)) > 1;
  });
  if (badTotalInvoices.length) {
    findings.push({
      severity: "high",
      label: `${badTotalInvoices.length} invoice${badTotalInvoices.length === 1 ? "" : "s"} where subtotal + GST doesn't add up to the total`,
      hint: "Likely edited after creation - double-check the numbers before sending or filing.",
    });
  }

  return findings;
}

// "Fix my books" one-tap audit: continuous anomaly + missing-data checks
// beyond the existing save-time duplicate warning - this scans the whole
// ledger retroactively, all from data already captured.
function fixMyBooksFindings(entries) {
  const findings = [];
  const noVendor = entries.filter((e) => !e.vendor || !e.vendor.trim());
  if (noVendor.length) {
    findings.push({ severity: "medium", label: `${noVendor.length} bill${noVendor.length === 1 ? "" : "s"} missing a vendor name`, hint: "Open each in the Ledger tab and fill it in." });
  }
  const noDate = entries.filter((e) => !e.date);
  if (noDate.length) {
    findings.push({ severity: "medium", label: `${noDate.length} bill${noDate.length === 1 ? "" : "s"} missing a date`, hint: "Dates drive your month totals and GST period - worth fixing." });
  }
  const noTotal = entries.filter((e) => e.total == null || Number(e.total) === 0);
  if (noTotal.length) {
    findings.push({ severity: "high", label: `${noTotal.length} bill${noTotal.length === 1 ? "" : "s"} with no amount recorded`, hint: "These won't count toward any totals until fixed." });
  }

  // Retroactive duplicate scan (same vendor, total within ₹1) across the
  // whole ledger - the save-time check only compares against what existed
  // at the time, so older duplicates can still slip through.
  const withVendorTotal = entries.filter((e) => e.vendor && e.total != null);
  let dupPairs = 0;
  for (let i = 0; i < withVendorTotal.length; i++) {
    for (let j = i + 1; j < withVendorTotal.length; j++) {
      const a = withVendorTotal[i], b = withVendorTotal[j];
      if (a.vendor.trim().toLowerCase() === b.vendor.trim().toLowerCase() && Math.abs(Number(a.total) - Number(b.total)) < 1) dupPairs++;
    }
  }
  if (dupPairs > 0) {
    findings.push({ severity: "medium", label: `${dupPairs} possible duplicate bill pair${dupPairs === 1 ? "" : "s"} found`, hint: "Same vendor and amount recorded more than once - check the Ledger tab for these." });
  }

  // Unusual amount: a bill for a vendor that's more than 3x that vendor's
  // own average (and meaningfully large in absolute terms) - a common
  // symptom of a misread total.
  const byVendor = {};
  for (const e of entries) {
    if (!e.vendor || e.total == null) continue;
    const key = e.vendor.trim().toLowerCase();
    (byVendor[key] = byVendor[key] || []).push(e);
  }
  let outlierCount = 0;
  for (const key in byVendor) {
    const rows = byVendor[key];
    if (rows.length < 3) continue;
    const avg = rows.reduce((s, e) => s + Number(e.total), 0) / rows.length;
    outlierCount += rows.filter((e) => Number(e.total) > avg * 3 && Number(e.total) - avg > 2000).length;
  }
  if (outlierCount > 0) {
    findings.push({ severity: "low", label: `${outlierCount} bill${outlierCount === 1 ? "" : "s"} unusually large for that vendor`, hint: "More than 3x that vendor's usual amount - could be correct, but worth a second look." });
  }

  return findings;
}

// Severity ranking shared by every finding list below so mixed-source
// arrays (Action Center) can be sorted high -> medium -> low consistently.
function severityRank(s) {
  return s === "high" ? 3 : s === "medium" ? 2 : 1;
}

// --- Smart Recurring Transactions (Phase 8) ---
// Finds vendors this business has billed on a fairly regular cadence (rent,
// utilities, subscriptions, wholesale restock) purely from gaps between that
// vendor's own past bill dates - no assumptions beyond what's already been
// scanned in. Vendors with too few bills, or whose gaps are too irregular
// (stdDev more than ~40% of the average gap), are left out rather than
// guessing at a pattern that isn't really there.
function detectRecurringVendors(entries) {
  const byVendor = {};
  for (const e of entries) {
    if (e.is_personal || !e.vendor || !e.vendor.trim() || !e.date || e.total == null) continue;
    const key = e.vendor.trim().toLowerCase();
    if (!byVendor[key]) byVendor[key] = { vendor: e.vendor.trim(), category: e.category, rows: [] };
    byVendor[key].rows.push(e);
  }
  const today = new Date();
  const results = [];
  for (const key in byVendor) {
    const g = byVendor[key];
    const rows = g.rows.slice().sort((a, b) => new Date(a.date) - new Date(b.date));
    if (rows.length < 3) continue;
    const gaps = [];
    for (let i = 1; i < rows.length; i++) {
      const d = Math.round((new Date(rows[i].date) - new Date(rows[i - 1].date)) / 86400000);
      if (d > 0) gaps.push(d);
    }
    if (gaps.length < 2) continue;
    const avgGap = gaps.reduce((s, d) => s + d, 0) / gaps.length;
    const stdDev = Math.sqrt(gaps.reduce((s, d) => s + Math.pow(d - avgGap, 2), 0) / gaps.length);
    if (avgGap < 4 || stdDev / avgGap > 0.4) continue;
    const cadence =
      avgGap <= 9 ? "Weekly" : avgGap <= 16 ? "Fortnightly" : avgGap <= 40 ? "Monthly" : avgGap <= 100 ? "Quarterly" : avgGap <= 200 ? "Half-yearly" : "Yearly";
    const amounts = rows.map((r) => Number(r.total));
    const avgAmount = amounts.reduce((s, a) => s + a, 0) / amounts.length;
    const lastRow = rows[rows.length - 1];
    const nextDue = new Date(lastRow.date);
    nextDue.setDate(nextDue.getDate() + Math.round(avgGap));
    const daysUntilDue = Math.round((nextDue - today) / 86400000);
    results.push({
      vendor: g.vendor,
      category: g.category || "other",
      cadence,
      avgGapDays: Math.round(avgGap),
      avgAmount,
      lastDate: lastRow.date,
      nextDueDate: nextDue.toISOString().slice(0, 10),
      daysUntilDue,
      occurrences: rows.length,
      status: daysUntilDue < 0 ? "overdue" : daysUntilDue <= 7 ? "due_soon" : "upcoming",
    });
  }
  return results.sort((a, b) => a.daysUntilDue - b.daysUntilDue);
}

// --- Vendor Price Watch (Phase 8) ---
// For each vendor+item pairing seen at least 3 times, compares the price
// paid on the most recent bill against that pairing's own earlier average.
// Falls back to the vendor's overall bill total (keyed on the vendor name
// itself) for single-line bills with no itemised breakdown, so vendors
// billed as one lump sum still get watched. Flags only a meaningful,
// sustained jump (>=15% and >=Rs 20 absolute) to avoid noise from normal
// quantity variation.
function vendorPriceWatch(entries) {
  const byKey = {};
  for (const e of entries) {
    if (e.is_personal || !e.vendor || !e.vendor.trim() || !e.date) continue;
    const vendor = e.vendor.trim();
    const lineItems =
      Array.isArray(e.items) && e.items.filter((it) => it.description && it.amount != null).length > 0
        ? e.items.filter((it) => it.description && it.amount != null)
        : e.total != null
        ? [{ description: vendor, amount: e.total }]
        : [];
    for (const it of lineItems) {
      const itemLabel = String(it.description).trim();
      if (!itemLabel) continue;
      const key = `${vendor.toLowerCase()}||${itemLabel.toLowerCase()}`;
      if (!byKey[key]) byKey[key] = { vendor, item: itemLabel, rows: [] };
      byKey[key].rows.push({ date: e.date, amount: Number(it.amount) });
    }
  }
  const results = [];
  for (const key in byKey) {
    const g = byKey[key];
    const rows = g.rows.filter((r) => r.amount > 0).sort((a, b) => new Date(a.date) - new Date(b.date));
    if (rows.length < 3) continue;
    const latest = rows[rows.length - 1];
    const earlier = rows.slice(0, -1);
    const earlierAvg = earlier.reduce((s, r) => s + r.amount, 0) / earlier.length;
    if (earlierAvg <= 0) continue;
    const pctChange = ((latest.amount - earlierAvg) / earlierAvg) * 100;
    if (pctChange >= 15 && latest.amount - earlierAvg >= 20) {
      results.push({
        vendor: g.vendor,
        item: g.item,
        earlierAvg,
        latestAmount: latest.amount,
        pctChange,
        latestDate: latest.date,
        occurrences: rows.length,
      });
    }
  }
  return results.sort((a, b) => b.pctChange - a.pctChange);
}

// --- Expense Leak Detector (Phase 8) ---
// Two independent checks over already-captured bills: (1) a spend category
// running meaningfully hotter this month than its own trailing 3-month
// average, and (2) near-duplicate charges - same vendor, same amount, a few
// days apart - which reads as an accidental double payment rather than the
// broader "same vendor+total ever" data-quality check Fix My Books already
// does further down the ledger.
function expenseLeakFindings(entries) {
  const findings = [];
  const businessEntries = entries.filter((e) => !e.is_personal && e.date && e.total != null);

  const now = new Date();
  const curKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const priorKeys = [];
  for (let i = 1; i <= 3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    priorKeys.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  }
  const spendByCatMonth = {};
  for (const e of businessEntries) {
    const mk = e.date.slice(0, 7);
    if (mk !== curKey && !priorKeys.includes(mk)) continue;
    const cat = e.category || "other";
    spendByCatMonth[cat] = spendByCatMonth[cat] || {};
    spendByCatMonth[cat][mk] = (spendByCatMonth[cat][mk] || 0) + Number(e.total);
  }
  for (const cat in spendByCatMonth) {
    const cur = spendByCatMonth[cat][curKey] || 0;
    const priorTotals = priorKeys.map((k) => spendByCatMonth[cat][k] || 0);
    if (priorTotals.filter((v) => v > 0).length < 2 || cur <= 0) continue;
    const priorAvg = priorTotals.reduce((s, v) => s + v, 0) / priorKeys.length;
    if (priorAvg <= 0) continue;
    const pctChange = ((cur - priorAvg) / priorAvg) * 100;
    if (pctChange >= 30 && cur - priorAvg >= 500) {
      findings.push({
        type: "category_creep",
        severity: pctChange >= 60 ? "high" : "medium",
        label: `${CATEGORY_LABELS[cat] || cat} spend up ${Math.round(pctChange)}% this month`,
        hint: `₹${Math.round(cur).toLocaleString("en-IN")} so far this month vs a ₹${Math.round(priorAvg).toLocaleString("en-IN")} average over the last 3 months.`,
        amount: cur - priorAvg,
      });
    }
  }

  const withVendorTotal = businessEntries.filter((e) => e.vendor && e.vendor.trim());
  const seenPairs = new Set();
  for (let i = 0; i < withVendorTotal.length; i++) {
    for (let j = i + 1; j < withVendorTotal.length; j++) {
      const a = withVendorTotal[i], b = withVendorTotal[j];
      if (a.vendor.trim().toLowerCase() !== b.vendor.trim().toLowerCase()) continue;
      if (Math.abs(Number(a.total) - Number(b.total)) >= 1) continue;
      const daysApart = Math.abs((new Date(a.date) - new Date(b.date)) / 86400000);
      if (daysApart > 3) continue;
      const pairKey = [a.id, b.id].sort().join("-");
      if (seenPairs.has(pairKey)) continue;
      seenPairs.add(pairKey);
      findings.push({
        type: "possible_duplicate",
        severity: "high",
        label: `${a.vendor.trim()}: possible double charge`,
        hint: `Two bills for ₹${Number(a.total).toLocaleString("en-IN")} within ${Math.round(daysApart)} day${Math.round(daysApart) === 1 ? "" : "s"} of each other (${a.date} and ${b.date}) - worth checking this wasn't scanned twice.`,
        amount: Number(a.total),
      });
    }
  }

  return findings.sort((x, y) => severityRank(y.severity) - severityRank(x.severity) || y.amount - x.amount);
}

// --- "What changed since yesterday?" (Phase 9) ---
// A rolling 24-hour digest built entirely from timestamps/dates already on
// entries and invoices (created_at, paid_date, closing_date) - no separate
// "last seen" state to persist, so it's always "since this time yesterday"
// rather than "since you last opened the app". Deliberately only reports
// things that are directly observable this way (new bills, new invoices,
// collections, cash closing) rather than guessing at GST movement or
// anomaly deltas that would need a stored snapshot to compare against.
// Generalized version of "what changed" - takes a window in days so the
// same logic covers "since yesterday" (1), "this week" (7), and "this
// month" (30) instead of three near-duplicate functions (Roadmap 3 #29).
function whatChangedForWindow(entries, invoices, dailyClosings, days) {
  const now = new Date();
  const since = new Date(now.getFullYear(), now.getMonth(), now.getDate() - days);
  const inWindow = (d) => !!d && new Date(d).getTime() >= since.getTime();

  const newBills = entries.filter((e) => !e.is_personal && inWindow(e.created_at));
  const newBillsTotal = newBills.reduce((s, e) => s + (Number(e.total) || 0), 0);

  const newInvoices = invoices.filter((inv) => inWindow(inv.created_at));
  const newInvoicesTotal = newInvoices.reduce((s, inv) => s + (Number(inv.total) || 0), 0);

  const collections = invoices.filter((inv) => inv.status === "paid" && inv.paid_date && inWindow(inv.paid_date));
  const collectionsTotal = collections.reduce((s, inv) => s + (Number(inv.total) || 0), 0);

  const closings = dailyClosings.filter((c) => inWindow(c.closing_date));
  const recentClosing = closings.length ? closings[closings.length - 1] : null;

  const hasActivity = newBills.length > 0 || newInvoices.length > 0 || collections.length > 0 || !!recentClosing;

  return { hasActivity, newBills, newBillsTotal, newInvoices, newInvoicesTotal, collections, collectionsTotal, recentClosing };
}
function whatChangedSinceYesterday(entries, invoices, dailyClosings) {
  return whatChangedForWindow(entries, invoices, dailyClosings, 1);
}

// --- Business Calendar (Phase 10) ---
// Merges GST filing dates, recurring vendor bills coming due, and
// receivables likely to land soon into one date-sorted "what's coming up"
// list - the same underlying signals GST Guardian, Smart Recurring, and
// Receivable Probability already compute, just organized by date instead of
// by feature, which is how an owner actually thinks about the next few
// weeks. Windowed to -3..+30 days so a just-missed date still shows briefly
// as overdue before dropping off.
function buildBusinessCalendar(gstReminders, recurringVendors, receivableRows) {
  const items = [];
  items.push({ date: gstReminders.gstr1.due, daysLeft: gstReminders.gstr1.daysLeft, type: "gst", label: `GSTR-1 filing due (${gstReminders.gstr1.period})` });
  items.push({ date: gstReminders.gstr3b.due, daysLeft: gstReminders.gstr3b.daysLeft, type: "gst", label: `GSTR-3B filing + payment due (${gstReminders.gstr3b.period})` });
  recurringVendors.forEach((r) => {
    items.push({ date: new Date(r.nextDueDate), daysLeft: r.daysUntilDue, type: "bill", label: `${r.vendor} bill expected, ~₹${Math.round(r.avgAmount).toLocaleString("en-IN")}` });
  });
  receivableRows
    .filter((r) => r.bucket === "likely_soon")
    .forEach((r) => {
      const expected = new Date(r.invoice.invoice_date);
      expected.setDate(expected.getDate() + r.expectedDelay);
      items.push({ date: expected, daysLeft: Math.round((expected - new Date()) / 86400000), type: "receivable", label: `${r.customerName} expected to pay ₹${r.total.toLocaleString("en-IN")} (${r.invoice.invoice_number})` });
    });
  return items.filter((it) => it.daysLeft >= -3 && it.daysLeft <= 30).sort((a, b) => a.date - b.date);
}

// --- Missing Money Detector (Phase 9) ---
// An invoice marked paid should leave a trace somewhere else BahiSathi
// already tracks: a matched bank/UPI transaction for non-cash payments, or
// a same-day evening closing for cash ones. When neither exists, the sale
// is recorded as collected but nothing independently confirms the money
// actually arrived - the same spirit as Evening Closing's own per-day
// mismatch check, just widened across the whole ledger instead of one day
// at a time.
function missingMoneyFindings(invoices, bankTransactions, dailyClosings) {
  const findings = [];
  const paidInvoices = invoices.filter((inv) => inv.status === "paid" && inv.paid_date);
  for (const inv of paidInvoices) {
    const total = Number(inv.total) || 0;
    if (!total) continue;
    if (inv.payment_method === "Cash") {
      const closing = dailyClosings.find((c) => c.closing_date === inv.paid_date);
      if (!closing) {
        findings.push({
          severity: total > 10000 ? "high" : "medium",
          label: `${inv.invoice_number}: ₹${total.toLocaleString("en-IN")} cash sale not yet reconciled`,
          hint: `Marked paid in cash on ${inv.paid_date}, but no evening closing was logged for that day - worth confirming the cash actually came in.`,
          amount: total,
        });
      }
    } else {
      const matched = bankTransactions.find((t) => t.matched_invoice_id === inv.id);
      if (!matched) {
        findings.push({
          severity: total > 10000 ? "high" : "medium",
          label: `${inv.invoice_number}: ₹${total.toLocaleString("en-IN")} marked paid via ${inv.payment_method || "bank/UPI"} with no matching bank record`,
          hint: `Paid on ${inv.paid_date} - link it to a bank/UPI transaction on the Bank tab, or confirm it was actually received.`,
          amount: total,
        });
      }
    }
  }
  return findings.sort((a, b) => severityRank(b.severity) - severityRank(a.severity) || b.amount - a.amount);
}

// --- Smart Customer Follow-up Tone (Phase 9) ---
// The reminder text gets progressively firmer the further a customer's
// oldest unpaid invoice has drifted past THEIR OWN usual payment pace
// (avgDelay from paymentProfileFor), not a flat day count - a customer who
// normally takes 45 days isn't "overdue" on day 40, but one who normally
// pays in 10 days and is now at 25 needs a firmer nudge than that day count
// alone would suggest. A worsening trend bumps the tone up one notch even
// if the ratio alone wouldn't. BahiSathi only ever drafts the text into a
// wa.me link - the owner still presses send themselves inside WhatsApp.
const FOLLOWUP_TIERS = [
  { key: "friendly", label: "Friendly nudge" },
  { key: "reminder", label: "Reminder" },
  { key: "overdue", label: "Overdue notice" },
  { key: "final", label: "Final follow-up" },
];
function followUpDraft(customer, profile, businessName) {
  if (!profile.unpaidCount || profile.outstanding <= 0) return null;
  const expectedDelay = profile.avgDelay != null && profile.avgDelay > 0 ? profile.avgDelay : 15;
  const ratio = profile.maxOverdueDays / expectedDelay;
  let tierIndex = ratio <= 1 ? 0 : ratio <= 1.5 ? 1 : ratio <= 2.5 ? 2 : 3;
  if (profile.trend === "worse" && tierIndex < 3) tierIndex += 1;
  const tier = FOLLOWUP_TIERS[tierIndex];
  const firstName = (customer.name || "there").split(" ")[0];
  const biz = businessName || "us";
  const amount = `₹${Math.round(profile.outstanding).toLocaleString("en-IN")}`;
  const days = profile.maxOverdueDays;
  const messages = {
    friendly: `Hi ${firstName}, hope business is going well! Just a friendly note that ${amount} is currently outstanding with ${biz}. No rush at all - let us know if you have any questions. Thank you!`,
    reminder: `Hi ${firstName}, this is a reminder from ${biz} - ${amount} has been outstanding for about ${days} day${days === 1 ? "" : "s"} now. Could you let us know an expected payment date? Thanks so much!`,
    overdue: `Hi ${firstName}, ${amount} owed to ${biz} is now ${days} days overdue, longer than usual for your account. Please arrange payment at the earliest, or let us know if there's an issue we can help with.`,
    final: `Hi ${firstName}, ${amount} owed to ${biz} has been outstanding for ${days} days now, well past the usual. This is a final follow-up - please clear this at the earliest to avoid any disruption to future orders. Thank you for your prompt attention.`,
  };
  return { tier: tier.key, tierLabel: tier.label, message: messages[tier.key] };
}

function ConfidenceDot({ score }) {
  const color = score == null ? C.textMuted : score >= 0.8 ? C.green : score >= 0.5 ? C.amber : C.red;
  return <span style={{ display: "inline-block", width: 7, height: 7, borderRadius: "50%", background: color, marginRight: 6, flexShrink: 0 }} />;
}

// ---------- Nav / layout ----------
function NavBar({ page, setPage, loggedIn }) {
  const { t } = useLang();
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
            {t(label)}
          </button>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <LangToggle />
        {!loggedIn && (
          <button className="cta" onClick={() => setPage("app")} style={{ background: "linear-gradient(135deg, #C9962F, #E8B44C)", border: "none", color: "#2A2005", borderRadius: 20, padding: "9px 18px", fontSize: 13.5, fontFamily: sans, fontWeight: 600, cursor: "pointer", boxShadow: "0 2px 10px rgba(232,180,76,0.35)" }}>
            {t("Log in / Sign up")}
          </button>
        )}
      </div>
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
  { icon: "🌐", title: "Speaks your language", body: "Confirm bills and ask about your money in Hindi, English, Tamil, Telugu, Bengali, Marathi, Gujarati, Kannada, and more." },
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
  "Hindi, English, and more Indian languages",
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
  ["🗣️", "More Indian languages", "Confirm bills and ask questions in Hindi, English, Tamil, Telugu, Bengali, Marathi, Gujarati, Kannada, and more - BahiSathi matches whatever language you write in."],
  ["🔔", "GST filing reminders", "A GST tab with live countdowns to GSTR-1 and GSTR-3B, plus proactive WhatsApp nudges before each deadline."],
  ["🤝", "CA sharing", "Generate a read-only link for your accountant - no login needed, ledger view plus Excel download."],
  ["🧾", "Invoicing", "Create GST-ready sales invoices with your own customers, track paid/unpaid, and print or save as PDF."],
];
const ABOUT_COMING = [];

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

      {ABOUT_COMING.length > 0 && (
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
      )}

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
  const { t } = useLang();
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
        <div style={{ fontFamily: display, fontSize: 20, fontWeight: 600, marginBottom: 8 }}>{t("Check your email")}</div>
        <div style={{ fontSize: 13.5, color: C.textMuted }}>{t("We sent a confirmation link to")} {email}. {t("Click it, then come back and log in.")}</div>
      </Section>
    );
  }

  return (
    <Section style={{ maxWidth: 380 }}>
      <div style={{ fontFamily: display, fontSize: 24, fontWeight: 600, marginBottom: 4 }}>{t("Welcome to BahiSathi")}</div>
      <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 22 }}>{t("Log in or create a free account to keep your ledger.")}</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <input type="email" placeholder={t("Email")} value={email} onChange={(e) => setEmail(e.target.value)} style={{ border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 12px", fontSize: 14, background: C.white, color: C.text, outline: "none" }} />
        <input
          type="password"
          placeholder={t("Password")}
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && email && password && !loading) handleSubmit(); }}
          style={{ border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 12px", fontSize: 14, background: C.white, color: C.text, outline: "none" }}
        />
        {mode === "signup" && (
          <label style={{ display: "flex", gap: 8, alignItems: "flex-start", cursor: "pointer" }}>
            <input type="checkbox" checked={agree} onChange={(e) => setAgree(e.target.checked)} style={{ marginTop: 3, flexShrink: 0 }} />
            <span style={{ fontSize: 12, color: C.textMuted, lineHeight: 1.5 }}>
              {t("I agree to the")}{" "}
              <button type="button" onClick={() => setShowTerms(true)} style={{ background: "transparent", border: "none", padding: 0, color: C.red, textDecoration: "underline", cursor: "pointer", fontSize: 12, fontFamily: sans }}>
                {t("Terms of Service")}
              </button>{" "}
              {t("and")}{" "}
              <button type="button" onClick={() => setShowPrivacy(true)} style={{ background: "transparent", border: "none", padding: 0, color: C.red, textDecoration: "underline", cursor: "pointer", fontSize: 12, fontFamily: sans }}>
                {t("Privacy Policy")}
              </button>
              {t(", including that Pro plan payments are GST-inclusive and non-refundable.")}
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
          {mode === "login" ? t("Log in") : t("Create account")}
        </button>
      </div>
      <button onClick={() => setMode(mode === "login" ? "signup" : "login")} style={{ marginTop: 14, background: "transparent", border: "none", color: C.inkLight, fontSize: 12.5, cursor: "pointer", textDecoration: "underline" }}>
        {mode === "login" ? t("New here? Create an account") : t("Already have an account? Log in")}
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
  const { t } = useLang();
  const [tab, setTab] = useState("scan");
  // Owner/CA mode toggle (Phase 3): CA mode narrows the nav to the
  // read-focused tabs a CA cares about and turns off ledger editing, so the
  // same logged-in account can hand the screen to their accountant without
  // the external read-only share link.
  const [viewMode, setViewMode] = useState("owner");
  // "What changed?" window for the Action Center summary card - 1/7/30 days
  // (Roadmap 3 #29: since yesterday / this week / this month).
  const [changedWindowDays, setChangedWindowDays] = useState(1);
  // Home-page IA redesign: which "Books" sub-tab was last visited, so
  // clicking the Books section button returns to where the owner left off
  // instead of always resetting to the same default tab.
  const [booksTab, setBooksTab] = useState("ledger");
  useEffect(() => {
    if (BOOKS_TABS.includes(tab)) setBooksTab(tab);
  }, [tab]);
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
  const [explainEntryId, setExplainEntryId] = useState(null);
  // Phase 10: persistent Audit Trail - backed by the audit_log table (see
  // logChange/restoreAuditEntry below), so unlike the first cut of this
  // feature it survives reload. auditLog holds the most recent rows
  // (audit_log_list caps at 50); restoringLogId/expandedLogId are UI-only.
  const [auditLog, setAuditLog] = useState([]);
  const [auditLogLoaded, setAuditLogLoaded] = useState(false);
  const [restoringLogId, setRestoringLogId] = useState(null);
  const [expandedLogId, setExpandedLogId] = useState(null);
  const [dupConfirmed, setDupConfirmed] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [manualEntryOpen, setManualEntryOpen] = useState(false);
  const [manualEntryText, setManualEntryText] = useState("");
  const [manualEntryLoading, setManualEntryLoading] = useState(false);
  const [manualEntryError, setManualEntryError] = useState(null);
  const [businessProfile, setBusinessProfile] = useState(null);
  const [savedBusinessProfile, setSavedBusinessProfile] = useState(null);
  const [businessLoaded, setBusinessLoaded] = useState(false);
  const [businessEditing, setBusinessEditing] = useState(false);
  const [businessSaving, setBusinessSaving] = useState(false);
  const [businessError, setBusinessError] = useState(null);
  const [businessSaved, setBusinessSaved] = useState(false);
  const [cashBalanceInput, setCashBalanceInput] = useState("");
  const [cashBalanceSaving, setCashBalanceSaving] = useState(false);
  // Phase 10: AI Business Goals - stored as a "goals" jsonb column on
  // business_profiles (see the goals SQL note near saveBusinessGoals below).
  // Defaults to an empty array so nothing breaks before that column exists.
  const [businessGoals, setBusinessGoals] = useState([]);
  const [goalsLoaded, setGoalsLoaded] = useState(false);
  const [savingGoal, setSavingGoal] = useState(false);
  const [goalsError, setGoalsError] = useState(null);
  const [newGoalType, setNewGoalType] = useState("min_cash");
  const [newGoalValue, setNewGoalValue] = useState("");
  const [cashBalanceSaved, setCashBalanceSaved] = useState(false);
  const [morningBriefSaving, setMorningBriefSaving] = useState(false);
  const [askQuestion, setAskQuestion] = useState("");
  const [askAnswer, setAskAnswer] = useState(null);
  const [askLoading, setAskLoading] = useState(false);
  const [askError, setAskError] = useState(null);
  const [shareLink, setShareLink] = useState(null);
  const [shareLoaded, setShareLoaded] = useState(false);
  const [shareBusy, setShareBusy] = useState(false);
  const [shareError, setShareError] = useState(null);
  const [shareCopied, setShareCopied] = useState(false);
  const [customers, setCustomers] = useState([]);
  const [customersLoaded, setCustomersLoaded] = useState(false);
  const [invoices, setInvoices] = useState([]);
  const [invoicesLoaded, setInvoicesLoaded] = useState(false);
  const [invoiceView, setInvoiceView] = useState("list"); // list | new
  const [invoiceDraft, setInvoiceDraft] = useState(null);
  const [invoiceSaving, setInvoiceSaving] = useState(false);
  const [invoiceError, setInvoiceError] = useState(null);
  const [previewInvoice, setPreviewInvoice] = useState(null);
  const [newCustomerMode, setNewCustomerMode] = useState(false);
  const [newCustomer, setNewCustomer] = useState({ name: "", gstin: "", address: "", phone: "" });
  const [customerSaving, setCustomerSaving] = useState(false);
  const [customerError, setCustomerError] = useState(null);
  const [deleteConfirming, setDeleteConfirming] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [deleteError, setDeleteError] = useState(null);
  const [transportSearch, setTransportSearch] = useState("");
  const [payMethodPickerId, setPayMethodPickerId] = useState(null);
  const [creditAdviceFor, setCreditAdviceFor] = useState(null);
  const [creditAdviceLoading, setCreditAdviceLoading] = useState(false);
  const [creditAdviceText, setCreditAdviceText] = useState({});
  const [pnlNarration, setPnlNarration] = useState(null);
  const [pnlNarrationLoading, setPnlNarrationLoading] = useState(false);
  // Phase 5: Inventory
  const [inventoryItems, setInventoryItems] = useState([]);
  const [inventoryMovements, setInventoryMovements] = useState([]);
  const [inventoryLoaded, setInventoryLoaded] = useState(false);
  const [newItemName, setNewItemName] = useState("");
  const [newItemUnit, setNewItemUnit] = useState("pcs");
  const [newItemQty, setNewItemQty] = useState("");
  const [newItemReorder, setNewItemReorder] = useState("");
  const [addingItem, setAddingItem] = useState(false);
  const [adjustQtyFor, setAdjustQtyFor] = useState(null);
  const [adjustQtyValue, setAdjustQtyValue] = useState("");
  const [adjustQtyReason, setAdjustQtyReason] = useState("");
  const [inventoryAdvice, setInventoryAdvice] = useState(null);
  const [inventoryAdviceLoading, setInventoryAdviceLoading] = useState(false);
  const [dailyClosings, setDailyClosings] = useState([]);
  const [closingsLoaded, setClosingsLoaded] = useState(false);
  const [closingDate, setClosingDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [enteredCashInput, setEnteredCashInput] = useState("");
  const [closingNote, setClosingNote] = useState("");
  const [closingSaving, setClosingSaving] = useState(false);
  const [closingSaved, setClosingSaved] = useState(false);
  // Phase 6: Staff access / Owner Vault. staffRecord is undefined while
  // loading, null once we know this account is a normal owner account, or
  // the staff_members row if this login belongs to an invited staff member.
  const [staffRecord, setStaffRecord] = useState(undefined);
  const [staffList, setStaffList] = useState([]);
  const [staffListLoaded, setStaffListLoaded] = useState(false);
  const [newStaffPhone, setNewStaffPhone] = useState("");
  const [newStaffName, setNewStaffName] = useState("");
  const [staffInviting, setStaffInviting] = useState(false);
  const [staffError, setStaffError] = useState(null);
  // Phase 4: Bank/UPI reconciliation (Owner Vault - no staff access)
  const [bankTransactions, setBankTransactions] = useState([]);
  const [bankLoaded, setBankLoaded] = useState(false);
  const [bankUploading, setBankUploading] = useState(false);
  const [bankUploadError, setBankUploadError] = useState(null);
  const [bankLinkPickerId, setBankLinkPickerId] = useState(null);
  const bankFileInputRef = useRef(null);
  // Phase 7: Business Memory notes + What if? simulator
  const [businessNotes, setBusinessNotes] = useState([]);
  const [notesLoaded, setNotesLoaded] = useState(false);
  const [newNoteText, setNewNoteText] = useState("");
  const [addingNote, setAddingNote] = useState(false);
  const [whatIfQuestion, setWhatIfQuestion] = useState("");
  const [whatIfAnswer, setWhatIfAnswer] = useState(null);
  const [whatIfLoading, setWhatIfLoading] = useState(false);
  const [whatIfError, setWhatIfError] = useState(null);
  // Phase 8: Smart Recurring / Price Watch / Leak Detector / Action Center -
  // all pure computations over entries/invoices already in state, so the
  // only new state needed is UI feedback for the recurring quick-add button.
  const [quickAddingVendor, setQuickAddingVendor] = useState(null);
  const [quickAddedVendors, setQuickAddedVendors] = useState([]);
  const fileInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const rec = await staffOfGet(session.access_token, session.user.id);
        setStaffRecord(rec || null);
      } catch (e) {
        console.error(e);
        setStaffRecord(null);
      }
    })();
  }, [session]);

  // Land owners on the new Home screen once we've positively confirmed this
  // login is a real owner (staffRecord === null, not just still-loading
  // undefined) - staff logins never get this tab, so the default stays
  // "scan" for them, same as before the Home-page IA redesign.
  useEffect(() => {
    if (staffRecord === null) setTab("home");
  }, [staffRecord]);

  // role/effectiveOwnerId drive everything below: a staff login still reads
  // and writes through the same tables, just scoped to the owner's data
  // (via RLS) instead of their own, and with a narrower nav.
  const role = staffRecord ? "staff" : "owner";
  const effectiveOwnerId = staffRecord ? staffRecord.owner_user_id : session.user.id;

  useEffect(() => {
    if (role !== "owner") return;
    (async () => {
      try {
        const rows = await staffMembersList(session.access_token);
        setStaffList(rows);
      } catch (e) {
        console.error(e);
      }
      setStaffListLoaded(true);
    })();
  }, [session, role]);

  async function inviteStaff() {
    setStaffError(null);
    // Digits only, no "+" - matches how the WhatsApp bot normalizes numbers
    // (see link-start in server.js), so the activation lookup by phone matches.
    const phone = newStaffPhone.replace(/[^0-9]/g, "");
    if (!phone) {
      setStaffError("Enter the staff member's WhatsApp number.");
      return;
    }
    setStaffInviting(true);
    try {
      const row = await staffInvite(session.access_token, session.user.id, phone, newStaffName.trim());
      setStaffList((prev) => [row, ...prev]);
      setNewStaffPhone("");
      setNewStaffName("");
    } catch (e) {
      setStaffError(e.message || "Could not send invite - please try again.");
    } finally {
      setStaffInviting(false);
    }
  }

  async function revokeStaffAccess(id) {
    try {
      const updated = await staffRevoke(session.access_token, id);
      setStaffList((prev) => prev.map((s) => (s.id === id ? updated : s)));
    } catch (e) {
      console.error(e);
    }
  }

  // Phase 4: bank statement rows are Owner Vault data, same as staff_members -
  // don't even try to load them for a staff login (the Bank tab is hidden
  // from staff anyway via ownerTabs/staffTabs below).
  useEffect(() => {
    if (role !== "owner") {
      setBankLoaded(true);
      return;
    }
    (async () => {
      try {
        const rows = await bankTransactionsList(session.access_token);
        setBankTransactions(rows);
      } catch (e) {
        console.error(e);
      }
      setBankLoaded(true);
    })();
  }, [session, role]);

  async function handleBankFileChange(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setBankUploadError(null);
    setBankUploading(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const parsed = parseBankStatement(wb);
      if (parsed.length === 0) throw new Error("No transactions found in this file.");
      const matchedBillIds = new Set(bankTransactions.filter((t) => t.matched_bill_id).map((t) => t.matched_bill_id));
      const matchedInvoiceIds = new Set(bankTransactions.filter((t) => t.matched_invoice_id).map((t) => t.matched_invoice_id));
      const batchLabel = `${file.name} · ${new Date().toLocaleDateString("en-IN")}`;
      const rowsToInsert = parsed.map((txn) => {
        const match = autoMatchBankTxn(txn, entries, invoices, matchedBillIds, matchedInvoiceIds);
        if (match.matched_bill_id) matchedBillIds.add(match.matched_bill_id);
        if (match.matched_invoice_id) matchedInvoiceIds.add(match.matched_invoice_id);
        return { user_id: effectiveOwnerId, txn_date: txn.txn_date, description: txn.description, amount: txn.amount, direction: txn.direction, batch_label: batchLabel, ...match };
      });
      const saved = await bankTransactionsInsertMany(session.access_token, rowsToInsert);
      setBankTransactions((prev) => [...saved, ...prev].sort((a, b) => (a.txn_date < b.txn_date ? 1 : -1)));
    } catch (e) {
      setBankUploadError(e.message || "Could not read this file - please check the format.");
    } finally {
      setBankUploading(false);
    }
  }

  // "Add as expense" for an unmatched debit - creates a bill from the bank
  // row so it counts in spending/reports the same as a scanned bill would.
  async function addBankTxnAsExpense(txn) {
    try {
      const saved = await billsInsert(session.access_token, {
        user_id: effectiveOwnerId,
        created_by: session.user.id,
        vendor: txn.description || "Bank debit",
        bill_date: txn.txn_date,
        total: txn.amount,
        category: "other",
        is_personal: false,
      });
      setEntries((prev) => [...prev, { id: saved.id, vendor: saved.vendor, date: saved.bill_date, total: saved.total, category: saved.category, items: saved.items, is_personal: saved.is_personal, created_at: saved.created_at }]);
      const updated = await bankTransactionsUpdate(session.access_token, txn.id, { status: "matched_bill", matched_bill_id: saved.id });
      setBankTransactions((prev) => prev.map((t) => (t.id === txn.id ? updated : t)));
    } catch (e) {
      console.error(e);
    }
  }

  // Quick-add for a Smart Recurring Transactions hit - logs today's
  // occurrence of a vendor BahiSathi has recognised a billing cadence for,
  // pre-filled with that vendor's own average amount and category so it's a
  // single tap rather than a full manual entry; the owner can still edit the
  // amount in the Ledger tab afterwards once the real bill arrives.
  async function quickAddRecurringBill(rec) {
    setQuickAddingVendor(rec.vendor);
    try {
      const today = new Date().toISOString().slice(0, 10);
      const saved = await billsInsert(session.access_token, {
        user_id: effectiveOwnerId,
        created_by: session.user.id,
        vendor: rec.vendor,
        bill_date: today,
        total: Math.round(rec.avgAmount),
        category: rec.category || "other",
        is_personal: false,
      });
      setEntries((prev) => [...prev, { id: saved.id, vendor: saved.vendor, date: saved.bill_date, total: saved.total, category: saved.category, items: saved.items, is_personal: saved.is_personal, created_at: saved.created_at }]);
      setQuickAddedVendors((prev) => [...prev, rec.vendor]);
    } catch (e) {
      console.error(e);
    } finally {
      setQuickAddingVendor(null);
    }
  }

  // Manually link an unmatched credit to a chosen unpaid invoice, marking
  // that invoice paid via bank transfer at the same time.
  async function linkBankTxnToInvoice(txn, invoiceId) {
    try {
      const inv = invoices.find((i) => i.id === invoiceId);
      if (inv && inv.status !== "paid") await toggleInvoiceStatus(inv, "Bank Transfer");
      const updated = await bankTransactionsUpdate(session.access_token, txn.id, { status: "matched_invoice", matched_invoice_id: invoiceId });
      setBankTransactions((prev) => prev.map((t) => (t.id === txn.id ? updated : t)));
      setBankLinkPickerId(null);
    } catch (e) {
      console.error(e);
    }
  }

  async function undoBankTxnMatch(txn) {
    try {
      const updated = await bankTransactionsUpdate(session.access_token, txn.id, { status: "unmatched", matched_bill_id: null, matched_invoice_id: null });
      setBankTransactions((prev) => prev.map((t) => (t.id === txn.id ? updated : t)));
    } catch (e) {
      console.error(e);
    }
  }

  async function deleteBankTxn(id) {
    try {
      await bankTransactionsDelete(session.access_token, id);
      setBankTransactions((prev) => prev.filter((t) => t.id !== id));
    } catch (e) {
      console.error(e);
    }
  }

  // Phase 7: Business Memory notes - owner-only, like bank_transactions.
  useEffect(() => {
    if (role !== "owner") {
      setNotesLoaded(true);
      return;
    }
    (async () => {
      try {
        const rows = await businessNotesList(session.access_token);
        setBusinessNotes(rows);
      } catch (e) {
        console.error(e);
      }
      setNotesLoaded(true);
    })();
  }, [session, role]);

  async function addBusinessNote() {
    const text = newNoteText.trim();
    if (!text) return;
    setAddingNote(true);
    try {
      const saved = await businessNotesInsert(session.access_token, session.user.id, text);
      setBusinessNotes((prev) => [saved, ...prev]);
      setNewNoteText("");
    } catch (e) {
      console.error(e);
    } finally {
      setAddingNote(false);
    }
  }

  async function deleteBusinessNote(id) {
    try {
      await businessNotesDelete(session.access_token, id);
      setBusinessNotes((prev) => prev.filter((n) => n.id !== id));
    } catch (e) {
      console.error(e);
    }
  }

  // What if? simulator - grounds the question in the same computed numbers
  // the Health/GST/P&L tabs already show (via pnlForPeriod/cashFlowForecast)
  // plus the Business Memory patterns/notes above, so the AI is reasoning
  // over this business's real figures rather than guessing generically.
  async function askWhatIf(question) {
    const q = (question ?? whatIfQuestion).trim();
    if (!q) return;
    setWhatIfLoading(true);
    setWhatIfError(null);
    setWhatIfAnswer(null);
    try {
      const now = new Date();
      const pKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
      const pnl = pnlForPeriod(entries, invoices, pKey);
      const cashFlow = cashFlowForecast(entries, invoices, businessProfile);
      const result = await getWhatIfAnswer(session.access_token, {
        question: q,
        pnl: { revenue: pnl.revenue, expenses: pnl.expenses, net: pnl.net, topCategories: pnl.topCategories },
        cashFlow: { avgExpense: cashFlow.avgExpense, avgIncome: cashFlow.avgIncome, available: cashFlow.available, status: cashFlow.status, monthsRunway: cashFlow.monthsRunway },
        seasonal: seasonalRevenuePattern(invoices),
        topVendors: topVendorsAllTime(entries),
        paymentBehavior: paymentBehaviorTrend(invoices),
        notes: businessNotes.map((n) => n.text),
      });
      setWhatIfAnswer(result.answer);
    } catch (e) {
      setWhatIfError(e.message || "Could not run that scenario - please try again.");
    } finally {
      setWhatIfLoading(false);
    }
  }

  useEffect(() => {
    (async () => {
      try {
        const data = await billsList(session.access_token);
        setEntries(
          data.map((row) => ({ id: row.id, vendor: row.vendor, date: row.bill_date, gstin: row.gstin, cgst: row.cgst, sgst: row.sgst, igst: row.igst, total: row.total, category: row.category, items: row.items, vendor_source: row.vendor_source, confidence: row.confidence, is_personal: row.is_personal, created_at: row.created_at, created_by: row.created_by }))
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
          website: (profile && profile.website) || "",
          email: (profile && profile.email) || "",
          cash_balance: profile && profile.cash_balance != null ? profile.cash_balance : null,
          cash_balance_updated_at: (profile && profile.cash_balance_updated_at) || null,
          morning_brief_enabled: !!(profile && profile.morning_brief_enabled),
        };
        setBusinessProfile(loaded);
        setSavedBusinessProfile(loaded);
        setCashBalanceInput(loaded.cash_balance != null ? String(loaded.cash_balance) : "");
        // Nothing saved yet - go straight to the form instead of an empty card.
        setBusinessEditing(!(loaded.business_name || loaded.gstin));
        // goals is a separate optional column - defaults to [] both if the
        // column doesn't exist yet on this account's table and if it's just
        // never been set, so the feature degrades gracefully either way.
        setBusinessGoals(profile && Array.isArray(profile.goals) ? profile.goals : []);
      } catch (e) {
        console.error(e);
        const empty = { business_name: "", business_type: "", gstin: "", address: "", website: "", email: "", cash_balance: null, cash_balance_updated_at: null, morning_brief_enabled: false };
        setBusinessProfile(empty);
        setSavedBusinessProfile(empty);
        setBusinessEditing(true);
        setBusinessGoals([]);
      }
      setGoalsLoaded(true);
      setBusinessLoaded(true);
    })();
  }, [session]);

  // Standalone quick-save for the GST tab's cash reserve card - doesn't touch
  // the rest of the business profile edit flow.
  async function saveCashBalance() {
    setCashBalanceSaving(true);
    setCashBalanceSaved(false);
    try {
      const value = cashBalanceInput.trim() === "" ? null : Number(cashBalanceInput);
      const now = new Date().toISOString();
      await businessProfileSave(session.access_token, session.user.id, {
        business_name: businessProfile.business_name || null,
        business_type: businessProfile.business_type || null,
        gstin: businessProfile.gstin || null,
        address: businessProfile.address || null,
        website: businessProfile.website || null,
        email: businessProfile.email || null,
        cash_balance: value,
        cash_balance_updated_at: now,
      });
      const updated = { ...businessProfile, cash_balance: value, cash_balance_updated_at: now };
      setBusinessProfile(updated);
      setSavedBusinessProfile(updated);
      setCashBalanceSaved(true);
    } catch (e) {
      console.error(e);
    } finally {
      setCashBalanceSaving(false);
    }
  }

  // AI Business Goals (Phase 10) - stored as a "goals" jsonb array on the
  // same business_profiles row as everything else on the Business tab. If
  // this column doesn't exist yet, businessProfileSave will fail with a
  // clear Postgres error surfaced in goalsError; run this once in the
  // Supabase SQL editor to add it:
  //   ALTER TABLE business_profiles ADD COLUMN IF NOT EXISTS goals jsonb DEFAULT '[]'::jsonb;
  async function addGoal() {
    const target = Number(newGoalValue);
    if (!target || target <= 0) return;
    setSavingGoal(true);
    setGoalsError(null);
    try {
      const goal = { id: `${Date.now()}`, type: newGoalType, target, created_at: new Date().toISOString() };
      const nextGoals = [...businessGoals, goal];
      await businessProfileSave(session.access_token, session.user.id, {
        business_name: businessProfile.business_name || null,
        business_type: businessProfile.business_type || null,
        gstin: businessProfile.gstin || null,
        address: businessProfile.address || null,
        website: businessProfile.website || null,
        email: businessProfile.email || null,
        goals: nextGoals,
      });
      setBusinessGoals(nextGoals);
      setNewGoalValue("");
    } catch (e) {
      setGoalsError(e.message && e.message.includes("column") ? "Goals need a one-time setup step in Supabase - see the note in the code, or ask Claude to walk you through it." : e.message || "Could not save this goal - please try again.");
    } finally {
      setSavingGoal(false);
    }
  }
  async function deleteGoal(id) {
    const nextGoals = businessGoals.filter((g) => g.id !== id);
    try {
      await businessProfileSave(session.access_token, session.user.id, {
        business_name: businessProfile.business_name || null,
        business_type: businessProfile.business_type || null,
        gstin: businessProfile.gstin || null,
        address: businessProfile.address || null,
        website: businessProfile.website || null,
        email: businessProfile.email || null,
        goals: nextGoals,
      });
      setBusinessGoals(nextGoals);
    } catch (e) {
      console.error(e);
    }
  }

  async function toggleMorningBrief() {
    const next = !businessProfile.morning_brief_enabled;
    setMorningBriefSaving(true);
    try {
      await businessProfileSave(session.access_token, session.user.id, {
        business_name: businessProfile.business_name || null,
        business_type: businessProfile.business_type || null,
        gstin: businessProfile.gstin || null,
        address: businessProfile.address || null,
        website: businessProfile.website || null,
        email: businessProfile.email || null,
        morning_brief_enabled: next,
      });
      const updated = { ...businessProfile, morning_brief_enabled: next };
      setBusinessProfile(updated);
      setSavedBusinessProfile(updated);
    } catch (e) {
      console.error(e);
    } finally {
      setMorningBriefSaving(false);
    }
  }

  async function saveClosing() {
    setClosingSaving(true);
    setClosingSaved(false);
    try {
      const expected = expectedCashFor(closingDate);
      const entered = enteredCashInput.trim() === "" ? 0 : Number(enteredCashInput);
      const saved = await dailyClosingUpsert(session.access_token, session.user.id, {
        closing_date: closingDate,
        expected_cash: expected,
        entered_cash: entered,
        difference: entered - expected,
        note: closingNote.trim() || null,
      });
      setDailyClosings((prev) => {
        const others = prev.filter((c) => c.closing_date !== closingDate);
        return [saved, ...others].sort((a, b) => (a.closing_date < b.closing_date ? 1 : -1));
      });
      setClosingSaved(true);
    } catch (e) {
      console.error(e);
    } finally {
      setClosingSaving(false);
    }
  }

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
        website: businessProfile.website || null,
        email: businessProfile.email || null,
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

  useEffect(() => {
    (async () => {
      try {
        const link = await shareLinkGet(session.access_token);
        setShareLink(link);
      } catch (e) {
        console.error(e);
      }
      setShareLoaded(true);
    })();
  }, [session]);

  async function generateShareLink() {
    setShareBusy(true);
    setShareError(null);
    try {
      const link = await shareLinkCreate(session.access_token, session.user.id);
      setShareLink(link);
    } catch (e) {
      setShareError(e.message || "Could not create link - please try again.");
    } finally {
      setShareBusy(false);
    }
  }

  async function revokeShareLink() {
    if (!shareLink) return;
    setShareBusy(true);
    setShareError(null);
    try {
      await shareLinkRevoke(session.access_token, shareLink.token);
      setShareLink(null);
    } catch (e) {
      setShareError(e.message || "Could not revoke - please try again.");
    } finally {
      setShareBusy(false);
    }
  }

  function copyShareLink() {
    if (!shareLink) return;
    const url = `${window.location.origin}${window.location.pathname}?share=${shareLink.token}`;
    navigator.clipboard.writeText(url).then(() => {
      setShareCopied(true);
      setTimeout(() => setShareCopied(false), 2000);
    });
  }

  useEffect(() => {
    (async () => {
      try {
        const rows = await customersList(session.access_token);
        setCustomers(rows);
      } catch (e) {
        console.error(e);
      }
      setCustomersLoaded(true);
    })();
  }, [session]);

  useEffect(() => {
    (async () => {
      try {
        const rows = await invoicesList(session.access_token);
        setInvoices(rows);
      } catch (e) {
        console.error(e);
      }
      setInvoicesLoaded(true);
    })();
  }, [session]);

  useEffect(() => {
    (async () => {
      try {
        const [items, movements] = await Promise.all([inventoryItemsList(session.access_token), inventoryMovementsList(session.access_token)]);
        setInventoryItems(items);
        setInventoryMovements(movements);
      } catch (e) {
        console.error(e);
      }
      setInventoryLoaded(true);
    })();
  }, [session]);

  useEffect(() => {
    (async () => {
      try {
        const rows = await dailyClosingsList(session.access_token);
        setDailyClosings(rows);
      } catch (e) {
        console.error(e);
      }
      setClosingsLoaded(true);
    })();
  }, [session]);

  // Fails quietly if the audit_log table hasn't been created yet - the
  // trail just stays empty until the one-time SQL migration is run, same
  // graceful-degrade approach as AI Business Goals' "goals" column.
  useEffect(() => {
    (async () => {
      try {
        const rows = await auditLogList(session.access_token);
        setAuditLog(rows);
      } catch (e) {
        console.error(e);
      }
      setAuditLogLoaded(true);
    })();
  }, [session]);

  // Whenever the selected closing date changes (or closings finish loading),
  // prefill the form from any already-saved closing for that date.
  useEffect(() => {
    const existing = dailyClosings.find((c) => c.closing_date === closingDate);
    setEnteredCashInput(existing ? String(existing.entered_cash) : "");
    setClosingNote(existing && existing.note ? existing.note : "");
    setClosingSaved(false);
  }, [closingDate, closingsLoaded]);

  function customerNameFor(id) {
    const c = customers.find((x) => x.id === id);
    return c ? c.name : "Customer";
  }

  // Distinct transporter combos this customer has ever used, most recent
  // first - lets the owner reuse "the usual transport" for a repeat buyer
  // instead of retyping it every time. Looks across every invoice ever saved
  // (no time limit / no cap), since a customer may want to look up a
  // transporter they used years ago, not just recently.
  function pastTransportsFor(customerId) {
    if (!customerId) return [];
    const seen = new Set();
    const out = [];
    for (const inv of invoices) {
      if (inv.customer_id !== customerId) continue;
      if (!inv.transport_name && !inv.vehicle_number) continue;
      const key = `${inv.transport_name || ""}|${inv.transport_contact || ""}|${inv.vehicle_number || ""}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ transport_name: inv.transport_name || "", transport_contact: inv.transport_contact || "", vehicle_number: inv.vehicle_number || "", invoice_date: inv.invoice_date });
    }
    return out;
  }

  function startNewInvoice() {
    setInvoiceDraft({
      customer_id: customers[0]?.id || "",
      invoice_date: new Date().toISOString().slice(0, 10),
      items: [{ description: "", qty: 1, rate: "" }],
      gstMode: "same",
      gstRate: 18,
      notes: "",
      transport_name: "",
      transport_contact: "",
      vehicle_number: "",
    });
    setInvoiceError(null);
    setInvoiceView("new");
  }

  function updateInvoiceItem(idx, field, value) {
    setInvoiceDraft((d) => ({ ...d, items: d.items.map((it, i) => (i === idx ? { ...it, [field]: value } : it)) }));
  }
  function addInvoiceItem() {
    setInvoiceDraft((d) => ({ ...d, items: [...d.items, { description: "", qty: 1, rate: "" }] }));
  }
  function removeInvoiceItem(idx) {
    setInvoiceDraft((d) => ({ ...d, items: d.items.filter((_, i) => i !== idx) }));
  }

  async function saveNewCustomer() {
    setCustomerSaving(true);
    setCustomerError(null);
    try {
      if (!newCustomer.name.trim()) throw new Error("Customer name is required.");
      const saved = await customersInsert(session.access_token, {
        user_id: effectiveOwnerId,
        name: newCustomer.name.trim(),
        gstin: newCustomer.gstin || null,
        address: newCustomer.address || null,
        phone: newCustomer.phone || null,
      });
      setCustomers((prev) => [saved, ...prev]);
      setInvoiceDraft((d) => (d ? { ...d, customer_id: saved.id } : d));
      setNewCustomer({ name: "", gstin: "", address: "", phone: "" });
      setNewCustomerMode(false);
    } catch (e) {
      setCustomerError(e.message || "Could not save customer - please try again.");
    } finally {
      setCustomerSaving(false);
    }
  }

  async function saveInvoice() {
    setInvoiceSaving(true);
    setInvoiceError(null);
    try {
      if (!invoiceDraft.customer_id) throw new Error("Please choose or add a customer.");
      const validItems = invoiceDraft.items.filter((it) => it.description.trim() && Number(it.rate) > 0);
      if (validItems.length === 0) throw new Error("Add at least one item with a description and rate.");
      const totals = computeInvoiceTotals(validItems, invoiceDraft.gstMode, invoiceDraft.gstRate);
      const invoiceNumber = `INV-${String(invoices.length + 1).padStart(4, "0")}`;
      const saved = await invoicesInsert(session.access_token, {
        user_id: effectiveOwnerId,
        created_by: session.user.id,
        customer_id: invoiceDraft.customer_id,
        invoice_number: invoiceNumber,
        invoice_date: invoiceDraft.invoice_date,
        items: validItems.map((it) => ({
          description: it.description,
          qty: Number(it.qty) || 0,
          rate: Number(it.rate) || 0,
          amount: (Number(it.qty) || 0) * (Number(it.rate) || 0),
        })),
        subtotal: totals.subtotal,
        cgst: totals.cgst,
        sgst: totals.sgst,
        igst: totals.igst,
        total: totals.total,
        status: "unpaid",
        notes: invoiceDraft.notes || null,
        transport_name: invoiceDraft.transport_name || null,
        transport_contact: invoiceDraft.transport_contact || null,
        vehicle_number: invoiceDraft.vehicle_number || null,
      });
      setInvoices((prev) => [saved, ...prev]);
      setInvoiceView("list");
      setInvoiceDraft(null);
    } catch (e) {
      setInvoiceError(e.message || "Could not save invoice - please try again.");
    } finally {
      setInvoiceSaving(false);
    }
  }

  // paymentMethod is only meaningful when marking an invoice paid (the
  // picker popover in the invoice list passes Cash/UPI/Card/Other) - it
  // feeds the Evening Closing tab's expected-cash breakdown below.
  async function toggleInvoiceStatus(inv, paymentMethod) {
    const newStatus = inv.status === "paid" ? "unpaid" : "paid";
    // Record when it was actually paid (not just the invoice date) so payment
    // delay/behavior can be computed per customer - reset it if un-marking.
    const paidDate = newStatus === "paid" ? new Date().toISOString().slice(0, 10) : null;
    const before = { status: inv.status, paid_date: inv.paid_date, payment_method: inv.payment_method };
    try {
      const updated = await invoicesUpdate(session.access_token, inv.id, {
        status: newStatus,
        paid_date: paidDate,
        payment_method: newStatus === "paid" ? paymentMethod || "Other" : null,
      });
      setInvoices((prev) => prev.map((x) => (x.id === inv.id ? updated : x)));
      logChange("invoice", inv.id, "status_changed", `Marked ${inv.invoice_number || "an invoice"} ${newStatus}`, before, { status: newStatus, paid_date: paidDate, payment_method: updated.payment_method });
    } catch (e) {
      console.error(e);
    }
  }

  // --- Inventory (Phase 5) ---
  async function addInventoryItem() {
    if (!newItemName.trim()) return;
    setAddingItem(true);
    try {
      const row = await inventoryItemsInsert(session.access_token, {
        user_id: effectiveOwnerId,
        created_by: session.user.id,
        name: newItemName.trim(),
        unit: newItemUnit.trim() || "pcs",
        quantity_on_hand: Number(newItemQty) || 0,
        reorder_level: newItemReorder === "" ? 0 : Number(newItemReorder),
      });
      setInventoryItems((prev) => [...prev, row].sort((a, b) => a.name.localeCompare(b.name)));
      setNewItemName("");
      setNewItemUnit("pcs");
      setNewItemQty("");
      setNewItemReorder("");
    } catch (e) {
      console.error(e);
    }
    setAddingItem(false);
  }

  async function deleteInventoryItem(id) {
    try {
      await inventoryItemsDelete(session.access_token, id);
      setInventoryItems((prev) => prev.filter((it) => it.id !== id));
    } catch (e) {
      console.error(e);
    }
  }

  // Logs a stock in/out movement and updates the item's running quantity in
  // the same action - the two rows (item + movement) are the whole of the
  // Phase 5 MVP's "transaction chain": no automatic linking to bills/invoices
  // yet, just a fast manual adjustment with a reason for the audit trail.
  async function adjustInventoryQty(item) {
    const delta = Number(adjustQtyValue);
    if (!delta) return;
    try {
      const newQty = Number(item.quantity_on_hand) + delta;
      const updated = await inventoryItemsUpdate(session.access_token, item.id, { quantity_on_hand: newQty });
      const movement = await inventoryMovementsInsert(session.access_token, {
        user_id: effectiveOwnerId,
        created_by: session.user.id,
        item_id: item.id,
        change_qty: delta,
        reason: adjustQtyReason.trim() || (delta > 0 ? "Stock in" : "Stock out"),
      });
      setInventoryItems((prev) => prev.map((it) => (it.id === item.id ? updated : it)));
      setInventoryMovements((prev) => [movement, ...prev]);
      setAdjustQtyFor(null);
      setAdjustQtyValue("");
      setAdjustQtyReason("");
    } catch (e) {
      console.error(e);
    }
  }

  async function askInventoryAdvice() {
    setInventoryAdviceLoading(true);
    try {
      const lowStock = inventoryItems.filter((it) => Number(it.quantity_on_hand) <= Number(it.reorder_level || 0));
      const res = await getInventoryAdvice(session.access_token, {
        items: inventoryItems.map((it) => ({ name: it.name, unit: it.unit, quantity_on_hand: it.quantity_on_hand, reorder_level: it.reorder_level })),
        lowStockCount: lowStock.length,
      });
      setInventoryAdvice(res.answer);
    } catch (e) {
      console.error(e);
      setInventoryAdvice("Couldn't get advice right now - please try again.");
    }
    setInventoryAdviceLoading(false);
  }

  // Sums invoices paid in cash on a given date - the baseline for the
  // Evening Closing "expected cash" figure. Only counts invoices whose
  // payment_method was captured as Cash (UPI/Card don't touch the till).
  function expectedCashFor(dateStr) {
    return invoices
      .filter((inv) => inv.status === "paid" && inv.paid_date === dateStr && inv.payment_method === "Cash")
      .reduce((s, inv) => s + (Number(inv.total) || 0), 0);
  }
  function closingFor(dateStr) {
    return dailyClosings.find((c) => c.closing_date === dateStr) || null;
  }

  // Payment Behaviour Profile for one customer - built entirely from
  // invoices already stored (no new data source). avgDelay only counts
  // invoices marked paid after this feature shipped (paid_date is set going
  // forward); older paid invoices without a paid_date are excluded from the
  // delay average but still count toward invoice totals.
  function paymentProfileFor(customerId) {
    const custInvoices = invoices.filter((inv) => inv.customer_id === customerId);
    const paidWithDate = custInvoices
      .filter((inv) => inv.status === "paid" && inv.paid_date)
      .sort((a, b) => new Date(a.invoice_date) - new Date(b.invoice_date));
    const delayOf = (inv) => Math.round((new Date(inv.paid_date) - new Date(inv.invoice_date)) / 86400000);
    const delays = paidWithDate.map(delayOf).filter((d) => Number.isFinite(d) && d >= 0);
    const avgDelay = delays.length ? Math.round(delays.reduce((s, d) => s + d, 0) / delays.length) : null;
    let trend = "steady";
    if (delays.length >= 4) {
      const mid = Math.floor(delays.length / 2);
      const avg = (arr) => arr.reduce((s, d) => s + d, 0) / arr.length;
      const diff = avg(delays.slice(mid)) - avg(delays.slice(0, mid));
      if (diff > 3) trend = "worse";
      else if (diff < -3) trend = "better";
    }
    const unpaid = custInvoices.filter((inv) => inv.status !== "paid");
    const outstanding = unpaid.reduce((s, inv) => s + (Number(inv.total) || 0), 0);
    const today = new Date();
    const maxOverdueDays = unpaid.length
      ? Math.max(...unpaid.map((inv) => Math.round((today - new Date(inv.invoice_date)) / 86400000)))
      : 0;
    return { invoiceCount: custInvoices.length, unpaidCount: unpaid.length, outstanding, avgDelay, trend, maxOverdueDays };
  }

  // Bahi Business Trust Score: a 0-100 credit-worthiness score built entirely
  // from a customer's own payment history already in the invoices table - no
  // external bureau, no new data source. Starts at 100 and loses points for
  // slow/overdue payment; a customer with too little history to judge gets
  // no score at all rather than a misleadingly neutral one.
  function trustScoreFor(profile) {
    if (profile.invoiceCount === 0) return null;
    const paidCount = profile.invoiceCount - profile.unpaidCount;
    if (paidCount === 0 && profile.unpaidCount < 2) return null; // too little history yet
    let score = 100;
    if (profile.avgDelay != null) score -= Math.min(50, profile.avgDelay * 1.5);
    if (profile.trend === "worse") score -= 10;
    else if (profile.trend === "better") score += 5;
    if (profile.invoiceCount > 0) score -= Math.round((profile.unpaidCount / profile.invoiceCount) * 20);
    if (profile.maxOverdueDays > 0) score -= Math.min(20, profile.maxOverdueDays / 3);
    score = Math.max(0, Math.min(100, Math.round(score)));
    const label = score >= 80 ? "Excellent" : score >= 60 ? "Good" : score >= 40 ? "Fair" : score >= 20 ? "Poor" : "Risky";
    const color = score >= 80 ? C.green : score >= 60 ? C.green : score >= 40 ? C.amber : C.red;
    return { score, label, color };
  }

  // Customer Churn Radar (Roadmap 3 #14): detects customers whose ordering
  // cadence has lapsed vs their own normal pattern - same gap-analysis
  // approach as Recurring Vendor detection, applied to a customer's invoice
  // dates instead of a vendor's bill dates. Needs at least 3 invoices with a
  // reasonably consistent gap to establish "normal" before it will flag
  // anything, so a customer with too little history is silently skipped
  // rather than falsely flagged.
  function customerChurnRadar() {
    const results = [];
    const today = new Date();
    for (const c of customers) {
      const rows = invoices
        .filter((inv) => inv.customer_id === c.id && inv.invoice_date)
        .slice()
        .sort((a, b) => new Date(a.invoice_date) - new Date(b.invoice_date));
      if (rows.length < 3) continue;
      const gaps = [];
      for (let i = 1; i < rows.length; i++) {
        const d = Math.round((new Date(rows[i].invoice_date) - new Date(rows[i - 1].invoice_date)) / 86400000);
        if (d > 0) gaps.push(d);
      }
      if (gaps.length < 2) continue;
      const avgGap = gaps.reduce((s, d) => s + d, 0) / gaps.length;
      const stdDev = Math.sqrt(gaps.reduce((s, d) => s + Math.pow(d - avgGap, 2), 0) / gaps.length);
      if (avgGap < 4 || stdDev / avgGap > 0.45) continue; // too irregular to call "normal"
      const lastRow = rows[rows.length - 1];
      const daysSinceLast = Math.round((today - new Date(lastRow.invoice_date)) / 86400000);
      if (daysSinceLast <= avgGap * 1.6) continue; // still within a normal-ish window
      const amounts = rows.map((r) => Number(r.total) || 0);
      const avgAmount = amounts.reduce((s, a) => s + a, 0) / amounts.length;
      const cadenceLow = Math.round(avgGap - stdDev);
      const cadenceHigh = Math.round(avgGap + stdDev);
      const revenueAtRisk = (avgAmount / avgGap) * 365;
      results.push({
        customer: c,
        cadenceLow: Math.max(1, cadenceLow),
        cadenceHigh: Math.max(cadenceLow + 1, cadenceHigh),
        daysSinceLast,
        avgAmount,
        revenueAtRiskAnnual: revenueAtRisk,
        lastOrderDate: lastRow.invoice_date,
      });
    }
    return results.sort((a, b) => b.revenueAtRiskAnnual - a.revenueAtRiskAnnual);
  }

  // Receivable Probability (Phase 9): buckets every unpaid invoice by how
  // likely it is to be collected soon, using that customer's own average
  // days-to-pay rather than one flat overdue cutoff for everyone - a
  // customer who reliably pays in 45 days isn't "at risk" on day 40, but a
  // customer whose average is 15 days and is now at 40 almost certainly is.
  // Customers with no payment history yet fall back to a neutral 30-day
  // assumption rather than being silently excluded.
  function receivableForecast() {
    const unpaid = invoices.filter((inv) => inv.status !== "paid" && inv.total != null);
    const today = new Date();
    let likelySoon = 0, atRisk = 0, dueLater = 0;
    const rows = unpaid.map((inv) => {
      const profile = paymentProfileFor(inv.customer_id);
      const ageDays = Math.round((today - new Date(inv.invoice_date)) / 86400000);
      const expectedDelay = profile.avgDelay != null ? profile.avgDelay : 30;
      const daysRemaining = expectedDelay - ageDays;
      let bucket;
      if (profile.trend === "worse" || ageDays > expectedDelay * 1.5) bucket = "at_risk";
      else if (daysRemaining <= 7) bucket = "likely_soon";
      else bucket = "due_later";
      const total = Number(inv.total) || 0;
      if (bucket === "likely_soon") likelySoon += total;
      else if (bucket === "at_risk") atRisk += total;
      else dueLater += total;
      const customer = customers.find((c) => c.id === inv.customer_id);
      return { invoice: inv, bucket, ageDays, expectedDelay, customerName: customer ? customer.name : "Unknown customer", total };
    });
    return { rows, likelySoon, atRisk, dueLater, total: likelySoon + atRisk + dueLater };
  }

  // Customer Profitability proxy (Phase 10): BahiSathi doesn't track
  // per-sale cost of goods, discounts, or logistics cost, so a true
  // profit-margin ranking isn't possible from what's recorded today - this
  // is a payment-quality-adjusted revenue proxy instead: paid revenue minus
  // a notional cost of capital for that customer's own average payment
  // delay, minus a discount for their currently at-risk/likely-soon
  // outstanding balance. A ₹5L customer who pays in 90 days and has ₹1L
  // stuck at-risk can rank behind a ₹2L customer who always pays in 10 days
  // - directionally useful, not a substitute for real margin data once
  // purchase costs are tracked per sale.
  function customerProfitabilityProxy() {
    const COST_OF_CAPITAL_ANNUAL = 0.12;
    const rf = receivableForecast();
    return customers
      .map((c) => {
        const paid = invoices.filter((inv) => inv.customer_id === c.id && inv.status === "paid");
        const paidRevenue = paid.reduce((s, inv) => s + (Number(inv.subtotal) || Number(inv.total) || 0), 0);
        if (paidRevenue <= 0) return null;
        const profile = paymentProfileFor(c.id);
        const carryingCost = profile.avgDelay != null ? paidRevenue * (profile.avgDelay / 365) * COST_OF_CAPITAL_ANNUAL : 0;
        const riskRows = rf.rows.filter((r) => r.invoice.customer_id === c.id);
        const riskDiscount = riskRows.reduce((s, r) => s + (r.bucket === "at_risk" ? r.total * 0.3 : r.bucket === "likely_soon" ? r.total * 0.05 : 0), 0);
        const effectiveValue = paidRevenue - carryingCost - riskDiscount;
        return { customer: c, paidRevenue, carryingCost, riskDiscount, effectiveValue, avgDelay: profile.avgDelay };
      })
      .filter(Boolean)
      .sort((a, b) => b.effectiveValue - a.effectiveValue);
  }

  // "Can I give him credit?" - layered on the Trust Score above. Sends the
  // already-computed payment profile (not raw invoices) to a small AI
  // endpoint so the backend doesn't need to re-query invoices/customers.
  async function askCreditAdvice(customer, profile, trust) {
    setCreditAdviceFor(customer.id);
    setCreditAdviceLoading(true);
    try {
      const result = await getCreditAdvice(session.access_token, {
        customerName: customer.name,
        profile: {
          invoiceCount: profile.invoiceCount,
          unpaidCount: profile.unpaidCount,
          outstanding: profile.outstanding,
          avgDelay: profile.avgDelay,
          trend: profile.trend,
          maxOverdueDays: profile.maxOverdueDays,
          trustScore: trust ? trust.score : null,
        },
      });
      setCreditAdviceText((prev) => ({ ...prev, [customer.id]: result.answer }));
    } catch (e) {
      setCreditAdviceText((prev) => ({ ...prev, [customer.id]: "Could not get advice right now - please try again." }));
    } finally {
      setCreditAdviceLoading(false);
    }
  }

  // "Explain my P&L" - sends the already-computed P&L figures (see
  // pnlForPeriod above) to a small AI endpoint for a plain-language writeup.
  async function explainPnl(pnlData, periodLabel, prevPnl) {
    setPnlNarrationLoading(true);
    setPnlNarration(null);
    try {
      const result = await getPnlNarration(session.access_token, {
        pnl: {
          period: periodLabel,
          revenue: pnlData.revenue,
          expenses: pnlData.expenses,
          net: pnlData.net,
          topCategories: pnlData.topCategories,
          revenueDelta: prevPnl ? pnlData.revenue - prevPnl.revenue : 0,
          expenseDelta: prevPnl ? pnlData.expenses - prevPnl.expenses : 0,
        },
      });
      setPnlNarration(result.answer);
    } catch (e) {
      setPnlNarration("Could not generate an explanation right now - please try again.");
    } finally {
      setPnlNarrationLoading(false);
    }
  }

  async function deleteInvoiceRow(id) {
    try {
      await invoicesDelete(session.access_token, id);
      setInvoices((prev) => prev.filter((x) => x.id !== id));
    } catch (e) {
      console.error(e);
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

  async function deleteAccount() {
    setDeleting(true);
    setDeleteError(null);
    try {
      await deleteAccountApi(session.access_token);
      onLogout();
    } catch (e) {
      setDeleteError(e.message || "Could not delete account - please try again.");
      setDeleting(false);
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

  // Plain-language manual transaction ("Paid Rs.18,500 rent to Sharma
  // Properties") - parsed by the same backend Claude call as bill photos,
  // just from typed text, then dropped into the exact same review/edit/save
  // flow (setDraft + step "review") so there's no separate UI to maintain.
  async function submitManualEntry() {
    const text = manualEntryText.trim();
    if (!text) return;
    setManualEntryError(null);
    setManualEntryLoading(true);
    try {
      const parsed = await parseManualEntry(session.access_token, text);
      if (parsed.total == null) {
        setManualEntryError("Couldn't find an amount in that - try including the number, e.g. \"Paid 18500 rent to Sharma Properties\".");
        return;
      }
      setDraft({ ...parsed, items: parsed.items || [] });
      setStep("review");
      setManualEntryOpen(false);
      setManualEntryText("");
    } catch (e) {
      setManualEntryError(e.message || "Could not understand that - please try rephrasing.");
    } finally {
      setManualEntryLoading(false);
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
        user_id: effectiveOwnerId,
        created_by: session.user.id,
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
        is_personal: !!draft.is_personal,
      });
      setEntries((prev) => [...prev, { id: saved.id, vendor: saved.vendor, date: saved.bill_date, gstin: saved.gstin, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst, total: saved.total, category: saved.category, items: saved.items, vendor_source: saved.vendor_source, confidence: saved.confidence, is_personal: saved.is_personal, created_at: saved.created_at }]);
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

  // Audit Trail (Phase 10) - persists to the audit_log table and prepends
  // to local state. The underlying bill/invoice change has already been
  // applied by the caller before this runs, so a failure here (most likely
  // because the audit_log table hasn't been created yet - see the SQL note
  // above auditLogList) never blocks the actual save/delete/toggle.
  async function logChange(entityType, entityId, action, summary, before, after) {
    try {
      const saved = await auditLogInsert(session.access_token, {
        user_id: effectiveOwnerId,
        created_by: session.user.id,
        entity_type: entityType,
        entity_id: String(entityId),
        action,
        summary,
        before: before ?? null,
        after: after ?? null,
      });
      setAuditLog((prev) => [saved, ...prev].slice(0, 50));
    } catch (e) {
      console.error("Audit log not available yet:", e.message);
    }
  }

  // Reverts a bill/invoice back to its "before" snapshot from an audit_log
  // row - works regardless of how long ago the change was made, since it
  // reads from Supabase rather than in-memory state. Deleted bills come
  // back as a new row (the old id no longer exists in the database);
  // edits and status changes revert in place.
  async function restoreAuditEntry(log) {
    setRestoringLogId(log.id);
    try {
      if (log.entity_type === "bill" && log.action === "deleted") {
        const b = log.before;
        const restored = await billsInsert(session.access_token, {
          user_id: effectiveOwnerId,
          created_by: session.user.id,
          vendor: b.vendor,
          bill_date: b.date,
          gstin: b.gstin,
          cgst: b.cgst,
          sgst: b.sgst,
          igst: b.igst,
          total: b.total,
          category: b.category,
          items: b.items,
          is_personal: b.is_personal,
        });
        setEntries((prev) => [...prev, { id: restored.id, vendor: restored.vendor, date: restored.bill_date, gstin: restored.gstin, cgst: restored.cgst, sgst: restored.sgst, igst: restored.igst, total: restored.total, category: restored.category, items: restored.items, vendor_source: restored.vendor_source, confidence: restored.confidence, is_personal: restored.is_personal, created_at: restored.created_at, created_by: restored.created_by }]);
        await logChange("bill", restored.id, "restored", `Restored: ${log.summary}`, log.after, log.before);
      } else if (log.entity_type === "bill") {
        const b = log.before;
        const saved = await billsUpdate(session.access_token, log.entity_id, {
          vendor: b.vendor,
          bill_date: b.date,
          category: b.category,
          items: b.items,
          total: b.total,
          cgst: b.cgst,
          sgst: b.sgst,
          igst: b.igst,
          is_personal: b.is_personal,
        });
        setEntries((prev) => prev.map((e) => (e.id === log.entity_id ? { ...e, vendor: saved.vendor, date: saved.bill_date, category: saved.category, items: saved.items, total: saved.total, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst, is_personal: saved.is_personal } : e)));
        await logChange("bill", log.entity_id, "restored", `Restored: ${log.summary}`, log.after, log.before);
      } else if (log.entity_type === "invoice") {
        const saved = await invoicesUpdate(session.access_token, log.entity_id, log.before);
        setInvoices((prev) => prev.map((x) => (x.id === log.entity_id ? saved : x)));
        await logChange("invoice", log.entity_id, "restored", `Restored: ${log.summary}`, log.after, log.before);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setRestoringLogId(null);
    }
  }

  function startEditEntry(entry) {
    setDeleteConfirmId(null);
    setEditError(null);
    setEditingId(entry.id);
    setEditDraft({ vendor: entry.vendor || "", date: entry.date || "", category: entry.category || "other", total: entry.total ?? "", cgst: entry.cgst ?? "", sgst: entry.sgst ?? "", igst: entry.igst ?? "", items: Array.isArray(entry.items) ? entry.items : [], is_personal: !!entry.is_personal });
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
      const before = entries.find((e) => e.id === id);
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
        is_personal: !!editDraft.is_personal,
      };
      const saved = await billsUpdate(session.access_token, id, patch);
      setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, vendor: saved.vendor, date: saved.bill_date, category: saved.category, items: saved.items, total: saved.total, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst, is_personal: saved.is_personal } : e)));
      if (before) {
        logChange(
          "bill",
          id,
          "updated",
          `Edited ${before.vendor || "a bill"}`,
          { vendor: before.vendor, date: before.date, category: before.category, total: before.total, cgst: before.cgst, sgst: before.sgst, igst: before.igst, is_personal: before.is_personal, items: before.items },
          { vendor: saved.vendor, date: saved.bill_date, category: saved.category, total: saved.total, cgst: saved.cgst, sgst: saved.sgst, igst: saved.igst, is_personal: saved.is_personal, items: saved.items }
        );
      }
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
      const before = entries.find((e) => e.id === id);
      await billsDelete(session.access_token, id);
      setEntries((prev) => prev.filter((e) => e.id !== id));
      if (editingId === id) {
        setEditingId(null);
        setEditDraft(null);
      }
      if (before) {
        logChange(
          "bill",
          id,
          "deleted",
          `Deleted ${before.vendor || "a bill"}`,
          { vendor: before.vendor, date: before.date, gstin: before.gstin, category: before.category, total: before.total, cgst: before.cgst, sgst: before.sgst, igst: before.igst, is_personal: before.is_personal, items: before.items },
          null
        );
      }
    } catch (e) {
      setEditError(e.message || "Could not delete this entry - please try again.");
    } finally {
      setDeleteConfirmId(null);
    }
  }

  async function askAboutSpending(questionOverride) {
    const q = (questionOverride ?? askQuestion).trim();
    if (!q) return;
    setAskQuestion(q);
    setAskLoading(true);
    setAskError(null);
    setAskAnswer(null);
    try {
      const result = await getSpendSummary(session.access_token, q);
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

  // P&L statement export (Phase 9) - same header-row pattern as the ledger
  // export above, handed to a CA as an actual statement rather than a
  // screenshot of the Health tab.
  function exportPnlTrend(trend) {
    const rows = trend.map((t) => ({
      Month: t.label,
      Revenue: Math.round(t.revenue),
      Expenses: Math.round(t.expenses),
      Net: Math.round(t.net),
      Invoices: t.invoiceCount,
      Bills: t.billCount,
    }));
    const wb = XLSX.utils.book_new();
    let ws;
    if (businessProfile && (businessProfile.business_name || businessProfile.gstin)) {
      const headerRows = [
        [businessProfile.business_name || "BahiSathi P&L"],
        ...(businessProfile.gstin ? [[`GSTIN: ${businessProfile.gstin}`]] : []),
        ["Profit & Loss - last 6 months"],
        [],
      ];
      ws = XLSX.utils.aoa_to_sheet(headerRows);
      XLSX.utils.sheet_add_json(ws, rows, { origin: -1 });
    } else {
      ws = XLSX.utils.json_to_sheet(rows);
    }
    XLSX.utils.book_append_sheet(wb, ws, "P&L");
    XLSX.writeFile(wb, "bahisathi-pnl.xlsx");
  }

  const monthTotal = entries.reduce((sum, e) => sum + (Number(e.total) || 0), 0);
  const personalTotal = entries.reduce((sum, e) => sum + (e.is_personal ? Number(e.total) || 0 : 0), 0);

  // Phase 6 attribution: who actually added this row. Only meaningful for
  // the owner (staffList is only loaded for owners) - returns null for the
  // owner's own entries and pre-Phase-6 rows with no created_by, so the tag
  // only shows up when it's actually informative (a staff member added it).
  function addedByLabel(createdBy) {
    if (!createdBy || createdBy === session.user.id) return null;
    const s = staffList.find((x) => x.staff_user_id === createdBy);
    return s ? s.name || s.phone : null;
  }

  return (
    <div>
      {(() => {
        const caTabs = ["ledger", "gst", "invoices", "pnl", "audit"];
        // Owner Vault (Phase 6): staff logins never see profit/health/GST/
        // audit/business-settings tabs - just the day-to-day operational
        // ones, and the CA-mode toggle doesn't apply to them at all.
        const staffTabs = ["scan", "ledger", "invoices", "inventory"];
        const tabLabel = (key) =>
          key === "action" ? "Action Center" : key === "scan" ? "Scan bill" : key === "ledger" ? "Ledger" : key === "gst" ? "GST" : key === "invoices" ? "Invoices" : key === "bank" ? "Bank" : key === "inventory" ? "Inventory" : key === "closing" ? "Closing" : key === "health" ? "Health" : key === "pnl" ? "P&L" : key === "memory" ? "Memory" : key === "whatif" ? "What if?" : key === "audit" ? "Audit" : key === "plan" ? "Plan" : "Business";

        const rightControls = (
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {role === "owner" && (
              <button
                onClick={() => {
                  const next = viewMode === "owner" ? "ca" : "owner";
                  setViewMode(next);
                  if (next === "ca" && !caTabs.includes(tab)) setTab("ledger");
                }}
                title={viewMode === "owner" ? "Switch to a simplified, read-focused view for your CA/accountant" : "Switch back to the full owner view"}
                style={{ background: viewMode === "ca" ? C.gold : "transparent", border: "1px solid #5C6B87", color: viewMode === "ca" ? C.white : "#B9C3D6", borderRadius: 20, padding: "5px 14px", fontSize: 12.5, cursor: "pointer", fontFamily: sans }}
              >
                {t(viewMode === "ca" ? "👤 CA mode" : "🧑‍💼 Owner mode")}
              </button>
            )}
            <LangToggle />
            <button onClick={onLogout} style={{ background: "transparent", border: "1px solid #5C6B87", color: "#B9C3D6", borderRadius: 20, padding: "5px 14px", fontSize: 12.5, cursor: "pointer", fontFamily: sans }}>
              {t("Log out")}
            </button>
          </div>
        );

        // CA mode and staff logins are already narrow (5 and 4 tabs), so
        // they keep the original flat tab bar - the Home/Ask/To-do/Books
        // split below is only needed for the owner view, where the tab
        // count keeps growing as features ship.
        if (role === "staff" || viewMode === "ca") {
          const visibleTabs = role === "staff" ? staffTabs : caTabs;
          return (
            <div style={{ display: "flex", background: C.ink, paddingLeft: 24, gap: 4, justifyContent: "space-between", alignItems: "center", paddingRight: 24 }}>
              <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                {visibleTabs.map((key) => (
                  <button key={key} onClick={() => setTab(key)} style={{ background: tab === key ? C.paper : "transparent", color: tab === key ? C.ink : "#B9C3D6", border: "none", borderRadius: "8px 8px 0 0", padding: "9px 18px", fontSize: 13.5, fontWeight: 500, cursor: "pointer", fontFamily: sans }}>
                    {t(tabLabel(key))}
                  </button>
                ))}
                {sub?.plan === "pro" && (
                  <span style={{ background: C.gold, color: C.white, fontSize: 10.5, fontWeight: 600, borderRadius: 20, padding: "3px 10px", letterSpacing: 0.3 }}>PRO</span>
                )}
              </div>
              {rightControls}
            </div>
          );
        }

        const SECTIONS = [
          { key: "home", tab: "home", icon: "🏠", label: "Home" },
          { key: "ask", tab: "ask", icon: "💬", label: "Ask BahiSathi" },
          { key: "todo", tab: "action", icon: "✅", label: "To-do" },
          { key: "books", tab: booksTab, icon: "📒", label: "Books" },
        ];
        const activeSection = tab === "home" ? "home" : tab === "ask" ? "ask" : tab === "action" ? "todo" : "books";

        return (
          <div>
            <div style={{ display: "flex", background: C.ink, paddingLeft: 24, gap: 4, justifyContent: "space-between", alignItems: "center", paddingRight: 24 }}>
              <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                {SECTIONS.map((s) => (
                  <button key={s.key} onClick={() => setTab(s.tab)} style={{ background: activeSection === s.key ? C.paper : "transparent", color: activeSection === s.key ? C.ink : "#B9C3D6", border: "none", borderRadius: "8px 8px 0 0", padding: "9px 18px", fontSize: 13.5, fontWeight: 500, cursor: "pointer", fontFamily: sans }}>
                    {s.icon} {t(s.label)}
                  </button>
                ))}
                {sub?.plan === "pro" && (
                  <span style={{ background: C.gold, color: C.white, fontSize: 10.5, fontWeight: 600, borderRadius: 20, padding: "3px 10px", letterSpacing: 0.3 }}>PRO</span>
                )}
              </div>
              {rightControls}
            </div>
            {activeSection === "books" && (
              <div style={{ display: "flex", background: C.inkLight, paddingLeft: 24, paddingRight: 24, gap: 4, overflowX: "auto" }}>
                {BOOKS_TABS.map((key) => (
                  <button key={key} onClick={() => setTab(key)} style={{ background: "transparent", color: tab === key ? C.white : "#C7D2E6", border: "none", borderBottom: tab === key ? `2px solid ${C.gold}` : "2px solid transparent", padding: "8px 14px", fontSize: 12.5, fontWeight: 500, cursor: "pointer", fontFamily: sans, whiteSpace: "nowrap" }}>
                    {t(tabLabel(key))}
                  </button>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      <div style={{ padding: 24, maxWidth: 640, margin: "0 auto" }}>
        {viewMode === "ca" && (
          <div style={{ background: "#FDF3DC", border: `1px solid ${C.gold}`, borderRadius: 8, padding: "10px 14px", fontSize: 12.5, color: C.text, marginBottom: 18 }}>
            CA mode - a simplified, read-focused view of the same books. Editing and new entries are turned off here; switch back to Owner mode to make changes.
          </div>
        )}
        {role === "staff" && (
          <div style={{ background: "#EAF0FF", border: `1px solid ${C.inkLight}`, borderRadius: 8, padding: "10px 14px", fontSize: 12.5, color: C.text, marginBottom: 18 }}>
            Staff access - you can scan bills, log ledger entries, and manage invoices and inventory. Profit, health, GST, audit, and business settings are only visible to the owner.
          </div>
        )}
        {tab === "ask" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>{t("Ask BahiSathi")}</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18 }}>{t("About your spending, or how BahiSathi works - ask anything.")}</div>
            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "14px 16px" }}>
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
                  {askLoading ? t("Thinking…") : t("Ask")}
                </button>
              </div>
              {!askAnswer && entries.length > 0 && (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 10 }}>
                  {["Why did my spend go up this month?", "Which category jumped the most?"].map((q) => (
                    <button key={q} onClick={() => askAboutSpending(q)} disabled={askLoading} style={{ background: C.paper, border: `1px solid ${C.cardBorder}`, color: C.text, borderRadius: 20, padding: "5px 11px", fontSize: 11.5, cursor: askLoading ? "default" : "pointer" }}>
                      {q}
                    </button>
                  ))}
                </div>
              )}
              {askError && <div style={{ fontSize: 12.5, color: C.red, marginTop: 8 }}>{askError}</div>}
              {askAnswer && (
                <div style={{ marginTop: 10, background: "#E7FFDB", border: "1px solid #BCE5A8", borderRadius: "10px 10px 10px 2px", padding: "10px 14px", fontSize: 13, color: "#1F2C1A", lineHeight: 1.55 }}>
                  {askAnswer}
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "home" && (() => {
          const cashFlow = cashFlowForecast(entries, invoices, businessProfile);
          const trustScores = customers.map((c) => trustScoreFor(paymentProfileFor(c.id))).filter((s) => s != null);
          const collectionsScore = trustScores.length ? Math.round(trustScores.reduce((s, tt) => s + tt.score, 0) / trustScores.length) : 75;
          const cashScore = cashFlow.status === "healthy" ? 100 : cashFlow.status === "insufficient_data" ? 60 : cashFlow.status === "watch" ? 70 : cashFlow.status === "at_risk" ? 40 : 15;
          const gstScore = cashFlow.available == null ? 75 : cashFlow.available >= 0 ? 100 : cashFlow.available > -cashFlow.netReserve ? 40 : 10;
          const overall = Math.round(cashScore * 0.4 + collectionsScore * 0.35 + gstScore * 0.25);
          const overallLabel = overall >= 80 ? "Strong" : overall >= 60 ? "Steady" : overall >= 40 ? "Needs attention" : "At risk";
          const overallColor = overall >= 80 ? C.green : overall >= 60 ? C.green : overall >= 40 ? C.amber : C.red;
          const mm = moneyMapFor(businessProfile, invoices);
          const sts = safeToSpendFor(businessProfile, entries, invoices);

          // "Today's priority" - the single most urgent signal, using the
          // same underlying detectors as Action Center but only asking for
          // the top one instead of the full list, since Home is meant to be
          // a glance, not another list to scroll.
          const worstTrust = customers
            .map((c) => ({ c, tr: trustScoreFor(paymentProfileFor(c.id)) }))
            .filter((x) => x.tr && x.tr.score < 50)
            .sort((a, b) => a.tr.score - b.tr.score);
          const overdueRecurring = detectRecurringVendors(entries).filter((r) => r.status === "overdue");
          const rf = receivableForecast();
          let priority = null;
          if (cashFlow.status === "immediate_risk") priority = { text: t("Cash is at or below your GST reserve - review collections or spending today."), action: () => setTab("health") };
          else if (rf.atRisk > 0) priority = { text: `₹${Math.round(rf.atRisk).toLocaleString("en-IN")} ${t("in receivables looks at risk - worth a follow-up.")}`, action: () => setTab("invoices") };
          else if (overdueRecurring.length > 0) priority = { text: `${overdueRecurring[0].vendor} ${t("bill looks overdue.")}`, action: () => setTab("action") };
          else if (worstTrust.length > 0) priority = { text: `${worstTrust[0].c.name} ${t("has a low trust score - consider requesting advance payment.")}`, action: () => setTab("invoices") };
          else if (cashFlow.status === "at_risk" || cashFlow.status === "watch") priority = { text: t("Spending is outpacing income lately - worth a look."), action: () => setTab("health") };

          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>{t("Home")}</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>{t("Your business at a glance.")}</div>

              {priority && (
                <div
                  onClick={priority.action}
                  style={{ background: "#FDF3DC", border: `1px solid ${C.gold}`, borderRadius: 8, padding: "14px 16px", marginBottom: 16, cursor: "pointer" }}
                >
                  <div style={{ fontSize: 11, fontWeight: 600, color: "#5C4413", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 4 }}>{t("Today's priority")}</div>
                  <div style={{ fontSize: 13.5, color: "#3A2C0A" }}>{priority.text}</div>
                </div>
              )}

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginBottom: 16, display: "flex", alignItems: "center", gap: 18 }}>
                <div style={{ width: 64, height: 64, borderRadius: "50%", border: `5px solid ${overallColor}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <span style={{ fontFamily: display, fontSize: 19, fontWeight: 700, color: overallColor }}>{overall}</span>
                </div>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 600, color: C.text }}>{t("Business pulse")}: {overallLabel}</div>
                  <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3 }}>
                    {t("Cash flow")} {cashScore} · {t("Collections")} {collectionsScore} · {t("GST readiness")} {gstScore}
                  </div>
                  <button onClick={() => { setBooksTab("health"); setTab("health"); }} style={{ marginTop: 6, background: "transparent", border: "none", color: C.ink, textDecoration: "underline", fontSize: 11.5, cursor: "pointer", padding: 0, fontFamily: sans }}>
                    {t("Full health breakdown")}
                  </button>
                </div>
              </div>

              {sts.available != null && (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontSize: 13.5, fontWeight: 600, color: C.text }}>{t("Safe to spend")}</span>
                    <span style={{ fontFamily: mono, fontSize: 18, fontWeight: 700, color: sts.available < 0 ? C.red : C.green }}>₹{Math.round(sts.available).toLocaleString("en-IN")}</span>
                  </div>
                  <div style={{ fontSize: 11, color: C.textMuted, marginTop: 4 }}>{t("Bank balance minus your estimated GST reserve.")}</div>
                </div>
              )}

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>{t("Your business money")} {mm.hasCash ? `— ₹${Math.round(mm.total).toLocaleString("en-IN")}` : ""}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
                    <span style={{ color: C.textMuted }}>{t("Bank & Cash")}</span>
                    <span style={{ fontFamily: mono, color: C.text }}>{mm.hasCash ? `₹${Math.round(mm.cash).toLocaleString("en-IN")}` : "—"}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
                    <span style={{ color: C.textMuted }}>{t("Customers owe you")}</span>
                    <span style={{ fontFamily: mono, color: C.text }}>₹{Math.round(mm.outstanding).toLocaleString("en-IN")}</span>
                  </div>
                </div>
              </div>

              <button
                onClick={() => { setStep("upload"); setTab("scan"); }}
                style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 8, padding: "14px 0", fontSize: 14.5, fontWeight: 600, cursor: "pointer" }}
              >
                {t("Scan a bill")}
              </button>
            </div>
          );
        })()}

        {tab === "action" && (() => {
          // Same weighting as the Health tab's overall score, kept in sync
          // deliberately - Action Center is meant to be the front door onto
          // that same number, not a second opinion.
          const cashFlow = cashFlowForecast(entries, invoices, businessProfile);
          const trustScores = customers.map((c) => trustScoreFor(paymentProfileFor(c.id))).filter((s) => s != null);
          const collectionsScore = trustScores.length ? Math.round(trustScores.reduce((s, t) => s + t.score, 0) / trustScores.length) : 75;
          const cashScore = cashFlow.status === "healthy" ? 100 : cashFlow.status === "insufficient_data" ? 60 : cashFlow.status === "watch" ? 70 : cashFlow.status === "at_risk" ? 40 : 15;
          const gstScore = cashFlow.available == null ? 75 : cashFlow.available >= 0 ? 100 : cashFlow.available > -cashFlow.netReserve ? 40 : 10;
          const overall = Math.round(cashScore * 0.4 + collectionsScore * 0.35 + gstScore * 0.25);
          const overallLabel = overall >= 80 ? "Strong" : overall >= 60 ? "Steady" : overall >= 40 ? "Needs attention" : "At risk";
          const overallColor = overall >= 80 ? C.green : overall >= 60 ? C.green : overall >= 40 ? C.amber : C.red;

          const behavior = paymentBehaviorTrend(invoices);
          const gstFindings = gstGuardianFindings(entries, invoices, customers);
          const bookFindings = fixMyBooksFindings(entries);
          const recurring = detectRecurringVendors(entries).filter((r) => r.status === "overdue" || r.status === "due_soon");
          const priceWatch = vendorPriceWatch(entries);
          const leaks = expenseLeakFindings(entries);
          const worstTrust = customers
            .map((c) => ({ c, t: trustScoreFor(paymentProfileFor(c.id)) }))
            .filter((x) => x.t && x.t.score < 50)
            .sort((a, b) => a.t.score - b.t.score);
          const rf = receivableForecast();
          const missingMoney = missingMoneyFindings(invoices, bankTransactions, dailyClosings);
          const sinceYesterday = whatChangedForWindow(entries, invoices, dailyClosings, changedWindowDays);
          const gstReminders = getGstReminders();
          const calendar = buildBusinessCalendar(gstReminders, detectRecurringVendors(entries), rf.rows);
          const now2 = new Date();
          const pKey2 = `${now2.getFullYear()}-${String(now2.getMonth() + 1).padStart(2, "0")}`;
          const prevDate2 = new Date(now2.getFullYear(), now2.getMonth() - 1, 1);
          const prevKey2 = `${prevDate2.getFullYear()}-${String(prevDate2.getMonth() + 1).padStart(2, "0")}`;
          const monthPnl = pnlForPeriod(entries, invoices, pKey2);
          const prevMonthPnl = pnlForPeriod(entries, invoices, prevKey2);

          // Every signal the app already computes, folded into one list of
          // {severity, source, title, hint, ctaLabel, onCta} rows - nothing
          // here is a new finding, it's the existing tabs' own functions
          // called once more and merged, same "no new data store" approach
          // as everything else in this file.
          const items = [];
          if (cashFlow.status === "immediate_risk" || cashFlow.status === "at_risk") {
            items.push({
              severity: cashFlow.status === "immediate_risk" ? "high" : "medium",
              source: "Cash-Flow Guardian",
              title: cashFlow.status === "immediate_risk" ? "Cash is already at risk" : "Cash flow needs attention",
              hint: cashFlow.status === "immediate_risk" ? "Available cash is at or below your GST reserve - collections or spending decisions need attention now." : `Spending is outpacing income - at this pace, cash runs low in about ${cashFlow.monthsRunway.toFixed(1)} months.`,
              ctaLabel: "View Health",
              onCta: () => setTab("health"),
            });
          }
          if (gstScore <= 40 && cashFlow.available != null) {
            items.push({
              severity: gstScore <= 10 ? "high" : "medium",
              source: "GST readiness",
              title: "GST reserve may be short",
              hint: "Available cash may not fully cover your estimated GST liability for this period.",
              ctaLabel: "View Health",
              onCta: () => setTab("health"),
            });
          }
          if (behavior && behavior.trend === "worse") {
            items.push({
              severity: "medium",
              source: "Payment Behaviour",
              title: "Customers are paying slower lately",
              hint: `Averaging ${behavior.avgDelay} days across the last ${behavior.sampleSize} paid invoices - worth a nudge to regular late payers.`,
              ctaLabel: "View Memory",
              onCta: () => setTab("memory"),
            });
          }
          worstTrust.slice(0, 3).forEach(({ c, t }) => {
            items.push({
              severity: t.score < 30 ? "high" : "medium",
              source: "Trust Score",
              title: `${c.name}: ${t.label.toLowerCase()} trust score (${t.score})`,
              hint: "Based on this customer's own payment history - consider requesting advance payment on new orders.",
              ctaLabel: "View Invoices",
              onCta: () => setTab("invoices"),
            });
          });
          recurring.forEach((r) => {
            items.push({
              severity: r.status === "overdue" ? "high" : "medium",
              source: "Recurring",
              title: `${r.vendor} - ${r.status === "overdue" ? `${Math.abs(r.daysUntilDue)} day${Math.abs(r.daysUntilDue) === 1 ? "" : "s"} overdue` : `due in ${r.daysUntilDue} day${r.daysUntilDue === 1 ? "" : "s"}`}`,
              hint: `~₹${Math.round(r.avgAmount).toLocaleString("en-IN")}, billed ${r.cadence.toLowerCase()} based on your history.`,
              ctaLabel: quickAddedVendors.includes(r.vendor) ? "✓ Added" : quickAddingVendor === r.vendor ? "Adding…" : "+ Quick add",
              onCta: quickAddedVendors.includes(r.vendor) || quickAddingVendor === r.vendor ? null : () => quickAddRecurringBill(r),
              isQuickAdd: true,
            });
          });
          priceWatch.slice(0, 5).forEach((p) => {
            items.push({
              severity: p.pctChange >= 30 ? "medium" : "low",
              source: "Price Watch",
              title: `${p.vendor}${p.item.toLowerCase() !== p.vendor.toLowerCase() ? ` - ${p.item}` : ""}: up ${Math.round(p.pctChange)}%`,
              hint: `₹${Math.round(p.earlierAvg).toLocaleString("en-IN")} avg before → ₹${Math.round(p.latestAmount).toLocaleString("en-IN")} on ${p.latestDate}.`,
              ctaLabel: "View Audit",
              onCta: () => setTab("audit"),
            });
          });
          leaks.forEach((f) => {
            items.push({ severity: f.severity, source: "Leak Detector", title: f.label, hint: f.hint, ctaLabel: "View Audit", onCta: () => setTab("audit") });
          });
          gstFindings.forEach((f) => {
            items.push({ severity: f.severity, source: "GST Guardian", title: f.label, hint: f.hint, ctaLabel: "View Audit", onCta: () => setTab("audit") });
          });
          bookFindings.forEach((f) => {
            items.push({ severity: f.severity, source: "Fix My Books", title: f.label, hint: f.hint, ctaLabel: "View Audit", onCta: () => setTab("audit") });
          });
          if (rf.atRisk > 0) {
            items.push({
              severity: rf.atRisk > rf.likelySoon ? "high" : "medium",
              source: "Receivable Probability",
              title: `₹${Math.round(rf.atRisk).toLocaleString("en-IN")} in receivables looks at risk`,
              hint: `Based on each customer's own usual payment pace, not just a flat overdue date - ₹${Math.round(rf.likelySoon).toLocaleString("en-IN")} still looks likely within 7 days.`,
              ctaLabel: "View Health",
              onCta: () => setTab("health"),
            });
          }
          missingMoney.forEach((f) => {
            items.push({ severity: f.severity, source: "Missing Money", title: f.label, hint: f.hint, ctaLabel: "View Closing", onCta: () => setTab("closing") });
          });

          items.sort((a, b) => severityRank(b.severity) - severityRank(a.severity));
          const topItems = items.slice(0, 12);
          const sevDot = (s) => (s === "high" ? "🔴" : s === "medium" ? "🟡" : "⚪️");
          const review = monthEndReview(monthPnl, prevMonthPnl, items);
          const calDot = (t) => (t === "gst" ? "🏛️" : t === "bill" ? "📄" : "💵");

          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>{t("Action Center")}</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                {t("Your business pulse in one place - Health, GST Guardian, Cash-Flow Guardian, Trust Score, Payment Behaviour, Fix My Books, Recurring Transactions, Price Watch, Leak Detector, Receivable Probability, and the Missing Money Detector, sorted by what's worth doing first.")}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 18px", marginBottom: 18 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: sinceYesterday.hasActivity ? 8 : 0, flexWrap: "wrap", gap: 8 }}>
                  <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text }}>{t("What changed?")}</div>
                  <div style={{ display: "flex", gap: 4 }}>
                    {[[1, "Since yesterday"], [7, "This week"], [30, "This month"]].map(([d, label]) => (
                      <button
                        key={d}
                        onClick={() => setChangedWindowDays(d)}
                        style={{ background: changedWindowDays === d ? C.paper : "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 14, padding: "3px 10px", fontSize: 11, fontWeight: 500, cursor: "pointer" }}
                      >
                        {t(label)}
                      </button>
                    ))}
                  </div>
                </div>
                {!sinceYesterday.hasActivity ? (
                  <div style={{ fontSize: 12, color: C.textMuted }}>{t("Nothing new logged in this period.")}</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    {sinceYesterday.newBills.length > 0 && (
                      <div style={{ fontSize: 12, color: C.text }}>
                        📄 {sinceYesterday.newBills.length} new bill{sinceYesterday.newBills.length === 1 ? "" : "s"} scanned, ₹{Math.round(sinceYesterday.newBillsTotal).toLocaleString("en-IN")} total
                      </div>
                    )}
                    {sinceYesterday.newInvoices.length > 0 && (
                      <div style={{ fontSize: 12, color: C.text }}>
                        🧾 {sinceYesterday.newInvoices.length} new invoice{sinceYesterday.newInvoices.length === 1 ? "" : "s"} raised, ₹{Math.round(sinceYesterday.newInvoicesTotal).toLocaleString("en-IN")} total
                      </div>
                    )}
                    {sinceYesterday.collections.length > 0 && (
                      <div style={{ fontSize: 12, color: C.green }}>
                        💰 ₹{Math.round(sinceYesterday.collectionsTotal).toLocaleString("en-IN")} collected across {sinceYesterday.collections.length} invoice{sinceYesterday.collections.length === 1 ? "" : "s"}
                      </div>
                    )}
                    {sinceYesterday.recentClosing && (
                      <div style={{ fontSize: 12, color: Math.abs(Number(sinceYesterday.recentClosing.difference)) > 0.5 ? C.amber : C.green }}>
                        🧮 Evening closing logged for {sinceYesterday.recentClosing.closing_date} - {Number(sinceYesterday.recentClosing.difference) === 0 ? "matched exactly" : `${Number(sinceYesterday.recentClosing.difference) > 0 ? "cash over" : "cash short"} by ₹${Math.abs(Number(sinceYesterday.recentClosing.difference)).toLocaleString("en-IN")}`}
                      </div>
                    )}
                  </div>
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 18px", marginBottom: 18 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text, marginBottom: calendar.length ? 8 : 0 }}>{t("Coming up (next 30 days)")}</div>
                {calendar.length === 0 ? (
                  <div style={{ fontSize: 12, color: C.textMuted }}>{t("Nothing dated on the calendar right now.")}</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    {calendar.slice(0, 8).map((it, i) => (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 10, fontSize: 12 }}>
                        <span style={{ color: C.text }}>{calDot(it.type)} {it.label}</span>
                        <span style={{ fontFamily: mono, color: it.daysLeft < 0 ? C.red : it.daysLeft <= 3 ? C.amber : C.textMuted, flexShrink: 0, whiteSpace: "nowrap" }}>
                          {it.daysLeft < 0 ? `${Math.abs(it.daysLeft)}d overdue` : it.daysLeft === 0 ? "today" : `in ${it.daysLeft}d`}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginBottom: 18, display: "flex", alignItems: "center", gap: 18, flexWrap: "wrap" }}>
                <div style={{ width: 64, height: 64, borderRadius: "50%", border: `5px solid ${overallColor}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <span style={{ fontFamily: display, fontSize: 19, fontWeight: 700, color: overallColor }}>{overall}</span>
                </div>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 600, color: C.text }}>{overallLabel}</div>
                  <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3, lineHeight: 1.5 }}>
                    {items.length === 0 ? "Nothing urgent right now." : `${items.length} item${items.length === 1 ? "" : "s"} worth a look, most pressing first.`}
                  </div>
                  <button onClick={() => setTab("health")} style={{ marginTop: 6, background: "transparent", border: "none", color: C.ink, textDecoration: "underline", fontSize: 11.5, cursor: "pointer", padding: 0, fontFamily: sans }}>
                    {t("Full health breakdown")}
                  </button>
                </div>
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 18px", marginBottom: 18 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text, marginBottom: 8 }}>{t("Month-end review")}</div>
                {review.improved.length === 0 && review.attention.length === 0 ? (
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: review.topActions.length ? 8 : 0 }}>{t("Not enough of a change vs last month to call out yet.")}</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 4, marginBottom: review.topActions.length ? 10 : 0 }}>
                    {review.improved.map((s, i) => (
                      <div key={`i${i}`} style={{ fontSize: 12, color: C.green }}>↑ {s}</div>
                    ))}
                    {review.attention.map((s, i) => (
                      <div key={`a${i}`} style={{ fontSize: 12, color: C.amber }}>↓ {s}</div>
                    ))}
                  </div>
                )}
                {review.topActions.length > 0 && (
                  <div style={{ paddingTop: review.improved.length || review.attention.length ? 8 : 0, borderTop: review.improved.length || review.attention.length ? `1px solid ${C.paperLine}` : "none" }}>
                    <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 3 }}>Top {review.topActions.length} for next month:</div>
                    {review.topActions.map((t, i) => (
                      <div key={i} style={{ fontSize: 12, color: C.text }}>{i + 1}. {t}</div>
                    ))}
                  </div>
                )}
                <button onClick={() => setTab("pnl")} style={{ marginTop: 8, background: "transparent", border: "none", color: C.ink, textDecoration: "underline", fontSize: 11, cursor: "pointer", padding: 0, fontFamily: sans }}>
                  {t("View full P&L statement")}
                </button>
              </div>

              {topItems.length === 0 ? (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "20px 18px", textAlign: "center" }}>
                  <div style={{ fontSize: 13.5, color: C.green }}>✓ {t("All clear - no priority actions right now.")}</div>
                </div>
              ) : (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "6px 18px" }}>
                  {topItems.map((it, i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, padding: "12px 0", borderBottom: i < topItems.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div>
                        <div style={{ fontSize: 10, fontWeight: 600, color: C.gold, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 2 }}>{it.source}</div>
                        <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{sevDot(it.severity)} {it.title}</div>
                        <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2, lineHeight: 1.5 }}>{it.hint}</div>
                      </div>
                      {it.onCta && (
                        <button
                          onClick={it.onCta}
                          style={{
                            background: it.isQuickAdd ? C.gold : "transparent",
                            color: it.isQuickAdd ? C.white : C.ink,
                            border: it.isQuickAdd ? "none" : `1px solid ${C.cardBorder}`,
                            borderRadius: 16,
                            padding: "5px 12px",
                            fontSize: 11.5,
                            fontWeight: it.isQuickAdd ? 500 : 400,
                            cursor: "pointer",
                            flexShrink: 0,
                            whiteSpace: "nowrap",
                          }}
                        >
                          {it.ctaLabel}
                        </button>
                      )}
                      {!it.onCta && it.ctaLabel && <span style={{ fontSize: 11.5, color: C.green, flexShrink: 0, whiteSpace: "nowrap" }}>{it.ctaLabel}</span>}
                    </div>
                  ))}
                  {items.length > topItems.length && (
                    <div style={{ padding: "10px 0", fontSize: 11.5, color: C.textMuted }}>+{items.length - topItems.length} more in Audit and Health.</div>
                  )}
                </div>
              )}

              {auditLog.length > 0 && (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 18px", marginTop: 18 }}>
                  <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text, marginBottom: 8 }}>{t("Recent changes")}</div>
                  {auditLog.slice(0, 5).map((u) => (
                    <div key={u.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${C.paperLine}`, gap: 10 }}>
                      <div style={{ fontSize: 12, color: C.text }}>
                        {u.summary}
                        <span style={{ color: C.textMuted }}> · {new Date(u.created_at).toLocaleString("en-IN", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}</span>
                      </div>
                      {u.before && (
                        <button onClick={() => restoreAuditEntry(u)} disabled={restoringLogId === u.id} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 16, padding: "4px 10px", fontSize: 11, cursor: restoringLogId === u.id ? "default" : "pointer", flexShrink: 0 }}>
                          {restoringLogId === u.id ? t("Undoing…") : t("Undo")}
                        </button>
                      )}
                    </div>
                  ))}
                  <button onClick={() => setTab("audit")} style={{ marginTop: 8, background: "transparent", border: "none", color: C.ink, textDecoration: "underline", fontSize: 11, cursor: "pointer", padding: 0, fontFamily: sans }}>
                    {t("View full audit trail")}
                  </button>
                </div>
              )}
            </div>
          );
        })()}

        {tab === "scan" && step === "upload" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>{t("Snap or upload a bill")}</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18 }}>{t("Photo of any invoice, receipt, or bill")}</div>
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

            <div style={{ marginTop: 18, textAlign: "center" }}>
              {!manualEntryOpen ? (
                <button onClick={() => setManualEntryOpen(true)} style={{ background: "transparent", border: "none", color: C.ink, textDecoration: "underline", cursor: "pointer", fontSize: 13, fontFamily: sans, padding: 0 }}>
                  + {t("No photo? Add a manual entry")}
                </button>
              ) : (
                <div style={{ textAlign: "left", background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "14px 16px" }}>
                  <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text, marginBottom: 8 }}>{t("Describe the transaction in your own words")}</div>
                  <textarea
                    value={manualEntryText}
                    onChange={(e) => setManualEntryText(e.target.value)}
                    placeholder={`e.g. "Paid ₹18,500 rent to Sharma Properties"`}
                    rows={2}
                    style={{ display: "block", width: "100%", boxSizing: "border-box", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none", fontFamily: sans, resize: "vertical" }}
                  />
                  {manualEntryError && <div style={{ fontSize: 12, color: C.red, marginTop: 8 }}>{manualEntryError}</div>}
                  <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                    <button onClick={submitManualEntry} disabled={manualEntryLoading || !manualEntryText.trim()} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: manualEntryLoading ? "default" : "pointer", opacity: manualEntryLoading || !manualEntryText.trim() ? 0.7 : 1 }}>
                      {manualEntryLoading ? t("Reading…") : t("Add entry")}
                    </button>
                    <button onClick={() => { setManualEntryOpen(false); setManualEntryText(""); setManualEntryError(null); }} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 14px", fontSize: 13, cursor: "pointer" }}>
                      {t("Cancel")}
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "scan" && step === "processing" && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ width: 34, height: 34, border: `3px solid ${C.paperLine}`, borderTopColor: C.gold, borderRadius: "50%", margin: "0 auto 16px", animation: "spin 0.9s linear infinite" }} />
            <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
            <div style={{ fontSize: 14, color: C.textMuted }}>{t("Reading your bill…")}</div>
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
                      {t(label)}
                    </div>
                    <input value={draft[field] ?? ""} onChange={(e) => updateDraft(field, e.target.value)} style={{ flex: 1, border: "none", background: "transparent", fontSize: 14, color: C.text, outline: "none" }} />
                  </div>
                ))}
                <div style={{ display: "flex", alignItems: "center", padding: "10px 0", gap: 12 }}>
                  <div style={{ width: 90, flexShrink: 0, fontSize: 12.5, color: C.ink, fontWeight: 500 }}>{t("Total")}</div>
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
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12, fontSize: 12.5, color: C.text, cursor: "pointer" }}>
              <input type="checkbox" checked={!!draft.is_personal} onChange={(e) => updateDraft("is_personal", e.target.checked)} style={{ width: 15, height: 15, cursor: "pointer" }} />
              🏠 This is a personal expense, not business
            </label>
            <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
              <button onClick={saveDraft} style={{ flex: 1, background: dupConfirmed ? C.red : C.green, color: C.white, border: "none", borderRadius: 6, padding: "11px 0", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>
                {dupConfirmed ? t("Yes, save anyway") : t("Save to ledger")}
              </button>
              <button onClick={discardDraft} style={{ background: "transparent", color: C.red, border: `1px solid ${C.red}`, borderRadius: 6, padding: "11px 18px", fontSize: 14, cursor: "pointer" }}>
                {t("Discard")}
              </button>
            </div>
          </div>
        )}

        {tab === "ledger" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: personalTotal > 0 ? 4 : 14 }}>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600 }}>{t("This month")}</div>
              <div style={{ fontFamily: mono, fontSize: 18, color: C.ink }}>₹{monthTotal.toLocaleString("en-IN")}</div>
            </div>
            {personalTotal > 0 && (
              <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 14 }}>
                Of which 🏠 personal: ₹{personalTotal.toLocaleString("en-IN")} — kept separate from your business books.
              </div>
            )}

            {entriesLoaded && entries.length === 0 && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>{t("No bills saved yet. Scan your first one.")}</div>}
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
                              {t("Split into items")}
                            </button>
                          )}
                          <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, fontSize: 12.5, color: C.text, cursor: "pointer" }}>
                            <input type="checkbox" checked={!!editDraft.is_personal} onChange={(ev) => updateEditField("is_personal", ev.target.checked)} style={{ width: 15, height: 15, cursor: "pointer" }} />
                            🏠 {t("Personal expense, not business")}
                          </label>
                          {editError && <div style={{ fontSize: 12, color: C.red, marginBottom: 8 }}>{editError}</div>}
                          <div style={{ display: "flex", gap: 8 }}>
                            <button onClick={() => saveEditEntry(e.id)} disabled={savingEdit} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 12.5, fontWeight: 500, cursor: savingEdit ? "default" : "pointer", opacity: savingEdit ? 0.7 : 1 }}>
                              {savingEdit ? t("Saving…") : t("Save")}
                            </button>
                            <button onClick={cancelEditEntry} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
                              {t("Cancel")}
                            </button>
                            <button onClick={() => deleteEntry(e.id)} style={{ marginLeft: "auto", background: deleteConfirmId === e.id ? C.red : "transparent", color: deleteConfirmId === e.id ? C.white : C.red, border: `1px solid ${C.red}`, borderRadius: 6, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
                              {deleteConfirmId === e.id ? t("Confirm delete?") : t("Delete")}
                            </button>
                          </div>
                        </div>
                      );
                    }
                    return (
                      <div key={e.id || i} onClick={viewMode === "owner" && role === "owner" ? () => startEditEntry(e) : undefined} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5, cursor: viewMode === "owner" && role === "owner" ? "pointer" : "default" }}>
                        <div style={{ color: C.textMuted, width: 78, flexShrink: 0 }}>{e.date || "—"}</div>
                        <div style={{ flex: 1 }}>
                          <div style={{ color: C.text }}>
                            {e.vendor || "—"}
                            {e.is_personal && (
                              <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600, color: C.amber, background: "#FFF1DB", borderRadius: 10, padding: "1px 7px" }}>🏠 Personal</span>
                            )}
                            {addedByLabel(e.created_by) && (
                              <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600, color: C.inkLight, background: "#EAF0FF", borderRadius: 10, padding: "1px 7px" }}>👤 {addedByLabel(e.created_by)}</span>
                            )}
                          </div>
                          <div style={{ fontSize: 11, color: C.textMuted, display: "flex", alignItems: "center", gap: 5 }}>
                            <span>
                              {CATEGORY_LABELS[e.category] || e.category}
                              {Array.isArray(e.items) && e.items.length > 1 ? ` · ${e.items.length} items` : ""}
                            </span>
                            <button
                              onClick={(ev) => { ev.stopPropagation(); setExplainEntryId(explainEntryId === e.id ? null : e.id); }}
                              title="Why this category?"
                              style={{ background: "transparent", border: "none", color: C.textMuted, cursor: "pointer", padding: 0, fontSize: 11, lineHeight: 1, opacity: 0.7 }}
                            >
                              ⓘ
                            </button>
                          </div>
                          {explainEntryId === e.id && (
                            <div onClick={(ev) => ev.stopPropagation()} style={{ fontSize: 11, color: C.textMuted, fontStyle: "italic", marginTop: 4, background: C.paper, borderRadius: 4, padding: "6px 9px", lineHeight: 1.5, cursor: "default" }}>
                              {categoryExplanation(e, entries)}
                            </div>
                          )}
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
                {t("Export to Excel")}
              </button>
            )}
          </div>
        )}

        {tab === "gst" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>{t("GST filing reminders")}</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
              {t("Standard monthly filing dates - if you're on the quarterly QRMP scheme, your actual dates differ. Confirm with your CA.")}
            </div>
            {(() => {
              const reminders = getGstReminders();
              const returns = [
                { key: "gstr1", label: "GSTR-1", sub: "Outward supplies (sales)", info: reminders.gstr1 },
                { key: "gstr3b", label: "GSTR-3B", sub: "Summary return + tax payment", info: reminders.gstr3b },
              ];
              return (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
                  {returns.map((r) => {
                    const urgent = r.info.daysLeft <= 3;
                    return (
                      <div key={r.key} style={{ background: C.white, border: `1px solid ${urgent ? C.red : C.cardBorder}`, borderRadius: 8, padding: "16px 16px" }}>
                        <div style={{ fontSize: 14, fontWeight: 600, color: C.text }}>{r.label}</div>
                        <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10 }}>{r.sub}</div>
                        <div style={{ fontFamily: display, fontSize: 22, fontWeight: 700, color: urgent ? C.red : C.ink }}>
                          {r.info.daysLeft === 0 ? "Due today" : r.info.daysLeft === 1 ? "1 day left" : `${r.info.daysLeft} days left`}
                        </div>
                        <div style={{ fontSize: 12, color: C.textMuted, marginTop: 4 }}>
                          Due {r.info.due.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })} · for {r.info.period}
                        </div>
                      </div>
                    );
                  })}
                </div>
              );
            })()}
            {(() => {
              const reminders = getGstReminders();
              const gst = purchaseGstForPeriod(entries, reminders.gstr3b.periodKey);
              return (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px" }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>{t("Your recorded purchase GST")} - {reminders.gstr3b.period}</div>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                    From {gst.count} bill{gst.count === 1 ? "" : "s"} scanned this period. This is GST on what you bought - useful for input tax credit and your CA. See below for your sales/output GST.
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
                    {[["CGST", gst.cgst], ["SGST", gst.sgst], ["IGST", gst.igst], ["Purchases", gst.total]].map(([label, val]) => (
                      <div key={label} style={{ textAlign: "center" }}>
                        <div style={{ fontFamily: mono, fontSize: 15, fontWeight: 600, color: C.ink }}>₹{val.toLocaleString("en-IN")}</div>
                        <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>{label}</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

            {(() => {
              const reminders = getGstReminders();
              const purchaseGst = purchaseGstForPeriod(entries, reminders.gstr3b.periodKey);
              const salesGst = salesGstForPeriod(invoices, reminders.gstr3b.periodKey);
              const netReserve = Math.max(salesGst.total - purchaseGst.total, 0);
              const balance = businessProfile && businessProfile.cash_balance != null ? Number(businessProfile.cash_balance) : null;
              const available = balance != null ? balance - netReserve : null;
              return (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginTop: 16 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>{t("GST reserve")} - {reminders.gstr3b.period}</div>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                    Sales GST (₹{salesGst.total.toLocaleString("en-IN")}) minus input tax credit from purchases (₹{purchaseGst.total.toLocaleString("en-IN")}) - a rough estimate of what to set aside before filing, not a substitute for your CA's calculation.
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderTop: `1px solid ${C.paperLine}` }}>
                    <span style={{ fontSize: 13, color: C.textMuted }}>{t("Estimated GST to reserve")}</span>
                    <span style={{ fontFamily: mono, fontSize: 16, fontWeight: 600, color: C.amber }}>₹{netReserve.toLocaleString("en-IN")}</span>
                  </div>

                  <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${C.paperLine}` }}>
                    <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 6 }}>{t("Your current cash/bank balance (entered manually - BahiSathi doesn't connect to your bank yet)")}</div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <input
                        type="number"
                        value={cashBalanceInput}
                        onChange={(e) => { setCashBalanceInput(e.target.value); setCashBalanceSaved(false); }}
                        placeholder="e.g. 842000"
                        style={{ flex: "1 1 160px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13.5, fontFamily: mono, background: C.white, color: C.text, outline: "none" }}
                      />
                      <button onClick={saveCashBalance} disabled={cashBalanceSaving} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: cashBalanceSaving ? "default" : "pointer", opacity: cashBalanceSaving ? 0.7 : 1 }}>
                        {cashBalanceSaving ? t("Saving…") : t("Save")}
                      </button>
                      {cashBalanceSaved && !cashBalanceSaving && <span style={{ fontSize: 12.5, color: C.green, alignSelf: "center" }}>{t("Saved")} ✓</span>}
                    </div>
                    {businessProfile && businessProfile.cash_balance_updated_at && (
                      <div style={{ fontSize: 11, color: C.textMuted, marginTop: 6 }}>
                        Last updated {new Date(businessProfile.cash_balance_updated_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })} - update it whenever you check your balance.
                      </div>
                    )}
                  </div>

                  {available != null && (
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 14px", marginTop: 12, background: C.paper, borderRadius: 6 }}>
                      <span style={{ fontSize: 13.5, fontWeight: 600, color: C.text }}>{t("Available for business")}</span>
                      <span style={{ fontFamily: mono, fontSize: 17, fontWeight: 700, color: available < 0 ? C.red : C.green }}>₹{available.toLocaleString("en-IN")}</span>
                    </div>
                  )}
                </div>
              );
            })()}
          </div>
        )}

        {tab === "health" && (() => {
          const cashFlow = cashFlowForecast(entries, invoices, businessProfile);
          const trustScores = customers
            .map((c) => trustScoreFor(paymentProfileFor(c.id)))
            .filter((s) => s != null);
          const collectionsScore = trustScores.length ? Math.round(trustScores.reduce((s, t) => s + t.score, 0) / trustScores.length) : 75;
          const cashScore = cashFlow.status === "healthy" ? 100 : cashFlow.status === "insufficient_data" ? 60 : cashFlow.status === "watch" ? 70 : cashFlow.status === "at_risk" ? 40 : 15;
          const gstScore = cashFlow.available == null ? 75 : cashFlow.available >= 0 ? 100 : cashFlow.available > -cashFlow.netReserve ? 40 : 10;
          const overall = Math.round(cashScore * 0.4 + collectionsScore * 0.35 + gstScore * 0.25);
          const overallLabel = overall >= 80 ? "Strong" : overall >= 60 ? "Steady" : overall >= 40 ? "Needs attention" : "At risk";
          const overallColor = overall >= 80 ? C.green : overall >= 60 ? C.green : overall >= 40 ? C.amber : C.red;

          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Business health</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                A rough, automatic read on cash, collections, and GST readiness - built from your own data, not a substitute for professional advice.
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginBottom: 16, display: "flex", alignItems: "center", gap: 18 }}>
                <div style={{ width: 72, height: 72, borderRadius: "50%", border: `5px solid ${overallColor}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <span style={{ fontFamily: display, fontSize: 22, fontWeight: 700, color: overallColor }}>{overall}</span>
                </div>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 600, color: C.text }}>{overallLabel}</div>
                  <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3, lineHeight: 1.5 }}>
                    Cash flow {cashScore} · Collections {collectionsScore} · GST readiness {gstScore}
                  </div>
                </div>
              </div>

              {(() => {
                const mm = moneyMapFor(businessProfile, invoices);
                return (
                  <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>{t("Your business money")} {mm.hasCash ? `— ₹${Math.round(mm.total).toLocaleString("en-IN")}` : ""}</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 10 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
                        <span style={{ color: C.textMuted }}>{t("Bank & Cash")}</span>
                        <span style={{ fontFamily: mono, color: C.text }}>{mm.hasCash ? `₹${Math.round(mm.cash).toLocaleString("en-IN")}` : "—"}</span>
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
                        <span style={{ color: C.textMuted }}>{t("Customers owe you")}</span>
                        <span style={{ fontFamily: mono, color: C.text }}>₹{Math.round(mm.outstanding).toLocaleString("en-IN")}</span>
                      </div>
                    </div>
                    <div style={{ fontSize: 11, color: C.textMuted, marginTop: 10, lineHeight: 1.5 }}>
                      {!mm.hasCash
                        ? t("Enter your cash balance on the GST tab to see your full money map.")
                        : t("Inventory value and other assets aren't tracked yet.")}
                    </div>
                  </div>
                );
              })()}

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>Cash-Flow Guardian</div>
                {cashFlow.status === "insufficient_data" && (
                  <div style={{ fontSize: 12.5, color: C.textMuted, lineHeight: 1.5 }}>
                    Not enough history yet to forecast - keep scanning bills, marking invoices paid, and entering your cash balance on the GST tab, and this will fill in over the next few months.
                  </div>
                )}
                {cashFlow.status === "healthy" && (
                  <div style={{ fontSize: 12.5, color: C.green, lineHeight: 1.5 }}>
                    Money coming in (₹{Math.round(cashFlow.avgIncome).toLocaleString("en-IN")}/mo average) is covering money going out (₹{Math.round(cashFlow.avgExpense).toLocaleString("en-IN")}/mo average) - cash isn't shrinking at the current pace.
                  </div>
                )}
                {(cashFlow.status === "watch" || cashFlow.status === "at_risk" || cashFlow.status === "immediate_risk") && (
                  <div style={{ fontSize: 12.5, color: cashFlow.status === "immediate_risk" ? C.red : C.amber, lineHeight: 1.5 }}>
                    Spending (₹{Math.round(cashFlow.avgExpense).toLocaleString("en-IN")}/mo average) is outpacing income (₹{Math.round(cashFlow.avgIncome).toLocaleString("en-IN")}/mo average).{" "}
                    {cashFlow.status === "immediate_risk"
                      ? "Available cash is already at or below your GST reserve - collections or new spending decisions need attention now."
                      : `At this pace, available cash runs out around ${cashFlow.riskDate.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })} (about ${cashFlow.monthsRunway.toFixed(1)} months).`}
                  </div>
                )}
              </div>

              {(() => {
                const now3 = new Date();
                const pKey3 = `${now3.getFullYear()}-${String(now3.getMonth() + 1).padStart(2, "0")}`;
                const monthPnl3 = pnlForPeriod(entries, invoices, pKey3);
                if (monthPnl3.revenue <= 0) return null;
                const gstReserveThisMonth = Math.max(salesGstForPeriod(invoices, pKey3).total - purchaseGstForPeriod(entries, pKey3).total, 0);
                const bridge = profitToCashBridge(monthPnl3, invoices, pKey3, gstReserveThisMonth);
                return (
                  <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>{t("Where did my profit go?")}</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
                        <span style={{ color: C.text, fontWeight: 500 }}>{t("Profit this month")}</span>
                        <span style={{ fontFamily: mono, color: C.ink, fontWeight: 600 }}>₹{Math.round(bridge.profit).toLocaleString("en-IN")}</span>
                      </div>
                      {bridge.uncollected > 0 && (
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, paddingLeft: 10 }}>
                          <span style={{ color: C.textMuted }}>{t("Sitting with customers (not yet paid)")}</span>
                          <span style={{ fontFamily: mono, color: C.amber }}>-₹{Math.round(bridge.uncollected).toLocaleString("en-IN")}</span>
                        </div>
                      )}
                      {bridge.gstReserve > 0 && (
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, paddingLeft: 10 }}>
                          <span style={{ color: C.textMuted }}>{t("Set aside for GST")}</span>
                          <span style={{ fontFamily: mono, color: C.amber }}>-₹{Math.round(bridge.gstReserve).toLocaleString("en-IN")}</span>
                        </div>
                      )}
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, paddingTop: 6, borderTop: `1px solid ${C.paperLine}` }}>
                        <span style={{ color: C.text, fontWeight: 600 }}>{t("Actually in hand from this month")}</span>
                        <span style={{ fontFamily: mono, color: bridge.actualCashFromThisMonth < 0 ? C.red : C.green, fontWeight: 700 }}>₹{Math.round(bridge.actualCashFromThisMonth).toLocaleString("en-IN")}</span>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {(() => {
                const rf = receivableForecast();
                if (rf.total <= 0) return null;
                return (
                  <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>Receivable Probability</div>
                    <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 12, lineHeight: 1.5 }}>
                      What's outstanding, split by how likely it is to come in soon - based on each customer's own payment history, not one flat overdue cutoff.
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                      <div style={{ textAlign: "center" }}>
                        <div style={{ fontFamily: mono, fontSize: 15, fontWeight: 700, color: C.green }}>₹{Math.round(rf.likelySoon).toLocaleString("en-IN")}</div>
                        <div style={{ fontSize: 10.5, color: C.textMuted, marginTop: 2 }}>Likely within 7 days</div>
                      </div>
                      <div style={{ textAlign: "center" }}>
                        <div style={{ fontFamily: mono, fontSize: 15, fontWeight: 700, color: C.ink }}>₹{Math.round(rf.dueLater).toLocaleString("en-IN")}</div>
                        <div style={{ fontSize: 10.5, color: C.textMuted, marginTop: 2 }}>Due later</div>
                      </div>
                      <div style={{ textAlign: "center" }}>
                        <div style={{ fontFamily: mono, fontSize: 15, fontWeight: 700, color: rf.atRisk > 0 ? C.red : C.ink }}>₹{Math.round(rf.atRisk).toLocaleString("en-IN")}</div>
                        <div style={{ fontSize: 10.5, color: C.textMuted, marginTop: 2 }}>At risk</div>
                      </div>
                    </div>
                    {rf.atRisk > 0 && (
                      <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${C.paperLine}` }}>
                        {rf.rows
                          .filter((r) => r.bucket === "at_risk")
                          .sort((a, b) => b.total - a.total)
                          .slice(0, 5)
                          .map((r) => (
                            <div key={r.invoice.id} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 12 }}>
                              <span style={{ color: C.text }}>{r.customerName} · {r.invoice.invoice_number}</span>
                              <span style={{ fontFamily: mono, color: C.red }}>₹{r.total.toLocaleString("en-IN")} ({r.ageDays}d old)</span>
                            </div>
                          ))}
                      </div>
                    )}
                  </div>
                );
              })()}

              <button onClick={() => setTab("pnl")} style={{ display: "block", width: "100%", textAlign: "left", background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", cursor: "pointer", fontFamily: sans }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text }}>Profit &amp; Loss statement →</div>
                  <span style={{ fontSize: 11.5, color: C.ink, textDecoration: "underline" }}>Open P&amp;L tab</span>
                </div>
                <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 4, lineHeight: 1.5 }}>
                  This month's revenue/expenses, a plain-language explanation, a 6-month trend, and an Excel export - now its own tab.
                </div>
              </button>
            </div>
          );
        })()}

        {tab === "pnl" && (() => {
          const now = new Date();
          const pKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
          const prevDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
          const prevKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, "0")}`;
          const pnl = pnlForPeriod(entries, invoices, pKey);
          const prevPnl = pnlForPeriod(entries, invoices, prevKey);
          const periodLabel = now.toLocaleDateString("en-IN", { month: "long", year: "numeric" });
          const trend = pnlTrend(entries, invoices, 6);
          const trendMax = Math.max(1, ...trend.map((t) => Math.max(t.revenue, t.expenses)));

          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Profit &amp; Loss</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                Revenue from invoices, minus business expenses from scanned bills (personal-tagged ones excluded) - a genuine estimate for the owner, not a formal statement your CA would prepare.
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>{periodLabel}</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 12 }}>
                  {[["Revenue", pnl.revenue, C.ink], ["Expenses", pnl.expenses, C.ink], ["Net", pnl.net, pnl.net >= 0 ? C.green : C.red]].map(([label, val, color]) => (
                    <div key={label} style={{ textAlign: "center" }}>
                      <div style={{ fontFamily: mono, fontSize: 16, fontWeight: 700, color }}>{pnl.net < 0 && label === "Net" ? "-" : ""}₹{Math.abs(val).toLocaleString("en-IN")}</div>
                      <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>{label}</div>
                    </div>
                  ))}
                </div>
                {pnl.topCategories.length > 0 && (
                  <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 12 }}>
                    Top expenses: {pnl.topCategories.map((c) => `${c.category} ₹${c.total.toLocaleString("en-IN")}`).join(" · ")}
                  </div>
                )}
                {!pnlNarration ? (
                  <button onClick={() => explainPnl(pnl, periodLabel, prevPnl)} disabled={pnlNarrationLoading} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: pnlNarrationLoading ? "default" : "pointer", opacity: pnlNarrationLoading ? 0.7 : 1 }}>
                    {pnlNarrationLoading ? "Thinking…" : "Explain my P&L"}
                  </button>
                ) : (
                  <div style={{ background: "#E7FFDB", border: "1px solid #BCE5A8", borderRadius: "10px 10px 10px 2px", padding: "10px 14px", fontSize: 13, color: "#1F2C1A", lineHeight: 1.55 }}>
                    {pnlNarration}
                  </div>
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text }}>Last 6 months</div>
                  <button onClick={() => exportPnlTrend(trend)} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "7px 14px", fontSize: 12.5, fontWeight: 500, cursor: "pointer" }}>
                    Export to Excel
                  </button>
                </div>
                {trend.map((t) => (
                  <div key={t.key} style={{ marginBottom: 8 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: C.textMuted, marginBottom: 3 }}>
                      <span>{t.label}</span>
                      <span style={{ fontFamily: mono, color: t.net >= 0 ? C.green : C.red }}>
                        {t.net < 0 ? "-" : ""}₹{Math.abs(Math.round(t.net)).toLocaleString("en-IN")}
                      </span>
                    </div>
                    <div style={{ height: 5, background: C.paper, borderRadius: 2, overflow: "hidden" }}>
                      <div style={{ width: `${Math.min(100, (t.revenue / trendMax) * 100)}%`, height: "100%", background: C.ink }} />
                    </div>
                    <div style={{ height: 5, background: C.paper, borderRadius: 2, overflow: "hidden", marginTop: 2 }}>
                      <div style={{ width: `${Math.min(100, (t.expenses / trendMax) * 100)}%`, height: "100%", background: C.amber }} />
                    </div>
                  </div>
                ))}
                <div style={{ display: "flex", gap: 14, marginTop: 6, fontSize: 10.5, color: C.textMuted }}>
                  <span><span style={{ display: "inline-block", width: 8, height: 8, background: C.ink, borderRadius: 2, marginRight: 4, verticalAlign: "middle" }} />Revenue</span>
                  <span><span style={{ display: "inline-block", width: 8, height: 8, background: C.amber, borderRadius: 2, marginRight: 4, verticalAlign: "middle" }} />Expenses</span>
                </div>
              </div>
            </div>
          );
        })()}

        {tab === "plan" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 18 }}>{t("Your plan")}</div>

            {sub === null && <div style={{ fontSize: 13.5, color: C.textMuted }}>{t("Loading…")}</div>}

            {sub && sub.plan === "pro" && (
              <div style={{ background: "#FBF4DF", border: `1px solid ${C.gold}`, borderRadius: 8, padding: "20px 22px" }}>
                <div style={{ fontFamily: display, fontSize: 17, fontWeight: 600, color: C.ink, marginBottom: 6 }}>{t("You're on Pro")} ✨</div>
                <div style={{ fontSize: 13.5, color: C.textMuted, lineHeight: 1.6 }}>
                  {t("Unlimited scans, every month.")}
                  {sub.current_period_end && (
                    <> {t("Renews on")} {new Date(sub.current_period_end).toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}.</>
                  )}
                </div>
              </div>
            )}

            {sub && sub.plan === "free" && (
              <div>
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginBottom: 18 }}>
                  <div style={{ fontSize: 13.5, color: C.text, fontWeight: 500, marginBottom: 4 }}>{t("Free plan")}</div>
                  <div style={{ fontSize: 13, color: C.textMuted }}>{sub.scans_used} / {sub.scans_limit} {t("scans used this month")}</div>
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
                          {t(label)}
                        </button>
                      ))}
                    </div>
                  </div>
                  <div style={{ fontSize: 13.5, color: "#C7CFDF", lineHeight: 1.7, margin: "10px 0 18px" }}>
                    {t("Unlimited bill scans on web and WhatsApp. No caps, no waiting for the 1st of the month. Price shown is inclusive of GST.")}
                  </div>
                  <button onClick={() => startUpgrade(upgradeCycle)} disabled={upgrading} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "11px 22px", fontSize: 14, fontWeight: 600, cursor: upgrading ? "default" : "pointer", opacity: upgrading ? 0.7 : 1 }}>
                    {upgrading ? t("Opening checkout…") : `${t("Upgrade to Pro")} — ${PRO_PRICING[upgradeCycle].amount}${PRO_PRICING[upgradeCycle].per}`}
                  </button>
                  {upgradeError && <div style={{ marginTop: 12, fontSize: 12.5, color: "#F3B8B8" }}>{upgradeError}</div>}
                  <div style={{ fontSize: 11.5, color: "#9FADC7", marginTop: 12 }}>
                    {t("Non-refundable, does not auto-renew. Full terms on the")} <button onClick={() => setPage("plans")} style={{ background: "transparent", border: "none", padding: 0, color: "#C7CFDF", textDecoration: "underline", cursor: "pointer", fontSize: 11.5, fontFamily: sans }}>{t("Plans page")}</button>.
                  </div>
                </div>
              </div>
            )}

            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginTop: 18 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>{t("Link your WhatsApp number")}</div>
              <div style={{ fontSize: 12.5, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                {t("Scan bills on WhatsApp and the web with one shared free-scan limit and one Pro subscription. Bills you've already sent on WhatsApp will move into this ledger too.")}
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

            {linkedNumbers.length > 0 && businessLoaded && businessProfile && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginTop: 18 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                  <div>
                    <div style={{ fontSize: 14.5, fontWeight: 600, color: C.text, marginBottom: 4 }}>AI Morning Brief</div>
                    <div style={{ fontSize: 12.5, color: C.textMuted, lineHeight: 1.5 }}>
                      A daily WhatsApp message with yesterday's sales, overdue collections, and your next GST due date - sent every morning to your linked number(s).
                    </div>
                  </div>
                  <button
                    onClick={toggleMorningBrief}
                    disabled={morningBriefSaving}
                    style={{
                      flexShrink: 0,
                      background: businessProfile.morning_brief_enabled ? C.green : C.cardBorder,
                      border: "none",
                      borderRadius: 14,
                      width: 46,
                      height: 26,
                      position: "relative",
                      cursor: morningBriefSaving ? "default" : "pointer",
                      opacity: morningBriefSaving ? 0.7 : 1,
                    }}
                    aria-label="Toggle AI Morning Brief"
                  >
                    <span style={{ position: "absolute", top: 3, left: businessProfile.morning_brief_enabled ? 23 : 3, width: 20, height: 20, borderRadius: "50%", background: C.white, transition: "left 0.15s" }} />
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {tab === "business" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>{t("Business details")}</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
              {t("Optional - not needed to scan bills or use your ledger. Fill this in if you want your business name and GSTIN to appear on your Excel exports and, later, on GST-ready reports.")}
            </div>

            {businessLoaded && goalsLoaded && businessProfile && (() => {
              const now = new Date();
              const pKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
              const monthPnl = pnlForPeriod(entries, invoices, pKey);
              const totalOutstanding = invoices.filter((inv) => inv.status !== "paid").reduce((s, inv) => s + (Number(inv.total) || 0), 0);
              const ctx = {
                cashBalance: businessProfile.cash_balance,
                totalOutstanding,
                monthRevenue: monthPnl.revenue,
                monthExpenseRatio: monthPnl.revenue > 0 ? (monthPnl.expenses / monthPnl.revenue) * 100 : null,
              };
              return (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 18 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>{t("AI Business Goals")}</div>
                  <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 12, lineHeight: 1.5 }}>
                    {t("Set a target, BahiSathi tracks it against numbers it already computes elsewhere.")}
                  </div>

                  {businessGoals.length > 0 && (
                    <div style={{ marginBottom: 14 }}>
                      {businessGoals.map((g) => {
                        const gp = goalProgress(g, ctx);
                        const barColor = gp.status === "met" ? C.green : gp.status === "behind" ? C.amber : C.textMuted;
                        return (
                          <div key={g.id} style={{ padding: "8px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                              <span style={{ fontSize: 12.5, color: C.text, fontWeight: 500 }}>
                                {gp.status === "met" ? "✓" : gp.status === "behind" ? "○" : "—"} {GOAL_TYPES[g.type].label}
                              </span>
                              <button onClick={() => deleteGoal(g.id)} style={{ background: "transparent", border: "none", color: C.red, fontSize: 11, cursor: "pointer", flexShrink: 0 }}>{t("Remove")}</button>
                            </div>
                            <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{gp.label}</div>
                            {gp.pct != null && (
                              <div style={{ height: 5, background: C.paper, borderRadius: 2, overflow: "hidden", marginTop: 6 }}>
                                <div style={{ width: `${gp.pct}%`, height: "100%", background: barColor }} />
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                    <select value={newGoalType} onChange={(e) => setNewGoalType(e.target.value)} style={{ border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 12.5, background: C.white, color: C.text }}>
                      {Object.entries(GOAL_TYPES).map(([key, t]) => (
                        <option key={key} value={key}>{t.label}</option>
                      ))}
                    </select>
                    <input
                      type="number"
                      value={newGoalValue}
                      onChange={(e) => setNewGoalValue(e.target.value)}
                      placeholder={GOAL_TYPES[newGoalType].unit === "%" ? "e.g. 18" : "e.g. 500000"}
                      style={{ width: 130, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 12.5, background: C.white, color: C.text, outline: "none" }}
                    />
                    <button onClick={addGoal} disabled={savingGoal || !newGoalValue} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 12.5, fontWeight: 500, cursor: savingGoal ? "default" : "pointer", opacity: savingGoal || !newGoalValue ? 0.7 : 1 }}>
                      {savingGoal ? "Saving…" : "+ Add goal"}
                    </button>
                  </div>
                  {goalsError && <div style={{ marginTop: 10, fontSize: 11.5, color: C.red }}>{goalsError}</div>}
                </div>
              );
            })()}

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
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5, gap: 16 }}>
                    <span style={{ color: C.textMuted, flexShrink: 0 }}>Address</span>
                    <span style={{ color: C.text, textAlign: "right" }}>{businessProfile.address || "Not set"}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
                    <span style={{ color: C.textMuted }}>Website</span>
                    <span style={{ color: C.text }}>{businessProfile.website || "Not set"}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", fontSize: 13.5 }}>
                    <span style={{ color: C.textMuted }}>Email</span>
                    <span style={{ color: C.text }}>{businessProfile.email || "Not set"}</span>
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
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Address</div>
                  <textarea
                    value={businessProfile.address}
                    onChange={(e) => updateBusinessField("address", e.target.value)}
                    placeholder="Shop / office address"
                    rows={3}
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none", fontFamily: sans, resize: "vertical" }}
                  />
                </div>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Website</div>
                  <input
                    value={businessProfile.website}
                    onChange={(e) => updateBusinessField("website", e.target.value)}
                    placeholder="e.g. www.sharmatraders.in"
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
                  />
                </div>
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Email address</div>
                  <input
                    value={businessProfile.email}
                    onChange={(e) => updateBusinessField("email", e.target.value)}
                    placeholder="e.g. shop@sharmatraders.in"
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, outline: "none" }}
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

            <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 14, lineHeight: 1.5, marginBottom: 24 }}>
              No company ID or GST verification is required to use BahiSathi - these details are only used to label your own exports.
            </div>

            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px" }}>
              <div style={{ fontFamily: display, fontSize: 16, fontWeight: 600, marginBottom: 4 }}>Share with your CA</div>
              <div style={{ fontSize: 12.5, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                Generate a read-only link your accountant can open without logging in - it shows your ledger and lets them download the Excel export.
              </div>
              {!shareLoaded && <div style={{ fontSize: 13, color: C.textMuted }}>Loading…</div>}
              {shareLoaded && !shareLink && (
                <button onClick={generateShareLink} disabled={shareBusy} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "9px 18px", fontSize: 13, fontWeight: 500, cursor: shareBusy ? "default" : "pointer", opacity: shareBusy ? 0.7 : 1 }}>
                  {shareBusy ? "Generating…" : "Generate share link"}
                </button>
              )}
              {shareLoaded && shareLink && (
                <div>
                  <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                    <input
                      readOnly
                      value={`${window.location.origin}${window.location.pathname}?share=${shareLink.token}`}
                      onFocus={(e) => e.target.select()}
                      style={{ flex: 1, display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 12, fontFamily: mono, background: C.paper, color: C.text }}
                    />
                    <button onClick={copyShareLink} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 14px", fontSize: 12.5, fontWeight: 500, cursor: "pointer", whiteSpace: "nowrap" }}>
                      {shareCopied ? "Copied ✓" : "Copy"}
                    </button>
                  </div>
                  <button onClick={revokeShareLink} disabled={shareBusy} style={{ background: "transparent", color: C.red, border: `1px solid ${C.red}`, borderRadius: 6, padding: "7px 14px", fontSize: 12, cursor: shareBusy ? "default" : "pointer" }}>
                    {shareBusy ? "Working…" : "Revoke link"}
                  </button>
                </div>
              )}
              {shareError && <div style={{ fontSize: 12, color: C.red, marginTop: 8 }}>{shareError}</div>}
            </div>

            <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px", marginTop: 16 }}>
              <div style={{ fontFamily: display, fontSize: 16, fontWeight: 600, marginBottom: 4 }}>Staff access</div>
              <div style={{ fontSize: 12.5, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                Invite someone by their WhatsApp number to help with scanning bills, invoices, and inventory. They sign up for their own BahiSathi login and link that same number - once linked, their access activates automatically. They can't see profit, Health, GST, audit, or these business settings, and can't edit or delete ledger entries.
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
                <input value={newStaffName} onChange={(e) => setNewStaffName(e.target.value)} placeholder="Name (optional)" style={{ flex: "1 1 140px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                <input value={newStaffPhone} onChange={(e) => setNewStaffPhone(e.target.value)} placeholder="WhatsApp number, e.g. +919812345678" style={{ flex: "2 1 200px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                <button onClick={inviteStaff} disabled={staffInviting || !newStaffPhone.trim()} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: staffInviting ? "default" : "pointer", opacity: staffInviting ? 0.7 : 1 }}>
                  {staffInviting ? "Inviting…" : "+ Invite"}
                </button>
              </div>
              {staffError && <div style={{ fontSize: 12.5, color: C.red, marginBottom: 10 }}>{staffError}</div>}
              {!staffListLoaded && <div style={{ fontSize: 13, color: C.textMuted }}>Loading…</div>}
              {staffListLoaded && staffList.length === 0 && (
                <div style={{ fontSize: 13, color: C.textMuted }}>No staff invited yet.</div>
              )}
              {staffListLoaded && staffList.length > 0 && (
                <div>
                  {staffList.map((s) => (
                    <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13 }}>
                      <div>
                        <span style={{ color: C.text, fontWeight: 500 }}>{s.name || s.phone}</span>
                        {s.name && <span style={{ color: C.textMuted, marginLeft: 8 }}>{s.phone}</span>}
                        <span
                          style={{
                            marginLeft: 10,
                            fontSize: 10.5,
                            fontWeight: 600,
                            borderRadius: 10,
                            padding: "1px 8px",
                            color: s.status === "active" ? C.green : s.status === "revoked" ? C.red : C.amber,
                            background: s.status === "active" ? "#E7FFDB" : s.status === "revoked" ? "#FBE6E6" : "#FFF1DB",
                          }}
                        >
                          {s.status === "active" ? "Active" : s.status === "revoked" ? "Revoked" : "Invited - not linked yet"}
                        </span>
                      </div>
                      {s.status !== "revoked" && (
                        <button onClick={() => revokeStaffAccess(s.id)} style={{ background: "transparent", border: `1px solid ${C.red}`, color: C.red, borderRadius: 6, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>
                          Revoke
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div style={{ background: C.white, border: `1px solid ${C.red}`, borderRadius: 8, padding: "18px 20px", marginTop: 16 }}>
              <div style={{ fontFamily: display, fontSize: 16, fontWeight: 600, marginBottom: 4, color: C.red }}>Delete account</div>
              <div style={{ fontSize: 12.5, color: C.textMuted, marginBottom: 14, lineHeight: 1.5 }}>
                Permanently deletes your account and everything in it - bills, business details, customers, invoices, share links, and any linked WhatsApp number's data. This cannot be undone.
              </div>
              {!deleteConfirming ? (
                <button onClick={() => setDeleteConfirming(true)} style={{ background: "transparent", color: C.red, border: `1px solid ${C.red}`, borderRadius: 6, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>
                  Delete my account
                </button>
              ) : (
                <div>
                  <div style={{ fontSize: 13, color: C.text, marginBottom: 10, fontWeight: 500 }}>
                    Are you sure? This deletes everything and cannot be undone.
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={deleteAccount} disabled={deleting} style={{ background: C.red, color: C.white, border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 13, fontWeight: 500, cursor: deleting ? "default" : "pointer", opacity: deleting ? 0.7 : 1 }}>
                      {deleting ? "Deleting…" : "Yes, permanently delete"}
                    </button>
                    <button onClick={() => { setDeleteConfirming(false); setDeleteError(null); }} disabled={deleting} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 16px", fontSize: 13, cursor: deleting ? "default" : "pointer" }}>
                      Cancel
                    </button>
                  </div>
                </div>
              )}
              {deleteError && <div style={{ fontSize: 12, color: C.red, marginTop: 8 }}>{deleteError}</div>}
            </div>
          </div>
        )}

        {tab === "invoices" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "baseline", gap: 16 }}>
                <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600 }}>{t("Invoices")}</div>
                {(invoiceView === "list" || invoiceView === "customers") && (
                  <div style={{ display: "flex", gap: 4 }}>
                    {[["list", "Invoices"], ["customers", "Customers"]].map(([key, label]) => (
                      <button key={key} onClick={() => setInvoiceView(key)} style={{ background: invoiceView === key ? C.paper : "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 16, padding: "5px 12px", fontSize: 12, fontWeight: 500, cursor: "pointer" }}>
                        {t(label)}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              {invoiceView === "list" && viewMode === "owner" && (
                <button onClick={startNewInvoice} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
                  + {t("New invoice")}
                </button>
              )}
            </div>

            {invoiceView === "customers" && (
              <div>
                {customers.length === 0 && (
                  <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>{t("No customers yet - add one from a new invoice.")}</div>
                )}

                {(() => {
                  const ranked = customerProfitabilityProxy();
                  if (ranked.length < 2) return null;
                  return (
                    <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 16px", marginBottom: 16 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 2 }}>{t("Customer Profitability")}</div>
                      <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                        {t("Paid revenue adjusted for payment delay and at-risk balances - not true margin (BahiSathi doesn't track cost per sale yet), but a sales total alone can be misleading.")}
                      </div>
                      {ranked.slice(0, 6).map((r, i) => (
                        <div key={r.customer.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", borderBottom: i < Math.min(ranked.length, 6) - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                          <span style={{ fontSize: 12, color: C.text }}>{r.customer.name}</span>
                          <span style={{ textAlign: "right" }}>
                            <span style={{ fontFamily: mono, fontSize: 12, color: C.ink }}>₹{Math.round(r.effectiveValue).toLocaleString("en-IN")}</span>
                            {(r.carryingCost + r.riskDiscount) > 0 && (
                              <span style={{ fontSize: 10, color: C.textMuted }}> (of ₹{Math.round(r.paidRevenue).toLocaleString("en-IN")})</span>
                            )}
                          </span>
                        </div>
                      ))}
                    </div>
                  );
                })()}

                {(() => {
                  const churn = customerChurnRadar();
                  if (churn.length === 0) return null;
                  return (
                    <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 16px", marginBottom: 16 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 2 }}>{t("Customers going quiet")}</div>
                      <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                        {t("Ordering pattern has lapsed vs their own normal cadence - worth a check-in before it becomes lost revenue.")}
                      </div>
                      {churn.slice(0, 6).map((r) => (
                        <div key={r.customer.id} style={{ padding: "7px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span style={{ fontSize: 12.5, fontWeight: 500, color: C.text }}>{r.customer.name}</span>
                            <span style={{ fontFamily: mono, fontSize: 11.5, color: C.amber }}>~₹{Math.round(r.revenueAtRiskAnnual).toLocaleString("en-IN")}/{t("annually")}</span>
                          </div>
                          <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>
                            {t("normally orders every")} {r.cadenceLow}-{r.cadenceHigh} {t("days")} · {t("Last order")} {r.daysSinceLast} {t("days ago")}
                          </div>
                        </div>
                      ))}
                    </div>
                  );
                })()}

                {customers.length > 0 && (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {customers
                      .map((c) => ({ c, p: paymentProfileFor(c.id) }))
                      .filter(({ p }) => p.invoiceCount > 0)
                      .sort((a, b) => b.p.outstanding - a.p.outstanding)
                      .map(({ c, p }) => {
                        const trendLabel = p.trend === "worse" ? "Getting worse ↓" : p.trend === "better" ? "Improving ↑" : "Steady";
                        const trendColor = p.trend === "worse" ? C.red : p.trend === "better" ? C.green : C.textMuted;
                        const phoneDigits = c.phone ? c.phone.replace(/\D/g, "") : null;
                        const draft = followUpDraft(c, p, businessProfile?.business_name);
                        const waLink = phoneDigits && draft ? `https://wa.me/${phoneDigits}?text=${encodeURIComponent(draft.message)}` : null;
                        const tierColor = draft ? { friendly: C.green, reminder: C.textMuted, overdue: C.amber, final: C.red }[draft.tier] : C.textMuted;
                        const trust = trustScoreFor(p);
                        return (
                          <div key={c.id} style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "14px 16px" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                              <div>
                                <div style={{ fontSize: 14.5, fontWeight: 600, color: C.text, display: "flex", alignItems: "center", gap: 8 }}>
                                  {c.name}
                                  {trust && (
                                    <span style={{ fontSize: 10.5, fontWeight: 600, color: C.white, background: trust.color, borderRadius: 10, padding: "2px 8px" }}>
                                      Trust {trust.score} · {trust.label}
                                    </span>
                                  )}
                                </div>
                                <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>
                                  {p.avgDelay != null ? `Average payment: ${p.avgDelay} day${p.avgDelay === 1 ? "" : "s"}` : "No payment history yet"}
                                  {p.avgDelay != null && <> · <span style={{ color: trendColor }}>{trendLabel}</span></>}
                                </div>
                              </div>
                              <div style={{ textAlign: "right" }}>
                                <div style={{ fontFamily: mono, fontSize: 16, fontWeight: 700, color: p.outstanding > 0 ? C.amber : C.green }}>
                                  ₹{p.outstanding.toLocaleString("en-IN")}
                                </div>
                                <div style={{ fontSize: 11, color: C.textMuted }}>
                                  {p.unpaidCount > 0 ? `${p.unpaidCount} unpaid invoice${p.unpaidCount === 1 ? "" : "s"}` : "Fully paid"}
                                </div>
                              </div>
                            </div>
                            {p.unpaidCount > 0 && draft && (
                              <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${C.paperLine}` }}>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                                  <div style={{ fontSize: 11.5, color: p.maxOverdueDays > 30 ? C.red : C.textMuted, display: "flex", alignItems: "center", gap: 6 }}>
                                    Oldest unpaid invoice: {p.maxOverdueDays} day{p.maxOverdueDays === 1 ? "" : "s"} ago
                                    <span style={{ fontSize: 10, fontWeight: 600, color: C.white, background: tierColor, borderRadius: 10, padding: "2px 8px" }}>{draft.tierLabel}</span>
                                  </div>
                                  {waLink ? (
                                    <a href={waLink} target="_blank" rel="noopener noreferrer" style={{ background: "#128C7E", color: C.white, borderRadius: 6, padding: "6px 12px", fontSize: 12, fontWeight: 500, textDecoration: "none" }}>
                                      Send WhatsApp reminder
                                    </a>
                                  ) : (
                                    <span style={{ fontSize: 11, color: C.textMuted }}>Add a phone number to send reminders</span>
                                  )}
                                </div>
                                <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 8, background: C.paper, borderRadius: 6, padding: "8px 10px", lineHeight: 1.5, fontStyle: "italic" }}>
                                  "{draft.message}"
                                </div>
                              </div>
                            )}
                            <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${C.paperLine}` }}>
                              {!creditAdviceText[c.id] ? (
                                <button
                                  onClick={() => askCreditAdvice(c, p, trust)}
                                  disabled={creditAdviceLoading && creditAdviceFor === c.id}
                                  style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}
                                >
                                  {creditAdviceLoading && creditAdviceFor === c.id ? "Thinking…" : `Can I give ${c.name.split(" ")[0]} more credit?`}
                                </button>
                              ) : (
                                <div style={{ background: "#E7FFDB", border: "1px solid #BCE5A8", borderRadius: "10px 10px 10px 2px", padding: "9px 12px", fontSize: 12.5, color: "#1F2C1A", lineHeight: 1.5 }}>
                                  {creditAdviceText[c.id]}
                                </div>
                              )}
                            </div>
                          </div>
                        );
                      })}
                  </div>
                )}
              </div>
            )}

            {invoiceView === "list" && (
              <div>
                {!invoicesLoaded && <div style={{ fontSize: 13.5, color: C.textMuted }}>Loading…</div>}
                {invoicesLoaded && invoices.length === 0 && (
                  <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>No invoices yet. Create your first one to bill a customer.</div>
                )}
                {invoices.length > 0 && (
                  <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "6px 16px" }}>
                    {invoices.map((inv) => (
                      <div key={inv.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5, flexWrap: "wrap", gap: 6 }}>
                        <div style={{ flex: 1, minWidth: 140 }}>
                          <div style={{ color: C.text, fontWeight: 500 }}>
                            {inv.invoice_number} · {customerNameFor(inv.customer_id)}
                            {addedByLabel(inv.created_by) && (
                              <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600, color: C.inkLight, background: "#EAF0FF", borderRadius: 10, padding: "1px 7px" }}>👤 {addedByLabel(inv.created_by)}</span>
                            )}
                          </div>
                          <div style={{ fontSize: 11, color: C.textMuted }}>{inv.invoice_date}</div>
                        </div>
                        <div style={{ fontFamily: mono, color: C.ink, fontWeight: 500, marginRight: 4 }}>₹{Number(inv.total).toLocaleString("en-IN")}</div>
                        <div style={{ position: "relative" }}>
                          <button
                            onClick={viewMode === "owner" ? () => (inv.status === "paid" ? toggleInvoiceStatus(inv) : setPayMethodPickerId(inv.id)) : undefined}
                            style={{ background: inv.status === "paid" ? "#E7FFDB" : "#FFF1DB", color: inv.status === "paid" ? C.green : C.amber, border: "none", borderRadius: 20, padding: "4px 10px", fontSize: 11, fontWeight: 600, cursor: viewMode === "owner" ? "pointer" : "default" }}
                          >
                            {inv.status === "paid" ? `Paid${inv.payment_method ? ` · ${inv.payment_method}` : ""}` : "Unpaid"}
                          </button>
                          {viewMode === "owner" && payMethodPickerId === inv.id && (
                            <>
                              <div onClick={() => setPayMethodPickerId(null)} style={{ position: "fixed", inset: 0, zIndex: 10 }} />
                              <div style={{ position: "absolute", top: "calc(100% + 4px)", right: 0, background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, boxShadow: "0 4px 14px rgba(0,0,0,0.12)", padding: 6, zIndex: 11, minWidth: 110 }}>
                                <div style={{ fontSize: 10.5, color: C.textMuted, padding: "2px 8px 4px" }}>Paid by</div>
                                {["Cash", "UPI", "Card", "Other"].map((m) => (
                                  <button
                                    key={m}
                                    onClick={() => { toggleInvoiceStatus(inv, m); setPayMethodPickerId(null); }}
                                    style={{ display: "block", width: "100%", textAlign: "left", background: "transparent", border: "none", color: C.text, borderRadius: 5, padding: "6px 8px", fontSize: 12.5, cursor: "pointer" }}
                                    onMouseEnter={(e) => (e.currentTarget.style.background = C.paper)}
                                    onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                                  >
                                    {m}
                                  </button>
                                ))}
                              </div>
                            </>
                          )}
                        </div>
                        <button onClick={() => setPreviewInvoice(inv)} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 6, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>
                          View
                        </button>
                        {role === "owner" && (
                          <button onClick={() => deleteInvoiceRow(inv.id)} style={{ background: "transparent", border: `1px solid ${C.red}`, color: C.red, borderRadius: 6, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>
                            Delete
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {invoiceView === "new" && invoiceDraft && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "18px 20px" }}>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Customer</div>
                  {!newCustomerMode ? (
                    <div style={{ display: "flex", gap: 8 }}>
                      <select
                        value={invoiceDraft.customer_id}
                        onChange={(e) => { setInvoiceDraft((d) => ({ ...d, customer_id: e.target.value })); setTransportSearch(""); }}
                        style={{ flex: 1, display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text }}
                      >
                        <option value="">Select a customer</option>
                        {customers.map((c) => (
                          <option key={c.id} value={c.id}>{c.name}</option>
                        ))}
                      </select>
                      <button onClick={() => setNewCustomerMode(true)} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 6, padding: "9px 14px", fontSize: 12.5, cursor: "pointer", whiteSpace: "nowrap" }}>
                        + New
                      </button>
                    </div>
                  ) : (
                    <div style={{ background: C.paper, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: 12 }}>
                      <input value={newCustomer.name} onChange={(e) => setNewCustomer((c) => ({ ...c, name: e.target.value }))} placeholder="Customer name" style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, marginBottom: 8, background: C.white, color: C.text }} />
                      <input value={newCustomer.gstin} onChange={(e) => setNewCustomer((c) => ({ ...c, gstin: e.target.value.toUpperCase() }))} placeholder="GSTIN (optional)" style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, fontFamily: mono, marginBottom: 8, background: C.white, color: C.text }} />
                      <input value={newCustomer.address} onChange={(e) => setNewCustomer((c) => ({ ...c, address: e.target.value }))} placeholder="Address (optional)" style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, marginBottom: 8, background: C.white, color: C.text }} />
                      <input value={newCustomer.phone} onChange={(e) => setNewCustomer((c) => ({ ...c, phone: e.target.value }))} placeholder="Phone (optional)" style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, marginBottom: 8, background: C.white, color: C.text }} />
                      {customerError && <div style={{ fontSize: 12, color: C.red, marginBottom: 8 }}>{customerError}</div>}
                      <div style={{ display: "flex", gap: 8 }}>
                        <button onClick={saveNewCustomer} disabled={customerSaving} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "7px 14px", fontSize: 12.5, cursor: customerSaving ? "default" : "pointer" }}>
                          {customerSaving ? "Saving…" : "Save customer"}
                        </button>
                        <button onClick={() => { setNewCustomerMode(false); setCustomerError(null); }} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.textMuted, borderRadius: 6, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
                          Cancel
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Invoice date</div>
                  <input
                    type="date"
                    value={invoiceDraft.invoice_date}
                    onChange={(e) => setInvoiceDraft((d) => ({ ...d, invoice_date: e.target.value }))}
                    style={{ display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text }}
                  />
                </div>

                <div style={{ marginBottom: 8 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Items</div>
                  {invoiceDraft.items.map((it, i) => (
                    <div key={i} style={{ display: "flex", gap: 6, marginBottom: 6, alignItems: "center" }}>
                      <input value={it.description} onChange={(e) => updateInvoiceItem(i, "description", e.target.value)} placeholder="Description" style={{ flex: 1, display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 13, background: C.white, color: C.text }} />
                      <input value={it.qty} onChange={(e) => updateInvoiceItem(i, "qty", e.target.value)} placeholder="Qty" type="number" style={{ width: 56, display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 8px", fontSize: 13, fontFamily: mono, background: C.white, color: C.text }} />
                      <input value={it.rate} onChange={(e) => updateInvoiceItem(i, "rate", e.target.value)} placeholder="Rate" type="number" style={{ width: 90, display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 8px", fontSize: 13, fontFamily: mono, background: C.white, color: C.text }} />
                      <button onClick={() => removeInvoiceItem(i)} style={{ background: "transparent", border: "none", color: C.red, cursor: "pointer", fontSize: 16, padding: "0 4px" }}>×</button>
                    </div>
                  ))}
                  <button onClick={addInvoiceItem} style={{ background: "transparent", border: "none", color: C.ink, textDecoration: "underline", cursor: "pointer", fontSize: 12, padding: 0 }}>
                    + Add item
                  </button>
                </div>

                <div style={{ display: "flex", gap: 12, marginBottom: 14, marginTop: 10 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>GST rate</div>
                    <select value={invoiceDraft.gstRate} onChange={(e) => setInvoiceDraft((d) => ({ ...d, gstRate: e.target.value }))} style={{ width: "100%", display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text }}>
                      {[0, 5, 12, 18, 28].map((r) => (
                        <option key={r} value={r}>{r}%</option>
                      ))}
                    </select>
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Customer state</div>
                    <select value={invoiceDraft.gstMode} onChange={(e) => setInvoiceDraft((d) => ({ ...d, gstMode: e.target.value }))} style={{ width: "100%", display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text }}>
                      <option value="same">Same state (CGST+SGST)</option>
                      <option value="different">Different state (IGST)</option>
                    </select>
                  </div>
                </div>

                <div style={{ marginBottom: 14, background: C.paper, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: 12 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 8, fontWeight: 600 }}>Transport (optional)</div>
                  {invoiceDraft.customer_id && pastTransportsFor(invoiceDraft.customer_id).length > 0 && (
                    <div style={{ marginBottom: 10 }}>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 6 }}>
                        Used before for {customerNameFor(invoiceDraft.customer_id)} - search or tap to reuse:
                      </div>
                      {pastTransportsFor(invoiceDraft.customer_id).length > 4 && (
                        <input
                          value={transportSearch}
                          onChange={(e) => setTransportSearch(e.target.value)}
                          placeholder="Search past transport / vehicle number…"
                          style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, marginBottom: 8, background: C.white, color: C.text }}
                        />
                      )}
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 160, overflowY: "auto" }}>
                        {pastTransportsFor(invoiceDraft.customer_id)
                          .filter((t) => {
                            const q = transportSearch.trim().toLowerCase();
                            if (!q) return true;
                            return (t.transport_name || "").toLowerCase().includes(q) || (t.vehicle_number || "").toLowerCase().includes(q);
                          })
                          .map((t, i) => (
                            <button
                              key={i}
                              onClick={() => setInvoiceDraft((d) => ({ ...d, transport_name: t.transport_name, transport_contact: t.transport_contact, vehicle_number: t.vehicle_number }))}
                              style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 20, padding: "6px 12px", fontSize: 12, color: C.ink, cursor: "pointer", textAlign: "left" }}
                              title={t.invoice_date ? `Last used ${t.invoice_date}` : undefined}
                            >
                              {t.transport_name || "Transport"}{t.vehicle_number ? ` · ${t.vehicle_number}` : ""}
                            </button>
                          ))}
                      </div>
                    </div>
                  )}
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <input
                      value={invoiceDraft.transport_name}
                      onChange={(e) => setInvoiceDraft((d) => ({ ...d, transport_name: e.target.value }))}
                      placeholder="Transport / courier name"
                      style={{ flex: "1 1 160px", display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.white, color: C.text }}
                    />
                    <input
                      value={invoiceDraft.transport_contact}
                      onChange={(e) => setInvoiceDraft((d) => ({ ...d, transport_contact: e.target.value }))}
                      placeholder="Transport contact (optional)"
                      style={{ flex: "1 1 140px", display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.white, color: C.text }}
                    />
                    <input
                      value={invoiceDraft.vehicle_number}
                      onChange={(e) => setInvoiceDraft((d) => ({ ...d, vehicle_number: e.target.value.toUpperCase() }))}
                      placeholder="Vehicle number (optional)"
                      style={{ flex: "1 1 140px", display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, fontFamily: mono, background: C.white, color: C.text }}
                    />
                  </div>
                </div>

                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Notes (optional)</div>
                  <textarea
                    value={invoiceDraft.notes}
                    onChange={(e) => setInvoiceDraft((d) => ({ ...d, notes: e.target.value }))}
                    rows={2}
                    style={{ display: "block", width: "100%", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text, fontFamily: sans, resize: "vertical" }}
                  />
                </div>

                {(() => {
                  const totals = computeInvoiceTotals(invoiceDraft.items, invoiceDraft.gstMode, invoiceDraft.gstRate);
                  return (
                    <div style={{ background: C.paper, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 14px", marginBottom: 14, fontSize: 13 }}>
                      <div style={{ display: "flex", justifyContent: "space-between" }}>
                        <span style={{ color: C.textMuted }}>Subtotal</span>
                        <span style={{ fontFamily: mono }}>₹{totals.subtotal.toLocaleString("en-IN")}</span>
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 600, marginTop: 4 }}>
                        <span>Total</span>
                        <span style={{ fontFamily: mono }}>₹{totals.total.toLocaleString("en-IN")}</span>
                      </div>
                    </div>
                  );
                })()}

                {invoiceError && <div style={{ fontSize: 12.5, color: C.red, marginBottom: 10 }}>{invoiceError}</div>}
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={saveInvoice} disabled={invoiceSaving} style={{ background: C.green, color: C.white, border: "none", borderRadius: 6, padding: "10px 20px", fontSize: 13.5, fontWeight: 500, cursor: invoiceSaving ? "default" : "pointer", opacity: invoiceSaving ? 0.7 : 1 }}>
                    {invoiceSaving ? "Saving…" : "Save invoice"}
                  </button>
                  <button onClick={() => { setInvoiceView("list"); setInvoiceDraft(null); setInvoiceError(null); }} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "10px 16px", fontSize: 13.5, cursor: "pointer" }}>
                    Cancel
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {tab === "closing" && (
          <div>
            <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Evening closing</div>
            <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
              Compare the cash you actually counted against what BahiSathi expects from invoices marked "Paid by Cash" that day.
            </div>

            {(() => {
              const missing = missingMoneyFindings(invoices, bankTransactions, dailyClosings);
              if (missing.length === 0) return null;
              const sevDot = (s) => (s === "high" ? "🔴" : s === "medium" ? "🟡" : "⚪️");
              return (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>Missing Money Detector</div>
                  <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                    Invoices marked paid with no independent trace elsewhere BahiSathi tracks - a matched bank/UPI transaction, or a same-day evening closing for cash sales.
                  </div>
                  {missing.map((f, i) => (
                    <div key={i} style={{ padding: "8px 0", borderBottom: i < missing.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{sevDot(f.severity)} {f.label}</div>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{f.hint}</div>
                    </div>
                  ))}
                </div>
              );
            })()}

            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 5 }}>Date</div>
              <input
                type="date"
                value={closingDate}
                onChange={(e) => setClosingDate(e.target.value)}
                style={{ display: "block", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.white, color: C.text }}
              />
            </div>

            {(() => {
              const expected = expectedCashFor(closingDate);
              const entered = enteredCashInput.trim() === "" ? null : Number(enteredCashInput);
              const difference = entered != null ? entered - expected : null;
              const mismatch = difference != null && Math.abs(difference) > 0.5;
              return (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                    <span style={{ fontSize: 13, color: C.textMuted }}>Expected cash (from paid invoices)</span>
                    <span style={{ fontFamily: mono, fontSize: 15, fontWeight: 600, color: C.ink }}>₹{expected.toLocaleString("en-IN")}</span>
                  </div>

                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 6 }}>Cash counted</div>
                    <input
                      type="number"
                      value={enteredCashInput}
                      onChange={(e) => { setEnteredCashInput(e.target.value); setClosingSaved(false); }}
                      placeholder="e.g. 12500"
                      style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, fontFamily: mono, background: C.white, color: C.text, outline: "none" }}
                    />
                  </div>

                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 6 }}>Note (optional)</div>
                    <input
                      value={closingNote}
                      onChange={(e) => { setClosingNote(e.target.value); setClosingSaved(false); }}
                      placeholder="e.g. ₹200 short - customer paid Rahim with a torn note"
                      style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13, background: C.white, color: C.text, outline: "none" }}
                    />
                  </div>

                  {difference != null && (
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 14px", marginTop: 14, background: mismatch ? "#FFF1DB" : "#E7FFDB", borderRadius: 6 }}>
                      <span style={{ fontSize: 13.5, fontWeight: 600, color: C.text }}>
                        {difference === 0 ? "Matches exactly" : difference > 0 ? "Cash over" : "Cash short"}
                      </span>
                      <span style={{ fontFamily: mono, fontSize: 17, fontWeight: 700, color: difference === 0 ? C.green : mismatch ? C.amber : C.green }}>
                        {difference > 0 ? "+" : ""}₹{difference.toLocaleString("en-IN")}
                      </span>
                    </div>
                  )}

                  <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 14 }}>
                    <button onClick={saveClosing} disabled={closingSaving} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "9px 18px", fontSize: 13, fontWeight: 500, cursor: closingSaving ? "default" : "pointer", opacity: closingSaving ? 0.7 : 1 }}>
                      {closingSaving ? "Saving…" : "Save closing"}
                    </button>
                    {closingSaved && !closingSaving && <span style={{ fontSize: 12.5, color: C.green }}>Saved ✓</span>}
                  </div>
                </div>
              );
            })()}

            {closingsLoaded && dailyClosings.length > 0 && (
              <div style={{ marginTop: 20 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: C.text, marginBottom: 8 }}>Recent closings</div>
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "4px 16px" }}>
                  {dailyClosings.slice(0, 14).map((c) => (
                    <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13 }}>
                      <button onClick={() => setClosingDate(c.closing_date)} style={{ background: "transparent", border: "none", color: C.ink, textDecoration: "underline", cursor: "pointer", fontSize: 13, padding: 0 }}>
                        {c.closing_date}
                      </button>
                      <span style={{ fontFamily: mono, color: Math.abs(Number(c.difference)) > 0.5 ? C.amber : C.green, fontWeight: 500 }}>
                        {Number(c.difference) > 0 ? "+" : ""}₹{Number(c.difference).toLocaleString("en-IN")}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {tab === "audit" && (() => {
          const gstFindings = gstGuardianFindings(entries, invoices, customers);
          const bookFindings = fixMyBooksFindings(entries);
          const recurring = detectRecurringVendors(entries);
          const dueRecurring = recurring.filter((r) => r.status === "overdue" || r.status === "due_soon");
          const priceWatch = vendorPriceWatch(entries);
          const leaks = expenseLeakFindings(entries);
          const sevColor = (s) => (s === "high" ? C.red : s === "medium" ? C.amber : C.textMuted);
          const sevDot = (s) => (s === "high" ? "🔴" : s === "medium" ? "🟡" : "⚪️");
          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Audit</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                Automatic checks over your own recorded data - no GST portal or bank connection needed. Nothing here is filed or changed for you; each is worth a quick look.
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>Smart Recurring Transactions</div>
                <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                  Vendors billed on a regular cadence, based on your own bill history - a nudge before the next one's due, not a subscription tracker.
                </div>
                {recurring.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.textMuted }}>Not enough repeat bills from the same vendor yet to spot a cadence - keep scanning and this fills in.</div>
                ) : (
                  recurring.map((r, i) => (
                    <div key={r.vendor} style={{ padding: "8px 0", borderBottom: i < recurring.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
                        <div>
                          <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>
                            {r.status === "overdue" ? "🔴" : r.status === "due_soon" ? "🟡" : "⚪️"} {r.vendor}
                          </div>
                          <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>
                            {r.cadence.toLowerCase()} · ~₹{Math.round(r.avgAmount).toLocaleString("en-IN")} avg · last billed {r.lastDate}
                            {" · "}
                            {r.status === "overdue"
                              ? `${Math.abs(r.daysUntilDue)} day${Math.abs(r.daysUntilDue) === 1 ? "" : "s"} overdue`
                              : r.status === "due_soon"
                              ? `due in ${r.daysUntilDue} day${r.daysUntilDue === 1 ? "" : "s"}`
                              : `next expected ${r.nextDueDate}`}
                          </div>
                        </div>
                        {quickAddedVendors.includes(r.vendor) ? (
                          <span style={{ fontSize: 11.5, color: C.green, flexShrink: 0 }}>✓ Added</span>
                        ) : (
                          (r.status === "overdue" || r.status === "due_soon") && (
                            <button
                              onClick={() => quickAddRecurringBill(r)}
                              disabled={quickAddingVendor === r.vendor}
                              style={{ background: C.gold, color: C.white, border: "none", borderRadius: 16, padding: "5px 12px", fontSize: 11.5, fontWeight: 500, cursor: quickAddingVendor === r.vendor ? "default" : "pointer", opacity: quickAddingVendor === r.vendor ? 0.7 : 1, flexShrink: 0 }}
                            >
                              {quickAddingVendor === r.vendor ? "Adding…" : "+ Quick add"}
                            </button>
                          )
                        )}
                      </div>
                    </div>
                  ))
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>Vendor Price Watch</div>
                <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                  Items or vendors whose price has crept up compared to what you'd paid before.
                </div>
                {priceWatch.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.green }}>✓ No meaningful price increases spotted against your own history.</div>
                ) : (
                  priceWatch.map((p, i) => (
                    <div key={`${p.vendor}-${p.item}`} style={{ padding: "8px 0", borderBottom: i < priceWatch.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>🟡 {p.vendor}{p.item.toLowerCase() !== p.vendor.toLowerCase() ? ` - ${p.item}` : ""}</div>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>
                        ₹{Math.round(p.earlierAvg).toLocaleString("en-IN")} avg before → ₹{Math.round(p.latestAmount).toLocaleString("en-IN")} on {p.latestDate} ({p.pctChange >= 0 ? "+" : ""}{Math.round(p.pctChange)}%)
                      </div>
                    </div>
                  ))
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>Expense Leak Detector</div>
                <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                  A spend category running hotter than usual, or bills that look like the same charge recorded twice.
                </div>
                {leaks.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.green }}>✓ No category spend creep or possible duplicate charges found.</div>
                ) : (
                  leaks.map((f, i) => (
                    <div key={i} style={{ padding: "8px 0", borderBottom: i < leaks.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{sevDot(f.severity)} {f.label}</div>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{f.hint}</div>
                    </div>
                  ))
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>GST Guardian</div>
                {gstFindings.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.green }}>✓ No GST inconsistencies found in your recorded bills and invoices.</div>
                ) : (
                  gstFindings.map((f, i) => (
                    <div key={i} style={{ padding: "8px 0", borderBottom: i < gstFindings.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{sevDot(f.severity)} {f.label}</div>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{f.hint}</div>
                    </div>
                  ))
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px" }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Fix my books</div>
                {bookFindings.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.green }}>✓ No missing data, duplicates, or unusual amounts found.</div>
                ) : (
                  bookFindings.map((f, i) => (
                    <div key={i} style={{ padding: "8px 0", borderBottom: i < bookFindings.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{sevDot(f.severity)} {f.label}</div>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{f.hint}</div>
                    </div>
                  ))
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginTop: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>Audit Trail</div>
                <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10, lineHeight: 1.5 }}>
                  Every edit, delete, and payment-status change to your bills and invoices, with what changed and a one-click restore - persists across sessions, unlike a browser undo.
                </div>
                {!auditLogLoaded ? (
                  <div style={{ fontSize: 12.5, color: C.textMuted }}>Loading…</div>
                ) : auditLog.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.textMuted }}>No changes logged yet - edits, deletes, and payment-status changes will show up here.</div>
                ) : (
                  auditLog.map((u, i) => {
                    const diffLines = auditDiffLines(u);
                    return (
                      <div key={u.id} style={{ padding: "8px 0", borderBottom: i < auditLog.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
                          <div>
                            <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{u.summary}</div>
                            <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>
                              {new Date(u.created_at).toLocaleString("en-IN", { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}
                              {diffLines.length > 0 && (
                                <button onClick={() => setExpandedLogId(expandedLogId === u.id ? null : u.id)} style={{ marginLeft: 8, background: "transparent", border: "none", color: C.ink, textDecoration: "underline", fontSize: 11, cursor: "pointer", padding: 0, fontFamily: sans }}>
                                  {expandedLogId === u.id ? "Hide changes" : "View changes"}
                                </button>
                              )}
                            </div>
                          </div>
                          {u.before && (
                            <button onClick={() => restoreAuditEntry(u)} disabled={restoringLogId === u.id} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 16, padding: "4px 10px", fontSize: 11, cursor: restoringLogId === u.id ? "default" : "pointer", flexShrink: 0 }}>
                              {restoringLogId === u.id ? "Restoring…" : "Restore"}
                            </button>
                          )}
                        </div>
                        {expandedLogId === u.id && diffLines.length > 0 && (
                          <div style={{ marginTop: 6, background: C.paper, borderRadius: 4, padding: "8px 10px" }}>
                            {diffLines.map((line, li) => (
                              <div key={li} style={{ fontSize: 11.5, color: C.text, fontFamily: mono }}>{line}</div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          );
        })()}

        {tab === "inventory" && (() => {
          const findings = inventoryFindings(inventoryItems, inventoryMovements);
          const sevDot = (s) => (s === "high" ? "🔴" : s === "medium" ? "🟡" : "⚪️");
          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Inventory</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                Track stock on hand and get a nudge before you run out - a manual log for now, not yet linked automatically to bills or invoices.
              </div>

              {viewMode === "owner" && (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Add item</div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <input value={newItemName} onChange={(e) => setNewItemName(e.target.value)} placeholder="Item name" style={{ flex: "2 1 140px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                    <input value={newItemUnit} onChange={(e) => setNewItemUnit(e.target.value)} placeholder="Unit (pcs, kg, box)" style={{ flex: "1 1 100px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                    <input value={newItemQty} onChange={(e) => setNewItemQty(e.target.value)} placeholder="Qty on hand" type="number" style={{ flex: "1 1 100px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                    <input value={newItemReorder} onChange={(e) => setNewItemReorder(e.target.value)} placeholder="Reorder level" type="number" style={{ flex: "1 1 100px", border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }} />
                    <button onClick={addInventoryItem} disabled={addingItem || !newItemName.trim()} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: addingItem ? "default" : "pointer", opacity: addingItem ? 0.7 : 1 }}>
                      {addingItem ? "Adding…" : "+ Add"}
                    </button>
                  </div>
                </div>
              )}

              {findings.length > 0 && (
                <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Worth a look</div>
                  {findings.map((f, i) => (
                    <div key={i} style={{ padding: "8px 0", borderBottom: i < findings.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{sevDot(f.severity)} {f.label}</div>
                      <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{f.hint}</div>
                    </div>
                  ))}
                  {viewMode === "owner" && (
                    <div style={{ marginTop: 10 }}>
                      <button onClick={askInventoryAdvice} disabled={inventoryAdviceLoading} style={{ background: "transparent", color: C.ink, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 14px", fontSize: 12.5, cursor: inventoryAdviceLoading ? "default" : "pointer" }}>
                        {inventoryAdviceLoading ? "Thinking…" : "🤖 What should I reorder?"}
                      </button>
                      {inventoryAdvice && <div style={{ marginTop: 10, fontSize: 12.5, color: C.text, lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{inventoryAdvice}</div>}
                    </div>
                  )}
                </div>
              )}

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "6px 16px" }}>
                {!inventoryLoaded && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "18px 0" }}>Loading…</div>}
                {inventoryLoaded && inventoryItems.length === 0 && (
                  <div style={{ fontSize: 13.5, color: C.textMuted, padding: "18px 0" }}>No items yet - add your first one above.</div>
                )}
                {inventoryItems.map((item) => {
                  const low = Number(item.quantity_on_hand) <= Number(item.reorder_level || 0);
                  return (
                    <div key={item.id} style={{ padding: "10px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                        <div>
                          <span style={{ fontSize: 13.5, fontWeight: 500, color: C.text }}>{item.name}</span>
                          {low && <span style={{ marginLeft: 8, fontSize: 10, fontWeight: 600, color: C.amber, background: "#FFF1DB", borderRadius: 10, padding: "1px 7px" }}>Low stock</span>}
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <span style={{ fontFamily: mono, fontSize: 13.5, color: C.ink }}>{item.quantity_on_hand} {item.unit || "pcs"}</span>
                          {viewMode === "owner" && (
                            <>
                              <button onClick={() => { setAdjustQtyFor(adjustQtyFor === item.id ? null : item.id); setAdjustQtyValue(""); setAdjustQtyReason(""); }} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 16, padding: "4px 10px", fontSize: 11.5, cursor: "pointer" }}>
                                Adjust
                              </button>
                              {role === "owner" && (
                                <button onClick={() => deleteInventoryItem(item.id)} style={{ background: "transparent", border: "none", color: C.red, fontSize: 11.5, cursor: "pointer" }}>
                                  Delete
                                </button>
                              )}
                            </>
                          )}
                        </div>
                      </div>
                      {adjustQtyFor === item.id && (
                        <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                          <input value={adjustQtyValue} onChange={(e) => setAdjustQtyValue(e.target.value)} type="number" placeholder="+10 or -5" style={{ width: 100, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, background: C.paper, color: C.text, outline: "none" }} />
                          <input value={adjustQtyReason} onChange={(e) => setAdjustQtyReason(e.target.value)} placeholder="Reason (optional)" style={{ flex: 1, minWidth: 140, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, background: C.paper, color: C.text, outline: "none" }} />
                          <button onClick={() => adjustInventoryQty(item)} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
                            Save
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {tab === "bank" && (() => {
          const matchedCount = bankTransactions.filter((t) => t.status !== "unmatched").length;
          const totalCredit = bankTransactions.filter((t) => t.direction === "credit").reduce((s, t) => s + (Number(t.amount) || 0), 0);
          const totalDebit = bankTransactions.filter((t) => t.direction === "debit").reduce((s, t) => s + (Number(t.amount) || 0), 0);
          const unpaidInvoices = invoices.filter((i) => i.status !== "paid");
          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Bank & UPI</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                Upload a bank or UPI statement (CSV/Excel export from your bank's app or netbanking) to match payments against your bills and invoices - no live bank connection needed.
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text }}>Upload statement</div>
                  <button onClick={() => bankFileInputRef.current && bankFileInputRef.current.click()} disabled={bankUploading} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: bankUploading ? "default" : "pointer", opacity: bankUploading ? 0.7 : 1 }}>
                    {bankUploading ? "Reading…" : "+ Upload CSV/Excel"}
                  </button>
                  <input ref={bankFileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleBankFileChange} style={{ display: "none" }} />
                </div>
                {bankUploadError && <div style={{ marginTop: 10, color: C.red, fontSize: 12.5 }}>{bankUploadError}</div>}
              </div>

              {bankTransactions.length > 0 && (
                <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
                  <div style={{ flex: "1 1 140px", background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 11, color: C.textMuted }}>Matched</div>
                    <div style={{ fontFamily: mono, fontSize: 16, fontWeight: 600, color: C.ink }}>{matchedCount} / {bankTransactions.length}</div>
                  </div>
                  <div style={{ flex: "1 1 140px", background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 11, color: C.textMuted }}>Money in</div>
                    <div style={{ fontFamily: mono, fontSize: 16, fontWeight: 600, color: C.green }}>₹{totalCredit.toLocaleString("en-IN")}</div>
                  </div>
                  <div style={{ flex: "1 1 140px", background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 11, color: C.textMuted }}>Money out</div>
                    <div style={{ fontFamily: mono, fontSize: 16, fontWeight: 600, color: C.red }}>₹{totalDebit.toLocaleString("en-IN")}</div>
                  </div>
                </div>
              )}

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "6px 16px" }}>
                {!bankLoaded && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "18px 0" }}>Loading…</div>}
                {bankLoaded && bankTransactions.length === 0 && (
                  <div style={{ fontSize: 13.5, color: C.textMuted, padding: "18px 0" }}>No statement uploaded yet - upload a CSV or Excel export above.</div>
                )}
                {bankTransactions.map((t) => {
                  const linkedBill = t.matched_bill_id ? entries.find((e) => e.id === t.matched_bill_id) : null;
                  const linkedInvoice = t.matched_invoice_id ? invoices.find((i) => i.id === t.matched_invoice_id) : null;
                  return (
                    <div key={t.id} style={{ padding: "10px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 6 }}>
                        <div>
                          <div style={{ fontSize: 13.5, fontWeight: 500, color: C.text }}>{t.description || "(no description)"}</div>
                          <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{t.txn_date}</div>
                        </div>
                        <div style={{ textAlign: "right" }}>
                          <div style={{ fontFamily: mono, fontSize: 13.5, fontWeight: 600, color: t.direction === "credit" ? C.green : C.text }}>
                            {t.direction === "credit" ? "+" : "-"}₹{Number(t.amount).toLocaleString("en-IN")}
                          </div>
                          <span style={{ fontSize: 10, fontWeight: 600, borderRadius: 10, padding: "1px 8px", background: t.status === "unmatched" ? "#FFF1DB" : "#E7FFDB", color: t.status === "unmatched" ? C.amber : C.green }}>
                            {t.status === "unmatched" ? "Unmatched" : t.status === "matched_bill" ? "Matched to bill" : "Matched to invoice"}
                          </span>
                        </div>
                      </div>
                      {linkedBill && <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 4 }}>→ {linkedBill.vendor}</div>}
                      {linkedInvoice && <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 4 }}>→ {linkedInvoice.invoice_number}</div>}
                      {t.status === "unmatched" && t.direction === "debit" && (
                        <div style={{ marginTop: 8 }}>
                          <button onClick={() => addBankTxnAsExpense(t)} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 16, padding: "4px 10px", fontSize: 11.5, cursor: "pointer" }}>
                            + Add as expense
                          </button>
                        </div>
                      )}
                      {t.status === "unmatched" && t.direction === "credit" && (
                        <div style={{ marginTop: 8 }}>
                          <button onClick={() => setBankLinkPickerId(bankLinkPickerId === t.id ? null : t.id)} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 16, padding: "4px 10px", fontSize: 11.5, cursor: "pointer" }}>
                            Link to invoice
                          </button>
                          {bankLinkPickerId === t.id && (
                            <div style={{ marginTop: 8, display: "flex", gap: 8, flexWrap: "wrap" }}>
                              {unpaidInvoices.length === 0 && <div style={{ fontSize: 12, color: C.textMuted }}>No unpaid invoices to match.</div>}
                              {unpaidInvoices.map((inv) => (
                                <button key={inv.id} onClick={() => linkBankTxnToInvoice(t, inv.id)} style={{ background: C.paper, border: `1px solid ${C.cardBorder}`, color: C.text, borderRadius: 6, padding: "5px 10px", fontSize: 11.5, cursor: "pointer" }}>
                                  {inv.invoice_number} · ₹{Number(inv.total).toLocaleString("en-IN")}
                                </button>
                              ))}
                            </div>
                          )}
                        </div>
                      )}
                      {t.status !== "unmatched" && (
                        <div style={{ marginTop: 8 }}>
                          <button onClick={() => undoBankTxnMatch(t)} style={{ background: "transparent", border: "none", color: C.textMuted, fontSize: 11.5, cursor: "pointer" }}>Undo match</button>
                        </div>
                      )}
                      <div style={{ marginTop: 4 }}>
                        <button onClick={() => deleteBankTxn(t.id)} style={{ background: "transparent", border: "none", color: C.red, fontSize: 11, cursor: "pointer" }}>Remove</button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}

        {tab === "memory" && (() => {
          const seasonal = seasonalRevenuePattern(invoices);
          const vendors = topVendorsAllTime(entries);
          const behavior = paymentBehaviorTrend(invoices);
          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>Business Memory</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                Patterns BahiSathi has noticed in your own numbers, plus anything you want to jot down and have it remember for later - none of this leaves your account, and it's what grounds the "What if?" tool.
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Seasonal revenue pattern</div>
                {!seasonal ? (
                  <div style={{ fontSize: 12.5, color: C.textMuted }}>Not enough invoice history spread across different months yet - keep invoicing and this will fill in.</div>
                ) : (
                  <div>
                    {seasonal.slice(0, 6).map((row, i) => (
                      <div key={row.month} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: i < 5 ? `1px solid ${C.paperLine}` : "none" }}>
                        <span style={{ fontSize: 12.5, color: C.text }}>{row.month}</span>
                        <span style={{ fontFamily: mono, fontSize: 12.5, color: C.ink }}>₹{Math.round(row.avg).toLocaleString("en-IN")} avg</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Top vendors, all-time</div>
                {vendors.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: C.textMuted }}>No business bills recorded yet.</div>
                ) : (
                  vendors.map((v, i) => (
                    <div key={v.vendor} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: i < vendors.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                      <span style={{ fontSize: 12.5, color: C.text }}>{v.vendor}</span>
                      <span style={{ fontFamily: mono, fontSize: 12.5, color: C.ink }}>₹{Math.round(v.total).toLocaleString("en-IN")}</span>
                    </div>
                  ))
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px", marginBottom: 16 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Customer payment behaviour</div>
                {!behavior ? (
                  <div style={{ fontSize: 12.5, color: C.textMuted }}>Not enough paid invoices with recorded payment dates yet to spot a trend.</div>
                ) : (
                  <div style={{ fontSize: 12.5, color: C.text, lineHeight: 1.6 }}>
                    Customers pay in {behavior.avgDelay} days on average, based on the last {behavior.sampleSize} paid invoices.{" "}
                    {behavior.trend === "worse" ? "That's gotten slower recently." : behavior.trend === "better" ? "That's gotten faster recently." : "That's been steady recently."}
                  </div>
                )}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px" }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 10 }}>Notes</div>
                <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
                  <input
                    value={newNoteText}
                    onChange={(e) => setNewNoteText(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && addBusinessNote()}
                    placeholder='e.g. "Raw material prices usually rise before Diwali"'
                    style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 10px", fontSize: 13, background: C.paper, color: C.text, outline: "none" }}
                  />
                  <button onClick={addBusinessNote} disabled={addingNote || !newNoteText.trim()} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: addingNote ? "default" : "pointer", opacity: addingNote ? 0.7 : 1 }}>
                    {addingNote ? "Saving…" : "+ Add"}
                  </button>
                </div>
                {!notesLoaded && <div style={{ fontSize: 12.5, color: C.textMuted }}>Loading…</div>}
                {notesLoaded && businessNotes.length === 0 && <div style={{ fontSize: 12.5, color: C.textMuted }}>No notes yet - add anything worth remembering about your business.</div>}
                {businessNotes.map((n) => (
                  <div key={n.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, padding: "8px 0", borderBottom: `1px solid ${C.paperLine}` }}>
                    <div style={{ fontSize: 12.5, color: C.text, lineHeight: 1.5 }}>{n.text}</div>
                    <button onClick={() => deleteBusinessNote(n.id)} style={{ background: "transparent", border: "none", color: C.red, fontSize: 11, cursor: "pointer", flexShrink: 0 }}>Remove</button>
                  </div>
                ))}
              </div>
            </div>
          );
        })()}

        {tab === "whatif" && (() => {
          const quickScenarios = [
            "What if I raise my prices by 10%?",
            "What if I hire one more staff member at ₹15,000/month?",
            "What if my top customer stops ordering from me?",
          ];
          return (
            <div>
              <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600, marginBottom: 4 }}>What if?</div>
              <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 18, lineHeight: 1.5 }}>
                Ask about a business decision and get an answer grounded in your real revenue, expenses, cash, and GST numbers - not generic advice.
              </div>

              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
                {quickScenarios.map((s) => (
                  <button key={s} onClick={() => { setWhatIfQuestion(s); askWhatIf(s); }} disabled={whatIfLoading} style={{ background: "transparent", border: `1px solid ${C.cardBorder}`, color: C.ink, borderRadius: 20, padding: "6px 14px", fontSize: 12, cursor: whatIfLoading ? "default" : "pointer" }}>
                    {s}
                  </button>
                ))}
              </div>

              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: "16px 18px" }}>
                <div style={{ display: "flex", gap: 8 }}>
                  <input
                    value={whatIfQuestion}
                    onChange={(e) => setWhatIfQuestion(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && askWhatIf()}
                    placeholder="Ask your own question…"
                    style={{ flex: 1, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "9px 12px", fontSize: 13.5, background: C.paper, color: C.text, outline: "none" }}
                  />
                  <button onClick={() => askWhatIf()} disabled={whatIfLoading || !whatIfQuestion.trim()} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "9px 18px", fontSize: 13, fontWeight: 500, cursor: whatIfLoading ? "default" : "pointer", opacity: whatIfLoading ? 0.7 : 1 }}>
                    {whatIfLoading ? "Thinking…" : "Ask"}
                  </button>
                </div>
                {whatIfError && <div style={{ marginTop: 12, color: C.red, fontSize: 12.5 }}>{whatIfError}</div>}
                {whatIfAnswer && <div style={{ marginTop: 14, fontSize: 13.5, color: C.text, lineHeight: 1.7, whiteSpace: "pre-wrap" }}>{whatIfAnswer}</div>}
              </div>
            </div>
          );
        })()}
      </div>

      {previewInvoice && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(30,51,88,0.55)", zIndex: 50, display: "flex", alignItems: "flex-start", justifyContent: "center", overflowY: "auto", padding: "30px 16px" }}>
          <style>{`
            @media print {
              body * { visibility: hidden; }
              #invoice-print-area, #invoice-print-area * { visibility: visible; }
              #invoice-print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 0; }
              .no-print { display: none !important; }
            }
          `}</style>
          <div id="invoice-print-area" style={{ background: C.white, borderRadius: 8, padding: "32px 36px", maxWidth: 560, width: "100%" }}>
            <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 16 }}>
              <button onClick={() => window.print()} style={{ background: C.ink, color: C.white, border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
                Print / Save as PDF
              </button>
              <button onClick={() => setPreviewInvoice(null)} style={{ background: "transparent", color: C.textMuted, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "8px 14px", fontSize: 13, cursor: "pointer" }}>
                Close
              </button>
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
              <div>
                <div style={{ fontFamily: display, fontSize: 20, fontWeight: 700 }}>{businessProfile?.business_name || "Your Business"}</div>
                {businessProfile?.gstin && <div style={{ fontSize: 12, color: C.textMuted, fontFamily: mono }}>GSTIN: {businessProfile.gstin}</div>}
                {businessProfile?.address && <div style={{ fontSize: 12, color: C.textMuted }}>{businessProfile.address}</div>}
                {businessProfile?.website && <div style={{ fontSize: 12, color: C.textMuted }}>{businessProfile.website}</div>}
                {businessProfile?.email && <div style={{ fontSize: 12, color: C.textMuted }}>{businessProfile.email}</div>}
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontFamily: display, fontSize: 18, fontWeight: 700, color: C.ink }}>INVOICE</div>
                <div style={{ fontSize: 12.5, color: C.textMuted }}>{previewInvoice.invoice_number}</div>
                <div style={{ fontSize: 12.5, color: C.textMuted }}>{previewInvoice.invoice_date}</div>
              </div>
            </div>

            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 11, color: C.textMuted, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 4 }}>Bill to</div>
              <div style={{ fontSize: 14, fontWeight: 600 }}>{customerNameFor(previewInvoice.customer_id)}</div>
              {(() => {
                const cust = customers.find((c) => c.id === previewInvoice.customer_id);
                return cust ? (
                  <div style={{ fontSize: 12, color: C.textMuted }}>
                    {cust.gstin && <div style={{ fontFamily: mono }}>GSTIN: {cust.gstin}</div>}
                    {cust.address && <div>{cust.address}</div>}
                  </div>
                ) : null;
              })()}
            </div>

            {(previewInvoice.transport_name || previewInvoice.vehicle_number) && (
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 11, color: C.textMuted, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 4 }}>Transport</div>
                <div style={{ fontSize: 12.5, color: C.text }}>
                  {previewInvoice.transport_name && <div>{previewInvoice.transport_name}{previewInvoice.transport_contact ? ` · ${previewInvoice.transport_contact}` : ""}</div>}
                  {previewInvoice.vehicle_number && <div style={{ fontFamily: mono }}>Vehicle: {previewInvoice.vehicle_number}</div>}
                </div>
              </div>
            )}

            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 16, fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: `1.5px solid ${C.ink}` }}>
                  <th style={{ textAlign: "left", padding: "6px 4px" }}>Description</th>
                  <th style={{ textAlign: "right", padding: "6px 4px" }}>Qty</th>
                  <th style={{ textAlign: "right", padding: "6px 4px" }}>Rate</th>
                  <th style={{ textAlign: "right", padding: "6px 4px" }}>Amount</th>
                </tr>
              </thead>
              <tbody>
                {(previewInvoice.items || []).map((it, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.paperLine}` }}>
                    <td style={{ padding: "6px 4px" }}>{it.description}</td>
                    <td style={{ textAlign: "right", padding: "6px 4px" }}>{it.qty}</td>
                    <td style={{ textAlign: "right", padding: "6px 4px", fontFamily: mono }}>₹{Number(it.rate).toLocaleString("en-IN")}</td>
                    <td style={{ textAlign: "right", padding: "6px 4px", fontFamily: mono }}>₹{Number(it.amount).toLocaleString("en-IN")}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div style={{ marginLeft: "auto", width: 220 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}>
                <span style={{ color: C.textMuted }}>Subtotal</span>
                <span style={{ fontFamily: mono }}>₹{Number(previewInvoice.subtotal).toLocaleString("en-IN")}</span>
              </div>
              {Number(previewInvoice.cgst) > 0 && (
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}>
                  <span style={{ color: C.textMuted }}>CGST</span>
                  <span style={{ fontFamily: mono }}>₹{Number(previewInvoice.cgst).toLocaleString("en-IN")}</span>
                </div>
              )}
              {Number(previewInvoice.sgst) > 0 && (
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}>
                  <span style={{ color: C.textMuted }}>SGST</span>
                  <span style={{ fontFamily: mono }}>₹{Number(previewInvoice.sgst).toLocaleString("en-IN")}</span>
                </div>
              )}
              {Number(previewInvoice.igst) > 0 && (
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "3px 0" }}>
                  <span style={{ color: C.textMuted }}>IGST</span>
                  <span style={{ fontFamily: mono }}>₹{Number(previewInvoice.igst).toLocaleString("en-IN")}</span>
                </div>
              )}
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, padding: "8px 0 0", borderTop: `1.5px solid ${C.ink}`, marginTop: 6 }}>
                <span>Total</span>
                <span style={{ fontFamily: mono }}>₹{Number(previewInvoice.total).toLocaleString("en-IN")}</span>
              </div>
            </div>

            {previewInvoice.notes && (
              <div style={{ marginTop: 20, fontSize: 12, color: C.textMuted, borderTop: `1px solid ${C.paperLine}`, paddingTop: 10 }}>
                {previewInvoice.notes}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// Public read-only view for whoever holds a CA share link - no login, no
// nav, no session. Shows the ledger and lets them download the same Excel
// export the owner would get.
function SharedLedgerView({ token }) {
  const [data, setData] = useState(undefined); // undefined = loading
  const [error, setError] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const result = await fetchSharedLedger(token);
        setData(result);
      } catch (e) {
        setError(e.message || "This link is invalid or has been revoked.");
      }
    })();
  }, [token]);

  function exportSharedToExcel() {
    if (!data) return;
    const rows = data.entries.map((e) => ({
      Date: e.date || "",
      Vendor: e.vendor || "",
      Category: CATEGORY_LABELS[e.category] || e.category || "",
      CGST: e.cgst ?? "",
      SGST: e.sgst ?? "",
      IGST: e.igst ?? "",
      Total: e.total ?? "",
    }));
    const wb = XLSX.utils.book_new();
    let ws;
    const bp = data.business || {};
    if (bp.business_name || bp.gstin) {
      const headerRows = [
        [bp.business_name || "BahiSathi Ledger"],
        ...(bp.gstin ? [[`GSTIN: ${bp.gstin}`]] : []),
        ...(bp.address ? [[bp.address]] : []),
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

  const total = data ? data.entries.reduce((s, e) => s + (Number(e.total) || 0), 0) : 0;

  return (
    <div style={{ fontFamily: sans, background: C.paper, color: C.text, minHeight: "100vh" }}>
      <style>{`*, *::before, *::after { box-sizing: border-box; }`}</style>
      <div style={{ background: C.ink, padding: "16px 24px", display: "flex", alignItems: "baseline", gap: 10 }}>
        <span style={{ fontFamily: display, fontSize: 18, fontWeight: 600, color: C.white }}>BahiSathi</span>
        <span style={{ fontSize: 12, color: "#B9C3D6" }}>Shared ledger — read-only</span>
      </div>
      <div style={{ padding: 24, maxWidth: 640, margin: "0 auto" }}>
        {data === undefined && !error && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>Loading…</div>}
        {error && <div style={{ fontSize: 13.5, color: C.red, padding: "24px 0" }}>{error}</div>}
        {data && (
          <div>
            {(data.business?.business_name || data.business?.gstin) && (
              <div style={{ marginBottom: 18 }}>
                <div style={{ fontFamily: display, fontSize: 19, fontWeight: 600 }}>{data.business.business_name || "Business"}</div>
                {data.business.gstin && <div style={{ fontSize: 12.5, color: C.textMuted, fontFamily: mono }}>GSTIN: {data.business.gstin}</div>}
                {data.business.address && <div style={{ fontSize: 12.5, color: C.textMuted }}>{data.business.address}</div>}
              </div>
            )}
            {Array.isArray(data.checklist) && data.checklist.length > 0 && (
              <div style={{ background: C.white, border: `1px solid ${C.gold}`, borderRadius: 8, padding: "16px 18px", marginBottom: 18 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, color: C.text, marginBottom: 2 }}>For your CA</div>
                <div style={{ fontSize: 11.5, color: C.textMuted, marginBottom: 10 }}>Automatically generated from this ledger - worth a look before filing.</div>
                {data.checklist.map((item, i) => (
                  <div key={i} style={{ padding: "8px 0", borderBottom: i < data.checklist.length - 1 ? `1px solid ${C.paperLine}` : "none" }}>
                    <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{item.label}</div>
                    <div style={{ fontSize: 11.5, color: C.textMuted, marginTop: 2 }}>{item.detail}</div>
                  </div>
                ))}
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 14 }}>
              <div style={{ fontFamily: display, fontSize: 17, fontWeight: 600 }}>Ledger</div>
              <div style={{ fontFamily: mono, fontSize: 16, color: C.ink }}>₹{total.toLocaleString("en-IN")}</div>
            </div>
            {data.entries.length === 0 && <div style={{ fontSize: 13.5, color: C.textMuted, padding: "24px 0" }}>No bills recorded yet.</div>}
            {data.entries.length > 0 && (
              <div style={{ background: C.white, border: `1px solid ${C.cardBorder}`, borderRadius: 6, padding: "6px 16px" }}>
                {data.entries.slice().reverse().map((e, i) => (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${C.paperLine}`, fontSize: 13.5 }}>
                    <div style={{ color: C.textMuted, width: 78, flexShrink: 0 }}>{e.date || "—"}</div>
                    <div style={{ flex: 1 }}>
                      <div style={{ color: C.text }}>{e.vendor || "—"}</div>
                      <div style={{ fontSize: 11, color: C.textMuted }}>{CATEGORY_LABELS[e.category] || e.category}</div>
                    </div>
                    <div style={{ fontFamily: mono, color: C.ink, fontWeight: 500 }}>₹{Number(e.total || 0).toLocaleString("en-IN")}</div>
                  </div>
                ))}
              </div>
            )}
            {data.entries.length > 0 && (
              <button onClick={exportSharedToExcel} style={{ marginTop: 16, background: C.gold, color: C.white, border: "none", borderRadius: 6, padding: "10px 20px", fontSize: 13.5, fontWeight: 500, cursor: "pointer", width: "100%" }}>
                Download Excel
              </button>
            )}
            <div style={{ fontSize: 11, color: C.textMuted, marginTop: 24, textAlign: "center" }}>
              Shared via BahiSathi — AI se bill padho, khata khud bane
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
  const [lang, setLang] = useState(() => (typeof window !== "undefined" && localStorage.getItem("bahisathi_lang")) || "en");
  useEffect(() => {
    localStorage.setItem("bahisathi_lang", lang);
  }, [lang]);
  const t = (s) => (lang === "hi" && HI[s]) || s;
  const langValue = { lang, setLang, t };

  // A CA share link (?share=TOKEN) bypasses login entirely - check for it
  // before anything session-related renders.
  const shareToken = typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("share") : null;

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

  if (shareToken) {
    return (
      <LangContext.Provider value={langValue}>
        <SharedLedgerView token={shareToken} />
      </LangContext.Provider>
    );
  }

  return (
    <LangContext.Provider value={langValue}>
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
    </LangContext.Provider>
  );
}
