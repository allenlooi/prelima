import React, { useState, useEffect, useRef, useMemo } from "react";
import {
  ArrowRight, ArrowLeft, Link2, Copy, Check, Sparkles, LayoutDashboard,
  FolderKanban, Users, FileText, Receipt, CreditCard, Settings, Sun, Moon,
  Plus, Globe, Upload, ChevronRight, Clock, CircleDollarSign, X, Loader2,
  CheckCircle2, AlertTriangle, HelpCircle, Search, MoreHorizontal, Send,
  FileStack, Wallet, Activity, ExternalLink, PenLine, Mic, Square,
  Undo2, Download, Trash2, Eye, Mail, Lock, AlertCircle
} from "lucide-react";
import { supabase, supabaseConfigured } from "./supabaseClient.js";
import { fetchProjects, syncProjects, fetchQuotes, syncQuotes, fetchTaskBriefs, syncTaskBriefs, ensureProfile, saveWorkspaceName } from "./db.js";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";

/* ------------------------------------------------------------------ */
/* Theme                                                               */
/* ------------------------------------------------------------------ */

const ThemeStyles = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400;12..96,500;12..96,600;12..96,700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

    [data-app="prelima"] {
      --bg: #F6F6F3;
      --surface: #FFFFFF;
      --surface-2: #FBFBF9;
      --ink: #17171C;
      --muted: #6E6E78;
      --line: #E8E8E2;
      --accent: #4353FF;
      --accent-ink: #FFFFFF;
      --accent-soft: #EEF0FF;
      --good: #0E9F6E;
      --warn: #C27803;
      --shadow: 0 1px 2px rgba(20,20,30,.04), 0 8px 24px rgba(20,20,30,.06);
      --shadow-lg: 0 2px 4px rgba(20,20,30,.05), 0 24px 64px rgba(20,20,30,.12);
      font-family: 'Inter', system-ui, sans-serif;
      color: var(--ink);
    }
    [data-app="prelima"][data-theme="dark"] {
      --bg: #0F0F13;
      --surface: #17171D;
      --surface-2: #1C1C23;
      --ink: #F1F1EE;
      --muted: #8F8F9A;
      --line: #272730;
      --accent: #7180FF;
      --accent-ink: #0F0F13;
      --accent-soft: #1D2038;
      --good: #31C48D;
      --warn: #E3A008;
      --shadow: 0 1px 2px rgba(0,0,0,.4), 0 8px 24px rgba(0,0,0,.35);
      --shadow-lg: 0 2px 4px rgba(0,0,0,.5), 0 24px 64px rgba(0,0,0,.5);
    }
    [data-app="prelima"] .display { font-family: 'Bricolage Grotesque', 'Inter', sans-serif; }
    [data-app="prelima"] .mono { font-family: 'JetBrains Mono', monospace; }
    [data-app="prelima"] * { box-sizing: border-box; }
    [data-app="prelima"] ::selection { background: var(--accent); color: var(--accent-ink); }
    [data-app="prelima"] button:focus-visible, [data-app="prelima"] a:focus-visible,
    [data-app="prelima"] input:focus-visible, [data-app="prelima"] textarea:focus-visible {
      outline: 2px solid var(--accent); outline-offset: 2px;
    }
    @keyframes pr-rise { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
    @keyframes pr-fade { from { opacity: 0; } to { opacity: 1; } }
    [data-app="prelima"] .rise { animation: pr-rise .45s cubic-bezier(.2,.7,.2,1) both; }
    [data-app="prelima"] .fade { animation: pr-fade .3s ease both; }
    @media (prefers-reduced-motion: reduce) {
      [data-app="prelima"] .rise, [data-app="prelima"] .fade { animation: none; }
    }
    [data-app="prelima"] input[type="range"] {
      -webkit-appearance: none; appearance: none; width: 100%; height: 6px;
      border-radius: 999px; background: var(--line); cursor: pointer;
    }
    [data-app="prelima"] input[type="range"]::-webkit-slider-thumb {
      -webkit-appearance: none; appearance: none; width: 26px; height: 26px; border-radius: 999px;
      background: var(--accent); border: 4px solid var(--surface); box-shadow: var(--shadow);
    }
    [data-app="prelima"] textarea, [data-app="prelima"] input[type="text"],
    [data-app="prelima"] input[type="email"], [data-app="prelima"] input[type="url"],
    [data-app="prelima"] input[type="date"], [data-app="prelima"] input[type="number"],
    [data-app="prelima"] select {
      background: var(--surface); color: var(--ink); border: 1px solid var(--line);
    }
    [data-app="prelima"] select option { background: var(--surface); color: var(--ink); }
    [data-app="prelima"][data-theme="dark"] { color-scheme: dark; }
    [data-app="prelima"] ::placeholder { color: var(--muted); opacity: .7; }
    .pr-print-only { display: none; }
    @media print {
      body * { visibility: hidden !important; }
      .pr-print-area, .pr-print-area * { visibility: visible !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      .pr-print-area { position: absolute !important; left: 0; top: 0; width: 100%; margin: 0; box-shadow: none !important; }
      .pr-print-only { display: block !important; }
    }
  `}</style>
);

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

const money = (n) => "RM" + Number(n).toLocaleString();
const cx = (...a) => a.filter(Boolean).join(" ");

/* Project value: quotation total when one exists (scoped, accurate),
   otherwise the client's stated budget from the brief (estimate).
   Pipeline value: sum of project values across projects that are not
   yet paid — once an invoice is paid it becomes revenue, not pipeline. */
const quoteSubtotal = (q) => q ? q.items.reduce((s, i) => s + (Number(i.qty) || 0) * (Number(i.unit) || 0), 0) : 0;
const quoteTotal = (q) => q ? Math.round(quoteSubtotal(q) * (1 + (q.sst || 0) / 100)) : 0;
const quoteFor = (p, quotes) => quotes.find(q => q.projectId === p.id) || null;
const projectValue = (p, quotes) => { const q = quoteFor(p, quotes); return q ? quoteTotal(q) : (p.budget || 0); };
const isPaid = (p) => p.invoice && p.invoice.status === "Paid";

async function callClaude(messages, { useSearch = false } = {}) {
  const res = await fetch("/api/claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages, useSearch }),
  });
  const data = await res.json();
  return (data.content || []).filter(b => b.type === "text").map(b => b.text).join("\n");
}

function parseJSON(text) {
  try {
    const clean = text.replace(/```json|```/g, "").trim();
    const start = clean.indexOf("{");
    const end = clean.lastIndexOf("}");
    return JSON.parse(clean.slice(start, end + 1));
  } catch { return null; }
}

const downloadBrief = () => window.print();

/* Heuristic "did they type something real?" guard for required free-text —
   blocks empty, too-short, vowel-less keyboard mashes and single repeated chars. */
function looksReal(s, min = 3) {
  const t = (s || "").trim();
  if (t.length < min) return false;
  const letters = t.replace(/[^a-zA-Z]/g, "");
  if (letters.length >= 3) {
    if (!/[aeiou]/i.test(letters)) return false;
    if (/^(.)\1+$/.test(letters.toLowerCase())) return false;
  }
  return true;
}

async function downloadTaskBriefDocx(a, result, typeLabel) {
  const bullets = (items) => (items || []).map(t => new Paragraph({ text: t, bullet: { level: 0 } }));
  const detail = (label, value) => new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text: `${label}: `, bold: true }), new TextRun(value || "—")],
  });

  const problemLabel = (a.problem || []).map(p => p === "Other" ? (a.problemOther || "Other") : p).join(", ");
  const deliverablesLabel = (a.deliverables || []).filter(d => d.format).map(d => `${d.qty}× ${d.format === "Other" ? (d.other || "Other") : d.format}`).join(", ");
  const locationLabel = (a.audienceLocation || []).map(l => l === "Other" ? (a.audienceLocationOther || "Other") : l).join(", ");
  const hobbyLabel = (a.audienceHobbies || []).map(h => h === "Other" ? (a.audienceHobbiesOther || "Other") : h).join(", ");
  const audienceSummary = [
    (a.audienceAgeRange || []).join(", "),
    (a.audienceGender || []).join("/"),
    locationLabel, hobbyLabel,
  ].filter(Boolean).join(" · ");

  const insightParagraphs = (a.insightQuestions || [])
    .map((q, i) => ((a.insightAnswers || [])[i] ? detail(q, a.insightAnswers[i]) : null))
    .filter(Boolean);

  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ text: a.title || "Task brief", heading: HeadingLevel.TITLE }),
        new Paragraph({ text: (result && result.creativeBrief) || "", spacing: { after: 200 } }),
        ...(result && result.deliverables && result.deliverables.length ? [
          new Paragraph({ text: "Deliverables", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }),
          ...bullets(result.deliverables),
        ] : []),
        ...(result && result.keyRequirements && result.keyRequirements.length ? [
          new Paragraph({ text: "Key requirements", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }),
          ...bullets(result.keyRequirements),
        ] : []),
        new Paragraph({ text: "Details", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }),
        detail("Brand", a.brandName),
        detail("Website", a.brandWebsite),
        detail("Type", typeLabel),
        detail("Deliverables", deliverablesLabel),
        detail("Deadline", a.deadline),
        detail("Problem", problemLabel),
        detail("Audience", audienceSummary),
        ...insightParagraphs,
        detail("References", a.references),
        detail("Avoid", a.referencesAvoid),
        detail("Working files", a.workingFiles),
        detail("Working deck", a.workingDeck),
        detail("Extra links", a.extraLinks),
        detail("Briefed by", a.briefer),
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const el = document.createElement("a");
  el.href = url; el.download = `${(a.title || "brief").replace(/[^\w-]+/g, "-")}.docx`;
  document.body.appendChild(el); el.click();
  document.body.removeChild(el); URL.revokeObjectURL(url);
}

/* ------------------------------------------------------------------ */
/* Atoms                                                               */
/* ------------------------------------------------------------------ */

const Btn = ({ children, variant = "primary", className, ...p }) => (
  <button
    {...p}
    className={cx("inline-flex items-center justify-center gap-2 rounded-xl px-5 py-3 text-sm font-medium transition-all duration-150 active:scale-[.98] disabled:opacity-50 disabled:pointer-events-none", className)}
    style={
      variant === "primary" ? { background: "var(--accent)", color: "var(--accent-ink)" }
      : variant === "ghost" ? { background: "transparent", color: "var(--ink)" }
      : { background: "var(--surface)", color: "var(--ink)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }
    }
  >{children}</button>
);

const Card = ({ children, className, style, ...p }) => (
  <div {...p} className={cx("rounded-2xl", className)}
    style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)", ...style }}>
    {children}
  </div>
);

const Tag = ({ children, tone = "muted" }) => (
  <span className="mono inline-flex items-center rounded-md px-2 py-0.5 text-[11px] font-medium tracking-wide"
    style={tone === "accent" ? { background: "var(--accent-soft)", color: "var(--accent)" }
      : tone === "good" ? { background: "color-mix(in srgb, var(--good) 12%, transparent)", color: "var(--good)" }
      : tone === "warn" ? { background: "color-mix(in srgb, var(--warn) 14%, transparent)", color: "var(--warn)" }
      : { background: "var(--surface-2)", color: "var(--muted)", border: "1px solid var(--line)" }}>
    {children}
  </span>
);

const SectionLabel = ({ children }) => (
  <div className="mono text-[11px] font-medium uppercase tracking-[0.14em] mb-2" style={{ color: "var(--muted)" }}>{children}</div>
);

/* Muted helper shown under any chip group that has an "Other"/"Others" option,
   so people know exactly what to do when their choice isn't listed. */
const OtherHint = () => (
  <p className="mt-3 text-xs" style={{ color: "var(--muted)" }}>Can't see yours? Pick “Other” and type it in.</p>
);

/* A titled section inside the designed brief document (accent heading + body). */
const BriefSection = ({ title, children }) => (
  <div className="mb-6">
    <div className="mono text-[11px] font-medium uppercase tracking-[0.14em] mb-2.5" style={{ color: "var(--accent)" }}>{title}</div>
    {children}
  </div>
);

/* ------------------------------------------------------------------ */
/* Client intake flow — the signature experience                       */
/* ------------------------------------------------------------------ */

const OBJECTIVES = ["Brand awareness", "Lead generation", "Sales / conversions", "Product launch", "Rebrand", "Community growth", "Recruitment", "Education & Understanding", "Engagement & Participation", "Traffic & Discovery", "Retention & Loyalty", "Crisis Management"];
const DELIVERABLE_OPTIONS = ["Logo & identity", "Website / landing page", "Social content", "Video", "Campaign", "Copywriting", "Packaging", "Paid advertising", "Content Creators / KOL", "PR", "Marketing Collaterals", "App development", "Social media management", "Photography", "Animation", "Event Management", "E-Commerce management", "Others"];
const PLATFORMS = ["Instagram", "TikTok", "LinkedIn", "YouTube", "Facebook", "Website", "Email", "Xiaohongshu", "Threads", "X", "WhatsApp", "Telegram", "Blog", "Radio", "Cinema", "PR / Media", "OOH / DOOH", "In-store POSM", "Print", "Offline / print"];

// Budget brackets by currency. Same bracket structure, currency-appropriate rounding.
const CURRENCIES = ["USD", "GBP", "SGD", "MYR", "EUR", "AUD"];
const BUDGET_BRACKETS = {
  USD: [500, 1000, 2000, 5000, 10000],
  GBP: [500, 1000, 2000, 5000, 10000],
  SGD: [500, 1000, 2000, 5000, 10000],
  MYR: [500, 1000, 2000, 5000, 10000],
  EUR: [500, 1000, 2000, 5000, 10000],
  AUD: [500, 1000, 2000, 5000, 10000],
};
const CURRENCY_SYMBOL = { USD: "$", GBP: "£", SGD: "S$", MYR: "RM", EUR: "€", AUD: "A$" };
const budgetRangeOptions = (cur) => {
  const sym = CURRENCY_SYMBOL[cur] || cur + " ";
  const b = BUDGET_BRACKETS[cur] || BUDGET_BRACKETS.USD;
  const fmt = (n) => sym + n.toLocaleString();
  const opts = [[`${fmt(0)} – ${fmt(b[0])}`, b[0]]];
  for (let i = 1; i < b.length; i++) opts.push([`${fmt(b[i - 1])} – ${fmt(b[i])}`, b[i]]);
  opts.push([`Above ${fmt(b[b.length - 1])}`, "custom"]);
  return opts;
};

const blankAnswers = {
  website: "", overview: "", background: "", products: "", objectives: [],
  audienceAgeRange: [], audienceLocation: [], audienceLocationOther: "",
  audienceGender: [], audienceHobbies: [], audienceHobbiesOther: "",
  deliverables: [], platforms: [], timelineMode: "weeks", timelineWeeks: 6, timelineDate: "",
  currency: "MYR", budget: 0, budgetRange: "", budgetFlexible: true,
  revisions: 2, deliverablesOther: "",
  workingFiles: "", briefingDeck: "", brandGuidelines: "",
  briefer: "",
  customFields: [],
};

const TASK_TYPES = ["Design", "Video", "Videography", "Photography", "Animation", "Copywriting", "Social content", "Illustration", "Web Dev", "Ads management", "SEO", "GEO", "Creative Concept", "Campaign Management", "Creators / KOLs", "PR & Media", "Other"];
const DELIVERABLE_FORMATS = ["Carousel (4:5)", "Carousel (1:1)", "Reel / Short video (9:16)", "Video (16:9)", "Story (9:16)", "Static post (1:1)", "Static post (4:5)", "Poster / Print (A4)", "Other"];
const PROBLEM_OPTIONS = ["Brand awareness", "Sales / conversions", "Product launch", "Customer retention", "Rebrand / repositioning", "Education / explaining the product", "Community engagement", "Engagement / participation", "Crisis management", "Other"];
const AGE_BRACKETS = ["13-17", "18-24", "25-34", "35-44", "45-54", "55-64", "65+"];
const LOCATIONS = ["Malaysia", "Singapore", "Southeast Asia", "Asia Pacific", "Europe", "North America", "Global / Worldwide", "Other"];
const GENDERS = ["Male", "Female"];
const HOBBIES = ["Fitness", "Travel", "Fashion & beauty", "Food & dining", "Technology", "Gaming", "Music", "Reading", "Outdoors", "Wellness", "Other"];

const blankTaskAnswers = {
  brandName: "", brandWebsite: "",
  title: "", type: [], typeOther: "",
  deliverables: [{ qty: 1, format: "", other: "" }],
  description: "",
  problem: [], problemOther: "",
  audienceAgeRange: [], audienceLocation: [], audienceLocationOther: "",
  audienceGender: [], audienceHobbies: [], audienceHobbiesOther: "",
  insightQuestions: [], insightAnswers: [],
  references: "", referencesAvoid: "",
  workingFiles: "", workingDeck: "", extraLinks: "",
  deadline: "",
  briefer: "",
};

const Chip = ({ active, children, onClick }) => (
  <button onClick={onClick}
    className="rounded-full px-4 py-2.5 text-sm font-medium transition-all duration-150 active:scale-[.97]"
    style={active
      ? { background: "var(--accent)", color: "var(--accent-ink)", border: "1px solid var(--accent)" }
      : { background: "var(--surface)", color: "var(--ink)", border: "1px solid var(--line)" }}>
    {children}
  </button>
);

const Q = ({ idx, total, title, hint, children }) => (
  <div className="rise" key={idx}>
    <div className="mono text-xs mb-3" style={{ color: "var(--accent)" }}>{String(idx + 1).padStart(2, "0")} / {total}</div>
    <h2 className="display text-2xl md:text-4xl font-semibold leading-tight mb-2">{title}</h2>
    {hint && <p className="text-sm md:text-base mb-6" style={{ color: "var(--muted)" }}>{hint}</p>}
    {!hint && <div className="mb-6" />}
    {children}
  </div>
);

const TA = ({ value, onChange, placeholder, rows = 5, autoFocus = true }) => (
  <textarea rows={rows} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
    className="w-full rounded-2xl p-4 md:p-5 text-base leading-relaxed resize-none"
    style={{ boxShadow: "var(--shadow)" }} autoFocus={autoFocus} />
);

const Slider = ({ value, min, max, step: st = 1, onChange, format }) => (
  <div>
    <div className="display text-4xl md:text-5xl font-semibold mb-6" style={{ color: "var(--accent)" }}>{format(value)}</div>
    <input type="range" min={min} max={max} step={st} value={value} onChange={e => onChange(Number(e.target.value))} />
    <div className="flex justify-between mt-2 mono text-xs" style={{ color: "var(--muted)" }}>
      <span>{format(min)}</span><span>{format(max)}</span>
    </div>
  </div>
);

const VoiceTA = ({ value, onChange, placeholder, rows = 5, cleanHint }) => {
  const [rec, setRec] = useState(false);
  const [busy, setBusy] = useState(false);
  const [prev, setPrev] = useState(null);
  const [note, setNote] = useState("");
  const valRef = useRef(value);
  valRef.current = value;
  const recRef = useRef(null);
  const supported = typeof window !== "undefined" && (window.SpeechRecognition || window.webkitSpeechRecognition);

  const tidy = async (text) => {
    if (!text || !text.trim()) return;
    setBusy(true); setNote("");
    try {
      const out = await callClaude([{
        role: "user",
        content: `Clean up this client's ${cleanHint} into clear, natural sentences. Fix rambling, filler words and speech-to-text errors. Keep their meaning and every detail; do not add anything new. Respond ONLY with the cleaned text, nothing else.\n\n---\n${text}`
      }]);
      const cleaned = (out || "").trim();
      if (cleaned) { setPrev(text); onChange(cleaned); setNote("tidied"); } else setNote("failed");
    } catch { setNote("failed"); }
    setBusy(false);
  };

  const toggleRec = () => {
    if (rec) { try { recRef.current && recRef.current.stop(); } catch {} return; }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) return;
    const r = new SR();
    r.lang = "en-US"; r.continuous = true; r.interimResults = false;
    r.onresult = (e) => {
      let t = "";
      for (let i = e.resultIndex; i < e.results.length; i++) t += e.results[i][0].transcript;
      const nv = (valRef.current ? valRef.current.trim() + " " : "") + t.trim();
      onChange(nv);
    };
    r.onend = () => { setRec(false); if (valRef.current && valRef.current.trim()) tidy(valRef.current); };
    r.onerror = () => { setRec(false); setNote("mic-error"); };
    recRef.current = r;
    try { r.start(); setRec(true); setNote(""); } catch { setNote("mic-error"); }
  };

  return (
    <div>
      <textarea rows={rows} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
        className="w-full rounded-2xl p-4 md:p-5 text-base leading-relaxed resize-none"
        style={{ boxShadow: "var(--shadow)" }} autoFocus />
      <div className="mt-3 flex flex-wrap items-center gap-2.5">
        {supported ? (
          <button onClick={toggleRec}
            className="inline-flex items-center gap-2 rounded-xl px-4 py-2.5 text-sm font-medium transition-all active:scale-[.98]"
            style={rec
              ? { background: "#E02424", color: "#fff" }
              : { background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)", color: "var(--ink)" }}>
            {rec ? <Square className="w-3.5 h-3.5" /> : <Mic className="w-4 h-4" />}
            {rec ? "Stop — I'm listening" : "Speak instead"}
          </button>
        ) : (
          <span className="mono text-[11px]" style={{ color: "var(--muted)" }}>Voice input needs Chrome or Safari</span>
        )}
        <button onClick={() => tidy(value)} disabled={busy || !value.trim()}
          className="inline-flex items-center gap-2 rounded-xl px-4 py-2.5 text-sm font-medium transition-all active:scale-[.98] disabled:opacity-40"
          style={{ background: "var(--accent-soft)", color: "var(--accent)" }}>
          {busy ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />}
          {busy ? "Tidying…" : "Tidy with AI"}
        </button>
        {note === "tidied" && prev !== null && (
          <button onClick={() => { onChange(prev); setPrev(null); setNote(""); }}
            className="inline-flex items-center gap-1.5 text-sm underline underline-offset-4" style={{ color: "var(--muted)" }}>
            <Undo2 className="w-3.5 h-3.5" /> Undo
          </button>
        )}
        {note === "failed" && <span className="text-sm" style={{ color: "var(--warn)" }}>AI tidy unavailable — your text is untouched.</span>}
        {note === "mic-error" && <span className="text-sm" style={{ color: "var(--warn)" }}>Couldn't access the microphone.</span>}
      </div>
    </div>
  );
};

function IntakeFlow({ projectName = "New project", freelancer = "My Studio", onDone, onExit }) {
  const [step, setStep] = useState(-1);
  const [a, setA] = useState(blankAnswers);
  const [saved, setSaved] = useState(true);
  const [aiBusy, setAiBusy] = useState(false);
  const [aiNote, setAiNote] = useState("");
  const [showBg, setShowBg] = useState(false);
  const [submitState, setSubmitState] = useState(null); // null | 'working' | 'done' | 'error'
  const [result, setResult] = useState(null);
  const [timelineWarning, setTimelineWarning] = useState("");
  const [timelineChecking, setTimelineChecking] = useState(false);
  const [editing, setEditing] = useState(false);
  const [preview, setPreview] = useState(false);

  const set = (patch) => { setA(prev => ({ ...prev, ...patch })); setSaved(false); };
  const setResultField = (patch) => setResult(r => ({ ...(r || {}), ...patch }));
  const linesToArr = (str) => str.split("\n").map(x => x.replace(/^[-•]\s*/, "").trim()).filter(Boolean);
  useEffect(() => { if (!saved) { const t = setTimeout(() => setSaved(true), 900); return () => clearTimeout(t); } }, [a, saved]);

  const steps = useMemo(() => ([
    { id: "website", label: "Website & background" },
    { id: "overview", label: "Project overview" },
    { id: "objectives", label: "Objectives" },
    { id: "audience", label: "Target audience" },
    { id: "deliverables", label: "Deliverables" },
    { id: "platforms", label: "Platforms" },
    { id: "timeline", label: "Timeline" },
    { id: "budget", label: "Budget" },
    { id: "revisions", label: "Revisions" },
    { id: "links", label: "Files & links" },
    { id: "briefer", label: "Your name" },
    { id: "review", label: "Review" },
  ]), []);

  const idx = step;
  const total = steps.length;
  const pct = step < 0 ? 0 : Math.round(((step) / total) * 100);

  const locationLabel = a.audienceLocation.map(l => l === "Other" ? (a.audienceLocationOther || "Other") : l).join(", ");
  const hobbyLabel = a.audienceHobbies.map(h => h === "Other" ? (a.audienceHobbiesOther || "Other") : h).join(", ");
  const audienceSummary = [a.audienceAgeRange.join(", "), a.audienceGender.join("/"), locationLabel, hobbyLabel].filter(Boolean).join(" · ");
  const deliverablesLabel = a.deliverables.map(d => d === "Others" ? (a.deliverablesOther || "Others") : d).join(", ");
  const curSym = CURRENCY_SYMBOL[a.currency] || (a.currency + " ");
  const fmtBudget = (n) => curSym + Number(n || 0).toLocaleString();
  const budgetOpts = budgetRangeOptions(a.currency);
  const budgetIsCustom = (() => { const o = budgetOpts.find(r => r[0] === a.budgetRange); return o && o[1] === "custom"; })();

  // On the timeline step, if the timeframe looks tight for the scope, ask AI to sanity-check
  // it and surface a friendly, non-blocking caution.
  useEffect(() => {
    if (steps[idx]?.id !== "timeline") { setTimelineChecking(false); return; }
    const weeks = a.timelineMode === "date"
      ? (a.timelineDate ? Math.round((new Date(a.timelineDate).getTime() - Date.now()) / (7 * 24 * 3600 * 1000)) : null)
      : a.timelineWeeks;
    if (weeks == null || weeks > 2) { setTimelineWarning(""); setTimelineChecking(false); return; }
    let cancelled = false;
    setTimelineChecking(true);
    const t = setTimeout(async () => {
      try {
        const text = await callClaude([{
          role: "user",
          content: `A client wants this delivered in about ${weeks <= 0 ? "less than a week" : weeks + " week(s)"}.\n\nProject: ${a.overview}\nDeliverables: ${deliverablesLabel || "—"}\nObjectives: ${a.objectives.join(", ") || "—"}\n\nIs that realistic for this scope? If it's too short, reply with ONE short friendly caution (under 25 words) naming the risk. If it's fine, reply exactly "OK".\n\nRespond ONLY with JSON: {"warning": "..."}`
        }]);
        const j = parseJSON(text);
        if (!cancelled) setTimelineWarning(j && j.warning && j.warning.trim().toUpperCase() !== "OK" ? j.warning.trim() : "");
      } catch {
        if (!cancelled) setTimelineWarning(weeks <= 1 ? "That's a very tight timeline for this scope — expect rushed delivery and limited revisions." : "");
      }
      if (!cancelled) setTimelineChecking(false);
    }, 700);
    return () => { cancelled = true; clearTimeout(t); };
  }, [idx, a.timelineWeeks, a.timelineDate, a.timelineMode]);

  async function analyseWebsite() {
    if (!a.website) return;
    setAiBusy(true); setAiNote("");
    try {
      const text = await callClaude([{
        role: "user",
        content: `Research this company's website: ${a.website}\n\nWrite a business background ABOUT THE COMPANY — what they do, what they sell or offer, who their customers are, and how they position themselves. Write it in first person plural, as if the business is describing itself (e.g. "We're a specialty coffee roaster with three cafés\u2026").\n\nStrict rules: never mention or describe the website itself. No phrases like "the website", "the site", "their online presence", or comments on design/polish. Only talk about the business. Keep the background to 2-3 sentences.\n\nSeparately, list their products/services as short bullet points, one per line, each starting with "- ". If there are many, group and summarise into a handful of clear bullets rather than listing every item. Do NOT cram products into the background paragraph.\n\nRespond ONLY with a JSON object, no preamble, no markdown fences:\n{"background": "the 2-3 sentence background", "products": "- product or service one\\n- product or service two"}`
      }], { useSearch: true });
      const j = parseJSON(text);
      if (j && j.background) {
        set({ background: j.background, products: j.products || "" });
        setShowBg(true);
        setAiNote("Here's your business background — edit anything that's off.");
      } else { setShowBg(true); setAiNote("Couldn't read that site — write a couple of lines about the business instead."); }
    } catch {
      setShowBg(true);
      setAiNote("Couldn't reach the analysis service — write a couple of lines about the business instead.");
    }
    setAiBusy(false);
  }

  async function submit() {
    setSubmitState("working");
    const payload = { ...a, projectName };
    try {
      const text = await callClaude([{
        role: "user",
        content: `You are structuring a client intake into a professional creative brief for a freelancer.\n\nClient answers (JSON):\n${JSON.stringify(payload, null, 2)}\n\nRespond ONLY with JSON, no preamble, no markdown fences:\n{"professionalBrief": "a well-written multi-paragraph creative brief in professional plain English",\n"missingInfo": ["specific information the client did not provide"],\n"followUpQuestions": ["sharp questions the freelancer should ask before quoting"],\n"unclearRequirements": ["requirements that are ambiguous, quoting the vague phrase"],\n"scopeGaps": ["work implied by the answers but not explicitly scoped"]}`
      }]);
      const j = parseJSON(text);
      if (j && j.professionalBrief) { setResult(j); setSubmitState("done"); onDone && onDone(j, payload); return; }
      throw new Error("bad json");
    } catch {
      const fallback = {
        professionalBrief: `${payload.overview}\n\n${payload.background}${payload.products ? `\n\nProducts / services:\n${payload.products}` : ""}\n\nObjectives: ${payload.objectives.join(", ") || "—"}. Audience: ${audienceSummary || "—"}.\nDeliverables: ${deliverablesLabel || "—"} across ${payload.platforms.join(", ") || "—"}.\nTimeline: ${payload.timelineMode === "date" ? (payload.timelineDate || "—") : `${payload.timelineWeeks} weeks`}. Budget: ${fmtBudget(payload.budget)}${payload.budgetFlexible ? " (flexible)" : " (fixed)"} with ${payload.revisions} revision rounds.`,
        missingInfo: [], followUpQuestions: [], unclearRequirements: [], scopeGaps: [],
        _fallback: true,
      };
      setResult(fallback); setSubmitState("done"); onDone && onDone(fallback, payload);
    }
  }

  const canNext = () => {
    const s = steps[idx]?.id;
    if (s === "website") return looksReal(a.background, 10);
    if (s === "overview") return looksReal(a.overview, 10);
    if (s === "objectives") return a.objectives.length > 0;
    if (s === "audience") return a.audienceAgeRange.length > 0 && a.audienceLocation.length > 0 && a.audienceGender.length > 0;
    if (s === "deliverables") return a.deliverables.length > 0 && (!a.deliverables.includes("Others") || looksReal(a.deliverablesOther));
    if (s === "budget") return !!a.budgetRange && (!budgetIsCustom || a.budget > 0);
    if (s === "timeline") return a.timelineMode !== "date" || a.timelineDate;
    if (s === "briefer") return a.briefer.trim().length >= 2;
    return true;
  };

  const next = () => setStep(s => Math.min(s + 1, total - 1));
  const back = () => setStep(s => Math.max(s - 1, 0));

  const toggle = (key, val) => set({ [key]: a[key].includes(val) ? a[key].filter(v => v !== val) : [...a[key], val] });

  /* ---- screens ---- */

  const briefDoc = (res, canEdit) => {
    const clarify = res ? [...(res.missingInfo || []), ...(res.unclearRequirements || []), ...(res.scopeGaps || [])] : [];
    return (
    <div className="pr-print-area rounded-2xl overflow-hidden" style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }}>
      <div style={{ background: "var(--accent-soft)", padding: "28px 32px", borderBottom: "3px solid var(--accent)" }}>
        <div className="mono text-[11px] uppercase tracking-[0.18em] mb-2" style={{ color: "var(--accent)" }}>Project Brief</div>
        <h1 className="display text-2xl md:text-3xl font-semibold leading-tight" style={{ color: "var(--ink)" }}>{projectName}</h1>
        <div className="mt-4 flex flex-wrap gap-2">
          {a.objectives.length > 0 && <Tag tone="accent">{a.objectives[0]}{a.objectives.length > 1 ? ` +${a.objectives.length - 1}` : ""}</Tag>}
          {a.budgetRange && <Tag>{fmtBudget(a.budget)}</Tag>}
          <Tag>{a.timelineMode === "date" ? (a.timelineDate || "—") : `${a.timelineWeeks} weeks`}</Tag>
        </div>
      </div>

      <div style={{ padding: "28px 32px" }}>
        <BriefSection title="The brief">
          {canEdit
            ? <textarea value={(res && res.professionalBrief) || ""} onChange={e => setResultField({ professionalBrief: e.target.value })} rows={9}
                className="w-full rounded-xl p-3 text-sm leading-relaxed resize-y" />
            : <p className="text-sm leading-relaxed whitespace-pre-line">{res && res.professionalBrief ? res.professionalBrief : ""}</p>}
        </BriefSection>

        {a.products && a.products.trim() && (
          <BriefSection title="Products & services">
            <div className="text-sm leading-relaxed whitespace-pre-line">{a.products}</div>
          </BriefSection>
        )}

        <BriefSection title="Key details">
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Objectives: </span>{a.objectives.join(", ") || "—"}</div>
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Audience: </span>{audienceSummary || "—"}</div>
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Deliverables: </span>{deliverablesLabel || "—"}</div>
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Platforms: </span>{a.platforms.join(", ") || "—"}</div>
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Timeline: </span>{a.timelineMode === "date" ? (a.timelineDate || "—") : `${a.timelineWeeks} weeks`}</div>
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Budget: </span>{fmtBudget(a.budget)}{a.budgetFlexible ? " (flexible)" : " (fixed)"}</div>
          <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Revisions: </span>{a.revisions}</div>
        </BriefSection>

        {(canEdit || clarify.length > 0) && (
          <BriefSection title="Worth clarifying">
            {canEdit
              ? <textarea value={clarify.join("\n")} onChange={e => setResultField({ missingInfo: linesToArr(e.target.value), unclearRequirements: [], scopeGaps: [] })} rows={4}
                  placeholder={"One point per line"} className="w-full rounded-xl p-3 text-sm resize-y" />
              : <ul className="space-y-1.5">{clarify.map((d, i) => <li key={i} className="text-sm flex gap-2.5"><span style={{ color: "var(--warn)" }}>•</span><span>{d}</span></li>)}</ul>}
          </BriefSection>
        )}

        {(canEdit || (res && res.followUpQuestions && res.followUpQuestions.length > 0)) && (
          <BriefSection title="Questions before quoting">
            {canEdit
              ? <textarea value={((res && res.followUpQuestions) || []).join("\n")} onChange={e => setResultField({ followUpQuestions: linesToArr(e.target.value) })} rows={4}
                  placeholder={"One question per line"} className="w-full rounded-xl p-3 text-sm resize-y" />
              : <ul className="space-y-1.5">{res.followUpQuestions.map((d, i) => <li key={i} className="text-sm flex gap-2.5"><span style={{ color: "var(--accent)" }}>•</span><span>{d}</span></li>)}</ul>}
          </BriefSection>
        )}

        {(a.workingFiles || a.briefingDeck || a.brandGuidelines) && (
          <BriefSection title="Files & links">
            {a.workingFiles && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Working files: </span>{a.workingFiles}</div>}
            {a.briefingDeck && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Briefing deck: </span>{a.briefingDeck}</div>}
            {a.brandGuidelines && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Brand guidelines: </span>{a.brandGuidelines}</div>}
          </BriefSection>
        )}

        <div className="mt-8 pt-5 flex items-center justify-between" style={{ borderTop: "1px solid var(--line)" }}>
          <div className="text-sm"><span style={{ color: "var(--muted)" }}>Briefed by </span><span className="font-medium">{a.briefer || "—"}</span></div>
          <div className="mono text-[11px]" style={{ color: "var(--muted)" }}>Prelima · {new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}</div>
        </div>
      </div>
    </div>
    );
  };

  if (submitState === "working" || submitState === "done") {
    return (
      <IntakeShell pct={100} saved onExit={onExit} projectName={projectName} freelancer={freelancer}>
        <div className="rise max-w-2xl">
          {submitState === "working" ? (<>
            <Loader2 className="w-8 h-8 animate-spin mb-6" style={{ color: "var(--accent)" }} />
            <h2 className="display text-3xl font-semibold mb-3">Turning your answers into a brief…</h2>
            <p style={{ color: "var(--muted)" }}>AI is rewriting your responses into a professional creative brief and checking for anything missing. This takes a few seconds.</p>
          </>) : (<>
            <CheckCircle2 className="w-10 h-10 mb-6" style={{ color: "var(--good)" }} />
            <h2 className="display text-3xl md:text-4xl font-semibold mb-3">Your brief is ready.</h2>
            <p className="mb-8" style={{ color: "var(--muted)" }}>{editing ? "Edit the AI-written parts below, then save." : "Download a copy to share with whoever you're briefing."}</p>
            <div className="flex flex-wrap items-center gap-3">
              {editing ? (
                <Btn onClick={() => setEditing(false)}><Check className="w-4 h-4" /> Save changes</Btn>
              ) : (
                <Btn variant="secondary" onClick={() => setEditing(true)}><PenLine className="w-4 h-4" /> Edit brief</Btn>
              )}
              <Btn onClick={downloadBrief}><Download className="w-4 h-4" /> Download my brief</Btn>
              {onExit && <Btn variant="secondary" onClick={onExit}>Close</Btn>}
            </div>
            <p className="mono text-[11px] mt-5 mb-1" style={{ color: "var(--muted)" }}>Keep a copy — handy if you're briefing other vendors too.</p>
            {briefDoc(result, editing)}
          </>)}
        </div>
      </IntakeShell>
    );
  }

  if (preview) {
    const previewRes = { professionalBrief: [a.overview, a.background, a.products ? `Products / services:\n${a.products}` : ""].filter(Boolean).join("\n\n") };
    return (
      <IntakeShell pct={100} saved onExit={onExit} projectName={projectName} freelancer={freelancer}>
        <div className="rise max-w-2xl">
          <div className="mb-5">
            <div className="mono text-[11px] uppercase tracking-wider mb-1" style={{ color: "var(--accent)" }}>Preview</div>
            <h2 className="display text-2xl font-semibold">How your brief will look</h2>
            <p className="text-sm mt-1" style={{ color: "var(--muted)" }}>This uses your own words. Send to have AI polish the writing and flag anything missing.</p>
          </div>
          {briefDoc(previewRes, false)}
          <div className="mt-6 flex flex-wrap gap-3">
            <Btn variant="secondary" onClick={() => setPreview(false)}><ArrowLeft className="w-4 h-4" /> Back to edit</Btn>
            <Btn onClick={() => { setPreview(false); submit(); }}><Sparkles className="w-4 h-4" /> Send my brief</Btn>
          </div>
        </div>
      </IntakeShell>
    );
  }

  if (step === -1) {
    return (
      <IntakeShell pct={0} saved onExit={onExit} projectName={projectName} freelancer={freelancer}>
        <div className="rise max-w-xl">
          <div className="mono text-xs mb-4" style={{ color: "var(--accent)" }}>NEW PROJECT BRIEF</div>
          <h1 className="display text-3xl md:text-5xl font-semibold leading-tight mb-4">Let's shape this project properly.</h1>
          <p className="text-base md:text-lg mb-8" style={{ color: "var(--muted)" }}>
            A few guided questions — one at a time, about five minutes. Your answers save automatically, so you can leave and come back.
          </p>
          <Btn onClick={() => setStep(0)} className="text-base px-7 py-4">Start <ArrowRight className="w-4 h-4" /></Btn>
          <div className="mt-4 mono text-xs" style={{ color: "var(--muted)" }}>No account needed</div>
        </div>
      </IntakeShell>
    );
  }

  const s = steps[idx].id;

  return (
    <IntakeShell pct={pct} saved={saved} onExit={onExit} projectName={projectName} freelancer={freelancer}>
      <div className="max-w-2xl w-full">
        {s === "website" && (
          <Q idx={idx} total={total} title="Let's start with your business." hint="Paste your website and AI writes your business background for you. No website? You can type it instead.">
            <div className="flex flex-col md:flex-row gap-3">
              <input type="url" value={a.website} onChange={e => set({ website: e.target.value })}
                placeholder="https://yourbusiness.com"
                className="flex-1 rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
              <Btn onClick={analyseWebsite} disabled={!a.website || aiBusy}>
                {aiBusy ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />}
                {aiBusy ? "Reading your site…" : "Analyse"}
              </Btn>
            </div>
            {!(showBg || a.background.trim()) && (
              <button className="mt-5 text-sm underline underline-offset-4" style={{ color: "var(--muted)" }}
                onClick={() => { set({ website: "" }); setAiNote(""); setShowBg(true); }}>
                I don't have a website
              </button>
            )}
            {(showBg || a.background.trim()) && (
              <div className="mt-6 fade">
                <div className="flex items-center gap-2 mb-2">
                  <SectionLabel>Business background</SectionLabel>
                  {aiNote && <span className="mono text-[11px] mb-2" style={{ color: "var(--accent)" }}>· AI draft</span>}
                </div>
                {aiNote && <p className="text-sm mb-3 flex items-start gap-2" style={{ color: "var(--accent)" }}><Sparkles className="w-4 h-4 shrink-0 mt-0.5" /> {aiNote}</p>}
                <VoiceTA value={a.background} onChange={v => set({ background: v })}
                  placeholder="What you do, who you serve, and anything that makes you different…"
                  rows={4} cleanHint="business background" />
                <div className="mt-5">
                  <SectionLabel>Products & services</SectionLabel>
                  <TA value={a.products} onChange={v => set({ products: v })}
                    placeholder={"- Your first product or service\n- Another one"} rows={4} autoFocus={false} />
                </div>
              </div>
            )}
          </Q>
        )}

        {s === "overview" && (
          <Q idx={idx} total={total} title="What's this project about?"
            hint={'In your own words — one or two sentences is fine. Type or speak. Example: "We’re launching a new cold-brew coffee range in July and need a 6-week Instagram + TikTok campaign to drive online pre-orders and hit 500 sales in the first month."'}>
            <VoiceTA value={a.overview} onChange={v => set({ overview: v })} placeholder="We need help with…" cleanHint="project overview" />
          </Q>
        )}

        {s === "objectives" && (
          <Q idx={idx} total={total} title="What should this project achieve?" hint="Pick everything that applies.">
            <div className="flex flex-wrap gap-2.5">
              {OBJECTIVES.map(o => <Chip key={o} active={a.objectives.includes(o)} onClick={() => toggle("objectives", o)}>{o}</Chip>)}
            </div>
          </Q>
        )}

        {s === "audience" && (
          <Q idx={idx} total={total} title="Who are we trying to reach?" hint="A quick sketch of who this is for — age, location and gender at least.">
            <div className="space-y-6">
              <div>
                <SectionLabel>Age range</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {AGE_BRACKETS.map(r => <Chip key={r} active={a.audienceAgeRange.includes(r)} onClick={() => toggle("audienceAgeRange", r)}>{r}</Chip>)}
                </div>
              </div>
              <div>
                <SectionLabel>Location</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {LOCATIONS.map(l => <Chip key={l} active={a.audienceLocation.includes(l)} onClick={() => toggle("audienceLocation", l)}>{l}</Chip>)}
                </div>
                {a.audienceLocation.includes("Other") ? (
                  <input type="text" value={a.audienceLocationOther} onChange={e => set({ audienceLocationOther: e.target.value })}
                    placeholder="Type the location here…" className="mt-3 w-full rounded-xl px-4 py-3 text-sm fade" />
                ) : <OtherHint />}
              </div>
              <div>
                <SectionLabel>Gender</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {GENDERS.map(g => <Chip key={g} active={a.audienceGender.includes(g)} onClick={() => toggle("audienceGender", g)}>{g}</Chip>)}
                </div>
              </div>
              <div>
                <SectionLabel>Hobbies & interests</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {HOBBIES.map(h => <Chip key={h} active={a.audienceHobbies.includes(h)} onClick={() => toggle("audienceHobbies", h)}>{h}</Chip>)}
                </div>
                {a.audienceHobbies.includes("Other") ? (
                  <input type="text" value={a.audienceHobbiesOther} onChange={e => set({ audienceHobbiesOther: e.target.value })}
                    placeholder="Type interests here…" className="mt-3 w-full rounded-xl px-4 py-3 text-sm fade" />
                ) : <OtherHint />}
              </div>
            </div>
          </Q>
        )}

        {s === "deliverables" && (
          <Q idx={idx} total={total} title="What do you need made?" hint="Select all that apply.">
            <div className="flex flex-wrap gap-2.5">
              {DELIVERABLE_OPTIONS.map(d => <Chip key={d} active={a.deliverables.includes(d)} onClick={() => toggle("deliverables", d)}>{d}</Chip>)}
            </div>
            {a.deliverables.includes("Others") ? (
              <input type="text" value={a.deliverablesOther} onChange={e => set({ deliverablesOther: e.target.value })}
                placeholder="Type what else you need here…"
                className="mt-4 w-full rounded-2xl px-5 py-4 text-base fade" style={{ boxShadow: "var(--shadow)" }} autoFocus />
            ) : <OtherHint />}
          </Q>
        )}

        {s === "platforms" && (
          <Q idx={idx} total={total} title="Where will this live?" hint="Platforms and channels the work is for.">
            <div className="flex flex-wrap gap-2.5">
              {PLATFORMS.map(p => <Chip key={p} active={a.platforms.includes(p)} onClick={() => toggle("platforms", p)}>{p}</Chip>)}
            </div>
          </Q>
        )}

        {s === "timeline" && (
          <Q idx={idx} total={total} title="When do you need it?" hint="A rough timeline works, or pick an exact date if you have one.">
            <div className="flex rounded-xl p-0.5 mb-6 w-fit" style={{ background: "var(--surface-2)", border: "1px solid var(--line)" }}>
              {[["weeks", "Rough timeframe"], ["date", "Specific date"]].map(([m, label]) => (
                <button key={m} onClick={() => set({ timelineMode: m })}
                  className="px-3.5 py-1.5 rounded-[10px] text-xs font-medium transition-all"
                  style={a.timelineMode === m ? { background: "var(--surface)", boxShadow: "var(--shadow)", color: "var(--ink)" } : { color: "var(--muted)" }}>
                  {label}
                </button>
              ))}
            </div>
            {a.timelineMode === "date" ? (
              <input type="date" value={a.timelineDate} onChange={e => set({ timelineDate: e.target.value })}
                className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
            ) : (
              <Slider value={a.timelineWeeks} min={1} max={24} onChange={v => set({ timelineWeeks: v })}
                format={v => v === 24 ? "24+ weeks" : `${v} week${v > 1 ? "s" : ""}`} />
            )}
            {timelineChecking && (
              <p className="mt-5 text-sm flex items-center gap-2" style={{ color: "var(--muted)" }}>
                <Loader2 className="w-4 h-4 animate-spin" /> Checking if that's enough time…
              </p>
            )}
            {!timelineChecking && timelineWarning && (
              <div className="mt-5 rounded-xl p-4 flex items-start gap-3 fade"
                style={{ background: "color-mix(in srgb, var(--warn) 12%, transparent)", border: "1px solid color-mix(in srgb, var(--warn) 30%, transparent)" }}>
                <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" style={{ color: "var(--warn)" }} />
                <span className="text-sm" style={{ color: "var(--ink)" }}>{timelineWarning}</span>
              </div>
            )}
          </Q>
        )}

        {s === "budget" && (
          <Q idx={idx} total={total} title="What's the budget for this project?" hint="A one-time budget for this project, not a monthly figure — a range helps scope the work honestly, nothing is locked in.">
            <SectionLabel>Currency</SectionLabel>
            <div className="flex flex-wrap gap-2.5 mb-6">
              {CURRENCIES.map(c => <Chip key={c} active={a.currency === c} onClick={() => set({ currency: c, budgetRange: "", budget: 0 })}>{c}</Chip>)}
            </div>
            <SectionLabel>Range</SectionLabel>
            <select value={a.budgetRange} onChange={e => {
                const range = budgetOpts.find(r => r[0] === e.target.value);
                set({ budgetRange: e.target.value, budget: range && range[1] !== "custom" ? range[1] : a.budget });
              }}
              className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }}>
              <option value="" disabled>Select a budget range…</option>
              {budgetOpts.map(([label]) => <option key={label} value={label}>{label}</option>)}
            </select>
            {budgetIsCustom && (
              <div className="mt-3 flex items-center rounded-2xl px-5 fade" style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }}>
                <span className="mono text-sm mr-2" style={{ color: "var(--muted)" }}>{curSym}</span>
                <input type="number" min="0" step="500" value={a.budget || ""} onChange={e => set({ budget: Number(e.target.value) })}
                  placeholder="Enter your exact budget" className="flex-1 py-4 text-base bg-transparent border-0" style={{ boxShadow: "none" }} autoFocus />
              </div>
            )}
            <div className="mt-8 flex items-center gap-3">
              <button onClick={() => set({ budgetFlexible: !a.budgetFlexible })}
                className="relative w-12 h-7 rounded-full transition-colors"
                style={{ background: a.budgetFlexible ? "var(--accent)" : "var(--line)" }}
                aria-label="Budget is flexible">
                <span className="absolute top-1 w-5 h-5 rounded-full transition-all"
                  style={{ background: "var(--surface)", left: a.budgetFlexible ? "26px" : "4px", boxShadow: "var(--shadow)" }} />
              </button>
              <span className="text-sm">This budget is flexible for the right scope</span>
            </div>
          </Q>
        )}

        {s === "revisions" && (
          <Q idx={idx} total={total} title="How many revision rounds do you expect?" hint="This helps quote fairly — most projects run smoothly on two.">
            <Slider value={a.revisions} min={1} max={5} onChange={v => set({ revisions: v })}
              format={v => `${v} round${v > 1 ? "s" : ""}`} />
          </Q>
        )}

        {s === "links" && (
          <Q idx={idx} total={total} title="Any files or links to share?" hint="Optional — anything that helps whoever picks this up.">
            <SectionLabel>Working files</SectionLabel>
            <input type="url" value={a.workingFiles} onChange={e => set({ workingFiles: e.target.value })}
              placeholder="Link to source files, assets…" className="w-full rounded-xl px-4 py-3 text-sm" autoFocus />
            <div className="mt-4">
              <SectionLabel>Briefing deck</SectionLabel>
              <input type="url" value={a.briefingDeck} onChange={e => set({ briefingDeck: e.target.value })}
                placeholder="Link to a deck or strategy doc…" className="w-full rounded-xl px-4 py-3 text-sm" />
            </div>
            <div className="mt-4">
              <SectionLabel>Brand guidelines</SectionLabel>
              <input type="url" value={a.brandGuidelines} onChange={e => set({ brandGuidelines: e.target.value })}
                placeholder="Link to brand guidelines / brand kit…" className="w-full rounded-xl px-4 py-3 text-sm" />
            </div>
          </Q>
        )}

        {s === "briefer" && (
          <Q idx={idx} total={total} title="Last thing — who's briefing this?" hint="Your name goes on the brief, so whoever picks up the work knows who to ask.">
            <input type="text" value={a.briefer} onChange={e => set({ briefer: e.target.value })}
              placeholder="Your name" className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
          </Q>
        )}

        {s === "review" && (
          <Q idx={idx} total={total} title="Quick check before we send." hint="Tap any answer to change it.">
            <div className="space-y-2">
              {[
                ["Website", a.website || "None", 0],
                ["Background", a.background, 0],
                ["Overview", a.overview, 1],
                ["Objectives", a.objectives.join(", "), 2],
                ["Audience", audienceSummary, 3],
                ["Deliverables", deliverablesLabel, 4],
                ["Timeline", a.timelineMode === "date" ? (a.timelineDate || "—") : `${a.timelineWeeks} weeks`, 6],
                ["Budget", `${fmtBudget(a.budget)}${a.budgetFlexible ? " · flexible" : " · fixed"}`, 7],
                ["Briefed by", a.briefer || "—", 10],
              ].map(([k, v, go]) => (
                <button key={k} onClick={() => setStep(go)} className="w-full text-left rounded-xl px-4 py-3 flex items-start gap-4 transition-colors"
                  style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
                  <span className="mono text-[11px] uppercase tracking-wider w-24 shrink-0 pt-0.5" style={{ color: "var(--muted)" }}>{k}</span>
                  <span className="text-sm flex-1 line-clamp-2">{v || "—"}</span>
                  <PenLine className="w-3.5 h-3.5 shrink-0 mt-1" style={{ color: "var(--muted)" }} />
                </button>
              ))}
            </div>
            <div className="mt-6 flex flex-wrap gap-3">
              <Btn variant="secondary" onClick={() => setPreview(true)} className="text-base px-6 py-4">
                <Eye className="w-4 h-4" /> Preview
              </Btn>
              <Btn onClick={submit} className="text-base px-7 py-4">
                <Sparkles className="w-4 h-4" /> Send my brief
              </Btn>
            </div>
          </Q>
        )}

        {s !== "review" && (
          <div className="mt-8 flex items-center gap-3">
            {idx > 0 && <Btn variant="secondary" onClick={back} aria-label="Back"><ArrowLeft className="w-4 h-4" /></Btn>}
            <Btn onClick={next} disabled={!canNext()}>Continue <ArrowRight className="w-4 h-4" /></Btn>
            {steps[idx].optional && <button onClick={next} className="text-sm underline underline-offset-4" style={{ color: "var(--muted)" }}>Skip</button>}
          </div>
        )}
      </div>
    </IntakeShell>
  );
}

function IntakeShell({ children, pct, saved, onExit, projectName, freelancer }) {
  return (
    <div className="min-h-screen flex flex-col" style={{ background: "var(--bg)" }}>
      <div className="h-1 w-full" style={{ background: "var(--line)" }}>
        <div className="h-1 transition-all duration-500" style={{ width: `${pct}%`, background: "var(--accent)" }} />
      </div>
      <header className="flex items-center justify-between px-5 md:px-10 py-4">
        <div className="flex items-center gap-2">
          <Wordmark small />
        </div>
        <div className="flex items-center gap-4">
          <span className="mono text-[11px] flex items-center gap-1.5" style={{ color: saved ? "var(--good)" : "var(--muted)" }}>
            {saved ? <Check className="w-3 h-3" /> : <Loader2 className="w-3 h-3 animate-spin" />}
            {saved ? "Saved" : "Saving"}
          </span>
          {onExit && <button onClick={onExit} aria-label="Exit intake"><X className="w-4 h-4" style={{ color: "var(--muted)" }} /></button>}
        </div>
      </header>
      <main className="flex-1 flex items-center px-5 md:px-10 pb-16 pt-4 md:pt-0">
        <div className="w-full max-w-3xl mx-auto">{children}</div>
      </main>
    </div>
  );
}

const Wordmark = ({ small }) => (
  <span className={cx("display font-semibold tracking-tight", small ? "text-base" : "text-xl")}>
    prelima<span style={{ color: "var(--accent)" }}>.</span>
  </span>
);

/* ------------------------------------------------------------------ */
/* Creative / task brief — freelancer briefing their own team          */
/* ------------------------------------------------------------------ */

function TaskBriefFlow({ freelancer = "My Studio", onDone, onExit }) {
  const [step, setStep] = useState(-1);
  const [a, setA] = useState(blankTaskAnswers);
  const [saved, setSaved] = useState(true);
  const [submitState, setSubmitState] = useState(null); // null | 'working' | 'done'
  const [result, setResult] = useState(null);
  const [link] = useState(() => `prelima.app/t/${Math.random().toString(36).slice(2, 6)}`);
  const [copied, setCopied] = useState(false);
  const [insightsLoading, setInsightsLoading] = useState(false);
  const [editing, setEditing] = useState(false);
  const [preview, setPreview] = useState(false);

  const set = (patch) => { setA(prev => ({ ...prev, ...patch })); setSaved(false); };
  const setResultField = (patch) => setResult(r => ({ ...(r || {}), ...patch }));
  const linesToArr = (s) => s.split("\n").map(x => x.replace(/^[-•]\s*/, "").trim()).filter(Boolean);
  const toggle = (key, val) => set({ [key]: a[key].includes(val) ? a[key].filter(v => v !== val) : [...a[key], val] });
  const setDeliv = (i, patch) => set({ deliverables: a.deliverables.map((d, idx) => idx === i ? { ...d, ...patch } : d) });
  const addDeliv = () => set({ deliverables: [...a.deliverables, { qty: 1, format: "", other: "" }] });
  const removeDeliv = (i) => set({ deliverables: a.deliverables.filter((_, idx) => idx !== i) });
  useEffect(() => { if (!saved) { const t = setTimeout(() => setSaved(true), 900); return () => clearTimeout(t); } }, [a, saved]);

  const steps = useMemo(() => ([
    { id: "brand", label: "Brand or website" },
    { id: "title", label: "Project name" },
    { id: "type", label: "Type of work" },
    { id: "format", label: "Deliverables" },
    { id: "description", label: "Description" },
    { id: "problem", label: "Problem to solve" },
    { id: "audience", label: "Target audience" },
    { id: "insights", label: "Insights" },
    { id: "references", label: "References" },
    { id: "links", label: "Working files & links" },
    { id: "deadline", label: "Deadline" },
    { id: "briefer", label: "Your name" },
    { id: "review", label: "Review" },
  ]), []);

  const idx = step;
  const total = steps.length;
  const pct = step < 0 ? 0 : Math.round((step / total) * 100);
  const typeLabel = a.type.map(t => t === "Other" ? (a.typeOther || "Other") : t).join(", ");
  const deliverablesLabel = a.deliverables.filter(d => d.format).map(d => `${d.qty}× ${d.format === "Other" ? (d.other || "Other") : d.format}`).join(", ");
  const problemLabel = a.problem.map(p => p === "Other" ? (a.problemOther || "Other") : p).join(", ");
  const locationLabel = a.audienceLocation.map(l => l === "Other" ? (a.audienceLocationOther || "Other") : l).join(", ");
  const hobbyLabel = a.audienceHobbies.map(h => h === "Other" ? (a.audienceHobbiesOther || "Other") : h).join(", ");
  const ageRangeLabel = a.audienceAgeRange.join(", ");
  const audienceSummary = [
    ageRangeLabel, a.audienceGender.join("/"),
    locationLabel, hobbyLabel,
  ].filter(Boolean).join(" · ");

  // Generate 3 tailored quick-fire questions once the user reaches Insights, so it's
  // guided instead of a blank box — most people know the context, just not how to phrase it.
  // The questions lean on the audience persona: their relationship with the problem,
  // what's been tried before, why they'd pick this brand.
  useEffect(() => {
    if (steps[idx]?.id !== "insights" || a.insightQuestions.length > 0 || insightsLoading) return;
    setInsightsLoading(true);
    (async () => {
      try {
        const text = await callClaude([{
          role: "user",
          content: `You're helping someone write a creative brief. Based on what they've shared so far, ask 3 short, specific quick-fire questions that surface useful context they know but haven't put into words yet. Cover angles like: the audience's relationship with the problem, what's been tried before (worked or didn't), and why this audience would choose this brand over alternatives. Make each question concrete to THIS brand and audience, under 14 words.\n\nBrand: ${a.brandName || "—"}${a.brandWebsite ? ` (${a.brandWebsite})` : ""}\nTitle: ${a.title}\nType: ${typeLabel}\nDescription: ${a.description}\nProblem: ${problemLabel || "—"}\nAudience: ${audienceSummary || "—"}\n\nRespond ONLY with JSON, no preamble: {"questions": ["...", "...", "..."]}`
        }]);
        const j = parseJSON(text);
        if (j && Array.isArray(j.questions) && j.questions.length) {
          set({ insightQuestions: j.questions, insightAnswers: j.questions.map(() => "") });
        } else throw new Error("bad json");
      } catch {
        const fallback = ["What's the audience's relationship with this problem?", "What's been tried before — and did it work?", "Why would they pick you over alternatives?"];
        set({ insightQuestions: fallback, insightAnswers: fallback.map(() => "") });
      }
      setInsightsLoading(false);
    })();
  }, [idx]);

  function finish(brief) {
    setResult(brief);
    setSubmitState("done");
    onDone && onDone({
      id: "t" + Date.now(),
      title: a.title || "Untitled brief",
      brandName: a.brandName, brandWebsite: a.brandWebsite,
      type: a.type, typeOther: a.typeOther, description: a.description,
      deliverables: a.deliverables,
      problem: a.problem, problemOther: a.problemOther,
      audienceAgeRange: a.audienceAgeRange, audienceLocation: a.audienceLocation, audienceLocationOther: a.audienceLocationOther,
      audienceGender: a.audienceGender, audienceHobbies: a.audienceHobbies, audienceHobbiesOther: a.audienceHobbiesOther,
      insightQuestions: a.insightQuestions, insightAnswers: a.insightAnswers,
      references: a.references, referencesAvoid: a.referencesAvoid,
      workingFiles: a.workingFiles, workingDeck: a.workingDeck, extraLinks: a.extraLinks,
      deadline: a.deadline, briefer: a.briefer,
      status: "Ready", created: new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" }),
      link, brief,
    });
  }

  async function submit() {
    setSubmitState("working");
    try {
      const text = await callClaude([{
        role: "user",
        content: `You are structuring a task brief a freelancer/agency is handing off to a designer or creative freelancer on their team.\n\nTask answers (JSON):\n${JSON.stringify(a, null, 2)}\n\nRespond ONLY with JSON, no preamble, no markdown fences:\n{"creativeBrief": "a well-written, clear creative brief in professional plain English that weaves in the problem, target audience and any insights provided", "keyRequirements": ["specific must-follow requirements"], "deliverables": ["concrete expected outputs"]}`
      }]);
      const j = parseJSON(text);
      if (j && j.creativeBrief) { finish(j); return; }
      throw new Error("bad json");
    } catch {
      finish({
        creativeBrief: `${a.description}${problemLabel ? `\n\nProblem: ${problemLabel}` : ""}${audienceSummary ? `\n\nAudience: ${audienceSummary}` : ""}\n\nType: ${typeLabel || "—"}. Deadline: ${a.deadline || "—"}.`,
        keyRequirements: [], deliverables: [],
        _fallback: true,
      });
    }
  }

  const canNext = () => {
    const s = steps[idx]?.id;
    if (s === "brand") return looksReal(a.brandName);
    if (s === "title") return looksReal(a.title);
    if (s === "type") return a.type.length > 0 && (!a.type.includes("Other") || looksReal(a.typeOther));
    if (s === "format") return a.deliverables.some(d => d.format && (d.format !== "Other" || looksReal(d.other)));
    if (s === "description") return looksReal(a.description, 10);
    if (s === "problem") return a.problem.length > 0 && (!a.problem.includes("Other") || looksReal(a.problemOther));
    if (s === "audience") return a.audienceAgeRange.length > 0 && a.audienceLocation.length > 0 && a.audienceGender.length > 0;
    if (s === "insights") return a.insightQuestions.length > 0 && a.insightAnswers.every(x => (x || "").trim().length > 0);
    if (s === "deadline") return a.deadline.trim().length > 0;
    if (s === "briefer") return a.briefer.trim().length >= 2;
    // references & links stay optional — forcing URLs that may not exist creates junk data.
    return true;
  };

  const next = () => setStep(s => Math.min(s + 1, total - 1));
  const back = () => setStep(s => Math.max(s - 1, 0));

  const briefDoc = (res, canEdit) => (
    <div className="pr-print-area rounded-2xl overflow-hidden" style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }}>
      <div style={{ background: "var(--accent-soft)", padding: "28px 32px", borderBottom: "3px solid var(--accent)" }}>
        <div className="mono text-[11px] uppercase tracking-[0.18em] mb-2" style={{ color: "var(--accent)" }}>Creative Brief</div>
        <h1 className="display text-2xl md:text-3xl font-semibold leading-tight" style={{ color: "var(--ink)" }}>{a.title || "Task brief"}</h1>
        <div className="mt-2 text-sm" style={{ color: "var(--muted)" }}>
          {a.brandName}{a.brandWebsite ? ` · ${a.brandWebsite}` : ""}
        </div>
        <div className="mt-4 flex flex-wrap gap-2">
          {typeLabel && <Tag tone="accent">{typeLabel}</Tag>}
          {a.deadline && <Tag>Due {a.deadline}</Tag>}
        </div>
      </div>

      <div style={{ padding: "28px 32px" }}>
        <BriefSection title="The brief">
          {canEdit
            ? <textarea value={(res && res.creativeBrief) || ""} onChange={e => setResultField({ creativeBrief: e.target.value })} rows={8}
                className="w-full rounded-xl p-3 text-sm leading-relaxed resize-y" />
            : <p className="text-sm leading-relaxed whitespace-pre-line">{res && res.creativeBrief ? res.creativeBrief : ""}</p>}
        </BriefSection>

        {deliverablesLabel && (
          <BriefSection title="Deliverables">
            <ul className="space-y-1.5">
              {a.deliverables.filter(d => d.format).map((d, i) => (
                <li key={i} className="text-sm flex items-baseline gap-2.5">
                  <span className="mono font-semibold" style={{ color: "var(--accent)" }}>{d.qty}×</span>
                  <span>{d.format === "Other" ? (d.other || "Other") : d.format}</span>
                </li>
              ))}
            </ul>
          </BriefSection>
        )}

        {(canEdit || (res && res.deliverables && res.deliverables.length > 0)) && (
          <BriefSection title="Expected outputs">
            {canEdit
              ? <textarea value={((res && res.deliverables) || []).join("\n")} onChange={e => setResultField({ deliverables: linesToArr(e.target.value) })} rows={4}
                  placeholder={"One output per line"} className="w-full rounded-xl p-3 text-sm resize-y" />
              : <ul className="space-y-1.5">{res.deliverables.map((d, i) => <li key={i} className="text-sm flex gap-2.5"><span style={{ color: "var(--accent)" }}>•</span><span>{d}</span></li>)}</ul>}
          </BriefSection>
        )}

        {(canEdit || (res && res.keyRequirements && res.keyRequirements.length > 0)) && (
          <BriefSection title="Key requirements">
            {canEdit
              ? <textarea value={((res && res.keyRequirements) || []).join("\n")} onChange={e => setResultField({ keyRequirements: linesToArr(e.target.value) })} rows={4}
                  placeholder={"One requirement per line"} className="w-full rounded-xl p-3 text-sm resize-y" />
              : <ul className="space-y-1.5">{res.keyRequirements.map((d, i) => <li key={i} className="text-sm flex gap-2.5"><span style={{ color: "var(--accent)" }}>•</span><span>{d}</span></li>)}</ul>}
          </BriefSection>
        )}

        {(problemLabel || audienceSummary || a.insightAnswers.some(Boolean)) && (
          <BriefSection title="Context">
            {problemLabel && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Problem: </span>{problemLabel}</div>}
            {audienceSummary && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Audience: </span>{audienceSummary}</div>}
            {a.insightQuestions.map((q, i) => a.insightAnswers[i] && (
              <div key={i} className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>{q} </span>{a.insightAnswers[i]}</div>
            ))}
          </BriefSection>
        )}

        {(a.references || a.referencesAvoid) && (
          <BriefSection title="References">
            {a.references && <div className="text-sm mb-1.5 whitespace-pre-line"><span style={{ color: "var(--muted)" }}>Follow: </span>{a.references}</div>}
            {a.referencesAvoid && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Avoid: </span>{a.referencesAvoid}</div>}
          </BriefSection>
        )}

        {(a.workingFiles || a.workingDeck || a.extraLinks) && (
          <BriefSection title="Files & links">
            {a.workingFiles && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Working files: </span>{a.workingFiles}</div>}
            {a.workingDeck && <div className="text-sm mb-1.5"><span style={{ color: "var(--muted)" }}>Working deck: </span>{a.workingDeck}</div>}
            {a.extraLinks && <div className="text-sm mb-1.5 whitespace-pre-line"><span style={{ color: "var(--muted)" }}>Extra links: </span>{a.extraLinks}</div>}
          </BriefSection>
        )}

        <div className="mt-8 pt-5 flex items-center justify-between" style={{ borderTop: "1px solid var(--line)" }}>
          <div className="text-sm"><span style={{ color: "var(--muted)" }}>Briefed by </span><span className="font-medium">{a.briefer || "—"}</span></div>
          <div className="mono text-[11px]" style={{ color: "var(--muted)" }}>Prelima · {new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}</div>
        </div>
      </div>
    </div>
  );

  if (submitState === "working" || submitState === "done") {
    return (
      <IntakeShell pct={100} saved onExit={onExit} freelancer={freelancer}>
        <div className="rise max-w-2xl">
          {submitState === "working" ? (<>
            <Loader2 className="w-8 h-8 animate-spin mb-6" style={{ color: "var(--accent)" }} />
            <h2 className="display text-3xl font-semibold mb-3">Structuring your brief…</h2>
            <p style={{ color: "var(--muted)" }}>AI is turning your notes into a clear creative brief. This takes a few seconds.</p>
          </>) : (<>
            <CheckCircle2 className="w-10 h-10 mb-6" style={{ color: "var(--good)" }} />
            <h2 className="display text-3xl md:text-4xl font-semibold mb-3">Brief ready.</h2>
            <p className="mb-8" style={{ color: "var(--muted)" }}>{editing ? "Edit the AI-written parts below, then save." : "Share the link below with whoever's doing the work, or download a copy."}</p>
            <div className="flex flex-wrap items-center gap-3 mb-3">
              <button onClick={() => { navigator.clipboard?.writeText(link); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
                className="flex items-center gap-2 rounded-xl px-4 py-3 mono text-xs"
                style={{ background: "var(--accent-soft)", color: "var(--accent)" }}>
                {copied ? <Check className="w-3.5 h-3.5" /> : <Link2 className="w-3.5 h-3.5" />}
                {copied ? "Link copied" : link}
              </button>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              {editing ? (
                <Btn onClick={() => setEditing(false)}><Check className="w-4 h-4" /> Save changes</Btn>
              ) : (
                <Btn variant="secondary" onClick={() => setEditing(true)}><PenLine className="w-4 h-4" /> Edit brief</Btn>
              )}
              <Btn onClick={downloadBrief}><Download className="w-4 h-4" /> Download PDF</Btn>
              <Btn variant="secondary" onClick={() => downloadTaskBriefDocx(a, result, typeLabel)}><FileText className="w-4 h-4" /> Download Word doc</Btn>
              {onExit && <Btn variant="secondary" onClick={onExit}>Close</Btn>}
            </div>
            {briefDoc(result, editing)}
          </>)}
        </div>
      </IntakeShell>
    );
  }

  if (preview) {
    const previewRes = { creativeBrief: a.description || "", deliverables: [], keyRequirements: [] };
    return (
      <IntakeShell pct={100} saved onExit={onExit} freelancer={freelancer}>
        <div className="rise max-w-2xl">
          <div className="mb-5">
            <div className="mono text-[11px] uppercase tracking-wider mb-1" style={{ color: "var(--accent)" }}>Preview</div>
            <h2 className="display text-2xl font-semibold">How your brief will look</h2>
            <p className="text-sm mt-1" style={{ color: "var(--muted)" }}>This uses your own words. Generate to have AI polish the writing and add expected outputs.</p>
          </div>
          {briefDoc(previewRes, false)}
          <div className="mt-6 flex flex-wrap gap-3">
            <Btn variant="secondary" onClick={() => setPreview(false)}><ArrowLeft className="w-4 h-4" /> Back to edit</Btn>
            <Btn onClick={() => { setPreview(false); submit(); }}><Sparkles className="w-4 h-4" /> Generate brief</Btn>
          </div>
        </div>
      </IntakeShell>
    );
  }

  if (step === -1) {
    return (
      <IntakeShell pct={0} saved onExit={onExit} freelancer={freelancer}>
        <div className="rise max-w-xl">
          <div className="mono text-xs mb-4" style={{ color: "var(--accent)" }}>NEW CREATIVE BRIEF</div>
          <h1 className="display text-3xl md:text-5xl font-semibold leading-tight mb-4">Brief someone on a piece of work.</h1>
          <p className="text-base md:text-lg mb-8" style={{ color: "var(--muted)" }}>
            A few quick questions — about two minutes. Answers save automatically.
          </p>
          <Btn onClick={() => setStep(0)} className="text-base px-7 py-4">Start <ArrowRight className="w-4 h-4" /></Btn>
        </div>
      </IntakeShell>
    );
  }

  const s = steps[idx].id;

  return (
    <IntakeShell pct={pct} saved={saved} onExit={onExit} freelancer={freelancer}>
      <div className="max-w-2xl w-full">
        {s === "brand" && (
          <Q idx={idx} total={total} title="Who's this for?" hint="Your brand or company. If you don't have a website yet, just give us the name.">
            <SectionLabel>Brand / company name</SectionLabel>
            <input type="text" value={a.brandName} onChange={e => set({ brandName: e.target.value })}
              placeholder="e.g. Acme Skincare" className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
            <div className="mt-4">
              <SectionLabel>Website</SectionLabel>
              <input type="url" value={a.brandWebsite} onChange={e => set({ brandWebsite: e.target.value })}
                placeholder="https://…  (optional if you gave a brand name)" className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} />
            </div>
          </Q>
        )}

        {s === "title" && (
          <Q idx={idx} total={total} title="What's the name of this project?" hint={'A short name for this task — e.g. "Instagram carousel for product launch."'}>
            <input type="text" value={a.title} onChange={e => set({ title: e.target.value })}
              placeholder="Task title" className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
          </Q>
        )}

        {s === "type" && (
          <Q idx={idx} total={total} title="What kind of work is this?" hint="Pick everything that applies.">
            <div className="flex flex-wrap gap-2.5">
              {TASK_TYPES.map(t => <Chip key={t} active={a.type.includes(t)} onClick={() => toggle("type", t)}>{t}</Chip>)}
            </div>
            {a.type.includes("Other") ? (
              <input type="text" value={a.typeOther} onChange={e => set({ typeOther: e.target.value })}
                placeholder="Type the kind of work here…"
                className="mt-4 w-full rounded-2xl px-5 py-4 text-base fade" style={{ boxShadow: "var(--shadow)" }} autoFocus />
            ) : <OtherHint />}
          </Q>
        )}

        {s === "format" && (
          <Q idx={idx} total={total} title="What are the deliverables?" hint="How many of each, and in what format. Add a row for each type — e.g. 5× Reel (9:16), 2× Carousel.">
            <div className="space-y-3">
              {a.deliverables.map((d, i) => (
                <div key={i} className="fade">
                  <div className="flex items-center gap-2.5">
                    <div className="flex items-center rounded-xl overflow-hidden shrink-0" style={{ border: "1px solid var(--line)", background: "var(--surface)" }}>
                      <button onClick={() => setDeliv(i, { qty: Math.max(1, d.qty - 1) })} className="px-3 py-3 text-base leading-none" style={{ color: "var(--muted)" }} aria-label="Decrease quantity">−</button>
                      <span className="w-8 text-center text-sm font-medium mono">{d.qty}</span>
                      <button onClick={() => setDeliv(i, { qty: d.qty + 1 })} className="px-3 py-3 text-base leading-none" style={{ color: "var(--muted)" }} aria-label="Increase quantity">+</button>
                    </div>
                    <span className="mono text-sm shrink-0" style={{ color: "var(--muted)" }}>×</span>
                    <select value={d.format} onChange={e => setDeliv(i, { format: e.target.value })}
                      className="flex-1 rounded-xl px-4 py-3 text-sm" style={{ background: "var(--surface)", border: "1px solid var(--line)", color: "var(--ink)" }}>
                      <option value="">Choose a format…</option>
                      {DELIVERABLE_FORMATS.map(f => <option key={f} value={f}>{f}</option>)}
                    </select>
                    {a.deliverables.length > 1 && (
                      <button onClick={() => removeDeliv(i)} className="p-2 shrink-0" style={{ color: "var(--muted)" }} aria-label="Remove row"><X className="w-4 h-4" /></button>
                    )}
                  </div>
                  {d.format === "Other" && (
                    <input type="text" value={d.other || ""} onChange={e => setDeliv(i, { other: e.target.value })}
                      placeholder="Type your format here (e.g. Billboard, Menu, Email header)…" className="mt-2 w-full rounded-xl px-4 py-3 text-sm fade" autoFocus />
                  )}
                </div>
              ))}
            </div>
            <p className="mt-3 text-xs" style={{ color: "var(--muted)" }}>Format not in the list? Choose “Other” and type your own.</p>
            <button onClick={addDeliv} className="mt-3 inline-flex items-center gap-1.5 text-sm font-medium" style={{ color: "var(--accent)" }}>
              <Plus className="w-4 h-4" /> Add another
            </button>
          </Q>
        )}

        {s === "description" && (
          <Q idx={idx} total={total} title="What needs to be made?" hint="Describe the task like you would to a teammate. Type or speak.">
            <VoiceTA value={a.description} onChange={v => set({ description: v })} placeholder="We need…" cleanHint="task description" />
          </Q>
        )}

        {s === "problem" && (
          <Q idx={idx} total={total} title="What problem are we trying to solve?" hint="Pick everything that applies.">
            <div className="flex flex-wrap gap-2.5">
              {PROBLEM_OPTIONS.map(p => <Chip key={p} active={a.problem.includes(p)} onClick={() => toggle("problem", p)}>{p}</Chip>)}
            </div>
            {a.problem.includes("Other") ? (
              <input type="text" value={a.problemOther} onChange={e => set({ problemOther: e.target.value })}
                placeholder="Type the problem here…" className="mt-4 w-full rounded-2xl px-5 py-4 text-base fade" style={{ boxShadow: "var(--shadow)" }} autoFocus />
            ) : <OtherHint />}
          </Q>
        )}

        {s === "audience" && (
          <Q idx={idx} total={total} title="Who's the target audience?" hint="A quick sketch of who this is for — age, location and gender at least. We'll use it to ask smarter questions next.">
            <div className="space-y-6">
              <div>
                <SectionLabel>Age range</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {AGE_BRACKETS.map(r => <Chip key={r} active={a.audienceAgeRange.includes(r)} onClick={() => toggle("audienceAgeRange", r)}>{r}</Chip>)}
                </div>
              </div>
              <div>
                <SectionLabel>Location</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {LOCATIONS.map(l => <Chip key={l} active={a.audienceLocation.includes(l)} onClick={() => toggle("audienceLocation", l)}>{l}</Chip>)}
                </div>
                {a.audienceLocation.includes("Other") && (
                  <input type="text" value={a.audienceLocationOther} onChange={e => set({ audienceLocationOther: e.target.value })}
                    placeholder="Where else?" className="mt-3 w-full rounded-xl px-4 py-3 text-sm fade" />
                )}
              </div>
              <div>
                <SectionLabel>Gender</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {GENDERS.map(g => <Chip key={g} active={a.audienceGender.includes(g)} onClick={() => toggle("audienceGender", g)}>{g}</Chip>)}
                </div>
              </div>
              <div>
                <SectionLabel>Hobbies & interests</SectionLabel>
                <div className="flex flex-wrap gap-2.5">
                  {HOBBIES.map(h => <Chip key={h} active={a.audienceHobbies.includes(h)} onClick={() => toggle("audienceHobbies", h)}>{h}</Chip>)}
                </div>
                {a.audienceHobbies.includes("Other") && (
                  <input type="text" value={a.audienceHobbiesOther} onChange={e => set({ audienceHobbiesOther: e.target.value })}
                    placeholder="What else are they into?" className="mt-3 w-full rounded-xl px-4 py-3 text-sm fade" />
                )}
              </div>
            </div>
          </Q>
        )}

        {s === "insights" && (
          <Q idx={idx} total={total} title="Any insights that could help?" hint="A few quick questions based on what you've told us so far.">
            {insightsLoading ? (
              <div className="flex items-center gap-2 text-sm" style={{ color: "var(--muted)" }}>
                <Loader2 className="w-4 h-4 animate-spin" /> Thinking of a few questions…
              </div>
            ) : (
              <div className="space-y-4">
                {a.insightQuestions.map((q, i) => (
                  <div key={i}>
                    <SectionLabel>{q}</SectionLabel>
                    <input type="text" value={a.insightAnswers[i] || ""} onChange={e => {
                        const next = [...a.insightAnswers]; next[i] = e.target.value; set({ insightAnswers: next });
                      }}
                      placeholder="Your answer…" className="w-full rounded-xl px-4 py-3 text-sm" autoFocus={i === 0} />
                  </div>
                ))}
              </div>
            )}
          </Q>
        )}

        {s === "references" && (
          <Q idx={idx} total={total} title="Any references?" hint="Optional, but it speeds things up.">
            <SectionLabel>Links or examples to follow</SectionLabel>
            <TA value={a.references} onChange={v => set({ references: v })}
              placeholder={"https://…  — like this style\nhttps://…  — similar layout"} rows={3} />
            <div className="mt-6">
              <SectionLabel>What to avoid</SectionLabel>
              <TA value={a.referencesAvoid} onChange={v => set({ referencesAvoid: v })}
                placeholder="Too busy, wrong brand colours, stock photos…" rows={3} autoFocus={false} />
            </div>
          </Q>
        )}

        {s === "links" && (
          <Q idx={idx} total={total} title="Any files or links to share?" hint="Optional — working files, a deck, anything relevant.">
            <SectionLabel>Working files</SectionLabel>
            <input type="url" value={a.workingFiles} onChange={e => set({ workingFiles: e.target.value })}
              placeholder="Link to source files, assets, brand kit…" className="w-full rounded-xl px-4 py-3 text-sm" autoFocus />
            <div className="mt-4">
              <SectionLabel>Working deck</SectionLabel>
              <input type="url" value={a.workingDeck} onChange={e => set({ workingDeck: e.target.value })}
                placeholder="Link to a brief deck, strategy doc…" className="w-full rounded-xl px-4 py-3 text-sm" />
            </div>
            <div className="mt-4">
              <SectionLabel>Extra links</SectionLabel>
              <TA value={a.extraLinks} onChange={v => set({ extraLinks: v })} placeholder={"https://…\nhttps://…"} rows={3} autoFocus={false} />
            </div>
          </Q>
        )}

        {s === "deadline" && (
          <Q idx={idx} total={total} title="When do you need it by?" hint="Pick the date this is due.">
            <input type="date" value={a.deadline} onChange={e => set({ deadline: e.target.value })}
              className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
          </Q>
        )}

        {s === "briefer" && (
          <Q idx={idx} total={total} title="Last thing — who's briefing this?" hint="Your name goes on the brief, so whoever picks up the work knows who to ask.">
            <input type="text" value={a.briefer} onChange={e => set({ briefer: e.target.value })}
              placeholder="Your name" className="w-full rounded-2xl px-5 py-4 text-base" style={{ boxShadow: "var(--shadow)" }} autoFocus />
          </Q>
        )}

        {s === "review" && (
          <Q idx={idx} total={total} title="Quick check before we generate it." hint="Tap any answer to change it.">
            <div className="space-y-2">
              {[
                ["Brand", a.brandName || a.brandWebsite || "—", 0],
                ["Title", a.title, 1],
                ["Type", typeLabel || "—", 2],
                ["Deliverables", deliverablesLabel || "—", 3],
                ["Description", a.description, 4],
                ["Problem", problemLabel || "—", 5],
                ["Audience", audienceSummary || "—", 6],
                ["Insights", a.insightAnswers.filter(Boolean).length > 0 ? `${a.insightAnswers.filter(Boolean).length} answered` : "—", 7],
                ["Deadline", a.deadline || "—", 10],
                ["Briefed by", a.briefer || "—", 11],
              ].map(([k, v, go]) => (
                <button key={k} onClick={() => setStep(go)} className="w-full text-left rounded-xl px-4 py-3 flex items-start gap-4 transition-colors"
                  style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
                  <span className="mono text-[11px] uppercase tracking-wider w-24 shrink-0 pt-0.5" style={{ color: "var(--muted)" }}>{k}</span>
                  <span className="text-sm flex-1 line-clamp-2">{v || "—"}</span>
                  <PenLine className="w-3.5 h-3.5 shrink-0 mt-1" style={{ color: "var(--muted)" }} />
                </button>
              ))}
            </div>
            <div className="mt-6 flex flex-wrap gap-3">
              <Btn variant="secondary" onClick={() => setPreview(true)} className="text-base px-6 py-4">
                <Eye className="w-4 h-4" /> Preview
              </Btn>
              <Btn onClick={submit} className="text-base px-7 py-4">
                <Sparkles className="w-4 h-4" /> Generate brief
              </Btn>
            </div>
          </Q>
        )}

        {s !== "review" && !(s === "insights" && insightsLoading) && (
          <div className="mt-8 flex items-center gap-3">
            {idx > 0 && <Btn variant="secondary" onClick={back} aria-label="Back"><ArrowLeft className="w-4 h-4" /></Btn>}
            <Btn onClick={next} disabled={!canNext()}>Continue <ArrowRight className="w-4 h-4" /></Btn>
          </div>
        )}
      </div>
    </IntakeShell>
  );
}

/* ------------------------------------------------------------------ */
/* Landing page                                                        */
/* ------------------------------------------------------------------ */

function Landing({ onStart, onStartTaskBrief, dark, setDark }) {
  return (
    <div className="min-h-screen flex flex-col" style={{ background: "var(--bg)" }}>
      <header className="flex items-center justify-between px-5 md:px-12 py-5 max-w-5xl mx-auto w-full">
        <Wordmark />
        <div className="flex items-center gap-2">
          <button onClick={() => setDark(!dark)} aria-label="Toggle dark mode" className="p-2 rounded-lg" style={{ color: "var(--muted)" }}>
            {dark ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
          </button>
        </div>
      </header>

      <main className="flex-1 flex flex-col items-center justify-center px-5 text-center py-16 md:py-24">
        <h1 className="rise display text-4xl md:text-6xl font-semibold tracking-tight leading-tight mb-3">
          What are we briefing today<span style={{ color: "var(--accent)" }}>?</span>
        </h1>
        <p className="rise mono text-[11px] uppercase tracking-[0.16em] mb-12" style={{ color: "var(--muted)", animationDelay: ".05s" }}>
          Every successful project starts with a good brief
        </p>
        <div className="rise grid grid-cols-1 md:grid-cols-2 gap-4 w-full max-w-2xl" style={{ animationDelay: ".1s" }}>
          <button onClick={onStart} className="group text-left rounded-2xl p-6 md:p-8 transition-transform duration-150 hover:scale-[1.01] active:scale-[.99]"
            style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }}>
            <Users className="w-6 h-6 mb-4" style={{ color: "var(--accent)" }} />
            <div className="display text-xl font-semibold mb-1.5">Write a brief</div>
            <div className="text-sm leading-relaxed mb-4" style={{ color: "var(--muted)" }}>
              For a client to tell a freelancer or agency what they need. One link, no login required.
            </div>
            <div className="inline-flex items-center gap-1.5 text-sm font-medium" style={{ color: "var(--accent)" }}>
              Start <ArrowRight className="w-3.5 h-3.5" />
            </div>
          </button>
          <button onClick={onStartTaskBrief} className="group text-left rounded-2xl p-6 md:p-8 transition-transform duration-150 hover:scale-[1.01] active:scale-[.99]"
            style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }}>
            <Sparkles className="w-6 h-6 mb-4" style={{ color: "var(--accent)" }} />
            <div className="display text-xl font-semibold mb-1.5">Write a creative brief</div>
            <div className="text-sm leading-relaxed mb-4" style={{ color: "var(--muted)" }}>
              For briefing a designer, writer or freelancer on your own team. No login required.
            </div>
            <div className="inline-flex items-center gap-1.5 text-sm font-medium" style={{ color: "var(--accent)" }}>
              Start <ArrowRight className="w-3.5 h-3.5" />
            </div>
          </button>
        </div>
        <p className="rise mt-8 text-xs flex items-center gap-1.5" style={{ color: "var(--muted)", animationDelay: ".15s" }}>
          <Lock className="w-3 h-3" /> No account needed — nothing is stored on our end. You download your brief, we don't keep a copy.
        </p>
      </main>

      <footer className="border-t px-5 md:px-12 py-8" style={{ borderColor: "var(--line)" }}>
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <Wordmark small />
          <span className="mono text-[11px]" style={{ color: "var(--muted)" }}>Free for life</span>
        </div>
      </footer>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Dashboard shell + pages                                             */
/* ------------------------------------------------------------------ */

function ComingSoonPage({ title, body }) {
  return (
    <div className="fade">
      <PageTitle title={title} />
      <EmptyState icon={Clock} title="Coming soon" body={body} />
    </div>
  );
}

const NAV = [
  ["dashboard", "Dashboard", LayoutDashboard],
  ["briefs", "Creative Briefs", FileStack],
  ["settings", "Settings", Settings],
];

function StatusTag({ s }) {
  const tone = s === "Brief received" ? "accent" : s === "Quoted" || s === "Accepted" || s === "Paid" || s === "Ready" ? "good" : s === "Awaiting brief" || s === "Awaiting payment" || s === "Sent" ? "warn" : "muted";
  return <Tag tone={tone}>{s}</Tag>;
}

function AppShell({ projects, setProjects, quotes, setQuotes, taskBriefs, setTaskBriefs, onNewTaskBrief, onLogout, onPreviewIntake, dark, setDark, wsName, setWsName }) {
  const [page, setPage] = useState("dashboard");
  const [openProject, setOpenProject] = useState(null);
  const [copied, setCopied] = useState("");
  const [showNew, setShowNew] = useState(false);
  const [mobileNav, setMobileNav] = useState(false);
  const [openQuote, setOpenQuote] = useState(null);
  const [openBrief, setOpenBrief] = useState(null);

  const createQuote = (projectId) => {
    const p = projects.find(x => x.id === projectId);
    const id = "q" + Date.now();
    const number = `Q-2026-${String(quotes.length + 15).padStart(3, "0")}`;
    const nq = {
      id, number, projectId: projectId || null,
      client: p ? p.client : "", title: p ? p.name : "",
      mode: "itemised", consolidatedLabel: "", status: "Draft",
      created: "6 Jul 2026", sst: 0, logo: null, paymentMethod: "Bank transfer", paymentTerms: "",
      bank: { name: "", account: "", holder: "" },
      items: [{ title: "", details: "", qty: 1, unit: 0 }],
    };
    setQuotes(qs => [nq, ...qs]);
    setPage("quotations"); setOpenProject(null); setOpenQuote(id);
  };

  const copy = (link) => { setCopied(link); setTimeout(() => setCopied(""), 1500); };

  const project = projects.find(p => p.id === openProject);

  return (
    <div className="min-h-screen flex" style={{ background: "var(--bg)" }}>
      {/* Sidebar */}
      <aside className={cx("fixed md:static inset-y-0 left-0 z-40 w-60 flex-col p-4 transition-transform md:translate-x-0 md:flex",
        mobileNav ? "flex translate-x-0" : "hidden md:flex")}
        style={{ background: "var(--surface)", borderRight: "1px solid var(--line)" }}>
        <div className="flex items-center justify-between px-2 py-2 mb-6">
          <Wordmark small />
          <button className="md:hidden" onClick={() => setMobileNav(false)} aria-label="Close menu"><X className="w-4 h-4" /></button>
        </div>
        <nav className="space-y-1 flex-1">
          {NAV.map(([id, label, Icon]) => (
            <button key={id} onClick={() => { setPage(id); setOpenProject(null); setOpenBrief(null); setMobileNav(false); }}
              className="w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition-colors"
              style={page === id ? { background: "var(--accent-soft)", color: "var(--accent)" } : { color: "var(--muted)" }}>
              <Icon className="w-4 h-4" /> {label}
            </button>
          ))}
        </nav>
        <div className="rounded-xl p-3 flex items-center gap-3" style={{ background: "var(--surface-2)", border: "1px solid var(--line)" }}>
          <div className="w-8 h-8 rounded-full flex items-center justify-center text-xs font-semibold shrink-0" style={{ background: "var(--accent)", color: "var(--accent-ink)" }}>{wsName[0] || "?"}</div>
          <div className="min-w-0 flex-1">
            <div className="text-sm font-medium truncate">{wsName}</div>
            <button onClick={onLogout} className="text-xs" style={{ color: "var(--muted)" }}>Sign out</button>
          </div>
        </div>
      </aside>
      {mobileNav && <div className="fixed inset-0 z-30 md:hidden" style={{ background: "rgba(0,0,0,.4)" }} onClick={() => setMobileNav(false)} />}

      {/* Main */}
      <div className="flex-1 min-w-0">
        <header className="flex items-center justify-between px-5 md:px-8 py-4 border-b sticky top-0 z-20" style={{ borderColor: "var(--line)", background: "var(--bg)" }}>
          <div className="flex items-center gap-3">
            <button className="md:hidden" onClick={() => setMobileNav(true)} aria-label="Open menu"><LayoutDashboard className="w-5 h-5" /></button>
            <div className="hidden md:flex items-center gap-2 rounded-xl px-3 py-2 w-72" style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
              <Search className="w-4 h-4" style={{ color: "var(--muted)" }} />
              <input placeholder="Search projects, clients…" className="bg-transparent text-sm flex-1 outline-none border-0" style={{ color: "var(--ink)" }} />
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={() => setDark(!dark)} aria-label="Toggle dark mode" className="p-2 rounded-lg" style={{ color: "var(--muted)" }}>
              {dark ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>
            <Btn onClick={() => setShowNew(true)} className="px-4 py-2"><Plus className="w-4 h-4" /> New project</Btn>
          </div>
        </header>

        <main className="p-5 md:p-8 max-w-5xl">
          {project ? (
            <ProjectDetail project={project} quotes={quotes} onBack={() => setOpenProject(null)} copy={copy} copied={copied} onPreviewIntake={() => onPreviewIntake(project.name)} />
          ) : page === "dashboard" ? (
            <DashboardHome projects={projects} open={setOpenProject} />
          ) : page === "briefs" ? (
            <CreativeBriefsBoard briefs={taskBriefs} setBriefs={setTaskBriefs} openBrief={openBrief} setOpenBrief={setOpenBrief} onNew={onNewTaskBrief} />
          ) : page === "quotations" ? (
            <QuotationsBoard quotes={quotes} setQuotes={setQuotes} projects={projects}
              openQuote={openQuote} setOpenQuote={setOpenQuote} createQuote={createQuote} />
          ) : page === "invoices" ? (
            <ComingSoonPage title="Invoices" body="Deposits and balances, generated from accepted quotations." />
          ) : page === "payments" ? (
            <PaymentsPage />
          ) : (
            <SettingsPage dark={dark} setDark={setDark} wsName={wsName} setWsName={setWsName} />
          )}
        </main>
      </div>

      {showNew && <NewProjectModal onClose={() => setShowNew(false)} onCreate={(p) => { setProjects(ps => [p, ...ps]); setShowNew(false); setPage("projects"); setOpenProject(p.id); }} />}
    </div>
  );
}

function PageTitle({ title, sub, action }) {
  return (
    <div className="flex items-end justify-between mb-6">
      <div>
        <h1 className="display text-2xl md:text-3xl font-semibold tracking-tight">{title}</h1>
        {sub && <p className="text-sm mt-1" style={{ color: "var(--muted)" }}>{sub}</p>}
      </div>
      {action}
    </div>
  );
}

function DashboardHome({ projects, open }) {
  const briefs = projects.filter(p => p.briefComplete).length;
  const awaiting = projects.length - briefs;
  return (
    <div className="fade">
      <PageTitle title="Dashboard" sub="Your briefs, one view." />

      <div className="grid grid-cols-3 gap-3 md:gap-4 mb-8">
        {[
          ["Projects", projects.length, FolderKanban],
          ["Briefs received", briefs, FileText],
          ["Awaiting brief", awaiting, Clock],
        ].map(([label, val, Icon]) => (
          <Card key={label} className="p-4 md:p-5">
            <Icon className="w-4 h-4 mb-3" style={{ color: "var(--accent)" }} />
            <div className="display text-xl md:text-2xl font-semibold truncate">{val}</div>
            <div className="text-xs md:text-sm mt-0.5" style={{ color: "var(--muted)" }}>{label}</div>
          </Card>
        ))}
      </div>

      <SectionLabel>Projects</SectionLabel>
      {projects.length === 0 ? (
        <EmptyState icon={FolderKanban} title="No projects yet" body="Create your first project to generate a client intake link." />
      ) : (
      <Card className="overflow-hidden">
        <div className="hidden md:grid grid-cols-12 gap-3 px-5 py-3 mono text-[11px] uppercase tracking-wider border-b"
          style={{ color: "var(--muted)", borderColor: "var(--line)" }}>
          <span className="col-span-6">Project</span>
          <span className="col-span-3">Status</span>
          <span className="col-span-3 text-right">Budget</span>
        </div>
        {projects.map((p, i) => (
          <button key={p.id} onClick={() => open(p.id)}
            className="w-full text-left grid grid-cols-2 md:grid-cols-12 gap-2 md:gap-3 items-center px-4 md:px-5 py-4 transition-colors hover:opacity-90"
            style={{ borderTop: i > 0 ? "1px solid var(--line)" : "none" }}>
            <div className="col-span-2 md:col-span-6 min-w-0">
              <div className="font-medium text-sm truncate">{p.name}</div>
              <div className="text-xs truncate" style={{ color: "var(--muted)" }}>{p.client}</div>
            </div>
            <div className="md:col-span-3"><StatusTag s={p.status} /></div>
            <div className="md:col-span-3 text-right mono text-sm">
              {p.budget ? money(p.budget) : <span style={{ color: "var(--muted)" }}>—</span>}
            </div>
          </button>
        ))}
      </Card>
      )}
    </div>
  );
}

function ProjectDetail({ project: p, quotes, onBack, copy, copied, onPreviewIntake }) {
  const [tab, setTab] = useState("overview");
  const [followUpsSent, setFollowUpsSent] = useState(false);
  const tabs = ["overview", "activity"];
  return (
    <div className="fade">
      <button onClick={onBack} className="mono text-[11px] uppercase tracking-wider mb-4 flex items-center gap-1.5" style={{ color: "var(--muted)" }}>
        <ArrowLeft className="w-3 h-3" /> Dashboard
      </button>
      <div className="flex flex-wrap items-center gap-3 mb-1">
        <h1 className="display text-2xl md:text-3xl font-semibold tracking-tight">{p.name}</h1>
        <StatusTag s={p.status} />
      </div>
      <p className="text-sm mb-5" style={{ color: "var(--muted)" }}>{p.client} · created {p.created}</p>

      <div className="flex flex-wrap items-center gap-2 mb-6">
        <button onClick={() => copy(p.link)} className="flex items-center gap-2 rounded-xl px-4 py-2.5 mono text-xs"
          style={{ background: "var(--accent-soft)", color: "var(--accent)" }}>
          {copied === p.link ? <Check className="w-3.5 h-3.5" /> : <Link2 className="w-3.5 h-3.5" />}
          {copied === p.link ? "Link copied" : p.link}
        </button>
        <Btn variant="secondary" className="px-4 py-2.5 text-xs" onClick={onPreviewIntake}><ExternalLink className="w-3.5 h-3.5" /> Preview client flow</Btn>
      </div>

      <div className="flex gap-1 overflow-x-auto border-b mb-6" style={{ borderColor: "var(--line)" }}>
        {tabs.map(t => (
          <button key={t} onClick={() => setTab(t)}
            className="px-4 py-2.5 text-sm font-medium capitalize whitespace-nowrap border-b-2 -mb-px transition-colors"
            style={tab === t ? { borderColor: "var(--accent)", color: "var(--ink)" } : { borderColor: "transparent", color: "var(--muted)" }}>
            {t}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <div className="space-y-4">
          <div className="grid grid-cols-3 gap-3 md:gap-4">
            {[["Value", projectValue(p, quotes) ? money(projectValue(p, quotes)) : "—"], ["Timeline", p.timelineMode === "date" ? (p.timelineDate || "—") : (p.timelineWeeks ? `${p.timelineWeeks} weeks` : "—")], ["Brief", p.briefComplete ? "Complete" : "Awaiting client"]].map(([k, v]) => (
              <Card key={k} className="p-4 md:p-5"><SectionLabel>{k}</SectionLabel><div className="display text-lg md:text-xl font-semibold">{v}</div></Card>
            ))}
          </div>

          {!p.brief ? (
            <EmptyState icon={Clock} title="Waiting on the client" body="They've opened the link but haven't finished. Answers autosave, so they can pick up where they left off." action={<Btn variant="secondary" onClick={onPreviewIntake}>Preview what they see</Btn>} />
          ) : (
            <>
              <Card className="p-6">
                <div className="flex items-center justify-between mb-4">
                  <SectionLabel>The brief</SectionLabel>
                  <Btn variant="secondary" className="px-3 py-1.5 text-xs"><PenLine className="w-3.5 h-3.5" /> Edit</Btn>
                </div>
                <div className="text-sm leading-relaxed whitespace-pre-line">{p.brief.professionalBrief}</div>
              </Card>

              <div className="flex items-center justify-between">
                <SectionLabel>AI checks — generated when the client submitted</SectionLabel>
                {p.brief.followUpQuestions?.length > 0 && (
                  <button onClick={() => setFollowUpsSent(true)}
                    className="inline-flex items-center gap-2 rounded-xl px-4 py-2 text-xs font-medium"
                    style={followUpsSent ? { background: "color-mix(in srgb, var(--good) 12%, transparent)", color: "var(--good)" } : { background: "var(--accent)", color: "var(--accent-ink)" }}>
                    {followUpsSent ? <><Check className="w-3.5 h-3.5" /> Follow-ups sent to client</> : <><Send className="w-3.5 h-3.5" /> Send follow-ups to client</>}
                  </button>
                )}
              </div>
              <div className="grid md:grid-cols-2 gap-4">
                <InsightCard icon={HelpCircle} tone="accent" title="Follow-up questions" items={p.brief.followUpQuestions} />
                <InsightCard icon={AlertTriangle} tone="warn" title="Missing information" items={p.brief.missingInfo} />
                <InsightCard icon={Search} tone="muted" title="Unclear requirements" items={p.brief.unclearRequirements} />
                <InsightCard icon={FolderKanban} tone="warn" title="Possible scope gaps" items={p.brief.scopeGaps} />
              </div>
            </>
          )}
        </div>
      )}

      {tab === "activity" && (
        <div className="space-y-0">
          {p.activity.map((ev, i) => (
            <div key={i} className="flex gap-4 pb-6 relative">
              <div className="flex flex-col items-center">
                <div className="w-2 h-2 rounded-full mt-1.5 shrink-0" style={{ background: "var(--accent)" }} />
                {i < p.activity.length - 1 && <div className="w-px flex-1" style={{ background: "var(--line)" }} />}
              </div>
              <div>
                <div className="text-sm">{ev.e}</div>
                <div className="mono text-[11px] mt-0.5" style={{ color: "var(--muted)" }}>{ev.t}</div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const InsightCard = ({ icon: Icon, tone, title, items }) => (
  <Card className="p-5">
    <div className="flex items-center gap-2 mb-3">
      <Icon className="w-4 h-4" style={{ color: tone === "warn" ? "var(--warn)" : tone === "accent" ? "var(--accent)" : "var(--muted)" }} />
      <span className="text-sm font-semibold">{title}</span>
      <Tag tone={items?.length ? tone : "good"}>{items?.length || 0}</Tag>
    </div>
    {items?.length ? (
      <ul className="space-y-2">
        {items.map((it, i) => <li key={i} className="text-sm leading-relaxed" style={{ color: "var(--muted)" }}>{it}</li>)}
      </ul>
    ) : <div className="text-sm" style={{ color: "var(--muted)" }}>Nothing flagged.</div>}
  </Card>
);

const EmptyState = ({ icon: Icon, title, body, action }) => (
  <Card className="p-10 text-center max-w-lg mx-auto" style={{ background: "var(--surface-2)" }}>
    <Icon className="w-6 h-6 mx-auto mb-4" style={{ color: "var(--muted)" }} />
    <div className="font-semibold mb-1.5">{title}</div>
    <p className="text-sm leading-relaxed mb-5" style={{ color: "var(--muted)" }}>{body}</p>
    {action}
  </Card>
);

function NewProjectModal({ onClose, onCreate }) {
  const [name, setName] = useState("");
  const [client, setClient] = useState("");
  const create = () => {
    const slug = Math.random().toString(36).slice(2, 6);
    onCreate({
      id: "p" + Date.now(), name: name || "Untitled project", client: client || "New client",
      status: "Awaiting brief", created: "6 Jul 2026", budget: null, timelineWeeks: null,
      link: `prelima.app/b/${slug}`, briefComplete: false, brief: null,
      activity: [{ t: "6 Jul", e: "Project created — intake link generated" }],
      files: [], quote: null, invoice: null,
    });
  };
  return (
    <div className="fixed inset-0 z-50 flex items-end md:items-center justify-center p-4" style={{ background: "rgba(0,0,0,.45)" }} onClick={onClose}>
      <Card className="w-full max-w-md p-6 fade" onClick={e => e.stopPropagation()} style={{ boxShadow: "var(--shadow-lg)" }}>
        <div className="flex items-center justify-between mb-5">
          <h2 className="display text-xl font-semibold">New project</h2>
          <button onClick={onClose} aria-label="Close"><X className="w-4 h-4" style={{ color: "var(--muted)" }} /></button>
        </div>
        <div className="space-y-3 mb-5">
          <input type="text" value={name} onChange={e => setName(e.target.value)} placeholder="Project name — e.g. Website copy refresh" className="w-full rounded-xl px-4 py-3 text-sm" autoFocus />
          <input type="text" value={client} onChange={e => setClient(e.target.value)} placeholder="Client or company name" className="w-full rounded-xl px-4 py-3 text-sm" />
        </div>
        <Btn onClick={create} className="w-full"><Link2 className="w-4 h-4" /> Create & generate intake link</Btn>
      </Card>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Quotations                                                          */
/* ------------------------------------------------------------------ */

function QuotationsBoard({ quotes, setQuotes, projects, openQuote, setOpenQuote, createQuote }) {
  const q = quotes.find(x => x.id === openQuote);
  const update = (nq) => setQuotes(qs => qs.map(x => x.id === nq.id ? nq : x));
  const remove = (id) => { setQuotes(qs => qs.filter(x => x.id !== id)); setOpenQuote(null); };

  if (q) return <QuoteEditor quote={q} update={update} remove={remove} onBack={() => setOpenQuote(null)} projects={projects} />;

  return (
    <div className="fade">
      <PageTitle title="Quotations" sub="Tag one to a project to keep things grouped."
        action={<Btn onClick={() => createQuote(null)} className="px-4 py-2"><Plus className="w-4 h-4" /> New quotation</Btn>} />
      {quotes.length === 0 ? (
        <EmptyState icon={FileText} title="No quotations yet" body="Create one and tag it to a project, or keep it standalone." action={<Btn onClick={() => createQuote(null)}><Plus className="w-4 h-4" /> New quotation</Btn>} />
      ) : (
        <div className="space-y-2">
          {quotes.map(qt => (
            <Card key={qt.id} className="p-4 flex items-center gap-4 cursor-pointer hover:opacity-90" onClick={() => setOpenQuote(qt.id)}>
              <span className="mono text-xs w-24 shrink-0" style={{ color: "var(--muted)" }}>{qt.number}</span>
              <div className="flex-1 min-w-0">
                <div className="text-sm font-medium truncate">{qt.title || "Untitled quotation"}</div>
                <div className="text-xs truncate" style={{ color: "var(--muted)" }}>{qt.client || "No client"}{qt.projectId ? " · tagged to project" : ""}{qt.sst ? ` · SST ${qt.sst}%` : ""}</div>
              </div>
              <span className="mono text-sm">{money(quoteTotal(qt))}</span>
              <StatusTag s={qt.status} />
              <ChevronRight className="w-4 h-4 shrink-0" style={{ color: "var(--muted)" }} />
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function CreativeBriefsBoard({ briefs, setBriefs, openBrief, setOpenBrief, onNew }) {
  const t = briefs.find(x => x.id === openBrief);
  const remove = (id) => { setBriefs(bs => bs.filter(x => x.id !== id)); setOpenBrief(null); };

  if (t) return <CreativeBriefViewer brief={t} onBack={() => setOpenBrief(null)} remove={remove} />;

  return (
    <div className="fade">
      <PageTitle title="Creative Briefs" sub="Quick briefs for designers, writers and other freelancers on your team."
        action={<Btn onClick={onNew} className="px-4 py-2"><Plus className="w-4 h-4" /> New brief</Btn>} />
      {briefs.length === 0 ? (
        <EmptyState icon={FileStack} title="No briefs yet" body="Create one to hand off a task to a designer or freelancer." action={<Btn onClick={onNew}><Plus className="w-4 h-4" /> New brief</Btn>} />
      ) : (
        <div className="space-y-2">
          {briefs.map(t => (
            <Card key={t.id} className="p-4 flex items-center gap-4 cursor-pointer hover:opacity-90" onClick={() => setOpenBrief(t.id)}>
              <div className="flex-1 min-w-0">
                <div className="text-sm font-medium truncate">{t.title || "Untitled brief"}</div>
                <div className="text-xs truncate" style={{ color: "var(--muted)" }}>
                  {(t.type || []).join(", ") || "—"}{t.deadline ? ` · due ${t.deadline}` : ""}
                </div>
              </div>
              <StatusTag s={t.status} />
              <ChevronRight className="w-4 h-4 shrink-0" style={{ color: "var(--muted)" }} />
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function CreativeBriefViewer({ brief: t, onBack, remove }) {
  const [copied, setCopied] = useState(false);
  const copyLink = () => { navigator.clipboard?.writeText(t.link); setCopied(true); setTimeout(() => setCopied(false), 1500); };
  const problemLabel = (t.problem || []).map(p => p === "Other" ? (t.problemOther || "Other") : p).join(", ");
  const deliverablesLabel = (t.deliverables || []).filter(d => d.format).map(d => `${d.qty}× ${d.format === "Other" ? (d.other || "Other") : d.format}`).join(", ");
  const locationLabel = (t.audienceLocation || []).map(l => l === "Other" ? (t.audienceLocationOther || "Other") : l).join(", ");
  const hobbyLabel = (t.audienceHobbies || []).map(h => h === "Other" ? (t.audienceHobbiesOther || "Other") : h).join(", ");
  const audienceSummary = [
    (t.audienceAgeRange || []).join(", "), (t.audienceGender || []).join("/"),
    locationLabel, hobbyLabel,
  ].filter(Boolean).join(" · ");
  return (
    <div className="fade max-w-2xl">
      <button onClick={onBack} className="mono text-[11px] uppercase tracking-wider mb-4 flex items-center gap-1.5" style={{ color: "var(--muted)" }}>
        <ArrowLeft className="w-3 h-3" /> Creative Briefs
      </button>
      <div className="flex flex-wrap items-center gap-3 mb-1">
        <h1 className="display text-2xl md:text-3xl font-semibold tracking-tight">{t.title}</h1>
        <StatusTag s={t.status} />
      </div>
      <p className="text-sm mb-5" style={{ color: "var(--muted)" }}>{(t.type || []).join(", ") || "—"} · created {t.created}{t.deadline ? ` · due ${t.deadline}` : ""}</p>

      <div className="flex flex-wrap items-center gap-2 mb-6">
        <button onClick={copyLink} className="flex items-center gap-2 rounded-xl px-4 py-2.5 mono text-xs"
          style={{ background: "var(--accent-soft)", color: "var(--accent)" }}>
          {copied ? <Check className="w-3.5 h-3.5" /> : <Link2 className="w-3.5 h-3.5" />}
          {copied ? "Link copied" : t.link}
        </button>
      </div>

      <div className="pr-print-area">
        <Card className="p-6 mb-4">
          <SectionLabel>The brief</SectionLabel>
          <div className="text-sm leading-relaxed whitespace-pre-line mt-2">{t.brief?.creativeBrief}</div>
        </Card>

        {t.brief?.deliverables?.length > 0 && (
          <Card className="p-6 mb-4">
            <SectionLabel>Deliverables</SectionLabel>
            <ul className="mt-2 space-y-1">
              {t.brief.deliverables.map((d, i) => <li key={i} className="text-sm flex gap-2"><span>•</span><span>{d}</span></li>)}
            </ul>
          </Card>
        )}

        {t.brief?.keyRequirements?.length > 0 && (
          <Card className="p-6 mb-4">
            <SectionLabel>Key requirements</SectionLabel>
            <ul className="mt-2 space-y-1">
              {t.brief.keyRequirements.map((d, i) => <li key={i} className="text-sm flex gap-2"><span>•</span><span>{d}</span></li>)}
            </ul>
          </Card>
        )}

        {(t.brandName || t.brandWebsite || deliverablesLabel || problemLabel || audienceSummary || t.insightAnswers?.some(Boolean)) && (
          <Card className="p-6 mb-4">
            <SectionLabel>Context</SectionLabel>
            {(t.brandName || t.brandWebsite) && <div className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>Brand: </span>{t.brandName || t.brandWebsite}</div>}
            {deliverablesLabel && <div className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>Deliverables: </span>{deliverablesLabel}</div>}
            {problemLabel && <div className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>Problem: </span>{problemLabel}</div>}
            {audienceSummary && <div className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>Audience: </span>{audienceSummary}</div>}
            {(t.insightQuestions || []).map((q, i) => (t.insightAnswers || [])[i] && (
              <div key={i} className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>{q} </span>{t.insightAnswers[i]}</div>
            ))}
          </Card>
        )}

        {(t.workingFiles || t.workingDeck || t.extraLinks) && (
          <Card className="p-6 mb-4">
            <SectionLabel>Files & links</SectionLabel>
            {t.workingFiles && <div className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>Working files: </span>{t.workingFiles}</div>}
            {t.workingDeck && <div className="text-sm mt-2"><span style={{ color: "var(--muted)" }}>Working deck: </span>{t.workingDeck}</div>}
            {t.extraLinks && <div className="text-sm mt-2 whitespace-pre-line"><span style={{ color: "var(--muted)" }}>Extra links: </span>{t.extraLinks}</div>}
          </Card>
        )}
      </div>

      <div className="flex flex-wrap items-center gap-2 mt-5">
        <Btn variant="secondary" onClick={downloadBrief}><Download className="w-4 h-4" /> Download PDF</Btn>
        <Btn variant="secondary" onClick={() => downloadTaskBriefDocx(t, t.brief, (t.type || []).join(", "))}><FileText className="w-4 h-4" /> Download Word doc</Btn>
        <button onClick={() => remove(t.id)} className="inline-flex items-center gap-1.5 text-sm ml-auto" style={{ color: "var(--muted)" }}>
          <Trash2 className="w-3.5 h-3.5" /> Delete
        </button>
      </div>
    </div>
  );
}

function LogoDrop({ logo, onLogo }) {
  const inputRef = useRef(null);
  const [over, setOver] = useState(false);
  const handle = (file) => {
    if (!file || !file.type.startsWith("image/")) return;
    const r = new FileReader();
    r.onload = () => onLogo(r.result);
    r.readAsDataURL(file);
  };
  return (
    <div
      onDragOver={e => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={e => { e.preventDefault(); setOver(false); handle(e.dataTransfer.files[0]); }}
      onClick={() => inputRef.current && inputRef.current.click()}
      className="relative w-32 h-24 rounded-2xl flex items-center justify-center cursor-pointer transition-colors shrink-0"
      style={{ border: `2px dashed ${over ? "var(--accent)" : "var(--line)"}`, background: over ? "var(--accent-soft)" : "var(--surface-2)" }}
      role="button" aria-label="Add your logo">
      {logo ? (
        <>
          <img src={logo} alt="Your logo" className="max-w-full max-h-full object-contain p-2 rounded-2xl" />
          <button onClick={e => { e.stopPropagation(); onLogo(null); }} aria-label="Remove logo"
            className="absolute -top-2 -right-2 w-6 h-6 rounded-full flex items-center justify-center"
            style={{ background: "var(--surface)", border: "1px solid var(--line)", boxShadow: "var(--shadow)" }}>
            <X className="w-3 h-3" style={{ color: "var(--muted)" }} />
          </button>
        </>
      ) : (
        <div className="text-center px-2">
          <Upload className="w-4 h-4 mx-auto mb-1" style={{ color: "var(--muted)" }} />
          <div className="text-[11px] leading-tight" style={{ color: "var(--muted)" }}>Drop your logo here</div>
        </div>
      )}
      <input ref={inputRef} type="file" accept="image/*" className="hidden" onChange={e => handle(e.target.files[0])} />
    </div>
  );
}

function QuoteEditor({ quote: q, update, remove, onBack, projects }) {
  const project = projects.find(p => p.id === q.projectId);
  const set = (patch) => update({ ...q, ...patch });
  const subtotal = quoteSubtotal(q);
  const total = quoteTotal(q);
  const sstAmount = total - subtotal;

  const isBlankItem = (it) => !it.title && !it.details && !it.unit;
  const [openItem, setOpenItem] = useState(() => (q.items.length === 1 && isBlankItem(q.items[0]) ? 0 : null));
  const [previewing, setPreviewing] = useState(false);
  const [openTerms, setOpenTerms] = useState(() => !q.paymentTerms);
  const bankFilled = (b) => b.name.trim() && b.account.trim() && b.holder.trim();
  const [openBank, setOpenBank] = useState(() => !bankFilled(q.bank));

  const setItem = (i, patch) => set({ items: q.items.map((it, j) => j === i ? { ...it, ...patch } : it) });
  const addItem = () => {
    const items = [...q.items, { title: "", details: "", qty: 1, unit: 0 }];
    set({ items });
    setOpenItem(items.length - 1);
  };
  const removeItem = (i) => { set({ items: q.items.filter((_, j) => j !== i) }); setOpenItem(null); };
  const tagProject = (pid) => {
    const p = projects.find(x => x.id === pid);
    set({ projectId: pid || null, client: q.client || (p ? p.client : ""), title: q.title || (p ? p.name : "") });
  };
  const setConsolidatedAmount = (v) => {
    const unit = Number(v) || 0;
    if (q.items.length === 0) set({ items: [{ title: q.consolidatedLabel || "Project fee", details: "", qty: 1, unit }] });
    else set({ items: [{ ...q.items[0], qty: 1, unit }, ...q.items.slice(1).map(it => ({ ...it, qty: 0 }))] });
  };

  const printOnMount = useRef(false);
  useEffect(() => {
    if (previewing && printOnMount.current) {
      printOnMount.current = false;
      const t = setTimeout(() => window.print(), 50);
      return () => clearTimeout(t);
    }
  }, [previewing]);

  const downloadQuote = () => {
    if (previewing) { window.print(); return; }
    printOnMount.current = true;
    setPreviewing(true);
  };

  if (previewing) {
    return (
      <div className="fade max-w-3xl">
        <button onClick={() => setPreviewing(false)} className="mono text-[11px] uppercase tracking-wider mb-4 flex items-center gap-1.5" style={{ color: "var(--muted)" }}>
          <ArrowLeft className="w-3 h-3" /> Back to edit
        </button>
        <Card className="p-6 md:p-10 pr-print-area">
          <div className="flex items-start justify-between gap-4 mb-8">
            <div className="min-w-0">
              {q.logo && <img src={q.logo} alt="Logo" className="h-12 mb-4 object-contain" />}
              <div className="display text-2xl md:text-3xl font-semibold tracking-tight">{q.title || "Untitled quotation"}</div>
              <div className="text-sm mt-1" style={{ color: "var(--muted)" }}>{q.client || "—"}</div>
            </div>
            <div className="text-right shrink-0">
              <div className="mono text-xs" style={{ color: "var(--muted)" }}>{q.number}</div>
              <div className="mono text-xs mt-0.5" style={{ color: "var(--muted)" }}>{q.created}</div>
              <div className="mt-2"><StatusTag s={q.status} /></div>
            </div>
          </div>

          {q.mode === "itemised" ? (
            <div className="space-y-4 mb-6">
              {q.items.map((it, i) => (
                <div key={i} className="flex justify-between items-start gap-4 pb-4" style={{ borderBottom: "1px solid var(--line)" }}>
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-medium">{it.title || "Untitled"}</div>
                    {it.details && <div className="text-xs mt-1 whitespace-pre-line leading-relaxed" style={{ color: "var(--muted)" }}>{it.details}</div>}
                  </div>
                  <div className="mono text-xs shrink-0 pt-0.5" style={{ color: "var(--muted)" }}>{it.qty || 0} × {money(it.unit || 0)}</div>
                  <div className="mono text-sm shrink-0 w-24 text-right">{money((it.qty || 0) * (it.unit || 0))}</div>
                </div>
              ))}
            </div>
          ) : (
            <div className="flex justify-between items-center pb-4 mb-6" style={{ borderBottom: "1px solid var(--line)" }}>
              <div className="text-sm font-medium">{q.consolidatedLabel || "Project fee"}</div>
              <div className="mono text-sm">{money(subtotal)}</div>
            </div>
          )}

          <div className="space-y-2 mb-8 max-w-xs ml-auto">
            <div className="flex justify-between text-sm"><span style={{ color: "var(--muted)" }}>Subtotal</span><span className="mono">{money(subtotal)}</span></div>
            {q.sst > 0 && <div className="flex justify-between text-sm"><span style={{ color: "var(--muted)" }}>SST {q.sst}%</span><span className="mono">{money(sstAmount)}</span></div>}
            <div className="flex justify-between pt-2 border-t" style={{ borderColor: "var(--line)" }}>
              <span className="font-semibold">Total</span><span className="display text-xl font-semibold">{money(total)}</span>
            </div>
          </div>

          <div>
            <SectionLabel>Payment</SectionLabel>
            {q.paymentTerms && <div className="text-sm mb-2" style={{ color: "var(--muted)" }}>{q.paymentTerms}</div>}
            <div className="text-sm">{q.paymentMethod}</div>
            <div className="text-sm" style={{ color: "var(--muted)" }}>{q.bank.name || "—"} · {q.bank.account || "—"} · {q.bank.holder || "—"}</div>
          </div>
        </Card>
        <div className="flex items-center gap-2 mt-5">
          <Btn onClick={downloadQuote}><Download className="w-4 h-4" /> Download PDF</Btn>
          <Btn variant="secondary" onClick={() => setPreviewing(false)}>Back to edit</Btn>
        </div>
      </div>
    );
  }

  return (
    <div className="fade max-w-3xl">
      <button onClick={onBack} className="mono text-[11px] uppercase tracking-wider mb-4 flex items-center gap-1.5" style={{ color: "var(--muted)" }}>
        <ArrowLeft className="w-3 h-3" /> Quotations
      </button>

      <div className="flex gap-4 items-start mb-4">
        <LogoDrop logo={q.logo} onLogo={(logo) => set({ logo })} />
        <div className="flex-1 min-w-0">
          <div className="flex flex-wrap items-center gap-3 mb-1">
            <span className="mono text-sm" style={{ color: "var(--muted)" }}>{q.number}</span>
            <StatusTag s={q.status} />
          </div>
          <input type="text" value={q.title} onChange={e => set({ title: e.target.value })} placeholder="Quotation title"
            className="display w-full text-2xl md:text-3xl font-semibold tracking-tight bg-transparent border-0 px-0 py-1" style={{ boxShadow: "none" }} />
          <input type="text" value={q.client} onChange={e => set({ client: e.target.value })} placeholder="Client or company name"
            className="w-full text-sm bg-transparent border-0 px-0" style={{ boxShadow: "none", color: "var(--muted)" }} />
        </div>
      </div>

      {/* Tag to a project */}
      <div className="flex items-center gap-3 mb-5">
        <SectionLabel>Tag to a brief</SectionLabel>
        <select value={q.projectId || ""} onChange={e => tagProject(e.target.value)}
          className="rounded-xl px-3 py-2 text-sm" style={{ background: "var(--surface)", border: "1px solid var(--line)", color: "var(--ink)" }}>
          <option value="">Not tagged</option>
          {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        {project && <Tag tone="accent">Grouped with {project.client}</Tag>}
      </div>

      {/* Billing */}
      <Card className="p-5 md:p-6">
        <div className="flex items-center justify-between mb-5">
          <SectionLabel>Billing</SectionLabel>
          <div className="flex rounded-xl p-0.5" style={{ background: "var(--surface-2)", border: "1px solid var(--line)" }}>
            {[["itemised", "Itemised"], ["consolidated", "One total"]].map(([m, label]) => (
              <button key={m} onClick={() => set({ mode: m })}
                className="px-3.5 py-1.5 rounded-[10px] text-xs font-medium transition-all"
                style={q.mode === m ? { background: "var(--surface)", boxShadow: "var(--shadow)", color: "var(--ink)" } : { color: "var(--muted)" }}>
                {label}
              </button>
            ))}
          </div>
        </div>

        {q.mode === "itemised" ? (
          <>
            <div className="space-y-2.5">
              {q.items.map((it, i) => (
                openItem === i ? (
                  <div key={i} className="rounded-2xl p-4 rise" style={{ background: "var(--surface-2)", border: "1px solid var(--accent)" }}
                    onBlur={e => { if (!e.currentTarget.contains(e.relatedTarget) && it.title.trim()) setOpenItem(null); }}>
                    <div className="flex gap-2 items-center mb-2.5">
                      <input type="text" value={it.title} onChange={e => setItem(i, { title: e.target.value })} autoFocus
                        placeholder="Field — e.g. Social media content" className="flex-1 rounded-xl px-4 py-3 text-sm font-medium min-w-0" />
                      <button onClick={() => setOpenItem(null)} aria-label="Collapse field" className="p-2 shrink-0">
                        <ChevronRight className="w-4 h-4" style={{ color: "var(--muted)", transform: "rotate(90deg)" }} />
                      </button>
                      <button onClick={() => removeItem(i)} aria-label="Remove item" className="p-2 shrink-0"><Trash2 className="w-4 h-4" style={{ color: "var(--muted)" }} /></button>
                    </div>
                    <textarea rows={2} value={it.details} onChange={e => setItem(i, { details: e.target.value })}
                      placeholder={"Items — one per line, shown as bullets:\n10x carousels (up to 5 frames)\n5x videos"}
                      className="w-full rounded-xl px-4 py-3 text-sm leading-relaxed resize-none mb-2.5" />
                    <div className="flex flex-wrap items-center gap-2">
                      <div className="flex items-center rounded-xl px-3" style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
                        <span className="mono text-[11px] mr-2" style={{ color: "var(--muted)" }}>QTY</span>
                        <input type="number" min="0" value={it.qty || ""} onChange={e => setItem(i, { qty: Number(e.target.value) })}
                          placeholder="1" className="w-14 py-2.5 text-sm mono bg-transparent border-0 text-right" style={{ boxShadow: "none" }} />
                      </div>
                      <span className="mono text-xs" style={{ color: "var(--muted)" }}>×</span>
                      <div className="flex items-center rounded-xl px-3" style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
                        <span className="mono text-[11px] mr-2" style={{ color: "var(--muted)" }}>RM</span>
                        <input type="number" min="0" value={it.unit || ""} onChange={e => setItem(i, { unit: Number(e.target.value) })}
                          placeholder="0" className="w-24 py-2.5 text-sm mono bg-transparent border-0 text-right" style={{ boxShadow: "none" }} />
                        <span className="mono text-[11px] ml-2" style={{ color: "var(--muted)" }}>/ unit</span>
                      </div>
                      <span className="mono text-sm ml-auto">{money((Number(it.qty) || 0) * (Number(it.unit) || 0))}</span>
                    </div>
                  </div>
                ) : (
                  <div key={i} role="button" tabIndex={0} onClick={() => setOpenItem(i)}
                    onKeyDown={e => { if (e.key === "Enter") setOpenItem(i); }}
                    className="rounded-2xl p-4 flex items-center gap-3 cursor-pointer transition-colors hover:opacity-90"
                    style={{ background: "var(--surface-2)", border: "1px solid var(--line)" }}>
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium truncate">{it.title || "Untitled field"}</div>
                      {it.details && <div className="text-xs truncate mt-0.5" style={{ color: "var(--muted)" }}>{it.details.split("\n")[0]}</div>}
                    </div>
                    <span className="mono text-xs shrink-0" style={{ color: "var(--muted)" }}>{it.qty || 0} × {money(it.unit || 0)}</span>
                    <span className="mono text-sm shrink-0 w-20 text-right">{money((Number(it.qty) || 0) * (Number(it.unit) || 0))}</span>
                    <span role="button" aria-label="Remove item" onClick={e => { e.stopPropagation(); removeItem(i); }} className="p-1 shrink-0">
                      <Trash2 className="w-3.5 h-3.5" style={{ color: "var(--muted)" }} />
                    </span>
                    <PenLine className="w-3.5 h-3.5 shrink-0" style={{ color: "var(--muted)" }} />
                  </div>
                )
              ))}
            </div>
            <button onClick={addItem} className="mt-4 inline-flex items-center gap-2 text-sm font-medium" style={{ color: "var(--accent)" }}>
              <Plus className="w-4 h-4" /> Add billing field
            </button>
          </>
        ) : (
          <>
            <input type="text" value={q.consolidatedLabel} onChange={e => set({ consolidatedLabel: e.target.value })}
              placeholder="What the client sees — e.g. Complete brand refresh package" className="w-full rounded-xl px-4 py-3 text-sm mb-3" />
            {q.items.filter(it => (it.qty || 0) * (it.unit || 0) > 0).length <= 1 ? (
              <div className="flex items-center rounded-xl px-3 w-44" style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
                <span className="mono text-xs mr-1" style={{ color: "var(--muted)" }}>RM</span>
                <input type="number" min="0" value={q.items[0] ? q.items[0].unit || "" : ""} onChange={e => setConsolidatedAmount(e.target.value)}
                  placeholder="0" className="flex-1 py-3 text-sm mono bg-transparent border-0 text-right" style={{ boxShadow: "none" }} />
              </div>
            ) : (
              <p className="text-sm" style={{ color: "var(--muted)" }}>Total rolls up from your {q.items.length} billing fields — they stay saved, the client just sees one line. Switch to Itemised to edit them.</p>
            )}
          </>
        )}

        {/* Totals + SST */}
        <div className="mt-6 pt-5 border-t space-y-2" style={{ borderColor: "var(--line)" }}>
          <div className="flex justify-between items-center text-sm">
            <span style={{ color: "var(--muted)" }}>Subtotal</span>
            <span className="mono">{money(subtotal)}</span>
          </div>
          <div className="flex justify-between items-center text-sm">
            <span className="flex items-center gap-2" style={{ color: "var(--muted)" }}>
              SST
              <select value={q.sst || 0} onChange={e => set({ sst: Number(e.target.value) })}
                className="rounded-lg px-2 py-1 text-xs" style={{ background: "var(--surface)", border: "1px solid var(--line)", color: "var(--ink)" }}>
                <option value={0}>None</option>
                <option value={6}>6%</option>
                <option value={9}>9%</option>
              </select>
            </span>
            <span className="mono">{q.sst ? money(sstAmount) : "—"}</span>
          </div>
          <div className="flex justify-between items-center pt-2">
            <span className="font-semibold">Total</span>
            <span className="display text-2xl font-semibold">{money(total)}</span>
          </div>
        </div>
      </Card>

      {/* Payment details */}
      <Card className="p-5 md:p-6 mt-4">
        <SectionLabel>Payment</SectionLabel>

        <div className="mt-2">
          {openTerms ? (
            <textarea rows={2} value={q.paymentTerms} onChange={e => set({ paymentTerms: e.target.value })} autoFocus
              onBlur={e => { if (!e.currentTarget.contains(e.relatedTarget) && q.paymentTerms.trim()) setOpenTerms(false); }}
              placeholder="Payment terms — e.g. 50% deposit to start, balance on delivery"
              className="w-full rounded-xl px-4 py-3 text-sm leading-relaxed resize-none" />
          ) : (
            <div role="button" tabIndex={0} onClick={() => setOpenTerms(true)}
              onKeyDown={e => { if (e.key === "Enter") setOpenTerms(true); }}
              className="rounded-xl px-4 py-3 text-sm flex items-center gap-2 cursor-pointer transition-colors hover:opacity-90"
              style={{ background: "var(--surface-2)", border: "1px solid var(--line)" }}>
              <span className="flex-1 truncate" style={{ color: q.paymentTerms ? "var(--ink)" : "var(--muted)" }}>{q.paymentTerms || "Add payment terms"}</span>
              <PenLine className="w-3.5 h-3.5 shrink-0" style={{ color: "var(--muted)" }} />
            </div>
          )}
        </div>

        <div className="mt-3">
          {openBank ? (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-2.5"
              onBlur={e => { if (!e.currentTarget.contains(e.relatedTarget) && bankFilled(q.bank)) setOpenBank(false); }}>
              <select value={q.paymentMethod} onChange={e => set({ paymentMethod: e.target.value })}
                className="rounded-xl px-4 py-3 text-sm" style={{ background: "var(--surface)", border: "1px solid var(--line)", color: "var(--ink)" }}>
                {["Bank transfer", "DuitNow", "Cheque", "Cash"].map(m => <option key={m} value={m}>{m}</option>)}
              </select>
              <input type="text" value={q.bank.name} onChange={e => set({ bank: { ...q.bank, name: e.target.value } })} placeholder="Bank name — e.g. Maybank" className="rounded-xl px-4 py-3 text-sm" />
              <input type="text" value={q.bank.account} onChange={e => set({ bank: { ...q.bank, account: e.target.value } })} placeholder="Account number" className="rounded-xl px-4 py-3 text-sm" />
              <input type="text" value={q.bank.holder} onChange={e => set({ bank: { ...q.bank, holder: e.target.value } })} placeholder="Account holder name" className="rounded-xl px-4 py-3 text-sm" />
            </div>
          ) : (
            <div role="button" tabIndex={0} onClick={() => setOpenBank(true)}
              onKeyDown={e => { if (e.key === "Enter") setOpenBank(true); }}
              className="rounded-xl px-4 py-3 text-sm flex items-center gap-2 cursor-pointer transition-colors hover:opacity-90"
              style={{ background: "var(--surface-2)", border: "1px solid var(--line)" }}>
              <span className="flex-1 truncate">{q.paymentMethod} · {q.bank.name || "—"} · {q.bank.account || "—"} · {q.bank.holder || "—"}</span>
              <PenLine className="w-3.5 h-3.5 shrink-0" style={{ color: "var(--muted)" }} />
            </div>
          )}
        </div>
      </Card>

      <div className="flex flex-wrap items-center gap-2 mt-5">
        {q.status === "Draft" && <Btn onClick={() => set({ status: "Sent" })} disabled={!total}><Send className="w-4 h-4" /> Send to client</Btn>}
        {q.status === "Sent" && <Btn onClick={() => set({ status: "Accepted" })}><Check className="w-4 h-4" /> Mark accepted</Btn>}
        <Btn variant="secondary" onClick={() => setPreviewing(true)}><Eye className="w-4 h-4" /> Preview</Btn>
        <Btn variant="secondary" onClick={downloadQuote}><Download className="w-4 h-4" /> Download PDF</Btn>
        <button onClick={() => remove(q.id)} className="inline-flex items-center gap-1.5 text-sm ml-auto" style={{ color: "var(--muted)" }}>
          <Trash2 className="w-3.5 h-3.5" /> Delete
        </button>
      </div>
    </div>
  );
}

function PaymentsPage() {
  return (
    <div className="fade">
      <PageTitle title="Payments" sub="Coming soon — the data model is ready for it." />
      <div className="grid md:grid-cols-3 gap-4">
        {[["Payment links", "Send a link with any invoice and get paid by card or bank transfer."],
          ["Milestone payments", "Split a project into stages, each with its own release."],
          ["Instalments", "Let clients pay larger projects across scheduled dates."]].map(([t, d]) => (
          <Card key={t} className="p-5" style={{ background: "var(--surface-2)" }}>
            <Tag tone="accent">Planned</Tag>
            <div className="font-semibold mt-3 mb-1.5">{t}</div>
            <div className="text-sm leading-relaxed" style={{ color: "var(--muted)" }}>{d}</div>
          </Card>
        ))}
      </div>
    </div>
  );
}

function SettingsPage({ dark, setDark, wsName, setWsName }) {
  const [name, setName] = useState(wsName);
  const [savedName, setSavedName] = useState(false);
  const [pw, setPw] = useState({ current: "", next: "", confirm: "" });
  const [pwNote, setPwNote] = useState("");

  const saveName = () => {
    setWsName(name.trim() || wsName);
    setSavedName(true); setTimeout(() => setSavedName(false), 1800);
  };
  const savePw = () => {
    if (!pw.current || !pw.next) { setPwNote("Fill in your current and new password."); return; }
    if (pw.next.length < 8) { setPwNote("New password needs at least 8 characters."); return; }
    if (pw.next !== pw.confirm) { setPwNote("New passwords don't match."); return; }
    setPw({ current: "", next: "", confirm: "" });
    setPwNote("done");
  };

  return (
    <div className="fade max-w-lg">
      <PageTitle title="Settings" />
      <div className="space-y-3">
        <Card className="p-5">
          <div className="font-medium mb-1">Workspace name</div>
          <div className="text-sm mb-4" style={{ color: "var(--muted)" }}>Shown to clients on your intake links.</div>
          <div className="flex gap-2">
            <input type="text" value={name} onChange={e => setName(e.target.value)} className="flex-1 rounded-xl px-4 py-3 text-sm" />
            <Btn onClick={saveName} className="px-4 py-2 text-xs">{savedName ? <><Check className="w-3.5 h-3.5" /> Saved</> : "Save changes"}</Btn>
          </div>
        </Card>

        <Card className="p-5 flex items-center justify-between">
          <div>
            <div className="font-medium">Dark mode</div>
            <div className="text-sm" style={{ color: "var(--muted)" }}>Applies across the app and intake links.</div>
          </div>
          <button onClick={() => setDark(!dark)} className="relative w-12 h-7 rounded-full transition-colors" style={{ background: dark ? "var(--accent)" : "var(--line)" }} aria-label="Toggle dark mode">
            <span className="absolute top-1 w-5 h-5 rounded-full transition-all" style={{ background: "var(--surface)", left: dark ? "26px" : "4px", boxShadow: "var(--shadow)" }} />
          </button>
        </Card>

        <Card className="p-5">
          <div className="font-medium mb-1">Change password</div>
          <div className="text-sm mb-4" style={{ color: "var(--muted)" }}>At least 8 characters. In production this runs through Supabase Auth.</div>
          <div className="space-y-2.5">
            <input type="password" value={pw.current} onChange={e => { setPw({ ...pw, current: e.target.value }); setPwNote(""); }} placeholder="Current password" className="w-full rounded-xl px-4 py-3 text-sm" />
            <input type="password" value={pw.next} onChange={e => { setPw({ ...pw, next: e.target.value }); setPwNote(""); }} placeholder="New password" className="w-full rounded-xl px-4 py-3 text-sm" />
            <input type="password" value={pw.confirm} onChange={e => { setPw({ ...pw, confirm: e.target.value }); setPwNote(""); }} placeholder="Confirm new password" className="w-full rounded-xl px-4 py-3 text-sm" />
          </div>
          <div className="mt-4 flex items-center gap-3">
            <Btn onClick={savePw} className="px-4 py-2 text-xs">Update password</Btn>
            {pwNote === "done"
              ? <span className="text-sm flex items-center gap-1.5" style={{ color: "var(--good)" }}><Check className="w-3.5 h-3.5" /> Password updated</span>
              : pwNote && <span className="text-sm" style={{ color: "var(--warn)" }}>{pwNote}</span>}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Auth                                                                */
/* ------------------------------------------------------------------ */

function AuthScreen({ onDone, onBack, dark, setDark }) {
  const [mode, setMode] = useState("signin"); // signin | signup
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [note, setNote] = useState("");

  const submit = async () => {
    if (!email.trim() || !password) { setError("Enter your email and password."); return; }
    setBusy(true); setError(""); setNote("");
    try {
      if (mode === "signup") {
        const { data, error: err } = await supabase.auth.signUp({ email: email.trim(), password });
        if (err) throw err;
        if (data.session) { onDone(); return; }
        setNote("Check your email to confirm your account, then sign in.");
        setMode("signin");
      } else {
        const { error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
        if (err) throw err;
        onDone();
      }
    } catch (err) {
      setError(err.message || "Something went wrong.");
    }
    setBusy(false);
  };

  if (!supabaseConfigured) {
    return (
      <div className="min-h-screen flex items-center justify-center px-5" style={{ background: "var(--bg)" }}>
        <Card className="max-w-sm p-6 text-center">
          <AlertCircle className="w-6 h-6 mx-auto mb-3" style={{ color: "var(--warn)" }} />
          <p className="text-sm mb-4" style={{ color: "var(--muted)" }}>Accounts aren't set up yet — the app isn't connected to a database in this environment.</p>
          <Btn variant="secondary" onClick={onBack}>Back</Btn>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "var(--bg)" }}>
      <header className="flex items-center justify-between px-5 md:px-12 py-5 max-w-5xl mx-auto w-full">
        <button onClick={onBack}><Wordmark /></button>
        <button onClick={() => setDark(!dark)} aria-label="Toggle dark mode" className="p-2 rounded-lg" style={{ color: "var(--muted)" }}>
          {dark ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
        </button>
      </header>
      <main className="flex-1 flex items-center justify-center px-5 py-12">
        <Card className="w-full max-w-sm p-6 md:p-8 fade">
          <h1 className="display text-2xl font-semibold mb-1">{mode === "signup" ? "Create your account" : "Welcome back"}</h1>
          <p className="text-sm mb-6" style={{ color: "var(--muted)" }}>
            {mode === "signup" ? "Free for life — no card required." : "Sign in to your workspace."}
          </p>
          <div className="space-y-3">
            <div className="flex items-center gap-2 rounded-xl px-4" style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
              <Mail className="w-4 h-4 shrink-0" style={{ color: "var(--muted)" }} />
              <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="you@studio.com"
                className="flex-1 py-3 text-sm bg-transparent border-0" style={{ boxShadow: "none" }} autoFocus />
            </div>
            <div className="flex items-center gap-2 rounded-xl px-4" style={{ background: "var(--surface)", border: "1px solid var(--line)" }}>
              <Lock className="w-4 h-4 shrink-0" style={{ color: "var(--muted)" }} />
              <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="Password"
                onKeyDown={e => { if (e.key === "Enter") submit(); }}
                className="flex-1 py-3 text-sm bg-transparent border-0" style={{ boxShadow: "none" }} />
            </div>
          </div>
          {error && <p className="text-sm mt-3 flex items-start gap-2" style={{ color: "var(--warn)" }}><AlertCircle className="w-4 h-4 shrink-0 mt-0.5" /> {error}</p>}
          {note && <p className="text-sm mt-3" style={{ color: "var(--good)" }}>{note}</p>}
          <Btn onClick={submit} disabled={busy} className="w-full mt-5">
            {busy ? <Loader2 className="w-4 h-4 animate-spin" /> : null}
            {mode === "signup" ? "Create account" : "Sign in"}
          </Btn>
          <button onClick={() => { setMode(mode === "signup" ? "signin" : "signup"); setError(""); setNote(""); }}
            className="text-sm mt-4 underline underline-offset-4 block mx-auto" style={{ color: "var(--muted)" }}>
            {mode === "signup" ? "Already have an account? Sign in" : "New here? Create an account"}
          </button>
        </Card>
      </main>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Root                                                                */
/* ------------------------------------------------------------------ */

export default function App() {
  const [view, setView] = useState("landing"); // landing | auth | app | intake | taskBrief
  const [intakeReturn, setIntakeReturn] = useState("landing");
  const [previewProjectName, setPreviewProjectName] = useState("New project");
  const [intakeIsPreview, setIntakeIsPreview] = useState(false);
  const startIntake = (from, name, isPreview = false) => {
    setIntakeReturn(from); if (name) setPreviewProjectName(name); setIntakeIsPreview(isPreview); setView("intake");
  };
  const [dark, setDark] = useState(false);

  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [projects, setProjects] = useState([]);
  const [quotes, setQuotes] = useState([]);
  const [taskBriefs, setTaskBriefs] = useState([]);
  const [wsName, setWsName] = useState("My Studio");
  const loadedRef = useRef(false);

  useEffect(() => {
    if (!supabaseConfigured) { setAuthLoading(false); return; }
    supabase.auth.getSession().then(({ data }) => { setSession(data.session); setAuthLoading(false); });
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session?.user?.id) { loadedRef.current = false; setProjects([]); setQuotes([]); setTaskBriefs([]); return; }
    let cancelled = false;
    loadedRef.current = false;
    (async () => {
      try {
        const [p, q, tb, name] = await Promise.all([
          fetchProjects(session.user.id),
          fetchQuotes(session.user.id),
          fetchTaskBriefs(session.user.id),
          ensureProfile(session.user.id, "My Studio"),
        ]);
        if (cancelled) return;
        setProjects(p);
        setQuotes(q);
        setTaskBriefs(tb);
        setWsName(name);
      } finally {
        if (!cancelled) loadedRef.current = true;
      }
    })();
    return () => { cancelled = true; };
  }, [session?.user?.id]);

  useEffect(() => {
    if (!loadedRef.current || !session?.user?.id) return;
    syncProjects(session.user.id, projects).catch(() => {});
  }, [projects, session?.user?.id]);

  useEffect(() => {
    if (!loadedRef.current || !session?.user?.id) return;
    syncQuotes(session.user.id, quotes).catch(() => {});
  }, [quotes, session?.user?.id]);

  useEffect(() => {
    if (!loadedRef.current || !session?.user?.id) return;
    syncTaskBriefs(session.user.id, taskBriefs).catch(() => {});
  }, [taskBriefs, session?.user?.id]);

  useEffect(() => {
    if (view === "app" && !authLoading && !session) setView("landing");
  }, [view, authLoading, session]);

  const setWsNamePersisted = (name) => {
    setWsName(name);
    if (session?.user?.id) saveWorkspaceName(session.user.id, name).catch(() => {});
  };

  const handleIntakeDone = (brief, payload) => {
    // "Preview client flow" is just the freelancer looking at their own intake link —
    // it shouldn't create or touch a real project.
    if (intakeIsPreview || !session?.user?.id) return;
    const slug = Math.random().toString(36).slice(2, 6);
    const newProject = {
      id: "p" + Date.now(),
      name: payload.overview ? payload.overview.slice(0, 60) : (previewProjectName || "New project"),
      client: "New client",
      status: "Brief received",
      created: new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" }),
      budget: payload.budget, timelineMode: payload.timelineMode, timelineWeeks: payload.timelineWeeks, timelineDate: payload.timelineDate,
      link: `prelima.app/b/${slug}`, briefComplete: true, brief,
      activity: [
        { t: "now", e: "AI brief generated and structured" },
        { t: "now", e: "Brief submitted by client" },
      ],
      files: [], quote: null, invoice: null,
    };
    setProjects(ps => [newProject, ...ps]);
    setIntakeReturn("app");
  };

  const handleTaskBriefDone = (newBrief) => {
    setTaskBriefs(tb => [newBrief, ...tb]);
  };

  if (authLoading) {
    return (
      <div data-app="prelima" data-theme={dark ? "dark" : "light"} className="min-h-screen flex items-center justify-center antialiased" style={{ background: "var(--bg)" }}>
        <ThemeStyles />
        <Loader2 className="w-6 h-6 animate-spin" style={{ color: "var(--accent)" }} />
      </div>
    );
  }

  return (
    <div data-app="prelima" data-theme={dark ? "dark" : "light"} className="min-h-screen antialiased" style={{ background: "var(--bg)" }}>
      <ThemeStyles />
      {view === "landing" && <Landing
        onStart={() => startIntake("landing")}
        onStartTaskBrief={() => setView("taskBrief")}
        dark={dark} setDark={setDark} />}
      {view === "auth" && <AuthScreen onDone={() => setView("app")} onBack={() => setView("landing")} dark={dark} setDark={setDark} />}
      {view === "app" && session && <AppShell projects={projects} setProjects={setProjects} quotes={quotes} setQuotes={setQuotes} taskBriefs={taskBriefs} setTaskBriefs={setTaskBriefs} onNewTaskBrief={() => setView("taskBrief")} wsName={wsName} setWsName={setWsNamePersisted} onLogout={() => { supabase.auth.signOut(); setView("landing"); }} onPreviewIntake={(name) => startIntake("app", name, true)} dark={dark} setDark={setDark} />}
      {view === "intake" && <IntakeFlow projectName={previewProjectName} freelancer={wsName} onDone={handleIntakeDone} onExit={() => setView(intakeReturn)} />}
      {view === "taskBrief" && <TaskBriefFlow freelancer={wsName} onDone={handleTaskBriefDone} onExit={() => setView("app")} />}
    </div>
  );
}
