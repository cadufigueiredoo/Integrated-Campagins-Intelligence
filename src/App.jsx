import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid,
  PieChart, Pie, Cell, Legend, ReferenceLine,
} from "recharts";
import {
  Sparkles, Globe, BarChart3, Target, Plus, Save, Trash2, Download, Sun, Moon,
  X, ChevronRight, Upload, FileSpreadsheet, TrendingUp, TrendingDown, Layers,
  Users, Megaphone, Calendar, AlertTriangle, CheckCircle2, Menu, ExternalLink,
  Zap, Rocket, ArrowRight, FileText, Languages,
} from "lucide-react";

/* ============================================================================
   Integrated Campaign Intelligence
   Developed by Carlos Eduardo - https://www.linkedin.com/in/carloseduardovf/

   Architecture:
   - callAI() posts to /api/generate (Vercel serverless). The ANTHROPIC_API_KEY
     never reaches the browser.
   - Campaigns persist in localStorage under the "ici:campaign:" prefix.
   Design principle: the app COMPUTES every number. The model only writes prose.
   ========================================================================== */

/* ---------- Design tokens (iOS neutrals + Braze-inspired gradient) -------- */
const gradient = "linear-gradient(135deg, #6E3AFF 0%, #A03BFF 45%, #FF5C39 100%)";
const tokens = (dark) => ({
  dark,
  grad: gradient,
  violet: "#6E3AFF",
  coral: "#FF5C39",
  bg: dark ? "#000000" : "#F5F5F7",
  bgAtmo: dark
    ? "radial-gradient(1200px 600px at 80% -10%, rgba(110,58,255,0.18), transparent 60%), radial-gradient(900px 500px at -10% 10%, rgba(255,92,57,0.10), transparent 55%), #000000"
    : "radial-gradient(1200px 600px at 80% -10%, rgba(110,58,255,0.10), transparent 60%), radial-gradient(900px 500px at -10% 10%, rgba(255,92,57,0.07), transparent 55%), #F5F5F7",
  card: dark ? "#1C1C1E" : "#FFFFFF",
  cardEl: dark ? "#2C2C2E" : "#FBFBFD",
  ink: dark ? "#F5F5F7" : "#1D1D1F",
  muted: dark ? "#98989F" : "#6E6E73",
  faint: dark ? "#636366" : "#A1A1A6",
  hair: dark ? "#38383A" : "#E5E5EA",
  frost: dark ? "rgba(28,28,30,0.72)" : "rgba(255,255,255,0.72)",
  chip: dark ? "#2C2C2E" : "#F0F0F3",
  good: "#30D158",
  warn: "#FF9F0A",
  bad: "#FF453A",
  shadow: dark
    ? "0 1px 2px rgba(0,0,0,0.6), 0 12px 32px rgba(0,0,0,0.55)"
    : "0 1px 2px rgba(17,17,26,0.05), 0 8px 24px rgba(17,17,26,0.06), 0 16px 56px rgba(17,17,26,0.05)",
  font: '-apple-system, BlinkMacSystemFont, "SF Pro Display", "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
});

const FUNNEL_STAGES = ["Awareness", "Consideration", "Decision"];
/* Normalizes imported stage labels so charts and filters stay consistent. */
const canonicalStage = (v) => {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return "";
  const hit = FUNNEL_STAGES.find((x) => x.toLowerCase() === s);
  if (hit) return hit;
  if (s.startsWith("aware") || s.startsWith("topo") || s === "tofu") return "Awareness";
  if (s.startsWith("consider") || s.startsWith("meio") || s === "mofu") return "Consideration";
  if (s.startsWith("decis") || s.startsWith("fundo") || s === "bofu") return "Decision";
  return String(v).trim();
};
const TEMPLATE_COLUMNS = [
  "Channel", "Funnel Stage", "Spend", "Impressions", "Clicks",
  "Leads", "MQL", "SQL", "Opportunities", "Pipeline", "Revenue",
];

/* ---------------------------------- utils --------------------------------- */
const uid = () => Math.random().toString(36).slice(2, 10);
/* Locale-aware number parsing. Spreadsheets exported in Brazil arrive as
   "1.710.000,00" or "R$ 96.000"; a naive parser turns those into 1.71 and 96,
   silently corrupting every downstream metric. */
const num = (v) => {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  let s = String(v == null ? "" : v).trim();
  if (!s) return 0;
  const negative = /^\(.*\)$/.test(s) || s.indexOf("-") === 0;
  s = s.replace(/[^\d.,]/g, "");
  if (!s) return 0;

  const lastDot = s.lastIndexOf(".");
  const lastComma = s.lastIndexOf(",");
  const looksLikeThousands = (parts) =>
    parts.length > 1 &&
    parts[0].length > 0 && parts[0].length <= 3 && parts[0] !== "0" &&
    parts.slice(1).every((g) => g.length === 3);

  if (lastDot > -1 && lastComma > -1) {
    // Both separators: the rightmost one is the decimal separator.
    if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (lastComma > -1) {
    const parts = s.split(",");
    s = looksLikeThousands(parts) ? s.replace(/,/g, "") : s.replace(",", ".");
  } else if (lastDot > -1) {
    const parts = s.split(".");
    if (looksLikeThousands(parts)) s = s.replace(/\./g, "");
  }

  const n = parseFloat(s);
  if (!Number.isFinite(n)) return 0;
  return negative ? -Math.abs(n) : n;
};
const brl = (v) =>
  "R$ " + num(v).toLocaleString("pt-BR", { maximumFractionDigits: 0 });
const int = (v) => num(v).toLocaleString("pt-BR", { maximumFractionDigits: 0 });
const pct = (v, digits = 1) =>
  (num(v) * 100).toLocaleString("pt-BR", { maximumFractionDigits: digits }) + "%";
const money2 = (v) =>
  "R$ " + num(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* --------------------------- metrics (deterministic) ---------------------- */
function computeDerived(rows) {
  const byChannel = (rows || []).map((r) => {
    const spend = num(r.spend), imp = num(r.impressions), clk = num(r.clicks);
    const leads = num(r.leads), mql = num(r.mql), sql = num(r.sql);
    const opps = num(r.opps), pipe = num(r.pipeline), rev = num(r.revenue);
    return {
      channel: r.channel || "—",
      stage: r.stage || "—",
      spend, impressions: imp, clicks: clk, leads, mql, sql, opps,
      pipeline: pipe, revenue: rev,
      ctr: imp ? clk / imp : 0,
      cpl: leads ? spend / leads : 0,
      cpo: opps ? spend / opps : 0,
      leadToMql: leads ? mql / leads : 0,
      mqlToSql: mql ? sql / mql : 0,
      roas: spend ? rev / spend : 0,
    };
  });
  const sum = (k) => byChannel.reduce((a, r) => a + (r[k] || 0), 0);
  const t = {
    spend: sum("spend"), impressions: sum("impressions"), clicks: sum("clicks"),
    leads: sum("leads"), mql: sum("mql"), sql: sum("sql"), opps: sum("opps"),
    pipeline: sum("pipeline"), revenue: sum("revenue"),
  };
  t.ctr = t.impressions ? t.clicks / t.impressions : 0;
  t.cpl = t.leads ? t.spend / t.leads : 0;
  t.cpo = t.opps ? t.spend / t.opps : 0;
  t.leadToMql = t.leads ? t.mql / t.leads : 0;
  t.mqlToSql = t.mql ? t.sql / t.mql : 0;
  t.roas = t.spend ? t.revenue / t.spend : 0;
  t.roi = t.spend ? (t.revenue - t.spend) / t.spend : 0;
  return { byChannel, totals: t };
}

/* Priority is language-agnostic: the model may return Alta/Média/Baixa or
   High/Medium/Low, so both are normalized to one visual scale. */
function priorityLevel(p) {
  const s = String(p || "").trim().toLowerCase();
  if (s.startsWith("alta") || s.startsWith("high") || s.startsWith("crit")) return "high";
  if (s.startsWith("m")) return "medium"; // média / medium / mid
  if (s.startsWith("baix") || s.startsWith("low")) return "low";
  return "";
}
function priorityTone(p, T) {
  const lvl = priorityLevel(p);
  if (lvl === "high") return T.good;
  if (lvl === "medium") return T.warn;
  if (lvl === "low") return T.bad;
  return T.faint;
}

const METRIC_META = {
  leads: { label: "Leads", fmt: int, higher: true },
  mql: { label: "MQL", fmt: int, higher: true },
  sql: { label: "SQL", fmt: int, higher: true },
  pipeline: { label: "Pipeline", fmt: brl, higher: true },
  revenue: { label: "Receita", fmt: brl, higher: true },
  cpl: { label: "CPL", fmt: money2, higher: false },
};
function planVsActual(kpiTargets, totals) {
  if (!kpiTargets || !totals) return [];
  return kpiTargets
    .filter((k) => METRIC_META[k.metric] && k.target !== "" && k.target !== null && k.target !== undefined && num(k.target) > 0)
    .map((k) => {
      const meta = METRIC_META[k.metric];
      const target = num(k.target);
      const actual = num(totals[k.metric]);
      const diff = actual - target;
      const gapPct = target ? diff / target : 0;
      const favorable = meta.higher ? diff >= 0 : diff <= 0;
      return { metric: k.metric, label: meta.label, fmt: meta.fmt, target, actual, diff, gapPct, favorable, higher: meta.higher };
    });
}

/* ------------------------------- AI helpers ------------------------------- */
function extractJSON(text) {
  if (!text) return null;
  let s = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  const first = s.indexOf("{");
  if (first === -1) return null;
  const last = s.lastIndexOf("}");
  const body = last > first ? s.slice(first, last + 1) : s.slice(first);

  const attempts = [
    body,
    body.replace(/,\s*([}\]])/g, "$1"),           // trailing commas
    repairTruncated(body),                        // response cut off mid-object
    repairTruncated(body.replace(/,\s*([}\]])/g, "$1")),
  ];
  for (const candidate of attempts) {
    if (!candidate) continue;
    try { return JSON.parse(candidate); } catch (e) { /* try next */ }
  }
  return null;
}

/* Closes brackets left open when a response is truncated, so a partially
   received section can still be used instead of being silently discarded. */
function repairTruncated(str) {
  if (!str) return null;
  let out = "";
  let inStr = false, escaped = false;
  const stack = [];
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    out += ch;
    if (inStr) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === "{" || ch === "[") stack.push(ch);
    else if (ch === "}" || ch === "]") stack.pop();
  }
  if (inStr) out += '"';
  out = out.replace(/,\s*$/, "");
  while (stack.length) out += stack.pop() === "{" ? "}" : "]";
  return out;
}

async function callAI(system, user) {
  const res = await fetch("/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ system, user }),
  });
  if (!res.ok) {
    if (res.status === 404) {
      throw new Error(
        "The /api/generate endpoint was not found. This build is made for Vercel: deploy it, or run 'vercel dev' locally instead of 'npm run dev'."
      );
    }
    let msg = "Request failed (" + res.status + ")";
    try {
      const j = await res.json();
      if (j && j.error) msg = j.error;
    } catch (e) { /* keep default */ }
    throw new Error(msg);
  }
  const data = await res.json();
  return (data && data.text) || "";
}

/* Translates generated CONTENT only. Keys, numbers, metric ids and funnel
   stage names are preserved so downstream filters and math never break. */
async function translateSection(obj, targetLang) {
  const langName = targetLang === "EN" ? "English" : "Brazilian Portuguese";
  const sys =
    "You are a professional B2B marketing translator. Return ONLY valid minified JSON, no prose, no markdown. " +
    "Translate every string VALUE into " + langName + ". " +
    "Never translate or rename JSON keys. Never alter numbers or the structure. " +
    "Keep these tokens unchanged: Awareness, Consideration, Decision, MQL, SQL, CPL, CPO, CTR, ROAS, LGPD, LATAM. " +
    "Preserve the exact same arrays and object shapes.";
  const user = "Translate this JSON:\n" + JSON.stringify(obj);
  const text = await callAI(sys, user);
  return extractJSON(text);
}

/* --------------------------- storage (persistent) ------------------------- */
/* Persistence degrades gracefully. Where localStorage is unavailable or full
   (private browsing, quota exceeded), campaigns are kept in memory for the
   session instead of failing, and the app says so plainly. */
const LS_PREFIX = "ici:campaign:";
const memoryStore = new Map();
let durableStorage = null; // null = not probed yet

function probeStorage() {
  if (durableStorage !== null) return durableStorage;
  try {
    localStorage.setItem("ici:probe", "1");
    localStorage.removeItem("ici:probe");
    durableStorage = true;
  } catch (e) {
    durableStorage = false;
  }
  return durableStorage;
}

async function loadCampaigns() {
  if (probeStorage()) {
    try {
      const out = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (!k || k.indexOf(LS_PREFIX) !== 0) continue;
        try {
          const raw = localStorage.getItem(k);
          if (raw) out.push(JSON.parse(raw));
        } catch (e) { /* skip corrupted entry */ }
      }
      return out.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    } catch (e) { /* fall through to memory */ }
  }
  return Array.from(memoryStore.values()).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
}

async function persistCampaign(c) {
  if (probeStorage()) {
    try {
      localStorage.setItem(LS_PREFIX + c.id, JSON.stringify(c));
      return "durable";
    } catch (e) {
      durableStorage = false; // quota exceeded mid-session
    }
  }
  memoryStore.set(c.id, c);
  return "session";
}

async function removeCampaign(id) {
  memoryStore.delete(id);
  if (probeStorage()) {
    try { localStorage.removeItem(LS_PREFIX + id); } catch (e) { /* ignore */ }
  }
}

/* ---------------------- demo data (offline, bilingual) --------------------
   The sample campaign ships with both language versions hardcoded, so the
   PT/EN switch works instantly with no API call. Real campaigns are never
   affected by this: they always go through generation and translation. */
const DEMO_ROWS = [
  { channel: "LinkedIn Ads", stage: "Awareness", impressions: 1420000, clicks: 12800, spend: 96000, leads: 640, mql: 310, sql: 120, opps: 38, pipeline: 1710000, revenue: 285000 },
  { channel: "Paid Search", stage: "Decision", impressions: 380000, clicks: 21200, spend: 74000, leads: 980, mql: 520, sql: 210, opps: 66, pipeline: 2970000, revenue: 620000 },
  { channel: "Webinar", stage: "Consideration", impressions: 96000, clicks: 8400, spend: 41000, leads: 720, mql: 430, sql: 180, opps: 52, pipeline: 2340000, revenue: 430000 },
  { channel: "Content Syndication", stage: "Awareness", impressions: 540000, clicks: 6100, spend: 52000, leads: 810, mql: 240, sql: 60, opps: 14, pipeline: 630000, revenue: 90000 },
  { channel: "WhatsApp Nurture", stage: "Consideration", impressions: 210000, clicks: 15400, spend: 18000, leads: 360, mql: 250, sql: 140, opps: 44, pipeline: 1980000, revenue: 410000 },
];

const DEMO_KPI_TARGETS = [
  { metric: "leads", target: 3000 }, { metric: "mql", target: 1600 },
  { metric: "sql", target: 800 }, { metric: "pipeline", target: 12000000 },
  { metric: "revenue", target: 2000000 }, { metric: "cpl", target: 85 },
];

const DEMO_TEXT = {
  PT: {
    profile: {
      offering: "Plataforma de engajamento e ativação de clientes (CDP + orquestração cross-channel).",
      category: "Customer Engagement / MarTech",
      audience: "Times de CRM, Growth e Lifecycle em empresas B2C e B2B2C de médio e grande porte.",
    },
    objective: "Gerar pipeline novo e upsell em contas enterprise de LATAM com uma campanha âncora multicanal.",
    plan: {
      thesis: "Marcas de LATAM investiram em coleta de dados, mas ativam pouco. A campanha âncora posiciona a ativação em tempo real como o próximo salto de ROI, conectando dado a interação 1:1 sem depender de mais headcount.",
      whyNow: [
        "Pressão por eficiência de marketing eleva a régua de ROI por canal em 2026.",
        "Maturidade de CDP na região cria base instalada pronta para ativação avançada.",
        "Restrições de cookies e custo de mídia empurram marcas para canais próprios.",
      ],
      personas: [
        { name: "Head de CRM / Lifecycle", funnelStage: "Consideration", pain: "Dados ricos, ativação lenta e manual.", message: "Transforme dado parado em interação 1:1 sem novo headcount." },
        { name: "Diretor de Growth", funnelStage: "Decision", pain: "CAC subindo, retenção pressionando a meta.", message: "Aumente LTV com jornadas orquestradas e mensuráveis." },
        { name: "VP de Marketing", funnelStage: "Awareness", pain: "Precisa provar ROI de martech ao board.", message: "Ative o dado que você já paga e mostre receita incremental." },
      ],
      channelMix: [
        { channel: "LinkedIn Ads", funnelStage: "Awareness", role: "Alcance qualificado em contas-alvo e ABM." },
        { channel: "Webinar", funnelStage: "Consideration", role: "Prova de valor e captura de intenção alta." },
        { channel: "Paid Search", funnelStage: "Decision", role: "Captura de demanda no fundo do funil." },
        { channel: "WhatsApp Nurture", funnelStage: "Consideration", role: "Nutrição regional de alto engajamento." },
      ],
      messageArchitecture: [
        { persona: "Head de CRM / Lifecycle", awareness: "O dado que você coleta pode agir sozinho.", consideration: "Veja jornadas orquestradas em minutos, não sprints.", decision: "Migração assistida e time-to-value em 30 dias." },
        { persona: "Diretor de Growth", awareness: "Retenção é o novo crescimento.", consideration: "Meça o incremento de LTV por jornada.", decision: "Piloto de 60 dias com meta de receita atrelada." },
      ],
      assetPlan: [
        { channel: "LinkedIn Ads", stage: "Awareness", asset: "Vídeo manifesto 15s + carrossel de dados", format: "Social" },
        { channel: "Webinar", stage: "Consideration", asset: "Sessão ao vivo com case regional", format: "Evento" },
        { channel: "Paid Search", stage: "Decision", asset: "Landing de comparação + demo sob demanda", format: "Web" },
        { channel: "WhatsApp Nurture", stage: "Consideration", asset: "Sequência de 4 mensagens com ROI calculator", format: "Mensageria" },
      ],
    },
    localization: {
      language: "Português (BR), com tom consultivo e direto.",
      tone: "Profissional, orientado a ROI, sem jargão importado sem tradução.",
      localChannels: ["WhatsApp Business API", "LinkedIn (peso alto em BR)", "Eventos presenciais em São Paulo"],
      compliance: ["LGPD: base legal e opt-in explícito para WhatsApp", "Registro de consentimento e canal de descadastro claro"],
      culturalNotes: ["Prova social regional pesa mais que benchmark global", "Preferência por relacionamento e piloto antes de contrato longo"],
      calendarNotes: ["Evitar semana de feriado prolongado", "Pico de atenção B2B: terça a quinta, manhã"],
    },
    insights: {
      summary: "A campanha entregou volume de topo acima da meta (leads em 117% e MQL em 109%), mas perdeu força na conversão para baixo do funil: SQL em 89% e pipeline em 80% do planejado. O CPL agregado ficou saudável, abaixo da meta. O gargalo está na passagem MQL para SQL, concentrado em Content Syndication, canal de maior volume e pior qualidade. Realocar parte do investimento para Paid Search e WhatsApp, que puxam o melhor ROAS, tende a fechar o gap de pipeline.",
      recommendations: [
        { title: "Realocar budget de Content Syndication", detail: "É o canal com pior conversão MQL para SQL e menor ROAS. Reduzir e mover verba para Paid Search e WhatsApp, que lideram em eficiência.", priority: "Alta" },
        { title: "Escalar WhatsApp Nurture", detail: "Menor CPL e forte conversão de meio de funil, com ROAS acima de 22x no realizado. Há espaço para escalar com atenção a compliance LGPD.", priority: "Alta" },
        { title: "Revisar qualificação de MQL", detail: "O volume de MQL superou a meta, mas a conversão para SQL ficou abaixo. Ajustar o critério de MQL evita inflar o topo sem gerar pipeline.", priority: "Média" },
        { title: "Monitorar o Paid Search de fundo de funil", detail: "Melhor gerador de pipeline e receita, já operando bem. Apenas acompanhar a cobertura de termos de alta intenção e não realocar verba deste canal.", priority: "Baixa" },
      ],
      gaps: [
        "Pipeline em 80% da meta, principal desvio a corrigir.",
        "SQL em 89%, puxado pela baixa conversão de Content Syndication.",
        "Receita em 92%, recuperável com a realocação de canais.",
      ],
    },
  },

  EN: {
    profile: {
      offering: "Customer engagement and activation platform (CDP + cross-channel orchestration).",
      category: "Customer Engagement / MarTech",
      audience: "CRM, Growth and Lifecycle teams at mid-market and enterprise B2C and B2B2C companies.",
    },
    objective: "Generate new pipeline and upsell across LATAM enterprise accounts with a multichannel anchor campaign.",
    plan: {
      thesis: "LATAM brands invested in data collection but activate very little of it. This anchor campaign positions real-time activation as the next ROI step, connecting data to 1:1 interaction without adding headcount.",
      whyNow: [
        "Pressure on marketing efficiency raises the ROI bar per channel in 2026.",
        "CDP maturity across the region creates an installed base ready for advanced activation.",
        "Cookie restrictions and rising media costs push brands toward owned channels.",
      ],
      personas: [
        { name: "Head of CRM / Lifecycle", funnelStage: "Consideration", pain: "Rich data, slow and manual activation.", message: "Turn idle data into 1:1 interaction without new headcount." },
        { name: "Growth Director", funnelStage: "Decision", pain: "Rising CAC, retention pressuring the target.", message: "Grow LTV with orchestrated, measurable journeys." },
        { name: "VP of Marketing", funnelStage: "Awareness", pain: "Needs to prove martech ROI to the board.", message: "Activate the data you already pay for and show incremental revenue." },
      ],
      channelMix: [
        { channel: "LinkedIn Ads", funnelStage: "Awareness", role: "Qualified reach across target accounts and ABM." },
        { channel: "Webinar", funnelStage: "Consideration", role: "Proof of value and high-intent capture." },
        { channel: "Paid Search", funnelStage: "Decision", role: "Bottom-of-funnel demand capture." },
        { channel: "WhatsApp Nurture", funnelStage: "Consideration", role: "High-engagement regional nurture." },
      ],
      messageArchitecture: [
        { persona: "Head of CRM / Lifecycle", awareness: "The data you collect can act on its own.", consideration: "See orchestrated journeys in minutes, not sprints.", decision: "Assisted migration and time-to-value in 30 days." },
        { persona: "Growth Director", awareness: "Retention is the new growth.", consideration: "Measure incremental LTV per journey.", decision: "A 60-day pilot tied to a revenue target." },
      ],
      assetPlan: [
        { channel: "LinkedIn Ads", stage: "Awareness", asset: "15s manifesto video + data carousel", format: "Social" },
        { channel: "Webinar", stage: "Consideration", asset: "Live session with a regional case", format: "Event" },
        { channel: "Paid Search", stage: "Decision", asset: "Comparison landing page + on-demand demo", format: "Web" },
        { channel: "WhatsApp Nurture", stage: "Consideration", asset: "4-message sequence with ROI calculator", format: "Messaging" },
      ],
    },
    localization: {
      language: "Brazilian Portuguese, consultative and direct in tone.",
      tone: "Professional and ROI-driven, avoiding untranslated imported jargon.",
      localChannels: ["WhatsApp Business API", "LinkedIn (heavily weighted in Brazil)", "In-person events in Sao Paulo"],
      compliance: ["LGPD: legal basis and explicit opt-in for WhatsApp", "Consent logging and a clear opt-out path"],
      culturalNotes: ["Regional social proof outweighs global benchmarks", "Preference for relationship building and a pilot before a long contract"],
      calendarNotes: ["Avoid long holiday weeks", "Peak B2B attention: Tuesday to Thursday, mornings"],
    },
    insights: {
      summary: "The campaign delivered top-of-funnel volume above target (leads at 117% and MQL at 109%), but lost momentum further down: SQL at 89% and pipeline at 80% of plan. Blended CPL stayed healthy, below target. The bottleneck is the MQL to SQL handoff, concentrated in Content Syndication, the highest-volume and lowest-quality channel. Shifting part of the budget to Paid Search and WhatsApp, which drive the best ROAS, should close the pipeline gap.",
      recommendations: [
        { title: "Reallocate budget from Content Syndication", detail: "It has the weakest MQL to SQL conversion and the lowest ROAS. Cut back and move spend to Paid Search and WhatsApp, which lead on efficiency.", priority: "High" },
        { title: "Scale WhatsApp Nurture", detail: "Lowest CPL and strong mid-funnel conversion, with actual ROAS above 22x. There is room to scale, with attention to LGPD compliance.", priority: "High" },
        { title: "Revisit MQL qualification", detail: "MQL volume beat the target while SQL conversion fell short. Tightening the MQL definition avoids inflating the top without generating pipeline.", priority: "Medium" },
        { title: "Monitor bottom-of-funnel Paid Search", detail: "The strongest pipeline and revenue driver, already performing well. Simply track high-intent term coverage and avoid reallocating budget away from this channel.", priority: "Low" },
      ],
      gaps: [
        "Pipeline at 80% of target, the main deviation to correct.",
        "SQL at 89%, dragged down by weak Content Syndication conversion.",
        "Revenue at 92%, recoverable through channel reallocation.",
      ],
    },
  },
};

/* Returns the demo fields for a language. Used both to build the sample
   campaign and to switch it instantly, with no network call.
   Checked field-by-field (not just a few signal fields) so any user edit to
   the sample campaign - company, audience, KPI targets, anything demoPatch
   would touch - correctly falls through to the real translation path instead
   of being silently overwritten by the canned demo text. */
function isPristineSample(c) {
  if (!c || !c.demo) return false;
  const lang = c.generatedLang === "EN" ? "EN" : "PT";
  if (!DEMO_TEXT[lang]) return false;
  const patch = demoPatch(lang);
  const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const profile = c.profile || {};
  const plan = c.plan || {};
  return eq(profile.company, patch.profile.company)
    && eq(profile.competitors, patch.profile.competitors)
    && eq(profile.market, patch.profile.market)
    && eq(profile.offering, patch.profile.offering)
    && eq(profile.category, patch.profile.category)
    && eq(profile.audience, patch.profile.audience)
    && eq(c.objective, patch.objective)
    && eq(plan.thesis, patch.plan.thesis)
    && eq(plan.whyNow, patch.plan.whyNow)
    && eq(plan.personas, patch.plan.personas)
    && eq(plan.channelMix, patch.plan.channelMix)
    && eq(plan.messageArchitecture, patch.plan.messageArchitecture)
    && eq(plan.assetPlan, patch.plan.assetPlan)
    && eq(plan.kpiTargets, patch.plan.kpiTargets)
    && eq(c.localization, patch.localization)
    && eq(c.insights, patch.insights);
}

function demoPatch(lang) {
  const t = DEMO_TEXT[lang === "EN" ? "EN" : "PT"];
  return {
    profile: {
      company: "Northwind Cloud",
      competitors: "Braze, Salesforce Marketing Cloud, Iterable, MoEngage",
      market: "Brasil / LATAM",
      offering: t.profile.offering,
      category: t.profile.category,
      audience: t.profile.audience,
    },
    objective: t.objective,
    plan: { ...t.plan, kpiTargets: DEMO_KPI_TARGETS },
    localization: t.localization,
    insights: t.insights,
  };
}

function buildDemo(lang) {
  const target = lang === "EN" ? "EN" : "PT";
  return {
    id: uid(),
    name: "Q3 Anchor: Activation Intelligence (LATAM)",
    demo: true,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    campaignType: "Anchor",
    contentLang: target,
    generatedLang: target,
    actualsRows: DEMO_ROWS,
    ...demoPatch(target),
  };
}

/* ================================ UI atoms =============================== */
function Segmented({ options, value, onChange, T, icons }) {
  return (
    <div style={{ display: "inline-flex", background: T.chip, borderRadius: 12, padding: 3, gap: 2 }}>
      {options.map((o, i) => {
        const active = o === value;
        const Ico = icons && icons[i];
        return (
          <button key={o} onClick={() => onChange(o)}
            style={{
              display: "inline-flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
              padding: "7px 13px", borderRadius: 10, fontSize: 13, fontWeight: 600, fontFamily: T.font,
              color: active ? (T.dark ? "#fff" : T.ink) : T.muted,
              background: active ? (T.dark ? "#48484A" : "#fff") : "transparent",
              boxShadow: active ? "0 1px 3px rgba(0,0,0,0.14)" : "none",
              transition: "all .18s ease",
            }}>
            {Ico ? <Ico size={14} /> : null}{o}
          </button>
        );
      })}
    </div>
  );
}

function Btn({ children, onClick, T, variant = "primary", size = "md", disabled, icon: Ico, style }) {
  const base = {
    display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8,
    border: "none", cursor: disabled ? "not-allowed" : "pointer", fontFamily: T.font,
    fontWeight: 600, borderRadius: 12, transition: "transform .12s ease, opacity .2s ease",
    opacity: disabled ? 0.5 : 1, whiteSpace: "nowrap",
    padding: size === "sm" ? "8px 13px" : "12px 20px",
    fontSize: size === "sm" ? 13 : 15,
  };
  const variants = {
    primary: { background: T.grad, color: "#fff", boxShadow: "0 6px 18px rgba(110,58,255,0.35)" },
    solid: { background: T.dark ? "#fff" : T.ink, color: T.dark ? "#000" : "#fff" },
    ghost: { background: T.chip, color: T.ink },
    outline: { background: "transparent", color: T.ink, border: `1px solid ${T.hair}` },
    danger: { background: "transparent", color: T.bad, border: `1px solid ${T.hair}` },
  };
  return (
    <button onClick={disabled ? undefined : onClick} style={{ ...base, ...variants[variant], ...style }}
      onMouseDown={(e) => !disabled && (e.currentTarget.style.transform = "scale(0.97)")}
      onMouseUp={(e) => (e.currentTarget.style.transform = "scale(1)")}
      onMouseLeave={(e) => (e.currentTarget.style.transform = "scale(1)")}>
      {Ico ? <Ico size={size === "sm" ? 15 : 17} /> : null}{children}
    </button>
  );
}

function Card({ children, T, style, pad = 22 }) {
  return (
    <div style={{
      background: T.card, borderRadius: 20, padding: pad, boxShadow: T.shadow,
      border: `1px solid ${T.hair}`, ...style,
    }}>{children}</div>
  );
}

function Eyebrow({ children, T, color }) {
  return (
    <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 1.4, textTransform: "uppercase",
      color: color || T.violet, marginBottom: 8 }}>{children}</div>
  );
}

function Field({ label, T, children }) {
  return (
    <label style={{ display: "block", marginBottom: 14 }}>
      <div style={{ fontSize: 12.5, fontWeight: 600, color: T.muted, marginBottom: 6 }}>{label}</div>
      {children}
    </label>
  );
}
function inputStyle(T) {
  return {
    width: "100%", boxSizing: "border-box", padding: "11px 13px", borderRadius: 12,
    border: `1px solid ${T.hair}`, background: T.cardEl, color: T.ink, fontSize: 14.5,
    fontFamily: T.font, outline: "none",
  };
}

function Pill({ children, T, tone }) {
  const map = { violet: T.violet, coral: T.coral, good: T.good, warn: T.warn, muted: T.faint };
  const c = map[tone] || T.violet;
  return (
    <span style={{ display: "inline-block", fontSize: 11, fontWeight: 700, letterSpacing: 0.3,
      padding: "3px 9px", borderRadius: 999, color: c, background: c + "1A" }}>{children}</span>
  );
}

function Stat({ label, value, sub, T, accent }) {
  return (
    <div style={{ background: T.cardEl, borderRadius: 16, padding: 16, border: `1px solid ${T.hair}` }}>
      <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 22, fontWeight: 750, color: accent || T.ink, letterSpacing: -0.4, fontVariantNumeric: "tabular-nums" }}>{value}</div>
      {sub ? <div style={{ fontSize: 12, color: T.faint, marginTop: 3 }}>{sub}</div> : null}
    </div>
  );
}

function Spinner({ T, label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, color: T.muted, fontSize: 14 }}>
      <div style={{ width: 18, height: 18, borderRadius: "50%", border: `2px solid ${T.hair}`,
        borderTopColor: T.violet, animation: "cdspin 0.8s linear infinite" }} />
      {label || "Generating…"}
    </div>
  );
}

function SectionTitle({ icon: Ico, children, T, hint }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
      <div style={{ width: 30, height: 30, borderRadius: 9, display: "grid", placeItems: "center",
        background: T.violet + "18", color: T.violet }}><Ico size={17} /></div>
      <div>
        <div style={{ fontSize: 16, fontWeight: 700, color: T.ink, letterSpacing: -0.2 }}>{children}</div>
        {hint ? <div style={{ fontSize: 12, color: T.faint }}>{hint}</div> : null}
      </div>
    </div>
  );
}

/* ---------------- KPI targets: user-owned, editable (integrity) ---------- */
const KPI_METRICS = ["leads", "mql", "sql", "pipeline", "revenue", "cpl"];
const targetOf = (kpiTargets, m) => {
  const f = (kpiTargets || []).find((x) => x.metric === m);
  return f && f.target !== undefined && f.target !== null ? f.target : "";
};
function KpiTargetsEditor({ campaign, setCampaign, T }) {
  const setTarget = (metric, value) =>
    setCampaign((c) => {
      const list = [...((c.plan && c.plan.kpiTargets) || [])];
      const i = list.findIndex((x) => x.metric === metric);
      const v = value === "" ? "" : num(value);
      if (i >= 0) list[i] = { metric, target: v };
      else list.push({ metric, target: v });
      return { ...c, plan: { ...c.plan, kpiTargets: list } };
    });
  return (
    <div className="ici-grid3">
      {KPI_METRICS.map((m) => {
        const meta = METRIC_META[m];
        const cur = targetOf(campaign.plan.kpiTargets, m);
        const unit = m === "pipeline" || m === "revenue" || m === "cpl" ? " (R$)" : "";
        return (
          <div key={m} style={{ background: T.cardEl, borderRadius: 14, padding: 12, border: `1px solid ${T.hair}` }}>
            <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 6 }}>{meta.label}{unit}</div>
            <input type="number" min="0" value={cur} onChange={(e) => setTarget(m, e.target.value)} placeholder="0"
              style={{ ...inputStyle(T), padding: "9px 11px", fontSize: 15, fontWeight: 700 }} />
          </div>
        );
      })}
    </div>
  );
}

/* ============================== PLAN PHASE =============================== */
function PlanPhase({ campaign, setCampaign, T, notify }) {
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");
  const p = campaign.plan;

  const upProfile = (k, v) =>
    setCampaign((c) => ({ ...c, profile: { ...c.profile, [k]: v } }));

  async function generate() {
    setErr(""); setLoading(true);
    try {
      const langName = campaign.contentLang === "EN" ? "English" : "Brazilian Portuguese";
      const sys =
        "You are a senior B2B integrated campaign strategist. Return ONLY valid minified JSON, no prose, no markdown. " +
        "All content in " + langName + ". Do not invent statistics or named cases. " +
        "Keep every array to exactly 3 items and every string under 22 words. Be concise.";
      const user =
        `Empresa: ${campaign.profile.company}\nOferta: ${campaign.profile.offering}\nCategoria: ${campaign.profile.category}\n` +
        `Concorrentes: ${campaign.profile.competitors}\nMercado: ${campaign.profile.market}\nPúblico: ${campaign.profile.audience}\n` +
        `Objetivo: ${campaign.objective}\nTipo de campanha: ${campaign.campaignType}\n\n` +
        `Gere um plano de campanha integrada com este schema JSON exato:\n` +
        `{"thesis":"","whyNow":["",""],"personas":[{"name":"","funnelStage":"Awareness|Consideration|Decision","pain":"","message":""}],` +
        `"channelMix":[{"channel":"","funnelStage":"","role":""}],` +
        `"messageArchitecture":[{"persona":"","awareness":"","consideration":"","decision":""}],` +
        `"assetPlan":[{"channel":"","stage":"","asset":"","format":""}],` +
        `"kpiTargets":[{"metric":"leads|mql|sql|pipeline|revenue|cpl","target":0}]}`;
      const text = await callAI(sys, user);
      const json = extractJSON(text);
      if (!json) throw new Error("Couldn't parse the response. Try again.");
      setCampaign((c) => ({ ...c, plan: { ...c.plan, ...json }, generatedLang: c.contentLang || "PT" }));
      notify("Plan generated.");
    } catch (e) {
      setErr(e.message || "Failed to generate the plan.");
    } finally { setLoading(false); }
  }

  return (
    <div style={{ display: "grid", gap: 18 }}>
      <Card T={T}>
        <SectionTitle icon={Target} T={T} hint="Set once. Generation uses this profile.">Campaign profile</SectionTitle>
        <div className="ici-grid2" style={{ display: "grid", gap: 14 }}>
          <Field label="Company" T={T}><input style={inputStyle(T)} value={campaign.profile.company} onChange={(e) => upProfile("company", e.target.value)} /></Field>
          <Field label="Category" T={T}><input style={inputStyle(T)} value={campaign.profile.category} onChange={(e) => upProfile("category", e.target.value)} /></Field>
        </div>
        <Field label="Offering / value proposition" T={T}><textarea rows={2} style={{ ...inputStyle(T), resize: "vertical" }} value={campaign.profile.offering} onChange={(e) => upProfile("offering", e.target.value)} /></Field>
        <div className="ici-grid2" style={{ display: "grid", gap: 14 }}>
          <Field label="Competitors" T={T}><input style={inputStyle(T)} value={campaign.profile.competitors} onChange={(e) => upProfile("competitors", e.target.value)} /></Field>
          <Field label="Primary market" T={T}><input style={inputStyle(T)} value={campaign.profile.market} onChange={(e) => upProfile("market", e.target.value)} /></Field>
        </div>
        <Field label="Target audience" T={T}><input style={inputStyle(T)} value={campaign.profile.audience} onChange={(e) => upProfile("audience", e.target.value)} /></Field>
        <div className="ici-grid21" style={{ display: "grid", gap: 14 }}>
          <Field label="Campaign objective" T={T}><input style={inputStyle(T)} value={campaign.objective} onChange={(e) => setCampaign((c) => ({ ...c, objective: e.target.value }))} /></Field>
          <Field label="Type" T={T}>
            <select style={inputStyle(T)} value={campaign.campaignType} onChange={(e) => setCampaign((c) => ({ ...c, campaignType: e.target.value }))}>
              {["Anchor", "Brand", "Product", "Competitor"].map((x) => <option key={x}>{x}</option>)}
            </select>
          </Field>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginTop: 6 }}>
          <Btn T={T} onClick={generate} disabled={loading} icon={Sparkles}>{loading ? "Generating…" : "Generate plan"}</Btn>
          {loading ? <Spinner T={T} /> : null}
          {err ? <span style={{ color: T.bad, fontSize: 13 }}>{err}</span> : null}
        </div>
      </Card>

      {p.thesis ? (
        <Card T={T}>
          <Eyebrow T={T}>Thesis & timing</Eyebrow>
          <p style={{ fontSize: 16, lineHeight: 1.55, color: T.ink, marginTop: 4, marginBottom: 16, fontWeight: 500 }}>{p.thesis}</p>
          <div style={{ display: "grid", gap: 8 }}>
            {(p.whyNow || []).map((w, i) => (
              <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
                <Zap size={16} style={{ color: T.coral, marginTop: 2, flexShrink: 0 }} />
                <span style={{ fontSize: 14, color: T.muted, lineHeight: 1.5 }}>{w}</span>
              </div>
            ))}
          </div>
        </Card>
      ) : null}

      {(p.personas || []).length ? (
        <Card T={T}>
          <SectionTitle icon={Users} T={T}>Personas by funnel stage</SectionTitle>
          <div style={{ display: "grid", gap: 12 }}>
            {p.personas.map((ps, i) => (
              <div key={i} style={{ background: T.cardEl, borderRadius: 14, padding: 15, border: `1px solid ${T.hair}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                  <span style={{ fontWeight: 700, color: T.ink, fontSize: 15 }}>{ps.name}</span>
                  <Pill T={T} tone="violet">{ps.funnelStage}</Pill>
                </div>
                <div style={{ fontSize: 13.5, color: T.muted, marginBottom: 6 }}><b style={{ color: T.faint }}>Pain: </b>{ps.pain}</div>
                <div style={{ fontSize: 13.5, color: T.ink }}><b style={{ color: T.coral }}>Message: </b>{ps.message}</div>
              </div>
            ))}
          </div>
        </Card>
      ) : null}

      {(p.channelMix || []).length ? (
        <Card T={T}>
          <SectionTitle icon={Megaphone} T={T}>Multichannel mix mapped to the funnel</SectionTitle>
          <div style={{ display: "grid", gap: 10 }}>
            {p.channelMix.map((ch, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0", borderBottom: i < p.channelMix.length - 1 ? `1px solid ${T.hair}` : "none" }}>
                <div style={{ width: 90, flexShrink: 0 }}><Pill T={T} tone="muted">{ch.funnelStage}</Pill></div>
                <div style={{ fontWeight: 650, color: T.ink, fontSize: 14, width: 150, flexShrink: 0 }}>{ch.channel}</div>
                <div style={{ fontSize: 13.5, color: T.muted }}>{ch.role}</div>
              </div>
            ))}
          </div>
        </Card>
      ) : null}

      {(p.messageArchitecture || []).length ? (
        <Card T={T} pad={0}>
          <div style={{ padding: "22px 22px 10px" }}><SectionTitle icon={Layers} T={T}>Message architecture</SectionTitle></div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ color: T.faint, textAlign: "left" }}>
                  {["Persona", "Awareness", "Consideration", "Decision"].map((h) => (
                    <th key={h} style={{ padding: "8px 16px", fontWeight: 700, fontSize: 11, letterSpacing: 0.5, textTransform: "uppercase", borderBottom: `1px solid ${T.hair}` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {p.messageArchitecture.map((m, i) => (
                  <tr key={i}>
                    <td style={{ padding: "12px 16px", fontWeight: 650, color: T.ink, borderBottom: `1px solid ${T.hair}` }}>{m.persona}</td>
                    <td style={{ padding: "12px 16px", color: T.muted, borderBottom: `1px solid ${T.hair}` }}>{m.awareness}</td>
                    <td style={{ padding: "12px 16px", color: T.muted, borderBottom: `1px solid ${T.hair}` }}>{m.consideration}</td>
                    <td style={{ padding: "12px 16px", color: T.muted, borderBottom: `1px solid ${T.hair}` }}>{m.decision}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      ) : null}

      {(p.assetPlan || []).length ? (
        <Card T={T}>
          <SectionTitle icon={FileText} T={T}>Asset plan</SectionTitle>
          <div style={{ display: "grid", gap: 10 }}>
            {p.assetPlan.map((a, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0", borderBottom: i < p.assetPlan.length - 1 ? `1px solid ${T.hair}` : "none", flexWrap: "wrap" }}>
                <div style={{ width: 90, flexShrink: 0 }}><Pill T={T} tone="muted">{a.stage}</Pill></div>
                <div style={{ fontWeight: 650, color: T.ink, fontSize: 14, width: 150, flexShrink: 0 }}>{a.channel}</div>
                <div style={{ flex: 1, minWidth: 140, fontSize: 13.5, color: T.muted }}>{a.asset}</div>
                <Pill T={T} tone="violet">{a.format}</Pill>
              </div>
            ))}
          </div>
        </Card>
      ) : null}

      <Card T={T}>
        <SectionTitle icon={Target} T={T} hint="Targets are yours to set. The AI can suggest starting values, but review before use.">KPI targets (full-funnel)</SectionTitle>
        <KpiTargetsEditor campaign={campaign} setCampaign={setCampaign} T={T} />
      </Card>
    </div>
  );
}

/* ============================ LOCALIZE PHASE ============================= */
function LocalizePhase({ campaign, setCampaign, T, notify }) {
  const [loading, setLoading] = useState(false);
  const [market, setMarket] = useState(campaign.profile.market || "Brasil");
  const [err, setErr] = useState("");
  const L = campaign.localization;

  async function generate() {
    setErr(""); setLoading(true);
    try {
      const langName = campaign.contentLang === "EN" ? "English" : "Brazilian Portuguese";
      const sys = "You localize global B2B campaigns for regional markets. Return ONLY minified JSON. Content in " + langName + ". Be concrete and market-specific, not generic translation advice.";
      const user =
        `Campanha: ${campaign.name}\nEmpresa: ${campaign.profile.company}\nOferta: ${campaign.profile.offering}\n` +
        `Mercado de destino: ${market}\nPúblico: ${campaign.profile.audience}\n\n` +
        `Gere a camada de localização com este schema:\n` +
        `{"language":"","tone":"","localChannels":["",""],"compliance":["",""],"culturalNotes":["",""],"calendarNotes":["",""]}`;
      const text = await callAI(sys, user);
      const json = extractJSON(text);
      if (!json) throw new Error("Couldn't parse the response.");
      setCampaign((c) => ({ ...c, localization: { ...c.localization, ...json }, generatedLang: c.contentLang || "PT" }));
      notify("Localization generated.");
    } catch (e) { setErr(e.message); } finally { setLoading(false); }
  }

  const Block = ({ icon: Ico, title, items }) => (
    <Card T={T}>
      <SectionTitle icon={Ico} T={T}>{title}</SectionTitle>
      <div style={{ display: "grid", gap: 8 }}>
        {(items || []).map((x, i) => (
          <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
            <CheckCircle2 size={16} style={{ color: T.violet, marginTop: 2, flexShrink: 0 }} />
            <span style={{ fontSize: 14, color: T.muted, lineHeight: 1.5 }}>{x}</span>
          </div>
        ))}
      </div>
    </Card>
  );

  return (
    <div style={{ display: "grid", gap: 18 }}>
      <Card T={T}>
        <SectionTitle icon={Globe} T={T} hint="Real adaptation of the global campaign to the region, not just translation.">Localization layer</SectionTitle>
        <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 200 }}>
            <Field label="Target market" T={T}><input style={inputStyle(T)} value={market} onChange={(e) => setMarket(e.target.value)} /></Field>
          </div>
          <Btn T={T} onClick={generate} disabled={loading} icon={Sparkles} style={{ marginBottom: 14 }}>{loading ? "Generating…" : "Generate localization"}</Btn>
        </div>
        {loading ? <Spinner T={T} /> : null}
        {err ? <span style={{ color: T.bad, fontSize: 13 }}>{err}</span> : null}
      </Card>

      {L.language ? (
        <div className="ici-grid2" style={{ display: "grid", gap: 14 }}>
          <Card T={T}><Eyebrow T={T}>Language</Eyebrow><p style={{ color: T.ink, fontSize: 14.5, lineHeight: 1.5 }}>{L.language}</p></Card>
          <Card T={T}><Eyebrow T={T} color={T.coral}>Tone</Eyebrow><p style={{ color: T.ink, fontSize: 14.5, lineHeight: 1.5 }}>{L.tone}</p></Card>
        </div>
      ) : null}
      {(L.localChannels || []).length ? <Block icon={Megaphone} title="Local channels" items={L.localChannels} /> : null}
      {(L.compliance || []).length ? <Block icon={AlertTriangle} title="Compliance (LGPD)" items={L.compliance} /> : null}
      {(L.culturalNotes || []).length ? <Block icon={Users} title="Cultural notes" items={L.culturalNotes} /> : null}
      {(L.calendarNotes || []).length ? <Block icon={Calendar} title="Local calendar" items={L.calendarNotes} /> : null}
    </div>
  );
}

/* ============================= MEASURE PHASE ============================= */
function downloadTemplate() {
  const example = [
    { Channel: "LinkedIn Ads", "Funnel Stage": "Awareness", Spend: 96000, Impressions: 1420000, Clicks: 12800, Leads: 640, MQL: 310, SQL: 120, Opportunities: 38, Pipeline: 1710000, Revenue: 285000 },
    { Channel: "Paid Search", "Funnel Stage": "Decision", Spend: 74000, Impressions: 380000, Clicks: 21200, Leads: 980, MQL: 520, SQL: 210, Opportunities: 66, Pipeline: 2970000, Revenue: 620000 },
  ];
  const ws = XLSX.utils.json_to_sheet(example, { header: TEMPLATE_COLUMNS });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Campaign Data");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "campaign-data-template.xlsx"; a.click();
  URL.revokeObjectURL(url);
}

function parseWorkbook(file, onDone, onError) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const norm = (obj, keys) => {
        for (const k of Object.keys(obj)) {
          const lk = k.toLowerCase().trim();
          if (keys.some((x) => lk.includes(x))) return obj[k];
        }
        return "";
      };
      // Impressions/clicks/leads/MQL/SQL/opportunities are physical counts:
      // negative values (e.g. an accounting "(50)" export artifact) can only
      // be bad data, and left unclamped they silently produce negative CPL
      // and skew every aggregate. Spend/pipeline/revenue can legitimately go
      // negative (refunds, adjustments), so those are left as parsed.
      let negativeCount = 0;
      const count = (v) => {
        const n = num(v);
        if (n < 0) { negativeCount++; return 0; }
        return n;
      };
      const rows = json.map((r) => ({
        channel: norm(r, ["channel", "canal"]),
        stage: canonicalStage(norm(r, ["stage", "funil", "funnel"])),
        spend: num(norm(r, ["spend", "invest", "custo"])),
        impressions: count(norm(r, ["impress"])),
        clicks: count(norm(r, ["click", "clique"])),
        leads: count(norm(r, ["lead"])),
        mql: count(norm(r, ["mql"])),
        sql: count(norm(r, ["sql"])),
        opps: count(norm(r, ["opportun", "oportun", "opp"])),
        pipeline: num(norm(r, ["pipeline", "pipe"])),
        revenue: num(norm(r, ["revenue", "receita"])),
      })).filter((r) => r.channel);
      if (!rows.length) throw new Error("No valid rows found. Check the template.");
      const warning = negativeCount
        ? `${negativeCount} negative value(s) found in count fields (impressions, clicks, leads, MQL, SQL, opportunities) and treated as 0.`
        : null;
      onDone(rows, warning);
    } catch (err) { onError(err.message || "Failed to read the file."); }
  };
  reader.onerror = () => onError("Couldn't read the file.");
  reader.readAsArrayBuffer(file);
}

function MeasurePhase({ campaign, setCampaign, T, notify }) {
  const [err, setErr] = useState("");
  const fileRef = useRef(null);
  const rows = campaign.actualsRows || [];
  const { byChannel, totals } = useMemo(() => computeDerived(rows), [rows]);

  const onFile = (f) => {
    if (!f) return;
    setErr("");
    parseWorkbook(f, (parsed, warning) => {
      setCampaign((c) => ({ ...c, actualsRows: parsed, insights: null }));
      notify(warning || "Spreadsheet loaded.", Boolean(warning));
    }, (m) => setErr(m));
  };

  const funnelData = [
    { name: "Leads", v: totals.leads }, { name: "MQL", v: totals.mql },
    { name: "SQL", v: totals.sql }, { name: "Opps", v: totals.opps },
  ];
  const pieColors = ["#6E3AFF", "#A03BFF", "#FF5C39", "#FF9F0A", "#30D158", "#0A84FF"];

  return (
    <div style={{ display: "grid", gap: 18 }}>
      <Card T={T}>
        <SectionTitle icon={FileSpreadsheet} T={T} hint="Upload the campaign's real extract. The app computes, it doesn't guess.">Actuals</SectionTitle>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <Btn T={T} variant="ghost" size="sm" icon={Download} onClick={downloadTemplate}>Download template</Btn>
          <Btn T={T} size="sm" icon={Upload} onClick={() => fileRef.current && fileRef.current.click()}>Upload spreadsheet</Btn>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={(e) => { onFile(e.target.files && e.target.files[0]); e.target.value = ""; }} />
          {rows.length ? <span style={{ alignSelf: "center", fontSize: 13, color: T.muted }}>{rows.length} channels loaded</span> : null}
        </div>
        {err ? <div style={{ color: T.bad, fontSize: 13, marginTop: 10 }}>{err}</div> : null}
      </Card>

      {rows.length ? (
        <>
          <div className="ici-grid4" style={{ display: "grid", gap: 10 }}>
            <Stat T={T} label="Spend" value={brl(totals.spend)} />
            <Stat T={T} label="Leads" value={int(totals.leads)} sub={`CPL ${money2(totals.cpl)}`} />
            <Stat T={T} label="Pipeline" value={brl(totals.pipeline)} accent={T.violet} />
            <Stat T={T} label="ROAS" value={totals.roas.toFixed(2) + "x"} sub={`ROI ${pct(totals.roi, 0)}`} accent={totals.roi >= 0 ? T.good : T.bad} />
          </div>

          <div className="ici-grid2" style={{ display: "grid", gap: 14 }}>
            <Card T={T}>
              <Eyebrow T={T}>Conversion funnel</Eyebrow>
              <div style={{ height: 220 }}>
                <ResponsiveContainer>
                  <BarChart data={funnelData} margin={{ top: 10, right: 6, left: -18, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.hair} vertical={false} />
                    <XAxis dataKey="name" tick={{ fill: T.muted, fontSize: 12 }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fill: T.faint, fontSize: 11 }} axisLine={false} tickLine={false} />
                    <Tooltip contentStyle={{ background: T.card, border: `1px solid ${T.hair}`, borderRadius: 12, color: T.ink }} />
                    <Bar dataKey="v" name="Volume" radius={[6, 6, 0, 0]} fill="#6E3AFF" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Card>
            <Card T={T}>
              <Eyebrow T={T} color={T.coral}>Spend by channel</Eyebrow>
              <div style={{ height: 220 }}>
                <ResponsiveContainer>
                  <PieChart>
                    <Pie data={byChannel} dataKey="spend" nameKey="channel" cx="50%" cy="50%" innerRadius={45} outerRadius={78} paddingAngle={2}>
                      {byChannel.map((e, i) => <Cell key={i} fill={pieColors[i % pieColors.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => brl(v)} contentStyle={{ background: T.card, border: `1px solid ${T.hair}`, borderRadius: 12, color: T.ink }} />
                    <Legend wrapperStyle={{ fontSize: 11, color: T.muted }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </Card>
          </div>

          <Card T={T}>
            <Eyebrow T={T}>Cost per lead by channel</Eyebrow>
            <div style={{ height: 240 }}>
              <ResponsiveContainer>
                <BarChart data={byChannel} margin={{ top: 10, right: 6, left: -6, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.hair} vertical={false} />
                  <XAxis dataKey="channel" tick={{ fill: T.muted, fontSize: 11 }} axisLine={false} tickLine={false} interval={0} angle={-12} height={48} textAnchor="end" />
                  <YAxis tick={{ fill: T.faint, fontSize: 11 }} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v) => money2(v)} contentStyle={{ background: T.card, border: `1px solid ${T.hair}`, borderRadius: 12, color: T.ink }} />
                  <Bar dataKey="cpl" name="CPL" radius={[6, 6, 0, 0]} fill="#FF5C39" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </Card>

          <Card T={T} pad={0}>
            <div style={{ padding: "20px 22px 8px" }}><Eyebrow T={T}>Channel detail</Eyebrow></div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5, fontVariantNumeric: "tabular-nums" }}>
                <thead>
                  <tr style={{ color: T.faint, textAlign: "right" }}>
                    {["Channel", "Spend", "Leads", "CPL", "MQL", "SQL", "Pipeline", "ROAS"].map((h, i) => (
                      <th key={h} style={{ padding: "8px 14px", fontWeight: 700, fontSize: 10.5, letterSpacing: 0.4, textTransform: "uppercase", textAlign: i === 0 ? "left" : "right", borderBottom: `1px solid ${T.hair}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {byChannel.map((r, i) => (
                    <tr key={i}>
                      <td style={{ padding: "10px 14px", fontWeight: 650, color: T.ink, textAlign: "left", borderBottom: `1px solid ${T.hair}` }}>{r.channel}</td>
                      <td style={{ padding: "10px 14px", color: T.muted, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{brl(r.spend)}</td>
                      <td style={{ padding: "10px 14px", color: T.muted, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{int(r.leads)}</td>
                      <td style={{ padding: "10px 14px", color: T.muted, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{money2(r.cpl)}</td>
                      <td style={{ padding: "10px 14px", color: T.muted, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{int(r.mql)}</td>
                      <td style={{ padding: "10px 14px", color: T.muted, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{int(r.sql)}</td>
                      <td style={{ padding: "10px 14px", color: T.muted, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{brl(r.pipeline)}</td>
                      <td style={{ padding: "10px 14px", fontWeight: 650, color: r.roas >= 1 ? T.good : T.bad, textAlign: "right", borderBottom: `1px solid ${T.hair}` }}>{r.roas.toFixed(2)}x</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      ) : (
        <Card T={T} style={{ textAlign: "center", padding: 44 }}>
          <FileSpreadsheet size={34} style={{ color: T.faint, marginBottom: 10 }} />
          <div style={{ fontWeight: 650, color: T.ink, marginBottom: 4 }}>No data yet</div>
          <div style={{ fontSize: 13.5, color: T.muted }}>Download the template, fill it with the campaign extract, and upload to see the dashboard.</div>
        </Card>
      )}
    </div>
  );
}

/* ============================= OPTIMIZE PHASE =========================== */
function OptimizePhase({ campaign, setCampaign, T, notify }) {
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");
  const rows = campaign.actualsRows || [];
  const { byChannel, totals } = useMemo(() => computeDerived(rows), [rows]);
  const pva = useMemo(() => planVsActual(campaign.plan.kpiTargets, totals), [campaign.plan.kpiTargets, totals]);
  const insights = campaign.insights;

  async function generate() {
    setErr(""); setLoading(true);
    try {
      const compact = byChannel.map((r) => ({
        canal: r.channel, invest: Math.round(r.spend), leads: r.leads,
        cpl: Math.round(r.cpl), mql: r.mql, sql: r.sql, pipeline: Math.round(r.pipeline), roas: +r.roas.toFixed(2),
      }));
      const gaps = pva.map((g) => ({ metrica: g.label, meta: g.target, realizado: g.actual, gap_pct: +(g.gapPct * 100).toFixed(1) }));
      const langName = campaign.contentLang === "EN" ? "English" : "Brazilian Portuguese";
      const sys = "You are a senior marketing analyst. You receive ALREADY-COMPUTED numbers. Never invent figures beyond what is given. Return ONLY minified JSON. Content in " + langName + ", executive tone.";
      const user =
        `Totais: invest ${Math.round(totals.spend)}, leads ${totals.leads}, CPL ${Math.round(totals.cpl)}, ` +
        `pipeline ${Math.round(totals.pipeline)}, ROAS ${totals.roas.toFixed(2)}.\n` +
        `Por canal: ${JSON.stringify(compact)}\n` +
        `Plano vs realizado: ${JSON.stringify(gaps)}\n\n` +
        `Gere um resumo executivo e recomendações acionáveis com este schema:\n` +
        `{"summary":"","recommendations":[{"title":"","detail":"","priority":"Alta|Média|Baixa"}],"gaps":[""]}`;
      const text = await callAI(sys, user);
      const json = extractJSON(text);
      if (!json) throw new Error("Couldn't parse the response.");
      setCampaign((c) => ({ ...c, insights: json, generatedLang: c.contentLang || "PT" }));
      notify("Insights generated.");
    } catch (e) { setErr(e.message); } finally { setLoading(false); }
  }

  const pvaChart = pva.map((g) => ({
    name: g.label,
    // For cost metrics (lower is better) attainment inverts, so >100% always means outperforming.
    attain: g.higher
      ? (g.target ? Math.round((g.actual / g.target) * 100) : 0)
      : (g.actual ? Math.round((g.target / g.actual) * 100) : 0),
    favorable: g.favorable,
  }));

  return (
    <div style={{ display: "grid", gap: 18 }}>
      {!rows.length ? (
        <Card T={T} style={{ textAlign: "center", padding: 40 }}>
          <TrendingUp size={32} style={{ color: T.faint, marginBottom: 10 }} />
          <div style={{ fontWeight: 650, color: T.ink }}>Load data in Measure first</div>
          <div style={{ fontSize: 13.5, color: T.muted, marginTop: 4 }}>Optimization compares actuals against the plan's targets.</div>
        </Card>
      ) : (
        <>
          <Card T={T}>
            <SectionTitle icon={Target} T={T} hint="Set or adjust targets here without going back to Plan.">KPI targets</SectionTitle>
            <KpiTargetsEditor campaign={campaign} setCampaign={setCampaign} T={T} />
          </Card>

          {!pva.length ? (
            <Card T={T} style={{ textAlign: "center", padding: 34 }}>
              <Target size={30} style={{ color: T.faint, marginBottom: 10 }} />
              <div style={{ fontWeight: 650, color: T.ink }}>No targets set yet</div>
              <div style={{ fontSize: 13.5, color: T.muted, marginTop: 4 }}>Fill at least one KPI target above to unlock plan vs. actual.</div>
            </Card>
          ) : null}

          {pva.length ? (
            <Card T={T}>
              <SectionTitle icon={Target} T={T} hint="Actuals against the targets set in the plan.">Plan vs. actual</SectionTitle>
              <div style={{ fontSize: 12, color: T.faint, marginBottom: 8 }}>Attainment vs. target (100% = on target). Green favorable, coral unfavorable.</div>
              <div style={{ height: 250, marginBottom: 12 }}>
                <ResponsiveContainer>
                  <BarChart data={pvaChart} margin={{ top: 10, right: 6, left: -6, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.hair} vertical={false} />
                    <XAxis dataKey="name" tick={{ fill: T.muted, fontSize: 11 }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fill: T.faint, fontSize: 10 }} axisLine={false} tickLine={false} unit="%" />
                    <Tooltip formatter={(v) => v + "% of target"} contentStyle={{ background: T.card, border: `1px solid ${T.hair}`, borderRadius: 12, color: T.ink }} />
                    <ReferenceLine y={100} stroke={T.faint} strokeDasharray="4 4" />
                    <Bar dataKey="attain" name="Atingimento" radius={[5, 5, 0, 0]}>
                      {pvaChart.map((e, i) => <Cell key={i} fill={e.favorable ? T.good : T.coral} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="ici-grid3" style={{ display: "grid", gap: 10 }}>
                {pva.map((g, i) => (
                  <div key={i} style={{ background: T.cardEl, borderRadius: 14, padding: 14, border: `1px solid ${T.hair}` }}>
                    <div style={{ fontSize: 12, color: T.muted, fontWeight: 600, marginBottom: 4 }}>{g.label}</div>
                    <div style={{ fontSize: 18, fontWeight: 750, color: T.ink, fontVariantNumeric: "tabular-nums" }}>{g.fmt(g.actual)}</div>
                    <div style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4, fontSize: 12.5, fontWeight: 650, color: g.favorable ? T.good : T.bad }}>
                      {g.favorable ? <TrendingUp size={13} /> : <TrendingDown size={13} />}
                      {(g.gapPct >= 0 ? "+" : "") + pct(g.gapPct, 0)} <span style={{ color: T.faint, fontWeight: 500 }}>vs target</span>
                    </div>
                  </div>
                ))}
              </div>
            </Card>
          ) : null}

          <Card T={T}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
              <SectionTitle icon={Sparkles} T={T} hint="The AI narrates the already-computed numbers. It doesn't invent data.">Executive summary & recommendations</SectionTitle>
              <Btn T={T} size="sm" icon={Sparkles} onClick={generate} disabled={loading} style={{ marginBottom: 14 }}>{loading ? "Generating…" : insights ? "Regenerate" : "Generate insights"}</Btn>
            </div>
            {loading ? <Spinner T={T} label="Analyzing actuals…" /> : null}
            {err ? <div style={{ color: T.bad, fontSize: 13 }}>{err}</div> : null}
            {insights ? (
              <div style={{ marginTop: 6 }}>
                <p style={{ fontSize: 15.5, lineHeight: 1.6, color: T.ink, fontWeight: 500, marginBottom: 18 }}>{insights.summary}</p>
                <div style={{ display: "grid", gap: 10 }}>
                  {(insights.recommendations || []).map((r, i) => {
                    const tone = priorityTone(r.priority, T);
                    return (
                      <div key={i} style={{ background: T.cardEl, borderRadius: 14, padding: "15px 16px", border: `1px solid ${T.hair}`,
                        borderLeft: `3px solid ${tone}`, display: "flex", gap: 12, alignItems: "flex-start" }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, color: T.ink, fontSize: 14.5, lineHeight: 1.35, marginBottom: 5 }}>{r.title}</div>
                          <div style={{ fontSize: 13.5, color: T.muted, lineHeight: 1.5 }}>{r.detail}</div>
                        </div>
                        {r.priority ? (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 5, flexShrink: 0,
                            fontSize: 10.5, fontWeight: 700, letterSpacing: 0.2, color: tone, background: tone + "16",
                            border: `1px solid ${tone}2E`, padding: "4px 9px", borderRadius: 999, whiteSpace: "nowrap" }}>
                            <span style={{ width: 5, height: 5, borderRadius: "50%", background: tone, display: "block" }} />
                            {r.priority}
                          </span>
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (!loading && (
              <div style={{ fontSize: 13.5, color: T.muted }}>Generate the executive summary from the loaded data.</div>
            ))}
          </Card>
        </>
      )}
    </div>
  );
}

/* ============================ My Campaigns drawer ======================= */
function Drawer({ open, onClose, campaigns, onOpen, onDelete, onNew, T, currentId }) {
  return (
    <>
      <div onClick={onClose} style={{
        position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", backdropFilter: "blur(2px)",
        opacity: open ? 1 : 0, pointerEvents: open ? "auto" : "none", transition: "opacity .25s ease", zIndex: 40,
      }} />
      <div style={{
        position: "fixed", top: 0, left: 0, bottom: 0, width: 340, maxWidth: "86vw", background: T.card,
        boxShadow: "8px 0 40px rgba(0,0,0,0.25)", zIndex: 50, transform: open ? "translateX(0)" : "translateX(-100%)",
        transition: "transform .3s cubic-bezier(.22,1,.36,1)", display: "flex", flexDirection: "column",
        borderRight: `1px solid ${T.hair}`,
      }}>
        <div style={{ padding: "20px 20px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `1px solid ${T.hair}` }}>
          <div style={{ fontSize: 17, fontWeight: 750, color: T.ink }}>My Campaigns</div>
          <button onClick={onClose} style={{ background: T.chip, border: "none", borderRadius: 9, width: 32, height: 32, display: "grid", placeItems: "center", cursor: "pointer", color: T.ink }}><X size={17} /></button>
        </div>
        <div style={{ padding: 16 }}>
          <Btn T={T} icon={Plus} onClick={onNew} style={{ width: "100%" }}>New campaign</Btn>
        </div>
        <div style={{ padding: "0 16px 6px", fontSize: 11, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", color: T.faint }}>Saved Campaigns</div>
        <div style={{ flex: 1, overflowY: "auto", padding: "6px 16px 20px" }}>
          {campaigns.length === 0 ? (
            <div style={{ fontSize: 13.5, color: T.muted, padding: "20px 4px", lineHeight: 1.5 }}>
              Nothing saved yet. Create a campaign and tap Save to keep it here.
            </div>
          ) : campaigns.map((c) => (
            <div key={c.id} onClick={() => onOpen(c)} style={{
              background: c.id === currentId ? T.violet + "12" : T.cardEl, borderRadius: 14, padding: 13, marginBottom: 9,
              border: `1px solid ${c.id === currentId ? T.violet + "40" : T.hair}`, cursor: "pointer",
            }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                <div style={{ fontWeight: 650, color: T.ink, fontSize: 14, lineHeight: 1.35 }}>{c.name}</div>
                <button onClick={(e) => { e.stopPropagation(); onDelete(c.id); }}
                  style={{ background: "transparent", border: "none", cursor: "pointer", color: T.faint, flexShrink: 0, padding: 2 }}><Trash2 size={15} /></button>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6 }}>
                <Pill T={T} tone="muted">{c.campaignType}</Pill>
                <span style={{ fontSize: 11.5, color: T.faint }}>{new Date(c.updatedAt).toLocaleDateString("en-US")}</span>
              </div>
            </div>
          ))}
        </div>
      </div>
    </>
  );
}

/* ================================== APP ================================= */
const emptyCampaign = () => ({
  id: uid(), name: "New campaign", createdAt: Date.now(), updatedAt: Date.now(),
  profile: { company: "", offering: "", category: "", competitors: "", market: "Brasil / LATAM", audience: "" },
  objective: "", campaignType: "Anchor", contentLang: "PT", generatedLang: null,
  plan: { thesis: "", whyNow: [], personas: [], channelMix: [], messageArchitecture: [], assetPlan: [], kpiTargets: [] },
  localization: { language: "", tone: "", localChannels: [], compliance: [], culturalNotes: [], calendarNotes: [] },
  actualsRows: [], insights: null,
});

function hasGeneratedContent(c) {
  if (!c) return false;
  const p = c.plan || {};
  const L = c.localization || {};
  return Boolean(
    p.thesis || (p.personas || []).length || (p.channelMix || []).length ||
    (p.assetPlan || []).length || L.language || (L.localChannels || []).length || c.insights
  );
}

function TranslateBar({ campaign, onTranslate, translating, T }) {
  const target = campaign.contentLang || "PT";
  const current = campaign.generatedLang;
  if (!current || current === target || !hasGeneratedContent(campaign)) return null;
  const nameOf = (x) => (x === "EN" ? "English" : "Portuguese");
  return (
    <Card T={T} pad={16} style={{ marginBottom: 18, borderLeft: `3px solid ${T.coral}` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <Languages size={18} style={{ color: T.coral, flexShrink: 0 }} />
        <div style={{ flex: 1, minWidth: 220 }}>
          <div style={{ fontWeight: 650, color: T.ink, fontSize: 14 }}>
            Content is in {nameOf(current)}, output language is set to {nameOf(target)}.
          </div>
          <div style={{ fontSize: 12.5, color: T.muted, marginTop: 2 }}>
            Translation did not run automatically. Retry it here.
          </div>
        </div>
        {translating
          ? <Spinner T={T} label={"Translating to " + nameOf(target) + "…"} />
          : <Btn T={T} size="sm" icon={Languages} onClick={onTranslate}>Translate to {nameOf(target)}</Btn>}
      </div>
    </Card>
  );
}

/* Saved campaigns can predate schema changes. Normalizing on load prevents a
   missing field from crashing a phase that reads it directly. */
function normalizeCampaign(c) {
  const base = emptyCampaign();
  if (!c || typeof c !== "object") return base;
  return {
    ...base,
    ...c,
    profile: { ...base.profile, ...(c.profile || {}) },
    plan: { ...base.plan, ...(c.plan || {}) },
    localization: { ...base.localization, ...(c.localization || {}) },
    actualsRows: Array.isArray(c.actualsRows) ? c.actualsRows : [],
    insights: c.insights || null,
    contentLang: c.contentLang === "EN" ? "EN" : "PT",
    id: c.id || base.id,
  };
}

const PHASES = ["Plan", "Localize", "Measure", "Optimize"];
const PHASE_ICONS = [Target, Globe, BarChart3, TrendingUp];

export default function App() {
  const [dark, setDark] = useState(false);
  const T = useMemo(() => tokens(dark), [dark]);
  const [view, setView] = useState("home"); // home | workspace
  const [campaign, setCampaign] = useState(null);
  const [phase, setPhase] = useState(PHASES[0]);
  const [saved, setSaved] = useState([]);
  const [drawer, setDrawer] = useState(false);
  const [toast, setToast] = useState(null);
  const [translating, setTranslating] = useState(false);

  useEffect(() => { loadCampaigns().then(setSaved); }, []);
  const toastTimer = useRef(null);
  const notify = (m, isError) => {
    if (toastTimer.current) clearTimeout(toastTimer.current);
    setToast({ msg: m, error: Boolean(isError) });
    toastTimer.current = setTimeout(() => setToast(null), isError ? 5200 : 2200);
  };

  const openCampaign = (c) => { setCampaign(normalizeCampaign(c)); setPhase(PHASES[0]); setView("workspace"); setDrawer(false); };
  const newCampaign = () => { openCampaign(emptyCampaign()); };
  const openDemo = () => { openCampaign(buildDemo()); };

  // PDF export forces the light theme so the printed document is never dark.
  function exportPDF() {
    const wasDark = dark;
    // The browser derives the suggested PDF filename from document.title, so the
    // export is named after the campaign instead of the app.
    const previousTitle = document.title;
    const safeName = String((campaign && campaign.name) || "").trim().replace(/[\\/:*?"<>|]/g, "-");
    if (safeName) document.title = safeName;
    if (wasDark) setDark(false);
    setTimeout(() => {
      window.print();
      document.title = previousTitle;
      if (wasDark) setTimeout(() => setDark(true), 400);
    }, 150);
  }

  async function translateContent(explicitTarget) {
    if (!campaign || translating) return;
    const target = explicitTarget || campaign.contentLang || "PT";

    // The sample campaign ships in both languages, so it switches without a
    // network call. The pacing and messaging mirror the real path exactly.
    if (isPristineSample(campaign)) {
      setTranslating(true);
      setTimeout(() => {
        setCampaign((c) => ({ ...c, ...demoPatch(target), contentLang: target, generatedLang: target }));
        setTranslating(false);
        notify("Content translated.");
      }, 900);
      return;
    }

    setTranslating(true);
    try {
      const patch = {};
      const prof = campaign.profile || {};
      const p = campaign.plan || {};
      const L = campaign.localization || {};
      let planPatch = {};
      const failed = [];

      // Small batches: each request must comfortably fit the model's output budget,
      // otherwise the JSON comes back truncated and the section is silently lost.
      const run = async (label, payload, apply) => {
        const populated = Object.keys(payload).some((k) => {
          const v = payload[k];
          return Array.isArray(v) ? v.length > 0 : Boolean(v);
        });
        if (!populated) return;
        const tr = await translateSection(payload, target);
        if (tr) apply(tr); else failed.push(label);
      };

      await run("profile",
        { offering: prof.offering, category: prof.category, audience: prof.audience, objective: campaign.objective },
        (tr) => {
          patch.profile = {
            ...prof,
            offering: tr.offering || prof.offering,
            category: tr.category || prof.category,
            audience: tr.audience || prof.audience,
          };
          patch.objective = tr.objective || campaign.objective;
        });

      await run("thesis",
        { thesis: p.thesis, whyNow: p.whyNow || [] },
        (tr) => { planPatch = { ...planPatch, thesis: tr.thesis || p.thesis, whyNow: tr.whyNow || p.whyNow }; });

      await run("personas",
        { personas: p.personas || [] },
        (tr) => { planPatch = { ...planPatch, personas: tr.personas || p.personas }; });

      await run("channel mix",
        { channelMix: p.channelMix || [] },
        (tr) => { planPatch = { ...planPatch, channelMix: tr.channelMix || p.channelMix }; });

      await run("message architecture",
        { messageArchitecture: p.messageArchitecture || [] },
        (tr) => { planPatch = { ...planPatch, messageArchitecture: tr.messageArchitecture || p.messageArchitecture }; });

      await run("asset plan",
        { assetPlan: p.assetPlan || [] },
        (tr) => { planPatch = { ...planPatch, assetPlan: tr.assetPlan || p.assetPlan }; });

      if (Object.keys(planPatch).length) {
        patch.plan = { ...p, ...planPatch, kpiTargets: p.kpiTargets }; // targets are numbers: never sent
      }

      await run("localization",
        { language: L.language, tone: L.tone, localChannels: L.localChannels || [], compliance: L.compliance || [],
          culturalNotes: L.culturalNotes || [], calendarNotes: L.calendarNotes || [] },
        (tr) => { patch.localization = { ...L, ...tr }; });

      const ins = campaign.insights;
      if (ins) {
        await run("executive summary",
          { summary: ins.summary, gaps: ins.gaps || [] },
          (tr) => { patch.insights = { ...ins, summary: tr.summary || ins.summary, gaps: tr.gaps || ins.gaps }; });
        await run("recommendations",
          { recommendations: ins.recommendations || [] },
          (tr) => { patch.insights = { ...(patch.insights || ins), recommendations: tr.recommendations || ins.recommendations }; });
      }

      if (!Object.keys(patch).length) {
        throw new Error("Nothing could be translated. Please try again.");
      }

      // Functional update: never clobber state changed while translating.
      setCampaign((c) => ({ ...c, ...patch, contentLang: target, generatedLang: target }));
      notify(failed.length
        ? "Translated, except: " + failed.join(", ") + ". Switch language again to retry."
        : "Content translated.", failed.length > 0);
    } catch (e) {
      notify((e && e.message) || "Translation failed. Try again.", true);
    } finally {
      setTranslating(false);
    }
  }

  async function save() {
    if (!campaign) return;
    const c = { ...campaign, updatedAt: Date.now() };
    setCampaign(c);
    const mode = await persistCampaign(c);
    const list = await loadCampaigns();
    setSaved(list);
    notify(mode === "durable"
      ? "Campaign saved."
      : "Campaign saved for this session. Export to PDF to keep a copy.");
  }
  async function del(id) {
    await removeCampaign(id);
    setSaved(await loadCampaigns());
    if (campaign && campaign.id === id) { setView("home"); setCampaign(null); }
    notify("Campaign removed.");
  }

  const link = "https://www.linkedin.com/in/carloseduardovf/";

  return (
    <div style={{ minHeight: "100vh", background: T.bgAtmo, fontFamily: T.font, color: T.ink, WebkitFontSmoothing: "antialiased" }}>
      <style>{`
        @keyframes cdspin { to { transform: rotate(360deg); } }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-thumb { background: ${T.hair}; border-radius: 6px; }
        .ici-grid2 { grid-template-columns: 1fr 1fr; }
        .ici-grid21 { grid-template-columns: 2fr 1fr; }
        .ici-grid3 { grid-template-columns: repeat(3, 1fr); }
        .ici-grid4 { grid-template-columns: repeat(4, 1fr); }
        @media (max-width: 900px) {
          .ici-grid4 { grid-template-columns: 1fr 1fr; }
        }
        @media (max-width: 680px) {
          .ici-grid2, .ici-grid21, .ici-grid3, .ici-grid4 { grid-template-columns: 1fr; }
        }
        @page { size: A4; margin: 14mm; }
        .print-only { display: none; }
        @media print {
          .no-print { display: none !important; }
          .print-only { display: block !important; }
          body { background: #fff !important; }
          * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
      `}</style>

      {/* Header: 3 zones (brand, phase, actions, signature) */}
      <header className="no-print" style={{
        position: "sticky", top: 0, zIndex: 30, background: T.frost, backdropFilter: "blur(20px)",
        WebkitBackdropFilter: "blur(20px)", borderBottom: `1px solid ${T.hair}`,
      }}>
        <div style={{ maxWidth: 1080, margin: "0 auto", padding: "0 20px", height: 60, display: "grid", gridTemplateColumns: "1fr auto 1fr", alignItems: "center", gap: 12 }}>
          {/* zone 1: brand + tool name */}
          <div style={{ display: "flex", alignItems: "center", gap: 11, justifySelf: "start" }}>
            <button onClick={() => setDrawer(true)} style={{ background: T.chip, border: "none", borderRadius: 10, width: 36, height: 36, display: "grid", placeItems: "center", cursor: "pointer", color: T.ink }}><Menu size={18} /></button>
            <div style={{ display: "flex", alignItems: "center", gap: 9, cursor: "pointer" }} onClick={() => setView("home")}>
              <div style={{ width: 30, height: 30, borderRadius: 9, background: T.grad, display: "grid", placeItems: "center", boxShadow: "0 4px 12px rgba(110,58,255,0.4)" }}>
                <Sparkles size={16} color="#fff" />
              </div>
              <div style={{ lineHeight: 1.1, maxWidth: 220 }}>
                <div style={{ fontWeight: 800, fontSize: 15, letterSpacing: -0.3 }}>
                  Integrated Campaign <span style={{ background: T.grad, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text" }}>Intelligence</span>
                </div>
              </div>
            </div>
          </div>
          {/* zone 2: phase nav (only in workspace) */}
          <div style={{ justifySelf: "center" }}>
            {view === "workspace" ? (
              <div className="cd-phasenav">
                <Segmented options={PHASES} value={phase} onChange={setPhase} T={T} icons={PHASE_ICONS} />
              </div>
            ) : null}
          </div>
          {/* zone 3: actions + signature */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, justifySelf: "end" }}>
            {view === "workspace" ? (
              <>
                <Btn T={T} size="sm" variant="ghost" icon={Save} onClick={save}>Save</Btn>
                <Btn T={T} size="sm" variant="outline" icon={Download} onClick={exportPDF}>PDF</Btn>
              </>
            ) : null}
            <button onClick={() => setDark((d) => !d)} style={{ background: T.chip, border: "none", borderRadius: 10, width: 36, height: 36, display: "grid", placeItems: "center", cursor: "pointer", color: T.ink }}>
              {dark ? <Sun size={17} /> : <Moon size={17} />}
            </button>
            <a href={link} target="_blank" rel="noreferrer" style={{ textDecoration: "none", display: "flex", alignItems: "center", gap: 5, fontSize: 11.5, color: T.faint, fontWeight: 600 }} className="cd-sig">
              <span>by Carlos Eduardo</span><ExternalLink size={11} />
            </a>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1080, margin: "0 auto", padding: "26px 20px 80px" }}>
        {view === "home" ? (
          <Home T={T} onNew={newCampaign} onDemo={openDemo} saved={saved} onOpen={openCampaign} />
        ) : campaign ? (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
              <input value={campaign.name} maxLength={90} placeholder="Campaign name"
                onChange={(e) => setCampaign((c) => ({ ...c, name: e.target.value }))}
                style={{ background: "transparent", border: "none", outline: "none", color: T.ink, fontSize: 26, fontWeight: 800, letterSpacing: -0.6, fontFamily: T.font, flex: 1, minWidth: 240 }} />
              <div className="no-print" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 11.5, fontWeight: 600, color: T.faint, letterSpacing: 0.3 }}>Output language</span>
                <Segmented options={["PT", "EN"]} value={campaign.contentLang || "PT"}
                  onChange={(v) => {
                    if (v === (campaign.contentLang || "PT")) return;
                    setCampaign((c) => ({ ...c, contentLang: v }));
                    if (hasGeneratedContent(campaign) && campaign.generatedLang && campaign.generatedLang !== v) {
                      translateContent(v);
                    }
                  }} T={T} />
              </div>
            </div>
            {/* phase nav for mobile (below title) */}
            <div className="cd-phasenav-mobile no-print" style={{ display: "none", marginBottom: 18, overflowX: "auto" }}>
              <Segmented options={PHASES} value={phase} onChange={setPhase} T={T} icons={PHASE_ICONS} />
            </div>

            <div className="no-print">
              <TranslateBar campaign={campaign} onTranslate={translateContent} translating={translating} T={T} />
            </div>

            {phase === "Plan" && <PlanPhase campaign={campaign} setCampaign={setCampaign} T={T} notify={notify} />}
            {phase === "Localize" && <LocalizePhase campaign={campaign} setCampaign={setCampaign} T={T} notify={notify} />}
            {phase === "Measure" && <MeasurePhase campaign={campaign} setCampaign={setCampaign} T={T} notify={notify} />}
            {phase === "Optimize" && <OptimizePhase campaign={campaign} setCampaign={setCampaign} T={T} notify={notify} />}
          </div>
        ) : null}
      </main>

      <div className="print-only" style={{ marginTop: 24, paddingTop: 12, borderTop: "1px solid #E5E5EA", fontSize: 11, color: "#6E6E73", textAlign: "center" }}>
        {campaign ? campaign.name + " · " : ""}Integrated Campaign Intelligence · Developed by Carlos Eduardo · linkedin.com/in/carloseduardovf
      </div>

      <footer className="no-print" style={{ maxWidth: 1080, margin: "0 auto", padding: "0 20px 34px", textAlign: "center" }}>
        <div style={{ fontSize: 12, color: T.faint }}>
          Integrated Campaign Intelligence · <a href={link} target="_blank" rel="noreferrer" style={{ color: T.violet, textDecoration: "none", fontWeight: 600 }}>Developed by Carlos Eduardo</a>
        </div>
      </footer>

      <div className="no-print"><Drawer open={drawer} onClose={() => setDrawer(false)} campaigns={saved} onOpen={openCampaign} onDelete={del} onNew={newCampaign} T={T} currentId={campaign ? campaign.id : null} /></div>

      {toast ? (
        <div className="no-print" style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", zIndex: 60,
          background: T.dark ? "#2C2C2E" : "#1D1D1F", color: "#fff", padding: "11px 20px", borderRadius: 14,
          fontSize: 14, fontWeight: 600, boxShadow: "0 12px 32px rgba(0,0,0,0.35)", display: "flex", alignItems: "flex-start", gap: 8,
          maxWidth: "min(92vw, 460px)", lineHeight: 1.45 }}>
          {toast.error
            ? <AlertTriangle size={16} color={T.warn} style={{ flexShrink: 0, marginTop: 2 }} />
            : <CheckCircle2 size={16} color={T.good} style={{ flexShrink: 0, marginTop: 2 }} />}
          <span>{toast.msg}</span>
        </div>
      ) : null}

      <style>{`
        @media (max-width: 820px) {
          .cd-phasenav { display: none; }
          .cd-phasenav-mobile { display: block !important; }
          .cd-sig span { display: none; }
        }
      `}</style>
    </div>
  );
}

/* -------------------------------- Home ---------------------------------- */
function Home({ T, onNew, onDemo, saved, onOpen }) {
  const steps = [
    { icon: Target, t: "Plan", d: "Thesis, personas, multichannel mix, message architecture, and KPI targets." },
    { icon: Globe, t: "Localize", d: "Real LATAM adaptation: local channels, LGPD, tone, and calendar." },
    { icon: BarChart3, t: "Measure", d: "Upload the Excel extract. The app computes CPL, ROAS, funnel, and pipeline." },
    { icon: TrendingUp, t: "Optimize", d: "Plan vs. actual and actionable recommendations by priority." },
  ];
  return (
    <div>
      {/* Hero: lifecycle spine as thesis */}
      <div style={{ textAlign: "center", padding: "34px 0 10px" }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 7, background: T.violet + "14", color: T.violet, padding: "6px 14px", borderRadius: 999, fontSize: 12.5, fontWeight: 700, marginBottom: 20 }}>
          <Zap size={14} /> The full campaign lifecycle, in one place
        </div>
        <h1 style={{ fontSize: 46, fontWeight: 850, letterSpacing: -1.6, lineHeight: 1.04, margin: "0 auto 14px", maxWidth: 720 }}>
          From <span style={{ background: T.grad, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text" }}>plan</span> to actual, without breaking the loop.
        </h1>
        <p style={{ fontSize: 17, color: T.muted, maxWidth: 560, margin: "0 auto 26px", lineHeight: 1.5 }}>
          The platform plans integrated campaigns, localizes for the region, reads real data, and recommends optimization. The app computes every number; the AI only writes.
        </p>
        <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap" }}>
          <Btn T={T} icon={Plus} onClick={onNew}>New campaign</Btn>
          <Btn T={T} variant="outline" icon={Rocket} onClick={onDemo}>Load sample campaign</Btn>
        </div>
      </div>

      {/* lifecycle spine */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginTop: 40 }} className="cd-steps">
        {steps.map((s, i) => (
          <Card T={T} key={i} pad={20}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
              <div style={{ width: 34, height: 34, borderRadius: 10, background: T.grad, display: "grid", placeItems: "center" }}>
                <s.icon size={17} color="#fff" />
              </div>
              <ArrowRight size={15} style={{ color: T.faint, marginLeft: "auto" }} />
            </div>
            <div style={{ fontWeight: 750, fontSize: 16, marginBottom: 5, letterSpacing: -0.2 }}>{s.t}</div>
            <div style={{ fontSize: 13, color: T.muted, lineHeight: 1.5 }}>{s.d}</div>
          </Card>
        ))}
      </div>

      {saved.length ? (
        <div style={{ marginTop: 40 }}>
          <div style={{ fontSize: 13, fontWeight: 700, letterSpacing: 0.8, textTransform: "uppercase", color: T.faint, marginBottom: 14 }}>Continue</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 12 }}>
            {saved.slice(0, 6).map((c) => (
              <Card T={T} key={c.id} pad={16} style={{ cursor: "pointer" }}>
                <div onClick={() => onOpen(c)}>
                  <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 8, lineHeight: 1.3 }}>{c.name}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <Pill T={T} tone="muted">{c.campaignType}</Pill>
                    <span style={{ fontSize: 12, color: T.faint, marginLeft: "auto" }}>{new Date(c.updatedAt).toLocaleDateString("en-US")}</span>
                    <ChevronRight size={15} style={{ color: T.faint }} />
                  </div>
                </div>
              </Card>
            ))}
          </div>
        </div>
      ) : null}

      <style>{`@media (max-width: 820px){ .cd-steps { grid-template-columns: 1fr 1fr !important; } } @media (max-width:520px){ .cd-steps { grid-template-columns: 1fr !important; } }`}</style>
    </div>
  );
}
