import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, ReferenceLine,
} from "recharts";
import {
  LayoutDashboard, CalendarRange, PiggyBank, Layers, BookOpen, TrendingUp,
  TrendingDown, Filter, X, Plus, Pencil, Trash2, Save, RotateCcw, Search,
  ArrowUpDown, Wallet, Target, AlertTriangle, Info, Check, Circle, ChevronRight,
  SlidersHorizontal, Workflow, CalendarDays, BarChart3, Briefcase, HandCoins, Clock,
  Users, Repeat, ClipboardList, UploadCloud, CheckSquare, Square, Menu, ChevronDown,
  Download, Printer, Bell, Sparkles, Gauge, ArrowRight, Percent, Upload, Mail,
  FileSpreadsheet, FileText, Loader2, Minus, GitCompare, HelpCircle, PieChart as PieChartIcon, Activity,
} from "lucide-react";

// ============================================================
// DESIGN TOKENS — "grand livre" ledger aesthetic
// ============================================================
const COLOR = {
  bg: "#0e1611",
  surface: "#151f19",
  surfaceRaised: "#1b2620",
  surfaceInput: "#101a14",
  hairline: "#2a362e",
  ink: "#eef2ea",
  inkMuted: "#7c8a7f",
  gold: "#c9a227",
  goldSoft: "#e0c15a",
  emerald: "#3f9c7a",
  emeraldSoft: "#5fc298",
  clay: "#c1543f",
  claySoft: "#dd7b64",
  slateBlue: "#5b7ea6",
  slateBlueSoft: "#8badd1",
  violet: "#8b7bc2",
  violetSoft: "#ab9fd6",
};

const fontImport = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
@media print {
  .gl-noprint { display: none !important; }
  .gl-print-full { max-width: 100% !important; }
  body { background: white !important; }
}
.gl-scroll::-webkit-scrollbar { height: 5px; width: 5px; }
.gl-scroll::-webkit-scrollbar-thumb { background: #2a362e; border-radius: 4px; }
@supports (padding: max(0px)) {
  .gl-safe-bottom { padding-bottom: max(10px, env(safe-area-inset-bottom)); }
  .gl-safe-top { padding-top: max(0px, env(safe-area-inset-top)); }
}
input[type=number]::-webkit-inner-spin-button,
input[type=number]::-webkit-outer-spin-button {
  -webkit-appearance: none;
  margin: 0;
}
input[type=number] {
  -moz-appearance: textfield;
}
@keyframes spin {
  from { transform: rotate(0deg); }
  to { transform: rotate(360deg); }
}
`;

function useIsMobile(breakpoint = 860) {
  const [isMobile, setIsMobile] = useState(() => (typeof window !== "undefined" ? window.innerWidth < breakpoint : false));
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < breakpoint);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, [breakpoint]);
  return isMobile;
}

// ============================================================
// SYNCHRONISATION CLOUD (Supabase) — désactivée automatiquement si les
// variables d'environnement VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY
// ne sont pas définies (ex: aperçu dans Claude). Aucune dépendance ajoutée :
// simples appels fetch() vers l'API REST de Supabase (PostgREST).
// ============================================================
function getEnvVar(name: string): string {
  try {
    // @ts-ignore — import.meta.env n'existe que dans un contexte Vite
    return (typeof import.meta !== "undefined" && (import.meta as any).env && (import.meta as any).env[name]) || "";
  } catch { return ""; }
}
const SUPABASE_URL = getEnvVar("VITE_SUPABASE_URL");
const SUPABASE_ANON_KEY = getEnvVar("VITE_SUPABASE_ANON_KEY");
const SYNC_ENABLED = !!(SUPABASE_URL && SUPABASE_ANON_KEY);

async function fetchRemoteState(syncCode: string): Promise<any | null> {
  if (!SYNC_ENABLED || !syncCode) return null;
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/app_state?sync_code=eq.${encodeURIComponent(syncCode)}&select=data,updated_at`, {
      headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}` },
    });
    if (!res.ok) return null;
    const rows = await res.json();
    return rows?.[0]?.data || null;
  } catch { return null; }
}
async function pushRemoteState(syncCode: string, data: any): Promise<boolean> {
  if (!SYNC_ENABLED || !syncCode) return false;
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/app_state?on_conflict=sync_code`, {
      method: "POST",
      headers: {
        apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        "Content-Type": "application/json", Prefer: "resolution=merge-duplicates",
      },
      body: JSON.stringify({ sync_code: syncCode, data, updated_at: new Date().toISOString() }),
    });
    return res.ok;
  } catch { return false; }
}

// Abonnement en temps réel (WebSocket) aux changements de la ligne sync_code, via
// @supabase/supabase-js chargé dynamiquement — n'existe pas dans l'aperçu Claude (fallback
// silencieux vers le mode "pull au chargement + push différé" déjà en place, qui continue
// de fonctionner seul). Retourne une fonction de désabonnement, ou null si indisponible.
let realtimeClientPromise: Promise<any> | null = null;
async function getRealtimeClient(): Promise<any | null> {
  if (!SYNC_ENABLED) return null;
  if (!realtimeClientPromise) {
    realtimeClientPromise = (async () => {
      try {
        const mod: any = await import(/* @vite-ignore */ "@supabase/supabase-js");
        return mod.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
      } catch {
        return null;
      }
    })();
  }
  return realtimeClientPromise;
}
async function subscribeRealtime(syncCode: string, onRemoteChange: () => void): Promise<(() => void) | null> {
  const client = await getRealtimeClient();
  if (!client || !syncCode) return null;
  try {
    const channel = client
      .channel(`app_state_${syncCode}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "app_state", filter: `sync_code=eq.${syncCode}` }, () => {
        onRemoteChange();
      })
      .subscribe();
    return () => { try { client.removeChannel(channel); } catch {} };
  } catch {
    return null;
  }
}


// ============================================================
// TYPES
// ============================================================
type TxType = "Dépense" | "Revenu";
type Group = "Nécessaire" | "Productif" | "Non-productif" | "Non classifié";
type Scope = "Personnel" | "Business";
type LoanStatus = "En attente" | "Remboursé";

interface Transaction {
  id: string;
  date: string; // "YYYY-MM-DD" — le mois est dérivé automatiquement pour tous les rapports
  time?: string; // "HH:MM" — heure d'enregistrement, optionnelle (absente sur les données historiques)
  category: string;
  subcategory?: string;
  type: TxType;
  amount: number;
  account?: string;
  payee?: string;
  note?: string;
  tags?: string;
}
interface Account {
  id: string;
  name: string;
  kind: "Espèces" | "Banque" | "Mobile Money" | "Carte de crédit" | "Autre";
  openingBalance: number; // solde au moment où le suivi dans l'app a commencé
}
interface CategoryBudget {
  id: string;
  category: string;
  amount: number;
  rollover: boolean;
}
interface Goal {
  id: string;
  name: string;
  target: number;
  current: number;
  date: string;
}
interface RecurringTemplate {
  id: string;
  category: string;
  type: TxType;
  amount: number;
  frequency: "Hebdomadaire" | "Mensuelle" | "Annuelle";
  nextDate: string;
  account?: string;
  payee?: string;
}
interface Loan {
  id: string;
  person: string;
  amount: number;
  dateGiven: string;
  status: LoanStatus;
  notes: string;
}
interface CategorizationRule {
  id: string;
  keyword: string;
  group: Group;
}

// ============================================================
// SEED DATA (633 transactions extraites des exports MoneyCoach)
// ============================================================
const seedTransactions: Transaction[] = [{"id":"t3026","date":"2024-06-01","category":"GRUNDFOS","type":"Dépense","amount":494562,"account":"SIB","payee":"-SIB","note":"Ajustement Petty cash"},{"id":"t3046","date":"2024-06-01","category":"Petty Cash","type":"Revenu","amount":494562,"account":"PETTY CASH","subcategory":"Ajustement Petty Cash"},{"id":"t3023","date":"2024-07-01","category":"Revenus Location Mazda","type":"Revenu","amount":730147,"account":"SIB"},{"id":"t3024","date":"2024-07-01","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t3025","date":"2024-07-01","category":"Ajustement","type":"Revenu","amount":31000,"account":"SIB"},{"id":"t3045","date":"2024-07-01","category":"Revenu général","type":"Revenu","amount":2630203,"account":"SIB","subcategory":"Solde 1er juillet 2024"},{"id":"t3005","date":"2024-07-31","category":"GRUNDFOS","type":"Dépense","amount":65000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t3006","date":"2024-07-31","category":"GRUNDFOS","type":"Dépense","amount":28200,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre"},{"id":"t3007","date":"2024-07-31","category":"Divertissement","type":"Dépense","amount":34500,"account":"SIB"},{"id":"t3008","date":"2024-07-31","category":"Utilitaires","type":"Dépense","amount":20000,"account":"SIB","subcategory":"la télé","note":"Canal"},{"id":"t3009","date":"2024-07-31","category":"Épargne","type":"Dépense","amount":147351,"account":"SIB"},{"id":"t3010","date":"2024-07-31","category":"Aliments","type":"Dépense","amount":80000,"account":"SIB"},{"id":"t3011","date":"2024-07-31","category":"Shopping","type":"Dépense","amount":58657,"account":"SIB"},{"id":"t3012","date":"2024-07-31","category":"Utilitaires","type":"Dépense","amount":50500,"account":"SIB","subcategory":"Téléphones","payee":"12","note":"Réparation Iphone"},{"id":"t3013","date":"2024-07-31","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t3014","date":"2024-07-31","category":"Voiture","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Lavage"},{"id":"t3015","date":"2024-07-31","category":"Invitation","type":"Dépense","amount":179000,"account":"SIB"},{"id":"t3016","date":"2024-07-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t3017","date":"2024-07-31","category":"Générales","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Police"},{"id":"t3018","date":"2024-07-31","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t3019","date":"2024-07-31","category":"Santé","type":"Dépense","amount":912,"account":"SIB","subcategory":"Assurance"},{"id":"t3020","date":"2024-07-31","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t3021","date":"2024-07-31","category":"Cadeaux","type":"Dépense","amount":127700,"account":"SIB"},{"id":"t3022","date":"2024-07-31","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2976","date":"2024-08-31","category":"GRUNDFOS","type":"Dépense","amount":15000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t2977","date":"2024-08-31","category":"GRUNDFOS","type":"Dépense","amount":155000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2978","date":"2024-08-31","category":"Ajustement","type":"Dépense","amount":32542,"account":"SIB"},{"id":"t2979","date":"2024-08-31","category":"Aliments","type":"Dépense","amount":359740,"account":"SIB"},{"id":"t2980","date":"2024-08-31","category":"Générales","type":"Dépense","amount":7874,"account":"SIB"},{"id":"t2981","date":"2024-08-31","category":"Ajustement","type":"Revenu","amount":90089,"account":"SIB"},{"id":"t2982","date":"2024-08-31","category":"Invitation","type":"Dépense","amount":127915,"account":"SIB"},{"id":"t2983","date":"2024-08-31","category":"Vêtements","type":"Dépense","amount":47000,"account":"SIB"},{"id":"t2984","date":"2024-08-31","category":"Logement","type":"Dépense","amount":550550,"account":"SIB","subcategory":"Location"},{"id":"t2985","date":"2024-08-31","category":"Enfants & Maman","type":"Dépense","amount":57200,"account":"SIB","subcategory":"Hemra"},{"id":"t2986","date":"2024-08-31","category":"Générales","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Police"},{"id":"t2987","date":"2024-08-31","category":"Vente Pompe","type":"Revenu","amount":1000500,"account":"SIB"},{"id":"t2988","date":"2024-08-31","category":"Shopping","type":"Dépense","amount":41000,"account":"SIB"},{"id":"t2989","date":"2024-08-31","category":"Revenus Location Mazda","type":"Revenu","amount":732311,"account":"SIB"},{"id":"t2990","date":"2024-08-31","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t2991","date":"2024-08-31","category":"Emprunt Bancaire","type":"Revenu","amount":6664569,"account":"SIB"},{"id":"t2992","date":"2024-08-31","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2993","date":"2024-08-31","category":"Santé","type":"Dépense","amount":22500,"account":"SIB"},{"id":"t2994","date":"2024-08-31","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t2995","date":"2024-08-31","category":"Dette","type":"Dépense","amount":797415,"account":"SIB"},{"id":"t2996","date":"2024-08-31","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2997","date":"2024-08-31","category":"Ajustement","type":"Revenu","amount":188323,"account":"SIB"},{"id":"t2998","date":"2024-08-31","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2999","date":"2024-08-31","category":"Divertissement","type":"Dépense","amount":360600,"account":"SIB"},{"id":"t3000","date":"2024-08-31","category":"Cadeaux","type":"Dépense","amount":278650,"account":"SIB"},{"id":"t3001","date":"2024-08-31","category":"Voiture","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Lavage"},{"id":"t3002","date":"2024-08-31","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t3003","date":"2024-08-31","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","subcategory":"Cotisations"},{"id":"t3004","date":"2024-08-31","category":"Abonnements","type":"Dépense","amount":12300,"account":"SIB","subcategory":"Spotify"},{"id":"t2941","date":"2024-09-30","category":"Payement Maison Bingerville","type":"Dépense","amount":3000000,"account":"SIB"},{"id":"t2942","date":"2024-09-30","category":"GRUNDFOS","type":"Dépense","amount":91731,"account":"PETTY CASH","subcategory":"Voyage","note":"Nairobi"},{"id":"t2943","date":"2024-09-30","category":"GRUNDFOS","type":"Dépense","amount":10953,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t2944","date":"2024-09-30","category":"Voiture","type":"Dépense","amount":14500,"account":"SIB","subcategory":"Lavage"},{"id":"t2945","date":"2024-09-30","category":"Achat Mazda","type":"Dépense","amount":6851100,"account":"SIB"},{"id":"t2946","date":"2024-09-30","category":"Voiture","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Installation GPS"},{"id":"t2947","date":"2024-09-30","category":"Ajustement","type":"Dépense","amount":2875,"account":"SIB"},{"id":"t2948","date":"2024-09-30","category":"Revenus Location Mazda","type":"Revenu","amount":732311,"account":"SIB"},{"id":"t2949","date":"2024-09-30","category":"Transport","type":"Dépense","amount":22000,"account":"SIB","subcategory":"Taxi"},{"id":"t2950","date":"2024-09-30","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2951","date":"2024-09-30","category":"Générales","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Police"},{"id":"t2952","date":"2024-09-30","category":"Ajustement","type":"Dépense","amount":17912,"account":"SIB"},{"id":"t2953","date":"2024-09-30","category":"Éducation","type":"Dépense","amount":270000,"account":"SIB","subcategory":"Nesher","note":"Scolarité"},{"id":"t2954","date":"2024-09-30","category":"Enfants & Maman","type":"Dépense","amount":55800,"account":"SIB","subcategory":"Nesher"},{"id":"t2955","date":"2024-09-30","category":"GRUNDFOS","type":"Dépense","amount":13270,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2956","date":"2024-09-30","category":"GRUNDFOS","type":"Dépense","amount":115000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2957","date":"2024-09-30","category":"Utilitaires","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Téléphones","note":"Chargeur"},{"id":"t2958","date":"2024-09-30","category":"Cadeaux","type":"Dépense","amount":163550,"account":"SIB"},{"id":"t2959","date":"2024-09-30","category":"Logement","type":"Dépense","amount":550550,"account":"SIB","subcategory":"Location"},{"id":"t2960","date":"2024-09-30","category":"Divertissement","type":"Dépense","amount":325500,"account":"SIB"},{"id":"t2961","date":"2024-09-30","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2962","date":"2024-09-30","category":"Utilitaires","type":"Dépense","amount":20000,"account":"SIB","subcategory":"la télé","note":"Canal"},{"id":"t2963","date":"2024-09-30","category":"Invitation","type":"Dépense","amount":50000,"account":"SIB"},{"id":"t2964","date":"2024-09-30","category":"Plan Éducation","type":"Dépense","amount":2457,"account":"SIB"},{"id":"t2965","date":"2024-09-30","category":"Petty Cash","type":"Revenu","amount":2500000,"account":"SIB"},{"id":"t2966","date":"2024-09-30","category":"Ajustement","type":"Revenu","amount":7666,"account":"SIB"},{"id":"t2967","date":"2024-09-30","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2968","date":"2024-09-30","category":"Dette","type":"Dépense","amount":631166,"account":"SIB"},{"id":"t2969","date":"2024-09-30","category":"Aliments","type":"Dépense","amount":267231,"account":"SIB"},{"id":"t2970","date":"2024-09-30","category":"Enfants & Maman","type":"Dépense","amount":69500,"account":"SIB","subcategory":"Hemra"},{"id":"t2971","date":"2024-09-30","category":"Voiture","type":"Dépense","amount":112000,"account":"SIB","subcategory":"Entretien","note":"Visite technique"},{"id":"t2972","date":"2024-09-30","category":"Santé","type":"Dépense","amount":500,"account":"SIB"},{"id":"t2973","date":"2024-09-30","category":"Utilitaires","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Téléphones","note":"Móbile"},{"id":"t2974","date":"2024-09-30","category":"Loyer","type":"Revenu","amount":1500000,"account":"SIB"},{"id":"t2975","date":"2024-09-30","category":"Shopping","type":"Dépense","amount":62500,"account":"SIB"},{"id":"t2937","date":"2024-10-30","category":"Revenus Location Mazda","type":"Revenu","amount":946800,"account":"SIB"},{"id":"t2938","date":"2024-10-30","category":"Revenus Location Mazda","type":"Revenu","amount":732311,"account":"SIB"},{"id":"t2939","date":"2024-10-30","category":"Revenus Location Mazda","type":"Revenu","amount":18000,"account":"SIB"},{"id":"t2940","date":"2024-10-30","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2909","date":"2024-10-31","category":"GRUNDFOS","type":"Dépense","amount":125000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2910","date":"2024-10-31","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2911","date":"2024-10-31","category":"Générales","type":"Dépense","amount":2520,"account":"SIB","subcategory":"Police"},{"id":"t2912","date":"2024-10-31","category":"Ajustement","type":"Dépense","amount":41475,"account":"SIB"},{"id":"t2913","date":"2024-10-31","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2914","date":"2024-10-31","category":"Santé","type":"Dépense","amount":13500,"account":"SIB"},{"id":"t2915","date":"2024-10-31","category":"Logement","type":"Dépense","amount":551500,"account":"SIB","subcategory":"Location"},{"id":"t2916","date":"2024-10-31","category":"Invitation","type":"Dépense","amount":125500,"account":"SIB"},{"id":"t2917","date":"2024-10-31","category":"GRUNDFOS","type":"Dépense","amount":76230,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2918","date":"2024-10-31","category":"Shopping","type":"Dépense","amount":140500,"account":"SIB"},{"id":"t2919","date":"2024-10-31","category":"Enfants & Maman","type":"Dépense","amount":67935,"account":"SIB","subcategory":"Hemra"},{"id":"t2920","date":"2024-10-31","category":"Personnel","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2921","date":"2024-10-31","category":"Cadeaux","type":"Dépense","amount":464080,"account":"SIB"},{"id":"t2922","date":"2024-10-31","category":"Générales","type":"Dépense","amount":1000,"account":"SIB"},{"id":"t2923","date":"2024-10-31","category":"Générales","type":"Dépense","amount":11284,"account":"SIB"},{"id":"t2924","date":"2024-10-31","category":"Transport","type":"Dépense","amount":84360,"account":"SIB","subcategory":"Taxi"},{"id":"t2925","date":"2024-10-31","category":"Divertissement","type":"Dépense","amount":340075,"account":"SIB"},{"id":"t2926","date":"2024-10-31","category":"Achat Mazda","type":"Dépense","amount":456000,"account":"SIB","note":"Clé Mazda"},{"id":"t2927","date":"2024-10-31","category":"Utilitaires","type":"Dépense","amount":17000,"account":"SIB","subcategory":"Téléphones","note":"Batterie"},{"id":"t2928","date":"2024-10-31","category":"Générales","type":"Dépense","amount":16600,"account":"SIB"},{"id":"t2929","date":"2024-10-31","category":"Voiture","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Lavage"},{"id":"t2930","date":"2024-10-31","category":"Aliments","type":"Dépense","amount":124170,"account":"SIB"},{"id":"t2931","date":"2024-10-31","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t2932","date":"2024-10-31","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t2933","date":"2024-10-31","category":"Dette","type":"Dépense","amount":631166,"account":"SIB"},{"id":"t2934","date":"2024-10-31","category":"Utilitaires","type":"Dépense","amount":16500,"account":"SIB","subcategory":"Téléphones","note":"Móbile"},{"id":"t2935","date":"2024-10-31","category":"Enfants & Maman","type":"Dépense","amount":50500,"account":"SIB","subcategory":"Nesher"},{"id":"t2936","date":"2024-10-31","category":"Voiture","type":"Dépense","amount":93500,"account":"SIB","subcategory":"Entretien"},{"id":"t2908","date":"2024-11-29","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2880","date":"2024-11-30","category":"GRUNDFOS","type":"Dépense","amount":150000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2881","date":"2024-11-30","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2882","date":"2024-11-30","category":"Shopping","type":"Dépense","amount":96000,"account":"SIB"},{"id":"t2883","date":"2024-11-30","category":"Utilitaires","type":"Dépense","amount":20000,"account":"SIB","subcategory":"la télé","note":"Canal"},{"id":"t2884","date":"2024-11-30","category":"Générales","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Police"},{"id":"t2885","date":"2024-11-30","category":"Pack Club","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2886","date":"2024-11-30","category":"Dette","type":"Dépense","amount":631166,"account":"SIB"},{"id":"t2887","date":"2024-11-30","category":"Personnel","type":"Dépense","amount":210000,"account":"SIB","subcategory":"Produits de beauté","note":"Parfum"},{"id":"t2888","date":"2024-11-30","category":"Utilitaires","type":"Dépense","amount":11500,"account":"SIB","subcategory":"Téléphones","note":"Móbile"},{"id":"t2889","date":"2024-11-30","category":"Enfants & Maman","type":"Dépense","amount":65000,"account":"SIB","subcategory":"Hemra"},{"id":"t2890","date":"2024-11-30","category":"Vêtements","type":"Dépense","amount":73020,"account":"SIB"},{"id":"t2891","date":"2024-11-30","category":"Logement","type":"Dépense","amount":560000,"account":"SIB","subcategory":"Location"},{"id":"t2892","date":"2024-11-30","category":"Divertissement","type":"Dépense","amount":298750,"account":"SIB"},{"id":"t2893","date":"2024-11-30","category":"Enfants & Maman","type":"Dépense","amount":60600,"account":"SIB","subcategory":"Nesher"},{"id":"t2894","date":"2024-11-30","category":"Invitation","type":"Dépense","amount":240500,"account":"SIB"},{"id":"t2895","date":"2024-11-30","category":"Vêtements","type":"Dépense","amount":13000,"account":"SIB"},{"id":"t2896","date":"2024-11-30","category":"Ajustement","type":"Dépense","amount":68147,"account":"SIB"},{"id":"t2897","date":"2024-11-30","category":"Voiture","type":"Dépense","amount":16500,"account":"SIB","subcategory":"Lavage"},{"id":"t2898","date":"2024-11-30","category":"Personnel","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Coiffure"},{"id":"t2899","date":"2024-11-30","category":"Âge D'or Retraite","type":"Dépense","amount":30000,"account":"SIB"},{"id":"t2900","date":"2024-11-30","category":"Santé","type":"Dépense","amount":14500,"account":"SIB"},{"id":"t2901","date":"2024-11-30","category":"GRUNDFOS","type":"Dépense","amount":3500,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t2902","date":"2024-11-30","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2903","date":"2024-11-30","category":"Transport","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Péage"},{"id":"t2904","date":"2024-11-30","category":"Cadeaux","type":"Dépense","amount":180850,"account":"SIB"},{"id":"t2905","date":"2024-11-30","category":"Aliments","type":"Dépense","amount":212500,"account":"SIB"},{"id":"t2906","date":"2024-11-30","category":"Voiture","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Entretien"},{"id":"t2907","date":"2024-11-30","category":"Cadeaux","type":"Dépense","amount":7500,"account":"SIB","note":"Aurore"},{"id":"t2875","date":"2024-12-30","category":"Vente Pompe","type":"Revenu","amount":721307,"account":"SIB"},{"id":"t2876","date":"2024-12-30","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"SIB"},{"id":"t2877","date":"2024-12-30","category":"Petty Cash","type":"Revenu","amount":2000000,"account":"SIB"},{"id":"t2878","date":"2024-12-30","category":"Loyer","type":"Revenu","amount":1500000,"account":"SIB"},{"id":"t2879","date":"2024-12-30","category":"Un salaire","type":"Revenu","amount":2391204,"account":"SIB"},{"id":"t2849","date":"2024-12-31","category":"GRUNDFOS","type":"Dépense","amount":85000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2850","date":"2024-12-31","category":"GRUNDFOS","type":"Dépense","amount":7000,"account":"PETTY CASH","subcategory":"Internet","note":"Móbile"},{"id":"t2851","date":"2024-12-31","category":"Divertissement","type":"Dépense","amount":340000,"account":"SIB"},{"id":"t2852","date":"2024-12-31","category":"Enfants & Maman","type":"Dépense","amount":34000,"account":"SIB","subcategory":"Hemra"},{"id":"t2853","date":"2024-12-31","category":"Générales","type":"Dépense","amount":5000,"account":"SIB"},{"id":"t2854","date":"2024-12-31","category":"Enfants & Maman","type":"Dépense","amount":77000,"account":"SIB","note":"Jeux"},{"id":"t2855","date":"2024-12-31","category":"GRUNDFOS","type":"Dépense","amount":203000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2856","date":"2024-12-31","category":"Cadeaux","type":"Dépense","amount":517020,"account":"SIB"},{"id":"t2857","date":"2024-12-31","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2858","date":"2024-12-31","category":"Voiture","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Entretien"},{"id":"t2859","date":"2024-12-31","category":"Invitation","type":"Dépense","amount":21000,"account":"SIB"},{"id":"t2860","date":"2024-12-31","category":"Voiture","type":"Dépense","amount":14100,"account":"SIB","subcategory":"Lavage"},{"id":"t2861","date":"2024-12-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t2862","date":"2024-12-31","category":"Plan Éducation","type":"Dépense","amount":64914,"account":"SIB"},{"id":"t2863","date":"2024-12-31","category":"Enfants & Maman","type":"Dépense","amount":50500,"account":"SIB","subcategory":"Nesher"},{"id":"t2864","date":"2024-12-31","category":"Dette","type":"Dépense","amount":740573,"account":"SIB"},{"id":"t2865","date":"2024-12-31","category":"Utilitaires","type":"Dépense","amount":20000,"account":"SIB","subcategory":"la télé","note":"Canal"},{"id":"t2866","date":"2024-12-31","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2867","date":"2024-12-31","category":"Shopping","type":"Dépense","amount":203000,"account":"SIB"},{"id":"t2868","date":"2024-12-31","category":"Logement","type":"Dépense","amount":550500,"account":"SIB","subcategory":"Location"},{"id":"t2869","date":"2024-12-31","category":"Ajustement","type":"Dépense","amount":90161,"account":"SIB"},{"id":"t2870","date":"2024-12-31","category":"Générales","type":"Dépense","amount":14200,"account":"SIB"},{"id":"t2871","date":"2024-12-31","category":"Générales","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Police"},{"id":"t2872","date":"2024-12-31","category":"Aliments","type":"Dépense","amount":479300,"account":"SIB"},{"id":"t2873","date":"2024-12-31","category":"Vêtements","type":"Dépense","amount":54500,"account":"SIB"},{"id":"t2874","date":"2024-12-31","category":"Santé","type":"Dépense","amount":30000,"account":"SIB"},{"id":"t2811","date":"2025-01-31","category":"GRUNDFOS","type":"Dépense","amount":213000,"account":"SIB","subcategory":"Carburant"},{"id":"t2812","date":"2025-01-31","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2813","date":"2025-01-31","category":"Santé","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Médicaments"},{"id":"t2814","date":"2025-01-31","category":"Générales","type":"Dépense","amount":40400,"account":"SIB"},{"id":"t2815","date":"2025-01-31","category":"Invitation","type":"Dépense","amount":70000,"account":"SIB"},{"id":"t2816","date":"2025-01-31","category":"GRUNDFOS","type":"Dépense","amount":5200,"account":"PETTY CASH","subcategory":"Internet","note":"Internet mobile"},{"id":"t2817","date":"2025-01-31","category":"Voyage","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Un hôtel","note":"Yamoussoukro"},{"id":"t2818","date":"2025-01-31","category":"Épargne","type":"Dépense","amount":310000,"account":"SIB"},{"id":"t2819","date":"2025-01-31","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Cotisations"},{"id":"t2820","date":"2025-01-31","category":"Dette","type":"Dépense","amount":631166,"account":"SIB","payee":"SIB","note":"Remboursement dettes"},{"id":"t2821","date":"2025-01-31","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t2822","date":"2025-01-31","category":"Personnel","type":"Dépense","amount":4500,"account":"SIB","subcategory":"Coiffure"},{"id":"t2823","date":"2025-01-31","category":"Générales","type":"Dépense","amount":1800,"account":"SIB"},{"id":"t2824","date":"2025-01-31","category":"Vente Pompe","type":"Revenu","amount":410000,"account":"SIB"},{"id":"t2825","date":"2025-01-31","category":"Ajustement","type":"Revenu","amount":20100,"account":"SIB"},{"id":"t2826","date":"2025-01-31","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2827","date":"2025-01-31","category":"Enfants & Maman","type":"Dépense","amount":134500,"account":"SIB","subcategory":"Hemra"},{"id":"t2828","date":"2025-01-31","category":"Logement","type":"Dépense","amount":551000,"account":"SIB"},{"id":"t2829","date":"2025-01-31","category":"Aliments","type":"Dépense","amount":365100,"account":"SIB"},{"id":"t2830","date":"2025-01-31","category":"Utilitaires","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Téléphones","note":"Coque"},{"id":"t2831","date":"2025-01-31","category":"Abonnements","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Netflix"},{"id":"t2832","date":"2025-01-31","category":"Voiture","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Lavage"},{"id":"t2833","date":"2025-01-31","category":"Générales","type":"Dépense","amount":150000,"account":"SIB"},{"id":"t2834","date":"2025-01-31","category":"Revenus Location Mazda","type":"Revenu","amount":2238058,"account":"SIB"},{"id":"t2835","date":"2025-01-31","category":"Cadeaux","type":"Dépense","amount":613832,"account":"SIB"},{"id":"t2836","date":"2025-01-31","category":"Divertissement","type":"Dépense","amount":479500,"account":"SIB"},{"id":"t2837","date":"2025-01-31","category":"Ajustement","type":"Dépense","amount":68682,"account":"SIB"},{"id":"t2838","date":"2025-01-31","category":"Abonnements","type":"Dépense","amount":1877,"account":"SIB","note":"Assurance SAF"},{"id":"t2839","date":"2025-01-31","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","subcategory":"Cotisations","note":"Décès Affoue"},{"id":"t2840","date":"2025-01-31","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2841","date":"2025-01-31","category":"Générales","type":"Dépense","amount":5000,"account":"SIB","note":"Certificat de résidence"},{"id":"t2842","date":"2025-01-31","category":"Shopping","type":"Dépense","amount":166200,"account":"SIB"},{"id":"t2843","date":"2025-01-31","category":"Utilitaires","type":"Dépense","amount":20000,"account":"SIB","subcategory":"la télé","note":"Canal"},{"id":"t2844","date":"2025-01-31","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Cotisations"},{"id":"t2845","date":"2025-01-31","category":"Générales","type":"Dépense","amount":4600,"account":"SIB"},{"id":"t2846","date":"2025-01-31","category":"Des sports","type":"Dépense","amount":74000,"account":"SIB","subcategory":"Équipement"},{"id":"t2847","date":"2025-01-31","category":"Vêtements","type":"Dépense","amount":182000,"account":"SIB"},{"id":"t2848","date":"2025-01-31","category":"Personnel","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2777","date":"2025-02-28","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t2778","date":"2025-02-28","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2779","date":"2025-02-28","category":"GRUNDFOS","type":"Dépense","amount":180000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2780","date":"2025-02-28","category":"GRUNDFOS","type":"Dépense","amount":95000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2781","date":"2025-02-28","category":"Vêtements","type":"Dépense","amount":22000,"account":"SIB"},{"id":"t2782","date":"2025-02-28","category":"Générales","type":"Dépense","amount":2020,"account":"SIB","subcategory":"Carte Djamo"},{"id":"t2783","date":"2025-02-28","category":"Utilitaires","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Téléphones","note":"Coque"},{"id":"t2784","date":"2025-02-28","category":"Abonnements","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Spotify"},{"id":"t2785","date":"2025-02-28","category":"Enfants & Maman","type":"Dépense","amount":60600,"account":"SIB","subcategory":"Nesher"},{"id":"t2786","date":"2025-02-28","category":"GRUNDFOS","type":"Dépense","amount":11500,"account":"PETTY CASH","subcategory":"Cachet"},{"id":"t2787","date":"2025-02-28","category":"Dette","type":"Dépense","amount":631166,"account":"SIB"},{"id":"t2788","date":"2025-02-28","category":"Aliments","type":"Dépense","amount":249500,"account":"SIB"},{"id":"t2789","date":"2025-02-28","category":"Shopping","type":"Dépense","amount":65500,"account":"SIB"},{"id":"t2790","date":"2025-02-28","category":"Épargne","type":"Dépense","amount":280000,"account":"SIB"},{"id":"t2791","date":"2025-02-28","category":"Petty Cash","type":"Revenu","amount":2500000,"account":"SIB"},{"id":"t2792","date":"2025-02-28","category":"Enfants & Maman","type":"Dépense","amount":76000,"account":"SIB","subcategory":"Hemra"},{"id":"t2793","date":"2025-02-28","category":"Logement","type":"Dépense","amount":550500,"account":"SIB","subcategory":"Location"},{"id":"t2794","date":"2025-02-28","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2795","date":"2025-02-28","category":"Voiture","type":"Dépense","amount":9000,"account":"SIB","subcategory":"Lavage"},{"id":"t2796","date":"2025-02-28","category":"Personnel","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2797","date":"2025-02-28","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"SIB"},{"id":"t2798","date":"2025-02-28","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2799","date":"2025-02-28","category":"Générales","type":"Dépense","amount":1000,"account":"SIB","note":"Impression"},{"id":"t2800","date":"2025-02-28","category":"Divertissement","type":"Dépense","amount":388000,"account":"SIB"},{"id":"t2801","date":"2025-02-28","category":"Vente Pompe","type":"Revenu","amount":250000,"account":"SIB"},{"id":"t2802","date":"2025-02-28","category":"Voiture","type":"Dépense","amount":45000,"account":"SIB","subcategory":"Entretien"},{"id":"t2803","date":"2025-02-28","category":"Invitation","type":"Dépense","amount":25000,"account":"SIB"},{"id":"t2804","date":"2025-02-28","category":"Cadeaux","type":"Dépense","amount":355350,"account":"SIB"},{"id":"t2805","date":"2025-02-28","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t2806","date":"2025-02-28","category":"Ajustement","type":"Dépense","amount":57258,"account":"SIB"},{"id":"t2807","date":"2025-02-28","category":"Santé","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2808","date":"2025-02-28","category":"Ajustement","type":"Revenu","amount":1242,"account":"SIB"},{"id":"t2809","date":"2025-02-28","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SIB","note":"Michelle"},{"id":"t2810","date":"2025-02-28","category":"Générales","type":"Dépense","amount":8050,"account":"SIB","note":"Transfert bose"},{"id":"t3044","date":"2025-02-28","category":"Générales","type":"Dépense","amount":50500,"account":"SIB","subcategory":"Création Entreprise ECO PUMP","note":"Site internet"},{"id":"t2740","date":"2025-03-31","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2741","date":"2025-03-31","category":"Logement","type":"Dépense","amount":550500,"account":"Dépôt LOYER","subcategory":"Location"},{"id":"t2742","date":"2025-03-31","category":"GRUNDFOS","type":"Dépense","amount":175000,"account":"SIB","subcategory":"Carburant"},{"id":"t2743","date":"2025-03-31","category":"GRUNDFOS","type":"Dépense","amount":2704,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t2744","date":"2025-03-31","category":"GRUNDFOS","type":"Dépense","amount":6500,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t2745","date":"2025-03-31","category":"Dette","type":"Dépense","amount":631166,"account":"SIB"},{"id":"t2746","date":"2025-03-31","category":"Utilitaires","type":"Dépense","amount":35000,"account":"SIB","subcategory":"la télé","note":"IPTV"},{"id":"t2747","date":"2025-03-31","category":"Un salaire","type":"Revenu","amount":1419055,"account":"SIB"},{"id":"t2748","date":"2025-03-31","category":"Divertissement","type":"Dépense","amount":556000,"account":"SIB"},{"id":"t2749","date":"2025-03-31","category":"Abonnements","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Spotify"},{"id":"t2750","date":"2025-03-31","category":"Pack Club","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2751","date":"2025-03-31","category":"Transport","type":"Dépense","amount":4200,"account":"SIB","subcategory":"Taxi"},{"id":"t2752","date":"2025-03-31","category":"Enfants & Maman","type":"Dépense","amount":36000,"account":"SIB","subcategory":"Hemra"},{"id":"t2753","date":"2025-03-31","category":"Voiture","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Lavage"},{"id":"t2754","date":"2025-03-31","category":"Cadeaux","type":"Dépense","amount":294150,"account":"SIB"},{"id":"t2755","date":"2025-03-31","category":"Voiture","type":"Dépense","amount":21000,"account":"SIB","subcategory":"Entretien"},{"id":"t2756","date":"2025-03-31","category":"Utilitaires","type":"Dépense","amount":24000,"account":"SIB","subcategory":"Téléphones","note":"Réparation Iphone"},{"id":"t2757","date":"2025-03-31","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t2758","date":"2025-03-31","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2759","date":"2025-03-31","category":"Vêtements","type":"Dépense","amount":15000,"account":"SIB"},{"id":"t2760","date":"2025-03-31","category":"Vente Pompe","type":"Revenu","amount":1447000,"account":"SIB"},{"id":"t2761","date":"2025-03-31","category":"Aliments","type":"Dépense","amount":224500,"account":"SIB"},{"id":"t2762","date":"2025-03-31","category":"Épargne","type":"Dépense","amount":310000,"account":"SIB"},{"id":"t2763","date":"2025-03-31","category":"Payement Maison Bingerville","type":"Dépense","amount":5000000,"account":"SIB"},{"id":"t2764","date":"2025-03-31","category":"Personnel","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2765","date":"2025-03-31","category":"Générales","type":"Dépense","amount":8030,"account":"SIB","subcategory":"Police"},{"id":"t2766","date":"2025-03-31","category":"Ajustement","type":"Dépense","amount":42699,"account":"SIB"},{"id":"t2767","date":"2025-03-31","category":"Générales","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Livraison Pompe"},{"id":"t2768","date":"2025-03-31","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"SIB"},{"id":"t2769","date":"2025-03-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t2770","date":"2025-03-31","category":"Shopping","type":"Dépense","amount":185000,"account":"SIB"},{"id":"t2771","date":"2025-03-31","category":"Transport","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Péage"},{"id":"t2772","date":"2025-03-31","category":"Santé","type":"Dépense","amount":13000,"account":"SIB"},{"id":"t2773","date":"2025-03-31","category":"Ajustement","type":"Dépense","amount":22000,"account":"SIB"},{"id":"t2774","date":"2025-03-31","category":"Générales","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Réparation Robinet"},{"id":"t2775","date":"2025-03-31","category":"GRUNDFOS","type":"Dépense","amount":51000,"account":"SIB","note":"Visa Pieter"},{"id":"t2776","date":"2025-03-31","category":"Loyer","type":"Revenu","amount":1500000,"account":"SIB"},{"id":"t2737","date":"2025-04-01","category":"Vente Pompe","type":"Revenu","amount":230900,"account":"SIB"},{"id":"t2738","date":"2025-04-01","category":"Épargne","type":"Revenu","amount":900000,"account":"SIB"},{"id":"t2739","date":"2025-04-01","category":"Ajustement","type":"Revenu","amount":1280,"account":"SIB"},{"id":"t2720","date":"2025-04-17","category":"Voiture","type":"Dépense","amount":47510,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Batterie"},{"id":"t2721","date":"2025-04-17","category":"Voiture","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2722","date":"2025-04-17","category":"GRUNDFOS","type":"Dépense","amount":80000,"account":"PETTY CASH","subcategory":"Hotel","note":"Hotel commissioning"},{"id":"t2723","date":"2025-04-17","category":"Générales","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Police"},{"id":"t2724","date":"2025-04-17","category":"Aliments","type":"Dépense","amount":89200,"account":"SIB"},{"id":"t2725","date":"2025-04-17","category":"Divertissement","type":"Dépense","amount":153000,"account":"SIB"},{"id":"t2726","date":"2025-04-17","category":"GRUNDFOS","type":"Dépense","amount":63000,"account":"PETTY CASH","subcategory":"Restaurant","note":"Meal Service Team"},{"id":"t2727","date":"2025-04-17","category":"Shopping","type":"Dépense","amount":8234,"account":"SIB"},{"id":"t2728","date":"2025-04-17","category":"GRUNDFOS","type":"Dépense","amount":4500,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t2729","date":"2025-04-17","category":"Enfants & Maman","type":"Dépense","amount":33000,"account":"SIB","subcategory":"Hemra"},{"id":"t2730","date":"2025-04-17","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Impression"},{"id":"t2731","date":"2025-04-17","category":"Utilitaires","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Téléphones","note":"Cable chargeur"},{"id":"t2732","date":"2025-04-17","category":"Cadeaux","type":"Dépense","amount":215900,"account":"SIB"},{"id":"t2733","date":"2025-04-17","category":"Vêtements","type":"Dépense","amount":60000,"account":"SIB"},{"id":"t2734","date":"2025-04-17","category":"Personnel","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2735","date":"2025-04-17","category":"GRUNDFOS","type":"Dépense","amount":120000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2736","date":"2025-04-17","category":"Générales","type":"Dépense","amount":72000,"account":"SIB","subcategory":"Réparation Robinet"},{"id":"t2717","date":"2025-04-18","category":"Divertissement","type":"Dépense","amount":77000,"account":"SIB"},{"id":"t2718","date":"2025-04-18","category":"Aliments","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Dîner"},{"id":"t2719","date":"2025-04-18","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SIB"},{"id":"t2708","date":"2025-04-19","category":"Revenu général","type":"Revenu","amount":0,"account":"Dépôt LOYER","note":"Montant initial"},{"id":"t2709","date":"2025-04-19","category":"Revenu général","type":"Revenu","amount":0,"account":"PETTY CASH","note":"Montant initial"},{"id":"t2710","date":"2025-04-19","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2711","date":"2025-04-19","category":"Abonnements","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Money Coach"},{"id":"t2712","date":"2025-04-19","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Augustin"},{"id":"t2713","date":"2025-04-19","category":"Générales","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Vol Djamo"},{"id":"t2714","date":"2025-04-19","category":"Ajustement","type":"Dépense","amount":17484,"account":"SIB"},{"id":"t2715","date":"2025-04-19","category":"Aliments","type":"Dépense","amount":17000,"account":"SIB"},{"id":"t2716","date":"2025-04-19","category":"Utilitaires","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Téléphones","note":"Accessoires"},{"id":"t2693","date":"2025-04-20","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Elvira"},{"id":"t2694","date":"2025-04-20","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Carburant","note":"Sans reçu"},{"id":"t2695","date":"2025-04-20","category":"Shopping","type":"Dépense","amount":33000,"account":"SIB"},{"id":"t2696","date":"2025-04-20","category":"Divertissement","type":"Dépense","amount":20000,"account":"SIB"},{"id":"t2697","date":"2025-04-20","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2698","date":"2025-04-20","category":"Générales","type":"Dépense","amount":6095,"account":"SIB"},{"id":"t2699","date":"2025-04-20","category":"General","type":"Dépense","amount":550500,"account":"SIB","payee":"comptes","note":"Rapprochement des"},{"id":"t2700","date":"2025-04-20","category":"General","type":"Revenu","amount":550500,"account":"Dépôt LOYER","payee":"comptes","note":"Rapprochement des"},{"id":"t2701","date":"2025-04-20","category":"General","type":"Dépense","amount":1000000,"account":"SIB","payee":"comptes","note":"Rapprochement des"},{"id":"t2702","date":"2025-04-20","category":"General","type":"Revenu","amount":1000000,"account":"Dépôt LOYER","payee":"comptes","note":"Rapprochement des"},{"id":"t2703","date":"2025-04-20","category":"General","type":"Dépense","amount":1132180,"account":"SIB","payee":"comptes","note":"Rapprochement des"},{"id":"t2704","date":"2025-04-20","category":"General","type":"Revenu","amount":583785,"account":"SIB","payee":"comptes","note":"Rapprochement des"},{"id":"t2705","date":"2025-04-20","category":"General","type":"Revenu","amount":548395,"account":"PETTY CASH","payee":"comptes","note":"Rapprochement des"},{"id":"t2706","date":"2025-04-20","category":"Invitation","type":"Dépense","amount":50000,"account":"SIB","note":"Andrea"},{"id":"t2707","date":"2025-04-20","category":"Divertissement","type":"Dépense","amount":45000,"account":"SIB"},{"id":"t2688","date":"2025-04-21","category":"Transport","type":"Dépense","amount":3200,"account":"SIB","subcategory":"Taxi","note":"Livraison Andrea"},{"id":"t2689","date":"2025-04-21","category":"Aliments","type":"Dépense","amount":5200,"account":"SIB"},{"id":"t2690","date":"2025-04-21","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB"},{"id":"t2691","date":"2025-04-21","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB"},{"id":"t2692","date":"2025-04-21","category":"Divertissement","type":"Dépense","amount":17880,"account":"SIB","subcategory":"BAP"},{"id":"t2685","date":"2025-04-22","category":"Enfants & Maman","type":"Dépense","amount":40000,"account":"SIB","subcategory":"Hemra"},{"id":"t2686","date":"2025-04-22","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Dîner"},{"id":"t2687","date":"2025-04-22","category":"Shopping","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2681","date":"2025-04-23","category":"Aliments","type":"Dépense","amount":20000,"account":"SIB"},{"id":"t2682","date":"2025-04-23","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2683","date":"2025-04-23","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Móbile"},{"id":"t2684","date":"2025-04-23","category":"Aliments","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Dîner"},{"id":"t2673","date":"2025-04-24","category":"Aliments","type":"Dépense","amount":9000,"account":"SIB","subcategory":"Le déjeuner"},{"id":"t2674","date":"2025-04-24","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SIB","note":"Mariage Maurice"},{"id":"t2675","date":"2025-04-24","category":"Cadeaux","type":"Dépense","amount":85000,"account":"SIB","note":"Andrea Four"},{"id":"t2676","date":"2025-04-24","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2677","date":"2025-04-24","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t2678","date":"2025-04-24","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2679","date":"2025-04-24","category":"Aliments","type":"Dépense","amount":17000,"account":"SIB","subcategory":"Dîner"},{"id":"t2680","date":"2025-04-24","category":"Dette","type":"Dépense","amount":511166,"account":"SIB"},{"id":"t2665","date":"2025-04-25","category":"Cadeaux","type":"Dépense","amount":900,"account":"SIB"},{"id":"t2666","date":"2025-04-25","category":"Logement","type":"Dépense","amount":450550,"account":"Dépôt LOYER","subcategory":"Location","note":"Loyer Mai"},{"id":"t2667","date":"2025-04-25","category":"Générales","type":"Dépense","amount":2020,"account":"SIB","subcategory":"Police"},{"id":"t2668","date":"2025-04-25","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB"},{"id":"t2669","date":"2025-04-25","category":"GRUNDFOS","type":"Dépense","amount":9000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Impression - IScanner"},{"id":"t2670","date":"2025-04-25","category":"Abonnements","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Spotify"},{"id":"t2671","date":"2025-04-25","category":"Aliments","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Dîner"},{"id":"t2672","date":"2025-04-25","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Mahora"},{"id":"t2660","date":"2025-04-26","category":"Dette","type":"Dépense","amount":120000,"account":"SIB","subcategory":"PEL"},{"id":"t2661","date":"2025-04-26","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2662","date":"2025-04-26","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Dîner"},{"id":"t2663","date":"2025-04-26","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2664","date":"2025-04-26","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Le déjeuner"},{"id":"t2654","date":"2025-04-27","category":"Formation","type":"Dépense","amount":5171,"account":"SIB","subcategory":"Piano & Guitare"},{"id":"t2655","date":"2025-04-27","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Devy"},{"id":"t2656","date":"2025-04-27","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Raquine"},{"id":"t2657","date":"2025-04-27","category":"Divertissement","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Fête","note":"Alfred"},{"id":"t2658","date":"2025-04-27","category":"Aliments","type":"Dépense","amount":11000,"account":"SIB","subcategory":"Dîner"},{"id":"t2659","date":"2025-04-27","category":"Voiture","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Crevaison"},{"id":"t2645","date":"2025-04-28","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Alcool","note":"Olo"},{"id":"t2646","date":"2025-04-28","category":"Invitation","type":"Dépense","amount":27000,"account":"SIB","note":"Maï"},{"id":"t2647","date":"2025-04-28","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t2648","date":"2025-04-28","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2649","date":"2025-04-28","category":"GRUNDFOS","type":"Dépense","amount":91920,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2650","date":"2025-04-28","category":"Voiture","type":"Dépense","amount":9000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Fixation"},{"id":"t2651","date":"2025-04-28","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","note":"Mardochee"},{"id":"t2652","date":"2025-04-28","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2653","date":"2025-04-28","category":"Voiture","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Pneu"},{"id":"t2635","date":"2025-04-29","category":"Divertissement","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Alcool"},{"id":"t2636","date":"2025-04-29","category":"Enfants & Maman","type":"Dépense","amount":8500,"account":"SIB","subcategory":"Hemra","note":"Santé"},{"id":"t2637","date":"2025-04-29","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SIB","subcategory":"Maman"},{"id":"t2638","date":"2025-04-29","category":"Enfants & Maman","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Hemra","note":"Carte d'assurance"},{"id":"t2639","date":"2025-04-29","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t2640","date":"2025-04-29","category":"Shopping","type":"Dépense","amount":15000,"account":"SIB"},{"id":"t2641","date":"2025-04-29","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2642","date":"2025-04-29","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t2643","date":"2025-04-29","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Julie"},{"id":"t2644","date":"2025-04-29","category":"Un salaire","type":"Revenu","amount":2804390,"account":"SIB","note":"Salaire + Bonus + Augmentation"},{"id":"t2632","date":"2025-04-30","category":"Shopping","type":"Dépense","amount":38000,"account":"SIB"},{"id":"t2633","date":"2025-04-30","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Dîner"},{"id":"t2634","date":"2025-04-30","category":"Shopping","type":"Dépense","amount":6600,"account":"SIB"},{"id":"t3042","date":"2025-04-30","category":"Petty Cash","type":"Revenu","amount":434075,"account":"SIB","subcategory":"Ajustement Petty Cash"},{"id":"t3043","date":"2025-04-30","category":"GRUNDFOS","type":"Dépense","amount":434075,"account":"PETTY CASH","subcategory":"Ajustement Petty Cash"},{"id":"t2631","date":"2025-05-01","category":"Divertissement","type":"Dépense","amount":9000,"account":"SIB","subcategory":"BAP"},{"id":"t2625","date":"2025-05-02","category":"Divertissement","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Alcool"},{"id":"t2626","date":"2025-05-02","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2627","date":"2025-05-02","category":"Vente Pompe","type":"Revenu","amount":49000,"account":"SIB","note":"Moteur 2.2KW"},{"id":"t2628","date":"2025-05-02","category":"Cadeaux","type":"Dépense","amount":10310,"account":"SIB","note":"Grâce"},{"id":"t2629","date":"2025-05-02","category":"Divertissement","type":"Dépense","amount":13000,"account":"SIB","subcategory":"BAP"},{"id":"t2630","date":"2025-05-02","category":"Divertissement","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Alcool"},{"id":"t2622","date":"2025-05-03","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2623","date":"2025-05-03","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Cotisations","note":"Bb Yves kouassi"},{"id":"t2624","date":"2025-05-03","category":"Transport","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Taxi","note":"Yango Schela"},{"id":"t2618","date":"2025-05-04","category":"Aliments","type":"Dépense","amount":3500,"account":"SIB"},{"id":"t2619","date":"2025-05-04","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alcool","note":"Schela"},{"id":"t2620","date":"2025-05-04","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Dîner","note":"Etran"},{"id":"t2621","date":"2025-05-04","category":"Divertissement","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Alcool","note":"Etran"},{"id":"t2611","date":"2025-05-05","category":"Aliments","type":"Dépense","amount":600,"account":"SIB"},{"id":"t2612","date":"2025-05-05","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t2613","date":"2025-05-05","category":"Utilitaires","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Téléphones","note":"Antichoc"},{"id":"t2614","date":"2025-05-05","category":"Aliments","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2615","date":"2025-05-05","category":"Ajustement","type":"Revenu","amount":20000,"account":"SIB","note":"Assa"},{"id":"t2616","date":"2025-05-05","category":"Divertissement","type":"Dépense","amount":21500,"account":"SIB","subcategory":"Fête","note":"Anniversaire OLO"},{"id":"t2617","date":"2025-05-05","category":"Générales","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Souris Sans Fil"},{"id":"t2607","date":"2025-05-06","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alcool"},{"id":"t2608","date":"2025-05-06","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2609","date":"2025-05-06","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Dîner"},{"id":"t2610","date":"2025-05-06","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t2602","date":"2025-05-07","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2603","date":"2025-05-07","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2604","date":"2025-05-07","category":"Invitation","type":"Dépense","amount":35000,"account":"SIB","subcategory":"Femmes","note":"Marie AKA"},{"id":"t2605","date":"2025-05-07","category":"Divertissement","type":"Dépense","amount":16000,"account":"SIB","subcategory":"Alcool"},{"id":"t2606","date":"2025-05-07","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Internet","note":"Móbile"},{"id":"t2597","date":"2025-05-08","category":"Aliments","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Dîner"},{"id":"t2598","date":"2025-05-08","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Ruth"},{"id":"t2599","date":"2025-05-08","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2600","date":"2025-05-08","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Alcool","note":"Rumba"},{"id":"t2601","date":"2025-05-08","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2593","date":"2025-05-09","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Alcool"},{"id":"t2594","date":"2025-05-09","category":"Aliments","type":"Dépense","amount":5500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2595","date":"2025-05-09","category":"Cadeaux","type":"Dépense","amount":120000,"account":"SIB","note":"Hammal"},{"id":"t2596","date":"2025-05-09","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Alcool"},{"id":"t2586","date":"2025-05-10","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Móbile"},{"id":"t2587","date":"2025-05-10","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"BAP"},{"id":"t2588","date":"2025-05-10","category":"Aliments","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Dîner"},{"id":"t2589","date":"2025-05-10","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Femmes","note":"Fatim"},{"id":"t2590","date":"2025-05-10","category":"Vente Pompe","type":"Revenu","amount":41090,"account":"SIB","note":"MS4000 4KW"},{"id":"t2591","date":"2025-05-10","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2592","date":"2025-05-10","category":"Enfants & Maman","type":"Dépense","amount":20200,"account":"SIB","subcategory":"Nesher"},{"id":"t2578","date":"2025-05-11","category":"Abonnements","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Onfray"},{"id":"t2579","date":"2025-05-11","category":"Shopping","type":"Dépense","amount":6000,"account":"SIB"},{"id":"t2580","date":"2025-05-11","category":"Aliments","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2581","date":"2025-05-11","category":"Shopping","type":"Dépense","amount":26550,"account":"SIB"},{"id":"t2582","date":"2025-05-11","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2583","date":"2025-05-11","category":"Santé","type":"Dépense","amount":11210,"account":"SIB","subcategory":"VG"},{"id":"t2584","date":"2025-05-11","category":"Ajustement","type":"Dépense","amount":19188,"account":"SIB"},{"id":"t2585","date":"2025-05-11","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t2577","date":"2025-05-12","category":"Invitation","type":"Dépense","amount":55000,"account":"SIB","subcategory":"Femmes","note":"Elvira"},{"id":"t2574","date":"2025-05-13","category":"Cadeaux","type":"Dépense","amount":17000,"account":"SIB","note":"Elvira"},{"id":"t2575","date":"2025-05-13","category":"Shopping","type":"Dépense","amount":30000,"account":"SIB"},{"id":"t2576","date":"2025-05-13","category":"Cadeaux","type":"Dépense","amount":6600,"account":"SIB","note":"Aliments Elvira"},{"id":"t2570","date":"2025-05-14","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Elvira"},{"id":"t2571","date":"2025-05-14","category":"Divertissement","type":"Dépense","amount":1800,"account":"SIB","subcategory":"Alcool"},{"id":"t2572","date":"2025-05-14","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2573","date":"2025-05-14","category":"Ajustement","type":"Dépense","amount":4372,"account":"SIB"},{"id":"t2568","date":"2025-05-15","category":"Santé","type":"Dépense","amount":13300,"account":"SIB","subcategory":"Médicaments"},{"id":"t2569","date":"2025-05-15","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t2562","date":"2025-05-16","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2563","date":"2025-05-16","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2564","date":"2025-05-16","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB"},{"id":"t2565","date":"2025-05-16","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Mardochee"},{"id":"t2566","date":"2025-05-16","category":"Invitation","type":"Dépense","amount":15000,"account":"SIB","note":"Etran"},{"id":"t2567","date":"2025-05-16","category":"Cadeaux","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Pourboire"},{"id":"t2558","date":"2025-05-17","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2559","date":"2025-05-17","category":"Aliments","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Le déjeuner"},{"id":"t2560","date":"2025-05-17","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t2561","date":"2025-05-17","category":"Invitation","type":"Dépense","amount":35000,"account":"SIB","subcategory":"Femmes","note":"Mai Konate"},{"id":"t2551","date":"2025-05-18","category":"Divertissement","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Alcool"},{"id":"t2552","date":"2025-05-18","category":"Shopping","type":"Dépense","amount":13000,"account":"SIB","note":"Andrea"},{"id":"t2553","date":"2025-05-18","category":"Invitation","type":"Dépense","amount":18000,"account":"SIB","note":"Risenata"},{"id":"t2554","date":"2025-05-18","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2555","date":"2025-05-18","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Anniversaire","note":"Risenata"},{"id":"t2556","date":"2025-05-18","category":"Cadeaux","type":"Dépense","amount":7000,"account":"SIB","note":"Andrea"},{"id":"t2557","date":"2025-05-18","category":"Personnel","type":"Dépense","amount":500,"account":"SIB","subcategory":"Coiffure"},{"id":"t2547","date":"2025-05-19","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SIB","note":"Sylvia"},{"id":"t2548","date":"2025-05-19","category":"Aliments","type":"Dépense","amount":800,"account":"SIB"},{"id":"t2549","date":"2025-05-19","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Ange ora"},{"id":"t2550","date":"2025-05-19","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Obed"},{"id":"t2545","date":"2025-05-20","category":"Ajustement","type":"Revenu","amount":538,"account":"SIB"},{"id":"t2546","date":"2025-05-20","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2539","date":"2025-05-21","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":600000,"account":"SIB"},{"id":"t2540","date":"2025-05-21","category":"Loyer","type":"Revenu","amount":1500000,"account":"Dépôt LOYER"},{"id":"t2541","date":"2025-05-21","category":"Aliments","type":"Dépense","amount":3900,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2542","date":"2025-05-21","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2543","date":"2025-05-21","category":"Aliments","type":"Dépense","amount":24000,"account":"SIB","subcategory":"Invitation","note":"Ouanlo"},{"id":"t2544","date":"2025-05-21","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"BAP"},{"id":"t2532","date":"2025-05-22","category":"Aliments","type":"Dépense","amount":7070,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2533","date":"2025-05-22","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Amira"},{"id":"t2534","date":"2025-05-22","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2535","date":"2025-05-22","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Impression"},{"id":"t2536","date":"2025-05-22","category":"Aliments","type":"Dépense","amount":599,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2537","date":"2025-05-22","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t2538","date":"2025-05-22","category":"Divertissement","type":"Dépense","amount":22000,"account":"SIB","subcategory":"Fête"},{"id":"t2524","date":"2025-05-23","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Péage"},{"id":"t2525","date":"2025-05-23","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t2526","date":"2025-05-23","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2527","date":"2025-05-23","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2528","date":"2025-05-23","category":"Vente Pompe","type":"Revenu","amount":72200,"account":"SIB","note":"Moteur 2.2 Tinin 375000"},{"id":"t2529","date":"2025-05-23","category":"Dette","type":"Dépense","amount":511166,"account":"SIB"},{"id":"t2530","date":"2025-05-23","category":"Aliments","type":"Dépense","amount":8500,"account":"SIB","subcategory":"Dîner"},{"id":"t2531","date":"2025-05-23","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Pourboire","note":"Dydime"},{"id":"t2522","date":"2025-05-24","category":"Invitation","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Femmes","note":"Mahora"},{"id":"t2523","date":"2025-05-24","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2516","date":"2025-05-25","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Mahora"},{"id":"t2517","date":"2025-05-25","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2518","date":"2025-05-25","category":"Invitation","type":"Dépense","amount":42000,"account":"SIB","subcategory":"Femmes","note":"Mahora"},{"id":"t2519","date":"2025-05-25","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t2520","date":"2025-05-25","category":"Aliments","type":"Dépense","amount":700,"account":"SIB","subcategory":"Dîner"},{"id":"t2521","date":"2025-05-25","category":"Revenu général","type":"Revenu","amount":1,"account":"SALAIRE","note":"Montant initial"},{"id":"t2510","date":"2025-05-26","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2511","date":"2025-05-26","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","note":"Tantie Esther"},{"id":"t2512","date":"2025-05-26","category":"Cadeaux","type":"Dépense","amount":5100,"account":"SIB","note":"MB"},{"id":"t2513","date":"2025-05-26","category":"Invitation","type":"Dépense","amount":17000,"account":"SIB","subcategory":"Femmes","note":"Ruth"},{"id":"t2514","date":"2025-05-26","category":"Aliments","type":"Dépense","amount":4500,"account":"SIB","subcategory":"Dîner"},{"id":"t2515","date":"2025-05-26","category":"Dette","type":"Dépense","amount":120000,"account":"SIB","subcategory":"PEL"},{"id":"t2508","date":"2025-05-27","category":"Logement","type":"Dépense","amount":450550,"account":"Dépôt LOYER","subcategory":"Location","note":"Juin 2025"},{"id":"t2509","date":"2025-05-27","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t2502","date":"2025-05-28","category":"Enfants & Maman","type":"Dépense","amount":7110,"account":"SIB","subcategory":"Hemra"},{"id":"t2503","date":"2025-05-28","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2504","date":"2025-05-28","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2505","date":"2025-05-28","category":"Shopping","type":"Dépense","amount":11000,"account":"SIB"},{"id":"t2506","date":"2025-05-28","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB"},{"id":"t2507","date":"2025-05-28","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Dîner"},{"id":"t2495","date":"2025-05-29","category":"Aliments","type":"Dépense","amount":25000,"account":"SIB","note":"Elvis YAO"},{"id":"t2496","date":"2025-05-29","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre"},{"id":"t2497","date":"2025-05-29","category":"Cadeaux","type":"Dépense","amount":5205,"account":"SIB","note":"Augustin"},{"id":"t2498","date":"2025-05-29","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t2499","date":"2025-05-29","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SIB","subcategory":"Maman"},{"id":"t2500","date":"2025-05-29","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"BAP","note":"Vera"},{"id":"t2501","date":"2025-05-29","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Alcool"},{"id":"t2480","date":"2025-05-30","category":"Invitation","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Femmes","note":"Lynda"},{"id":"t2481","date":"2025-05-30","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2482","date":"2025-05-30","category":"Générales","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Vol Djamo"},{"id":"t2483","date":"2025-05-30","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SIB","note":"Lynda"},{"id":"t2484","date":"2025-05-30","category":"Divertissement","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Alcool"},{"id":"t2485","date":"2025-05-30","category":"Shopping","type":"Dépense","amount":18000,"account":"SIB"},{"id":"t2486","date":"2025-05-30","category":"Enfants & Maman","type":"Dépense","amount":16000,"account":"SIB","subcategory":"Hemra"},{"id":"t2487","date":"2025-05-30","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Mardochee"},{"id":"t2488","date":"2025-05-30","category":"Invitation","type":"Dépense","amount":26000,"account":"SIB","subcategory":"Femmes","note":"Nancy"},{"id":"t2489","date":"2025-05-30","category":"Invitation","type":"Dépense","amount":80000,"account":"SIB","subcategory":"Femmes","note":"Nancy"},{"id":"t2490","date":"2025-05-30","category":"Personnel","type":"Dépense","amount":500,"account":"SIB","subcategory":"Produits de beauté","note":"Peigne"},{"id":"t2491","date":"2025-05-30","category":"Utilitaires","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Téléphones","note":"Antichoc"},{"id":"t2492","date":"2025-05-30","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t2493","date":"2025-05-30","category":"Un salaire","type":"Revenu","amount":1553752,"account":"SALAIRE"},{"id":"t2494","date":"2025-05-30","category":"Personnel","type":"Dépense","amount":345000,"account":"SIB","subcategory":"Produits de beauté","note":"Richkoff"},{"id":"t2478","date":"2025-05-31","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB"},{"id":"t2479","date":"2025-05-31","category":"Ajustement","type":"Dépense","amount":18070,"account":"SIB"},{"id":"t2476","date":"2025-06-01","category":"Divertissement","type":"Dépense","amount":19000,"account":"SIB","subcategory":"BAP"},{"id":"t2477","date":"2025-06-01","category":"Aliments","type":"Dépense","amount":9000,"account":"SIB"},{"id":"t2467","date":"2025-06-02","category":"Générales","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Carte Djamo"},{"id":"t2468","date":"2025-06-02","category":"Cadeaux","type":"Dépense","amount":102000,"account":"SIB","subcategory":"MJO"},{"id":"t2469","date":"2025-06-02","category":"Ajustement","type":"Revenu","amount":880,"account":"SIB"},{"id":"t2470","date":"2025-06-02","category":"Shopping","type":"Dépense","amount":1300,"account":"SIB"},{"id":"t2471","date":"2025-06-02","category":"Abonnements","type":"Dépense","amount":6066,"account":"SIB","subcategory":"Money Coach"},{"id":"t2472","date":"2025-06-02","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2473","date":"2025-06-02","category":"Shopping","type":"Dépense","amount":26000,"account":"SIB"},{"id":"t2474","date":"2025-06-02","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2475","date":"2025-06-02","category":"Cadeaux","type":"Dépense","amount":26500,"account":"SIB","note":"Mahora"},{"id":"t2461","date":"2025-06-03","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Cotisations","note":"Blanche"},{"id":"t2462","date":"2025-06-03","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Thomas"},{"id":"t2463","date":"2025-06-03","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2464","date":"2025-06-03","category":"Divertissement","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Alcool"},{"id":"t2465","date":"2025-06-03","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2466","date":"2025-06-03","category":"Utilitaires","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Téléphones"},{"id":"t2455","date":"2025-06-04","category":"Revenu général","type":"Revenu","amount":0,"account":"SGO","note":"Montant initial"},{"id":"t2456","date":"2025-06-04","category":"INVEST SGO","type":"Dépense","amount":5050,"account":"SGO","subcategory":"DJAMO"},{"id":"t2457","date":"2025-06-04","category":"Ajustement","type":"Revenu","amount":1600,"account":"SIB"},{"id":"t2458","date":"2025-06-04","category":"Vente Pompe","type":"Revenu","amount":161660,"account":"SIB"},{"id":"t2459","date":"2025-06-04","category":"GRUNDFOS","type":"Dépense","amount":4000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t2460","date":"2025-06-04","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Lynda"},{"id":"t2451","date":"2025-06-05","category":"Vente Pompe","type":"Revenu","amount":47000,"account":"SIB"},{"id":"t2452","date":"2025-06-05","category":"Divertissement","type":"Dépense","amount":32500,"account":"SIB","subcategory":"Alcool"},{"id":"t2453","date":"2025-06-05","category":"INVEST SGO","type":"Dépense","amount":10000,"account":"SGO","subcategory":"Daba Finance"},{"id":"t2454","date":"2025-06-05","category":"Abonnements","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Daba Finance"},{"id":"t2447","date":"2025-06-06","category":"Aliments","type":"Dépense","amount":12000,"account":"SIB"},{"id":"t2448","date":"2025-06-06","category":"Cadeaux","type":"Dépense","amount":17000,"account":"SIB","subcategory":"Anniversaire","note":"Kessy"},{"id":"t2449","date":"2025-06-06","category":"Cadeaux","type":"Dépense","amount":64000,"account":"SIB","note":"Kessy"},{"id":"t2450","date":"2025-06-06","category":"Cadeaux","type":"Dépense","amount":320000,"account":"SIB","subcategory":"MJO"},{"id":"t2443","date":"2025-06-07","category":"INVEST SGO","type":"Dépense","amount":4858,"account":"SGO","subcategory":"NSIA"},{"id":"t2444","date":"2025-06-07","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2445","date":"2025-06-07","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Triade"},{"id":"t2446","date":"2025-06-07","category":"Divertissement","type":"Dépense","amount":5500,"account":"SIB","subcategory":"Alcool"},{"id":"t2442","date":"2025-06-08","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2437","date":"2025-06-09","category":"INVEST SGO","type":"Dépense","amount":32000,"account":"SGO","subcategory":"NSIA"},{"id":"t2438","date":"2025-06-09","category":"Abonnements","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Financial AFRIK","note":"3 mois"},{"id":"t2439","date":"2025-06-09","category":"Shopping","type":"Dépense","amount":2500,"account":"SIB"},{"id":"t2440","date":"2025-06-09","category":"Shopping","type":"Dépense","amount":13500,"account":"SIB"},{"id":"t2441","date":"2025-06-09","category":"Shopping","type":"Dépense","amount":5000,"account":"SIB"},{"id":"t2429","date":"2025-06-10","category":"Divertissement","type":"Dépense","amount":1200,"account":"SIB","subcategory":"Alcool"},{"id":"t2430","date":"2025-06-10","category":"Vente Pompe","type":"Revenu","amount":44000,"account":"SIB"},{"id":"t2431","date":"2025-06-10","category":"Abonnements","type":"Dépense","amount":2300,"account":"SIB","subcategory":"Spotify"},{"id":"t2432","date":"2025-06-10","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Marie belle"},{"id":"t2433","date":"2025-06-10","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2434","date":"2025-06-10","category":"GRUNDFOS","type":"Dépense","amount":18000,"account":"PETTY CASH","note":"Extrait de compte Visa"},{"id":"t2435","date":"2025-06-10","category":"Abonnements","type":"Dépense","amount":23500,"account":"SGO","subcategory":"Richbourse"},{"id":"t2436","date":"2025-06-10","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2422","date":"2025-06-11","category":"Invitation","type":"Dépense","amount":10500,"account":"SIB","subcategory":"Femmes","note":"Marie belle"},{"id":"t2423","date":"2025-06-11","category":"GRUNDFOS","type":"Dépense","amount":2453,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t2424","date":"2025-06-11","category":"Vente Pompe","type":"Revenu","amount":46500,"account":"SIB"},{"id":"t2425","date":"2025-06-11","category":"INVEST SGO","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Daba Finance","note":"Filtisac"},{"id":"t2426","date":"2025-06-11","category":"Aliments","type":"Dépense","amount":10500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2427","date":"2025-06-11","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2428","date":"2025-06-11","category":"Divertissement","type":"Dépense","amount":25000,"account":"SIB","subcategory":"BAP"},{"id":"t2419","date":"2025-06-12","category":"Santé","type":"Dépense","amount":8500,"account":"SIB","subcategory":"Médicaments"},{"id":"t2420","date":"2025-06-12","category":"Divertissement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Alcool"},{"id":"t2421","date":"2025-06-12","category":"Invitation","type":"Dépense","amount":10500,"account":"SIB","subcategory":"Femmes"},{"id":"t2417","date":"2025-06-13","category":"Invitation","type":"Dépense","amount":11000,"account":"SIB","note":"Achi"},{"id":"t2418","date":"2025-06-13","category":"Ajustement","type":"Dépense","amount":7809,"account":"SIB"},{"id":"t3041","date":"2025-06-13","category":"Petty Cash","type":"Revenu","amount":2500000,"account":"PETTY CASH","subcategory":"Ajustement Petty Cash"},{"id":"t2413","date":"2025-06-14","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Raquine"},{"id":"t2414","date":"2025-06-14","category":"Formation","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Finelo Invest"},{"id":"t2415","date":"2025-06-14","category":"Invitation","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Femmes","note":"Maï Konate"},{"id":"t2416","date":"2025-06-14","category":"Vente Pompe","type":"Revenu","amount":36500,"account":"SIB"},{"id":"t2412","date":"2025-06-15","category":"Invitation","type":"Dépense","amount":22000,"account":"SIB","subcategory":"Femmes","note":"Andrea"},{"id":"t2408","date":"2025-06-16","category":"Shopping","type":"Dépense","amount":4000,"account":"SIB"},{"id":"t2409","date":"2025-06-16","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t2410","date":"2025-06-16","category":"Aliments","type":"Dépense","amount":9500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2411","date":"2025-06-16","category":"GRUNDFOS","type":"Dépense","amount":3500,"account":"PETTY CASH","payee":"MSC","note":"Document/Comissioning"},{"id":"t2406","date":"2025-06-17","category":"Aliments","type":"Dépense","amount":9000,"account":"SIB","subcategory":"Dîner"},{"id":"t2407","date":"2025-06-17","category":"Vente Pompe","type":"Revenu","amount":185000,"account":"SIB"},{"id":"t2403","date":"2025-06-18","category":"Aliments","type":"Dépense","amount":4500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2404","date":"2025-06-18","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2405","date":"2025-06-18","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Dîner"},{"id":"t2395","date":"2025-06-19","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2396","date":"2025-06-19","category":"Ajustement","type":"Revenu","amount":10000,"account":"SIB"},{"id":"t2397","date":"2025-06-19","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Entretien"},{"id":"t2398","date":"2025-06-19","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2399","date":"2025-06-19","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2400","date":"2025-06-19","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","note":"Obed"},{"id":"t2401","date":"2025-06-19","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t2402","date":"2025-06-19","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2389","date":"2025-06-20","category":"Enfants & Maman","type":"Dépense","amount":45000,"account":"SIB","subcategory":"Hemra"},{"id":"t2390","date":"2025-06-20","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Pourboire"},{"id":"t2391","date":"2025-06-20","category":"Aliments","type":"Dépense","amount":7000,"account":"SIB"},{"id":"t2392","date":"2025-06-20","category":"Divertissement","type":"Dépense","amount":1700,"account":"SIB","subcategory":"Alcool"},{"id":"t2393","date":"2025-06-20","category":"Divertissement","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Alcool"},{"id":"t2394","date":"2025-06-20","category":"Invitation","type":"Dépense","amount":36000,"account":"SIB","subcategory":"Femmes"},{"id":"t2382","date":"2025-06-21","category":"Voiture","type":"Dépense","amount":84000,"account":"Revenus MAZDA","subcategory":"Entretien"},{"id":"t2383","date":"2025-06-21","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Parker"},{"id":"t2384","date":"2025-06-21","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Entretien"},{"id":"t2385","date":"2025-06-21","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t2386","date":"2025-06-21","category":"Vente Pompe","type":"Revenu","amount":350000,"account":"SIB","note":"DWK"},{"id":"t2387","date":"2025-06-21","category":"Divertissement","type":"Dépense","amount":12100,"account":"SIB","subcategory":"BAP"},{"id":"t2388","date":"2025-06-21","category":"Invitation","type":"Dépense","amount":21000,"account":"SIB","subcategory":"Femmes"},{"id":"t2381","date":"2025-06-22","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t2373","date":"2025-06-23","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2374","date":"2025-06-23","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Petty cash"},{"id":"t2375","date":"2025-06-23","category":"Logement","type":"Dépense","amount":450550,"account":"Dépôt LOYER","subcategory":"Location"},{"id":"t2376","date":"2025-06-23","category":"Divertissement","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Alcool"},{"id":"t2377","date":"2025-06-23","category":"Shopping","type":"Dépense","amount":3100,"account":"SIB"},{"id":"t2378","date":"2025-06-23","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2379","date":"2025-06-23","category":"Voiture","type":"Dépense","amount":500,"account":"Revenus MAZDA","subcategory":"Assurance"},{"id":"t2380","date":"2025-06-23","category":"Voiture","type":"Dépense","amount":19000,"account":"Revenus MAZDA","subcategory":"Assurance"},{"id":"t2369","date":"2025-06-24","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Emmanuella"},{"id":"t2370","date":"2025-06-24","category":"GRUNDFOS","type":"Dépense","amount":36500,"account":"PETTY CASH","subcategory":"Restaurant","note":"Clifford/Rene"},{"id":"t2371","date":"2025-06-24","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Franck"},{"id":"t2372","date":"2025-06-24","category":"Dette","type":"Dépense","amount":511166,"account":"SIB"},{"id":"t2364","date":"2025-06-25","category":"Divertissement","type":"Dépense","amount":28000,"account":"SIB","subcategory":"Alcool"},{"id":"t2365","date":"2025-06-25","category":"Un salaire","type":"Revenu","amount":1553752,"account":"SALAIRE"},{"id":"t2366","date":"2025-06-25","category":"Vente Pompe","type":"Revenu","amount":43000,"account":"SIB"},{"id":"t2367","date":"2025-06-25","category":"GRUNDFOS","type":"Dépense","amount":33000,"account":"PETTY CASH","subcategory":"Restaurant","note":"Clifford"},{"id":"t2368","date":"2025-06-25","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Police"},{"id":"t2361","date":"2025-06-26","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t2362","date":"2025-06-26","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","note":"Ismo Etran"},{"id":"t2363","date":"2025-06-26","category":"Utilitaires","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Téléphones","note":"Cable"},{"id":"t2350","date":"2025-06-27","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Allarienne"},{"id":"t2351","date":"2025-06-27","category":"Cadeaux","type":"Dépense","amount":6060,"account":"SIB","subcategory":"Femme","note":"Allarienne"},{"id":"t2352","date":"2025-06-27","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alcool"},{"id":"t2353","date":"2025-06-27","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Dîner"},{"id":"t2354","date":"2025-06-27","category":"Aliments","type":"Dépense","amount":2900,"account":"SIB"},{"id":"t2355","date":"2025-06-27","category":"Divertissement","type":"Dépense","amount":16000,"account":"SIB","subcategory":"Alcool"},{"id":"t2356","date":"2025-06-27","category":"Divertissement","type":"Dépense","amount":3700,"account":"SIB","subcategory":"Alcool"},{"id":"t2357","date":"2025-06-27","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2358","date":"2025-06-27","category":"Ajustement","type":"Dépense","amount":36080,"account":"SIB"},{"id":"t2359","date":"2025-06-27","category":"Dette","type":"Dépense","amount":120000,"account":"SIB","subcategory":"PEL"},{"id":"t2360","date":"2025-06-27","category":"Bourse","type":"Revenu","amount":37300,"account":"SGO"},{"id":"t2346","date":"2025-06-28","category":"Aliments","type":"Dépense","amount":2600,"account":"SIB"},{"id":"t2347","date":"2025-06-28","category":"Achat Terrain Port Bouet","type":"Dépense","amount":121200,"account":"SIB"},{"id":"t2348","date":"2025-06-28","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Le déjeuner"},{"id":"t2349","date":"2025-06-28","category":"Divertissement","type":"Dépense","amount":26000,"account":"SIB","subcategory":"BAP"},{"id":"t3040","date":"2025-06-28","category":"Générales","type":"Dépense","amount":3090,"account":"SIB","subcategory":"Carte Money Fusion"},{"id":"t2343","date":"2025-06-29","category":"Aliments","type":"Dépense","amount":5500,"account":"SIB","note":"Joel"},{"id":"t2344","date":"2025-06-29","category":"Aliments","type":"Dépense","amount":3000,"account":"SIB"},{"id":"t2345","date":"2025-06-29","category":"Shopping","type":"Dépense","amount":37000,"account":"SIB"},{"id":"t2329","date":"2025-06-30","category":"Divertissement","type":"Dépense","amount":4050,"account":"SIB","subcategory":"Alcool"},{"id":"t2330","date":"2025-06-30","category":"Divertissement","type":"Dépense","amount":8030,"account":"SIB","subcategory":"Femme","note":"Yango Fatim"},{"id":"t2331","date":"2025-06-30","category":"Ajustement","type":"Dépense","amount":470,"account":"SIB"},{"id":"t2332","date":"2025-06-30","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t2333","date":"2025-06-30","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","note":"Marie belle"},{"id":"t2334","date":"2025-06-30","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Maman"},{"id":"t2335","date":"2025-06-30","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2336","date":"2025-06-30","category":"Shopping","type":"Dépense","amount":13000,"account":"SIB"},{"id":"t2337","date":"2025-06-30","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t2338","date":"2025-06-30","category":"GRUNDFOS","type":"Dépense","amount":87340,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2339","date":"2025-06-30","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t2340","date":"2025-06-30","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"BAP"},{"id":"t2341","date":"2025-06-30","category":"Divertissement","type":"Dépense","amount":14000,"account":"SIB","subcategory":"BAP"},{"id":"t2342","date":"2025-06-30","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Dîner"},{"id":"t2325","date":"2025-07-02","category":"Aliments","type":"Dépense","amount":12210,"account":"SIB","subcategory":"Le déjeuner"},{"id":"t2326","date":"2025-07-02","category":"Aliments","type":"Dépense","amount":1200,"account":"SIB","subcategory":"Dîner"},{"id":"t2327","date":"2025-07-02","category":"Ajustement","type":"Dépense","amount":1750,"account":"SIB"},{"id":"t2328","date":"2025-07-02","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alcool"},{"id":"t2318","date":"2025-07-03","category":"Aliments","type":"Dépense","amount":3535,"account":"SIB"},{"id":"t2319","date":"2025-07-03","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2320","date":"2025-07-03","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2321","date":"2025-07-03","category":"Divertissement","type":"Dépense","amount":11000,"account":"SIB","subcategory":"Alcool"},{"id":"t2322","date":"2025-07-03","category":"Abonnements","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Money Coach"},{"id":"t2323","date":"2025-07-03","category":"Abonnements","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Jeu D'affaire"},{"id":"t2324","date":"2025-07-03","category":"Abonnements","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Chat GPT"},{"id":"t2315","date":"2025-07-04","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Dîner"},{"id":"t2316","date":"2025-07-04","category":"Cadeaux","type":"Dépense","amount":5364,"account":"SIB","subcategory":"Femme","note":"Alvy"},{"id":"t2317","date":"2025-07-04","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Mardochee"},{"id":"t2311","date":"2025-07-05","category":"Aliments","type":"Dépense","amount":8280,"account":"SIB"},{"id":"t2312","date":"2025-07-05","category":"Ajustement","type":"Dépense","amount":3100,"account":"SIB","note":"Frais SIB"},{"id":"t2313","date":"2025-07-05","category":"Voiture","type":"Dépense","amount":500,"account":"Revenus MAZDA","subcategory":"Entretien"},{"id":"t2314","date":"2025-07-05","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Risenata"},{"id":"t2309","date":"2025-07-06","category":"Ajustement","type":"Revenu","amount":1637,"account":"SIB"},{"id":"t2310","date":"2025-07-06","category":"Divertissement","type":"Dépense","amount":3030,"account":"SIB","subcategory":"Alcool"},{"id":"t2302","date":"2025-07-07","category":"Divertissement","type":"Dépense","amount":1200,"account":"SIB","subcategory":"Alcool"},{"id":"t2303","date":"2025-07-07","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":27513,"account":"SIB","subcategory":"Boîte Postale"},{"id":"t2304","date":"2025-07-07","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SIB","subcategory":"Cotisations","note":"Mo loukou"},{"id":"t2305","date":"2025-07-07","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2306","date":"2025-07-07","category":"Abonnements","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Tinder"},{"id":"t2307","date":"2025-07-07","category":"Divertissement","type":"Dépense","amount":14000,"account":"SIB","subcategory":"BAP"},{"id":"t2308","date":"2025-07-07","category":"Aliments","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Dîner"},{"id":"t3039","date":"2025-07-07","category":"Générales","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Certificat De Perte SIB"},{"id":"t2298","date":"2025-07-08","category":"Divertissement","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Femme","note":"Tinder"},{"id":"t2299","date":"2025-07-08","category":"Divertissement","type":"Dépense","amount":6050,"account":"SIB","subcategory":"Alcool"},{"id":"t2300","date":"2025-07-08","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2301","date":"2025-07-08","category":"Cadeaux","type":"Dépense","amount":1900,"account":"SIB","payee":"Cimelec","note":"Livraison électrode"},{"id":"t2295","date":"2025-07-09","category":"Vente Pompe","type":"Revenu","amount":50000,"account":"SIB"},{"id":"t2296","date":"2025-07-09","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":700,"account":"SIB"},{"id":"t2297","date":"2025-07-09","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Cachet"},{"id":"t2290","date":"2025-07-10","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Internet","note":"Móbile"},{"id":"t2291","date":"2025-07-10","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2292","date":"2025-07-10","category":"Aliments","type":"Dépense","amount":1010,"account":"SIB","subcategory":"Dîner"},{"id":"t2293","date":"2025-07-10","category":"Divertissement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Alcool"},{"id":"t2294","date":"2025-07-10","category":"Shopping","type":"Dépense","amount":10225,"account":"SIB"},{"id":"t2283","date":"2025-07-11","category":"Divertissement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Alcool"},{"id":"t2284","date":"2025-07-11","category":"Santé","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Médicaments","note":"Deparasitant"},{"id":"t2285","date":"2025-07-11","category":"Enfants & Maman","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Hemra"},{"id":"t2286","date":"2025-07-11","category":"Aliments","type":"Dépense","amount":8000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2287","date":"2025-07-11","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2288","date":"2025-07-11","category":"Voiture","type":"Dépense","amount":2020,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2289","date":"2025-07-11","category":"Divertissement","type":"Dépense","amount":6900,"account":"SIB","subcategory":"Alcool"},{"id":"t2279","date":"2025-07-12","category":"Shopping","type":"Dépense","amount":19500,"account":"SIB","subcategory":"Alimentation"},{"id":"t2280","date":"2025-07-12","category":"Santé","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Médicaments"},{"id":"t2281","date":"2025-07-12","category":"Divertissement","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Femme","note":"Tind"},{"id":"t2282","date":"2025-07-12","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Dîner"},{"id":"t2278","date":"2025-07-13","category":"Divertissement","type":"Dépense","amount":18365,"account":"SIB","subcategory":"Alcool"},{"id":"t2277","date":"2025-07-14","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SIB","note":"Roland"},{"id":"t2273","date":"2025-07-15","category":"Invitation","type":"Dépense","amount":30500,"account":"SIB","subcategory":"Femmes","note":"Mahora"},{"id":"t2274","date":"2025-07-15","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SIB","note":"Mahora"},{"id":"t2275","date":"2025-07-15","category":"Abonnements","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Spotify"},{"id":"t2276","date":"2025-07-15","category":"Divertissement","type":"Dépense","amount":1515,"account":"SIB","subcategory":"Alcool"},{"id":"t2268","date":"2025-07-16","category":"Divertissement","type":"Dépense","amount":2600,"account":"SIB","subcategory":"Alcool"},{"id":"t2269","date":"2025-07-16","category":"Enfants & Maman","type":"Dépense","amount":13220,"account":"SIB","subcategory":"Hemra"},{"id":"t2270","date":"2025-07-16","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Cotisations","note":"bb KAMBOU"},{"id":"t2271","date":"2025-07-16","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SIB","subcategory":"Femme","note":"Tind"},{"id":"t2272","date":"2025-07-16","category":"Divertissement","type":"Dépense","amount":12000,"account":"SIB","subcategory":"BAP"},{"id":"t2265","date":"2025-07-17","category":"Aliments","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2266","date":"2025-07-17","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire","note":"Frais retrait SIB"},{"id":"t2267","date":"2025-07-17","category":"Ajustement","type":"Dépense","amount":501,"account":"SIB"},{"id":"t2263","date":"2025-07-18","category":"Formation","type":"Dépense","amount":26400,"account":"SIB","subcategory":"Emergent"},{"id":"t2264","date":"2025-07-18","category":"Invitation","type":"Dépense","amount":11000,"account":"SIB","subcategory":"Femmes","note":"Mahora"},{"id":"t2260","date":"2025-07-19","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2261","date":"2025-07-19","category":"Invitation","type":"Dépense","amount":17000,"account":"SIB","subcategory":"Femmes","note":"Tcholodjo"},{"id":"t2262","date":"2025-07-19","category":"Formation","type":"Dépense","amount":44100,"account":"SIB","subcategory":"Emergent"},{"id":"t2256","date":"2025-07-20","category":"Divertissement","type":"Dépense","amount":37000,"account":"SIB"},{"id":"t2257","date":"2025-07-20","category":"Divertissement","type":"Dépense","amount":11000,"account":"SIB"},{"id":"t2258","date":"2025-07-20","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Dîner"},{"id":"t2259","date":"2025-07-20","category":"Aliments","type":"Dépense","amount":5150,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2250","date":"2025-07-21","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t2251","date":"2025-07-21","category":"Formation","type":"Dépense","amount":33000,"account":"SIB","subcategory":"Emergent"},{"id":"t2252","date":"2025-07-21","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2253","date":"2025-07-21","category":"Formation","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Emergent"},{"id":"t2254","date":"2025-07-21","category":"Divertissement","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Femme","note":"Blessing"},{"id":"t2255","date":"2025-07-21","category":"Formation","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Emergent"},{"id":"t2248","date":"2025-07-22","category":"Formation","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Emergent"},{"id":"t2249","date":"2025-07-22","category":"Formation","type":"Dépense","amount":26000,"account":"SIB","subcategory":"Emergent"},{"id":"t2238","date":"2025-07-23","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"FIBRE"},{"id":"t2239","date":"2025-07-23","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Obed"},{"id":"t2240","date":"2025-07-23","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Tcholodjo"},{"id":"t2241","date":"2025-07-23","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t2242","date":"2025-07-23","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t2243","date":"2025-07-23","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire","note":"Frais retrait SIB"},{"id":"t2244","date":"2025-07-23","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2245","date":"2025-07-23","category":"Voiture","type":"Dépense","amount":52000,"account":"Revenus MAZDA","subcategory":"Assurance","note":"3 mois - Octobre 2025"},{"id":"t2246","date":"2025-07-23","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2247","date":"2025-07-23","category":"Voiture","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Lavage"},{"id":"t2233","date":"2025-07-24","category":"Formation","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Emergent"},{"id":"t2234","date":"2025-07-24","category":"Formation","type":"Dépense","amount":26000,"account":"SIB","subcategory":"Emergent"},{"id":"t2235","date":"2025-07-24","category":"Divertissement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Alcool"},{"id":"t2236","date":"2025-07-24","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Dîner"},{"id":"t2237","date":"2025-07-24","category":"Logement","type":"Dépense","amount":500550,"account":"Dépôt LOYER","subcategory":"Location"},{"id":"t2222","date":"2025-07-25","category":"Divertissement","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Alcool","note":"Rumba"},{"id":"t2223","date":"2025-07-25","category":"Divertissement","type":"Dépense","amount":20200,"account":"SIB","subcategory":"Alcool","note":"Rumba"},{"id":"t2224","date":"2025-07-25","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t2225","date":"2025-07-25","category":"Divertissement","type":"Dépense","amount":11948,"account":"SIB","subcategory":"Alcool"},{"id":"t2226","date":"2025-07-25","category":"Ajustement","type":"Dépense","amount":550,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t2227","date":"2025-07-25","category":"Dette","type":"Dépense","amount":120000,"account":"SIB","subcategory":"PEL"},{"id":"t2228","date":"2025-07-25","category":"Aliments","type":"Dépense","amount":10500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2229","date":"2025-07-25","category":"Un salaire","type":"Revenu","amount":1553752,"account":"SALAIRE"},{"id":"t2230","date":"2025-07-25","category":"Dette","type":"Dépense","amount":511166,"account":"SIB"},{"id":"t2231","date":"2025-07-25","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Pourboire"},{"id":"t2232","date":"2025-07-25","category":"Cadeaux","type":"Dépense","amount":11000,"account":"SIB","subcategory":"Pourboire"},{"id":"t2219","date":"2025-07-26","category":"Invitation","type":"Dépense","amount":26000,"account":"SIB","subcategory":"Femmes","note":"Déjeuner Tcholodjo"},{"id":"t2220","date":"2025-07-26","category":"Ajustement","type":"Dépense","amount":2000,"account":"SIB"},{"id":"t2221","date":"2025-07-26","category":"Invitation","type":"Dépense","amount":17000,"account":"SIB","note":"Zoh Cataleya"},{"id":"t2215","date":"2025-07-27","category":"Aliments","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2216","date":"2025-07-27","category":"Shopping","type":"Dépense","amount":4550,"account":"SIB"},{"id":"t2217","date":"2025-07-27","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t2218","date":"2025-07-27","category":"Shopping","type":"Dépense","amount":6500,"account":"SIB"},{"id":"t2211","date":"2025-07-28","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t2212","date":"2025-07-28","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Nesher"},{"id":"t2213","date":"2025-07-28","category":"Invitation","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Femmes","note":"Desi cave"},{"id":"t2214","date":"2025-07-28","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Desi cave"},{"id":"t2207","date":"2025-07-29","category":"Aliments","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2208","date":"2025-07-29","category":"Cadeaux","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Femme","note":"Emmanuella"},{"id":"t2209","date":"2025-07-29","category":"Ajustement","type":"Revenu","amount":2150,"account":"SIB"},{"id":"t2210","date":"2025-07-29","category":"Invitation","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Femmes","note":"Emmanuella"},{"id":"t2202","date":"2025-07-30","category":"Aliments","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2203","date":"2025-07-30","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t2204","date":"2025-07-30","category":"Divertissement","type":"Dépense","amount":2020,"account":"SIB","subcategory":"Femme","note":"Transport Mahora"},{"id":"t2205","date":"2025-07-30","category":"Aliments","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Dîner"},{"id":"t2206","date":"2025-07-30","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t3038","date":"2025-07-30","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":33100,"account":"SIB","subcategory":"Dédouanement"},{"id":"t2192","date":"2025-07-31","category":"Payement Maison Bingerville","type":"Dépense","amount":5300000,"account":"SALAIRE","note":"Payement 23.300.000"},{"id":"t2193","date":"2025-07-31","category":"Divertissement","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Femme","note":"Transport Fadi"},{"id":"t2194","date":"2025-07-31","category":"Ajustement","type":"Dépense","amount":420,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t2195","date":"2025-07-31","category":"Petty Cash","type":"Revenu","amount":2500000,"account":"PETTY CASH"},{"id":"t2196","date":"2025-07-31","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Triade"},{"id":"t2197","date":"2025-07-31","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Femme","note":"Fadi"},{"id":"t2198","date":"2025-07-31","category":"Shopping","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Alimentation"},{"id":"t2199","date":"2025-07-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SIB","subcategory":"Maman"},{"id":"t2200","date":"2025-07-31","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2201","date":"2025-07-31","category":"Divertissement","type":"Dépense","amount":5500,"account":"SIB","subcategory":"Alcool"},{"id":"t2186","date":"2025-08-01","category":"Ajustement","type":"Dépense","amount":120,"account":"SIB"},{"id":"t2187","date":"2025-08-01","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Internet mobile"},{"id":"t2188","date":"2025-08-01","category":"Voiture","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2189","date":"2025-08-01","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Impression"},{"id":"t2190","date":"2025-08-01","category":"Aliments","type":"Dépense","amount":3100,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2191","date":"2025-08-01","category":"Divertissement","type":"Dépense","amount":23000,"account":"SIB","subcategory":"Alcool"},{"id":"t2184","date":"2025-08-02","category":"Shopping","type":"Dépense","amount":27000,"account":"SIB","subcategory":"Alimentation"},{"id":"t2185","date":"2025-08-02","category":"Abonnements","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Money Coach"},{"id":"t2182","date":"2025-08-03","category":"Divertissement","type":"Dépense","amount":24000,"account":"SIB","subcategory":"BAP"},{"id":"t2183","date":"2025-08-03","category":"Divertissement","type":"Dépense","amount":1300,"account":"SIB","subcategory":"Alcool"},{"id":"t3037","date":"2025-08-03","category":"General","type":"Revenu","amount":28000,"account":"SIB","subcategory":"Commission Vente Peugeot 307"},{"id":"t2177","date":"2025-08-04","category":"Ajustement","type":"Dépense","amount":350,"account":"SIB","subcategory":"Frais"},{"id":"t2178","date":"2025-08-04","category":"Ajustement","type":"Dépense","amount":1870,"account":"SIB"},{"id":"t2179","date":"2025-08-04","category":"Ajustement","type":"Dépense","amount":700,"account":"SIB","subcategory":"Frais"},{"id":"t2180","date":"2025-08-04","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t2181","date":"2025-08-04","category":"Divertissement","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Femme","note":"Moon"},{"id":"t2172","date":"2025-08-05","category":"Aliments","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Dîner"},{"id":"t2173","date":"2025-08-05","category":"Divertissement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Alcool"},{"id":"t2174","date":"2025-08-05","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB"},{"id":"t2175","date":"2025-08-05","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2176","date":"2025-08-05","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t2168","date":"2025-08-06","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SIB","subcategory":"MJO","note":"Bouquet de roses"},{"id":"t2169","date":"2025-08-06","category":"Aliments","type":"Dépense","amount":6100,"account":"SIB","subcategory":"Déjeuner"},{"id":"t2170","date":"2025-08-06","category":"Vêtements","type":"Dépense","amount":16000,"account":"SIB","subcategory":"Chaussures"},{"id":"t2171","date":"2025-08-06","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","note":"Guy Nea"},{"id":"t2165","date":"2025-08-07","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Ceba"},{"id":"t2166","date":"2025-08-07","category":"Invitation","type":"Dépense","amount":15000,"account":"SIB"},{"id":"t2167","date":"2025-08-07","category":"Divertissement","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Alcool"},{"id":"t2161","date":"2025-08-08","category":"Divertissement","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Kouao"},{"id":"t2162","date":"2025-08-08","category":"Shopping","type":"Dépense","amount":30000,"account":"SIB","note":"Voyage lahou"},{"id":"t2163","date":"2025-08-08","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2164","date":"2025-08-08","category":"Divertissement","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Alcool"},{"id":"t2158","date":"2025-08-09","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","note":"Déjeuner Lahou"},{"id":"t2159","date":"2025-08-09","category":"Shopping","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Alimentation","note":"Achat poisson"},{"id":"t2160","date":"2025-08-09","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Dîner"},{"id":"t2152","date":"2025-08-10","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t2153","date":"2025-08-10","category":"Invitation","type":"Dépense","amount":25000,"account":"SIB","subcategory":"Femmes","note":"Gisèle"},{"id":"t2154","date":"2025-08-10","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","note":"Domi"},{"id":"t2155","date":"2025-08-10","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Hotel"},{"id":"t2156","date":"2025-08-10","category":"Divertissement","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Alcool"},{"id":"t2157","date":"2025-08-10","category":"Cadeaux","type":"Dépense","amount":14000,"account":"SIB","subcategory":"Femme","note":"Gisèle"},{"id":"t2146","date":"2025-08-11","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t2147","date":"2025-08-11","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t2148","date":"2025-08-11","category":"Aliments","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Dîner"},{"id":"t2149","date":"2025-08-11","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Femme","note":"Steph Lahou"},{"id":"t2150","date":"2025-08-11","category":"Ajustement","type":"Dépense","amount":62238,"account":"SIB","subcategory":"Lahou"},{"id":"t2151","date":"2025-08-11","category":"Cadeaux","type":"Dépense","amount":6000,"account":"SIB","subcategory":"Femme","note":"Sarah"},{"id":"t2141","date":"2025-08-12","category":"Shopping","type":"Dépense","amount":6000,"account":"SIB"},{"id":"t2142","date":"2025-08-12","category":"Enfants & Maman","type":"Dépense","amount":8000,"account":"SIB","subcategory":"Hemra"},{"id":"t2143","date":"2025-08-12","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t2144","date":"2025-08-12","category":"Enfants & Maman","type":"Dépense","amount":26500,"account":"SIB","subcategory":"Hemra"},{"id":"t2145","date":"2025-08-12","category":"Aliments","type":"Dépense","amount":4500,"account":"SIB"},{"id":"t2135","date":"2025-08-13","category":"Aliments","type":"Dépense","amount":11500,"account":"SIB","subcategory":"Dîner"},{"id":"t2136","date":"2025-08-13","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Cotisations","note":"Jean Privat"},{"id":"t2137","date":"2025-08-13","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":10100,"account":"SIB"},{"id":"t2138","date":"2025-08-13","category":"INVEST SGO","type":"Dépense","amount":84600,"account":"SGO","subcategory":"NSIA"},{"id":"t2139","date":"2025-08-13","category":"Ajustement","type":"Dépense","amount":3615,"account":"SIB","subcategory":"Frais"},{"id":"t2140","date":"2025-08-13","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2130","date":"2025-08-14","category":"Aliments","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2131","date":"2025-08-14","category":"Divertissement","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Femme","note":"Moon"},{"id":"t2132","date":"2025-08-14","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Alcool","note":"Ismo"},{"id":"t2133","date":"2025-08-14","category":"Vêtements","type":"Dépense","amount":55000,"account":"SALAIRE"},{"id":"t2134","date":"2025-08-14","category":"Cadeaux","type":"Dépense","amount":107000,"account":"SALAIRE","subcategory":"Femme","note":"Anniversaire Juliet"},{"id":"t2128","date":"2025-08-15","category":"Divertissement","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2129","date":"2025-08-15","category":"Invitation","type":"Dépense","amount":11000,"account":"SALAIRE","note":"Privat / Deladet"},{"id":"t2123","date":"2025-08-16","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t2124","date":"2025-08-16","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2125","date":"2025-08-16","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2126","date":"2025-08-16","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Femme","note":"Transport Grâce"},{"id":"t2127","date":"2025-08-16","category":"Invitation","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Femmes","note":"Lesly"},{"id":"t2115","date":"2025-08-17","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2116","date":"2025-08-17","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE"},{"id":"t2117","date":"2025-08-17","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Femme Drissa"},{"id":"t2118","date":"2025-08-17","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Cotisations","note":"Village"},{"id":"t2119","date":"2025-08-17","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Oncle"},{"id":"t2120","date":"2025-08-17","category":"Divertissement","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2121","date":"2025-08-17","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2122","date":"2025-08-17","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Femme","note":"Ela"},{"id":"t2103","date":"2025-08-18","category":"INVEST SGO","type":"Dépense","amount":30543,"account":"SGO","subcategory":"NSIA"},{"id":"t2104","date":"2025-08-18","category":"Aliments","type":"Dépense","amount":3600,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2105","date":"2025-08-18","category":"GRUNDFOS","type":"Dépense","amount":38500,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2106","date":"2025-08-18","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","note":"Tonton Hassan"},{"id":"t2107","date":"2025-08-18","category":"Générales","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Yango Livraison"},{"id":"t2108","date":"2025-08-18","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2109","date":"2025-08-18","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2110","date":"2025-08-18","category":"Ajustement","type":"Dépense","amount":2200,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t2111","date":"2025-08-18","category":"Enfants & Maman","type":"Dépense","amount":25250,"account":"SALAIRE","subcategory":"Nesher","note":"Inscription Nesher"},{"id":"t2112","date":"2025-08-18","category":"Ajustement","type":"Dépense","amount":10046,"account":"SALAIRE"},{"id":"t2113","date":"2025-08-18","category":"Cadeaux","type":"Dépense","amount":13100,"account":"SALAIRE","subcategory":"Femme","note":"Desi"},{"id":"t2114","date":"2025-08-18","category":"Divertissement","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2102","date":"2025-08-19","category":"GRUNDFOS","type":"Dépense","amount":51000,"account":"PETTY CASH","subcategory":"Voyage","note":"Passport David"},{"id":"t2098","date":"2025-08-20","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t2099","date":"2025-08-20","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"Sarah"},{"id":"t2100","date":"2025-08-20","category":"Shopping","type":"Dépense","amount":45000,"account":"SALAIRE"},{"id":"t2101","date":"2025-08-20","category":"Divertissement","type":"Dépense","amount":50000,"account":"SALAIRE","subcategory":"Anniversaire","note":"Kindoki"},{"id":"t2090","date":"2025-08-21","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Femme","note":"Lesly"},{"id":"t2091","date":"2025-08-21","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2092","date":"2025-08-21","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Internet"},{"id":"t2093","date":"2025-08-21","category":"Ajustement","type":"Dépense","amount":3003,"account":"SALAIRE"},{"id":"t2094","date":"2025-08-21","category":"Générales","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Police"},{"id":"t2095","date":"2025-08-21","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2096","date":"2025-08-21","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Saturnin"},{"id":"t2097","date":"2025-08-21","category":"Aliments","type":"Dépense","amount":10500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2087","date":"2025-08-22","category":"Ajustement","type":"Dépense","amount":2200,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t2088","date":"2025-08-22","category":"Logement","type":"Dépense","amount":475550,"account":"Dépôt LOYER","subcategory":"Location","note":"Septembre"},{"id":"t2089","date":"2025-08-22","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2079","date":"2025-08-23","category":"Shopping","type":"Dépense","amount":9000,"account":"SALAIRE","note":"Drap"},{"id":"t2080","date":"2025-08-23","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2081","date":"2025-08-23","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t2082","date":"2025-08-23","category":"Générales","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Visite Maison"},{"id":"t2083","date":"2025-08-23","category":"Enfants & Maman","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t2084","date":"2025-08-23","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2085","date":"2025-08-23","category":"GRUNDFOS","type":"Dépense","amount":50840,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t2086","date":"2025-08-23","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t2070","date":"2025-08-24","category":"Shopping","type":"Dépense","amount":13000,"account":"SALAIRE"},{"id":"t2071","date":"2025-08-24","category":"Divertissement","type":"Dépense","amount":25500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2072","date":"2025-08-24","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2073","date":"2025-08-24","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t2074","date":"2025-08-24","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2075","date":"2025-08-24","category":"Générales","type":"Dépense","amount":2020,"account":"SIB","note":"Câble chargeur"},{"id":"t2076","date":"2025-08-24","category":"Divertissement","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2077","date":"2025-08-24","category":"Divertissement","type":"Dépense","amount":14700,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2078","date":"2025-08-24","category":"Divertissement","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2060","date":"2025-08-25","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Steph"},{"id":"t2061","date":"2025-08-25","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Impression"},{"id":"t2062","date":"2025-08-25","category":"Générales","type":"Dépense","amount":500,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t2063","date":"2025-08-25","category":"Santé","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t2064","date":"2025-08-25","category":"INVEST SGO","type":"Dépense","amount":101808,"account":"SGO","subcategory":"NSIA"},{"id":"t2065","date":"2025-08-25","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t2066","date":"2025-08-25","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SALAIRE"},{"id":"t2067","date":"2025-08-25","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SALAIRE"},{"id":"t2068","date":"2025-08-25","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2069","date":"2025-08-25","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2056","date":"2025-08-26","category":"Divertissement","type":"Dépense","amount":33500,"account":"SALAIRE","subcategory":"Femme"},{"id":"t2057","date":"2025-08-26","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2058","date":"2025-08-26","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2059","date":"2025-08-26","category":"Invitation","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Femmes","note":"Steph"},{"id":"t2052","date":"2025-08-27","category":"Ajustement","type":"Dépense","amount":1834,"account":"SIB","subcategory":"Frais","note":"Carte bancaire"},{"id":"t2053","date":"2025-08-27","category":"Aliments","type":"Dépense","amount":5700,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2054","date":"2025-08-27","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2055","date":"2025-08-27","category":"GRUNDFOS","type":"Dépense","amount":2253,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t2049","date":"2025-08-28","category":"Divertissement","type":"Dépense","amount":2100,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2050","date":"2025-08-28","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2051","date":"2025-08-28","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","note":"Franck"},{"id":"t2041","date":"2025-08-29","category":"Divertissement","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2042","date":"2025-08-29","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"Lesly"},{"id":"t2043","date":"2025-08-29","category":"Ajustement","type":"Dépense","amount":12634,"account":"SALAIRE"},{"id":"t2044","date":"2025-08-29","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t2045","date":"2025-08-29","category":"Aliments","type":"Dépense","amount":6500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2046","date":"2025-08-29","category":"GRUNDFOS","type":"Dépense","amount":360500,"account":"PETTY CASH","subcategory":"Hotel","note":"Réservation Azalaï"},{"id":"t2047","date":"2025-08-29","category":"Ajustement","type":"Dépense","amount":2200,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t2048","date":"2025-08-29","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Infraction"},{"id":"t2036","date":"2025-08-30","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2037","date":"2025-08-30","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t2038","date":"2025-08-30","category":"Aliments","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2039","date":"2025-08-30","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Moon"},{"id":"t2040","date":"2025-08-30","category":"Divertissement","type":"Dépense","amount":55000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t2020","date":"2025-08-31","category":"Loyer","type":"Revenu","amount":1500000,"account":"Dépôt LOYER"},{"id":"t2021","date":"2025-08-31","category":"Invitation","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Femmes","note":"MB déjeuner 01/09"},{"id":"t2022","date":"2025-08-31","category":"Aliments","type":"Dépense","amount":17000,"account":"SALAIRE","subcategory":"Déjeuner","note":"01/09"},{"id":"t2023","date":"2025-08-31","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Femme","note":"Maeva 01/09"},{"id":"t2024","date":"2025-08-31","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"MB 01/09"},{"id":"t2025","date":"2025-08-31","category":"Enfants & Maman","type":"Dépense","amount":60000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t2026","date":"2025-08-31","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Alcool","note":"Olo 01/09"},{"id":"t2027","date":"2025-08-31","category":"Cadeaux","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Pourboire","note":"01/09"},{"id":"t2028","date":"2025-08-31","category":"Santé","type":"Dépense","amount":14000,"account":"SALAIRE","subcategory":"VG","note":"01/09"},{"id":"t2029","date":"2025-08-31","category":"Ajustement","type":"Revenu","amount":28882,"account":"SALAIRE"},{"id":"t2030","date":"2025-08-31","category":"Divertissement","type":"Dépense","amount":27000,"account":"SALAIRE","subcategory":"Alcool","note":"Olo le 01/09"},{"id":"t2031","date":"2025-08-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Maman"},{"id":"t2032","date":"2025-08-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Nesher","note":"01/09"},{"id":"t2033","date":"2025-08-31","category":"Cadeaux","type":"Dépense","amount":10365,"account":"SALAIRE","note":"Tonton Moussa"},{"id":"t2034","date":"2025-08-31","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Cotisations","note":"Marc"},{"id":"t2035","date":"2025-08-31","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2019","date":"2025-09-01","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t2013","date":"2025-09-02","category":"Shopping","type":"Dépense","amount":15150,"account":"SALAIRE"},{"id":"t2014","date":"2025-09-02","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t2015","date":"2025-09-02","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2016","date":"2025-09-02","category":"Shopping","type":"Dépense","amount":12000,"account":"SALAIRE"},{"id":"t2017","date":"2025-09-02","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t2018","date":"2025-09-02","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t2009","date":"2025-09-03","category":"Santé","type":"Dépense","amount":8500,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t2010","date":"2025-09-03","category":"GRUNDFOS","type":"Dépense","amount":7000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Calepin"},{"id":"t2011","date":"2025-09-03","category":"Aliments","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t2012","date":"2025-09-03","category":"GRUNDFOS","type":"Dépense","amount":18000,"account":"PETTY CASH","note":"Pocket wifi"},{"id":"t2008","date":"2025-09-04","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","note":"Rasoir David"},{"id":"t2001","date":"2025-09-05","category":"Abonnements","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t2002","date":"2025-09-05","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t2003","date":"2025-09-05","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Metty"},{"id":"t2004","date":"2025-09-05","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Hotel","note":"Transport Delali"},{"id":"t2005","date":"2025-09-05","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Restaurant","note":"Azalai"},{"id":"t2006","date":"2025-09-05","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","note":"Zoh"},{"id":"t2007","date":"2025-09-05","category":"GRUNDFOS","type":"Dépense","amount":70000,"account":"PETTY CASH","subcategory":"Enjoy","note":"David"},{"id":"t1997","date":"2025-09-06","category":"Ajustement","type":"Dépense","amount":1500,"account":"SIB"},{"id":"t1998","date":"2025-09-06","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1999","date":"2025-09-06","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t2000","date":"2025-09-06","category":"Aliments","type":"Dépense","amount":1800,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1990","date":"2025-09-07","category":"Générales","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Visite Maison"},{"id":"t1991","date":"2025-09-07","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1992","date":"2025-09-07","category":"Abonnements","type":"Dépense","amount":2400,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t1993","date":"2025-09-07","category":"Abonnements","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t1994","date":"2025-09-07","category":"Ajustement","type":"Dépense","amount":300,"account":"SIB"},{"id":"t1995","date":"2025-09-07","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1996","date":"2025-09-07","category":"Divertissement","type":"Dépense","amount":16050,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1985","date":"2025-09-08","category":"Abonnements","type":"Dépense","amount":1877,"account":"SIB","subcategory":"Assurance SAF"},{"id":"t1986","date":"2025-09-08","category":"Invitation","type":"Dépense","amount":25000,"account":"SALAIRE","note":"Elvis"},{"id":"t1987","date":"2025-09-08","category":"Divertissement","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1988","date":"2025-09-08","category":"Aliments","type":"Dépense","amount":3530,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1989","date":"2025-09-08","category":"Cadeaux","type":"Dépense","amount":7610,"account":"SALAIRE","subcategory":"Metty"},{"id":"t1974","date":"2025-09-09","category":"Vêtements","type":"Dépense","amount":3000,"account":"SALAIRE"},{"id":"t1975","date":"2025-09-09","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1976","date":"2025-09-09","category":"Cadeaux","type":"Dépense","amount":200000,"account":"SALAIRE","subcategory":"Metty","note":"Metty"},{"id":"t1977","date":"2025-09-09","category":"Divertissement","type":"Dépense","amount":6500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1978","date":"2025-09-09","category":"Générales","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Visite Maison"},{"id":"t1979","date":"2025-09-09","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1980","date":"2025-09-09","category":"Aliments","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1981","date":"2025-09-09","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1982","date":"2025-09-09","category":"Ajustement","type":"Dépense","amount":3101,"account":"SALAIRE"},{"id":"t1983","date":"2025-09-09","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1984","date":"2025-09-09","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1969","date":"2025-09-10","category":"Cadeaux","type":"Dépense","amount":25250,"account":"SALAIRE","subcategory":"Femme","note":"Lesly"},{"id":"t1970","date":"2025-09-10","category":"Aliments","type":"Dépense","amount":6010,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1971","date":"2025-09-10","category":"Générales","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Visite Maison"},{"id":"t1972","date":"2025-09-10","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1973","date":"2025-09-10","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1958","date":"2025-09-11","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Metty","note":"Déjeuner"},{"id":"t1959","date":"2025-09-11","category":"INVEST SGO","type":"Dépense","amount":1808,"account":"SGO","subcategory":"Frais"},{"id":"t1960","date":"2025-09-11","category":"INVEST SGO","type":"Dépense","amount":100000,"account":"SGO","subcategory":"NSIA"},{"id":"t1961","date":"2025-09-11","category":"Générales","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Visite Maison"},{"id":"t1962","date":"2025-09-11","category":"Déménagement","type":"Dépense","amount":2125000,"account":"SALAIRE","note":"MORIJAH"},{"id":"t1963","date":"2025-09-11","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1964","date":"2025-09-11","category":"Ajustement","type":"Dépense","amount":905,"account":"SALAIRE"},{"id":"t1965","date":"2025-09-11","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t1966","date":"2025-09-11","category":"Cadeaux","type":"Dépense","amount":2020,"account":"SALAIRE","note":"Marc"},{"id":"t1967","date":"2025-09-11","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1968","date":"2025-09-11","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1955","date":"2025-09-12","category":"Divertissement","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1956","date":"2025-09-12","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Moon"},{"id":"t1957","date":"2025-09-12","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Cotisations","note":"Dawkins"},{"id":"t1947","date":"2025-09-13","category":"Cadeaux","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Metty","note":"Déjeuner"},{"id":"t1948","date":"2025-09-13","category":"Utilitaires","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Téléphones","note":"Antichoc"},{"id":"t1949","date":"2025-09-13","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1950","date":"2025-09-13","category":"Aliments","type":"Dépense","amount":2600,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1951","date":"2025-09-13","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1952","date":"2025-09-13","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1953","date":"2025-09-13","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alcool"},{"id":"t1954","date":"2025-09-13","category":"Invitation","type":"Dépense","amount":30500,"account":"SALAIRE","subcategory":"Femmes","note":"Metty"},{"id":"t1944","date":"2025-09-14","category":"Aliments","type":"Dépense","amount":5200,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1945","date":"2025-09-14","category":"Divertissement","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1946","date":"2025-09-14","category":"Divertissement","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1941","date":"2025-09-15","category":"Revenu général","type":"Revenu","amount":1443037,"account":"SIB","subcategory":"Rappel PEL"},{"id":"t1942","date":"2025-09-15","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Metty"},{"id":"t1943","date":"2025-09-15","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1934","date":"2025-09-16","category":"Divertissement","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1935","date":"2025-09-16","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t1936","date":"2025-09-16","category":"GRUNDFOS","type":"Dépense","amount":80000,"account":"PETTY CASH","subcategory":"Hotel"},{"id":"t1937","date":"2025-09-16","category":"INVEST SGO","type":"Dépense","amount":181,"account":"SGO","subcategory":"Frais"},{"id":"t1938","date":"2025-09-16","category":"INVEST SGO","type":"Dépense","amount":10000,"account":"SGO","subcategory":"NSIA"},{"id":"t1939","date":"2025-09-16","category":"INVEST SGO","type":"Dépense","amount":723,"account":"SGO","subcategory":"Frais"},{"id":"t1940","date":"2025-09-16","category":"INVEST SGO","type":"Dépense","amount":40000,"account":"SGO","subcategory":"NSIA"},{"id":"t1927","date":"2025-09-17","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Femme","note":"Ruth Panini"},{"id":"t1928","date":"2025-09-17","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t1929","date":"2025-09-17","category":"Plan Éducation","type":"Dépense","amount":30000,"account":"SIB","subcategory":"PEL"},{"id":"t1930","date":"2025-09-17","category":"Plan Éducation","type":"Dépense","amount":2457,"account":"SIB","subcategory":"PEL"},{"id":"t1931","date":"2025-09-17","category":"Aliments","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1932","date":"2025-09-17","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1933","date":"2025-09-17","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Péage","note":"Péage"},{"id":"t1925","date":"2025-09-18","category":"Divertissement","type":"Dépense","amount":17000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1926","date":"2025-09-18","category":"Cadeaux","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Metty","note":"Dîner"},{"id":"t1912","date":"2025-09-19","category":"Shopping","type":"Dépense","amount":16802,"account":"SALAIRE"},{"id":"t1913","date":"2025-09-19","category":"Générales","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Péage"},{"id":"t1914","date":"2025-09-19","category":"Aliments","type":"Dépense","amount":9000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1915","date":"2025-09-19","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1916","date":"2025-09-19","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Lia"},{"id":"t1917","date":"2025-09-19","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Nesher","note":"Scolarité"},{"id":"t1918","date":"2025-09-19","category":"GRUNDFOS","type":"Dépense","amount":20200,"account":"PETTY CASH","subcategory":"AUTRES","note":"Navette Cisse"},{"id":"t1919","date":"2025-09-19","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1920","date":"2025-09-19","category":"Aliments","type":"Dépense","amount":650,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1921","date":"2025-09-19","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":520000,"account":"SIB"},{"id":"t1922","date":"2025-09-19","category":"Aliments","type":"Dépense","amount":1210,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1923","date":"2025-09-19","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1924","date":"2025-09-19","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1908","date":"2025-09-20","category":"Enfants & Maman","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Hemra","note":"Yaourt"},{"id":"t1909","date":"2025-09-20","category":"Invitation","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Femmes","note":"Yango Ornella"},{"id":"t1910","date":"2025-09-20","category":"Invitation","type":"Dépense","amount":44500,"account":"SALAIRE","note":"Charly"},{"id":"t1911","date":"2025-09-20","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"La musique","note":"Rumba"},{"id":"t1905","date":"2025-09-21","category":"Invitation","type":"Dépense","amount":26000,"account":"SALAIRE"},{"id":"t1906","date":"2025-09-21","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1907","date":"2025-09-21","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Ruth"},{"id":"t1894","date":"2025-09-22","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1895","date":"2025-09-22","category":"Cadeaux","type":"Dépense","amount":32700,"account":"SALAIRE","subcategory":"Metty","note":"Clinique"},{"id":"t1896","date":"2025-09-22","category":"GRUNDFOS","type":"Dépense","amount":5700,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1897","date":"2025-09-22","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Cotisations","note":"Mère Patrice"},{"id":"t1898","date":"2025-09-22","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1899","date":"2025-09-22","category":"Allocation","type":"Revenu","amount":120000,"account":"SALAIRE","subcategory":"Avoir Azalai"},{"id":"t1900","date":"2025-09-22","category":"Générales","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Péage"},{"id":"t1901","date":"2025-09-22","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1902","date":"2025-09-22","category":"Divertissement","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1903","date":"2025-09-22","category":"Cadeaux","type":"Dépense","amount":18500,"account":"SALAIRE","subcategory":"Metty","note":"Shopping"},{"id":"t1904","date":"2025-09-22","category":"Cadeaux","type":"Dépense","amount":18500,"account":"SALAIRE","subcategory":"Metty","note":"Pharmacie"},{"id":"t1887","date":"2025-09-23","category":"Voiture","type":"Dépense","amount":112100,"account":"Revenus MAZDA","subcategory":"Visite Technique"},{"id":"t1888","date":"2025-09-23","category":"Voiture","type":"Dépense","amount":45000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Pneus"},{"id":"t1889","date":"2025-09-23","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1890","date":"2025-09-23","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1891","date":"2025-09-23","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"Lesly"},{"id":"t1892","date":"2025-09-23","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1893","date":"2025-09-23","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1882","date":"2025-09-24","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t1883","date":"2025-09-24","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1884","date":"2025-09-24","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":65650,"account":"SIB","note":"Elvis"},{"id":"t1885","date":"2025-09-24","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1886","date":"2025-09-24","category":"Ajustement","type":"Dépense","amount":4110,"account":"SIB"},{"id":"t1877","date":"2025-09-25","category":"Vente Pompe","type":"Revenu","amount":150000,"account":"SALAIRE"},{"id":"t1878","date":"2025-09-25","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1879","date":"2025-09-25","category":"Déménagement","type":"Dépense","amount":15150,"account":"Revenus MAZDA"},{"id":"t1880","date":"2025-09-25","category":"Déménagement","type":"Dépense","amount":1250000,"account":"Revenus MAZDA"},{"id":"t1881","date":"2025-09-25","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1873","date":"2025-09-26","category":"Divertissement","type":"Dépense","amount":41000,"account":"SIB","subcategory":"La musique"},{"id":"t1874","date":"2025-09-26","category":"Invitation","type":"Dépense","amount":23500,"account":"SIB","note":"Prisca YAO"},{"id":"t1875","date":"2025-09-26","category":"Invitation","type":"Dépense","amount":8000,"account":"SIB","subcategory":"Femmes","note":"Sylvia"},{"id":"t1876","date":"2025-09-26","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SIB","subcategory":"Metty"},{"id":"t1869","date":"2025-09-27","category":"Divertissement","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Alcool"},{"id":"t1870","date":"2025-09-27","category":"Shopping","type":"Dépense","amount":1500,"account":"SIB"},{"id":"t1871","date":"2025-09-27","category":"Aliments","type":"Dépense","amount":16000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1872","date":"2025-09-27","category":"Divertissement","type":"Dépense","amount":52200,"account":"SIB","subcategory":"La musique"},{"id":"t3035","date":"2025-09-27","category":"Déménagement","type":"Dépense","amount":85500,"account":"SIB","subcategory":"Installation Clim Chauffe Eau"},{"id":"t3036","date":"2025-09-27","category":"Déménagement","type":"Dépense","amount":140100,"account":"SIB","subcategory":"Remplacement Gaziniere"},{"id":"t1864","date":"2025-09-28","category":"Déménagement","type":"Dépense","amount":40000,"account":"SIB","subcategory":"Nettoyage"},{"id":"t1865","date":"2025-09-28","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t1866","date":"2025-09-28","category":"Divertissement","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Alcool","note":"Lingot d'or"},{"id":"t1867","date":"2025-09-28","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Dîner","note":"Elvis"},{"id":"t1868","date":"2025-09-28","category":"Déménagement","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Lits"},{"id":"t1861","date":"2025-09-29","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1862","date":"2025-09-29","category":"Invitation","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Femmes","note":"Emmanuella"},{"id":"t1863","date":"2025-09-29","category":"Divertissement","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Femme","note":"Nadia"},{"id":"t1852","date":"2025-09-30","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Femme","note":"Emmanuella"},{"id":"t1853","date":"2025-09-30","category":"Déménagement","type":"Dépense","amount":2000,"account":"SIB","note":"Aide"},{"id":"t1854","date":"2025-09-30","category":"Aliments","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1855","date":"2025-09-30","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1856","date":"2025-09-30","category":"Déménagement","type":"Dépense","amount":38000,"account":"SIB"},{"id":"t1857","date":"2025-09-30","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t1858","date":"2025-09-30","category":"Cadeaux","type":"Dépense","amount":25250,"account":"SIB","subcategory":"Mardochee"},{"id":"t1859","date":"2025-09-30","category":"Enfants & Maman","type":"Dépense","amount":50500,"account":"SIB","subcategory":"Nesher"},{"id":"t1860","date":"2025-09-30","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"Revenus MAZDA","subcategory":"Maman"},{"id":"t1842","date":"2025-10-01","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Dîner"},{"id":"t1843","date":"2025-10-01","category":"Personnel","type":"Dépense","amount":11615,"account":"Revenus MAZDA","subcategory":"Produits de beauté","note":"Savon bio"},{"id":"t1844","date":"2025-10-01","category":"Aliments","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1845","date":"2025-10-01","category":"Déménagement","type":"Dépense","amount":18000,"account":"SIB","subcategory":"Deco & Senteur"},{"id":"t1846","date":"2025-10-01","category":"Déménagement","type":"Dépense","amount":25500,"account":"Dépôt LOYER","subcategory":"Electricien"},{"id":"t1847","date":"2025-10-01","category":"Voiture","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Lavage"},{"id":"t1848","date":"2025-10-01","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Metty"},{"id":"t1849","date":"2025-10-01","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1850","date":"2025-10-01","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Olokpacha"},{"id":"t1851","date":"2025-10-01","category":"Invitation","type":"Dépense","amount":46000,"account":"Revenus MAZDA","note":"Paula"},{"id":"t1835","date":"2025-10-02","category":"Shopping","type":"Dépense","amount":3000,"account":"SIB"},{"id":"t1836","date":"2025-10-02","category":"Déménagement","type":"Dépense","amount":34000,"account":"Revenus MAZDA","note":"Bouteille de Gaz"},{"id":"t1837","date":"2025-10-02","category":"Cadeaux","type":"Dépense","amount":10100,"account":"Revenus MAZDA","subcategory":"Metty"},{"id":"t1838","date":"2025-10-02","category":"Divertissement","type":"Dépense","amount":1500,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1839","date":"2025-10-02","category":"Aliments","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1840","date":"2025-10-02","category":"Enfants & Maman","type":"Dépense","amount":40000,"account":"Revenus MAZDA","subcategory":"Hemra"},{"id":"t1841","date":"2025-10-02","category":"Déménagement","type":"Dépense","amount":70000,"account":"Revenus MAZDA","note":"Deco/gazon/Senteur"},{"id":"t3034","date":"2025-10-02","category":"Déménagement","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Installation Clim Chauffe Eau"},{"id":"t1828","date":"2025-10-03","category":"Shopping","type":"Dépense","amount":15000,"account":"Revenus MAZDA","note":"Jus Paula"},{"id":"t1829","date":"2025-10-03","category":"Générales","type":"Dépense","amount":1600,"account":"Revenus MAZDA","note":"Yango transport"},{"id":"t1830","date":"2025-10-03","category":"Invitation","type":"Dépense","amount":40000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Paula"},{"id":"t1831","date":"2025-10-03","category":"Déménagement","type":"Dépense","amount":43000,"account":"Revenus MAZDA"},{"id":"t1832","date":"2025-10-03","category":"Aliments","type":"Dépense","amount":7000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1833","date":"2025-10-03","category":"Cadeaux","type":"Dépense","amount":20000,"account":"Revenus MAZDA","note":"Mosso"},{"id":"t1834","date":"2025-10-03","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t1825","date":"2025-10-04","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1826","date":"2025-10-04","category":"Cadeaux","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Metty"},{"id":"t1827","date":"2025-10-04","category":"Aliments","type":"Dépense","amount":500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1821","date":"2025-10-05","category":"Divertissement","type":"Dépense","amount":9000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1822","date":"2025-10-05","category":"Cadeaux","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Desiree"},{"id":"t1823","date":"2025-10-05","category":"Aliments","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1824","date":"2025-10-05","category":"Aliments","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1814","date":"2025-10-06","category":"Aliments","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1815","date":"2025-10-06","category":"Divertissement","type":"Dépense","amount":8000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1816","date":"2025-10-06","category":"Divertissement","type":"Dépense","amount":30000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Nikki"},{"id":"t1817","date":"2025-10-06","category":"Divertissement","type":"Dépense","amount":15000,"account":"Revenus MAZDA","subcategory":"Residence"},{"id":"t1818","date":"2025-10-06","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1819","date":"2025-10-06","category":"Cadeaux","type":"Dépense","amount":10100,"account":"Revenus MAZDA","subcategory":"Femme","note":"Kim"},{"id":"t1820","date":"2025-10-06","category":"Enfants & Maman","type":"Dépense","amount":15150,"account":"Revenus MAZDA","subcategory":"Nesher","note":"Cantine Nesher"},{"id":"t1807","date":"2025-10-07","category":"Abonnements","type":"Dépense","amount":2400,"account":"Revenus MAZDA","subcategory":"Spotify"},{"id":"t1808","date":"2025-10-07","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t1809","date":"2025-10-07","category":"Shopping","type":"Dépense","amount":1500,"account":"Revenus MAZDA"},{"id":"t1810","date":"2025-10-07","category":"Cadeaux","type":"Dépense","amount":5050,"account":"Revenus MAZDA","note":"Parker"},{"id":"t1811","date":"2025-10-07","category":"Aliments","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1812","date":"2025-10-07","category":"Divertissement","type":"Dépense","amount":30000,"account":"Revenus MAZDA","subcategory":"Alcool","note":"Emmanuella"},{"id":"t1813","date":"2025-10-07","category":"Divertissement","type":"Dépense","amount":20000,"account":"Revenus MAZDA","note":"Dydime"},{"id":"t1802","date":"2025-10-08","category":"Divertissement","type":"Dépense","amount":15000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Emmanuella"},{"id":"t1803","date":"2025-10-08","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1804","date":"2025-10-08","category":"Cadeaux","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Paula"},{"id":"t1805","date":"2025-10-08","category":"Divertissement","type":"Dépense","amount":6000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1806","date":"2025-10-08","category":"Aliments","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1796","date":"2025-10-09","category":"Invitation","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Sandra"},{"id":"t1797","date":"2025-10-09","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1798","date":"2025-10-09","category":"Générales","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t1799","date":"2025-10-09","category":"Voiture","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1800","date":"2025-10-09","category":"Abonnements","type":"Dépense","amount":5500,"account":"Revenus MAZDA","subcategory":"Money Coach"},{"id":"t1801","date":"2025-10-09","category":"Abonnements","type":"Dépense","amount":6000,"account":"Revenus MAZDA","subcategory":"Tinder"},{"id":"t1794","date":"2025-10-10","category":"Divertissement","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1795","date":"2025-10-10","category":"Aliments","type":"Dépense","amount":2500,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1791","date":"2025-10-11","category":"Divertissement","type":"Dépense","amount":15000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1792","date":"2025-10-11","category":"Aliments","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1793","date":"2025-10-11","category":"Déménagement","type":"Dépense","amount":40500,"account":"Revenus MAZDA","subcategory":"Étagère Cuisine"},{"id":"t1787","date":"2025-10-12","category":"Aliments","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1788","date":"2025-10-12","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Pneu"},{"id":"t1789","date":"2025-10-12","category":"Aliments","type":"Dépense","amount":2500,"account":"Revenus MAZDA"},{"id":"t1790","date":"2025-10-12","category":"Divertissement","type":"Dépense","amount":22000,"account":"Revenus MAZDA","subcategory":"BAP"},{"id":"t1784","date":"2025-10-13","category":"Aliments","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1785","date":"2025-10-13","category":"Divertissement","type":"Dépense","amount":1500,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1786","date":"2025-10-13","category":"Cadeaux","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Sandra"},{"id":"t1781","date":"2025-10-14","category":"Vente Pompe","type":"Revenu","amount":90000,"account":"SALAIRE"},{"id":"t1782","date":"2025-10-14","category":"Aliments","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1783","date":"2025-10-14","category":"Santé","type":"Dépense","amount":3500,"account":"Revenus MAZDA","subcategory":"Médicaments"},{"id":"t1775","date":"2025-10-15","category":"Shopping","type":"Dépense","amount":2500,"account":"Revenus MAZDA"},{"id":"t1776","date":"2025-10-15","category":"Divertissement","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Rue des B"},{"id":"t1777","date":"2025-10-15","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1778","date":"2025-10-15","category":"Cadeaux","type":"Dépense","amount":14000,"account":"Revenus MAZDA","subcategory":"Pourboire"},{"id":"t1779","date":"2025-10-15","category":"Divertissement","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1780","date":"2025-10-15","category":"Divertissement","type":"Dépense","amount":19000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1772","date":"2025-10-16","category":"Générales","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t1773","date":"2025-10-16","category":"Aliments","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1774","date":"2025-10-16","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t1764","date":"2025-10-17","category":"Enfants & Maman","type":"Dépense","amount":5050,"account":"Revenus MAZDA","subcategory":"Nesher"},{"id":"t1765","date":"2025-10-17","category":"Cadeaux","type":"Dépense","amount":1000,"account":"Revenus MAZDA"},{"id":"t1766","date":"2025-10-17","category":"Divertissement","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1767","date":"2025-10-17","category":"Aliments","type":"Dépense","amount":12100,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1768","date":"2025-10-17","category":"Invitation","type":"Dépense","amount":20000,"account":"Revenus MAZDA","note":"Chez Donia"},{"id":"t1769","date":"2025-10-17","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1770","date":"2025-10-17","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1771","date":"2025-10-17","category":"Divertissement","type":"Dépense","amount":35000,"account":"Revenus MAZDA","subcategory":"La musique"},{"id":"t1752","date":"2025-10-18","category":"Ajustement","type":"Dépense","amount":12508,"account":"SIB"},{"id":"t1753","date":"2025-10-18","category":"Aliments","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Invitation"},{"id":"t1754","date":"2025-10-18","category":"Cadeaux","type":"Dépense","amount":7260,"account":"Revenus MAZDA","subcategory":"Femme","note":"Internet Vanessa"},{"id":"t1755","date":"2025-10-18","category":"Aliments","type":"Dépense","amount":6500,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1756","date":"2025-10-18","category":"Divertissement","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1757","date":"2025-10-18","category":"Divertissement","type":"Dépense","amount":30000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Bad"},{"id":"t1758","date":"2025-10-18","category":"Divertissement","type":"Dépense","amount":40000,"account":"Revenus MAZDA","subcategory":"La musique"},{"id":"t1759","date":"2025-10-18","category":"Invitation","type":"Dépense","amount":25000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Vanessa"},{"id":"t1760","date":"2025-10-18","category":"Cadeaux","type":"Dépense","amount":15150,"account":"Revenus MAZDA","subcategory":"Femme","note":"Yasmine Keita"},{"id":"t1761","date":"2025-10-18","category":"Cadeaux","type":"Dépense","amount":26000,"account":"Revenus MAZDA","subcategory":"Anniversaire","note":"Bb Yannick"},{"id":"t1762","date":"2025-10-18","category":"Prêt Orange","type":"Revenu","amount":63464,"account":"SIB"},{"id":"t1763","date":"2025-10-18","category":"Ajustement","type":"Dépense","amount":122286,"account":"Revenus MAZDA","subcategory":"Étrange"},{"id":"t1747","date":"2025-10-19","category":"Cadeaux","type":"Dépense","amount":50000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Vanessa"},{"id":"t1748","date":"2025-10-19","category":"Cadeaux","type":"Dépense","amount":21000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Médicaments Paula"},{"id":"t1749","date":"2025-10-19","category":"Invitation","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Femmes","note":"Vanessa"},{"id":"t1750","date":"2025-10-19","category":"Invitation","type":"Dépense","amount":21000,"account":"SIB","subcategory":"Femmes","note":"Vanessa"},{"id":"t1751","date":"2025-10-19","category":"Aliments","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1740","date":"2025-10-20","category":"Vêtements","type":"Dépense","amount":33030,"account":"Revenus MAZDA","subcategory":"Chaussures"},{"id":"t1741","date":"2025-10-20","category":"Aliments","type":"Dépense","amount":4500,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1742","date":"2025-10-20","category":"Invitation","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Femmes"},{"id":"t1743","date":"2025-10-20","category":"Vêtements","type":"Dépense","amount":42500,"account":"Revenus MAZDA","subcategory":"Chemises"},{"id":"t1744","date":"2025-10-20","category":"Shopping","type":"Dépense","amount":37000,"account":"Revenus MAZDA"},{"id":"t1745","date":"2025-10-20","category":"Cadeaux","type":"Dépense","amount":50000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Cadeau Vanessa"},{"id":"t1746","date":"2025-10-20","category":"Voiture","type":"Dépense","amount":51500,"account":"Revenus MAZDA","subcategory":"Assurance"},{"id":"t1735","date":"2025-10-21","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"Revenus MAZDA","subcategory":"PEL"},{"id":"t1736","date":"2025-10-21","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"Revenus MAZDA"},{"id":"t1737","date":"2025-10-21","category":"Enfants & Maman","type":"Dépense","amount":10100,"account":"Revenus MAZDA","subcategory":"Nesher"},{"id":"t1738","date":"2025-10-21","category":"Invitation","type":"Dépense","amount":36000,"account":"Revenus MAZDA"},{"id":"t1739","date":"2025-10-21","category":"Invitation","type":"Dépense","amount":15000,"account":"Revenus MAZDA"},{"id":"t1724","date":"2025-10-22","category":"Ajustement","type":"Dépense","amount":44564,"account":"Revenus MAZDA","subcategory":"Étrange"},{"id":"t1725","date":"2025-10-22","category":"Générales","type":"Dépense","amount":700,"account":"SIB","subcategory":"Yango Livraison","note":"Djeneba"},{"id":"t1726","date":"2025-10-22","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t1727","date":"2025-10-22","category":"Petty Cash","type":"Revenu","amount":2500000,"account":"PETTY CASH"},{"id":"t1728","date":"2025-10-22","category":"Cadeaux","type":"Dépense","amount":57000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Djeneba"},{"id":"t1729","date":"2025-10-22","category":"Personnel","type":"Dépense","amount":42000,"account":"Revenus MAZDA","subcategory":"Produits de beauté","note":"Bracelet"},{"id":"t1730","date":"2025-10-22","category":"Enfants & Maman","type":"Dépense","amount":222200,"account":"Revenus MAZDA","subcategory":"Maman","note":"Déménagement maman"},{"id":"t1731","date":"2025-10-22","category":"Aliments","type":"Dépense","amount":7000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1732","date":"2025-10-22","category":"Shopping","type":"Dépense","amount":1500,"account":"Revenus MAZDA"},{"id":"t1733","date":"2025-10-22","category":"Aliments","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1734","date":"2025-10-22","category":"Divertissement","type":"Dépense","amount":35000,"account":"Revenus MAZDA","subcategory":"La musique"},{"id":"t1718","date":"2025-10-23","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1719","date":"2025-10-23","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1720","date":"2025-10-23","category":"Shopping","type":"Dépense","amount":10758,"account":"Revenus MAZDA"},{"id":"t1721","date":"2025-10-23","category":"Santé","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Médicaments"},{"id":"t1722","date":"2025-10-23","category":"Cadeaux","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Vanessa"},{"id":"t1723","date":"2025-10-23","category":"Utilitaires","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Téléphones"},{"id":"t1711","date":"2025-10-24","category":"Shopping","type":"Dépense","amount":9200,"account":"Revenus MAZDA"},{"id":"t1712","date":"2025-10-24","category":"Shopping","type":"Dépense","amount":90300,"account":"PETTY CASH"},{"id":"t1713","date":"2025-10-24","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1714","date":"2025-10-24","category":"Utilitaires","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Téléphones"},{"id":"t1715","date":"2025-10-24","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1716","date":"2025-10-24","category":"Divertissement","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Favor"},{"id":"t1717","date":"2025-10-24","category":"Divertissement","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Alcool"},{"id":"t1707","date":"2025-10-26","category":"Cadeaux","type":"Dépense","amount":25250,"account":"Revenus MAZDA","subcategory":"Metty"},{"id":"t1708","date":"2025-10-26","category":"Ajustement","type":"Dépense","amount":2000,"account":"SIB"},{"id":"t1709","date":"2025-10-26","category":"Divertissement","type":"Dépense","amount":12000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1710","date":"2025-10-26","category":"Divertissement","type":"Dépense","amount":11000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1701","date":"2025-10-27","category":"Invitation","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Vanessa"},{"id":"t1702","date":"2025-10-27","category":"Cadeaux","type":"Dépense","amount":23000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Vanessa"},{"id":"t1703","date":"2025-10-27","category":"Aliments","type":"Dépense","amount":10500,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1704","date":"2025-10-27","category":"Cadeaux","type":"Dépense","amount":10575,"account":"Revenus MAZDA","subcategory":"Femme","note":"Mahora"},{"id":"t1705","date":"2025-10-27","category":"Ajustement","type":"Dépense","amount":1410,"account":"SIB"},{"id":"t1706","date":"2025-10-27","category":"Divertissement","type":"Dépense","amount":21000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1693","date":"2025-10-28","category":"Cadeaux","type":"Dépense","amount":170000,"account":"Revenus MAZDA","subcategory":"Metty","note":"Grossesse"},{"id":"t1694","date":"2025-10-28","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1695","date":"2025-10-28","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"Revenus MAZDA","subcategory":"Nesher"},{"id":"t1696","date":"2025-10-28","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t1697","date":"2025-10-28","category":"Ajustement","type":"Dépense","amount":2220,"account":"Revenus MAZDA","subcategory":"Frais Bancaire"},{"id":"t1698","date":"2025-10-28","category":"Invitation","type":"Dépense","amount":29000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Vanessa"},{"id":"t1699","date":"2025-10-28","category":"Aliments","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Dîner"},{"id":"t1700","date":"2025-10-28","category":"Divertissement","type":"Dépense","amount":1500,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1686","date":"2025-10-29","category":"Vente Pompe","type":"Revenu","amount":250000,"account":"SALAIRE"},{"id":"t1687","date":"2025-10-29","category":"Ajustement","type":"Revenu","amount":1755,"account":"SIB"},{"id":"t1688","date":"2025-10-29","category":"Vêtements","type":"Dépense","amount":19000,"account":"Revenus MAZDA","subcategory":"Chemises"},{"id":"t1689","date":"2025-10-29","category":"Cadeaux","type":"Dépense","amount":16000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Vanessa"},{"id":"t1690","date":"2025-10-29","category":"Enfants & Maman","type":"Dépense","amount":40200,"account":"Revenus MAZDA","subcategory":"Hemra"},{"id":"t1691","date":"2025-10-29","category":"Aliments","type":"Dépense","amount":6000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1692","date":"2025-10-29","category":"Invitation","type":"Dépense","amount":19000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Désirée Ismo"},{"id":"t1682","date":"2025-10-30","category":"Divertissement","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"BAP","note":"Favour"},{"id":"t1683","date":"2025-10-30","category":"Cadeaux","type":"Dépense","amount":11000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Mahy"},{"id":"t1684","date":"2025-10-30","category":"Divertissement","type":"Dépense","amount":62000,"account":"Revenus MAZDA","subcategory":"Alcool"},{"id":"t1685","date":"2025-10-30","category":"Aliments","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1674","date":"2025-10-31","category":"Voiture","type":"Dépense","amount":2500,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1675","date":"2025-10-31","category":"Ajustement","type":"Dépense","amount":5551,"account":"SIB"},{"id":"t1676","date":"2025-10-31","category":"Ajustement","type":"Revenu","amount":10000,"account":"SIB","note":"Don Joel"},{"id":"t1677","date":"2025-10-31","category":"Cadeaux","type":"Dépense","amount":10000,"account":"Revenus MAZDA","note":"Vanessa"},{"id":"t1678","date":"2025-10-31","category":"Divertissement","type":"Dépense","amount":11000,"account":"Revenus MAZDA","subcategory":"Alcool","note":"Vanessa"},{"id":"t1679","date":"2025-10-31","category":"Cadeaux","type":"Dépense","amount":2000,"account":"Revenus MAZDA","note":"Vanessa"},{"id":"t1680","date":"2025-10-31","category":"Aliments","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1681","date":"2025-10-31","category":"Invitation","type":"Dépense","amount":13500,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Vanessa"},{"id":"t1667","date":"2025-11-01","category":"Enfants & Maman","type":"Dépense","amount":12000,"account":"Revenus MAZDA","subcategory":"Hemra","note":"Jeux"},{"id":"t1668","date":"2025-11-01","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1669","date":"2025-11-01","category":"Voiture","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Lavage"},{"id":"t1670","date":"2025-11-01","category":"Invitation","type":"Dépense","amount":10500,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Grâce"},{"id":"t1671","date":"2025-11-01","category":"Invitation","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Femmes","note":"Djeneba"},{"id":"t1672","date":"2025-11-01","category":"Aliments","type":"Dépense","amount":9000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1673","date":"2025-11-01","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1658","date":"2025-11-02","category":"Invitation","type":"Dépense","amount":14000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Carole"},{"id":"t1659","date":"2025-11-02","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1660","date":"2025-11-02","category":"Générales","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Péage"},{"id":"t1661","date":"2025-11-02","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB"},{"id":"t1662","date":"2025-11-02","category":"Voiture","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1663","date":"2025-11-02","category":"Ajustement","type":"Dépense","amount":40000,"account":"Revenus MAZDA","note":"Accident"},{"id":"t1664","date":"2025-11-02","category":"Divertissement","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Anniversaire","note":"Toussaint"},{"id":"t1665","date":"2025-11-02","category":"Divertissement","type":"Dépense","amount":15000,"account":"Revenus MAZDA","subcategory":"Anniversaire","note":"Toussaint"},{"id":"t1666","date":"2025-11-02","category":"Générales","type":"Dépense","amount":5000,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t1649","date":"2025-11-03","category":"Cadeaux","type":"Dépense","amount":1000,"account":"Revenus MAZDA"},{"id":"t1650","date":"2025-11-03","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1651","date":"2025-11-03","category":"ECO PUMP","type":"Revenu","amount":585000,"account":"PUMP"},{"id":"t1652","date":"2025-11-03","category":"Cadeaux","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Vanessa"},{"id":"t1653","date":"2025-11-03","category":"Aliments","type":"Dépense","amount":6700,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1654","date":"2025-11-03","category":"Shopping","type":"Dépense","amount":3000,"account":"Revenus MAZDA"},{"id":"t1655","date":"2025-11-03","category":"Aliments","type":"Dépense","amount":700,"account":"Revenus MAZDA"},{"id":"t1656","date":"2025-11-03","category":"Shopping","type":"Dépense","amount":20000,"account":"Revenus MAZDA"},{"id":"t1657","date":"2025-11-03","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1648","date":"2025-11-04","category":"Santé","type":"Dépense","amount":2730,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t1646","date":"2025-11-05","category":"Abonnements","type":"Dépense","amount":1975,"account":"Revenus MAZDA","subcategory":"Spotify"},{"id":"t1647","date":"2025-11-05","category":"Aliments","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Déjeuner"},{"id":"t1645","date":"2025-11-06","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1640","date":"2025-11-07","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1641","date":"2025-11-07","category":"Invitation","type":"Dépense","amount":15000,"account":"SALAIRE","note":"Hommes"},{"id":"t1642","date":"2025-11-07","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1643","date":"2025-11-07","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1644","date":"2025-11-07","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1637","date":"2025-11-08","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1638","date":"2025-11-08","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1639","date":"2025-11-08","category":"Cadeaux","type":"Dépense","amount":10500,"account":"SALAIRE","subcategory":"Pourboire","note":"Malan"},{"id":"t1628","date":"2025-11-09","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1629","date":"2025-11-09","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1630","date":"2025-11-09","category":"Ajustement","type":"Dépense","amount":3391,"account":"SIB"},{"id":"t1631","date":"2025-11-09","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1632","date":"2025-11-09","category":"Invitation","type":"Dépense","amount":17000,"account":"SALAIRE","note":"OLO"},{"id":"t1633","date":"2025-11-09","category":"Divertissement","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1634","date":"2025-11-09","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Olo"},{"id":"t1635","date":"2025-11-09","category":"Vêtements","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Accessoires"},{"id":"t1636","date":"2025-11-09","category":"Divertissement","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1625","date":"2025-11-10","category":"Générales","type":"Dépense","amount":35000,"account":"SALAIRE"},{"id":"t1626","date":"2025-11-10","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1627","date":"2025-11-10","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t1619","date":"2025-11-11","category":"Vêtements","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Un pantalon"},{"id":"t1620","date":"2025-11-11","category":"GRUNDFOS","type":"Dépense","amount":4000,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango"},{"id":"t1621","date":"2025-11-11","category":"Générales","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Puce GHANA"},{"id":"t1622","date":"2025-11-11","category":"GRUNDFOS","type":"Dépense","amount":50000,"account":"PETTY CASH","subcategory":"Divertissement"},{"id":"t1623","date":"2025-11-11","category":"Divertissement","type":"Dépense","amount":60000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Awa Salifou"},{"id":"t1624","date":"2025-11-11","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Voyage","note":"Déjeuner"},{"id":"t1618","date":"2025-11-12","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Divertissement"},{"id":"t1616","date":"2025-11-13","category":"Ajustement","type":"Dépense","amount":16555,"account":"SIB"},{"id":"t1617","date":"2025-11-13","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA","note":"November"},{"id":"t1610","date":"2025-11-15","category":"Divertissement","type":"Dépense","amount":45000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Richie"},{"id":"t1611","date":"2025-11-15","category":"Cadeaux","type":"Dépense","amount":30000,"account":"Revenus MAZDA","note":"Bb Roland"},{"id":"t1612","date":"2025-11-15","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango airport"},{"id":"t1613","date":"2025-11-15","category":"Cadeaux","type":"Dépense","amount":13000,"account":"SALAIRE","note":"Chaussures Roland"},{"id":"t1614","date":"2025-11-15","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Voyage","note":"Déjeuner"},{"id":"t1615","date":"2025-11-15","category":"Shopping","type":"Dépense","amount":29000,"account":"Revenus MAZDA","note":"Kenya"},{"id":"t1598","date":"2025-11-16","category":"Loyer","type":"Revenu","amount":395020,"account":"SIB","note":"Caution"},{"id":"t1599","date":"2025-11-16","category":"Shopping","type":"Dépense","amount":158000,"account":"Revenus MAZDA","note":"Achat duty free"},{"id":"t1600","date":"2025-11-16","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1601","date":"2025-11-16","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Pourboire","note":"Virgile Bernabe"},{"id":"t1602","date":"2025-11-16","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango"},{"id":"t1603","date":"2025-11-16","category":"Générales","type":"Dépense","amount":2100,"account":"SALAIRE","subcategory":"Yango Livraison","note":"Livraison clé"},{"id":"t1604","date":"2025-11-16","category":"Cadeaux","type":"Dépense","amount":10000,"account":"Revenus MAZDA","note":"Alvy"},{"id":"t1605","date":"2025-11-16","category":"Cadeaux","type":"Dépense","amount":20000,"account":"Revenus MAZDA","note":"Maman Alvy"},{"id":"t1606","date":"2025-11-16","category":"Divertissement","type":"Dépense","amount":40000,"account":"Revenus MAZDA","subcategory":"Alcool","note":"Vin Toussaint Alvy"},{"id":"t1607","date":"2025-11-16","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Femme","note":"Yango Alvy"},{"id":"t1608","date":"2025-11-16","category":"Santé","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Médicaments","note":"Para"},{"id":"t1609","date":"2025-11-16","category":"Invitation","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Alvy"},{"id":"t1591","date":"2025-11-17","category":"Voiture","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Pneu"},{"id":"t1592","date":"2025-11-17","category":"Enfants & Maman","type":"Dépense","amount":20200,"account":"SIB","subcategory":"Nesher","note":"Education"},{"id":"t1593","date":"2025-11-17","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1594","date":"2025-11-17","category":"Abonnements","type":"Dépense","amount":5336,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t1595","date":"2025-11-17","category":"Ajustement","type":"Dépense","amount":1600,"account":"SIB","subcategory":"Frais"},{"id":"t1596","date":"2025-11-17","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","subcategory":"Femme","note":"Emmanuella"},{"id":"t1597","date":"2025-11-17","category":"Aliments","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Dîner"},{"id":"t1581","date":"2025-11-18","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SIB","subcategory":"Femme","note":"Alvy"},{"id":"t1582","date":"2025-11-18","category":"Divertissement","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Alcool"},{"id":"t1583","date":"2025-11-18","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1584","date":"2025-11-18","category":"Générales","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t1585","date":"2025-11-18","category":"Shopping","type":"Dépense","amount":10500,"account":"SIB"},{"id":"t1586","date":"2025-11-18","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1587","date":"2025-11-18","category":"Générales","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Police"},{"id":"t1588","date":"2025-11-18","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1589","date":"2025-11-18","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1590","date":"2025-11-18","category":"Divertissement","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Alcool"},{"id":"t1575","date":"2025-11-19","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t1576","date":"2025-11-19","category":"Shopping","type":"Dépense","amount":2100,"account":"SIB"},{"id":"t1577","date":"2025-11-19","category":"Divertissement","type":"Dépense","amount":11000,"account":"SIB","subcategory":"Alcool"},{"id":"t1578","date":"2025-11-19","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1579","date":"2025-11-19","category":"Divertissement","type":"Dépense","amount":25000,"account":"SIB","subcategory":"BAP"},{"id":"t1580","date":"2025-11-19","category":"Divertissement","type":"Dépense","amount":1500,"account":"SIB","subcategory":"Alcool"},{"id":"t1573","date":"2025-11-20","category":"GRUNDFOS","type":"Dépense","amount":617500,"account":"PETTY CASH","subcategory":"iPhone 16 Pro"},{"id":"t1574","date":"2025-11-20","category":"Aliments","type":"Dépense","amount":4240,"account":"SIB","subcategory":"Déjeuner"},{"id":"t3033","date":"2025-11-20","category":"Générales","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Réparation iPhone 13"},{"id":"t1571","date":"2025-11-21","category":"Divertissement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Alcool"},{"id":"t1572","date":"2025-11-21","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1561","date":"2025-11-22","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1562","date":"2025-11-22","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme"},{"id":"t1563","date":"2025-11-22","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1564","date":"2025-11-22","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1565","date":"2025-11-22","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Ela"},{"id":"t1566","date":"2025-11-22","category":"Shopping","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Alimentation","note":"Jus Paula"},{"id":"t1567","date":"2025-11-22","category":"Shopping","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alimentation"},{"id":"t1568","date":"2025-11-22","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Ndjore"},{"id":"t1569","date":"2025-11-22","category":"Invitation","type":"Dépense","amount":19000,"account":"SIB","subcategory":"Femmes","note":"Ela"},{"id":"t1570","date":"2025-11-22","category":"Voyage","type":"Dépense","amount":65000,"account":"SIB","note":"Bassam"},{"id":"t1553","date":"2025-11-23","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Femme","note":"Sarah"},{"id":"t1554","date":"2025-11-23","category":"Divertissement","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"BAP"},{"id":"t1555","date":"2025-11-23","category":"Cadeaux","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Pourboire","note":"485 Muzik"},{"id":"t1556","date":"2025-11-23","category":"Divertissement","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Alcool","note":"Rumba"},{"id":"t1557","date":"2025-11-23","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Dîner"},{"id":"t1558","date":"2025-11-23","category":"Générales","type":"Dépense","amount":10000,"account":"SIB","note":"Œuf accident"},{"id":"t1559","date":"2025-11-23","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Femme","note":"Mahy"},{"id":"t1560","date":"2025-11-23","category":"Divertissement","type":"Dépense","amount":27000,"account":"SIB","subcategory":"Alcool"},{"id":"t1544","date":"2025-11-24","category":"Securicompte","type":"Dépense","amount":109407,"account":"Revenus MAZDA"},{"id":"t1545","date":"2025-11-24","category":"Vente Pompe","type":"Revenu","amount":59500,"account":"SIB"},{"id":"t1546","date":"2025-11-24","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t1547","date":"2025-11-24","category":"Aliments","type":"Dépense","amount":8000,"account":"Revenus MAZDA","subcategory":"Déjeuner","note":"Hammal"},{"id":"t1548","date":"2025-11-24","category":"Cadeaux","type":"Dépense","amount":21500,"account":"Revenus MAZDA","note":"Cadeau Delali"},{"id":"t1549","date":"2025-11-24","category":"Vêtements","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Chemises"},{"id":"t1550","date":"2025-11-24","category":"Cadeaux","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Femme","note":"Kerene"},{"id":"t1551","date":"2025-11-24","category":"Vêtements","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Chemises"},{"id":"t1552","date":"2025-11-24","category":"Invitation","type":"Dépense","amount":36000,"account":"SIB","subcategory":"Femmes","note":"Kerene"},{"id":"t1541","date":"2025-11-25","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t1542","date":"2025-11-25","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"Revenus MAZDA"},{"id":"t1543","date":"2025-11-25","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"Revenus MAZDA","subcategory":"PEL"},{"id":"t1531","date":"2025-11-26","category":"GRUNDFOS","type":"Dépense","amount":80000,"account":"PETTY CASH","subcategory":"Voyage","note":"Team building"},{"id":"t1532","date":"2025-11-26","category":"Ajustement","type":"Revenu","amount":25002,"account":"SIB"},{"id":"t1533","date":"2025-11-26","category":"Securicompte","type":"Dépense","amount":912,"account":"SIB"},{"id":"t1534","date":"2025-11-26","category":"Abonnements","type":"Dépense","amount":1191,"account":"SIB","subcategory":"Google Espace"},{"id":"t1535","date":"2025-11-26","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"Voyage","note":"Pourboire"},{"id":"t1536","date":"2025-11-26","category":"GRUNDFOS","type":"Dépense","amount":8000,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango"},{"id":"t1537","date":"2025-11-26","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","note":"Tantie Esther"},{"id":"t1538","date":"2025-11-26","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","note":"Parker"},{"id":"t1539","date":"2025-11-26","category":"Aliments","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1540","date":"2025-11-26","category":"Voiture","type":"Dépense","amount":240000,"account":"Revenus MAZDA","subcategory":"Peinture-Retouche","note":"Inclus accident JP"},{"id":"t1527","date":"2025-11-27","category":"GRUNDFOS","type":"Dépense","amount":850,"account":"PETTY CASH","subcategory":"Voyage","note":"Divertissement Bella"},{"id":"t1528","date":"2025-11-27","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Voyage","note":"Taxi airport Accra"},{"id":"t1529","date":"2025-11-27","category":"Ajustement","type":"Dépense","amount":2334,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1530","date":"2025-11-27","category":"GRUNDFOS","type":"Dépense","amount":36000,"account":"PETTY CASH","subcategory":"Voyage","payee":"Bella","note":"Dîner Accra invitation"},{"id":"t1524","date":"2025-11-28","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Voyage","note":"Medicament Nzulezo"},{"id":"t1525","date":"2025-11-28","category":"GRUNDFOS","type":"Dépense","amount":4000,"account":"PETTY CASH","subcategory":"Voyage","note":"Taxis dans Accra"},{"id":"t1526","date":"2025-11-28","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Voyage","note":"Taxi dans Accra"},{"id":"t1522","date":"2025-11-29","category":"GRUNDFOS","type":"Dépense","amount":53000,"account":"PETTY CASH","subcategory":"Voyage","note":"Entertainment Diane"},{"id":"t1523","date":"2025-11-29","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Voyage","note":"Dîner avec Nigeria"},{"id":"t1507","date":"2025-11-30","category":"Ajustement","type":"Revenu","amount":32649,"account":"SIB"},{"id":"t1508","date":"2025-11-30","category":"GRUNDFOS","type":"Dépense","amount":4000,"account":"PETTY CASH","subcategory":"Voyage","note":"Taxi airport kotoka"},{"id":"t1509","date":"2025-11-30","category":"Cadeaux","type":"Dépense","amount":60600,"account":"Revenus MAZDA","subcategory":"Mardochee","note":"Inscription"},{"id":"t1510","date":"2025-11-30","category":"Enfants & Maman","type":"Dépense","amount":60600,"account":"Revenus MAZDA","subcategory":"Nesher"},{"id":"t1511","date":"2025-11-30","category":"GRUNDFOS","type":"Dépense","amount":8075,"account":"PETTY CASH","subcategory":"Voyage","note":"Déjeuner airport Kotoka"},{"id":"t1512","date":"2025-11-30","category":"GRUNDFOS","type":"Dépense","amount":25700,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1513","date":"2025-11-30","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Phone"},{"id":"t1514","date":"2025-11-30","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Voyage","payee":"TB","note":"Yango Aeroport GGH"},{"id":"t1515","date":"2025-11-30","category":"Logement","type":"Dépense","amount":380000,"account":"Revenus MAZDA","subcategory":"Location","note":"Décembre"},{"id":"t1516","date":"2025-11-30","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"Revenus MAZDA","subcategory":"Maman"},{"id":"t1517","date":"2025-11-30","category":"Cadeaux","type":"Dépense","amount":151500,"account":"Revenus MAZDA","note":"Accident Jean Philippe"},{"id":"t1518","date":"2025-11-30","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t1519","date":"2025-11-30","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Pourboire","note":"Paulo"},{"id":"t1520","date":"2025-11-30","category":"Divertissement","type":"Dépense","amount":35000,"account":"SIB","subcategory":"La musique"},{"id":"t1521","date":"2025-11-30","category":"Aliments","type":"Dépense","amount":10000,"account":"SIB","subcategory":"Dîner"},{"id":"t1501","date":"2025-12-01","category":"Revenu général","type":"Revenu","amount":2500000,"account":"Dépôt LOYER","note":"Petty cash"},{"id":"t1502","date":"2025-12-01","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1503","date":"2025-12-01","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1504","date":"2025-12-01","category":"Shopping","type":"Dépense","amount":24500,"account":"SALAIRE"},{"id":"t1505","date":"2025-12-01","category":"Santé","type":"Dépense","amount":11110,"account":"SALAIRE","subcategory":"VG"},{"id":"t1506","date":"2025-12-01","category":"Loyer","type":"Revenu","amount":1500000,"account":"Dépôt LOYER","note":"Décembre janvier Février"},{"id":"t1497","date":"2025-12-02","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1498","date":"2025-12-02","category":"Santé","type":"Dépense","amount":11500,"account":"SALAIRE","subcategory":"Médicaments","note":"Ronfle"},{"id":"t1499","date":"2025-12-02","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Alcool","note":"Olo"},{"id":"t1500","date":"2025-12-02","category":"Enfants & Maman","type":"Dépense","amount":37000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t1489","date":"2025-12-03","category":"Ajustement","type":"Dépense","amount":8603,"account":"SALAIRE"},{"id":"t1490","date":"2025-12-03","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire","note":"Hassan"},{"id":"t1491","date":"2025-12-03","category":"Voiture","type":"Dépense","amount":17000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Essuie glace"},{"id":"t1492","date":"2025-12-03","category":"Générales","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Police","note":"Radar"},{"id":"t1493","date":"2025-12-03","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"BAP","note":"Favour"},{"id":"t1494","date":"2025-12-03","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Alcool","note":"Ismo"},{"id":"t1495","date":"2025-12-03","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Tonton Jacques"},{"id":"t1496","date":"2025-12-03","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1486","date":"2025-12-04","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1487","date":"2025-12-04","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1488","date":"2025-12-04","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1480","date":"2025-12-05","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1481","date":"2025-12-05","category":"Aliments","type":"Dépense","amount":23500,"account":"SALAIRE","note":"Leati"},{"id":"t1482","date":"2025-12-05","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1483","date":"2025-12-05","category":"Shopping","type":"Dépense","amount":1500,"account":"SALAIRE"},{"id":"t1484","date":"2025-12-05","category":"Personnel","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Produits de beauté","note":"Chaîne"},{"id":"t1485","date":"2025-12-05","category":"Invitation","type":"Dépense","amount":63000,"account":"SALAIRE","subcategory":"Femmes","note":"Paola"},{"id":"t1470","date":"2025-12-06","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SALAIRE"},{"id":"t1471","date":"2025-12-06","category":"Ajustement","type":"Dépense","amount":295,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1472","date":"2025-12-06","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1473","date":"2025-12-06","category":"Ajustement","type":"Dépense","amount":295,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1474","date":"2025-12-06","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1475","date":"2025-12-06","category":"Voyage","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Pourboire","note":"Vigile lycée"},{"id":"t1476","date":"2025-12-06","category":"Voyage","type":"Dépense","amount":17000,"account":"SALAIRE","subcategory":"Pourboire","note":"Medine"},{"id":"t1477","date":"2025-12-06","category":"Voyage","type":"Dépense","amount":62000,"account":"SALAIRE","subcategory":"Divertissement","note":"Henessy Mano"},{"id":"t1478","date":"2025-12-06","category":"Voyage","type":"Dépense","amount":2020,"account":"SALAIRE","subcategory":"Divertissement"},{"id":"t1479","date":"2025-12-06","category":"Voyage","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Divertissement","note":"Ameyao"},{"id":"t1459","date":"2025-12-07","category":"Shopping","type":"Dépense","amount":8404,"account":"SALAIRE"},{"id":"t1460","date":"2025-12-07","category":"Santé","type":"Dépense","amount":9000,"account":"SALAIRE","subcategory":"Médicaments","note":"Franko"},{"id":"t1461","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":26260,"account":"SALAIRE","subcategory":"Divertissement"},{"id":"t1462","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":26000,"account":"SALAIRE","subcategory":"Un hôtel","note":"Royaume"},{"id":"t1463","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Pourboire","note":"Parker"},{"id":"t1464","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Pourboire","note":"DJ"},{"id":"t1465","date":"2025-12-07","category":"Enfants & Maman","type":"Dépense","amount":40000,"account":"Dépôt LOYER","subcategory":"Maman","note":"Loyer maman"},{"id":"t1466","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Aliment","note":"Dîner"},{"id":"t1467","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Un hôtel","note":"Residence"},{"id":"t1468","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Aliment"},{"id":"t1469","date":"2025-12-07","category":"Voyage","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Pourboire","note":"Fille Mano"},{"id":"t1451","date":"2025-12-08","category":"Voyage","type":"Dépense","amount":1000,"account":"Revenus MAZDA"},{"id":"t1452","date":"2025-12-08","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1453","date":"2025-12-08","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1454","date":"2025-12-08","category":"Voyage","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Pourboire","note":"Élève"},{"id":"t1455","date":"2025-12-08","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1456","date":"2025-12-08","category":"Voyage","type":"Dépense","amount":3030,"account":"SALAIRE","subcategory":"Pourboire","note":"Juliana"},{"id":"t1457","date":"2025-12-08","category":"Voyage","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Pourboire","note":"Aude"},{"id":"t1458","date":"2025-12-08","category":"Voyage","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Pourboire","note":"Juliana"},{"id":"t1444","date":"2025-12-09","category":"Ajustement","type":"Dépense","amount":499,"account":"SIB","subcategory":"Frais"},{"id":"t1445","date":"2025-12-09","category":"Ajustement","type":"Dépense","amount":1877,"account":"SIB","subcategory":"Frais"},{"id":"t1446","date":"2025-12-09","category":"GRUNDFOS","type":"Dépense","amount":300,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1447","date":"2025-12-09","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA","note":"Janvier"},{"id":"t1448","date":"2025-12-09","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1449","date":"2025-12-09","category":"Payement Maison Bingerville","type":"Dépense","amount":419000,"account":"Dépôt LOYER","note":"Finition appartement"},{"id":"t1450","date":"2025-12-09","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1437","date":"2025-12-10","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1438","date":"2025-12-10","category":"Cadeaux","type":"Dépense","amount":65000,"account":"SALAIRE","subcategory":"Ndjore"},{"id":"t1439","date":"2025-12-10","category":"Voiture","type":"Dépense","amount":63000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Vidange 71222"},{"id":"t1440","date":"2025-12-10","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1441","date":"2025-12-10","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1442","date":"2025-12-10","category":"GRUNDFOS","type":"Dépense","amount":5185,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1443","date":"2025-12-10","category":"Ajustement","type":"Dépense","amount":2955,"account":"SALAIRE","subcategory":"Frais"},{"id":"t1434","date":"2025-12-11","category":"Aliments","type":"Dépense","amount":700,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1435","date":"2025-12-11","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1436","date":"2025-12-11","category":"Divertissement","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1431","date":"2025-12-12","category":"Personnel","type":"Dépense","amount":1515,"account":"Revenus MAZDA","subcategory":"Produits de beauté"},{"id":"t1432","date":"2025-12-12","category":"Invitation","type":"Dépense","amount":38000,"account":"SALAIRE","note":"Anniversaire Faten"},{"id":"t1433","date":"2025-12-12","category":"Invitation","type":"Dépense","amount":19000,"account":"SALAIRE","note":"Elvis dandy"},{"id":"t1427","date":"2025-12-13","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1428","date":"2025-12-13","category":"Divertissement","type":"Dépense","amount":100000,"account":"SALAIRE","subcategory":"La musique","note":"Alvy"},{"id":"t1429","date":"2025-12-13","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1430","date":"2025-12-13","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t3032","date":"2025-12-13","category":"Personnel","type":"Dépense","amount":320000,"account":"Revenus MAZDA","subcategory":"Hygiène personnelle"},{"id":"t1425","date":"2025-12-14","category":"Aliments","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1426","date":"2025-12-14","category":"Divertissement","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1422","date":"2025-12-15","category":"Shopping","type":"Dépense","amount":10000,"account":"SALAIRE"},{"id":"t1423","date":"2025-12-15","category":"Divertissement","type":"Dépense","amount":2515,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1424","date":"2025-12-15","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1418","date":"2025-12-16","category":"GRUNDFOS","type":"Dépense","amount":5200,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1419","date":"2025-12-16","category":"Aliments","type":"Dépense","amount":9500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1420","date":"2025-12-16","category":"Ajustement","type":"Dépense","amount":18234,"account":"Revenus MAZDA"},{"id":"t1421","date":"2025-12-16","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1414","date":"2025-12-17","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1415","date":"2025-12-17","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1416","date":"2025-12-17","category":"Aliments","type":"Dépense","amount":3900,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1417","date":"2025-12-17","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1407","date":"2025-12-18","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","note":"Jean Marc"},{"id":"t1408","date":"2025-12-18","category":"Personnel","type":"Dépense","amount":39000,"account":"SALAIRE","subcategory":"Produits de beauté","note":"Chaîne et Bracelet"},{"id":"t1409","date":"2025-12-18","category":"Divertissement","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1410","date":"2025-12-18","category":"Aliments","type":"Dépense","amount":7500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1411","date":"2025-12-18","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","note":"Augustin"},{"id":"t1412","date":"2025-12-18","category":"Un salaire","type":"Revenu","amount":2636156,"account":"SALAIRE"},{"id":"t1413","date":"2025-12-18","category":"Invitation","type":"Dépense","amount":18000,"account":"SALAIRE","subcategory":"Femmes","note":"Ange"},{"id":"t1402","date":"2025-12-19","category":"Enfants & Maman","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t1403","date":"2025-12-19","category":"Aliments","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1404","date":"2025-12-19","category":"Divertissement","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"BAP","note":"Moon"},{"id":"t1405","date":"2025-12-19","category":"Vêtements","type":"Dépense","amount":22500,"account":"SALAIRE","subcategory":"Chaussures"},{"id":"t1406","date":"2025-12-19","category":"Shopping","type":"Dépense","amount":15000,"account":"SALAIRE"},{"id":"t1391","date":"2025-12-20","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1392","date":"2025-12-20","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool","note":"REX"},{"id":"t1393","date":"2025-12-20","category":"Divertissement","type":"Dépense","amount":31000,"account":"SALAIRE","note":"Richard résidence"},{"id":"t1394","date":"2025-12-20","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1395","date":"2025-12-20","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1396","date":"2025-12-20","category":"Divertissement","type":"Dépense","amount":30500,"account":"SALAIRE","subcategory":"Anniversaire"},{"id":"t1397","date":"2025-12-20","category":"Invitation","type":"Dépense","amount":18000,"account":"SALAIRE","note":"Famille"},{"id":"t1398","date":"2025-12-20","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Tonton Kouassi"},{"id":"t1399","date":"2025-12-20","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1400","date":"2025-12-20","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1401","date":"2025-12-20","category":"Cadeaux","type":"Dépense","amount":45000,"account":"SALAIRE","note":"Patrice"},{"id":"t1384","date":"2025-12-21","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1385","date":"2025-12-21","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1386","date":"2025-12-21","category":"Divertissement","type":"Dépense","amount":33000,"account":"SALAIRE","subcategory":"Alcool","note":"Ismo"},{"id":"t1387","date":"2025-12-21","category":"Cadeaux","type":"Dépense","amount":25000,"account":"SALAIRE","note":"Femme Ismo"},{"id":"t1388","date":"2025-12-21","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","note":"Patrice & Richard"},{"id":"t1389","date":"2025-12-21","category":"Aliments","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1390","date":"2025-12-21","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1373","date":"2025-12-22","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1374","date":"2025-12-22","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1375","date":"2025-12-22","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"Medine"},{"id":"t1376","date":"2025-12-22","category":"Shopping","type":"Dépense","amount":33000,"account":"SALAIRE"},{"id":"t1377","date":"2025-12-22","category":"Ajustement","type":"Dépense","amount":5362,"account":"SALAIRE"},{"id":"t1378","date":"2025-12-22","category":"Générales","type":"Dépense","amount":500,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1379","date":"2025-12-22","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SALAIRE"},{"id":"t1380","date":"2025-12-22","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SALAIRE"},{"id":"t1381","date":"2025-12-22","category":"Abonnements","type":"Dépense","amount":5300,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t1382","date":"2025-12-22","category":"Enfants & Maman","type":"Dépense","amount":70700,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t1383","date":"2025-12-22","category":"Divertissement","type":"Dépense","amount":22000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1359","date":"2025-12-23","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","note":"Vigile"},{"id":"t1360","date":"2025-12-23","category":"Abonnements","type":"Dépense","amount":5885,"account":"PETTY CASH","subcategory":"Gamma App"},{"id":"t1361","date":"2025-12-23","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","note":"Juliana"},{"id":"t1362","date":"2025-12-23","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","note":"Armande"},{"id":"t1363","date":"2025-12-23","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1364","date":"2025-12-23","category":"Cadeaux","type":"Dépense","amount":500000,"account":"SALAIRE","subcategory":"Adrien"},{"id":"t1365","date":"2025-12-23","category":"Logement","type":"Dépense","amount":380000,"account":"Dépôt LOYER","subcategory":"Location"},{"id":"t1366","date":"2025-12-23","category":"Générales","type":"Dépense","amount":2200,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t1367","date":"2025-12-23","category":"Ajustement","type":"Dépense","amount":2200,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t1368","date":"2025-12-23","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE"},{"id":"t1369","date":"2025-12-23","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1370","date":"2025-12-23","category":"Ajustement","type":"Revenu","amount":80,"account":"SIB"},{"id":"t1371","date":"2025-12-23","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1372","date":"2025-12-23","category":"Aliments","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1350","date":"2025-12-24","category":"Divertissement","type":"Dépense","amount":9500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1351","date":"2025-12-24","category":"Abonnements","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Canal"},{"id":"t1352","date":"2025-12-24","category":"Cadeaux","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Femme","note":"Taxi Medine"},{"id":"t1353","date":"2025-12-24","category":"Aliments","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1354","date":"2025-12-24","category":"Shopping","type":"Dépense","amount":50000,"account":"SALAIRE","note":"Matelas"},{"id":"t1355","date":"2025-12-24","category":"Shopping","type":"Dépense","amount":46000,"account":"SALAIRE","note":"Cabri"},{"id":"t1356","date":"2025-12-24","category":"Invitation","type":"Dépense","amount":32000,"account":"SALAIRE","note":"Triade"},{"id":"t1357","date":"2025-12-24","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Francko"},{"id":"t1358","date":"2025-12-24","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","note":"Tantie Esther"},{"id":"t1341","date":"2025-12-25","category":"Ajustement","type":"Dépense","amount":8731,"account":"SALAIRE"},{"id":"t1342","date":"2025-12-25","category":"Divertissement","type":"Dépense","amount":28500,"account":"SALAIRE","subcategory":"Alcool","note":"Baileys"},{"id":"t1343","date":"2025-12-25","category":"Cadeaux","type":"Dépense","amount":4020,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1344","date":"2025-12-25","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Femme","note":"Ruth"},{"id":"t1345","date":"2025-12-25","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1346","date":"2025-12-25","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Femme","note":"Charlene"},{"id":"t1347","date":"2025-12-25","category":"Shopping","type":"Dépense","amount":20000,"account":"SALAIRE"},{"id":"t1348","date":"2025-12-25","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1349","date":"2025-12-25","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1335","date":"2025-12-26","category":"Voyage","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Péage"},{"id":"t1336","date":"2025-12-26","category":"Voyage","type":"Dépense","amount":39000,"account":"SALAIRE","subcategory":"Shopping"},{"id":"t1337","date":"2025-12-26","category":"Voyage","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Aliment"},{"id":"t1338","date":"2025-12-26","category":"Securicompte","type":"Dépense","amount":912,"account":"SALAIRE"},{"id":"t1339","date":"2025-12-26","category":"Pack Club","type":"Dépense","amount":9087,"account":"SALAIRE"},{"id":"t1340","date":"2025-12-26","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Femme","note":"Medine"},{"id":"t1329","date":"2025-12-27","category":"Voyage","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Pourboire","note":"Autres pourboires"},{"id":"t1330","date":"2025-12-27","category":"Ajustement","type":"Dépense","amount":2679,"account":"SALAIRE"},{"id":"t1331","date":"2025-12-27","category":"Voyage","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Pourboire","note":"Tantes"},{"id":"t1332","date":"2025-12-27","category":"Voyage","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Pourboire","note":"Saturnin"},{"id":"t1333","date":"2025-12-27","category":"Voyage","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Pourboire","note":"Carolle et enfants"},{"id":"t1334","date":"2025-12-27","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Cotisations","note":"Nziyè"},{"id":"t1323","date":"2025-12-28","category":"Aliments","type":"Dépense","amount":505,"account":"SIB"},{"id":"t1324","date":"2025-12-28","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1325","date":"2025-12-28","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1326","date":"2025-12-28","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1327","date":"2025-12-28","category":"Shopping","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Alimentation","note":"Poissons"},{"id":"t1328","date":"2025-12-28","category":"Divertissement","type":"Dépense","amount":61500,"account":"SALAIRE","subcategory":"Alcool","note":"Match"},{"id":"t1317","date":"2025-12-29","category":"Divertissement","type":"Dépense","amount":19020,"account":"SALAIRE","subcategory":"Femme","note":"Ange mahou"},{"id":"t1318","date":"2025-12-29","category":"Générales","type":"Dépense","amount":2700,"account":"SIB","subcategory":"Livraison Pompe"},{"id":"t1319","date":"2025-12-29","category":"Vente Pompe","type":"Revenu","amount":200000,"account":"SIB"},{"id":"t1320","date":"2025-12-29","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1321","date":"2025-12-29","category":"Divertissement","type":"Dépense","amount":1300,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1322","date":"2025-12-29","category":"Aliments","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1306","date":"2025-12-30","category":"Ajustement","type":"Dépense","amount":9418,"account":"SIB"},{"id":"t1307","date":"2025-12-30","category":"GRUNDFOS","type":"Dépense","amount":340760,"account":"PETTY CASH","subcategory":"FedEx"},{"id":"t1308","date":"2025-12-30","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1309","date":"2025-12-30","category":"Santé","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t1310","date":"2025-12-30","category":"Aliments","type":"Dépense","amount":8500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1311","date":"2025-12-30","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Parker"},{"id":"t1312","date":"2025-12-30","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","note":"Tonton Martin"},{"id":"t1313","date":"2025-12-30","category":"Enfants & Maman","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Maman"},{"id":"t1314","date":"2025-12-30","category":"Divertissement","type":"Dépense","amount":9000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1315","date":"2025-12-30","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1316","date":"2025-12-30","category":"Invitation","type":"Dépense","amount":33000,"account":"SALAIRE","subcategory":"Femmes","note":"Nella"},{"id":"t1299","date":"2025-12-31","category":"Enfants & Maman","type":"Dépense","amount":57000,"account":"SIB","subcategory":"Hemra","note":"Jeu Hemra"},{"id":"t1300","date":"2025-12-31","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1301","date":"2025-12-31","category":"Cadeaux","type":"Dépense","amount":28000,"account":"SALAIRE","subcategory":"Pourboire","note":"Achat cadeaux"},{"id":"t1302","date":"2025-12-31","category":"Cadeaux","type":"Dépense","amount":167000,"account":"SALAIRE","subcategory":"Femme","note":"Ruth"},{"id":"t1303","date":"2025-12-31","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SALAIRE","subcategory":"Femme","note":"Kessy"},{"id":"t1304","date":"2025-12-31","category":"Enfants & Maman","type":"Dépense","amount":60600,"account":"SIB","subcategory":"Nesher"},{"id":"t1305","date":"2025-12-31","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1295","date":"2026-01-01","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1296","date":"2026-01-01","category":"Invitation","type":"Dépense","amount":67000,"account":"SIB","subcategory":"Femmes"},{"id":"t1297","date":"2026-01-01","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1298","date":"2026-01-01","category":"Petty Cash","type":"Revenu","amount":2500000,"account":"PETTY CASH"},{"id":"t1288","date":"2026-01-02","category":"Ajustement","type":"Dépense","amount":6038,"account":"SALAIRE"},{"id":"t1289","date":"2026-01-02","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire","note":"Vigile"},{"id":"t1290","date":"2026-01-02","category":"Shopping","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Jus"},{"id":"t1291","date":"2026-01-02","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Déjeuner","note":"Olo"},{"id":"t1292","date":"2026-01-02","category":"Cadeaux","type":"Dépense","amount":10500,"account":"SALAIRE","subcategory":"Femme","note":"Poulet Ruth"},{"id":"t1293","date":"2026-01-02","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1294","date":"2026-01-02","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","note":"Esther & Mamy"},{"id":"t1287","date":"2026-01-03","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Pourboire","note":"Ladji"},{"id":"t1283","date":"2026-01-04","category":"Shopping","type":"Dépense","amount":5000,"account":"SALAIRE"},{"id":"t1284","date":"2026-01-04","category":"Santé","type":"Dépense","amount":81000,"account":"SALAIRE","subcategory":"Hôpital","note":"Mobidoc + médoc 65K+16K"},{"id":"t1285","date":"2026-01-04","category":"Shopping","type":"Dépense","amount":71200,"account":"SALAIRE"},{"id":"t1286","date":"2026-01-04","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t1281","date":"2026-01-05","category":"Shopping","type":"Dépense","amount":2500,"account":"SALAIRE"},{"id":"t1282","date":"2026-01-05","category":"Cadeaux","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Ndjore"},{"id":"t1276","date":"2026-01-06","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Tonton Roger"},{"id":"t1277","date":"2026-01-06","category":"Utilitaires","type":"Dépense","amount":22000,"account":"SALAIRE","subcategory":"Nettoyage"},{"id":"t1278","date":"2026-01-06","category":"Shopping","type":"Dépense","amount":134000,"account":"SALAIRE"},{"id":"t1279","date":"2026-01-06","category":"Ajustement","type":"Dépense","amount":111,"account":"SALAIRE"},{"id":"t1280","date":"2026-01-06","category":"Shopping","type":"Dépense","amount":894,"account":"SALAIRE"},{"id":"t1271","date":"2026-01-07","category":"Aliments","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1272","date":"2026-01-07","category":"Shopping","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Pumpkin Nespresso"},{"id":"t1273","date":"2026-01-07","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1274","date":"2026-01-07","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1275","date":"2026-01-07","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1267","date":"2026-01-08","category":"GRUNDFOS","type":"Dépense","amount":5200,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1268","date":"2026-01-08","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1269","date":"2026-01-08","category":"Ajustement","type":"Dépense","amount":2874,"account":"SALAIRE"},{"id":"t1270","date":"2026-01-08","category":"Aliments","type":"Dépense","amount":600,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1259","date":"2026-01-09","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1260","date":"2026-01-09","category":"Vente Pompe","type":"Revenu","amount":1137500,"account":"SIB","note":"CR10-16 Badiel"},{"id":"t1261","date":"2026-01-09","category":"Ajustement","type":"Dépense","amount":6208,"account":"SALAIRE"},{"id":"t1262","date":"2026-01-09","category":"Générales","type":"Dépense","amount":1500,"account":"Revenus MAZDA","subcategory":"Péage","note":"Attente Joel"},{"id":"t1263","date":"2026-01-09","category":"Divertissement","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1264","date":"2026-01-09","category":"Santé","type":"Dépense","amount":10480,"account":"SALAIRE","subcategory":"VG","note":"AT"},{"id":"t1265","date":"2026-01-09","category":"Invitation","type":"Dépense","amount":16000,"account":"SALAIRE","note":"Joel"},{"id":"t1266","date":"2026-01-09","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1257","date":"2026-01-10","category":"Générales","type":"Dépense","amount":500,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1258","date":"2026-01-10","category":"Divertissement","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1252","date":"2026-01-11","category":"Shopping","type":"Dépense","amount":2000,"account":"SALAIRE"},{"id":"t1253","date":"2026-01-11","category":"Shopping","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t1254","date":"2026-01-11","category":"Prêt","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Drissa"},{"id":"t1255","date":"2026-01-11","category":"Divertissement","type":"Dépense","amount":800,"account":"SALAIRE","subcategory":"Boisson"},{"id":"t1256","date":"2026-01-11","category":"Voiture","type":"Dépense","amount":3000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1247","date":"2026-01-12","category":"Divertissement","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1248","date":"2026-01-12","category":"Enfants & Maman","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t1249","date":"2026-01-12","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1250","date":"2026-01-12","category":"Générales","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t1251","date":"2026-01-12","category":"Ajustement","type":"Dépense","amount":1140,"account":"SIB"},{"id":"t1244","date":"2026-01-13","category":"Shopping","type":"Dépense","amount":15200,"account":"SALAIRE"},{"id":"t1245","date":"2026-01-13","category":"Aliments","type":"Dépense","amount":6500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1246","date":"2026-01-13","category":"Générales","type":"Dépense","amount":500,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1234","date":"2026-01-14","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1235","date":"2026-01-14","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Restaurant"},{"id":"t1236","date":"2026-01-14","category":"Ajustement","type":"Dépense","amount":13490,"account":"SIB"},{"id":"t1237","date":"2026-01-14","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1238","date":"2026-01-14","category":"Dette","type":"Dépense","amount":12276,"account":"SIB","note":"Remboursement prêt"},{"id":"t1239","date":"2026-01-14","category":"Vente Pompe","type":"Revenu","amount":337000,"account":"SIB"},{"id":"t1240","date":"2026-01-14","category":"GRUNDFOS","type":"Dépense","amount":80000,"account":"PETTY CASH","subcategory":"Divertissement","note":"Ogadinma"},{"id":"t1241","date":"2026-01-14","category":"Divertissement","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1242","date":"2026-01-14","category":"Invitation","type":"Dépense","amount":20000,"account":"SALAIRE","note":"Ismo"},{"id":"t1243","date":"2026-01-14","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t1229","date":"2026-01-15","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1230","date":"2026-01-15","category":"Cadeaux","type":"Dépense","amount":2020,"account":"SALAIRE","subcategory":"Femme"},{"id":"t1231","date":"2026-01-15","category":"Shopping","type":"Dépense","amount":15000,"account":"SALAIRE","note":"Spray insectes"},{"id":"t1232","date":"2026-01-15","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","subcategory":"Femme","note":"Charlene"},{"id":"t1233","date":"2026-01-15","category":"Logement","type":"Dépense","amount":50500,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t1223","date":"2026-01-16","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1224","date":"2026-01-16","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA","note":"Fevrier"},{"id":"t1225","date":"2026-01-16","category":"Ajustement","type":"Dépense","amount":101,"account":"SIB"},{"id":"t1226","date":"2026-01-16","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1227","date":"2026-01-16","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1228","date":"2026-01-16","category":"Abonnements","type":"Dépense","amount":1955,"account":"SIB","subcategory":"Spotify"},{"id":"t1217","date":"2026-01-17","category":"Shopping","type":"Dépense","amount":9500,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t1218","date":"2026-01-17","category":"Divertissement","type":"Dépense","amount":88000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1219","date":"2026-01-17","category":"Shopping","type":"Dépense","amount":3200,"account":"SALAIRE"},{"id":"t1220","date":"2026-01-17","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1221","date":"2026-01-17","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1222","date":"2026-01-17","category":"Divertissement","type":"Dépense","amount":21515,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1214","date":"2026-01-18","category":"Ajustement","type":"Dépense","amount":10854,"account":"SIB"},{"id":"t1215","date":"2026-01-18","category":"Divertissement","type":"Dépense","amount":45000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1216","date":"2026-01-18","category":"Cadeaux","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Pourboire","note":"Dydime"},{"id":"t1210","date":"2026-01-19","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1211","date":"2026-01-19","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","note":"Ladji"},{"id":"t1212","date":"2026-01-19","category":"Divertissement","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1213","date":"2026-01-19","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1205","date":"2026-01-20","category":"Aliments","type":"Dépense","amount":14000,"account":"SIB","subcategory":"Dîner"},{"id":"t1206","date":"2026-01-20","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","note":"Pass Juliana"},{"id":"t1207","date":"2026-01-20","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1208","date":"2026-01-20","category":"Aliments","type":"Dépense","amount":4500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1209","date":"2026-01-20","category":"Revenu général","type":"Revenu","amount":29305,"account":"SIB"},{"id":"t1202","date":"2026-01-21","category":"Voiture","type":"Dépense","amount":51500,"account":"Revenus MAZDA","subcategory":"Assurance"},{"id":"t1203","date":"2026-01-21","category":"Shopping","type":"Dépense","amount":34000,"account":"SIB","note":"6 Draps"},{"id":"t1204","date":"2026-01-21","category":"Aliments","type":"Dépense","amount":600,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1196","date":"2026-01-22","category":"Divertissement","type":"Dépense","amount":30000,"account":"SIB","subcategory":"BAP","note":"Moon"},{"id":"t1197","date":"2026-01-22","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":500,"account":"PUMP","subcategory":"Timbre","note":"Achat timbre"},{"id":"t1198","date":"2026-01-22","category":"Générales","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Péage"},{"id":"t1199","date":"2026-01-22","category":"Aliments","type":"Dépense","amount":7500,"account":"SIB","subcategory":"Dîner"},{"id":"t1200","date":"2026-01-22","category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":50000,"account":"PUMP","subcategory":"FNE"},{"id":"t1201","date":"2026-01-22","category":"Shopping","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t1191","date":"2026-01-23","category":"Cadeaux","type":"Dépense","amount":11000,"account":"SIB","subcategory":"Femme","note":"Josepha"},{"id":"t1192","date":"2026-01-23","category":"Aliments","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Dîner"},{"id":"t1193","date":"2026-01-23","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","note":"Elvis"},{"id":"t1194","date":"2026-01-23","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t1195","date":"2026-01-23","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t1182","date":"2026-01-24","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1183","date":"2026-01-24","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1184","date":"2026-01-24","category":"Divertissement","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Alcool"},{"id":"t1185","date":"2026-01-24","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1186","date":"2026-01-24","category":"Cadeaux","type":"Dépense","amount":40400,"account":"SIB","subcategory":"Femme","note":"Josepha"},{"id":"t1187","date":"2026-01-24","category":"Shopping","type":"Dépense","amount":102000,"account":"SIB","subcategory":"Draps"},{"id":"t1188","date":"2026-01-24","category":"Divertissement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Alcool"},{"id":"t1189","date":"2026-01-24","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1190","date":"2026-01-24","category":"Shopping","type":"Dépense","amount":32000,"account":"SIB"},{"id":"t1177","date":"2026-01-25","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SIB","subcategory":"Ruth"},{"id":"t1178","date":"2026-01-25","category":"Divertissement","type":"Dépense","amount":13000,"account":"SIB","subcategory":"Alcool"},{"id":"t1179","date":"2026-01-25","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Femme","note":"Josepha"},{"id":"t1180","date":"2026-01-25","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","note":"Olo"},{"id":"t1181","date":"2026-01-25","category":"Invitation","type":"Dépense","amount":28000,"account":"SIB","subcategory":"Femmes","note":"Fatim"},{"id":"t1172","date":"2026-01-26","category":"Pack Club","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t1173","date":"2026-01-26","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t1174","date":"2026-01-26","category":"Ajustement","type":"Revenu","amount":362,"account":"SIB"},{"id":"t1175","date":"2026-01-26","category":"Abonnements","type":"Dépense","amount":5280,"account":"SIB","subcategory":"Money Coach"},{"id":"t1176","date":"2026-01-26","category":"Aliments","type":"Dépense","amount":2020,"account":"SIB","subcategory":"Dîner"},{"id":"t1170","date":"2026-01-27","category":"Aliments","type":"Dépense","amount":3500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1171","date":"2026-01-27","category":"Aliments","type":"Dépense","amount":3650,"account":"SIB","subcategory":"Dîner"},{"id":"t1167","date":"2026-01-28","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Pourboire"},{"id":"t1168","date":"2026-01-28","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","note":"Roland"},{"id":"t1169","date":"2026-01-28","category":"Aliments","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1159","date":"2026-01-29","category":"Invitation","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Femmes","note":"Letissia"},{"id":"t1160","date":"2026-01-29","category":"Logement","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Location","note":"Armande"},{"id":"t1161","date":"2026-01-29","category":"Invitation","type":"Dépense","amount":30030,"account":"SIB","note":"Déjeuner Abengourou"},{"id":"t1162","date":"2026-01-29","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1163","date":"2026-01-29","category":"Personnel","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Coiffure"},{"id":"t1164","date":"2026-01-29","category":"Aliments","type":"Dépense","amount":3630,"account":"SIB","subcategory":"Déjeuner"},{"id":"t1165","date":"2026-01-29","category":"Shopping","type":"Dépense","amount":500,"account":"SIB"},{"id":"t1166","date":"2026-01-29","category":"Divertissement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Alcool"},{"id":"t1157","date":"2026-01-30","category":"Invitation","type":"Dépense","amount":15000,"account":"SIB","note":"Ndjore"},{"id":"t1158","date":"2026-01-30","category":"Enfants & Maman","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Maman","note":"Glacière"},{"id":"t1137","date":"2026-01-31","category":"Shopping","type":"Dépense","amount":2500,"account":"SIB","note":"Lub"},{"id":"t1138","date":"2026-01-31","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Femme","note":"Juliana"},{"id":"t1139","date":"2026-01-31","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t1140","date":"2026-01-31","category":"Cadeaux","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Pourboire","note":"Juliana"},{"id":"t1141","date":"2026-01-31","category":"Voyage","type":"Dépense","amount":28700,"account":"SIB","subcategory":"Un hôtel","note":"Hotel Akabla"},{"id":"t1142","date":"2026-01-31","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1143","date":"2026-01-31","category":"Shopping","type":"Dépense","amount":20000,"account":"SALAIRE"},{"id":"t1144","date":"2026-01-31","category":"Divertissement","type":"Dépense","amount":13130,"account":"SALAIRE","subcategory":"Alcool","note":"Atto abg"},{"id":"t1145","date":"2026-01-31","category":"Aliments","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Déjeuner","note":"Parker"},{"id":"t1146","date":"2026-01-31","category":"Divertissement","type":"Dépense","amount":50000,"account":"SIB","subcategory":"Alcool","note":"Flex abg"},{"id":"t1147","date":"2026-01-31","category":"Cadeaux","type":"Dépense","amount":40000,"account":"SIB","subcategory":"Femme","note":"Juliana"},{"id":"t1148","date":"2026-01-31","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","note":"Fabrice"},{"id":"t1149","date":"2026-01-31","category":"Cadeaux","type":"Dépense","amount":25000,"account":"SIB","note":"Parker"},{"id":"t1150","date":"2026-01-31","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1151","date":"2026-01-31","category":"Enfants & Maman","type":"Dépense","amount":110000,"account":"SIB","subcategory":"Maman"},{"id":"t1152","date":"2026-01-31","category":"Enfants & Maman","type":"Dépense","amount":55550,"account":"SIB","subcategory":"Nesher"},{"id":"t1153","date":"2026-01-31","category":"Logement","type":"Dépense","amount":380000,"account":"Dépôt LOYER","subcategory":"Location","note":"Fevrier"},{"id":"t1154","date":"2026-01-31","category":"Ajustement","type":"Dépense","amount":36094,"account":"SIB","note":"Voyage Abengourou"},{"id":"t1155","date":"2026-01-31","category":"Vente Pompe","type":"Revenu","amount":50000,"account":"SIB","note":"Moteur 5.5"},{"id":"t1156","date":"2026-01-31","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre"},{"id":"t1135","date":"2026-02-01","category":"Générales","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t1136","date":"2026-02-01","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1130","date":"2026-02-02","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1131","date":"2026-02-02","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1132","date":"2026-02-02","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Femme","note":"Binta Barry"},{"id":"t1133","date":"2026-02-02","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","note":"JP"},{"id":"t1134","date":"2026-02-02","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1127","date":"2026-02-03","category":"Ajustement","type":"Dépense","amount":8103,"account":"SALAIRE"},{"id":"t1128","date":"2026-02-03","category":"Divertissement","type":"Dépense","amount":21000,"account":"SALAIRE","subcategory":"Alcool","note":"Toussaint"},{"id":"t1129","date":"2026-02-03","category":"Aliments","type":"Dépense","amount":900,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1122","date":"2026-02-04","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1123","date":"2026-02-04","category":"Invitation","type":"Dépense","amount":45000,"account":"SALAIRE","subcategory":"Femmes","note":"Binta"},{"id":"t1124","date":"2026-02-04","category":"Vêtements","type":"Dépense","amount":9090,"account":"SALAIRE"},{"id":"t1125","date":"2026-02-04","category":"Aliments","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1126","date":"2026-02-04","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Cotisations","note":"Ndjimsi"},{"id":"t1116","date":"2026-02-05","category":"Shopping","type":"Dépense","amount":500,"account":"SALAIRE"},{"id":"t1117","date":"2026-02-05","category":"Vêtements","type":"Dépense","amount":5000,"account":"SALAIRE"},{"id":"t1118","date":"2026-02-05","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1119","date":"2026-02-05","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t1120","date":"2026-02-05","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1121","date":"2026-02-05","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t1110","date":"2026-02-06","category":"Aliments","type":"Dépense","amount":23000,"account":"SALAIRE","subcategory":"Invitation","note":"Ouanlo"},{"id":"t1111","date":"2026-02-06","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SALAIRE","note":"Augustin"},{"id":"t1112","date":"2026-02-06","category":"Cadeaux","type":"Dépense","amount":60600,"account":"SALAIRE","subcategory":"Femme","note":"Juliana"},{"id":"t1113","date":"2026-02-06","category":"Divertissement","type":"Dépense","amount":45000,"account":"SALAIRE","subcategory":"Alcool","note":"Basi"},{"id":"t1114","date":"2026-02-06","category":"Shopping","type":"Dépense","amount":36000,"account":"SALAIRE"},{"id":"t1115","date":"2026-02-06","category":"Prêt","type":"Dépense","amount":505000,"account":"Revenus MAZDA","note":"Toussaint"},{"id":"t3031","date":"2026-02-06","category":"Personnel","type":"Dépense","amount":17170,"account":"SALAIRE","subcategory":"Hygiène personnelle","note":"Nettoyage maison"},{"id":"t1103","date":"2026-02-07","category":"Ajustement","type":"Dépense","amount":13084,"account":"SALAIRE"},{"id":"t1104","date":"2026-02-07","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Invitation","note":"Basi"},{"id":"t1105","date":"2026-02-07","category":"Shopping","type":"Dépense","amount":10000,"account":"SALAIRE"},{"id":"t1106","date":"2026-02-07","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1107","date":"2026-02-07","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1108","date":"2026-02-07","category":"Divertissement","type":"Dépense","amount":6365,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1109","date":"2026-02-07","category":"Prêt Orange","type":"Revenu","amount":76984,"account":"SIB"},{"id":"t1098","date":"2026-02-08","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1099","date":"2026-02-08","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1100","date":"2026-02-08","category":"Abonnements","type":"Dépense","amount":5848,"account":"SALAIRE","note":"Dazn"},{"id":"t1101","date":"2026-02-08","category":"Ajustement","type":"Dépense","amount":2877,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1102","date":"2026-02-08","category":"Ajustement","type":"Dépense","amount":8051,"account":"SALAIRE"},{"id":"t1097","date":"2026-02-09","category":"Invitation","type":"Dépense","amount":11500,"account":"SALAIRE","note":"Gnakri"},{"id":"t1095","date":"2026-02-10","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1096","date":"2026-02-10","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1092","date":"2026-02-11","category":"Shopping","type":"Dépense","amount":13425,"account":"SALAIRE"},{"id":"t1093","date":"2026-02-11","category":"Cadeaux","type":"Dépense","amount":3500,"account":"SALAIRE","note":"Binta Medoc"},{"id":"t1094","date":"2026-02-11","category":"Invitation","type":"Dépense","amount":13500,"account":"SALAIRE","subcategory":"Femmes","note":"Binta"},{"id":"t1077","date":"2026-02-12","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1078","date":"2026-02-12","category":"Cadeaux","type":"Dépense","amount":800,"account":"SALAIRE","note":"Roland"},{"id":"t1079","date":"2026-02-12","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1080","date":"2026-02-12","category":"INVEST SGO","type":"Dépense","amount":271,"account":"SGO","subcategory":"Frais"},{"id":"t1081","date":"2026-02-12","category":"INVEST SGO","type":"Dépense","amount":15000,"account":"SGO","subcategory":"NSIA"},{"id":"t1082","date":"2026-02-12","category":"INVEST SGO","type":"Dépense","amount":633,"account":"SGO","subcategory":"Frais"},{"id":"t1083","date":"2026-02-12","category":"INVEST SGO","type":"Dépense","amount":35000,"account":"SGO","subcategory":"NSIA"},{"id":"t1084","date":"2026-02-12","category":"Voiture","type":"Dépense","amount":100000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Batterie"},{"id":"t1085","date":"2026-02-12","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t1086","date":"2026-02-12","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1087","date":"2026-02-12","category":"Générales","type":"Dépense","amount":300,"account":"SALAIRE"},{"id":"t1088","date":"2026-02-12","category":"Voiture","type":"Dépense","amount":45450,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Plaquettes de frein"},{"id":"t1089","date":"2026-02-12","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1090","date":"2026-02-12","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1091","date":"2026-02-12","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1068","date":"2026-02-13","category":"Abonnements","type":"Dépense","amount":1987,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t1069","date":"2026-02-13","category":"Ajustement","type":"Dépense","amount":1899,"account":"SALAIRE"},{"id":"t1070","date":"2026-02-13","category":"Voiture","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Remplacement Plaquettes"},{"id":"t1071","date":"2026-02-13","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Femme","note":"Binta Medoc"},{"id":"t1072","date":"2026-02-13","category":"Cadeaux","type":"Dépense","amount":800,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t1073","date":"2026-02-13","category":"Enfants & Maman","type":"Dépense","amount":27000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t1074","date":"2026-02-13","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1075","date":"2026-02-13","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1076","date":"2026-02-13","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA","note":"Mars"},{"id":"t1055","date":"2026-02-14","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t1056","date":"2026-02-14","category":"Cadeaux","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Femme","note":"Binta"},{"id":"t1057","date":"2026-02-14","category":"Invitation","type":"Dépense","amount":25500,"account":"SALAIRE","subcategory":"Femmes","note":"Binta"},{"id":"t1058","date":"2026-02-14","category":"Aliments","type":"Dépense","amount":7020,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1059","date":"2026-02-14","category":"Invitation","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femmes","note":"Ruth"},{"id":"t1060","date":"2026-02-14","category":"Générales","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Péage"},{"id":"t1061","date":"2026-02-14","category":"Enfants & Maman","type":"Dépense","amount":1600,"account":"SALAIRE","subcategory":"Nesher","payee":"+ autres","note":"Envoi Carte d'assurance"},{"id":"t1062","date":"2026-02-14","category":"Shopping","type":"Dépense","amount":7000,"account":"SALAIRE"},{"id":"t1063","date":"2026-02-14","category":"Enfants & Maman","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t1064","date":"2026-02-14","category":"Cadeaux","type":"Dépense","amount":35000,"account":"SALAIRE","subcategory":"Femme","note":"Fleurs Binta"},{"id":"t1065","date":"2026-02-14","category":"Cadeaux","type":"Dépense","amount":178500,"account":"SALAIRE","subcategory":"Ruth","note":"Cadeau"},{"id":"t1066","date":"2026-02-14","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1067","date":"2026-02-14","category":"Générales","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Police","note":"Rackette"},{"id":"t1048","date":"2026-02-15","category":"Aliments","type":"Dépense","amount":6060,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1049","date":"2026-02-15","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1050","date":"2026-02-15","category":"Ajustement","type":"Dépense","amount":12880,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t1051","date":"2026-02-15","category":"Vêtements","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Accessoires","note":"Chapeau"},{"id":"t1052","date":"2026-02-15","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1053","date":"2026-02-15","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool","note":"Lingot"},{"id":"t1054","date":"2026-02-15","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1042","date":"2026-02-16","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1043","date":"2026-02-16","category":"GRUNDFOS","type":"Dépense","amount":9500,"account":"PETTY CASH","note":"Souris"},{"id":"t1044","date":"2026-02-16","category":"Cadeaux","type":"Dépense","amount":61000,"account":"SALAIRE","note":"Roland"},{"id":"t1045","date":"2026-02-16","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Femme","note":"Syntyche"},{"id":"t1046","date":"2026-02-16","category":"Shopping","type":"Dépense","amount":2800,"account":"SALAIRE"},{"id":"t1047","date":"2026-02-16","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1040","date":"2026-02-17","category":"Aliments","type":"Dépense","amount":400,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1041","date":"2026-02-17","category":"GRUNDFOS","type":"Dépense","amount":3570,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t1033","date":"2026-02-18","category":"GRUNDFOS","type":"Dépense","amount":70000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Paniers ramadan"},{"id":"t1034","date":"2026-02-18","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1035","date":"2026-02-18","category":"Ajustement","type":"Dépense","amount":2322,"account":"SALAIRE"},{"id":"t1036","date":"2026-02-18","category":"Divertissement","type":"Dépense","amount":25300,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1037","date":"2026-02-18","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1038","date":"2026-02-18","category":"Aliments","type":"Dépense","amount":2700,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1039","date":"2026-02-18","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t1029","date":"2026-02-19","category":"Utilitaires","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Nettoyage"},{"id":"t1030","date":"2026-02-19","category":"Aliments","type":"Dépense","amount":1050,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1031","date":"2026-02-19","category":"Aliments","type":"Dépense","amount":6500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1032","date":"2026-02-19","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t1019","date":"2026-02-20","category":"Invitation","type":"Dépense","amount":9000,"account":"SALAIRE","note":"Visite Parker"},{"id":"t1020","date":"2026-02-20","category":"Cadeaux","type":"Dépense","amount":42000,"account":"SALAIRE","note":"Visite Parker"},{"id":"t1021","date":"2026-02-20","category":"Enfants & Maman","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Maman","note":"Telephone"},{"id":"t1022","date":"2026-02-20","category":"GRUNDFOS","type":"Dépense","amount":98644,"account":"PETTY CASH","subcategory":"AUTRES","note":"JTB Visa David et Felistas"},{"id":"t1023","date":"2026-02-20","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1024","date":"2026-02-20","category":"Cadeaux","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Pourboire","note":"Etran et autres"},{"id":"t1025","date":"2026-02-20","category":"Divertissement","type":"Dépense","amount":45000,"account":"SALAIRE","subcategory":"Alcool","note":"Etran"},{"id":"t1026","date":"2026-02-20","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1027","date":"2026-02-20","category":"GRUNDFOS","type":"Dépense","amount":505000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Marketing JTB"},{"id":"t1028","date":"2026-02-20","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t1016","date":"2026-02-21","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t1017","date":"2026-02-21","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t1018","date":"2026-02-21","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","note":"Hamed"},{"id":"t1013","date":"2026-02-22","category":"Aliments","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1014","date":"2026-02-22","category":"Divertissement","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1015","date":"2026-02-22","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t1010","date":"2026-02-23","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t1011","date":"2026-02-23","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t1012","date":"2026-02-23","category":"Aliments","type":"Dépense","amount":1800,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t1000","date":"2026-02-24","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire","note":"Mahou"},{"id":"t1001","date":"2026-02-24","category":"Divertissement","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t1002","date":"2026-02-24","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","note":"Robokauf"},{"id":"t1003","date":"2026-02-24","category":"Shopping","type":"Dépense","amount":11500,"account":"SALAIRE","note":"Rasoir"},{"id":"t1004","date":"2026-02-24","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t1005","date":"2026-02-24","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Cousins"},{"id":"t1006","date":"2026-02-24","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t1007","date":"2026-02-24","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1008","date":"2026-02-24","category":"Divertissement","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t1009","date":"2026-02-24","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t991","date":"2026-02-25","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t992","date":"2026-02-25","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre"},{"id":"t993","date":"2026-02-25","category":"Enfants & Maman","type":"Dépense","amount":55550,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t994","date":"2026-02-25","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t995","date":"2026-02-25","category":"Cadeaux","type":"Dépense","amount":1515,"account":"SALAIRE","note":"Armani"},{"id":"t996","date":"2026-02-25","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t997","date":"2026-02-25","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t998","date":"2026-02-25","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t999","date":"2026-02-25","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t983","date":"2026-02-26","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Certificat résidence"},{"id":"t984","date":"2026-02-26","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Entretien"},{"id":"t985","date":"2026-02-26","category":"Générales","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Police"},{"id":"t986","date":"2026-02-26","category":"Ajustement","type":"Dépense","amount":17094,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t987","date":"2026-02-26","category":"Pack Club","type":"Dépense","amount":10000,"account":"SIB"},{"id":"t988","date":"2026-02-26","category":"GRUNDFOS","type":"Dépense","amount":3500,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t989","date":"2026-02-26","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t990","date":"2026-02-26","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t974","date":"2026-02-27","category":"Ajustement","type":"Dépense","amount":1834,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t975","date":"2026-02-27","category":"Ajustement","type":"Dépense","amount":3400,"account":"SALAIRE"},{"id":"t976","date":"2026-02-27","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"Binta"},{"id":"t977","date":"2026-02-27","category":"Aliments","type":"Dépense","amount":6500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t978","date":"2026-02-27","category":"Logement","type":"Dépense","amount":380000,"account":"Dépôt LOYER","subcategory":"Location"},{"id":"t979","date":"2026-02-27","category":"Vêtements","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Accessoires"},{"id":"t980","date":"2026-02-27","category":"Cadeaux","type":"Dépense","amount":15195,"account":"SALAIRE","subcategory":"Femme","note":"Diane Accra"},{"id":"t981","date":"2026-02-27","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t982","date":"2026-02-27","category":"Invitation","type":"Dépense","amount":17000,"account":"SALAIRE","subcategory":"Femmes"},{"id":"t968","date":"2026-02-28","category":"GRUNDFOS","type":"Dépense","amount":34000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t969","date":"2026-02-28","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Maman"},{"id":"t970","date":"2026-02-28","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Maman"},{"id":"t971","date":"2026-02-28","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t972","date":"2026-02-28","category":"Cadeaux","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t973","date":"2026-02-28","category":"Invitation","type":"Dépense","amount":99000,"account":"SALAIRE","subcategory":"Femmes","note":"Fatim Mounkaila"},{"id":"t964","date":"2026-03-01","category":"Loyer","type":"Revenu","amount":1500000,"account":"Dépôt LOYER","note":"Mars avril Mai 2026"},{"id":"t965","date":"2026-03-01","category":"Aliments","type":"Dépense","amount":2505,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t966","date":"2026-03-01","category":"Abonnements","type":"Dépense","amount":5254,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t967","date":"2026-03-01","category":"Abonnements","type":"Dépense","amount":11909,"account":"SALAIRE","note":"Claude AI"},{"id":"t955","date":"2026-03-02","category":"Personnel","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Produits de beauté"},{"id":"t956","date":"2026-03-02","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t957","date":"2026-03-02","category":"Divertissement","type":"Dépense","amount":9000,"account":"SALAIRE"},{"id":"t958","date":"2026-03-02","category":"Aliments","type":"Dépense","amount":7500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t959","date":"2026-03-02","category":"Ajustement","type":"Dépense","amount":499,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t960","date":"2026-03-02","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Femme","note":"Aliments"},{"id":"t961","date":"2026-03-02","category":"Divertissement","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t962","date":"2026-03-02","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t963","date":"2026-03-02","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Femme","note":"Miracle parf Yop"},{"id":"t953","date":"2026-03-03","category":"Shopping","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t954","date":"2026-03-03","category":"Shopping","type":"Dépense","amount":1800,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t944","date":"2026-03-04","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"BAP","note":"Residence"},{"id":"t945","date":"2026-03-04","category":"Enfants & Maman","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t946","date":"2026-03-04","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t947","date":"2026-03-04","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t948","date":"2026-03-04","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t949","date":"2026-03-04","category":"Aliments","type":"Dépense","amount":600,"account":"SALAIRE"},{"id":"t950","date":"2026-03-04","category":"Shopping","type":"Dépense","amount":56003,"account":"SALAIRE"},{"id":"t951","date":"2026-03-04","category":"GRUNDFOS","type":"Dépense","amount":98165,"account":"PETTY CASH","subcategory":"AUTRES","note":"Paniers ramadan Vinci"},{"id":"t952","date":"2026-03-04","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Femme","note":"Juliana"},{"id":"t935","date":"2026-03-05","category":"Divertissement","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t936","date":"2026-03-05","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Stéphanie"},{"id":"t937","date":"2026-03-05","category":"Vente Pompe","type":"Revenu","amount":150000,"account":"SIB"},{"id":"t938","date":"2026-03-05","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t939","date":"2026-03-05","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t940","date":"2026-03-05","category":"Santé","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t941","date":"2026-03-05","category":"Shopping","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t942","date":"2026-03-05","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t943","date":"2026-03-05","category":"Cadeaux","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Femme"},{"id":"t923","date":"2026-03-06","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Femme","note":"Michou"},{"id":"t924","date":"2026-03-06","category":"Shopping","type":"Dépense","amount":14000,"account":"SALAIRE"},{"id":"t925","date":"2026-03-06","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Guigui"},{"id":"t926","date":"2026-03-06","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t927","date":"2026-03-06","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t928","date":"2026-03-06","category":"GRUNDFOS","type":"Dépense","amount":2055,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t929","date":"2026-03-06","category":"Utilitaires","type":"Dépense","amount":17000,"account":"SALAIRE","subcategory":"Nettoyage"},{"id":"t930","date":"2026-03-06","category":"Divertissement","type":"Dépense","amount":72000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t931","date":"2026-03-06","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Femme","note":"Binta"},{"id":"t932","date":"2026-03-06","category":"Ajustement","type":"Dépense","amount":14536,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t933","date":"2026-03-06","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t934","date":"2026-03-06","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Pourboire","note":"Dydime"},{"id":"t920","date":"2026-03-07","category":"Générales","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Police"},{"id":"t921","date":"2026-03-07","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t922","date":"2026-03-07","category":"Invitation","type":"Dépense","amount":20500,"account":"SALAIRE","note":"Charly"},{"id":"t917","date":"2026-03-08","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t918","date":"2026-03-08","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t919","date":"2026-03-08","category":"Invitation","type":"Dépense","amount":27000,"account":"SALAIRE"},{"id":"t911","date":"2026-03-09","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t912","date":"2026-03-09","category":"Générales","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Péage"},{"id":"t913","date":"2026-03-09","category":"Aliments","type":"Dépense","amount":7500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t914","date":"2026-03-09","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t915","date":"2026-03-09","category":"Prêt","type":"Dépense","amount":250000,"account":"SALAIRE","note":"Sanogo"},{"id":"t916","date":"2026-03-09","category":"GRUNDFOS","type":"Dépense","amount":50000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Panier Christian"},{"id":"t908","date":"2026-03-10","category":"Divertissement","type":"Dépense","amount":2525,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t909","date":"2026-03-10","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t910","date":"2026-03-10","category":"Aliments","type":"Dépense","amount":6500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t898","date":"2026-03-11","category":"Divertissement","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Femme","note":"Syntyche"},{"id":"t899","date":"2026-03-11","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA","note":"Avril"},{"id":"t900","date":"2026-03-11","category":"Santé","type":"Dépense","amount":12900,"account":"SALAIRE","subcategory":"Médicaments","note":"ATT"},{"id":"t901","date":"2026-03-11","category":"Ajustement","type":"Dépense","amount":185,"account":"SIB","subcategory":"Frais"},{"id":"t902","date":"2026-03-11","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t903","date":"2026-03-11","category":"Ajustement","type":"Revenu","amount":5945,"account":"SALAIRE"},{"id":"t904","date":"2026-03-11","category":"Divertissement","type":"Dépense","amount":9000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t905","date":"2026-03-11","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t906","date":"2026-03-11","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Pourboire","note":"Artiste"},{"id":"t907","date":"2026-03-11","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t891","date":"2026-03-12","category":"Divertissement","type":"Dépense","amount":13130,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t892","date":"2026-03-12","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","note":"Olokpacha"},{"id":"t893","date":"2026-03-12","category":"Ajustement","type":"Revenu","amount":9000,"account":"SIB"},{"id":"t894","date":"2026-03-12","category":"Divertissement","type":"Dépense","amount":33535,"account":"SALAIRE","subcategory":"Femme","note":"Precious"},{"id":"t895","date":"2026-03-12","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t896","date":"2026-03-12","category":"Abonnements","type":"Dépense","amount":1960,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t897","date":"2026-03-12","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Appel"},{"id":"t883","date":"2026-03-13","category":"GRUNDFOS","type":"Dépense","amount":500,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t884","date":"2026-03-13","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t885","date":"2026-03-13","category":"GRUNDFOS","type":"Dépense","amount":580000,"account":"PETTY CASH","subcategory":"AUTRES","note":"Marketing JTB"},{"id":"t886","date":"2026-03-13","category":"Ajustement","type":"Dépense","amount":2200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t887","date":"2026-03-13","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t888","date":"2026-03-13","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t889","date":"2026-03-13","category":"Cadeaux","type":"Dépense","amount":21000,"account":"SALAIRE","subcategory":"Pourboire","note":"Rumba"},{"id":"t890","date":"2026-03-13","category":"Invitation","type":"Dépense","amount":55000,"account":"SALAIRE","subcategory":"Femmes","note":"Ruth"},{"id":"t876","date":"2026-03-14","category":"Générales","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Livraison Pompe"},{"id":"t877","date":"2026-03-14","category":"Invitation","type":"Dépense","amount":31010,"account":"SALAIRE","subcategory":"Femmes","note":"Blanche"},{"id":"t878","date":"2026-03-14","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Pourboire","note":"Filles Armande"},{"id":"t879","date":"2026-03-14","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t880","date":"2026-03-14","category":"Vêtements","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Accessoires"},{"id":"t881","date":"2026-03-14","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t882","date":"2026-03-14","category":"Divertissement","type":"Dépense","amount":74000,"account":"SALAIRE","subcategory":"Alcool","note":"Anniversaire HEMRA"},{"id":"t873","date":"2026-03-15","category":"Divertissement","type":"Dépense","amount":35000,"account":"SALAIRE","subcategory":"Femme","note":"Precious"},{"id":"t874","date":"2026-03-15","category":"Invitation","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Femmes","note":"Precious"},{"id":"t875","date":"2026-03-15","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t868","date":"2026-03-16","category":"Santé","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t869","date":"2026-03-16","category":"Ajustement","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t870","date":"2026-03-16","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t871","date":"2026-03-16","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t872","date":"2026-03-16","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","note":"Jean Philippe"},{"id":"t866","date":"2026-03-17","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t867","date":"2026-03-17","category":"Aliments","type":"Dépense","amount":11700,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t862","date":"2026-03-18","category":"Cadeaux","type":"Dépense","amount":1515,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t863","date":"2026-03-18","category":"Aliments","type":"Dépense","amount":2100,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t864","date":"2026-03-18","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t865","date":"2026-03-18","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t852","date":"2026-03-19","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t853","date":"2026-03-19","category":"INVEST SGO","type":"Dépense","amount":29,"account":"SGO","subcategory":"Frais","note":"Monetaris"},{"id":"t854","date":"2026-03-19","category":"INVEST SGO","type":"Dépense","amount":3000,"account":"SGO","note":"Monetaris"},{"id":"t855","date":"2026-03-19","category":"Abonnements","type":"Dépense","amount":3000,"account":"SALAIRE","note":"API Key"},{"id":"t856","date":"2026-03-19","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t857","date":"2026-03-19","category":"GRUNDFOS","type":"Dépense","amount":33000,"account":"PETTY CASH","subcategory":"Impression","note":"Carte de visite"},{"id":"t858","date":"2026-03-19","category":"Cadeaux","type":"Dépense","amount":60600,"account":"SALAIRE","subcategory":"Femme","note":"Sylvia"},{"id":"t859","date":"2026-03-19","category":"Déménagement","type":"Dépense","amount":15000,"account":"PETTY CASH","note":"Demarcheur"},{"id":"t860","date":"2026-03-19","category":"Ajustement","type":"Dépense","amount":2003,"account":"SALAIRE"},{"id":"t861","date":"2026-03-19","category":"Divertissement","type":"Dépense","amount":6020,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t848","date":"2026-03-21","category":"Aliments","type":"Dépense","amount":9000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t849","date":"2026-03-21","category":"Invitation","type":"Dépense","amount":8500,"account":"SALAIRE","note":"Etran"},{"id":"t850","date":"2026-03-21","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t851","date":"2026-03-21","category":"Divertissement","type":"Dépense","amount":33000,"account":"SALAIRE"},{"id":"t847","date":"2026-03-22","category":"Invitation","type":"Dépense","amount":31500,"account":"SALAIRE","subcategory":"Femmes","note":"Roxane"},{"id":"t841","date":"2026-03-23","category":"Aliments","type":"Dépense","amount":4700,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t842","date":"2026-03-23","category":"Générales","type":"Dépense","amount":300,"account":"SALAIRE","subcategory":"Impression"},{"id":"t843","date":"2026-03-23","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t844","date":"2026-03-23","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Impression"},{"id":"t845","date":"2026-03-23","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t846","date":"2026-03-23","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Femme","note":"Vicky internet"},{"id":"t831","date":"2026-03-24","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Femme","note":"Fatim internet"},{"id":"t832","date":"2026-03-24","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t833","date":"2026-03-24","category":"Ajustement","type":"Revenu","amount":390528,"account":"SIB","note":"JTB"},{"id":"t834","date":"2026-03-24","category":"Petty Cash","type":"Revenu","amount":1085000,"account":"PETTY CASH","note":"Gadget JTB"},{"id":"t835","date":"2026-03-24","category":"Ajustement","type":"Revenu","amount":9355,"account":"SALAIRE"},{"id":"t836","date":"2026-03-24","category":"GRUNDFOS","type":"Dépense","amount":7575,"account":"PETTY CASH","note":"David"},{"id":"t837","date":"2026-03-24","category":"GRUNDFOS","type":"Dépense","amount":3000,"account":"PETTY CASH","subcategory":"Impression","note":"JTB"},{"id":"t838","date":"2026-03-24","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t839","date":"2026-03-24","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t840","date":"2026-03-24","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t827","date":"2026-03-25","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre"},{"id":"t828","date":"2026-03-25","category":"GRUNDFOS","type":"Dépense","amount":9000,"account":"PETTY CASH","note":"JTB"},{"id":"t829","date":"2026-03-25","category":"GRUNDFOS","type":"Dépense","amount":52000,"account":"PETTY CASH","subcategory":"Restaurant","note":"Dîner avec David"},{"id":"t830","date":"2026-03-25","category":"GRUNDFOS","type":"Dépense","amount":8000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t822","date":"2026-03-26","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool","note":"Ismo"},{"id":"t823","date":"2026-03-26","category":"Vêtements","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Accessoires"},{"id":"t824","date":"2026-03-26","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE","note":"Avril"},{"id":"t825","date":"2026-03-26","category":"GRUNDFOS","type":"Dépense","amount":13130,"account":"PETTY CASH","subcategory":"Restaurant"},{"id":"t826","date":"2026-03-26","category":"GRUNDFOS","type":"Dépense","amount":500,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t814","date":"2026-03-27","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t815","date":"2026-03-27","category":"GRUNDFOS","type":"Dépense","amount":64800,"account":"PETTY CASH","subcategory":"Restaurant","note":"David"},{"id":"t816","date":"2026-03-27","category":"Ajustement","type":"Revenu","amount":276,"account":"SALAIRE"},{"id":"t817","date":"2026-03-27","category":"GRUNDFOS","type":"Dépense","amount":60000,"account":"PETTY CASH","subcategory":"Divertissement","note":"David Djimmys"},{"id":"t818","date":"2026-03-27","category":"GRUNDFOS","type":"Dépense","amount":15000,"account":"PETTY CASH","subcategory":"Divertissement","note":"Pourboire David rumba"},{"id":"t819","date":"2026-03-27","category":"Aliments","type":"Dépense","amount":1010,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t820","date":"2026-03-27","category":"Déménagement","type":"Dépense","amount":3177430,"account":"Revenus MAZDA","payee":"Immobilier","note":"Déménagement Drayât"},{"id":"t821","date":"2026-03-27","category":"GRUNDFOS","type":"Dépense","amount":63000,"account":"PETTY CASH","subcategory":"Divertissement","note":"David"},{"id":"t802","date":"2026-03-28","category":"General","type":"Revenu","amount":200000,"account":"SALAIRE","payee":"Sanogo","note":"Remboursement Prêt"},{"id":"t803","date":"2026-03-28","category":"GRUNDFOS","type":"Dépense","amount":1000,"account":"PETTY CASH","subcategory":"Eau"},{"id":"t804","date":"2026-03-28","category":"GRUNDFOS","type":"Dépense","amount":10100,"account":"PETTY CASH","subcategory":"Restaurant","note":"Aliments"},{"id":"t805","date":"2026-03-28","category":"Securicompte","type":"Dépense","amount":912,"account":"SIB"},{"id":"t806","date":"2026-03-28","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t807","date":"2026-03-28","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t808","date":"2026-03-28","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t809","date":"2026-03-28","category":"Générales","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Péage"},{"id":"t810","date":"2026-03-28","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Pourboire","note":"Miracle"},{"id":"t811","date":"2026-03-28","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire","note":"Yop"},{"id":"t812","date":"2026-03-28","category":"Aliments","type":"Dépense","amount":8600,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t813","date":"2026-03-28","category":"Divertissement","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t795","date":"2026-03-29","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t796","date":"2026-03-29","category":"Utilitaires","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Téléphones","note":"Câble"},{"id":"t797","date":"2026-03-29","category":"Invitation","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Femmes","note":"Alvy"},{"id":"t798","date":"2026-03-29","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t799","date":"2026-03-29","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t800","date":"2026-03-29","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Cotisations","note":"Toussaint"},{"id":"t801","date":"2026-03-29","category":"Ajustement","type":"Revenu","amount":75,"account":"SALAIRE"},{"id":"t788","date":"2026-03-30","category":"Aliments","type":"Dépense","amount":100,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t789","date":"2026-03-30","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t790","date":"2026-03-30","category":"Enfants & Maman","type":"Dépense","amount":65065,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t791","date":"2026-03-30","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Femme"},{"id":"t792","date":"2026-03-30","category":"Logement","type":"Dépense","amount":380000,"account":"Dépôt LOYER","subcategory":"Location","note":"Avril"},{"id":"t793","date":"2026-03-30","category":"Invitation","type":"Dépense","amount":13500,"account":"SALAIRE","subcategory":"Femmes","note":"Fatim"},{"id":"t794","date":"2026-03-30","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Maman"},{"id":"t778","date":"2026-03-31","category":"Enfants & Maman","type":"Dépense","amount":9500,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t779","date":"2026-03-31","category":"Enfants & Maman","type":"Dépense","amount":18955,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t780","date":"2026-03-31","category":"Ajustement","type":"Dépense","amount":300,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t781","date":"2026-03-31","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t782","date":"2026-03-31","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t783","date":"2026-03-31","category":"Ajustement","type":"Dépense","amount":7912,"account":"SALAIRE"},{"id":"t784","date":"2026-03-31","category":"Invitation","type":"Dépense","amount":11000,"account":"SALAIRE","note":"Saturnin"},{"id":"t785","date":"2026-03-31","category":"Divertissement","type":"Dépense","amount":3535,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t786","date":"2026-03-31","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Maman"},{"id":"t787","date":"2026-03-31","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t771","date":"2026-04-01","category":"Ajustement","type":"Dépense","amount":14500,"account":"SIB","subcategory":"Frais Bancaire","note":"AFG"},{"id":"t772","date":"2026-04-01","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Mardochee"},{"id":"t773","date":"2026-04-01","category":"Cadeaux","type":"Dépense","amount":25250,"account":"SALAIRE","note":"Etran"},{"id":"t774","date":"2026-04-01","category":"Cadeaux","type":"Dépense","amount":21500,"account":"SALAIRE","subcategory":"Femme","note":"Jose"},{"id":"t775","date":"2026-04-01","category":"Ajustement","type":"Revenu","amount":4000,"account":"SIB"},{"id":"t776","date":"2026-04-01","category":"Aliments","type":"Dépense","amount":800,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t777","date":"2026-04-01","category":"Aliments","type":"Dépense","amount":700,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t764","date":"2026-04-02","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t765","date":"2026-04-02","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t766","date":"2026-04-02","category":"Générales","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Police"},{"id":"t767","date":"2026-04-02","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t768","date":"2026-04-02","category":"Abonnements","type":"Dépense","amount":145,"account":"SALAIRE","note":"Netflix"},{"id":"t769","date":"2026-04-02","category":"Abonnements","type":"Dépense","amount":5961,"account":"SALAIRE","subcategory":"Netflix"},{"id":"t770","date":"2026-04-02","category":"Ajustement","type":"Dépense","amount":498,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t759","date":"2026-04-03","category":"Cadeaux","type":"Dépense","amount":2020,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t760","date":"2026-04-03","category":"Aliments","type":"Dépense","amount":2830,"account":"SALAIRE","note":"Fruits"},{"id":"t761","date":"2026-04-03","category":"Ajustement","type":"Dépense","amount":200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t762","date":"2026-04-03","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Franck"},{"id":"t763","date":"2026-04-03","category":"Aliments","type":"Dépense","amount":700,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t748","date":"2026-04-04","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t749","date":"2026-04-04","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Le déjeuner","note":"Déjeuner"},{"id":"t750","date":"2026-04-04","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Femme","note":"Fatim"},{"id":"t751","date":"2026-04-04","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t752","date":"2026-04-04","category":"Divertissement","type":"Dépense","amount":26000,"account":"SALAIRE","subcategory":"Alcool","note":"Fatim"},{"id":"t753","date":"2026-04-04","category":"Aliments","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t754","date":"2026-04-04","category":"Vêtements","type":"Dépense","amount":45000,"account":"SALAIRE","subcategory":"Draps"},{"id":"t755","date":"2026-04-04","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t756","date":"2026-04-04","category":"Personnel","type":"Dépense","amount":8050,"account":"SALAIRE","subcategory":"Produits de beauté"},{"id":"t757","date":"2026-04-04","category":"GRUNDFOS","type":"Dépense","amount":10300,"account":"PETTY CASH","subcategory":"AUTRES","note":"EPI"},{"id":"t758","date":"2026-04-04","category":"Abonnements","type":"Dépense","amount":5379,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t744","date":"2026-04-05","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t745","date":"2026-04-05","category":"GRUNDFOS","type":"Dépense","amount":26090,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t746","date":"2026-04-05","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t747","date":"2026-04-05","category":"Invitation","type":"Dépense","amount":37000,"account":"SALAIRE","subcategory":"Femmes","note":"Fatim"},{"id":"t740","date":"2026-04-06","category":"Cadeaux","type":"Dépense","amount":250,"account":"SALAIRE","subcategory":"Femme","note":"Fatim"},{"id":"t741","date":"2026-04-06","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Dîner","note":"Ruth"},{"id":"t742","date":"2026-04-06","category":"Cadeaux","type":"Dépense","amount":1250,"account":"SALAIRE","subcategory":"Femme","note":"Internet Fatim"},{"id":"t743","date":"2026-04-06","category":"Cadeaux","type":"Dépense","amount":4040,"account":"SALAIRE","subcategory":"Femme","note":"Fatim"},{"id":"t738","date":"2026-04-07","category":"GRUNDFOS","type":"Dépense","amount":7200,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango SMI"},{"id":"t739","date":"2026-04-07","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Restaurant","note":"SMI"},{"id":"t733","date":"2026-04-08","category":"Ajustement","type":"Dépense","amount":4150,"account":"SALAIRE","note":"Smi"},{"id":"t734","date":"2026-04-08","category":"Ajustement","type":"Dépense","amount":499,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t735","date":"2026-04-08","category":"Santé","type":"Dépense","amount":1877,"account":"SIB","subcategory":"Assurance"},{"id":"t736","date":"2026-04-08","category":"Ajustement","type":"Dépense","amount":100,"account":"SALAIRE","subcategory":"Frais Bancaire","note":"Claude"},{"id":"t737","date":"2026-04-08","category":"Abonnements","type":"Dépense","amount":11955,"account":"SALAIRE","note":"Claude"},{"id":"t728","date":"2026-04-09","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t729","date":"2026-04-09","category":"Voiture","type":"Dépense","amount":1000,"account":"Dépôt LOYER","subcategory":"Parking"},{"id":"t730","date":"2026-04-09","category":"General","type":"Revenu","amount":150000,"account":"SIB","note":"Mission SMI"},{"id":"t731","date":"2026-04-09","category":"GRUNDFOS","type":"Dépense","amount":1600,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango SMI"},{"id":"t732","date":"2026-04-09","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t724","date":"2026-04-10","category":"Ajustement","type":"Revenu","amount":200,"account":"SALAIRE"},{"id":"t725","date":"2026-04-10","category":"Divertissement","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Femme","note":"Syntiche"},{"id":"t726","date":"2026-04-10","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t727","date":"2026-04-10","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t718","date":"2026-04-11","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t719","date":"2026-04-11","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t720","date":"2026-04-11","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t721","date":"2026-04-11","category":"Cadeaux","type":"Dépense","amount":5150,"account":"SALAIRE","subcategory":"Mardochee","note":"Documents ENA"},{"id":"t722","date":"2026-04-11","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t723","date":"2026-04-11","category":"Divertissement","type":"Dépense","amount":8500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t711","date":"2026-04-12","category":"Ajustement","type":"Dépense","amount":2600,"account":"SALAIRE","payee":"fatim","note":"Ajustement sortie yop"},{"id":"t712","date":"2026-04-12","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Femme","note":"Internet fatim"},{"id":"t713","date":"2026-04-12","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Femme","note":"Ami Fatim"},{"id":"t714","date":"2026-04-12","category":"Invitation","type":"Dépense","amount":4000,"account":"SALAIRE","note":"Diner Fatim"},{"id":"t715","date":"2026-04-12","category":"Divertissement","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t716","date":"2026-04-12","category":"Divertissement","type":"Dépense","amount":11500,"account":"SALAIRE","subcategory":"Alcool","note":"Lingot dor"},{"id":"t717","date":"2026-04-12","category":"Invitation","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Femmes","note":"Fatim"},{"id":"t707","date":"2026-04-13","category":"Cadeaux","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Femme","note":"Fatym"},{"id":"t708","date":"2026-04-13","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE"},{"id":"t709","date":"2026-04-13","category":"Divertissement","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t710","date":"2026-04-13","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t701","date":"2026-04-14","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Cotisations","note":"Palma"},{"id":"t702","date":"2026-04-14","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Femme","note":"Juliana"},{"id":"t703","date":"2026-04-14","category":"Ajustement","type":"Dépense","amount":750,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t704","date":"2026-04-14","category":"Aliments","type":"Dépense","amount":1150,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t705","date":"2026-04-14","category":"Divertissement","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t706","date":"2026-04-14","category":"Invitation","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Femmes","note":"Sylvia"},{"id":"t689","date":"2026-04-15","category":"Shopping","type":"Dépense","amount":1250,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t690","date":"2026-04-15","category":"Santé","type":"Dépense","amount":17675,"account":"SALAIRE","subcategory":"VG"},{"id":"t691","date":"2026-04-15","category":"Shopping","type":"Dépense","amount":3900,"account":"SALAIRE"},{"id":"t692","date":"2026-04-15","category":"Cadeaux","type":"Dépense","amount":2010,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t693","date":"2026-04-15","category":"Shopping","type":"Dépense","amount":49000,"account":"SALAIRE"},{"id":"t694","date":"2026-04-15","category":"Générales","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Payement robinets"},{"id":"t695","date":"2026-04-15","category":"Aliments","type":"Dépense","amount":13500,"account":"SALAIRE","subcategory":"Invitation","note":"Dandy"},{"id":"t696","date":"2026-04-15","category":"Ajustement","type":"Dépense","amount":6380,"account":"SALAIRE"},{"id":"t697","date":"2026-04-15","category":"Ajustement","type":"Dépense","amount":250,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t698","date":"2026-04-15","category":"Aliments","type":"Dépense","amount":505,"account":"SALAIRE","note":"Oeuf"},{"id":"t699","date":"2026-04-15","category":"Aliments","type":"Dépense","amount":13500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t700","date":"2026-04-15","category":"Cadeaux","type":"Dépense","amount":5550,"account":"SALAIRE","subcategory":"Femme","note":"Taxi juliana"},{"id":"t684","date":"2026-04-16","category":"Ajustement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t685","date":"2026-04-16","category":"Cadeaux","type":"Dépense","amount":25250,"account":"SALAIRE","subcategory":"Femme","note":"Juliana"},{"id":"t686","date":"2026-04-16","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA","note":"Mai"},{"id":"t687","date":"2026-04-16","category":"Aliments","type":"Dépense","amount":9400,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t688","date":"2026-04-16","category":"Divertissement","type":"Dépense","amount":5400,"account":"SALAIRE","subcategory":"Alcool","note":"Joel"},{"id":"t679","date":"2026-04-17","category":"Shopping","type":"Dépense","amount":3000,"account":"SALAIRE"},{"id":"t680","date":"2026-04-17","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t681","date":"2026-04-17","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t682","date":"2026-04-17","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t683","date":"2026-04-17","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Vigile"},{"id":"t674","date":"2026-04-18","category":"Générales","type":"Dépense","amount":580,"account":"SALAIRE","note":"Car buddy"},{"id":"t675","date":"2026-04-18","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t676","date":"2026-04-18","category":"Ajustement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t677","date":"2026-04-18","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t678","date":"2026-04-18","category":"Divertissement","type":"Dépense","amount":4700,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t671","date":"2026-04-19","category":"Aliments","type":"Dépense","amount":3840,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t672","date":"2026-04-19","category":"Cadeaux","type":"Dépense","amount":250000,"account":"SALAIRE","subcategory":"Ruth","note":"Phone"},{"id":"t673","date":"2026-04-19","category":"Enfants & Maman","type":"Dépense","amount":35500,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t665","date":"2026-04-20","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Femme","note":"Fatim"},{"id":"t666","date":"2026-04-20","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Femme","note":"Fatim"},{"id":"t667","date":"2026-04-20","category":"Aliments","type":"Dépense","amount":2100,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t668","date":"2026-04-20","category":"Divertissement","type":"Dépense","amount":505,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t669","date":"2026-04-20","category":"GRUNDFOS","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Internet","note":"Phone"},{"id":"t670","date":"2026-04-20","category":"Ajustement","type":"Dépense","amount":3931,"account":"SALAIRE"},{"id":"t655","date":"2026-04-21","category":"Ajustement","type":"Dépense","amount":185,"account":"PETTY CASH","subcategory":"Frais","note":"Internet"},{"id":"t656","date":"2026-04-21","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Phone"},{"id":"t657","date":"2026-04-21","category":"Ajustement","type":"Dépense","amount":2815,"account":"Revenus MAZDA","subcategory":"Frais","note":"Payement assurance"},{"id":"t658","date":"2026-04-21","category":"Voiture","type":"Dépense","amount":281555,"account":"Revenus MAZDA","subcategory":"Assurance","note":"Assurance 1 an"},{"id":"t659","date":"2026-04-21","category":"Ajustement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t660","date":"2026-04-21","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t661","date":"2026-04-21","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t662","date":"2026-04-21","category":"Santé","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Médicaments","note":"Toux"},{"id":"t663","date":"2026-04-21","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t664","date":"2026-04-21","category":"Ajustement","type":"Dépense","amount":9999,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t647","date":"2026-04-22","category":"Voiture","type":"Dépense","amount":82000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"77392 kms"},{"id":"t648","date":"2026-04-22","category":"Aliments","type":"Dépense","amount":2900,"account":"SALAIRE","subcategory":"Le déjeuner","note":"Yango"},{"id":"t649","date":"2026-04-22","category":"Aliments","type":"Dépense","amount":19190,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t650","date":"2026-04-22","category":"Prêt","type":"Dépense","amount":82290,"account":"SALAIRE","note":"Remboursement orange"},{"id":"t651","date":"2026-04-22","category":"GRUNDFOS","type":"Dépense","amount":8018,"account":"PETTY CASH","subcategory":"Eau","note":"Mars"},{"id":"t652","date":"2026-04-22","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre"},{"id":"t653","date":"2026-04-22","category":"Ajustement","type":"Dépense","amount":550,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t654","date":"2026-04-22","category":"Abonnements","type":"Dépense","amount":1962,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t644","date":"2026-04-23","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t645","date":"2026-04-23","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t646","date":"2026-04-23","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t638","date":"2026-04-24","category":"Aliments","type":"Dépense","amount":900,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t639","date":"2026-04-24","category":"Un salaire","type":"Revenu","amount":1555362,"account":"SALAIRE"},{"id":"t640","date":"2026-04-24","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t641","date":"2026-04-24","category":"Ajustement","type":"Dépense","amount":1129,"account":"SALAIRE"},{"id":"t642","date":"2026-04-24","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t643","date":"2026-04-24","category":"Invitation","type":"Dépense","amount":32500,"account":"SALAIRE","subcategory":"Femmes","note":"Ruth"},{"id":"t631","date":"2026-04-25","category":"Ajustement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t632","date":"2026-04-25","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t633","date":"2026-04-25","category":"Invitation","type":"Dépense","amount":48500,"account":"SALAIRE","subcategory":"Femmes","note":"Mariam bouake"},{"id":"t634","date":"2026-04-25","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t635","date":"2026-04-25","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Mardochee"},{"id":"t636","date":"2026-04-25","category":"Ajustement","type":"Dépense","amount":250,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t637","date":"2026-04-25","category":"Divertissement","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t623","date":"2026-04-26","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","note":"Cola"},{"id":"t624","date":"2026-04-26","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire","note":"Alloco Fatim"},{"id":"t625","date":"2026-04-26","category":"Divertissement","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Alcool","note":"Fatim lingot"},{"id":"t626","date":"2026-04-26","category":"Shopping","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Alimentation","note":"Yakro"},{"id":"t627","date":"2026-04-26","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","subcategory":"Carburant","note":"Yakro"},{"id":"t628","date":"2026-04-26","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Carburant","note":"Yakro"},{"id":"t629","date":"2026-04-26","category":"Voyage","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Péage"},{"id":"t630","date":"2026-04-26","category":"Aliments","type":"Dépense","amount":23000,"account":"SALAIRE","subcategory":"Invitation","note":"Fatim"},{"id":"t620","date":"2026-04-27","category":"Ajustement","type":"Revenu","amount":49,"account":"SALAIRE"},{"id":"t621","date":"2026-04-27","category":"Divertissement","type":"Dépense","amount":22000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t622","date":"2026-04-27","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t609","date":"2026-04-28","category":"Ajustement","type":"Dépense","amount":1834,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t610","date":"2026-04-28","category":"Securicompte","type":"Dépense","amount":912,"account":"SIB"},{"id":"t611","date":"2026-04-28","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t612","date":"2026-04-28","category":"Aliments","type":"Dépense","amount":1700,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t613","date":"2026-04-28","category":"Enfants & Maman","type":"Dépense","amount":200000,"account":"SALAIRE","subcategory":"Hemra","note":"Scolarité"},{"id":"t614","date":"2026-04-28","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","note":"Armande"},{"id":"t615","date":"2026-04-28","category":"Ajustement","type":"Dépense","amount":50,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t616","date":"2026-04-28","category":"Ajustement","type":"Dépense","amount":200,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t617","date":"2026-04-28","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Femme","note":"Jose"},{"id":"t618","date":"2026-04-28","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","subcategory":"Pourboire","note":"Seba"},{"id":"t619","date":"2026-04-28","category":"GRUNDFOS","type":"Dépense","amount":33500,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t602","date":"2026-04-29","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t603","date":"2026-04-29","category":"Divertissement","type":"Dépense","amount":25250,"account":"SALAIRE","subcategory":"Femme","note":"Lush"},{"id":"t604","date":"2026-04-29","category":"Divertissement","type":"Dépense","amount":1200,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t605","date":"2026-04-29","category":"GRUNDFOS","type":"Dépense","amount":2500,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t606","date":"2026-04-29","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t607","date":"2026-04-29","category":"Enfants & Maman","type":"Dépense","amount":55550,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t608","date":"2026-04-29","category":"Enfants & Maman","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t596","date":"2026-04-30","category":"Déménagement","type":"Dépense","amount":1000,"account":"Dépôt LOYER","subcategory":"Autres","note":"Scotch"},{"id":"t597","date":"2026-04-30","category":"Aliments","type":"Dépense","amount":3840,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t598","date":"2026-04-30","category":"Ajustement","type":"Dépense","amount":1800,"account":"SIB","subcategory":"Frais Bancaire","note":"Déménagement"},{"id":"t599","date":"2026-04-30","category":"Déménagement","type":"Dépense","amount":80000,"account":"Dépôt LOYER","subcategory":"Nettoyage"},{"id":"t600","date":"2026-04-30","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Maman"},{"id":"t601","date":"2026-04-30","category":"Enfants & Maman","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Maman"},{"id":"t587","date":"2026-05-01","category":"Déménagement","type":"Dépense","amount":48700,"account":"Dépôt LOYER","subcategory":"Autres","note":"Thermos"},{"id":"t588","date":"2026-05-01","category":"Déménagement","type":"Dépense","amount":10000,"account":"Dépôt LOYER","subcategory":"Lits"},{"id":"t589","date":"2026-05-01","category":"Ajustement","type":"Dépense","amount":14500,"account":"SIB","subcategory":"Frais Bancaire","note":"Frais package AFG"},{"id":"t590","date":"2026-05-01","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t591","date":"2026-05-01","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t592","date":"2026-05-01","category":"GRUNDFOS","type":"Dépense","amount":20000,"account":"PETTY CASH","payee":"Sodeci Ifat","note":"Payement LOGISTIC"},{"id":"t593","date":"2026-05-01","category":"Aliments","type":"Dépense","amount":7070,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t594","date":"2026-05-01","category":"Déménagement","type":"Dépense","amount":45000,"account":"Dépôt LOYER","note":"Yango"},{"id":"t595","date":"2026-05-01","category":"Déménagement","type":"Dépense","amount":30000,"account":"Dépôt LOYER","note":"Aide"},{"id":"t577","date":"2026-05-02","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Parker"},{"id":"t578","date":"2026-05-02","category":"Cadeaux","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Pourboire","note":"Déménagement"},{"id":"t579","date":"2026-05-02","category":"Déménagement","type":"Dépense","amount":1350000,"account":"Revenus MAZDA","note":"TV"},{"id":"t580","date":"2026-05-02","category":"Déménagement","type":"Dépense","amount":242000,"account":"PETTY CASH","note":"Shopping"},{"id":"t581","date":"2026-05-02","category":"Cadeaux","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t582","date":"2026-05-02","category":"Déménagement","type":"Dépense","amount":3700000,"account":"Dépôt LOYER","note":"Chez madi"},{"id":"t583","date":"2026-05-02","category":"Ajustement","type":"Dépense","amount":19000,"account":"SIB","subcategory":"Frais Bancaire","note":"Chez Madi"},{"id":"t584","date":"2026-05-02","category":"Shopping","type":"Dépense","amount":30400,"account":"SALAIRE"},{"id":"t585","date":"2026-05-02","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SALAIRE","subcategory":"Femme","note":"Marie Dominique"},{"id":"t586","date":"2026-05-02","category":"Invitation","type":"Dépense","amount":70000,"account":"SALAIRE","subcategory":"Femmes","note":"Marie Dominique"},{"id":"t568","date":"2026-05-03","category":"Divertissement","type":"Dépense","amount":50000,"account":"SALAIRE","subcategory":"La musique","note":"5e avenue"},{"id":"t569","date":"2026-05-03","category":"Invitation","type":"Dépense","amount":36000,"account":"SALAIRE","note":"Cousins"},{"id":"t570","date":"2026-05-03","category":"Déménagement","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Nettoyage"},{"id":"t571","date":"2026-05-03","category":"Déménagement","type":"Dépense","amount":25000,"account":"Revenus MAZDA","subcategory":"Autres"},{"id":"t572","date":"2026-05-03","category":"Déménagement","type":"Dépense","amount":4000,"account":"Revenus MAZDA","subcategory":"Autres"},{"id":"t573","date":"2026-05-03","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Augustin"},{"id":"t574","date":"2026-05-03","category":"Cadeaux","type":"Dépense","amount":150000,"account":"SALAIRE","subcategory":"Mardochee","note":"Scolarite Mardochee"},{"id":"t575","date":"2026-05-03","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t576","date":"2026-05-03","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t560","date":"2026-05-04","category":"Déménagement","type":"Dépense","amount":5050,"account":"Revenus MAZDA","subcategory":"Autres","note":"Aide déménagement"},{"id":"t561","date":"2026-05-04","category":"Invitation","type":"Dépense","amount":20100,"account":"SALAIRE","subcategory":"Femmes","note":"Kessy"},{"id":"t562","date":"2026-05-04","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t563","date":"2026-05-04","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Fibre optique"},{"id":"t564","date":"2026-05-04","category":"Déménagement","type":"Dépense","amount":17170,"account":"Revenus MAZDA","subcategory":"Splits","note":"Demontage clim"},{"id":"t565","date":"2026-05-04","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Dorgeo"},{"id":"t566","date":"2026-05-04","category":"Aliments","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t567","date":"2026-05-04","category":"Abonnements","type":"Dépense","amount":5438,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t551","date":"2026-05-05","category":"Divertissement","type":"Dépense","amount":800,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t552","date":"2026-05-05","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t553","date":"2026-05-05","category":"Abonnements","type":"Dépense","amount":1996,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t554","date":"2026-05-05","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t555","date":"2026-05-05","category":"Ajustement","type":"Dépense","amount":545,"account":"SIB","subcategory":"Frais","note":"Frais easy"},{"id":"t556","date":"2026-05-05","category":"GRUNDFOS","type":"Dépense","amount":80000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t557","date":"2026-05-05","category":"Prêt Orange","type":"Revenu","amount":75050,"account":"SIB"},{"id":"t558","date":"2026-05-05","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t559","date":"2026-05-05","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t544","date":"2026-05-06","category":"Ajustement","type":"Dépense","amount":69607,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t545","date":"2026-05-06","category":"GRUNDFOS","type":"Dépense","amount":10000,"account":"PETTY CASH","subcategory":"Internet","note":"Achat phone"},{"id":"t546","date":"2026-05-06","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet","note":"Moov"},{"id":"t547","date":"2026-05-06","category":"Déménagement","type":"Dépense","amount":25000,"account":"Revenus MAZDA","subcategory":"Nettoyage","note":"Fauteuil"},{"id":"t548","date":"2026-05-06","category":"Ajustement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t549","date":"2026-05-06","category":"Divertissement","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t550","date":"2026-05-06","category":"Aliments","type":"Dépense","amount":2600,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t532","date":"2026-05-07","category":"Divertissement","type":"Dépense","amount":1700,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t533","date":"2026-05-07","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t534","date":"2026-05-07","category":"Shopping","type":"Dépense","amount":6000,"account":"SALAIRE","note":"Eau"},{"id":"t535","date":"2026-05-07","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t536","date":"2026-05-07","category":"Shopping","type":"Dépense","amount":145529,"account":"Revenus MAZDA"},{"id":"t537","date":"2026-05-07","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t538","date":"2026-05-07","category":"Générales","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Péage"},{"id":"t539","date":"2026-05-07","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t540","date":"2026-05-07","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t541","date":"2026-05-07","category":"Générales","type":"Dépense","amount":1000,"account":"Revenus MAZDA","subcategory":"Péage"},{"id":"t542","date":"2026-05-07","category":"Divertissement","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t543","date":"2026-05-07","category":"Abonnements","type":"Dépense","amount":1877,"account":"SIB","subcategory":"Assurance SAF"},{"id":"t3030","date":"2026-05-07","category":"Personnel","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Hygiène personnelle","payee":"Parfum"},{"id":"t525","date":"2026-05-08","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE"},{"id":"t526","date":"2026-05-08","category":"Vente Pompe","type":"Revenu","amount":255000,"account":"SIB","note":"Moteur 7.5KW 850.000"},{"id":"t527","date":"2026-05-08","category":"Générales","type":"Dépense","amount":1400,"account":"SALAIRE","subcategory":"Yango Livraison","note":"Vente pompe"},{"id":"t528","date":"2026-05-08","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Femme","note":"Binta"},{"id":"t529","date":"2026-05-08","category":"Ajustement","type":"Dépense","amount":200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t530","date":"2026-05-08","category":"Abonnements","type":"Dépense","amount":12088,"account":"SALAIRE","subcategory":"Claude"},{"id":"t531","date":"2026-05-08","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t518","date":"2026-05-09","category":"Shopping","type":"Dépense","amount":3000,"account":"SALAIRE"},{"id":"t519","date":"2026-05-09","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t520","date":"2026-05-09","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Femme","note":"Binta"},{"id":"t521","date":"2026-05-09","category":"Invitation","type":"Dépense","amount":22000,"account":"SALAIRE","subcategory":"Femmes","note":"Binta"},{"id":"t522","date":"2026-05-09","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Femme"},{"id":"t523","date":"2026-05-09","category":"GRUNDFOS","type":"Dépense","amount":1445,"account":"PETTY CASH","subcategory":"Électricité","note":"Mutation"},{"id":"t524","date":"2026-05-09","category":"Invitation","type":"Dépense","amount":65000,"account":"SALAIRE","subcategory":"Femmes","note":"Binta"},{"id":"t514","date":"2026-05-10","category":"Générales","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Yango Livraison"},{"id":"t515","date":"2026-05-10","category":"Ajustement","type":"Dépense","amount":100,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t516","date":"2026-05-10","category":"Déménagement","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Autres","note":"Feronnier"},{"id":"t517","date":"2026-05-10","category":"Shopping","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alimentation","note":"Jus paula"},{"id":"t505","date":"2026-05-11","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t506","date":"2026-05-11","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA"},{"id":"t507","date":"2026-05-11","category":"Divertissement","type":"Dépense","amount":1200,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t508","date":"2026-05-11","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Saturnin"},{"id":"t509","date":"2026-05-11","category":"Déménagement","type":"Dépense","amount":10000,"account":"Revenus MAZDA","subcategory":"Autres","note":"Cales de portes"},{"id":"t510","date":"2026-05-11","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t511","date":"2026-05-11","category":"Déménagement","type":"Dépense","amount":4500,"account":"Revenus MAZDA","subcategory":"Autres","note":"Barre de porte"},{"id":"t512","date":"2026-05-11","category":"Ajustement","type":"Dépense","amount":8905,"account":"SALAIRE"},{"id":"t513","date":"2026-05-11","category":"Ajustement","type":"Dépense","amount":1497,"account":"SIB","subcategory":"Frais"},{"id":"t499","date":"2026-05-12","category":"Cadeaux","type":"Dépense","amount":200000,"account":"SIB","note":"Ordinateur Augustin"},{"id":"t500","date":"2026-05-12","category":"Utilitaires","type":"Dépense","amount":230000,"account":"SIB","subcategory":"Ordinateur HP"},{"id":"t501","date":"2026-05-12","category":"Vente Pompe","type":"Revenu","amount":10000,"account":"SIB","note":"5.5 KW"},{"id":"t502","date":"2026-05-12","category":"Générales","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Yango Livraison"},{"id":"t503","date":"2026-05-12","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t504","date":"2026-05-12","category":"Divertissement","type":"Dépense","amount":28000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t491","date":"2026-05-13","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t492","date":"2026-05-13","category":"Divertissement","type":"Dépense","amount":9000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t493","date":"2026-05-13","category":"General","type":"Revenu","amount":35000,"account":"SIB","subcategory":"Deladet ti"},{"id":"t494","date":"2026-05-13","category":"Shopping","type":"Dépense","amount":500,"account":"SALAIRE","note":"Lotus"},{"id":"t495","date":"2026-05-13","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t496","date":"2026-05-13","category":"Ajustement","type":"Dépense","amount":3500,"account":"SALAIRE"},{"id":"t497","date":"2026-05-13","category":"Ajustement","type":"Dépense","amount":200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t498","date":"2026-05-13","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Femme","note":"Internet Binta"},{"id":"t488","date":"2026-05-14","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE"},{"id":"t489","date":"2026-05-14","category":"Invitation","type":"Dépense","amount":21000,"account":"SALAIRE","subcategory":"Femmes","note":"Miss yasmine"},{"id":"t490","date":"2026-05-14","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t3029","date":"2026-05-14","category":"Personnel","type":"Dépense","amount":250000,"account":"Revenus MAZDA","subcategory":"Hygiène personnelle","payee":"Richkoff"},{"id":"t479","date":"2026-05-15","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t480","date":"2026-05-15","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Femme","note":"Fatim"},{"id":"t481","date":"2026-05-15","category":"Aliments","type":"Dépense","amount":11500,"account":"SALAIRE","subcategory":"Invitation"},{"id":"t482","date":"2026-05-15","category":"Abonnements","type":"Dépense","amount":6067,"account":"SALAIRE","subcategory":"Netflix"},{"id":"t483","date":"2026-05-15","category":"Ajustement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t484","date":"2026-05-15","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Elvis"},{"id":"t485","date":"2026-05-15","category":"Divertissement","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t486","date":"2026-05-15","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t487","date":"2026-05-15","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t471","date":"2026-05-16","category":"Cadeaux","type":"Dépense","amount":7220,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t472","date":"2026-05-16","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t473","date":"2026-05-16","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t474","date":"2026-05-16","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Pourboire","note":"Lunik"},{"id":"t475","date":"2026-05-16","category":"Divertissement","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t476","date":"2026-05-16","category":"Cadeaux","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t477","date":"2026-05-16","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t478","date":"2026-05-16","category":"Divertissement","type":"Dépense","amount":57000,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t465","date":"2026-05-17","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Offrande"},{"id":"t466","date":"2026-05-17","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Pourboire","note":"Musique"},{"id":"t467","date":"2026-05-17","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Femme","note":"Pamela"},{"id":"t468","date":"2026-05-17","category":"Divertissement","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t469","date":"2026-05-17","category":"Aliments","type":"Dépense","amount":43000,"account":"SALAIRE","subcategory":"Invitation","note":"Pamela Assa"},{"id":"t470","date":"2026-05-17","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","subcategory":"Femme","note":"Yasmine"},{"id":"t459","date":"2026-05-18","category":"Divertissement","type":"Dépense","amount":1200,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t460","date":"2026-05-18","category":"Invitation","type":"Dépense","amount":15500,"account":"Revenus MAZDA","subcategory":"Femmes","note":"Lety"},{"id":"t461","date":"2026-05-18","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t462","date":"2026-05-18","category":"Cadeaux","type":"Dépense","amount":202000,"account":"Revenus MAZDA","subcategory":"Cotisations","note":"Tonton martin"},{"id":"t463","date":"2026-05-18","category":"Utilitaires","type":"Dépense","amount":20000,"account":"Revenus MAZDA","subcategory":"Nettoyage"},{"id":"t464","date":"2026-05-18","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t453","date":"2026-05-19","category":"GRUNDFOS","type":"Dépense","amount":3000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t454","date":"2026-05-19","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","subcategory":"Femme","note":"Binta"},{"id":"t455","date":"2026-05-19","category":"Shopping","type":"Dépense","amount":6565,"account":"SIB","note":"Barre de porte"},{"id":"t456","date":"2026-05-19","category":"Shopping","type":"Dépense","amount":3030,"account":"SIB","note":"Ecailleur"},{"id":"t457","date":"2026-05-19","category":"General","type":"Revenu","amount":50000,"account":"SIB","subcategory":"Cadeau","note":"Ruth"},{"id":"t458","date":"2026-05-19","category":"Ajustement","type":"Dépense","amount":5,"account":"Revenus MAZDA","subcategory":"Étrange"},{"id":"t449","date":"2026-05-20","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Produits de beauté"},{"id":"t450","date":"2026-05-20","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t451","date":"2026-05-20","category":"Divertissement","type":"Dépense","amount":18500,"account":"SIB","subcategory":"Alcool"},{"id":"t452","date":"2026-05-20","category":"Aliments","type":"Dépense","amount":11500,"account":"SIB","subcategory":"Invitation"},{"id":"t445","date":"2026-05-21","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t446","date":"2026-05-21","category":"Divertissement","type":"Dépense","amount":12000,"account":"SIB","subcategory":"Alcool"},{"id":"t447","date":"2026-05-21","category":"Aliments","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Dîner"},{"id":"t448","date":"2026-05-21","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SIB","note":"Francko"},{"id":"t436","date":"2026-05-22","category":"Shopping","type":"Dépense","amount":8500,"account":"SIB","subcategory":"Alimentation"},{"id":"t437","date":"2026-05-22","category":"Voyage","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Péage"},{"id":"t438","date":"2026-05-22","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","note":"Dominique Koffi"},{"id":"t439","date":"2026-05-22","category":"Aliments","type":"Dépense","amount":600,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t440","date":"2026-05-22","category":"Shopping","type":"Dépense","amount":17000,"account":"SIB"},{"id":"t441","date":"2026-05-22","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t442","date":"2026-05-22","category":"Personnel","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Coiffure"},{"id":"t443","date":"2026-05-22","category":"Divertissement","type":"Dépense","amount":20000,"account":"SIB","subcategory":"Funérailles","note":"Martin"},{"id":"t444","date":"2026-05-22","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Pourboire","note":"DJ obseques martin"},{"id":"t431","date":"2026-05-23","category":"Cadeaux","type":"Dépense","amount":3500,"account":"SIB","note":"Nicolas / Carolle"},{"id":"t432","date":"2026-05-23","category":"Aliments","type":"Dépense","amount":21000,"account":"SIB","subcategory":"Invitation","note":"Famille"},{"id":"t433","date":"2026-05-23","category":"Divertissement","type":"Dépense","amount":55500,"account":"SIB","subcategory":"Alcool","payee":"+ Pourboire","note":"Voyage Yakro Martin"},{"id":"t434","date":"2026-05-23","category":"Invitation","type":"Dépense","amount":4500,"account":"SIB","subcategory":"Femmes"},{"id":"t435","date":"2026-05-23","category":"Cadeaux","type":"Dépense","amount":90000,"account":"SIB","note":"Funerailles Martin"},{"id":"t422","date":"2026-05-24","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t423","date":"2026-05-24","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Femme","note":"Olive"},{"id":"t424","date":"2026-05-24","category":"Divertissement","type":"Dépense","amount":5050,"account":"SIB","subcategory":"Alcool"},{"id":"t425","date":"2026-05-24","category":"Voyage","type":"Dépense","amount":72000,"account":"SIB","note":"Hotel Aho Yakro"},{"id":"t426","date":"2026-05-24","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SIB","note":"Dorgeo"},{"id":"t427","date":"2026-05-24","category":"Divertissement","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Alcool"},{"id":"t428","date":"2026-05-24","category":"Aliments","type":"Dépense","amount":9500,"account":"SIB","subcategory":"Dîner"},{"id":"t429","date":"2026-05-24","category":"Voyage","type":"Dépense","amount":20200,"account":"SIB","subcategory":"Un hôtel"},{"id":"t430","date":"2026-05-24","category":"Ajustement","type":"Dépense","amount":5,"account":"SIB","note":"Martin"},{"id":"t414","date":"2026-05-25","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t415","date":"2026-05-25","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t416","date":"2026-05-25","category":"Voyage","type":"Dépense","amount":4000,"account":"SIB","subcategory":"Péage"},{"id":"t417","date":"2026-05-25","category":"Aliments","type":"Dépense","amount":7000,"account":"SIB","subcategory":"Dîner"},{"id":"t418","date":"2026-05-25","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t419","date":"2026-05-25","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t420","date":"2026-05-25","category":"Santé","type":"Dépense","amount":5130,"account":"SIB","subcategory":"Médicaments"},{"id":"t421","date":"2026-05-25","category":"Ajustement","type":"Dépense","amount":3200,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t408","date":"2026-05-26","category":"Cadeaux","type":"Dépense","amount":6060,"account":"SIB","subcategory":"Ruth"},{"id":"t409","date":"2026-05-26","category":"Aliments","type":"Dépense","amount":5000,"account":"SIB","subcategory":"Le déjeuner"},{"id":"t410","date":"2026-05-26","category":"Santé","type":"Dépense","amount":15000,"account":"SIB","subcategory":"Médicaments"},{"id":"t411","date":"2026-05-26","category":"Un salaire","type":"Revenu","amount":3340975,"account":"SALAIRE"},{"id":"t412","date":"2026-05-26","category":"GRUNDFOS","type":"Dépense","amount":16365,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t413","date":"2026-05-26","category":"Pack Club","type":"Dépense","amount":9087,"account":"SIB"},{"id":"t407","date":"2026-05-27","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SIB","subcategory":"Ruth"},{"id":"t405","date":"2026-05-28","category":"Santé","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Médicaments"},{"id":"t406","date":"2026-05-28","category":"Invitation","type":"Dépense","amount":5000,"account":"SIB","note":"Hammal alcool"},{"id":"t400","date":"2026-05-29","category":"Aliments","type":"Dépense","amount":33000,"account":"SIB","subcategory":"Invitation","note":"Binta barry"},{"id":"t401","date":"2026-05-29","category":"Déménagement","type":"Dépense","amount":27270,"account":"SIB","note":"Installation elect fibre"},{"id":"t402","date":"2026-05-29","category":"Aliments","type":"Dépense","amount":15150,"account":"SIB"},{"id":"t403","date":"2026-05-29","category":"Aliments","type":"Dépense","amount":2500,"account":"SIB","subcategory":"Déjeuner"},{"id":"t404","date":"2026-05-29","category":"Divertissement","type":"Dépense","amount":64000,"account":"SIB","subcategory":"Alcool","note":"Olo"},{"id":"t389","date":"2026-05-30","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t390","date":"2026-05-30","category":"Aliments","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t391","date":"2026-05-30","category":"Enfants & Maman","type":"Dépense","amount":135000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t392","date":"2026-05-30","category":"Shopping","type":"Dépense","amount":33000,"account":"SIB"},{"id":"t393","date":"2026-05-30","category":"Shopping","type":"Dépense","amount":15010,"account":"SIB"},{"id":"t394","date":"2026-05-30","category":"Abonnements","type":"Dépense","amount":35350,"account":"SIB","subcategory":"Canal","note":"Hot IPTV"},{"id":"t395","date":"2026-05-30","category":"GRUNDFOS","type":"Dépense","amount":10365,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t396","date":"2026-05-30","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t397","date":"2026-05-30","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","note":"Mami"},{"id":"t398","date":"2026-05-30","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t399","date":"2026-05-30","category":"Shopping","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t388","date":"2026-05-31","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Maman","note":"Loyer maman"},{"id":"t384","date":"2026-06-01","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t385","date":"2026-06-01","category":"Invitation","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Femmes"},{"id":"t386","date":"2026-06-01","category":"GRUNDFOS","type":"Dépense","amount":35000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t387","date":"2026-06-01","category":"Aliments","type":"Dépense","amount":5500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t379","date":"2026-06-02","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Internet mobile"},{"id":"t380","date":"2026-06-02","category":"Divertissement","type":"Dépense","amount":12500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t381","date":"2026-06-02","category":"Aliments","type":"Dépense","amount":7500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t382","date":"2026-06-02","category":"Vêtements","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Accessoires"},{"id":"t383","date":"2026-06-02","category":"Cadeaux","type":"Dépense","amount":18000,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t374","date":"2026-06-03","category":"Petty Cash","type":"Revenu","amount":2482797,"account":"PETTY CASH"},{"id":"t375","date":"2026-06-03","category":"Invitation","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Femmes","note":"Lety"},{"id":"t376","date":"2026-06-03","category":"Cadeaux","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Femme","note":"Hermine"},{"id":"t377","date":"2026-06-03","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t378","date":"2026-06-03","category":"Divertissement","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t371","date":"2026-06-04","category":"Aliments","type":"Dépense","amount":8500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t372","date":"2026-06-04","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t373","date":"2026-06-04","category":"Ajustement","type":"Dépense","amount":10,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t365","date":"2026-06-05","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t366","date":"2026-06-05","category":"Aliments","type":"Dépense","amount":10200,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t367","date":"2026-06-05","category":"Cadeaux","type":"Dépense","amount":12000,"account":"SALAIRE","note":"Prince Elie - etran"},{"id":"t368","date":"2026-06-05","category":"Enfants & Maman","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t369","date":"2026-06-05","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t370","date":"2026-06-05","category":"Divertissement","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t359","date":"2026-06-06","category":"Abonnements","type":"Dépense","amount":5531,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t360","date":"2026-06-06","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SALAIRE","note":"Christelle soeur"},{"id":"t361","date":"2026-06-06","category":"Aliments","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Invitation"},{"id":"t362","date":"2026-06-06","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t363","date":"2026-06-06","category":"Shopping","type":"Dépense","amount":5000,"account":"SALAIRE"},{"id":"t364","date":"2026-06-06","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t350","date":"2026-06-07","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Femme","note":"Yasmine"},{"id":"t351","date":"2026-06-07","category":"Ajustement","type":"Dépense","amount":39173,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t352","date":"2026-06-07","category":"INVEST SGO","type":"Dépense","amount":1725,"account":"SGO","subcategory":"Frais"},{"id":"t353","date":"2026-06-07","category":"INVEST SGO","type":"Dépense","amount":100000,"account":"SALAIRE"},{"id":"t354","date":"2026-06-07","category":"Divertissement","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"BAP"},{"id":"t355","date":"2026-06-07","category":"Divertissement","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t356","date":"2026-06-07","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t357","date":"2026-06-07","category":"Ajustement","type":"Dépense","amount":2750,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t358","date":"2026-06-07","category":"GRUNDFOS","type":"Dépense","amount":20730,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t346","date":"2026-06-08","category":"GRUNDFOS","type":"Dépense","amount":1700,"account":"PETTY CASH","subcategory":"Voyage","note":"Yango"},{"id":"t347","date":"2026-06-08","category":"Abonnements","type":"Dépense","amount":12303,"account":"SALAIRE","subcategory":"Claude"},{"id":"t348","date":"2026-06-08","category":"GRUNDFOS","type":"Dépense","amount":5200,"account":"PETTY CASH","subcategory":"Restaurant","note":"Voyage ity"},{"id":"t349","date":"2026-06-08","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t340","date":"2026-06-09","category":"Shopping","type":"Dépense","amount":4000,"account":"PETTY CASH","note":"Voyage SMI"},{"id":"t341","date":"2026-06-09","category":"Abonnements","type":"Dépense","amount":2027,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t342","date":"2026-06-09","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Flora"},{"id":"t343","date":"2026-06-09","category":"Générales","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Police","note":"Voyage SMI"},{"id":"t344","date":"2026-06-09","category":"Générales","type":"Dépense","amount":1834,"account":"SALAIRE","note":"Cotisation carte"},{"id":"t345","date":"2026-06-09","category":"Abonnements","type":"Dépense","amount":1877,"account":"SIB","subcategory":"Assurance SAF"},{"id":"t325","date":"2026-06-10","category":"Ajustement","type":"Dépense","amount":9691,"account":"SALAIRE","subcategory":"Frais"},{"id":"t326","date":"2026-06-10","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t327","date":"2026-06-10","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Enfants Maï"},{"id":"t328","date":"2026-06-10","category":"GRUNDFOS","type":"Dépense","amount":4300,"account":"PETTY CASH","subcategory":"Voyage","note":"Ity retour"},{"id":"t329","date":"2026-06-10","category":"Ajustement","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t330","date":"2026-06-10","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE"},{"id":"t331","date":"2026-06-10","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Pourboire","note":"Musique"},{"id":"t332","date":"2026-06-10","category":"General","type":"Revenu","amount":150000,"account":"SIB","note":"Voyage Ity"},{"id":"t333","date":"2026-06-10","category":"GRUNDFOS","type":"Dépense","amount":31000,"account":"PETTY CASH","subcategory":"Hotel"},{"id":"t334","date":"2026-06-10","category":"GRUNDFOS","type":"Dépense","amount":2000,"account":"PETTY CASH","subcategory":"Restaurant","note":"PD"},{"id":"t335","date":"2026-06-10","category":"Cadeaux","type":"Dépense","amount":8040,"account":"SALAIRE","note":"Etran"},{"id":"t336","date":"2026-06-10","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","note":"JJK"},{"id":"t337","date":"2026-06-10","category":"Enfants & Maman","type":"Dépense","amount":50500,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t338","date":"2026-06-10","category":"Aliments","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t339","date":"2026-06-10","category":"Divertissement","type":"Dépense","amount":126000,"account":"SALAIRE","subcategory":"Alcool","note":"Nicolas"},{"id":"t319","date":"2026-06-11","category":"Santé","type":"Dépense","amount":10080,"account":"SALAIRE","subcategory":"VG"},{"id":"t320","date":"2026-06-11","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t321","date":"2026-06-11","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t322","date":"2026-06-11","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t323","date":"2026-06-11","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t324","date":"2026-06-11","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t311","date":"2026-06-12","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t312","date":"2026-06-12","category":"Vêtements","type":"Dépense","amount":90800,"account":"SALAIRE","subcategory":"Chaussures"},{"id":"t313","date":"2026-06-12","category":"Abonnements","type":"Dépense","amount":28000,"account":"SALAIRE","subcategory":"Claude"},{"id":"t314","date":"2026-06-12","category":"Enfants & Maman","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t315","date":"2026-06-12","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t316","date":"2026-06-12","category":"Ajustement","type":"Dépense","amount":400,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t317","date":"2026-06-12","category":"Enfants & Maman","type":"Dépense","amount":40400,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t318","date":"2026-06-12","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t300","date":"2026-06-13","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","note":"Olo"},{"id":"t301","date":"2026-06-13","category":"GRUNDFOS","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Restaurant","note":"Christiane"},{"id":"t302","date":"2026-06-13","category":"Cadeaux","type":"Dépense","amount":30300,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t303","date":"2026-06-13","category":"Divertissement","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t304","date":"2026-06-13","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Anniversaire","note":"Bekanty"},{"id":"t305","date":"2026-06-13","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Cotisations","note":"Martin"},{"id":"t306","date":"2026-06-13","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","note":"Bb didier"},{"id":"t307","date":"2026-06-13","category":"Invitation","type":"Dépense","amount":10500,"account":"SALAIRE","note":"Didier"},{"id":"t308","date":"2026-06-13","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t309","date":"2026-06-13","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t310","date":"2026-06-13","category":"Ajustement","type":"Dépense","amount":2410,"account":"SALAIRE"},{"id":"t294","date":"2026-06-14","category":"Shopping","type":"Dépense","amount":44000,"account":"SALAIRE"},{"id":"t295","date":"2026-06-14","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t296","date":"2026-06-14","category":"GRUNDFOS","type":"Dépense","amount":5055,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t297","date":"2026-06-14","category":"Enfants & Maman","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t298","date":"2026-06-14","category":"Aliments","type":"Dépense","amount":15300,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t299","date":"2026-06-14","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t291","date":"2026-06-15","category":"Cadeaux","type":"Dépense","amount":700,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t292","date":"2026-06-15","category":"Cadeaux","type":"Dépense","amount":50000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t293","date":"2026-06-15","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t282","date":"2026-06-17","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t283","date":"2026-06-17","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB"},{"id":"t284","date":"2026-06-17","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t285","date":"2026-06-17","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Flamand"},{"id":"t286","date":"2026-06-17","category":"Loyer","type":"Revenu","amount":1490164,"account":"Dépôt LOYER"},{"id":"t287","date":"2026-06-17","category":"Ajustement","type":"Dépense","amount":934,"account":"SIB","subcategory":"Frais"},{"id":"t288","date":"2026-06-17","category":"Aliments","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Invitation","note":"Docteur Assamoi"},{"id":"t289","date":"2026-06-17","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Pneu"},{"id":"t290","date":"2026-06-17","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Kouakou EIE"},{"id":"t272","date":"2026-06-18","category":"Shopping","type":"Dépense","amount":5000,"account":"SALAIRE"},{"id":"t273","date":"2026-06-18","category":"Cadeaux","type":"Dépense","amount":8444,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t274","date":"2026-06-18","category":"Cadeaux","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t275","date":"2026-06-18","category":"Aliments","type":"Dépense","amount":5200,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t276","date":"2026-06-18","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"BAP"},{"id":"t277","date":"2026-06-18","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t278","date":"2026-06-18","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t279","date":"2026-06-18","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t280","date":"2026-06-18","category":"Cadeaux","type":"Dépense","amount":111100,"account":"SALAIRE","note":"Flamand"},{"id":"t281","date":"2026-06-18","category":"Ajustement","type":"Dépense","amount":1100,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t264","date":"2026-06-19","category":"GRUNDFOS","type":"Dépense","amount":500,"account":"PETTY CASH","subcategory":"Péage"},{"id":"t265","date":"2026-06-19","category":"GRUNDFOS","type":"Dépense","amount":116000,"account":"PETTY CASH","subcategory":"Infraction"},{"id":"t266","date":"2026-06-19","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t267","date":"2026-06-19","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t268","date":"2026-06-19","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t269","date":"2026-06-19","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Femme","note":"Flora"},{"id":"t270","date":"2026-06-19","category":"Divertissement","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t271","date":"2026-06-19","category":"Invitation","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Femmes"},{"id":"t259","date":"2026-06-20","category":"Ajustement","type":"Dépense","amount":800,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t260","date":"2026-06-20","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t261","date":"2026-06-20","category":"Cadeaux","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t262","date":"2026-06-20","category":"Divertissement","type":"Dépense","amount":52360,"account":"SALAIRE","note":"Match"},{"id":"t263","date":"2026-06-20","category":"Shopping","type":"Dépense","amount":1600,"account":"SALAIRE"},{"id":"t254","date":"2026-06-21","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t255","date":"2026-06-21","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t256","date":"2026-06-21","category":"Aliments","type":"Dépense","amount":4200,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t257","date":"2026-06-21","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t258","date":"2026-06-21","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t251","date":"2026-06-22","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t252","date":"2026-06-22","category":"Éducation","type":"Dépense","amount":9300,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t253","date":"2026-06-22","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","note":"Augustin"},{"id":"t242","date":"2026-06-23","category":"Revenu général","type":"Revenu","amount":10000,"account":"SIB","note":"Deladet"},{"id":"t243","date":"2026-06-23","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t244","date":"2026-06-23","category":"Cadeaux","type":"Dépense","amount":800,"account":"SALAIRE","note":"Yasmine"},{"id":"t245","date":"2026-06-23","category":"Aliments","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t246","date":"2026-06-23","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Mardochee"},{"id":"t247","date":"2026-06-23","category":"GRUNDFOS","type":"Dépense","amount":3000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t248","date":"2026-06-23","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE"},{"id":"t249","date":"2026-06-23","category":"Ajustement","type":"Dépense","amount":255,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t250","date":"2026-06-23","category":"Aliments","type":"Dépense","amount":1200,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t237","date":"2026-06-24","category":"Ajustement","type":"Dépense","amount":23084,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t238","date":"2026-06-24","category":"Revenu général","type":"Revenu","amount":5000,"account":"SIB","note":"Deladet Brou"},{"id":"t239","date":"2026-06-24","category":"Aliments","type":"Dépense","amount":200,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t240","date":"2026-06-24","category":"Ajustement","type":"Dépense","amount":500,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t241","date":"2026-06-24","category":"Shopping","type":"Dépense","amount":21325,"account":"SALAIRE"},{"id":"t234","date":"2026-06-25","category":"Logement","type":"Dépense","amount":1100550,"account":"Dépôt LOYER","subcategory":"Location","note":"Juillet/Aout"},{"id":"t235","date":"2026-06-25","category":"Un salaire","type":"Revenu","amount":1629526,"account":"SALAIRE"},{"id":"t236","date":"2026-06-25","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t225","date":"2026-06-26","category":"Générales","type":"Dépense","amount":65050,"account":"SALAIRE","note":"Installation Camera"},{"id":"t226","date":"2026-06-26","category":"Vacance Nesher","type":"Dépense","amount":32000,"account":"SALAIRE","note":"Table Nesher"},{"id":"t227","date":"2026-06-26","category":"Vacance Nesher","type":"Dépense","amount":151500,"account":"SALAIRE"},{"id":"t228","date":"2026-06-26","category":"Vacance Nesher","type":"Dépense","amount":13130,"account":"SALAIRE"},{"id":"t229","date":"2026-06-26","category":"Aliments","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Invitation"},{"id":"t230","date":"2026-06-26","category":"Divertissement","type":"Dépense","amount":12000,"account":"SALAIRE","subcategory":"BAP","note":"Rue des bars"},{"id":"t231","date":"2026-06-26","category":"Divertissement","type":"Dépense","amount":39000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t232","date":"2026-06-26","category":"Aliments","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t233","date":"2026-06-26","category":"Vacance Nesher","type":"Dépense","amount":25000,"account":"SALAIRE","note":"Matelas"},{"id":"t212","date":"2026-06-27","category":"Ajustement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Frais Bancaire"},{"id":"t213","date":"2026-06-27","category":"Vacance Nesher","type":"Dépense","amount":250000,"account":"SALAIRE","note":"Lit"},{"id":"t214","date":"2026-06-27","category":"Ajustement","type":"Dépense","amount":2700,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t215","date":"2026-06-27","category":"GRUNDFOS","type":"Dépense","amount":45000,"account":"PETTY CASH","subcategory":"Divertissement","note":"Solomon"},{"id":"t216","date":"2026-06-27","category":"Divertissement","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Femme","note":"Julie"},{"id":"t217","date":"2026-06-27","category":"Cadeaux","type":"Dépense","amount":1500,"account":"PETTY CASH","subcategory":"Pourboire","note":"Solomon"},{"id":"t218","date":"2026-06-27","category":"GRUNDFOS","type":"Dépense","amount":38500,"account":"PETTY CASH","note":"Invitation solomon"},{"id":"t219","date":"2026-06-27","category":"Shopping","type":"Dépense","amount":12000,"account":"SALAIRE"},{"id":"t220","date":"2026-06-27","category":"Vacance Nesher","type":"Dépense","amount":2000,"account":"SALAIRE","note":"Livraison tablette"},{"id":"t221","date":"2026-06-27","category":"Vacance Nesher","type":"Dépense","amount":80800,"account":"SALAIRE","note":"Tablette"},{"id":"t222","date":"2026-06-27","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t223","date":"2026-06-27","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","note":"Monteur lit"},{"id":"t224","date":"2026-06-27","category":"Vacance Nesher","type":"Dépense","amount":8900,"account":"SALAIRE","note":"Lits"},{"id":"t208","date":"2026-06-28","category":"Aliments","type":"Dépense","amount":10500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t209","date":"2026-06-28","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t210","date":"2026-06-28","category":"Vacance Nesher","type":"Dépense","amount":1500,"account":"SALAIRE","note":"Savon"},{"id":"t211","date":"2026-06-28","category":"Divertissement","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t198","date":"2026-06-29","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t199","date":"2026-06-29","category":"Vacance Nesher","type":"Dépense","amount":2000,"account":"SALAIRE","note":"Education"},{"id":"t200","date":"2026-06-29","category":"Aliments","type":"Dépense","amount":3600,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t201","date":"2026-06-29","category":"Vacance Nesher","type":"Dépense","amount":4500,"account":"SALAIRE","note":"Shopping"},{"id":"t202","date":"2026-06-29","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","note":"Bekanty"},{"id":"t203","date":"2026-06-29","category":"Shopping","type":"Dépense","amount":18175,"account":"SALAIRE"},{"id":"t204","date":"2026-06-29","category":"Aliments","type":"Dépense","amount":600,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t205","date":"2026-06-29","category":"Ajustement","type":"Revenu","amount":1000,"account":"SIB","note":"Wave"},{"id":"t206","date":"2026-06-29","category":"Ajustement","type":"Dépense","amount":300,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t207","date":"2026-06-29","category":"Shopping","type":"Dépense","amount":22000,"account":"SALAIRE","note":"Gaz"},{"id":"t188","date":"2026-06-30","category":"Shopping","type":"Dépense","amount":1300,"account":"SALAIRE"},{"id":"t189","date":"2026-06-30","category":"Enfants & Maman","type":"Dépense","amount":60000,"account":"SALAIRE","subcategory":"Nesher"},{"id":"t190","date":"2026-06-30","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","note":"Mohamed"},{"id":"t191","date":"2026-06-30","category":"Logement","type":"Dépense","amount":101000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t192","date":"2026-06-30","category":"GRUNDFOS","type":"Dépense","amount":100000,"account":"PETTY CASH","subcategory":"Électricité"},{"id":"t193","date":"2026-06-30","category":"Enfants & Maman","type":"Dépense","amount":70700,"account":"SALAIRE","subcategory":"Maman"},{"id":"t194","date":"2026-06-30","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Femme"},{"id":"t195","date":"2026-06-30","category":"Divertissement","type":"Dépense","amount":23000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t196","date":"2026-06-30","category":"Ajustement","type":"Dépense","amount":1200,"account":"SIB","subcategory":"Frais"},{"id":"t197","date":"2026-06-30","category":"Cadeaux","type":"Dépense","amount":120000,"account":"SALAIRE","note":"Flamand"},{"id":"t176","date":"2026-07-01","category":"Ajustement","type":"Dépense","amount":1200,"account":"SIB","subcategory":"Frais"},{"id":"t177","date":"2026-07-01","category":"Shopping","type":"Dépense","amount":60000,"account":"SALAIRE"},{"id":"t178","date":"2026-07-01","category":"Aliments","type":"Dépense","amount":3400,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t179","date":"2026-07-01","category":"Shopping","type":"Dépense","amount":20200,"account":"SALAIRE"},{"id":"t180","date":"2026-07-01","category":"Ajustement","type":"Dépense","amount":22500,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t181","date":"2026-07-01","category":"Divertissement","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t182","date":"2026-07-01","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t183","date":"2026-07-01","category":"Enfants & Maman","type":"Dépense","amount":27270,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t184","date":"2026-07-01","category":"Santé","type":"Dépense","amount":17565,"account":"SALAIRE","subcategory":"Médicaments"},{"id":"t185","date":"2026-07-01","category":"Enfants & Maman","type":"Dépense","amount":19300,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t186","date":"2026-07-01","category":"Divertissement","type":"Dépense","amount":25250,"account":"SALAIRE","subcategory":"Femme","note":"Lilly Tinder"},{"id":"t187","date":"2026-07-01","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t3028","date":"2026-07-01","category":"Personnel","type":"Dépense","amount":15960,"account":"SALAIRE","subcategory":"Hygiène personnelle"},{"id":"t172","date":"2026-07-02","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t173","date":"2026-07-02","category":"Shopping","type":"Dépense","amount":12120,"account":"SALAIRE"},{"id":"t174","date":"2026-07-02","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t175","date":"2026-07-02","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t171","date":"2026-07-03","category":"Vente Pompe","type":"Revenu","amount":260000,"account":"PUMP"},{"id":"t169","date":"2026-07-04","category":"Vacance Nesher","type":"Dépense","amount":70700,"account":"SALAIRE","note":"Atelier de vacances"},{"id":"t170","date":"2026-07-04","category":"Vacance Nesher","type":"Dépense","amount":1100,"account":"SALAIRE"},{"id":"t163","date":"2026-07-05","category":"Divertissement","type":"Dépense","amount":40000,"account":"SALAIRE","subcategory":"Alcool","note":"Kehrann"},{"id":"t164","date":"2026-07-05","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","note":"Leati"},{"id":"t165","date":"2026-07-05","category":"Vêtements","type":"Dépense","amount":216642,"account":"SALAIRE","subcategory":"Chaussures","note":"Girotti"},{"id":"t166","date":"2026-07-05","category":"Ajustement","type":"Dépense","amount":43,"account":"SALAIRE"},{"id":"t167","date":"2026-07-05","category":"Générales","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Police"},{"id":"t168","date":"2026-07-05","category":"Shopping","type":"Dépense","amount":8000,"account":"SALAIRE"},{"id":"t158","date":"2026-07-06","category":"Cadeaux","type":"Dépense","amount":9755,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t159","date":"2026-07-06","category":"Aliments","type":"Dépense","amount":500,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t160","date":"2026-07-06","category":"Aliments","type":"Dépense","amount":1900,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t161","date":"2026-07-06","category":"Voiture","type":"Dépense","amount":3020,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t162","date":"2026-07-06","category":"Aliments","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t156","date":"2026-07-07","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t157","date":"2026-07-07","category":"Abonnements","type":"Dépense","amount":2043,"account":"SALAIRE","subcategory":"Spotify"},{"id":"t149","date":"2026-07-08","category":"Divertissement","type":"Dépense","amount":4000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t150","date":"2026-07-08","category":"Personnel","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t151","date":"2026-07-08","category":"Shopping","type":"Dépense","amount":21593,"account":"SALAIRE"},{"id":"t152","date":"2026-07-08","category":"Shopping","type":"Dépense","amount":21000,"account":"SALAIRE"},{"id":"t153","date":"2026-07-08","category":"Aliments","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t154","date":"2026-07-08","category":"Achat Terrain Port Bouet","type":"Dépense","amount":10100,"account":"SALAIRE","note":"Diakite"},{"id":"t155","date":"2026-07-08","category":"Ajustement","type":"Dépense","amount":600,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t140","date":"2026-07-09","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t141","date":"2026-07-09","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t142","date":"2026-07-09","category":"Vacance Nesher","type":"Dépense","amount":4500,"account":"SALAIRE","note":"Livre"},{"id":"t143","date":"2026-07-09","category":"Cadeaux","type":"Dépense","amount":4747,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t144","date":"2026-07-09","category":"Abonnements","type":"Dépense","amount":12439,"account":"SALAIRE","subcategory":"Claude"},{"id":"t145","date":"2026-07-09","category":"Abonnements","type":"Dépense","amount":5591,"account":"SALAIRE","subcategory":"Money Coach"},{"id":"t146","date":"2026-07-09","category":"Shopping","type":"Dépense","amount":14000,"account":"SALAIRE","note":"Lapin"},{"id":"t147","date":"2026-07-09","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"BAP","note":"Tin"},{"id":"t148","date":"2026-07-09","category":"Divertissement","type":"Dépense","amount":6000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t132","date":"2026-07-10","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t133","date":"2026-07-10","category":"Shopping","type":"Dépense","amount":21500,"account":"SALAIRE"},{"id":"t134","date":"2026-07-10","category":"Générales","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Police"},{"id":"t135","date":"2026-07-10","category":"Voiture","type":"Dépense","amount":108000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Plaquettes"},{"id":"t136","date":"2026-07-10","category":"Ajustement","type":"Dépense","amount":2000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t137","date":"2026-07-10","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t138","date":"2026-07-10","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t139","date":"2026-07-10","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t120","date":"2026-07-11","category":"Ajustement","type":"Dépense","amount":9301,"account":"SALAIRE"},{"id":"t121","date":"2026-07-11","category":"Aliments","type":"Dépense","amount":10500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t122","date":"2026-07-11","category":"Cadeaux","type":"Dépense","amount":2020,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t123","date":"2026-07-11","category":"Cadeaux","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Cotisations","note":"Ousmane"},{"id":"t124","date":"2026-07-11","category":"Cadeaux","type":"Dépense","amount":2020,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t125","date":"2026-07-11","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE"},{"id":"t126","date":"2026-07-11","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","subcategory":"Femme","note":"Chancelle ring"},{"id":"t127","date":"2026-07-11","category":"Ajustement","type":"Dépense","amount":3000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t128","date":"2026-07-11","category":"Divertissement","type":"Dépense","amount":155070,"account":"SALAIRE","subcategory":"Alcool","note":"Mariage Ousmane"},{"id":"t129","date":"2026-07-11","category":"Cadeaux","type":"Dépense","amount":13000,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t130","date":"2026-07-11","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t131","date":"2026-07-11","category":"Cadeaux","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t115","date":"2026-07-12","category":"GRUNDFOS","type":"Dépense","amount":25000,"account":"PETTY CASH","subcategory":"Internet"},{"id":"t116","date":"2026-07-12","category":"Générales","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Péage"},{"id":"t117","date":"2026-07-12","category":"Cadeaux","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t118","date":"2026-07-12","category":"Cadeaux","type":"Dépense","amount":15015,"account":"SALAIRE","note":"Dorgeo"},{"id":"t119","date":"2026-07-12","category":"Invitation","type":"Dépense","amount":34500,"account":"SALAIRE","subcategory":"Femmes","note":"Christiane"},{"id":"t113","date":"2026-07-13","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t114","date":"2026-07-13","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","note":"Franck"},{"id":"t110","date":"2026-07-14","category":"Shopping","type":"Dépense","amount":50500,"account":"SALAIRE"},{"id":"t111","date":"2026-07-14","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Pneu"},{"id":"t112","date":"2026-07-14","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Déjeuner"},{"id":"t100","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t101","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":4450,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t102","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t103","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":10100,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t104","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE"},{"id":"t105","date":"2026-07-15","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t106","date":"2026-07-15","category":"Divertissement","type":"Dépense","amount":13500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t107","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":15150,"account":"SALAIRE","subcategory":"Cotisations","note":"Famille"},{"id":"t108","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":67670,"account":"SALAIRE","subcategory":"Cotisations","note":"Famille"},{"id":"t109","date":"2026-07-15","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Femme","note":"Lety"},{"id":"t94","date":"2026-07-16","category":"Divertissement","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t95","date":"2026-07-16","category":"Aliments","type":"Dépense","amount":6060,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t96","date":"2026-07-16","category":"Vacance Nesher","type":"Dépense","amount":8080,"account":"SALAIRE","note":"Tenue traditionnel"},{"id":"t97","date":"2026-07-16","category":"Cadeaux","type":"Dépense","amount":3030,"account":"SALAIRE","subcategory":"Femme","note":"Dab"},{"id":"t98","date":"2026-07-16","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Mardochee"},{"id":"t99","date":"2026-07-16","category":"Éducation","type":"Dépense","amount":8900,"account":"SALAIRE","subcategory":"Livres","note":"Alex fezan"},{"id":"t88","date":"2026-07-17","category":"Cadeaux","type":"Dépense","amount":5000,"account":"SALAIRE","note":"Grace Niamke"},{"id":"t89","date":"2026-07-17","category":"Vêtements","type":"Dépense","amount":82000,"account":"SALAIRE","subcategory":"Chaussures"},{"id":"t90","date":"2026-07-17","category":"Cadeaux","type":"Dépense","amount":30000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t91","date":"2026-07-17","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"La musique"},{"id":"t92","date":"2026-07-17","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t93","date":"2026-07-17","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE"},{"id":"t83","date":"2026-07-18","category":"Ajustement","type":"Dépense","amount":1000,"account":"SIB","subcategory":"Frais Bancaire"},{"id":"t84","date":"2026-07-18","category":"Cadeaux","type":"Dépense","amount":50500,"account":"SALAIRE","note":"Abdoul"},{"id":"t85","date":"2026-07-18","category":"Invitation","type":"Dépense","amount":28000,"account":"SALAIRE","subcategory":"Femmes","note":"Ruth"},{"id":"t86","date":"2026-07-18","category":"Cadeaux","type":"Dépense","amount":700,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t87","date":"2026-07-18","category":"Cadeaux","type":"Dépense","amount":36000,"account":"SALAIRE","note":"Niamke Grace"},{"id":"t80","date":"2026-07-19","category":"Vacance Nesher","type":"Dépense","amount":17000,"account":"SALAIRE","note":"Sante"},{"id":"t81","date":"2026-07-19","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t82","date":"2026-07-19","category":"Divertissement","type":"Dépense","amount":5540,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t72","date":"2026-07-20","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Femme","note":"Christiane"},{"id":"t73","date":"2026-07-20","category":"Invitation","type":"Dépense","amount":29000,"account":"SALAIRE","subcategory":"Femmes","note":"Christiane"},{"id":"t74","date":"2026-07-20","category":"Cadeaux","type":"Dépense","amount":2020,"account":"SALAIRE"},{"id":"t75","date":"2026-07-20","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t76","date":"2026-07-20","category":"Aliments","type":"Dépense","amount":13500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t77","date":"2026-07-20","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t78","date":"2026-07-20","category":"Shopping","type":"Dépense","amount":11400,"account":"SALAIRE"},{"id":"t79","date":"2026-07-20","category":"Divertissement","type":"Dépense","amount":21000,"account":"SALAIRE","note":"Alcool"},{"id":"t66","date":"2026-07-21","category":"GRUNDFOS","type":"Dépense","amount":31060,"account":"PETTY CASH","subcategory":"Hinoter"},{"id":"t67","date":"2026-07-21","category":"Achat Terrain Port Bouet","type":"Dépense","amount":412000,"account":"Revenus MAZDA"},{"id":"t68","date":"2026-07-21","category":"Vêtements","type":"Dépense","amount":41000,"account":"SALAIRE","subcategory":"Chaussures"},{"id":"t69","date":"2026-07-21","category":"Générales","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Péage"},{"id":"t70","date":"2026-07-21","category":"Aliments","type":"Dépense","amount":3500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t71","date":"2026-07-21","category":"GRUNDFOS","type":"Dépense","amount":5000,"account":"PETTY CASH","subcategory":"Internet","note":"Mobile"},{"id":"t62","date":"2026-07-22","category":"Divertissement","type":"Dépense","amount":25000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t63","date":"2026-07-22","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t64","date":"2026-07-22","category":"Divertissement","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t65","date":"2026-07-22","category":"Ajustement","type":"Dépense","amount":23303,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t58","date":"2026-07-23","category":"Aliments","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t59","date":"2026-07-23","category":"Abonnements","type":"Dépense","amount":12621,"account":"PETTY CASH","subcategory":"Claude"},{"id":"t60","date":"2026-07-23","category":"Plan Éducation","type":"Dépense","amount":32457,"account":"SIB","subcategory":"PEL"},{"id":"t61","date":"2026-07-23","category":"Âge D'or Retraite","type":"Dépense","amount":40000,"account":"SIB"},{"id":"t51","date":"2026-07-24","category":"Vacance Nesher","type":"Dépense","amount":1010,"account":"SALAIRE"},{"id":"t52","date":"2026-07-24","category":"Shopping","type":"Dépense","amount":58100,"account":"SALAIRE","subcategory":"Alimentation"},{"id":"t53","date":"2026-07-24","category":"Cadeaux","type":"Dépense","amount":20200,"account":"SALAIRE","subcategory":"Cotisations","note":"BAO"},{"id":"t54","date":"2026-07-24","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t55","date":"2026-07-24","category":"Un salaire","type":"Revenu","amount":1629526,"account":"SALAIRE"},{"id":"t56","date":"2026-07-24","category":"Aliments","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t57","date":"2026-07-24","category":"Aliments","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t45","date":"2026-07-25","category":"GRUNDFOS","type":"Dépense","amount":137595,"account":"PETTY CASH","subcategory":"Divertissement","note":"Dot Jo"},{"id":"t46","date":"2026-07-25","category":"GRUNDFOS","type":"Dépense","amount":62000,"account":"PETTY CASH","subcategory":"Divertissement","note":"Dot Jo"},{"id":"t47","date":"2026-07-25","category":"GRUNDFOS","type":"Dépense","amount":44000,"account":"PETTY CASH","subcategory":"Divertissement","note":"Dot Jo"},{"id":"t48","date":"2026-07-25","category":"Aliments","type":"Dépense","amount":6800,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t49","date":"2026-07-25","category":"Personnel","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Coiffure"},{"id":"t50","date":"2026-07-25","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t39","date":"2026-07-26","category":"Invitation","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Femmes"},{"id":"t40","date":"2026-07-26","category":"Voiture","type":"Dépense","amount":30000,"account":"Revenus MAZDA","subcategory":"Entretien","note":"Huile"},{"id":"t41","date":"2026-07-26","category":"Cadeaux","type":"Dépense","amount":5050,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t42","date":"2026-07-26","category":"Enfants & Maman","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t43","date":"2026-07-26","category":"GRUNDFOS","type":"Dépense","amount":40000,"account":"PETTY CASH","subcategory":"Carburant"},{"id":"t44","date":"2026-07-26","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t35","date":"2026-07-27","category":"Aliments","type":"Dépense","amount":3000,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t36","date":"2026-07-27","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Entretien"},{"id":"t37","date":"2026-07-27","category":"Invitation","type":"Dépense","amount":11000,"account":"SALAIRE","subcategory":"Femmes","note":"Alvy"},{"id":"t38","date":"2026-07-27","category":"Divertissement","type":"Dépense","amount":20000,"account":"SALAIRE","subcategory":"BAP","note":"Dab"},{"id":"t33","date":"2026-07-28","category":"Aliments","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t34","date":"2026-07-28","category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"account":"Revenus MAZDA"},{"id":"t23","date":"2026-07-29","category":"Shopping","type":"Dépense","amount":5500,"account":"SALAIRE"},{"id":"t24","date":"2026-07-29","category":"Aliments","type":"Dépense","amount":1500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t25","date":"2026-07-29","category":"Ajustement","type":"Dépense","amount":22292,"account":"SALAIRE","subcategory":"Étrange"},{"id":"t26","date":"2026-07-29","category":"Santé","type":"Dépense","amount":8000,"account":"SALAIRE","subcategory":"Médicaments","note":"Palu"},{"id":"t27","date":"2026-07-29","category":"Pack Club","type":"Dépense","amount":10000,"account":"SALAIRE"},{"id":"t28","date":"2026-07-29","category":"Aliments","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t29","date":"2026-07-29","category":"Logement","type":"Dépense","amount":100000,"account":"Dépôt LOYER","subcategory":"Location","note":"Armande"},{"id":"t30","date":"2026-07-29","category":"Enfants & Maman","type":"Dépense","amount":16000,"account":"SALAIRE","subcategory":"Hemra"},{"id":"t31","date":"2026-07-29","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Pourboire","note":"Mon phone"},{"id":"t32","date":"2026-07-29","category":"Cadeaux","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Femme","note":"Christianne"},{"id":"t3027","date":"2026-07-29","category":"Petty Cash","type":"Revenu","amount":2482797,"account":"PETTY CASH","subcategory":"Ajustement Petty Cash"},{"id":"t21","date":"2026-07-30","category":"Cadeaux","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Ruth"},{"id":"t22","date":"2026-07-30","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t14","date":"2026-07-31","category":"Aliments","type":"Dépense","amount":3635,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t15","date":"2026-07-31","category":"Ajustement","type":"Revenu","amount":900,"account":"SALAIRE"},{"id":"t16","date":"2026-07-31","category":"Vacance Nesher","type":"Dépense","amount":60000,"account":"SALAIRE","note":"Nounou"},{"id":"t17","date":"2026-07-31","category":"Ajustement","type":"Dépense","amount":14500,"account":"PETTY CASH","subcategory":"Frais Bancaire","note":"AFG"},{"id":"t18","date":"2026-07-31","category":"Aliments","type":"Dépense","amount":9500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t19","date":"2026-07-31","category":"Prêt","type":"Dépense","amount":80333,"account":"SALAIRE","note":"Tik tak"},{"id":"t20","date":"2026-07-31","category":"Enfants & Maman","type":"Dépense","amount":70700,"account":"SALAIRE","subcategory":"Maman"},{"id":"t7","date":"2026-08-01","category":"Voiture","type":"Dépense","amount":2000,"account":"Revenus MAZDA","subcategory":"Lavage"},{"id":"t8","date":"2026-08-01","category":"Aliments","type":"Dépense","amount":4500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t9","date":"2026-08-01","category":"Cadeaux","type":"Dépense","amount":405,"account":"SALAIRE","subcategory":"Femme","note":"Alvy"},{"id":"t10","date":"2026-08-01","category":"Divertissement","type":"Dépense","amount":7000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t11","date":"2026-08-01","category":"Cadeaux","type":"Dépense","amount":2000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t12","date":"2026-08-01","category":"Aliments","type":"Dépense","amount":17000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t13","date":"2026-08-01","category":"Générales","type":"Dépense","amount":3999,"account":"SALAIRE","note":"Recettes cuisine"},{"id":"t1","date":"2026-08-02","category":"Cadeaux","type":"Dépense","amount":1000,"account":"SALAIRE","subcategory":"Pourboire"},{"id":"t2","date":"2026-08-02","category":"Aliments","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"Dîner"},{"id":"t3","date":"2026-08-02","category":"Aliments","type":"Dépense","amount":2500,"account":"SALAIRE","subcategory":"Le déjeuner"},{"id":"t4","date":"2026-08-02","category":"Divertissement","type":"Dépense","amount":10000,"account":"SALAIRE","subcategory":"Alcool"},{"id":"t5","date":"2026-08-02","category":"Divertissement","type":"Dépense","amount":15000,"account":"SALAIRE","subcategory":"La musique","note":"Dydime"},{"id":"t6","date":"2026-08-02","category":"Divertissement","type":"Dépense","amount":5000,"account":"SALAIRE","subcategory":"Alcool"}];
const defaultCategoryGroups: Record<string, Group> = {"Logement":"Nécessaire","Aliments":"Nécessaire","Santé":"Nécessaire","Utilitaires":"Nécessaire","Voiture":"Nécessaire","Transport":"Nécessaire","Enfants & Maman":"Nécessaire","Déménagement":"Nécessaire","Securicompte":"Nécessaire","Des sports":"Non-productif","Dette":"Productif","Épargne":"Productif","Âge D'or Retraite":"Productif","Plan Éducation":"Productif","Achat Mazda":"Productif","Payement Maison Bingerville":"Productif","Achat Terrain Port Bouet":"Productif","GRUNDFOS":"Productif","INVEST SGO":"Productif","Création Entreprise ECO PUMP AFRIK":"Productif","Prêt":"Productif","Formation":"Productif","Éducation":"Productif","Cadeaux":"Non-productif","Divertissement":"Non-productif","Invitation":"Non-productif","Shopping":"Non-productif","Personnel":"Non-productif","Vêtements":"Non-productif","Générales":"Non-productif","Vacance Nesher":"Non-productif","Voyage":"Non-productif","Abonnements":"Non-productif","Pack Club":"Non-productif","Ajustement":"Non classifié","General":"Non classifié"};

const defaultCategoryScope: Record<string, Scope> = {
  "GRUNDFOS": "Business", "INVEST SGO": "Business", "Création Entreprise ECO PUMP AFRIK": "Business",
  "Vente Pompe": "Business", "ECO PUMP": "Business",
};

// Hiérarchie catégorie → sous-catégories, telle que définie dans MoneyCoach.
// Clés séparées par type car certains noms (ex: "Ajustement") existent des deux côtés
// avec des sous-catégories différentes.
const depSubcategories: Record<string, string[]> = {
  "Invitation": ["Femmes", "Triade"],
  "Logement": ["Location"],
  "Personnel": ["Coiffure", "Produits de beauté", "Hygiène personnelle"],
  "Plan Éducation": ["PEL"],
  "Santé": ["Médicaments", "VG", "Dentaire", "Vision", "Hôpital", "Assurance"],
  "Shopping": ["Draps", "Alimentation"],
  "Transport": ["Péage", "Souterrain", "Autobus", "Taxi", "Train"],
  "Utilitaires": ["Nettoyage", "Électricité", "Eau", "Gaz", "Chauffage", "Des ordures", "l'Internet", "Téléphones", "la télé", "Ordinateur HP"],
  "Voiture": ["Lavage", "Peinture-Retouche", "Installation GPS", "Carburant", "Visite Technique", "Entretien", "La lessive", "Parking", "Assurance"],
  "Voyage": ["Pourboire", "Péage", "Divertissement", "Aliment", "Shopping", "Vol", "Un hôtel", "Location de voiture"],
  "Vêtements": ["Chemises", "Chaussures", "Un pantalon", "Tops", "Des sacs", "Accessoires", "Draps"],
  "Éducation": ["Nesher", "Cours", "Livres", "Fournitures scolaires", "Prêt étudiant"],
  "Dette": ["PEL"],
  "Divertissement": ["Femme", "Residence", "Alcool", "BAP", "Films", "Boisson", "Anniversaire", "La musique", "Jeux", "Performance", "Fête", "Funérailles"],
  "Déménagement": ["Lits", "Installation Clim Chauffe Eau", "Nettoyage", "Micro-ondes Four", "Chauffe-eau", "Réfrigérateur", "Gaziniere", "Remplacement Gaziniere", "Deco & Senteur", "Electricien", "Étagère Cuisine", "Splits", "Autres"],
  "Aliments": ["Déjeuner", "Invitation", "Le déjeuner", "Dîner", "Les courses", "Dîner à l'extérieur"],
  "Cadeaux": ["Mardochee", "Olokpacha", "Femme", "Ruth", "Enfants Nesher", "Ndjore", "Pourboire", "Adrien", "Cotisations", "Metty", "Obed", "Anniversaire", "Noël", "Juste pour le fun", "Dot Jo", "MJO"],
  "Création Entreprise ECO PUMP AFRIK": ["FNE", "Documents", "Dédouanement", "Timbre", "Application De Gestion", "Cachet", "Boîte Postale"],
  "Des sports": ["Gym", "Équipement", "Piscine"],
  "Enfants & Maman": ["Maman", "Nesher", "Hemra"],
  "Formation": ["Finelo Invest", "Emergent", "Piano & Guitare"],
  "GRUNDFOS": ["Appel", "Restaurant", "Location Prado", "Eau", "Enjoy", "Impression", "Impression Allowance", "Internet", "Ajustement Petty Cash", "Voyage", "FedEx", "Hotel", "Électricité", "Péage", "iPhone 16 Pro", "Divertissement", "Cachet", "Carburant", "Infraction", "AUTRES", "Hinoter"],
  "Générales": ["Cachet Grundfos", "Impression", "Abonnement IScanner", "Visite Maison", "Cachet OMÉGA", "Certificat De Perte SIB", "Carte Money Fusion", "Police", "Réparation iPhone 13", "Souris Sans Fil", "Carte Djamo", "Vol Djamo", "Badoo", "Yango Livraison", "Péage", "Réparation Robinet", "Création Entreprise ECO PUMP", "Yango Transport Pompe", "Payement Pour Terrain Port Bouet", "Livraison Pompe"],
  "INVEST SGO": ["DJAMO", "Daba Finance", "Frais", "NSIA"],
  "Abonnements": ["Gamma App", "Daba Finance", "Emergent", "Netflix", "Financial AFRIK", "Assurance SAF", "Spotify", "Money Coach", "Richbourse", "Tinder", "Chat GPT", "Onfray", "Google Espace", "Canal", "Jeu D'affaire", "Claude"],
  "Ajustement": ["Lahou", "Frais", "Frais Bancaire", "Étrange"],
};

const revSubcategories: Record<string, string[]> = {
  "Petty Cash": ["Ajustement Petty Cash"],
  "Revenu général": ["Solde 1er juillet 2024", "Rappel PEL", "Solde 1er Jan 2025"],
  "Allocation": ["Avoir Azalai"],
  "General": ["Commission Vente Peugeot 307", "Commission Vente Nissan", "Deladet ti", "Cadeau"],
};

const defaultRules: CategorizationRule[] = [
  { id: "r1", keyword: "grundfos", group: "Productif" },
  { id: "r2", keyword: "pompe", group: "Productif" },
  { id: "r3", keyword: "cadeau", group: "Non-productif" },
  { id: "r4", keyword: "invitation", group: "Non-productif" },
  { id: "r5", keyword: "divertissement", group: "Non-productif" },
];

const netWorthRaw: [string, number][] = [
  ["2024_6", 0], ["2024_7", 3898741], ["2024_8", 10993608], ["2024_9", 4301045],
  ["2024_10", 3842272], ["2024_11", 2212794], ["2024_12", 6237266],
  ["2025_1", 6047015], ["2025_2", 8000440], ["2025_3", 5113318], ["2025_4", 6796780],
  ["2025_5", 7243083], ["2025_6", 9846075], ["2025_7", 7092614], ["2025_8", 7951891],
  ["2025_9", 5592325], ["2025_10", 7975751], ["2025_11", 7980393], ["2025_12", 10457325],
  ["2026_1", 14173506], ["2026_2", 12972957], ["2026_3", 12095271], ["2026_4", 12378388],
  ["2026_5", 6722858], ["2026_6", 10048723], ["2026_7", 11992195], ["2026_8", 11988196],
];

// Solde réel d'un compte = solde de départ + mouvements des transactions qui lui sont liées.
// Depuis l'import complet de l'historique (3046 transactions, chacune avec son vrai compte),
// le solde de départ de chaque compte est à 0 : tout est désormais reconstitué depuis les
// transactions elles-mêmes, plutôt que depuis un solde figé au moment de la mise en place.
function accountBalance(acc: Account, transactions: Transaction[]): number {
  let net = 0;
  transactions.forEach((t) => {
    if (t.account !== acc.name) return;
    net += t.type === "Revenu" ? t.amount : -t.amount;
  });
  return acc.openingBalance + net;
}
function totalAccountsBalance(accounts: Account[], transactions: Transaction[]): number {
  return accounts.reduce((a, acc) => a + accountBalance(acc, transactions), 0);
}

// Remplace (ou ajoute) le dernier point par le total réel des comptes — la valeur nette
// affichée reflète alors l'état actuel des comptes plutôt qu'un relevé historique figé.
function liveNetWorthSeries(accounts: Account[], transactions: Transaction[]): [string, number][] {
  const total = totalAccountsBalance(accounts, transactions);
  const curKey = dateToMonthKey(todayISO());
  const lastKey = netWorthRaw[netWorthRaw.length - 1][0];
  if (lastKey === curKey) return [...netWorthRaw.slice(0, -1), [lastKey, total]];
  return [...netWorthRaw, [curKey, total]];
}

const seedLoans: Loan[] = [
  { id: "l1", person: "Ami (février 2026)", amount: 500000, dateGiven: "2026_2", status: "En attente", notes: "Prêt personnel accordé — à suivre" },
];

const seedAccounts: Account[] = [
  { id: "a1", name: "SIB", kind: "Banque", openingBalance: 0 },
  { id: "a2", name: "PETTY CASH", kind: "Banque", openingBalance: 0 },
  { id: "a3", name: "Dépôt LOYER", kind: "Banque", openingBalance: 0 },
  { id: "a4", name: "Revenus MAZDA", kind: "Banque", openingBalance: 0 },
  { id: "a5", name: "SALAIRE", kind: "Banque", openingBalance: 0 },
  { id: "a6", name: "SGO", kind: "Banque", openingBalance: 0 },
  { id: "a7", name: "PUMP", kind: "Espèces", openingBalance: 0 },
];

const seedBudgets: CategoryBudget[] = [
  { id: "b1", category: "Cadeaux", amount: 300000, rollover: false },
  { id: "b2", category: "Divertissement", amount: 200000, rollover: false },
  { id: "b3", category: "Invitation", amount: 100000, rollover: false },
];

const seedGoals: Goal[] = [
  { id: "g1", name: "Valeur nette cible", target: 20000000, current: 11957163, date: "déc. 2027" },
];

const seedRecurring: RecurringTemplate[] = [
  { id: "r1", category: "Logement", type: "Dépense", amount: 480000, frequency: "Mensuelle", nextDate: "2026-09-01" },
  { id: "r2", category: "Plan Éducation", type: "Dépense", amount: 32457, frequency: "Mensuelle", nextDate: "2026-09-01" },
  { id: "r3", category: "Un salaire", type: "Revenu", amount: 1629526, frequency: "Mensuelle", nextDate: "2026-09-01" },
];

// ============================================================
// HELPERS
// ============================================================
const MONTH_NAMES = ["janv","fév","mars","avr","mai","juin","juil","août","sept","oct","nov","déc"];
function monthLabel(m: string) {
  if (!m || typeof m !== "string" || !m.includes("_")) return "—";
  const [y, mm] = m.split("_");
  return `${MONTH_NAMES[parseInt(mm, 10) - 1]} ${y.slice(2)}`;
}
function monthSortKey(m: string) {
  if (!m || typeof m !== "string" || !m.includes("_")) return 0;
  const [y, mm] = m.split("_");
  return parseInt(y, 10) * 12 + parseInt(mm, 10);
}
const fmt = (n: number) => new Intl.NumberFormat("fr-FR").format(Math.round(n));
const fmtShort = (n: number) => {
  const abs = Math.abs(n);
  if (abs >= 1000000) return `${(n / 1000000).toFixed(1)}M`;
  if (abs >= 1000) return `${(n / 1000).toFixed(0)}k`;
  return `${n}`;
};
const uid = (p = "t") => `${p}${Date.now()}${Math.floor(Math.random() * 10000)}`;
const mean = (arr: number[]) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);
function pad2(n: number) { return n < 10 ? `0${n}` : `${n}`; }
function todayISO() { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function nowTime() { const d = new Date(); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; }
function dateToMonthKey(date: string) {
  if (!date || typeof date !== "string" || !date.includes("-")) return dateToMonthKey(todayISO());
  const [y, m] = date.split("-");
  return `${parseInt(y, 10)}_${parseInt(m, 10)}`;
}
function monthKeyToFirstDate(monthKey: string) {
  if (!monthKey || typeof monthKey !== "string" || !monthKey.includes("_")) return todayISO();
  const [y, m] = monthKey.split("_");
  return `${y}-${pad2(parseInt(m, 10))}-01`;
}
function dateLabelFull(date: string) {
  try {
    const d = new Date(date + "T00:00:00");
    return d.toLocaleDateString("fr-FR", { day: "numeric", month: "short", year: "2-digit" });
  } catch { return date; }
}
function dateLabelShort(date: string) {
  try {
    const d = new Date(date + "T00:00:00");
    return d.toLocaleDateString("fr-FR", { day: "numeric", month: "short" });
  } catch { return date; }
}
function addInterval(date: string, freq: "Hebdomadaire" | "Mensuelle" | "Annuelle") {
  const d = new Date(date + "T00:00:00");
  if (freq === "Hebdomadaire") d.setDate(d.getDate() + 7);
  else if (freq === "Mensuelle") d.setMonth(d.getMonth() + 1);
  else d.setFullYear(d.getFullYear() + 1);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function daysBetween(a: string, b: string) {
  const d1 = new Date(a + "T00:00:00"), d2 = new Date(b + "T00:00:00");
  return Math.round((d2.getTime() - d1.getTime()) / 86400000);
}
const stdev = (arr: number[]) => {
  if (arr.length < 2) return 0;
  const m = mean(arr);
  return Math.sqrt(mean(arr.map((x) => (x - m) ** 2)));
};

const GROUPS: Group[] = ["Nécessaire", "Productif", "Non-productif", "Non classifié"];
function getSubcategories(type: TxType, category: string): string[] {
  return (type === "Dépense" ? depSubcategories[category] : revSubcategories[category]) || [];
}
function normalizeText(s: string): string {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}
function categoriesForType(transactions: Transaction[], type: TxType): string[] {
  const used = new Set(transactions.filter((t) => t.type === type).map((t) => t.category));
  const known = Object.keys(type === "Dépense" ? depSubcategories : revSubcategories);
  known.forEach((c) => used.add(c));
  return Array.from(used).sort((a, b) => a.localeCompare(b, "fr"));
}
function defaultQuickCategory(transactions: Transaction[], type: TxType): string {
  const list = categoriesForType(transactions, type);
  if (type === "Dépense" && list.includes("Aliments")) return "Aliments";
  return list[0] || "";
}
function defaultQuickAccount(accounts: Account[]): string {
  const found = accounts.find((a) => a.name === "SALAIRE");
  return found ? found.name : (accounts[0]?.name || "");
}
const groupColor: Record<string, string> = {
  "Nécessaire": COLOR.slateBlue, "Productif": COLOR.emerald, "Non-productif": COLOR.clay,
  "Non classifié": COLOR.inkMuted, "Revenu": COLOR.goldSoft,
};

function CustomTooltip({ active, payload, label }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, padding: "10px 14px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLOR.ink }}>
      <div style={{ color: COLOR.inkMuted, marginBottom: 4 }}>{label}</div>
      {payload.map((p: any, i: number) => <div key={i} style={{ color: p.color || p.fill }}>{p.name}: {fmt(p.value)} FCFA</div>)}
    </div>
  );
}

function usePersistentState<T>(key: string, initial: T): [T, (v: T) => void, boolean] {
  const [state, setState] = useState<T>(initial);
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    try {
      const raw = localStorage.getItem(key);
      if (raw) setState(JSON.parse(raw));
    } catch {}
    setLoaded(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => {
    if (!loaded) return;
    try {
      localStorage.setItem(key, JSON.stringify(state));
    } catch {}
  }, [state, loaded]);
  return [state, setState, loaded];
}

// ============================================================
// UI PRIMITIVES
// ============================================================
function Panel({ title, subtitle, right, children, style = {} }: {
  title?: string; subtitle?: string; right?: React.ReactNode; children: React.ReactNode; style?: React.CSSProperties;
}) {
  return (
    <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 24, ...style }}>
      {(title || right) && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18, gap: 12, flexWrap: "wrap" }}>
          <div>
            {title && <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 500, margin: 0, color: COLOR.ink }}>{title}</h3>}
            {subtitle && <div style={{ color: COLOR.inkMuted, fontSize: 12.5, marginTop: 4 }}>{subtitle}</div>}
          </div>
          {right}
        </div>
      )}
      {children}
    </div>
  );
}

// Panneau avec bouton d'aide (?) — affiche une explication du graphique au clic,
// juste au-dessus du contenu, sans quitter la page.
function PanelWithHelp({ title, subtitle, explain, right, children, style = {} }: {
  title?: string; subtitle?: string; explain: string; right?: React.ReactNode; children: React.ReactNode; style?: React.CSSProperties;
}) {
  const [showHelp, setShowHelp] = useState(false);
  return (
    <Panel
      title={title}
      subtitle={subtitle}
      style={style}
      right={
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {right}
          <button onClick={() => setShowHelp((s) => !s)} title="Comprendre ce graphique" style={{
            width: 24, height: 24, borderRadius: "50%", border: `1px solid ${showHelp ? COLOR.gold : COLOR.hairline}`,
            background: showHelp ? "rgba(201,162,39,0.15)" : "transparent", color: showHelp ? COLOR.goldSoft : COLOR.inkMuted,
            cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
          }}>
            <HelpCircle size={14} />
          </button>
        </div>
      }
    >
      {showHelp && (
        <div style={{ background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: "12px 14px", marginBottom: 16, fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
          {explain}
        </div>
      )}
      {children}
    </Panel>
  );
}

function Kpi({ label, value, suffix = "FCFA", tone = COLOR.ink, icon: Icon, hint }: {
  label: string; value: string; suffix?: string; tone?: string; icon?: any; hint?: string;
}) {
  return (
    <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "16px 18px", flex: 1, minWidth: 190 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLOR.inkMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10 }}>
        {Icon && <Icon size={12.5} />} {label}
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 600, color: tone }}>
        {value}<span style={{ fontSize: 11.5, color: COLOR.inkMuted, marginLeft: 6 }}>{suffix}</span>
      </div>
      {hint && <div style={{ fontSize: 11, color: COLOR.inkMuted, marginTop: 6 }}>{hint}</div>}
    </div>
  );
}

function Select({ value, onChange, options, label }: {
  value: string; onChange: (v: string) => void; options: { value: string; label: string }[]; label?: string;
}) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      {label && <label style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</label>}
      <select value={value} onChange={(e) => onChange(e.target.value)} style={{
        background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink,
        padding: "8px 10px", fontSize: 12.5, fontFamily: "'Inter', sans-serif", minWidth: 130, cursor: "pointer",
      }}>
        {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </div>
  );
}

function EmptyState({ text = "Aucune donnée pour cette combinaison de filtres." }: { text?: string }) {
  return <div style={{ padding: "40px 0", textAlign: "center", color: COLOR.inkMuted, fontSize: 13 }}>{text}</div>;
}

function iconBtnStyle(color: string): React.CSSProperties {
  return { background: "transparent", border: "none", color, cursor: "pointer", padding: 5, marginRight: 2, display: "inline-flex", alignItems: "center" };
}
function pagerBtn(disabled: boolean): React.CSSProperties {
  return { background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: disabled ? COLOR.hairline : COLOR.inkMuted, padding: "6px 12px", fontSize: 12, cursor: disabled ? "default" : "pointer" };
}
const inputStyle: React.CSSProperties = {
  background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink,
  padding: "6px 9px", fontSize: 12.5, fontFamily: "'Inter', sans-serif", width: "100%", boxSizing: "border-box",
};

// ============================================================
// FILTER BAR
// ============================================================
interface Filters { from: string; to: string; type: string; group: string; category: string; subcategory: string; search: string; scope: string; }

function FilterBar({ filters, setFilters, allMonths, allCategories, onReset }: {
  filters: Filters; setFilters: (f: Filters) => void; allMonths: string[]; allCategories: string[]; onReset: () => void;
}) {
  const patch = (p: Partial<Filters>) => setFilters({ ...filters, ...p });
  return (
    <div className="gl-noprint" style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "16px 20px", display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLOR.gold, fontSize: 11.5, marginBottom: 5, textTransform: "uppercase", letterSpacing: "0.08em" }}>
        <Filter size={13} /> Filtres
      </div>
      <Select label="Du mois" value={filters.from} onChange={(v) => patch({ from: v })} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
      <Select label="Au mois" value={filters.to} onChange={(v) => patch({ to: v })} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
      <Select label="Type" value={filters.type} onChange={(v) => patch({ type: v })} options={[{ value: "Tous", label: "Tous" }, { value: "Dépense", label: "Dépenses" }, { value: "Revenu", label: "Revenus" }]} />
      <Select label="Groupe" value={filters.group} onChange={(v) => patch({ group: v })} options={[{ value: "Tous", label: "Tous" }, ...GROUPS.map((g) => ({ value: g, label: g })), { value: "Revenu", label: "Revenu" }]} />
      <Select label="Portée" value={filters.scope} onChange={(v) => patch({ scope: v })} options={[{ value: "Tous", label: "Tous" }, { value: "Personnel", label: "Personnel" }, { value: "Business", label: "Business" }]} />
      <Select label="Catégorie" value={filters.category} onChange={(v) => patch({ category: v, subcategory: "Toutes" })} options={[{ value: "Toutes", label: "Toutes" }, ...allCategories.map((c) => ({ value: c, label: c }))]} />
      {filters.category !== "Toutes" && (
        <Select label="Sous-catégorie" value={filters.subcategory} onChange={(v) => patch({ subcategory: v })}
          options={[{ value: "Toutes", label: "Toutes" }, ...Array.from(new Set([...depSubcategories[filters.category] || [], ...revSubcategories[filters.category] || []])).map((s) => ({ value: s, label: s }))]} />
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 160 }}>
        <label style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>Recherche</label>
        <div style={{ position: "relative" }}>
          <Search size={13} color={COLOR.inkMuted} style={{ position: "absolute", left: 9, top: 9 }} />
          <input value={filters.search} onChange={(e) => patch({ search: e.target.value })} placeholder="ex: Déjeuner, GRUNDFOS…" style={{ width: "100%", background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink, padding: "8px 10px 8px 28px", fontSize: 12.5, fontFamily: "'Inter', sans-serif", boxSizing: "border-box" }} />
        </div>
      </div>
      <button onClick={onReset} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 12px", fontSize: 12, cursor: "pointer", height: 34 }}>
        <RotateCcw size={12} /> Réinitialiser
      </button>
    </div>
  );
}

// ============================================================ END OF PART 1 — continued below in same file
// ============================================================
// ANALYTICS HELPERS
// ============================================================
function computeHealthScore(tauxEpargne: number, pctNonProd: number, monthlyRevenues: number[]) {
  const savingsScore = Math.max(0, Math.min(100, (tauxEpargne / 20) * 100));
  const nonProdScore = Math.max(0, Math.min(100, 100 - pctNonProd * 2));
  const cv = mean(monthlyRevenues) > 0 ? stdev(monthlyRevenues) / mean(monthlyRevenues) : 1;
  const stabilityScore = Math.max(0, Math.min(100, 100 - cv * 80));
  const overall = (savingsScore + nonProdScore + stabilityScore) / 3;
  let grade = "À surveiller", gradeColor = COLOR.clay;
  if (overall >= 80) { grade = "Excellent"; gradeColor = COLOR.emerald; }
  else if (overall >= 60) { grade = "Bon"; gradeColor = COLOR.emeraldSoft; }
  else if (overall >= 40) { grade = "Moyen"; gradeColor = COLOR.gold; }
  return { savingsScore, nonProdScore, stabilityScore, overall, grade, gradeColor };
}

function projectNetWorth(months = 12, series: [string, number][] = netWorthRaw) {
  const recent = series.slice(-6);
  const deltas = recent.slice(1).map((v, i) => v[1] - recent[i][1]);
  const avgDelta = mean(deltas);
  const sd = stdev(deltas);
  const last = series[series.length - 1][1];
  const points = [];
  for (let i = 0; i <= months; i++) {
    points.push({
      mois: i === 0 ? "aujourd'hui" : `+${i}m`,
      central: last + avgDelta * i,
      haut: last + (avgDelta + sd) * i,
      bas: last + (avgDelta - sd) * i,
    });
  }
  return { points, avgDelta, sd };
}

function detectAnomalies(filtered: any[]) {
  const byCategory: Record<string, number[]> = {};
  filtered.filter((t) => t.type === "Dépense").forEach((t) => {
    (byCategory[t.category] = byCategory[t.category] || []).push(t.amount);
  });
  const anomalies: any[] = [];
  filtered.filter((t) => t.type === "Dépense").forEach((t) => {
    const arr = byCategory[t.category];
    if (arr.length < 3) return;
    const m = mean(arr);
    if (t.amount > m * 2 && t.amount > 50000) {
      anomalies.push({ ...t, avg: m, ratio: t.amount / m });
    }
  });
  return anomalies.sort((a, b) => b.ratio - a.ratio).slice(0, 12);
}

// ============================================================
// MOTEUR D'ANALYSE — alertes, conseils, constats positifs générés
// automatiquement à partir des transactions filtrées (règles déterministes).
// ============================================================
type InsightKind = "alerte" | "conseil" | "positif";
interface Insight { kind: InsightKind; title: string; text: string; }
interface CatFocus {
  typeView: TxType; periodLabel: string;
  catList: { name: string; value: number; pct: number }[];
  prevByCat: Record<string, number>;
  total: number; prevTotal: number; delta: number; deltaPct: number;
}

function generateInsights(filtered: any[], catFocus?: CatFocus): Insight[] {
  const insights: Insight[] = [];
  const revenus = filtered.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const depenses = filtered.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const solde = revenus - depenses;
  const tauxEpargne = revenus > 0 ? (solde / revenus) * 100 : 0;
  const nonProd = filtered.filter((t) => t.type === "Dépense" && t.group === "Non-productif").reduce((a, t) => a + t.amount, 0);
  const pctNonProd = revenus > 0 ? (nonProd / revenus) * 100 : 0;
  const productif = filtered.filter((t) => t.type === "Dépense" && t.group === "Productif").reduce((a, t) => a + t.amount, 0);
  const pctProductif = depenses > 0 ? (productif / depenses) * 100 : 0;

  // Solde / taux d'épargne
  if (solde < 0) {
    insights.push({ kind: "alerte", title: "Solde négatif", text: `Les dépenses dépassent les revenus de ${fmt(Math.abs(solde))} FCFA sur cette période. À surveiller de près si la tendance se confirme sur les mois suivants.` });
  } else if (tauxEpargne < 10) {
    insights.push({ kind: "conseil", title: "Taux d'épargne faible", text: `${tauxEpargne.toFixed(1)}% du revenu est épargné, en dessous du repère habituel de 15-20%. Une réduction ciblée du non-productif pourrait combler l'écart.` });
  } else if (tauxEpargne >= 20) {
    insights.push({ kind: "positif", title: "Bon taux d'épargne", text: `${tauxEpargne.toFixed(1)}% du revenu est épargné sur cette période — au-dessus du repère de 20% généralement recommandé.` });
  }

  // Non-productif
  if (pctNonProd > 30) {
    insights.push({ kind: "alerte", title: "Non-productif élevé", text: `Les dépenses non-productives (cadeaux, sorties, shopping…) représentent ${pctNonProd.toFixed(1)}% du revenu — au-dessus du seuil de vigilance de 30%.` });
  } else if (pctNonProd > 15) {
    insights.push({ kind: "conseil", title: "Non-productif à surveiller", text: `${pctNonProd.toFixed(1)}% du revenu part en dépenses non-productives. Réduire ce poste de quelques points libérerait une marge d'épargne significative.` });
  } else if (revenus > 0) {
    insights.push({ kind: "positif", title: "Non-productif maîtrisé", text: `Seulement ${pctNonProd.toFixed(1)}% du revenu part en dépenses non-productives — une discipline budgétaire solide.` });
  }

  // Productif
  if (depenses > 0 && pctProductif >= 35) {
    insights.push({ kind: "positif", title: "Investissement soutenu", text: `${pctProductif.toFixed(1)}% des dépenses sont classées productives (immobilier, activité, épargne…) — la stratégie patrimoniale reste la priorité budgétaire.` });
  }

  // Tendance mensuelle (comparaison 1re moitié vs 2e moitié de la période)
  const monthKeys = Array.from(new Set(filtered.map((t: any) => t.month))).sort((a: any, b: any) => monthSortKey(a) - monthSortKey(b));
  if (monthKeys.length >= 4) {
    const mid = Math.floor(monthKeys.length / 2);
    const firstHalf = monthKeys.slice(0, mid);
    const secondHalf = monthKeys.slice(mid);
    const depFirst = filtered.filter((t: any) => t.type === "Dépense" && firstHalf.includes(t.month)).reduce((a: number, t: any) => a + t.amount, 0) / firstHalf.length;
    const depSecond = filtered.filter((t: any) => t.type === "Dépense" && secondHalf.includes(t.month)).reduce((a: number, t: any) => a + t.amount, 0) / secondHalf.length;
    if (depFirst > 0) {
      const delta = ((depSecond - depFirst) / depFirst) * 100;
      if (delta > 20) insights.push({ kind: "alerte", title: "Dépenses en hausse", text: `Les dépenses mensuelles moyennes ont augmenté de ${delta.toFixed(0)}% entre la première et la seconde moitié de la période.` });
      else if (delta < -20) insights.push({ kind: "positif", title: "Dépenses en baisse", text: `Les dépenses mensuelles moyennes ont diminué de ${Math.abs(delta).toFixed(0)}% entre la première et la seconde moitié de la période.` });
    }
  }

  // Catégorie dominante côté non-productif
  const nonProdCats: Record<string, number> = {};
  filtered.filter((t: any) => t.type === "Dépense" && t.group === "Non-productif").forEach((t: any) => { nonProdCats[t.category] = (nonProdCats[t.category] || 0) + t.amount; });
  const topNonProd = Object.entries(nonProdCats).sort((a, b) => b[1] - a[1])[0];
  if (topNonProd && nonProd > 0 && topNonProd[1] / nonProd > 0.4) {
    insights.push({ kind: "conseil", title: `"${topNonProd[0]}" concentre le non-productif`, text: `Cette catégorie représente à elle seule ${((topNonProd[1] / nonProd) * 100).toFixed(0)}% des dépenses non-productives (${fmt(topNonProd[1])} FCFA) — le premier levier d'économie à considérer.` });
  }

  // Anomalies ponctuelles
  const anomalies = detectAnomalies(filtered);
  if (anomalies.length > 0) {
    insights.push({ kind: "alerte", title: `${anomalies.length} transaction(s) atypique(s)`, text: `Certaines dépenses dépassent largement la moyenne habituelle de leur catégorie (ex: ${anomalies[0].category} à ${fmt(anomalies[0].amount)} FCFA, ${anomalies[0].ratio.toFixed(1)}× la moyenne). Détail dans l'onglet Catégories.` });
  }

  if (!insights.length) {
    insights.push({ kind: "positif", title: "Rien à signaler", text: "Aucune alerte particulière sur cette période — les indicateurs sont dans des plages normales." });
  }

  // Analyse ciblée sur les "Principales catégories" (période + type actuellement filtrés)
  if (catFocus && catFocus.catList.length) {
    const { typeView, periodLabel, catList, prevByCat, total: cfTotal, delta: cfDelta, deltaPct: cfDeltaPct } = catFocus;
    const verbeType = typeView === "Dépense" ? "dépenses" : "revenus";
    const cfImproved = typeView === "Dépense" ? cfDelta <= 0 : cfDelta >= 0;

    // Tendance globale sur la période filtrée vs période précédente équivalente
    if (Math.abs(cfDeltaPct) >= 8) {
      insights.push({
        kind: cfImproved ? "positif" : (Math.abs(cfDeltaPct) >= 20 ? "alerte" : "conseil"),
        title: `${verbeType[0].toUpperCase()}${verbeType.slice(1)} ${cfDelta >= 0 ? "en hausse" : "en baisse"} de ${Math.abs(cfDeltaPct).toFixed(0)}%`,
        text: `Sur ${periodLabel}, les ${verbeType} s'élèvent à ${fmt(cfTotal)} FCFA contre ${fmt(catFocus.prevTotal)} FCFA sur la période comparable précédente, soit ${cfDelta >= 0 ? "+" : "−"}${fmt(Math.abs(cfDelta))} FCFA.`,
      });
    }

    // Concentration sur la catégorie dominante
    const top = catList[0];
    if (top) {
      if (top.pct >= 40) {
        insights.push({ kind: "alerte", title: `"${top.name}" concentre l'essentiel`, text: `Cette catégorie représente à elle seule ${top.pct.toFixed(0)}% des ${verbeType} de la période (${fmt(top.value)} FCFA) — une forte concentration qui mérite d'être questionnée si elle n'est pas structurelle (loyer, salaire fixe…).` });
      } else if (top.pct >= 25) {
        insights.push({ kind: "conseil", title: `"${top.name}" en tête`, text: `Premier poste de la période avec ${top.pct.toFixed(0)}% du total (${fmt(top.value)} FCFA). À garder à l'œil si elle continue de progresser.` });
      }
    }

    // Catégorie avec la plus forte progression vs période précédente
    let biggestMover: { name: string; prevVal: number; curVal: number; pct: number } | null = null;
    catList.forEach((c) => {
      const prevVal = prevByCat[c.name] || 0;
      if (prevVal > 0) {
        const pct = ((c.value - prevVal) / prevVal) * 100;
        if (pct >= 50 && (!biggestMover || pct > biggestMover.pct)) biggestMover = { name: c.name, prevVal, curVal: c.value, pct };
      } else if (c.value > 0 && !biggestMover) {
        biggestMover = { name: c.name, prevVal: 0, curVal: c.value, pct: 100 };
      }
    });
    if (biggestMover) {
      const bm = biggestMover as { name: string; prevVal: number; curVal: number; pct: number };
      insights.push({
        kind: typeView === "Dépense" ? "alerte" : "positif",
        title: `Forte progression sur "${bm.name}"`,
        text: bm.prevVal > 0
          ? `Passée de ${fmt(bm.prevVal)} à ${fmt(bm.curVal)} FCFA (+${bm.pct.toFixed(0)}%) par rapport à la période précédente — le principal moteur de la variation observée.`
          : `Nouvelle sur cette période, pour ${fmt(bm.curVal)} FCFA — absente ou négligeable lors de la période précédente.`,
      });
    }

    // Catégorie en forte baisse (bonne nouvelle côté dépenses)
    let biggestDrop: { name: string; prevVal: number; curVal: number; pct: number } | null = null;
    Object.entries(prevByCat).forEach(([name, prevVal]) => {
      const curVal = catList.find((c) => c.name === name)?.value || 0;
      if (prevVal > 0 && curVal < prevVal) {
        const pct = ((curVal - prevVal) / prevVal) * 100;
        if (pct <= -40 && (!biggestDrop || pct < biggestDrop.pct)) biggestDrop = { name, prevVal, curVal, pct };
      }
    });
    if (biggestDrop && typeView === "Dépense") {
      const bd = biggestDrop as { name: string; prevVal: number; curVal: number; pct: number };
      insights.push({ kind: "positif", title: `Nette baisse sur "${bd.name}"`, text: `Passée de ${fmt(bd.prevVal)} à ${fmt(bd.curVal)} FCFA (${bd.pct.toFixed(0)}%) — un effort payant à maintenir.` });
    }

    // Diversification : beaucoup de petites catégories résiduelles
    const smallCats = catList.filter((c) => c.pct < 3);
    if (smallCats.length >= 8) {
      const smallTotal = smallCats.reduce((a, c) => a + c.value, 0);
      insights.push({ kind: "conseil", title: "Beaucoup de petites lignes éparses", text: `${smallCats.length} catégories pèsent chacune moins de 3% (${fmt(smallTotal)} FCFA cumulés). Les regrouper ou les revoir pourrait simplifier le suivi sans perte d'information utile.` });
    }
  }

  const order: Record<InsightKind, number> = { alerte: 0, conseil: 1, positif: 2 };
  return insights.sort((a, b) => order[a.kind] - order[b.kind]);
}

// ============================================================
// SÉLECTEUR CATÉGORIE + SOUS-CATÉGORIE AVEC RECHERCHE
// ============================================================
function CategoryPickerSheet({ open, onClose, transactions, type, value, subvalue, onSelect }: {
  open: boolean; onClose: () => void; transactions: Transaction[]; type: TxType; value: string; subvalue: string;
  onSelect: (cat: string, sub: string) => void;
}) {
  const [query, setQuery] = useState("");
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  useEffect(() => { if (open) { setQuery(""); setExpanded(new Set(value ? [value] : [])); } }, [open, value]);
  if (!open) return null;

  const cats = categoriesForType(transactions, type);
  const q = normalizeText(query.trim());
  const filteredCats = q
    ? cats.filter((c) => normalizeText(c).includes(q) || getSubcategories(type, c).some((s) => normalizeText(s).includes(q)))
    : cats;

  const toggleExpand = (c: string) => setExpanded((prev) => { const n = new Set(prev); n.has(c) ? n.delete(c) : n.add(c); return n; });

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 400, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "100%", maxWidth: 480, maxHeight: "80vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
        display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 18px", borderBottom: `1px solid ${COLOR.hairline}` }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16 }}>Choisir une catégorie</div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex" }}><X size={18} /></button>
        </div>
        <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: "8px 10px", WebkitOverflowScrolling: "touch" }}>
          {filteredCats.map((c) => {
            const subs = getSubcategories(type, c);
            const matchesQuery = q && (normalizeText(c).includes(q) || subs.some((s) => normalizeText(s).includes(q)));
            const isExpanded = expanded.has(c) || !!matchesQuery;
            const isCatSelected = value === c && !subvalue;
            const visibleSubs = q ? subs.filter((s) => normalizeText(c).includes(q) || normalizeText(s).includes(q)) : subs;
            return (
              <div key={c}>
                <button
                  onClick={() => { onSelect(c, ""); if (subs.length) toggleExpand(c); else onClose(); }}
                  style={{
                    width: "100%", textAlign: "left", padding: "12px 14px", borderRadius: 8, marginBottom: 2,
                    background: isCatSelected ? "rgba(201,162,39,0.12)" : "transparent", border: "none", cursor: "pointer",
                    display: "flex", justifyContent: "space-between", alignItems: "center",
                  }}>
                  <span style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, color: COLOR.ink, fontWeight: isCatSelected ? 600 : 400 }}>
                    {subs.length > 0 && <ChevronRight size={13} color={COLOR.inkMuted} style={{ transform: isExpanded ? "rotate(90deg)" : "none", transition: "transform 0.15s", flexShrink: 0 }} />}
                    {c}
                  </span>
                  {isCatSelected && <Check size={14} color={COLOR.goldSoft} />}
                </button>
                {isExpanded && subs.length > 0 && (
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8, padding: "4px 8px 14px 33px" }}>
                    {visibleSubs.map((s) => {
                      const active = value === c && subvalue === s;
                      return (
                        <button key={s} onClick={() => { onSelect(c, s); onClose(); }} style={{
                          padding: "7px 14px", borderRadius: 16, fontSize: 12.5, cursor: "pointer", whiteSpace: "nowrap",
                          border: `1px solid ${active ? COLOR.gold : COLOR.hairline}`,
                          background: active ? COLOR.gold : COLOR.surfaceInput,
                          color: active ? COLOR.bg : COLOR.inkMuted, fontWeight: active ? 600 : 400,
                        }}>{s}</button>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
          {!filteredCats.length && <div style={{ padding: "30px 10px", textAlign: "center", color: COLOR.inkMuted, fontSize: 13 }}>Aucun résultat pour "{query}"</div>}
        </div>
        <div className="gl-safe-bottom" style={{ padding: "12px 16px", borderTop: `1px solid ${COLOR.hairline}` }}>
          <div style={{ position: "relative" }}>
            <Search size={14} color={COLOR.inkMuted} style={{ position: "absolute", left: 12, top: 12 }} />
            <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Rechercher catégorie ou sous-catégorie…"
              style={{ width: "100%", background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "11px 14px 11px 34px", color: COLOR.ink, fontSize: 14, boxSizing: "border-box", outline: "none" }} />
          </div>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// FICHE DE MODIFICATION — même style visuel que "Saisie rapide",
// réutilisable partout où on veut éditer une transaction existante.
// ============================================================
function TransactionEditSheet({ open, transaction, transactions, accounts, onClose, onSave, onDelete }: {
  open: boolean; transaction: Transaction | null; transactions: Transaction[]; accounts: Account[];
  onClose: () => void; onSave: (t: Transaction) => void; onDelete: (id: string) => void;
}) {
  const [date, setDate] = useState(todayISO());
  const [time, setTime] = useState(nowTime());
  const [type, setType] = useState<TxType>("Dépense");
  const [category, setCategory] = useState("");
  const [subcategory, setSubcategory] = useState("");
  const [amount, setAmount] = useState<number | "">("");
  const [account, setAccount] = useState("");
  const [catPickerOpen, setCatPickerOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    if (open && transaction) {
      setDate(transaction.date); setTime(transaction.time || nowTime()); setType(transaction.type);
      setCategory(transaction.category); setSubcategory(transaction.subcategory || "");
      setAmount(transaction.amount); setAccount(transaction.account || defaultQuickAccount(accounts));
      setSaved(false);
    }
  }, [open, transaction]);

  if (!open || !transaction) return null;
  const typeColor = type === "Revenu" ? COLOR.emerald : COLOR.clay;
  const fieldLabel: React.CSSProperties = { fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 };
  const nakedSelect: React.CSSProperties = {
    background: "transparent", border: "none", color: COLOR.ink, fontSize: 16, fontWeight: 600,
    fontFamily: "'Fraunces', serif", padding: 0, cursor: "pointer", width: "100%", appearance: "none", WebkitAppearance: "none",
  };

  const submit = () => {
    if (!category || !amount || Number(amount) <= 0) return;
    onSave({ ...transaction, date, time, type, category, subcategory: subcategory || undefined, amount: Number(amount), account: account || undefined });
    setSaved(true);
    setTimeout(() => { setSaved(false); onClose(); }, 700);
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 450, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center", padding: 0 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: "100%", maxWidth: 480, maxHeight: "92vh", overflowY: "auto" }} className="gl-scroll">
        <div style={{ background: `linear-gradient(180deg, ${COLOR.surfaceRaised} 0%, ${COLOR.surface} 70%)`, border: `1px solid ${COLOR.hairline}`, borderBottom: "none", borderRadius: "16px 16px 0 0", overflow: "hidden" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 20px 0 20px" }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 15, color: COLOR.goldSoft }}>Modifier la transaction</div>
            <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex" }}><X size={18} /></button>
          </div>
          {/* Type + Date */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 20px 6px 20px", flexWrap: "wrap", gap: 10 }}>
            <div style={{ display: "flex", gap: 8 }}>
              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "8px 16px" }}>
                <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
                  style={{ background: "transparent", border: "none", color: COLOR.ink, fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" }} />
              </div>
              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "8px 16px" }}>
                <input type="time" value={time} onChange={(e) => setTime(e.target.value)}
                  style={{ background: "transparent", border: "none", color: COLOR.ink, fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" }} />
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, background: COLOR.surface, borderRadius: 24, padding: 5, border: `1px solid ${COLOR.hairline}` }}>
              <button onClick={() => { setType("Dépense"); setSubcategory(""); }} title="Dépense" style={{
                width: 34, height: 34, borderRadius: "50%", border: "none", cursor: "pointer",
                background: type === "Dépense" ? COLOR.clay : "transparent", color: type === "Dépense" ? COLOR.bg : COLOR.claySoft,
                display: "flex", alignItems: "center", justifyContent: "center",
              }}><Minus size={16} strokeWidth={2.5} /></button>
              <button onClick={() => { setType("Revenu"); setSubcategory(""); }} title="Revenu" style={{
                width: 34, height: 34, borderRadius: "50%", border: "none", cursor: "pointer",
                background: type === "Revenu" ? COLOR.emerald : "transparent", color: type === "Revenu" ? COLOR.bg : COLOR.emeraldSoft,
                display: "flex", alignItems: "center", justifyContent: "center",
              }}><Plus size={16} strokeWidth={2.5} /></button>
            </div>
          </div>

          {/* Montant */}
          <div style={{ position: "relative", display: "flex", flexDirection: "column", alignItems: "center", padding: "14px 24px 8px 24px" }}>
            <div style={{ position: "absolute", fontSize: 56, fontWeight: 700, color: typeColor, opacity: 0.07, fontFamily: "'Fraunces', serif", pointerEvents: "none", userSelect: "none", top: 6 }}>FCFA</div>
            <input type="number" inputMode="numeric" value={amount} placeholder="0"
              onChange={(e) => setAmount(e.target.value === "" ? "" : Number(e.target.value))}
              onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
              style={{ position: "relative", background: "transparent", border: "none", outline: "none", color: COLOR.ink, fontSize: 42, fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace", textAlign: "center", width: "100%", maxWidth: 260 }} />
          </div>

          {/* Compte / Catégorie */}
          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "16px 20px" }}>
            <div style={{ display: "flex", gap: 24 }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={fieldLabel}>Compte</div>
                <select value={account} onChange={(e) => setAccount(e.target.value)} style={nakedSelect}>
                  {!accounts.length && <option value="">Aucun compte créé</option>}
                  {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                </select>
              </div>
              <div style={{ flex: 1, minWidth: 0, textAlign: "right" }}>
                <div style={{ ...fieldLabel, textAlign: "right" }}>Catégorie</div>
                <button onClick={() => setCatPickerOpen(true)} style={{ ...nakedSelect, textAlign: "right", cursor: "pointer", display: "block" }}>
                  {category || "Choisir…"}
                </button>
                {subcategory && <div style={{ textAlign: "right", fontSize: 13, color: COLOR.inkMuted, marginTop: 4 }}>{subcategory}</div>}
              </div>
            </div>
            <CategoryPickerSheet open={catPickerOpen} onClose={() => setCatPickerOpen(false)} transactions={transactions} type={type}
              value={category} subvalue={subcategory} onSelect={(c, s) => { setCategory(c); setSubcategory(s); }} />

            <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
              <button onClick={() => setConfirmDelete(true)} style={{
                display: "flex", alignItems: "center", justifyContent: "center", gap: 6, padding: "14px 18px", borderRadius: 12,
                border: `1px solid ${COLOR.clay}`, background: "rgba(193,84,63,0.1)", color: COLOR.claySoft, fontSize: 13.5, cursor: "pointer",
              }}><Trash2 size={15} /> Supprimer</button>
              <button onClick={submit} disabled={!amount || Number(amount) <= 0} style={{
                flex: 1, padding: "14px 0", borderRadius: 12, border: "none",
                background: saved ? COLOR.emerald : (!amount || Number(amount) <= 0) ? COLOR.hairline : COLOR.gold,
                color: saved ? COLOR.bg : (!amount || Number(amount) <= 0) ? COLOR.inkMuted : COLOR.bg,
                fontSize: 14.5, fontWeight: 700, cursor: (!amount || Number(amount) <= 0) ? "default" : "pointer",
                display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "background 0.15s",
              }}>{saved ? <Check size={17} /> : null} {saved ? "Mis à jour" : "Mettre à jour"}</button>
            </div>
          </div>
        </div>
      </div>
      <ConfirmDialog
        open={confirmDelete}
        title="Supprimer cette transaction ?"
        message="Cette action est définitive. Le montant ne sera plus comptabilisé nulle part dans l'app."
        onConfirm={() => { setConfirmDelete(false); onDelete(transaction.id); onClose(); }}
        onCancel={() => setConfirmDelete(false)}
      />
    </div>
  );
}


function ConfirmDialog({ open, title, message, onConfirm, onCancel, confirmLabel = "Supprimer" }: {
  open: boolean; title: string; message: string; onConfirm: () => void; onCancel: () => void; confirmLabel?: string;
}) {
  if (!open) return null;
  return (
    <div onClick={onCancel} style={{ position: "fixed", inset: 0, zIndex: 500, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: "100%", maxWidth: 360, background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 16, padding: 24 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
          <AlertTriangle size={20} color={COLOR.claySoft} />
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>{title}</div>
        </div>
        <div style={{ fontSize: 13.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 22 }}>{message}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onCancel} style={{ flex: 1, padding: "11px 0", borderRadius: 10, border: `1px solid ${COLOR.hairline}`, background: "transparent", color: COLOR.inkMuted, fontSize: 13.5, cursor: "pointer" }}>
            Annuler
          </button>
          <button onClick={onConfirm} style={{ flex: 1, padding: "11px 0", borderRadius: 10, border: "none", background: COLOR.clay, color: COLOR.bg, fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

function InsightsPanel({ filtered, catFocus, title, subtitle }: { filtered: any[]; catFocus?: CatFocus; title?: string; subtitle?: string }) {
  const insights = useMemo(() => generateInsights(filtered, catFocus), [filtered, catFocus]);
  const styleFor: Record<InsightKind, { bg: string; border: string; color: string; icon: any }> = {
    alerte: { bg: "rgba(193,84,63,0.08)", border: COLOR.clay, color: COLOR.claySoft, icon: AlertTriangle },
    conseil: { bg: "rgba(201,162,39,0.08)", border: COLOR.gold, color: COLOR.goldSoft, icon: Info },
    positif: { bg: "rgba(63,156,122,0.08)", border: COLOR.emerald, color: COLOR.emeraldSoft, icon: Check },
  };
  return (
    <Panel title={title || "Analyse & conseils"} subtitle={subtitle || "Généré automatiquement à partir de la période et des filtres actifs"}>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {insights.map((ins, i) => {
          const s = styleFor[ins.kind];
          const Icon = s.icon;
          return (
            <div key={i} style={{ display: "flex", gap: 10, padding: "12px 14px", background: s.bg, border: `1px solid ${s.border}`, borderRadius: 8 }}>
              <Icon size={15} color={s.color} style={{ flexShrink: 0, marginTop: 1 }} />
              <div>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: s.color, marginBottom: 3 }}>{ins.title}</div>
                <div style={{ fontSize: 12, color: COLOR.inkMuted, lineHeight: 1.55 }}>{ins.text}</div>
              </div>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

// ============================================================
// SANKEY-LIKE FLOW (custom SVG — Revenus → Groupes de dépenses)
// ============================================================
function FlowDiagram({ filtered }: { filtered: any[] }) {
  const totalRevenus = filtered.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const depByGroup: Record<string, number> = {};
  filtered.filter((t) => t.type === "Dépense").forEach((t) => { depByGroup[t.group] = (depByGroup[t.group] || 0) + t.amount; });
  const totalDep = Object.values(depByGroup).reduce((a, b) => a + b, 0);
  const solde = totalRevenus - totalDep;

  const nodesRight = [
    ...GROUPS.map((g) => ({ name: g, value: depByGroup[g] || 0, color: groupColor[g] })),
    { name: "Solde (épargne)", value: Math.max(0, solde), color: COLOR.goldSoft },
  ].filter((n) => n.value > 0);

  const total = Math.max(1, totalRevenus);
  const width = 720, height = 320, leftX = 40, rightX = width - 220, nodeW = 90, midX = (leftX + nodeW + rightX) / 2;

  let yCursor = 20;
  const rightNodes = nodesRight.map((n) => {
    const h = Math.max((n.value / total) * (height - 40), 4);
    const node = { ...n, y0: yCursor, h };
    yCursor += h + 6;
    return node;
  });
  const leftH = Math.max(yCursor - 20 - 6, height - 40);

  return (
    <div style={{ overflowX: "auto" }}>
      <svg width={width} height={Math.max(height, yCursor + 20)} style={{ display: "block", margin: "0 auto" }}>
        {/* left node: Revenus */}
        <rect x={leftX} y={20} width={nodeW} height={leftH} rx={4} fill={COLOR.emerald} opacity={0.9} />
        <text x={leftX + nodeW / 2} y={20 + leftH / 2 - 6} textAnchor="middle" fill={COLOR.bg} fontSize={12} fontWeight={700} fontFamily="Inter">Revenus</text>
        <text x={leftX + nodeW / 2} y={20 + leftH / 2 + 12} textAnchor="middle" fill={COLOR.bg} fontSize={10.5} fontFamily="IBM Plex Mono">{fmtShort(totalRevenus)}</text>

        {/* flows — ribbon between matching y-ranges on left and right (shared cursor keeps them aligned) */}
        {rightNodes.map((n) => {
          const path = `M ${leftX + nodeW} ${n.y0}
                        C ${midX} ${n.y0}, ${midX} ${n.y0}, ${rightX} ${n.y0}
                        L ${rightX} ${n.y0 + n.h}
                        C ${midX} ${n.y0 + n.h}, ${midX} ${n.y0 + n.h}, ${leftX + nodeW} ${n.y0 + n.h}
                        Z`;
          return <path key={n.name} d={path} fill={n.color} opacity={0.3} />;
        })}

        {/* right nodes */}
        {rightNodes.map((n) => (
          <g key={n.name}>
            <rect x={rightX} y={n.y0} width={nodeW} height={n.h} rx={4} fill={n.color} />
            <text x={rightX + nodeW + 10} y={n.y0 + n.h / 2 - 5} fontSize={11.5} fill={COLOR.ink} fontFamily="Inter">{n.name}</text>
            <text x={rightX + nodeW + 10} y={n.y0 + n.h / 2 + 10} fontSize={10.5} fill={COLOR.inkMuted} fontFamily="IBM Plex Mono">{fmt(n.value)}</text>
          </g>
        ))}
      </svg>
    </div>
  );
}

// ============================================================
// HEATMAP CALENDRIER
// ============================================================
function HeatmapCalendar({ filtered }: { filtered: any[] }) {
  const byMonth: Record<string, number> = {};
  filtered.filter((t) => t.type === "Dépense" && t.group === "Non-productif").forEach((t) => {
    byMonth[t.month] = (byMonth[t.month] || 0) + t.amount;
  });
  const years = Array.from(new Set(filtered.map((t) => t.date.slice(0, 4)))).sort();
  const max = Math.max(1, ...Object.values(byMonth));

  const cellColor = (v: number) => {
    if (!v) return COLOR.surfaceRaised;
    const ratio = v / max;
    if (ratio > 0.66) return COLOR.clay;
    if (ratio > 0.33) return COLOR.gold;
    return COLOR.emerald;
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "50px repeat(12, 1fr)", gap: 5, alignItems: "center" }}>
        <div />
        {MONTH_NAMES.map((m) => (
          <div key={m} style={{ fontSize: 10, color: COLOR.inkMuted, textAlign: "center" }}>{m}</div>
        ))}
        {years.map((y) => (
          <React.Fragment key={y}>
            <div style={{ fontSize: 11.5, color: COLOR.inkMuted, fontFamily: "IBM Plex Mono, monospace" }}>{y}</div>
            {MONTH_NAMES.map((_, mi) => {
              const key = `${y}_${mi + 1}`;
              const v = byMonth[key];
              return (
                <div key={key} title={v ? `${monthLabel(key)}: ${fmt(v)} FCFA` : monthLabel(key)} style={{
                  height: 28, borderRadius: 4, background: cellColor(v), border: `1px solid ${COLOR.hairline}`,
                  display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, color: v ? COLOR.bg : COLOR.hairline, cursor: "default",
                }}>
                  {v ? fmtShort(v) : ""}
                </div>
              );
            })}
          </React.Fragment>
        ))}
      </div>
      <div style={{ display: "flex", gap: 14, marginTop: 14, fontSize: 11, color: COLOR.inkMuted, alignItems: "center" }}>
        <span>Intensité du non-productif :</span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}><i style={{ width: 10, height: 10, background: COLOR.emerald, display: "inline-block", borderRadius: 2 }} /> faible</span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}><i style={{ width: 10, height: 10, background: COLOR.gold, display: "inline-block", borderRadius: 2 }} /> moyenne</span>
        <span style={{ display: "flex", alignItems: "center", gap: 5 }}><i style={{ width: 10, height: 10, background: COLOR.clay, display: "inline-block", borderRadius: 2 }} /> élevée</span>
      </div>
    </div>
  );
}

// ============================================================ END OF PART 2 — continued below in same file
// ============================================================
// APERÇU TAB (KPIs + valeur nette + revenu/dépense + groupes + santé + comparaison)
// ============================================================
function ApercuTab({ filtered, filters, accounts, transactions }: { filtered: any[]; filters: Filters; accounts: Account[]; transactions: Transaction[] }) {
  const totalRevenus = filtered.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const totalDepenses = filtered.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const solde = totalRevenus - totalDepenses;
  const tauxEpargne = totalRevenus > 0 ? (solde / totalRevenus) * 100 : 0;
  const nonProd = filtered.filter((t) => t.group === "Non-productif").reduce((a, t) => a + t.amount, 0);
  const pctNonProd = totalRevenus > 0 ? (nonProd / totalRevenus) * 100 : 0;
  const monthsInRange = new Set(filtered.map((t) => t.month)).size || 1;

  const byMonth = useMemo(() => {
    const m: Record<string, { revenus: number; depenses: number }> = {};
    filtered.forEach((t) => {
      if (!m[t.month]) m[t.month] = { revenus: 0, depenses: 0 };
      if (t.type === "Revenu") m[t.month].revenus += t.amount; else m[t.month].depenses += t.amount;
    });
    return Object.keys(m).sort((a, b) => monthSortKey(a) - monthSortKey(b)).map((k) => ({ key: k, mois: monthLabel(k), revenus: m[k].revenus, depenses: m[k].depenses }));
  }, [filtered]);

  const monthlyRevenues = byMonth.map((m) => m.revenus);
  const health = computeHealthScore(tauxEpargne, pctNonProd, monthlyRevenues);

  const groupBreakdown = useMemo(() => {
    const m: Record<string, number> = {};
    filtered.filter((t) => t.type === "Dépense").forEach((t) => { m[t.group] = (m[t.group] || 0) + t.amount; });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [filtered]);

  const nwFiltered = liveNetWorthSeries(accounts, transactions).filter(([m]) => {
    const k = monthSortKey(m);
    return k >= monthSortKey(filters.from) && k <= monthSortKey(filters.to);
  }).map(([m, v]) => ({ mois: monthLabel(m), valeur: v }));

  // comparaison période sur période
  const cur = byMonth[byMonth.length - 1];
  const prev = byMonth[byMonth.length - 2];
  const yearAgo = byMonth.find((m) => cur && m.key === `${parseInt(cur.key.split("_")[0]) - 1}_${cur.key.split("_")[1]}`);

  const delta = (a?: number, b?: number) => (a !== undefined && b !== undefined && b !== 0 ? ((a - b) / b) * 100 : null);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Revenus (période)" value={fmt(totalRevenus)} tone={COLOR.emeraldSoft} icon={TrendingUp} />
        <Kpi label="Dépenses (période)" value={fmt(totalDepenses)} tone={COLOR.claySoft} icon={TrendingDown} />
        <Kpi label="Solde net" value={fmt(solde)} tone={solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} />
        <Kpi label="Taux d'épargne" value={tauxEpargne.toFixed(1)} suffix="%" tone={COLOR.gold} icon={Target} />
      </div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Non-productif (période)" value={fmt(nonProd)} tone={COLOR.claySoft} icon={AlertTriangle} />
        <Kpi label="Non-productif / revenu" value={pctNonProd.toFixed(1)} suffix="%" tone={COLOR.gold} />
        <Kpi label="Revenu moyen / mois" value={fmt(totalRevenus / monthsInRange)} tone={COLOR.emeraldSoft} />
        <Kpi label="Dépense moyenne / mois" value={fmt(totalDepenses / monthsInRange)} tone={COLOR.claySoft} />
      </div>

      <Panel title="Score de santé financière" subtitle="Composite : taux d'épargne, poids du non-productif, stabilité des revenus">
        <div style={{ display: "flex", gap: 28, alignItems: "center", flexWrap: "wrap" }}>
          <div style={{ position: "relative", width: 130, height: 130, flexShrink: 0 }}>
            <svg width={130} height={130}>
              <circle cx={65} cy={65} r={54} fill="none" stroke={COLOR.hairline} strokeWidth={12} />
              <circle cx={65} cy={65} r={54} fill="none" stroke={health.gradeColor} strokeWidth={12}
                strokeDasharray={`${(health.overall / 100) * 339} 339`} strokeLinecap="round" transform="rotate(-90 65 65)" />
            </svg>
            <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 26, fontWeight: 600, color: health.gradeColor }}>{health.overall.toFixed(0)}</div>
              <div style={{ fontSize: 10.5, color: COLOR.inkMuted }}>/ 100</div>
            </div>
          </div>
          <div style={{ flex: 1, minWidth: 260 }}>
            <div style={{ fontSize: 15, fontWeight: 500, color: health.gradeColor, marginBottom: 12 }}>{health.grade}</div>
            {[
              { label: "Taux d'épargne", val: health.savingsScore },
              { label: "Maîtrise du non-productif", val: health.nonProdScore },
              { label: "Stabilité des revenus", val: health.stabilityScore },
            ].map((s) => (
              <div key={s.label} style={{ marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                  <span style={{ color: COLOR.inkMuted }}>{s.label}</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s.val.toFixed(0)}</span>
                </div>
                <div style={{ height: 6, background: COLOR.hairline, borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ width: `${s.val}%`, height: "100%", background: COLOR.gold }} />
                </div>
              </div>
            ))}
          </div>
        </div>
      </Panel>

      {cur && (
        <Panel title="Comparaison période sur période" subtitle="Dernier mois filtré vs mois précédent / vs même mois l'an dernier">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16 }}>
            {[
              { label: "vs mois précédent", ref: prev },
              { label: "vs même mois -1 an", ref: yearAgo },
            ].map(({ label, ref }) => {
              const dRev = delta(cur.revenus, ref?.revenus);
              const dDep = delta(cur.depenses, ref?.depenses);
              return (
                <div key={label} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: 16 }}>
                  <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginBottom: 10 }}>{label}{ref ? ` (${ref.mois})` : ""}</div>
                  {ref ? (
                    <>
                      <DeltaRow label="Revenus" val={dRev} />
                      <DeltaRow label="Dépenses" val={dDep} inverse />
                    </>
                  ) : <div style={{ fontSize: 12, color: COLOR.inkMuted }}>Pas de données de référence</div>}
                </div>
              );
            })}
          </div>
        </Panel>
      )}

      {nwFiltered.length > 1 && (
        <PanelWithHelp title="Évolution de la valeur nette" subtitle="Relevé de solde — rapport séparé, non affecté par les filtres de type/groupe"
          explain="Cette courbe retrace le solde total de tes comptes mois par mois, tel qu'enregistré dans le relevé historique MoneyCoach puis mis à jour en direct par le solde réel de tes comptes actuels. Elle ne dépend pas des filtres Type/Groupe/Catégorie ci-dessus — seule la période (Du mois/Au mois) l'affecte. Une pente montante signifie que ton patrimoine augmente sur la période ; les creux correspondent souvent à de grosses dépenses ponctuelles (achat, investissement).">
          <ResponsiveContainer width="100%" height={260}>
            <AreaChart data={nwFiltered} margin={{ left: 0, right: 10, top: 10 }}>
              <defs>
                <linearGradient id="nwGrad" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={COLOR.gold} stopOpacity={0.35} />
                  <stop offset="100%" stopColor={COLOR.gold} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} interval={1} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <Area type="monotone" dataKey="valeur" name="Valeur nette" stroke={COLOR.goldSoft} strokeWidth={2} fill="url(#nwGrad)" />
            </AreaChart>
          </ResponsiveContainer>
        </PanelWithHelp>
      )}

      <PanelWithHelp title="Revenus vs Dépenses" subtitle={`${byMonth.length} mois dans la période filtrée`}
        explain="Chaque paire de barres compare, pour un même mois, le total encaissé (vert) au total dépensé (rouge). Quand la barre rouge dépasse la verte, ce mois-là a coûté plus qu'il n'a rapporté — pas forcément un problème si c'est un mois d'investissement ponctuel (achat, travaux), mais à surveiller si ça se répète plusieurs mois de suite.">
        {byMonth.length ? (
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={byMonth} margin={{ left: 0, right: 10, top: 10 }}>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} interval={byMonth.length > 14 ? 1 : 0} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 12, color: COLOR.inkMuted }} />
              <Bar dataKey="revenus" name="Revenus" fill={COLOR.emerald} radius={[3, 3, 0, 0]} />
              <Bar dataKey="depenses" name="Dépenses" fill={COLOR.clay} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
      </PanelWithHelp>

      <PanelWithHelp title="Répartition des dépenses par groupe" subtitle="Selon la classification des catégories"
        explain="Chaque dépense est classée dans l'un de ces 4 groupes : Nécessaire (logement, alimentation…), Productif (investissements, épargne, activité), Non-productif (cadeaux, sorties, shopping — sans retour), ou Non classifié (catégories comme 'Ajustement' pas encore triées). Un cercle doré épais indique une part élevée de dépenses non-productives par rapport au reste — c'est le principal levier si tu veux augmenter ton épargne.">
        {groupBreakdown.length ? (
          <div style={{ display: "flex", gap: 28, flexWrap: "wrap", alignItems: "center" }}>
            <ResponsiveContainer width={240} height={240}>
              <PieChart>
                <Pie data={groupBreakdown} dataKey="value" nameKey="name" innerRadius={58} outerRadius={96} paddingAngle={2}>
                  {groupBreakdown.map((g) => <Cell key={g.name} fill={groupColor[g.name]} stroke={COLOR.surface} strokeWidth={2} />)}
                </Pie>
                <Tooltip content={<CustomTooltip />} />
              </PieChart>
            </ResponsiveContainer>
            <div style={{ flex: 1, minWidth: 220 }}>
              {groupBreakdown.sort((a, b) => b.value - a.value).map((g) => (
                <div key={g.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ width: 10, height: 10, borderRadius: "50%", background: groupColor[g.name], display: "inline-block" }} />
                    <span style={{ fontSize: 13.5 }}>{g.name}</span>
                  </div>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}>{fmt(g.value)}</span>
                </div>
              ))}
            </div>
          </div>
        ) : <EmptyState />}
      </PanelWithHelp>
    </div>
  );
}

function DeltaRow({ label, val, inverse = false }: { label: string; val: number | null; inverse?: boolean }) {
  const good = val === null ? true : inverse ? val <= 0 : val >= 0;
  const color = val === null ? COLOR.inkMuted : good ? COLOR.emeraldSoft : COLOR.claySoft;
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0" }}>
      <span style={{ fontSize: 12.5 }}>{label}</span>
      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color, display: "flex", alignItems: "center", gap: 4 }}>
        {val !== null && (val >= 0 ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
        {val !== null ? `${val >= 0 ? "+" : ""}${val.toFixed(1)}%` : "—"}
      </span>
    </div>
  );
}

// ============================================================
// FLUX TAB (Sankey + Heatmap)
// ============================================================
function FluxTab({ filtered }: { filtered: any[] }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Flux Revenus → Dépenses" subtitle="Comment le revenu de la période se répartit entre les groupes de dépenses et le solde"
        explain="Le bloc vert à gauche représente 100% du revenu de la période. Chaque ruban qui en part montre la part qui va vers un groupe de dépenses (Nécessaire, Productif, Non-productif, Non classifié) ou vers le Solde (épargne restante). Plus un ruban est épais, plus ce poste absorbe une grande partie du revenu.">
        <FlowDiagram filtered={filtered} />
      </PanelWithHelp>
      <PanelWithHelp title="Calendrier d'intensité — dépenses non-productives" subtitle="Repérer les mois et saisons à risque"
        explain="Chaque case représente un mois. Plus la couleur est intense (vert → or → rouge), plus les dépenses non-productives (cadeaux, sorties, shopping) ont été élevées ce mois-là. Pratique pour repérer des périodes récurrentes à risque — fêtes de fin d'année, rentrée scolaire, anniversaires groupés…">
        <HeatmapCalendar filtered={filtered} />
      </PanelWithHelp>
    </div>
  );
}

// ============================================================ END OF PART 3 — continued below
// ============================================================
// RAPPORT MENSUEL TAB
// ============================================================
function MensuelTab({ filtered }: { filtered: any[] }) {
  const [sortKey, setSortKey] = useState<"mois" | "revenus" | "depenses" | "solde" | "nonProd">("mois");
  const [sortDir, setSortDir] = useState<1 | -1>(1);

  const rows = useMemo(() => {
    const m: Record<string, { revenus: number; depenses: number; nonProd: number }> = {};
    filtered.forEach((t) => {
      if (!m[t.month]) m[t.month] = { revenus: 0, depenses: 0, nonProd: 0 };
      if (t.type === "Revenu") m[t.month].revenus += t.amount;
      else { m[t.month].depenses += t.amount; if (t.group === "Non-productif") m[t.month].nonProd += t.amount; }
    });
    let arr = Object.keys(m).map((k) => ({ mois: k, label: monthLabel(k), revenus: m[k].revenus, depenses: m[k].depenses, solde: m[k].revenus - m[k].depenses, nonProd: m[k].nonProd }));
    arr.sort((a, b) => {
      let av: number = sortKey === "mois" ? monthSortKey(a.mois) : (a as any)[sortKey];
      let bv: number = sortKey === "mois" ? monthSortKey(b.mois) : (b as any)[sortKey];
      return (av - bv) * sortDir;
    });
    return arr;
  }, [filtered, sortKey, sortDir]);

  const chartData = useMemo(() => rows.slice().sort((a, b) => monthSortKey(a.mois) - monthSortKey(b.mois)), [rows]);
  const toggleSort = (k: typeof sortKey) => { if (sortKey === k) setSortDir((d) => (d === 1 ? -1 : 1)); else { setSortKey(k); setSortDir(1); } };
  const th = (label: string, key: typeof sortKey) => (
    <th onClick={() => toggleSort(key)} style={{ textAlign: key === "mois" ? "left" : "right", padding: "10px 12px", fontSize: 11, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", cursor: "pointer", userSelect: "none", borderBottom: `1px solid ${COLOR.hairline}`, whiteSpace: "nowrap" }}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>{label} <ArrowUpDown size={10} style={{ opacity: sortKey === key ? 1 : 0.35 }} /></span>
    </th>
  );
  const totals = rows.reduce((a, r) => ({ revenus: a.revenus + r.revenus, depenses: a.depenses + r.depenses, solde: a.solde + r.solde, nonProd: a.nonProd + r.nonProd }), { revenus: 0, depenses: 0, solde: 0, nonProd: 0 });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Revenus, dépenses et solde par mois" subtitle="Cliquer sur une colonne pour trier"
        explain="La ligne dorée (Solde net) trace la différence revenus−dépenses de chaque mois : au-dessus de zéro, le mois est excédentaire. La ligne rouge pointillée (Non-productif) montre le poids des dépenses sans retour (cadeaux, sorties…) — si elle suit de près ou dépasse le solde net, c'est souvent le premier poste à réduire pour améliorer l'épargne.">
        {chartData.length ? (
          <ResponsiveContainer width="100%" height={240}>
            <LineChart data={chartData} margin={{ left: 0, right: 10, top: 10 }}>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="label" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} interval={chartData.length > 14 ? 1 : 0} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 12, color: COLOR.inkMuted }} />
              <Line type="monotone" dataKey="solde" name="Solde net" stroke={COLOR.goldSoft} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="nonProd" name="Non-productif" stroke={COLOR.clay} strokeWidth={2} dot={false} strokeDasharray="4 3" />
            </LineChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
      </PanelWithHelp>
      <Panel title="Tableau mensuel" subtitle={`${rows.length} mois`}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'IBM Plex Mono', monospace" }}>
            <thead><tr>{th("Mois", "mois")}{th("Revenus", "revenus")}{th("Dépenses", "depenses")}{th("Solde", "solde")}{th("Non-productif", "nonProd")}</tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.mois}>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${COLOR.hairline}`, fontFamily: "'Inter', sans-serif" }}>{r.label}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, textAlign: "right", color: COLOR.emeraldSoft, borderBottom: `1px solid ${COLOR.hairline}` }}>{fmt(r.revenus)}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, textAlign: "right", color: COLOR.claySoft, borderBottom: `1px solid ${COLOR.hairline}` }}>{fmt(r.depenses)}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, textAlign: "right", color: r.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft, borderBottom: `1px solid ${COLOR.hairline}` }}>{fmt(r.solde)}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, textAlign: "right", color: COLOR.inkMuted, borderBottom: `1px solid ${COLOR.hairline}` }}>{fmt(r.nonProd)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr>
                <td style={{ padding: "10px 12px", fontSize: 12.5, fontWeight: 600, fontFamily: "'Inter', sans-serif" }}>Total</td>
                <td style={{ padding: "10px 12px", fontSize: 12.5, textAlign: "right", fontWeight: 600, color: COLOR.emeraldSoft }}>{fmt(totals.revenus)}</td>
                <td style={{ padding: "10px 12px", fontSize: 12.5, textAlign: "right", fontWeight: 600, color: COLOR.claySoft }}>{fmt(totals.depenses)}</td>
                <td style={{ padding: "10px 12px", fontSize: 12.5, textAlign: "right", fontWeight: 600, color: totals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(totals.solde)}</td>
                <td style={{ padding: "10px 12px", fontSize: 12.5, textAlign: "right", fontWeight: 600, color: COLOR.inkMuted }}>{fmt(totals.nonProd)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// CATÉGORIES TAB (reclassification + détection d'anomalies)
// ============================================================
function CategoriesTab({ filtered, categoryGroups, resolvedGroups, setCategoryGroups }: {
  filtered: any[]; categoryGroups: Record<string, Group>; resolvedGroups: Record<string, Group>; setCategoryGroups: (g: Record<string, Group>) => void;
}) {
  const [sortDir, setSortDir] = useState<1 | -1>(-1);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [selected, setSelected] = useState<{ name: string; type: TxType } | null>(null);
  const rows = useMemo(() => {
    const m: Record<string, { value: number; count: number; type: TxType; group: string }> = {};
    filtered.forEach((t) => {
      if (!m[t.category]) m[t.category] = { value: 0, count: 0, type: t.type, group: t.group };
      m[t.category].value += t.amount; m[t.category].count += 1;
    });
    return Object.entries(m).map(([name, d]) => ({ name, ...d })).sort((a, b) => (a.value - b.value) * sortDir);
  }, [filtered, sortDir]);
  const maxVal = Math.max(1, ...rows.map((r) => r.value));
  const total = rows.reduce((a, r) => a + r.value, 0);
  const anomalies = useMemo(() => detectAnomalies(filtered), [filtered]);

  const toggleExpand = (name: string) => setExpanded((prev) => { const n = new Set(prev); n.has(name) ? n.delete(name) : n.add(name); return n; });

  const subcatBreakdown = (categoryName: string, type: TxType) => {
    const m: Record<string, number> = {};
    filtered.filter((t) => t.category === categoryName && t.type === type).forEach((t) => {
      const key = t.subcategory || "Sans sous-catégorie";
      m[key] = (m[key] || 0) + t.amount;
    });
    return Object.entries(m).sort((a, b) => b[1] - a[1]);
  };

  if (selected) {
    return <CategoryDetailView category={selected.name} type={selected.type} filtered={filtered} onBack={() => setSelected(null)} />;
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {anomalies.length > 0 && (
        <Panel title="Transactions atypiques détectées" subtitle="Montant supérieur au double de la moyenne habituelle de la catégorie">
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {anomalies.map((a) => (
              <div key={a.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 12px", background: "rgba(193,84,63,0.08)", border: `1px solid ${COLOR.clay}`, borderRadius: 6 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <AlertTriangle size={14} color={COLOR.claySoft} />
                  <span style={{ fontSize: 12.5 }}>{monthLabel(a.month)} · {a.category}</span>
                </div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLOR.claySoft }}>
                  {fmt(a.amount)} <span style={{ color: COLOR.inkMuted }}>({a.ratio.toFixed(1)}× la moyenne de {fmt(a.avg)})</span>
                </div>
              </div>
            ))}
          </div>
        </Panel>
      )}
      <Panel title="Détail par catégorie" subtitle={`${rows.length} catégorie(s) · cliquez sur la loupe pour l'analyse complète d'une catégorie · le chevron déplie ses sous-catégories`}
        right={
          <button onClick={() => setSortDir((d) => (d === 1 ? -1 : 1))} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
            <ArrowUpDown size={12} /> {sortDir === -1 ? "Plus élevé d'abord" : "Plus faible d'abord"}
          </button>
        }>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {rows.map((r) => {
            const isOpen = expanded.has(r.name);
            const subs = isOpen ? subcatBreakdown(r.name, r.type) : [];
            const hasRealSubs = subs.some(([name]) => name !== "Sans sous-catégorie");
            return (
              <div key={r.name}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <button onClick={() => setSelected({ name: r.name, type: r.type })} title="Analyser cette catégorie" style={{ background: "transparent", border: "none", cursor: "pointer", display: "flex", flexShrink: 0, color: COLOR.goldSoft, padding: 2 }}>
                    <BarChart3 size={14} />
                  </button>
                  <div onClick={() => toggleExpand(r.name)} style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer", flex: 1, minWidth: 0 }}>
                    <ChevronRight size={13} color={COLOR.inkMuted} style={{ flexShrink: 0, transform: isOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                    <div style={{ width: 166, fontSize: 12.5, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.name}>{r.name}</div>
                    <div style={{ flex: 1, background: COLOR.hairline, borderRadius: 4, height: 16, position: "relative" }}>
                      <div style={{ width: `${(r.value / maxVal) * 100}%`, height: "100%", borderRadius: 4, background: groupColor[r.group] || COLOR.inkMuted }} />
                    </div>
                    <div style={{ width: 95, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, flexShrink: 0 }}>{fmt(r.value)}</div>
                    <div style={{ width: 42, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: COLOR.inkMuted, flexShrink: 0 }}>{total ? ((r.value / total) * 100).toFixed(1) : "0"}%</div>
                  </div>
                  {r.type === "Dépense" ? (
                    <select value={resolvedGroups[r.name] || "Non classifié"} onClick={(e) => e.stopPropagation()} onChange={(e) => setCategoryGroups({ ...categoryGroups, [r.name]: e.target.value as Group })}
                      style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: groupColor[resolvedGroups[r.name] || "Non classifié"], padding: "5px 8px", fontSize: 11.5, fontFamily: "'Inter', sans-serif", flexShrink: 0, width: 128, cursor: "pointer" }}>
                      {GROUPS.map((g) => <option key={g} value={g}>{g}</option>)}
                    </select>
                  ) : <div style={{ width: 128, flexShrink: 0, fontSize: 11.5, color: COLOR.goldSoft, textAlign: "center" }}>Revenu</div>}
                </div>
                {isOpen && (
                  <div style={{ marginLeft: 47, marginTop: 8, marginBottom: 4, paddingLeft: 12, borderLeft: `2px solid ${COLOR.hairline}`, display: "flex", flexDirection: "column", gap: 5 }}>
                    {hasRealSubs ? subs.map(([name, val]) => (
                      <div key={name} style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5 }}>
                        <span style={{ color: name === "Sans sous-catégorie" ? COLOR.inkMuted : COLOR.ink, fontStyle: name === "Sans sous-catégorie" ? "italic" : "normal" }}>{name}</span>
                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.inkMuted }}>{fmt(val)}</span>
                      </div>
                    )) : (
                      <span style={{ fontSize: 11.5, color: COLOR.inkMuted, fontStyle: "italic" }}>Aucune sous-catégorie renseignée sur ces transactions</span>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          {!rows.length && <EmptyState />}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// FICHE DÉTAILLÉE PAR CATÉGORIE — évolution, sous-catégories, analyse, impression
// ============================================================
function CategoryDetailView({ category, type, filtered, onBack }: { category: string; type: TxType; filtered: any[]; onBack: () => void }) {
  const catTx = useMemo(() => filtered.filter((t) => t.category === category && t.type === type), [filtered, category, type]);
  const group = catTx[0]?.group || "Non classifié";
  const total = catTx.reduce((a, t) => a + t.amount, 0);
  const monthsPresent = Array.from(new Set(catTx.map((t) => t.month)));
  const avgMonth = total / (monthsPresent.length || 1);
  const grandTotal = filtered.filter((t) => t.type === type).reduce((a, t) => a + t.amount, 0);
  const pctOfTotal = grandTotal ? (total / grandTotal) * 100 : 0;

  const byMonth = useMemo(() => {
    const m: Record<string, number> = {};
    catTx.forEach((t) => { m[t.month] = (m[t.month] || 0) + t.amount; });
    return Object.keys(m).sort((a, b) => monthSortKey(a) - monthSortKey(b)).map((k) => ({ mois: monthLabel(k), key: k, montant: m[k] }));
  }, [catTx]);

  const bySubcat = useMemo(() => {
    const m: Record<string, number> = {};
    catTx.forEach((t) => { const key = t.subcategory || "Sans sous-catégorie"; m[key] = (m[key] || 0) + t.amount; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
  }, [catTx]);

  const catAnomalies = useMemo(() => {
    if (catTx.length < 3) return [];
    const m = mean(catTx.map((t) => t.amount));
    return catTx.filter((t) => t.amount > m * 2 && t.amount > 20000).sort((a, b) => b.amount - a.amount).slice(0, 5).map((t) => ({ ...t, avg: m, ratio: t.amount / m }));
  }, [catTx]);

  const trend = useMemo(() => {
    if (byMonth.length < 4) return null;
    const mid = Math.floor(byMonth.length / 2);
    const first = byMonth.slice(0, mid).reduce((a, m) => a + m.montant, 0) / mid;
    const second = byMonth.slice(mid).reduce((a, m) => a + m.montant, 0) / (byMonth.length - mid);
    if (first === 0) return null;
    return ((second - first) / first) * 100;
  }, [byMonth]);

  const maxSub = Math.max(1, ...bySubcat.map((s) => s.value));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div className="gl-noprint" style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <button onClick={onBack} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
          ← Retour aux catégories
        </button>
        <button onClick={() => window.print()} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, color: COLOR.goldSoft, padding: "8px 16px", fontSize: 13, cursor: "pointer" }}>
          <Printer size={15} /> Imprimer cette fiche
        </button>
      </div>

      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ fontSize: 11, color: groupColor[type === "Revenu" ? "Revenu" : group], textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>{type === "Revenu" ? "Revenu" : group}</div>
            <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 26, margin: 0 }}>{category}</h2>
          </div>
        </div>
      </Panel>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Total (période)" value={fmt(total)} tone={type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} />
        <Kpi label="Moyenne / mois" value={fmt(avgMonth)} tone={COLOR.gold} icon={CalendarRange} />
        <Kpi label="Transactions" value={String(catTx.length)} suffix="" icon={BookOpen} />
        <Kpi label={`% du total ${type === "Revenu" ? "revenus" : "dépenses"}`} value={pctOfTotal.toFixed(1)} suffix="%" tone={COLOR.goldSoft} icon={Percent} />
      </div>

      <PanelWithHelp title="Évolution mensuelle" subtitle={`${byMonth.length} mois avec activité`}
        explain={`Chaque barre montre le total dépensé (ou reçu) en "${category}" ce mois-là. Une hausse progressive peut indiquer une habitude qui s'installe ; des pics isolés correspondent souvent à des achats ponctuels plutôt qu'à une tendance de fond.`}>
        {byMonth.length ? (
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={byMonth} margin={{ left: 0, right: 10, top: 10 }}>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} interval={byMonth.length > 14 ? 1 : 0} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <Bar dataKey="montant" name={category} radius={[3, 3, 0, 0]} fill={type === "Revenu" ? COLOR.emerald : groupColor[group] || COLOR.clay} />
            </BarChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
      </PanelWithHelp>

      {bySubcat.length > 1 && (
        <Panel title="Répartition par sous-catégorie">
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {bySubcat.map((s) => (
              <div key={s.name} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ width: 150, fontSize: 12, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: s.name === "Sans sous-catégorie" ? COLOR.inkMuted : COLOR.ink, fontStyle: s.name === "Sans sous-catégorie" ? "italic" : "normal" }}>{s.name}</div>
                <div style={{ flex: 1, background: COLOR.hairline, borderRadius: 4, height: 14 }}>
                  <div style={{ width: `${(s.value / maxSub) * 100}%`, height: "100%", borderRadius: 4, background: type === "Revenu" ? COLOR.emerald : groupColor[group] || COLOR.clay }} />
                </div>
                <div style={{ width: 90, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>{fmt(s.value)}</div>
              </div>
            ))}
          </div>
        </Panel>
      )}

      <Panel title="Analyse de cette catégorie" subtitle="Constats spécifiques générés automatiquement">
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {trend !== null && (
            <div style={{ display: "flex", gap: 10, padding: "12px 14px", background: trend > 20 ? "rgba(193,84,63,0.08)" : trend < -20 ? "rgba(63,156,122,0.08)" : "rgba(201,162,39,0.06)", border: `1px solid ${trend > 20 ? COLOR.clay : trend < -20 ? COLOR.emerald : COLOR.hairline}`, borderRadius: 8 }}>
              {trend > 20 ? <TrendingUp size={15} color={COLOR.claySoft} style={{ flexShrink: 0, marginTop: 1 }} /> : trend < -20 ? <TrendingDown size={15} color={COLOR.emeraldSoft} style={{ flexShrink: 0, marginTop: 1 }} /> : <Info size={15} color={COLOR.goldSoft} style={{ flexShrink: 0, marginTop: 1 }} />}
              <div style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.55 }}>
                {trend > 20 && <>Tendance à la hausse : +{trend.toFixed(0)}% entre la première et la seconde moitié de la période.</>}
                {trend < -20 && <>Tendance à la baisse : {trend.toFixed(0)}% entre la première et la seconde moitié de la période.</>}
                {trend >= -20 && trend <= 20 && <>Dépense relativement stable sur la période (variation de {trend.toFixed(0)}% entre les deux moitiés).</>}
              </div>
            </div>
          )}
          {catAnomalies.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {catAnomalies.map((a) => (
                <div key={a.id} style={{ display: "flex", justifyContent: "space-between", padding: "10px 14px", background: "rgba(193,84,63,0.08)", border: `1px solid ${COLOR.clay}`, borderRadius: 8, fontSize: 12 }}>
                  <span style={{ color: COLOR.inkMuted }}>{dateLabelFull(a.date)}{a.subcategory ? ` · ${a.subcategory}` : ""}</span>
                  <span style={{ color: COLOR.claySoft, fontFamily: "'IBM Plex Mono', monospace" }}>{fmt(a.amount)} ({a.ratio.toFixed(1)}×)</span>
                </div>
              ))}
            </div>
          )}
          {trend === null && !catAnomalies.length && (
            <span style={{ fontSize: 12.5, color: COLOR.inkMuted, fontStyle: "italic" }}>Pas assez de données sur la période pour dégager une tendance ou détecter des anomalies.</span>
          )}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// GROUPES TAB
// ============================================================
function GroupesTab({ filtered }: { filtered: any[] }) {
  const dep = filtered.filter((t) => t.type === "Dépense");
  const rev = filtered.filter((t) => t.type === "Revenu");
  const totalDep = dep.reduce((a, t) => a + t.amount, 0);
  const totalRev = rev.reduce((a, t) => a + t.amount, 0);
  const solde = totalRev - totalDep;
  const tauxEpargne = totalRev > 0 ? (solde / totalRev) * 100 : 0;
  const cards = GROUPS.map((g) => {
    const items = dep.filter((t) => t.group === g);
    const value = items.reduce((a, t) => a + t.amount, 0);
    const cats: Record<string, number> = {};
    items.forEach((t) => { cats[t.category] = (cats[t.category] || 0) + t.amount; });
    const top = Object.entries(cats).sort((a, b) => b[1] - a[1]).slice(0, 4);
    return { group: g, value, count: items.length, top, pct: totalDep ? (value / totalDep) * 100 : 0 };
  });

  const byMonth = useMemo(() => {
    const m: Record<string, { revenus: number; depenses: number }> = {};
    filtered.forEach((t) => {
      if (!m[t.month]) m[t.month] = { revenus: 0, depenses: 0 };
      if (t.type === "Revenu") m[t.month].revenus += t.amount; else m[t.month].depenses += t.amount;
    });
    return Object.keys(m).sort((a, b) => monthSortKey(a) - monthSortKey(b)).map((k) => ({ mois: monthLabel(k), revenus: m[k].revenus, depenses: m[k].depenses }));
  }, [filtered]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Revenus" value={fmt(totalRev)} tone={COLOR.emeraldSoft} icon={TrendingUp} />
        <Kpi label="Dépenses" value={fmt(totalDep)} tone={COLOR.claySoft} icon={TrendingDown} />
        <Kpi label="Solde" value={fmt(solde)} tone={solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} />
        <Kpi label="Taux d'épargne" value={tauxEpargne.toFixed(1)} suffix="%" tone={COLOR.gold} icon={Target} />
      </div>

      <PanelWithHelp title="Revenus vs Dépenses" subtitle={`${byMonth.length} mois dans la période filtrée`}
        explain="Chaque paire de barres compare, pour un même mois, le total encaissé (vert) au total dépensé (rouge). Utile pour repérer d'un coup d'œil les mois où les dépenses ont dépassé les revenus.">
        {byMonth.length ? (
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={byMonth} margin={{ left: 0, right: 10, top: 10 }}>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} interval={byMonth.length > 14 ? 1 : 0} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 12, color: COLOR.inkMuted }} />
              <Bar dataKey="revenus" name="Revenus" fill={COLOR.emerald} radius={[3, 3, 0, 0]} />
              <Bar dataKey="depenses" name="Dépenses" fill={COLOR.clay} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
      </PanelWithHelp>

      <InsightsPanel filtered={filtered} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16 }}>
        {cards.map((c) => (
          <div key={c.group} style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ width: 10, height: 10, borderRadius: "50%", background: groupColor[c.group] }} />
                <span style={{ fontSize: 14, fontWeight: 500 }}>{c.group}</span>
              </div>
              <span style={{ fontSize: 11, color: COLOR.inkMuted }}>{c.count} transactions</span>
            </div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 600, color: groupColor[c.group] }}>
              {fmt(c.value)} <span style={{ fontSize: 11, color: COLOR.inkMuted }}>FCFA · {c.pct.toFixed(1)}%</span>
            </div>
            <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 6 }}>
              {c.top.map(([name, val]) => (
                <div key={name} style={{ display: "flex", justifyContent: "space-between", fontSize: 12 }}>
                  <span style={{ color: COLOR.inkMuted }}>{name}</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmt(val)}</span>
                </div>
              ))}
              {!c.top.length && <span style={{ fontSize: 12, color: COLOR.inkMuted }}>Aucune donnée</span>}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ============================================================
// COMPARATIF ANNUEL TAB
// ============================================================
function ComparatifTab({ transactions, categoryGroups }: { transactions: Transaction[]; categoryGroups: Record<string, Group> }) {
  const years = Array.from(new Set(transactions.map((t) => t.date.slice(0, 4)))).sort();
  const monthCounts: Record<string, number> = {};
  years.forEach((y) => {
    monthCounts[y] = new Set(transactions.filter((t) => t.date.startsWith(y)).map((t) => t.date.slice(0, 7))).size || 1;
  });

  const byYearCategory = useMemo(() => {
    const m: Record<string, Record<string, number>> = {};
    transactions.filter((t) => t.type === "Dépense").forEach((t) => {
      const y = t.date.slice(0, 4);
      m[t.category] = m[t.category] || {};
      m[t.category][y] = (m[t.category][y] || 0) + t.amount;
    });
    return m;
  }, [transactions]);

  const topCategories = useMemo(() => {
    const totals = Object.entries(byYearCategory).map(([name, y]) => ({ name, total: Object.values(y).reduce((a, b) => a + b, 0) }));
    return totals.sort((a, b) => b.total - a.total).slice(0, 10).map((t) => t.name);
  }, [byYearCategory]);

  const chartData = topCategories.map((cat) => {
    const row: any = { name: cat };
    years.forEach((y) => { row[y] = byYearCategory[cat]?.[y] || 0; });
    return row;
  });

  const yearTotals = years.map((y) => {
    const dep = transactions.filter((t) => t.type === "Dépense" && t.date.startsWith(y)).reduce((a, t) => a + t.amount, 0);
    const rev = transactions.filter((t) => t.type === "Revenu" && t.date.startsWith(y)).reduce((a, t) => a + t.amount, 0);
    return { year: y, dep, rev, depPerMonth: dep / monthCounts[y], revPerMonth: rev / monthCounts[y] };
  });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Comparatif annuel — moyenne mensuelle" subtitle="2024 et 2026 sont des années partielles ; la comparaison se fait donc par moyenne mensuelle, pas par total brut">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16 }}>
          {yearTotals.map((y) => (
            <div key={y.year} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: 16 }}>
              <div style={{ fontFamily: "'Fraunces', serif", fontSize: 18, marginBottom: 10 }}>{y.year} <span style={{ fontSize: 11, color: COLOR.inkMuted, fontFamily: "'Inter', sans-serif" }}>({monthCounts[y.year]} mois)</span></div>
              <div style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Revenu moyen/mois</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, color: COLOR.emeraldSoft, marginBottom: 8 }}>{fmt(y.revPerMonth)}</div>
              <div style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Dépense moyenne/mois</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, color: COLOR.claySoft }}>{fmt(y.depPerMonth)}</div>
            </div>
          ))}
        </div>
      </Panel>
      <PanelWithHelp title="Top 10 catégories — comparaison par année" subtitle="Totaux bruts par catégorie et par année"
        explain="Pour chacune de tes 10 catégories de dépenses les plus importantes, une barre par année (bleu=2024, or=2025, vert=2026…) montre le total dépensé. Comme 2024 et l'année en cours sont partielles, compare plutôt les tendances/ordres de grandeur que les totaux bruts — le panneau du dessus donne les moyennes mensuelles, plus justes pour ce type de comparaison.">
        <ResponsiveContainer width="100%" height={420}>
          <BarChart data={chartData} layout="vertical" margin={{ left: 10, right: 20, top: 10 }}>
            <CartesianGrid stroke={COLOR.hairline} horizontal={false} />
            <XAxis type="number" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
            <YAxis type="category" dataKey="name" tick={{ fill: COLOR.inkMuted, fontSize: 10.5 }} axisLine={false} tickLine={false} width={150} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ fontSize: 12, color: COLOR.inkMuted }} />
            {years.map((y, i) => (
              <Bar key={y} dataKey={y} fill={[COLOR.slateBlue, COLOR.gold, COLOR.emerald, COLOR.violet, COLOR.clay][i % 5]} radius={[0, 3, 3, 0]} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </PanelWithHelp>
    </div>
  );
}

// ============================================================
// COMPARATEUR — deux périodes personnalisées, côte à côte
// ============================================================
function ComparateurTab({ transactions, categoryGroups, allMonths }: {
  transactions: Transaction[]; categoryGroups: Record<string, Group>; allMonths: string[];
}) {
  const withGroup = useMemo(
    () => transactions.map((t) => ({ ...t, month: dateToMonthKey(t.date), group: t.type === "Revenu" ? "Revenu" : (categoryGroups[t.category] || "Non classifié") })),
    [transactions, categoryGroups]
  );

  const lastMonth = allMonths[allMonths.length - 1] || "";
  const idxLast = allMonths.indexOf(lastMonth);
  const prevMonth = idxLast > 0 ? allMonths[idxLast - 1] : lastMonth;

  const [aFrom, setAFrom] = useState(prevMonth);
  const [aTo, setATo] = useState(prevMonth);
  const [bFrom, setBFrom] = useState(lastMonth);
  const [bTo, setBTo] = useState(lastMonth);

  const statsFor = (from: string, to: string) => {
    const fk = monthSortKey(from), tk = monthSortKey(to);
    const tx = withGroup.filter((t) => { const k = monthSortKey(t.month); return k >= fk && k <= tk; });
    const revenus = tx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
    const depenses = tx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
    const nonProd = tx.filter((t) => t.type === "Dépense" && t.group === "Non-productif").reduce((a, t) => a + t.amount, 0);
    const solde = revenus - depenses;
    const tauxEpargne = revenus > 0 ? (solde / revenus) * 100 : 0;
    const byCat: Record<string, number> = {};
    tx.filter((t) => t.type === "Dépense").forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
    return { revenus, depenses, solde, tauxEpargne, nonProd, byCat, count: tx.length };
  };

  const A = useMemo(() => statsFor(aFrom, aTo), [withGroup, aFrom, aTo]);
  const B = useMemo(() => statsFor(bFrom, bTo), [withGroup, bFrom, bTo]);

  const pctDelta = (a: number, b: number) => (a !== 0 ? ((b - a) / Math.abs(a)) * 100 : (b !== 0 ? 100 : 0));

  const chartData = [
    { name: "Revenus", A: A.revenus, B: B.revenus },
    { name: "Dépenses", A: A.depenses, B: B.depenses },
    { name: "Solde", A: A.solde, B: B.solde },
    { name: "Non-productif", A: A.nonProd, B: B.nonProd },
  ];

  const allCats = Array.from(new Set([...Object.keys(A.byCat), ...Object.keys(B.byCat)]));
  const catRows = allCats
    .map((c) => ({ cat: c, a: A.byCat[c] || 0, b: B.byCat[c] || 0, delta: (B.byCat[c] || 0) - (A.byCat[c] || 0) }))
    .filter((r) => r.a !== 0 || r.b !== 0)
    .sort((x, y) => Math.abs(y.delta) - Math.abs(x.delta))
    .slice(0, 15);

  const metricRow = (label: string, a: number, b: number, lowerIsBetter: boolean, suffix = "FCFA") => {
    const d = pctDelta(a, b);
    const good = d === 0 ? null : lowerIsBetter ? d < 0 : d > 0;
    return (
      <div style={{ display: "flex", alignItems: "center", padding: "12px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
        <div style={{ width: 150, fontSize: 12.5, color: COLOR.inkMuted }}>{label}</div>
        <div style={{ flex: 1, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.slateBlueSoft }}>{suffix === "%" ? a.toFixed(1) + "%" : fmt(a)}</div>
        <div style={{ width: 30, textAlign: "center", color: COLOR.inkMuted }}>→</div>
        <div style={{ flex: 1, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.goldSoft }}>{suffix === "%" ? b.toFixed(1) + "%" : fmt(b)}</div>
        <div style={{ width: 90, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: good === null ? COLOR.inkMuted : good ? COLOR.emeraldSoft : COLOR.claySoft }}>
          {d >= 0 ? "+" : ""}{d.toFixed(0)}%
        </div>
      </div>
    );
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Choisir les deux périodes à comparer">
        <div style={{ display: "flex", gap: 30, flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: COLOR.slateBlue }} />
              <span style={{ fontSize: 12.5, color: COLOR.slateBlueSoft, fontWeight: 600 }}>Période A</span>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <Select label="Du mois" value={aFrom} onChange={setAFrom} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
              <Select label="Au mois" value={aTo} onChange={setATo} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
            </div>
          </div>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: COLOR.gold }} />
              <span style={{ fontSize: 12.5, color: COLOR.goldSoft, fontWeight: 600 }}>Période B</span>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <Select label="Du mois" value={bFrom} onChange={setBFrom} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
              <Select label="Au mois" value={bTo} onChange={setBTo} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
            </div>
          </div>
        </div>
      </Panel>

      <Panel title="Comparaison des indicateurs clés" subtitle={`A : ${monthLabel(aFrom)} — ${monthLabel(aTo)} (${A.count} tx)  ·  B : ${monthLabel(bFrom)} — ${monthLabel(bTo)} (${B.count} tx)`}>
        <div style={{ display: "flex", padding: "4px 0", fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>
          <div style={{ width: 150 }}></div>
          <div style={{ flex: 1, textAlign: "right" }}>A</div>
          <div style={{ width: 30 }}></div>
          <div style={{ flex: 1, textAlign: "right" }}>B</div>
          <div style={{ width: 90, textAlign: "right" }}>Évolution</div>
        </div>
        {metricRow("Revenus", A.revenus, B.revenus, false)}
        {metricRow("Dépenses", A.depenses, B.depenses, true)}
        {metricRow("Solde", A.solde, B.solde, false)}
        {metricRow("Taux d'épargne", A.tauxEpargne, B.tauxEpargne, false, "%")}
        {metricRow("Non-productif", A.nonProd, B.nonProd, true)}
      </Panel>

      <PanelWithHelp title="Comparaison visuelle" explain="Chaque paire de barres met côte à côte la Période A (bleu) et la Période B (or) pour Revenus, Dépenses, Solde et Non-productif. Ça permet de voir en un coup d'œil quels indicateurs ont le plus bougé entre les deux périodes choisies ci-dessus.">
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={chartData} margin={{ left: 0, right: 10, top: 10 }}>
            <CartesianGrid stroke={COLOR.hairline} vertical={false} />
            <XAxis dataKey="name" tick={{ fill: COLOR.inkMuted, fontSize: 11 }} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
            <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ fontSize: 12, color: COLOR.inkMuted }} />
            <Bar dataKey="A" name={`A (${monthLabel(aFrom)}${aFrom !== aTo ? "…" + monthLabel(aTo) : ""})`} fill={COLOR.slateBlue} radius={[3, 3, 0, 0]} />
            <Bar dataKey="B" name={`B (${monthLabel(bFrom)}${bFrom !== bTo ? "…" + monthLabel(bTo) : ""})`} fill={COLOR.gold} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </PanelWithHelp>

      <Panel title="Catégories qui ont le plus évolué" subtitle="Triées par variation absolue entre A et B">
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {catRows.map((r) => (
            <div key={r.cat} style={{ display: "flex", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <div style={{ flex: 1, fontSize: 12.5, color: COLOR.ink }}>{r.cat}</div>
              <div style={{ width: 90, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLOR.slateBlueSoft }}>{fmt(r.a)}</div>
              <div style={{ width: 20, textAlign: "center", color: COLOR.inkMuted }}>→</div>
              <div style={{ width: 90, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLOR.goldSoft }}>{fmt(r.b)}</div>
              <div style={{ width: 100, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: r.delta > 0 ? COLOR.claySoft : COLOR.emeraldSoft }}>
                {r.delta >= 0 ? "+" : ""}{fmt(r.delta)}
              </div>
            </div>
          ))}
          {!catRows.length && <EmptyState text="Aucune dépense sur ces deux périodes." />}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// PRINCIPALES CATÉGORIES — anneau coloré, pastilles de mois, comparaison
// ============================================================
const DONUT_COLORS = [COLOR.emerald, COLOR.slateBlue, COLOR.gold, COLOR.violet, COLOR.clay, COLOR.emeraldSoft, COLOR.slateBlueSoft, COLOR.goldSoft];

function TopCategoriesTab({ transactions, setTransactions, categoryGroups, allMonths, accounts }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; categoryGroups: Record<string, Group>; allMonths: string[]; accounts: Account[];
}) {
  const withGroup = useMemo(
    () => transactions.map((t) => ({ ...t, month: dateToMonthKey(t.date), group: t.type === "Revenu" ? "Revenu" : (categoryGroups[t.category] || "Non classifié") })),
    [transactions, categoryGroups]
  );

  const lastMonth = allMonths[allMonths.length - 1] || "";
  const pillMonths = allMonths;
  const [selectedMonth, setSelectedMonth] = useState(lastMonth);
  const [customOpen, setCustomOpen] = useState(false);
  const [customFrom, setCustomFrom] = useState(lastMonth);
  const [customTo, setCustomTo] = useState(lastMonth);
  const [typeView, setTypeView] = useState<TxType>("Dépense");
  const [expandedCat, setExpandedCat] = useState<string | null>(null);
  const [expandedSub, setExpandedSub] = useState<string | null>(null);
  const [editingTx, setEditingTx] = useState<Transaction | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const pillScrollRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    pillScrollRef.current?.scrollTo({ left: pillScrollRef.current.scrollWidth, behavior: "auto" });
  }, []);

  useEffect(() => { setExpandedCat(null); setExpandedSub(null); }, [typeView, selectedMonth, customFrom, customTo, customOpen]);

  const subcatFor = (catName: string) => {
    const rows: Record<string, number> = {};
    periodTx.filter((t) => t.category === catName).forEach((t) => {
      const key = t.subcategory || "Sans sous-catégorie";
      rows[key] = (rows[key] || 0) + t.amount;
    });
    const catTotal = byCat[catName] || 0;
    return Object.entries(rows)
      .map(([name, value]) => ({ name, value, pct: catTotal ? (value / catTotal) * 100 : 0 }))
      .sort((a, b) => b.value - a.value);
  };
  const txFor = (catName: string, subName: string) =>
    periodTx
      .filter((t) => t.category === catName && (t.subcategory || "Sans sous-catégorie") === subName)
      .sort((a, b) => b.date.localeCompare(a.date));

  const startEdit = (t: Transaction) => setEditingTx(t);
  const saveEdit = (t: Transaction) => setTransactions(transactions.map((x) => (x.id === t.id ? t : x)));
  const removeTx = (id: string) => setTransactions(transactions.filter((t) => t.id !== id));

  const range = customOpen ? { from: customFrom, to: customTo } : { from: selectedMonth, to: selectedMonth };
  const fk = monthSortKey(range.from), tk = monthSortKey(range.to);
  const periodTx = withGroup.filter((t) => { const k = monthSortKey(t.month); return k >= fk && k <= tk && t.type === typeView; });
  const total = periodTx.reduce((a, t) => a + t.amount, 0);

  // Période précédente de même longueur, immédiatement avant
  const spanMonths = Math.max(1, tk - fk + 1);
  const prevTo = fk - 1, prevFrom = fk - spanMonths;
  const prevTx = withGroup.filter((t) => { const k = monthSortKey(t.month); return k >= prevFrom && k <= prevTo && t.type === typeView; });
  const prevTotal = prevTx.reduce((a, t) => a + t.amount, 0);
  const delta = total - prevTotal;
  const deltaPct = prevTotal !== 0 ? (delta / prevTotal) * 100 : (total !== 0 ? 100 : 0);
  const improved = typeView === "Dépense" ? delta <= 0 : delta >= 0;

  const periodAllTypesTx = withGroup.filter((t) => { const k = monthSortKey(t.month); return k >= fk && k <= tk; });
  const prevByCat: Record<string, number> = {};
  prevTx.forEach((t) => { prevByCat[t.category] = (prevByCat[t.category] || 0) + t.amount; });

  const byCat: Record<string, number> = {};
  periodTx.forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
  const catList = Object.entries(byCat)
    .map(([name, value]) => ({ name, value, pct: total ? (value / total) * 100 : 0 }))
    .sort((a, b) => b.value - a.value);
  const donutData = catList.map((c, i) => ({ ...c, color: DONUT_COLORS[i % DONUT_COLORS.length] }));

  const periodLabel = customOpen
    ? (customFrom === customTo ? monthLabel(customFrom) : `${monthLabel(customFrom)} — ${monthLabel(customTo)}`)
    : monthLabel(selectedMonth);

  const [pdfState, setPdfState] = useState<"idle" | "loading" | "error">("idle");

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const summaryRows: any[][] = [
      ["Grand Livre — Principales catégories"],
      ["Période", periodLabel],
      ["Type", typeView],
      ["Généré le", dateLabelFull(todayISO())],
      [],
      [`Total ${typeView === "Dépense" ? "dépenses" : "revenus"}`, total],
      ["Vs. période précédente (FCFA)", delta],
      ["Vs. période précédente (%)", `${delta >= 0 ? "+" : "−"}${Math.abs(deltaPct).toFixed(1)}%`],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    wsSummary["!cols"] = [{ wch: 30 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Résumé");

    const catHeader = ["Catégorie", "Montant (FCFA)", "% du total"];
    const catRows = catList.map((c) => [c.name, c.value, `${c.pct.toFixed(1)}%`]);
    const wsCat = XLSX.utils.aoa_to_sheet([catHeader, ...catRows]);
    wsCat["!cols"] = [{ wch: 26 }, { wch: 16 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, wsCat, "Par catégorie");

    const subHeader = ["Catégorie", "Sous-catégorie", "Montant (FCFA)", "% de la catégorie"];
    const subRows: any[][] = [];
    catList.forEach((c) => { subcatFor(c.name).forEach((s) => subRows.push([c.name, s.name, s.value, `${s.pct.toFixed(1)}%`])); });
    const wsSub = XLSX.utils.aoa_to_sheet([subHeader, ...subRows]);
    wsSub["!cols"] = [{ wch: 26 }, { wch: 22 }, { wch: 16 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsSub, "Par sous-catégorie");

    XLSX.writeFile(wb, `grand-livre_principales-categories_${periodLabel.replace(/\s/g, "-")}.xlsx`);
  };

  const exportPDF = async () => {
    setPdfState("loading");
    try {
      const [{ default: jsPDF }, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      const autoTable = (autoTableModule as any).default || autoTableModule;
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();

      doc.setFillColor(26, 43, 76);
      doc.rect(0, 0, pageWidth, 34, "F");
      doc.setFontSize(17); doc.setTextColor(255, 255, 255);
      doc.text("Grand Livre — Principales catégories", 14, 16);
      doc.setFontSize(9.5); doc.setTextColor(200, 210, 225);
      doc.text(`${typeView} · ${periodLabel} · Généré le ${dateLabelFull(todayISO())}`, 14, 24);
      doc.text(`${catList.length} catégorie(s)`, 14, 29);

      const kpiY = 40, kpiH = 20, kpiW = (pageWidth - 28 - 16) / 3;
      const drawKpiBox = (x: number, label: string, value: string, r: number, g: number, b: number) => {
        doc.setFillColor(r, g, b);
        doc.roundedRect(x, kpiY, kpiW, kpiH, 2, 2, "F");
        doc.setFontSize(7.5); doc.setTextColor(255, 255, 255);
        doc.text(label, x + 5, kpiY + 7);
        doc.setFontSize(11); doc.setFont("helvetica", "bold");
        doc.text(value, x + 5, kpiY + 15);
        doc.setFont("helvetica", "normal");
      };
      drawKpiBox(14, `TOTAL ${typeView.toUpperCase()}`, `${fmt(total)} FCFA`, typeView === "Revenu" ? 63 : 193, typeView === "Revenu" ? 156 : 84, typeView === "Revenu" ? 122 : 63);
      drawKpiBox(14 + kpiW + 8, "VS PÉRIODE PRÉC.", `${delta >= 0 ? "+" : "−"}${fmt(Math.abs(delta))}`, improved ? 63 : 193, improved ? 156 : 84, improved ? 122 : 63);
      drawKpiBox(14 + (kpiW + 8) * 2, "VARIATION", `${delta >= 0 ? "+" : "−"}${Math.abs(deltaPct).toFixed(0)}%`, improved ? 63 : 193, improved ? 156 : 84, improved ? 122 : 63);

      autoTable(doc, {
        startY: 68,
        head: [["Catégorie", "Montant (FCFA)", "% du total"]],
        body: catList.map((c) => [c.name, fmt(c.value), `${c.pct.toFixed(0)}%`]),
        headStyles: { fillColor: [26, 43, 76] },
        styles: { fontSize: 8.5 },
        columnStyles: { 1: { halign: "right" }, 2: { halign: "right" } },
      });

      let y = (doc as any).lastAutoTable.finalY + 10;
      catList.forEach((c) => {
        const subs = subcatFor(c.name);
        if (!subs.length) return;
        if (y > 260) { doc.addPage(); y = 20; }
        doc.setFontSize(10); doc.setTextColor(26, 43, 76); doc.setFont("helvetica", "bold");
        doc.text(c.name, 14, y);
        doc.setFont("helvetica", "normal");
        autoTable(doc, {
          startY: y + 3,
          head: [["Sous-catégorie", "Montant (FCFA)", "% de la catégorie"]],
          body: subs.map((s) => [s.name, fmt(s.value), `${s.pct.toFixed(0)}%`]),
          headStyles: { fillColor: [201, 162, 39], textColor: [26, 26, 26] },
          styles: { fontSize: 8 },
          columnStyles: { 1: { halign: "right" }, 2: { halign: "right" } },
          margin: { left: 14, right: 14 },
        });
        y = (doc as any).lastAutoTable.finalY + 10;
      });

      doc.save(`grand-livre_principales-categories_${periodLabel.replace(/\s/g, "-")}.pdf`);
      setPdfState("idle");
    } catch {
      setPdfState("error");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel>
        <div ref={pillScrollRef} className="gl-noprint gl-scroll" style={{ display: "flex", gap: 8, flexWrap: "nowrap", overflowX: "auto", marginBottom: 20, paddingBottom: 4, WebkitOverflowScrolling: "touch" }}>
          {pillMonths.map((m) => (
            <button key={m} onClick={() => { setSelectedMonth(m); setCustomOpen(false); }} style={{
              padding: "8px 18px", borderRadius: 20, fontSize: 13, cursor: "pointer", whiteSpace: "nowrap", flexShrink: 0,
              border: `1px solid ${!customOpen && selectedMonth === m ? COLOR.gold : COLOR.hairline}`,
              background: !customOpen && selectedMonth === m ? COLOR.gold : "transparent",
              color: !customOpen && selectedMonth === m ? COLOR.bg : COLOR.inkMuted, fontWeight: !customOpen && selectedMonth === m ? 600 : 400,
            }}>{monthLabel(m)}</button>
          ))}
          <button onClick={() => setCustomOpen((o) => !o)} style={{
            padding: "8px 18px", borderRadius: 20, fontSize: 13, cursor: "pointer", whiteSpace: "nowrap", flexShrink: 0,
            border: `1px solid ${customOpen ? COLOR.gold : COLOR.hairline}`,
            background: customOpen ? COLOR.gold : "transparent", color: customOpen ? COLOR.bg : COLOR.inkMuted, fontWeight: customOpen ? 600 : 400,
          }}>Personnalisé…</button>
        </div>

        {customOpen && (
          <div className="gl-noprint" style={{ display: "flex", gap: 10, marginBottom: 20 }}>
            <Select label="Du mois" value={customFrom} onChange={setCustomFrom} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
            <Select label="Au mois" value={customTo} onChange={setCustomTo} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
          </div>
        )}

        <div className="gl-noprint" style={{ display: "flex", gap: 6, marginBottom: 20 }}>
          {(["Dépense", "Revenu"] as TxType[]).map((ty) => (
            <button key={ty} onClick={() => setTypeView(ty)} style={{
              padding: "6px 14px", borderRadius: 6, fontSize: 12, cursor: "pointer",
              border: `1px solid ${typeView === ty ? (ty === "Revenu" ? COLOR.emerald : COLOR.clay) : COLOR.hairline}`,
              background: typeView === ty ? (ty === "Revenu" ? "rgba(63,156,122,0.15)" : "rgba(193,84,63,0.15)") : "transparent",
              color: typeView === ty ? (ty === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft) : COLOR.inkMuted,
            }}>{ty === "Dépense" ? "Dépenses" : "Revenus"}</button>
          ))}
        </div>

        {donutData.length ? (
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={donutData} dataKey="value" nameKey="name" innerRadius={85} outerRadius={130} paddingAngle={2} startAngle={90} endAngle={-270}>
                {donutData.map((d) => <Cell key={d.name} fill={d.color} stroke={COLOR.surface} strokeWidth={2} />)}
              </Pie>
              <Tooltip content={<CustomTooltip />} />
            </PieChart>
          </ResponsiveContainer>
        ) : <EmptyState />}

        <div style={{ textAlign: "center", marginTop: 4 }}>
          <div style={{ fontSize: 12, color: COLOR.inkMuted, marginBottom: 6 }}>Vs. période précédente</div>
          <div style={{ fontSize: 18, fontWeight: 600, display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
            <span style={{ color: improved ? COLOR.emeraldSoft : COLOR.claySoft, display: "flex", alignItems: "center", gap: 4 }}>
              {delta >= 0 ? <TrendingUp size={15} /> : <TrendingDown size={15} />} {fmt(Math.abs(delta))} FCFA
            </span>
            <span style={{ color: COLOR.inkMuted }}>|</span>
            <span style={{ color: improved ? COLOR.emeraldSoft : COLOR.claySoft }}>
              {delta >= 0 ? "+" : "−"}{Math.abs(deltaPct).toFixed(0)}%
            </span>
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginTop: 24, paddingTop: 18, borderTop: `1px solid ${COLOR.hairline}` }}>
          <span style={{ fontSize: 15, color: COLOR.inkMuted }}>{typeView === "Dépense" ? "Dépenses" : "Revenus"} · {periodLabel}</span>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 600, color: typeView === "Dépense" ? COLOR.claySoft : COLOR.emeraldSoft }}>{fmt(total)} FCFA</span>
        </div>
      </Panel>

      <InsightsPanel
        filtered={periodAllTypesTx}
        catFocus={{ typeView, periodLabel, catList, prevByCat, total, prevTotal, delta, deltaPct }}
        title="Analyse d'expert financier"
        subtitle={`Critique et recommandations sur ${typeView === "Dépense" ? "les dépenses" : "les revenus"} · ${periodLabel}`}
      />

      <Panel
        title="Détail par catégorie"
        right={
          <div className="gl-noprint" style={{ display: "flex", gap: 8 }}>
            <button onClick={exportExcel} title="Exporter en Excel" style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(63,156,122,0.14)", border: `1px solid ${COLOR.emerald}`, borderRadius: 8, color: COLOR.emeraldSoft, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}>
              <FileSpreadsheet size={13} /> Excel
            </button>
            <button onClick={exportPDF} disabled={pdfState === "loading"} title="Exporter en PDF" style={{ display: "flex", alignItems: "center", gap: 6, background: pdfState === "error" ? "rgba(193,84,63,0.14)" : "rgba(201,162,39,0.14)", border: `1px solid ${pdfState === "error" ? COLOR.clay : COLOR.gold}`, borderRadius: 8, color: pdfState === "error" ? COLOR.claySoft : COLOR.goldSoft, padding: "6px 12px", fontSize: 12, cursor: pdfState === "loading" ? "default" : "pointer" }}>
              {pdfState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <FileText size={13} />}
              {pdfState === "loading" ? "Génération…" : pdfState === "error" ? "Réessayer" : "PDF"}
            </button>
          </div>
        }
      >
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          {donutData.map((c) => {
            const isOpen = expandedCat === c.name;
            const subs = isOpen ? subcatFor(c.name) : [];
            return (
              <div key={c.name}>
                <div
                  onClick={() => setExpandedCat(isOpen ? null : c.name)}
                  style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 0", borderBottom: isOpen ? "none" : `1px solid ${COLOR.hairline}`, cursor: "pointer" }}
                >
                  <div style={{ width: 40, height: 40, borderRadius: 10, background: `${c.color}22`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <span style={{ width: 12, height: 12, borderRadius: "50%", background: c.color, display: "inline-block" }} />
                  </div>
                  <div style={{ flex: 1, fontSize: 14, color: COLOR.ink, display: "flex", alignItems: "center", gap: 6 }}>
                    {c.name}
                    <ChevronDown size={14} color={COLOR.inkMuted} style={{ transform: isOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, fontWeight: 600, color: COLOR.ink }}>{fmt(c.value)} FCFA</div>
                    <div style={{ fontSize: 11.5, color: COLOR.inkMuted }}>{c.pct.toFixed(0)}%</div>
                  </div>
                </div>
                {isOpen && (
                  <div style={{ padding: "0 0 14px 54px", display: "flex", flexDirection: "column", gap: 2, borderBottom: `1px solid ${COLOR.hairline}` }}>
                    {subs.map((s) => {
                      const subOpen = expandedSub === `${c.name}::${s.name}`;
                      const subTx = subOpen ? txFor(c.name, s.name) : [];
                      return (
                        <div key={s.name}>
                          <div
                            onClick={() => setExpandedSub(subOpen ? null : `${c.name}::${s.name}`)}
                            style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", fontSize: 13, cursor: "pointer" }}
                          >
                            <span style={{ color: COLOR.inkMuted, display: "flex", alignItems: "center", gap: 5 }}>
                              <ChevronDown size={11} color={COLOR.inkMuted} style={{ transform: subOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                              {s.name}
                            </span>
                            <span style={{ display: "flex", gap: 10 }}>
                              <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.ink }}>{fmt(s.value)} FCFA</span>
                              <span style={{ color: COLOR.inkMuted, minWidth: 32, textAlign: "right" }}>{s.pct.toFixed(0)}%</span>
                            </span>
                          </div>
                          {subOpen && (
                            <div style={{ padding: "2px 0 10px 16px", display: "flex", flexDirection: "column", gap: 6 }}>
                              {subTx.map((t) => (
                                <div key={t.id} style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: "8px 10px" }}>
                                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                                    <div style={{ fontSize: 12.5, color: COLOR.ink, fontFamily: "'IBM Plex Mono', monospace" }}>
                                      {dateLabelFull(t.date)}
                                      {t.account && <span style={{ color: COLOR.inkMuted }}> · {t.account}</span>}
                                      {t.payee && <span style={{ color: COLOR.inkMuted }}> · {t.payee}</span>}
                                    </div>
                                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(t.amount)} FCFA</span>
                                      <button onClick={() => startEdit(t)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={12} /></button>
                                      <button onClick={() => setConfirmDeleteId(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={12} /></button>
                                    </div>
                                  </div>
                                </div>
                              ))}
                              {!subTx.length && <div style={{ fontSize: 12.5, color: COLOR.inkMuted }}>Aucune transaction.</div>}
                            </div>
                          )}
                        </div>
                      );
                    })}
                    {!subs.length && <div style={{ fontSize: 13, color: COLOR.inkMuted, padding: "7px 0" }}>Aucune sous-catégorie pour cette période.</div>}
                  </div>
                )}
              </div>
            );
          })}
          {!donutData.length && <EmptyState text="Aucune transaction pour cette période." />}
        </div>
      </Panel>

      <TransactionEditSheet
        open={!!editingTx}
        transaction={editingTx}
        transactions={transactions}
        accounts={accounts}
        onClose={() => setEditingTx(null)}
        onSave={saveEdit}
        onDelete={removeTx}
      />
      <ConfirmDialog
        open={!!confirmDeleteId}
        title="Supprimer cette transaction ?"
        message="Cette action est définitive. Le montant ne sera plus comptabilisé nulle part dans l'app."
        onConfirm={() => { if (confirmDeleteId) removeTx(confirmDeleteId); setConfirmDeleteId(null); }}
        onCancel={() => setConfirmDeleteId(null)}
      />
    </div>
  );
}

// ============================================================
// APERÇU DE CATÉGORIE — courbe d'évolution, moyenne, sélecteur avec "Changer"
// ============================================================
function trailingRange(allMonths: string[], n: number): [string, string] {
  const cur = dateToMonthKey(todayISO());
  const curK = monthSortKey(cur);
  const from = allMonths.find((m) => monthSortKey(m) >= curK - (n - 1)) || allMonths[0] || cur;
  return [from, cur];
}

function CategoryOverviewTab({ transactions, categoryGroups, allMonths }: {
  transactions: Transaction[]; categoryGroups: Record<string, Group>; allMonths: string[];
}) {
  const [pickerOpen, setPickerOpen] = useState(false);
  const [type, setType] = useState<TxType>("Dépense");
  const [category, setCategory] = useState(() => defaultQuickCategory(transactions, "Dépense"));
  const [subcategory, setSubcategory] = useState("");
  const [presetKey, setPresetKey] = useState("6m");
  const [granularity, setGranularity] = useState<"mois" | "jour">("mois");

  const presets: { key: string; label: string; range: () => [string, string] }[] = [
    { key: "mtd", label: "MTD", range: () => { const k = dateToMonthKey(todayISO()); return [k, k]; } },
    { key: "ytd", label: "Depuis le début de l'année", range: () => { const y = todayISO().slice(0, 4); const jan = allMonths.find((m) => m.startsWith(y)) || dateToMonthKey(todayISO()); return [jan, dateToMonthKey(todayISO())]; } },
    { key: "1m", label: "1M", range: () => trailingRange(allMonths, 1) },
    { key: "3m", label: "3M", range: () => trailingRange(allMonths, 3) },
    { key: "6m", label: "6M", range: () => trailingRange(allMonths, 6) },
    { key: "1a", label: "1A", range: () => trailingRange(allMonths, 12) },
    { key: "all", label: "Tout", range: () => [allMonths[0] || dateToMonthKey(todayISO()), allMonths[allMonths.length - 1] || dateToMonthKey(todayISO())] },
  ];
  const activePreset = presets.find((p) => p.key === presetKey) || presets[4];
  const [from, to] = activePreset.range();

  const catTxAll = useMemo(
    () => transactions.filter((t) => t.category === category && t.type === type && (!subcategory || t.subcategory === subcategory)).map((t) => ({ ...t, month: dateToMonthKey(t.date) })),
    [transactions, category, type, subcategory]
  );
  const fk = monthSortKey(from), tk = monthSortKey(to);
  const catTx = catTxAll.filter((t) => { const k = monthSortKey(t.month); return k >= fk && k <= tk; });
  const total = catTx.reduce((a, t) => a + t.amount, 0);

  // Détecte le cas où la sous-catégorie choisie n'a jamais été renseignée sur les transactions
  // existantes, alors que la catégorie parente, elle, contient bien des données.
  const hasAnySubcatData = subcategory ? transactions.some((t) => t.category === category && t.type === type && t.subcategory === subcategory) : true;
  const parentCatTotal = useMemo(() => transactions.filter((t) => t.category === category && t.type === type).reduce((a, t) => a + t.amount, 0), [transactions, category, type]);

  const monthsInRange = allMonths.filter((m) => { const k = monthSortKey(m); return k >= fk && k <= tk; });
  const byMonth = monthsInRange.map((m) => ({ key: m, label: monthLabel(m), value: catTxAll.filter((t) => t.month === m).reduce((a, t) => a + t.amount, 0) }));

  // Vue journalière : un point par jour civil couvert par la fenêtre sélectionnée
  // (bornes du 1er jour du premier mois au dernier jour du dernier mois de la plage),
  // avec la somme des transactions de la catégorie/sous-catégorie tombant ce jour-là.
  const byDay = useMemo(() => {
    if (granularity !== "jour") return [];
    const start = monthKeyToFirstDate(from);
    const [endY, endM] = to.split("_");
    const lastDayOfEndMonth = new Date(parseInt(endY, 10), parseInt(endM, 10), 0).getDate();
    const end = `${endY}-${pad2(parseInt(endM, 10))}-${pad2(lastDayOfEndMonth)}`;
    const byDate: Record<string, number> = {};
    catTxAll.forEach((t) => { byDate[t.date] = (byDate[t.date] || 0) + t.amount; });
    const days: { key: string; label: string; value: number }[] = [];
    const cur = new Date(start + "T00:00:00");
    const endD = new Date(end + "T00:00:00");
    while (cur <= endD) {
      const iso = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      days.push({ key: iso, label: dateLabelShort(iso), value: byDate[iso] || 0 });
      cur.setDate(cur.getDate() + 1);
    }
    return days;
  }, [granularity, catTxAll, from, to]);

  const series = granularity === "jour" ? byDay : byMonth;
  const avg = series.length ? mean(series.map((m) => m.value)) : 0;

  const changeType = (ty: TxType) => { setType(ty); setCategory(defaultQuickCategory(transactions, ty)); setSubcategory(""); };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel>
        <div className="gl-noprint gl-scroll" style={{ display: "flex", gap: 8, overflowX: "auto", marginBottom: 18, paddingBottom: 4, WebkitOverflowScrolling: "touch" }}>
          {presets.map((p) => (
            <button key={p.key} onClick={() => setPresetKey(p.key)} style={{
              padding: "8px 16px", borderRadius: 20, fontSize: 12.5, cursor: "pointer", whiteSpace: "nowrap", flexShrink: 0,
              border: `1px solid ${presetKey === p.key ? COLOR.gold : COLOR.hairline}`,
              background: presetKey === p.key ? COLOR.gold : "transparent",
              color: presetKey === p.key ? COLOR.bg : COLOR.inkMuted, fontWeight: presetKey === p.key ? 600 : 400,
            }}>{p.label}</button>
          ))}
        </div>

        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "12px 16px", background: COLOR.surfaceRaised, borderRadius: 12, marginBottom: 18, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 0 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: `${type === "Revenu" ? COLOR.emerald : COLOR.clay}22`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <Layers size={17} color={type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft} />
            </div>
            <div style={{ fontSize: 15, fontWeight: 600, color: COLOR.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {category}{subcategory && <span style={{ color: COLOR.inkMuted, fontWeight: 400 }}> · {subcategory}</span>}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
            <div style={{ display: "flex", gap: 4, background: COLOR.surface, borderRadius: 16, padding: 3, border: `1px solid ${COLOR.hairline}` }}>
              {(["mois", "jour"] as const).map((g) => (
                <button key={g} onClick={() => setGranularity(g)} style={{
                  padding: "4px 10px", borderRadius: 12, fontSize: 11, cursor: "pointer", border: "none",
                  background: granularity === g ? COLOR.gold : "transparent",
                  color: granularity === g ? COLOR.bg : COLOR.inkMuted,
                }}>{g === "mois" ? "Mois" : "Jour"}</button>
              ))}
            </div>
            <div style={{ display: "flex", gap: 4, background: COLOR.surface, borderRadius: 16, padding: 3, border: `1px solid ${COLOR.hairline}` }}>
              {(["Dépense", "Revenu"] as TxType[]).map((ty) => (
                <button key={ty} onClick={() => changeType(ty)} style={{
                  padding: "4px 10px", borderRadius: 12, fontSize: 11, cursor: "pointer", border: "none",
                  background: type === ty ? (ty === "Revenu" ? COLOR.emerald : COLOR.clay) : "transparent",
                  color: type === ty ? COLOR.bg : COLOR.inkMuted,
                }}>{ty}</button>
              ))}
            </div>
            <button onClick={() => setPickerOpen(true)} style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 16, color: COLOR.goldSoft, padding: "7px 16px", fontSize: 12.5, cursor: "pointer" }}>
              Changer
            </button>
          </div>
        </div>
        <CategoryPickerSheet open={pickerOpen} onClose={() => setPickerOpen(false)} transactions={transactions} type={type}
          value={category} subvalue={subcategory} onSelect={(c, s) => { setCategory(c); setSubcategory(s); }} />

        {subcategory && !hasAnySubcatData && parentCatTotal > 0 && (
          <div style={{ display: "flex", gap: 10, padding: "12px 14px", background: "rgba(201,162,39,0.08)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, marginBottom: 18 }}>
            <Info size={15} color={COLOR.goldSoft} style={{ flexShrink: 0, marginTop: 1 }} />
            <div style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.55 }}>
              Aucune transaction n'a la sous-catégorie <b style={{ color: COLOR.ink }}>"{subcategory}"</b> renseignée pour l'instant — c'est normal pour les données
              historiques importées, qui n'ont que la catégorie. <b style={{ color: COLOR.goldSoft }}>{fmt(parentCatTotal)} FCFA</b> existent au total dans "{category}" toutes sous-catégories confondues.{" "}
              <button onClick={() => setSubcategory("")} style={{ background: "none", border: "none", color: COLOR.goldSoft, textDecoration: "underline", cursor: "pointer", padding: 0, fontSize: 12.5 }}>
                Voir toute la catégorie
              </button>
            </div>
          </div>
        )}

        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 34, fontWeight: 600, color: type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>
          {type === "Dépense" ? "−" : "+"}{fmt(total)} <span style={{ fontSize: 15, color: COLOR.inkMuted }}>FCFA</span>
        </div>
        <div style={{ fontSize: 12.5, color: COLOR.inkMuted, marginTop: 4, marginBottom: 20 }}>
          {monthsInRange.length ? `${monthLabel(monthsInRange[0])} — ${monthLabel(monthsInRange[monthsInRange.length - 1])}` : "Aucune donnée"}
          {granularity === "jour" && series.length ? ` · ${series.length} jours` : ""}
        </div>

        {series.length ? (
          <ResponsiveContainer width="100%" height={240}>
            <AreaChart data={series} margin={{ left: 0, right: 10, top: 20 }}>
              <defs>
                <linearGradient id="catGrad" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={type === "Revenu" ? COLOR.emerald : COLOR.clay} stopOpacity={0.4} />
                  <stop offset="100%" stopColor={type === "Revenu" ? COLOR.emerald : COLOR.clay} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="label" tick={{ fill: COLOR.inkMuted, fontSize: 9.5 }} interval={series.length > 12 ? Math.floor(series.length / (granularity === "jour" ? 10 : 8)) : 0} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <ReferenceLine y={avg} stroke={COLOR.slateBlueSoft} strokeDasharray="5 4"
                label={{ value: `Moyenne : ${fmt(avg)} FCFA`, position: "insideTopRight", fill: COLOR.slateBlueSoft, fontSize: 10.5 }} />
              <Area type="monotone" dataKey="value" name={category} stroke={type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft} strokeWidth={2} fill="url(#catGrad)" />
            </AreaChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
      </Panel>

      <Panel title={granularity === "jour" ? "Détail journalier" : "Détail mensuel"}>
        <div style={{ display: "flex", flexDirection: "column" }}>
          {series.slice().reverse().filter((m) => granularity === "mois" || m.value).map((m) => (
            <div key={m.key} style={{ display: "flex", justifyContent: "space-between", padding: "12px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <span style={{ fontSize: 14, color: COLOR.ink }}>{m.label}</span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, fontWeight: 600, color: m.value ? (type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft) : COLOR.inkMuted }}>
                {m.value ? `${type === "Dépense" ? "−" : "+"}${fmt(m.value)} FCFA` : "—"}
              </span>
            </div>
          ))}
          {!series.length && <EmptyState text="Aucune donnée pour cette période." />}
          {granularity === "jour" && series.length > 0 && series.every((m) => !m.value) && (
            <EmptyState text="Aucune transaction certains jours de cette période." />
          )}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================ END OF PART 4 — continued below
// ============================================================
// ENVELOPPES TAB (avec alertes)
// ============================================================
function EnveloppesTab({ filtered, cap, setCap }: { filtered: any[]; cap: number; setCap: (n: number) => void }) {
  const envCats = ["Cadeaux", "Divertissement", "Invitation"];
  const byMonth = useMemo(() => {
    const m: Record<string, number> = {};
    filtered.filter((t) => t.type === "Dépense" && envCats.includes(t.category)).forEach((t) => { m[t.month] = (m[t.month] || 0) + t.amount; });
    return Object.keys(m).sort((a, b) => monthSortKey(a) - monthSortKey(b)).map((k) => ({ mois: monthLabel(k), key: k, total: m[k] }));
  }, [filtered]);

  const last = byMonth[byMonth.length - 1];
  const pct = last ? Math.min(100, (last.total / cap) * 100) : 0;
  const status = pct < 70 ? COLOR.emerald : pct < 100 ? COLOR.gold : COLOR.clay;

  const perCat = envCats.map((cat) => {
    const total = filtered.filter((t) => t.type === "Dépense" && t.category === cat && (!last || t.month === last.key)).reduce((a, t) => a + t.amount, 0);
    return { name: cat, value: total };
  });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {last && last.total > cap && (
        <div style={{ display: "flex", alignItems: "center", gap: 10, background: "rgba(193,84,63,0.12)", border: `1px solid ${COLOR.clay}`, borderRadius: 8, padding: "12px 16px" }}>
          <Bell size={16} color={COLOR.claySoft} />
          <span style={{ fontSize: 13, color: COLOR.claySoft }}>
            Alerte : l'enveloppe a dépassé son plafond en {last.mois} ({fmt(last.total)} / {fmt(cap)} FCFA)
          </span>
        </div>
      )}
      <Panel title="Enveloppe combinée : Cadeaux + Divertissement + Invitation" subtitle={last ? `Dernier mois filtré (${last.mois}) : ${fmt(last.total)} FCFA` : "Aucune donnée"}>
        <div style={{ display: "flex", gap: 24, alignItems: "center", flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 260 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
              <span style={{ fontSize: 12.5, color: COLOR.inkMuted }}>Consommé</span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: status }}>{last ? fmt(last.total) : 0} / {fmt(cap)} FCFA</span>
            </div>
            <div style={{ height: 12, background: COLOR.hairline, borderRadius: 6, overflow: "hidden" }}>
              <div style={{ width: `${pct}%`, height: "100%", background: status, transition: "width 0.3s" }} />
            </div>
            <div style={{ marginTop: 14, display: "flex", alignItems: "center", gap: 10 }}>
              <Target size={14} color={COLOR.inkMuted} />
              <label style={{ fontSize: 12.5, color: COLOR.inkMuted }}>Plafond mensuel (FCFA)</label>
              <input type="number" inputMode="numeric" value={cap} onChange={(e) => setCap(Number(e.target.value) || 0)} style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink, padding: "6px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, width: 130 }} />
            </div>
          </div>
        </div>
      </Panel>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16 }}>
        {perCat.map((c) => (
          <div key={c.name} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, overflow: "hidden" }}>
            <div style={{ height: 28, background: `linear-gradient(135deg, ${COLOR.hairline} 49.5%, transparent 50%), linear-gradient(-135deg, ${COLOR.hairline} 49.5%, transparent 50%)`, backgroundSize: "50% 100%", backgroundPosition: "left top, right top", backgroundRepeat: "no-repeat" }} />
            <div style={{ padding: 16 }}>
              <div style={{ fontSize: 13, marginBottom: 8 }}>{c.name}</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 18, fontWeight: 600 }}>{fmt(c.value)}</div>
            </div>
          </div>
        ))}
      </div>
      <Panel title="Rythme de l'enveloppe dans le temps" subtitle="Total mensuel vs plafond">
        {byMonth.length ? (
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={byMonth} margin={{ left: 0, right: 10, top: 10 }}>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} interval={byMonth.length > 12 ? 1 : 0} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <ReferenceLine y={cap} stroke={COLOR.goldSoft} strokeDasharray="4 3" />
              <Bar dataKey="total" name="Cadeaux+Divert.+Invit." radius={[3, 3, 0, 0]}>
                {byMonth.map((d, i) => <Cell key={i} fill={d.total > cap ? COLOR.clay : d.total > cap * 0.7 ? COLOR.gold : COLOR.emerald} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
      </Panel>
    </div>
  );
}

// ============================================================
// SIMULATEUR "ET SI"
// ============================================================
function SimulateurTab({ filtered, accounts, transactions }: { filtered: any[]; accounts: Account[]; transactions: Transaction[] }) {
  const nonProdCats = useMemo(() => {
    const m: Record<string, number> = {};
    filtered.filter((t) => t.type === "Dépense" && t.group === "Non-productif").forEach((t) => { m[t.category] = (m[t.category] || 0) + t.amount; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 8);
  }, [filtered]);

  const [reductions, setReductions] = useState<Record<string, number>>({});
  useEffect(() => {
    const init: Record<string, number> = {};
    nonProdCats.forEach(([name]) => { if (!(name in reductions)) init[name] = 20; });
    if (Object.keys(init).length) setReductions((r) => ({ ...init, ...r }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nonProdCats.map((c) => c[0]).join(",")]);

  const totalRevenus = filtered.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const totalDepenses = filtered.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const solde = totalRevenus - totalDepenses;
  const monthsInRange = new Set(filtered.map((t) => t.month)).size || 1;

  const savings = nonProdCats.reduce((a, [name, val]) => a + val * ((reductions[name] || 0) / 100), 0);
  const newSolde = solde + savings;
  const newTaux = totalRevenus > 0 ? (newSolde / totalRevenus) * 100 : 0;
  const oldTaux = totalRevenus > 0 ? (solde / totalRevenus) * 100 : 0;
  const monthlySaving = savings / monthsInRange;

  const projection = useMemo(() => {
    const series = liveNetWorthSeries(accounts, transactions);
    const last = series[series.length - 1][1];
    return [0, 6, 12, 24].map((n) => ({ mois: n === 0 ? "aujourd'hui" : `+${n}m`, sansAction: last + (n * (projectNetWorth(1, series).avgDelta)), avecAction: last + (n * (projectNetWorth(1, series).avgDelta + monthlySaving)) }));
  }, [monthlySaving, accounts]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Simulateur « et si… »" subtitle="Ajustez le curseur de réduction par catégorie non-productive, sur la période filtrée">
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          {nonProdCats.map(([name, val]) => (
            <div key={name}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 6 }}>
                <span>{name}</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.inkMuted }}>
                  {fmt(val)} → <span style={{ color: COLOR.emeraldSoft }}>{fmt(val * (1 - (reductions[name] || 0) / 100))}</span> ({reductions[name] || 0}% de réduction)
                </span>
              </div>
              <input type="range" min={0} max={100} value={reductions[name] || 0}
                onChange={(e) => setReductions({ ...reductions, [name]: Number(e.target.value) })}
                style={{ width: "100%", accentColor: COLOR.gold }} />
            </div>
          ))}
          {!nonProdCats.length && <EmptyState text="Aucune dépense non-productive dans cette période." />}
        </div>
      </Panel>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Économies simulées (période)" value={fmt(savings)} tone={COLOR.emeraldSoft} icon={Sparkles} />
        <Kpi label="Nouveau taux d'épargne" value={newTaux.toFixed(1)} suffix="%" tone={COLOR.gold} hint={`vs ${oldTaux.toFixed(1)}% actuellement`} />
        <Kpi label="Économie mensuelle moyenne" value={fmt(monthlySaving)} tone={COLOR.emeraldSoft} />
        <Kpi label="Nouveau solde (période)" value={fmt(newSolde)} tone={newSolde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} />
      </div>

      <PanelWithHelp title="Impact projeté sur la valeur nette" subtitle="Comparaison : trajectoire actuelle vs trajectoire avec les réductions appliquées chaque mois"
        explain="La ligne grise prolonge ta trajectoire actuelle sans rien changer. La ligne verte simule l'effet, mois après mois, si tu appliquais réellement les réductions réglées avec les curseurs ci-dessus. L'écart entre les deux, au bout de 24 mois, donne une idée concrète du gain potentiel sur ton patrimoine.">
        <ResponsiveContainer width="100%" height={260}>
          <LineChart data={projection} margin={{ left: 0, right: 10, top: 10 }}>
            <CartesianGrid stroke={COLOR.hairline} vertical={false} />
            <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
            <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ fontSize: 12, color: COLOR.inkMuted }} />
            <Line type="monotone" dataKey="sansAction" name="Sans changement" stroke={COLOR.inkMuted} strokeWidth={2} dot />
            <Line type="monotone" dataKey="avecAction" name="Avec réductions" stroke={COLOR.emeraldSoft} strokeWidth={2.5} dot />
          </LineChart>
        </ResponsiveContainer>
      </PanelWithHelp>
    </div>
  );
}

// ============================================================
// PROJECTION TAB CONTENT (utilisé dans Aperçu section additionnelle — intégré ici pour Simulateur avancé)
// ============================================================
function ProjectionPanel({ accounts, transactions }: { accounts: Account[]; transactions: Transaction[] }) {
  const [months, setMonths] = useState(12);
  const { points } = projectNetWorth(months, liveNetWorthSeries(accounts, transactions));
  return (
    <PanelWithHelp title="Projection de valeur nette" subtitle={`Basée sur la tendance des 6 derniers relevés — bande optimiste/pessimiste (±1 écart-type)`}
      explain="La ligne dorée centrale prolonge la moyenne d'évolution de ta valeur nette sur les 6 derniers mois. Les deux lignes pointillées (vert=optimiste, rouge=prudent) montrent une fourchette réaliste autour de cette projection, basée sur la variabilité récente de ton patrimoine. C'est une extrapolation statistique, pas une garantie — un gros achat ou une rentrée d'argent imprévue peut la faire dévier."
      right={
        <div style={{ display: "flex", gap: 6 }}>
          {[6, 12, 24].map((m) => (
            <button key={m} onClick={() => setMonths(m)} style={{
              background: months === m ? "rgba(201,162,39,0.15)" : "transparent", border: `1px solid ${months === m ? COLOR.gold : COLOR.hairline}`,
              color: months === m ? COLOR.goldSoft : COLOR.inkMuted, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer",
            }}>{m} mois</button>
          ))}
        </div>
      }>
      <ResponsiveContainer width="100%" height={260}>
        <AreaChart data={points} margin={{ left: 0, right: 10, top: 10 }}>
          <CartesianGrid stroke={COLOR.hairline} vertical={false} />
          <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
          <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
          <Tooltip content={<CustomTooltip />} />
          <Area type="monotone" dataKey="haut" stroke="none" fill={COLOR.gold} fillOpacity={0.08} />
          <Area type="monotone" dataKey="bas" stroke="none" fill={COLOR.bg} fillOpacity={1} />
          <Line type="monotone" dataKey="central" name="Projection centrale" stroke={COLOR.goldSoft} strokeWidth={2.5} dot={false} />
          <Line type="monotone" dataKey="haut" name="Scénario optimiste" stroke={COLOR.emeraldSoft} strokeWidth={1.5} strokeDasharray="4 3" dot={false} />
          <Line type="monotone" dataKey="bas" name="Scénario prudent" stroke={COLOR.claySoft} strokeWidth={1.5} strokeDasharray="4 3" dot={false} />
        </AreaChart>
      </ResponsiveContainer>
    </PanelWithHelp>
  );
}

// ============================================================
// OBJECTIF D'ÉPARGNE
// ============================================================
function GoalsPanel({ goals, setGoals, accounts, transactions }: { goals: Goal[]; setGoals: (g: Goal[]) => void; accounts: Account[]; transactions: Transaction[] }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Goal, "id">>({ name: "", target: 1000000, current: 0, date: "" });
  const { avgDelta } = projectNetWorth(1, liveNetWorthSeries(accounts, transactions));

  const add = () => { if (!form.name || form.target <= 0) return; setGoals([...goals, { ...form, id: uid("g") }]); setForm({ name: "", target: 1000000, current: 0, date: "" }); setAdding(false); };
  const update = (id: string, patch: Partial<Goal>) => setGoals(goals.map((g) => (g.id === id ? { ...g, ...patch } : g)));
  const remove = (id: string) => setGoals(goals.filter((g) => g.id !== id));

  return (
    <Panel title="Objectifs d'épargne" subtitle="Plusieurs cibles simultanées — patrimoine global, projets spécifiques…"
      right={
        <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
          {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Nouvel objectif"}
        </button>
      }>
      {adding && (
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Nom</label><input style={{ ...inputStyle, width: 180 }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Voyage, voiture, patrimoine…" /></div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Cible (FCFA)</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 150 }} value={form.target} onChange={(e) => setForm({ ...form, target: Number(e.target.value) || 0 })} /></div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Déjà atteint (FCFA)</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 150 }} value={form.current} onChange={(e) => setForm({ ...form, current: Number(e.target.value) || 0 })} /></div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Date cible</label><input style={{ ...inputStyle, width: 120 }} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} placeholder="déc. 2027" /></div>
          <button onClick={add} style={{ background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}>Créer</button>
        </div>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        {goals.map((g) => {
          const pct = Math.min(100, (g.current / g.target) * 100);
          const remaining = g.target - g.current;
          const monthsNeeded = avgDelta > 0 ? Math.ceil(remaining / avgDelta) : null;
          return (
            <div key={g.id} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                <span style={{ fontSize: 14, fontWeight: 500 }}>{g.name}</span>
                <button onClick={() => remove(g.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                <input type="number" inputMode="numeric" value={g.current} onChange={(e) => update(g.id, { current: Number(e.target.value) || 0 })} style={{ ...inputStyle, width: 140 }} />
                <span style={{ alignSelf: "center", color: COLOR.inkMuted, fontSize: 12 }}>/ {fmt(g.target)} FCFA {g.date && `· ${g.date}`}</span>
              </div>
              <div style={{ height: 10, background: COLOR.hairline, borderRadius: 5, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: COLOR.gold }} />
              </div>
              <div style={{ marginTop: 8, fontSize: 11.5, color: COLOR.inkMuted }}>
                {pct.toFixed(1)}% atteint {monthsNeeded !== null && monthsNeeded > 0 && remaining > 0 ? `· ~${monthsNeeded} mois au rythme actuel du patrimoine global` : ""}
              </div>
            </div>
          );
        })}
        {!goals.length && <EmptyState text="Aucun objectif défini." />}
      </div>
    </Panel>
  );
}

// ============================================================
// BUSINESS / PERSONNEL TAB
// ============================================================
function BusinessTab({ transactions, categoryGroups, categoryScope, setCategoryScope, allCategories }: {
  transactions: Transaction[]; categoryGroups: Record<string, Group>; categoryScope: Record<string, Scope>;
  setCategoryScope: (s: Record<string, Scope>) => void; allCategories: string[];
}) {
  const withScope = transactions.map((t) => ({ ...t, scope: categoryScope[t.category] || "Personnel" }));
  const bizRev = withScope.filter((t) => t.scope === "Business" && t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const bizDep = withScope.filter((t) => t.scope === "Business" && t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const bizMargin = bizRev - bizDep;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Revenus Business (total)" value={fmt(bizRev)} tone={COLOR.emeraldSoft} icon={Briefcase} />
        <Kpi label="Dépenses Business (total)" value={fmt(bizDep)} tone={COLOR.claySoft} icon={Briefcase} />
        <Kpi label="Marge Business" value={fmt(bizMargin)} tone={bizMargin >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} />
      </div>
      <Panel title="Compte de résultat — activité (GRUNDFOS / ECO PUMP AFRIK / INVEST SGO)" subtitle="Isole l'activité commerciale du budget personnel">
        <div style={{ fontSize: 13, color: COLOR.inkMuted, marginBottom: 14, lineHeight: 1.6 }}>
          Mélanger trésorerie personnelle et activité commerciale déforme les deux analyses. Cette vue les sépare :
          la marge business ci-dessus ne doit pas être confondue avec le solde personnel des autres onglets.
        </div>
      </Panel>
      <Panel title="Assigner la portée de chaque catégorie" subtitle="Personnel ou Business — modifiable à tout moment">
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {allCategories.map((c) => (
            <div key={c} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <span style={{ fontSize: 12.5 }}>{c}</span>
              <select value={categoryScope[c] || "Personnel"} onChange={(e) => setCategoryScope({ ...categoryScope, [c]: e.target.value as Scope })}
                style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: (categoryScope[c] || "Personnel") === "Business" ? COLOR.violetSoft : COLOR.inkMuted, padding: "5px 8px", fontSize: 11.5, width: 120 }}>
                <option value="Personnel">Personnel</option>
                <option value="Business">Business</option>
              </select>
            </div>
          ))}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// CRÉANCES (PRÊTS) TAB
// ============================================================
function CreancesTab({ loans, setLoans }: { loans: Loan[]; setLoans: (l: Loan[]) => void }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Loan, "id">>({ person: "", amount: 0, dateGiven: "2026_8", status: "En attente", notes: "" });

  const add = () => { if (!form.person || form.amount <= 0) return; setLoans([...loans, { ...form, id: uid("l") }]); setForm({ person: "", amount: 0, dateGiven: "2026_8", status: "En attente", notes: "" }); setAdding(false); };
  const toggleStatus = (id: string) => setLoans(loans.map((l) => l.id === id ? { ...l, status: l.status === "En attente" ? "Remboursé" : "En attente" } : l));
  const remove = (id: string) => setLoans(loans.filter((l) => l.id !== id));

  const totalOutstanding = loans.filter((l) => l.status === "En attente").reduce((a, l) => a + l.amount, 0);
  const totalRepaid = loans.filter((l) => l.status === "Remboursé").reduce((a, l) => a + l.amount, 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Créances en attente" value={fmt(totalOutstanding)} tone={COLOR.gold} icon={HandCoins} />
        <Kpi label="Créances remboursées" value={fmt(totalRepaid)} tone={COLOR.emeraldSoft} icon={Check} />
      </div>
      <Panel title="Suivi des prêts accordés" subtitle="Ces montants ne sont pas des dépenses perdues — ce sont des créances récupérables"
        right={
          <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
            {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Ajouter un prêt"}
          </button>
        }>
        {adding && (
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Personne</label><input style={{ ...inputStyle, width: 160 }} value={form.person} onChange={(e) => setForm({ ...form, person: e.target.value })} /></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 130 }} value={form.amount || ""} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) })} /></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Mois (AAAA_M)</label><input style={{ ...inputStyle, width: 100 }} value={form.dateGiven} onChange={(e) => setForm({ ...form, dateGiven: e.target.value })} /></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Note</label><input style={{ ...inputStyle, width: 180 }} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></div>
            <button onClick={add} style={{ display: "flex", alignItems: "center", gap: 6, background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}><Save size={13} /> Enregistrer</button>
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {loans.map((l) => (
            <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}` }}>
              <div>
                <div style={{ fontSize: 13 }}>{l.person} <span style={{ color: COLOR.inkMuted, fontSize: 11.5 }}>· {monthLabel(l.dateGiven)}</span></div>
                {l.notes && <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 2 }}>{l.notes}</div>}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}>{fmt(l.amount)}</span>
                <button onClick={() => toggleStatus(l.id)} style={{ fontSize: 11, padding: "4px 10px", borderRadius: 12, border: `1px solid ${l.status === "Remboursé" ? COLOR.emerald : COLOR.gold}`, background: "transparent", color: l.status === "Remboursé" ? COLOR.emeraldSoft : COLOR.goldSoft, cursor: "pointer" }}>{l.status}</button>
                <button onClick={() => remove(l.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
              </div>
            </div>
          ))}
          {!loans.length && <EmptyState text="Aucune créance enregistrée." />}
        </div>
      </Panel>
    </div>
  );
}
// ============================================================
// COMPTES (ACCOUNTS)
// ============================================================
function ComptesTab({ accounts, setAccounts, transactions }: { accounts: Account[]; setAccounts: (a: Account[]) => void; transactions: Transaction[] }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Account, "id">>({ name: "", kind: "Banque", openingBalance: 0 });
  const [editingOpening, setEditingOpening] = useState<string | null>(null);
  const kinds: Account["kind"][] = ["Espèces", "Banque", "Mobile Money", "Carte de crédit", "Autre"];

  const add = () => { if (!form.name) return; setAccounts([...accounts, { ...form, id: uid("a") }]); setForm({ name: "", kind: "Banque", openingBalance: 0 }); setAdding(false); };
  const update = (id: string, patch: Partial<Account>) => setAccounts(accounts.map((a) => (a.id === id ? { ...a, ...patch } : a)));
  const remove = (id: string) => setAccounts(accounts.filter((a) => a.id !== id));
  const total = totalAccountsBalance(accounts, transactions);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Kpi label="Total des comptes (temps réel)" value={fmt(total)} tone={COLOR.goldSoft} icon={Wallet} />
      <Panel title="Comptes" subtitle="Le solde de chaque compte se met à jour automatiquement dès qu'une transaction lui est liée"
        right={
          <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
            {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Ajouter un compte"}
          </button>
        }>
        {adding && (
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Nom</label><input style={{ ...inputStyle, width: 170 }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Type</label><select style={{ ...inputStyle, width: 150 }} value={form.kind} onChange={(e) => setForm({ ...form, kind: e.target.value as Account["kind"] })}>{kinds.map((k) => <option key={k} value={k}>{k}</option>)}</select></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Solde de départ (FCFA)</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 160 }} value={form.openingBalance} onChange={(e) => setForm({ ...form, openingBalance: Number(e.target.value) || 0 })} /></div>
            <button onClick={add} style={{ background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}>Créer</button>
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {accounts.map((a) => {
            const linked = transactions.filter((t) => t.account === a.name);
            const current = accountBalance(a, transactions);
            const isEditingOpening = editingOpening === a.id;
            return (
              <div key={a.id} style={{ padding: "12px 14px", background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div><div style={{ fontSize: 13 }}>{a.name}</div><div style={{ fontSize: 11, color: COLOR.inkMuted }}>{a.kind} · {linked.length} transaction(s) liée(s)</div></div>
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: current >= 0 ? COLOR.ink : COLOR.claySoft }}>{fmt(current)}</div>
                    <button onClick={() => setEditingOpening(isEditingOpening ? null : a.id)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={13} /></button>
                    <button onClick={() => remove(a.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
                  </div>
                </div>
                {isEditingOpening && (
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10, paddingTop: 10, borderTop: `1px solid ${COLOR.hairline}` }}>
                    <label style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Solde de départ (avant suivi dans l'app)</label>
                    <input type="number" inputMode="numeric" value={a.openingBalance} onChange={(e) => update(a.id, { openingBalance: Number(e.target.value) || 0 })} style={{ ...inputStyle, width: 150 }} />
                  </div>
                )}
              </div>
            );
          })}
          {!accounts.length && <EmptyState text="Aucun compte." />}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// BUDGETS PAR CATÉGORIE (avec reconduction, transfert)
// ============================================================
function BudgetsTab({ transactions, categoryGroups, budgets, setBudgets, allCategories }: {
  transactions: Transaction[]; categoryGroups: Record<string, Group>; budgets: CategoryBudget[]; setBudgets: (b: CategoryBudget[]) => void; allCategories: string[];
}) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState({ category: categoriesForType(transactions, "Dépense")[0] || "", amount: 100000, rollover: false });
  const [transferFrom, setTransferFrom] = useState<string | null>(null);
  const [transferAmount, setTransferAmount] = useState(0);
  const [transferTo, setTransferTo] = useState("");

  const currentMonth = dateToMonthKey(todayISO());
  const prevDate = new Date(); prevDate.setMonth(prevDate.getMonth() - 1);
  const prevMonth = dateToMonthKey(`${prevDate.getFullYear()}-${pad2(prevDate.getMonth() + 1)}-01`);

  const spentInMonth = (cat: string, monthKey: string) => transactions.filter((t) => t.type === "Dépense" && t.category === cat && dateToMonthKey(t.date) === monthKey).reduce((a, t) => a + t.amount, 0);

  const add = () => { if (!form.category) return; setBudgets([...budgets, { ...form, id: uid("b") }]); setForm({ category: categoriesForType(transactions, "Dépense")[0] || "", amount: 100000, rollover: false }); setAdding(false); };
  const update = (id: string, patch: Partial<CategoryBudget>) => setBudgets(budgets.map((b) => (b.id === id ? { ...b, ...patch } : b)));
  const remove = (id: string) => setBudgets(budgets.filter((b) => b.id !== id));

  const doTransfer = () => {
    if (!transferFrom || !transferTo || transferAmount <= 0) return;
    setBudgets(budgets.map((b) => {
      if (b.id === transferFrom) return { ...b, amount: Math.max(0, b.amount - transferAmount) };
      if (b.id === transferTo) return { ...b, amount: b.amount + transferAmount };
      return b;
    }));
    setTransferFrom(null); setTransferAmount(0); setTransferTo("");
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Budgets par catégorie" subtitle="Limite mensuelle, reconduction du solde non utilisé, transfert entre catégories"
        right={
          <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
            {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Nouveau budget"}
          </button>
        }>
        {adding && (
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Catégorie</label>
              <select style={{ ...inputStyle, width: 180 }} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>
                {categoriesForType(transactions, "Dépense").map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Limite mensuelle</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 150 }} value={form.amount} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) || 0 })} /></div>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: COLOR.inkMuted }}><input type="checkbox" checked={form.rollover} onChange={(e) => setForm({ ...form, rollover: e.target.checked })} /> Reconduction</label>
            <button onClick={add} style={{ background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}>Créer</button>
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {budgets.map((b) => {
            const spent = spentInMonth(b.category, currentMonth);
            const prevSpent = spentInMonth(b.category, prevMonth);
            const carry = b.rollover ? (b.amount - prevSpent) : 0;
            const limit = b.amount + (b.rollover ? carry : 0);
            const pct = Math.min(100, (spent / Math.max(1, limit)) * 100);
            const color = pct < 70 ? COLOR.emerald : pct < 100 ? COLOR.gold : COLOR.clay;
            return (
              <div key={b.id} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 13.5, fontWeight: 500 }}>{b.category}</span>
                    <span style={{ fontSize: 10.5, color: groupColor[categoryGroups[b.category] || "Non classifié"] }}>{categoryGroups[b.category] || "Non classifié"}</span>
                    {b.rollover && <span style={{ fontSize: 10, color: COLOR.slateBlueSoft, border: `1px solid ${COLOR.slateBlue}`, borderRadius: 10, padding: "1px 7px" }}>reconduit</span>}
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={() => setTransferFrom(transferFrom === b.id ? null : b.id)} style={{ fontSize: 11, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "3px 9px", cursor: "pointer" }}>Transférer</button>
                    <button onClick={() => remove(b.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
                  </div>
                </div>
                {transferFrom === b.id && (
                  <div style={{ display: "flex", gap: 8, alignItems: "flex-end", marginBottom: 10, padding: 10, background: COLOR.surface, borderRadius: 6 }}>
                    <select style={{ ...inputStyle, width: 160 }} value={transferTo} onChange={(e) => setTransferTo(e.target.value)}>
                      <option value="">Vers…</option>
                      {budgets.filter((x) => x.id !== b.id).map((x) => <option key={x.id} value={x.id}>{x.category}</option>)}
                    </select>
                    <input type="number" inputMode="numeric" style={{ ...inputStyle, width: 120 }} value={transferAmount || ""} onChange={(e) => setTransferAmount(Number(e.target.value) || 0)} placeholder="Montant" />
                    <button onClick={doTransfer} style={{ background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}>Confirmer</button>
                  </div>
                )}
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 6 }}>
                  <span style={{ color: COLOR.inkMuted }}>{fmt(spent)} / {fmt(limit)} FCFA{b.rollover && carry !== 0 ? ` (dont ${carry > 0 ? "+" : ""}${fmt(carry)} reconduit)` : ""}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color }}>{pct.toFixed(0)}%</span>
                </div>
                <div style={{ height: 8, background: COLOR.hairline, borderRadius: 4, overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: color }} />
                </div>
              </div>
            );
          })}
          {!budgets.length && <EmptyState text="Aucun budget défini." />}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// BÉNÉFICIAIRES (PAYEES)
// ============================================================
function PayeesTab({ transactions }: { transactions: Transaction[] }) {
  const rows = useMemo(() => {
    const m: Record<string, { dep: number; rev: number; count: number }> = {};
    transactions.filter((t) => t.payee && t.payee.trim()).forEach((t) => {
      const p = t.payee!.trim();
      if (!m[p]) m[p] = { dep: 0, rev: 0, count: 0 };
      m[p].count += 1;
      if (t.type === "Dépense") m[p].dep += t.amount; else m[p].rev += t.amount;
    });
    return Object.entries(m).map(([name, d]) => ({ name, ...d })).sort((a, b) => (b.dep + b.rev) - (a.dep + a.rev));
  }, [transactions]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Bénéficiaires & Sociétés" subtitle="Ajoute un bénéficiaire lors de la saisie d'une transaction (Journal ou Saisie du jour) pour le voir apparaître ici">
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {rows.map((r) => (
            <div key={r.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}` }}>
              <div><div style={{ fontSize: 13 }}>{r.name}</div><div style={{ fontSize: 11, color: COLOR.inkMuted }}>{r.count} transaction(s)</div></div>
              <div style={{ display: "flex", gap: 16 }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: COLOR.claySoft }}>−{fmt(r.dep)}</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: COLOR.emeraldSoft }}>+{fmt(r.rev)}</span>
              </div>
            </div>
          ))}
          {!rows.length && <EmptyState text="Aucun bénéficiaire renseigné pour le moment." />}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// RÉCURRENCES & ÉCHÉANCES
// ============================================================
function RecurrencesTab({ recurring, setRecurring, transactions, setTransactions, allCategories, accounts }: {
  recurring: RecurringTemplate[]; setRecurring: (r: RecurringTemplate[]) => void;
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; allCategories: string[]; accounts: Account[];
}) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<RecurringTemplate, "id">>({ category: categoriesForType(transactions, "Dépense")[0] || "", type: "Dépense", amount: 0, frequency: "Mensuelle", nextDate: todayISO(), account: accounts[0]?.name });

  const today = todayISO();
  const upcoming = recurring.filter((r) => daysBetween(today, r.nextDate) <= 14).sort((a, b) => a.nextDate.localeCompare(b.nextDate));

  const add = () => { if (!form.category || form.amount <= 0) return; setRecurring([...recurring, { ...form, id: uid("r") }]); setForm({ category: categoriesForType(transactions, "Dépense")[0] || "", type: "Dépense", amount: 0, frequency: "Mensuelle", nextDate: todayISO(), account: accounts[0]?.name }); setAdding(false); };
  const remove = (id: string) => setRecurring(recurring.filter((r) => r.id !== id));

  const enregistrer = (r: RecurringTemplate) => {
    setTransactions([...transactions, { id: uid(), date: r.nextDate, time: nowTime(), category: r.category, type: r.type, amount: r.amount, account: r.account, payee: r.payee }]);
    setRecurring(recurring.map((x) => (x.id === r.id ? { ...x, nextDate: addInterval(x.nextDate, x.frequency) } : x)));
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Échéances à venir" subtitle="Prochains 14 jours — enregistrez d'un clic quand la transaction se réalise">
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {upcoming.map((r) => {
            const days = daysBetween(today, r.nextDate);
            return (
              <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <Repeat size={14} color={COLOR.goldSoft} />
                  <div>
                    <div style={{ fontSize: 13 }}>{r.category} <span style={{ color: COLOR.inkMuted, fontSize: 11 }}>· {r.frequency}</span></div>
                    <div style={{ fontSize: 11, color: days <= 0 ? COLOR.claySoft : COLOR.inkMuted }}>{days <= 0 ? "Aujourd'hui / en retard" : `dans ${days} jour(s)`} · {dateLabelFull(r.nextDate)}</div>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: r.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(r.amount)}</span>
                  <button onClick={() => enregistrer(r)} style={{ display: "flex", alignItems: "center", gap: 5, background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}><Check size={12} /> Enregistrer</button>
                </div>
              </div>
            );
          })}
          {!upcoming.length && <EmptyState text="Aucune échéance dans les 14 prochains jours." />}
        </div>
      </Panel>

      <Panel title="Modèles de transactions récurrentes" subtitle="Loyer, salaire, abonnements… tout ce qui revient à intervalle régulier"
        right={
          <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
            {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Nouveau modèle"}
          </button>
        }>
        {adding && (
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Catégorie</label>
              <select style={{ ...inputStyle, width: 160 }} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>
                {categoriesForType(transactions, form.type).map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Type</label><select style={inputStyle} value={form.type} onChange={(e) => { const ty = e.target.value as TxType; setForm({ ...form, type: ty, category: categoriesForType(transactions, ty)[0] || "" }); }}><option value="Dépense">Dépense</option><option value="Revenu">Revenu</option></select></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 130 }} value={form.amount || ""} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) })} /></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Fréquence</label><select style={inputStyle} value={form.frequency} onChange={(e) => setForm({ ...form, frequency: e.target.value as RecurringTemplate["frequency"] })}><option>Hebdomadaire</option><option>Mensuelle</option><option>Annuelle</option></select></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Prochaine échéance</label><input type="date" style={inputStyle} value={form.nextDate} onChange={(e) => setForm({ ...form, nextDate: e.target.value })} /></div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte</label>
              <select style={{ ...inputStyle, width: 140 }} value={form.account || ""} onChange={(e) => setForm({ ...form, account: e.target.value })}>
                {!accounts.length && <option value="">Aucun compte</option>}
                {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
              </select>
            </div>
            <button onClick={add} style={{ background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}>Créer</button>
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {recurring.map((r) => (
            <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: COLOR.surfaceRaised, borderRadius: 6 }}>
              <span style={{ fontSize: 12.5 }}>{r.category} · {r.frequency} · prochaine le {dateLabelFull(r.nextDate)}</span>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: r.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(r.amount)}</span>
                <button onClick={() => remove(r.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
              </div>
            </div>
          ))}
          {!recurring.length && <EmptyState text="Aucun modèle récurrent." />}
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// SAUVEGARDE & RESTAURATION
// ============================================================
function SauvegardeTab({ getSnapshot, restore, syncCode, setSyncCode, syncStatus, lastSyncedAt, onForceSync, realtimeConnected }: {
  getSnapshot: () => any; restore: (data: any) => void; syncCode: string; setSyncCode: (c: string) => void;
  syncStatus: "idle" | "syncing" | "synced" | "error" | "disabled"; lastSyncedAt: string | null; onForceSync: () => void; realtimeConnected: boolean;
}) {
  const [status, setStatus] = useState<string | null>(null);
  const [linkCopied, setLinkCopied] = useState(false);
  const [codeInput, setCodeInput] = useState(syncCode);
  const download = () => {
    const data = getSnapshot();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `grand-livre-sauvegarde-${todayISO()}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setStatus("Sauvegarde téléchargée.");
  };
  const onFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result as string);
        restore(data);
        setStatus("Données restaurées avec succès.");
      } catch {
        setStatus("Fichier invalide — restauration annulée.");
      }
    };
    reader.readAsText(file);
  };

  const generateCode = () => {
    const words = ["atelier", "lagune", "baobab", "kora", "orage", "savane", "azur", "grelot", "brume", "corail"];
    const code = `${words[Math.floor(Math.random() * words.length)]}-${words[Math.floor(Math.random() * words.length)]}-${Math.floor(Math.random() * 900 + 100)}`;
    setCodeInput(code);
  };

  const statusLabel: Record<string, { text: string; color: string }> = {
    idle: { text: "Non connecté", color: COLOR.inkMuted },
    syncing: { text: "Synchronisation…", color: COLOR.goldSoft },
    synced: { text: `Synchronisé${lastSyncedAt ? " à " + lastSyncedAt : ""}`, color: COLOR.emeraldSoft },
    error: { text: "Erreur de synchronisation", color: COLOR.claySoft },
    disabled: { text: "Non configuré", color: COLOR.inkMuted },
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Synchronisation entre appareils" subtitle="Un même code de synchronisation, saisi sur ton iPhone et ton ordinateur, relie automatiquement tes données">
        {syncStatus === "disabled" ? (
          <div style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.7, display: "flex", gap: 10 }}>
            <Info size={16} color={COLOR.gold} style={{ flexShrink: 0, marginTop: 2 }} />
            <span>
              La synchronisation cloud n'est pas configurée sur ce déploiement — les variables d'environnement
              <code style={{ color: COLOR.goldSoft }}> VITE_SUPABASE_URL</code> et <code style={{ color: COLOR.goldSoft }}>VITE_SUPABASE_ANON_KEY</code> sont
              absentes. Ajoute-les dans les réglages Vercel du projet, puis redéploie. En attendant, utilise
              l'export/import manuel ci-dessous pour transférer tes données entre appareils.
            </span>
          </div>
        ) : (
          <>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: statusLabel[syncStatus].color, display: "inline-block" }} />
              <span style={{ fontSize: 12.5, color: statusLabel[syncStatus].color }}>{statusLabel[syncStatus].text}</span>
              {syncCode && (
                <span style={{
                  display: "flex", alignItems: "center", gap: 5, fontSize: 11, padding: "3px 9px", borderRadius: 20,
                  border: `1px solid ${realtimeConnected ? COLOR.emerald : COLOR.hairline}`,
                  color: realtimeConnected ? COLOR.emeraldSoft : COLOR.inkMuted,
                  background: realtimeConnected ? "rgba(63,156,122,0.1)" : "transparent",
                }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: realtimeConnected ? COLOR.emerald : COLOR.inkMuted }} />
                  {realtimeConnected ? "Temps réel actif" : "Temps réel indisponible"}
                </span>
              )}
            </div>
            <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: 1, minWidth: 220 }}>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Code de synchronisation</label>
                <input style={inputStyle} value={codeInput} onChange={(e) => setCodeInput(e.target.value)} placeholder="ex: atelier-lagune-482"
                  autoCapitalize="off" autoCorrect="off" spellCheck={false} autoComplete="off" />
              </div>
              <button onClick={generateCode} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 12px", fontSize: 12, cursor: "pointer", height: 34 }}>Générer</button>
              <button onClick={() => setSyncCode(codeInput.trim().toLowerCase())} disabled={!codeInput.trim()} style={{ background: codeInput.trim() ? COLOR.emerald : COLOR.hairline, border: "none", borderRadius: 6, color: codeInput.trim() ? COLOR.bg : COLOR.inkMuted, padding: "9px 16px", fontSize: 12.5, fontWeight: 600, cursor: codeInput.trim() ? "pointer" : "default", height: 34 }}>
                {syncCode ? "Mettre à jour" : "Se connecter"}
              </button>
              {syncCode && (
                <button onClick={() => { setSyncCode(""); setCodeInput(""); }} style={{ background: "transparent", border: `1px solid ${COLOR.clay}`, borderRadius: 6, color: COLOR.claySoft, padding: "9px 16px", fontSize: 12.5, cursor: "pointer", height: 34 }}>
                  Déconnecter
                </button>
              )}
            </div>
            {syncCode && (
              <div style={{ marginTop: 14, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <button onClick={onForceSync} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
                  <RotateCcw size={12} /> Forcer la synchronisation
                </button>
                <span style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Saisis exactement le même code sur ton autre appareil pour le relier.</span>
              </div>
            )}
            {syncCode && !realtimeConnected && (
              <div style={{ marginTop: 10, fontSize: 11.5, color: COLOR.inkMuted, display: "flex", gap: 8, alignItems: "flex-start" }}>
                <Info size={13} color={COLOR.gold} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>
                  Sans temps réel actif, les mises à jour d'un autre appareil apparaissent quand même — juste au prochain
                  chargement de la page, pas instantanément. Pour activer le temps réel : dans Supabase, <b>Database → Replication</b>,
                  active la réplication pour la table <code style={{ color: COLOR.goldSoft }}>app_state</code>.
                </span>
              </div>
            )}
            {syncCode && (
              <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px solid ${COLOR.hairline}` }}>
                <div style={{ fontSize: 12, color: COLOR.goldSoft, marginBottom: 6 }}>💡 Astuce anti-déconnexion (iPhone)</div>
                <div style={{ fontSize: 11.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 10 }}>
                  Safari sur iPhone efface parfois la mémoire du navigateur, ce qui oblige à retaper le code. Mets ce lien
                  en favori (ou sur l'écran d'accueil) à la place de l'adresse simple — il contient ton code et te reconnecte
                  automatiquement à chaque ouverture, même après un effacement.
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <code style={{ flex: 1, minWidth: 220, background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, padding: "8px 10px", fontSize: 11, color: COLOR.ink, overflowX: "auto", whiteSpace: "nowrap", display: "block" }}>
                    {typeof window !== "undefined" ? `${window.location.origin}${window.location.pathname}?sync=${syncCode}` : ""}
                  </code>
                  <button onClick={() => {
                    const link = `${window.location.origin}${window.location.pathname}?sync=${syncCode}`;
                    navigator.clipboard?.writeText(link).then(() => { setLinkCopied(true); setTimeout(() => setLinkCopied(false), 1500); });
                  }} style={{ display: "flex", alignItems: "center", gap: 6, background: linkCopied ? COLOR.emerald : "rgba(201,162,39,0.14)", border: `1px solid ${linkCopied ? COLOR.emerald : COLOR.gold}`, borderRadius: 6, color: linkCopied ? COLOR.bg : COLOR.goldSoft, padding: "8px 14px", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>
                    {linkCopied ? <Check size={13} /> : <ClipboardList size={13} />} {linkCopied ? "Copié" : "Copier le lien"}
                  </button>
                </div>
              </div>
            )}
          </>
        )}
      </Panel>
      <Panel title="Sauvegarde" subtitle="Toutes tes données vivent uniquement dans ce navigateur — exporte-les régulièrement pour ne rien perdre">
        <button onClick={download} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(63,156,122,0.14)", border: `1px solid ${COLOR.emerald}`, borderRadius: 8, color: COLOR.emeraldSoft, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
          <Download size={15} /> Télécharger une sauvegarde complète (.json)
        </button>
      </Panel>
      <Panel title="Restauration" subtitle="Remplace toutes les données actuelles par celles du fichier importé">
        <label style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, color: COLOR.inkMuted, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
          <UploadCloud size={15} /> Choisir un fichier de sauvegarde
          <input type="file" accept="application/json" onChange={onFile} style={{ display: "none" }} />
        </label>
        {status && <div style={{ marginTop: 12, fontSize: 12.5, color: COLOR.goldSoft }}>{status}</div>}
      </Panel>
    </div>
  );
}

// ============================================================ END OF PART 5 — continued below
// ============================================================
// JOURNAL TAB (CRUD + import texte + règles de catégorisation)
// ============================================================
function emptyForm(transactions: Transaction[], accounts: Account[]): Omit<Transaction, "id"> {
  return { date: todayISO(), time: nowTime(), category: categoriesForType(transactions, "Dépense")[0] || "Cadeaux", type: "Dépense", amount: 0, account: accounts[0]?.name };
}

function JournalTab({ filtered, allCategories, categoryGroups, transactions, setTransactions, rules, setRules, accounts }: {
  filtered: any[]; allCategories: string[]; categoryGroups: Record<string, Group>;
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void;
  rules: CategorizationRule[]; setRules: (r: CategorizationRule[]) => void; accounts: Account[];
}) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Transaction, "id">>(emptyForm(transactions, accounts));
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [editingTx, setEditingTx] = useState<Transaction | null>(null);
  const [page, setPage] = useState(0);
  const [showImport, setShowImport] = useState(false);
  const [importText, setImportText] = useState("");
  const [importPreview, setImportPreview] = useState<Transaction[]>([]);
  const [showRules, setShowRules] = useState(false);
  const [newRule, setNewRule] = useState({ keyword: "", group: "Non classifié" as Group });
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const [confirmBulkDelete, setConfirmBulkDelete] = useState(false);
  const [bulkCategory, setBulkCategory] = useState("");
  const pageSize = 25;

  const sorted = useMemo(() => filtered.slice().sort((a, b) => b.date.localeCompare(a.date)), [filtered]);
  const pageCount = Math.max(1, Math.ceil(sorted.length / pageSize));
  const pageRows = sorted.slice(page * pageSize, page * pageSize + pageSize);

  const toggleSelect = (id: string) => setSelected((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const allPageSelected = pageRows.length > 0 && pageRows.every((t) => selected.has(t.id));
  const toggleSelectAllPage = () => setSelected((prev) => {
    const n = new Set(prev);
    if (allPageSelected) pageRows.forEach((t) => n.delete(t.id)); else pageRows.forEach((t) => n.add(t.id));
    return n;
  });
  const bulkDelete = () => { setTransactions(transactions.filter((t) => !selected.has(t.id))); setSelected(new Set()); };
  const bulkChangeCategory = () => {
    if (!bulkCategory) return;
    setTransactions(transactions.map((t) => (selected.has(t.id) ? { ...t, category: bulkCategory } : t)));
    setSelected(new Set()); setBulkCategory("");
  };

  const addTransaction = () => {
    if (!form.category || form.amount <= 0) return;
    setTransactions([...transactions, { ...form, id: uid() }]);
    setForm(emptyForm(transactions, accounts)); setAdding(false);
  };
  const startEdit = (t: Transaction) => setEditingTx(t);
  const saveEdit = (t: Transaction) => setTransactions(transactions.map((x) => (x.id === t.id ? t : x)));
  const remove = (id: string) => setTransactions(transactions.filter((t) => t.id !== id));

  // parsing du texte collé (format MoneyCoach : "2026_1 Logement Dépense 480500 €")
  const parseImport = (text: string) => {
    const lines = text.split("\n");
    const out: Transaction[] = [];
    const re = /^(\d{4}_\d{1,2})\s+(.+?)\s+(Dépense|Revenu)\s+(-?\d[\d\s]*)\s*€?\s*$/;
    lines.forEach((line) => {
      const m = line.trim().match(re);
      if (m) {
        const amount = parseInt(m[4].replace(/\s/g, ""), 10);
        if (!isNaN(amount)) out.push({ id: uid("imp"), date: monthKeyToFirstDate(m[1]), category: m[2].trim(), type: m[3] as TxType, amount });
      }
    });
    return out;
  };

  const runPreview = (text: string) => { setImportText(text); setImportPreview(parseImport(text)); };
  const confirmImport = () => { setTransactions([...transactions, ...importPreview]); setImportPreview([]); setImportText(""); setShowImport(false); };

  const addRule = () => { if (!newRule.keyword) return; setRules([...rules, { id: uid("r"), keyword: newRule.keyword, group: newRule.group }]); setNewRule({ keyword: "", group: "Non classifié" }); };
  const removeRule = (id: string) => setRules(rules.filter((r) => r.id !== id));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Journal des transactions" subtitle={`${sorted.length} transaction(s) filtrée(s) sur ${transactions.length} au total`}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => setShowRules((r) => !r)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
              <SlidersHorizontal size={13} /> Règles
            </button>
            <button onClick={() => setShowImport((s) => !s)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
              <Upload size={13} /> Importer
            </button>
            <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
              {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Ajouter"}
            </button>
          </div>
        }>

        {showRules && (
          <div style={{ padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ fontSize: 12.5, color: COLOR.inkMuted, marginBottom: 12 }}>
              Règles de catégorisation automatique — appliquées aux nouvelles catégories (mot-clé contenu dans le nom, insensible à la casse).
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
              {rules.map((r) => (
                <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5 }}>
                  <span>si la catégorie contient <b>"{r.keyword}"</b> → <span style={{ color: groupColor[r.group] }}>{r.group}</span></span>
                  <button onClick={() => removeRule(r.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={12} /></button>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
              <input style={{ ...inputStyle, width: 160 }} placeholder="mot-clé" value={newRule.keyword} onChange={(e) => setNewRule({ ...newRule, keyword: e.target.value })} />
              <select style={{ ...inputStyle, width: 140 }} value={newRule.group} onChange={(e) => setNewRule({ ...newRule, group: e.target.value as Group })}>
                {GROUPS.map((g) => <option key={g} value={g}>{g}</option>)}
              </select>
              <button onClick={addRule} style={{ background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>Ajouter</button>
            </div>
          </div>
        )}

        {showImport && (
          <div style={{ padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ fontSize: 12.5, color: COLOR.inkMuted, marginBottom: 10 }}>
              Collez le texte copié d'un export MoneyCoach (format : <code style={{ color: COLOR.goldSoft }}>2026_1 Logement Dépense 480500 €</code>, une ligne par transaction).
            </div>
            <textarea value={importText} onChange={(e) => runPreview(e.target.value)} rows={6}
              placeholder="2026_8 Logement Dépense 480000 €&#10;2026_8 Un salaire Revenu 1629000 €"
              style={{ ...inputStyle, resize: "vertical", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
              <span style={{ fontSize: 12, color: COLOR.inkMuted }}>{importPreview.length} transaction(s) reconnue(s)</span>
              <button onClick={confirmImport} disabled={!importPreview.length} style={{ display: "flex", alignItems: "center", gap: 6, background: importPreview.length ? COLOR.emerald : COLOR.hairline, border: "none", borderRadius: 6, color: importPreview.length ? COLOR.bg : COLOR.inkMuted, padding: "8px 14px", fontSize: 12.5, cursor: importPreview.length ? "pointer" : "default" }}>
                <Save size={13} /> Importer {importPreview.length} transaction(s)
              </button>
            </div>
          </div>
        )}

        {adding && (
          <div style={{ padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Date</label><input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Heure</label><input type="time" style={{ ...inputStyle, width: 100 }} value={form.time || ""} onChange={(e) => setForm({ ...form, time: e.target.value })} /></div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Catégorie</label>
                <select style={{ ...inputStyle, width: 180 }} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value, subcategory: "" })}>
                  {categoriesForType(transactions, form.type).map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              {getSubcategories(form.type, form.category).length > 0 && (
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Sous-catégorie</label>
                  <select style={{ ...inputStyle, width: 160 }} value={form.subcategory || ""} onChange={(e) => setForm({ ...form, subcategory: e.target.value })}>
                    <option value="">—</option>
                    {getSubcategories(form.type, form.category).map((s) => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Type</label><select style={inputStyle} value={form.type} onChange={(e) => { const ty = e.target.value as TxType; setForm({ ...form, type: ty, subcategory: "", category: categoriesForType(transactions, ty)[0] || "" }); }}><option value="Dépense">Dépense</option><option value="Revenu">Revenu</option></select></div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant (FCFA)</label><input style={inputStyle} type="number" inputMode="numeric" value={form.amount || ""} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) })} /></div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte</label>
                <select style={{ ...inputStyle, width: 150 }} value={form.account || ""} onChange={(e) => setForm({ ...form, account: e.target.value })}>
                  {!accounts.length && <option value="">Aucun compte créé</option>}
                  {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                </select>
              </div>
              <button onClick={() => setShowAdvanced((s) => !s)} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 12px", fontSize: 12, cursor: "pointer", height: 32 }}>{showAdvanced ? "− options" : "+ bénéficiaire / note"}</button>
              <button onClick={addTransaction} style={{ display: "flex", alignItems: "center", gap: 6, background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}><Save size={13} /> Enregistrer</button>
            </div>
            {showAdvanced && (
              <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Bénéficiaire</label><input style={{ ...inputStyle, width: 180 }} value={form.payee || ""} onChange={(e) => setForm({ ...form, payee: e.target.value })} /></div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4, flex: 1 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Note</label><input style={inputStyle} value={form.note || ""} onChange={(e) => setForm({ ...form, note: e.target.value })} /></div>
              </div>
            )}
          </div>
        )}

        {selected.size > 0 && (
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", background: "rgba(201,162,39,0.1)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, marginBottom: 14, flexWrap: "wrap" }}>
            <span style={{ fontSize: 12.5, color: COLOR.goldSoft }}>{selected.size} sélectionnée(s)</span>
            <select style={{ ...inputStyle, width: 170 }} value={bulkCategory} onChange={(e) => setBulkCategory(e.target.value)}>
              <option value="">Changer la catégorie…</option>
              {allCategories.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <button onClick={bulkChangeCategory} disabled={!bulkCategory} style={{ background: bulkCategory ? COLOR.emerald : COLOR.hairline, border: "none", borderRadius: 6, color: bulkCategory ? COLOR.bg : COLOR.inkMuted, padding: "6px 12px", fontSize: 11.5, cursor: bulkCategory ? "pointer" : "default" }}>Appliquer</button>
            <button onClick={() => setConfirmBulkDelete(true)} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: `1px solid ${COLOR.clay}`, borderRadius: 6, color: COLOR.claySoft, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}><Trash2 size={12} /> Supprimer la sélection</button>
            <button onClick={() => setSelected(new Set())} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, fontSize: 11.5, cursor: "pointer" }}>Désélectionner</button>
          </div>
        )}

        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={{ padding: "8px 10px", borderBottom: `1px solid ${COLOR.hairline}` }}>
                <button onClick={toggleSelectAllPage} style={{ background: "transparent", border: "none", cursor: "pointer", display: "flex" }}>
                  {allPageSelected ? <CheckSquare size={14} color={COLOR.goldSoft} /> : <Square size={14} color={COLOR.inkMuted} />}
                </button>
              </th>
              {["Date", "Catégorie", "Type", "Groupe", "Montant", ""].map((h, i) => (
              <th key={h} style={{ textAlign: i === 4 ? "right" : "left", padding: "8px 10px", fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", borderBottom: `1px solid ${COLOR.hairline}` }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {pageRows.map((t) => {
                return (
                  <tr key={t.id}>
                    <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}` }}>
                      <button onClick={() => toggleSelect(t.id)} style={{ background: "transparent", border: "none", cursor: "pointer", display: "flex" }}>
                        {selected.has(t.id) ? <CheckSquare size={14} color={COLOR.goldSoft} /> : <Square size={14} color={COLOR.inkMuted} />}
                      </button>
                    </td>
                    <td style={{ padding: "9px 10px", fontSize: 12.5, borderBottom: `1px solid ${COLOR.hairline}`, fontFamily: "'IBM Plex Mono', monospace" }}>
                      {dateLabelFull(t.date)}{t.time && <div style={{ fontSize: 10.5, color: COLOR.inkMuted }}>{t.time}</div>}
                    </td>
                    <td style={{ padding: "9px 10px", fontSize: 12.5, borderBottom: `1px solid ${COLOR.hairline}` }}>
                      {t.category}{t.subcategory && <span style={{ color: COLOR.inkMuted }}> · {t.subcategory}</span>}
                      {t.payee && <div style={{ fontSize: 10.5, color: COLOR.inkMuted }}>{t.payee}</div>}
                    </td>
                    <td style={{ padding: "9px 10px", fontSize: 12.5, borderBottom: `1px solid ${COLOR.hairline}`, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{t.type}</td>
                    <td style={{ padding: "9px 10px", fontSize: 11.5, borderBottom: `1px solid ${COLOR.hairline}`, color: groupColor[t.group] }}>
                      {t.group}
                      {t.account ? (
                        <div style={{ color: COLOR.inkMuted, fontSize: 10.5, marginTop: 2 }}>{t.account}</div>
                      ) : (
                        <div style={{ color: COLOR.claySoft, fontSize: 10.5, marginTop: 2, display: "flex", alignItems: "center", gap: 3 }}><AlertTriangle size={9} /> sans compte</div>
                      )}
                    </td>
                    <td style={{ padding: "9px 10px", fontSize: 12.5, textAlign: "right", borderBottom: `1px solid ${COLOR.hairline}`, fontFamily: "'IBM Plex Mono', monospace" }}>{fmt(t.amount)}</td>
                    <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}`, whiteSpace: "nowrap" }}><button onClick={() => startEdit(t)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={13} /></button><button onClick={() => setConfirmDeleteId(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button></td>
                  </tr>
                );
              })}
              {!pageRows.length && <tr><td colSpan={7}><EmptyState /></td></tr>}
            </tbody>
          </table>
        </div>
        {pageCount > 1 && (
          <div style={{ display: "flex", justifyContent: "center", gap: 8, marginTop: 16 }}>
            <button disabled={page === 0} onClick={() => setPage((p) => p - 1)} style={pagerBtn(page === 0)}>Précédent</button>
            <span style={{ fontSize: 12, color: COLOR.inkMuted, alignSelf: "center" }}>Page {page + 1} / {pageCount}</span>
            <button disabled={page >= pageCount - 1} onClick={() => setPage((p) => p + 1)} style={pagerBtn(page >= pageCount - 1)}>Suivant</button>
          </div>
        )}
      </Panel>
      <TransactionEditSheet
        open={!!editingTx}
        transaction={editingTx}
        transactions={transactions}
        accounts={accounts}
        onClose={() => setEditingTx(null)}
        onSave={saveEdit}
        onDelete={remove}
      />
      <ConfirmDialog
        open={!!confirmDeleteId}
        title="Supprimer cette transaction ?"
        message="Cette action est définitive. Le montant ne sera plus comptabilisé nulle part dans l'app."
        onConfirm={() => { if (confirmDeleteId) remove(confirmDeleteId); setConfirmDeleteId(null); }}
        onCancel={() => setConfirmDeleteId(null)}
      />
      <ConfirmDialog
        open={confirmBulkDelete}
        title={`Supprimer ${selected.size} transaction(s) ?`}
        message="Cette action est définitive et concerne toutes les transactions actuellement sélectionnées."
        onConfirm={() => { bulkDelete(); setConfirmBulkDelete(false); }}
        onCancel={() => setConfirmBulkDelete(false)}
      />
    </div>
  );
}

// ============================================================
// RAPPORTS & EXPORT TAB
// ============================================================
function ExportTab({ filtered, filters, setFilters, allMonths }: { filtered: any[]; filters: Filters; setFilters: (f: Filters) => void; allMonths: string[] }) {
  const [pdfState, setPdfState] = useState<"idle" | "loading" | "error">("idle");
  const totalRevenus = filtered.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const totalDepenses = filtered.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const solde = totalRevenus - totalDepenses;

  const byMonth = useMemo(() => {
    const m: Record<string, { revenus: number; depenses: number }> = {};
    filtered.forEach((t) => { if (!m[t.month]) m[t.month] = { revenus: 0, depenses: 0 }; if (t.type === "Revenu") m[t.month].revenus += t.amount; else m[t.month].depenses += t.amount; });
    return Object.keys(m).sort((a, b) => monthSortKey(a) - monthSortKey(b)).map((k) => ({ key: k, ...m[k] }));
  }, [filtered]);

  const cur = byMonth[byMonth.length - 1];
  const prev = byMonth[byMonth.length - 2];

  const summary = useMemo(() => {
    if (!cur) return "Aucune donnée disponible pour la période filtrée.";
    let s = `En ${monthLabel(cur.key)}, les revenus se sont élevés à ${fmt(cur.revenus)} FCFA et les dépenses à ${fmt(cur.depenses)} FCFA, soit un solde de ${fmt(cur.revenus - cur.depenses)} FCFA. `;
    if (prev) {
      const depDelta = ((cur.depenses - prev.depenses) / (prev.depenses || 1)) * 100;
      s += `Les dépenses ont ${depDelta >= 0 ? "augmenté" : "diminué"} de ${Math.abs(depDelta).toFixed(0)}% par rapport au mois précédent. `;
    }
    const nonProd = filtered.filter((t) => t.type === "Dépense" && t.group === "Non-productif" && t.month === cur.key).reduce((a, t) => a + t.amount, 0);
    if (cur.revenus > 0) s += `Les dépenses non-productives représentent ${((nonProd / cur.revenus) * 100).toFixed(0)}% du revenu de ce mois.`;
    return s;
  }, [cur, prev, filtered]);

  const exportCSV = () => {
    const header = "Date,Heure,Mois,Catégorie,Sous-catégorie,Type,Groupe,Compte,Montant\n";
    const rows = filtered.map((t) => `${t.date},${t.time || ""},${monthLabel(t.month)},"${t.category}","${t.subcategory || ""}",${t.type},${t.group},"${t.account || ""}",${t.amount}`).join("\n");
    const blob = new Blob([header + rows], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `grand-livre_${filters.from}_${filters.to}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportExcelSimple = () => {
    const wb = XLSX.utils.book_new();
    const byMonthXl: Record<string, { revenus: number; depenses: number }> = {};
    filtered.forEach((t) => {
      if (!byMonthXl[t.month]) byMonthXl[t.month] = { revenus: 0, depenses: 0 };
      if (t.type === "Revenu") byMonthXl[t.month].revenus += t.amount; else byMonthXl[t.month].depenses += t.amount;
    });
    const summaryRows: any[][] = [
      ["Grand Livre — Rapport financier"], ["Période", `${monthLabel(filters.from)} — ${monthLabel(filters.to)}`],
      ["Généré le", dateLabelFull(todayISO())], [], ["Revenus", totalRevenus], ["Dépenses", totalDepenses],
      ["Solde", solde], ["Nombre de transactions", filtered.length],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    wsSummary["!cols"] = [{ wch: 26 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Résumé");
    const txHeader = ["Date", "Heure", "Catégorie", "Sous-catégorie", "Type", "Groupe", "Compte", "Bénéficiaire", "Note", "Montant (FCFA)"];
    const txRows = filtered.slice().sort((a, b) => b.date.localeCompare(a.date))
      .map((t) => [t.date, t.time || "", t.category, t.subcategory || "", t.type, t.group, t.account || "", t.payee || "", t.note || "", t.amount]);
    const wsTx = XLSX.utils.aoa_to_sheet([txHeader, ...txRows]);
    wsTx["!cols"] = [{ wch: 12 }, { wch: 8 }, { wch: 22 }, { wch: 18 }, { wch: 10 }, { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsTx, "Transactions");
    const monthHeader = ["Mois", "Revenus", "Dépenses", "Solde"];
    const monthRows = Object.keys(byMonthXl).sort((a, b) => monthSortKey(a) - monthSortKey(b))
      .map((k) => [monthLabel(k), byMonthXl[k].revenus, byMonthXl[k].depenses, byMonthXl[k].revenus - byMonthXl[k].depenses]);
    const wsMonth = XLSX.utils.aoa_to_sheet([monthHeader, ...monthRows]);
    wsMonth["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsMonth, "Par mois");
    const catTotals: Record<string, { value: number; type: string }> = {};
    filtered.forEach((t) => { if (!catTotals[t.category]) catTotals[t.category] = { value: 0, type: t.type }; catTotals[t.category].value += t.amount; });
    const catHeader = ["Catégorie", "Type", "Total (FCFA)"];
    const catRows = Object.entries(catTotals).sort((a, b) => b[1].value - a[1].value).map(([name, d]) => [name, d.type, d.value]);
    const wsCat = XLSX.utils.aoa_to_sheet([catHeader, ...catRows]);
    wsCat["!cols"] = [{ wch: 26 }, { wch: 10 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsCat, "Par catégorie");
    XLSX.writeFile(wb, `grand-livre_${filters.from}_${filters.to}.xlsx`);
  };

  const exportExcel = async () => {
    let ExcelJS: any;
    try {
      ExcelJS = await import(/* @vite-ignore */ "exceljs");
    } catch {
      exportExcelSimple();
      return;
    }
    const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", SUBTLE = "FF232F27", WHITE = "FFFFFFFF", MUTED = "FF8A9A8E";
    const wb = new ExcelJS.Workbook();
    wb.creator = "Grand Livre"; wb.created = new Date();

    const styleHeaderRow = (row: any) => {
      row.eachCell((c: any) => {
        c.font = { bold: true, color: { argb: WHITE } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
        c.alignment = { vertical: "middle" };
      });
      row.height = 22;
    };
    const styleSubtotalRow = (row: any) => {
      row.eachCell((c: any) => {
        c.font = { bold: true, color: { argb: GOLD } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUBTLE } };
      });
    };

    // ===== Feuille Résumé =====
    const byMonthXl: Record<string, { revenus: number; depenses: number }> = {};
    filtered.forEach((t) => {
      if (!byMonthXl[t.month]) byMonthXl[t.month] = { revenus: 0, depenses: 0 };
      if (t.type === "Revenu") byMonthXl[t.month].revenus += t.amount; else byMonthXl[t.month].depenses += t.amount;
    });
    const wsSummary = wb.addWorksheet("Résumé");
    wsSummary.columns = [{ width: 30 }, { width: 24 }];
    wsSummary.mergeCells("A1:B1");
    const titleCell = wsSummary.getCell("A1");
    titleCell.value = "Grand Livre — Rapport financier";
    titleCell.font = { bold: true, size: 15, color: { argb: WHITE } };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    titleCell.alignment = { vertical: "middle", indent: 1 };
    wsSummary.getRow(1).height = 30;
    const addSum = (label: string, value: any, color?: string, fmtNum?: boolean) => {
      const r = wsSummary.addRow([label, value]);
      r.getCell(1).font = { color: { argb: MUTED } };
      r.getCell(2).font = { bold: true, color: { argb: color || NAVY } };
      if (fmtNum) r.getCell(2).numFmt = "#,##0";
      r.getCell(2).alignment = { horizontal: "right" };
    };
    wsSummary.addRow([]);
    addSum("Période", `${monthLabel(filters.from)} — ${monthLabel(filters.to)}`);
    addSum("Généré le", dateLabelFull(todayISO()));
    wsSummary.addRow([]);
    addSum("Revenus", totalRevenus, EMERALD, true);
    addSum("Dépenses", totalDepenses, CLAY, true);
    addSum("Solde", solde, solde >= 0 ? EMERALD : CLAY, true);
    addSum("Taux d'épargne", `${(totalRevenus > 0 ? (solde / totalRevenus) * 100 : 0).toFixed(1)}%`);
    addSum("Nombre de transactions", filtered.length);

    // ===== Feuille Transactions, avec sous-total automatique à chaque changement de mois =====
    const wsTx = wb.addWorksheet("Transactions");
    wsTx.columns = [
      { header: "Date", key: "date", width: 12 }, { header: "Heure", key: "time", width: 8 },
      { header: "Catégorie", key: "cat", width: 22 }, { header: "Sous-catégorie", key: "sub", width: 18 },
      { header: "Type", key: "type", width: 10 }, { header: "Groupe", key: "group", width: 16 },
      { header: "Compte", key: "account", width: 14 }, { header: "Montant (FCFA)", key: "amount", width: 16 },
    ];
    styleHeaderRow(wsTx.getRow(1));
    const sortedTx = filtered.slice().sort((a, b) => a.date.localeCompare(b.date));
    let curMonth: string | null = null, monthRev = 0, monthDep = 0;
    const flushMonthSubtotal = () => {
      if (curMonth === null) return;
      const r = wsTx.addRow({ cat: `— Sous-total ${monthLabel(curMonth)} —`, account: `Rev: ${fmt(monthRev)}`, amount: monthDep });
      styleSubtotalRow(r);
      r.getCell("amount").numFmt = "#,##0";
    };
    sortedTx.forEach((t) => {
      if (curMonth !== null && t.month !== curMonth) { flushMonthSubtotal(); monthRev = 0; monthDep = 0; }
      curMonth = t.month;
      if (t.type === "Revenu") monthRev += t.amount; else monthDep += t.amount;
      const r = wsTx.addRow({ date: t.date, time: t.time || "", cat: t.category, sub: t.subcategory || "", type: t.type, group: t.group, account: t.account || "", amount: t.amount });
      r.getCell("amount").font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY } };
      r.getCell("amount").numFmt = "#,##0";
    });
    flushMonthSubtotal();

    // ===== Feuille Par mois =====
    const wsMonth = wb.addWorksheet("Par mois");
    wsMonth.columns = [{ header: "Mois", key: "mois", width: 14 }, { header: "Revenus", key: "rev", width: 16 }, { header: "Dépenses", key: "dep", width: 16 }, { header: "Solde", key: "solde", width: 16 }];
    styleHeaderRow(wsMonth.getRow(1));
    Object.keys(byMonthXl).sort((a, b) => monthSortKey(a) - monthSortKey(b)).forEach((k) => {
      const r = wsMonth.addRow({ mois: monthLabel(k), rev: byMonthXl[k].revenus, dep: byMonthXl[k].depenses, solde: byMonthXl[k].revenus - byMonthXl[k].depenses });
      r.getCell("rev").font = { color: { argb: EMERALD } }; r.getCell("rev").numFmt = "#,##0";
      r.getCell("dep").font = { color: { argb: CLAY } }; r.getCell("dep").numFmt = "#,##0";
      const s = byMonthXl[k].revenus - byMonthXl[k].depenses;
      r.getCell("solde").font = { bold: true, color: { argb: s >= 0 ? EMERALD : CLAY } }; r.getCell("solde").numFmt = "#,##0";
    });

    // ===== Feuille Par catégorie, groupée avec sous-total par groupe =====
    const wsCat = wb.addWorksheet("Par catégorie");
    wsCat.columns = [{ header: "Catégorie", key: "cat", width: 26 }, { header: "Groupe", key: "group", width: 18 }, { header: "Total (FCFA)", key: "total", width: 16 }];
    styleHeaderRow(wsCat.getRow(1));
    const groupOrder = ["Revenu", ...GROUPS];
    groupOrder.forEach((g) => {
      const items = filtered.filter((t) => t.group === g);
      if (!items.length) return;
      const byCat: Record<string, number> = {};
      items.forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
      const groupTitleRow = wsCat.addRow({ cat: g.toUpperCase() });
      groupTitleRow.getCell("cat").font = { bold: true, color: { argb: g === "Revenu" ? EMERALD : g === "Non-productif" ? CLAY : GOLD } };
      Object.entries(byCat).sort((a, b) => b[1] - a[1]).forEach(([name, val]) => {
        const r = wsCat.addRow({ cat: `   ${name}`, group: g, total: val });
        r.getCell("total").numFmt = "#,##0";
      });
      const groupTotal = Object.values(byCat).reduce((a, b) => a + b, 0);
      const subRow = wsCat.addRow({ cat: `— Sous-total ${g} —`, total: groupTotal });
      styleSubtotalRow(subRow);
      subRow.getCell("total").numFmt = "#,##0";
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `grand-livre_${filters.from}_${filters.to}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportPDF = async () => {
    setPdfState("loading");
    try {
      const [{ default: jsPDF }, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      const autoTable = (autoTableModule as any).default || autoTableModule;
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();

      // Bandeau d'en-tête coloré
      doc.setFillColor(26, 43, 76);
      doc.rect(0, 0, pageWidth, 34, "F");
      doc.setFontSize(17);
      doc.setTextColor(255, 255, 255);
      doc.text("Grand Livre — Rapport financier", 14, 16);
      doc.setFontSize(9.5);
      doc.setTextColor(200, 210, 225);
      doc.text(`Période : ${monthLabel(filters.from)} — ${monthLabel(filters.to)}  ·  Généré le ${dateLabelFull(todayISO())}`, 14, 24);
      doc.text(`${filtered.length} transaction(s)`, 14, 29);

      // Cartes KPI colorées
      const kpiY = 40, kpiH = 20, kpiW = (pageWidth - 28 - 16) / 3;
      const drawKpiBox = (x: number, label: string, value: string, r: number, g: number, b: number) => {
        doc.setFillColor(r, g, b); doc.setDrawColor(r, g, b);
        doc.roundedRect(x, kpiY, kpiW, kpiH, 2, 2, "F");
        doc.setFontSize(7.5); doc.setTextColor(255, 255, 255);
        doc.text(label, x + 5, kpiY + 7);
        doc.setFontSize(11); doc.setFont("helvetica", "bold");
        doc.text(value, x + 5, kpiY + 15);
        doc.setFont("helvetica", "normal");
      };
      drawKpiBox(14, "REVENUS", `${fmt(totalRevenus)} FCFA`, 63, 156, 122);
      drawKpiBox(14 + kpiW + 8, "DÉPENSES", `${fmt(totalDepenses)} FCFA`, 193, 84, 63);
      drawKpiBox(14 + (kpiW + 8) * 2, "SOLDE", `${fmt(solde)} FCFA`, solde >= 0 ? 63 : 193, solde >= 0 ? 156 : 84, solde >= 0 ? 122 : 63);

      // Graphique en barres dessiné à la main (revenus vs dépenses par mois, jusqu'à 12 mois)
      let chartBottom = 68;
      const monthsChart = byMonth.slice(-12);
      if (monthsChart.length > 1) {
        const chartX = 14, chartY = 70, chartW = pageWidth - 28, chartH = 42;
        const maxVal = Math.max(1, ...monthsChart.map((m) => Math.max(m.revenus, m.depenses)));
        const bw = chartW / monthsChart.length;
        doc.setDrawColor(220, 220, 220);
        doc.line(chartX, chartY + chartH, chartX + chartW, chartY + chartH);
        monthsChart.forEach((m, i) => {
          const slot = chartX + i * bw;
          const hRev = (m.revenus / maxVal) * chartH;
          const hDep = (m.depenses / maxVal) * chartH;
          doc.setFillColor(63, 156, 122);
          doc.rect(slot + bw * 0.15, chartY + chartH - hRev, bw * 0.3, hRev, "F");
          doc.setFillColor(193, 84, 63);
          doc.rect(slot + bw * 0.5, chartY + chartH - hDep, bw * 0.3, hDep, "F");
          doc.setFontSize(6); doc.setTextColor(120, 120, 120);
          doc.text(monthLabel(m.key), slot + bw * 0.5, chartY + chartH + 5, { align: "center", maxWidth: bw });
        });
        doc.setFontSize(7); doc.setTextColor(63, 156, 122); doc.text("■ Revenus", chartX, chartY - 3);
        doc.setTextColor(193, 84, 63); doc.text("■ Dépenses", chartX + 28, chartY - 3);
        chartBottom = chartY + chartH + 10;
      }

      // Tableau des transactions, groupé avec sous-total à chaque changement de mois
      const sortedTx = filtered.slice().sort((a, b) => a.date.localeCompare(b.date));
      const rows: any[] = [];
      let curMonth: string | null = null, monthRev = 0, monthDep = 0;
      const pushSubtotal = () => {
        if (curMonth === null) return;
        rows.push([{ content: `▸ Sous-total ${monthLabel(curMonth)}`, colSpan: 4 }, { content: `Rev: ${fmt(monthRev)} / Dép: ${fmt(monthDep)}`, styles: { halign: "right" } }]);
      };
      sortedTx.forEach((t) => {
        if (curMonth !== null && t.month !== curMonth) { pushSubtotal(); monthRev = 0; monthDep = 0; }
        curMonth = t.month;
        if (t.type === "Revenu") monthRev += t.amount; else monthDep += t.amount;
        rows.push([dateLabelFull(t.date), t.time || "—", t.category + (t.subcategory ? ` · ${t.subcategory}` : ""), t.type, fmt(t.amount)]);
      });
      pushSubtotal();

      autoTable(doc, {
        startY: chartBottom,
        head: [["Date", "Heure", "Catégorie", "Type", "Montant (FCFA)"]],
        body: rows,
        styles: { fontSize: 8, cellPadding: 2.5 },
        headStyles: { fillColor: [26, 43, 76], textColor: 255, fontStyle: "bold" },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        columnStyles: { 4: { halign: "right" } },
        margin: { left: 14, right: 14 },
        didParseCell: (data: any) => {
          const first = data.row.raw?.[0];
          if (first && typeof first === "object" && first.content && String(first.content).startsWith("▸")) {
            data.cell.styles.fillColor = [27, 38, 32];
            data.cell.styles.textColor = [201, 162, 39];
            data.cell.styles.fontStyle = "bold";
          }
        },
      });

      doc.save(`grand-livre_${filters.from}_${filters.to}.pdf`);
      setPdfState("idle");
    } catch (e) {
      setPdfState("error");
      setTimeout(() => setPdfState("idle"), 4000);
    }
  };

  const tauxEpargne = totalRevenus > 0 ? (solde / totalRevenus) * 100 : 0;

  const filterDescriptions: string[] = [];
  filterDescriptions.push(`du ${monthLabel(filters.from)} au ${monthLabel(filters.to)}`);
  if (filters.type !== "Tous") filterDescriptions.push(`type : ${filters.type}`);
  if (filters.group !== "Tous") filterDescriptions.push(`groupe : ${filters.group}`);
  if (filters.scope !== "Tous") filterDescriptions.push(`portée : ${filters.scope}`);
  if (filters.category !== "Toutes") filterDescriptions.push(`catégorie : ${filters.category}`);
  if (filters.subcategory && filters.subcategory !== "Toutes") filterDescriptions.push(`sous-catégorie : ${filters.subcategory}`);
  if (filters.search) filterDescriptions.push(`recherche : "${filters.search}"`);
  const hasExtraFilters = filters.type !== "Tous" || filters.group !== "Tous" || filters.scope !== "Tous" || filters.category !== "Toutes" || !!filters.search;

  const applyPreset = (from: string, to: string) => {
    setFilters({ ...filters, from, to, type: "Tous", group: "Tous", scope: "Tous", category: "Toutes", subcategory: "Toutes", search: "" });
  };
  const presetThisMonth = () => { const k = dateToMonthKey(todayISO()); applyPreset(k, k); };
  const presetLast3Months = () => {
    const cur = monthSortKey(dateToMonthKey(todayISO()));
    const from = allMonths.find((m) => monthSortKey(m) >= cur - 2) || allMonths[0];
    applyPreset(from || dateToMonthKey(todayISO()), dateToMonthKey(todayISO()));
  };
  const presetThisYear = () => {
    const year = todayISO().slice(0, 4);
    const monthsThisYear = allMonths.filter((m) => m.startsWith(year));
    if (monthsThisYear.length) applyPreset(monthsThisYear[0], monthsThisYear[monthsThisYear.length - 1]);
  };
  const presetAll = () => { if (allMonths.length) applyPreset(allMonths[0], allMonths[allMonths.length - 1]); };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Ce que couvre ce rapport" subtitle="Basé exactement sur les filtres actifs en haut de l'écran">
        <div style={{ fontSize: 13, color: COLOR.ink, lineHeight: 1.7, marginBottom: 16 }}>
          Ce rapport et tous les exports ci-dessous portent sur <b>{filtered.length} transaction(s)</b> {filterDescriptions.join(" · ")}.
          {!hasExtraFilters && " Aucun filtre restrictif (type, groupe, catégorie) n'est actif au-delà de la période — l'export couvre toutes les données de cet intervalle."}
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={presetThisMonth} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 20, color: COLOR.inkMuted, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>Mois en cours</button>
          <button onClick={presetLast3Months} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 20, color: COLOR.inkMuted, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>3 derniers mois</button>
          <button onClick={presetThisYear} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 20, color: COLOR.inkMuted, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>Année en cours</button>
          <button onClick={presetAll} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 20, color: COLOR.inkMuted, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>Tout l'historique</button>
        </div>
      </Panel>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Revenus (sélection)" value={fmt(totalRevenus)} tone={COLOR.emeraldSoft} icon={TrendingUp} />
        <Kpi label="Dépenses (sélection)" value={fmt(totalDepenses)} tone={COLOR.claySoft} icon={TrendingDown} />
        <Kpi label="Solde" value={fmt(solde)} tone={solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} />
        <Kpi label="Taux d'épargne" value={tauxEpargne.toFixed(1)} suffix="%" tone={COLOR.gold} icon={Target} />
      </div>

      <Panel title="Résumé automatique" subtitle="Généré à partir de la période filtrée" right={<Sparkles size={16} color={COLOR.goldSoft} />}>
        <p style={{ fontSize: 13.5, lineHeight: 1.7, color: COLOR.ink, margin: 0 }}>{summary}</p>
      </Panel>

      <Panel title="Que contient chaque format ?" subtitle="Choisis selon ce que tu veux en faire">
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {[
            { icon: FileSpreadsheet, color: COLOR.emeraldSoft, name: "Excel (.xlsx)", desc: "4 feuilles mises en couleur : Résumé, Transactions (avec sous-total automatique à chaque changement de mois), Par mois, Par catégorie (regroupée par groupe avec sous-totaux). Idéal pour retravailler les chiffres." },
            { icon: FileText, color: COLOR.goldSoft, name: "PDF", desc: "Document coloré prêt à partager : bandeau d'en-tête, cartes KPI, un graphique en barres Revenus/Dépenses par mois dessiné directement, puis le tableau détaillé avec sous-totaux mensuels en surbrillance dorée." },
            { icon: Download, color: COLOR.inkMuted, name: "CSV", desc: "Liste brute des transactions, une ligne par opération. Le format le plus simple à réimporter dans un autre outil ou tableur." },
            { icon: Printer, color: COLOR.inkMuted, name: "Imprimer", desc: "Utilise la fenêtre d'impression de ton navigateur — pratique pour un aperçu rapide papier ou un export PDF alternatif." },
          ].map((f) => (
            <div key={f.name} style={{ display: "flex", gap: 12, padding: "10px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <f.icon size={16} color={f.color} style={{ flexShrink: 0, marginTop: 2 }} />
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: COLOR.ink, marginBottom: 2 }}>{f.name}</div>
                <div style={{ fontSize: 12, color: COLOR.inkMuted, lineHeight: 1.5 }}>{f.desc}</div>
              </div>
            </div>
          ))}
        </div>
      </Panel>

      <Panel title="Exporter le rapport filtré" subtitle={`${filtered.length} transaction(s) dans la sélection actuelle`}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <button onClick={exportExcel} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(63,156,122,0.14)", border: `1px solid ${COLOR.emerald}`, borderRadius: 8, color: COLOR.emeraldSoft, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
            <FileSpreadsheet size={15} /> Exporter en Excel (.xlsx)
          </button>
          <button onClick={exportPDF} disabled={pdfState === "loading"} style={{ display: "flex", alignItems: "center", gap: 8, background: pdfState === "error" ? "rgba(193,84,63,0.14)" : "rgba(201,162,39,0.14)", border: `1px solid ${pdfState === "error" ? COLOR.clay : COLOR.gold}`, borderRadius: 8, color: pdfState === "error" ? COLOR.claySoft : COLOR.goldSoft, padding: "10px 18px", fontSize: 13, cursor: pdfState === "loading" ? "default" : "pointer" }}>
            {pdfState === "loading" ? <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> : <FileText size={15} />}
            {pdfState === "loading" ? "Génération…" : pdfState === "error" ? "Échec — réessayer" : "Télécharger en PDF"}
          </button>
          <button onClick={exportCSV} style={{ display: "flex", alignItems: "center", gap: 8, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, color: COLOR.inkMuted, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
            <Download size={15} /> Exporter en CSV
          </button>
          <button onClick={() => window.print()} style={{ display: "flex", alignItems: "center", gap: 8, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, color: COLOR.inkMuted, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
            <Printer size={15} /> Imprimer
          </button>
        </div>
        {pdfState === "error" && (
          <div style={{ marginTop: 12, fontSize: 11.5, color: COLOR.claySoft, display: "flex", alignItems: "center", gap: 6 }}>
            <AlertTriangle size={12} /> Le générateur PDF n'a pas pu se charger (aperçu Claude sans accès à cette librairie). Fonctionne normalement sur le site déployé.
          </div>
        )}
      </Panel>
      <Panel title="Aperçu du rapport imprimable" subtitle="Ce contenu apparaît lors de l'impression">
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>
          <div>Période : {monthLabel(filters.from)} — {monthLabel(filters.to)}</div>
          <div style={{ color: COLOR.emeraldSoft }}>Revenus : {fmt(totalRevenus)} FCFA</div>
          <div style={{ color: COLOR.claySoft }}>Dépenses : {fmt(totalDepenses)} FCFA</div>
          <div style={{ color: solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>Solde : {fmt(solde)} FCFA</div>
        </div>
      </Panel>
    </div>
  );
}
// ============================================================
// SAISIE QUOTIDIENNE (entrée rapide, jour par jour)
// ============================================================
function SaisieQuotidienneTab({ transactions, setTransactions, allCategories, categoryGroups, accounts }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; allCategories: string[]; categoryGroups: Record<string, Group>; accounts: Account[];
}) {
  const [quickDate, setQuickDate] = useState(todayISO());
  const [quickTime, setQuickTime] = useState(nowTime());
  const [quickCategory, setQuickCategory] = useState(() => defaultQuickCategory(transactions, "Dépense"));
  const [quickSubcategory, setQuickSubcategory] = useState("");
  const [quickType, setQuickType] = useState<TxType>("Dépense");
  const [quickAmount, setQuickAmount] = useState<number | "">("");
  const [quickAccount, setQuickAccount] = useState(() => defaultQuickAccount(accounts));
  const [justAdded, setJustAdded] = useState(false);
  const [catPickerOpen, setCatPickerOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  const today = todayISO();
  const currentMonthKey = dateToMonthKey(today);
  const weekAgo = (() => { const d = new Date(); d.setDate(d.getDate() - 6); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; })();

  const withGroup = transactions.map((t) => ({ ...t, group: t.type === "Revenu" ? "Revenu" : (categoryGroups[t.category] || "Non classifié") }));

  const sumFor = (pred: (t: any) => boolean) => {
    const arr = withGroup.filter(pred);
    const rev = arr.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
    const dep = arr.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
    return { rev, dep, solde: rev - dep };
  };

  const todayTotals = sumFor((t) => t.date === today);
  const weekTotals = sumFor((t) => t.date >= weekAgo && t.date <= today);
  const monthTotals = sumFor((t) => dateToMonthKey(t.date) === currentMonthKey);

  const quickDateEntries = withGroup.filter((t) => t.date === quickDate).sort((a, b) => b.id.localeCompare(a.id));

  const resetForm = () => {
    setQuickAmount(""); setQuickTime(nowTime()); setEditingId(null);
  };

  const submit = () => {
    if (!quickCategory || !quickAmount || Number(quickAmount) <= 0) return;
    if (editingId) {
      setTransactions(transactions.map((t) => t.id === editingId ? {
        ...t, date: quickDate, time: quickTime, category: quickCategory, subcategory: quickSubcategory || undefined,
        type: quickType, amount: Number(quickAmount), account: quickAccount || undefined,
      } : t));
    } else {
      setTransactions([...transactions, { id: uid(), date: quickDate, time: quickTime, category: quickCategory, subcategory: quickSubcategory || undefined, type: quickType, amount: Number(quickAmount), account: quickAccount || undefined }]);
    }
    resetForm();
    setJustAdded(true);
    setTimeout(() => setJustAdded(false), 1200);
  };

  const editEntry = (t: Transaction) => {
    setEditingId(t.id);
    setQuickDate(t.date);
    setQuickTime(t.time || nowTime());
    setQuickType(t.type);
    setQuickCategory(t.category);
    setQuickSubcategory(t.subcategory || "");
    setQuickAmount(t.amount);
    setQuickAccount(t.account || defaultQuickAccount(accounts));
  };

  const remove = (id: string) => setTransactions(transactions.filter((t) => t.id !== id));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Aujourd'hui — solde" value={fmt(todayTotals.solde)} tone={todayTotals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Clock} />
        <Kpi label="7 derniers jours — solde" value={fmt(weekTotals.solde)} tone={weekTotals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={CalendarDays} />
        <Kpi label="Mois en cours — solde" value={fmt(monthTotals.solde)} tone={monthTotals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={CalendarRange} />
      </div>

      <Panel title="Saisie rapide" subtitle="Ajoutez vos dépenses et revenus au fil de la journée — comptabilisés instantanément">
        {(() => {
          const typeColor = quickType === "Revenu" ? COLOR.emerald : COLOR.clay;
          const fieldLabel: React.CSSProperties = { fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 };
          const nakedSelect: React.CSSProperties = {
            background: "transparent", border: "none", color: COLOR.ink, fontSize: 16, fontWeight: 600,
            fontFamily: "'Fraunces', serif", padding: 0, cursor: "pointer", width: "100%", appearance: "none", WebkitAppearance: "none",
          };
          return (
            <div style={{ background: `linear-gradient(180deg, ${COLOR.surfaceRaised} 0%, ${COLOR.surface} 70%)`, border: `1px solid ${COLOR.hairline}`, borderRadius: 16, overflow: "hidden" }}>
              {/* Type + Date */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 20px 6px 20px", flexWrap: "wrap", gap: 10 }}>
                <div style={{ display: "flex", gap: 8 }}>
                  <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "8px 16px" }}>
                    <input type="date" value={quickDate} onChange={(e) => setQuickDate(e.target.value)}
                      style={{ background: "transparent", border: "none", color: COLOR.ink, fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" }} />
                  </div>
                  <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "8px 16px" }}>
                    <input type="time" value={quickTime} onChange={(e) => setQuickTime(e.target.value)}
                      style={{ background: "transparent", border: "none", color: COLOR.ink, fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" }} />
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10, background: COLOR.surface, borderRadius: 24, padding: 5, border: `1px solid ${COLOR.hairline}` }}>
                  <button onClick={() => { setQuickType("Dépense"); setQuickSubcategory(""); setQuickCategory(defaultQuickCategory(transactions, "Dépense")); }} title="Dépense" style={{
                    width: 34, height: 34, borderRadius: "50%", border: "none", cursor: "pointer",
                    background: quickType === "Dépense" ? COLOR.clay : "transparent", color: quickType === "Dépense" ? COLOR.bg : COLOR.claySoft,
                    display: "flex", alignItems: "center", justifyContent: "center",
                  }}><Minus size={16} strokeWidth={2.5} /></button>
                  <button onClick={() => { setQuickType("Revenu"); setQuickSubcategory(""); setQuickCategory(defaultQuickCategory(transactions, "Revenu")); }} title="Revenu" style={{
                    width: 34, height: 34, borderRadius: "50%", border: "none", cursor: "pointer",
                    background: quickType === "Revenu" ? COLOR.emerald : "transparent", color: quickType === "Revenu" ? COLOR.bg : COLOR.emeraldSoft,
                    display: "flex", alignItems: "center", justifyContent: "center",
                  }}><Plus size={16} strokeWidth={2.5} /></button>
                </div>
              </div>

              {/* Montant */}
              <div style={{ position: "relative", display: "flex", flexDirection: "column", alignItems: "center", padding: "18px 24px 8px 24px" }}>
                <div style={{ position: "absolute", fontSize: 56, fontWeight: 700, color: typeColor, opacity: 0.07, fontFamily: "'Fraunces', serif", pointerEvents: "none", userSelect: "none", top: 10 }}>FCFA</div>
                <input type="number" inputMode="numeric" value={quickAmount} placeholder="0"
                  onChange={(e) => setQuickAmount(e.target.value === "" ? "" : Number(e.target.value))}
                  onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
                  style={{ position: "relative", background: "transparent", border: "none", outline: "none", color: COLOR.ink, fontSize: 42, fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace", textAlign: "center", width: "100%", maxWidth: 260 }} />
              </div>

              {/* Compte / Catégorie */}
              <div style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "16px 20px" }}>
                <div style={{ display: "flex", gap: 24 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={fieldLabel}>Compte</div>
                    <select value={quickAccount} onChange={(e) => setQuickAccount(e.target.value)} style={nakedSelect}>
                      {!accounts.length && <option value="">Aucun compte créé</option>}
                      {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: 1, minWidth: 0, textAlign: "right" }}>
                    <div style={{ ...fieldLabel, textAlign: "right" }}>Catégorie</div>
                    <button onClick={() => setCatPickerOpen(true)} style={{ ...nakedSelect, textAlign: "right", cursor: "pointer", display: "block" }}>
                      {quickCategory || "Choisir…"}
                    </button>
                    {quickSubcategory && (
                      <div style={{ textAlign: "right", fontSize: 13, color: COLOR.inkMuted, marginTop: 4 }}>{quickSubcategory}</div>
                    )}
                  </div>
                </div>
                <CategoryPickerSheet open={catPickerOpen} onClose={() => setCatPickerOpen(false)} transactions={transactions} type={quickType}
                  value={quickCategory} subvalue={quickSubcategory} onSelect={(c, s) => { setQuickCategory(c); setQuickSubcategory(s); }} />

                {editingId && (
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 14, padding: "8px 12px", background: "rgba(201,162,39,0.08)", border: `1px solid ${COLOR.gold}`, borderRadius: 8 }}>
                    <span style={{ fontSize: 12, color: COLOR.goldSoft }}>Modification d'une entrée existante</span>
                    <button onClick={resetForm} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, fontSize: 11.5, cursor: "pointer", textDecoration: "underline" }}>Annuler</button>
                  </div>
                )}

                <button onClick={submit} disabled={!quickAmount || Number(quickAmount) <= 0} style={{
                  width: "100%", marginTop: 12, padding: "14px 0", borderRadius: 12, border: "none",
                  background: justAdded ? COLOR.emerald : (!quickAmount || Number(quickAmount) <= 0) ? COLOR.hairline : COLOR.gold,
                  color: justAdded ? COLOR.bg : (!quickAmount || Number(quickAmount) <= 0) ? COLOR.inkMuted : COLOR.bg,
                  fontSize: 14.5, fontWeight: 700, cursor: (!quickAmount || Number(quickAmount) <= 0) ? "default" : "pointer",
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "background 0.15s",
                }}>
                  {justAdded ? <Check size={17} /> : null} {justAdded ? (editingId ? "Mis à jour" : "Ajouté") : (editingId ? "Mettre à jour" : "Sauvegarder")}
                </button>
              </div>
            </div>
          );
        })()}
      </Panel>

      <Panel title={`Entrées du ${dateLabelFull(quickDate)}`} subtitle={`Revenus ${fmt(sumFor((t) => t.date === quickDate).rev)} · Dépenses ${fmt(sumFor((t) => t.date === quickDate).dep)} · Solde ${fmt(sumFor((t) => t.date === quickDate).solde)}`}>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {quickDateEntries.map((t) => (
            <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: editingId === t.id ? "rgba(201,162,39,0.08)" : COLOR.surfaceRaised, border: editingId === t.id ? `1px solid ${COLOR.gold}` : "1px solid transparent", borderRadius: 6 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                {t.time && <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: COLOR.inkMuted, width: 38 }}>{t.time}</span>}
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: groupColor[t.group] || COLOR.inkMuted, display: "inline-block" }} />
                <span style={{ fontSize: 12.5 }}>{t.category}{t.subcategory && ` · ${t.subcategory}`}</span>
                <span style={{ fontSize: 11, color: COLOR.inkMuted }}>{t.type}</span>
                {t.account && <span style={{ fontSize: 10.5, color: COLOR.slateBlueSoft }}>{t.account}</span>}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(t.amount)}</span>
                <button onClick={() => editEntry(t)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={13} /></button>
                <button onClick={() => setConfirmDeleteId(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
              </div>
            </div>
          ))}
          {!quickDateEntries.length && <EmptyState text="Aucune entrée pour cette date." />}
        </div>
      </Panel>
      <ConfirmDialog
        open={!!confirmDeleteId}
        title="Supprimer cette transaction ?"
        message="Cette action est définitive. Le montant ne sera plus comptabilisé nulle part dans l'app."
        onConfirm={() => { if (confirmDeleteId) { remove(confirmDeleteId); if (editingId === confirmDeleteId) resetForm(); } setConfirmDeleteId(null); }}
        onCancel={() => setConfirmDeleteId(null)}
      />
    </div>
  );
}

// ============================================================
// RAPPORT JOURNALIER
// ============================================================
function JournalierTab({ filtered }: { filtered: any[] }) {
  const rows = useMemo(() => {
    const m: Record<string, { revenus: number; depenses: number }> = {};
    filtered.forEach((t) => {
      if (!m[t.date]) m[t.date] = { revenus: 0, depenses: 0 };
      if (t.type === "Revenu") m[t.date].revenus += t.amount; else m[t.date].depenses += t.amount;
    });
    return Object.keys(m).sort().map((d) => ({ date: d, label: dateLabelShort(d), revenus: m[d].revenus, depenses: m[d].depenses, solde: m[d].revenus - m[d].depenses }));
  }, [filtered]);

  const totals = rows.reduce((a, r) => ({ revenus: a.revenus + r.revenus, depenses: a.depenses + r.depenses, solde: a.solde + r.solde }), { revenus: 0, depenses: 0, solde: 0 });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Solde net par jour" subtitle={`${rows.length} jour(s) avec activité dans la période filtrée`}
        explain="Chaque point représente le solde (revenus−dépenses) d'un seul jour, pas un cumul. Les pointes vers le haut sont des jours de rentrée d'argent (salaire, loyer perçu…) ; les creux vers le bas sont des jours de grosse dépense. La ligne de zéro sépare les jours excédentaires des jours déficitaires.">
        {rows.length ? (
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={rows} margin={{ left: 0, right: 10, top: 10 }}>
              <CartesianGrid stroke={COLOR.hairline} vertical={false} />
              <XAxis dataKey="label" tick={{ fill: COLOR.inkMuted, fontSize: 9.5 }} interval={Math.max(0, Math.floor(rows.length / 20))} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
              <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
              <Tooltip content={<CustomTooltip />} />
              <ReferenceLine y={0} stroke={COLOR.hairline} />
              <Line type="monotone" dataKey="solde" name="Solde du jour" stroke={COLOR.goldSoft} strokeWidth={2} dot={rows.length < 60} />
            </LineChart>
          </ResponsiveContainer>
        ) : <EmptyState />}
        {rows.length > 60 && (
          <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 10, display: "flex", alignItems: "center", gap: 6 }}>
            <Info size={12} /> Période large — restreignez le filtre "Du mois / Au mois" pour une lecture jour par jour plus lisible.
          </div>
        )}
      </PanelWithHelp>
      <Panel title="Tableau journalier" subtitle={`${rows.length} jour(s)`}>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'IBM Plex Mono', monospace" }}>
            <thead>
              <tr>
                {["Date", "Revenus", "Dépenses", "Solde"].map((h, i) => (
                  <th key={h} style={{ textAlign: i === 0 ? "left" : "right", padding: "8px 10px", fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", borderBottom: `1px solid ${COLOR.hairline}`, position: "sticky", top: 0, background: COLOR.surface }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.slice().reverse().map((r) => (
                <tr key={r.date}>
                  <td style={{ padding: "8px 10px", fontSize: 12, borderBottom: `1px solid ${COLOR.hairline}`, fontFamily: "'Inter', sans-serif" }}>{r.label}</td>
                  <td style={{ padding: "8px 10px", fontSize: 12, textAlign: "right", color: COLOR.emeraldSoft, borderBottom: `1px solid ${COLOR.hairline}` }}>{r.revenus ? fmt(r.revenus) : "—"}</td>
                  <td style={{ padding: "8px 10px", fontSize: 12, textAlign: "right", color: COLOR.claySoft, borderBottom: `1px solid ${COLOR.hairline}` }}>{r.depenses ? fmt(r.depenses) : "—"}</td>
                  <td style={{ padding: "8px 10px", fontSize: 12, textAlign: "right", color: r.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft, borderBottom: `1px solid ${COLOR.hairline}` }}>{fmt(r.solde)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr>
                <td style={{ padding: "9px 10px", fontSize: 12, fontWeight: 600, fontFamily: "'Inter', sans-serif" }}>Total</td>
                <td style={{ padding: "9px 10px", fontSize: 12, textAlign: "right", fontWeight: 600, color: COLOR.emeraldSoft }}>{fmt(totals.revenus)}</td>
                <td style={{ padding: "9px 10px", fontSize: 12, textAlign: "right", fontWeight: 600, color: COLOR.claySoft }}>{fmt(totals.depenses)}</td>
                <td style={{ padding: "9px 10px", fontSize: 12, textAlign: "right", fontWeight: 600, color: totals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(totals.solde)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </Panel>
    </div>
  );
}

// ============================================================
// SAISIE RAPIDE FLOTTANTE — accessible depuis n'importe quel onglet
// ============================================================
function QuickAddFAB({ transactions, setTransactions, accounts, categoryGroups, isMobile }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; accounts: Account[]; categoryGroups: Record<string, Group>; isMobile: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [date, setDate] = useState(todayISO());
  const [time, setTime] = useState(nowTime());
  const [type, setType] = useState<TxType>("Dépense");
  const [category, setCategory] = useState(() => defaultQuickCategory(transactions, "Dépense"));
  const [subcategory, setSubcategory] = useState("");
  const [amount, setAmount] = useState<number | "">("");
  const [account, setAccount] = useState(() => defaultQuickAccount(accounts));
  const [note, setNote] = useState("");
  const [justAdded, setJustAdded] = useState(false);
  const [catPickerOpen, setCatPickerOpen] = useState(false);

  useEffect(() => {
    if (open) {
      setTime(nowTime());
      const prevHtmlOverflow = document.documentElement.style.overflow;
      const prevBodyOverflow = document.body.style.overflow;
      document.documentElement.style.overflow = "hidden";
      document.body.style.overflow = "hidden";
      return () => {
        document.documentElement.style.overflow = prevHtmlOverflow;
        document.body.style.overflow = prevBodyOverflow;
      };
    }
  }, [open]);

  const submit = () => {
    if (!category || !amount || Number(amount) <= 0) return;
    setTransactions([...transactions, {
      id: uid(), date, time, category, subcategory: subcategory || undefined, type, amount: Number(amount),
      account: account || undefined, note: note || undefined,
    }]);
    setAmount(""); setNote("");
    setJustAdded(true);
    setTimeout(() => setJustAdded(false), 1000);
  };

  const changeType = (ty: TxType) => {
    setType(ty);
    setSubcategory("");
    setCategory(defaultQuickCategory(transactions, ty));
  };

  const typeColor = type === "Revenu" ? COLOR.emerald : COLOR.clay;

  const fieldLabel: React.CSSProperties = { fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 };
  const nakedSelect: React.CSSProperties = {
    background: "transparent", border: "none", color: COLOR.ink, fontSize: 16, fontWeight: 600,
    fontFamily: "'Fraunces', serif", padding: 0, cursor: "pointer", width: "100%", appearance: "none",
    WebkitAppearance: "none",
  };

  return (
    <>
      <button onClick={() => setOpen(true)} className="gl-noprint" style={{
        position: "fixed", bottom: isMobile ? 82 : 28, right: 24, zIndex: 155, width: 56, height: 56, borderRadius: "50%",
        background: COLOR.gold, border: "none", color: COLOR.bg, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 6px 20px rgba(201,162,39,0.4)",
      }}>
        <Plus size={26} />
      </button>

      {open && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 300, background: COLOR.bg,
          display: "flex", justifyContent: "center", alignItems: isMobile ? "stretch" : "center",
        }}>
          <div style={{
            width: "100%", maxWidth: isMobile ? "100%" : 440, height: isMobile ? "100%" : "min(720px, 92vh)",
            display: "flex", flexDirection: "column", background: `linear-gradient(180deg, ${COLOR.surfaceRaised} 0%, ${COLOR.bg} 55%)`,
            borderRadius: isMobile ? 0 : 20, border: isMobile ? "none" : `1px solid ${COLOR.hairline}`,
            overflowY: "auto", overflowX: "hidden", WebkitOverflowScrolling: "touch", overscrollBehavior: "contain",
          }}>
            {/* Header : fermer + sélecteur de type */}
            <div className="gl-safe-top" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "20px 20px 8px 20px" }}>
              <button onClick={() => setOpen(false)} style={{
                width: 40, height: 40, borderRadius: "50%", background: COLOR.surface, border: `1px solid ${COLOR.hairline}`,
                color: COLOR.inkMuted, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
              }}>
                <X size={18} />
              </button>
              <div style={{ display: "flex", gap: 10, background: COLOR.surface, borderRadius: 24, padding: 5, border: `1px solid ${COLOR.hairline}` }}>
                <button onClick={() => changeType("Dépense")} title="Dépense" style={{
                  width: 38, height: 38, borderRadius: "50%", border: "none", cursor: "pointer",
                  background: type === "Dépense" ? COLOR.clay : "transparent", color: type === "Dépense" ? COLOR.bg : COLOR.claySoft,
                  display: "flex", alignItems: "center", justifyContent: "center", transition: "background 0.15s",
                }}>
                  <Minus size={18} strokeWidth={2.5} />
                </button>
                <button onClick={() => changeType("Revenu")} title="Revenu" style={{
                  width: 38, height: 38, borderRadius: "50%", border: "none", cursor: "pointer",
                  background: type === "Revenu" ? COLOR.emerald : "transparent", color: type === "Revenu" ? COLOR.bg : COLOR.emeraldSoft,
                  display: "flex", alignItems: "center", justifyContent: "center", transition: "background 0.15s",
                }}>
                  <Plus size={18} strokeWidth={2.5} />
                </button>
              </div>
              <div style={{ width: 40 }} />
            </div>

            {/* Date + heure */}
            <div style={{ display: "flex", justifyContent: "center", gap: 10, marginTop: 8 }}>
              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "8px 18px" }}>
                <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
                  style={{ background: "transparent", border: "none", color: COLOR.ink, fontSize: 14, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" }} />
              </div>
              <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "8px 18px" }}>
                <input type="time" value={time} onChange={(e) => setTime(e.target.value)}
                  style={{ background: "transparent", border: "none", color: COLOR.ink, fontSize: 14, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" }} />
              </div>
            </div>

            {/* Montant — zone centrale */}
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", position: "relative", padding: "10px 24px", minHeight: 0 }}>
              <div style={{
                position: "absolute", fontSize: 72, fontWeight: 700, color: typeColor, opacity: 0.07,
                fontFamily: "'Fraunces', serif", pointerEvents: "none", userSelect: "none",
              }}>FCFA</div>
              <input
                type="number" inputMode="numeric" value={amount} placeholder="0" autoFocus
                onChange={(e) => setAmount(e.target.value === "" ? "" : Number(e.target.value))}
                onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
                style={{
                  position: "relative", background: "transparent", border: "none", outline: "none",
                  color: COLOR.ink, fontSize: 52, fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace",
                  textAlign: "center", width: "100%", maxWidth: 280,
                }}
              />
              <input
                value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ajouter une note"
                style={{
                  position: "relative", background: "transparent", border: "none", outline: "none", marginTop: 14,
                  color: COLOR.inkMuted, fontSize: 14, fontFamily: "'Inter', sans-serif", textAlign: "center", width: "100%",
                }}
              />
            </div>

            {/* Bas : compte / catégorie / bénéficiaire / enregistrer */}
            <div style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "18px 20px", background: COLOR.surface }} className="gl-safe-bottom">
              <div style={{ display: "flex", gap: 20 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={fieldLabel}>Compte</div>
                  <select value={account} onChange={(e) => setAccount(e.target.value)} style={nakedSelect}>
                    {!accounts.length && <option value="">Aucun</option>}
                    {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                  </select>
                </div>
                <div style={{ flex: 1, minWidth: 0, textAlign: "right" }}>
                  <div style={{ ...fieldLabel, textAlign: "right" }}>Catégorie</div>
                  <button onClick={() => setCatPickerOpen(true)} style={{ ...nakedSelect, textAlign: "right", cursor: "pointer", display: "block" }}>
                    {category || "Choisir…"}
                  </button>
                  {subcategory && (
                    <div style={{ textAlign: "right", fontSize: 13, color: COLOR.inkMuted, marginTop: 4 }}>{subcategory}</div>
                  )}
                </div>
              </div>
              <CategoryPickerSheet open={catPickerOpen} onClose={() => setCatPickerOpen(false)} transactions={transactions} type={type}
                value={category} subvalue={subcategory} onSelect={(c, s) => { setCategory(c); setSubcategory(s); }} />

              <button onClick={submit} disabled={!amount || Number(amount) <= 0} style={{
                width: "100%", marginTop: 16, padding: "15px 0", borderRadius: 14, border: "none",
                background: justAdded ? COLOR.emerald : (!amount || Number(amount) <= 0) ? COLOR.hairline : COLOR.gold,
                color: justAdded ? COLOR.bg : (!amount || Number(amount) <= 0) ? COLOR.inkMuted : COLOR.bg,
                fontSize: 15.5, fontWeight: 700, cursor: (!amount || Number(amount) <= 0) ? "default" : "pointer",
                display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "background 0.15s",
              }}>
                {justAdded ? <Check size={18} /> : null}
                {justAdded ? "Ajouté" : "Sauvegarder"}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ============================================================ END OF PART 6 — App component follows in part 7
// ============================================================
// MAIN APP
// ============================================================
type Tab = "saisie" | "apercu" | "flux" | "comparatif" | "comparateur" | "topcategories" | "categoryoverview" | "mensuel" | "journalier" | "categories" | "groupes" | "enveloppes" | "budgets" | "simulateur" | "objectif" | "business" | "creances" | "comptes" | "payees" | "recurrences" | "journal" | "export" | "sauvegarde";

const NAV: { section: string; items: { id: Tab; label: string; icon: any }[] }[] = [
  { section: "Saisie rapide", items: [
    { id: "saisie", label: "Saisie du jour", icon: Clock },
  ]},
  { section: "Tableau de bord", items: [
    { id: "apercu", label: "Aperçu", icon: LayoutDashboard },
    { id: "flux", label: "Flux & Calendrier", icon: Workflow },
    { id: "comparatif", label: "Comparatif annuel", icon: BarChart3 },
    { id: "comparateur", label: "Comparateur", icon: GitCompare },
    { id: "topcategories", label: "Principales catégories", icon: PieChartIcon },
    { id: "categoryoverview", label: "Aperçu de catégorie", icon: Activity },
  ]},
  { section: "Budget", items: [
    { id: "mensuel", label: "Rapport mensuel", icon: CalendarRange },
    { id: "journalier", label: "Rapport journalier", icon: CalendarDays },
    { id: "categories", label: "Catégories", icon: PiggyBank },
    { id: "groupes", label: "Groupes", icon: Layers },
    { id: "enveloppes", label: "Enveloppes", icon: Mail },
    { id: "budgets", label: "Budgets par catégorie", icon: ClipboardList },
  ]},
  { section: "Outils", items: [
    { id: "simulateur", label: "Simulateur", icon: SlidersHorizontal },
    { id: "objectif", label: "Objectifs & Projection", icon: Gauge },
    { id: "business", label: "Business / Personnel", icon: Briefcase },
    { id: "creances", label: "Créances", icon: HandCoins },
    { id: "comptes", label: "Comptes", icon: Wallet },
    { id: "payees", label: "Bénéficiaires", icon: Users },
    { id: "recurrences", label: "Récurrences", icon: Repeat },
  ]},
  { section: "Données", items: [
    { id: "journal", label: "Journal", icon: BookOpen },
    { id: "export", label: "Rapports & Export", icon: Download },
    { id: "sauvegarde", label: "Sauvegarde", icon: UploadCloud },
  ]},
];

export default function GrandLivre() {
  const [transactions, setTransactions, txLoaded] = usePersistentState<Transaction[]>("gl-transactions", seedTransactions);
  const [categoryGroups, setCategoryGroups, groupsLoaded] = usePersistentState<Record<string, Group>>("gl-category-groups", defaultCategoryGroups);
  const [categoryScope, setCategoryScope, scopeLoaded] = usePersistentState<Record<string, Scope>>("gl-category-scope", defaultCategoryScope);
  const [rules, setRules, rulesLoaded] = usePersistentState<CategorizationRule[]>("gl-rules", defaultRules);
  const [loans, setLoans, loansLoaded] = usePersistentState<Loan[]>("gl-loans", seedLoans);
  const [envelopeCap, setEnvelopeCap, capLoaded] = usePersistentState<number>("gl-envelope-cap", 600000);
  const [accounts, setAccounts, accountsLoaded] = usePersistentState<Account[]>("gl-accounts", seedAccounts);
  const [budgets, setBudgets, budgetsLoaded] = usePersistentState<CategoryBudget[]>("gl-budgets", seedBudgets);
  const [goals, setGoals, goalsLoaded] = usePersistentState<Goal[]>("gl-goals", seedGoals);
  const [recurring, setRecurring, recurringLoaded] = usePersistentState<RecurringTemplate[]>("gl-recurring", seedRecurring);
  const [tab, setTab] = useState<Tab>("saisie");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const isMobile = useIsMobile();
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  useEffect(() => {
    if (isMobile && mobileMenuOpen) {
      const prevHtmlOverflow = document.documentElement.style.overflow;
      const prevBodyOverflow = document.body.style.overflow;
      document.documentElement.style.overflow = "hidden";
      document.body.style.overflow = "hidden";
      return () => {
        document.documentElement.style.overflow = prevHtmlOverflow;
        document.body.style.overflow = prevBodyOverflow;
      };
    }
  }, [isMobile, mobileMenuOpen]);
  const [filtersOpen, setFiltersOpen] = useState(!isMobile);
  const [syncCode, setSyncCode, syncCodeLoaded] = usePersistentState<string>("gl-sync-code", "");

  // Récupère le code de synchronisation depuis l'URL (?sync=code), si présent — utile si le
  // localStorage a été effacé par iOS. Marquer cette page en favori avec ce paramètre rend
  // la reconnexion automatique, sans jamais retaper le code.
  useEffect(() => {
    if (!syncCodeLoaded) return;
    try {
      const params = new URLSearchParams(window.location.search);
      const urlCode = params.get("sync");
      if (urlCode && urlCode.trim() && urlCode.trim().toLowerCase() !== syncCode) {
        setSyncCode(urlCode.trim().toLowerCase());
      }
    } catch {}
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [syncCodeLoaded]);
  const [syncStatus, setSyncStatus] = useState<"idle" | "syncing" | "synced" | "error" | "disabled">(SYNC_ENABLED ? "idle" : "disabled");
  const [lastSyncedAt, setLastSyncedAt] = useState<string | null>(null);
  const [realtimeConnected, setRealtimeConnected] = useState(false);
  const skipNextPush = useRef(false);
  const pushTimer = useRef<any>(null);

  const allMonths = useMemo(() => {
    const s = new Set(transactions.filter((t) => t).map((t) => dateToMonthKey(t.date)));
    return Array.from(s).sort((a, b) => monthSortKey(a) - monthSortKey(b));
  }, [transactions]);

  const allCategories = useMemo(() => {
    const s = new Set(transactions.map((t) => t.category));
    return Array.from(s).sort();
  }, [transactions]);

  const defaultFilters: Filters = {
    from: allMonths[0] || "2024_6", to: allMonths[allMonths.length - 1] || "2026_8",
    type: "Tous", group: "Tous", category: "Toutes", subcategory: "Toutes", search: "", scope: "Tous",
  };
  const [filters, setFilters] = useState<Filters>(defaultFilters);

  // Se re-synchronise sur la vraie période disponible dès que les données réelles sont chargées
  // (localStorage / Supabase) — évite que le filtre reste bloqué sur la période calculée au tout
  // premier rendu (avant que les données réelles ne soient prêtes), qui peut différer une fois
  // les données effectivement chargées.
  const allMonthsKey = allMonths.join(",");
  useEffect(() => {
    if (!allMonths.length) return;
    const rangeStillValid = allMonths.includes(filters.from) && allMonths.includes(filters.to);
    if (!rangeStillValid) {
      setFilters((f) => ({ ...f, from: allMonths[0], to: allMonths[allMonths.length - 1] }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allMonthsKey]);

  // applique les règles de catégorisation aux catégories inconnues
  const resolvedGroups = useMemo(() => {
    const out = { ...categoryGroups };
    allCategories.forEach((c) => {
      if (out[c]) return;
      const match = rules.find((r) => c.toLowerCase().includes(r.keyword.toLowerCase()));
      if (match) out[c] = match.group;
    });
    return out;
  }, [categoryGroups, allCategories, rules]);

  const txWithGroup = useMemo(() => transactions.filter((t) => t && typeof t.amount === "number").map((t) => ({
    ...t,
    date: t.date || todayISO(),
    month: dateToMonthKey(t.date),
    group: t.type === "Revenu" ? "Revenu" : (resolvedGroups[t.category] || "Non classifié"),
    scope: categoryScope[t.category] || "Personnel",
  })), [transactions, resolvedGroups, categoryScope]);

  const filtered = useMemo(() => {
    const fromKey = monthSortKey(filters.from);
    const toKey = monthSortKey(filters.to);
    return txWithGroup.filter((t) => {
      const mk = monthSortKey(t.month);
      if (mk < fromKey || mk > toKey) return false;
      if (filters.type !== "Tous" && t.type !== filters.type) return false;
      if (filters.group !== "Tous" && t.group !== filters.group) return false;
      if (filters.scope !== "Tous" && t.scope !== filters.scope) return false;
      if (filters.category !== "Toutes" && t.category !== filters.category) return false;
      if (filters.subcategory && filters.subcategory !== "Toutes" && t.subcategory !== filters.subcategory) return false;
      if (filters.search && !normalizeText(t.category).includes(normalizeText(filters.search)) && !normalizeText(t.subcategory || "").includes(normalizeText(filters.search))) return false;
      return true;
    });
  }, [txWithGroup, filters]);

  const allLoaded = txLoaded && groupsLoaded && scopeLoaded && rulesLoaded && loansLoaded && capLoaded && accountsLoaded && budgetsLoaded && goalsLoaded && recurringLoaded && syncCodeLoaded;

  // Tire l'état distant et hydrate les états locaux. Réutilisé au chargement,
  // sur demande (forcer la sync) et à chaque notification temps réel.
  const pullAndHydrate = React.useCallback(async (code: string) => {
    setSyncStatus("syncing");
    const remote = await fetchRemoteState(code);
    if (remote) {
      skipNextPush.current = true;
      if (remote.transactions) setTransactions(remote.transactions);
      if (remote.categoryGroups) setCategoryGroups(remote.categoryGroups);
      if (remote.categoryScope) setCategoryScope(remote.categoryScope);
      if (remote.rules) setRules(remote.rules);
      if (remote.loans) setLoans(remote.loans);
      if (typeof remote.envelopeCap === "number") setEnvelopeCap(remote.envelopeCap);
      if (remote.accounts) setAccounts(remote.accounts);
      if (remote.budgets) setBudgets(remote.budgets);
      if (remote.goals) setGoals(remote.goals);
      if (remote.recurring) setRecurring(remote.recurring);
      setLastSyncedAt(new Date().toLocaleTimeString("fr-FR"));
    }
    setSyncStatus("synced");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Tire l'état distant au chargement si un code de synchronisation est défini.
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) return;
    let cancelled = false;
    (async () => { if (!cancelled) await pullAndHydrate(syncCode); })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allLoaded, syncCode]);

  // Abonnement temps réel : dès qu'un autre appareil modifie la ligne distante, on la
  // retire immédiatement — sans avoir à recharger la page. Repli silencieux si le canal
  // temps réel n'est pas disponible (ex: aperçu Claude) ; le pull différé continue seul.
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) { setRealtimeConnected(false); return; }
    let unsub: (() => void) | null = null;
    let cancelled = false;
    (async () => {
      const fn = await subscribeRealtime(syncCode, () => { pullAndHydrate(syncCode); });
      if (cancelled) { fn?.(); return; }
      unsub = fn;
      setRealtimeConnected(!!fn);
    })();
    return () => { cancelled = true; unsub?.(); setRealtimeConnected(false); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allLoaded, syncCode]);

  // Pousse l'état local vers le cloud après chaque modification (avec un court délai
  // pour regrouper les changements rapprochés et éviter de spammer l'API).
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) return;
    if (skipNextPush.current) { skipNextPush.current = false; return; }
    if (pushTimer.current) clearTimeout(pushTimer.current);
    pushTimer.current = setTimeout(async () => {
      setSyncStatus("syncing");
      const ok = await pushRemoteState(syncCode, {
        transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
      });
      setSyncStatus(ok ? "synced" : "error");
      if (ok) setLastSyncedAt(new Date().toLocaleTimeString("fr-FR"));
    }, 1500);
    return () => { if (pushTimer.current) clearTimeout(pushTimer.current); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring, syncCode, allLoaded]);

  if (!allLoaded) {
    return <div style={{ minHeight: "100vh", background: COLOR.bg, color: COLOR.inkMuted, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>Chargement…</div>;
  }

  const lastNW = (() => { const s = liveNetWorthSeries(accounts, transactions); return s[s.length - 1][1]; })();

  return (
    <div style={{ minHeight: "100vh", background: COLOR.bg, color: COLOR.ink, fontFamily: "'Inter', sans-serif", display: isMobile ? "block" : "flex" }}>
      <style>{fontImport}</style>

      {/* SIDEBAR — poussé sur desktop, tiroir superposé sur mobile */}
      {(!isMobile || mobileMenuOpen) && (
        <>
          {isMobile && (
            <div onClick={() => setMobileMenuOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 210, touchAction: "none" }} />
          )}
          <aside className="gl-noprint gl-scroll" style={{
            width: isMobile ? 268 : (sidebarOpen ? 226 : 0), flexShrink: 0, borderRight: `1px solid ${COLOR.hairline}`,
            transition: isMobile ? "none" : "width 0.2s", overflowY: isMobile ? "auto" : "hidden", overflowX: "hidden",
            position: isMobile ? "fixed" : "static", top: 0, left: 0, height: isMobile ? "100dvh" : "auto", bottom: isMobile ? 0 : "auto",
            zIndex: isMobile ? 220 : "auto", background: isMobile ? COLOR.bg : "transparent",
            WebkitOverflowScrolling: "touch", overscrollBehavior: "contain",
          }}>
            <div className="gl-safe-top" style={{ width: isMobile ? 268 : 226, padding: "24px 16px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, letterSpacing: "0.16em", color: COLOR.gold, textTransform: "uppercase", paddingLeft: 8 }}>XOF</div>
                {isMobile && (
                  <button onClick={() => setMobileMenuOpen(false)} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", padding: 6 }}>
                    <X size={18} />
                  </button>
                )}
              </div>
              <h1 style={{ fontFamily: "'Fraunces', serif", fontWeight: 500, fontSize: 24, margin: "0 0 22px 0", paddingLeft: 8 }}>Grand Livre</h1>
              {NAV.map((section) => (
                <div key={section.section} style={{ marginBottom: 18 }}>
                  <div style={{ fontSize: 10, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.08em", padding: "0 8px 8px 8px" }}>{section.section}</div>
                  {section.items.map((item) => {
                    const Icon = item.icon; const active = tab === item.id;
                    return (
                      <button key={item.id} onClick={() => { setTab(item.id); setMobileMenuOpen(false); }} style={{
                        display: "flex", alignItems: "center", gap: 10, width: "100%", textAlign: "left",
                        background: active ? "rgba(201,162,39,0.1)" : "transparent", border: "none",
                        borderLeft: active ? `2px solid ${COLOR.gold}` : "2px solid transparent",
                        color: active ? COLOR.ink : COLOR.inkMuted, padding: "10px 8px 10px 10px", fontSize: 13.5, cursor: "pointer",
                        borderRadius: 4, marginBottom: 2, fontFamily: "'Inter', sans-serif",
                      }}>
                        <Icon size={15} /> {item.label}
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          </aside>
        </>
      )}

      {/* MAIN */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <header className="gl-safe-top" style={{ borderBottom: `1px solid ${COLOR.hairline}`, padding: isMobile ? "16px 16px" : "20px 32px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: isMobile ? "flex-start" : "center", flexWrap: "wrap", gap: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <button className="gl-noprint" onClick={() => (isMobile ? setMobileMenuOpen(true) : setSidebarOpen((s) => !s))} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: 7, cursor: "pointer", display: "flex" }}>
                {isMobile ? <Menu size={16} /> : <Layers size={14} />}
              </button>
              {!isMobile && (
                <div style={{ fontSize: 12.5, color: COLOR.inkMuted }}>
                  {allMonths.length ? `${monthLabel(allMonths[0])} — ${monthLabel(allMonths[allMonths.length - 1])}` : ""} · {transactions.length} transactions
                </div>
              )}
            </div>
            <div style={{ textAlign: isMobile ? "left" : "right" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, justifyContent: isMobile ? "flex-start" : "flex-end" }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, letterSpacing: "0.1em", color: COLOR.inkMuted, textTransform: "uppercase" }}>Valeur nette</div>
                {SYNC_ENABLED && syncCode && (
                  <span title={syncStatus === "synced" ? "Synchronisé" : syncStatus === "syncing" ? "Synchronisation…" : "Erreur de synchronisation"}
                    style={{ width: 6, height: 6, borderRadius: "50%", display: "inline-block", background: syncStatus === "synced" ? COLOR.emerald : syncStatus === "syncing" ? COLOR.gold : COLOR.clay }} />
                )}
              </div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: isMobile ? 21 : 26, fontWeight: 600, color: COLOR.goldSoft }}>{fmt(lastNW)}<span style={{ fontSize: 12, color: COLOR.inkMuted, marginLeft: 6 }}>FCFA</span></div>
            </div>
          </div>
        </header>

        <main className="gl-print-full" style={{ maxWidth: 1180, padding: isMobile ? "16px 14px 100px 14px" : "24px 32px 60px 32px" }}>
          {tab !== "saisie" && (
            <div className="gl-noprint" style={{ marginBottom: 20 }}>
              {isMobile ? (
                <div>
                  <button onClick={() => setFiltersOpen((o) => !o)} style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%",
                    background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "12px 16px",
                    color: COLOR.goldSoft, fontSize: 13, cursor: "pointer",
                  }}>
                    <span style={{ display: "flex", alignItems: "center", gap: 8 }}><Filter size={14} /> Filtres</span>
                    <ChevronRight size={14} style={{ transform: filtersOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                  </button>
                  {filtersOpen && (
                    <div style={{ marginTop: 10 }}>
                      <FilterBar filters={filters} setFilters={setFilters} allMonths={allMonths} allCategories={allCategories} onReset={() => setFilters(defaultFilters)} />
                    </div>
                  )}
                </div>
              ) : (
                <FilterBar filters={filters} setFilters={setFilters} allMonths={allMonths} allCategories={allCategories} onReset={() => setFilters(defaultFilters)} />
              )}
            </div>
          )}

          {tab === "apercu" && <ApercuTab filtered={filtered} filters={filters} accounts={accounts} transactions={transactions} />}
          {tab === "flux" && <FluxTab filtered={filtered} />}
          {tab === "comparatif" && <ComparatifTab transactions={transactions} categoryGroups={resolvedGroups} />}
          {tab === "comparateur" && <ComparateurTab transactions={transactions} categoryGroups={resolvedGroups} allMonths={allMonths} />}
          {tab === "topcategories" && <TopCategoriesTab transactions={transactions} setTransactions={setTransactions} categoryGroups={resolvedGroups} allMonths={allMonths} accounts={accounts} />}
          {tab === "categoryoverview" && <CategoryOverviewTab transactions={transactions} categoryGroups={resolvedGroups} allMonths={allMonths} />}
          {tab === "saisie" && <SaisieQuotidienneTab transactions={transactions} setTransactions={setTransactions} allCategories={allCategories} categoryGroups={resolvedGroups} accounts={accounts} />}
          {tab === "mensuel" && <MensuelTab filtered={filtered} />}
          {tab === "journalier" && <JournalierTab filtered={filtered} />}
          {tab === "categories" && <CategoriesTab filtered={filtered} categoryGroups={categoryGroups} resolvedGroups={resolvedGroups} setCategoryGroups={setCategoryGroups} />}
          {tab === "groupes" && <GroupesTab filtered={filtered} />}
          {tab === "enveloppes" && <EnveloppesTab filtered={filtered} cap={envelopeCap} setCap={setEnvelopeCap} />}
          {tab === "budgets" && <BudgetsTab transactions={transactions} categoryGroups={resolvedGroups} budgets={budgets} setBudgets={setBudgets} allCategories={allCategories} />}
          {tab === "simulateur" && <SimulateurTab filtered={filtered} accounts={accounts} transactions={transactions} />}
          {tab === "objectif" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              <GoalsPanel goals={goals} setGoals={setGoals} accounts={accounts} transactions={transactions} />
              <ProjectionPanel accounts={accounts} transactions={transactions} />
            </div>
          )}
          {tab === "business" && <BusinessTab transactions={transactions} categoryGroups={resolvedGroups} categoryScope={categoryScope} setCategoryScope={setCategoryScope} allCategories={allCategories} />}
          {tab === "creances" && <CreancesTab loans={loans} setLoans={setLoans} />}
          {tab === "comptes" && <ComptesTab accounts={accounts} setAccounts={setAccounts} transactions={transactions} />}
          {tab === "payees" && <PayeesTab transactions={transactions} />}
          {tab === "recurrences" && <RecurrencesTab recurring={recurring} setRecurring={setRecurring} transactions={transactions} setTransactions={setTransactions} allCategories={allCategories} accounts={accounts} />}
          {tab === "journal" && <JournalTab filtered={filtered} allCategories={allCategories} categoryGroups={resolvedGroups} transactions={transactions} setTransactions={setTransactions} rules={rules} setRules={setRules} accounts={accounts} />}
          {tab === "export" && <ExportTab filtered={filtered} filters={filters} setFilters={setFilters} allMonths={allMonths} />}
          {tab === "sauvegarde" && (
            <SauvegardeTab
              getSnapshot={() => ({
                transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
              })}
              restore={(data: any) => {
                if (data.transactions) setTransactions(data.transactions);
                if (data.categoryGroups) setCategoryGroups(data.categoryGroups);
                if (data.categoryScope) setCategoryScope(data.categoryScope);
                if (data.rules) setRules(data.rules);
                if (data.loans) setLoans(data.loans);
                if (typeof data.envelopeCap === "number") setEnvelopeCap(data.envelopeCap);
                if (data.accounts) setAccounts(data.accounts);
                if (data.budgets) setBudgets(data.budgets);
                if (data.goals) setGoals(data.goals);
                if (data.recurring) setRecurring(data.recurring);
              }}
              syncCode={syncCode}
              setSyncCode={setSyncCode}
              syncStatus={syncStatus}
              lastSyncedAt={lastSyncedAt}
              realtimeConnected={realtimeConnected}
              onForceSync={async () => {
                if (!syncCode) return;
                setSyncStatus("syncing");
                const ok = await pushRemoteState(syncCode, {
                  transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
                });
                setSyncStatus(ok ? "synced" : "error");
                if (ok) setLastSyncedAt(new Date().toLocaleTimeString("fr-FR"));
              }}
            />
          )}
        </main>

        {!isMobile && (
          <footer className="gl-noprint" style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "18px 32px", textAlign: "center", color: COLOR.inkMuted, fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace" }}>
            Grand Livre · {transactions.length} transactions · {loans.length} créance(s) suivie(s) · données stockées uniquement dans ce navigateur
          </footer>
        )}
      </div>

      {isMobile && (
        <nav className="gl-noprint gl-safe-bottom" style={{
          position: "fixed", bottom: 0, left: 0, right: 0, zIndex: 150,
          background: "rgba(14,22,17,0.96)", backdropFilter: "blur(10px)", borderTop: `1px solid ${COLOR.hairline}`,
          display: "flex", justifyContent: "space-around", padding: "8px 4px 4px 4px",
        }}>
          {[
            { id: "saisie" as Tab, label: "Saisie", icon: Clock },
            { id: "apercu" as Tab, label: "Aperçu", icon: LayoutDashboard },
            { id: "journal" as Tab, label: "Journal", icon: BookOpen },
            { id: "comptes" as Tab, label: "Comptes", icon: Wallet },
          ].map((item) => {
            const Icon = item.icon; const active = tab === item.id;
            return (
              <button key={item.id} onClick={() => setTab(item.id)} style={{
                display: "flex", flexDirection: "column", alignItems: "center", gap: 3, background: "transparent",
                border: "none", color: active ? COLOR.goldSoft : COLOR.inkMuted, cursor: "pointer", padding: "6px 10px", flex: 1,
              }}>
                <Icon size={19} />
                <span style={{ fontSize: 10 }}>{item.label}</span>
              </button>
            );
          })}
          <button onClick={() => setMobileMenuOpen(true)} style={{
            display: "flex", flexDirection: "column", alignItems: "center", gap: 3, background: "transparent",
            border: "none", color: COLOR.inkMuted, cursor: "pointer", padding: "6px 10px", flex: 1,
          }}>
            <Menu size={19} />
            <span style={{ fontSize: 10 }}>Menu</span>
          </button>
        </nav>
      )}

      {tab !== "saisie" && (
        <QuickAddFAB transactions={transactions} setTransactions={setTransactions} accounts={accounts} categoryGroups={resolvedGroups} isMobile={isMobile} />
      )}
    </div>
  );
}
