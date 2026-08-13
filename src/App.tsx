import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import seedTransactionsData from "./seedTransactions.json";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, ReferenceLine,
} from "recharts";
import {
  LayoutDashboard, CalendarRange, PiggyBank, Layers, BookOpen, TrendingUp,
  TrendingDown, Filter, X, Plus, Pencil, Trash2, Save, RotateCcw, Search,
  ArrowUpDown, Wallet, Target, AlertTriangle, Info, Check, Circle, ChevronRight, ChevronLeft,
  SlidersHorizontal, Workflow, CalendarDays, BarChart3, Briefcase, HandCoins, Clock,
  Users, Repeat, ClipboardList, UploadCloud, CheckSquare, Square, Menu, ChevronDown,
  Download, Printer, Bell, Sparkles, Gauge, ArrowRight, Percent, Upload, Mail, Rocket, Compass,
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
.gl-journal-row { transition: background 0.1s ease; }
.gl-journal-row:hover { background: rgba(201,162,39,0.06) !important; }
.gl-journal-row:hover td { border-color: rgba(201,162,39,0.25) !important; }
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

// Fusionne deux listes (locale + distante) par identifiant unique, au lieu de laisser
// l'une écraser l'autre — sur demande explicite de l'utilisateur (11/08/2026) : "je veux
// une solution qui récupère toutes les transactions saisies et corrige sur tous [les
// appareils]". Ce qui n'existe QUE d'un côté est toujours conservé : rien n'est jamais
// perdu par simple écrasement.
// Corrigé le 12/08/2026 : en cas de conflit (même id des deux côtés — typiquement une
// transaction MODIFIÉE localement, pas ajoutée), la version distante gagnait toujours,
// sans condition — une modification pas encore poussée au moment du cycle de
// synchronisation suivant était donc silencieusement écrasée par l'ancienne version.
// Comparaison par horodatage INDIVIDUEL par élément (updatedAt, posé automatiquement à
// chaque création/édition — voir setTransactionsTracked) plutôt qu'un seul verdict
// global pour toute la synchronisation : plus juste si plusieurs éléments ont été
// modifiés à des moments différents. Sans horodatage d'aucun côté (données historiques
// jamais rééditées), on garde la version distante par défaut, comme avant.
function mergeById<T extends { id: string; updatedAt?: string }>(local: T[], remote: T[]): T[] {
  const byId = new Map<string, T>();
  local.forEach((item) => byId.set(item.id, item));
  remote.forEach((item) => {
    const existing = byId.get(item.id);
    if (!existing) { byId.set(item.id, item); return; }
    if (existing.updatedAt || item.updatedAt) {
      byId.set(item.id, (item.updatedAt || "") > (existing.updatedAt || "") ? item : existing);
      return;
    }
    byId.set(item.id, item);
  });
  return Array.from(byId.values());
}

async function fetchRemoteState(syncCode: string): Promise<{ data: any; updatedAt: string } | null> {
  if (!SYNC_ENABLED || !syncCode) return null;
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/app_state?sync_code=eq.${encodeURIComponent(syncCode)}&select=data,updated_at`, {
      headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}` },
    });
    if (!res.ok) return null;
    const rows = await res.json();
    if (!rows?.[0]?.data) return null;
    return { data: rows[0].data, updatedAt: rows[0].updated_at };
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
type LoanStatus = "En attente" | "Partiellement remboursé" | "Remboursé";
interface LoanRepayment { id: string; date: string; amount: number; note?: string; }

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
  reconciled?: boolean; // pointée face au relevé bancaire réel (rapprochement bancaire)
  // Avance entre comptes : cette dépense a réellement été payée depuis "account", mais
  // concerne en réalité "onBehalfOf" (compte vide au moment du paiement, avancé par un
  // autre). Permet de suivre qui doit quoi à qui entre comptes, sans fausser les soldes
  // réels (qui restent basés sur "account", où l'argent a vraiment bougé).
  onBehalfOf?: string;
  settled?: boolean; // avance entre comptes déjà réglée entre les deux comptes concernés
  // Horodatage de la dernière modification — posé automatiquement (via
  // setTransactionsTracked, pas au cas par cas) à chaque création ou édition, quel que
  // soit le point de saisie utilisé. Permet à la fusion de synchronisation de départager
  // deux versions d'une même transaction déjà connue des deux côtés. Absent sur les
  // données historiques jamais rééditées depuis l'ajout de ce champ.
  updatedAt?: string;
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
  repayments?: LoanRepayment[];
}
interface CategorizationRule {
  id: string;
  keyword: string;
  group: Group;
}

// ============================================================
// SEED DATA (633 transactions extraites des exports MoneyCoach)
// ============================================================
const seedTransactions = seedTransactionsData as Transaction[];
const defaultCategoryGroups: Record<string, Group> = {"Logement":"Nécessaire","Aliments":"Nécessaire","Santé":"Nécessaire","Utilitaires":"Nécessaire","Voiture":"Nécessaire","Transport":"Nécessaire","Enfants & Maman":"Nécessaire","Déménagement":"Nécessaire","Securicompte":"Nécessaire","Payement Dette Orange":"Nécessaire","Des sports":"Non-productif","Dette":"Productif","Épargne":"Productif","Âge D’or Retraite":"Productif","Plan Éducation":"Productif","Achat Mazda":"Productif","Payement Maison Bingerville":"Productif","Achat Terrain Port Bouet":"Productif","GRUNDFOS":"Productif","INVEST SGO":"Productif","Création Entreprise ECO PUMP AFRIK":"Productif","Prêt":"Productif","Formation":"Productif","Éducation":"Productif","Cadeaux":"Non-productif","Divertissement":"Non-productif","Invitation":"Non-productif","Shopping":"Non-productif","Personnel":"Non-productif","Vêtements":"Non-productif","Générales":"Non-productif","Vacance Nesher":"Non-productif","Voyage":"Non-productif","Abonnements":"Non-productif","Pack Club":"Non-productif","Ajustement":"Non classifié","General":"Non classifié"};

const defaultCategoryScope: Record<string, Scope> = {
  "GRUNDFOS": "Business", "INVEST SGO": "Business", "Création Entreprise ECO PUMP AFRIK": "Business",
  "Vente Pompe": "Business", "ECO PUMP": "Business",
};

// Activités : suivi de rentabilité par activité réelle plutôt que par compte —
// l'argent circule souvent entre comptes (ex: salaire épuisé → on puise sur Petty
// Cash ou Revenus MAZDA), donc le compte n'est pas un indicateur fiable de
// l'activité. La catégorie, elle, garde son sens économique réel.
const defaultActivities: string[] = ["Personnel", "Mazda", "GRUNDFOS", "Vente Pompe", "Logement"];
const defaultCategoryActivity: Record<string, string> = {
  "Achat Mazda": "Mazda", "Voiture": "Mazda", "Revenus Location Mazda": "Mazda",
  "GRUNDFOS": "GRUNDFOS",
  "Vente Pompe": "Vente Pompe",
  "Logement": "Logement",
};

// Hiérarchie catégorie → sous-catégories, telle que définie dans MoneyCoach.
// Clés séparées par type car certains noms (ex: "Ajustement") existent des deux côtés
// avec des sous-catégories différentes.
const depSubcategories: Record<string, string[]> = {
  "Invitation": ["Femmes", "Triade"],
  "Logement": ["Location", "Résidence principale", "Deuxième logement"],
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
// Historique RÉEL de valeur nette, mois par mois, reconstitué à partir des soldes de
// départ des comptes + du cumul des vraies transactions (plutôt qu'une série fictive).
// Remplace l'ancienne série de démonstration figée, qui produisait des variations
// erratiques sans rapport avec les données réelles de l'utilisateur.
function computeMonthlyNetWorthSeries(accounts: Account[], transactions: Transaction[]): [string, number][] {
  const base = accounts.reduce((a, acc) => a + acc.openingBalance, 0);
  if (!transactions.length) return [[dateToMonthKey(todayISO()), base]];
  const months = Array.from(new Set(transactions.map((t) => dateToMonthKey(t.date)))).sort((a, b) => monthSortKey(a) - monthSortKey(b));
  const curKey = dateToMonthKey(todayISO());
  const allMonths = monthSortKey(curKey) <= monthSortKey(months[months.length - 1]) ? months : [...months, curKey];
  // Complète les mois manquants entre le premier et le dernier pour une courbe continue.
  const seq: string[] = [];
  let cursor = allMonths[0];
  const lastM = allMonths[allMonths.length - 1];
  while (monthSortKey(cursor) <= monthSortKey(lastM)) {
    seq.push(cursor);
    const [y, m] = cursor.split("_").map(Number);
    const next = new Date(y, m, 1); // m est déjà 1-indexé ici -> avance d'un mois
    cursor = `${next.getFullYear()}_${next.getMonth() + 1}`;
  }
  let running = base;
  const byMonthDelta: Record<string, number> = {};
  transactions.forEach((t) => {
    const mk = dateToMonthKey(t.date);
    byMonthDelta[mk] = (byMonthDelta[mk] || 0) + (t.type === "Revenu" ? t.amount : -t.amount);
  });
  return seq.map((mk) => { running += byMonthDelta[mk] || 0; return [mk, running] as [string, number]; });
}

function liveNetWorthSeries(accounts: Account[], transactions: Transaction[]): [string, number][] {
  const total = totalAccountsBalance(accounts, transactions);
  const series = computeMonthlyNetWorthSeries(accounts, transactions);
  const curKey = dateToMonthKey(todayISO());
  const lastKey = series[series.length - 1][0];
  if (lastKey === curKey) return [...series.slice(0, -1), [lastKey, total]];
  return [...series, [curKey, total]];
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
// Montant applicable pour un mois donné d'après un historique de segments (loyer qui a
// changé plusieurs fois par exemple) — null si aucun segment ne couvre ce mois.
function scheduleAmountFor(schedule: ChargeScheduleEntry[], mk: string): number | null {
  const entry = schedule.find((e) => monthSortKey(mk) >= monthSortKey(e.from) && (e.to === null || monthSortKey(mk) <= monthSortKey(e.to)));
  return entry ? entry.amount : null;
}
// Montant "actuel" d'un historique : le segment qui couvre aujourd'hui, sinon le plus
// récent (le plus probable pour représenter l'engagement mensuel en cours).
function currentScheduleAmount(schedule: ChargeScheduleEntry[]): number {
  const todayMk = dateToMonthKey(todayISO());
  const covering = scheduleAmountFor(schedule, todayMk);
  if (covering !== null) return covering;
  const sorted = [...schedule].sort((a, b) => monthSortKey(b.from) - monthSortKey(a.from));
  return sorted[0]?.amount ?? 0;
}
// Montant retenu pour une fenêtre de mois donnée : moyenne pondérée des segments qui la
// couvrent (ex : 8 mois à 550 000 + 2 mois à 480 000 sur une fenêtre de 10 mois → moyenne
// pondérée, pas juste la moyenne des deux montants). Les mois non couverts par aucun
// segment (trous dans l'historique) sont ignorés plutôt que comblés arbitrairement. Si
// aucun mois de la fenêtre n'est couvert, retombe sur le montant actuel.
function scheduleAmountForWindow(schedule: ChargeScheduleEntry[], lookback: string[]): number {
  const covered = lookback.map((m) => scheduleAmountFor(schedule, m)).filter((v): v is number => v !== null);
  if (!covered.length) return currentScheduleAmount(schedule);
  return covered.reduce((a, v) => a + v, 0) / covered.length;
}
const fmt = (n: number) => new Intl.NumberFormat("fr-FR").format(Math.round(n));
// Les PDF utilisent les polices standard de jsPDF (Helvetica), qui ne supportent
// que l'encodage WinAnsi/Latin-1 — l'espace fine insécable utilisée par fmt() pour
// séparer les milliers (format français réel) n'en fait pas partie et s'affichait
// comme "/" ou coupait le nombre. On utilise ici une espace normale, sûre dans
// n'importe quelle police PDF, uniquement pour ce qui part dans un export PDF.
const fmtPdf = (n: number) => Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
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
function weekdayLabel(date: string) {
  try {
    const d = new Date(date + "T00:00:00");
    const s = d.toLocaleDateString("fr-FR", { weekday: "long" });
    return s.charAt(0).toUpperCase() + s.slice(1);
  } catch { return ""; }
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
// Overrides modifiables par l'utilisateur (Gestion des catégories) — synchronisés
// depuis l'état persistant de l'app à chaque rendu. Tant qu'aucune modification
// n'a été faite, on retombe sur les listes intégrées ci-dessus.
let CUSTOM_DEP_SUBCATS: Record<string, string[]> | null = null;
let CUSTOM_REV_SUBCATS: Record<string, string[]> | null = null;
function activeSubcatMap(type: TxType): Record<string, string[]> {
  return (type === "Dépense" ? CUSTOM_DEP_SUBCATS : CUSTOM_REV_SUBCATS) || (type === "Dépense" ? depSubcategories : revSubcategories);
}
function getSubcategories(type: TxType, category: string): string[] {
  return activeSubcatMap(type)[category] || [];
}
function normalizeText(s: string): string {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}
function categoriesForType(transactions: Transaction[], type: TxType): string[] {
  const used = new Set(transactions.filter((t) => t.type === type).map((t) => t.category));
  const known = Object.keys(activeSubcatMap(type));
  known.forEach((c) => used.add(c));
  return Array.from(used).sort((a, b) => a.localeCompare(b, "fr"));
}
// Construit une liste d'options de catégorie regroupée "Dépenses" / "Revenus" (une
// catégorie qui existe dans les deux — ex: "Ajustement" — apparaît dans les deux
// groupes) — sur demande explicite de l'utilisateur (10/08/2026) : la liste plate,
// mélangée et triée alphabétiquement toutes catégories confondues était trop désordonnée.
function groupedCategoryOptions(transactions: Transaction[]): { value: string; label: string; group: string }[] {
  const dep = categoriesForType(transactions, "Dépense").map((c) => ({ value: c, label: c, group: "Dépenses" }));
  const rev = categoriesForType(transactions, "Revenu").map((c) => ({ value: c, label: c, group: "Revenus" }));
  return [...dep, ...rev];
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
// Résout le groupe (Nécessaire/Productif/Non-productif) d'une transaction : regarde
// d'abord si sa SOUS-catégorie a un groupe qui lui est propre (clé "Catégorie::Sous-catégorie"
// dans categoryGroups), sinon retombe sur le groupe de la catégorie entière. Permet à des
// sous-catégories d'une même catégorie d'avoir des natures différentes (ex : "Abonnements"
// globalement Non-productif, mais "Abonnements · Assurance SAF" classée Nécessaire).
function groupFor(t: { category: string; subcategory?: string }, categoryGroups: Record<string, Group>): Group {
  if (t.subcategory) {
    const subKey = `${t.category}::${t.subcategory}`;
    if (categoryGroups[subKey]) return categoryGroups[subKey];
  }
  return categoryGroups[t.category] || "Non classifié";
}

function CustomTooltip({ active, payload, label }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, padding: "10px 14px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLOR.ink }}>
      <div style={{ color: COLOR.inkMuted, marginBottom: 4 }}>{label}</div>
      {payload.map((p: any, i: number) => <div key={i} style={{ color: p.color || p.fill }}>{p.name}: {fmt(p.value)} FCFA</div>)}
    </div>
  );
}

// Corrigé le 11/08/2026 : des transactions récemment saisies disparaissaient parfois,
// notamment juste après un redéploiement. Deux garde-fous ajoutés :
// 1. Synchronisation entre onglets/fenêtres : si l'app est ouverte à deux endroits (ex :
//    icône PWA sur l'écran d'accueil + onglet navigateur), sans ce correctif, un onglet
//    avec des données périmées pouvait écraser silencieusement les données plus
//    récentes de l'autre en enregistrant son propre état obsolète après coup. On écoute
//    maintenant les changements de storage venant d'un autre onglet et on les adopte.
// 2. Détection de démarrage à vide : si le localStorage est vide au chargement (ex :
//    nouvel appareil, navigation privée, ou un déploiement qui change d'URL/domaine —
//    le localStorage est propre à chaque origine), l'app retombe sur les données de
//    démonstration au lieu des vraies données, ce qui RESSEMBLE à une perte de données
//    alors que rien n'a été effacé — juste chargé depuis le mauvais endroit. Signalé via
//    le 4e élément retourné pour que l'app puisse avertir clairement l'utilisateur.
// Corrigé le 11/08/2026 : le vrai défaut de conception qui causait la perte de données
// n'était pas juste "le local peut se vider" (ça, on ne peut pas l'empêcher — c'est le
// navigateur/l'OS qui décide) — c'est que dès que le local était vide, l'app se mettait
// À SAUVEGARDER SILENCIEUSEMENT les données de secours comme si c'était les vraies,
// sans jamais demander confirmation. Le paramètre "holdSave" permet de bloquer
// complètement l'écriture tant qu'un choix explicite n'a pas été fait (voir l'écran de
// blocage plus bas dans le composant racine).
function usePersistentState<T>(key: string, initial: T, gateResolved?: boolean): [T, (v: T) => void, boolean, boolean] {
  const [state, setState] = useState<T>(initial);
  const [readDone, setReadDone] = useState(false);
  const [startedEmpty, setStartedEmpty] = useState(false);
  const [foundReal, setFoundReal] = useState(false);
  useEffect(() => {
    try {
      const raw = localStorage.getItem(key);
      if (raw) { setState(JSON.parse(raw)); setFoundReal(true); }
      else setStartedEmpty(true);
    } catch {}
    setReadDone(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // On n'écrit dans le localStorage que si on a trouvé de vraies données au départ, OU
  // si le blocage a été explicitement levé (voir l'écran de choix au niveau racine) —
  // jamais automatiquement juste parce que la lecture est terminée.
  const safeToSave = readDone && (foundReal || !!gateResolved);
  useEffect(() => {
    if (!safeToSave) return;
    try {
      localStorage.setItem(key, JSON.stringify(state));
    } catch {}
  }, [state, safeToSave]);
  useEffect(() => {
    const onStorage = (e: StorageEvent) => {
      if (e.key !== key || e.newValue === null) return;
      try { setState(JSON.parse(e.newValue)); } catch {}
    };
    window.addEventListener("storage", onStorage);
    return () => window.removeEventListener("storage", onStorage);
  }, [key]);
  return [state, setState, readDone, startedEmpty];
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

// Panneau repliable — même carte visuelle que Panel, mais fermé (ou ouvert) par
// défaut avec un clic sur l'en-tête pour dérouler. Utilisé pour raccourcir les
// pages qui empilent beaucoup de contenu, sans rien retirer : tout reste
// accessible, juste replié tant qu'on n'en a pas besoin.
function CollapsibleSection({ title, subtitle, defaultOpen = false, badge, badgeColor, children }: {
  title: string; subtitle?: string; defaultOpen?: boolean; badge?: string; badgeColor?: string; children: React.ReactNode;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, overflow: "hidden" }}>
      <button onClick={() => setOpen((v) => !v)} style={{
        width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12,
        padding: "20px 24px", background: "transparent", border: "none", cursor: "pointer", textAlign: "left",
      }}>
        <div style={{ minWidth: 0 }}>
          <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 500, margin: 0, color: COLOR.ink }}>{title}</h3>
          {subtitle && <div style={{ color: COLOR.inkMuted, fontSize: 12.5, marginTop: 4 }}>{subtitle}</div>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          {badge && <span style={{ fontSize: 11, fontWeight: 700, color: badgeColor || COLOR.goldSoft, background: `${badgeColor || COLOR.gold}22`, borderRadius: 20, padding: "3px 10px", whiteSpace: "nowrap" }}>{badge}</span>}
          <ChevronDown size={17} color={COLOR.inkMuted} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s", flexShrink: 0 }} />
        </div>
      </button>
      {open && <div style={{ padding: "0 24px 24px 24px" }}>{children}</div>}
    </div>
  );
}

// Panneau avec bouton d'aide (?) — affiche une explication du graphique au clic,
// juste au-dessus du contenu, sans quitter la page. Optionnellement repliable
// (collapsible + defaultOpen) pour raccourcir les pages qui empilent plusieurs
// de ces panneaux — comportement par défaut inchangé si non précisé.
function PanelWithHelp({ title, subtitle, explain, right, children, style = {}, collapsible = false, defaultOpen = true, badge, badgeColor }: {
  title?: string; subtitle?: string; explain: string; right?: React.ReactNode; children: React.ReactNode; style?: React.CSSProperties;
  collapsible?: boolean; defaultOpen?: boolean; badge?: string; badgeColor?: string;
}) {
  const [showHelp, setShowHelp] = useState(false);
  const [open, setOpen] = useState(defaultOpen);

  const helpButton = (
    <button onClick={() => setShowHelp((s) => !s)} title="Comprendre ce graphique" style={{
      width: 24, height: 24, borderRadius: "50%", border: `1px solid ${showHelp ? COLOR.gold : COLOR.hairline}`,
      background: showHelp ? "rgba(201,162,39,0.15)" : "transparent", color: showHelp ? COLOR.goldSoft : COLOR.inkMuted,
      cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
    }}>
      <HelpCircle size={14} />
    </button>
  );
  const explainBlock = showHelp && (
    <div style={{ background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: "12px 14px", marginBottom: 16, fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
      {explain}
    </div>
  );

  if (collapsible) {
    return (
      <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, overflow: "hidden", ...style }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <button onClick={() => setOpen((v) => !v)} style={{
            flex: 1, minWidth: 0, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12,
            padding: "20px 24px", background: "transparent", border: "none", cursor: "pointer", textAlign: "left",
          }}>
            <div style={{ minWidth: 0 }}>
              {title && <h3 style={{ fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 500, margin: 0, color: COLOR.ink }}>{title}</h3>}
              {subtitle && <div style={{ color: COLOR.inkMuted, fontSize: 12.5, marginTop: 4 }}>{subtitle}</div>}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
              {badge && <span style={{ fontSize: 11, fontWeight: 700, color: badgeColor || COLOR.goldSoft, background: `${badgeColor || COLOR.gold}22`, borderRadius: 20, padding: "3px 10px", whiteSpace: "nowrap" }}>{badge}</span>}
              <ChevronDown size={17} color={COLOR.inkMuted} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s", flexShrink: 0 }} />
            </div>
          </button>
          {open && <div style={{ display: "flex", alignItems: "center", gap: 8, paddingRight: 24, flexWrap: "wrap", justifyContent: "flex-end" }}>{right}{helpButton}</div>}
        </div>
        {open && <div style={{ padding: "0 24px 24px 24px" }}>{explainBlock}{children}</div>}
      </div>
    );
  }

  return (
    <Panel
      title={title}
      subtitle={subtitle}
      style={style}
      right={
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {right}
          {helpButton}
        </div>
      }
    >
      {explainBlock}
      {children}
    </Panel>
  );
}

// Calcule le badge de comparaison "vs période passée" — flèche + mot, avec un sens qui
// s'inverse selon l'indicateur : pour un solde/revenu, monter est bon ; pour une dépense,
// monter est mauvais. Seuils : ±5% = stable, ±5-20% = bon/à surveiller, >20% = excellent/mauvais.
function compareLabel(diffPct: number | null, goodDirection: "up" | "down"): { text: string; tone: string; Icon: any } | null {
  if (diffPct === null) return null;
  const effective = goodDirection === "up" ? diffPct : -diffPct; // ramène toujours à "positif = amélioration"
  if (Math.abs(diffPct) < 5) return { text: "Stable", tone: COLOR.inkMuted, Icon: Minus };
  if (effective >= 20) return { text: "Excellent", tone: COLOR.emeraldSoft, Icon: diffPct >= 0 ? TrendingUp : TrendingDown };
  if (effective >= 5) return { text: "Bon", tone: COLOR.emeraldSoft, Icon: diffPct >= 0 ? TrendingUp : TrendingDown };
  if (effective <= -20) return { text: "Mauvais", tone: COLOR.claySoft, Icon: diffPct >= 0 ? TrendingUp : TrendingDown };
  return { text: "À surveiller", tone: COLOR.goldSoft, Icon: diffPct >= 0 ? TrendingUp : TrendingDown };
}

function Kpi({ label, value, suffix = "FCFA", tone = COLOR.ink, icon: Icon, hint, hintBadge, onDetailClick }: {
  label: string; value: string; suffix?: string; tone?: string; icon?: any; hint?: string;
  hintBadge?: { text: string; tone: string; Icon: any } | null; onDetailClick?: () => void;
}) {
  return (
    <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "16px 18px", flex: 1, minWidth: 190 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLOR.inkMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10 }}>
        {Icon && <Icon size={12.5} />} {label}{onDetailClick && <CalcDetailIcon onClick={onDetailClick} />}
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 600, color: tone }}>
        {value}<span style={{ fontSize: 11.5, color: COLOR.inkMuted, marginLeft: 6 }}>{suffix}</span>
      </div>
      {hint && (
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6, flexWrap: "wrap" }}>
          {hintBadge && (
            <span style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 10.5, fontWeight: 700, color: hintBadge.tone, background: `${hintBadge.tone}22`, borderRadius: 20, padding: "2px 7px" }}>
              <hintBadge.Icon size={11} /> {hintBadge.text}
            </span>
          )}
          <div style={{ fontSize: 11, color: COLOR.inkMuted }}>{hint}</div>
        </div>
      )}
    </div>
  );
}

function Select({ value, onChange, options, label }: {
  value: string; onChange: (v: string) => void; options: { value: string; label: string; group?: string }[]; label?: string;
}) {
  // Regroupe les options par "group" si fourni (ex: Dépenses / Revenus), en conservant
  // l'ordre d'apparition des groupes — sur demande explicite de l'utilisateur
  // (10/08/2026) : la liste de catégories mélangée dépenses/revenus était trop désordonnée.
  const hasGroups = options.some((o) => o.group);
  const groups: { name: string | null; items: typeof options }[] = [];
  options.forEach((o) => {
    const g = o.group || null;
    let bucket = groups.find((b) => b.name === g);
    if (!bucket) { bucket = { name: g, items: [] }; groups.push(bucket); }
    bucket.items.push(o);
  });
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      {label && <label style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</label>}
      <select value={value} onChange={(e) => onChange(e.target.value)} style={{
        background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink,
        padding: "8px 10px", fontSize: 12.5, fontFamily: "'Inter', sans-serif", minWidth: 130, cursor: "pointer",
      }}>
        {hasGroups
          ? groups.map((g, i) => g.name
              ? <optgroup key={g.name} label={g.name}>{g.items.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}</optgroup>
              : g.items.map((o) => <option key={o.value} value={o.value}>{o.label}</option>))
          : options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
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
interface Filters { from: string; to: string; type: string; group: string; category: string; subcategory: string; search: string; scope: string; accounts: string[]; }

// Sélecteur multi-choix simple (liste à cocher) — utilisé pour le filtre "Compte", qui
// doit permettre de choisir PLUSIEURS comptes à la fois (ex: "Petty Cash" + "Revenus
// MAZDA", pour repérer les dépenses d'une catégorie qui n'ont PAS été prélevées sur le
// compte habituel — sur demande explicite de l'utilisateur, 10/08/2026).
// Menu déroulant simple-sélection avec en-têtes de groupe colorés (Dépenses en rouille,
// Revenus en vert) — remplace un <select> natif avec <optgroup>, dont la couleur n'est
// pas fiable, en particulier sur mobile où l'OS prend le rendu en main et ignore le CSS
// de l'app. Sur demande explicite de l'utilisateur (10/08/2026).
function GroupedSingleSelect({ label, value, onChange, allLabel, options }: {
  label: string; value: string; onChange: (v: string) => void; allLabel: string;
  options: { value: string; label: string; group: string }[];
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const onClick = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false); };
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, []);
  const groupColor: Record<string, string> = { "Dépenses": COLOR.claySoft, "Revenus": COLOR.emeraldSoft };
  const groups: { name: string; items: typeof options }[] = [];
  options.forEach((o) => {
    let bucket = groups.find((b) => b.name === o.group);
    if (!bucket) { bucket = { name: o.group, items: [] }; groups.push(bucket); }
    bucket.items.push(o);
  });
  const current = options.find((o) => o.value === value);
  return (
    <div ref={ref} style={{ position: "relative", display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</label>
      <button onClick={() => setOpen((o) => !o)} style={{ background: COLOR.surfaceInput, border: `1px solid ${value ? COLOR.gold : COLOR.hairline}`, borderRadius: 6, color: current ? (groupColor[current.group] || COLOR.ink) : COLOR.ink, padding: "8px 10px", fontSize: 12.5, cursor: "pointer", minWidth: 130, textAlign: "left", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{current ? current.label : allLabel}</span> <ChevronDown size={12} color={COLOR.inkMuted} />
      </button>
      {open && (
        <div style={{ position: "absolute", top: "100%", left: 0, marginTop: 4, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: 8, zIndex: 20, minWidth: 200, maxHeight: 320, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
          <button onClick={() => { onChange(""); setOpen(false); }} style={{ display: "block", width: "100%", textAlign: "left", background: !value ? "rgba(201,162,39,0.12)" : "transparent", border: "none", color: COLOR.ink, fontSize: 12.5, cursor: "pointer", padding: "6px 8px", borderRadius: 4, marginBottom: 4 }}>{allLabel}</button>
          {groups.map((g) => (
            <div key={g.name}>
              <div style={{ fontSize: 10, fontWeight: 700, color: groupColor[g.name] || COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em", padding: "8px 8px 4px" }}>{g.name}</div>
              {g.items.map((o) => (
                <button key={o.value} onClick={() => { onChange(o.value); setOpen(false); }} style={{ display: "block", width: "100%", textAlign: "left", background: value === o.value ? "rgba(201,162,39,0.12)" : "transparent", border: "none", color: groupColor[g.name] || COLOR.ink, fontSize: 12.5, cursor: "pointer", padding: "6px 8px", borderRadius: 4 }}>{o.label}</button>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function MultiSelectDropdown({ label, options, selected, onChange }: { label: string; options: string[]; selected: string[]; onChange: (v: string[]) => void }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const onClick = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false); };
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, []);
  const toggle = (v: string) => onChange(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v]);
  const summary = selected.length === 0 ? "Tous" : selected.length === 1 ? selected[0] : `${selected.length} sélectionnés`;
  return (
    <div ref={ref} style={{ position: "relative", display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</label>
      <button onClick={() => setOpen((o) => !o)} style={{ background: COLOR.surfaceInput, border: `1px solid ${selected.length ? COLOR.gold : COLOR.hairline}`, borderRadius: 6, color: selected.length ? COLOR.goldSoft : COLOR.ink, padding: "8px 10px", fontSize: 12.5, cursor: "pointer", minWidth: 140, textAlign: "left", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
        {summary} <ChevronDown size={12} />
      </button>
      {open && (
        <div style={{ position: "absolute", top: "100%", left: 0, marginTop: 4, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: 8, zIndex: 20, minWidth: 200, maxHeight: 260, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
          {selected.length > 0 && (
            <button onClick={() => onChange([])} style={{ display: "block", width: "100%", textAlign: "left", background: "transparent", border: "none", color: COLOR.claySoft, fontSize: 11.5, cursor: "pointer", padding: "4px 6px", marginBottom: 4 }}>Tout désélectionner</button>
          )}
          {options.map((o) => (
            <label key={o} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, padding: "5px 6px", cursor: "pointer", borderRadius: 4 }}>
              <input type="checkbox" checked={selected.includes(o)} onChange={() => toggle(o)} />
              {o}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

function FilterBar({ filters, setFilters, allMonths, allCategories, categoryOptions, allAccounts, onReset }: {
  filters: Filters; setFilters: (f: Filters) => void; allMonths: string[]; allCategories: string[]; categoryOptions: { value: string; label: string; group: string }[]; allAccounts: string[]; onReset: () => void;
}) {
  // "Du mois" cale automatiquement "Au mois" sur la même valeur (l'utilisateur élargit
  // ensuite lui-même si besoin) — sur demande explicite de l'utilisateur, plutôt que de
  // laisser "Au mois" sur une ancienne valeur qui semblait ne "rien faire" visuellement.
  const setFrom = (v: string) => setFilters({ ...filters, from: v, to: v });
  // "Au mois" reste protégé contre une inversion (Au postérieur... même principe, dans
  // l'autre sens) — échange si besoin plutôt que de laisser une plage impossible.
  const setTo = (v: string) => {
    if (monthSortKey(v) < monthSortKey(filters.from)) setFilters({ ...filters, from: v, to: filters.from });
    else setFilters({ ...filters, to: v });
  };
  const patch = (p: Partial<Filters>) => setFilters({ ...filters, ...p });
  return (
    <div className="gl-noprint" style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "16px 20px", display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, color: COLOR.gold, fontSize: 11.5, marginBottom: 5, textTransform: "uppercase", letterSpacing: "0.08em" }}>
        <Filter size={13} /> Filtres
      </div>
      <Select label="Du mois" value={filters.from} onChange={setFrom} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
      <Select label="Au mois" value={filters.to} onChange={setTo} options={allMonths.filter((m) => monthSortKey(m) >= monthSortKey(filters.from)).map((m) => ({ value: m, label: monthLabel(m) }))} />
      <Select label="Type" value={filters.type} onChange={(v) => patch({ type: v })} options={[{ value: "Tous", label: "Tous" }, { value: "Dépense", label: "Dépenses" }, { value: "Revenu", label: "Revenus" }]} />
      <Select label="Groupe" value={filters.group} onChange={(v) => patch({ group: v })} options={[{ value: "Tous", label: "Tous" }, ...GROUPS.map((g) => ({ value: g, label: g })), { value: "Revenu", label: "Revenu" }]} />
      <Select label="Portée" value={filters.scope} onChange={(v) => patch({ scope: v })} options={[{ value: "Tous", label: "Tous" }, { value: "Personnel", label: "Personnel" }, { value: "Business", label: "Business" }]} />
      <GroupedSingleSelect label="Catégorie" allLabel="Toutes" value={filters.category === "Toutes" ? "" : filters.category}
        onChange={(v) => patch({ category: v || "Toutes", subcategory: "Toutes" })} options={categoryOptions} />
      {filters.category !== "Toutes" && (
        <Select label="Sous-catégorie" value={filters.subcategory} onChange={(v) => patch({ subcategory: v })}
          options={[{ value: "Toutes", label: "Toutes" }, ...Array.from(new Set([...depSubcategories[filters.category] || [], ...revSubcategories[filters.category] || []])).map((s) => ({ value: s, label: s }))]} />
      )}
      <MultiSelectDropdown label="Compte" options={allAccounts} selected={filters.accounts} onChange={(v) => patch({ accounts: v })} />
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

// Rapport narratif du Score de santé financière — explique les 3 composantes
// (taux d'épargne, poids du non-productif, stabilité des revenus) avec leurs
// vraies valeurs (pas juste le score sur 100), des exemples concrets, et une
// synthèse qui identifie la composante la plus faible.
function generateHealthScoreNarrative(tauxEpargne: number, pctNonProd: number, cv: number, health: { savingsScore: number; nonProdScore: number; stabilityScore: number; overall: number; grade: string }) {
  const sections = [
    {
      key: "savings", title: "Taux d'épargne", you: tauxEpargne, score: health.savingsScore, unit: "%",
      definition: "Part de ton revenu qui reste une fois toutes les dépenses de la période déduites — la référence utilisée ici est 20%, un repère courant en conseil patrimonial.",
      examples: ["Argent non dépensé en fin de mois", "Virement vers une épargne", "Investissement du surplus"],
      verdict: tauxEpargne >= 20 ? "Tu es au-dessus de la référence — c'est un moteur solide pour ta valeur nette." : tauxEpargne >= 10 ? "Tu es sous la référence de 20%, avec de la marge pour progresser." : "Ton taux d'épargne est faible — la quasi-totalité de ton revenu part en dépenses.",
    },
    {
      key: "nonprod", title: "Maîtrise du non-productif", you: pctNonProd, score: health.nonProdScore, unit: "%",
      definition: "Part de ton revenu partie dans des dépenses classées « Non-productif » — pas indispensables, pas un investissement.",
      examples: ["Divertissement", "Shopping non essentiel", "Sorties", "Cadeaux", "Abonnements de loisir"],
      verdict: pctNonProd <= 15 ? "Ces dépenses restent contenues — un bon signe de discipline budgétaire." : pctNonProd <= 30 ? "Ces dépenses prennent une place notable — un levier d'économie réaliste si besoin." : "Ces dépenses représentent une part importante de ton revenu — c'est probablement le levier le plus rapide pour augmenter ton épargne.",
    },
    {
      key: "stability", title: "Stabilité des revenus", you: cv * 100, score: health.stabilityScore, unit: "% de variation",
      definition: "Mesure à quel point tes revenus varient d'un mois à l'autre (coefficient de variation) — plus c'est bas, plus tes revenus sont prévisibles.",
      examples: ["Revenu fixe et régulier = stable", "Activité freelance ou saisonnière = plus variable", "Commissions ou primes ponctuelles = pics irréguliers"],
      verdict: cv <= 0.25 ? "Tes revenus sont réguliers d'un mois à l'autre, ce qui facilite la planification." : cv <= 0.5 ? "Tes revenus varient modérément — prévoir une marge de sécurité reste utile." : "Tes revenus varient beaucoup d'un mois à l'autre — difficile à planifier sans réserve de précaution solide.",
    },
  ];

  const weakest = [...sections].sort((a, b) => a.score - b.score)[0];
  const recommendation = `La composante la plus faible de ton score est "${weakest.title}" (${weakest.score.toFixed(0)}/100). ${
    weakest.key === "savings" ? "Augmenter ce taux, même progressivement, aurait le plus d'effet sur ton score global."
    : weakest.key === "nonprod" ? "Réduire un peu ces dépenses non-productives aurait le plus d'effet sur ton score global."
    : "Constituer une réserve de précaution pour absorber les mois plus faibles aiderait à compenser cette irrégularité."
  }`;

  return { sections, recommendation };
}

// Statistiques robustes : la médiane et l'écart absolu médian (MAD, mis à l'échelle
// pour être comparable à un écart-type) sont beaucoup moins sensibles à UN mois
// exceptionnel qu'une simple moyenne + écart-type classiques.
function median(arr: number[]): number {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  const n = s.length;
  return n % 2 === 0 ? (s[n / 2 - 1] + s[n / 2]) / 2 : s[(n - 1) / 2];
}
function madStdev(arr: number[]): number {
  const m = median(arr);
  const dev = arr.map((v) => Math.abs(v - m));
  return median(dev) * 1.4826; // facteur de mise à l'échelle standard pour approcher un écart-type
}

function projectNetWorth(months = 12, series: [string, number][] = netWorthRaw) {
  // Fenêtre plus large (jusqu'à 9 relevés au lieu de 6) pour qu'un seul mois
  // exceptionnel pèse moins sur la tendance et l'incertitude calculées.
  const windowSize = Math.min(9, series.length);
  const recent = series.slice(-windowSize);
  const deltas = recent.slice(1).map((v, i) => v[1] - recent[i][1]);
  const avgDelta = median(deltas);
  const sd = madStdev(deltas);
  const last = series[series.length - 1][1];
  const points = [];
  for (let i = 0; i <= months; i++) {
    points.push({
      mois: i === 0 ? "aujourd'hui" : `+${i}m`,
      central: last + avgDelta * i,
      haut: last + (avgDelta + sd) * i,
      bas: last + (avgDelta - sd) * i,
      range: [last + (avgDelta - sd) * i, last + (avgDelta + sd) * i],
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
  const [onBehalfOf, setOnBehalfOf] = useState("");
  const [note, setNote] = useState("");
  const [catPickerOpen, setCatPickerOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    if (open && transaction) {
      setDate(transaction.date); setTime(transaction.time || nowTime()); setType(transaction.type);
      setCategory(transaction.category); setSubcategory(transaction.subcategory || "");
      setAmount(transaction.amount); setAccount(transaction.account || defaultQuickAccount(accounts));
      setOnBehalfOf(transaction.onBehalfOf || "");
      setNote(transaction.note || "");
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
    onSave({ ...transaction, date, time, type, category, subcategory: subcategory || undefined, amount: Number(amount), account: account || undefined, onBehalfOf: (onBehalfOf && onBehalfOf !== account) ? onBehalfOf : undefined, note: note || undefined });
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
            <input
              value={note} onChange={(e) => setNote(e.target.value)} placeholder="Ajouter une note"
              style={{ position: "relative", background: "transparent", border: "none", outline: "none", marginTop: 10, color: COLOR.inkMuted, fontSize: 13.5, fontFamily: "'Inter', sans-serif", textAlign: "center", width: "100%", maxWidth: 320 }}
            />
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
            {accounts.length > 1 && (
              <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px dashed ${COLOR.hairline}` }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: COLOR.inkMuted, cursor: "pointer" }}>
                  <input type="checkbox" checked={!!onBehalfOf} onChange={(e) => setOnBehalfOf(e.target.checked ? (accounts.find((a) => a.name !== account)?.name || "") : "")} />
                  Payée depuis {account || "ce compte"} mais destinée à un autre compte (avance entre comptes)
                </label>
                {onBehalfOf && (
                  <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Compte réellement concerné :</span>
                    <select value={onBehalfOf} onChange={(e) => setOnBehalfOf(e.target.value)} style={{ ...inputStyle, flex: 1 }}>
                      {accounts.filter((a) => a.name !== account).map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                    </select>
                  </div>
                )}
              </div>
            )}
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

function ExpertAnalysisButton({ filtered, catFocus, title, subtitle }: { filtered: any[]; catFocus?: CatFocus; title?: string; subtitle?: string }) {
  const [open, setOpen] = useState(false);
  const insights = useMemo(() => generateInsights(filtered, catFocus), [filtered, catFocus]);
  const styleFor: Record<InsightKind, { bg: string; border: string; color: string; icon: any }> = {
    alerte: { bg: "rgba(193,84,63,0.08)", border: COLOR.clay, color: COLOR.claySoft, icon: AlertTriangle },
    conseil: { bg: "rgba(201,162,39,0.08)", border: COLOR.gold, color: COLOR.goldSoft, icon: Info },
    positif: { bg: "rgba(63,156,122,0.08)", border: COLOR.emerald, color: COLOR.emeraldSoft, icon: Check },
  };
  const alertCount = insights.filter((i) => i.kind === "alerte").length;
  const finalTitle = title || "Analyse d'expert financier";
  const finalSubtitle = subtitle || "Critique et recommandations générées à partir de la période et des filtres actifs";

  return (
    <>
      <button onClick={() => setOpen(true)} style={{
        display: "flex", alignItems: "center", gap: 14, width: "100%", textAlign: "left", cursor: "pointer",
        background: `linear-gradient(180deg, ${COLOR.surfaceRaised} 0%, ${COLOR.surface} 100%)`, border: `1px solid ${COLOR.hairline}`,
        borderRadius: 14, padding: "16px 18px",
      }}>
        <div style={{ width: 38, height: 38, borderRadius: 10, background: "rgba(201,162,39,0.12)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Gauge size={18} color={COLOR.goldSoft} />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 14, fontWeight: 600, color: COLOR.ink, fontFamily: "'Fraunces', serif" }}>{finalTitle}</div>
          <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{finalSubtitle}</div>
        </div>
        {alertCount > 0 && (
          <span style={{ background: "rgba(193,84,63,0.15)", color: COLOR.claySoft, fontSize: 11, fontWeight: 700, borderRadius: 20, padding: "3px 9px", flexShrink: 0 }}>{alertCount} alerte{alertCount > 1 ? "s" : ""}</span>
        )}
        <ChevronRight size={16} color={COLOR.inkMuted} style={{ flexShrink: 0 }} />
      </button>

      {open && (
        <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 460, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{
            width: "100%", maxWidth: 560, maxHeight: "86vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
            display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "18px 20px", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <div>
                <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, color: COLOR.ink, display: "flex", alignItems: "center", gap: 8 }}>
                  <Gauge size={17} color={COLOR.goldSoft} /> {finalTitle}
                </div>
                <div style={{ fontSize: 12, color: COLOR.inkMuted, marginTop: 4 }}>{finalSubtitle}</div>
              </div>
              <button onClick={() => setOpen(false)} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}><X size={18} /></button>
            </div>
            <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: 18, WebkitOverflowScrolling: "touch" }}>
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
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ============================================================
// CONSEILLER QUOTIDIEN — analyse le comportement réel (jours passés, même
// période le mois dernier) pour donner un seuil de dépense concret pour
// aujourd'hui, au regard d'un objectif mensuel défini par l'utilisateur.
// ============================================================
function prevMonthKey(mk: string): string {
  const [y, m] = mk.split("_").map(Number);
  return m === 1 ? `${y - 1}_12` : `${y}_${m - 1}`;
}
function nextMonthKey(mk: string): string {
  const [y, m] = mk.split("_").map(Number);
  return m === 12 ? `${y + 1}_1` : `${y}_${m + 1}`;
}
function addDays(dateISO: string, n: number): string {
  const d = new Date(dateISO + "T00:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}
function daysInMonthOf(mk: string): number {
  const [y, m] = mk.split("_").map(Number);
  return new Date(y, m, 0).getDate();
}

// ============================================================
// CHARGES FIXES vs VARIABLES — classification par régularité (mois présents
// sur 6) et coefficient de variation (CV = écart-type / moyenne). Un poste
// est "Fixe" si régulier et peu variable, "Variable régulière" s'il revient
// souvent mais avec un montant fluctuant, "Occasionnel" sinon. Certaines
// catégories (ex: Enfants & Maman) sont éclatées par sous-catégorie, et
// l'utilisateur peut surcharger le mode/montant pour chaque poste.
// ============================================================
type ChargeMode = "fixe" | "variable" | "occasionnelle" | "exclu";
// Un segment d'historique de montant : valable du mois "from" au mois "to" inclus
// (to: null = toujours en cours). Permet de représenter un loyer qui a changé plusieurs
// fois dans le temps plutôt qu'un montant fixe unique pour toute l'historique.
interface ChargeScheduleEntry { from: string; to: string | null; amount: number; }
interface ChargeOverride { mode: ChargeMode | "auto"; amount?: number; schedule?: ChargeScheduleEntry[]; }
interface SettingsLogEntry { at: string; text: string; }
// Catégories éclatées par sous-catégorie plutôt qu'agrégées — nécessaire pour
// distinguer par exemple GRUNDFOS·Carburant (fixe) de GRUNDFOS·Électricité (variable).
// "Logement" retiré le 07/08/2026 sur demande explicite de l'utilisateur : l'éclatement
// par sous-catégorie fragmentait ses transactions existantes en plusieurs postes
// ("Location", "(non précisé)") au lieu d'un seul "Logement" — il veut une seule ligne.
const EXPAND_SUBCATS_FOR_CHARGES: Record<string, boolean> = { "Enfants & Maman": true, "GRUNDFOS": true, "Voiture": true, "Abonnements": true };
// La catégorie que l'utilisateur veut pouvoir inclure/exclure en un clic — GRUNDFOS
// uniquement (précisé le 07/08/2026 : ça n'a rien à voir avec l'activité achat/vente
// de pompes ECO PUMP AFRIK). "Voiture" reste désormais TOUJOURS incluse dans le budget
// personnel, dans les deux options — elle n'est plus concernée par ce bouton.
const GRUNDFOS_VOITURE_CATEGORIES = ["GRUNDFOS"];
// Revenu directement lié à GRUNDFOS, confirmé par l'utilisateur le 07/08/2026 :
// "Petty Cash" finance GRUNDFOS. Exclure la dépense sans exclure ce revenu gonflait
// artificiellement tous les ratios (le "revenu" comptait de l'argent qui ne sert en
// réalité qu'à payer la dépense exclue) — désormais symétrique.
const GRUNDFOS_VOITURE_LINKED_REVENUE = ["Petty Cash"];

const defaultChargeOverrides: Record<string, ChargeOverride> = {
  // "Logement" éclaté par sous-catégorie le 07/08/2026 : deux logements distincts payés
  // séparément (résidence principale, deuxième logement), fusionnés à tort dans une
  // seule catégorie jusqu'ici. Sur demande explicite de l'utilisateur (07/08/2026) :
  // aucun montant "Fixe" n'est pré-rempli ici pour Logement — c'est à l'utilisateur de
  // décider si un poste est fixe et à quel montant, l'app se contente de proposer la
  // médiane calculée automatiquement tant qu'aucun choix n'a été fait.
  "Enfants & Maman::Maman": { mode: "fixe", amount: 70700 },
  "Enfants & Maman::Nesher": { mode: "variable" },
  "Enfants & Maman::Hemra": { mode: "variable" },
  "Utilitaires": { mode: "exclu" },
  // --- GRUNDFOS, analysé sous-catégorie par sous-catégorie ---
  "GRUNDFOS::Carburant": { mode: "fixe", amount: 154250 }, // CV 14%, 6/6 mois — vraiment régulier
  "GRUNDFOS::Internet": { mode: "fixe", amount: 35000 }, // médiane réelle ; le chiffre de 25 000 (fibre) donné ne correspond pas exactement aux données — à ajuster si tu factures la fibre séparément
  "GRUNDFOS::Électricité": { mode: "variable" }, // CV 72% — vraiment variable, pas une charge fixe malgré la régularité
  "GRUNDFOS::Eau": { mode: "occasionnelle" }, // seulement 3/6 mois et montants faibles (<8 000) — ne ressemble pas à une vraie facture mensuelle dans ces données
  // --- Voiture ---
  "Voiture::Assurance": { mode: "fixe", amount: 23698 }, // Prime annuelle (281 555 + 2 815 = 284 370 FCFA, transaction du 21 avril) ÷ 12 — mise à jour le 07/08/2026
  // --- Abonnements ---
  "Abonnements::Money Coach": { mode: "fixe", amount: 5450 }, // stable à ~5 300-5 600 sur 5 des 6 mois, un seul mois manqué
  "Abonnements::Claude": { mode: "variable" }, // abonnement récent (3/6 mois), montant encore instable — à reclasser en Fixe une fois stabilisé
};

function classifyCharges(transactions: Transaction[], overrides: Record<string, ChargeOverride>, includeGrundfosVoiture: boolean, monthsWindow: number = 6, explicitRange?: [string, string]) {
  // explicitRange (format "YYYY_M") permet de faire porter l'analyse sur la période
  // exacte sélectionnée dans le filtre global "Du mois / Au mois" plutôt que de
  // toujours recalculer une fenêtre glissante depuis aujourd'hui — pour que ce filtre
  // ait un effet réel sur Diagnostic Financier / Charges Fixes & Variables.
  let lookback: string[];
  if (explicitRange) {
    lookback = [];
    let mk = explicitRange[0];
    while (monthSortKey(mk) <= monthSortKey(explicitRange[1])) { lookback.push(mk); mk = nextMonthKey(mk); }
  } else {
    const today = todayISO();
    const curMonth = dateToMonthKey(today);
    lookback = [];
    let mk = prevMonthKey(curMonth);
    for (let i = 0; i < monthsWindow; i++) { lookback.push(mk); mk = prevMonthKey(mk); }
  }

  const byPosteMonth: Record<string, Record<string, number>> = {};
  // Repère aussi TOUTES les valeurs historiques par poste, mois par mois, indépendamment
  // de la fenêtre filtrée — sert de base à la CLASSIFICATION (Fixe/Variable/Occasionnelle)
  // ci-dessous, qui doit rester stable quel que soit le filtre "Du mois/Au mois" affiché.
  // Sans ça, resserrer le filtre sur une courte période peut faire passer par coïncidence
  // n'importe quelle catégorie pour "régulière" et la classer Fixe automatiquement — ce
  // n'est pas censé être une décision qui dépend de ce qu'on est juste en train de regarder.
  const byPosteMonthAllHistory: Record<string, Record<string, number>> = {};
  const byPosteAllTime: Record<string, number[]> = {};
  transactions.forEach((t) => {
    if (t.type !== "Dépense") return;
    if (!includeGrundfosVoiture && GRUNDFOS_VOITURE_CATEGORIES.includes(t.category)) return;
    const tmk = dateToMonthKey(t.date);
    const poste = EXPAND_SUBCATS_FOR_CHARGES[t.category] ? `${t.category}::${t.subcategory || "(non précisé)"}` : t.category;
    if (lookback.includes(tmk)) {
      byPosteMonth[poste] = byPosteMonth[poste] || {};
      byPosteMonth[poste][tmk] = (byPosteMonth[poste][tmk] || 0) + t.amount;
    }
    byPosteMonthAllHistory[poste] = byPosteMonthAllHistory[poste] || {};
    byPosteMonthAllHistory[poste][tmk] = (byPosteMonthAllHistory[poste][tmk] || 0) + t.amount;
    byPosteAllTime[poste] = byPosteAllTime[poste] || [];
    byPosteAllTime[poste].push(t.amount);
  });

  // Fenêtre de classification : tout l'historique, du premier mois de transactions au
  // mois précédant aujourd'hui — toujours la même, jamais celle du filtre affiché.
  const todayForClass = todayISO();
  const curMonthForClass = dateToMonthKey(todayForClass);
  const allMonthKeys = transactions.map((t) => dateToMonthKey(t.date));
  const firstMonthForClass = allMonthKeys.length ? allMonthKeys.reduce((a, b) => (monthSortKey(b) < monthSortKey(a) ? b : a)) : prevMonthKey(curMonthForClass);
  const classificationLookback: string[] = [];
  { let mk = firstMonthForClass; const end = prevMonthKey(curMonthForClass); while (monthSortKey(mk) <= monthSortKey(end)) { classificationLookback.push(mk); mk = nextMonthKey(mk); } }

  // Seuils de régularité exprimés en proportion du nombre de mois observés (≈83%
  // et ≈67%, les mêmes ratios qu'avec la fenêtre de 6 mois d'origine) plutôt qu'en
  // nombre de mois fixe — pour rester cohérents quelle que soit la taille de la fenêtre.
  const fixedRatio = 5 / 6, variableRatio = 4 / 6;
  const allPostes = new Set([...Object.keys(byPosteMonth), ...Object.keys(byPosteMonthAllHistory)]);
  const rows = Array.from(allPostes).map((poste) => {
    const months = byPosteMonth[poste] || {};
    const vals = lookback.map((m) => months[m] || 0);
    const present = vals.filter((v) => v > 0).length;
    // Sur demande explicite de l'utilisateur (07/08/2026) : la moyenne sur la période
    // (mois d'absence inclus comme des zéros) sert désormais de référence principale
    // pour "le montant retenu" AFFICHÉ, à la place de la médiane utilisée jusqu'ici.
    const meanV = mean(vals);
    const medianV = median(vals);
    const presentVals = vals.filter((v) => v > 0);
    const medianPresent = presentVals.length ? median(presentVals) : 0;
    const meanPresent = presentVals.length ? mean(presentVals) : 0;
    const allTimeVals = byPosteAllTime[poste] || [];
    const medianAllTime = allTimeVals.length ? median(allTimeVals) : 0;
    const meanAllTime = allTimeVals.length ? mean(allTimeVals) : 0;

    // Stats de CLASSIFICATION (mode Fixe/Variable/Occasionnelle) : toujours sur tout
    // l'historique (classificationLookback), jamais sur la fenêtre filtrée affichée.
    const classMonths = byPosteMonthAllHistory[poste] || {};
    const classVals = classificationLookback.map((m) => classMonths[m] || 0);
    const classPresent = classVals.filter((v) => v > 0).length;
    const classPresentRatio = classificationLookback.length ? classPresent / classificationLookback.length : 0;
    const classMean = mean(classVals);
    const classSd = stdev(classVals);
    const cv = classMean > 0 ? (classSd / classMean) * 100 : null;

    const override = overrides[poste];
    let mode: ChargeMode;
    let amount: number;
    if (override && override.mode !== "auto") {
      mode = override.mode;
      if (override.mode === "fixe" && override.schedule && override.schedule.length) {
        amount = currentScheduleAmount(override.schedule);
      } else {
        amount = override.mode === "fixe" && override.amount !== undefined ? override.amount : (present >= 3 ? (meanV || meanPresent) : meanAllTime);
      }
    } else if (classPresentRatio >= fixedRatio && cv !== null && cv < 20) {
      mode = "fixe"; amount = meanV || meanAllTime;
    } else if (classPresentRatio >= variableRatio) {
      mode = "variable"; amount = meanV || meanAllTime;
    } else {
      mode = "occasionnelle"; amount = meanV || meanPresent || meanAllTime;
    }
    return { poste, present, classPresent, classTotal: classificationLookback.length, mean: meanV, median: medianV, medianPresent, medianAllTime, meanPresent, meanAllTime, cv, mode, amount, overridden: !!(override && override.mode !== "auto") };
  }).sort((a, b) => b.amount - a.amount);

  const totalFixe = rows.filter((r) => r.mode === "fixe").reduce((a, r) => a + r.amount, 0);
  const totalVariable = rows.filter((r) => r.mode === "variable").reduce((a, r) => a + r.amount, 0);
  const totalOccasionnelle = rows.filter((r) => r.mode === "occasionnelle").reduce((a, r) => a + r.amount, 0);

  const revByMonth: Record<string, number> = {};
  transactions.forEach((t) => {
    if (t.type !== "Revenu") return;
    if (!includeGrundfosVoiture && GRUNDFOS_VOITURE_LINKED_REVENUE.includes(t.category)) return;
    const tmk = dateToMonthKey(t.date);
    if (lookback.includes(tmk)) revByMonth[tmk] = (revByMonth[tmk] || 0) + t.amount;
  });
  const avgRevenu = mean(lookback.map((m) => revByMonth[m] || 0));
  const resteAVivre = avgRevenu - totalFixe - totalVariable;

  return { rows, totalFixe, totalVariable, totalOccasionnelle, avgRevenu, resteAVivre, lookback };
}

// Construit la liste des mois depuis le tout premier mois de transactions
// disponible jusqu'au mois précédant aujourd'hui — pour les analyses qui doivent
// porter sur tout l'historique (ex : indicateurs asiatiques) plutôt que sur une
// fenêtre glissante de 6 mois, sur suggestion explicite de l'utilisateur.
function monthsSinceInception(transactions: Transaction[]): number {
  if (!transactions.length) return 6;
  const today = todayISO();
  const curMonth = dateToMonthKey(today);
  const firstMonth = transactions.map((t) => dateToMonthKey(t.date)).sort((a, b) => monthSortKey(a) - monthSortKey(b))[0];
  return Math.max(1, monthSortKey(prevMonthKey(curMonth)) - monthSortKey(firstMonth) + 1);
}

// ============================================================
// RATIOS FINANCIERS INSTITUTIONNELS — les mêmes repères que ceux utilisés par
// les banques, le CFPB (régulateur américain de la protection financière) et
// les cabinets de conseil en gestion de patrimoine, appliqués aux vraies
// données de l'utilisateur plutôt qu'à des hypothèses.
// ============================================================
type RatioVerdict = "sain" | "vigilance" | "risque";
interface RatioResult { key: string; label: string; value: number; unit: "%" | "mois" | "FCFA"; verdict: RatioVerdict; benchmark: string; explain: string; }

function computeFinancialRatios(
  transactions: Transaction[], accounts: Account[], chargeOverrides: Record<string, ChargeOverride>, includeGrundfosVoiture: boolean, explicitRange?: [string, string]
) {
  // Respecte le filtre global "Du mois / Au mois" quand il est fourni ; sinon, tout
  // l'historique depuis la première transaction (sur suggestion de l'utilisateur, 06/08/2026).
  const windowMonths = monthsSinceInception(transactions);
  const charges = classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, windowMonths, explicitRange);
  const netWorth = totalAccountsBalance(accounts, transactions);
  const essentialMonthly = charges.totalFixe + charges.totalVariable;

  const ratios: RatioResult[] = [];

  // 1. Taux d'épargne — référence académique et institutionnelle : 20% (règle popularisée
  // par les cabinets de conseil en gestion de patrimoine, ex. Fidelity, et enseignée en
  // finance personnelle).
  const tauxEpargne = charges.avgRevenu > 0 ? ((charges.avgRevenu - essentialMonthly) / charges.avgRevenu) * 100 : 0;
  ratios.push({
    key: "epargne", label: "Taux d'épargne", value: tauxEpargne, unit: "%",
    verdict: tauxEpargne >= 20 ? "sain" : tauxEpargne >= 10 ? "vigilance" : "risque",
    benchmark: "Référence : ≥ 20% (règle largement utilisée en conseil patrimonial)",
    explain: "Part du revenu mensuel moyen qui n'est ni consommée en charges fixes ni en dépenses variables régulières.",
  });

  // 2. Ratio charges fixes / revenu — équivalent du "debt-to-income ratio" (DTI) utilisé
  // par les banques pour l'octroi de crédit. Seuils standards : < 36% sain, 36-43% vigilance,
  // > 43% zone à risque (repère utilisé aux États-Unis pour les prêts hypothécaires qualifiés).
  const dti = charges.avgRevenu > 0 ? (charges.totalFixe / charges.avgRevenu) * 100 : 0;
  ratios.push({
    key: "dti", label: "Charges fixes / Revenu", value: dti, unit: "%",
    verdict: dti < 36 ? "sain" : dti <= 43 ? "vigilance" : "risque",
    benchmark: "Référence bancaire (DTI) : < 36% sain · 36-43% vigilance · > 43% risqué",
    explain: "Part du revenu mensuel absorbée par les seules charges fixes — l'équivalent du ratio d'endettement utilisé par les banques pour juger de ta capacité à emprunter.",
  });

  // 3. Ratio logement / revenu — règle des 30% utilisée par les organismes de logement
  // (ex. HUD aux États-Unis) et la plupart des banques pour évaluer un dossier de prêt.
  const logementCharge = charges.rows.find((r) => r.poste === "Logement");
  const logementRatio = logementCharge && charges.avgRevenu > 0 ? (logementCharge.amount / charges.avgRevenu) * 100 : null;
  if (logementRatio !== null) {
    ratios.push({
      key: "logement", label: "Logement / Revenu", value: logementRatio, unit: "%",
      verdict: logementRatio < 30 ? "sain" : logementRatio <= 40 ? "vigilance" : "risque",
      benchmark: "Règle des 30% (référence internationale logement) · > 40% considéré à risque",
      explain: "Part du revenu mensuel absorbée par le seul poste logement.",
    });
  }

  // 4. Fonds d'urgence — nombre de mois de charges essentielles couverts par la valeur
  // nette actuelle. Référence CFPB / conseillers financiers : 3 à 6 mois de dépenses
  // essentielles en réserve liquide.
  const moisCouverture = essentialMonthly > 0 ? netWorth / essentialMonthly : 0;
  ratios.push({
    key: "urgence", label: "Fonds d'urgence", value: moisCouverture, unit: "mois",
    verdict: moisCouverture >= 6 ? "sain" : moisCouverture >= 3 ? "vigilance" : "risque",
    benchmark: "Référence CFPB / conseil patrimonial : 3 à 6 mois de charges essentielles en réserve",
    explain: "Combien de mois ta valeur nette actuelle couvrirait tes charges fixes + variables régulières si tes revenus s'arrêtaient complètement.",
  });

  // 5. Concentration des revenus — logique inspirée de l'indice de concentration (type
  // Herfindahl-Hirschman) utilisé en gestion des risques : une seule source de revenu
  // dominante = fragilité en cas de choc sur cette source.
  const lookback = charges.lookback;
  const revByCat: Record<string, number> = {};
  transactions.forEach((t) => {
    if (t.type !== "Revenu" || !lookback.includes(dateToMonthKey(t.date))) return;
    if (!includeGrundfosVoiture && GRUNDFOS_VOITURE_LINKED_REVENUE.includes(t.category)) return;
    revByCat[t.category] = (revByCat[t.category] || 0) + t.amount;
  });
  const totalRevSix = Object.values(revByCat).reduce((a, v) => a + v, 0);
  const topShare = totalRevSix > 0 ? (Math.max(0, ...Object.values(revByCat)) / totalRevSix) * 100 : 0;
  const topSource = Object.entries(revByCat).sort((a, b) => b[1] - a[1])[0]?.[0] || "—";
  ratios.push({
    key: "concentration", label: "Concentration des revenus", value: topShare, unit: "%",
    verdict: topShare < 50 ? "sain" : topShare <= 75 ? "vigilance" : "risque",
    benchmark: "< 50% diversifié · 50-75% concentré · > 75% forte dépendance à une seule source",
    explain: `Part du revenu total (${charges.lookback.length} mois) apportée par la plus grosse source ("${topSource}") — plus c'est élevé, plus un choc sur cette seule source affecterait l'ensemble du budget.`,
  });

  return { ratios, netWorth, essentialMonthly, topSource, topShare };
}

// Rapport narratif pour les 5 ratios institutionnels (DTI bancaire, règle des
// 30% logement, fonds d'urgence CFPB, concentration des revenus, taux d'épargne)
// — même esprit que le rapport asiatique, mais avec un contenu propre à chaque
// ratio (définition, exemples concrets, verdict basé sur le seuil déjà calculé).
function generateRatiosNarrative(ratios: RatioResult[]): { sections: { key: string; title: string; you: number; unit: string; definition: string; examples: string[]; verdict: string; benchmark: string }[]; checks: { ok: boolean; text: string }[]; recommendation: string } {
  const meta: Record<string, { definition: string; examples: string[]; goodPhrase: string; okPhrase: string; badPhrase: string }> = {
    epargne: {
      definition: "Part de ton revenu mensuel qui n'est ni consommée en charges fixes ni en dépenses variables régulières — donc disponible pour être épargnée ou investie.",
      examples: ["Virement automatique vers une épargne", "Argent non dépensé en fin de mois", "Provision pour un projet futur"],
      goodPhrase: "C'est un très bon niveau, au-dessus de la référence de conseil patrimonial (20%).",
      okPhrase: "Tu es dans une zone correcte mais avec une marge de progression avant d'atteindre un niveau confortable.",
      badPhrase: "C'est un niveau bas — la plupart de ton revenu part en charges, laissant peu de marge pour te constituer un patrimoine.",
    },
    dti: {
      definition: "Part de ton revenu absorbée par tes seules charges fixes — l'équivalent du ratio d'endettement (debt-to-income) que les banques utilisent pour juger de ta capacité à emprunter.",
      examples: ["Mensualités de prêt", "Loyer fixe", "Abonnements obligatoires", "Remboursements de dettes"],
      goodPhrase: "Tes charges fixes laissent une vraie marge de manœuvre sur ton revenu.",
      okPhrase: "Tes charges fixes prennent une part significative de ton revenu — à surveiller si de nouvelles charges s'ajoutent.",
      badPhrase: "Tes charges fixes absorbent une très grosse part de ton revenu — une banque considérerait probablement ce niveau comme risqué pour un nouveau prêt.",
    },
    logement: {
      definition: "Part de ton revenu consacrée uniquement à ton logement.",
      examples: ["Loyer", "Mensualité de prêt immobilier", "Charges de copropriété"],
      goodPhrase: "Ton logement reste dans une proportion saine de ton budget.",
      okPhrase: "Ton logement pèse plus lourd que la référence habituelle, sans être alarmant.",
      badPhrase: "Ton logement absorbe une part très importante de ton revenu, ce qui réduit d'autant ta capacité à épargner ou investir ailleurs.",
    },
    urgence: {
      definition: "Combien de mois ta valeur nette actuelle couvrirait tes charges essentielles si tes revenus s'arrêtaient complètement.",
      examples: ["Perte d'emploi", "Accident ou maladie", "Panne majeure", "Réparation imprévue importante"],
      goodPhrase: "Tu as un vrai coussin de sécurité en cas de coup dur.",
      okPhrase: "Tu as un début de réserve, mais elle ne couvrirait qu'une partie d'un imprévu prolongé.",
      badPhrase: "Ta réserve de précaution est très faible — un imprévu sérieux (perte de revenu, urgence médicale) te mettrait rapidement en difficulté.",
    },
    concentration: {
      definition: "Part de ton revenu total qui provient d'une seule et même source — plus c'est élevé, plus un choc sur cette source affecterait tout ton budget d'un coup.",
      examples: ["Salaire d'un seul employeur", "Un seul client majeur", "Une seule activité locative"],
      goodPhrase: "Tes revenus sont bien diversifiés entre plusieurs sources.",
      okPhrase: "Une source domine sans être totalement exclusive — un vrai plan B existe si elle faiblit.",
      badPhrase: "Une seule source représente l'essentiel de ton revenu — si elle s'arrête, il n'y a presque rien pour compenser.",
    },
  };

  const sections = ratios.map((r) => {
    const m = meta[r.key];
    const verdict = r.verdict === "sain" ? m.goodPhrase : r.verdict === "vigilance" ? m.okPhrase : m.badPhrase;
    return { key: r.key, title: r.label, you: r.value, unit: r.unit, definition: m.definition, examples: m.examples, verdict, benchmark: r.benchmark };
  });

  const checks = ratios.map((r) => ({ ok: r.verdict === "sain", text: `${r.label} : ${r.unit === "mois" ? r.value.toFixed(1) + " mois" : Math.round(r.value) + r.unit} — ${r.verdict === "sain" ? "sain" : r.verdict === "vigilance" ? "à surveiller" : "en risque"}.` }));

  const worstRatio = ratios.filter((r) => r.verdict !== "sain").sort((a, b) => (a.verdict === "risque" ? -1 : 1) - (b.verdict === "risque" ? -1 : 1))[0];
  const recommendation = worstRatio
    ? `Le point le plus urgent à traiter est "${worstRatio.label}" (${worstRatio.unit === "mois" ? worstRatio.value.toFixed(1) + " mois" : Math.round(worstRatio.value) + worstRatio.unit}) — ${worstRatio.benchmark.toLowerCase()}. C'est ce ratio qui, une fois amélioré, réduirait le plus ta fragilité financière globale.`
    : "Tous tes ratios institutionnels sont dans la zone saine — la priorité devient de consolider cette position plutôt que de corriger un point faible.";

  return { sections, checks, recommendation };
}

// ============================================================
// INDICATEURS ASIATIQUES — cadres de gestion financière personnelle
// popularisés en Chine, au Japon et à Singapour, distincts des repères
// occidentaux (banques US/CFPB) déjà présents dans le Diagnostic Financier.
// ============================================================
const CN_4321_CATEGORIES = {
  // Achat d'actifs (immobilier, véhicule générant des revenus, terrain, placements) +
  // investissement en capital humain (formation, éducation) — tout ce qui construit un
  // patrimoine ou une capacité future, par opposition à une dépense consommée.
  // "Bourse" et "ECO PUMP" retirés le 06/08/2026 : confirmés par l'utilisateur comme
  // purement des revenus encaissés (dividende, vente) chez lui — aucune dépense
  // réelle dans ces catégories, ils ne servaient donc jamais à rien ici.
  // "Dette" ajoutée : remboursement d'un emprunt de 6 664 569 FCFA qui a directement
  // financé l'achat de la Mazda mise en location (1 119 029 FCFA/mois de revenus) —
  // c'est un coût de financement d'un actif, pas une dépense de vie courante.
  // "Prêt" ajouté : confirmé par l'utilisateur comme des prêts ACCORDÉS à des tiers
  // (505 000 + 250 000 + 10 100 FCFA) — du capital déployé avec attente de retour,
  // pas une dépense consommée (à ne pas confondre avec "Dette", qui est un emprunt
  // que l'utilisateur rembourse — voir aussi l'onglet Créances pour le suivi détaillé).
  investissement: ["INVEST SGO", "Épargne", "Achat Terrain Port", "Création Entreprise", "PAYEMENT MAISON", "Achat MAZDA", "FORMATION", "Éducation", "Dette", "Prêt"],
  protection: ["Âge D’or Retraite", "Plan Éducation", "Securicompte"],
};
const KAKEIBO_CATEGORIES = {
  // Structure ajustée par l'utilisateur le 06/08/2026, adaptée à son niveau de
  // revenus et ses habitudes réelles (ex : Pack Club = besoin bancaire courant,
  // Abonnements/Cadeaux/Invitations à envies plutôt qu'imprévus, Utilitaires
  // interprété comme téléphone/accessoires = une envie plutôt qu'un besoin).
  // "Loyer" retiré : confirmé par l'utilisateur comme un loyer ENCAISSÉ (trimestriel,
  // ~500 000 FCFA/mois équivalent) — un revenu, jamais une dépense personnelle.
  // "Voiture" volontairement absente : entretien/réparations de la Mazda mise en
  // location — un coût de l'activité Mazda, pas une dépense personnelle (même
  // logique que GRUNDFOS, déjà exclu). "Prêt" (prêts accordés à des tiers) volontairement
  // absent aussi — géré côté 4-3-2-1 (Investissement) et dans l'onglet Créances.
  // "Payement Dette Orange" ajouté : nouvelle catégorie créée par l'utilisateur pour
  // isoler le remboursement d'un emprunt bancaire (Orange Bank) — une obligation
  // récurrente, traitée comme "Dette" générique tant qu'aucun lien avec un actif
  // rentable n'est établi (contrairement à "Dette", qui finance la Mazda).
  // Orthographe vérifiée le 06/08/2026 contre le Journal réel de l'app (capture d'écran).
  // "Déménagement" ajouté le 07/08/2026 : confirmé par l'utilisateur comme un besoin
  // nécessaire pour se loger correctement — un oubli de ma part jusqu'ici, malgré sa
  // confirmation explicite (12 808 870 FCFA au total, non négligeable).
  survie: ["Logement", "PAYEMENT MAISON", "Enfants & Maman", "Aliments", "Santé", "Transport", "Pack Club", "Payement Dette Orange", "Déménagement"],
  optionnel: ["Utilitaires", "Divertissement", "Shopping", "VACANCE NESHER", "Vêtements", "Personnel", "Voyage", "Abonnements", "Cadeaux", "Invitation"],
  culture: ["Éducation", "FORMATION"],
  extra: ["Ajustement", "Générales", "General"],
};

function computeAsianIndicators(transactions: Transaction[], chargeOverrides: Record<string, ChargeOverride>, includeGrundfosVoiture: boolean, explicitRange?: [string, string]) {
  // Respecte le filtre global "Du mois / Au mois" quand il est fourni ; sinon, tout
  // l'historique depuis la toute première transaction (sur suggestion de l'utilisateur,
  // 06/08/2026) — plus représentatif pour des ratios censés refléter un comportement
  // de fond plutôt qu'un instantané récent.
  const windowMonths = monthsSinceInception(transactions);
  const charges = classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, windowMonths, explicitRange);
  const lookback = charges.lookback;
  const avgRevenu = charges.avgRevenu;

  const sumFor = (cats: string[]) => {
    const total = transactions
      .filter((t) => t.type === "Dépense" && cats.includes(t.category) && lookback.includes(dateToMonthKey(t.date)))
      .reduce((a, t) => a + t.amount, 0);
    return total / lookback.length; // moyenne mensuelle sur la fenêtre
  };

  // --- Règle chinoise du 4-3-2-1 (家庭资产配置法则), enseignée dans les
  // certifications chinoises de planification financière (AFP) : 40% investissement,
  // 30% vie courante, 20% protection/assurance, 10% épargne de précaution liquide.
  // Base commune : la dépense totale réelle sur la fenêtre (pas seulement la part
  // classée "fixe/variable" par ailleurs, sans quoi les achats ponctuels comme un
  // terrain, une maison ou une voiture — souvent "occasionnels" au sens de la
  // classification charges fixes/variables — seraient perdus du calcul).
  const totalDepensesMonthly = transactions
    .filter((t) => t.type === "Dépense" && lookback.includes(dateToMonthKey(t.date)) && (includeGrundfosVoiture || !GRUNDFOS_VOITURE_CATEGORIES.includes(t.category)))
    .reduce((a, t) => a + t.amount, 0) / lookback.length;
  const investMonthly = sumFor(CN_4321_CATEGORIES.investissement);
  const protectionMonthly = sumFor(CN_4321_CATEGORIES.protection);
  // Ni "vie courante" ni "épargne de précaution" ne sont plancherisées à 0 : si les
  // dépenses dépassent le revenu moyen, l'épargne de précaution devient RÉELLEMENT
  // négative (un déficit), et c'est affiché tel quel plutôt que masqué en silence.
  // Grâce à ça, les 4 pourcentages totalisent toujours exactement 100%.
  const vieCouranteMonthly = totalDepensesMonthly - investMonthly - protectionMonthly;
  const liquideMonthly = avgRevenu - totalDepensesMonthly;
  const hasDeficit = liquideMonthly < 0;

  const pct = (v: number) => (avgRevenu > 0 ? (v / avgRevenu) * 100 : 0);
  const rule4321 = [
    { label: "Investissement", value: pct(investMonthly), target: 40 },
    { label: "Vie courante", value: pct(vieCouranteMonthly), target: 30 },
    { label: "Protection / assurance", value: pct(protectionMonthly), target: 20 },
    { label: "Épargne de précaution", value: pct(liquideMonthly), target: 10 },
  ];

  // --- Taux d'épargne norme asiatique : les ménages chinois épargnent
  // traditionnellement 30 à 45% de leur revenu (contre ~20% recommandé en Occident) —
  // une des raisons souvent citées du haut niveau d'épargne des ménages en Chine.
  const tauxEpargne = avgRevenu > 0 ? ((avgRevenu - charges.totalFixe - charges.totalVariable) / avgRevenu) * 100 : 0;

  // --- Kakeibo (家計簿), méthode budgétaire japonaise (Hani Motoko, 1904),
  // toujours largement utilisée en Asie : classer chaque dépense en 4 catégories
  // et se poser 4 questions (combien j'ai / combien je veux épargner / combien je
  // dépense / comment je peux m'améliorer) plutôt que de suivre un simple total.
  const kakeiboTotal = sumFor(KAKEIBO_CATEGORIES.survie) + sumFor(KAKEIBO_CATEGORIES.optionnel) + sumFor(KAKEIBO_CATEGORIES.culture) + sumFor(KAKEIBO_CATEGORIES.extra) || 1;
  const kakeibo = [
    { label: "Survie (nécessités)", key: "survie", value: sumFor(KAKEIBO_CATEGORIES.survie) },
    { label: "Optionnel (envies)", key: "optionnel", value: sumFor(KAKEIBO_CATEGORIES.optionnel) },
    { label: "Culture (développement)", key: "culture", value: sumFor(KAKEIBO_CATEGORIES.culture) },
    { label: "Extra (imprévus)", key: "extra", value: sumFor(KAKEIBO_CATEGORIES.extra) },
  ].map((k) => ({ ...k, pct: (k.value / kakeiboTotal) * 100 }));

  const sumForDetail = (cats: string[]) => {
    const byCat: Record<string, number> = {};
    transactions.forEach((t) => {
      if (t.type === "Dépense" && cats.includes(t.category) && lookback.includes(dateToMonthKey(t.date))) {
        byCat[t.category] = (byCat[t.category] || 0) + t.amount;
      }
    });
    return Object.entries(byCat).map(([category, total]) => ({ category, monthly: total / lookback.length })).sort((a, b) => b.monthly - a.monthly);
  };
  const investDetail = sumForDetail(CN_4321_CATEGORIES.investissement);
  const protectionDetail = sumForDetail(CN_4321_CATEGORIES.protection);

  return { rule4321, tauxEpargne, kakeibo, avgRevenu, hasDeficit, windowMonths, totalDepensesMonthly, investMonthly, protectionMonthly, vieCouranteMonthly, liquideMonthly, investDetail, protectionDetail, lookback };
}

// ============================================================
// RAPPORT NARRATIF DÉTAILLÉ — transforme les chiffres bruts des indicateurs
// asiatiques en un vrai texte explicatif : définition de chaque catégorie,
// exemples concrets, comparaison chiffrée à l'objectif, verdict dynamique,
// puis une synthèse globale avec recommandation. Généré à partir des données
// réelles à chaque ouverture, pas un texte figé.
// ============================================================
interface NarrativeSection { title: string; you: number; target: number; direction: "plus_is_better" | "less_is_better"; definition: string; examples: string[]; verdict: string; }

function verdictFor(value: number, target: number, direction: "plus_is_better" | "less_is_better"): string {
  const gap = direction === "plus_is_better" ? value - target : target - value;
  if (gap >= -3) {
    return direction === "plus_is_better" && value > target + 5
      ? "Tu dépasses largement l'objectif."
      : "Tu es proche de l'objectif — c'est l'une de tes meilleures catégories.";
  }
  if (gap >= -15) return "Tu es en dessous de l'objectif, avec une marge de progression réelle.";
  return "Tu es très loin de l'objectif.";
}

function generateAsian4321Narrative(rule4321: { label: string; value: number; target: number }[]): NarrativeSection[] {
  const defs: Record<string, { definition: string; examples: string[]; direction: "plus_is_better" | "less_is_better" }> = {
    "Investissement": {
      definition: "Cela correspond à l'argent placé pour créer de la richesse à long terme, pas dépensé.",
      examples: ["Actions", "Immobilier", "ETF", "Entreprise", "Terrain", "Obligations", "Parts dans une société"],
      direction: "plus_is_better",
    },
    "Vie courante": {
      definition: "Ce sont les dépenses nécessaires pour vivre au quotidien.",
      examples: ["Nourriture", "Carburant", "Restaurants", "Vêtements", "Sorties", "Téléphone", "Internet", "Abonnements", "Loisirs", "Cadeaux"],
      direction: "less_is_better",
    },
    "Protection / assurance": {
      definition: "Ce sont les dépenses qui protègent ton avenir en cas de coup dur.",
      examples: ["Assurance maladie", "Assurance vie", "Retraite", "Assurance habitation", "Assurance automobile", "Assurance professionnelle"],
      direction: "plus_is_better",
    },
    "Épargne de précaution": {
      definition: "C'est l'argent qui reste disponible en cas d'imprévu.",
      examples: ["Accident", "Maladie", "Perte d'emploi", "Panne de voiture", "Grosse réparation"],
      direction: "plus_is_better",
    },
  };
  return rule4321.map((r) => {
    const d = defs[r.label] || { definition: "", examples: [], direction: "plus_is_better" as const };
    return { title: r.label, you: r.value, target: r.target, direction: d.direction, definition: d.definition, examples: d.examples, verdict: verdictFor(r.value, r.target, d.direction) };
  });
}

function generateKakeiboNarrative(kakeibo: { label: string; key: string; pct: number }[]): NarrativeSection[] {
  const defs: Record<string, { definition: string; examples: string[]; target: number; commentary: (pct: number) => string }> = {
    survie: {
      definition: "Ce sont les dépenses indispensables pour vivre.",
      examples: ["Loyer", "Alimentation", "Eau", "Électricité", "Carburant pour aller travailler", "Transport"],
      target: 50,
      commentary: (pct) => (pct <= 55 ? "C'est raisonnable." : "C'est élevé pour un poste censé être incompressible."),
    },
    optionnel: {
      definition: "Ce sont les envies, pas les besoins.",
      examples: ["Restaurants", "Shopping", "Netflix", "Sorties", "Gadgets", "Voyages loisirs"],
      target: 20,
      commentary: (pct) => (pct >= 30 ? "C'est presque autant que les dépenses essentielles — un vrai levier d'économie si besoin." : "C'est sous contrôle."),
    },
    culture: {
      definition: "Cette catégorie concerne l'investissement dans ton propre développement.",
      examples: ["Livres", "Formations", "MBA", "Cours en ligne", "Conférences", "Apprentissage"],
      target: 10,
      commentary: (pct) => (pct <= 5 ? "Pour quelqu'un qui souhaite progresser professionnellement, c'est assez faible." : "C'est un niveau correct d'investissement sur toi-même."),
    },
    extra: {
      definition: "Cette catégorie regroupe les dépenses exceptionnelles ou non planifiées.",
      examples: ["Réparations", "Cadeaux", "Urgences", "Frais imprévus", "Dépenses non planifiées"],
      target: 15,
      commentary: (pct) => (pct >= 25 ? "C'est élevé — cela signifie probablement beaucoup de dépenses inhabituelles, ou que certaines dépenses sont mal catégorisées dans l'app." : "C'est un niveau normal d'imprévus."),
    },
  };
  return kakeibo.map((k) => {
    const d = defs[k.key];
    return { title: k.label, you: k.pct, target: d.target, direction: "less_is_better" as const, definition: d.definition, examples: d.examples, verdict: d.commentary(k.pct) };
  });
}

function generateFinancialProfileSynthesis(rule4321Narr: NarrativeSection[], tauxEpargne: number, kakeiboNarr: NarrativeSection[]): { checks: { ok: boolean; text: string }[]; recommendation: string } {
  const invest = rule4321Narr.find((r) => r.title === "Investissement");
  const vieCourante = rule4321Narr.find((r) => r.title === "Vie courante");
  const protection = rule4321Narr.find((r) => r.title === "Protection / assurance");
  const extra = kakeiboNarr.find((r) => r.title.startsWith("Extra"));

  const checks: { ok: boolean; text: string }[] = [];
  checks.push({ ok: tauxEpargne >= 20, text: `Capacité d'épargne ${tauxEpargne >= 20 ? "correcte" : "insuffisante"} (${tauxEpargne.toFixed(0)}% selon l'indicateur de l'app).` });
  if (invest) checks.push({ ok: invest.you >= invest.target - 5, text: `Investissement ${invest.you >= invest.target - 5 ? "à un niveau sain" : `très faible (${invest.you.toFixed(0)}% contre un objectif de ${invest.target}%)`}.` });
  if (vieCourante) checks.push({ ok: vieCourante.you <= vieCourante.target + 10, text: vieCourante.you <= vieCourante.target + 10 ? "Dépenses de vie courante maîtrisées." : `Dépenses de vie courante beaucoup trop élevées (${vieCourante.you.toFixed(0)}% contre ${vieCourante.target}%).` });
  if (protection) checks.push({ ok: protection.you >= protection.target - 8, text: protection.you >= protection.target - 8 ? "Protection financière correcte." : `Très peu de protection financière (${protection.you.toFixed(0)}%).` });
  if (extra) checks.push({ ok: extra.you < 20, text: extra.you >= 20 ? `Les dépenses imprévues représentent une part importante (${extra.you.toFixed(0)}%), signe possible d'un manque de planification.` : "Dépenses imprévues sous contrôle." });

  const worst = [invest, vieCourante, protection].filter((s): s is NarrativeSection => !!s)
    .sort((a, b) => {
      const gapA = a.direction === "plus_is_better" ? a.target - a.you : a.you - a.target;
      const gapB = b.direction === "plus_is_better" ? b.target - b.you : b.you - b.target;
      return gapB - gapA;
    })[0];

  const recommendation = worst
    ? `Si tu souhaites utiliser cette analyse pour améliorer ta situation, l'objectif principal serait de ${worst.title === "Vie courante" ? "réduire progressivement les dépenses de vie courante et les imprévus" : `renforcer "${worst.title}"`} afin de libérer davantage de ressources pour l'investissement — la catégorie qui a le plus d'impact sur la création de patrimoine à long terme.`
    : "Ta répartition est globalement équilibrée par rapport aux références utilisées ici.";

  return { checks, recommendation };
}

// Détecte les dépenses périodiques (loyer, retraite, PEL...) qui reviennent la
// plupart des mois, souvent en fin de mois, et qui ne sont pas encore passées ce
// mois-ci — pour que le conseiller les anticipe plutôt que de les ignorer.
function detectRecurringExpenses(transactions: Transaction[], curMonth: string, dayNum: number, chargeOverrides: Record<string, ChargeOverride>, includeGrundfosVoiture: boolean) {
  const lookback: string[] = [];
  let mk = prevMonthKey(curMonth);
  for (let i = 0; i < 6; i++) { lookback.push(mk); mk = prevMonthKey(mk); }

  // Suivi au niveau "poste" (catégorie::sous-catégorie pour Abonnements, GRUNDFOS...) —
  // pas juste la catégorie entière. Sur demande explicite de l'utilisateur (08/08/2026) :
  // regrouper toute la catégorie masquait le fait qu'un abonnement précis (ex: Spotify)
  // n'était pas encore payé ce mois-ci simplement parce qu'un autre (ex: Money Coach)
  // l'était déjà — chaque charge fixe doit être suivie et rappelée individuellement.
  const byPosteMonth: Record<string, Record<string, { amount: number; day: number }>> = {};
  transactions.forEach((t) => {
    if (t.type !== "Dépense") return;
    const tmk = dateToMonthKey(t.date);
    if (!lookback.includes(tmk)) return;
    const day = new Date(t.date + "T00:00:00").getDate();
    const poste = EXPAND_SUBCATS_FOR_CHARGES[t.category] ? `${t.category}::${t.subcategory || "(non précisé)"}` : t.category;
    byPosteMonth[poste] = byPosteMonth[poste] || {};
    if (!byPosteMonth[poste][tmk]) byPosteMonth[poste][tmk] = { amount: 0, day };
    byPosteMonth[poste][tmk].amount += t.amount;
    byPosteMonth[poste][tmk].day = Math.min(byPosteMonth[poste][tmk].day, day);
  });

  // Ne retient que les postes réellement classés "Fixe" dans Charges Fixes &
  // Variables (sur demande explicite de l'utilisateur, 07/08/2026) — la seule régularité
  // statistique de présence (≥4/6 mois) ne suffit pas : Vêtements, Santé ou Ajustement
  // reviennent souvent sans être des charges sûres et incompressibles.
  const windowMonths = monthsSinceInception(transactions);
  const classified = classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, windowMonths);
  const fixedPostes = new Set(classified.rows.filter((r) => r.mode === "fixe").map((r) => r.poste));

  const results: { category: string; monthsPresent: number; typicalDay: number; typicalAmount: number }[] = [];
  Object.entries(byPosteMonth).forEach(([poste, monthsData]) => {
    if (!fixedPostes.has(poste)) return;
    const presentMonths = Object.keys(monthsData);
    if (presentMonths.length < 4) return; // pas assez régulier pour être qualifié de "périodique"
    const amounts = presentMonths.map((m) => monthsData[m].amount);
    const days = presentMonths.map((m) => monthsData[m].day);
    const typicalDay = Math.round(median(days));
    const typicalAmount = median(amounts);
    const [cat, sub] = poste.split("::");
    const alreadyThisMonth = transactions.some((t) => t.type === "Dépense" && dateToMonthKey(t.date) === curMonth && t.category === cat && (sub === undefined || (t.subcategory || "(non précisé)") === sub));
    if (!alreadyThisMonth && dayNum >= typicalDay - 4) {
      results.push({ category: poste.replace("::", " · "), monthsPresent: presentMonths.length, typicalDay, typicalAmount });
    }
  });
  return results.sort((a, b) => b.typicalAmount - a.typicalAmount).slice(0, 4);
}

// Repère, sous-catégorie par sous-catégorie, où le comportement du mois en cours
// dévie nettement de la moyenne des 3 derniers mois — pour nommer concrètement
// ce qui dérape (ex: "Divertissement · Alcool") plutôt que de rester générique.
function analyzeSubcategoryDrift(transactions: Transaction[], curMonth: string, dayNum: number, daysInMonth: number, chargeOverrides?: Record<string, ChargeOverride>, includeGrundfosVoiture: boolean = true) {
  const watched = ["Divertissement", "Cadeaux", "Shopping", "Invitation", "Vêtements", "Abonnements", "Voyage", "Personnel"];
  const lookback: string[] = [];
  let mk = prevMonthKey(curMonth);
  for (let i = 0; i < 3; i++) { lookback.push(mk); mk = prevMonthKey(mk); }

  // Exclut les postes classés "Fixe" dans Charges Fixes & Variables — sur demande
  // explicite de l'utilisateur (08/08/2026) : projeter une dépense en "rythme
  // quotidien × jours du mois" n'a aucun sens pour un abonnement payé une seule fois
  // par mois (Spotify, Claude, GRUNDFOS·Internet...). Ces postes-là sont déjà couverts
  // par la détection de charges périodiques (upcoming), pas par le dérapage journalier.
  const fixedPostes = new Set<string>();
  if (chargeOverrides) {
    const windowMonths = monthsSinceInception(transactions);
    const classified = classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, windowMonths);
    classified.rows.forEach((r) => { if (r.mode === "fixe") fixedPostes.add(r.poste); });
  }
  const isFixed = (cat: string, sub: string) => fixedPostes.has(cat) || fixedPostes.has(`${cat}::${sub}`);

  const results: { category: string; subcategory: string; thisMonth: number; projected: number; avgPast: number; diffPct: number }[] = [];
  watched.forEach((cat) => {
    const thisMonthTx = transactions.filter((t) => t.type === "Dépense" && t.category === cat && dateToMonthKey(t.date) === curMonth);
    const subs = Array.from(new Set(thisMonthTx.map((t) => t.subcategory || "(non précisé)")));
    subs.forEach((sub) => {
      if (isFixed(cat, sub)) return;
      const thisAmt = thisMonthTx.filter((t) => (t.subcategory || "(non précisé)") === sub).reduce((a, t) => a + t.amount, 0);
      const pastAmts = lookback.map((m) =>
        transactions.filter((t) => t.type === "Dépense" && t.category === cat && (t.subcategory || "(non précisé)") === sub && dateToMonthKey(t.date) === m).reduce((a, t) => a + t.amount, 0)
      ).filter((a) => a > 0);
      if (!pastAmts.length) return; // pas d'historique comparable
      const avgPast = mean(pastAmts);
      const projected = dayNum > 0 ? (thisAmt / dayNum) * daysInMonth : thisAmt;
      if (avgPast <= 0) return;
      const diffPct = ((projected - avgPast) / avgPast) * 100;
      if (Math.abs(diffPct) >= 40 && Math.abs(projected - avgPast) >= 3000) {
        results.push({ category: cat, subcategory: sub, thisMonth: thisAmt, projected, avgPast, diffPct });
      }
    });
  });
  return results.sort((a, b) => Math.abs(b.diffPct) - Math.abs(a.diffPct)).slice(0, 3);
}

function generateDailyAdvice(transactions: Transaction[], monthlyObjective: number, chargeOverrides?: Record<string, ChargeOverride>, includeGrundfosVoiture: boolean = true) {
  const today = todayISO();
  const curMonth = dateToMonthKey(today);
  const dayNum = new Date(today + "T00:00:00").getDate();
  const daysInMonth = daysInMonthOf(curMonth);
  const daysElapsed = dayNum;
  const daysRemaining = daysInMonth - dayNum + 1;

  const txThisMonthSoFar = transactions.filter((t) => dateToMonthKey(t.date) === curMonth && t.date <= today);
  const spent = txThisMonthSoFar.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const revenu = txThisMonthSoFar.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);

  const avgDailySpend = daysElapsed > 0 ? spent / daysElapsed : 0;
  const projectedEndOfMonth = avgDailySpend * daysInMonth;

  const yDate = new Date(today + "T00:00:00"); yDate.setDate(yDate.getDate() - 1);
  const yesterday = `${yDate.getFullYear()}-${pad2(yDate.getMonth() + 1)}-${pad2(yDate.getDate())}`;
  const spentYesterday = transactions.filter((t) => t.date === yesterday && t.type === "Dépense").reduce((a, t) => a + t.amount, 0);

  const pMonth = prevMonthKey(curMonth);
  const spentSamePeriodLastMonth = transactions.filter((t) => t.type === "Dépense" && dateToMonthKey(t.date) === pMonth && new Date(t.date + "T00:00:00").getDate() <= dayNum).reduce((a, t) => a + t.amount, 0);

  // Dépenses périodiques probablement encore à venir (loyer, retraite, PEL...) — anticipées
  // avant de calculer un vrai budget quotidien, sans quoi le seuil serait trompeur.
  const upcoming = detectRecurringExpenses(transactions, curMonth, dayNum, chargeOverrides || {}, includeGrundfosVoiture);
  const upcomingTotal = upcoming.reduce((a, u) => a + u.typicalAmount, 0);

  const hasObjective = monthlyObjective > 0;
  const remainingBudget = hasObjective ? monthlyObjective - spent : null;
  const remainingAfterUpcoming = hasObjective ? (remainingBudget as number) - upcomingTotal : null;
  const dailyThreshold = hasObjective && daysRemaining > 0 ? Math.max(0, ((remainingAfterUpcoming ?? remainingBudget) as number) / daysRemaining) : null;
  const onTrack = hasObjective ? projectedEndOfMonth <= monthlyObjective : null;

  const insights: Insight[] = [];

  // Le conseil principal : le seuil concret pour aujourd'hui, déjà net des charges périodiques attendues.
  let headline: Insight;
  if (hasObjective) {
    if ((remainingBudget as number) <= 0) {
      headline = { kind: "alerte", title: "Objectif du mois déjà dépassé", text: `Tu as dépensé ${fmt(spent)} FCFA depuis le début du mois, au-delà de ton objectif de ${fmt(monthlyObjective)} FCFA. Il reste ${daysRemaining} jour(s) — vise le zéro dépense non essentielle jusqu'à la fin du mois.` };
    } else {
      headline = {
        kind: onTrack && (remainingAfterUpcoming as number) > 0 ? "positif" : "conseil",
        title: `Seuil du jour : ${fmt(Math.max(0, dailyThreshold as number))} FCFA`,
        text: upcomingTotal > 0
          ? `Il te reste ${fmt(remainingBudget as number)} FCFA sur ton objectif, mais ${fmt(upcomingTotal)} FCFA sont probablement encore à sortir ce mois-ci en charges périodiques (détail ci-dessous). Une fois cette réserve mise de côté, il reste environ ${fmt(Math.max(0, remainingAfterUpcoming as number))} FCFA de marge réelle sur ${daysRemaining} jour(s), soit ${fmt(Math.max(0, dailyThreshold as number))} FCFA/jour.`
          : `Il te reste ${fmt(remainingBudget as number)} FCFA sur ton objectif de ${fmt(monthlyObjective)} FCFA pour les ${daysRemaining} jour(s) restants — soit environ ${fmt(dailyThreshold as number)} FCFA/jour maximum si tu veux rester dans les clous jusqu'à la fin du mois.`,
      };
    }
  } else {
    headline = { kind: "conseil", title: "Définis un objectif mensuel pour des conseils plus précis", text: `Sans objectif, je peux seulement observer ta tendance : à ${fmt(avgDailySpend)} FCFA/jour en moyenne depuis le début du mois, tu termines vers ${fmt(projectedEndOfMonth)} FCFA de dépenses ce mois-ci si le rythme se maintient.` };
  }
  insights.push(headline);

  // Charges périodiques attendues, nommées une par une (loyer, retraite, PEL...).
  if (upcoming.length) {
    const detail = upcoming.map((u) => `${u.category} (~${fmt(u.typicalAmount)} FCFA, généralement autour du ${u.typicalDay})`).join(", ");
    insights.push({
      kind: "conseil",
      title: `${upcoming.length} charge(s) périodique(s) probablement encore à venir`,
      text: `Sur tes 6 derniers mois, ces postes reviennent presque systématiquement et n'ont rien enregistré ce mois-ci : ${detail}. Historiquement, ce type de dépense arrive plutôt en fin de mois — mieux vaut réserver ces ${fmt(upcomingTotal)} FCFA maintenant que les découvrir le 30.`,
    });
  }

  // Analyse comportementale par sous-catégorie : ce qui dérape vraiment, nommément.
  const drift = analyzeSubcategoryDrift(transactions, curMonth, dayNum, daysInMonth, chargeOverrides, includeGrundfosVoiture);
  drift.forEach((d) => {
    insights.push({
      kind: d.diffPct > 0 ? "alerte" : "positif",
      title: `"${d.category} · ${d.subcategory}" ${d.diffPct > 0 ? "en dérapage" : "nettement maîtrisé"} ce mois-ci`,
      text: d.diffPct > 0
        ? `${fmt(d.thisMonth)} FCFA déjà dépensés sur "${d.subcategory}" en ${daysElapsed} jour(s) — au rythme actuel, ça projette à ${fmt(d.projected)} FCFA sur le mois complet, contre ${fmt(d.avgPast)} FCFA en moyenne les 3 derniers mois (+${d.diffPct.toFixed(0)}%). C'est ce type de ligne, précisément, qui grignote ta capacité d'épargne en fin de mois.`
        : `${fmt(d.thisMonth)} FCFA sur "${d.subcategory}" jusqu'ici, en rythme sur le mois ça donnerait ${fmt(d.projected)} FCFA — nettement sous ta moyenne de ${fmt(d.avgPast)} FCFA (${d.diffPct.toFixed(0)}%). Un vrai progrès sur cette ligne précise.`,
    });
  });

  // Comparaison avec la même période le mois dernier.
  if (spentSamePeriodLastMonth > 0) {
    const diffPct = ((spent - spentSamePeriodLastMonth) / spentSamePeriodLastMonth) * 100;
    if (Math.abs(diffPct) >= 10) {
      insights.push({
        kind: diffPct <= 0 ? "positif" : "alerte",
        title: `${diffPct <= 0 ? "Mieux" : "Moins bien"} que le mois dernier à la même date`,
        text: `Sur les ${daysElapsed} premiers jours, tu as dépensé ${fmt(spent)} FCFA — contre ${fmt(spentSamePeriodLastMonth)} FCFA sur la même période le mois dernier (${diffPct >= 0 ? "+" : ""}${diffPct.toFixed(0)}%).`,
      });
    }
  }

  // Comparaison avec hier.
  if (spentYesterday > 0 && avgDailySpend > 0) {
    const vsAvgPct = ((spentYesterday - avgDailySpend) / avgDailySpend) * 100;
    if (vsAvgPct >= 30) {
      insights.push({ kind: "alerte", title: "Hier, dépense au-dessus de ta moyenne", text: `${fmt(spentYesterday)} FCFA dépensés hier, contre ${fmt(avgDailySpend)} FCFA/jour en moyenne ce mois-ci (+${vsAvgPct.toFixed(0)}%). Un jour plus calme aujourd'hui aiderait à rééquilibrer.` });
    }
  }

  // Projection de fin de mois vs objectif.
  if (hasObjective && daysElapsed >= 3) {
    const projDiff = projectedEndOfMonth - monthlyObjective;
    if (projDiff > 0) {
      insights.push({ kind: "alerte", title: "Au rythme actuel, l'objectif sera dépassé", text: `En prolongeant ta moyenne quotidienne (${fmt(avgDailySpend)} FCFA/jour), tu finirais le mois à environ ${fmt(projectedEndOfMonth)} FCFA — soit ${fmt(projDiff)} FCFA au-dessus de ton objectif.` });
    } else if (daysElapsed >= 5) {
      insights.push({ kind: "positif", title: "Au rythme actuel, l'objectif est tenable", text: `En continuant sur ta lancée (${fmt(avgDailySpend)} FCFA/jour), tu terminerais le mois autour de ${fmt(projectedEndOfMonth)} FCFA — sous ton objectif de ${fmt(monthlyObjective)} FCFA. Continue ainsi.` });
    }
  }

  // Épargne du mois en cours.
  if (revenu > 0) {
    const soldeMois = revenu - spent;
    const tauxEpargne = (soldeMois / revenu) * 100;
    if (tauxEpargne < 10 && daysElapsed >= 10) {
      insights.push({ kind: "conseil", title: "Taux d'épargne encore faible ce mois-ci", text: `${tauxEpargne.toFixed(0)}% des revenus encaissés ce mois-ci sont conservés pour l'instant. Réduire les dépenses non essentielles dans les jours qui restent ferait grimper ce chiffre et ta valeur nette en fin de mois.` });
    }
  }

  // Reste à vivre, en s'appuyant sur la classification charges fixes / variables régulières
  // (page "Charges Fixes & Variables") plutôt que sur une simple moyenne de dépenses.
  if (chargeOverrides) {
    const charges = classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture);
    if (charges.avgRevenu > 0) {
      insights.push({
        kind: charges.resteAVivre >= 0 ? "positif" : "alerte",
        title: charges.resteAVivre >= 0 ? "Reste à vivre positif une fois les charges couvertes" : "Reste à vivre négatif une fois les charges couvertes",
        text: `Sur un revenu moyen de ${fmt(charges.avgRevenu)} FCFA/mois, une fois les charges fixes (${fmt(charges.totalFixe)} FCFA) et les variables régulières estimées (${fmt(charges.totalVariable)} FCFA) déduites, il reste environ ${fmt(charges.resteAVivre)} FCFA/mois de marge réellement discrétionnaire. ${charges.resteAVivre < 0 ? "Les charges dépassent le revenu moyen — à regarder de près sur la page Charges Fixes & Variables." : "C'est cette marge, pas le revenu brut, qui détermine ce que tu peux vraiment te permettre au jour le jour."}`,
      });
    }
  }

  return { insights, spent, revenu, daysElapsed, daysRemaining, daysInMonth, avgDailySpend, projectedEndOfMonth, dailyThreshold, remainingBudget, hasObjective, upcomingTotal };
}

// Note du jour : Excellent / Bon / Moyen / Faible — combine trois signaux concrets :
// le respect du seuil quotidien, la dépense du jour comparée à l'habitude, et l'absence
// de dérapage sur une sous-catégorie déjà repérée comme en dérive ce mois-ci.
function computeDayScore(transactions: Transaction[], monthlyObjective: number) {
  const today = todayISO();
  const curMonth = dateToMonthKey(today);
  const dayNum = new Date(today + "T00:00:00").getDate();
  const daysInMonth = daysInMonthOf(curMonth);

  const todayTx = transactions.filter((t) => t.date === today);
  const todayDep = todayTx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const todayRev = todayTx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const todaySolde = todayRev - todayDep;

  // Dépense quotidienne "habituelle" : médiane des 30 jours précédents (jours à 0 inclus,
  // pour refléter le vrai rythme, pas seulement les jours où il y a eu une dépense).
  const dailyDepenses: Record<string, number> = {};
  transactions.forEach((t) => { if (t.type === "Dépense" && t.date < today) dailyDepenses[t.date] = (dailyDepenses[t.date] || 0) + t.amount; });
  const last30: number[] = [];
  const cursor = new Date(today + "T00:00:00");
  for (let i = 1; i <= 30; i++) {
    cursor.setDate(cursor.getDate() - 1);
    const iso = `${cursor.getFullYear()}-${pad2(cursor.getMonth() + 1)}-${pad2(cursor.getDate())}`;
    last30.push(dailyDepenses[iso] || 0);
  }
  const avgHistDaily = median(last30);

  const advice = generateDailyAdvice(transactions, monthlyObjective);
  const dailyThreshold = advice.dailyThreshold;

  // A. Respect du seuil quotidien (si un objectif mensuel est défini)
  let seuilScore = 60;
  if (dailyThreshold !== null) {
    seuilScore = todayDep <= dailyThreshold ? 100 : Math.max(0, 100 - ((todayDep - dailyThreshold) / Math.max(dailyThreshold, 1)) * 100);
  }

  // B. Dépense du jour comparée à l'habitude (médiane des 30 derniers jours)
  let habitScore = 60;
  if (avgHistDaily > 0) {
    const ratio = todayDep / avgHistDaily;
    habitScore = Math.max(0, Math.min(100, 100 - (ratio - 1) * 60));
  } else if (todayDep === 0) {
    habitScore = 100;
  }

  // C. Absence de dérapage sur une sous-catégorie déjà repérée en dérive ce mois-ci
  const drift = analyzeSubcategoryDrift(transactions, curMonth, dayNum, daysInMonth);
  const driftedKeys = new Set(drift.filter((d) => d.diffPct > 0).map((d) => `${d.category}::${d.subcategory}`));
  const touchesDrift = todayTx.some((t) => t.type === "Dépense" && driftedKeys.has(`${t.category}::${t.subcategory || "(non précisé)"}`));
  const driftScore = drift.length ? (touchesDrift ? 25 : 100) : 70;

  const overall = (seuilScore + habitScore + driftScore) / 3;
  let grade = "Faible", gradeColor = COLOR.clay;
  if (overall >= 80) { grade = "Excellent"; gradeColor = COLOR.emerald; }
  else if (overall >= 60) { grade = "Bon"; gradeColor = COLOR.emeraldSoft; }
  else if (overall >= 40) { grade = "Moyen"; gradeColor = COLOR.gold; }

  return { overall, grade, gradeColor, seuilScore, habitScore, driftScore, todayDep, todayRev, todaySolde, avgHistDaily, dailyThreshold, touchesDrift, hasDriftData: drift.length > 0 };
}

// Note du mois : même logique que la note du jour, mais à l'échelle du mois en cours —
// rythme de dépense vs objectif, comparaison à la même période le mois dernier, épargne.
function computeMonthScore(transactions: Transaction[], monthlyObjective: number) {
  const today = todayISO();
  const curMonth = dateToMonthKey(today);
  const dayNum = new Date(today + "T00:00:00").getDate();
  const daysInMonth = daysInMonthOf(curMonth);
  const timeRatio = dayNum / daysInMonth;

  const txThisMonth = transactions.filter((t) => dateToMonthKey(t.date) === curMonth && t.date <= today);
  const spent = txThisMonth.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const revenu = txThisMonth.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const solde = revenu - spent;

  const pMonth = prevMonthKey(curMonth);
  const spentSamePeriodLastMonth = transactions.filter((t) => t.type === "Dépense" && dateToMonthKey(t.date) === pMonth && new Date(t.date + "T00:00:00").getDate() <= dayNum).reduce((a, t) => a + t.amount, 0);

  // A. Rythme de dépense vs objectif (compare la part du mois écoulée à la part du budget déjà consommée)
  let paceScore = 60;
  const hasObjective = monthlyObjective > 0;
  if (hasObjective) {
    const spendRatio = spent / monthlyObjective;
    paceScore = spendRatio <= timeRatio ? 100 : Math.max(0, 100 - ((spendRatio - timeRatio) / Math.max(timeRatio, 0.05)) * 100);
  }

  // B. Comparaison à la même période le mois dernier
  let vsLastMonthScore = 60;
  if (spentSamePeriodLastMonth > 0) {
    const ratio = spent / spentSamePeriodLastMonth;
    vsLastMonthScore = Math.max(0, Math.min(100, 100 - (ratio - 1) * 60));
  }

  // C. Taux d'épargne du mois en cours (20% = référence "excellent", comme le score de santé global)
  let savingsScore = 50;
  if (revenu > 0) {
    const tauxEpargne = (solde / revenu) * 100;
    savingsScore = Math.max(0, Math.min(100, (tauxEpargne / 20) * 100));
  }

  const overall = (paceScore + vsLastMonthScore + savingsScore) / 3;
  let grade = "Faible", gradeColor = COLOR.clay;
  if (overall >= 80) { grade = "Excellent"; gradeColor = COLOR.emerald; }
  else if (overall >= 60) { grade = "Bon"; gradeColor = COLOR.emeraldSoft; }
  else if (overall >= 40) { grade = "Moyen"; gradeColor = COLOR.gold; }

  return { overall, grade, gradeColor, paceScore, vsLastMonthScore, savingsScore, spent, revenu, solde, spentSamePeriodLastMonth, dayNum, daysInMonth, hasObjective };
}

function DayScoreBadge({ transactions, monthlyObjective, compact, scope = "jour" }: { transactions: Transaction[]; monthlyObjective: number; compact?: boolean; scope?: "jour" | "mois" }) {
  const dayScore = useMemo(() => computeDayScore(transactions, monthlyObjective), [transactions, monthlyObjective]);
  const monthScore = useMemo(() => computeMonthScore(transactions, monthlyObjective), [transactions, monthlyObjective]);
  const score = scope === "jour" ? dayScore : monthScore;
  const sub = scope === "jour"
    ? [
        { label: "Seuil du jour respecté", value: dayScore.seuilScore },
        { label: "Vs. dépense habituelle", value: dayScore.habitScore },
        { label: "Sous-catégories surveillées", value: dayScore.driftScore },
      ]
    : [
        { label: "Rythme vs objectif mensuel", value: monthScore.paceScore },
        { label: "Vs. même période le mois dernier", value: monthScore.vsLastMonthScore },
        { label: "Taux d'épargne du mois", value: monthScore.savingsScore },
      ];
  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 16, background: `linear-gradient(135deg, ${score.gradeColor}22 0%, ${COLOR.surfaceRaised} 65%)`,
      border: `1px solid ${score.gradeColor}`, borderRadius: 14, padding: compact ? "12px 16px" : "16px 20px", flex: 1, minWidth: 260,
    }}>
      <div style={{ width: compact ? 52 : 64, height: compact ? 52 : 64, borderRadius: "50%", border: `3px solid ${score.gradeColor}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: compact ? 14 : 17, fontWeight: 700, color: score.gradeColor }}>{Math.round(score.overall)}</span>
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
          <span style={{ fontFamily: "'Fraunces', serif", fontSize: compact ? 15 : 17, fontWeight: 600, color: score.gradeColor }}>
            {scope === "jour" ? "Journée" : "Mois"} {score.grade}
          </span>
          <span style={{ fontSize: 11.5, color: COLOR.inkMuted }}>{scope === "jour" ? dateLabelFull(todayISO()) : monthLabel(dateToMonthKey(todayISO()))}</span>
        </div>
        {!compact && (
          <div style={{ display: "flex", flexDirection: "column", gap: 4, marginTop: 8 }}>
            {sub.map((s) => (
              <div key={s.label} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 11, color: COLOR.inkMuted, width: 190, flexShrink: 0 }}>{s.label}</span>
                <div style={{ flex: 1, height: 5, background: COLOR.hairline, borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${s.value}%`, background: score.gradeColor }} />
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function DailyAdvisorButton({ transactions, monthlyObjective, setMonthlyObjective, chargeOverrides, includeGrundfosVoiture }: {
  transactions: Transaction[]; monthlyObjective: number; setMonthlyObjective: (n: number) => void; chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [editingObjective, setEditingObjective] = useState(false);
  const [draftObjective, setDraftObjective] = useState(monthlyObjective);
  const data = useMemo(() => generateDailyAdvice(transactions, monthlyObjective, chargeOverrides, includeGrundfosVoiture), [transactions, monthlyObjective, chargeOverrides, includeGrundfosVoiture]);
  const styleFor: Record<InsightKind, { bg: string; border: string; color: string; icon: any }> = {
    alerte: { bg: "rgba(193,84,63,0.08)", border: COLOR.clay, color: COLOR.claySoft, icon: AlertTriangle },
    conseil: { bg: "rgba(201,162,39,0.08)", border: COLOR.gold, color: COLOR.goldSoft, icon: Info },
    positif: { bg: "rgba(63,156,122,0.08)", border: COLOR.emerald, color: COLOR.emeraldSoft, icon: Check },
  };
  const progressPct = data.hasObjective ? Math.min(100, (data.spent / monthlyObjective) * 100) : 0;

  return (
    <>
      <button onClick={() => setOpen(true)} style={{
        display: "flex", alignItems: "center", gap: 14, width: "100%", textAlign: "left", cursor: "pointer",
        background: `linear-gradient(135deg, rgba(201,162,39,0.14) 0%, ${COLOR.surfaceRaised} 60%)`, border: `1px solid ${COLOR.gold}`,
        borderRadius: 14, padding: "16px 18px", marginBottom: 20,
      }}>
        <div style={{ width: 40, height: 40, borderRadius: 10, background: "rgba(201,162,39,0.18)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Compass size={19} color={COLOR.goldSoft} />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 14.5, fontWeight: 600, color: COLOR.ink, fontFamily: "'Fraunces', serif" }}>Conseiller quotidien</div>
          <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 2 }}>
            {data.hasObjective ? `Seuil du jour : ${fmt(data.dailyThreshold ?? 0)} FCFA` : "Définis un objectif mensuel pour un conseil personnalisé"}
          </div>
        </div>
        <ChevronRight size={16} color={COLOR.inkMuted} style={{ flexShrink: 0 }} />
      </button>

      {open && (
        <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 460, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{
            width: "100%", maxWidth: 560, maxHeight: "88vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
            display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "18px 20px", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <div>
                <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, color: COLOR.ink, display: "flex", alignItems: "center", gap: 8 }}>
                  <Compass size={17} color={COLOR.goldSoft} /> Conseiller quotidien
                </div>
                <div style={{ fontSize: 12, color: COLOR.inkMuted, marginTop: 4 }}>{dateLabelFull(todayISO())} · jour {data.daysElapsed} sur {data.daysInMonth}</div>
              </div>
              <button onClick={() => setOpen(false)} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}><X size={18} /></button>
            </div>

            <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: 18, WebkitOverflowScrolling: "touch" }}>

              <div style={{ marginBottom: 14, display: "flex", gap: 10, flexWrap: "wrap" }}>
                <DayScoreBadge transactions={transactions} monthlyObjective={monthlyObjective} compact scope="jour" />
                <DayScoreBadge transactions={transactions} monthlyObjective={monthlyObjective} compact scope="mois" />
              </div>

              <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: 11.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Objectif de dépenses ce mois</span>
                  <button onClick={() => { setDraftObjective(monthlyObjective); setEditingObjective((v) => !v); }} style={{ background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", display: "flex", alignItems: "center", gap: 4, fontSize: 11.5 }}>
                    <Pencil size={11} /> {editingObjective ? "Fermer" : "Modifier"}
                  </button>
                </div>
                {editingObjective ? (
                  <div style={{ display: "flex", gap: 8 }}>
                    <input type="number" inputMode="numeric" style={{ ...inputStyle, flex: 1 }} value={draftObjective} onChange={(e) => setDraftObjective(Number(e.target.value) || 0)} />
                    <button onClick={() => { setMonthlyObjective(draftObjective); setEditingObjective(false); }} style={{ background: COLOR.gold, border: "none", borderRadius: 6, color: COLOR.bg, padding: "0 16px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>Valider</button>
                  </div>
                ) : (
                  <>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 18, fontWeight: 700, color: COLOR.ink }}>
                      {data.hasObjective ? `${fmt(data.spent)} / ${fmt(monthlyObjective)} FCFA` : "Aucun objectif défini"}
                    </div>
                    {data.hasObjective && (
                      <div style={{ height: 7, background: COLOR.hairline, borderRadius: 4, overflow: "hidden", marginTop: 8 }}>
                        <div style={{ height: "100%", width: `${progressPct}%`, background: progressPct > 100 ? COLOR.clay : COLOR.gold }} />
                      </div>
                    )}
                  </>
                )}
              </div>

              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {data.insights.map((ins, i) => {
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
            </div>
          </div>
        </div>
      )}
    </>
  );
}


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
// Fiche de lecture pour le rapport narratif du Score de santé financière.
function HealthScoreNarrativeSheet({ open, onClose, tauxEpargne, pctNonProd, cv, health }: {
  open: boolean; onClose: () => void; tauxEpargne: number; pctNonProd: number; cv: number;
  health: { savingsScore: number; nonProdScore: number; stabilityScore: number; overall: number; grade: string; gradeColor: string };
}) {
  if (!open) return null;
  const report = generateHealthScoreNarrative(tauxEpargne, pctNonProd, cv, health);

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 480, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "100%", maxWidth: 620, maxHeight: "90vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
        display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "18px 22px", borderBottom: `1px solid ${COLOR.hairline}` }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, color: COLOR.ink }}>Rapport détaillé — Score de santé financière</div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}><X size={18} /></button>
        </div>
        <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: "20px 22px", WebkitOverflowScrolling: "touch" }}>
          <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 20 }}>
            Ton score global est de <strong style={{ color: health.gradeColor }}>{health.overall.toFixed(0)}/100 ({health.grade})</strong>, calculé comme la moyenne de 3 composantes. Voici ce que chacune signifie concrètement.
          </p>
          {report.sections.map((s) => {
            const color = s.score >= 70 ? COLOR.emeraldSoft : s.score >= 40 ? COLOR.goldSoft : COLOR.claySoft;
            return (
              <div key={s.key} style={{ marginBottom: 22 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4, flexWrap: "wrap", gap: 6 }}>
                  <span style={{ fontFamily: "'Fraunces', serif", fontSize: 15.5, color: COLOR.ink }}>{s.title}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color }}>{s.you.toFixed(0)}{s.unit} · score {s.score.toFixed(0)}/100</span>
                </div>
                <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, margin: "4px 0 8px 0" }}>{s.definition}</p>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                  {s.examples.map((ex) => (
                    <span key={ex} style={{ fontSize: 10.5, color: COLOR.inkMuted, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "3px 10px" }}>{ex}</span>
                  ))}
                </div>
                <div style={{ fontSize: 12.5, color, fontWeight: 600, display: "flex", alignItems: "flex-start", gap: 6 }}>
                  <ArrowRight size={13} style={{ flexShrink: 0, marginTop: 2 }} /> {s.verdict}
                </div>
              </div>
            );
          })}
          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, paddingTop: 18, marginTop: 4 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: COLOR.ink, marginBottom: 10 }}>Ce qui ferait le plus progresser ton score</div>
            <div style={{ padding: "12px 14px", background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
              {report.recommendation}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ApercuTab({ filtered, filters, accounts, transactions, categoryGroups, chargeOverrides, includeGrundfosVoiture, monthlyObjective }: {
  filtered: any[]; filters: Filters; accounts: Account[]; transactions: Transaction[]; categoryGroups: Record<string, Group>;
  chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean; monthlyObjective: number;
}) {
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
  const healthCv = mean(monthlyRevenues) > 0 ? stdev(monthlyRevenues) / mean(monthlyRevenues) : 1;
  const [healthNarrativeOpen, setHealthNarrativeOpen] = useState(false);

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

  // Rapport mois — synthèse complète du mois choisi (revenus, dépenses, nature, écarts vs
  // le mois précédent), indépendante des filtres globaux, sur demande explicite de
  // l'utilisateur (12/08/2026). Navigable (mois précédent/suivant), jamais au-delà du mois en cours.
  const [monthlyReportOpen, setMonthlyReportOpen] = useState(false);
  const [reportAnchor, setReportAnchor] = useState(() => dateToMonthKey(todayISO()));

  const buildMonthlyReport = (anchorMonthKey: string) => {
    const prevKey = prevMonthKey(anchorMonthKey);
    const withGroupFull = transactions.map((t) => ({ ...t, month: dateToMonthKey(t.date), group: t.type === "Revenu" ? "Revenu" : groupFor(t, categoryGroups) }));

    const sumForMonth = (mk: string) => {
      const arr = withGroupFull.filter((t) => t.month === mk);
      const rev = arr.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
      const dep = arr.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
      return { rev, dep, solde: rev - dep, tx: arr };
    };
    const cur = sumForMonth(anchorMonthKey);
    const prev = sumForMonth(prevKey);

    const groupsList = ["Nécessaire", "Productif", "Non-productif", "Non classifié"] as const;
    const byGroup = (arr: any[]) => {
      const byG: Record<string, number> = { "Nécessaire": 0, "Productif": 0, "Non-productif": 0, "Non classifié": 0 };
      arr.filter((t) => t.type === "Dépense").forEach((t) => { byG[t.group] = (byG[t.group] || 0) + t.amount; });
      return byG;
    };
    const curByGroup = byGroup(cur.tx);
    const prevByGroup = byGroup(prev.tx);
    const deltas = groupsList.map((g) => {
      const c = curByGroup[g], p = prevByGroup[g], d = c - p;
      const pct = p > 0 ? (d / p) * 100 : (c > 0 ? 100 : 0);
      return { group: g, cur: c, prev: p, delta: d, pct };
    }).filter((d) => d.cur > 0 || d.prev > 0);

    const depDelta = cur.dep - prev.dep;
    const depPct = prev.dep > 0 ? (depDelta / prev.dep) * 100 : (cur.dep > 0 ? 100 : 0);
    const revDelta = cur.rev - prev.rev;
    const soldeDelta = cur.solde - prev.solde;
    const curEpargne = cur.rev > 0 ? (cur.solde / cur.rev) * 100 : 0;
    const prevEpargne = prev.rev > 0 ? (prev.solde / prev.rev) * 100 : 0;
    const biggestMover = [...deltas].sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta))[0];

    let verdict = "";
    if (Math.abs(depDelta) < 1) {
      verdict = "Dépenses quasi stables par rapport au mois dernier — rien de notable à signaler.";
    } else if (depDelta < 0) {
      if (biggestMover?.group === "Non-productif") verdict = `Baisse fondée : elle est portée principalement par le "Non-productif" (${fmt(Math.abs(biggestMover.delta))} FCFA de moins), le poste le plus facile à maîtriser sans rien sacrifier d'essentiel.`;
      else if (biggestMover?.group === "Nécessaire") verdict = `À surveiller : cette baisse vient surtout du "Nécessaire" (${fmt(Math.abs(biggestMover.delta))} FCFA de moins) — vérifie qu'il ne s'agit pas d'une privation plutôt que d'une vraie économie.`;
      else if (biggestMover?.group === "Productif") verdict = `Baisse portée par le "Productif" (${fmt(Math.abs(biggestMover.delta))} FCFA de moins) — à vérifier que ce n'est pas un investissement simplement décalé dans le temps.`;
      else verdict = "Baisse des dépenses, sans nature clairement dominante.";
    } else {
      if (biggestMover?.group === "Productif") verdict = `Hausse plutôt saine : elle est portée principalement par le "Productif" (${fmt(biggestMover.delta)} FCFA de plus), donc probablement de l'investissement plutôt que du gaspillage.`;
      else if (biggestMover?.group === "Non-productif") verdict = `À surveiller : cette hausse vient surtout du "Non-productif" (${fmt(biggestMover.delta)} FCFA de plus) — c'est le premier poste à réduire si besoin.`;
      else if (biggestMover?.group === "Nécessaire") verdict = `Hausse portée par le "Nécessaire" (${fmt(biggestMover.delta)} FCFA de plus) — vérifie si c'est ponctuel ou un vrai changement de rythme.`;
      else verdict = "Hausse des dépenses, sans nature clairement dominante.";
    }

    return {
      title: `Rapport mois — ${monthLabel(anchorMonthKey)}`,
      headline: `Solde : ${fmt(cur.solde)} FCFA (${soldeDelta >= 0 ? "+" : ""}${fmt(soldeDelta)} FCFA vs ${monthLabel(prevKey)})`,
      formula: `${monthLabel(anchorMonthKey)} vs ${monthLabel(prevKey)} — synthèse revenus, dépenses et nature`,
      blocks: [
        {
          kind: "kv" as const,
          rows: [
            { label: "Revenus", value: `${fmt(cur.rev)} FCFA (${revDelta >= 0 ? "+" : ""}${fmt(revDelta)})` },
            { label: "Dépenses", value: `${fmt(cur.dep)} FCFA (${depDelta >= 0 ? "+" : ""}${fmt(depDelta)})`, warn: depDelta > 0 },
            { label: "Solde", value: `${fmt(cur.solde)} FCFA (${soldeDelta >= 0 ? "+" : ""}${fmt(soldeDelta)})`, strong: true, warn: cur.solde < 0 },
            { label: "Taux d'épargne", value: `${curEpargne.toFixed(1)}% (${curEpargne - prevEpargne >= 0 ? "+" : ""}${(curEpargne - prevEpargne).toFixed(1)} pts)` },
          ],
        },
        {
          kind: "table" as const,
          columns: ["Nature", "Ce mois", "Mois dernier", "Écart", "Évolution"],
          rows: deltas.map((d) => [d.group, fmt(d.cur), fmt(d.prev), `${d.delta >= 0 ? "+" : ""}${fmt(d.delta)}`, `${d.pct >= 0 ? "+" : ""}${d.pct.toFixed(0)}%`]),
          cellColors: deltas.map((d) => [undefined, undefined, undefined, undefined, Math.abs(d.delta) < 1 ? COLOR.inkMuted : d.delta < 0 ? COLOR.emeraldSoft : COLOR.claySoft]),
          footerRow: ["Total", fmt(cur.dep), fmt(prev.dep), `${depDelta >= 0 ? "+" : ""}${fmt(depDelta)}`, `${depPct >= 0 ? "+" : ""}${depPct.toFixed(0)}%`],
          footerColors: [undefined, undefined, undefined, undefined, Math.abs(depDelta) < 1 ? COLOR.inkMuted : depDelta < 0 ? COLOR.emeraldSoft : COLOR.claySoft],
        },
        { kind: "note" as const, tone: ((depDelta < 0 && biggestMover?.group === "Nécessaire") || (depDelta > 0 && biggestMover?.group !== "Productif") ? "warn" : "info") as "warn" | "info", text: verdict },
      ],
    };
  };


  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <button onClick={() => { setReportAnchor(dateToMonthKey(todayISO())); setMonthlyReportOpen(true); }} style={{
          display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`,
          borderRadius: 8, color: COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer", fontFamily: "'Inter', sans-serif", fontWeight: 600,
        }}>
          <CalendarRange size={14} /> Rapport mois
        </button>
      </div>
      <SignauxClesPanel transactions={transactions} accounts={accounts} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} monthlyObjective={monthlyObjective} />
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

      <Panel title="Score de santé financière" subtitle="Composite : taux d'épargne, poids du non-productif, stabilité des revenus"
        right={
          <button onClick={() => setHealthNarrativeOpen(true)} style={{
            display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`,
            borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer",
          }}>
            <BookOpen size={13} /> Rapport détaillé
          </button>
        }>
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

      <HealthScoreNarrativeSheet open={healthNarrativeOpen} onClose={() => setHealthNarrativeOpen(false)} tauxEpargne={tauxEpargne} pctNonProd={pctNonProd} cv={healthCv} health={health} />

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

      {monthlyReportOpen && (() => {
        const d = buildMonthlyReport(reportAnchor);
        const atCurrent = monthSortKey(reportAnchor) >= monthSortKey(dateToMonthKey(todayISO()));
        return (
          <CalcDetailSheet open={monthlyReportOpen} onClose={() => setMonthlyReportOpen(false)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks}
            onPrev={() => setReportAnchor(prevMonthKey(reportAnchor))}
            onNext={atCurrent ? undefined : () => setReportAnchor(nextMonthKey(reportAnchor))}
          />
        );
      })()}
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
  const [fluxNarrativeOpen, setFluxNarrativeOpen] = useState(false);
  const byMonthNonProd = useMemo(() => {
    const m: Record<string, number> = {};
    filtered.filter((t) => t.type === "Dépense" && t.group === "Non-productif").forEach((t) => { m[t.month] = (m[t.month] || 0) + t.amount; });
    return Object.keys(m).sort((a, b) => monthSortKey(a) - monthSortKey(b)).map((k) => ({ key: k, label: monthLabel(k), value: m[k] }));
  }, [filtered]);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Flux Revenus → Dépenses" subtitle="Comment le revenu de la période se répartit entre les groupes de dépenses et le solde"
        explain="Le bloc vert à gauche représente 100% du revenu de la période. Chaque ruban qui en part montre la part qui va vers un groupe de dépenses (Nécessaire, Productif, Non-productif, Non classifié) ou vers le Solde (épargne restante). Plus un ruban est épais, plus ce poste absorbe une grande partie du revenu.">
        <FlowDiagram filtered={filtered} />
      </PanelWithHelp>
      <PanelWithHelp title="Calendrier d'intensité — dépenses non-productives" subtitle="Repérer les mois et saisons à risque"
        explain="Chaque case représente un mois. Plus la couleur est intense (vert → or → rouge), plus les dépenses non-productives (cadeaux, sorties, shopping) ont été élevées ce mois-là. Pratique pour repérer des périodes récurrentes à risque — fêtes de fin d'année, rentrée scolaire, anniversaires groupés…"
        right={
          byMonthNonProd.length >= 2 && (
            <button onClick={() => setFluxNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
              <BookOpen size={13} /> Rapport détaillé
            </button>
          )
        }>
        <HeatmapCalendar filtered={filtered} />
      </PanelWithHelp>
      {fluxNarrativeOpen && (() => {
        const values = byMonthNonProd.map((m) => m.value);
        const avg = mean(values);
        const hot = [...byMonthNonProd].sort((a, b) => b.value - a.value).slice(0, 3);
        const attrib = (mk: string) => {
          const byCat: Record<string, number> = {};
          filtered.filter((t) => t.type === "Dépense" && t.group === "Non-productif" && t.month === mk).forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
          return Object.entries(byCat).sort((a: any, b: any) => b[1] - a[1]).slice(0, 4);
        };
        const blocks: CalcDetailBlock[] = [
          { kind: "kv", rows: [{ label: "Moyenne mensuelle (Non-productif)", value: `${fmt(avg)} FCFA`, strong: true }] },
        ];
        hot.forEach((m) => blocks.push({ kind: "note", tone: m.value > avg * 1.3 ? "warn" : "info", text: `${m.label} : ${fmt(m.value)} FCFA. Répartition : ${attrib(m.key).map(([c, v]: any) => `${c} ${fmt(v)} FCFA`).join(", ")}.` }));
        return <CalcDetailSheet open={fluxNarrativeOpen} onClose={() => setFluxNarrativeOpen(false)}
          title="Calendrier d'intensité — analyse détaillée" headline={`Moyenne ${fmt(avg)} FCFA/mois`}
          formula="Mois les plus intenses en dépenses Non-productif, attribués aux catégories qui les expliquent" blocks={blocks} />;
      })()}
    </div>
  );
}

// ============================================================ END OF PART 3 — continued below
// ============================================================
// RAPPORT MENSUEL TAB
// ============================================================
function MensuelTab({ filtered }: { filtered: any[] }) {
  const [sortKey, setSortKey] = useState<"mois" | "revenus" | "depenses" | "solde" | "nonProd">("mois");
  const [sortDir, setSortDir] = useState<1 | -1>(-1);
  const [mensuelNarrativeOpen, setMensuelNarrativeOpen] = useState(false);

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
        explain="La ligne dorée (Solde net) trace la différence revenus−dépenses de chaque mois : au-dessus de zéro, le mois est excédentaire. La ligne rouge pointillée (Non-productif) montre le poids des dépenses sans retour (cadeaux, sorties…) — si elle suit de près ou dépasse le solde net, c'est souvent le premier poste à réduire pour améliorer l'épargne."
        right={
          chartData.length >= 2 && (
            <button onClick={() => setMensuelNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
              <BookOpen size={13} /> Rapport détaillé
            </button>
          )
        }>
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
      {mensuelNarrativeOpen && (() => {
        const values = chartData.map((r) => r.solde);
        const bestIdx = values.indexOf(Math.max(...values)), worstIdx = values.indexOf(Math.min(...values));
        const deficits = chartData.filter((r) => r.solde < 0);
        const half = Math.floor(chartData.length / 2);
        const firstHalfAvg = mean(values.slice(0, half || 1));
        const secondHalfAvg = mean(values.slice(half));
        const attrib = (mk: string) => {
          const byCat: Record<string, number> = {};
          filtered.filter((t) => t.type === "Dépense" && t.month === mk).forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
          return Object.entries(byCat).sort((a: any, b: any) => b[1] - a[1]).slice(0, 4);
        };
        const blocks: CalcDetailBlock[] = [
          { kind: "kv", rows: [
            { label: "Solde total sur la période", value: `${fmt(totals.solde)} FCFA`, strong: true, warn: totals.solde < 0 },
            { label: "Mois en déficit", value: `${deficits.length} / ${chartData.length}`, warn: deficits.length > chartData.length / 3 },
            { label: "Tendance du solde (1ère moitié → 2e moitié)", value: `${fmt(firstHalfAvg)} → ${fmt(secondHalfAvg)} FCFA/mois` },
          ] },
          { kind: "note", tone: "info", text: `Meilleur mois : ${chartData[bestIdx]?.label} (solde +${fmt(values[bestIdx])} FCFA).` },
          { kind: "note", tone: "warn", text: `Pire mois : ${chartData[worstIdx]?.label} (solde ${fmt(values[worstIdx])} FCFA).${attrib(chartData[worstIdx]?.mois).length ? ` Principales dépenses ce mois-là : ${attrib(chartData[worstIdx]?.mois).map(([c, v]: any) => `${c} (${fmt(v)} FCFA)`).join(", ")}.` : ""}` },
        ];
        if (deficits.length) blocks.push({ kind: "table", columns: ["Mois en déficit", "Solde (FCFA)"], rows: deficits.map((r) => [r.label, fmt(r.solde)]) });
        return <CalcDetailSheet open={mensuelNarrativeOpen} onClose={() => setMensuelNarrativeOpen(false)}
          title="Rapport mensuel — analyse détaillée" headline={`${fmt(totals.solde)} FCFA de solde cumulé sur ${chartData.length} mois`}
          formula="Mois triés par solde net ; attribution du pire mois aux catégories qui l'expliquent" blocks={blocks} />;
      })()}
    </div>
  );
}

// ============================================================
// CATÉGORIES TAB (reclassification + détection d'anomalies)
// ============================================================
// ============================================================
// GESTION DES CATÉGORIES — créer, renommer, supprimer catégories et
// sous-catégories, avec répercussion automatique sur les transactions, les
// groupes (Nécessaire/Productif/Non-productif), les portées Business/Personnel,
// les activités et les budgets. Une catégorie/sous-catégorie déjà utilisée ne
// peut être supprimée qu'en fusionnant ses transactions vers une autre.
// ============================================================
// Fiche de saisie de nom, même esprit visuel que "Saisie rapide" — grand champ
// central, dégradé doré — réutilisée pour créer/renommer une catégorie ou une
// sous-catégorie plutôt qu'un simple champ texte en ligne.
function CategoryNameSheet({ open, title, subtitle, initialValue, confirmLabel, accentColor, onClose, onSave, showGroup, initialGroup, allowFollowCategory }: {
  open: boolean; title: string; subtitle?: string; initialValue: string; confirmLabel: string; accentColor: string;
  onClose: () => void; onSave: (value: string, group?: Group) => void; showGroup?: boolean; initialGroup?: Group; allowFollowCategory?: boolean;
}) {
  const [value, setValue] = useState(initialValue);
  const [group, setGroup] = useState<Group | null>(initialGroup || (allowFollowCategory ? null : "Nécessaire"));
  const [saved, setSaved] = useState(false);
  useEffect(() => { if (open) { setValue(initialValue); setGroup(initialGroup || (allowFollowCategory ? null : "Nécessaire")); setSaved(false); } }, [open, initialValue, initialGroup, allowFollowCategory]);
  if (!open) return null;

  const submit = () => {
    if (!value.trim()) return;
    onSave(value.trim(), showGroup && group ? group : undefined);
    setSaved(true);
    setTimeout(() => { setSaved(false); onClose(); }, 500);
  };

  const groupOptions: { key: Group; label: string; hint: string }[] = [
    { key: "Nécessaire", label: "Nécessaire", hint: "Logement, santé, obligations" },
    { key: "Productif", label: "Productif", hint: "Investissement, épargne, retraite" },
    { key: "Non-productif", label: "Non-productif", hint: "Loisirs, envies, discrétionnaire" },
  ];

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 470, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: "100%", maxWidth: 460 }}>
        <div style={{ background: `linear-gradient(180deg, ${COLOR.surfaceRaised} 0%, ${COLOR.surface} 70%)`, border: `1px solid ${COLOR.hairline}`, borderBottom: "none", borderRadius: "16px 16px 0 0", overflow: "hidden" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px 0 20px" }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 15, color: accentColor }}>{title}</div>
            <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex" }}><X size={18} /></button>
          </div>
          {subtitle && <div style={{ padding: "4px 20px 0 20px", fontSize: 11.5, color: COLOR.inkMuted, lineHeight: 1.5 }}>{subtitle}</div>}

          <div style={{ position: "relative", display: "flex", flexDirection: "column", alignItems: "center", padding: "22px 24px 10px 24px" }}>
            <Layers size={64} style={{ position: "absolute", top: 14, color: accentColor, opacity: 0.08, pointerEvents: "none" }} />
            <input
              autoFocus value={value} placeholder="Nom…" onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
              style={{ position: "relative", background: "transparent", border: "none", outline: "none", color: COLOR.ink, fontSize: 28, fontWeight: 600, fontFamily: "'Fraunces', serif", textAlign: "center", width: "100%" }}
            />
          </div>

          {showGroup && (
            <div style={{ padding: "4px 20px 16px 20px" }}>
              <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8, textAlign: "center" }}>Nature de cette {allowFollowCategory ? "sous-catégorie" : "catégorie"}</div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {allowFollowCategory && (
                  <button onClick={() => setGroup(null)} title="Ne définit rien de spécifique — hérite du groupe de la catégorie parente" style={{
                    flex: "1 1 100%", padding: "8px 6px", borderRadius: 10, cursor: "pointer", textAlign: "center",
                    border: `1px solid ${group === null ? COLOR.gold : COLOR.hairline}`,
                    background: group === null ? "rgba(201,162,39,0.14)" : "transparent",
                  }}>
                    <div style={{ fontSize: 11.5, fontWeight: 600, color: group === null ? COLOR.goldSoft : COLOR.ink }}>Suivre la catégorie (par défaut)</div>
                  </button>
                )}
                {groupOptions.map((g) => (
                  <button key={g.key} onClick={() => setGroup(g.key)} title={g.hint} style={{
                    flex: 1, padding: "10px 6px", borderRadius: 10, cursor: "pointer", textAlign: "center",
                    border: `1px solid ${group === g.key ? groupColor[g.key] : COLOR.hairline}`,
                    background: group === g.key ? `${groupColor[g.key]}22` : "transparent",
                  }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: group === g.key ? groupColor[g.key] : COLOR.ink }}>{g.label}</div>
                    <div style={{ fontSize: 9.5, color: COLOR.inkMuted, marginTop: 2 }}>{g.hint}</div>
                  </button>
                ))}
              </div>
            </div>
          )}

          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "16px 20px" }}>
            <button onClick={submit} disabled={!value.trim()} style={{
              width: "100%", padding: "14px 0", borderRadius: 12, border: "none",
              background: saved ? COLOR.emerald : !value.trim() ? COLOR.hairline : accentColor,
              color: saved ? COLOR.bg : !value.trim() ? COLOR.inkMuted : COLOR.bg,
              fontSize: 14.5, fontWeight: 700, cursor: !value.trim() ? "default" : "pointer",
              display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "background 0.15s",
            }}>{saved ? <Check size={17} /> : null} {saved ? "Enregistré" : confirmLabel}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function CategoryManagementTab({
  transactions, setTransactions, customDepSubcategories, setCustomDepSubcategories, customRevSubcategories, setCustomRevSubcategories,
  categoryGroups, setCategoryGroups, categoryScope, setCategoryScope, categoryActivity, setCategoryActivity, budgets, setBudgets,
}: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void;
  customDepSubcategories: Record<string, string[]>; setCustomDepSubcategories: (m: Record<string, string[]>) => void;
  customRevSubcategories: Record<string, string[]>; setCustomRevSubcategories: (m: Record<string, string[]>) => void;
  categoryGroups: Record<string, Group>; setCategoryGroups: (g: Record<string, Group>) => void;
  categoryScope: Record<string, Scope>; setCategoryScope: (s: Record<string, Scope>) => void;
  categoryActivity: Record<string, string>; setCategoryActivity: (a: Record<string, string>) => void;
  budgets: CategoryBudget[]; setBudgets: (b: CategoryBudget[]) => void;
}) {
  const [typeView, setTypeView] = useState<TxType>("Dépense");
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [catSheet, setCatSheet] = useState<{ mode: "new" | "rename"; oldName?: string } | null>(null);
  const [subSheet, setSubSheet] = useState<{ cat: string; mode: "new" | "rename"; oldName?: string } | null>(null);
  const [deleteTarget, setDeleteTarget] = useState<{ kind: "cat" | "sub"; cat: string; sub?: string } | null>(null);
  const [mergeInto, setMergeInto] = useState("");

  const activeMap = typeView === "Dépense" ? customDepSubcategories : customRevSubcategories;
  const setActiveMap = typeView === "Dépense" ? setCustomDepSubcategories : setCustomRevSubcategories;
  const categories = categoriesForType(transactions, typeView);
  const accentColor = typeView === "Revenu" ? COLOR.emeraldSoft : COLOR.goldSoft;

  const countForCat = (cat: string) => transactions.filter((t) => t.category === cat && t.type === typeView).length;
  const countForSub = (cat: string, sub: string) => transactions.filter((t) => t.category === cat && t.subcategory === sub && t.type === typeView).length;

  const addCategory = (name: string, group?: Group) => {
    if (activeMap[name] !== undefined) return;
    setActiveMap({ ...activeMap, [name]: [] });
    if (group) setCategoryGroups({ ...categoryGroups, [name]: group });
  };

  const addSubcategory = (cat: string, name: string, group?: Group) => {
    const subs = activeMap[cat] || [];
    if (subs.includes(name)) return;
    setActiveMap({ ...activeMap, [cat]: [...subs, name] });
    if (group) setCategoryGroups({ ...categoryGroups, [`${cat}::${name}`]: group });
  };

  const renameKeyed = (map: Record<string, any> | undefined, setMap: ((m: any) => void) | undefined, oldKey: string, newKey: string) => {
    if (!map || !setMap || map[oldKey] === undefined) return;
    const next = { ...map }; next[newKey] = next[oldKey]; delete next[oldKey]; setMap(next);
  };

  const renameCategory = (oldName: string, newName: string) => {
    if (newName === oldName || activeMap[newName] !== undefined) return;
    const next = { ...activeMap }; next[newName] = next[oldName] || []; delete next[oldName]; setActiveMap(next);
    setTransactions(transactions.map((t) => (t.category === oldName && t.type === typeView ? { ...t, category: newName } : t)));
    renameKeyed(categoryGroups, setCategoryGroups, oldName, newName);
    renameKeyed(categoryScope, setCategoryScope, oldName, newName);
    renameKeyed(categoryActivity, setCategoryActivity, oldName, newName);
    setBudgets(budgets.map((b) => (b.category === oldName ? { ...b, category: newName } : b)));
    // Répercute aussi les groupes par sous-catégorie ("Ancien::Sous" -> "Nouveau::Sous").
    const ng = { ...categoryGroups };
    Object.keys(ng).forEach((k) => { if (k.startsWith(`${oldName}::`)) { ng[k.replace(`${oldName}::`, `${newName}::`)] = ng[k]; delete ng[k]; } });
    setCategoryGroups(ng);
  };

  const renameSubcategory = (cat: string, oldSub: string, newSub: string, group?: Group) => {
    const subs = (activeMap[cat] || []).map((s) => (s === oldSub ? newSub : s));
    if (newSub !== oldSub) {
      setActiveMap({ ...activeMap, [cat]: subs });
      setTransactions(transactions.map((t) => (t.category === cat && t.subcategory === oldSub && t.type === typeView ? { ...t, subcategory: newSub } : t)));
    }
    if (group) {
      const next = { ...categoryGroups };
      delete next[`${cat}::${oldSub}`];
      next[`${cat}::${newSub}`] = group;
      setCategoryGroups(next);
    } else if (newSub !== oldSub) {
      renameKeyed(categoryGroups, setCategoryGroups, `${cat}::${oldSub}`, `${cat}::${newSub}`);
    }
  };

  const performDelete = () => {
    if (!deleteTarget) return;
    if (deleteTarget.kind === "cat") {
      const cat = deleteTarget.cat;
      const count = countForCat(cat);
      if (count > 0) {
        if (!mergeInto) return;
        setTransactions(transactions.map((t) => (t.category === cat && t.type === typeView ? { ...t, category: mergeInto, subcategory: undefined } : t)));
      }
      const next = { ...activeMap }; delete next[cat]; setActiveMap(next);
      const ng = { ...categoryGroups }; delete ng[cat]; setCategoryGroups(ng);
      const ns = { ...categoryScope }; delete ns[cat]; setCategoryScope(ns);
      const na = { ...categoryActivity }; delete na[cat]; setCategoryActivity(na);
      setBudgets(budgets.filter((b) => b.category !== cat));
    } else {
      const { cat, sub } = deleteTarget;
      if (!sub) return;
      const count = countForSub(cat, sub);
      if (count > 0) {
        setTransactions(transactions.map((t) => (t.category === cat && t.subcategory === sub && t.type === typeView ? { ...t, subcategory: mergeInto || undefined } : t)));
      }
      setActiveMap({ ...activeMap, [cat]: (activeMap[cat] || []).filter((s) => s !== sub) });
    }
    setDeleteTarget(null); setMergeInto("");
  };

  const deleteCount = deleteTarget ? (deleteTarget.kind === "cat" ? countForCat(deleteTarget.cat) : countForSub(deleteTarget.cat, deleteTarget.sub || "")) : 0;
  const mergeOptions = deleteTarget?.kind === "cat" ? categories.filter((c) => c !== deleteTarget.cat) : (deleteTarget ? activeMap[deleteTarget.cat] || [] : []).filter((s: string) => s !== deleteTarget?.sub);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Gestion des catégories" subtitle="Créer, renommer et supprimer des catégories et sous-catégories"
        explain="Renommer une catégorie ou sous-catégorie met à jour automatiquement toutes les transactions existantes qui l'utilisent, ainsi que son groupe (Nécessaire/Productif/Non-productif), sa portée Business/Personnel, son activité et ses budgets. Supprimer une catégorie déjà utilisée par des transactions exige de choisir une catégorie de remplacement — les transactions y sont alors basculées avant la suppression, pour ne jamais perdre de données silencieusement.">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 18 }}>
          <div style={{ display: "flex", gap: 4, background: COLOR.surface, borderRadius: 16, padding: 3, border: `1px solid ${COLOR.hairline}` }}>
            {(["Dépense", "Revenu"] as TxType[]).map((ty) => (
              <button key={ty} onClick={() => setTypeView(ty)} style={{
                padding: "6px 16px", borderRadius: 12, fontSize: 12.5, cursor: "pointer", border: "none",
                background: typeView === ty ? (ty === "Revenu" ? COLOR.emerald : COLOR.clay) : "transparent",
                color: typeView === ty ? COLOR.bg : COLOR.inkMuted,
              }}>{ty}</button>
            ))}
          </div>
          <button onClick={() => setCatSheet({ mode: "new" })} style={{
            display: "flex", alignItems: "center", gap: 8, background: `linear-gradient(135deg, ${accentColor}22 0%, ${COLOR.surfaceRaised} 70%)`,
            border: `1px solid ${accentColor}`, borderRadius: 10, color: accentColor, padding: "9px 18px", fontSize: 13, fontWeight: 600, cursor: "pointer",
          }}>
            <Plus size={15} /> Nouvelle catégorie
          </button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {categories.map((cat) => {
            const isOpen = expanded.has(cat);
            const subs = activeMap[cat] || [];
            const count = countForCat(cat);
            return (
              <div key={cat} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, overflow: "hidden" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px" }}>
                  <button onClick={() => setExpanded((prev) => { const n = new Set(prev); n.has(cat) ? n.delete(cat) : n.add(cat); return n; })} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}>
                    <ChevronDown size={14} style={{ transform: isOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                  </button>
                  <span style={{ flex: 1, fontSize: 13.5, color: COLOR.ink, fontWeight: 600, fontFamily: "'Fraunces', serif" }}>{cat} <span style={{ fontWeight: 400, fontSize: 11, color: COLOR.inkMuted, fontFamily: "'Inter', sans-serif" }}>({count} tx · {subs.length} sous-cat.)</span></span>
                  <button onClick={() => setCatSheet({ mode: "rename", oldName: cat })} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={13} /></button>
                  <button onClick={() => { setDeleteTarget({ kind: "cat", cat }); setMergeInto(""); }} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
                </div>
                {isOpen && (
                  <div style={{ padding: "0 14px 12px 40px", display: "flex", flexDirection: "column", gap: 6 }}>
                    {subs.map((sub) => (
                      <div key={sub} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ flex: 1, fontSize: 12.5, color: COLOR.inkMuted }}>{sub} <span style={{ fontSize: 10.5 }}>({countForSub(cat, sub)} tx)</span></span>
                        <button onClick={() => setSubSheet({ cat, mode: "rename", oldName: sub })} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={11} /></button>
                        <button onClick={() => { setDeleteTarget({ kind: "sub", cat, sub }); setMergeInto(""); }} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={11} /></button>
                      </div>
                    ))}
                    <button onClick={() => setSubSheet({ cat, mode: "new" })} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "none", color: COLOR.goldSoft, cursor: "pointer", fontSize: 11.5, padding: "4px 0", width: "fit-content" }}>
                      <Plus size={11} /> Ajouter une sous-catégorie
                    </button>
                    {!subs.length && <div style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Aucune sous-catégorie.</div>}
                  </div>
                )}
              </div>
            );
          })}
          {!categories.length && <EmptyState text="Aucune catégorie." />}
        </div>
      </PanelWithHelp>

      <CategoryNameSheet
        open={!!catSheet}
        title={catSheet?.mode === "new" ? "Nouvelle catégorie" : "Renommer la catégorie"}
        subtitle={typeView}
        initialValue={catSheet?.mode === "rename" ? catSheet.oldName || "" : ""}
        confirmLabel={catSheet?.mode === "new" ? "Créer" : "Renommer"}
        accentColor={accentColor}
        onClose={() => setCatSheet(null)}
        showGroup
        initialGroup={catSheet?.mode === "rename" && catSheet.oldName ? categoryGroups[catSheet.oldName] : undefined}
        onSave={(value, group) => {
          if (catSheet?.mode === "new") { addCategory(value, group); return; }
          if (!catSheet?.oldName) return;
          renameCategory(catSheet.oldName, value);
          if (group) setCategoryGroups({ ...categoryGroups, [value]: group });
        }}
      />
      <CategoryNameSheet
        open={!!subSheet}
        title={subSheet?.mode === "new" ? "Nouvelle sous-catégorie" : "Renommer la sous-catégorie"}
        subtitle={subSheet ? `Dans « ${subSheet.cat} » — laisse "Suivre la catégorie" pour garder le groupe de "${subSheet.cat}" (${categoryGroups[subSheet.cat] || "Non classifié"})` : undefined}
        initialValue={subSheet?.mode === "rename" ? subSheet.oldName || "" : ""}
        confirmLabel={subSheet?.mode === "new" ? "Créer" : "Renommer"}
        accentColor={accentColor}
        onClose={() => setSubSheet(null)}
        showGroup
        allowFollowCategory
        initialGroup={subSheet?.mode === "rename" && subSheet.oldName ? categoryGroups[`${subSheet.cat}::${subSheet.oldName}`] : undefined}
        onSave={(value, group) => { if (!subSheet) return; if (subSheet.mode === "new") addSubcategory(subSheet.cat, value, group); else if (subSheet.oldName) renameSubcategory(subSheet.cat, subSheet.oldName, value, group); }}
      />

      <ConfirmDialog
        open={!!deleteTarget && deleteCount === 0}
        title={deleteTarget?.kind === "cat" ? `Supprimer la catégorie "${deleteTarget.cat}" ?` : `Supprimer la sous-catégorie "${deleteTarget?.sub}" ?`}
        message="Aucune transaction ne l'utilise actuellement — suppression sans impact."
        onConfirm={performDelete}
        onCancel={() => setDeleteTarget(null)}
      />

      {deleteTarget && deleteCount > 0 && (
        <div onClick={() => setDeleteTarget(null)} style={{ position: "fixed", inset: 0, zIndex: 500, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ width: "100%", maxWidth: 420, background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 14, padding: 20 }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, color: COLOR.ink, marginBottom: 8 }}>
              {deleteTarget.kind === "cat" ? `"${deleteTarget.cat}" est utilisée par ${deleteCount} transaction(s)` : `"${deleteTarget.sub}" est utilisée par ${deleteCount} transaction(s)`}
            </div>
            <div style={{ fontSize: 12.5, color: COLOR.inkMuted, marginBottom: 14, lineHeight: 1.6 }}>
              Choisis {deleteTarget.kind === "cat" ? "une catégorie" : "une sous-catégorie (ou aucune)"} vers laquelle basculer ces transactions avant suppression — aucune donnée ne sera perdue.
            </div>
            <select value={mergeInto} onChange={(e) => setMergeInto(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 16 }}>
              <option value="">{deleteTarget.kind === "sub" ? "— aucune sous-catégorie —" : "— choisir —"}</option>
              {mergeOptions.map((o: string) => <option key={o} value={o}>{o}</option>)}
            </select>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => setDeleteTarget(null)} style={{ flex: 1, padding: "10px 0", borderRadius: 8, border: `1px solid ${COLOR.hairline}`, background: "transparent", color: COLOR.inkMuted, cursor: "pointer", fontSize: 13 }}>Annuler</button>
              <button onClick={performDelete} disabled={deleteTarget.kind === "cat" && !mergeInto} style={{
                flex: 1, padding: "10px 0", borderRadius: 8, border: "none", cursor: deleteTarget.kind === "cat" && !mergeInto ? "default" : "pointer",
                background: deleteTarget.kind === "cat" && !mergeInto ? COLOR.hairline : COLOR.clay, color: deleteTarget.kind === "cat" && !mergeInto ? COLOR.inkMuted : COLOR.bg, fontWeight: 700, fontSize: 13,
              }}>Fusionner et supprimer</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


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
    const allCats = Object.entries(cats).sort((a, b) => b[1] - a[1]);
    const top = allCats.slice(0, 4);
    return { group: g, value, count: items.length, top, allCats, pct: totalDep ? (value / totalDep) * 100 : 0 };
  });
  const [groupDetailKey, setGroupDetailKey] = useState<Group | null>(null);
  const buildGroupDetail = (g: Group) => {
    const c = cards.find((x) => x.group === g)!;
    const items = dep.filter((t) => t.group === g);
    const byMonthG: Record<string, number> = {};
    items.forEach((t) => { byMonthG[t.month] = (byMonthG[t.month] || 0) + t.amount; });
    const months = Object.keys(byMonthG).sort((a, b) => monthSortKey(a) - monthSortKey(b));
    const values = months.map((m) => byMonthG[m]);
    const half = Math.floor(months.length / 2);
    const trendPct = months.length >= 2 && mean(values.slice(0, half || 1)) > 0
      ? ((mean(values.slice(half)) - mean(values.slice(0, half || 1))) / mean(values.slice(0, half || 1))) * 100 : 0;
    const blocks: CalcDetailBlock[] = [
      { kind: "table" as const, columns: ["Catégorie", "Montant (FCFA)", "%"], rows: c.allCats.map(([name, val]) => [name, fmt(val), `${(val / c.value * 100).toFixed(1)}%`]) },
    ];
    if (months.length >= 2) {
      blocks.push({ kind: "kv", rows: [{ label: "Tendance (1ère moitié → 2e moitié de la période)", value: `${trendPct >= 0 ? "+" : ""}${trendPct.toFixed(1)}%`, warn: g !== "Productif" && trendPct > 20, strong: true }] });
      blocks.push({ kind: "table", columns: ["Mois", "Montant (FCFA)"], rows: months.map((m, i) => [monthLabel(m), fmt(values[i])]) });
    }
    return {
      title: g, headline: `${fmt(c.value)} FCFA · ${c.pct.toFixed(1)}%`,
      formula: `Somme des dépenses classées "${g}" — ${c.count} transactions`,
      blocks,
    };
  };

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

      <ExpertAnalysisButton filtered={filtered} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16 }}>
        {cards.map((c) => (
          <div key={c.group} style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ width: 10, height: 10, borderRadius: "50%", background: groupColor[c.group] }} />
                <span style={{ fontSize: 14, fontWeight: 500 }}>{c.group}</span>
                <CalcDetailIcon onClick={() => setGroupDetailKey(c.group)} />
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
      {groupDetailKey && (() => {
        const d = buildGroupDetail(groupDetailKey);
        return <CalcDetailSheet open={!!groupDetailKey} onClose={() => setGroupDetailKey(null)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} />;
      })()}
    </div>
  );
}

// ============================================================
// COMPARATIF ANNUEL TAB
// ============================================================
// Compare les deux années les plus récentes (en moyenne mensuelle, pour rester juste
// même si l'une des deux est partielle) et attribue l'écart aux catégories qui ont le
// plus contribué à la hausse ou à la baisse — pas juste "2025 vs 2026", mais "pourquoi".
function generateYearComparisonNarrative(yearTotals: { year: string; dep: number; rev: number; depPerMonth: number; revPerMonth: number }[], byYearCategory: Record<string, Record<string, number>>, monthCounts: Record<string, number>) {
  if (yearTotals.length < 2) return null;
  const prev = yearTotals[yearTotals.length - 2], cur = yearTotals[yearTotals.length - 1];
  const depChangePct = prev.depPerMonth > 0 ? ((cur.depPerMonth - prev.depPerMonth) / prev.depPerMonth) * 100 : 0;
  const revChangePct = prev.revPerMonth > 0 ? ((cur.revPerMonth - prev.revPerMonth) / prev.revPerMonth) * 100 : 0;

  const catDeltas = Object.entries(byYearCategory).map(([cat, byYear]) => {
    const prevPerMonth = (byYear[prev.year] || 0) / (monthCounts[prev.year] || 1);
    const curPerMonth = (byYear[cur.year] || 0) / (monthCounts[cur.year] || 1);
    return { cat, delta: curPerMonth - prevPerMonth, prevPerMonth, curPerMonth };
  }).filter((d) => Math.abs(d.delta) > 1000);
  const risers = catDeltas.filter((d) => d.delta > 0).sort((a, b) => b.delta - a.delta).slice(0, 4);
  const fallers = catDeltas.filter((d) => d.delta < 0).sort((a, b) => a.delta - b.delta).slice(0, 4);

  const blocks: CalcDetailBlock[] = [
    { kind: "kv", rows: [
      { label: `Dépense moyenne/mois ${prev.year} → ${cur.year}`, value: `${fmt(prev.depPerMonth)} → ${fmt(cur.depPerMonth)} FCFA`, strong: true },
      { label: "Variation", value: `${depChangePct >= 0 ? "+" : ""}${depChangePct.toFixed(1)}%`, warn: depChangePct > 15 },
      { label: `Revenu moyen/mois ${prev.year} → ${cur.year}`, value: `${fmt(prev.revPerMonth)} → ${fmt(cur.revPerMonth)} FCFA` },
      { label: "Variation", value: `${revChangePct >= 0 ? "+" : ""}${revChangePct.toFixed(1)}%` },
    ] },
  ];
  if (risers.length) blocks.push({ kind: "table", columns: ["Catégorie en hausse", `${prev.year}/mois`, `${cur.year}/mois`], rows: risers.map((r) => [r.cat, fmt(r.prevPerMonth), fmt(r.curPerMonth)]) });
  if (fallers.length) blocks.push({ kind: "table", columns: ["Catégorie en baisse", `${prev.year}/mois`, `${cur.year}/mois`], rows: fallers.map((r) => [r.cat, fmt(r.prevPerMonth), fmt(r.curPerMonth)]) });

  return {
    title: `Comparatif ${prev.year} → ${cur.year}`,
    headline: `Dépenses ${depChangePct >= 0 ? "+" : ""}${depChangePct.toFixed(1)}% · Revenus ${revChangePct >= 0 ? "+" : ""}${revChangePct.toFixed(1)}%`,
    formula: "Moyennes mensuelles (pas les totaux bruts, pour rester juste même avec une année partielle) ; catégories triées par écart de moyenne mensuelle",
    blocks,
  };
}


function ComparatifTab({ transactions, categoryGroups }: { transactions: Transaction[]; categoryGroups: Record<string, Group> }) {
  const [yearNarrativeOpen, setYearNarrativeOpen] = useState(false);
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
      <Panel title="Comparatif annuel — moyenne mensuelle" subtitle="2024 et 2026 sont des années partielles ; la comparaison se fait donc par moyenne mensuelle, pas par total brut"
        right={years.length >= 2 && (
          <button onClick={() => setYearNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
            <BookOpen size={13} /> Rapport détaillé
          </button>
        )}>
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
      {yearNarrativeOpen && (() => {
        const d = generateYearComparisonNarrative(yearTotals, byYearCategory, monthCounts);
        return d ? <CalcDetailSheet open={yearNarrativeOpen} onClose={() => setYearNarrativeOpen(false)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} /> : null;
      })()}
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
    () => transactions.map((t) => ({ ...t, month: dateToMonthKey(t.date), group: t.type === "Revenu" ? "Revenu" : groupFor(t, categoryGroups) })),
    [transactions, categoryGroups]
  );

  const lastMonth = allMonths[allMonths.length - 1] || "";
  const idxLast = allMonths.indexOf(lastMonth);
  const prevMonth = idxLast > 0 ? allMonths[idxLast - 1] : lastMonth;

  const [granularity, setGranularity] = useState<"mois" | "jour">("mois");
  const [comparateurNarrativeOpen, setComparateurNarrativeOpen] = useState(false);

  const [aFrom, setAFrom] = useState(prevMonth);
  const [aTo, setATo] = useState(prevMonth);
  const [bFrom, setBFrom] = useState(lastMonth);
  const [bTo, setBTo] = useState(lastMonth);
  // Période A : totalement libre (tous les mois proposés), juste protégée contre une
  // inversion interne (Du postérieur à Au). Période B : se cale automatiquement à partir
  // de la fin de la période A dès qu'elle change — sur demande explicite de l'utilisateur,
  // pour comparer "A puis ce qui vient après" sans avoir à retoucher B à chaque fois.
  const setAFromSafe = (v: string) => { if (monthSortKey(v) > monthSortKey(aTo)) { setAFrom(aTo); setATo(v); } else setAFrom(v); };
  const setAToSafe = (v: string) => { if (monthSortKey(v) < monthSortKey(aFrom)) { setATo(aFrom); setAFrom(v); } else setATo(v); };
  const setBFromSafe = (v: string) => { if (monthSortKey(v) > monthSortKey(bTo)) { setBFrom(bTo); setBTo(v); } else setBFrom(v); };
  const setBToSafe = (v: string) => { if (monthSortKey(v) < monthSortKey(bFrom)) { setBTo(bFrom); setBFrom(v); } else setBTo(v); };
  useEffect(() => {
    if (monthSortKey(bFrom) < monthSortKey(aTo)) {
      setBFrom(aTo);
      if (monthSortKey(bTo) < monthSortKey(aTo)) setBTo(aTo);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [aTo]);

  const allDates = useMemo(() => Array.from(new Set(transactions.map((t) => t.date))).sort(), [transactions]);
  const lastDate = allDates[allDates.length - 1] || todayISO();
  const idxLastDate = allDates.indexOf(lastDate);
  const prevDate = idxLastDate > 0 ? allDates[idxLastDate - 1] : lastDate;

  const [aFromDay, setAFromDay] = useState(prevDate);
  const [aToDay, setAToDay] = useState(prevDate);
  const [bFromDay, setBFromDay] = useState(lastDate);
  const [bToDay, setBToDay] = useState(lastDate);

  const statsFor = (from: string, to: string) => {
    const tx = granularity === "jour"
      ? withGroup.filter((t) => t.date >= from && t.date <= to)
      : withGroup.filter((t) => { const fk = monthSortKey(from), tk = monthSortKey(to); const k = monthSortKey(t.month); return k >= fk && k <= tk; });
    const revenus = tx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
    const depenses = tx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
    const nonProd = tx.filter((t) => t.type === "Dépense" && t.group === "Non-productif").reduce((a, t) => a + t.amount, 0);
    const solde = revenus - depenses;
    const tauxEpargne = revenus > 0 ? (solde / revenus) * 100 : 0;
    const byCat: Record<string, number> = {};
    tx.filter((t) => t.type === "Dépense").forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
    return { revenus, depenses, solde, tauxEpargne, nonProd, byCat, count: tx.length };
  };

  const A = useMemo(() => statsFor(granularity === "jour" ? aFromDay : aFrom, granularity === "jour" ? aToDay : aTo), [withGroup, aFrom, aTo, aFromDay, aToDay, granularity]);
  const B = useMemo(() => statsFor(granularity === "jour" ? bFromDay : bFrom, granularity === "jour" ? bToDay : bTo), [withGroup, bFrom, bTo, bFromDay, bToDay, granularity]);

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

  const labelA = granularity === "jour"
    ? `${dateLabelShort(aFromDay)}${aFromDay !== aToDay ? " — " + dateLabelShort(aToDay) : ""}`
    : `${monthLabel(aFrom)}${aFrom !== aTo ? " — " + monthLabel(aTo) : ""}`;
  const labelB = granularity === "jour"
    ? `${dateLabelShort(bFromDay)}${bFromDay !== bToDay ? " — " + dateLabelShort(bToDay) : ""}`
    : `${monthLabel(bFrom)}${bFrom !== bTo ? " — " + monthLabel(bTo) : ""}`;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Choisir les deux périodes à comparer">
        <div style={{ display: "flex", gap: 4, background: COLOR.surface, borderRadius: 16, padding: 3, border: `1px solid ${COLOR.hairline}`, marginBottom: 20, width: "fit-content" }}>
          {(["mois", "jour"] as const).map((g) => (
            <button key={g} onClick={() => setGranularity(g)} style={{
              padding: "6px 14px", borderRadius: 12, fontSize: 12, cursor: "pointer", border: "none",
              background: granularity === g ? COLOR.gold : "transparent",
              color: granularity === g ? COLOR.bg : COLOR.inkMuted,
            }}>{g === "mois" ? "Par mois" : "Par jour"}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 30, flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: COLOR.slateBlue }} />
              <span style={{ fontSize: 12.5, color: COLOR.slateBlueSoft, fontWeight: 600 }}>Période A</span>
            </div>
            {granularity === "mois" ? (
              <div style={{ display: "flex", gap: 10 }}>
                <Select label="Du mois" value={aFrom} onChange={setAFromSafe} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
                <Select label="Au mois" value={aTo} onChange={setAToSafe} options={allMonths.map((m) => ({ value: m, label: monthLabel(m) }))} />
              </div>
            ) : (
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Du jour</label>
                  <input type="date" style={inputStyle} value={aFromDay} onChange={(e) => setAFromDay(e.target.value)} />
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Au jour</label>
                  <input type="date" style={inputStyle} value={aToDay} onChange={(e) => setAToDay(e.target.value)} />
                </div>
              </div>
            )}
          </div>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: COLOR.gold }} />
              <span style={{ fontSize: 12.5, color: COLOR.goldSoft, fontWeight: 600 }}>Période B</span>
            </div>
            {granularity === "mois" ? (
              <div style={{ display: "flex", gap: 10 }}>
                <Select label="Du mois" value={bFrom} onChange={setBFromSafe} options={allMonths.filter((m) => monthSortKey(m) >= monthSortKey(aTo) && monthSortKey(m) <= monthSortKey(bTo)).map((m) => ({ value: m, label: monthLabel(m) }))} />
                <Select label="Au mois" value={bTo} onChange={setBToSafe} options={allMonths.filter((m) => monthSortKey(m) >= monthSortKey(bFrom)).map((m) => ({ value: m, label: monthLabel(m) }))} />
              </div>
            ) : (
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Du jour</label>
                  <input type="date" style={inputStyle} value={bFromDay} onChange={(e) => setBFromDay(e.target.value)} />
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Au jour</label>
                  <input type="date" style={inputStyle} value={bToDay} onChange={(e) => setBToDay(e.target.value)} />
                </div>
              </div>
            )}
          </div>
        </div>
      </Panel>

      <Panel title="Comparaison des indicateurs clés" subtitle={`A : ${labelA} (${A.count} tx)  ·  B : ${labelB} (${B.count} tx)`}>
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
            <Bar dataKey="A" name={`A (${labelA})`} fill={COLOR.slateBlue} radius={[3, 3, 0, 0]} />
            <Bar dataKey="B" name={`B (${labelB})`} fill={COLOR.gold} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </PanelWithHelp>

      <Panel title="Catégories qui ont le plus évolué" subtitle="Triées par variation absolue entre A et B"
        right={
          <button onClick={() => setComparateurNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
            <BookOpen size={13} /> Rapport détaillé
          </button>
        }>
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
      {comparateurNarrativeOpen && (() => {
        const depDelta = pctDelta(A.depenses, B.depenses);
        const revDelta = pctDelta(A.revenus, B.revenus);
        const blocks: CalcDetailBlock[] = [
          { kind: "kv", rows: [
            { label: `Dépenses A (${labelA}) → B (${labelB})`, value: `${fmt(A.depenses)} → ${fmt(B.depenses)} FCFA`, strong: true },
            { label: "Variation", value: `${depDelta >= 0 ? "+" : ""}${depDelta.toFixed(1)}%`, warn: depDelta > 20 },
            { label: `Revenus A → B`, value: `${fmt(A.revenus)} → ${fmt(B.revenus)} FCFA` },
            { label: "Variation", value: `${revDelta >= 0 ? "+" : ""}${revDelta.toFixed(1)}%` },
          ] },
          { kind: "table", columns: ["Catégorie", `A (${labelA})`, `B (${labelB})`, "Écart"], rows: catRows.slice(0, 10).map((r) => [r.cat, fmt(r.a), fmt(r.b), `${r.delta >= 0 ? "+" : ""}${fmt(r.delta)}`]) },
        ];
        return <CalcDetailSheet open={comparateurNarrativeOpen} onClose={() => setComparateurNarrativeOpen(false)}
          title="Comparatif détaillé" headline={`${depDelta >= 0 ? "+" : ""}${depDelta.toFixed(1)}% de dépenses entre A et B`}
          formula="Catégories triées par écart absolu entre les deux périodes choisies" blocks={blocks} />;
      })()}
    </div>
  );
}

// ============================================================
// PRINCIPALES CATÉGORIES — anneau coloré, pastilles de mois, comparaison
// ============================================================
const DONUT_COLORS = [COLOR.emerald, COLOR.slateBlue, COLOR.gold, COLOR.violet, COLOR.clay, COLOR.emeraldSoft, COLOR.slateBlueSoft, COLOR.goldSoft];

function TopCategoriesTab({ transactions, setTransactions, categoryGroups, allMonths, accounts, onNavigate }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; categoryGroups: Record<string, Group>; allMonths: string[]; accounts: Account[];
  onNavigate?: (tab: Tab, data?: any) => void;
}) {
  const withGroup = useMemo(
    () => transactions.map((t) => ({ ...t, month: dateToMonthKey(t.date), group: t.type === "Revenu" ? "Revenu" : groupFor(t, categoryGroups) })),
    [transactions, categoryGroups]
  );

  const lastMonth = allMonths[allMonths.length - 1] || "";
  const pillMonths = allMonths;
  const [selectedMonth, setSelectedMonth] = useState(lastMonth);
  const [customOpen, setCustomOpen] = useState(false);
  const [customFrom, setCustomFrom] = useState(lastMonth);
  const [customTo, setCustomTo] = useState(lastMonth);
  // Empêche une plage inversée (Du mois postérieur à Au mois) — échange les deux valeurs
  // plutôt que de laisser une plage impossible s'installer.
  const setCustomFromSafe = (v: string) => { if (monthSortKey(v) > monthSortKey(customTo)) { setCustomFrom(customTo); setCustomTo(v); } else setCustomFrom(v); };
  const setCustomToSafe = (v: string) => { if (monthSortKey(v) < monthSortKey(customFrom)) { setCustomTo(customFrom); setCustomFrom(v); } else setCustomTo(v); };
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
      const [jsPDFModule, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      // jspdf@2.5.x expose le vrai constructeur sur l'export NOMMÉ "jsPDF", pas sur
      // "default" (qui résout vers un objet inutilisable selon le mode d'interop
      // CJS/ESM) — cause du "Réessayer" systématique sur tous les boutons PDF de l'app.
      const jsPDF: any = (jsPDFModule as any).jsPDF || (jsPDFModule as any).default;
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
      drawKpiBox(14, `TOTAL ${typeView.toUpperCase()}`, `${fmtPdf(total)} FCFA`, typeView === "Revenu" ? 63 : 193, typeView === "Revenu" ? 156 : 84, typeView === "Revenu" ? 122 : 63);
      drawKpiBox(14 + kpiW + 8, "VS PÉRIODE PRÉC.", `${delta >= 0 ? "+" : "−"}${fmtPdf(Math.abs(delta))}`, improved ? 63 : 193, improved ? 156 : 84, improved ? 122 : 63);
      drawKpiBox(14 + (kpiW + 8) * 2, "VARIATION", `${delta >= 0 ? "+" : "−"}${Math.abs(deltaPct).toFixed(0)}%`, improved ? 63 : 193, improved ? 156 : 84, improved ? 122 : 63);

      doc.autoTable({
        startY: 68,
        head: [["Catégorie", "Montant (FCFA)", "% du total"]],
        body: catList.map((c) => [c.name, fmtPdf(c.value), `${c.pct.toFixed(0)}%`]),
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
        doc.autoTable({
          startY: y + 3,
          head: [["Sous-catégorie", "Montant (FCFA)", "% de la catégorie"]],
          body: subs.map((s) => [s.name, fmtPdf(s.value), `${s.pct.toFixed(0)}%`]),
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
            <Select label="Du mois" value={customFrom} onChange={setCustomFromSafe} options={allMonths.filter((m) => monthSortKey(m) <= monthSortKey(customTo)).map((m) => ({ value: m, label: monthLabel(m) }))} />
            <Select label="Au mois" value={customTo} onChange={setCustomToSafe} options={allMonths.filter((m) => monthSortKey(m) >= monthSortKey(customFrom)).map((m) => ({ value: m, label: monthLabel(m) }))} />
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

      <ExpertAnalysisButton
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
                  {onNavigate && (
                    <button onClick={(e) => { e.stopPropagation(); onNavigate("categoryoverview", { category: c.name, type: typeView }); }}
                      title="Voir la tendance mensuelle de cette catégorie" style={{ background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", display: "flex", flexShrink: 0, padding: 4 }}>
                      <Activity size={15} />
                    </button>
                  )}
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

// Analyse narrative pour une catégorie/sous-catégorie donnée : détecte les mois qui
// s'écartent nettement de la moyenne (pics et creux), les attribue aux transactions
// individuelles qui les expliquent, et mesure la tendance de fond (première moitié de
// la période vs seconde moitié) plutôt que de se contenter d'un total.
function generateCategoryNarrative(byMonth: { key: string; label: string; value: number }[], catTxAll: Transaction[], category: string, subcategory: string, type: TxType) {
  const nonZero = byMonth.filter((m) => m.value > 0);
  if (nonZero.length < 2) return null;
  const vals = byMonth.map((m) => m.value);
  const avg = mean(vals);
  const sd = stdev(vals);

  const spikes = byMonth.filter((m) => m.value > avg + sd * 0.8 && m.value > 0).sort((a, b) => b.value - a.value).slice(0, 3);
  const lows = byMonth.filter((m) => m.value > 0 && m.value < avg - sd * 0.8).sort((a, b) => a.value - b.value).slice(0, 2);

  const txForMonth = (mk: string) => catTxAll.filter((t) => dateToMonthKey(t.date) === mk);
  const topTxIn = (mk: string) => txForMonth(mk).sort((a, b) => b.amount - a.amount).slice(0, 3);

  const half = Math.floor(byMonth.length / 2);
  const firstHalfAvg = half > 0 ? mean(vals.slice(0, half)) : 0;
  const secondHalfAvg = byMonth.length - half > 0 ? mean(vals.slice(half)) : 0;
  const trendPct = firstHalfAvg > 0 ? ((secondHalfAvg - firstHalfAvg) / firstHalfAvg) * 100 : 0;

  const total = vals.reduce((a, v) => a + v, 0);
  const blocks: CalcDetailBlock[] = [
    { kind: "kv", rows: [
      { label: `Total sur la période (${byMonth.length} mois)`, value: `${fmt(total)} FCFA`, strong: true },
      { label: "Moyenne mensuelle", value: `${fmt(avg)} FCFA` },
      { label: `Tendance (1ère moitié → 2e moitié de la période)`, value: `${trendPct >= 0 ? "+" : ""}${trendPct.toFixed(1)}%`, warn: type === "Dépense" ? trendPct > 15 : trendPct < -15 },
    ] },
  ];

  if (spikes.length) {
    spikes.forEach((s) => {
      const top = topTxIn(s.key);
      blocks.push({ kind: "note", tone: "warn", text: `${s.label} : ${fmt(s.value)} FCFA, nettement au-dessus de la moyenne (${fmt(avg)} FCFA).${top.length ? ` Principales transactions ce mois-là : ${top.map((t) => `${t.subcategory ? `${t.subcategory} ` : ""}${fmt(t.amount)} FCFA${t.payee ? ` (${t.payee})` : ""}`).join(", ")}.` : ""}` });
    });
  }
  if (lows.length) {
    lows.forEach((l) => {
      blocks.push({ kind: "note", tone: "info", text: `${l.label} : ${fmt(l.value)} FCFA, nettement en dessous de la moyenne — mois calme pour cette catégorie.` });
    });
  }
  if (!spikes.length && !lows.length) {
    blocks.push({ kind: "note", tone: "info", text: "Pas de mois qui s'écarte nettement de la moyenne — comportement régulier sur cette période." });
  }

  return {
    title: `${category}${subcategory ? ` · ${subcategory}` : ""}`,
    headline: `${fmt(total)} FCFA sur ${byMonth.length} mois`,
    formula: "Mois détectés comme pics/creux si l'écart à la moyenne dépasse 0,8× l'écart-type ; tendance = 1ère moitié vs 2e moitié de la période affichée",
    blocks,
  };
}


function CategoryOverviewTab({ transactions, categoryGroups, allMonths, navContext }: {
  transactions: Transaction[]; categoryGroups: Record<string, Group>; allMonths: string[]; navContext?: { tab: Tab; data: any } | null;
}) {
  const [pickerOpen, setPickerOpen] = useState(false);
  const [catNarrativeOpen, setCatNarrativeOpen] = useState(false);
  const [type, setType] = useState<TxType>("Dépense");
  const [category, setCategory] = useState(() => defaultQuickCategory(transactions, "Dépense"));
  const [subcategory, setSubcategory] = useState("");
  const [presetKey, setPresetKey] = useState("6m");
  const [granularity, setGranularity] = useState<"mois" | "jour">("mois");

  // Arrivée depuis un lien contextuel (ex: clic sur une catégorie ailleurs dans l'app) —
  // pré-sélectionne cette catégorie/type au lieu du choix par défaut.
  useEffect(() => {
    if (navContext?.tab === "categoryoverview" && navContext.data?.category) {
      setType(navContext.data.type || "Dépense");
      setCategory(navContext.data.category);
      setSubcategory(navContext.data.subcategory || "");
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [navContext]);

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
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
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
            <button onClick={() => setCatNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 16, color: COLOR.goldSoft, padding: "7px 16px", fontSize: 12.5, cursor: "pointer" }}>
              <BookOpen size={13} /> Rapport détaillé
            </button>
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
      {catNarrativeOpen && (() => {
        const d = generateCategoryNarrative(byMonth, catTxAll, category, subcategory, type);
        return d ? <CalcDetailSheet open={catNarrativeOpen} onClose={() => setCatNarrativeOpen(false)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} /> : null;
      })()}
    </div>
  );
}

// ============================================================ END OF PART 4 — continued below
// ============================================================
// ENVELOPPES TAB (avec alertes)
// ============================================================
function EnveloppesTab({ filtered, cap, setCap }: { filtered: any[]; cap: number; setCap: (n: number) => void }) {
  const [envNarrativeOpen, setEnvNarrativeOpen] = useState(false);
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
      <Panel title="Rythme de l'enveloppe dans le temps" subtitle="Total mensuel vs plafond"
        right={
          byMonth.length >= 2 && (
            <button onClick={() => setEnvNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer" }}>
              <BookOpen size={13} /> Rapport détaillé
            </button>
          )
        }>
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
      {envNarrativeOpen && (() => {
        const overages = byMonth.filter((m) => m.total > cap);
        const attrib = (mk: string) => {
          const byCat: Record<string, number> = {};
          filtered.filter((t: any) => t.type === "Dépense" && envCats.includes(t.category) && t.month === mk).forEach((t: any) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
          return Object.entries(byCat).sort((a: any, b: any) => b[1] - a[1]);
        };
        const blocks: CalcDetailBlock[] = [
          { kind: "kv", rows: [
            { label: "Plafond mensuel", value: `${fmt(cap)} FCFA` },
            { label: "Mois où le plafond a été dépassé", value: `${overages.length} / ${byMonth.length}`, warn: overages.length > byMonth.length / 3, strong: true },
          ] },
        ];
        overages.forEach((m) => blocks.push({ kind: "note", tone: "warn", text: `${m.mois} : ${fmt(m.total)} FCFA (dépassement de ${fmt(m.total - cap)} FCFA). Répartition : ${attrib(m.key).map(([c, v]: any) => `${c} ${fmt(v)} FCFA`).join(", ")}.` }));
        if (!overages.length) blocks.push({ kind: "note", tone: "info", text: "Aucun dépassement détecté sur la période affichée — l'enveloppe est restée sous contrôle." });
        return <CalcDetailSheet open={envNarrativeOpen} onClose={() => setEnvNarrativeOpen(false)}
          title="Enveloppe — analyse détaillée" headline={`${overages.length} dépassement(s) sur ${byMonth.length} mois`}
          formula="Cadeaux + Divertissement + Invitation, comparés au plafond mensuel" blocks={blocks} />;
      })()}
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
function ProjectionTooltip({ active, payload }: any) {
  if (!active || !payload?.length) return null;
  const p = payload[0]?.payload;
  if (!p) return null;
  return (
    <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: "12px 16px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, minWidth: 190, boxShadow: "0 8px 24px rgba(0,0,0,0.35)" }}>
      <div style={{ color: COLOR.inkMuted, marginBottom: 8, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em" }}>
        {p.mois === "aujourd'hui" ? "Aujourd'hui" : `Dans ${p.mois.replace("+", "")}`}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 16, marginBottom: 4 }}>
        <span style={{ color: COLOR.goldSoft }}>● Central</span>
        <span style={{ color: COLOR.ink, fontWeight: 600 }}>{fmt(p.central)} FCFA</span>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 16, marginBottom: 4 }}>
        <span style={{ color: COLOR.emeraldSoft }}>● Optimiste</span>
        <span style={{ color: COLOR.ink }}>{fmt(p.haut)} FCFA</span>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 16 }}>
        <span style={{ color: COLOR.claySoft }}>● Prudent</span>
        <span style={{ color: COLOR.ink }}>{fmt(p.bas)} FCFA</span>
      </div>
    </div>
  );
}

// ============================================================
// RAPPORT VALEUR NETTE — historique mensuel complet, avec pour chaque mois
// une explication automatique (principal poste de dépense, principal poste de
// revenu), plus les plus grosses transactions individuelles toutes périodes
// confondues.
// ============================================================
function computeNetWorthReport(accounts: Account[], transactions: Transaction[]) {
  const series = liveNetWorthSeries(accounts, transactions);

  const rows = series.map(([month, netWorth], i) => {
    const prevNW = i > 0 ? series[i - 1][1] : accounts.reduce((a, acc) => a + acc.openingBalance, 0);
    const delta = netWorth - prevNW;
    const monthTx = transactions.filter((t) => dateToMonthKey(t.date) === month);
    const revenu = monthTx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
    const depense = monthTx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);

    const depByCat: Record<string, number> = {};
    monthTx.filter((t) => t.type === "Dépense").forEach((t) => { depByCat[t.category] = (depByCat[t.category] || 0) + t.amount; });
    const topDepCat = Object.entries(depByCat).sort((a, b) => b[1] - a[1])[0];

    const revByCat: Record<string, number> = {};
    monthTx.filter((t) => t.type === "Revenu").forEach((t) => { revByCat[t.category] = (revByCat[t.category] || 0) + t.amount; });
    const topRevCat = Object.entries(revByCat).sort((a, b) => b[1] - a[1])[0];

    let explanation = "";
    if (delta >= 0) {
      explanation = topRevCat
        ? `Hausse portée par "${topRevCat[0]}" (${fmt(topRevCat[1])} FCFA)${topDepCat ? `, malgré ${fmt(topDepCat[1])} FCFA sur "${topDepCat[0]}"` : ""}.`
        : "Peu de mouvement ce mois-ci.";
    } else {
      explanation = topDepCat
        ? `Baisse tirée par "${topDepCat[0]}" (${fmt(topDepCat[1])} FCFA)${topRevCat ? `, partiellement compensée par ${fmt(topRevCat[1])} FCFA de "${topRevCat[0]}"` : ""}.`
        : "Peu de mouvement ce mois-ci.";
    }

    return { month, netWorth, delta, revenu, depense, topDepCat, topRevCat, explanation, count: monthTx.length };
  });

  const depTx = transactions.filter((t) => t.type === "Dépense").sort((a, b) => b.amount - a.amount).slice(0, 15);
  const revTx = transactions.filter((t) => t.type === "Revenu").sort((a, b) => b.amount - a.amount).slice(0, 15);

  const best = rows.reduce((a, b) => (b.delta > a.delta ? b : a), rows[0]);
  const worst = rows.reduce((a, b) => (b.delta < a.delta ? b : a), rows[0]);

  return { rows, depTx, revTx, best, worst };
}

// Analyse narrative de la courbe de valeur nette dans l'esprit d'un rapport d'analyste
// patrimonial : détection du plus grand drawdown (creux après un sommet), attribution
// aux vraies catégories de transactions qui l'expliquent (pas des suppositions), et
// mesure de la récupération. Contrairement à une lecture générique de la courbe, ceci
// s'appuie sur les transactions réelles pour dire PRÉCISÉMENT ce qui explique chaque
// mouvement plutôt que de lister des causes possibles.
// Moteur narratif générique dans l'esprit d'un rapport d'analyste patrimonial : phases,
// forces/risques identifiés avec les VRAIES transactions qui les expliquent, scores de
// diagnostic, arc narratif chronologique, synthèse — tout en prose, sans tableau. Marche
// pour n'importe quelle série mensuelle (valeur nette, marge d'activité...) du moment
// qu'on lui donne comment attribuer une transaction à cette série.
function generateDeepNarrative(
  series: { month: string; value: number }[],
  transactions: Transaction[],
  belongsTo: (t: Transaction) => boolean,
  opts: { subject: string; cumulative: boolean } // cumulative=true pour une valeur nette (niveau), false pour une marge périodique (flux)
): string[] {
  if (series.length < 3) return [];
  const values = series.map((r) => r.value);
  const avg = mean(values);
  const sd = stdev(values);
  const cv = avg !== 0 ? Math.abs(sd / avg) * 100 : 0;
  // Échelle de référence pour exprimer une baisse en % : jamais juste "abs(pic)", qui peut
  // être proche de zéro (ex : une marge mensuelle qui frôle zéro un mois) et faire
  // exploser le pourcentage à des valeurs absurdes (vu : "-4897%"). On utilise plutôt
  // le plus grand entre le pic et la moyenne absolue de toute la série.
  const avgAbs = mean(values.map((v) => Math.abs(v))) || 1;

  const findMaxDrawdown = (slice: typeof series, offset: number) => {
    let runningPeak = slice[0]?.value ?? 0, runningPeakIdx = 0;
    let maxDD = 0, peakIdx = 0, troughIdx = 0;
    slice.forEach((r, i) => {
      if (r.value > runningPeak) { runningPeak = r.value; runningPeakIdx = i; }
      const scaleRef = Math.max(Math.abs(runningPeak), avgAbs);
      const dd = scaleRef !== 0 ? (runningPeak - r.value) / scaleRef : 0;
      if (dd > maxDD) { maxDD = dd; peakIdx = runningPeakIdx; troughIdx = i; }
    });
    return { maxDD, peakIdx: peakIdx + offset, troughIdx: troughIdx + offset };
  };
  const ep1 = findMaxDrawdown(series, 0);
  const ep2 = ep1.troughIdx < series.length - 1 ? findMaxDrawdown(series.slice(ep1.troughIdx), ep1.troughIdx) : { maxDD: 0, peakIdx: 0, troughIdx: 0 };
  const episodes = [ep1, ep2].filter((e) => e.maxDD > 0.15 && e.troughIdx > e.peakIdx);

  const attribMonth = (mk: string, n = 2) => {
    const byCat: Record<string, number> = {};
    transactions.filter((t) => t.type === "Dépense" && belongsTo(t) && dateToMonthKey(t.date) === mk)
      .forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
    return Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, n);
  };
  const attribRange = (fromIdx: number, toIdx: number, n = 3) => {
    const months = series.slice(fromIdx + 1, toIdx + 1).map((r) => r.month);
    const byCat: Record<string, number> = {};
    transactions.filter((t) => t.type === "Dépense" && belongsTo(t) && months.includes(dateToMonthKey(t.date)))
      .forEach((t) => { byCat[t.category] = (byCat[t.category] || 0) + t.amount; });
    return Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, n);
  };
  const fmtCats = (cats: [string, number][]) => cats.map(([c, v]) => `${c} (${fmt(v)} FCFA)`).join(", ");

  const half = Math.floor(series.length / 2);
  const firstHalfAvg = mean(values.slice(0, half || 1));
  const secondHalfAvg = mean(values.slice(half));
  const trendPct = firstHalfAvg !== 0 ? ((secondHalfAvg - firstHalfAvg) / Math.abs(firstHalfAvg)) * 100 : 0;
  const lastVal = values[values.length - 1];
  const first = series[0], last = series[series.length - 1];

  // Scores de diagnostic (0-10), calibrés par seuils simples plutôt que par formule
  // opaque — pour rester lisibles et défendables.
  const growthScore = trendPct >= 50 ? 9 : trendPct >= 20 ? 7 : trendPct >= 0 ? 5 : trendPct >= -20 ? 3 : 1;
  const stabilityScore = cv < 20 ? 9 : cv < 50 ? 7 : cv < 100 ? 5 : cv < 200 ? 3 : 1;
  const maxDD = episodes.length ? Math.max(...episodes.map((e) => e.maxDD)) : 0;
  const riskScore = maxDD < 0.1 ? 9 : maxDD < 0.25 ? 7 : maxDD < 0.5 ? 5 : maxDD < 0.75 ? 3 : 1;
  const recentTrend = values.length >= 4 ? mean(values.slice(-2)) - mean(values.slice(-4, -2)) : 0;
  const momentumScore = avg !== 0 ? (recentTrend / Math.abs(avg) > 0.3 ? 9 : recentTrend / Math.abs(avg) > 0 ? 6 : recentTrend / Math.abs(avg) > -0.3 ? 4 : 2) : 5;

  const out: string[] = [];
  const unit = opts.cumulative ? "niveau" : "flux mensuel";

  out.push(`De ${monthLabel(first.month)} à ${monthLabel(last.month)}, ${opts.subject} passe de ${fmt(first.value)} à ${fmt(last.value)} FCFA. ${opts.cumulative ? "La tendance de fond" : "Le rythme moyen"} sur la première moitié de la période (${fmt(firstHalfAvg)} FCFA) évolue vers ${fmt(secondHalfAvg)} FCFA sur la seconde (${trendPct >= 0 ? "+" : ""}${trendPct.toFixed(0)}%).`);

  // 🟢 Points forts
  const strengths: string[] = [];
  if (trendPct > 10) strengths.push(`tendance structurelle ${opts.cumulative ? "haussière" : "en amélioration"} sur la période`);
  if (episodes.length && episodes[episodes.length - 1].troughIdx < series.length - 1) {
    const lastEp = episodes[episodes.length - 1];
    const recoveryEndIdx = series.length - 1;
    const recoveryScaleRef = Math.max(Math.abs(series[lastEp.troughIdx].value), avgAbs);
    const recoveryPct = recoveryScaleRef !== 0 ? ((series[recoveryEndIdx].value - series[lastEp.troughIdx].value) / recoveryScaleRef) * 100 : 0;
    if (recoveryPct > 20) strengths.push(`récupération de ${recoveryPct >= 0 ? "+" : ""}${recoveryPct.toFixed(0)}% depuis le creux de ${monthLabel(series[lastEp.troughIdx].month)}`);
  }
  if (opts.cumulative && lastVal > avg) strengths.push(`niveau actuel (${fmt(lastVal)} FCFA) au-dessus de la moyenne de la période (${fmt(avg)} FCFA)`);
  if (strengths.length) out.push(`🟢 Points forts : ${strengths.join(" ; ")}.`);

  // 🔴 Points de vigilance, avec attribution réelle du principal choc
  const risks: string[] = [];
  if (cv > 50) risks.push(`amplitude des fluctuations élevée (variation d'environ ${cv.toFixed(0)}% autour de la moyenne)`);
  episodes.forEach((ep, i) => {
    const cats = attribRange(ep.peakIdx, ep.troughIdx);
    const causeText = cats.length ? ` — expliqué principalement par ${fmtCats(cats)}` : "";
    const swingAbs = series[ep.peakIdx].value - series[ep.troughIdx].value;
    risks.push(`baisse de ${fmt(swingAbs)} FCFA (${(ep.maxDD * 100).toFixed(0)}% de la moyenne des mouvements) entre ${monthLabel(series[ep.peakIdx].month)} et ${monthLabel(series[ep.troughIdx].month)}${causeText}`);
  });
  if (risks.length) out.push(`🔴 Points de vigilance : ${risks.join(" ; ")}.`);

  // Arc narratif chronologique — un repère tous les 2-3 mois avec la transaction/catégorie
  // dominante de ce mois-là quand elle est notable, pour ancrer le récit dans le réel.
  const arcStep = Math.max(1, Math.ceil(series.length / 6));
  const arcParts: string[] = [];
  for (let i = 0; i < series.length; i += arcStep) {
    const r = series[i];
    const top = attribMonth(r.month, 1);
    const isEpisodePoint = episodes.some((e) => e.peakIdx === i || e.troughIdx === i);
    arcParts.push(`${monthLabel(r.month)} : ${fmt(r.value)} FCFA${top.length && (isEpisodePoint || i === series.length - 1) ? ` (dominé par ${top[0][0]}, ${fmt(top[0][1])} FCFA)` : ""}`);
  }
  out.push(`Ce que racontent les chiffres, mois par mois : ${arcParts.join(" → ")}.`);

  // Diagnostic chiffré
  out.push(`Diagnostic — Tendance : ${growthScore}/10 · Stabilité : ${stabilityScore}/10 · Maîtrise des baisses : ${riskScore}/10 · Dynamique récente : ${momentumScore}/10.`);

  // Synthèse
  const overallTone = growthScore >= 6 && riskScore <= 4 ? "Une vraie capacité de progression, mais encore avec des à-coups importants à maîtriser."
    : growthScore >= 6 && riskScore >= 6 ? "Une progression solide et raisonnablement maîtrisée."
    : growthScore < 4 && riskScore <= 4 ? "Une tendance qui mérite une vraie attention, avec des baisses marquées."
    : "Une situation globalement stable, sans dynamique forte dans un sens ou l'autre.";
  out.push(`En résumé : ${overallTone}`);

  return out;
}


function generateNetWorthNarrative(fullReport: ReturnType<typeof computeNetWorthReport>, transactions: Transaction[]) {
  const rows = fullReport.rows;
  if (rows.length < 3) return null;

  const series = rows.map((r) => ({ month: r.month, value: r.netWorth }));
  const paragraphs = generateDeepNarrative(series, transactions, () => true, { subject: "la valeur nette", cumulative: true });
  if (!paragraphs.length) return null;

  const blocks: CalcDetailBlock[] = paragraphs.map((p) => ({ kind: "note", tone: p.startsWith("🔴") ? "warn" : "info", text: p }));

  // Meilleur et pire mois individuels, avec leur vraie explication (déjà calculée par
  // computeNetWorthReport à partir des transactions réelles de ce mois).
  blocks.push({ kind: "kv", rows: [
    { label: `Meilleur mois : ${monthLabel(fullReport.best.month)}`, value: `+${fmt(fullReport.best.delta)} FCFA` },
    { label: "Explication", value: fullReport.best.explanation.replace(/\.$/, "") },
    { label: `Pire mois : ${monthLabel(fullReport.worst.month)}`, value: `${fmt(fullReport.worst.delta)} FCFA`, warn: true },
    { label: "Explication", value: fullReport.worst.explanation.replace(/\.$/, "") },
  ] });

  const lastRow = rows[rows.length - 1];
  return {
    title: "Analyse de la courbe de valeur nette",
    headline: `${fmt(lastRow.netWorth)} FCFA`,
    formula: "Analyse narrative : phases, forces/risques attribués aux vraies transactions, scores de diagnostic",
    blocks,
  };
}


function ProjectionPanel({ accounts, transactions }: { accounts: Account[]; transactions: Transaction[] }) {
  const [months, setMonths] = useState(12);
  const { points, avgDelta } = projectNetWorth(months, liveNetWorthSeries(accounts, transactions));
  const startVal = points[0]?.central ?? 0;
  const endPoint = points[points.length - 1];
  const trendUp = avgDelta >= 0;

  return (
    <PanelWithHelp title="Projection de valeur nette" subtitle="Basée sur la tendance médiane des derniers relevés — bande optimiste/pessimiste robuste aux écarts ponctuels"
      collapsible defaultOpen={false}
      explain="La ligne dorée centrale prolonge la variation médiane de ta valeur nette sur tes derniers mois (jusqu'à 9). Les deux lignes pointillées (vert=optimiste, rouge=prudent) montrent une fourchette autour de cette projection, basée sur l'écart absolu médian plutôt qu'un écart-type classique — un seul mois exceptionnel (gros achat, rentrée imprévue) pèse donc beaucoup moins sur la prévision qu'avant. C'est une extrapolation statistique, pas une garantie."
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

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 22 }}>
        <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "12px 16px" }}>
          <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Aujourd'hui</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 17, fontWeight: 600, color: COLOR.ink }}>{fmt(startVal)}<span style={{ fontSize: 11, color: COLOR.inkMuted, marginLeft: 4 }}>FCFA</span></div>
        </div>
        <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.gold}`, borderRadius: 10, padding: "12px 16px" }}>
          <div style={{ fontSize: 10.5, color: COLOR.goldSoft, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Projection à {months} mois</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 17, fontWeight: 600, color: COLOR.ink, display: "flex", alignItems: "center", gap: 6 }}>
            {fmt(endPoint?.central ?? 0)}<span style={{ fontSize: 11, color: COLOR.inkMuted }}>FCFA</span>
            {trendUp ? <TrendingUp size={14} color={COLOR.emeraldSoft} /> : <TrendingDown size={14} color={COLOR.claySoft} />}
          </div>
        </div>
        <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: "12px 16px" }}>
          <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Fourchette à {months} mois</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.ink }}>
            <span style={{ color: COLOR.claySoft }}>{fmt(endPoint?.bas ?? 0)}</span>
            <span style={{ color: COLOR.inkMuted }}> — </span>
            <span style={{ color: COLOR.emeraldSoft }}>{fmt(endPoint?.haut ?? 0)}</span>
          </div>
        </div>
      </div>

      <ResponsiveContainer width="100%" height={260}>
        <AreaChart data={points} margin={{ left: 0, right: 10, top: 10 }}>
          <CartesianGrid stroke={COLOR.hairline} vertical={false} />
          <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
          <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
          <Tooltip content={<ProjectionTooltip />} />
          <Legend wrapperStyle={{ fontSize: 11.5, color: COLOR.inkMuted, paddingTop: 8 }} />
          <Area type="monotone" dataKey="range" stroke="none" fill={COLOR.gold} fillOpacity={0.14} legendType="none" />
          <Line type="monotone" dataKey="central" name="Projection centrale" stroke={COLOR.goldSoft} strokeWidth={2.5} dot={false} />
          <Line type="monotone" dataKey="haut" name="Scénario optimiste" stroke={COLOR.emeraldSoft} strokeWidth={1.5} strokeDasharray="4 3" dot={false} />
          <Line type="monotone" dataKey="bas" name="Scénario prudent" stroke={COLOR.claySoft} strokeWidth={1.5} strokeDasharray="4 3" dot={false} />
        </AreaChart>
      </ResponsiveContainer>
    </PanelWithHelp>
  );
}

// ============================================================
// PROJECTION SUR MESURE — simule la valeur nette à une échéance donnée en
// excluant certaines catégories (ex: dette bientôt soldée, déménagement
// ponctuel) et/ou en remplaçant une catégorie par un montant fixe mensuel
// (ex: nouveau loyer). Basée sur la médiane des 6 derniers mois pour rester
// robuste, avec une fourchette prudent/optimiste reflétant la vraie volatilité
// des revenus observée plutôt qu'un chiffre unique trompeur.
// ============================================================
function CustomProjectionPanel({ transactions, accounts, allCategories }: {
  transactions: Transaction[]; accounts: Account[]; allCategories: string[];
}) {
  const [excluded, setExcluded] = useState<string[]>([]);
  const [overrideCategory, setOverrideCategory] = useState("");
  const [overrideAmount, setOverrideAmount] = useState<number | "">("");
  const [monthsAhead, setMonthsAhead] = useState(24);

  const toggleExcluded = (cat: string) => setExcluded((prev) => (prev.includes(cat) ? prev.filter((c) => c !== cat) : [...prev, cat]));

  const result = useMemo(() => {
    const today = todayISO();
    const curMonth = dateToMonthKey(today);
    const lookback: string[] = [];
    let mk = prevMonthKey(curMonth);
    for (let i = 0; i < 6; i++) { lookback.push(mk); mk = prevMonthKey(mk); }

    const revByMonth: Record<string, number> = {};
    const depByMonth: Record<string, number> = {};
    transactions.forEach((t) => {
      const tmk = dateToMonthKey(t.date);
      if (!lookback.includes(tmk)) return;
      // Corrigé le 10/08/2026 : la liste "excluded" ne s'appliquait qu'aux dépenses —
      // cocher/décocher une catégorie de REVENU dans la liste n'avait donc aucun effet
      // sur la projection, alors qu'elle est proposée comme n'importe quelle catégorie.
      if (excluded.includes(t.category)) return;
      if (t.type === "Revenu") { revByMonth[tmk] = (revByMonth[tmk] || 0) + t.amount; return; }
      // Dépense : exclue si dans la liste, ou si c'est la catégorie remplacée par un montant fixe (comptée à part).
      if (overrideCategory && t.category === overrideCategory) return;
      depByMonth[tmk] = (depByMonth[tmk] || 0) + t.amount;
    });

    const revVals = lookback.map((m) => revByMonth[m] || 0);
    const depVals = lookback.map((m) => depByMonth[m] || 0);
    const medRev = median(revVals);
    const medDep = median(depVals) + (overrideCategory && overrideAmount !== "" ? Number(overrideAmount) : 0);
    const monthlyNet = medRev - medDep;

    const currentNW = totalAccountsBalance(accounts, transactions);
    const central = currentNW + monthlyNet * monthsAhead;
    const prudent = currentNW + (Math.min(...revVals) - medDep) * monthsAhead;
    const optimiste = currentNW + (Math.max(...revVals) - medDep) * monthsAhead;

    return { medRev, medDep, monthlyNet, currentNW, central, prudent, optimiste, lookback };
  }, [transactions, accounts, excluded, overrideCategory, overrideAmount, monthsAhead]);

  const targetDate = (() => {
    const d = new Date(todayISO() + "T00:00:00");
    d.setMonth(d.getMonth() + monthsAhead);
    return d.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  })();

  return (
    <PanelWithHelp title="Projection sur mesure" subtitle="Simule ta valeur nette à une échéance donnée en excluant ou en fixant certaines charges"
      collapsible defaultOpen={false}
      explain="Coche les catégories à exclure entièrement de la projection (ex : une dette qui sera soldée, un déménagement ponctuel qui ne se reproduira pas), et/ou remplace une catégorie par un montant fixe mensuel (ex : un nouveau loyer). Le calcul utilise la médiane des 6 derniers mois pour rester robuste face aux mois exceptionnels, avec une fourchette prudent/optimiste basée sur le pire et le meilleur mois de revenu réellement observés — plutôt qu'un chiffre unique qui masquerait la vraie volatilité.">
      <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>

        <div>
          <div style={{ fontSize: 12, color: COLOR.inkMuted, marginBottom: 8 }}>Exclure entièrement ces catégories de la projection :</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {allCategories.map((c) => (
              <button key={c} onClick={() => toggleExcluded(c)} style={{
                padding: "5px 12px", borderRadius: 16, fontSize: 11.5, cursor: "pointer",
                border: `1px solid ${excluded.includes(c) ? COLOR.clay : COLOR.hairline}`,
                background: excluded.includes(c) ? "rgba(193,84,63,0.14)" : "transparent",
                color: excluded.includes(c) ? COLOR.claySoft : COLOR.inkMuted,
              }}>{c}</button>
            ))}
          </div>
        </div>

        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
          <GroupedSingleSelect label="Remplacer une catégorie par un montant fixe" allLabel="— aucune —" value={overrideCategory} onChange={setOverrideCategory} options={groupedCategoryOptions(transactions)} />
          {overrideCategory && (
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Nouveau montant / mois</label>
              <input type="number" inputMode="numeric" style={{ ...inputStyle, width: 160, textAlign: "right" }} placeholder="Ex : 650000" value={overrideAmount}
                onChange={(e) => setOverrideAmount(e.target.value === "" ? "" : Number(e.target.value))} />
            </div>
          )}
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Échéance</label>
            <div style={{ display: "flex", gap: 6 }}>
              {[12, 24, 36].map((m) => (
                <button key={m} onClick={() => setMonthsAhead(m)} style={{
                  padding: "8px 14px", borderRadius: 6, fontSize: 12, cursor: "pointer",
                  border: `1px solid ${monthsAhead === m ? COLOR.gold : COLOR.hairline}`,
                  background: monthsAhead === m ? "rgba(201,162,39,0.14)" : "transparent",
                  color: monthsAhead === m ? COLOR.goldSoft : COLOR.inkMuted,
                }}>{m} mois</button>
              ))}
            </div>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12 }}>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Solde mensuel net projeté</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: result.monthlyNet >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(result.monthlyNet)} FCFA</div>
          </div>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.gold}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.goldSoft, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Valeur nette estimée — {targetDate}</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 18, fontWeight: 700, color: COLOR.ink }}>{fmt(result.central)} FCFA</div>
          </div>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Fourchette prudent — optimiste</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.ink }}>
              <span style={{ color: COLOR.claySoft }}>{fmt(result.prudent)}</span>
              <span style={{ color: COLOR.inkMuted }}> — </span>
              <span style={{ color: COLOR.emeraldSoft }}>{fmt(result.optimiste)}</span>
            </div>
          </div>
        </div>

        <div style={{ fontSize: 11.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
          Basé sur un revenu médian de {fmt(result.medRev)} FCFA/mois et des dépenses médianes de {fmt(result.medDep)} FCFA/mois (hors catégories exclues{overrideCategory ? `, "${overrideCategory}" fixée à ${overrideAmount || 0} FCFA/mois` : ""}), sur les 6 derniers mois complets. La valeur nette actuelle de départ est {fmt(result.currentNW)} FCFA.
        </div>
      </div>
    </PanelWithHelp>
  );
}

// ============================================================
// PAGE VALEUR NETTE — historique mensuel complet, rapport explicatif
// mois par mois, et plus grosses transactions individuelles.
// ============================================================
function NetWorthTab({ accounts, transactions, filters }: { accounts: Account[]; transactions: Transaction[]; filters: Filters }) {
  const [xlsState, setXlsState] = useState<"idle" | "loading" | "error">("idle");
  const [narrativeSheetOpen, setNarrativeSheetOpen] = useState(false);
  // La série complète est toujours calculée sur tout l'historique — un cumul de valeur
  // nette n'a de sens que reconstruit depuis le début. Seul l'AFFICHAGE (tableau, graphique,
  // KPI, plus grosses transactions) respecte ensuite le filtre "Du mois / Au mois" global,
  // comme sur les autres pages de l'app.
  const fullReport = useMemo(() => computeNetWorthReport(accounts, transactions), [accounts, transactions]);
  const fromKey = monthSortKey(filters.from), toKey = monthSortKey(filters.to);
  const rows = useMemo(() => fullReport.rows.filter((r) => { const k = monthSortKey(r.month); return k >= fromKey && k <= toKey; }), [fullReport, fromKey, toKey]);
  const report = useMemo(() => {
    const inRange = transactions.filter((t) => { const k = monthSortKey(dateToMonthKey(t.date)); return k >= fromKey && k <= toKey; });
    const depTx = inRange.filter((t) => t.type === "Dépense").sort((a, b) => b.amount - a.amount).slice(0, 15);
    const revTx = inRange.filter((t) => t.type === "Revenu").sort((a, b) => b.amount - a.amount).slice(0, 15);
    return { rows, depTx, revTx };
  }, [rows, transactions, fromKey, toKey]);
  const current = rows[rows.length - 1];
  const first = rows[0];
  const totalGrowth = current && first ? current.netWorth - (rows.length > 1 ? first.netWorth - first.delta : first.netWorth) : 0;
  const best = rows.length ? rows.reduce((a, b) => (b.delta > a.delta ? b : a), rows[0]) : undefined;
  const worst = rows.length ? rows.reduce((a, b) => (b.delta < a.delta ? b : a), rows[0]) : undefined;
  const [visibleMonths, setVisibleMonths] = useState(6);
  useEffect(() => { setVisibleMonths(6); }, [fromKey, toKey]);

  const exportNetWorthExcel = async () => {
    setXlsState("loading");
    try {
      const ExcelJS: any = await import(/* @vite-ignore */ "exceljs");
      const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", WHITE = "FFFFFFFF", MUTED = "FF8A9A8E";
      const wb = new ExcelJS.Workbook();
      wb.creator = "Grand Livre"; wb.created = new Date();
      const styleHeaderRow = (row: any) => {
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; c.alignment = { vertical: "middle" }; });
        row.height = 22;
      };
      const ws1 = wb.addWorksheet("Valeur nette mensuelle");
      ws1.columns = [
        { header: "Mois", key: "month", width: 16 }, { header: "Valeur nette (FCFA)", key: "nw", width: 20 },
        { header: "Variation (FCFA)", key: "delta", width: 18 }, { header: "Revenus", key: "rev", width: 16 }, { header: "Dépenses", key: "dep", width: 16 },
        { header: "Principal poste dépense", key: "topdep", width: 26 }, { header: "Principal poste revenu", key: "toprev", width: 26 }, { header: "Explication", key: "expl", width: 50 },
      ];
      styleHeaderRow(ws1.getRow(1));
      report.rows.forEach((r) => {
        const row = ws1.addRow({
          month: monthLabel(r.month), nw: r.netWorth, delta: r.delta, rev: r.revenu, dep: r.depense,
          topdep: r.topDepCat ? `${r.topDepCat[0]} (${fmt(r.topDepCat[1])})` : "", toprev: r.topRevCat ? `${r.topRevCat[0]} (${fmt(r.topRevCat[1])})` : "", expl: r.explanation,
        });
        row.getCell("nw").font = { bold: true, color: { argb: GOLD } };
        row.getCell("delta").font = { color: { argb: r.delta >= 0 ? EMERALD : CLAY } };
        row.getCell("rev").font = { color: { argb: EMERALD } };
        row.getCell("dep").font = { color: { argb: CLAY } };
        ["nw", "delta", "rev", "dep"].forEach((k) => { row.getCell(k).numFmt = "#,##0"; row.getCell(k).alignment = { horizontal: "right" }; });
        row.getCell("expl").alignment = { wrapText: true };
      });
      const totalRowNW = ws1.addRow({ month: "TOTAL (somme sur la période)", nw: "", delta: report.rows.reduce((a, r) => a + r.delta, 0), rev: report.rows.reduce((a, r) => a + r.revenu, 0), dep: report.rows.reduce((a, r) => a + r.depense, 0), topdep: "", toprev: "", expl: "" });
      totalRowNW.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      ["delta", "rev", "dep"].forEach((k) => { totalRowNW.getCell(k).numFmt = "#,##0"; totalRowNW.getCell(k).alignment = { horizontal: "right" }; });

      const ws2 = wb.addWorksheet("Plus grosses dépenses");
      ws2.columns = [{ header: "Date", key: "date", width: 12 }, { header: "Catégorie", key: "cat", width: 24 }, { header: "Sous-catégorie", key: "sub", width: 20 }, { header: "Montant (FCFA)", key: "amount", width: 18 }];
      styleHeaderRow(ws2.getRow(1));
      report.depTx.forEach((t) => {
        const row = ws2.addRow({ date: dateLabelFull(t.date), cat: t.category, sub: t.subcategory || "", amount: t.amount });
        row.getCell("amount").font = { color: { argb: CLAY }, bold: true }; row.getCell("amount").numFmt = "#,##0"; row.getCell("amount").alignment = { horizontal: "right" };
      });
      const totalRow2 = ws2.addRow({ date: `Total des ${report.depTx.length} éléments affichés`, cat: "", sub: "", amount: report.depTx.reduce((a, t) => a + t.amount, 0) });
      totalRow2.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      totalRow2.getCell("amount").numFmt = "#,##0"; totalRow2.getCell("amount").alignment = { horizontal: "right" };

      const ws3 = wb.addWorksheet("Plus gros revenus");
      ws3.columns = [{ header: "Date", key: "date", width: 12 }, { header: "Catégorie", key: "cat", width: 24 }, { header: "Sous-catégorie", key: "sub", width: 20 }, { header: "Montant (FCFA)", key: "amount", width: 18 }];
      styleHeaderRow(ws3.getRow(1));
      report.revTx.forEach((t) => {
        const row = ws3.addRow({ date: dateLabelFull(t.date), cat: t.category, sub: t.subcategory || "", amount: t.amount });
        row.getCell("amount").font = { color: { argb: EMERALD }, bold: true }; row.getCell("amount").numFmt = "#,##0"; row.getCell("amount").alignment = { horizontal: "right" };
      });
      const totalRow3nw = ws3.addRow({ date: `Total des ${report.revTx.length} éléments affichés`, cat: "", sub: "", amount: report.revTx.reduce((a, t) => a + t.amount, 0) });
      totalRow3nw.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      totalRow3nw.getCell("amount").numFmt = "#,##0"; totalRow3nw.getCell("amount").alignment = { horizontal: "right" };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `grand-livre_valeur-nette_${todayISO()}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      setXlsState("idle");
    } catch {
      setXlsState("error");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Valeur nette actuelle" value={fmt(current?.netWorth ?? 0)} tone={COLOR.goldSoft} icon={Wallet} />
        <Kpi label="Évolution totale (période)" value={fmt(totalGrowth)} tone={totalGrowth >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={totalGrowth >= 0 ? TrendingUp : TrendingDown} />
        <Kpi label="Meilleur mois" value={best ? `${monthLabel(best.month)} (+${fmt(best.delta)})` : "—"} tone={COLOR.emeraldSoft} icon={TrendingUp} />
        <Kpi label="Pire mois" value={worst ? `${monthLabel(worst.month)} (${fmt(worst.delta)})` : "—"} tone={COLOR.claySoft} icon={TrendingDown} />
      </div>

      <PanelWithHelp title="Valeur nette mensuelle" subtitle={`${monthLabel(filters.from)} — ${monthLabel(filters.to)} · explication automatique de chaque variation`}
        explain="Pour chaque mois, la variation de valeur nette est expliquée par le principal poste de dépense et le principal poste de revenu de ce mois-là — pas une simple observation du solde, mais une tentative de dire concrètement ce qui l'a fait bouger."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => setNarrativeSheetOpen(true)} style={{
              display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`,
              borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer",
            }}>
              <BookOpen size={13} /> Rapport détaillé
            </button>
            <button onClick={exportNetWorthExcel} disabled={xlsState === "loading"} style={{
              display: "flex", alignItems: "center", gap: 6, background: xlsState === "error" ? "rgba(193,84,63,0.14)" : "rgba(63,156,122,0.14)",
              border: `1px solid ${xlsState === "error" ? COLOR.clay : COLOR.emerald}`, borderRadius: 8,
              color: xlsState === "error" ? COLOR.claySoft : COLOR.emeraldSoft, padding: "7px 14px", fontSize: 12.5, cursor: xlsState === "loading" ? "default" : "pointer",
            }}>
              {xlsState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <FileSpreadsheet size={13} />}
              {xlsState === "loading" ? "Génération…" : xlsState === "error" ? "Réessayer" : "Rapport Excel"}
            </button>
          </div>
        }>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={report.rows.map((r) => ({ mois: monthLabel(r.month), valeur: r.netWorth }))} margin={{ left: 0, right: 10, top: 10 }}>
            <defs>
              <linearGradient id="nwGrad2" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={COLOR.gold} stopOpacity={0.35} />
                <stop offset="100%" stopColor={COLOR.gold} stopOpacity={0} />
              </linearGradient>
            </defs>
            <CartesianGrid stroke={COLOR.hairline} vertical={false} />
            <XAxis dataKey="mois" tick={{ fill: COLOR.inkMuted, fontSize: 9.5 }} interval={Math.floor(report.rows.length / 10)} axisLine={{ stroke: COLOR.hairline }} tickLine={false} />
            <YAxis tick={{ fill: COLOR.inkMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
            <Tooltip content={<CustomTooltip />} />
            <Area type="monotone" dataKey="valeur" name="Valeur nette" stroke={COLOR.goldSoft} strokeWidth={2.5} fill="url(#nwGrad2)" />
          </AreaChart>
        </ResponsiveContainer>

        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 18 }}>
          {report.rows.slice().reverse().slice(0, visibleMonths).map((r) => (
            <div key={r.month} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 6 }}>
                <span style={{ fontFamily: "'Fraunces', serif", fontSize: 14, color: COLOR.ink }}>{monthLabel(r.month)}</span>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.inkMuted }}>{fmt(r.netWorth)} FCFA</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, fontWeight: 700, color: r.delta >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{r.delta >= 0 ? "+" : ""}{fmt(r.delta)}</span>
                </div>
              </div>
              <div style={{ fontSize: 12, color: COLOR.inkMuted, lineHeight: 1.55 }}>{r.explanation}</div>
            </div>
          ))}
          {report.rows.length > visibleMonths && (
            <button onClick={() => setVisibleMonths((v) => v + 6)} style={{
              display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "transparent",
              border: `1px dashed ${COLOR.hairline}`, borderRadius: 10, color: COLOR.goldSoft, padding: "10px 0", fontSize: 12.5, cursor: "pointer", marginTop: 4,
            }}>
              <ChevronDown size={14} /> Afficher les {Math.min(6, report.rows.length - visibleMonths)} mois précédents ({report.rows.length - visibleMonths} restants)
            </button>
          )}
        </div>
      </PanelWithHelp>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 20 }}>
        <Panel title="Les plus grosses dépenses" subtitle={`${monthLabel(filters.from)} — ${monthLabel(filters.to)} · les 15 transactions les plus élevées`}>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {report.depTx.map((t) => (
              <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 10px", borderBottom: `1px solid ${COLOR.hairline}` }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 12.5, color: COLOR.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.category}{t.subcategory && ` · ${t.subcategory}`}</div>
                  <div style={{ fontSize: 10.5, color: COLOR.inkMuted }}>{dateLabelFull(t.date)}</div>
                </div>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, fontWeight: 600, color: COLOR.claySoft, flexShrink: 0, marginLeft: 10 }}>{fmt(t.amount)}</span>
              </div>
            ))}
            {!report.depTx.length && <EmptyState text="Aucune dépense." />}
          </div>
        </Panel>
        <Panel title="Les plus gros revenus" subtitle={`${monthLabel(filters.from)} — ${monthLabel(filters.to)} · les 15 transactions les plus élevées`}>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {report.revTx.map((t) => (
              <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 10px", borderBottom: `1px solid ${COLOR.hairline}` }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 12.5, color: COLOR.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.category}{t.subcategory && ` · ${t.subcategory}`}</div>
                  <div style={{ fontSize: 10.5, color: COLOR.inkMuted }}>{dateLabelFull(t.date)}</div>
                </div>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, fontWeight: 600, color: COLOR.emeraldSoft, flexShrink: 0, marginLeft: 10 }}>{fmt(t.amount)}</span>
              </div>
            ))}
            {!report.revTx.length && <EmptyState text="Aucun revenu." />}
          </div>
        </Panel>
      </div>
      {narrativeSheetOpen && (() => {
        const d = generateNetWorthNarrative(fullReport, transactions);
        return d ? <CalcDetailSheet open={narrativeSheetOpen} onClose={() => setNarrativeSheetOpen(false)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} /> : null;
      })()}
    </div>
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
  const [xlsState, setXlsState] = useState<"idle" | "loading" | "error">("idle");
  const [bizNarrativeOpen, setBizNarrativeOpen] = useState(false);
  const withScope = transactions.map((t) => ({ ...t, scope: categoryScope[t.category] || "Personnel" }));
  const bizRev = withScope.filter((t) => t.scope === "Business" && t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
  const bizDep = withScope.filter((t) => t.scope === "Business" && t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
  const bizMargin = bizRev - bizDep;

  const exportBusinessExcel = async () => {
    setXlsState("loading");
    try {
      const ExcelJS: any = await import(/* @vite-ignore */ "exceljs");
      const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", SUBTLE = "FF232F27", WHITE = "FFFFFFFF", MUTED = "FF8A9A8E", VIOLET = "FF7A6FB0";
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

      // ===== Feuille 1 : Résumé (les 3 montants demandés, en vert/rouge) =====
      const ws1 = wb.addWorksheet("Résumé");
      ws1.columns = [{ width: 30 }, { width: 24 }];
      ws1.mergeCells("A1:B1");
      const title = ws1.getCell("A1");
      title.value = "Grand Livre — Rapport Business / Personnel";
      title.font = { bold: true, size: 15, color: { argb: WHITE } };
      title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
      title.alignment = { vertical: "middle", indent: 1 };
      ws1.getRow(1).height = 30;

      const addSum = (label: string, value: any, color?: string, isAmount?: boolean) => {
        const r = ws1.addRow([label, value]);
        r.getCell(1).font = { color: { argb: MUTED } };
        r.getCell(2).font = { bold: true, size: 12, color: { argb: color || NAVY } };
        if (isAmount) r.getCell(2).numFmt = "#,##0 \"FCFA\"";
        r.getCell(2).alignment = { horizontal: "right" };
        return r;
      };
      ws1.addRow([]);
      const monthsInData = Array.from(new Set(transactions.map((t) => dateToMonthKey(t.date)))).sort((a, b) => monthSortKey(a) - monthSortKey(b));
      addSum("Période couverte", monthsInData.length ? `${monthLabel(monthsInData[0])} — ${monthLabel(monthsInData[monthsInData.length - 1])}` : "—");
      addSum("Généré le", dateLabelFull(todayISO()));
      ws1.addRow([]);
      addSum("Revenus Business (total)", bizRev, EMERALD, true);
      addSum("Dépenses Business (total)", bizDep, CLAY, true);
      addSum("Marge Business", bizMargin, bizMargin >= 0 ? EMERALD : CLAY, true);
      ws1.addRow([]);
      const noteRow = ws1.addRow(["Note", "La marge business ci-dessus est isolée du budget personnel — voir feuilles suivantes pour le détail justificatif."]);
      noteRow.getCell(1).font = { color: { argb: MUTED }, italic: true };
      noteRow.getCell(2).font = { color: { argb: MUTED }, italic: true };
      noteRow.getCell(2).alignment = { wrapText: true };
      ws1.mergeCells(`B${noteRow.number}:B${noteRow.number}`);

      // ===== Feuille 2 : Justificatif par catégorie (portée + revenus/dépenses de chaque catégorie) =====
      const ws2 = wb.addWorksheet("Justificatif — Catégories");
      ws2.columns = [
        { header: "Catégorie", key: "cat", width: 26 }, { header: "Portée", key: "scope", width: 14 },
        { header: "Revenus (FCFA)", key: "rev", width: 18 }, { header: "Dépenses (FCFA)", key: "dep", width: 18 },
        { header: "Solde (FCFA)", key: "solde", width: 18 }, { header: "Nb transactions", key: "count", width: 16 },
      ];
      styleHeaderRow(ws2.getRow(1));
      const byCat = allCategories.map((c) => {
        const tx = withScope.filter((t) => t.category === c);
        const rev = tx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
        const dep = tx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
        return { cat: c, scope: categoryScope[c] || "Personnel", rev, dep, solde: rev - dep, count: tx.length };
      }).filter((r) => r.count > 0).sort((a, b) => (a.scope === b.scope ? b.rev + b.dep - (a.rev + a.dep) : a.scope === "Business" ? -1 : 1));
      byCat.forEach((r, i) => {
        const prevScope = i > 0 ? byCat[i - 1].scope : null;
        if (prevScope !== null && r.scope !== prevScope) {
          const same = byCat.filter((x) => x.scope === prevScope);
          const subRow = ws2.addRow({ cat: `Sous-total ${prevScope}`, scope: "", rev: same.reduce((a, x) => a + x.rev, 0), dep: same.reduce((a, x) => a + x.dep, 0), solde: same.reduce((a, x) => a + x.solde, 0), count: same.reduce((a, x) => a + x.count, 0) });
          subRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUBTLE } }; });
          ["rev", "dep", "solde"].forEach((k) => { subRow.getCell(k).numFmt = "#,##0"; subRow.getCell(k).alignment = { horizontal: "right" }; });
          subRow.getCell("count").alignment = { horizontal: "center" };
        }
        const row = ws2.addRow(r);
        row.getCell("scope").font = { bold: true, color: { argb: r.scope === "Business" ? VIOLET : MUTED } };
        row.getCell("rev").font = { color: { argb: EMERALD } };
        row.getCell("dep").font = { color: { argb: CLAY } };
        row.getCell("solde").font = { bold: true, color: { argb: r.solde >= 0 ? EMERALD : CLAY } };
        ["rev", "dep", "solde"].forEach((k) => { row.getCell(k).numFmt = "#,##0"; row.getCell(k).alignment = { horizontal: "right" }; });
        row.getCell("count").alignment = { horizontal: "center" };
      });
      if (byCat.length) {
        const lastScope = byCat[byCat.length - 1].scope;
        const same = byCat.filter((x) => x.scope === lastScope);
        const subRow = ws2.addRow({ cat: `Sous-total ${lastScope}`, scope: "", rev: same.reduce((a, x) => a + x.rev, 0), dep: same.reduce((a, x) => a + x.dep, 0), solde: same.reduce((a, x) => a + x.solde, 0), count: same.reduce((a, x) => a + x.count, 0) });
        subRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUBTLE } }; });
        ["rev", "dep", "solde"].forEach((k) => { subRow.getCell(k).numFmt = "#,##0"; subRow.getCell(k).alignment = { horizontal: "right" }; });
        subRow.getCell("count").alignment = { horizontal: "center" };
      }
      const grandTotal2 = ws2.addRow({ cat: "TOTAL GÉNÉRAL", scope: "", rev: byCat.reduce((a, x) => a + x.rev, 0), dep: byCat.reduce((a, x) => a + x.dep, 0), solde: byCat.reduce((a, x) => a + x.solde, 0), count: byCat.reduce((a, x) => a + x.count, 0) });
      grandTotal2.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      ["rev", "dep", "solde"].forEach((k) => { grandTotal2.getCell(k).numFmt = "#,##0"; grandTotal2.getCell(k).alignment = { horizontal: "right" }; });
      grandTotal2.getCell("count").alignment = { horizontal: "center" };

      // ===== Feuille 3 : Justificatif ligne par ligne — uniquement les transactions Business =====
      const ws3 = wb.addWorksheet("Justificatif — Transactions Business");
      ws3.columns = [
        { header: "Date", key: "date", width: 12 }, { header: "Catégorie", key: "cat", width: 22 },
        { header: "Sous-catégorie", key: "sub", width: 20 }, { header: "Type", key: "type", width: 10 },
        { header: "Compte", key: "account", width: 16 }, { header: "Montant (FCFA)", key: "amount", width: 16 },
      ];
      styleHeaderRow(ws3.getRow(1));
      const bizTx = withScope.filter((t) => t.scope === "Business").sort((a, b) => a.date.localeCompare(b.date));
      let curBizMonth: string | null = null, bizMonthRev = 0, bizMonthDep = 0;
      const flushBizMonthSubtotal = () => {
        if (curBizMonth === null) return;
        const row = ws3.addRow({ date: `Sous-total ${monthLabel(curBizMonth)}`, cat: "", sub: "", type: "", account: "", amount: bizMonthRev - bizMonthDep });
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUBTLE } }; });
        row.getCell("amount").numFmt = "#,##0"; row.getCell("amount").alignment = { horizontal: "right" };
      };
      bizTx.forEach((t) => {
        const tMonth = dateToMonthKey(t.date);
        if (curBizMonth !== null && tMonth !== curBizMonth) { flushBizMonthSubtotal(); bizMonthRev = 0; bizMonthDep = 0; }
        curBizMonth = tMonth;
        if (t.type === "Revenu") bizMonthRev += t.amount; else bizMonthDep += t.amount;
        const row = ws3.addRow({ date: dateLabelFull(t.date), cat: t.category, sub: t.subcategory || "", type: t.type, account: t.account || "", amount: t.amount });
        row.getCell("type").font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY } };
        row.getCell("amount").font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY }, bold: true };
        row.getCell("amount").numFmt = "#,##0";
        row.getCell("amount").alignment = { horizontal: "right" };
      });
      flushBizMonthSubtotal();
      const totalRow = ws3.addRow({ date: "", cat: "", sub: "", type: "TOTAL GÉNÉRAL", account: "", amount: bizMargin });
      totalRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      totalRow.getCell("amount").numFmt = "#,##0";
      totalRow.getCell("amount").alignment = { horizontal: "right" };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `grand-livre_rapport-business_${todayISO()}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      setXlsState("idle");
    } catch {
      setXlsState("error");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Revenus Business (total)" value={fmt(bizRev)} tone={COLOR.emeraldSoft} icon={Briefcase} />
        <Kpi label="Dépenses Business (total)" value={fmt(bizDep)} tone={COLOR.claySoft} icon={Briefcase} />
        <Kpi label="Marge Business" value={fmt(bizMargin)} tone={bizMargin >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} onDetailClick={() => setBizNarrativeOpen(true)} />
      </div>
      <Panel title="Compte de résultat — activité (GRUNDFOS / ECO PUMP AFRIK / INVEST SGO)" subtitle="Isole l'activité commerciale du budget personnel"
        right={
          <button onClick={exportBusinessExcel} disabled={xlsState === "loading"} style={{
            display: "flex", alignItems: "center", gap: 6, background: xlsState === "error" ? "rgba(193,84,63,0.14)" : "rgba(63,156,122,0.14)",
            border: `1px solid ${xlsState === "error" ? COLOR.clay : COLOR.emerald}`, borderRadius: 8,
            color: xlsState === "error" ? COLOR.claySoft : COLOR.emeraldSoft, padding: "7px 14px", fontSize: 12.5, cursor: xlsState === "loading" ? "default" : "pointer",
          }}>
            {xlsState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <FileSpreadsheet size={13} />}
            {xlsState === "loading" ? "Génération…" : xlsState === "error" ? "Réessayer" : "Rapport Excel"}
          </button>
        }>
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
      {bizNarrativeOpen && (() => {
        const byMonth: Record<string, number> = {};
        withScope.forEach((t) => { if (t.scope === "Business") { const mk = dateToMonthKey(t.date); byMonth[mk] = (byMonth[mk] || 0) + (t.type === "Revenu" ? t.amount : -t.amount); } });
        const months = Object.keys(byMonth).sort((a, b) => monthSortKey(a) - monthSortKey(b));
        const values = months.map((m) => byMonth[m]);
        const blocks: CalcDetailBlock[] = [
          { kind: "kv", rows: [
            { label: "Marge Business totale", value: `${bizMargin >= 0 ? "+" : ""}${fmt(bizMargin)} FCFA`, strong: true },
            { label: "Revenus Business", value: `${fmt(bizRev)} FCFA` },
            { label: "Dépenses Business", value: `${fmt(bizDep)} FCFA` },
          ] },
        ];
        if (months.length >= 2) {
          const bestIdx = values.indexOf(Math.max(...values)), worstIdx = values.indexOf(Math.min(...values));
          blocks.push({ kind: "table", columns: ["Mois", "Marge Business (FCFA)"], rows: months.map((m, i) => [monthLabel(m), fmt(values[i])]) });
          blocks.push({ kind: "note", tone: values[worstIdx] < 0 ? "warn" : "info", text: `Meilleur mois : ${monthLabel(months[bestIdx])} (${values[bestIdx] >= 0 ? "+" : ""}${fmt(values[bestIdx])} FCFA). Pire mois : ${monthLabel(months[worstIdx])} (${fmt(values[worstIdx])} FCFA).` });
        }
        return <CalcDetailSheet open={bizNarrativeOpen} onClose={() => setBizNarrativeOpen(false)}
          title="Business — analyse détaillée" headline={`${bizMargin >= 0 ? "+" : ""}${fmt(bizMargin)} FCFA de marge`}
          formula="Marge Business = Revenus − Dépenses des catégories classées Business, mois par mois" blocks={blocks} />;
      })()}
    </div>
  );
}

// ============================================================
// ACTIVITÉS & RENTABILITÉ — suivi par activité réelle (pas par compte, car
// l'argent circule entre comptes) : marge, ROI et délai de remboursement estimé.
// ============================================================
function ActivitiesTab({ transactions, setTransactions, activities, setActivities, categoryActivity, setCategoryActivity, activityCapital, setActivityCapital, allCategories, categoryGroups, accounts, onNavigate, periodRange }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; activities: string[]; setActivities: (a: string[]) => void;
  categoryActivity: Record<string, string>; setCategoryActivity: (m: Record<string, string>) => void;
  activityCapital: Record<string, number>; setActivityCapital: (m: Record<string, number>) => void;
  allCategories: string[]; categoryGroups: Record<string, Group>; accounts: Account[]; onNavigate?: (tab: Tab, data?: any) => void; periodRange?: [string, string];
}) {
  const [newActivity, setNewActivity] = useState("");
  const [confirmDeleteActivity, setConfirmDeleteActivity] = useState<string | null>(null);
  const [xlsState, setXlsState] = useState<"idle" | "loading" | "error">("idle");
  const [expandedGroup, setExpandedGroup] = useState<string | null>(null); // clé: "activité::groupe"
  const [expandedGroupCat, setExpandedGroupCat] = useState<string | null>(null); // clé: "activité::groupe::catégorie"
  const [editingTx, setEditingTx] = useState<Transaction | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  const activityFor = (cat: string) => categoryActivity[cat] || "Personnel";
  const allActivities = Array.from(new Set(["Personnel", ...activities]));

  const addActivity = () => {
    const name = newActivity.trim();
    if (!name || allActivities.includes(name)) return;
    setActivities([...activities, name]);
    setNewActivity("");
  };
  const deleteActivity = (name: string) => {
    setActivities(activities.filter((a) => a !== name));
    const next = { ...categoryActivity };
    Object.keys(next).forEach((c) => { if (next[c] === name) delete next[c]; });
    setCategoryActivity(next);
    const nextCap = { ...activityCapital };
    delete nextCap[name];
    setActivityCapital(nextCap);
  };

  const stats = useMemo(() => {
    // Respecte le filtre global "Du mois/Au mois" pour la marge/dépenses/revenus de la
    // période — mais PAS pour le capital/ROI/remboursement, qui doivent rester cumulés
    // depuis le début (un remboursement de capital ne "redémarre" pas parce qu'on
    // resserre la période affichée).
    const inRange = (t: Transaction) => !periodRange || (monthSortKey(dateToMonthKey(t.date)) >= monthSortKey(periodRange[0]) && monthSortKey(dateToMonthKey(t.date)) <= monthSortKey(periodRange[1]));
    return allActivities.map((act) => {
      const txAllTime = transactions.filter((t) => activityFor(t.category) === act);
      const tx = txAllTime.filter(inRange);
      const revenus = tx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
      const depenses = tx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
      const marge = revenus - depenses;
      const margeAllTime = txAllTime.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0) - txAllTime.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
      const capital = activityCapital[act] || 0;

      // Moyenne mensuelle de marge sur les mois où l'activité a eu du mouvement, pour estimer un délai de remboursement.
      const byMonth: Record<string, number> = {};
      tx.forEach((t) => { const mk = dateToMonthKey(t.date); byMonth[mk] = (byMonth[mk] || 0) + (t.type === "Revenu" ? t.amount : -t.amount); });
      const monthlyMargins = Object.values(byMonth);
      const avgMonthly = monthlyMargins.length ? mean(monthlyMargins) : 0;

      const roiPct = capital > 0 ? (margeAllTime / capital) * 100 : null;
      const remaining = capital > 0 ? Math.max(0, capital - margeAllTime) : 0;
      const monthsLeft = capital > 0 && margeAllTime < capital && avgMonthly > 0 ? Math.ceil(remaining / avgMonthly) : null;
      const paidOff = capital > 0 && margeAllTime >= capital;

      // Répartition des dépenses par nature (Nécessaire / Productif / Non-productif)
      const groupTotals: Record<Group, number> = { "Nécessaire": 0, "Productif": 0, "Non-productif": 0, "Non classifié": 0 };
      tx.filter((t) => t.type === "Dépense").forEach((t) => { groupTotals[groupFor(t, categoryGroups)] += t.amount; });
      const depTotal = depenses || 1;
      const groups = (["Nécessaire", "Productif", "Non-productif"] as Group[])
        .map((g) => ({ group: g, value: groupTotals[g], pct: (groupTotals[g] / depTotal) * 100 }))
        .filter((g) => g.value > 0);
      if (groupTotals["Non classifié"] > 0) groups.push({ group: "Non classifié", value: groupTotals["Non classifié"], pct: (groupTotals["Non classifié"] / depTotal) * 100 });

      return { act, revenus, depenses, marge, margeAllTime, capital, roiPct, remaining, monthsLeft, paidOff, avgMonthly, count: tx.length, groups, byMonth };
    }).filter((s) => s.count > 0 || s.capital > 0);
  }, [transactions, categoryActivity, activityCapital, allActivities]);

  const [activityNarrativeKey, setActivityNarrativeKey] = useState<string | null>(null);
  const buildActivityNarrative = (act: string) => {
    const s = stats.find((x) => x.act === act);
    if (!s) return null;
    const months = Object.keys(s.byMonth).sort((a, b) => monthSortKey(a) - monthSortKey(b));
    if (months.length < 3) return null;
    const series = months.map((m) => ({ month: m, value: s.byMonth[m] }));
    const paragraphs = generateDeepNarrative(series, transactions, (t) => activityFor(t.category) === act, { subject: `la marge de ${act}`, cumulative: false });
    if (!paragraphs.length) return null;
    const blocks: CalcDetailBlock[] = paragraphs.map((p) => ({ kind: "note", tone: p.startsWith("🔴") ? "warn" : "info", text: p }));
    if (s.capital > 0) {
      blocks.push({ kind: "kv", rows: [
        { label: "Capital investi", value: `${fmt(s.capital)} FCFA` },
        { label: "ROI à date", value: `${s.roiPct?.toFixed(1)}%`, warn: (s.roiPct || 0) < 0 },
        { label: s.paidOff ? "Statut" : "Reste à rembourser", value: s.paidOff ? "Capital remboursé" : `${fmt(s.remaining)} FCFA${s.monthsLeft ? ` (~${s.monthsLeft} mois au rythme actuel)` : ""}` },
      ] });
    }
    return { title: `Activité — ${act}`, headline: `${s.marge >= 0 ? "+" : ""}${fmt(s.marge)} FCFA de marge`, formula: "Marge = Revenus − Dépenses de l'activité, mois par mois", blocks };
  };

  // Catégories qui composent un groupe (Nécessaire/Productif/Non-productif) au sein d'une activité donnée.
  const catsForGroup = (act: string, group: Group) => {
    // Corrigé le 08/08/2026 : filtrait sur "transactions" (tout l'historique) au lieu de
    // la fenêtre respectant le filtre "Du mois/Au mois" — les montants du détail ne
    // correspondaient plus au total affiché juste au-dessus dès que le filtre était resserré.
    const inRange = (t: Transaction) => !periodRange || (monthSortKey(dateToMonthKey(t.date)) >= monthSortKey(periodRange[0]) && monthSortKey(dateToMonthKey(t.date)) <= monthSortKey(periodRange[1]));
    const rows: Record<string, number> = {};
    transactions.filter((t) => t.type === "Dépense" && activityFor(t.category) === act && groupFor(t, categoryGroups) === group && inRange(t))
      .forEach((t) => { rows[t.category] = (rows[t.category] || 0) + t.amount; });
    const total = Object.values(rows).reduce((a, v) => a + v, 0) || 1;
    return Object.entries(rows).map(([name, value]) => ({ name, value, pct: (value / total) * 100 })).sort((a, b) => b.value - a.value);
  };
  const txForGroupCat = (act: string, group: Group, cat: string) => {
    const inRange = (t: Transaction) => !periodRange || (monthSortKey(dateToMonthKey(t.date)) >= monthSortKey(periodRange[0]) && monthSortKey(dateToMonthKey(t.date)) <= monthSortKey(periodRange[1]));
    return transactions.filter((t) => t.type === "Dépense" && t.category === cat && activityFor(t.category) === act && groupFor(t, categoryGroups) === group && inRange(t))
      .sort((a, b) => b.date.localeCompare(a.date));
  };

  const startEdit = (t: Transaction) => setEditingTx(t);
  const saveEdit = (t: Transaction) => setTransactions(transactions.map((x) => (x.id === t.id ? t : x)));
  const removeTx = (id: string) => setTransactions(transactions.filter((t) => t.id !== id));

  const exportActivitiesExcel = async () => {
    setXlsState("loading");
    try {
      const ExcelJS: any = await import(/* @vite-ignore */ "exceljs");
      const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", SUBTLE = "FF232F27", WHITE = "FFFFFFFF", MUTED = "FF8A9A8E";
      const wb = new ExcelJS.Workbook();
      wb.creator = "Grand Livre"; wb.created = new Date();

      const inRange = (t: Transaction) => !periodRange || (monthSortKey(dateToMonthKey(t.date)) >= monthSortKey(periodRange[0]) && monthSortKey(dateToMonthKey(t.date)) <= monthSortKey(periodRange[1]));
      const periodLabel = periodRange ? `${monthLabel(periodRange[0])} — ${monthLabel(periodRange[1])}` : "Tout l'historique";

      const styleHeaderRow = (row: any) => {
        row.eachCell((c: any) => {
          c.font = { bold: true, color: { argb: WHITE } };
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
          c.alignment = { vertical: "middle" };
        });
        row.height = 22;
      };
      const titleBanner = (ws: any, title: string) => {
        ws.mergeCells("A1:F1");
        const t = ws.getCell("A1");
        t.value = `Grand Livre — ${title}`;
        t.font = { bold: true, size: 14, color: { argb: WHITE } };
        t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
        t.alignment = { vertical: "middle", indent: 1 };
        ws.getRow(1).height = 28;
        ws.mergeCells("A2:F2");
        const s = ws.getCell("A2");
        s.value = `Période : ${periodLabel} · Généré le ${dateLabelFull(todayISO())}`;
        s.font = { italic: true, color: { argb: MUTED } };
        ws.addRow([]);
      };

      // ===== Feuille 1 : Synthèse (respecte la période filtrée) =====
      const ws1 = wb.addWorksheet("Synthèse");
      titleBanner(ws1, "Rentabilité par activité");
      const headerRow1 = ws1.addRow(["Activité", "Transactions", "Revenus (FCFA)", "Dépenses (FCFA)", "Marge période (FCFA)", "Marge/mois moy.", "Solde cumulé (tout l'historique)"]);
      styleHeaderRow(headerRow1);
      ws1.columns = [{ width: 22 }, { width: 14 }, { width: 18 }, { width: 18 }, { width: 20 }, { width: 20 }, { width: 26 }];
      const SLATE = "FF6E7FA8";
      stats.forEach((s) => {
        const row = ws1.addRow([s.act, s.count, s.revenus, s.depenses, s.marge, s.avgMonthly, s.margeAllTime]);
        row.getCell(3).font = { color: { argb: EMERALD } };
        row.getCell(4).font = { color: { argb: CLAY } };
        row.getCell(5).font = { bold: true, color: { argb: s.marge >= 0 ? EMERALD : CLAY } };
        row.getCell(6).font = { color: { argb: s.avgMonthly >= 0 ? EMERALD : CLAY } };
        row.getCell(7).font = { color: { argb: s.margeAllTime >= 0 ? EMERALD : CLAY } };
        [3, 4, 5, 6, 7].forEach((i) => { row.getCell(i).numFmt = "#,##0"; row.getCell(i).alignment = { horizontal: "right" }; });
        row.getCell(2).alignment = { horizontal: "center" };
      });
      ws1.addRow([]);
      const totRev = stats.reduce((a, s) => a + s.revenus, 0), totDep = stats.reduce((a, s) => a + s.depenses, 0);
      const totRow = ws1.addRow(["TOTAL", stats.reduce((a, s) => a + s.count, 0), totRev, totDep, totRev - totDep, "", ""]);
      totRow.eachCell((c: any) => { c.font = { bold: true }; });
      [3, 4, 5].forEach((i) => { totRow.getCell(i).numFmt = "#,##0"; totRow.getCell(i).alignment = { horizontal: "right" }; });

      // ===== Feuille 2 : Évolution mensuelle (les soldes, mois par mois) =====
      const ws3 = wb.addWorksheet("Évolution mensuelle");
      titleBanner(ws3, "Marge mensuelle par activité (soldes)");
      const allMonthsUnion = Array.from(new Set(stats.flatMap((s) => Object.keys(s.byMonth)))).sort((a, b) => monthSortKey(a) - monthSortKey(b));
      const activeStatsE = stats.filter((s) => s.count > 0);
      const headerRow3 = ws3.addRow(["Mois", ...activeStatsE.map((s) => s.act), "Total mois"]);
      styleHeaderRow(headerRow3);
      ws3.columns = [{ width: 14 }, ...activeStatsE.map(() => ({ width: 18 })), { width: 18 }];
      allMonthsUnion.forEach((m) => {
        const vals = activeStatsE.map((s) => s.byMonth[m] || 0);
        const row = ws3.addRow([monthLabel(m), ...vals, vals.reduce((a, v) => a + v, 0)]);
        for (let i = 2; i <= activeStatsE.length + 2; i++) {
          const cell = row.getCell(i);
          cell.numFmt = "#,##0"; cell.alignment = { horizontal: "right" };
          if (typeof cell.value === "number") cell.font = { color: { argb: cell.value >= 0 ? EMERALD : CLAY }, bold: i === activeStatsE.length + 2 };
        }
      });
      const colTotals = activeStatsE.map((s) => allMonthsUnion.reduce((a, m) => a + (s.byMonth[m] || 0), 0));
      const grandTotalE = colTotals.reduce((a, v) => a + v, 0);
      const totalRow3 = ws3.addRow(["TOTAL", ...colTotals, grandTotalE]);
      totalRow3.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      for (let i = 2; i <= activeStatsE.length + 2; i++) { totalRow3.getCell(i).numFmt = "#,##0"; totalRow3.getCell(i).alignment = { horizontal: "right" }; }

      // ===== Feuille 4 : Justificatif — quelle catégorie va dans quelle activité (période) =====
      const ws4 = wb.addWorksheet("Justificatif — Catégories");
      titleBanner(ws4, "Détail par catégorie");
      const headerRow4 = ws4.addRow(["Catégorie", "Activité assignée", "Revenus (FCFA)", "Dépenses (FCFA)", "Nb transactions"]);
      styleHeaderRow(headerRow4);
      ws4.columns = [{ width: 26 }, { width: 20 }, { width: 18 }, { width: 18 }, { width: 16 }];
      const byCatRows = allCategories.map((c) => {
        const tx = transactions.filter((t) => t.category === c && inRange(t));
        const rev = tx.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
        const dep = tx.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
        return { cat: c, act: activityFor(c), rev, dep, count: tx.length };
      }).filter((rr) => rr.count > 0).sort((a, b) => a.act.localeCompare(b.act) || (b.rev + b.dep) - (a.rev + a.dep));
      byCatRows.forEach((rr, i) => {
        const prevAct = i > 0 ? byCatRows[i - 1].act : null;
        if (prevAct !== null && rr.act !== prevAct) {
          const sameAct = byCatRows.filter((x) => x.act === prevAct);
          const subRow = ws4.addRow([`Sous-total ${prevAct}`, "", sameAct.reduce((a, x) => a + x.rev, 0), sameAct.reduce((a, x) => a + x.dep, 0), sameAct.reduce((a, x) => a + x.count, 0)]);
          subRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF232F27" } }; });
          [3, 4].forEach((k) => { subRow.getCell(k).numFmt = "#,##0"; subRow.getCell(k).alignment = { horizontal: "right" }; });
          subRow.getCell(5).alignment = { horizontal: "center" };
        }
        const row = ws4.addRow([rr.cat, rr.act, rr.rev, rr.dep, rr.count]);
        row.getCell(2).font = { bold: true, color: { argb: rr.act === "Personnel" ? MUTED : GOLD } };
        row.getCell(3).font = { color: { argb: EMERALD } };
        row.getCell(4).font = { color: { argb: CLAY } };
        [3, 4].forEach((k) => { row.getCell(k).numFmt = "#,##0"; row.getCell(k).alignment = { horizontal: "right" }; });
        row.getCell(5).alignment = { horizontal: "center" };
      });
      if (byCatRows.length) {
        const lastAct = byCatRows[byCatRows.length - 1].act;
        const sameAct = byCatRows.filter((x) => x.act === lastAct);
        const subRow = ws4.addRow([`Sous-total ${lastAct}`, "", sameAct.reduce((a, x) => a + x.rev, 0), sameAct.reduce((a, x) => a + x.dep, 0), sameAct.reduce((a, x) => a + x.count, 0)]);
        subRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF232F27" } }; });
        [3, 4].forEach((k) => { subRow.getCell(k).numFmt = "#,##0"; subRow.getCell(k).alignment = { horizontal: "right" }; });
        subRow.getCell(5).alignment = { horizontal: "center" };
      }
      const grandTotalRow4b = ws4.addRow(["TOTAL GÉNÉRAL", "", byCatRows.reduce((a, x) => a + x.rev, 0), byCatRows.reduce((a, x) => a + x.dep, 0), byCatRows.reduce((a, x) => a + x.count, 0)]);
      grandTotalRow4b.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      [3, 4].forEach((k) => { grandTotalRow4b.getCell(k).numFmt = "#,##0"; grandTotalRow4b.getCell(k).alignment = { horizontal: "right" }; });
      grandTotalRow4b.getCell(5).alignment = { horizontal: "center" };

      // ===== Feuille 5 : Justificatif ligne par ligne (période filtrée) =====
      const ws5 = wb.addWorksheet("Justificatif — Transactions");
      titleBanner(ws5, "Détail transaction par transaction");
      const headerRow5 = ws5.addRow(["Activité", "Date", "Catégorie", "Sous-catégorie", "Type", "Compte", "Montant (FCFA)"]);
      styleHeaderRow(headerRow5);
      ws5.columns = [{ width: 18 }, { width: 12 }, { width: 22 }, { width: 20 }, { width: 10 }, { width: 16 }, { width: 16 }];
      const sortedTx = transactions.filter(inRange).slice().sort((a, b) => activityFor(a.category).localeCompare(activityFor(b.category)) || a.date.localeCompare(b.date));
      let curAct: string | null = null, actRev = 0, actDep = 0;
      const flushActSubtotal = () => {
        if (curAct === null) return;
        const row = ws5.addRow([`Sous-total ${curAct}`, "", "", "", "", "", actRev - actDep]);
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SUBTLE } }; });
        row.getCell(7).numFmt = "#,##0"; row.getCell(7).alignment = { horizontal: "right" };
        ws5.addRow([]);
      };
      sortedTx.forEach((t) => {
        const act = activityFor(t.category);
        if (act !== curAct) { flushActSubtotal(); curAct = act; actRev = 0; actDep = 0; }
        if (t.type === "Revenu") actRev += t.amount; else actDep += t.amount;
        const row = ws5.addRow([act, dateLabelFull(t.date), t.category, t.subcategory || "", t.type, t.account || "", t.amount]);
        row.getCell(5).font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY } };
        row.getCell(7).font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY } };
        row.getCell(7).numFmt = "#,##0"; row.getCell(7).alignment = { horizontal: "right" };
      });
      flushActSubtotal();
      const grandTotalRow5 = ws5.addRow(["TOTAL GÉNÉRAL", "", "", "", "", "", sortedTx.reduce((a, t) => a + (t.type === "Revenu" ? t.amount : -t.amount), 0)]);
      grandTotalRow5.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      grandTotalRow5.getCell(7).numFmt = "#,##0"; grandTotalRow5.getCell(7).alignment = { horizontal: "right" };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `grand-livre_rapport-activites_${todayISO()}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      setXlsState("idle");
    } catch (e) {
      console.error(e);
      setXlsState("error");
    }
  };

  const [pdfState, setPdfState] = useState<"idle" | "loading" | "error">("idle");
  const exportActivitiesNarrativePdf = async () => {
    setPdfState("loading");
    try {
      const [jsPDFModule, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      // jspdf@2.5.x expose le vrai constructeur sur l'export NOMMÉ "jsPDF", pas sur
      // "default" (qui résout vers un objet inutilisable selon le mode d'interop
      // CJS/ESM) — cause du "Réessayer" systématique sur tous les boutons PDF de l'app.
      const jsPDF: any = (jsPDFModule as any).jsPDF || (jsPDFModule as any).default;
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const periodLabel = periodRange ? `${monthLabel(periodRange[0])} — ${monthLabel(periodRange[1])}` : "Tout l'historique";
      const ps = (s: any): string => String(s).replace(/[\u202F\u00A0]/g, " ").replace(/→/g, "->").replace(/—/g, "-").replace(/…/g, "...");

      doc.setFillColor(26, 43, 76);
      doc.rect(0, 0, pageWidth, 34, "F");
      doc.setFontSize(17); doc.setTextColor(255, 255, 255);
      doc.text("Grand Livre — Analyse narrative par activité", 14, 16);
      doc.setFontSize(9); doc.setTextColor(200, 210, 225);
      doc.text(ps(`Période : ${periodLabel} · Généré le ${dateLabelFull(todayISO())}`), 14, 24);

      let y = 44;
      stats.filter((s) => s.count > 0).forEach((s) => {
        const months = Object.keys(s.byMonth).sort((a, b) => monthSortKey(a) - monthSortKey(b));
        const series = months.map((m) => ({ month: m, value: s.byMonth[m] }));
        const paragraphs = generateDeepNarrative(series, transactions, (t) => activityFor(t.category) === s.act, { subject: `la marge de ${s.act}`, cumulative: false });

        if (y > pageHeight - 60) { doc.addPage(); y = 20; }

        doc.setFillColor(201, 162, 39);
        doc.rect(14, y, 3, 12, "F");
        doc.setFontSize(13); doc.setTextColor(26, 43, 76); doc.setFont("helvetica", "bold");
        doc.text(ps(s.act), 20, y + 8.5);
        doc.setFont("helvetica", "normal");
        y += 18;

        doc.setFillColor(245, 247, 245);
        doc.roundedRect(14, y, pageWidth - 28, 16, 2, 2, "F");
        doc.setFontSize(8); doc.setTextColor(90, 100, 95);
        doc.text("MARGE (PÉRIODE)", 20, y + 6);
        doc.setFontSize(12); doc.setTextColor(26, 26, 26); doc.setFont("helvetica", "bold");
        doc.text(ps(`${s.marge >= 0 ? "+" : ""}${fmt(s.marge)} FCFA  (${fmt(s.avgMonthly)} FCFA/mois en moyenne)`), 20, y + 12.5);
        doc.setFont("helvetica", "normal");
        y += 22;

        const allParagraphs: string[] = [`Sur ${periodLabel}, ${s.act} a généré ${fmt(s.revenus)} FCFA de revenus pour ${fmt(s.depenses)} FCFA de dépenses.`, ...paragraphs];
        if (s.capital > 0) {
          allParagraphs.push(`Capital investi : ${fmt(s.capital)} FCFA (suivi sur tout l'historique). ${s.paidOff ? "Investissement entièrement remboursé." : `Il reste ${fmt(s.remaining)} FCFA à rembourser${s.monthsLeft ? `, soit environ ${s.monthsLeft} mois au rythme actuel` : ""}.`} ROI à date : ${s.roiPct?.toFixed(1)}%.`);
        }

        doc.setFontSize(9.5); doc.setTextColor(40, 40, 40);
        allParagraphs.forEach((p) => {
          const isFlag = p.startsWith("🟢") || p.startsWith("🔴");
          if (isFlag) doc.setFont("helvetica", "bold"); else doc.setFont("helvetica", "normal");
          const lines = doc.splitTextToSize(ps(p), pageWidth - 34);
          if (y + lines.length * 5 > pageHeight - 16) { doc.addPage(); y = 20; }
          doc.text(lines, 20, y);
          y += lines.length * 5 + 5;
        });
        doc.setFont("helvetica", "normal");
        y += 6;
      });

      doc.save(`grand-livre_narratif-activites_${todayISO()}.pdf`);
      setPdfState("idle");
    } catch (e) {
      console.error(e);
      setPdfState("error");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Rentabilité par activité" subtitle="Basée sur la catégorie de chaque transaction, pas sur le compte — l'argent circule souvent entre comptes"
        explain="Chaque catégorie est rattachée à une activité (Mazda, GRUNDFOS, Personnel…) plutôt qu'à un compte, parce que les comptes se mélangent dans la réalité (ex : un salaire épuisé qui pousse à puiser sur Petty Cash ou Revenus MAZDA). La marge affichée est cumulée depuis la toute première transaction de cette activité. Si tu renseignes un capital investi (ex : prix d'achat de la voiture), l'app calcule un ROI et estime, au rythme actuel, dans combien de mois l'investissement sera remboursé."
        right={
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={exportActivitiesExcel} disabled={xlsState === "loading"} style={{
              display: "flex", alignItems: "center", gap: 6, background: xlsState === "error" ? "rgba(193,84,63,0.14)" : "rgba(63,156,122,0.14)",
              border: `1px solid ${xlsState === "error" ? COLOR.clay : COLOR.emerald}`, borderRadius: 8,
              color: xlsState === "error" ? COLOR.claySoft : COLOR.emeraldSoft, padding: "7px 14px", fontSize: 12.5, cursor: xlsState === "loading" ? "default" : "pointer",
            }}>
              {xlsState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <FileSpreadsheet size={13} />}
              {xlsState === "loading" ? "Génération…" : xlsState === "error" ? "Réessayer" : "Rapport Excel"}
            </button>
            <button onClick={exportActivitiesNarrativePdf} disabled={pdfState === "loading"} style={{
              display: "flex", alignItems: "center", gap: 6, background: pdfState === "error" ? "rgba(193,84,63,0.14)" : "rgba(201,162,39,0.14)",
              border: `1px solid ${pdfState === "error" ? COLOR.clay : COLOR.gold}`, borderRadius: 8,
              color: pdfState === "error" ? COLOR.claySoft : COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: pdfState === "loading" ? "default" : "pointer",
            }}>
              {pdfState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <BookOpen size={13} />}
              {pdfState === "loading" ? "Génération…" : pdfState === "error" ? "Réessayer" : "Rapport narratif PDF"}
            </button>
          </div>
        }>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {stats.map((s) => (
            <div key={s.act} style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 12, padding: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                <div style={{ fontFamily: "'Fraunces', serif", fontSize: 15, color: COLOR.ink, display: "flex", alignItems: "center", gap: 8 }}>
                  <Rocket size={15} color={COLOR.goldSoft} /> {s.act}
                  <span style={{ fontSize: 11, color: COLOR.inkMuted, fontFamily: "'IBM Plex Mono', monospace" }}>({s.count} tx)</span>
                  <CalcDetailIcon onClick={() => setActivityNarrativeKey(s.act)} />
                </div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 700, color: s.marge >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>
                  {s.marge >= 0 ? "+" : ""}{fmt(s.marge)} FCFA
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 10, marginBottom: s.capital > 0 ? 14 : 0 }}>
                <div>
                  <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Revenus</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.emeraldSoft }}>{fmt(s.revenus)}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Dépenses</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.claySoft }}>{fmt(s.depenses)}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Marge / mois (moy.)</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: s.avgMonthly >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(s.avgMonthly)}</div>
                </div>
                {s.capital > 0 && (
                  <div>
                    <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>ROI</div>
                    <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: (s.roiPct || 0) >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{s.roiPct?.toFixed(0)}%</div>
                  </div>
                )}
              </div>
              {s.capital > 0 && (
                <div>
                  <div style={{ height: 8, background: COLOR.hairline, borderRadius: 4, overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${Math.min(100, (s.margeAllTime / s.capital) * 100).toFixed(1)}%`, background: s.paidOff ? COLOR.emerald : COLOR.gold, transition: "width 0.3s" }} />
                  </div>
                  <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 6 }}>
                    Capital investi : {fmt(s.capital)} FCFA (suivi sur tout l'historique, indépendant du filtre de période) · {s.paidOff
                      ? <span style={{ color: COLOR.emeraldSoft }}>investissement remboursé ✓</span>
                      : s.monthsLeft !== null
                        ? `reste ${fmt(s.remaining)} FCFA — remboursement estimé dans ~${s.monthsLeft} mois au rythme actuel`
                        : `reste ${fmt(s.remaining)} FCFA — rythme actuel insuffisant pour estimer un délai`}
                  </div>
                </div>
              )}
              {s.groups.length > 0 && (
                <div style={{ marginTop: s.capital > 0 ? 14 : 4, paddingTop: 14, borderTop: `1px solid ${COLOR.hairline}` }}>
                  <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>
                    Répartition des dépenses par nature
                  </div>
                  <div style={{ display: "flex", height: 8, borderRadius: 4, overflow: "hidden", marginBottom: 10 }}>
                    {s.groups.map((g) => (
                      <div key={g.group} style={{ width: `${g.pct}%`, background: groupColor[g.group] }} title={`${g.group} : ${g.pct.toFixed(0)}%`} />
                    ))}
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                    {s.groups.map((g) => {
                      const gKey = `${s.act}::${g.group}`;
                      const gOpen = expandedGroup === gKey;
                      const cats = gOpen ? catsForGroup(s.act, g.group) : [];
                      return (
                        <div key={g.group}>
                          <div onClick={() => setExpandedGroup(gOpen ? null : gKey)} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, padding: "5px 0", cursor: "pointer" }}>
                            <ChevronDown size={11} color={COLOR.inkMuted} style={{ transform: gOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s", flexShrink: 0 }} />
                            <span style={{ width: 8, height: 8, borderRadius: "50%", background: groupColor[g.group], flexShrink: 0 }} />
                            <span style={{ color: COLOR.inkMuted }}>{g.group}</span>
                            <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.ink, fontWeight: 600 }}>{fmt(g.value)} FCFA</span>
                            <span style={{ color: COLOR.inkMuted }}>({g.pct.toFixed(0)}%)</span>
                          </div>
                          {gOpen && (
                            <div style={{ padding: "2px 0 8px 25px", display: "flex", flexDirection: "column", gap: 2 }}>
                              {cats.map((c) => {
                                const cKey = `${gKey}::${c.name}`;
                                const cOpen = expandedGroupCat === cKey;
                                const tx = cOpen ? txForGroupCat(s.act, g.group, c.name) : [];
                                return (
                                  <div key={c.name}>
                                    <div onClick={() => setExpandedGroupCat(cOpen ? null : cKey)} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "5px 0", cursor: "pointer" }}>
                                      <span style={{ color: COLOR.inkMuted, display: "flex", alignItems: "center", gap: 5 }}>
                                        <ChevronDown size={10} color={COLOR.inkMuted} style={{ transform: cOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s" }} />
                                        {c.name}
                                      </span>
                                      <span style={{ display: "flex", gap: 10 }}>
                                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.ink }}>{fmt(c.value)} FCFA</span>
                                        <span style={{ color: COLOR.inkMuted, minWidth: 30, textAlign: "right" }}>{c.pct.toFixed(0)}%</span>
                                      </span>
                                    </div>
                                    {cOpen && (
                                      <div style={{ padding: "2px 0 8px 18px", display: "flex", flexDirection: "column", gap: 6 }}>
                                        {tx.map((t) => (
                                          <div key={t.id} style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 8, padding: "8px 10px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                                            <div style={{ fontSize: 12, color: COLOR.ink, fontFamily: "'IBM Plex Mono', monospace" }}>
                                              {dateLabelFull(t.date)}
                                              {t.subcategory && <span style={{ color: COLOR.inkMuted }}> · {t.subcategory}</span>}
                                              {t.account && <span style={{ color: COLOR.inkMuted }}> · {t.account}</span>}
                                            </div>
                                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, fontWeight: 600, color: COLOR.claySoft }}>{fmt(t.amount)} FCFA</span>
                                              <button onClick={() => startEdit(t)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={12} /></button>
                                              <button onClick={() => setConfirmDeleteId(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={12} /></button>
                                            </div>
                                          </div>
                                        ))}
                                        {!tx.length && <div style={{ fontSize: 12, color: COLOR.inkMuted }}>Aucune transaction.</div>}
                                      </div>
                                    )}
                                  </div>
                                );
                              })}
                              {!cats.length && <div style={{ fontSize: 12, color: COLOR.inkMuted, padding: "5px 0" }}>Aucune catégorie.</div>}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          ))}
          {!stats.length && <EmptyState text="Aucune activité avec des transactions pour l'instant." />}
        </div>
      </PanelWithHelp>

      <CollapsibleSection title="Gérer les activités" subtitle="Ajoute une activité, et renseigne un capital investi si tu veux suivre un retour sur investissement (ex : achat d'un véhicule)">
        <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
          <input style={{ ...inputStyle, flex: 1 }} placeholder="Nouvelle activité (ex : Vente Pompe)" value={newActivity}
            onChange={(e) => setNewActivity(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") addActivity(); }} />
          <button onClick={addActivity} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 6, color: COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
            <Plus size={13} /> Ajouter
          </button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {activities.map((a) => (
            <div key={a} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${COLOR.hairline}`, gap: 10, flexWrap: "wrap" }}>
              <span style={{ fontSize: 13, color: COLOR.ink }}>{a}</span>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <label style={{ fontSize: 11, color: COLOR.inkMuted }}>Capital investi</label>
                <input type="number" inputMode="numeric" style={{ ...inputStyle, width: 130, textAlign: "right" }}
                  value={activityCapital[a] || ""} placeholder="0"
                  onChange={(e) => setActivityCapital({ ...activityCapital, [a]: Number(e.target.value) })} />
                <button onClick={() => setConfirmDeleteActivity(a)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
              </div>
            </div>
          ))}
          {!activities.length && <EmptyState text="Aucune activité personnalisée. « Personnel » reste le fourre-tout par défaut." />}
        </div>
      </CollapsibleSection>

      <CollapsibleSection title="Assigner chaque catégorie à une activité" subtitle="Par catégorie, indépendamment du compte utilisé pour la transaction">
        {onNavigate && (
          <button onClick={() => onNavigate("gestioncategories")} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", fontSize: 11.5, padding: "0 0 12px 0" }}>
            Une catégorie manque ou est mal nommée ? La créer/renommer dans Gestion des catégories <ArrowRight size={11} />
          </button>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {allCategories.map((c) => (
            <div key={c} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <span style={{ fontSize: 12.5 }}>{c}</span>
              <select value={activityFor(c)} onChange={(e) => setCategoryActivity({ ...categoryActivity, [c]: e.target.value })}
                style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: activityFor(c) === "Personnel" ? COLOR.inkMuted : COLOR.goldSoft, padding: "5px 8px", fontSize: 11.5, width: 160 }}>
                {allActivities.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
            </div>
          ))}
        </div>
      </CollapsibleSection>

      <ConfirmDialog
        open={!!confirmDeleteActivity}
        title={`Supprimer l'activité "${confirmDeleteActivity}" ?`}
        message="Les catégories qui y étaient rattachées repasseront automatiquement en « Personnel ». Le capital investi renseigné sera perdu."
        onConfirm={() => { if (confirmDeleteActivity) deleteActivity(confirmDeleteActivity); setConfirmDeleteActivity(null); }}
        onCancel={() => setConfirmDeleteActivity(null)}
      />
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
      {activityNarrativeKey && (() => {
        const d = buildActivityNarrative(activityNarrativeKey);
        return d ? <CalcDetailSheet open={!!activityNarrativeKey} onClose={() => setActivityNarrativeKey(null)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} /> : null;
      })()}
    </div>
  );
}


// ============================================================
// CHARGES FIXES vs VARIABLES — pilote classifyCharges, laisse l'utilisateur
// ajuster chaque poste, exporte un rapport, et donne un "reste à vivre" estimé.
// ============================================================
// Champ de saisie du montant "Fixe" avec enregistrement explicite : la frappe met à jour
// un brouillon local seulement, rien n'est enregistré tant que l'utilisateur ne clique
// pas ✓ (ou n'appuie pas sur Entrée) — évite l'ambiguïté de savoir si une valeur tapée a
// bien été prise en compte, avec une confirmation visuelle claire au moment où ça l'est.
function FixedAmountInput({ value, onSave, onRecalculate, recalcTitle }: {
  value: number; onSave: (v: number) => void; onRecalculate: () => void; recalcTitle: string;
}) {
  const [draft, setDraft] = useState(String(Math.round(value)));
  const [saved, setSaved] = useState(false);
  useEffect(() => { setDraft(String(Math.round(value))); }, [value]);
  const dirty = Number(draft) !== Math.round(value);
  const commit = () => {
    if (!dirty || draft === "") return;
    onSave(Number(draft));
    setSaved(true);
    setTimeout(() => setSaved(false), 1200);
  };
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
      <input type="number" inputMode="numeric" style={{ ...inputStyle, textAlign: "right", borderColor: dirty ? COLOR.gold : COLOR.hairline }}
        value={draft} onChange={(e) => setDraft(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") commit(); }} />
      <button onClick={commit} disabled={!dirty} title={dirty ? "Enregistrer ce montant" : saved ? "Enregistré" : "Aucun changement à enregistrer"} style={{
        background: saved ? "rgba(63,156,122,0.18)" : "transparent", border: "none", cursor: dirty ? "pointer" : "default",
        color: saved ? COLOR.emeraldSoft : dirty ? COLOR.goldSoft : COLOR.inkMuted, display: "flex", flexShrink: 0, padding: 4, borderRadius: 4, opacity: dirty || saved ? 1 : 0.4,
      }}>
        <Check size={13} />
      </button>
      <button onClick={onRecalculate} title={recalcTitle} style={{ background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", display: "flex", flexShrink: 0, padding: 4 }}>
        <RotateCcw size={13} />
      </button>
    </div>
  );
}

function ChargesTab({ transactions, chargeOverrides, setChargeOverrides, includeGrundfosVoiture, setIncludeGrundfosVoiture, onNavigate, periodRange }: {
  transactions: Transaction[]; chargeOverrides: Record<string, ChargeOverride>; setChargeOverrides: (o: Record<string, ChargeOverride>) => void;
  includeGrundfosVoiture: boolean; setIncludeGrundfosVoiture: (b: boolean) => void; onNavigate?: (tab: Tab, data?: any) => void; periodRange?: [string, string];
}) {
  const [xlsState, setXlsState] = useState<"idle" | "loading" | "error">("idle");
  const result = useMemo(() => classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, 6, periodRange), [transactions, chargeOverrides, includeGrundfosVoiture, periodRange]);
  // Détail mois par mois par poste — pas renvoyé par classifyCharges (qui n'expose que les
  // agrégats), recalculé ici uniquement pour l'affichage de la fiche de détail au clic.
  const posteMonthlyDetail = useMemo(() => {
    const byPosteMonth: Record<string, Record<string, number>> = {};
    transactions.forEach((t) => {
      if (t.type !== "Dépense") return;
      if (!includeGrundfosVoiture && GRUNDFOS_VOITURE_CATEGORIES.includes(t.category)) return;
      const tmk = dateToMonthKey(t.date);
      if (!result.lookback.includes(tmk)) return;
      const poste = EXPAND_SUBCATS_FOR_CHARGES[t.category] ? `${t.category}::${t.subcategory || "(non précisé)"}` : t.category;
      byPosteMonth[poste] = byPosteMonth[poste] || {};
      byPosteMonth[poste][tmk] = (byPosteMonth[poste][tmk] || 0) + t.amount;
    });
    return byPosteMonth;
  }, [transactions, includeGrundfosVoiture, result.lookback]);
  const [calcDetailPoste, setCalcDetailPoste] = useState<string | null>(null);
  const [kpiDetailKey, setKpiDetailKey] = useState<"revenu" | "fixe" | "variable" | "reste" | null>(null);
  const [scheduleSheetPoste, setScheduleSheetPoste] = useState<string | null>(null);

  const modeLabel: Record<ChargeMode, string> = { fixe: "Fixe", variable: "Variable régulière", occasionnelle: "Occasionnelle", exclu: "Exclu" };
  const modeColor: Record<ChargeMode, string> = { fixe: COLOR.emeraldSoft, variable: COLOR.goldSoft, occasionnelle: COLOR.inkMuted, exclu: COLOR.claySoft };

  const setOverride = (poste: string, patch: Partial<ChargeOverride>) => {
    const current = chargeOverrides[poste] || { mode: "auto" as const };
    setChargeOverrides({ ...chargeOverrides, [poste]: { ...current, ...patch } });
  };

  const buildChargeDetail = (poste: string): { title: string; headline: string; formula: string; blocks: CalcDetailBlock[] } | null => {
    const r = result.rows.find((row) => row.poste === poste);
    if (!r) return null;
    const schedule = chargeOverrides[poste]?.schedule;
    const monthly = posteMonthlyDetail[poste] || {};
    const rows = result.lookback.map((m) => [monthLabel(m), monthly[m] ? fmt(monthly[m]) : "—"]);
    const usesAllTimeFallback = r.overridden && r.present < 3 && !!chargeOverrides[poste] && chargeOverrides[poste].mode === "fixe" && chargeOverrides[poste].amount === undefined;
    const usesPresentFallback = !usesAllTimeFallback && r.mean === 0 && r.meanPresent > 0;
    const blocks: CalcDetailBlock[] = [
      { kind: "kv", rows: [
        { label: "Mode retenu", value: modeLabel[r.mode], strong: true },
        { label: "Origine", value: schedule?.length ? "historique de montants" : r.overridden ? "réglage manuel" : "automatique" },
        ...(schedule?.length ? [] : [
          { label: "Mois présents (classification, tout l'historique)", value: `${r.classPresent} / ${r.classTotal}` },
          { label: "Mois présents (fenêtre affichée)", value: `${r.present} / ${result.lookback.length}` },
          { label: "Moyenne (toute la fenêtre affichée, absences comptées comme 0)", value: `${fmt(r.mean)} FCFA` },
          { label: "Médiane (pour référence)", value: `${fmt(r.median)} FCFA` },
          ...(usesPresentFallback ? [{ label: "Moyenne (mois présents uniquement)", value: `${fmt(r.meanPresent)} FCFA`, warn: true }] : []),
          ...(usesAllTimeFallback ? [{ label: "Moyenne (tout l'historique, hors fenêtre)", value: `${fmt(r.meanAllTime)} FCFA`, warn: true }] : []),
          { label: "Coefficient de variation (CV)", value: r.cv !== null ? `${r.cv.toFixed(0)}%` : "—" },
        ]),
        { label: "Montant retenu pour les calculs", value: `${fmt(r.amount)} FCFA`, strong: true },
      ] },
      ...(schedule?.length ? [{ kind: "table" as const, columns: ["Depuis", "Jusqu'à", "Montant (FCFA)"], rows: [...schedule].sort((a, b) => monthSortKey(a.from) - monthSortKey(b.from)).map((s) => [monthLabel(s.from), s.to ? monthLabel(s.to) : "En cours", fmt(s.amount)]) }] : []),
      ...(usesAllTimeFallback ? [{ kind: "note" as const, tone: "warn" as const, text: `Seulement ${r.present} mois avec des données dans la période affichée — trop peu pour une moyenne fiable (un seul mois atypique la définirait entièrement). Le montant retenu s'appuie automatiquement sur la moyenne de TOUT l'historique disponible de ce poste (${fmt(r.meanAllTime)} FCFA), plus stable pour représenter un montant censé être fixe. Tape un montant explicite dans le champ si ce chiffre ne te convient pas.` }] : []),
      ...(usesPresentFallback ? [{ kind: "note" as const, tone: "warn" as const, text: `La moyenne sur toute la fenêtre tombe à 0 FCFA parce que ce poste est absent la plupart des mois (${r.present}/${result.lookback.length} présents) — souvent le signe d'une charge apparue récemment. Le montant retenu utilise automatiquement la moyenne des mois présents (${fmt(r.meanPresent)} FCFA) à la place.` }] : []),
      { kind: "note", tone: "info", text: schedule?.length ? "Historique de montants actif : le montant retenu est celui du segment couvrant le mois en cours (ou le plus récent si aucun ne le couvre). Modifiable via le bouton horloge à côté du champ." : "Règle automatique (mode) : \"Fixe\" si présent sur au moins 5/6 de la période avec CV < 20%. \"Variable régulière\" si présent sur au moins 4/6 mais montant qui fluctue davantage. \"Occasionnelle\" sinon. Un réglage manuel prime toujours sur cette règle. Le montant retenu utilise la moyenne sur la période, mois d'absence comptés comme 0 (sauf repli ci-dessus)." },
      { kind: "table", columns: ["Mois", "Montant (FCFA)"], rows },
    ];
    return {
      title: poste.replace("::", " · "), headline: `${fmt(r.amount)} FCFA/mois`,
      formula: schedule?.length ? "Historique de montants — segment en cours" : r.overridden ? "Montant fixé manuellement" : `${modeLabel[r.mode]} — moyenne sur ${result.lookback.length} mois`,
      blocks,
    };
  };

  const revByMonthDetail = useMemo(() => {
    const byMonth: Record<string, number> = {};
    transactions.forEach((t) => {
      if (t.type !== "Revenu") return;
      if (!includeGrundfosVoiture && GRUNDFOS_VOITURE_LINKED_REVENUE.includes(t.category)) return;
      const tmk = dateToMonthKey(t.date);
      if (!result.lookback.includes(tmk)) return;
      byMonth[tmk] = (byMonth[tmk] || 0) + t.amount;
    });
    return byMonth;
  }, [transactions, includeGrundfosVoiture, result.lookback]);

  const buildKpiDetail = (key: "revenu" | "fixe" | "variable" | "reste"): { title: string; headline: string; formula: string; blocks: CalcDetailBlock[] } => {
    if (key === "revenu") {
      return {
        title: "Revenu moyen / mois", headline: `${fmt(result.avgRevenu)} FCFA`,
        formula: `Somme des revenus (${includeGrundfosVoiture ? "GRUNDFOS inclus" : "Petty Cash exclu, lié à GRUNDFOS"}) / ${result.lookback.length} mois`,
        blocks: [{ kind: "table", columns: ["Mois", "Revenu (FCFA)"], rows: result.lookback.map((m) => [monthLabel(m), fmt(revByMonthDetail[m] || 0)]) }],
      };
    }
    if (key === "fixe") {
      const rows = result.rows.filter((r) => r.mode === "fixe").sort((a, b) => b.amount - a.amount);
      return {
        title: "Charges fixes / mois", headline: `${fmt(result.totalFixe)} FCFA`,
        formula: "Somme des postes classés \"Fixe\"",
        blocks: [{ kind: "table", columns: ["Poste", "Montant/mois (FCFA)", "Origine"], rows: rows.map((r) => [r.poste.replace("::", " · "), fmt(r.amount), r.overridden ? "manuel" : "auto"]) }],
      };
    }
    if (key === "variable") {
      const rows = result.rows.filter((r) => r.mode === "variable").sort((a, b) => b.amount - a.amount);
      return {
        title: "Variables régulières / mois", headline: `${fmt(result.totalVariable)} FCFA`,
        formula: "Somme des postes classés \"Variable régulière\"",
        blocks: [{ kind: "table", columns: ["Poste", "Montant/mois (FCFA)", "Origine"], rows: rows.map((r) => [r.poste.replace("::", " · "), fmt(r.amount), r.overridden ? "manuel" : "auto"]) }],
      };
    }
    return {
      title: "Reste à vivre estimé / mois", headline: `${fmt(result.resteAVivre)} FCFA`,
      formula: "Revenu moyen − Charges fixes − Variables régulières",
      blocks: [{ kind: "kv", rows: [
        { label: "Revenu moyen", value: `${fmt(result.avgRevenu)} FCFA` },
        { label: "− Charges fixes", value: `${fmt(result.totalFixe)} FCFA` },
        { label: "− Variables régulières", value: `${fmt(result.totalVariable)} FCFA` },
        { label: "= Reste à vivre estimé", value: `${fmt(result.resteAVivre)} FCFA`, strong: true, warn: result.resteAVivre < 0 },
      ] }],
    };
  };


  const exportChargesExcel = async () => {
    setXlsState("loading");
    try {
      const ExcelJS: any = await import(/* @vite-ignore */ "exceljs");
      const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", MUTED = "FF8A9A8E", WHITE = "FFFFFFFF";
      const wb = new ExcelJS.Workbook();
      wb.creator = "Grand Livre"; wb.created = new Date();
      const styleHeaderRow = (row: any) => {
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; c.alignment = { vertical: "middle" }; });
        row.height = 22;
      };
      const ws1 = wb.addWorksheet("Synthèse");
      ws1.columns = [{ width: 30 }, { width: 22 }];
      ws1.mergeCells("A1:B1");
      const title = ws1.getCell("A1");
      title.value = "Grand Livre — Charges fixes vs variables";
      title.font = { bold: true, size: 15, color: { argb: WHITE } };
      title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
      title.alignment = { vertical: "middle", indent: 1 };
      ws1.getRow(1).height = 30;
      const addSum = (label: string, value: any, color?: string, isAmount?: boolean) => {
        const r = ws1.addRow([label, value]);
        r.getCell(1).font = { color: { argb: MUTED } };
        r.getCell(2).font = { bold: true, size: 12, color: { argb: color || NAVY } };
        if (isAmount) r.getCell(2).numFmt = "#,##0 \"FCFA\"";
        r.getCell(2).alignment = { horizontal: "right" };
      };
      ws1.addRow([]);
      addSum("Généré le", dateLabelFull(todayISO()));
      addSum("Fenêtre d'analyse", `${monthLabel(result.lookback[0])} — ${monthLabel(result.lookback[result.lookback.length - 1])}`);
      addSum("GRUNDFOS", includeGrundfosVoiture ? "Inclus" : "Exclus");
      ws1.addRow([]);
      addSum("Revenu moyen mensuel", result.avgRevenu, undefined, true);
      addSum("Total charges fixes / mois", result.totalFixe, EMERALD, true);
      addSum("Total variables régulières / mois", result.totalVariable, GOLD, true);
      addSum("Reste à vivre estimé / mois", result.resteAVivre, result.resteAVivre >= 0 ? EMERALD : CLAY, true);

      const ws2 = wb.addWorksheet("Détail par poste");
      ws2.columns = [
        { header: "Poste", key: "poste", width: 30 }, { header: "Mode", key: "mode", width: 18 },
        { header: "Montant retenu (FCFA)", key: "amount", width: 20 }, { header: "Moyenne (période)", key: "mean", width: 18 },
        { header: "Médiane (période)", key: "median", width: 18 }, { header: "Mois présents", key: "present", width: 16 },
        { header: "CV", key: "cv", width: 10 }, { header: "Ajusté manuellement", key: "overridden", width: 18 },
      ];
      styleHeaderRow(ws2.getRow(1));
      result.rows.forEach((r) => {
        const row = ws2.addRow({
          poste: r.poste.replace("::", " · "), mode: modeLabel[r.mode], amount: r.amount, mean: Math.round(r.mean), median: Math.round(r.median),
          present: `${r.present}/${result.lookback.length}`, cv: r.cv !== null ? `${r.cv.toFixed(0)}%` : "—", overridden: r.overridden ? "Oui" : "",
        });
        row.getCell("mode").font = { bold: true, color: { argb: r.mode === "fixe" ? EMERALD : r.mode === "variable" ? GOLD : r.mode === "exclu" ? CLAY : MUTED } };
        ["amount", "mean", "median"].forEach((k) => { row.getCell(k).numFmt = "#,##0"; row.getCell(k).alignment = { horizontal: "right" }; });
        ["present", "cv", "overridden"].forEach((k) => { row.getCell(k).alignment = { horizontal: "center" }; });
      });
      ws2.addRow([]);
      (["fixe", "variable", "occasionnelle", "exclu"] as ChargeMode[]).forEach((mode) => {
        const rowsForMode = result.rows.filter((r) => r.mode === mode);
        if (!rowsForMode.length) return;
        const sub = ws2.addRow({ poste: `Total "${modeLabel[mode]}"`, mode: "", amount: rowsForMode.reduce((a, r) => a + r.amount, 0), mean: "", median: "", present: `${rowsForMode.length} poste(s)`, cv: "", overridden: "" });
        sub.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF232F27" } }; });
        sub.getCell("amount").numFmt = "#,##0"; sub.getCell("amount").alignment = { horizontal: "right" };
        sub.getCell("present").alignment = { horizontal: "center" };
      });
      const grandTotal2c = ws2.addRow({ poste: "TOTAL GÉNÉRAL (Fixe + Variable)", mode: "", amount: result.totalFixe + result.totalVariable, mean: "", median: "", present: "", cv: "", overridden: "" });
      grandTotal2c.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      grandTotal2c.getCell("amount").numFmt = "#,##0"; grandTotal2c.getCell("amount").alignment = { horizontal: "right" };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `grand-livre_charges-fixes-variables_${todayISO()}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      setXlsState("idle");
    } catch {
      setXlsState("error");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 12, padding: "12px 16px" }}>
        <div style={{ display: "flex", gap: 4, background: COLOR.surface, borderRadius: 16, padding: 3, border: `1px solid ${COLOR.hairline}` }}>
          <button onClick={() => setIncludeGrundfosVoiture(true)} style={{
            padding: "6px 14px", borderRadius: 12, fontSize: 12, cursor: "pointer", border: "none",
            background: includeGrundfosVoiture ? COLOR.gold : "transparent", color: includeGrundfosVoiture ? COLOR.bg : COLOR.inkMuted,
          }}>Inclure GRUNDFOS</button>
          <button onClick={() => setIncludeGrundfosVoiture(false)} style={{
            padding: "6px 14px", borderRadius: 12, fontSize: 12, cursor: "pointer", border: "none",
            background: !includeGrundfosVoiture ? COLOR.gold : "transparent", color: !includeGrundfosVoiture ? COLOR.bg : COLOR.inkMuted,
          }}>Exclure GRUNDFOS</button>
        </div>
        <span style={{ fontSize: 11.5, color: COLOR.inkMuted, flex: 1, minWidth: 200 }}>
          {includeGrundfosVoiture
            ? "Les charges GRUNDFOS (carburant, internet, électricité…) et Voiture sont comptées dans le budget personnel."
            : "GRUNDFOS est exclue — utile si tu veux voir ton budget personnel pur, séparé de cette activité. \"Voiture\" reste toujours incluse."}
        </span>
      </div>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label={`Revenu moyen / mois (${result.lookback.length}M)`} value={fmt(result.avgRevenu)} tone={COLOR.goldSoft} icon={TrendingUp} onDetailClick={() => setKpiDetailKey("revenu")} />
        <Kpi label="Charges fixes / mois" value={fmt(result.totalFixe)} tone={COLOR.emeraldSoft} icon={CalendarRange} onDetailClick={() => setKpiDetailKey("fixe")} />
        <Kpi label="Variables régulières / mois" value={fmt(result.totalVariable)} tone={COLOR.goldSoft} icon={Repeat} onDetailClick={() => setKpiDetailKey("variable")} />
        <Kpi label="Reste à vivre estimé / mois" value={fmt(result.resteAVivre)} tone={result.resteAVivre >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Wallet} onDetailClick={() => setKpiDetailKey("reste")} />
      </div>

      {onNavigate && (
        <button onClick={() => onNavigate("diagnostic")} style={{
          display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`,
          borderRadius: 8, color: COLOR.slateBlueSoft, padding: "8px 14px", fontSize: 12, cursor: "pointer", width: "fit-content",
        }}>
          <Gauge size={13} /> Voir ces charges dans le Diagnostic Financier (ratios DTI, logement/revenu…) <ArrowRight size={12} />
        </button>
      )}

      <PanelWithHelp title="Classification des charges" subtitle="Basée sur la régularité et la variabilité du montant sur tout l'historique (jamais sur le filtre de période affiché) — ajustable poste par poste"
        explain="Un poste est classé 'Fixe' automatiquement s'il apparaît sur au moins 5/6 de la période avec un montant peu variable (coefficient de variation < 20%). 'Variable régulière' s'il revient souvent mais avec un montant qui fluctue. 'Occasionnel' sinon. En pratique, beaucoup de vraies charges fixes (loyer qui augmente, assurance payée par trimestre, factures irrégulières) ont un montant trop variable pour être détectées automatiquement — le réglage manuel existe pour ces cas, pas par défaut de l'algorithme. Le bouton ↻ à côté de chaque montant recalcule sur la médiane actuelle sans perdre le classement 'Fixe' que tu as choisi."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => {
              const next = { ...chargeOverrides };
              result.rows.forEach((r) => { if (next[r.poste]?.mode === "fixe") next[r.poste] = { ...next[r.poste], amount: Math.round(r.present >= 3 ? (r.mean || r.meanPresent) : r.meanAllTime) }; });
              setChargeOverrides(next);
            }} style={{
              display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`,
              borderRadius: 8, color: COLOR.slateBlueSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer",
            }}>
              <RotateCcw size={13} /> Tout recalculer
            </button>
            <button onClick={exportChargesExcel} disabled={xlsState === "loading"} style={{
              display: "flex", alignItems: "center", gap: 6, background: xlsState === "error" ? "rgba(193,84,63,0.14)" : "rgba(63,156,122,0.14)",
              border: `1px solid ${xlsState === "error" ? COLOR.clay : COLOR.emerald}`, borderRadius: 8,
              color: xlsState === "error" ? COLOR.claySoft : COLOR.emeraldSoft, padding: "7px 14px", fontSize: 12.5, cursor: xlsState === "loading" ? "default" : "pointer",
            }}>
              {xlsState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <FileSpreadsheet size={13} />}
              {xlsState === "loading" ? "Génération…" : xlsState === "error" ? "Réessayer" : "Rapport Excel"}
            </button>
          </div>
        }>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr 0.8fr 0.6fr", gap: 8, padding: "6px 0", fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", borderBottom: `1px solid ${COLOR.hairline}` }}>
            <span>Poste</span><span>Mode</span><span style={{ textAlign: "right" }}>Montant retenu</span><span style={{ textAlign: "center" }}>Régularité</span><span style={{ textAlign: "center" }}>CV</span>
          </div>
          {result.rows.map((r) => {
            const override = chargeOverrides[r.poste];
            const isFixedOverride = override?.mode === "fixe";
            return (
              <div key={r.poste} style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr 0.8fr 0.6fr", gap: 8, alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
                <span style={{ fontSize: 12.5, color: COLOR.ink }}>{r.poste.replace("::", " · ")}<CalcDetailIcon onClick={() => setCalcDetailPoste(r.poste)} /></span>
                <select value={override?.mode || "auto"} onChange={(e) => setOverride(r.poste, { mode: e.target.value as ChargeMode | "auto" })}
                  style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: modeColor[r.mode], padding: "4px 6px", fontSize: 11 }}>
                  <option value="auto">Auto ({modeLabel[r.mode]})</option>
                  <option value="fixe">Fixe</option>
                  <option value="variable">Variable régulière</option>
                  <option value="occasionnelle">Occasionnelle</option>
                  <option value="exclu">Exclu</option>
                </select>
                {isFixedOverride ? (
                  override?.schedule && override.schedule.length ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 6, justifyContent: "flex-end" }}>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: COLOR.ink }}>{fmt(r.amount)} FCFA</span>
                      <button onClick={() => setScheduleSheetPoste(r.poste)} title="Historique des montants (plusieurs périodes)" style={{ background: "rgba(201,162,39,0.12)", border: `1px solid ${COLOR.gold}`, borderRadius: 5, color: COLOR.goldSoft, cursor: "pointer", display: "flex", flexShrink: 0, padding: 4 }}>
                        <Clock size={13} />
                      </button>
                    </div>
                  ) : (
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <FixedAmountInput
                        value={override?.amount ?? Math.round(r.mean)}
                        onSave={(v) => setOverride(r.poste, { amount: v })}
                        onRecalculate={() => setOverride(r.poste, { amount: Math.round(r.present >= 3 ? (r.mean || r.meanPresent) : r.meanAllTime) })}
                        recalcTitle={`Recalculer (${r.present >= 3 ? `moyenne sur la fenêtre affichée` : `trop peu de mois affichés, moyenne sur tout l'historique`} : ${fmt(r.present >= 3 ? (r.mean || r.meanPresent) : r.meanAllTime)} FCFA)`}
                      />
                      <button onClick={() => setScheduleSheetPoste(r.poste)} title="Ce montant a changé plusieurs fois dans le temps ? Définir un historique" style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0, padding: 4 }}>
                        <Clock size={13} />
                      </button>
                    </div>
                  )
                ) : (
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, textAlign: "right", color: r.mode === "exclu" ? COLOR.inkMuted : COLOR.ink }}>
                    {r.mode === "exclu" ? "—" : `${fmt(r.amount)} FCFA`}
                  </span>
                )}
                <span style={{ fontSize: 11.5, color: COLOR.inkMuted, textAlign: "center" }} title="Sur tout l'historique — la classification Fixe/Variable ne dépend jamais du filtre affiché">{r.classPresent}/{r.classTotal}</span>
                <span style={{ fontSize: 11.5, color: COLOR.inkMuted, textAlign: "center" }}>{r.cv !== null ? `${r.cv.toFixed(0)}%` : "—"}</span>
              </div>
            );
          })}
        </div>
      </PanelWithHelp>
      {calcDetailPoste && (() => {
        const d = buildChargeDetail(calcDetailPoste);
        return d ? <CalcDetailSheet open={!!calcDetailPoste} onClose={() => setCalcDetailPoste(null)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} /> : null;
      })()}
      {kpiDetailKey && (() => {
        const d = buildKpiDetail(kpiDetailKey);
        return <CalcDetailSheet open={!!kpiDetailKey} onClose={() => setKpiDetailKey(null)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} />;
      })()}
      {scheduleSheetPoste && (
        <ChargeScheduleSheet
          open={!!scheduleSheetPoste}
          onClose={() => setScheduleSheetPoste(null)}
          poste={scheduleSheetPoste}
          schedule={chargeOverrides[scheduleSheetPoste]?.schedule || []}
          onSave={(s) => setOverride(scheduleSheetPoste, { schedule: s.length ? s : undefined })}
        />
      )}
    </div>
  );
}

// ============================================================
// DIAGNOSTIC FINANCIER — ratios institutionnels (DTI bancaire, règle des 30%
// logement, fonds d'urgence CFPB, concentration des revenus) + simulateur de
// résilience (stress test simplifié : baisse de revenu × durée).
// ============================================================
// ============================================================
// SIGNAUX CLÉS — vue d'ensemble unifiée qui remonte les 5 informations les
// plus importantes issues de tous les moteurs d'analyse de l'app (Conseiller
// quotidien, notes du jour/mois, charges fixes/variables, ratios financiers),
// pour éviter d'avoir à visiter 6 pages différentes pour voir l'essentiel.
// ============================================================
function SignauxClesPanel({ transactions, accounts, chargeOverrides, includeGrundfosVoiture, monthlyObjective }: {
  transactions: Transaction[]; accounts: Account[]; chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean; monthlyObjective: number;
}) {
  const dayScore = useMemo(() => computeDayScore(transactions, monthlyObjective), [transactions, monthlyObjective]);
  const monthScore = useMemo(() => computeMonthScore(transactions, monthlyObjective), [transactions, monthlyObjective]);
  const advice = useMemo(() => generateDailyAdvice(transactions, monthlyObjective, chargeOverrides, includeGrundfosVoiture), [transactions, monthlyObjective, chargeOverrides, includeGrundfosVoiture]);
  const charges = useMemo(() => classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture), [transactions, chargeOverrides, includeGrundfosVoiture]);
  const { ratios } = useMemo(() => computeFinancialRatios(transactions, accounts, chargeOverrides, includeGrundfosVoiture), [transactions, accounts, chargeOverrides, includeGrundfosVoiture]);
  const today = todayISO();
  const curMonth = dateToMonthKey(today);
  const dayNum = new Date(today + "T00:00:00").getDate();
  const upcoming = useMemo(() => detectRecurringExpenses(transactions, curMonth, dayNum, chargeOverrides, includeGrundfosVoiture), [transactions, curMonth, dayNum, chargeOverrides, includeGrundfosVoiture]);

  const gradeColor: Record<string, string> = { Excellent: COLOR.emerald, Bon: COLOR.emeraldSoft, Moyen: COLOR.gold, Faible: COLOR.clay };
  const topAlert = advice.insights.find((i) => i.kind === "alerte") || advice.insights[0];
  const worstRatio = ratios.find((r) => r.verdict === "risque") || ratios.find((r) => r.verdict === "vigilance");

  const signals = [
    { icon: Clock, color: dayScore.gradeColor, title: `Journée : ${dayScore.grade} (${Math.round(dayScore.overall)}/100)`, text: `Note du mois : ${monthScore.grade} (${Math.round(monthScore.overall)}/100)` },
    topAlert ? { icon: topAlert.kind === "alerte" ? AlertTriangle : topAlert.kind === "positif" ? Check : Info, color: topAlert.kind === "alerte" ? COLOR.claySoft : topAlert.kind === "positif" ? COLOR.emeraldSoft : COLOR.goldSoft, title: topAlert.title, text: topAlert.text } : null,
    { icon: Wallet, color: charges.resteAVivre >= 0 ? COLOR.emeraldSoft : COLOR.claySoft, title: `Reste à vivre estimé : ${fmt(charges.resteAVivre)} FCFA/mois`, text: `Charges fixes ${fmt(charges.totalFixe)} + variables ${fmt(charges.totalVariable)} déduites du revenu moyen.` },
    worstRatio ? { icon: AlertTriangle, color: worstRatio.verdict === "risque" ? COLOR.claySoft : COLOR.goldSoft, title: `${worstRatio.label} : ${worstRatio.unit === "mois" ? worstRatio.value.toFixed(1) : Math.round(worstRatio.value)}${worstRatio.unit === "mois" ? " mois" : worstRatio.unit}`, text: worstRatio.benchmark } : { icon: Check, color: COLOR.emeraldSoft, title: "Tous les ratios financiers sont sains", text: "Aucun signal de risque sur les repères institutionnels (DTI, logement, fonds d'urgence, concentration des revenus)." },
    upcoming.length ? { icon: CalendarRange, color: COLOR.goldSoft, title: `${upcoming.length} charge(s) périodique(s) probablement encore à venir`, text: upcoming.map((u) => `${u.category} (~${fmt(u.typicalAmount)} FCFA)`).join(", ") } : null,
  ].filter(Boolean) as { icon: any; color: string; title: string; text: string }[];

  return (
    <Panel title="Signaux clés" subtitle="L'essentiel de tous les moteurs d'analyse de l'app, en un coup d'œil">
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {signals.map((s, i) => {
          const Icon = s.icon;
          return (
            <div key={i} style={{ display: "flex", gap: 12, padding: "10px 12px", background: `${s.color}14`, border: `1px solid ${s.color}`, borderRadius: 8 }}>
              <Icon size={15} color={s.color} style={{ flexShrink: 0, marginTop: 1 }} />
              <div style={{ minWidth: 0 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600, color: s.color, marginBottom: 2 }}>{s.title}</div>
                <div style={{ fontSize: 11.5, color: COLOR.inkMuted, lineHeight: 1.5, overflow: "hidden", textOverflow: "ellipsis" }}>{s.text}</div>
              </div>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

// ============================================================
// BANNIÈRE DE RAPPEL PROACTIF — visible sur tous les onglets (pas seulement
// dans le Conseiller quotidien qu'il faut aller consulter), pour les charges
// périodiques qui approchent sans être encore enregistrées ce mois-ci.
// Peut être masquée pour la journée ; réapparaît le lendemain si toujours
// non résolue.
// ============================================================
function GlobalReminderBanner({ transactions, dismissedDate, setDismissedDate, chargeOverrides, includeGrundfosVoiture }: {
  transactions: Transaction[]; dismissedDate: string | null; setDismissedDate: (d: string) => void;
  chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean;
}) {
  const today = todayISO();
  const curMonth = dateToMonthKey(today);
  const dayNum = new Date(today + "T00:00:00").getDate();
  const upcoming = useMemo(() => detectRecurringExpenses(transactions, curMonth, dayNum, chargeOverrides, includeGrundfosVoiture), [transactions, curMonth, dayNum, chargeOverrides, includeGrundfosVoiture]);

  if (!upcoming.length || dismissedDate === today) return null;
  const total = upcoming.reduce((a, u) => a + u.typicalAmount, 0);

  return (
    <div className="gl-noprint" style={{
      display: "flex", alignItems: "center", gap: 12, background: "rgba(201,162,39,0.1)", border: `1px solid ${COLOR.gold}`,
      borderRadius: 10, padding: "12px 16px", marginBottom: 16,
    }}>
      <CalendarRange size={17} color={COLOR.goldSoft} style={{ flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0, fontSize: 12.5, color: COLOR.ink }}>
        <strong style={{ color: COLOR.goldSoft }}>{upcoming.length} charge(s) périodique(s)</strong> probablement encore à venir ce mois-ci (~{fmt(total)} FCFA) :{" "}
        <span style={{ color: COLOR.inkMuted }}>{upcoming.map((u) => u.category).join(", ")}</span>
      </div>
      <button onClick={() => setDismissedDate(today)} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }} title="Masquer pour aujourd'hui">
        <X size={16} />
      </button>
    </div>
  );
}

// ============================================================
// RAPPROCHEMENT BANCAIRE — pointer les transactions face au relevé bancaire
// réel, compte par compte, pour repérer tout écart résiduel après import.
// ============================================================
function RapprochementTab({ transactions, setTransactions, accounts }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; accounts: Account[];
}) {
  const [selectedAccount, setSelectedAccount] = useState(accounts[0]?.name || "");
  const curMonth = dateToMonthKey(todayISO());
  const [periodFrom, setPeriodFrom] = useState(curMonth);
  const [periodTo, setPeriodTo] = useState(curMonth);
  // Empêche une plage inversée (Du mois postérieur à Au mois) — échange les deux valeurs
  // plutôt que de laisser une plage impossible s'installer.
  const setPeriodFromSafe = (v: string) => { if (monthSortKey(v) > monthSortKey(periodTo)) { setPeriodFrom(periodTo); setPeriodTo(v); } else setPeriodFrom(v); };
  const setPeriodToSafe = (v: string) => { if (monthSortKey(v) < monthSortKey(periodFrom)) { setPeriodTo(periodFrom); setPeriodFrom(v); } else setPeriodTo(v); };
  const [soldeReel, setSoldeReel] = useState<number | "">("");

  const allMonths = useMemo(() => Array.from(new Set(transactions.map((t) => dateToMonthKey(t.date)))).sort((a, b) => monthSortKey(a) - monthSortKey(b)), [transactions]);

  const periodTx = useMemo(() => {
    const fk = monthSortKey(periodFrom), tk = monthSortKey(periodTo);
    return transactions
      .filter((t) => t.account === selectedAccount && monthSortKey(dateToMonthKey(t.date)) >= fk && monthSortKey(dateToMonthKey(t.date)) <= tk)
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [transactions, selectedAccount, periodFrom, periodTo]);

  const account = accounts.find((a) => a.name === selectedAccount);
  // Solde théorique de l'app, à la fin de la période sélectionnée (pas juste sur la période affichée).
  const soldeAppFinPeriode = useMemo(() => {
    if (!account) return 0;
    const tk = monthSortKey(periodTo);
    const net = transactions
      .filter((t) => t.account === selectedAccount && monthSortKey(dateToMonthKey(t.date)) <= tk)
      .reduce((a, t) => a + (t.type === "Revenu" ? t.amount : -t.amount), 0);
    return account.openingBalance + net;
  }, [transactions, selectedAccount, periodTo, account]);

  const ecart = soldeReel !== "" ? (soldeReel as number) - soldeAppFinPeriode : null;

  const toggleReconciled = (id: string) => setTransactions(transactions.map((t) => (t.id === id ? { ...t, reconciled: !t.reconciled } : t)));
  const pointedCount = periodTx.filter((t) => t.reconciled).length;
  const pointedTotal = periodTx.filter((t) => t.reconciled).reduce((a, t) => a + (t.type === "Revenu" ? t.amount : -t.amount), 0);
  const nonPointedCount = periodTx.length - pointedCount;

  const markAllReconciled = (value: boolean) => setTransactions(transactions.map((t) => (periodTx.some((p) => p.id === t.id) ? { ...t, reconciled: value } : t)));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <PanelWithHelp title="Rapprochement bancaire" subtitle="Pointe chaque transaction face à ton relevé bancaire réel pour détecter tout écart résiduel"
        explain="Sélectionne un compte et une période, puis coche chaque transaction que tu retrouves sur ton relevé bancaire réel. Renseigne le solde réel indiqué sur ton relevé à la fin de la période : si un écart apparaît avec le solde théorique de l'app, c'est le signe d'une transaction manquante, en double, ou mal datée — repérable ligne par ligne plutôt qu'en devinant.">
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 20 }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte</label>
            <select value={selectedAccount} onChange={(e) => setSelectedAccount(e.target.value)} style={{ ...inputStyle, width: 160 }}>
              {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
            </select>
          </div>
          <Select label="Du mois" value={periodFrom} onChange={setPeriodFromSafe} options={allMonths.filter((m) => monthSortKey(m) <= monthSortKey(periodTo)).map((m) => ({ value: m, label: monthLabel(m) }))} />
          <Select label="Au mois" value={periodTo} onChange={setPeriodToSafe} options={allMonths.filter((m) => monthSortKey(m) >= monthSortKey(periodFrom)).map((m) => ({ value: m, label: monthLabel(m) }))} />
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Solde réel (relevé, fin de période)</label>
            <input type="number" inputMode="numeric" style={{ ...inputStyle, width: 160, textAlign: "right" }} placeholder="Ex : 1 250 000" value={soldeReel}
              onChange={(e) => setSoldeReel(e.target.value === "" ? "" : Number(e.target.value))} />
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 20 }}>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Solde théorique (app)</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLOR.ink }}>{fmt(soldeAppFinPeriode)} FCFA</div>
          </div>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${ecart === null ? COLOR.hairline : Math.abs(ecart) < 1 ? COLOR.emerald : COLOR.clay}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Écart avec le relevé réel</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: ecart === null ? COLOR.inkMuted : Math.abs(ecart) < 1 ? COLOR.emeraldSoft : COLOR.claySoft }}>
              {ecart === null ? "— saisis le solde réel" : `${ecart >= 0 ? "+" : ""}${fmt(ecart)} FCFA`}
            </div>
          </div>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Pointées / total</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLOR.ink }}>{pointedCount} / {periodTx.length}</div>
          </div>
          <div style={{ background: COLOR.surfaceRaised, border: `1px solid ${nonPointedCount > 0 ? COLOR.gold : COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Non pointées</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: nonPointedCount > 0 ? COLOR.goldSoft : COLOR.emeraldSoft }}>{nonPointedCount}</div>
          </div>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <button onClick={() => markAllReconciled(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(63,156,122,0.14)", border: `1px solid ${COLOR.emerald}`, borderRadius: 6, color: COLOR.emeraldSoft, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}>
            <CheckSquare size={12} /> Tout pointer
          </button>
          <button onClick={() => markAllReconciled(false)} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}>
            <Square size={12} /> Tout dépointer
          </button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {periodTx.map((t) => (
            <div key={t.id} onClick={() => toggleReconciled(t.id)} style={{
              display: "flex", alignItems: "center", gap: 12, padding: "9px 12px", borderRadius: 8, cursor: "pointer",
              background: t.reconciled ? "rgba(63,156,122,0.08)" : COLOR.surfaceRaised, border: `1px solid ${t.reconciled ? COLOR.emerald : COLOR.hairline}`,
            }}>
              {t.reconciled ? <CheckSquare size={16} color={COLOR.emeraldSoft} style={{ flexShrink: 0 }} /> : <Square size={16} color={COLOR.inkMuted} style={{ flexShrink: 0 }} />}
              <div style={{ flex: 1, minWidth: 0, fontSize: 12.5, color: COLOR.ink }}>
                {dateLabelFull(t.date)} — {t.category}{t.subcategory && <span style={{ color: COLOR.inkMuted }}> · {t.subcategory}</span>}
              </div>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft, flexShrink: 0 }}>
                {t.type === "Revenu" ? "+" : "−"}{fmt(t.amount)} FCFA
              </span>
            </div>
          ))}
          {!periodTx.length && <EmptyState text="Aucune transaction pour ce compte sur cette période." />}
        </div>
      </PanelWithHelp>
    </div>
  );
}

// Fiche de lecture pour le rapport narratif détaillé (4-3-2-1 + Kakeibo +
// synthèse) — ouverte à la demande pour ne pas alourdir la page.
function NarrativeReportSheet({ open, onClose, rule4321, tauxEpargne, kakeibo }: {
  open: boolean; onClose: () => void;
  rule4321: { label: string; value: number; target: number }[]; tauxEpargne: number; kakeibo: { label: string; key: string; pct: number }[];
}) {
  if (!open) return null;
  const narr4321 = generateAsian4321Narrative(rule4321);
  const narrKakeibo = generateKakeiboNarrative(kakeibo);
  const synthesis = generateFinancialProfileSynthesis(narr4321, tauxEpargne, narrKakeibo);

  const Section = ({ s }: { s: NarrativeSection }) => {
    const color = s.verdict.includes("loin") || s.verdict.includes("élevé") || s.verdict.includes("faible")
      ? COLOR.claySoft : s.verdict.includes("marge de progression") ? COLOR.goldSoft : COLOR.emeraldSoft;
    return (
      <div style={{ marginBottom: 22 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4, flexWrap: "wrap", gap: 6 }}>
          <span style={{ fontFamily: "'Fraunces', serif", fontSize: 15.5, color: COLOR.ink }}>{s.title}</span>
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: COLOR.inkMuted }}>Toi : <strong style={{ color }}>{s.you.toFixed(0)}%</strong> · Objectif : {s.target}%</span>
        </div>
        <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, margin: "4px 0 8px 0" }}>{s.definition}</p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
          {s.examples.map((ex) => (
            <span key={ex} style={{ fontSize: 10.5, color: COLOR.inkMuted, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "3px 10px" }}>{ex}</span>
          ))}
        </div>
        <div style={{ fontSize: 12.5, color, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
          <ArrowRight size={13} /> {s.verdict}
        </div>
      </div>
    );
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 480, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "100%", maxWidth: 620, maxHeight: "90vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
        display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "18px 22px", borderBottom: `1px solid ${COLOR.hairline}` }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, color: COLOR.ink }}>Rapport détaillé — Indicateurs asiatiques</div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}><X size={18} /></button>
        </div>
        <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: "20px 22px", WebkitOverflowScrolling: "touch" }}>

          <div style={{ fontSize: 13, fontWeight: 700, color: COLOR.goldSoft, marginBottom: 4 }}>1. Règle chinoise du 4-3-2-1</div>
          <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 18 }}>
            Cette règle indique comment chaque revenu mensuel devrait idéalement être réparti : 40% investissement, 30% vie courante, 20% protection, 10% épargne de précaution. Voici comment tes chiffres se comparent.
          </p>
          {narr4321.map((s) => <Section key={s.title} s={s} />)}

          <div style={{ fontSize: 13, fontWeight: 700, color: COLOR.goldSoft, marginBottom: 4, marginTop: 8 }}>2. Taux d'épargne — norme chinoise</div>
          <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 22 }}>
            Ton taux d'épargne est de <strong style={{ color: COLOR.ink }}>{tauxEpargne.toFixed(0)}%</strong>. Cela ne veut pas dire que tu mets ce pourcentage de côté sur un compte — cela signifie que {tauxEpargne.toFixed(0)}% de tes revenus ne sont pas consommés en dépenses courantes (épargne + investissements + protection cumulés). Référence : Chine 30 à 45%, Occident environ 20%.
            {" "}{tauxEpargne >= 30 ? "Tu dépasses la référence chinoise." : tauxEpargne >= 20 ? "Tu es dans la moyenne occidentale, mais sous la référence chinoise." : "Tu es sous les deux références."}
          </p>

          <div style={{ fontSize: 13, fontWeight: 700, color: COLOR.goldSoft, marginBottom: 4 }}>3. Répartition Kakeibo (méthode japonaise)</div>
          <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 18 }}>Le Kakeibo classe chaque dépense en quatre catégories, pour favoriser la réflexion plutôt qu'un simple total.</p>
          {narrKakeibo.map((s) => <Section key={s.title} s={s} />)}

          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, paddingTop: 18, marginTop: 4 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: COLOR.ink, marginBottom: 10 }}>Ce que ton profil financier raconte</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 }}>
              {synthesis.checks.map((c, i) => (
                <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, color: c.ok ? COLOR.emeraldSoft : COLOR.claySoft, lineHeight: 1.5 }}>
                  <span style={{ flexShrink: 0 }}>{c.ok ? "✅" : "⚠️"}</span>
                  <span>{c.text}</span>
                </div>
              ))}
            </div>
            <div style={{ padding: "12px 14px", background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
              {synthesis.recommendation}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// Fiche de lecture pour le rapport narratif des 5 ratios institutionnels.
type CalcDetailBlock =
  | { kind: "kv"; rows: { label: string; value: string; strong?: boolean; warn?: boolean }[] }
  | { kind: "table"; columns: string[]; rows: (string | number)[][]; warnRows?: number[]; cellColors?: (string | undefined)[][]; footerRow?: (string | number)[]; footerColors?: (string | undefined)[] }
  | { kind: "note"; text: string; tone?: "warn" | "info" };

// Petite icône cliquable placée à côté de chaque chiffre du Diagnostic Financier —
// ouvre CalcDetailSheet avec le détail exact (formule + composants) de ce chiffre-là,
// pour que chaque taux affiché soit vérifiable en un clic plutôt qu'à prendre pour acquis.
function CalcDetailIcon({ onClick }: { onClick: () => void }) {
  return (
    <button onClick={(e) => { e.stopPropagation(); onClick(); }} title="Voir le détail du calcul" style={{
      background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "inline-flex",
      padding: 2, marginLeft: 4, verticalAlign: "middle", opacity: 0.7,
    }}>
      <Info size={13} />
    </button>
  );
}

function CalcDetailSheet({ open, onClose, title, headline, formula, blocks, onPrev, onNext }: {
  open: boolean; onClose: () => void; title: string; headline: string; formula: string; blocks: CalcDetailBlock[];
  onPrev?: () => void; onNext?: () => void;
}) {
  const [pdfState, setPdfState] = useState<"idle" | "loading" | "error">("idle");
  if (!open) return null;
  // PDF structuré (même style que les autres exports de l'app : bandeau navy, encadré
  // doré pour le résultat, tableaux via autoTable) — remplace le CSV, jugé peu présentable.
  const downloadPdf = async () => {
    setPdfState("loading");
    try {
      const [jsPDFModule, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      // jspdf@2.5.x expose le vrai constructeur sur l'export NOMMÉ "jsPDF", pas sur
      // "default" (qui résout vers un objet inutilisable selon le mode d'interop
      // CJS/ESM) — cause du "Réessayer" systématique sur tous les boutons PDF de l'app.
      const jsPDF: any = (jsPDFModule as any).jsPDF || (jsPDFModule as any).default;
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();
      // Les polices standard de jsPDF (Helvetica) ne supportent pas l'espace fine
      // insécable (U+202F) que produit Intl.NumberFormat("fr-FR") pour les milliers —
      // elle s'affichait comme "/" (ex: "508/111" au lieu de "508 111"). Toute chaîne
      // affichée dans ce PDF passe par ce filtre avant écriture.
      const ps = (s: any): string => String(s).replace(/[\u202F\u00A0]/g, " ").replace(/→/g, "->").replace(/—/g, "-").replace(/…/g, "...");

      doc.setFillColor(26, 43, 76);
      doc.rect(0, 0, pageWidth, 34, "F");
      doc.setFontSize(16); doc.setTextColor(255, 255, 255);
      doc.text(ps(`Grand Livre — ${title}`), 14, 16);
      doc.setFontSize(9); doc.setTextColor(200, 210, 225);
      doc.text(ps(`Généré le ${dateLabelFull(todayISO())}`), 14, 24);
      doc.text(ps(formula), 14, 29);

      doc.setFillColor(201, 162, 39);
      doc.roundedRect(14, 40, pageWidth - 28, 16, 2, 2, "F");
      doc.setFontSize(8); doc.setTextColor(26, 26, 26);
      doc.text("MONTANT", 20, 46);
      doc.setFontSize(13); doc.setFont("helvetica", "bold");
      doc.text(ps(headline), 20, 53);
      doc.setFont("helvetica", "normal");

      let y = 66;
      blocks.forEach((b) => {
        if (y > 260) { doc.addPage(); y = 20; }
        if (b.kind === "kv") {
          doc.autoTable({
            startY: y,
            body: b.rows.map((r) => [ps(r.label), ps(r.value)]),
            theme: "plain",
            styles: { fontSize: 9, cellPadding: { top: 2.5, bottom: 2.5, left: 2, right: 2 } },
            columnStyles: { 1: { halign: "right", fontStyle: "bold" } },
            didParseCell: (data: any) => {
              const row = b.rows[data.row.index];
              if (row?.warn) data.cell.styles.textColor = [193, 84, 63];
              else if (row?.strong && data.column.index === 1) data.cell.styles.textColor = [201, 162, 39];
            },
          });
          y = (doc as any).lastAutoTable.finalY + 8;
        } else if (b.kind === "table") {
          doc.autoTable({
            startY: y,
            head: [b.columns.map(ps)],
            body: b.rows.map((row) => row.map((c) => ps(c))),
            foot: b.footerRow ? [b.footerRow.map(ps)] : undefined,
            headStyles: { fillColor: [26, 43, 76] },
            footStyles: { fillColor: [255, 255, 255], textColor: [20, 20, 20], fontStyle: "bold", lineWidth: 0.3 },
            styles: { fontSize: 8 },
            columnStyles: Object.fromEntries(b.columns.map((_, i) => [i, i === 0 ? { halign: "left" } : { halign: "right" }])),
            didParseCell: (data: any) => {
              if (data.section === "foot") {
                const footColor = b.footerColors?.[data.column.index];
                if (footColor) {
                  const hex = footColor.replace("#", "");
                  data.cell.styles.textColor = [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
                }
                return;
              }
              if (data.section !== "body") return;
              const cellColor = b.cellColors?.[data.row.index]?.[data.column.index];
              if (cellColor) {
                const hex = cellColor.replace("#", "");
                data.cell.styles.textColor = [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
              } else if (b.warnRows?.includes(data.row.index)) {
                data.cell.styles.textColor = [193, 84, 63];
              }
            },
          });
          y = (doc as any).lastAutoTable.finalY + 10;
        } else {
          doc.setFillColor(b.tone === "warn" ? 250 : 250, b.tone === "warn" ? 235 : 245, b.tone === "warn" ? 232 : 220);
          const lines = doc.splitTextToSize(ps(b.text), pageWidth - 32);
          const boxH = lines.length * 4.5 + 6;
          if (y + boxH > 275) { doc.addPage(); y = 20; }
          doc.roundedRect(14, y, pageWidth - 28, boxH, 2, 2, "F");
          doc.setFontSize(8); doc.setTextColor(b.tone === "warn" ? 150 : 120, b.tone === "warn" ? 60 : 100, b.tone === "warn" ? 50 : 40);
          doc.text(lines, 18, y + 6);
          doc.setTextColor(0, 0, 0);
          y += boxH + 8;
        }
      });

      doc.save(`grand-livre_${title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")}_${todayISO()}.pdf`);
      setPdfState("idle");
    } catch (e) {
      console.error(e);
      setPdfState("error");
    }
  };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 490, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "100%", maxWidth: 620, maxHeight: "85vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
        display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "18px 22px", borderBottom: `1px solid ${COLOR.hairline}` }}>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, color: COLOR.ink }}>{title}</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, color: COLOR.goldSoft, marginTop: 4 }}>{headline}</div>
            <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 4, fontStyle: "italic" }}>{formula}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 4, flexShrink: 0 }}>
            {onPrev && <button onClick={onPrev} title="Période précédente" style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", padding: 4 }}><ChevronLeft size={18} /></button>}
            {onNext && <button onClick={onNext} title="Période suivante" style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", padding: 4 }}><ChevronRight size={18} /></button>}
            <button onClick={downloadPdf} disabled={pdfState === "loading"} title="Télécharger cette fiche en PDF" style={{ background: "transparent", border: "none", color: pdfState === "error" ? COLOR.claySoft : COLOR.slateBlueSoft, cursor: pdfState === "loading" ? "default" : "pointer", display: "flex", padding: 4 }}>
              {pdfState === "loading" ? <Loader2 size={17} style={{ animation: "spin 1s linear infinite" }} /> : <Download size={17} />}
            </button>
            <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex" }}><X size={18} /></button>
          </div>
        </div>
        <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: "18px 22px", WebkitOverflowScrolling: "touch" }}>
          {blocks.map((b, i) => {
            if (b.kind === "kv") return (
              <div key={i} style={{ marginBottom: 16 }}>
                {b.rows.map((r, j) => (
                  <div key={j} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: j < b.rows.length - 1 ? `1px solid ${COLOR.hairline}` : "none" }}>
                    <span style={{ fontSize: 12.5, color: r.strong ? COLOR.ink : COLOR.inkMuted, fontWeight: r.strong ? 600 : 400 }}>{r.label}</span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: r.warn ? COLOR.claySoft : r.strong ? COLOR.goldSoft : COLOR.ink, fontWeight: r.strong ? 700 : 500 }}>{r.value}</span>
                  </div>
                ))}
              </div>
            );
            if (b.kind === "table") return (
              <div key={i} style={{ marginBottom: 16, overflowX: "auto", maxWidth: "100%" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11.5, tableLayout: "fixed" }}>
                  <thead>
                    <tr>{b.columns.map((c, ci) => <th key={ci} style={{ textAlign: ci === 0 ? "left" : "right", padding: "6px 4px", color: COLOR.inkMuted, fontWeight: 600, borderBottom: `1px solid ${COLOR.hairline}`, whiteSpace: "normal", wordBreak: "break-word", lineHeight: 1.3 }}>{c}</th>)}</tr>
                  </thead>
                  <tbody>
                    {b.rows.map((row, ri) => (
                      <tr key={ri} style={{ background: b.warnRows?.includes(ri) ? "rgba(193,84,63,0.08)" : "transparent" }}>
                        {row.map((cell, ci) => <td key={ci} style={{ textAlign: ci === 0 ? "left" : "right", padding: "6px 4px", fontFamily: ci === 0 ? "inherit" : "'IBM Plex Mono', monospace", fontWeight: b.cellColors?.[ri]?.[ci] ? 600 : 400, color: b.cellColors?.[ri]?.[ci] || (b.warnRows?.includes(ri) ? COLOR.claySoft : COLOR.ink), borderBottom: `1px solid ${COLOR.hairline}`, whiteSpace: "normal", wordBreak: "break-word" }}>{cell}</td>)}
                      </tr>
                    ))}
                  </tbody>
                  {b.footerRow && (
                    <tfoot>
                      <tr style={{ background: "rgba(91,126,166,0.08)" }}>
                        {b.footerRow.map((cell, ci) => <td key={ci} style={{
                          textAlign: ci === 0 ? "left" : "right", padding: "7px 4px", fontFamily: ci === 0 ? "'Inter', sans-serif" : "'IBM Plex Mono', monospace",
                          fontWeight: 700, color: b.footerColors?.[ci] || COLOR.ink, whiteSpace: "normal", wordBreak: "break-word",
                          borderTop: `1.5px solid ${COLOR.slateBlue}`, borderBottom: `1.5px solid ${COLOR.slateBlue}`,
                          borderLeft: ci === 0 ? `1.5px solid ${COLOR.slateBlue}` : "none",
                          borderRight: ci === (b.footerRow!.length - 1) ? `1.5px solid ${COLOR.slateBlue}` : "none",
                        }}>{cell}</td>)}
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            );
            return (
              <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-start", background: b.tone === "warn" ? "rgba(193,84,63,0.1)" : "rgba(201,162,39,0.08)", border: `1px solid ${b.tone === "warn" ? COLOR.clay : COLOR.hairline}`, borderRadius: 8, padding: "10px 12px", marginBottom: 16 }}>
                <AlertTriangle size={14} color={b.tone === "warn" ? COLOR.claySoft : COLOR.goldSoft} style={{ flexShrink: 0, marginTop: 1 }} />
                <div style={{ fontSize: 12, color: b.tone === "warn" ? COLOR.claySoft : COLOR.inkMuted, lineHeight: 1.5 }}>{b.text}</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}


const monthKeyToInputValue = (mk: string) => { const [y, m] = mk.split("_"); return `${y}-${String(m).padStart(2, "0")}`; };
const inputValueToMonthKey = (v: string) => { const [y, m] = v.split("-"); return `${parseInt(y, 10)}_${parseInt(m, 10)}`; };

// Fiche de gestion de l'historique des montants d'un poste "Fixe" — pour un loyer (ou
// toute charge) qui a changé plusieurs fois dans le temps plutôt qu'un montant unique
// valable pour toute l'historique. Le montant retenu pour les calculs (DTI, 4-3-2-1...)
// utilise toujours le segment couvrant aujourd'hui (ou le plus récent).
function ChargeScheduleSheet({ open, onClose, poste, schedule, onSave }: {
  open: boolean; onClose: () => void; poste: string; schedule: ChargeScheduleEntry[]; onSave: (s: ChargeScheduleEntry[]) => void;
}) {
  const [rows, setRows] = useState<ChargeScheduleEntry[]>(schedule.length ? schedule : []);
  useEffect(() => { if (open) setRows(schedule.length ? schedule : []); }, [open, schedule]);
  if (!open) return null;

  const sorted = [...rows].sort((a, b) => monthSortKey(a.from) - monthSortKey(b.from));
  const addRow = () => {
    const lastTo = sorted.length ? sorted[sorted.length - 1].to : null;
    const nextFrom = lastTo ? nextMonthKey(lastTo) : dateToMonthKey(todayISO());
    setRows([...rows, { from: nextFrom, to: null, amount: 0 }]);
  };
  const updateRow = (i: number, patch: Partial<ChargeScheduleEntry>) => setRows(rows.map((r, ri) => (ri === i ? { ...r, ...patch } : r)));
  const removeRow = (i: number) => setRows(rows.filter((_, ri) => ri !== i));
  const save = () => { onSave(rows.filter((r) => r.amount > 0)); onClose(); };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 495, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "100%", maxWidth: 560, maxHeight: "85vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
        display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 22px", borderBottom: `1px solid ${COLOR.hairline}` }}>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, color: COLOR.ink }}>Historique — {poste.replace("::", " · ")}</div>
            <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 2 }}>Montant retenu pour les calculs = le segment couvrant aujourd'hui</div>
          </div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}><X size={18} /></button>
        </div>
        <div style={{ flex: 1, overflowY: "auto", padding: "18px 22px" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {sorted.map((r) => {
              const i = rows.indexOf(r);
              return (
                <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap", padding: 12, background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}` }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    <label style={{ fontSize: 10, color: COLOR.inkMuted }}>Depuis</label>
                    <input type="month" style={{ ...inputStyle, width: 130 }} value={monthKeyToInputValue(r.from)} onChange={(e) => updateRow(i, { from: inputValueToMonthKey(e.target.value) })} />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    <label style={{ fontSize: 10, color: COLOR.inkMuted }}>Jusqu'à</label>
                    <input type="month" disabled={r.to === null} style={{ ...inputStyle, width: 130, opacity: r.to === null ? 0.4 : 1 }} value={r.to ? monthKeyToInputValue(r.to) : ""} onChange={(e) => updateRow(i, { to: inputValueToMonthKey(e.target.value) })} />
                  </div>
                  <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: COLOR.inkMuted, paddingBottom: 8 }}>
                    <input type="checkbox" checked={r.to === null} onChange={(e) => updateRow(i, { to: e.target.checked ? null : dateToMonthKey(todayISO()) })} /> En cours
                  </label>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    <label style={{ fontSize: 10, color: COLOR.inkMuted }}>Montant (FCFA)</label>
                    <input type="number" inputMode="numeric" style={{ ...inputStyle, width: 120, textAlign: "right" }} value={r.amount || ""} onChange={(e) => updateRow(i, { amount: Number(e.target.value) })} />
                  </div>
                  <button onClick={() => removeRow(i)} style={{ ...iconBtnStyle(COLOR.claySoft), marginBottom: 4 }}><Trash2 size={13} /></button>
                </div>
              );
            })}
            {!sorted.length && <EmptyState text="Aucun segment — ajoute la première période." />}
          </div>
          <button onClick={addRow} style={{ display: "flex", alignItems: "center", gap: 6, background: "transparent", border: `1px dashed ${COLOR.hairline}`, borderRadius: 8, color: COLOR.slateBlueSoft, padding: "10px 14px", fontSize: 12.5, cursor: "pointer", width: "100%", justifyContent: "center", marginTop: 12 }}>
            <Plus size={13} /> Ajouter une période
          </button>
        </div>
        <div style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "16px 22px", display: "flex", gap: 10 }}>
          <button onClick={() => { onSave([]); onClose(); }} style={{ padding: "12px 16px", borderRadius: 10, border: `1px solid ${COLOR.hairline}`, background: "transparent", color: COLOR.inkMuted, fontSize: 13, cursor: "pointer" }}>Effacer l'historique</button>
          <button onClick={save} style={{ flex: 1, padding: "12px 0", borderRadius: 10, border: "none", background: COLOR.gold, color: COLOR.bg, fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Enregistrer</button>
        </div>
      </div>
    </div>
  );
}


function RatiosNarrativeSheet({ open, onClose, ratios }: { open: boolean; onClose: () => void; ratios: RatioResult[] }) {
  if (!open) return null;
  const report = generateRatiosNarrative(ratios);

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 480, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "100%", maxWidth: 620, maxHeight: "90vh", background: COLOR.surface, borderRadius: "20px 20px 0 0",
        display: "flex", flexDirection: "column", border: `1px solid ${COLOR.hairline}`, borderBottom: "none",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "18px 22px", borderBottom: `1px solid ${COLOR.hairline}` }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17, color: COLOR.ink }}>Rapport détaillé — Ratios institutionnels</div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex", flexShrink: 0 }}><X size={18} /></button>
        </div>
        <div className="gl-scroll" style={{ flex: 1, overflowY: "auto", padding: "20px 22px", WebkitOverflowScrolling: "touch" }}>
          <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, marginBottom: 20 }}>
            Ces 5 repères viennent des banques, des régulateurs financiers (ex. CFPB américain) et des cabinets de conseil en gestion de patrimoine — pas des seuils inventés pour l'app. Voici ce que chacun signifie concrètement, et où tu te situes.
          </p>
          {report.sections.map((s) => {
            const color = s.verdict.startsWith("C'est un très bon") || s.verdict.startsWith("Tu as un vrai") || s.verdict.startsWith("Tes revenus sont bien") || s.verdict.startsWith("Tes charges fixes laissent") || s.verdict.startsWith("Ton logement reste")
              ? COLOR.emeraldSoft
              : s.verdict.startsWith("C'est un niveau bas") || s.verdict.startsWith("Tes charges fixes absorbent") || s.verdict.startsWith("Ton logement absorbe") || s.verdict.startsWith("Ta réserve") || s.verdict.startsWith("Une seule source")
                ? COLOR.claySoft : COLOR.goldSoft;
            return (
              <div key={s.key} style={{ marginBottom: 22 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4, flexWrap: "wrap", gap: 6 }}>
                  <span style={{ fontFamily: "'Fraunces', serif", fontSize: 15.5, color: COLOR.ink }}>{s.title}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color }}>{s.unit === "mois" ? s.you.toFixed(1) + " mois" : Math.round(s.you) + s.unit}</span>
                </div>
                <p style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6, margin: "4px 0 8px 0" }}>{s.definition}</p>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                  {s.examples.map((ex) => (
                    <span key={ex} style={{ fontSize: 10.5, color: COLOR.inkMuted, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "3px 10px" }}>{ex}</span>
                  ))}
                </div>
                <div style={{ fontSize: 11, color: COLOR.inkMuted, marginBottom: 6 }}>{s.benchmark}</div>
                <div style={{ fontSize: 12.5, color, fontWeight: 600, display: "flex", alignItems: "flex-start", gap: 6 }}>
                  <ArrowRight size={13} style={{ flexShrink: 0, marginTop: 2 }} /> {s.verdict}
                </div>
              </div>
            );
          })}

          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, paddingTop: 18, marginTop: 4 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: COLOR.ink, marginBottom: 10 }}>Ce que tes ratios racontent</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 }}>
              {report.checks.map((c, i) => (
                <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, color: c.ok ? COLOR.emeraldSoft : COLOR.claySoft, lineHeight: 1.5 }}>
                  <span style={{ flexShrink: 0 }}>{c.ok ? "✅" : "⚠️"}</span>
                  <span>{c.text}</span>
                </div>
              ))}
            </div>
            <div style={{ padding: "12px 14px", background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
              {report.recommendation}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function DiagnosticTab({ transactions, accounts, chargeOverrides, includeGrundfosVoiture, setIncludeGrundfosVoiture, onNavigate, periodRange }: {
  transactions: Transaction[]; accounts: Account[]; chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean;
  setIncludeGrundfosVoiture: (b: boolean) => void; onNavigate?: (tab: Tab, data?: any) => void; periodRange?: [string, string];
}) {
  const [dropPct, setDropPct] = useState(30);
  const [duration, setDuration] = useState(6);
  const [narrativeOpen, setNarrativeOpen] = useState(false);
  const [ratiosNarrativeOpen, setRatiosNarrativeOpen] = useState(false);
  const [xlsState, setXlsState] = useState<"idle" | "loading" | "error">("idle");

  const { ratios, netWorth, essentialMonthly } = useMemo(
    () => computeFinancialRatios(transactions, accounts, chargeOverrides, includeGrundfosVoiture, periodRange),
    [transactions, accounts, chargeOverrides, includeGrundfosVoiture, periodRange]
  );
  const windowMonths = useMemo(() => monthsSinceInception(transactions), [transactions]);
  const charges = useMemo(() => classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, windowMonths, periodRange), [transactions, chargeOverrides, includeGrundfosVoiture, windowMonths, periodRange]);
  const asian = useMemo(() => computeAsianIndicators(transactions, chargeOverrides, includeGrundfosVoiture, periodRange), [transactions, chargeOverrides, includeGrundfosVoiture, periodRange]);
  // Nombre réel de mois couverts par l'analyse ci-dessous (respecte le filtre global
  // "Du mois / Au mois" quand il restreint la période, sinon = tout l'historique).
  const effectiveMonths = charges.lookback.length;

  const accountBalances = useMemo(() => accounts.map((a) => ({ name: a.name, balance: accountBalance(a, transactions) })), [accounts, transactions]);
  const revByCatDetail = useMemo(() => {
    const byCat: Record<string, number> = {};
    transactions.forEach((t) => {
      if (t.type !== "Revenu") return;
      if (!includeGrundfosVoiture && GRUNDFOS_VOITURE_LINKED_REVENUE.includes(t.category)) return;
      if (!charges.lookback.includes(dateToMonthKey(t.date))) return;
      byCat[t.category] = (byCat[t.category] || 0) + t.amount;
    });
    return Object.entries(byCat).map(([category, total]) => ({ category, total })).sort((a, b) => b.total - a.total);
  }, [transactions, includeGrundfosVoiture, charges.lookback]);
  const fixedRows = useMemo(() => charges.rows.filter((r) => r.mode === "fixe").sort((a, b) => b.amount - a.amount), [charges.rows]);

  const [calcDetailKey, setCalcDetailKey] = useState<string | null>(null);
  // Construit le contenu de la fiche de détail pour chaque chiffre affiché sur cette page —
  // même logique que le rapport chat déjà donné à l'utilisateur, rejouée dynamiquement
  // pour toujours refléter le réglage GRUNDFOS et la période sélectionnée en cours.
  const buildCalcDetail = (key: string): { title: string; headline: string; formula: string; blocks: CalcDetailBlock[] } | null => {
    const totalRev = revByCatDetail.reduce((a, r) => a + r.total, 0);
    switch (key) {
      case "epargne_precaution":
        return {
          title: "Épargne de précaution (règle 4-3-2-1)", headline: `${fmt(asian.liquideMonthly)} FCFA · ${(asian.liquideMonthly / asian.avgRevenu * 100).toFixed(1)}%`,
          formula: "Revenu moyen − Dépenses totales moyennes",
          blocks: [{ kind: "kv", rows: [
            { label: `Revenu moyen (${effectiveMonths} mois)`, value: `${fmt(asian.avgRevenu)} FCFA` },
            { label: `Dépenses totales moyennes (${includeGrundfosVoiture ? "GRUNDFOS inclus" : "GRUNDFOS exclu"})`, value: `${fmt(asian.totalDepensesMonthly)} FCFA` },
            { label: "= Épargne de précaution", value: `${fmt(asian.liquideMonthly)} FCFA → ${(asian.liquideMonthly / asian.avgRevenu * 100).toFixed(1)}%`, strong: true, warn: asian.liquideMonthly < 0 },
          ] }],
        };
      case "investissement":
        return {
          title: "Investissement (règle 4-3-2-1)", headline: `${fmt(asian.investMonthly)} FCFA · ${(asian.investMonthly / asian.avgRevenu * 100).toFixed(1)}%`,
          formula: "Somme des catégories d'investissement / mois",
          blocks: [{ kind: "table", columns: ["Catégorie", "Montant/mois (FCFA)"], rows: asian.investDetail.map((d) => [d.category, fmt(d.monthly)]) }],
        };
      case "vie_courante":
        return {
          title: "Vie courante (règle 4-3-2-1)", headline: `${fmt(asian.vieCouranteMonthly)} FCFA · ${(asian.vieCouranteMonthly / asian.avgRevenu * 100).toFixed(1)}%`,
          formula: "Dépenses totales − Investissement − Protection",
          blocks: [{ kind: "kv", rows: [
            { label: "Dépenses totales", value: `${fmt(asian.totalDepensesMonthly)} FCFA` },
            { label: "− Investissement", value: `${fmt(asian.investMonthly)} FCFA` },
            { label: "− Protection", value: `${fmt(asian.protectionMonthly)} FCFA` },
            { label: "= Vie courante", value: `${fmt(asian.vieCouranteMonthly)} FCFA`, strong: true },
          ] }],
        };
      case "protection":
        return {
          title: "Protection / assurance (règle 4-3-2-1)", headline: `${fmt(asian.protectionMonthly)} FCFA · ${(asian.protectionMonthly / asian.avgRevenu * 100).toFixed(1)}%`,
          formula: "Somme des catégories de protection / mois",
          blocks: [{ kind: "table", columns: ["Catégorie", "Montant/mois (FCFA)"], rows: asian.protectionDetail.map((d) => [d.category, fmt(d.monthly)]) }],
        };
      case "taux_epargne":
        return {
          title: "Taux d'épargne (référence institutionnelle)", headline: `${asian.tauxEpargne.toFixed(1)}%`,
          formula: "(Revenu moyen − Charges essentielles) / Revenu moyen",
          blocks: [{ kind: "kv", rows: [
            { label: "Revenu moyen", value: `${fmt(charges.avgRevenu)} FCFA` },
            { label: "Charges fixes", value: `${fmt(charges.totalFixe)} FCFA` },
            { label: "+ Charges variables", value: `${fmt(charges.totalVariable)} FCFA` },
            { label: "= Charges essentielles", value: `${fmt(charges.totalFixe + charges.totalVariable)} FCFA` },
            { label: "Taux d'épargne", value: `${asian.tauxEpargne.toFixed(1)}%`, strong: true },
          ] }],
        };
      case "dti": {
        const hasZero = fixedRows.some((r) => r.amount === 0);
        const blocks: CalcDetailBlock[] = [{
          kind: "table", columns: ["Poste", "Montant/mois (FCFA)", "Origine"],
          rows: fixedRows.map((r) => [r.poste.replace("::", " · "), fmt(r.amount), r.overridden ? "réglage manuel" : "auto"]),
          warnRows: fixedRows.map((r, i) => r.amount === 0 ? i : -1).filter((i) => i >= 0),
        }];
        if (hasZero) blocks.push({ kind: "note", tone: "warn", text: "Au moins un poste classé \"Fixe\" a un montant de 0 FCFA/mois — soit une erreur de saisie dans Charges Fixes & Variables (montant jamais renseigné), soit volontaire. Ça ne fausse rien numériquement, mais autant vérifier." });
        return { title: "Charges fixes / Revenu (DTI)", headline: `${(charges.totalFixe / charges.avgRevenu * 100).toFixed(1)}%`, formula: "Total des postes \"Fixe\" / Revenu moyen", blocks };
      }
      case "logement": {
        const logementRow = charges.rows.find((r) => r.poste === "Logement");
        const amt = logementRow ? logementRow.amount : 0;
        return {
          title: "Logement / Revenu", headline: `${(amt / charges.avgRevenu * 100).toFixed(1)}%`, formula: "Logement (montant retenu) / Revenu moyen",
          blocks: [{ kind: "kv", rows: [
            { label: "Logement (montant retenu)", value: `${fmt(amt)} FCFA` },
            { label: "Revenu moyen", value: `${fmt(charges.avgRevenu)} FCFA` },
            { label: "Logement / Revenu", value: `${(amt / charges.avgRevenu * 100).toFixed(1)}%`, strong: true },
          ] }],
        };
      }
      case "urgence":
        return {
          title: "Fonds d'urgence", headline: `${(netWorth / essentialMonthly).toFixed(1)} mois`, formula: "Valeur nette / Charges essentielles mensuelles",
          blocks: [
            { kind: "table", columns: ["Compte", "Solde (FCFA)"], rows: [...accountBalances.map((a) => [a.name, fmt(a.balance)]), ["Total (Valeur nette)", fmt(netWorth)]] },
            { kind: "kv", rows: [
              { label: "Valeur nette", value: `${fmt(netWorth)} FCFA` },
              { label: "Charges essentielles/mois", value: `${fmt(essentialMonthly)} FCFA` },
              { label: "Fonds d'urgence", value: `${(netWorth / essentialMonthly).toFixed(1)} mois`, strong: true },
            ] },
          ],
        };
      case "concentration": {
        const top = revByCatDetail[0];
        return {
          title: "Concentration des revenus", headline: `${top ? (top.total / totalRev * 100).toFixed(1) : 0}% (${top?.category || "—"})`,
          formula: "Plus grosse source / Revenu total de la période",
          blocks: [{ kind: "table", columns: ["Source", "Montant (FCFA)", "%"], rows: revByCatDetail.map((r) => [r.category, fmt(r.total), `${(r.total / totalRev * 100).toFixed(1)}%`]) }],
        };
      }
      default: return null;
    }
  };

  const exportDiagnosticExcel = async () => {
    setXlsState("loading");
    try {
      const ExcelJS: any = await import(/* @vite-ignore */ "exceljs");
      const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", MUTED = "FF8A9A8E", WHITE = "FFFFFFFF";
      const wb = new ExcelJS.Workbook();
      wb.creator = "Grand Livre"; wb.created = new Date();
      const styleHeaderRow = (row: any) => {
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; c.alignment = { vertical: "middle" }; });
        row.height = 22;
      };
      const totalRev = revByCatDetail.reduce((a, r) => a + r.total, 0);

      // ===== Feuille 1 : Synthèse =====
      const ws1 = wb.addWorksheet("Synthèse");
      ws1.columns = [{ width: 34 }, { width: 22 }, { width: 40 }];
      ws1.mergeCells("A1:C1");
      const title = ws1.getCell("A1");
      title.value = "Grand Livre — Diagnostic Financier, rapport détaillé";
      title.font = { bold: true, size: 15, color: { argb: WHITE } };
      title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
      title.alignment = { vertical: "middle", indent: 1 };
      ws1.getRow(1).height = 30;
      const addSum = (label: string, value: any, color?: string, isAmount?: boolean) => {
        const r = ws1.addRow([label, value]);
        r.getCell(1).font = { color: { argb: MUTED } };
        r.getCell(2).font = { bold: true, size: 12, color: { argb: color || NAVY } };
        if (isAmount) r.getCell(2).numFmt = "#,##0 \"FCFA\"";
        r.getCell(2).alignment = { horizontal: "right" };
      };
      ws1.addRow([]);
      addSum("Généré le", dateLabelFull(todayISO()));
      addSum("Période", `${monthLabel(charges.lookback[0])} — ${monthLabel(charges.lookback[charges.lookback.length - 1])} (${effectiveMonths} mois)`);
      addSum("GRUNDFOS", includeGrundfosVoiture ? "Inclus" : "Exclu (Petty Cash exclu symétriquement)");
      ws1.addRow([]);

      const section = (text: string, color = NAVY) => {
        const r = ws1.addRow([text]);
        ws1.mergeCells(`A${r.number}:C${r.number}`);
        r.getCell(1).font = { bold: true, color: { argb: WHITE } };
        r.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
        r.getCell(1).alignment = { vertical: "middle", indent: 1 };
        r.height = 20;
      };

      section("ÉPARGNE DE PRÉCAUTION (RÈGLE 4-3-2-1)", CLAY);
      addSum("Revenu moyen", asian.avgRevenu, undefined, true);
      addSum("Dépenses totales moyennes", asian.totalDepensesMonthly, undefined, true);
      addSum("= Épargne de précaution", asian.liquideMonthly, asian.liquideMonthly >= 0 ? EMERALD : CLAY, true);
      addSum("  en % du revenu", `${(asian.liquideMonthly / asian.avgRevenu * 100).toFixed(1)}%`, asian.liquideMonthly >= 0 ? EMERALD : CLAY);
      ws1.addRow([]);

      section("RÈGLE 4-3-2-1 — LES 4 POSTES", CLAY);
      addSum("Investissement (cible 40%)", `${fmt(asian.investMonthly)} FCFA · ${(asian.investMonthly / asian.avgRevenu * 100).toFixed(1)}%`);
      addSum("Vie courante (cible 30%)", `${fmt(asian.vieCouranteMonthly)} FCFA · ${(asian.vieCouranteMonthly / asian.avgRevenu * 100).toFixed(1)}%`);
      addSum("Protection (cible 20%)", `${fmt(asian.protectionMonthly)} FCFA · ${(asian.protectionMonthly / asian.avgRevenu * 100).toFixed(1)}%`);
      addSum("Épargne précaution (cible 10%)", `${fmt(asian.liquideMonthly)} FCFA · ${(asian.liquideMonthly / asian.avgRevenu * 100).toFixed(1)}%`);
      ws1.addRow([]);

      section("LES 5 RATIOS INSTITUTIONNELS", GOLD);
      ratios.forEach((r) => {
        const val = r.unit === "mois" ? `${r.value.toFixed(1)} mois` : `${r.value.toFixed(1)}%`;
        addSum(r.label, val, r.verdict === "sain" ? EMERALD : r.verdict === "vigilance" ? GOLD : CLAY);
      });
      ws1.addRow([]);
      addSum("Taux d'épargne : formule", "(Revenu moyen − Charges essentielles) / Revenu moyen");
      addSum("  Charges essentielles", `${fmt(charges.totalFixe)} (fixe) + ${fmt(charges.totalVariable)} (variable) = ${fmt(charges.totalFixe + charges.totalVariable)}`);
      addSum("DTI : formule", "Total postes \"Fixe\" / Revenu moyen");
      addSum("Logement/Revenu : formule", "Logement (montant retenu) / Revenu moyen");
      addSum("Fonds d'urgence : formule", "Valeur nette / Charges essentielles mensuelles");
      addSum("Concentration : formule", "Plus grosse source de revenu / Revenu total de la période");

      // ===== Feuille 2 : Détail 4-3-2-1 =====
      const ws2 = wb.addWorksheet("Détail 4-3-2-1");
      ws2.columns = [{ header: "Poste", key: "poste", width: 20 }, { header: "Catégorie", key: "cat", width: 30 }, { header: "Montant/mois (FCFA)", key: "amt", width: 20 }];
      styleHeaderRow(ws2.getRow(1));
      asian.investDetail.forEach((d) => { const row = ws2.addRow({ poste: "Investissement", cat: d.category, amt: Math.round(d.monthly) }); row.getCell("amt").numFmt = "#,##0"; row.getCell("amt").alignment = { horizontal: "right" }; });
      if (asian.investDetail.length) {
        const subI = ws2.addRow({ poste: "Sous-total Investissement", cat: "", amt: Math.round(asian.investMonthly) });
        subI.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF232F27" } }; });
        subI.getCell("amt").numFmt = "#,##0"; subI.getCell("amt").alignment = { horizontal: "right" };
      }
      asian.protectionDetail.forEach((d) => { const row = ws2.addRow({ poste: "Protection", cat: d.category, amt: Math.round(d.monthly) }); row.getCell("amt").numFmt = "#,##0"; row.getCell("amt").alignment = { horizontal: "right" }; });
      if (asian.protectionDetail.length) {
        const subP = ws2.addRow({ poste: "Sous-total Protection", cat: "", amt: Math.round(asian.protectionMonthly) });
        subP.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF232F27" } }; });
        subP.getCell("amt").numFmt = "#,##0"; subP.getCell("amt").alignment = { horizontal: "right" };
      }
      const vcRow = ws2.addRow({ poste: "Vie courante", cat: "= Dépenses totales − Investissement − Protection", amt: Math.round(asian.vieCouranteMonthly) });
      vcRow.getCell("amt").numFmt = "#,##0"; vcRow.getCell("amt").alignment = { horizontal: "right" }; vcRow.getCell("amt").font = { bold: true };
      const epRow = ws2.addRow({ poste: "Épargne précaution", cat: "= Revenu moyen − Dépenses totales", amt: Math.round(asian.liquideMonthly) });
      epRow.getCell("amt").numFmt = "#,##0"; epRow.getCell("amt").alignment = { horizontal: "right" }; epRow.getCell("amt").font = { bold: true, color: { argb: asian.liquideMonthly >= 0 ? EMERALD : CLAY } };
      const grandTotal2d = ws2.addRow({ poste: "TOTAL GÉNÉRAL (= Revenu moyen)", cat: "", amt: Math.round(asian.investMonthly + asian.protectionMonthly + asian.vieCouranteMonthly + asian.liquideMonthly) });
      grandTotal2d.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      grandTotal2d.getCell("amt").numFmt = "#,##0"; grandTotal2d.getCell("amt").alignment = { horizontal: "right" };

      // ===== Feuille 3 : Détail DTI (postes fixes) =====
      const ws3 = wb.addWorksheet("Détail DTI (charges fixes)");
      ws3.columns = [{ header: "Poste", key: "poste", width: 32 }, { header: "Montant/mois (FCFA)", key: "amt", width: 20 }, { header: "Origine", key: "src", width: 18 }];
      styleHeaderRow(ws3.getRow(1));
      fixedRows.forEach((r) => {
        const row = ws3.addRow({ poste: r.poste.replace("::", " · "), amt: Math.round(r.amount), src: r.overridden ? "réglage manuel" : "auto" });
        row.getCell("amt").numFmt = "#,##0"; row.getCell("amt").alignment = { horizontal: "right" };
        if (r.amount === 0) row.eachCell((c: any) => { c.font = { color: { argb: CLAY } }; });
      });
      const totalFixeRow = ws3.addRow({ poste: "TOTAL", amt: Math.round(charges.totalFixe), src: `${(charges.totalFixe / charges.avgRevenu * 100).toFixed(1)}% du revenu` });
      totalFixeRow.eachCell((c: any) => { c.font = { bold: true }; }); totalFixeRow.getCell("amt").numFmt = "#,##0"; totalFixeRow.getCell("amt").alignment = { horizontal: "right" };

      // ===== Feuille 4 : Fonds d'urgence (comptes) =====
      const ws4 = wb.addWorksheet("Fonds d'urgence");
      ws4.columns = [{ header: "Compte", key: "acc", width: 24 }, { header: "Solde (FCFA)", key: "bal", width: 20 }];
      styleHeaderRow(ws4.getRow(1));
      accountBalances.forEach((a) => { const row = ws4.addRow({ acc: a.name, bal: Math.round(a.balance) }); row.getCell("bal").numFmt = "#,##0"; row.getCell("bal").alignment = { horizontal: "right" }; if (a.balance < 0) row.getCell("bal").font = { color: { argb: CLAY } }; });
      const vnRow = ws4.addRow({ acc: "TOTAL (Valeur nette)", bal: Math.round(netWorth) });
      vnRow.eachCell((c: any) => { c.font = { bold: true }; }); vnRow.getCell("bal").numFmt = "#,##0"; vnRow.getCell("bal").alignment = { horizontal: "right" };
      ws4.addRow([]);
      const eRow = ws4.addRow({ acc: "Charges essentielles/mois", bal: Math.round(essentialMonthly) });
      eRow.getCell("bal").numFmt = "#,##0"; eRow.getCell("bal").alignment = { horizontal: "right" };
      const fuRow = ws4.addRow({ acc: "Fonds d'urgence (mois)", bal: `${(netWorth / essentialMonthly).toFixed(1)} mois` as any });
      fuRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; });

      // ===== Feuille 5 : Concentration des revenus =====
      const ws5 = wb.addWorksheet("Concentration revenus");
      ws5.columns = [{ header: "Source", key: "src", width: 28 }, { header: "Montant (FCFA)", key: "amt", width: 20 }, { header: "%", key: "pct", width: 12 }];
      styleHeaderRow(ws5.getRow(1));
      revByCatDetail.forEach((r) => {
        const row = ws5.addRow({ src: r.category, amt: Math.round(r.total), pct: `${(r.total / totalRev * 100).toFixed(1)}%` });
        row.getCell("amt").numFmt = "#,##0"; row.getCell("amt").alignment = { horizontal: "right" }; row.getCell("pct").alignment = { horizontal: "right" };
      });
      const totalRevRow = ws5.addRow({ src: "TOTAL", amt: Math.round(totalRev), pct: "100%" });
      totalRevRow.eachCell((c: any) => { c.font = { bold: true }; });

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `grand-livre_diagnostic-financier-detaille_${todayISO()}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      setXlsState("idle");
    } catch (e) {
      console.error(e);
      setXlsState("error");
    }
  };

  const verdictStyle: Record<RatioVerdict, { color: string; bg: string; label: string; icon: any }> = {
    sain: { color: COLOR.emeraldSoft, bg: "rgba(63,156,122,0.1)", label: "Sain", icon: Check },
    vigilance: { color: COLOR.goldSoft, bg: "rgba(201,162,39,0.1)", label: "Vigilance", icon: AlertTriangle },
    risque: { color: COLOR.claySoft, bg: "rgba(193,84,63,0.1)", label: "Risque", icon: AlertTriangle },
  };
  const overallVerdict: RatioVerdict = ratios.some((r) => r.verdict === "risque") ? "risque" : ratios.some((r) => r.verdict === "vigilance") ? "vigilance" : "sain";

  // Simulateur de résilience : impact d'une baisse de revenu de X% pendant N mois.
  const reducedRevenu = charges.avgRevenu * (1 - dropPct / 100);
  const monthlyDeficit = essentialMonthly - reducedRevenu;
  const netWorthAfter = netWorth - monthlyDeficit * duration;
  const monthsToDepletion = monthlyDeficit > 0 ? netWorth / monthlyDeficit : Infinity;
  const survives = netWorthAfter >= 0;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 12, padding: "12px 16px", flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 4, background: COLOR.surface, borderRadius: 16, padding: 3, border: `1px solid ${COLOR.hairline}` }}>
          <button onClick={() => setIncludeGrundfosVoiture(true)} style={{
            padding: "6px 14px", borderRadius: 12, fontSize: 12, cursor: "pointer", border: "none",
            background: includeGrundfosVoiture ? COLOR.gold : "transparent", color: includeGrundfosVoiture ? COLOR.bg : COLOR.inkMuted,
          }}>Inclure GRUNDFOS</button>
          <button onClick={() => setIncludeGrundfosVoiture(false)} style={{
            padding: "6px 14px", borderRadius: 12, fontSize: 12, cursor: "pointer", border: "none",
            background: !includeGrundfosVoiture ? COLOR.gold : "transparent", color: !includeGrundfosVoiture ? COLOR.bg : COLOR.inkMuted,
          }}>Exclure GRUNDFOS</button>
        </div>
        <span style={{ fontSize: 11.5, color: COLOR.inkMuted, flex: 1, minWidth: 220 }}>
          {includeGrundfosVoiture
            ? "GRUNDFOS et le revenu qui la finance (Petty Cash) sont inclus dans tous les calculs de cette page. \"Voiture\" reste toujours incluse, quelle que soit l'option choisie."
            : "Vue \"personnel pur\" : GRUNDFOS ET le revenu qui la finance (Petty Cash) sont exclus symétriquement de tous les calculs ci-dessous. \"Voiture\" reste toujours incluse."}
        </span>
        <button onClick={exportDiagnosticExcel} disabled={xlsState === "loading"} style={{
          display: "flex", alignItems: "center", gap: 6, background: xlsState === "error" ? "rgba(193,84,63,0.14)" : "rgba(63,156,122,0.14)", border: `1px solid ${xlsState === "error" ? COLOR.clay : COLOR.emerald}`,
          borderRadius: 8, color: xlsState === "error" ? COLOR.claySoft : COLOR.emeraldSoft, padding: "8px 14px", fontSize: 12, cursor: xlsState === "loading" ? "default" : "pointer", flexShrink: 0,
        }}>
          {xlsState === "loading" ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <FileSpreadsheet size={13} />} {xlsState === "loading" ? "Génération…" : xlsState === "error" ? "Réessayer" : "Rapport Excel complet"}
        </button>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 14, background: `linear-gradient(135deg, ${verdictStyle[overallVerdict].bg} 0%, ${COLOR.surfaceRaised} 70%)`, border: `1px solid ${verdictStyle[overallVerdict].color}`, borderRadius: 14, padding: "16px 20px", flexWrap: "wrap" }}>
        <Gauge size={22} color={verdictStyle[overallVerdict].color} />
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, fontWeight: 600, color: verdictStyle[overallVerdict].color }}>
            Diagnostic global : {verdictStyle[overallVerdict].label}
          </div>
          <div style={{ fontSize: 12, color: COLOR.inkMuted, marginTop: 2 }}>
            Basé sur 5 ratios utilisés par les banques, régulateurs financiers et cabinets de conseil en gestion de patrimoine — appliqués à la période sélectionnée ({effectiveMonths} mois{effectiveMonths === windowMonths ? ", tout ton historique" : ""}).
          </div>
        </div>
        <button onClick={() => setRatiosNarrativeOpen(true)} style={{
          display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`,
          borderRadius: 8, color: COLOR.goldSoft, padding: "9px 16px", fontSize: 12.5, cursor: "pointer", flexShrink: 0,
        }}>
          <BookOpen size={13} /> Rapport détaillé
        </button>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14 }}>
        {ratios.map((r) => {
          const s = verdictStyle[r.verdict];
          const Icon = s.icon;
          return (
            <div key={r.key} style={{ background: COLOR.surfaceRaised, border: `1px solid ${s.color}`, borderRadius: 12, padding: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                <span style={{ fontSize: 12.5, color: COLOR.inkMuted, fontWeight: 600 }}>
                  {r.label}
                  {["epargne", "dti", "logement", "urgence", "concentration"].includes(r.key) && (
                    <CalcDetailIcon onClick={() => setCalcDetailKey(r.key === "epargne" ? "taux_epargne" : r.key)} />
                  )}
                </span>
                <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: s.color, background: s.bg, borderRadius: 20, padding: "2px 8px" }}>
                  <Icon size={11} /> {s.label}
                </span>
              </div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 700, color: s.color, marginBottom: 6 }}>
                {r.unit === "mois" ? r.value.toFixed(1) : Math.round(r.value)}
                <span style={{ fontSize: 13, fontWeight: 500, marginLeft: 3 }}>{r.unit === "mois" ? " mois" : r.unit}</span>
              </div>
              <div style={{ fontSize: 11, color: COLOR.inkMuted, marginBottom: 4 }}>{r.benchmark}</div>
              <div style={{ fontSize: 11.5, color: COLOR.inkMuted, lineHeight: 1.5, marginBottom: (r.key === "dti" || r.key === "logement") && onNavigate ? 8 : 0 }}>{r.explain}</div>
              {(r.key === "dti" || r.key === "logement") && onNavigate && (
                <button onClick={() => onNavigate("charges")} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", fontSize: 11, padding: 0 }}>
                  Voir le détail dans Charges Fixes & Variables <ArrowRight size={11} />
                </button>
              )}
            </div>
          );
        })}
      </div>

      <PanelWithHelp title="Simulateur de résilience" subtitle="Stress test simplifié : que se passerait-il si tes revenus baissaient pendant plusieurs mois ?"
        collapsible defaultOpen={false}
        explain="Méthode inspirée des tests de résistance utilisés par les régulateurs bancaires (ex. stress tests de la Fed ou de la BCE), adaptée à un budget personnel : on simule une baisse de revenu pendant une durée donnée, charges fixes et variables régulières inchangées, et on regarde ce qu'il resterait de ta valeur nette actuelle à la fin de la période.">
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ fontSize: 12.5, color: COLOR.inkMuted }}>Baisse de revenu simulée</span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.goldSoft, fontWeight: 600 }}>{dropPct}%</span>
            </div>
            <input type="range" min={0} max={100} step={5} value={dropPct} onChange={(e) => setDropPct(Number(e.target.value))} style={{ width: "100%", accentColor: COLOR.gold }} />
          </div>
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ fontSize: 12.5, color: COLOR.inkMuted }}>Durée</span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLOR.goldSoft, fontWeight: 600 }}>{duration} mois</span>
            </div>
            <input type="range" min={1} max={24} step={1} value={duration} onChange={(e) => setDuration(Number(e.target.value))} style={{ width: "100%", accentColor: COLOR.gold }} />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginTop: 4 }}>
            <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Revenu réduit simulé / mois</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLOR.ink }}>{fmt(reducedRevenu)} FCFA</div>
            </div>
            <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Déficit mensuel pendant le choc</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: monthlyDeficit > 0 ? COLOR.claySoft : COLOR.emeraldSoft }}>
                {monthlyDeficit > 0 ? `−${fmt(monthlyDeficit)}` : `+${fmt(-monthlyDeficit)}`} FCFA
              </div>
            </div>
            <div style={{ background: COLOR.surface, border: `1px solid ${survives ? COLOR.emerald : COLOR.clay}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Valeur nette après {duration} mois de choc</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: survives ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(netWorthAfter)} FCFA</div>
            </div>
            <div style={{ background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Tiendrait combien de temps au total</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLOR.ink }}>
                {monthlyDeficit <= 0 ? "Indéfiniment (pas de déficit)" : isFinite(monthsToDepletion) ? `~${monthsToDepletion.toFixed(1)} mois` : "—"}
              </div>
            </div>
          </div>

          <div style={{ padding: "12px 14px", background: survives ? "rgba(63,156,122,0.08)" : "rgba(193,84,63,0.08)", border: `1px solid ${survives ? COLOR.emerald : COLOR.clay}`, borderRadius: 8, fontSize: 12.5, color: survives ? COLOR.emeraldSoft : COLOR.claySoft, lineHeight: 1.6 }}>
            {survives
              ? `Avec une baisse de revenu de ${dropPct}% pendant ${duration} mois, ta valeur nette actuelle absorberait le choc sans être épuisée.`
              : `Avec une baisse de revenu de ${dropPct}% pendant ${duration} mois, ta valeur nette actuelle serait épuisée avant la fin de la période (~${isFinite(monthsToDepletion) ? monthsToDepletion.toFixed(1) : "?"} mois de marge réelle).`}
          </div>
        </div>
      </PanelWithHelp>

      <PanelWithHelp title="Indicateurs asiatiques de gestion financière" subtitle="Règle chinoise du 4-3-2-1, norme d'épargne chinoise, méthode Kakeibo japonaise"
        collapsible defaultOpen={false}
        right={
          <button onClick={(e) => { e.stopPropagation(); setNarrativeOpen(true); }} style={{
            display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`,
            borderRadius: 8, color: COLOR.goldSoft, padding: "7px 14px", fontSize: 12.5, cursor: "pointer",
          }}>
            <BookOpen size={13} /> Rapport détaillé
          </button>
        }
        explain="Ces cadres viennent de traditions de gestion financière personnelle distinctes des repères bancaires occidentaux ci-dessus. La règle du 4-3-2-1 (家庭资产配置法则) est enseignée dans les certifications chinoises de planification financière (AFP Chine). Le taux d'épargne chinois reflète le niveau d'épargne traditionnellement très élevé des ménages en Chine (30 à 45%, contre ~20% recommandé en Occident). Le Kakeibo (家計簿) est une méthode budgétaire japonaise créée en 1904 par Hani Motoko, toujours largement utilisée aujourd'hui — elle classe chaque dépense en 4 catégories plutôt que de suivre un simple total, pour favoriser la réflexion plutôt que le seul chiffrage.">
        <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>

          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: COLOR.ink, marginBottom: 4 }}>Règle chinoise du 4-3-2-1 (allocation du revenu mensuel)</div>
            <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginBottom: 12 }}>Cible : 40% investissement · 30% vie courante · 20% protection/assurance · 10% épargne de précaution</div>
            {asian.hasDeficit && (
              <div style={{ display: "flex", gap: 8, alignItems: "flex-start", background: "rgba(193,84,63,0.1)", border: `1px solid ${COLOR.clay}`, borderRadius: 8, padding: "10px 12px", marginBottom: 14 }}>
                <AlertTriangle size={14} color={COLOR.claySoft} style={{ flexShrink: 0, marginTop: 1 }} />
                <div style={{ fontSize: 12, color: COLOR.claySoft, lineHeight: 1.5 }}>
                  <strong>Déficit réel sur cette période :</strong> les dépenses moyennes dépassent le revenu moyen. "Épargne de précaution" est donc négative — ce n'est pas une erreur d'affichage, c'est le signe qu'il n'y a structurellement rien à mettre de côté sur cette période, et même un prélèvement sur l'existant.
                </div>
              </div>
            )}
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {asian.rule4321.map((r) => {
                const diff = r.value - r.target;
                const aligned = Math.abs(diff) <= 5;
                const color = r.value < 0 ? COLOR.claySoft : aligned ? COLOR.emeraldSoft : Math.abs(diff) <= 15 ? COLOR.goldSoft : COLOR.claySoft;
                const detailKey = r.label === "Investissement" ? "investissement" : r.label === "Vie courante" ? "vie_courante" : r.label === "Protection / assurance" ? "protection" : "epargne_precaution";
                return (
                  <div key={r.label}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4, fontSize: 12.5 }}>
                      <span style={{ color: COLOR.ink }}>{r.label}<CalcDetailIcon onClick={() => setCalcDetailKey(detailKey)} /></span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color }}>{r.value.toFixed(0)}% <span style={{ color: COLOR.inkMuted }}>(cible {r.target}%)</span></span>
                    </div>
                    <div style={{ position: "relative", height: 8, background: COLOR.hairline, borderRadius: 4, overflow: "hidden" }}>
                      <div style={{ position: "absolute", left: `${r.target}%`, top: 0, bottom: 0, width: 2, background: COLOR.inkMuted, zIndex: 2 }} />
                      <div style={{ height: "100%", width: `${Math.max(0, Math.min(100, r.value))}%`, background: color }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, paddingTop: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: COLOR.ink }}>Taux d'épargne — norme chinoise</div>
                <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 2 }}>Référence : 30 à 45% (moyenne des ménages chinois) — nettement plus exigeante que le standard occidental de 20%</div>
              </div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 700, color: asian.tauxEpargne >= 30 ? COLOR.emeraldSoft : asian.tauxEpargne >= 20 ? COLOR.goldSoft : COLOR.claySoft }}>
                {asian.tauxEpargne.toFixed(0)}%
              </div>
            </div>
          </div>

          <div style={{ borderTop: `1px solid ${COLOR.hairline}`, paddingTop: 20 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: COLOR.ink, marginBottom: 4 }}>Répartition Kakeibo (méthode japonaise)</div>
            <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginBottom: 12 }}>Chaque dépense classée en 4 questions plutôt qu'un simple total — moyenne mensuelle sur la période sélectionnée ({effectiveMonths} mois)</div>
            <div style={{ display: "flex", height: 10, borderRadius: 5, overflow: "hidden", marginBottom: 10 }}>
              {asian.kakeibo.map((k, i) => (
                <div key={k.key} style={{ width: `${k.pct}%`, background: [COLOR.slateBlue, COLOR.clay, COLOR.gold, COLOR.emerald][i] }} title={`${k.label} : ${k.pct.toFixed(0)}%`} />
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10 }}>
              {asian.kakeibo.map((k, i) => (
                <div key={k.key} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12 }}>
                  <span style={{ width: 8, height: 8, borderRadius: "50%", background: [COLOR.slateBlue, COLOR.clay, COLOR.gold, COLOR.emerald][i], flexShrink: 0 }} />
                  <span style={{ color: COLOR.inkMuted }}>{k.label}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.ink, fontWeight: 600, marginLeft: "auto" }}>{k.pct.toFixed(0)}%</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </PanelWithHelp>

      <NarrativeReportSheet
        open={narrativeOpen}
        onClose={() => setNarrativeOpen(false)}
        rule4321={asian.rule4321}
        tauxEpargne={asian.tauxEpargne}
        kakeibo={asian.kakeibo}
      />
      <RatiosNarrativeSheet open={ratiosNarrativeOpen} onClose={() => setRatiosNarrativeOpen(false)} ratios={ratios} />
      {calcDetailKey && (() => {
        const d = buildCalcDetail(calcDetailKey);
        return d ? <CalcDetailSheet open={!!calcDetailKey} onClose={() => setCalcDetailKey(null)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks} /> : null;
      })()}
    </div>
  );
}

function CreancesTab({ loans, setLoans }: { loans: Loan[]; setLoans: (l: Loan[]) => void }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Loan, "id">>({ person: "", amount: 0, dateGiven: "2026_8", status: "En attente", notes: "" });
  const [repayForm, setRepayForm] = useState<{ loanId: string; amount: number; date: string; note: string } | null>(null);
  const [expandedId, setExpandedId] = useState<string | null>(null);

  // Le statut se déduit désormais du montant réellement remboursé plutôt que d'être
  // basculé à la main — sur demande explicite de l'utilisateur (10/08/2026) : gérer des
  // remboursements partiels, pas juste "tout ou rien".
  const repaidOf = (l: Loan) => (l.repayments || []).reduce((a, r) => a + r.amount, 0);
  const remainingOf = (l: Loan) => Math.max(0, l.amount - repaidOf(l));
  const statusOf = (l: Loan): LoanStatus => {
    const repaid = repaidOf(l);
    if (repaid <= 0) return "En attente";
    if (repaid >= l.amount) return "Remboursé";
    return "Partiellement remboursé";
  };

  const add = () => { if (!form.person || form.amount <= 0) return; setLoans([...loans, { ...form, id: uid("l"), repayments: [] }]); setForm({ person: "", amount: 0, dateGiven: "2026_8", status: "En attente", notes: "" }); setAdding(false); };
  const remove = (id: string) => setLoans(loans.filter((l) => l.id !== id));

  const addRepayment = () => {
    if (!repayForm || repayForm.amount <= 0) return;
    setLoans(loans.map((l) => {
      if (l.id !== repayForm.loanId) return l;
      const repayments = [...(l.repayments || []), { id: uid("rp"), date: repayForm.date, amount: repayForm.amount, note: repayForm.note || undefined }];
      return { ...l, repayments, status: statusOf({ ...l, repayments }) };
    }));
    setRepayForm(null);
  };
  const removeRepayment = (loanId: string, repaymentId: string) => {
    setLoans(loans.map((l) => {
      if (l.id !== loanId) return l;
      const repayments = (l.repayments || []).filter((r) => r.id !== repaymentId);
      return { ...l, repayments, status: statusOf({ ...l, repayments }) };
    }));
  };

  const totalOutstanding = loans.reduce((a, l) => a + remainingOf(l), 0);
  const totalRepaid = loans.reduce((a, l) => a + repaidOf(l), 0);
  const statusStyle: Record<LoanStatus, { color: string; border: string }> = {
    "En attente": { color: COLOR.goldSoft, border: COLOR.gold },
    "Partiellement remboursé": { color: COLOR.slateBlueSoft, border: COLOR.slateBlue },
    "Remboursé": { color: COLOR.emeraldSoft, border: COLOR.emerald },
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Créances en attente (reste dû)" value={fmt(totalOutstanding)} tone={COLOR.gold} icon={HandCoins} />
        <Kpi label="Total remboursé à date" value={fmt(totalRepaid)} tone={COLOR.emeraldSoft} icon={Check} />
      </div>
      <Panel title="Suivi des prêts accordés" subtitle="Ces montants ne sont pas des dépenses perdues — ce sont des créances récupérables, remboursables en plusieurs fois"
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
          {loans.map((l) => {
            const status = statusOf(l);
            const repaid = repaidOf(l);
            const remaining = remainingOf(l);
            const isExpanded = expandedId === l.id;
            const repayments = l.repayments || [];
            return (
              <div key={l.id} style={{ background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}`, overflow: "hidden" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", flexWrap: "wrap", gap: 10 }}>
                  <div style={{ flex: 1, minWidth: 200 }}>
                    <div style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                      {l.person} <span style={{ color: COLOR.inkMuted, fontSize: 11.5 }}>· {monthLabel(l.dateGiven)}</span>
                      {repayments.length > 0 && (
                        <button onClick={() => setExpandedId(isExpanded ? null : l.id)} style={{ background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", fontSize: 11, display: "flex", alignItems: "center", gap: 2, padding: 0 }}>
                          {repayments.length} remboursement(s) {isExpanded ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                        </button>
                      )}
                    </div>
                    {l.notes && <div style={{ fontSize: 11.5, color: COLOR.inkMuted, marginTop: 2 }}>{l.notes}</div>}
                    {status !== "En attente" && (
                      <div style={{ marginTop: 6, display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ flex: 1, maxWidth: 160, height: 5, background: COLOR.hairline, borderRadius: 3, overflow: "hidden" }}>
                          <div style={{ height: "100%", width: `${Math.min(100, (repaid / l.amount) * 100)}%`, background: status === "Remboursé" ? COLOR.emerald : COLOR.slateBlue }} />
                        </div>
                        <span style={{ fontSize: 10.5, color: COLOR.inkMuted }}>{fmt(repaid)} / {fmt(l.amount)} FCFA</span>
                      </div>
                    )}
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}>{fmt(l.amount)}</div>
                      {status !== "Remboursé" && <div style={{ fontSize: 10.5, color: COLOR.claySoft }}>reste {fmt(remaining)}</div>}
                    </div>
                    <span style={{ fontSize: 11, padding: "4px 10px", borderRadius: 12, border: `1px solid ${statusStyle[status].border}`, color: statusStyle[status].color }}>{status}</span>
                    {status !== "Remboursé" && (
                      <button onClick={() => setRepayForm({ loanId: l.id, amount: remaining, date: todayISO(), note: "" })} title="Enregistrer un remboursement" style={{ display: "flex", alignItems: "center", gap: 5, background: "rgba(63,156,122,0.14)", border: `1px solid ${COLOR.emerald}`, borderRadius: 6, color: COLOR.emeraldSoft, padding: "5px 10px", fontSize: 11.5, cursor: "pointer" }}>
                        <Plus size={12} /> Rembourser
                      </button>
                    )}
                    <button onClick={() => remove(l.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
                  </div>
                </div>
                {repayForm?.loanId === l.id && (
                  <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", padding: 14, background: COLOR.surface, borderTop: `1px solid ${COLOR.hairline}` }}>
                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant remboursé</label><input type="number" inputMode="numeric" style={{ ...inputStyle, width: 130 }} value={repayForm.amount || ""} onChange={(e) => setRepayForm({ ...repayForm, amount: Number(e.target.value) })} /></div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Date</label><input type="date" style={{ ...inputStyle, width: 150 }} value={repayForm.date} onChange={(e) => setRepayForm({ ...repayForm, date: e.target.value })} /></div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Note (optionnel)</label><input style={{ ...inputStyle, width: 160 }} value={repayForm.note} onChange={(e) => setRepayForm({ ...repayForm, note: e.target.value })} /></div>
                    <button onClick={addRepayment} style={{ display: "flex", alignItems: "center", gap: 6, background: COLOR.emerald, border: "none", borderRadius: 6, color: COLOR.bg, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 32 }}><Save size={13} /> Confirmer</button>
                    <button onClick={() => setRepayForm(null)} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "8px 14px", fontSize: 12.5, cursor: "pointer", height: 32 }}>Annuler</button>
                  </div>
                )}
                {isExpanded && repayments.length > 0 && (
                  <div style={{ padding: "8px 14px 12px", borderTop: `1px solid ${COLOR.hairline}`, display: "flex", flexDirection: "column", gap: 4 }}>
                    {repayments.slice().sort((a, b) => b.date.localeCompare(a.date)).map((r) => (
                      <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 11.5, padding: "4px 0" }}>
                        <span style={{ color: COLOR.inkMuted }}>{dateLabelFull(r.date)}{r.note ? ` — ${r.note}` : ""}</span>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLOR.emeraldSoft }}>+{fmt(r.amount)}</span>
                          <button onClick={() => removeRepayment(l.id, r.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={11} /></button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
          {!loans.length && <EmptyState text="Aucune créance enregistrée." />}
        </div>
      </Panel>
    </div>
  );
}
// ============================================================
// COMPTES (ACCOUNTS)
// ============================================================
function ComptesTab({ accounts, setAccounts, transactions, setTransactions }: { accounts: Account[]; setAccounts: (a: Account[]) => void; transactions: Transaction[]; setTransactions: (t: Transaction[]) => void }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Account, "id">>({ name: "", kind: "Banque", openingBalance: 0 });
  const [editingOpening, setEditingOpening] = useState<string | null>(null);
  const kinds: Account["kind"][] = ["Espèces", "Banque", "Mobile Money", "Carte de crédit", "Autre"];
  const [merging, setMerging] = useState(false);
  const [mergeSourceId, setMergeSourceId] = useState("");
  const [mergeTargetId, setMergeTargetId] = useState("");
  const [confirmMerge, setConfirmMerge] = useState(false);

  const add = () => { if (!form.name) return; setAccounts([...accounts, { ...form, id: uid("a") }]); setForm({ name: "", kind: "Banque", openingBalance: 0 }); setAdding(false); };
  const update = (id: string, patch: Partial<Account>) => setAccounts(accounts.map((a) => (a.id === id ? { ...a, ...patch } : a)));
  const remove = (id: string) => setAccounts(accounts.filter((a) => a.id !== id));
  const total = totalAccountsBalance(accounts, transactions);

  // Fusion de deux comptes en un seul, en conservant le nom du compte "cible" — sur
  // demande explicite de l'utilisateur (10/08/2026). Ce n'est pas un simple renommage :
  // il faut aussi additionner les soldes de départ (sinon l'argent du compte absorbé
  // disparaît), réattribuer toute transaction qui référençait le compte absorbé comme
  // avance (onBehalfOf), et nettoyer les avances qui deviendraient incohérentes (un
  // compte ne peut pas s'être avancé de l'argent à lui-même une fois fusionné).
  const mergeAccounts = () => {
    const source = accounts.find((a) => a.id === mergeSourceId);
    const target = accounts.find((a) => a.id === mergeTargetId);
    if (!source || !target || source.id === target.id) return;
    setTransactions(transactions.map((t) => {
      let next = t;
      if (next.account === source.name) next = { ...next, account: target.name };
      if (next.onBehalfOf === source.name) next = { ...next, onBehalfOf: target.name };
      if (next.account === next.onBehalfOf) next = { ...next, onBehalfOf: undefined };
      return next;
    }));
    setAccounts(accounts.filter((a) => a.id !== source.id).map((a) => a.id === target.id ? { ...a, openingBalance: a.openingBalance + source.openingBalance } : a));
    setMerging(false); setMergeSourceId(""); setMergeTargetId(""); setConfirmMerge(false);
  };

  // Consommation par compte sur une période — indépendant du solde temps réel ci-dessus.
  const withMonth = useMemo(() => transactions.map((t) => ({ ...t, month: dateToMonthKey(t.date) })), [transactions]);
  const allMonths = useMemo(() => Array.from(new Set(withMonth.map((t) => t.month))).sort((a, b) => monthSortKey(a) - monthSortKey(b)), [withMonth]);
  const lastMonth = allMonths[allMonths.length - 1] || dateToMonthKey(todayISO());
  const [periodFrom, setPeriodFrom] = useState(allMonths[0] || lastMonth);
  const [periodTo, setPeriodTo] = useState(lastMonth);
  // Empêche une plage inversée (Du postérieur à Au) — échange les deux valeurs plutôt
  // que de laisser une plage impossible s'installer.
  const setPeriodFromSafe = (v: string) => { if (monthSortKey(v) > monthSortKey(periodTo)) { setPeriodFrom(periodTo); setPeriodTo(v); } else setPeriodFrom(v); };
  const setPeriodToSafe = (v: string) => { if (monthSortKey(v) < monthSortKey(periodFrom)) { setPeriodTo(periodFrom); setPeriodFrom(v); } else setPeriodTo(v); };

  const applyPreset = (key: string) => {
    if (key === "1m") { setPeriodFrom(lastMonth); setPeriodTo(lastMonth); }
    else if (key === "3m") { setPeriodFrom(trailingRange(allMonths, 3)[0]); setPeriodTo(lastMonth); }
    else if (key === "6m") { setPeriodFrom(trailingRange(allMonths, 6)[0]); setPeriodTo(lastMonth); }
    else if (key === "1a") { setPeriodFrom(trailingRange(allMonths, 12)[0]); setPeriodTo(lastMonth); }
    else if (key === "tout") { setPeriodFrom(allMonths[0] || lastMonth); setPeriodTo(lastMonth); }
  };

  const periodTx = useMemo(() => {
    const fk = monthSortKey(periodFrom), tk = monthSortKey(periodTo);
    return withMonth.filter((t) => { const k = monthSortKey(t.month); return k >= fk && k <= tk; });
  }, [withMonth, periodFrom, periodTo]);

  const [accountNarrativeId, setAccountNarrativeId] = useState<string | null>(null);
  const [xlsState, setXlsState] = useState<"idle" | "loading" | "error">("idle");
  const consoParAccount = useMemo(() => accounts.map((a) => {
    // Consommation réelle du compte : une dépense marquée "avance" (onBehalfOf) est
    // rattachée au compte BÉNÉFICIAIRE ici, pas au compte qui a réellement payé — sur
    // demande explicite de l'utilisateur (10/08/2026), pour que "qu'est-ce que ce compte
    // finance vraiment" ne soit pas faussé par un dépannage ponctuel entre comptes.
    // Le solde réel (soldeDebut/soldeFin plus bas) reste lui basé sur le compte qui a
    // vraiment payé, sans aucune réattribution — un solde ne doit jamais mentir.
    // Consommation ET revenus réels du compte : une transaction marquée "avance"
    // (onBehalfOf) est rattachée au compte BÉNÉFICIAIRE ici — que ce soit une dépense
    // payée pour un autre compte, ou un revenu encaissé qui appartient en fait à un
    // autre compte — sur demande explicite de l'utilisateur (10/08/2026). Le solde réel
    // (soldeDebut/soldeFin plus bas) reste lui basé sur le compte réellement mouvementé,
    // sans aucune réattribution — un solde ne doit jamais mentir.
    const tx = periodTx.filter((t) => (t.onBehalfOf || t.account) === a.name);
    const depenses = tx.filter((t) => t.type === "Dépense").reduce((s, t) => s + t.amount, 0);
    const revenus = tx.filter((t) => t.type === "Revenu").reduce((s, t) => s + t.amount, 0);
    const depByCat: Record<string, number> = {};
    const revByCat: Record<string, number> = {};
    tx.forEach((t) => { (t.type === "Dépense" ? depByCat : revByCat)[t.category] = ((t.type === "Dépense" ? depByCat : revByCat)[t.category] || 0) + t.amount; });
    // Montant reçu en avance d'un autre compte (déjà inclus dans "depenses"/"revenus"
    // ci-dessus, mais isolé ici pour l'affichage — utile pour comprendre l'écart avec le
    // solde réel, quel que soit le type de la transaction).
    const receivedAsAdvance = tx.filter((t) => t.account !== a.name).reduce((s, t) => s + t.amount, 0);
    // Solde réel du compte (pas juste le mouvement net de la période) : ce qu'il y avait
    // avant le début de la période, et ce qu'il y a à la fin — pour voir l'évolution du
    // solde effectif, pas seulement ce qui a transité pendant la fenêtre choisie.
    const beforePeriod = withMonth.filter((t) => t.account === a.name && monthSortKey(t.month) < monthSortKey(periodFrom));
    const uptoEndOfPeriod = withMonth.filter((t) => t.account === a.name && monthSortKey(t.month) <= monthSortKey(periodTo));
    const soldeDebut = accountBalance(a, beforePeriod);
    const soldeFin = accountBalance(a, uptoEndOfPeriod);
    const byMonth: Record<string, number> = {};
    tx.forEach((t) => { byMonth[t.month] = (byMonth[t.month] || 0) + (t.type === "Revenu" ? t.amount : -t.amount); });
    return { account: a, depenses, revenus, net: revenus - depenses, count: tx.length, depByCat, revByCat, soldeDebut, soldeFin, byMonth, receivedAsAdvance };
  }).sort((x, y) => y.depenses - x.depenses), [accounts, periodTx, withMonth, periodFrom, periodTo]);

  const totalConsoDep = consoParAccount.reduce((a, c) => a + c.depenses, 0);

  const exportComptesExcel = async () => {
    setXlsState("loading");
    try {
      const ExcelJS: any = await import(/* @vite-ignore */ "exceljs");
      const NAVY = "FF1A2B4C", GOLD = "FFC9A227", EMERALD = "FF3F9C7A", CLAY = "FFC1543F", WHITE = "FFFFFFFF", MUTED = "FF8A9A8E";
      const wb = new ExcelJS.Workbook();
      wb.creator = "Grand Livre"; wb.created = new Date();
      const periodLabel = `${monthLabel(periodFrom)} — ${monthLabel(periodTo)}`;

      const styleHeaderRow = (row: any) => {
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; c.alignment = { vertical: "middle" }; });
        row.height = 22;
      };
      const titleBanner = (ws: any, title: string) => {
        ws.mergeCells("A1:F1");
        const t = ws.getCell("A1");
        t.value = `Grand Livre — ${title}`;
        t.font = { bold: true, size: 14, color: { argb: WHITE } };
        t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
        t.alignment = { vertical: "middle", indent: 1 };
        ws.getRow(1).height = 28;
        ws.mergeCells("A2:F2");
        const s = ws.getCell("A2");
        s.value = `Période : ${periodLabel} · Généré le ${dateLabelFull(todayISO())}`;
        s.font = { italic: true, color: { argb: MUTED } };
        ws.addRow([]);
      };

      // ===== Feuille 1 : Synthèse — soldes début/fin + mouvement de la période =====
      const ws1 = wb.addWorksheet("Synthèse");
      titleBanner(ws1, "Consommation par compte");
      const headerRow1 = ws1.addRow(["Compte", "Solde début période", "Reçu (période)", "Consommé (période)", "Mouvement net", "Solde fin période", "Transactions"]);
      styleHeaderRow(headerRow1);
      ws1.columns = [{ width: 20 }, { width: 20 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 20 }, { width: 14 }];
      consoParAccount.forEach((c) => {
        const row = ws1.addRow([c.account.name, c.soldeDebut, c.revenus, c.depenses, c.net, c.soldeFin, c.count]);
        row.getCell(2).font = { color: { argb: MUTED } };
        row.getCell(3).font = { color: { argb: EMERALD } };
        row.getCell(4).font = { color: { argb: CLAY } };
        row.getCell(5).font = { bold: true, color: { argb: c.net >= 0 ? EMERALD : CLAY } };
        row.getCell(6).font = { bold: true, color: { argb: c.soldeFin >= 0 ? EMERALD : CLAY } };
        [2, 3, 4, 5, 6].forEach((i) => { row.getCell(i).numFmt = "#,##0"; row.getCell(i).alignment = { horizontal: "right" }; });
        row.getCell(7).alignment = { horizontal: "center" };
      });
      ws1.addRow([]);
      const totDebut = consoParAccount.reduce((a, c) => a + c.soldeDebut, 0), totFin = consoParAccount.reduce((a, c) => a + c.soldeFin, 0);
      const totRowC = ws1.addRow(["TOTAL", totDebut, consoParAccount.reduce((a, c) => a + c.revenus, 0), consoParAccount.reduce((a, c) => a + c.depenses, 0), totFin - totDebut, totFin, ""]);
      totRowC.eachCell((c: any) => { c.font = { bold: true }; });
      [2, 3, 4, 5, 6].forEach((i) => { totRowC.getCell(i).numFmt = "#,##0"; totRowC.getCell(i).alignment = { horizontal: "right" }; });

      // ===== Feuille 2 : Évolution mensuelle du mouvement net par compte =====
      const ws3 = wb.addWorksheet("Évolution mensuelle");
      titleBanner(ws3, "Mouvement net mensuel par compte");
      const activeAccounts = consoParAccount.filter((c) => c.count > 0);
      const allMonthsUnion = Array.from(new Set(activeAccounts.flatMap((c) => Object.keys(c.byMonth)))).sort((a, b) => monthSortKey(a) - monthSortKey(b));
      const headerRow3 = ws3.addRow(["Mois", ...activeAccounts.map((c) => c.account.name), "Total mois"]);
      styleHeaderRow(headerRow3);
      ws3.columns = [{ width: 14 }, ...activeAccounts.map(() => ({ width: 18 })), { width: 18 }];
      allMonthsUnion.forEach((m) => {
        const vals = activeAccounts.map((c) => c.byMonth[m] || 0);
        const row = ws3.addRow([monthLabel(m), ...vals, vals.reduce((a, v) => a + v, 0)]);
        for (let i = 2; i <= activeAccounts.length + 2; i++) {
          const cell = row.getCell(i);
          cell.numFmt = "#,##0"; cell.alignment = { horizontal: "right" };
          if (typeof cell.value === "number") cell.font = { color: { argb: cell.value >= 0 ? EMERALD : CLAY }, bold: i === activeAccounts.length + 2 };
        }
      });
      const colTotalsC = activeAccounts.map((c) => allMonthsUnion.reduce((a, m) => a + (c.byMonth[m] || 0), 0));
      const grandTotalC = colTotalsC.reduce((a, v) => a + v, 0);
      const totalRow3c = ws3.addRow(["TOTAL", ...colTotalsC, grandTotalC]);
      totalRow3c.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      for (let i = 2; i <= activeAccounts.length + 2; i++) { totalRow3c.getCell(i).numFmt = "#,##0"; totalRow3c.getCell(i).alignment = { horizontal: "right" }; }

      // ===== Feuille 4 : Détail transactions (période filtrée) =====
      const ws4 = wb.addWorksheet("Détail transactions");
      titleBanner(ws4, "Transactions de la période");
      const headerRow4 = ws4.addRow(["Compte", "Date", "Catégorie", "Sous-catégorie", "Type", "Montant (FCFA)"]);
      styleHeaderRow(headerRow4);
      ws4.columns = [{ width: 18 }, { width: 12 }, { width: 22 }, { width: 20 }, { width: 10 }, { width: 16 }];
      const sortedTx4 = periodTx.slice().sort((a, b) => (a.account || "").localeCompare(b.account || "") || a.date.localeCompare(b.date));
      let curAccount: string | null = null, accRev = 0, accDep = 0;
      const flushAccountSubtotal = () => {
        if (curAccount === null) return;
        const row = ws4.addRow([`Sous-total ${curAccount}`, "", "", "", "", accRev - accDep]);
        row.eachCell((c: any) => { c.font = { bold: true, color: { argb: GOLD } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF232F27" } }; });
        row.getCell(6).numFmt = "#,##0"; row.getCell(6).alignment = { horizontal: "right" };
      };
      sortedTx4.forEach((t) => {
        const acc = t.account || "—";
        if (curAccount !== null && acc !== curAccount) { flushAccountSubtotal(); accRev = 0; accDep = 0; }
        curAccount = acc;
        if (t.type === "Revenu") accRev += t.amount; else accDep += t.amount;
        const row = ws4.addRow([acc, dateLabelFull(t.date), t.category, t.subcategory || "", t.type, t.amount]);
        row.getCell(5).font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY } };
        row.getCell(6).font = { color: { argb: t.type === "Revenu" ? EMERALD : CLAY } };
        row.getCell(6).numFmt = "#,##0"; row.getCell(6).alignment = { horizontal: "right" };
      });
      flushAccountSubtotal();
      const grandTotalRow4 = ws4.addRow(["TOTAL GÉNÉRAL", "", "", "", "", sortedTx4.reduce((a, t) => a + (t.type === "Revenu" ? t.amount : -t.amount), 0)]);
      grandTotalRow4.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
      grandTotalRow4.getCell(6).numFmt = "#,##0"; grandTotalRow4.getCell(6).alignment = { horizontal: "right" };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `grand-livre_consommation-comptes_${todayISO()}.xlsx`; a.click();
      URL.revokeObjectURL(url);
      setXlsState("idle");
    } catch (e) {
      console.error(e);
      setXlsState("error");
    }
  };

  const [pdfState, setPdfState] = useState<"idle" | "loading" | "error">("idle");
  const exportComptesNarrativePdf = async () => {
    setPdfState("loading");
    try {
      const [jsPDFModule, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      // jspdf@2.5.x expose le vrai constructeur sur l'export NOMMÉ "jsPDF", pas sur
      // "default" (qui résout vers un objet inutilisable selon le mode d'interop
      // CJS/ESM) — cause du "Réessayer" systématique sur tous les boutons PDF de l'app.
      const jsPDF: any = (jsPDFModule as any).jsPDF || (jsPDFModule as any).default;
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const periodLabel = `${monthLabel(periodFrom)} — ${monthLabel(periodTo)}`;
      const ps = (s: any): string => String(s).replace(/[\u202F\u00A0]/g, " ").replace(/→/g, "->").replace(/—/g, "-").replace(/…/g, "...");

      doc.setFillColor(26, 43, 76);
      doc.rect(0, 0, pageWidth, 34, "F");
      doc.setFontSize(17); doc.setTextColor(255, 255, 255);
      doc.text("Grand Livre — Analyse narrative par compte", 14, 16);
      doc.setFontSize(9); doc.setTextColor(200, 210, 225);
      doc.text(ps(`Période : ${periodLabel} · Généré le ${dateLabelFull(todayISO())}`), 14, 24);

      let y = 44;
      const active = consoParAccount.filter((c) => c.count > 0 || c.soldeDebut !== c.soldeFin);
      active.forEach((c) => {
        const evolPct = c.soldeDebut !== 0 ? ((c.soldeFin - c.soldeDebut) / Math.abs(c.soldeDebut)) * 100 : null;
        const topDep = Object.entries(c.depByCat).sort((a, b) => b[1] - a[1])[0];
        const topRev = Object.entries(c.revByCat).sort((a, b) => b[1] - a[1])[0];
        const months = Object.keys(c.byMonth).sort((a, b) => monthSortKey(a) - monthSortKey(b));

        if (y > pageHeight - 60) { doc.addPage(); y = 20; }

        doc.setFillColor(201, 162, 39);
        doc.rect(14, y, 3, 12, "F");
        doc.setFontSize(13); doc.setTextColor(26, 43, 76); doc.setFont("helvetica", "bold");
        doc.text(ps(c.account.name), 20, y + 8.5);
        doc.setFont("helvetica", "normal");
        y += 18;

        doc.setFillColor(245, 247, 245);
        doc.roundedRect(14, y, pageWidth - 28, 16, 2, 2, "F");
        doc.setFontSize(8); doc.setTextColor(90, 100, 95);
        doc.text("SOLDE", 20, y + 6);
        doc.setFontSize(12); doc.setTextColor(26, 26, 26); doc.setFont("helvetica", "bold");
        doc.text(ps(`${fmt(c.soldeDebut)} → ${fmt(c.soldeFin)} FCFA${evolPct !== null ? `  (${evolPct >= 0 ? "+" : ""}${evolPct.toFixed(1)}%)` : ""}`), 20, y + 12.5);
        doc.setFont("helvetica", "normal");
        y += 22;

        const paragraphs: string[] = [
          `Ce compte a reçu ${fmt(c.revenus)} FCFA et vu sortir ${fmt(c.depenses)} FCFA sur la période, pour un mouvement net de ${c.net >= 0 ? "+" : ""}${fmt(c.net)} FCFA, sur ${c.count} transaction(s).`,
        ];
        if (topDep) paragraphs.push(`Principal poste de dépense : "${topDep[0]}" (${fmt(topDep[1])} FCFA, ${((topDep[1] / (c.depenses || 1)) * 100).toFixed(0)}% de la consommation de ce compte).`);
        if (topRev) paragraphs.push(`Principale source de revenu : "${topRev[0]}" (${fmt(topRev[1])} FCFA, ${((topRev[1] / (c.revenus || 1)) * 100).toFixed(0)}% des entrées de ce compte).`);
        if (months.length >= 2) {
          const values = months.map((m) => c.byMonth[m]);
          const bestIdx = values.indexOf(Math.max(...values)), worstIdx = values.indexOf(Math.min(...values));
          paragraphs.push(`Meilleur mois : ${monthLabel(months[bestIdx])} (${values[bestIdx] >= 0 ? "+" : ""}${fmt(values[bestIdx])} FCFA de mouvement net). Pire mois : ${monthLabel(months[worstIdx])} (${fmt(values[worstIdx])} FCFA).`);
        }

        doc.setFontSize(9.5); doc.setTextColor(40, 40, 40);
        paragraphs.forEach((p) => {
          const lines = doc.splitTextToSize(ps(p), pageWidth - 34);
          if (y + lines.length * 5 > pageHeight - 16) { doc.addPage(); y = 20; }
          doc.text(lines, 20, y);
          y += lines.length * 5 + 4;
        });

        if (months.length >= 2) {
          if (y > pageHeight - 40) { doc.addPage(); y = 20; }
          doc.autoTable({
            startY: y,
            head: [["Mois", "Mouvement net (FCFA)"]],
            body: months.map((m) => [monthLabel(m), ps(fmt(c.byMonth[m]))]),
            headStyles: { fillColor: [26, 43, 76], fontSize: 8 },
            styles: { fontSize: 8 },
            columnStyles: { 1: { halign: "right" } },
            margin: { left: 20, right: 14 },
          });
          y = (doc as any).lastAutoTable.finalY + 12;
        } else {
          y += 8;
        }
      });

      doc.save(`grand-livre_narratif-comptes_${todayISO()}.pdf`);
      setPdfState("idle");
    } catch (e) {
      console.error(e);
      setPdfState("error");
    }
  };

  // Avances entre comptes : regroupe toutes les transactions marquées "onBehalfOf" par
  // paire débiteur→créancier, net des règlements déjà marqués. Le sens de la dette
  // s'inverse selon le type — pour une DÉPENSE, le compte qui a payé est le créancier (on
  // lui doit) ; pour un REVENU, le compte qui a encaissé est le débiteur (il doit reverser
  // l'argent au bon compte).
  //
  // Compensation automatique : si A doit B ET B doit A en même temps (deux dettes en sens
  // opposé entre les deux mêmes comptes), ça n'a pas de sens d'afficher deux lignes
  // contradictoires — seul le SOLDE NET doit apparaître, dans un seul sens. Sur demande
  // explicite de l'utilisateur (10/08/2026), après avoir vu ce cas de figure en pratique.
  const advances = useMemo(() => {
    const groups: Record<string, { debtor: string; creditor: string; total: number; settled: number; tx: Transaction[] }> = {};
    transactions.forEach((t) => {
      if (!t.onBehalfOf || !t.account || t.onBehalfOf === t.account) return;
      const debtor = t.type === "Dépense" ? t.onBehalfOf : t.account;
      const creditor = t.type === "Dépense" ? t.account : t.onBehalfOf;
      const key = `${debtor}→${creditor}`;
      groups[key] = groups[key] || { debtor, creditor, total: 0, settled: 0, tx: [] };
      groups[key].total += t.amount;
      if (t.settled) groups[key].settled += t.amount;
      groups[key].tx.push(t);
    });
    const pairKeys = new Set<string>();
    Object.values(groups).forEach((g) => pairKeys.add([g.debtor, g.creditor].sort().join("|")));
    const netted: { debtor: string; creditor: string; outstanding: number; grossFwd: number; grossBwd: number; tx: Transaction[] }[] = [];
    pairKeys.forEach((pk) => {
      const [x, y] = pk.split("|");
      const fwd = groups[`${x}→${y}`];
      const bwd = groups[`${y}→${x}`];
      const fwdOut = fwd ? fwd.total - fwd.settled : 0;
      const bwdOut = bwd ? bwd.total - bwd.settled : 0;
      const net = fwdOut - bwdOut;
      const allTx = [...(fwd?.tx || []), ...(bwd?.tx || [])];
      if (Math.abs(net) < 1 && fwdOut === 0 && bwdOut === 0) return; // rien à afficher
      netted.push({
        debtor: net >= 0 ? x : y, creditor: net >= 0 ? y : x, outstanding: Math.abs(net),
        grossFwd: fwd?.total || 0, grossBwd: bwd?.total || 0, tx: allTx,
      });
    });
    return netted.sort((a, b) => b.outstanding - a.outstanding);
  }, [transactions]);
  const [expandedAdvance, setExpandedAdvance] = useState<string | null>(null);
  const toggleSettled = (txId: string) => setTransactions(transactions.map((t) => t.id === txId ? { ...t, settled: !t.settled } : t));

  // Solde "corrigé" par compte : ce que serait le solde de chaque compte si toutes les
  // avances en cours étaient réglées aujourd'hui — le payeur récupère ce qu'on lui doit,
  // le bénéficiaire paie ce qu'il doit. Permet de voir la vraie situation de fond de
  // chaque compte, au-delà du solde brut faussé par des paiements faits pour un autre.
  const correctedBalances = useMemo(() => {
    const owedTo: Record<string, number> = {}; // ce que ce compte doit récupérer (il a avancé pour d'autres)
    const owedBy: Record<string, number> = {}; // ce que ce compte doit payer (d'autres ont avancé pour lui)
    advances.forEach((a) => {
      owedTo[a.creditor] = (owedTo[a.creditor] || 0) + a.outstanding;
      owedBy[a.debtor] = (owedBy[a.debtor] || 0) + a.outstanding;
    });
    return accounts.map((a) => {
      const real = accountBalance(a, transactions);
      const receivable = owedTo[a.name] || 0;
      const payable = owedBy[a.name] || 0;
      return { account: a, real, corrected: real + receivable - payable, receivable, payable };
    }).filter((r) => r.receivable > 0 || r.payable > 0);
  }, [accounts, transactions, advances]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Kpi label="Total des comptes (temps réel)" value={fmt(total)} tone={COLOR.goldSoft} icon={Wallet} />

      {correctedBalances.length > 0 && (
        <Panel title="Soldes corrigés — si les avances étaient réglées" subtitle="Ce que serait le solde de chaque compte si les avances en cours (non encore marquées réglées) étaient soldées aujourd'hui">
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {correctedBalances.map((r) => (
              <div key={r.account.id} style={{ padding: "10px 14px", background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}` }}>
                <div style={{ fontSize: 13, marginBottom: 6 }}>{r.account.name}</div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, flexWrap: "wrap" }}>
                  <span style={{ color: COLOR.inkMuted }}>{fmt(r.real)}</span>
                  <ArrowRight size={12} color={COLOR.inkMuted} style={{ flexShrink: 0 }} />
                  <span style={{ fontWeight: 600, color: r.corrected >= r.real ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(r.corrected)} FCFA</span>
                </div>
                {(r.receivable > 0 || r.payable > 0) && (
                  <div style={{ display: "flex", gap: 10, marginTop: 4, flexWrap: "wrap" }}>
                    {r.receivable > 0 && <span style={{ fontSize: 10.5, color: COLOR.emeraldSoft }}>+{fmt(r.receivable)} à recevoir</span>}
                    {r.payable > 0 && <span style={{ fontSize: 10.5, color: COLOR.claySoft }}>−{fmt(r.payable)} à payer</span>}
                  </div>
                )}
              </div>
            ))}
          </div>
        </Panel>
      )}

      {advances.length > 0 && (
        <Panel title="Avances entre comptes" subtitle="Dépenses payées ou revenus encaissés sur le mauvais compte — le compte à gauche doit au compte à droite">
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {advances.map((adv) => {
              const key = `${adv.debtor}→${adv.creditor}`;
              const isExpanded = expandedAdvance === key;
              return (
                <div key={key} style={{ background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${adv.outstanding > 0 ? COLOR.gold : COLOR.hairline}`, overflow: "hidden" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px" }}>
                    <div>
                      <div style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{ color: COLOR.claySoft }}>{adv.debtor}</span>
                        <ArrowRight size={12} color={COLOR.inkMuted} />
                        <span style={{ color: COLOR.emeraldSoft }}>{adv.creditor}</span>
                      </div>
                      <button onClick={() => setExpandedAdvance(isExpanded ? null : key)} style={{ background: "transparent", border: "none", color: COLOR.slateBlueSoft, cursor: "pointer", fontSize: 11, display: "flex", alignItems: "center", gap: 2, padding: 0, marginTop: 2 }}>
                        {adv.tx.length} transaction(s) {isExpanded ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                      </button>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: adv.outstanding > 0 ? COLOR.goldSoft : COLOR.emeraldSoft }}>
                        {adv.outstanding > 0 ? `${fmt(adv.outstanding)} dû` : "Réglé"}
                      </div>
                      {adv.grossFwd > 0 && adv.grossBwd > 0 && (
                        <div style={{ fontSize: 10.5, color: COLOR.inkMuted }} title="Les deux comptes se devaient mutuellement — compensé automatiquement, seul le solde net reste dû">
                          brut {fmt(adv.grossFwd)} / {fmt(adv.grossBwd)} — compensé
                        </div>
                      )}
                    </div>
                  </div>
                  {isExpanded && (
                    <div style={{ padding: "8px 14px 12px", borderTop: `1px solid ${COLOR.hairline}`, display: "flex", flexDirection: "column", gap: 4 }}>
                      {adv.tx.slice().sort((a, b) => b.date.localeCompare(a.date)).map((t) => (
                        <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 11.5, padding: "4px 0" }}>
                          <span style={{ color: COLOR.inkMuted }}>{dateLabelFull(t.date)} — {t.category}{t.subcategory ? ` · ${t.subcategory}` : ""}</span>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: t.settled ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(t.amount)}</span>
                            <button onClick={() => toggleSettled(t.id)} style={{ fontSize: 10.5, padding: "3px 8px", borderRadius: 10, border: `1px solid ${t.settled ? COLOR.emerald : COLOR.hairline}`, background: "transparent", color: t.settled ? COLOR.emeraldSoft : COLOR.inkMuted, cursor: "pointer" }}>
                              {t.settled ? "Réglé ✓" : "Marquer réglé"}
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </Panel>
      )}

      <PanelWithHelp title="Consommation par compte" subtitle={`${monthLabel(periodFrom)} — ${monthLabel(periodTo)} · les avances entre comptes sont comptées ici sous le compte bénéficiaire, pas le compte payeur`}
        explain="Le solde ci-dessus est un instantané à aujourd'hui. Ce panneau, lui, montre ce qui a réellement transité sur chaque compte pendant une période choisie : combien en a été dépensé (consommation), combien y est entré (revenus), et le mouvement net. Utile pour voir, par exemple, à quel point Petty Cash ou SALAIRE ont été sollicités sur les 3 derniers mois."
        right={
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
            <button onClick={exportComptesExcel} disabled={xlsState === "loading"} style={{
              display: "flex", alignItems: "center", gap: 6, background: xlsState === "error" ? "rgba(193,84,63,0.14)" : "rgba(63,156,122,0.14)",
              border: `1px solid ${xlsState === "error" ? COLOR.clay : COLOR.emerald}`, borderRadius: 6,
              color: xlsState === "error" ? COLOR.claySoft : COLOR.emeraldSoft, padding: "6px 12px", fontSize: 11.5, cursor: xlsState === "loading" ? "default" : "pointer",
            }}>
              {xlsState === "loading" ? <Loader2 size={12} style={{ animation: "spin 1s linear infinite" }} /> : <FileSpreadsheet size={12} />}
              {xlsState === "loading" ? "Génération…" : xlsState === "error" ? "Réessayer" : "Rapport Excel"}
            </button>
            <button onClick={exportComptesNarrativePdf} disabled={pdfState === "loading"} style={{
              display: "flex", alignItems: "center", gap: 6, background: pdfState === "error" ? "rgba(193,84,63,0.14)" : "rgba(201,162,39,0.14)",
              border: `1px solid ${pdfState === "error" ? COLOR.clay : COLOR.gold}`, borderRadius: 6,
              color: pdfState === "error" ? COLOR.claySoft : COLOR.goldSoft, padding: "6px 12px", fontSize: 11.5, cursor: pdfState === "loading" ? "default" : "pointer",
            }}>
              {pdfState === "loading" ? <Loader2 size={12} style={{ animation: "spin 1s linear infinite" }} /> : <BookOpen size={12} />}
              {pdfState === "loading" ? "Génération…" : pdfState === "error" ? "Réessayer" : "Rapport narratif PDF"}
            </button>
            {[["1m", "1M"], ["3m", "3M"], ["6m", "6M"], ["1a", "1A"], ["tout", "Tout"]].map(([key, label]) => (
              <button key={key} onClick={() => applyPreset(key)} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, color: COLOR.inkMuted, borderRadius: 6, padding: "6px 10px", fontSize: 11.5, cursor: "pointer" }}>{label}</button>
            ))}
            <Select label="Du" value={periodFrom} onChange={setPeriodFromSafe} options={allMonths.filter((m) => monthSortKey(m) <= monthSortKey(periodTo)).map((m) => ({ value: m, label: monthLabel(m) }))} />
            <Select label="Au" value={periodTo} onChange={setPeriodToSafe} options={allMonths.filter((m) => monthSortKey(m) >= monthSortKey(periodFrom)).map((m) => ({ value: m, label: monthLabel(m) }))} />
          </div>
        }>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {consoParAccount.map((c) => (
            <div key={c.account.id} style={{ padding: "12px 14px", background: COLOR.surfaceRaised, borderRadius: 8, border: `1px solid ${COLOR.hairline}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontSize: 13 }}>{c.account.name}<CalcDetailIcon onClick={() => setAccountNarrativeId(c.account.id)} /> <span style={{ color: COLOR.inkMuted, fontSize: 11 }}>· {c.count} transaction(s)</span></div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 14, fontWeight: 600, color: c.net >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>
                  {c.net >= 0 ? "+" : ""}{fmt(c.net)} FCFA
                </div>
              </div>
              <div style={{ display: "flex", gap: 20, fontSize: 11.5, flexWrap: "wrap" }}>
                <span style={{ color: COLOR.claySoft }}>Consommé : <strong style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmt(c.depenses)}</strong></span>
                <span style={{ color: COLOR.emeraldSoft }}>Reçu : <strong style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmt(c.revenus)}</strong></span>
                {totalConsoDep > 0 && c.depenses > 0 && (
                  <span style={{ color: COLOR.inkMuted }}>{((c.depenses / totalConsoDep) * 100).toFixed(0)}% de la conso totale</span>
                )}
                {c.receivedAsAdvance > 0 && (
                  <span style={{ color: COLOR.goldSoft, display: "flex", alignItems: "center", gap: 3 }} title="Dépenses réellement payées depuis un autre compte, réattribuées ici">
                    <ArrowRight size={10} /> dont {fmt(c.receivedAsAdvance)} avancés par un autre compte
                  </span>
                )}
              </div>
              {(c.depenses > 0 || c.revenus > 0) && (
                <div style={{ display: "flex", height: 6, borderRadius: 3, overflow: "hidden", marginTop: 8, background: COLOR.hairline }}>
                  {c.depenses > 0 && <div style={{ width: `${(c.depenses / (c.depenses + c.revenus)) * 100}%`, background: COLOR.clay }} />}
                  {c.revenus > 0 && <div style={{ width: `${(c.revenus / (c.depenses + c.revenus)) * 100}%`, background: COLOR.emerald }} />}
                </div>
              )}
            </div>
          ))}
          {!consoParAccount.length && <EmptyState text="Aucun compte." />}
        </div>
      </PanelWithHelp>
      {accountNarrativeId && (() => {
        const c = consoParAccount.find((x) => x.account.id === accountNarrativeId);
        if (!c) return null;
        const depRows = Object.entries(c.depByCat).sort((a, b) => b[1] - a[1]);
        const revRows = Object.entries(c.revByCat).sort((a, b) => b[1] - a[1]);
        const months = Object.keys(c.byMonth).sort((a, b) => monthSortKey(a) - monthSortKey(b));
        const blocks: CalcDetailBlock[] = [
          { kind: "kv", rows: [
            { label: `Période`, value: `${monthLabel(periodFrom)} — ${monthLabel(periodTo)}` },
            { label: "Solde en début de période", value: `${fmt(c.soldeDebut)} FCFA` },
            { label: "Solde en fin de période", value: `${fmt(c.soldeFin)} FCFA`, strong: true },
            { label: "Reçu (revenus)", value: `${fmt(c.revenus)} FCFA` },
            { label: "Consommé (dépenses)", value: `${fmt(c.depenses)} FCFA` },
            { label: "Mouvement net", value: `${c.net >= 0 ? "+" : ""}${fmt(c.net)} FCFA`, warn: c.net < 0 },
          ] },
        ];
        if (depRows.length) blocks.push({ kind: "table", columns: ["Catégorie (dépenses)", "Montant (FCFA)"], rows: depRows.map(([cat, v]) => [cat, fmt(v)]) });
        if (revRows.length) blocks.push({ kind: "table", columns: ["Catégorie (revenus)", "Montant (FCFA)"], rows: revRows.map(([cat, v]) => [cat, fmt(v)]) });
        if (months.length >= 2) blocks.push({ kind: "table", columns: ["Mois", "Mouvement net (FCFA)"], rows: months.map((m) => [monthLabel(m), fmt(c.byMonth[m])]) });
        return <CalcDetailSheet open={!!accountNarrativeId} onClose={() => setAccountNarrativeId(null)}
          title={c.account.name} headline={`${fmt(c.soldeDebut)} → ${fmt(c.soldeFin)} FCFA`}
          formula="Solde réel du compte en début/fin de période, puis revenus et dépenses par catégorie" blocks={blocks} />;
      })()}

      <Panel title="Comptes" subtitle="Le solde de chaque compte se met à jour automatiquement dès qu'une transaction lui est liée"
        right={
          <div style={{ display: "flex", gap: 8 }}>
            {accounts.length >= 2 && (
              <button onClick={() => { setMerging((m) => !m); setAdding(false); }} style={{ display: "flex", alignItems: "center", gap: 6, background: merging ? COLOR.hairline : "rgba(110,127,168,0.14)", border: `1px solid ${merging ? COLOR.hairline : COLOR.slateBlue}`, borderRadius: 6, color: merging ? COLOR.inkMuted : COLOR.slateBlueSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
                <GitCompare size={13} /> {merging ? "Annuler" : "Fusionner deux comptes"}
              </button>
            )}
            <button onClick={() => { setAdding((a) => !a); setMerging(false); }} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
              {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Ajouter un compte"}
            </button>
          </div>
        }>
        {merging && (
          <div style={{ padding: 16, background: COLOR.surfaceRaised, borderRadius: 8, marginBottom: 16, border: `1px solid ${COLOR.hairline}` }}>
            <div style={{ fontSize: 12, color: COLOR.inkMuted, marginBottom: 10 }}>Le compte "à absorber" disparaît ; toutes ses transactions et son solde de départ rejoignent le compte "à garder", qui conserve son nom.</div>
            <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte à absorber (disparaît)</label>
                <select style={{ ...inputStyle, width: 200 }} value={mergeSourceId} onChange={(e) => setMergeSourceId(e.target.value)}>
                  <option value="">Choisir…</option>
                  {accounts.map((a) => <option key={a.id} value={a.id} disabled={a.id === mergeTargetId}>{a.name}</option>)}
                </select>
              </div>
              <ArrowRight size={16} color={COLOR.inkMuted} style={{ marginBottom: 8 }} />
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte à garder (nom conservé)</label>
                <select style={{ ...inputStyle, width: 200 }} value={mergeTargetId} onChange={(e) => setMergeTargetId(e.target.value)}>
                  <option value="">Choisir…</option>
                  {accounts.map((a) => <option key={a.id} value={a.id} disabled={a.id === mergeSourceId}>{a.name}</option>)}
                </select>
              </div>
              <button onClick={() => setConfirmMerge(true)} disabled={!mergeSourceId || !mergeTargetId} style={{ background: mergeSourceId && mergeTargetId ? COLOR.slateBlue : COLOR.hairline, border: "none", borderRadius: 6, color: mergeSourceId && mergeTargetId ? COLOR.bg : COLOR.inkMuted, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: mergeSourceId && mergeTargetId ? "pointer" : "default", height: 32 }}>Fusionner</button>
            </div>
            {mergeSourceId && mergeTargetId && (() => {
              const source = accounts.find((a) => a.id === mergeSourceId)!;
              const target = accounts.find((a) => a.id === mergeTargetId)!;
              const linkedCount = transactions.filter((t) => t.account === source.name || t.onBehalfOf === source.name).length;
              return (
                <div style={{ marginTop: 10, fontSize: 11.5, color: COLOR.goldSoft }}>
                  {linkedCount} transaction(s) de "{source.name}" rejoindront "{target.name}" · solde de départ combiné : {fmt(target.openingBalance + source.openingBalance)} FCFA
                </div>
              );
            })()}
          </div>
        )}
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
      <ConfirmDialog
        open={confirmMerge}
        title="Fusionner ces deux comptes ?"
        message={mergeSourceId && mergeTargetId ? `"${accounts.find((a) => a.id === mergeSourceId)?.name}" disparaîtra définitivement. Toutes ses transactions et son solde de départ rejoindront "${accounts.find((a) => a.id === mergeTargetId)?.name}". Cette action est irréversible.` : ""}
        confirmLabel="Fusionner"
        onConfirm={mergeAccounts}
        onCancel={() => setConfirmMerge(false)}
      />
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
  const [budgetNarrativeOpen, setBudgetNarrativeOpen] = useState(false);
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
          <div style={{ display: "flex", gap: 8 }}>
            {budgets.length > 0 && (
              <button onClick={() => setBudgetNarrativeOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 6, color: COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
                <BookOpen size={13} /> Rapport détaillé
              </button>
            )}
            <button onClick={() => setAdding((a) => !a)} style={{ display: "flex", alignItems: "center", gap: 6, background: adding ? COLOR.hairline : "rgba(201,162,39,0.14)", border: `1px solid ${adding ? COLOR.hairline : COLOR.gold}`, borderRadius: 6, color: adding ? COLOR.inkMuted : COLOR.goldSoft, padding: "8px 14px", fontSize: 12.5, cursor: "pointer" }}>
              {adding ? <X size={13} /> : <Plus size={13} />} {adding ? "Annuler" : "Nouveau budget"}
            </button>
          </div>
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
      {budgetNarrativeOpen && (() => {
        const monthsBack: string[] = [];
        { let mk = currentMonth; for (let i = 0; i < 6; i++) { monthsBack.push(mk); mk = prevMonthKey(mk); } }
        const rows = budgets.map((b) => {
          const overMonths = monthsBack.filter((m) => spentInMonth(b.category, m) > b.amount);
          const avgSpent = mean(monthsBack.map((m) => spentInMonth(b.category, m)));
          return { cat: b.category, limit: b.amount, avgSpent, overCount: overMonths.length, overRate: (overMonths.length / monthsBack.length) * 100 };
        }).sort((a, b) => b.overRate - a.overRate);
        const blocks: CalcDetailBlock[] = [
          { kind: "table", columns: ["Catégorie budgétée", "Limite (FCFA)", "Dépense moy./mois", "Mois dépassés (sur 6)"], rows: rows.map((r) => [r.cat, fmt(r.limit), fmt(r.avgSpent), `${r.overCount}/6`]) },
        ];
        const worst = rows[0];
        if (worst && worst.overCount > 0) blocks.push({ kind: "note", tone: "warn", text: `"${worst.cat}" est le budget le plus souvent dépassé (${worst.overCount} mois sur 6) — dépense moyenne ${fmt(worst.avgSpent)} FCFA/mois pour une limite de ${fmt(worst.limit)} FCFA.` });
        else blocks.push({ kind: "note", tone: "info", text: "Aucun budget dépassé sur les 6 derniers mois." });
        return <CalcDetailSheet open={budgetNarrativeOpen} onClose={() => setBudgetNarrativeOpen(false)}
          title="Budgets — analyse détaillée" headline={`${rows.filter((r) => r.overCount > 0).length} budget(s) dépassé(s) au moins une fois sur 6 mois`}
          formula="Comparaison dépense réelle vs limite, mois par mois, sur les 6 derniers mois" blocks={blocks} />;
      })()}
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
function RecurrencesTab({ recurring, setRecurring, transactions, setTransactions, allCategories, accounts, chargeOverrides, includeGrundfosVoiture }: {
  recurring: RecurringTemplate[]; setRecurring: (r: RecurringTemplate[]) => void;
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; allCategories: string[]; accounts: Account[];
  chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean;
}) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<RecurringTemplate, "id">>({ category: categoriesForType(transactions, "Dépense")[0] || "", type: "Dépense", amount: 0, frequency: "Mensuelle", nextDate: todayISO(), account: accounts[0]?.name });

  // Suggestions tirées de Charges Fixes & Variables : tout poste classé "Fixe" (auto ou
  // manuel) n'ayant pas encore de modèle récurrent correspondant — pont entre les deux
  // systèmes, jusqu'ici indépendants l'un de l'autre.
  const fixedCharges = useMemo(() => {
    const windowMonths = monthsSinceInception(transactions);
    const result = classifyCharges(transactions, chargeOverrides, includeGrundfosVoiture, windowMonths);
    return result.rows.filter((r) => r.mode === "fixe" && r.amount > 0);
  }, [transactions, chargeOverrides, includeGrundfosVoiture]);
  const existingRecurringCategories = useMemo(() => new Set(recurring.filter((r) => r.type === "Dépense").map((r) => r.category)), [recurring]);
  const suggestions = fixedCharges.filter((r) => !existingRecurringCategories.has(r.poste.split("::")[0]));

  const addFromSuggestion = (r: { poste: string; amount: number }) => {
    setRecurring([...recurring, { category: r.poste.split("::")[0], type: "Dépense", amount: Math.round(r.amount), frequency: "Mensuelle", nextDate: todayISO(), account: accounts[0]?.name, id: uid("r") }]);
  };

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

      {suggestions.length > 0 && (
        <Panel title="Suggestions depuis Charges Fixes & Variables" subtitle="Postes déjà classés « Fixe » sans modèle récurrent correspondant — ajoute-les en un clic si tu veux un rappel">
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {suggestions.map((r) => (
              <div key={r.poste} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: "rgba(201,162,39,0.06)", border: `1px solid ${COLOR.hairline}`, borderRadius: 6 }}>
                <span style={{ fontSize: 12.5 }}>{r.poste.replace("::", " · ")}</span>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: COLOR.claySoft }}>{fmt(r.amount)}</span>
                  <button onClick={() => addFromSuggestion(r)} style={{ display: "flex", alignItems: "center", gap: 5, background: "rgba(201,162,39,0.14)", border: `1px solid ${COLOR.gold}`, borderRadius: 6, color: COLOR.goldSoft, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}><Plus size={12} /> Ajouter</button>
                </div>
              </div>
            ))}
          </div>
        </Panel>
      )}

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
// Écran bloquant affiché quand le stockage local semble vide — force un choix explicite
// avant que quoi que ce soit ne soit sauvegardé, pour ne plus jamais transformer
// silencieusement des données de secours en "vraies" données. Sur demande explicite de
// l'utilisateur (11/08/2026), après une perte de données réelle causée par ce défaut.
function DataRecoveryGate({ onRestore, onConnectSync, onStartFresh }: {
  onRestore: (data: any) => void; onConnectSync: (code: string) => Promise<boolean>; onStartFresh: () => void;
}) {
  const [mode, setMode] = useState<"choix" | "sync" | "confirmFresh">("choix");
  const [codeInput, setCodeInput] = useState("");
  const [syncTrying, setSyncTrying] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [fileError, setFileError] = useState<string | null>(null);

  const onFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result as string);
        if (!data || typeof data !== "object") throw new Error("format");
        setFileError(null);
        onRestore(data);
      } catch {
        setFileError("Fichier invalide — vérifie que c'est bien un export Grand Livre (.json).");
      }
    };
    reader.readAsText(file);
  };

  const tryConnect = async () => {
    if (!codeInput.trim()) return;
    setSyncTrying(true); setSyncError(null);
    const ok = await onConnectSync(codeInput.trim());
    setSyncTrying(false);
    if (!ok) setSyncError("Aucune donnée trouvée pour ce code — vérifie qu'il est correct, ou essaie une autre option.");
  };

  const cardStyle: React.CSSProperties = { background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 12, padding: 20, cursor: "pointer", textAlign: "left", width: "100%" };

  return (
    <div style={{ minHeight: "100vh", background: COLOR.bg, color: COLOR.ink, fontFamily: "'Inter', sans-serif", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ maxWidth: 460, width: "100%" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
          <AlertTriangle size={22} color={COLOR.claySoft} />
          <h1 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 500, margin: 0 }}>Aucune donnée trouvée ici</h1>
        </div>
        <p style={{ color: COLOR.inkMuted, fontSize: 13, lineHeight: 1.5, marginBottom: 24 }}>
          Ce navigateur/appareil ne contient aucune donnée enregistrée. Si tu avais déjà utilisé l'app, tes données existent probablement encore ailleurs — choisis comment les retrouver avant de continuer. Rien ne sera sauvegardé tant que tu n'as pas fait un choix ici.
        </p>

        {mode === "choix" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <label style={cardStyle}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                <UploadCloud size={16} color={COLOR.goldSoft} />
                <strong style={{ fontSize: 14 }}>Importer une sauvegarde</strong>
              </div>
              <div style={{ fontSize: 12, color: COLOR.inkMuted }}>Un fichier .json déjà exporté depuis Sauvegarde, sur cet appareil ou un autre.</div>
              <input type="file" accept=".json" onChange={onFile} style={{ display: "none" }} />
            </label>
            {fileError && <div style={{ fontSize: 12, color: COLOR.claySoft }}>{fileError}</div>}

            <button onClick={() => setMode("sync")} style={cardStyle}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                <Repeat size={16} color={COLOR.slateBlueSoft} />
                <strong style={{ fontSize: 14 }}>Se connecter à une synchronisation existante</strong>
              </div>
              <div style={{ fontSize: 12, color: COLOR.inkMuted }}>Si tu avais déjà activé la synchronisation avant, entre ton code pour retrouver tes données.</div>
            </button>

            <button onClick={() => setMode("confirmFresh")} style={{ ...cardStyle, borderColor: COLOR.hairline, opacity: 0.85 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                <Plus size={16} color={COLOR.inkMuted} />
                <strong style={{ fontSize: 14, color: COLOR.inkMuted }}>Démarrer avec des données de démonstration</strong>
              </div>
              <div style={{ fontSize: 12, color: COLOR.inkMuted }}>Seulement si c'est vraiment la première fois que tu utilises l'app sur cet appareil.</div>
            </button>
          </div>
        )}

        {mode === "sync" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <input value={codeInput} onChange={(e) => setCodeInput(e.target.value)} placeholder="ex: atelier-lagune-482" style={{ ...inputStyle, width: "100%", padding: "12px 14px", fontSize: 14 }} />
            {syncError && <div style={{ fontSize: 12, color: COLOR.claySoft }}>{syncError}</div>}
            <button onClick={tryConnect} disabled={syncTrying || !codeInput.trim()} style={{ background: COLOR.emerald, border: "none", borderRadius: 8, color: COLOR.bg, padding: "12px 14px", fontSize: 13.5, fontWeight: 600, cursor: syncTrying ? "default" : "pointer" }}>
              {syncTrying ? "Connexion…" : "Se connecter"}
            </button>
            <button onClick={() => setMode("choix")} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, fontSize: 12.5, cursor: "pointer", padding: 6 }}>← Retour</button>
          </div>
        )}

        {mode === "confirmFresh" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ background: "rgba(193,84,63,0.12)", border: `1px solid ${COLOR.clay}`, borderRadius: 8, padding: 14, fontSize: 12.5, color: COLOR.claySoft }}>
              Confirme bien : si tu avais déjà de vraies données ailleurs et que tu choisis cette option par erreur, les données de démonstration commenceront à être sauvegardées ici à la place. Tu pourras toujours importer une sauvegarde plus tard depuis Sauvegarde.
            </div>
            <button onClick={onStartFresh} style={{ background: COLOR.clay, border: "none", borderRadius: 8, color: "#fff", padding: "12px 14px", fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}>
              Oui, démarrer avec les données de démonstration
            </button>
            <button onClick={() => setMode("choix")} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, fontSize: 12.5, cursor: "pointer", padding: 6 }}>← Retour</button>
          </div>
        )}
      </div>
    </div>
  );
}

function SauvegardeTab({ getSnapshot, restore, syncCode, setSyncCode, syncStatus, lastSyncedAt, onForceSync, realtimeConnected, undoSnapshotAt, onUndoRestore, settingsLog }: {
  getSnapshot: () => any; restore: (data: any) => void; syncCode: string; setSyncCode: (c: string) => void;
  syncStatus: "idle" | "syncing" | "synced" | "error" | "disabled"; lastSyncedAt: string | null; onForceSync: () => void; realtimeConnected: boolean;
  undoSnapshotAt: string | null; onUndoRestore: () => void; settingsLog: SettingsLogEntry[];
}) {
  const [confirmUndo, setConfirmUndo] = useState(false);
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
      {undoSnapshotAt && (
        <Panel title="Filet de sécurité" subtitle="Un instantané de tes données a été pris automatiquement juste avant ta dernière restauration">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, flexWrap: "wrap" }}>
            <div style={{ fontSize: 12.5, color: COLOR.inkMuted, lineHeight: 1.6 }}>
              Dernière restauration effectuée le <strong style={{ color: COLOR.ink }}>{undoSnapshotAt}</strong>. Si le résultat ne correspond pas à ce que tu attendais, tu peux revenir en un clic à l'état d'avant cette restauration.
            </div>
            <button onClick={() => setConfirmUndo(true)} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(193,84,63,0.14)", border: `1px solid ${COLOR.clay}`, borderRadius: 8, color: COLOR.claySoft, padding: "10px 18px", fontSize: 13, cursor: "pointer", whiteSpace: "nowrap" }}>
              <RotateCcw size={15} /> Annuler la dernière restauration
            </button>
          </div>
        </Panel>
      )}
      <ConfirmDialog
        open={confirmUndo}
        title="Annuler la dernière restauration ?"
        message="Tes données reviendront exactement à l'état où elles étaient juste avant ta dernière restauration. Ce qui a été fait depuis (nouvelles transactions, réglages) sera perdu."
        onConfirm={() => { onUndoRestore(); setConfirmUndo(false); setStatus("Retour à l'état précédent effectué."); }}
        onCancel={() => setConfirmUndo(false)}
      />
      <Panel title="Historique des ajustements manuels" subtitle="Tous les réglages que tu as modifiés toi-même (charges, activités, portées, objectifs), avec la date">
        <div className="gl-scroll" style={{ maxHeight: 420, overflowY: "auto", display: "flex", flexDirection: "column", gap: 8 }}>
          {settingsLog.map((entry, i) => (
            <div key={i} style={{ display: "flex", gap: 12, padding: "8px 0", borderBottom: `1px solid ${COLOR.hairline}` }}>
              <span style={{ fontSize: 11, color: COLOR.inkMuted, fontFamily: "'IBM Plex Mono', monospace", flexShrink: 0, width: 150 }}>{entry.at}</span>
              <span style={{ fontSize: 12.5, color: COLOR.ink }}>{entry.text}</span>
            </div>
          ))}
          {!settingsLog.length && <EmptyState text="Aucun ajustement manuel enregistré pour l'instant." />}
        </div>
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
  const [bulkOnBehalfOf, setBulkOnBehalfOf] = useState("");
  const pageSize = 25;

  const sorted = useMemo(() => filtered.slice().sort((a, b) => b.date.localeCompare(a.date)), [filtered]);
  const pageCount = Math.max(1, Math.ceil(sorted.length / pageSize));
  // Corrigé le 10/08/2026 : la page ne se réinitialisait jamais quand le filtre changeait
  // — rester sur la page 3 puis restreindre à une catégorie avec 1 seul résultat donnait
  // un tableau vide malgré un compteur correct ("1 transaction filtrée" mais rien affiché).
  // Garde-fou en deux temps : on réinitialise à la page 0 dès que la liste filtrée change,
  // et on protège quand même l'affichage avec un clamp au cas où le rendu arrive avant
  // que l'effet ait eu le temps de s'appliquer.
  useEffect(() => { setPage(0); }, [filtered]);
  const safePage = Math.min(page, pageCount - 1);
  const pageRows = sorted.slice(safePage * pageSize, safePage * pageSize + pageSize);

  // Regroupement par jour pour les bandeaux "rupture" (date + sous-total), à la manière
  // d'un journal comptable classique / grille WinDev — demande explicite de l'utilisateur
  // (13/08/2026). Regroupe seulement les lignes déjà présentes sur la page affichée : un
  // jour dont les écritures sont réparties sur deux pages aura un sous-total par page.
  const pageGroups = useMemo(() => {
    const groups: { date: string; rows: typeof pageRows }[] = [];
    pageRows.forEach((t) => {
      const last = groups[groups.length - 1];
      if (last && last.date === t.date) last.rows.push(t);
      else groups.push({ date: t.date, rows: [t] });
    });
    return groups;
  }, [pageRows]);
  const dayTotals = (rows: typeof pageRows) => {
    const rev = rows.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
    const dep = rows.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
    return { rev, dep, solde: rev - dep };
  };


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
  // Applique l'avance entre comptes à toute la sélection d'un coup — sur demande
  // explicite de l'utilisateur (10/08/2026), pour retagger efficacement les anciennes
  // transactions sans les rouvrir une par une. Ne touche que les Dépenses de la
  // sélection (les Revenus ne sont pas concernés par ce mécanisme).
  const bulkApplyOnBehalfOf = () => {
    if (!bulkOnBehalfOf) return;
    setTransactions(transactions.map((t) => (selected.has(t.id) && t.account !== bulkOnBehalfOf) ? { ...t, onBehalfOf: bulkOnBehalfOf } : t));
    setSelected(new Set()); setBulkOnBehalfOf("");
  };
  const bulkClearOnBehalfOf = () => {
    setTransactions(transactions.map((t) => (selected.has(t.id) ? { ...t, onBehalfOf: undefined } : t)));
    setSelected(new Set());
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
            <GroupedSingleSelect label="" allLabel="Changer la catégorie…" value={bulkCategory} onChange={setBulkCategory} options={groupedCategoryOptions(transactions)} />
            <button onClick={bulkChangeCategory} disabled={!bulkCategory} style={{ background: bulkCategory ? COLOR.emerald : COLOR.hairline, border: "none", borderRadius: 6, color: bulkCategory ? COLOR.bg : COLOR.inkMuted, padding: "6px 12px", fontSize: 11.5, cursor: bulkCategory ? "pointer" : "default" }}>Appliquer</button>
            <select style={{ ...inputStyle, width: 210 }} value={bulkOnBehalfOf} onChange={(e) => setBulkOnBehalfOf(e.target.value)} title="Marque ces dépenses comme réellement destinées à cet autre compte">
              <option value="">Marquer avance pour…</option>
              {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
            </select>
            <button onClick={bulkApplyOnBehalfOf} disabled={!bulkOnBehalfOf} style={{ background: bulkOnBehalfOf ? COLOR.emerald : COLOR.hairline, border: "none", borderRadius: 6, color: bulkOnBehalfOf ? COLOR.bg : COLOR.inkMuted, padding: "6px 12px", fontSize: 11.5, cursor: bulkOnBehalfOf ? "pointer" : "default" }}>Appliquer</button>
            <button onClick={bulkClearOnBehalfOf} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}>Retirer l'avance</button>
            <button onClick={() => setConfirmBulkDelete(true)} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: `1px solid ${COLOR.clay}`, borderRadius: 6, color: COLOR.claySoft, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}><Trash2 size={12} /> Supprimer la sélection</button>
            <button onClick={() => setSelected(new Set())} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, fontSize: 11.5, cursor: "pointer" }}>Désélectionner</button>
          </div>
        )}

        <div className="gl-scroll" style={{ overflowX: "auto", border: `1px solid ${COLOR.hairline}`, borderRadius: 10 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 720 }}>
            <thead><tr style={{ background: "linear-gradient(180deg, #1c2a22, #182119)" }}>
              <th style={{ padding: "10px 10px", borderBottom: `2px solid ${COLOR.gold}`, borderRight: `1px solid ${COLOR.hairline}` }}>
                <button onClick={toggleSelectAllPage} style={{ background: "transparent", border: "none", cursor: "pointer", display: "flex" }}>
                  {allPageSelected ? <CheckSquare size={14} color={COLOR.goldSoft} /> : <Square size={14} color={COLOR.inkMuted} />}
                </button>
              </th>
              {["Heure", "Catégorie", "Type", "Groupe", "Montant", ""].map((h, i) => (
              <th key={h} style={{ textAlign: i === 4 ? "right" : "left", padding: "10px 10px", fontSize: 10.5, color: COLOR.goldSoft, textTransform: "uppercase", letterSpacing: "0.06em", fontWeight: 700, borderBottom: `2px solid ${COLOR.gold}`, borderRight: i < 5 ? `1px solid ${COLOR.hairline}` : "none" }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {pageGroups.map((g) => {
                const dt = dayTotals(g.rows);
                return (
                  <React.Fragment key={g.date}>
                    <tr>
                      <td colSpan={7} style={{ padding: "8px 12px", background: "rgba(201,162,39,0.09)", borderTop: `1px solid ${COLOR.hairline}`, borderBottom: `1px solid ${COLOR.hairline}`, borderLeft: `3px solid ${COLOR.gold}` }}>
                        <div style={{ display: "flex", alignItems: "baseline", gap: 10, flexWrap: "wrap" }}>
                          <span style={{ fontFamily: "'Fraunces', serif", fontSize: 13.5, color: COLOR.goldSoft, fontWeight: 600 }}>{weekdayLabel(g.date)} {dateLabelFull(g.date)}</span>
                          <span style={{ fontSize: 10.5, color: COLOR.inkMuted, background: COLOR.surfaceRaised, border: `1px solid ${COLOR.hairline}`, borderRadius: 20, padding: "1px 9px" }}>{g.rows.length} écriture{g.rows.length > 1 ? "s" : ""}</span>
                        </div>
                      </td>
                    </tr>
                    {g.rows.map((t, ri) => (
                      <tr key={t.id} className="gl-journal-row" style={{ background: ri % 2 === 1 ? "rgba(255,255,255,0.015)" : "transparent" }}>
                        <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}`, borderRight: `1px solid ${COLOR.hairline}` }}>
                          <button onClick={() => toggleSelect(t.id)} style={{ background: "transparent", border: "none", cursor: "pointer", display: "flex" }}>
                            {selected.has(t.id) ? <CheckSquare size={14} color={COLOR.goldSoft} /> : <Square size={14} color={COLOR.inkMuted} />}
                          </button>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 11.5, borderBottom: `1px solid ${COLOR.hairline}`, borderRight: `1px solid ${COLOR.hairline}`, fontFamily: "'IBM Plex Mono', monospace", color: COLOR.inkMuted }}>
                          {t.time || "—"}
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12.5, borderBottom: `1px solid ${COLOR.hairline}`, borderRight: `1px solid ${COLOR.hairline}`, maxWidth: 260 }}>
                          {t.category}{t.subcategory && <span style={{ color: COLOR.inkMuted }}> · {t.subcategory}</span>}
                          {t.payee && <div style={{ fontSize: 10.5, color: COLOR.inkMuted }}>{t.payee}</div>}
                          {t.note && <div style={{ fontSize: 10.5, color: COLOR.inkMuted, fontStyle: "italic", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>« {t.note} »</div>}
                        </td>
                        <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}`, borderRight: `1px solid ${COLOR.hairline}` }}>
                          <span style={{ fontSize: 10.5, fontWeight: 600, padding: "2px 9px", borderRadius: 20, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft, background: t.type === "Revenu" ? "rgba(63,156,122,0.12)" : "rgba(193,84,63,0.12)" }}>{t.type}</span>
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 11.5, borderBottom: `1px solid ${COLOR.hairline}`, borderRight: `1px solid ${COLOR.hairline}`, color: groupColor[t.group] }}>
                          {t.group}
                          {t.account ? (
                            <div style={{ color: COLOR.inkMuted, fontSize: 10.5, marginTop: 2 }}>{t.account}</div>
                          ) : (
                            <div style={{ color: COLOR.claySoft, fontSize: 10.5, marginTop: 2, display: "flex", alignItems: "center", gap: 3 }}><AlertTriangle size={9} /> sans compte</div>
                          )}
                          {t.onBehalfOf && (
                            <div style={{ color: COLOR.goldSoft, fontSize: 10, marginTop: 2, display: "flex", alignItems: "center", gap: 3 }} title="Avance entre comptes">
                              <ArrowRight size={9} /> pour {t.onBehalfOf}
                            </div>
                          )}
                        </td>
                        <td style={{ padding: "9px 10px", fontSize: 12.5, textAlign: "right", borderBottom: `1px solid ${COLOR.hairline}`, borderRight: `1px solid ${COLOR.hairline}`, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{t.type === "Revenu" ? "+" : "−"}{fmt(t.amount)}</td>
                        <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}`, whiteSpace: "nowrap" }}><button onClick={() => startEdit(t)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={13} /></button><button onClick={() => setConfirmDeleteId(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button></td>
                      </tr>
                    ))}
                    <tr>
                      <td colSpan={7} style={{ padding: "7px 12px", background: "rgba(91,126,166,0.09)", borderBottom: `2px solid ${COLOR.slateBlue}`, borderLeft: `3px solid ${COLOR.slateBlue}` }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 18, flexWrap: "wrap", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }}>
                          <span style={{ color: COLOR.slateBlueSoft, fontFamily: "'Inter', sans-serif", fontWeight: 600, marginRight: "auto" }}>Sous-total du jour</span>
                          <span style={{ color: COLOR.inkMuted }}>Revenus <b style={{ color: COLOR.emeraldSoft }}>{fmt(dt.rev)}</b></span>
                          <span style={{ color: COLOR.inkMuted }}>Dépenses <b style={{ color: COLOR.claySoft }}>{fmt(dt.dep)}</b></span>
                          <span style={{ color: COLOR.inkMuted }}>Solde <b style={{ color: dt.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(dt.solde)}</b></span>
                        </div>
                      </td>
                    </tr>
                  </React.Fragment>
                );
              })}
              {!pageRows.length && <tr><td colSpan={7}><EmptyState /></td></tr>}
            </tbody>
          </table>
        </div>
        {pageCount > 1 && (
          <div style={{ display: "flex", justifyContent: "center", gap: 8, marginTop: 16 }}>
            <button disabled={safePage === 0} onClick={() => setPage((p) => p - 1)} style={pagerBtn(safePage === 0)}>Précédent</button>
            <span style={{ fontSize: 12, color: COLOR.inkMuted, alignSelf: "center" }}>Page {safePage + 1} / {pageCount}</span>
            <button disabled={safePage >= pageCount - 1} onClick={() => setPage((p) => p + 1)} style={pagerBtn(safePage >= pageCount - 1)}>Suivant</button>
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
    const grandTotalRow = wsTx.addRow({ cat: "TOTAL GÉNÉRAL", account: `Rev: ${fmt(totalRevenus)}`, amount: totalDepenses });
    grandTotalRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
    grandTotalRow.getCell("amount").numFmt = "#,##0";

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
    const totalRowMonth = wsMonth.addRow({ mois: "TOTAL", rev: totalRevenus, dep: totalDepenses, solde: solde });
    totalRowMonth.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
    ["rev", "dep", "solde"].forEach((k) => { totalRowMonth.getCell(k).numFmt = "#,##0"; });

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
    const grandTotalCatRow = wsCat.addRow({ cat: "TOTAL GÉNÉRAL DÉPENSES (hors Revenu)", total: totalDepenses });
    grandTotalCatRow.eachCell((c: any) => { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } }; });
    grandTotalCatRow.getCell("total").numFmt = "#,##0";

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
      const [jsPDFModule, autoTableModule] = await Promise.all([
        import(/* @vite-ignore */ "jspdf"),
        import(/* @vite-ignore */ "jspdf-autotable"),
      ]);
      // jspdf@2.5.x expose le vrai constructeur sur l'export NOMMÉ "jsPDF", pas sur
      // "default" (qui résout vers un objet inutilisable selon le mode d'interop
      // CJS/ESM) — cause du "Réessayer" systématique sur tous les boutons PDF de l'app.
      const jsPDF: any = (jsPDFModule as any).jsPDF || (jsPDFModule as any).default;
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
      drawKpiBox(14, "REVENUS", `${fmtPdf(totalRevenus)} FCFA`, 63, 156, 122);
      drawKpiBox(14 + kpiW + 8, "DÉPENSES", `${fmtPdf(totalDepenses)} FCFA`, 193, 84, 63);
      drawKpiBox(14 + (kpiW + 8) * 2, "SOLDE", `${fmtPdf(solde)} FCFA`, solde >= 0 ? 63 : 193, solde >= 0 ? 156 : 84, solde >= 0 ? 122 : 63);

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
        rows.push([{ content: `▸ Sous-total ${monthLabel(curMonth)}`, colSpan: 4 }, { content: `Rev: ${fmtPdf(monthRev)} / Dép: ${fmtPdf(monthDep)}`, styles: { halign: "right" } }]);
      };
      sortedTx.forEach((t) => {
        if (curMonth !== null && t.month !== curMonth) { pushSubtotal(); monthRev = 0; monthDep = 0; }
        curMonth = t.month;
        if (t.type === "Revenu") monthRev += t.amount; else monthDep += t.amount;
        rows.push([dateLabelFull(t.date), t.time || "—", t.category + (t.subcategory ? ` · ${t.subcategory}` : ""), t.type, fmtPdf(t.amount)]);
      });
      pushSubtotal();

      doc.autoTable({
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
function SaisieQuotidienneTab({ transactions, setTransactions, allCategories, categoryGroups, accounts, monthlyObjective, setMonthlyObjective, chargeOverrides, includeGrundfosVoiture }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; allCategories: string[]; categoryGroups: Record<string, Group>; accounts: Account[];
  monthlyObjective: number; setMonthlyObjective: (n: number) => void; chargeOverrides: Record<string, ChargeOverride>; includeGrundfosVoiture: boolean;
}) {
  const [quickDate, setQuickDate] = useState(todayISO());
  const [quickTime, setQuickTime] = useState(nowTime());
  const [quickCategory, setQuickCategory] = useState(() => defaultQuickCategory(transactions, "Dépense"));
  const [quickSubcategory, setQuickSubcategory] = useState("");
  const [quickType, setQuickType] = useState<TxType>("Dépense");
  const [quickAmount, setQuickAmount] = useState<number | "">("");
  const [quickAccount, setQuickAccount] = useState(() => defaultQuickAccount(accounts));
  const [quickOnBehalfOf, setQuickOnBehalfOf] = useState("");
  const [quickNote, setQuickNote] = useState("");
  const [justAdded, setJustAdded] = useState(false);
  const [catPickerOpen, setCatPickerOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  const today = todayISO();
  const currentMonthKey = dateToMonthKey(today);
  const weekAgo = (() => { const d = new Date(); d.setDate(d.getDate() - 6); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; })();

  const withGroup = transactions.map((t) => ({ ...t, group: t.type === "Revenu" ? "Revenu" : groupFor(t, categoryGroups) }));

  const sumFor = (pred: (t: any) => boolean) => {
    const arr = withGroup.filter(pred);
    const rev = arr.filter((t) => t.type === "Revenu").reduce((a, t) => a + t.amount, 0);
    const dep = arr.filter((t) => t.type === "Dépense").reduce((a, t) => a + t.amount, 0);
    return { rev, dep, solde: rev - dep };
  };

  const todayTotals = sumFor((t) => t.date === today);
  const weekTotals = sumFor((t) => t.date >= weekAgo && t.date <= today);
  const monthTotals = sumFor((t) => dateToMonthKey(t.date) === currentMonthKey);

  // Répartition des dépenses par nature (Nécessaire/Productif/Non-productif) pour une
  // fenêtre donnée — sur demande explicite de l'utilisateur (11/08/2026), pour les
  // cartes "Aujourd'hui" et "Mois en cours — dépenses".
  const groupBreakdown = (pred: (t: any) => boolean) => {
    const dep = withGroup.filter((t) => t.type === "Dépense" && pred(t));
    const total = dep.reduce((a, t) => a + t.amount, 0) || 1;
    const byGroup: Record<string, number> = { "Nécessaire": 0, "Productif": 0, "Non-productif": 0, "Non classifié": 0 };
    dep.forEach((t) => { byGroup[t.group] = (byGroup[t.group] || 0) + t.amount; });
    return (["Nécessaire", "Productif", "Non-productif", "Non classifié"] as const)
      .map((g) => ({ group: g, value: byGroup[g], pct: (byGroup[g] / total) * 100 }))
      .filter((r) => r.value > 0);
  };
  const [kpiDetail, setKpiDetail] = useState<{ mode: "today" | "month"; anchor: string } | null>(null);

  // Comparaisons "vs même période le mois dernier" — même logique que la carte déjà
  // existante dans le Conseiller quotidien ("Mieux que le mois dernier à la même date"),
  // reprise ici pour chacun des 5 indicateurs, sur demande explicite de l'utilisateur.
  const prevMonthKeyVal = prevMonthKey(currentMonthKey);
  const dayNum = new Date(today + "T00:00:00").getDate();
  const clampDay = (mk: string, d: number) => Math.min(d, daysInMonthOf(mk));
  const mkDate = (mk: string, d: number) => { const [y, m] = mk.split("_"); return `${y}-${pad2(Number(m))}-${pad2(d)}`; };
  const sameDayLastMonth = mkDate(prevMonthKeyVal, clampDay(prevMonthKeyVal, dayNum));
  const weekStartDayLastMonth = mkDate(prevMonthKeyVal, clampDay(prevMonthKeyVal, Math.max(1, dayNum - 6)));

  const todayLastMonthTotals = sumFor((t) => t.date === sameDayLastMonth);
  const weekLastMonthTotals = sumFor((t) => t.date >= weekStartDayLastMonth && t.date <= sameDayLastMonth);
  const monthLastMonthTotals = sumFor((t) => dateToMonthKey(t.date) === prevMonthKeyVal && new Date(t.date + "T00:00:00").getDate() <= dayNum);

  // Construit la fiche de détail avec, pour chaque nature (Nécessaire/Productif/
  // Non-productif), l'écart entre la période actuelle et la période comparative — pour
  // répondre à la vraie question posée par l'utilisateur (11/08/2026) : une progression
  // ou une régression du total est-elle fondée (portée par le Non-productif, qu'on
  // maîtrise) ou inquiétante (si elle vient d'une baisse du Nécessaire, ce qui peut
  // vouloir dire qu'on se prive de l'essentiel plutôt que de vraiment progresser) ?
  // "anchor" est navigable (jour ou mois précédent/suivant) — corrigé le 12/08/2026 :
  // la fiche était figée sur "aujourd'hui" et devenait invisible dès que le jour passait.
  const buildKpiGroupDetail = (mode: "today" | "month", anchor: string) => {
    const anchorMonthKey = mode === "today" ? dateToMonthKey(anchor) : anchor;
    const anchorDayNum = mode === "today" ? new Date(anchor + "T00:00:00").getDate() : dayNum;
    const prevAnchorMonthKey = prevMonthKey(anchorMonthKey);
    const anchorSameDayLastMonth = mode === "today" ? mkDate(prevAnchorMonthKey, clampDay(prevAnchorMonthKey, anchorDayNum)) : "";

    const curPred = mode === "today" ? (t: any) => t.date === anchor : (t: any) => dateToMonthKey(t.date) === anchorMonthKey;
    const prevPred = mode === "today"
      ? (t: any) => t.date === anchorSameDayLastMonth
      : (t: any) => dateToMonthKey(t.date) === prevAnchorMonthKey && new Date(t.date + "T00:00:00").getDate() <= anchorDayNum;
    const curRows = groupBreakdown(curPred);
    const prevRows = groupBreakdown(prevPred);
    const curTotal = curRows.reduce((a, r) => a + r.value, 0);
    const prevTotal = prevRows.reduce((a, r) => a + r.value, 0);
    const curLabel = mode === "today" ? dateLabelFull(anchor) : monthLabel(anchorMonthKey);
    const prevLabel = `Même ${mode === "today" ? "jour" : "période"} le mois dernier (${monthLabel(prevAnchorMonthKey)})`;
    // Intitulés courts pour les en-têtes du tableau — la version complète reste dans la
    // ligne "formula" juste au-dessus, pour ne pas forcer un défilement horizontal sur
    // mobile avec des en-têtes trop longs, sur demande explicite de l'utilisateur (11/08/2026).
    const curLabelShort = mode === "today" ? "Ce jour" : "Ce mois";
    const prevLabelShort = "Mois dernier";

    const groups = ["Nécessaire", "Productif", "Non-productif", "Non classifié"] as const;
    const valueOf = (rows: typeof curRows, g: string) => rows.find((r) => r.group === g)?.value || 0;
    const deltas = groups.map((g) => {
      const cur = valueOf(curRows, g), prev = valueOf(prevRows, g), delta = cur - prev;
      const pct = prev > 0 ? (delta / prev) * 100 : (cur > 0 ? 100 : 0);
      return { group: g, cur, prev, delta, pct };
    }).filter((d) => d.cur > 0 || d.prev > 0);

    const totalDelta = curTotal - prevTotal;
    // Repère la nature qui explique le plus l'évolution du total (en valeur absolue).
    const biggestMover = [...deltas].sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta))[0];
    let verdict = "";
    if (Math.abs(totalDelta) < 1) {
      verdict = "Total quasi stable sur les deux périodes — rien de notable à signaler.";
    } else if (totalDelta < 0) {
      // Baisse des dépenses = a priori une bonne nouvelle, mais tout dépend d'où elle vient.
      if (biggestMover?.group === "Non-productif") verdict = `Baisse fondée : elle est portée principalement par le "Non-productif" (${fmt(Math.abs(biggestMover.delta))} FCFA de moins), le poste le plus facile à maîtriser sans rien sacrifier d'essentiel.`;
      else if (biggestMover?.group === "Nécessaire") verdict = `À surveiller : cette baisse vient surtout du "Nécessaire" (${fmt(Math.abs(biggestMover.delta))} FCFA de moins) — vérifie qu'il ne s'agit pas d'une privation plutôt que d'une vraie économie.`;
      else if (biggestMover?.group === "Productif") verdict = `Baisse portée par le "Productif" (${fmt(Math.abs(biggestMover.delta))} FCFA de moins) — à vérifier que ce n'est pas un investissement ou un remboursement simplement décalé dans le temps.`;
      else verdict = "Baisse du total, sans nature clairement dominante.";
    } else {
      // Hausse des dépenses = a priori à surveiller, mais peut être un investissement sain.
      if (biggestMover?.group === "Productif") verdict = `Hausse plutôt saine : elle est portée principalement par le "Productif" (${fmt(biggestMover.delta)} FCFA de plus), donc probablement de l'investissement plutôt que du gaspillage.`;
      else if (biggestMover?.group === "Non-productif") verdict = `À surveiller : cette hausse vient surtout du "Non-productif" (${fmt(biggestMover.delta)} FCFA de plus) — c'est le premier poste à réduire si besoin.`;
      else if (biggestMover?.group === "Nécessaire") verdict = `Hausse portée par le "Nécessaire" (${fmt(biggestMover.delta)} FCFA de plus) — vérifie si c'est ponctuel (ex: une charge exceptionnelle) ou un vrai changement de rythme.`;
      else verdict = "Hausse du total, sans nature clairement dominante.";
    }

    const totalPct = prevTotal > 0 ? (totalDelta / prevTotal) * 100 : (curTotal > 0 ? 100 : 0);

    return {
      title: `${curLabel} — dépenses par nature`,
      headline: `${fmt(curTotal)} FCFA (${totalDelta >= 0 ? "+" : ""}${fmt(totalDelta)} FCFA vs période comparative)`,
      formula: `${curLabel} vs ${prevLabel} — écart par nature (Nécessaire / Productif / Non-productif)`,
      blocks: [
        {
          kind: "table" as const,
          columns: ["Nature", curLabelShort, prevLabelShort, "Écart", "Évolution"],
          rows: deltas.map((d) => [d.group, fmt(d.cur), fmt(d.prev), `${d.delta >= 0 ? "+" : ""}${fmt(d.delta)}`, `${d.pct >= 0 ? "+" : ""}${d.pct.toFixed(0)}%`]),
          // Vert quand les dépenses de cette nature ont baissé, rouge quand elles ont augmenté —
          // la vraie lecture (bonne ou mauvaise nouvelle selon la nature) reste dans le verdict ci-dessous.
          cellColors: deltas.map((d) => [undefined, undefined, undefined, undefined, Math.abs(d.delta) < 1 ? COLOR.inkMuted : d.delta < 0 ? COLOR.emeraldSoft : COLOR.claySoft]),
          // Ligne de totaux (toutes natures confondues) en bas du tableau, pour lire le
          // solde global sans avoir à additionner les lignes soi-même — demande explicite
          // de l'utilisateur (12/08/2026).
          footerRow: ["Total", fmt(curTotal), fmt(prevTotal), `${totalDelta >= 0 ? "+" : ""}${fmt(totalDelta)}`, `${totalPct >= 0 ? "+" : ""}${totalPct.toFixed(0)}%`],
          footerColors: [undefined, undefined, undefined, undefined, Math.abs(totalDelta) < 1 ? COLOR.inkMuted : totalDelta < 0 ? COLOR.emeraldSoft : COLOR.claySoft],
        },
        { kind: "note" as const, tone: ((totalDelta < 0 && biggestMover?.group === "Nécessaire") || (totalDelta > 0 && biggestMover?.group !== "Productif") ? "warn" : "info") as "warn" | "info", text: verdict },
      ],
    };
  };

  const pctDelta = (cur: number, prev: number) => (prev !== 0 ? ((cur - prev) / Math.abs(prev)) * 100 : null);
  const vsHint = (cur: number, prev: number, suffix: string) => {
    const d = pctDelta(cur, prev);
    return `vs ${fmt(prev)} FCFA ${suffix}${d !== null ? ` (${d >= 0 ? "+" : ""}${d.toFixed(0)}%)` : ""}`;
  };

  const quickDateEntries = withGroup.filter((t) => t.date === quickDate).sort((a, b) => b.id.localeCompare(a.id));

  const resetForm = () => {
    setQuickAmount(""); setQuickTime(nowTime()); setQuickNote(""); setEditingId(null); setQuickOnBehalfOf("");
  };

  const submit = () => {
    if (!quickCategory || !quickAmount || Number(quickAmount) <= 0) return;
    const onBehalfOfVal = (quickOnBehalfOf && quickOnBehalfOf !== quickAccount) ? quickOnBehalfOf : undefined;
    if (editingId) {
      setTransactions(transactions.map((t) => t.id === editingId ? {
        ...t, date: quickDate, time: quickTime, category: quickCategory, subcategory: quickSubcategory || undefined,
        type: quickType, amount: Number(quickAmount), account: quickAccount || undefined, onBehalfOf: onBehalfOfVal, note: quickNote || undefined,
      } : t));
    } else {
      setTransactions([...transactions, { id: uid(), date: quickDate, time: quickTime, category: quickCategory, subcategory: quickSubcategory || undefined, type: quickType, amount: Number(quickAmount), account: quickAccount || undefined, onBehalfOf: onBehalfOfVal, note: quickNote || undefined }]);
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
    setQuickOnBehalfOf(t.onBehalfOf || "");
    setQuickNote(t.note || "");
  };

  const remove = (id: string) => setTransactions(transactions.filter((t) => t.id !== id));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <DayScoreBadge transactions={transactions} monthlyObjective={monthlyObjective} scope="jour" />
        <DayScoreBadge transactions={transactions} monthlyObjective={monthlyObjective} scope="mois" />
      </div>
      <DailyAdvisorButton transactions={transactions} monthlyObjective={monthlyObjective} setMonthlyObjective={setMonthlyObjective} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} />
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Kpi label="Aujourd'hui — solde" value={fmt(todayTotals.solde)} tone={todayTotals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={Clock} hint={vsHint(todayTotals.solde, todayLastMonthTotals.solde, "même jour le mois dernier")} hintBadge={compareLabel(pctDelta(todayTotals.solde, todayLastMonthTotals.solde), "up")} onDetailClick={() => setKpiDetail({ mode: "today", anchor: today })} />
        <Kpi label="7 derniers jours — solde" value={fmt(weekTotals.solde)} tone={weekTotals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={CalendarDays} hint={vsHint(weekTotals.solde, weekLastMonthTotals.solde, "même période le mois dernier")} hintBadge={compareLabel(pctDelta(weekTotals.solde, weekLastMonthTotals.solde), "up")} />
        <Kpi label="Mois en cours — solde" value={fmt(monthTotals.solde)} tone={monthTotals.solde >= 0 ? COLOR.emeraldSoft : COLOR.claySoft} icon={CalendarRange} hint={vsHint(monthTotals.solde, monthLastMonthTotals.solde, "au même jour le mois dernier")} hintBadge={compareLabel(pctDelta(monthTotals.solde, monthLastMonthTotals.solde), "up")} />
        <Kpi label="Mois en cours — dépenses" value={fmt(monthTotals.dep)} tone={COLOR.claySoft} icon={TrendingDown} hint={vsHint(monthTotals.dep, monthLastMonthTotals.dep, "au même jour le mois dernier")} hintBadge={compareLabel(pctDelta(monthTotals.dep, monthLastMonthTotals.dep), "down")} onDetailClick={() => setKpiDetail({ mode: "month", anchor: currentMonthKey })} />
        <Kpi label="Mois en cours — revenus" value={fmt(monthTotals.rev)} tone={COLOR.emeraldSoft} icon={TrendingUp} hint={vsHint(monthTotals.rev, monthLastMonthTotals.rev, "au même jour le mois dernier")} hintBadge={compareLabel(pctDelta(monthTotals.rev, monthLastMonthTotals.rev), "up")} />
      </div>

      {kpiDetail && (() => {
        const d = buildKpiGroupDetail(kpiDetail.mode, kpiDetail.anchor);
        const goPrev = () => setKpiDetail({ mode: kpiDetail.mode, anchor: kpiDetail.mode === "today" ? addDays(kpiDetail.anchor, -1) : prevMonthKey(kpiDetail.anchor) });
        const goNext = () => setKpiDetail({ mode: kpiDetail.mode, anchor: kpiDetail.mode === "today" ? addDays(kpiDetail.anchor, 1) : nextMonthKey(kpiDetail.anchor) });
        const atToday = kpiDetail.mode === "today" ? kpiDetail.anchor >= today : kpiDetail.anchor >= currentMonthKey;
        return (
          <CalcDetailSheet open={!!kpiDetail} onClose={() => setKpiDetail(null)} title={d.title} headline={d.headline} formula={d.formula} blocks={d.blocks}
            onPrev={goPrev} onNext={atToday ? undefined : goNext} />
        );
      })()}

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
                <input
                  value={quickNote} onChange={(e) => setQuickNote(e.target.value)} placeholder="Ajouter une note"
                  style={{ position: "relative", background: "transparent", border: "none", outline: "none", marginTop: 10, color: COLOR.inkMuted, fontSize: 13.5, fontFamily: "'Inter', sans-serif", textAlign: "center", width: "100%", maxWidth: 320 }}
                />
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
                {accounts.length > 1 && (
                  <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px dashed ${COLOR.hairline}` }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: COLOR.inkMuted, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!quickOnBehalfOf} onChange={(e) => setQuickOnBehalfOf(e.target.checked ? (accounts.find((a) => a.name !== quickAccount)?.name || "") : "")} />
                      Payée depuis {quickAccount || "ce compte"} mais destinée à un autre compte (avance entre comptes)
                    </label>
                    {quickOnBehalfOf && (
                      <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Compte réellement concerné :</span>
                        <select value={quickOnBehalfOf} onChange={(e) => setQuickOnBehalfOf(e.target.value)} style={{ ...inputStyle, flex: 1 }}>
                          {accounts.filter((a) => a.name !== quickAccount).map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                        </select>
                      </div>
                    )}
                  </div>
                )}
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
            <div key={t.id} style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", gap: "4px 10px", padding: "8px 12px", background: editingId === t.id ? "rgba(201,162,39,0.08)" : COLOR.surfaceRaised, border: editingId === t.id ? `1px solid ${COLOR.gold}` : "1px solid transparent", borderRadius: 6 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0, flex: "1 1 160px", overflow: "hidden" }}>
                {t.time && <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: COLOR.inkMuted, flexShrink: 0 }}>{t.time}</span>}
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: groupColor[t.group] || COLOR.inkMuted, display: "inline-block", flexShrink: 0 }} />
                <span style={{ fontSize: 12.5, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 }}>
                  {t.category}{t.subcategory && ` · ${t.subcategory}`}
                  {t.account && <span style={{ color: COLOR.slateBlueSoft }}> · {t.account}</span>}
                  {t.note && <span style={{ color: COLOR.inkMuted, fontStyle: "italic" }}> · « {t.note} »</span>}
                </span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0, marginLeft: "auto" }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft, whiteSpace: "nowrap" }}>{fmt(t.amount)}</span>
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
  const [onBehalfOf, setOnBehalfOf] = useState("");
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
      account: account || undefined, onBehalfOf: (onBehalfOf && onBehalfOf !== account) ? onBehalfOf : undefined, note: note || undefined,
    }]);
    setAmount(""); setNote(""); setOnBehalfOf("");
    setJustAdded(true);
    setTimeout(() => setJustAdded(false), 1000);
  };

  const changeType = (ty: TxType) => {
    setType(ty);
    setSubcategory("");
    setCategory(defaultQuickCategory(transactions, ty));
    if (ty !== "Dépense") setOnBehalfOf("");
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
              {accounts.length > 1 && (
                <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px dashed ${COLOR.hairline}` }}>
                  <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: COLOR.inkMuted, cursor: "pointer" }}>
                    <input type="checkbox" checked={!!onBehalfOf} onChange={(e) => setOnBehalfOf(e.target.checked ? (accounts.find((a) => a.name !== account)?.name || "") : "")} />
                    Payée depuis {account || "ce compte"} mais destinée à un autre compte (avance entre comptes)
                  </label>
                  {onBehalfOf && (
                    <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ fontSize: 11.5, color: COLOR.inkMuted }}>Compte réellement concerné :</span>
                      <select value={onBehalfOf} onChange={(e) => setOnBehalfOf(e.target.value)} style={{ ...inputStyle, flex: 1 }}>
                        {accounts.filter((a) => a.name !== account).map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                      </select>
                    </div>
                  )}
                </div>
              )}
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
type Tab = "saisie" | "apercu" | "valeurnette" | "flux" | "comparatif" | "comparateur" | "topcategories" | "categoryoverview" | "mensuel" | "journalier" | "categories" | "gestioncategories" | "groupes" | "enveloppes" | "budgets" | "simulateur" | "objectif" | "business" | "activites" | "charges" | "diagnostic" | "rapprochement" | "creances" | "comptes" | "payees" | "recurrences" | "journal" | "export" | "sauvegarde";

const NAV: { section: string; items: { id: Tab; label: string; icon: any }[] }[] = [
  { section: "Saisie rapide", items: [
    { id: "saisie", label: "Saisie du jour", icon: Clock },
  ]},
  { section: "Tableau de bord", items: [
    { id: "apercu", label: "Aperçu", icon: LayoutDashboard },
    { id: "valeurnette", label: "Valeur nette", icon: Wallet },
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
    { id: "gestioncategories", label: "Gestion des catégories", icon: Layers },
    { id: "groupes", label: "Groupes", icon: Layers },
    { id: "enveloppes", label: "Enveloppes", icon: Mail },
    { id: "budgets", label: "Budgets par catégorie", icon: ClipboardList },
  ]},
  { section: "Outils", items: [
    { id: "simulateur", label: "Simulateur", icon: SlidersHorizontal },
    { id: "objectif", label: "Objectifs & Projection", icon: Gauge },
    { id: "business", label: "Business / Personnel", icon: Briefcase },
    { id: "activites", label: "Activités & Rentabilité", icon: Rocket },
    { id: "charges", label: "Charges Fixes & Variables", icon: CalendarRange },
    { id: "diagnostic", label: "Diagnostic Financier", icon: Gauge },
    { id: "rapprochement", label: "Rapprochement bancaire", icon: CheckSquare },
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
  // Contrôle le blocage complet des sauvegardes tant qu'un choix explicite n'a pas été
  // fait après un démarrage à vide (voir plus bas : écran de blocage). Tant que
  // holdSave=true, RIEN n'est écrit dans le localStorage, pour ne jamais transformer
  // silencieusement les données de secours en "vraies" données.
  const [dataGateResolved, setDataGateResolved] = useState(false);
  const [transactions, setTransactions, txLoaded, txStartedEmpty] = usePersistentState<Transaction[]>("gl-transactions", seedTransactions, dataGateResolved);
  // Corrigé le 11/08/2026 : ce verrou (voulu pour bloquer la sauvegarde tant que l'écran
  // de choix "aucune donnée trouvée" n'est pas résolu) bloquait aussi, par effet de bord,
  // toute clé fraîchement introduite dans le code (ex: "gl-deleted-tx-ids", qui n'existait
  // jamais avant) sur un appareil qui a pourtant déjà de vraies données ailleurs. Une
  // suppression de transaction n'était donc jamais vraiment mémorisée d'une session à
  // l'autre, et revenait à chaque rechargement de page. Le verrou ne doit s'appliquer que
  // si CE périphérique démarre vraiment à vide (txStartedEmpty) — jamais pour une clé
  // simplement nouvelle sur un appareil qui a déjà d'autres données établies.
  const canSaveGated = !txStartedEmpty || dataGateResolved;
  // Corrigé le 11/08/2026 : une transaction supprimée revenait toute seule après une
  // synchronisation — cause identifiée avec certitude : la fusion (mergeById) ne fait
  // qu'AJOUTER ce qui manque d'un côté, elle ne peut pas distinguer "l'autre appareil a
  // ajouté ça, je ne le connais pas encore" de "j'ai supprimé ça volontairement, l'autre
  // appareil n'est pas encore au courant". Sans mémoire des suppressions, toute
  // transaction supprimée localement mais encore présente côté serveur revenait à la
  // prochaine synchronisation. On garde donc la liste des identifiants supprimés
  // (jamais oubliée), qui a toujours le dernier mot sur une fusion.
  const [deletedTransactionIds, setDeletedTransactionIds] = usePersistentState<Record<string, string>>("gl-deleted-tx-ids", {}, canSaveGated);
  const setTransactionsTracked = (next: Transaction[]) => {
    const nextIds = new Set(next.map((t) => t.id));
    const removedIds = transactions.filter((t) => !nextIds.has(t.id)).map((t) => t.id);
    if (removedIds.length) {
      const now = new Date().toISOString();
      setDeletedTransactionIds({ ...deletedTransactionIds, ...Object.fromEntries(removedIds.map((id) => [id, now])) });
    }
    // Pose automatiquement un horodatage sur toute transaction NOUVELLE ou dont le
    // contenu a réellement changé par rapport à ce qu'on connaissait avant — peu importe
    // le formulaire ou le bouton utilisé pour la modifier. C'est cet horodatage que la
    // fusion de synchronisation utilise pour départager deux versions d'une même
    // transaction (voir mergeById) — poser l'horodatage ici, une seule fois, garantit
    // que ça marche pour TOUS les points de saisie sans avoir à y penser à chaque fois.
    const now = new Date().toISOString();
    const prevById = new Map(transactions.map((t) => [t.id, t]));
    const stamped = next.map((t) => {
      const prev = prevById.get(t.id);
      if (!prev) return { ...t, updatedAt: now };
      const { updatedAt: _p, ...prevRest } = prev;
      const { updatedAt: _n, ...tRest } = t;
      if (JSON.stringify(prevRest) !== JSON.stringify(tRest)) return { ...t, updatedAt: now };
      return t;
    });
    setTransactions(stamped);
  };
  const [categoryGroups, setCategoryGroups, groupsLoaded] = usePersistentState<Record<string, Group>>("gl-category-groups", defaultCategoryGroups, canSaveGated);
  const [categoryScope, setCategoryScope, scopeLoaded] = usePersistentState<Record<string, Scope>>("gl-category-scope", defaultCategoryScope, canSaveGated);
  const [activities, setActivities] = usePersistentState<string[]>("gl-activities", defaultActivities, canSaveGated);
  const [categoryActivity, setCategoryActivity] = usePersistentState<Record<string, string>>("gl-category-activity", defaultCategoryActivity, canSaveGated);
  const [activityCapital, setActivityCapital] = usePersistentState<Record<string, number>>("gl-activity-capital", {}, canSaveGated);
  const [monthlyObjective, setMonthlyObjective] = usePersistentState<number>("gl-monthly-objective", 0, canSaveGated);
  const [chargeOverrides, setChargeOverrides] = usePersistentState<Record<string, ChargeOverride>>("gl-charge-overrides", defaultChargeOverrides, canSaveGated);
  const [includeGrundfosVoiture, setIncludeGrundfosVoiture] = usePersistentState<boolean>("gl-include-grundfos-voiture", true, canSaveGated);
  const [preRestoreSnapshot, setPreRestoreSnapshot] = usePersistentState<any>("gl-pre-restore-snapshot", null);
  const [preRestoreSnapshotAt, setPreRestoreSnapshotAt] = usePersistentState<string | null>("gl-pre-restore-snapshot-at", null);
  const [settingsLog, setSettingsLog] = usePersistentState<SettingsLogEntry[]>("gl-settings-log", []);
  const [dismissedReminderDate, setDismissedReminderDate] = usePersistentState<string | null>("gl-dismissed-reminder-date", null);
  const [customDepSubcategories, setCustomDepSubcategories] = usePersistentState<Record<string, string[]>>("gl-custom-dep-subcats", depSubcategories, canSaveGated);
  const [customRevSubcategories, setCustomRevSubcategories] = usePersistentState<Record<string, string[]>>("gl-custom-rev-subcats", revSubcategories, canSaveGated);
  CUSTOM_DEP_SUBCATS = customDepSubcategories;
  CUSTOM_REV_SUBCATS = customRevSubcategories;
  const logChange = (text: string) => setSettingsLog([{ at: `${dateLabelFull(todayISO())} à ${nowTime()}`, text }, ...settingsLog].slice(0, 300));

  const chargeModeDesc = (o?: ChargeOverride) => !o || o.mode === "auto" ? "Auto" : o.mode === "fixe" ? `Fixe${o.amount !== undefined ? ` (${fmt(o.amount)} FCFA)` : ""}` : o.mode === "variable" ? "Variable régulière" : o.mode === "exclu" ? "Exclu" : "Occasionnelle";
  const setChargeOverridesLogged = (next: Record<string, ChargeOverride>) => {
    Object.keys(next).forEach((k) => {
      if (JSON.stringify(next[k]) !== JSON.stringify(chargeOverrides[k])) {
        logChange(`Charge "${k.replace("::", " · ")}" : ${chargeModeDesc(chargeOverrides[k])} → ${chargeModeDesc(next[k])}`);
      }
    });
    setChargeOverrides(next);
  };
  const setCategoryActivityLogged = (next: Record<string, string>) => {
    Object.keys(next).forEach((k) => {
      if (next[k] !== categoryActivity[k]) logChange(`Catégorie "${k}" rattachée à l'activité "${next[k]}" (avant : "${categoryActivity[k] || "Personnel"}")`);
    });
    setCategoryActivity(next);
  };
  const setActivityCapitalLogged = (next: Record<string, number>) => {
    Object.keys(next).forEach((k) => {
      if (next[k] !== activityCapital[k]) logChange(`Capital investi de l'activité "${k}" : ${fmt(activityCapital[k] || 0)} → ${fmt(next[k])} FCFA`);
    });
    setActivityCapital(next);
  };
  const setActivitiesLogged = (next: string[]) => {
    next.filter((a) => !activities.includes(a)).forEach((a) => logChange(`Activité "${a}" ajoutée`));
    activities.filter((a) => !next.includes(a)).forEach((a) => logChange(`Activité "${a}" supprimée`));
    setActivities(next);
  };
  const setMonthlyObjectiveLogged = (next: number) => {
    if (next !== monthlyObjective) logChange(`Objectif de dépenses mensuel : ${fmt(monthlyObjective)} → ${fmt(next)} FCFA`);
    setMonthlyObjective(next);
  };
  const setIncludeGrundfosVoitureLogged = (next: boolean) => {
    if (next !== includeGrundfosVoiture) logChange(`GRUNDFOS dans les charges : ${includeGrundfosVoiture ? "Inclus" : "Exclu"} → ${next ? "Inclus" : "Exclu"}`);
    setIncludeGrundfosVoiture(next);
  };
  const setCategoryScopeLogged = (next: Record<string, Scope>) => {
    Object.keys(next).forEach((k) => {
      if (next[k] !== categoryScope[k]) logChange(`Portée de "${k}" : ${categoryScope[k] || "Personnel"} → ${next[k]}`);
    });
    setCategoryScope(next);
  };
  const [rules, setRules, rulesLoaded] = usePersistentState<CategorizationRule[]>("gl-rules", defaultRules, canSaveGated);
  const [loans, setLoans, loansLoaded] = usePersistentState<Loan[]>("gl-loans", seedLoans, canSaveGated);
  const [envelopeCap, setEnvelopeCap, capLoaded] = usePersistentState<number>("gl-envelope-cap", 600000, canSaveGated);
  const [accounts, setAccounts, accountsLoaded] = usePersistentState<Account[]>("gl-accounts", seedAccounts, canSaveGated);
  const [budgets, setBudgets, budgetsLoaded] = usePersistentState<CategoryBudget[]>("gl-budgets", seedBudgets, canSaveGated);
  const [goals, setGoals, goalsLoaded] = usePersistentState<Goal[]>("gl-goals", seedGoals, canSaveGated);
  const [recurring, setRecurring, recurringLoaded] = usePersistentState<RecurringTemplate[]>("gl-recurring", seedRecurring, canSaveGated);
  const [tab, setTab] = useState<Tab>("saisie");
  // Navigation contextuelle entre pages : navigateTo("categoryoverview", { category: "Shopping" })
  // change d'onglet ET transmet un contexte que la page de destination applique à son
  // ouverture (ex : catégorie déjà sélectionnée) — pour relier les pages entre elles
  // selon leur logique plutôt que de forcer une navigation manuelle à chaque fois.
  const [navContext, setNavContext] = useState<{ tab: Tab; data: any } | null>(null);
  const navigateTo = (target: Tab, data?: any) => { setNavContext({ tab: target, data }); setTab(target); };
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
  const allMonths = useMemo(() => {
    const s = new Set(transactions.filter((t) => t).map((t) => dateToMonthKey(t.date)));
    return Array.from(s).sort((a, b) => monthSortKey(a) - monthSortKey(b));
  }, [transactions]);

  const allCategories = useMemo(() => {
    const s = new Set(transactions.map((t) => t.category));
    return Array.from(s).sort();
  }, [transactions]);
  const groupedCatOptions = useMemo(() => groupedCategoryOptions(transactions), [transactions]);

  const defaultFilters: Filters = {
    from: allMonths[0] || "2024_6", to: allMonths[allMonths.length - 1] || "2026_8",
    type: "Tous", group: "Tous", category: "Toutes", subcategory: "Toutes", search: "", scope: "Tous", accounts: [],
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
    group: t.type === "Revenu" ? "Revenu" : groupFor(t, resolvedGroups),
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
      if (filters.accounts.length && !filters.accounts.includes(t.account || "")) return false;
      return true;
    });
  }, [txWithGroup, filters]);

  const allLoaded = txLoaded && groupsLoaded && scopeLoaded && rulesLoaded && loansLoaded && capLoaded && accountsLoaded && budgetsLoaded && goalsLoaded && recurringLoaded && syncCodeLoaded;

  // ============================================================
  // MOTEUR DE SYNCHRONISATION UNIFIÉ — reconstruit le 11/08/2026
  // ============================================================
  // Avant : 6 mécanismes séparés (chargement, édition, reconnexion réseau, retour au
  // premier plan, intervalle, notification temps réel) coordonnés entre eux par un
  // drapeau partagé fragile ("skipNextPush") — d'où des incohérences observées en
  // pratique (suppressions qui reviennent, statut qui s'affole, appareils qui ne se
  // rejoignent pas). Reconstruit en un seul moteur (runSync), appelé par tous ces
  // déclencheurs, avec deux garanties strictes :
  //   1. Jamais deux synchronisations en même temps (une file d'attente d'une place).
  //   2. Toujours les données les plus fraîches, lues via une référence mise à jour à
  //      CHAQUE rendu — plus aucun risque de "fermeture périmée" (stale closure), quel
  //      que soit le déclencheur ou son ancienneté.

  // Référence toujours à jour de tout l'état synchronisé — mise à jour à chaque rendu,
  // jamais figée dans la fermeture d'un effet créé plus tôt.
  const syncedRef = useRef({
    transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
    activities, categoryActivity, activityCapital, monthlyObjective, chargeOverrides, includeGrundfosVoiture,
    customDepSubcategories, customRevSubcategories, deletedTransactionIds,
  });
  useEffect(() => {
    syncedRef.current = {
      transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
      activities, categoryActivity, activityCapital, monthlyObjective, chargeOverrides, includeGrundfosVoiture,
      customDepSubcategories, customRevSubcategories, deletedTransactionIds,
    };
  });
  const syncCodeRef = useRef(syncCode);
  useEffect(() => { syncCodeRef.current = syncCode; }, [syncCode]);
  // Retient la date de la dernière VRAIE modification locale (persistée, pour survivre à
  // un rechargement) — comparée à la date de dernière mise à jour distante pour les
  // réglages scalaires (pas de fusion possible pour un simple nombre ou booléen) :
  // celui modifié le plus récemment l'emporte.
  const lastLocalChangeRef = useRef<string>((() => { try { return localStorage.getItem("gl-last-local-change") || ""; } catch { return ""; } })());
  const localChangeTrackInit = useRef(false);
  const syncInFlight = useRef(false);
  const syncQueued = useRef(false);

  // Le moteur lui-même : AUCUNE dépendance (tableau vide) — reste la même fonction sur
  // toute la durée de vie du composant, donc toujours sûr à passer à n'importe quel
  // écouteur ou intervalle créé une seule fois, sans jamais devenir périmé, puisqu'il
  // ne lit l'état qu'au travers de syncedRef.current au moment de l'exécution.
  const runSync = React.useCallback(async () => {
    const code = syncCodeRef.current;
    if (!SYNC_ENABLED || !code) return;
    if (syncInFlight.current) { syncQueued.current = true; return; }
    syncInFlight.current = true;
    setSyncStatus("syncing");
    try {
      const s = syncedRef.current;
      const remote = await fetchRemoteState(code);
      if (!remote) { setSyncStatus("error"); return; }
      const d = remote.data || {};
      const eq = (a: any, b: any) => JSON.stringify(a) === JSON.stringify(b);

      const mergedDeletedIds = d.deletedTransactionIds ? { ...d.deletedTransactionIds, ...s.deletedTransactionIds } : s.deletedTransactionIds;
      // Calculée en premier : décide qui gagne un vrai conflit (même id modifié des deux
      // côtés) dans les fusions ci-dessous — voir la note sur mergeById plus haut.
      const localIsNewer = !!(lastLocalChangeRef.current && (!remote.updatedAt || new Date(lastLocalChangeRef.current) > new Date(remote.updatedAt)));
      const mergedTransactions = (d.transactions ? mergeById(s.transactions, d.transactions) : s.transactions).filter((t: Transaction) => !mergedDeletedIds[t.id]);
      const mergedLoans = d.loans ? mergeById(s.loans, d.loans) : s.loans;
      const mergedBudgets = d.budgets ? mergeById(s.budgets, d.budgets) : s.budgets;
      const mergedGoals = d.goals ? mergeById(s.goals, d.goals) : s.goals;
      const mergedRecurring = d.recurring ? mergeById(s.recurring, d.recurring) : s.recurring;
      const mergedRules = d.rules ? mergeById(s.rules, d.rules) : s.rules;
      const mergedAccounts = d.accounts ? mergeById(s.accounts, d.accounts) : s.accounts;
      const mergedActivities = d.activities ? Array.from(new Set([...s.activities, ...d.activities])) : s.activities;
      const mergedCategoryGroups = d.categoryGroups ? { ...d.categoryGroups, ...s.categoryGroups } : s.categoryGroups;
      const mergedCategoryScope = d.categoryScope ? { ...d.categoryScope, ...s.categoryScope } : s.categoryScope;
      const mergedChargeOverrides = d.chargeOverrides ? { ...d.chargeOverrides, ...s.chargeOverrides } : s.chargeOverrides;
      const mergedCategoryActivity = d.categoryActivity ? { ...d.categoryActivity, ...s.categoryActivity } : s.categoryActivity;
      const mergedActivityCapital = d.activityCapital ? { ...d.activityCapital, ...s.activityCapital } : s.activityCapital;
      const mergedCustomDep = d.customDepSubcategories ? { ...d.customDepSubcategories, ...s.customDepSubcategories } : s.customDepSubcategories;
      const mergedCustomRev = d.customRevSubcategories ? { ...d.customRevSubcategories, ...s.customRevSubcategories } : s.customRevSubcategories;

      const finalEnvelopeCap = localIsNewer || typeof d.envelopeCap !== "number" ? s.envelopeCap : d.envelopeCap;
      const finalMonthlyObjective = localIsNewer || typeof d.monthlyObjective !== "number" ? s.monthlyObjective : d.monthlyObjective;
      const finalIncludeGrundfos = localIsNewer || typeof d.includeGrundfosVoiture !== "boolean" ? s.includeGrundfosVoiture : d.includeGrundfosVoiture;

      const localChanged = !eq(mergedTransactions, s.transactions) || !eq(mergedLoans, s.loans) || !eq(mergedBudgets, s.budgets)
        || !eq(mergedGoals, s.goals) || !eq(mergedRecurring, s.recurring) || !eq(mergedRules, s.rules) || !eq(mergedAccounts, s.accounts)
        || !eq(mergedActivities, s.activities) || !eq(mergedCategoryGroups, s.categoryGroups) || !eq(mergedCategoryScope, s.categoryScope)
        || !eq(mergedChargeOverrides, s.chargeOverrides) || !eq(mergedCategoryActivity, s.categoryActivity) || !eq(mergedActivityCapital, s.activityCapital)
        || !eq(mergedCustomDep, s.customDepSubcategories) || !eq(mergedCustomRev, s.customRevSubcategories) || !eq(mergedDeletedIds, s.deletedTransactionIds)
        || finalEnvelopeCap !== s.envelopeCap || finalMonthlyObjective !== s.monthlyObjective || finalIncludeGrundfos !== s.includeGrundfosVoiture;
      const remoteNeedsUpdate = !eq(mergedTransactions, d.transactions) || !eq(mergedLoans, d.loans) || !eq(mergedBudgets, d.budgets)
        || !eq(mergedGoals, d.goals) || !eq(mergedRecurring, d.recurring) || !eq(mergedRules, d.rules) || !eq(mergedAccounts, d.accounts)
        || !eq(mergedActivities, d.activities) || !eq(mergedCategoryGroups, d.categoryGroups) || !eq(mergedCategoryScope, d.categoryScope)
        || !eq(mergedChargeOverrides, d.chargeOverrides) || !eq(mergedCategoryActivity, d.categoryActivity) || !eq(mergedActivityCapital, d.activityCapital)
        || !eq(mergedCustomDep, d.customDepSubcategories) || !eq(mergedCustomRev, d.customRevSubcategories) || !eq(mergedDeletedIds, d.deletedTransactionIds)
        || finalEnvelopeCap !== d.envelopeCap || finalMonthlyObjective !== d.monthlyObjective || finalIncludeGrundfos !== d.includeGrundfosVoiture;

      if (localChanged) {
        skipNextPush.current = true;
        setTransactions(mergedTransactions);
        setLoans(mergedLoans);
        setBudgets(mergedBudgets);
        setGoals(mergedGoals);
        setRecurring(mergedRecurring);
        setRules(mergedRules);
        setAccounts(mergedAccounts);
        setActivities(mergedActivities);
        setCategoryGroups(mergedCategoryGroups);
        setCategoryScope(mergedCategoryScope);
        setChargeOverrides(mergedChargeOverrides);
        setCategoryActivity(mergedCategoryActivity);
        setActivityCapital(mergedActivityCapital);
        setCustomDepSubcategories(mergedCustomDep);
        setCustomRevSubcategories(mergedCustomRev);
        setEnvelopeCap(finalEnvelopeCap);
        setMonthlyObjective(finalMonthlyObjective);
        setIncludeGrundfosVoiture(finalIncludeGrundfos);
        setDeletedTransactionIds(mergedDeletedIds);
        // Garde le ref à jour immédiatement (avant même le prochain rendu), pour que si
        // une synchronisation en attente se déclenche tout de suite après, elle reparte
        // bien de ce résultat fusionné plutôt que de l'ancien état.
        syncedRef.current = {
          transactions: mergedTransactions, loans: mergedLoans, budgets: mergedBudgets, goals: mergedGoals, recurring: mergedRecurring,
          rules: mergedRules, accounts: mergedAccounts, activities: mergedActivities, categoryGroups: mergedCategoryGroups,
          categoryScope: mergedCategoryScope, chargeOverrides: mergedChargeOverrides, categoryActivity: mergedCategoryActivity,
          activityCapital: mergedActivityCapital, customDepSubcategories: mergedCustomDep, customRevSubcategories: mergedCustomRev,
          envelopeCap: finalEnvelopeCap, monthlyObjective: finalMonthlyObjective, includeGrundfosVoiture: finalIncludeGrundfos,
          deletedTransactionIds: mergedDeletedIds,
        };
      }

      if (!remoteNeedsUpdate) {
        setSyncStatus("synced");
        setLastSyncedAt(new Date().toLocaleTimeString("fr-FR"));
      } else {
        const ok = await pushRemoteState(code, {
          transactions: mergedTransactions, categoryGroups: mergedCategoryGroups, categoryScope: mergedCategoryScope, rules: mergedRules,
          loans: mergedLoans, envelopeCap: finalEnvelopeCap, accounts: mergedAccounts, budgets: mergedBudgets, goals: mergedGoals, recurring: mergedRecurring,
          activities: mergedActivities, categoryActivity: mergedCategoryActivity, activityCapital: mergedActivityCapital, monthlyObjective: finalMonthlyObjective,
          chargeOverrides: mergedChargeOverrides, includeGrundfosVoiture: finalIncludeGrundfos, customDepSubcategories: mergedCustomDep, customRevSubcategories: mergedCustomRev,
          deletedTransactionIds: mergedDeletedIds,
        });
        setSyncStatus(ok ? "synced" : "error");
        if (ok) setLastSyncedAt(new Date().toLocaleTimeString("fr-FR"));
      }
    } catch {
      setSyncStatus("error");
    } finally {
      syncInFlight.current = false;
      // Une demande est arrivée pendant qu'on synchronisait déjà — on la rejoue une
      // seule fois, avec les données les plus fraîches (pas d'accumulation infinie).
      if (syncQueued.current) { syncQueued.current = false; runSync(); }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Chargement initial (ou reconnexion via ?sync=... dans l'URL) : synchronise dès que
  // possible, et débloque l'écran de choix (données de démarrage à vide) si ça réussit.
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) return;
    (async () => {
      await runSync();
      if (txStartedEmpty) setDataGateResolved(true);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allLoaded, syncCode]);

  // Note la date de toute VRAIE modification locale (pas une fusion venue d'ailleurs,
  // filtrée via skipNextPush, ni le tout premier rendu au chargement) — persistée pour
  // survivre à un rechargement.
  useEffect(() => {
    if (!allLoaded) return;
    if (!localChangeTrackInit.current) { localChangeTrackInit.current = true; return; }
    if (skipNextPush.current) { skipNextPush.current = false; return; }
    const now = new Date().toISOString();
    lastLocalChangeRef.current = now;
    try { localStorage.setItem("gl-last-local-change", now); } catch {}
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring, activities, categoryActivity, activityCapital, monthlyObjective, chargeOverrides, includeGrundfosVoiture, customDepSubcategories, customRevSubcategories, deletedTransactionIds, allLoaded]);

  // Pousse/tire après chaque modification locale, avec un court délai pour regrouper les
  // changements rapprochés (ex: plusieurs champs modifiés d'un coup).
  const editSyncTimer = useRef<any>(null);
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) return;
    if (editSyncTimer.current) clearTimeout(editSyncTimer.current);
    editSyncTimer.current = setTimeout(runSync, 500);
    return () => { if (editSyncTimer.current) clearTimeout(editSyncTimer.current); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring, activities, categoryActivity, activityCapital, monthlyObjective, chargeOverrides, includeGrundfosVoiture, customDepSubcategories, customRevSubcategories, deletedTransactionIds, syncCode, allLoaded]);

  // Trois filets de sécurité indépendants, tous sûrs à utiliser puisque runSync ne
  // devient jamais périmé (aucune dépendance) : reconnexion réseau, retour au premier
  // plan de l'app, et un sondage toutes les 20 secondes en dernier recours.
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) return;
    const onOnline = () => runSync();
    const onVisible = () => { if (document.visibilityState === "visible") runSync(); };
    window.addEventListener("online", onOnline);
    document.addEventListener("visibilitychange", onVisible);
    const interval = setInterval(runSync, 20000);
    return () => {
      window.removeEventListener("online", onOnline);
      document.removeEventListener("visibilitychange", onVisible);
      clearInterval(interval);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allLoaded, syncCode]);

  // Abonnement temps réel : dès qu'un autre appareil modifie la ligne distante, on
  // resynchronise immédiatement — sans avoir à recharger la page.
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) { setRealtimeConnected(false); return; }
    let unsub: (() => void) | null = null;
    let cancelled = false;
    (async () => {
      const fn = await subscribeRealtime(syncCode, () => runSync());
      if (cancelled) { fn?.(); return; }
      unsub = fn;
      setRealtimeConnected(!!fn);
    })();
    return () => { cancelled = true; unsub?.(); setRealtimeConnected(false); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allLoaded, syncCode]);

  // Filet de sécurité final : si l'app se ferme/passe en arrière-plan alors qu'une
  // synchronisation est encore en attente, tentative d'envoi immédiat "best effort" via
  // fetch+keepalive (fonctionne pendant la fermeture, contrairement à un fetch normal —
  // et contrairement à sendBeacon, supporte les en-têtes d'authentification requis).
  useEffect(() => {
    if (!allLoaded || !SYNC_ENABLED || !syncCode) return;
    const flush = () => {
      if (!editSyncTimer.current) return;
      const s = syncedRef.current;
      try {
        fetch(`${SUPABASE_URL}/rest/v1/app_state?on_conflict=sync_code`, {
          method: "POST",
          keepalive: true,
          headers: {
            apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
            "Content-Type": "application/json", Prefer: "resolution=merge-duplicates",
          },
          body: JSON.stringify({ sync_code: syncCode, updated_at: new Date().toISOString(), data: s }),
        }).catch(() => {});
      } catch {}
    };
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") flush(); });
    window.addEventListener("pagehide", flush);
    window.addEventListener("blur", flush);
    return () => {
      window.removeEventListener("pagehide", flush);
      window.removeEventListener("blur", flush);
    };
  }, [allLoaded, syncCode]);

  // Réutilisée par le Sauvegarde et par l'écran de choix ci-dessous.
  const restoreFromBackup = (data: any) => {
    setPreRestoreSnapshot({
      transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
      activities, categoryActivity, activityCapital, monthlyObjective, chargeOverrides, includeGrundfosVoiture, customDepSubcategories, customRevSubcategories,
    });
    setPreRestoreSnapshotAt(`${dateLabelFull(todayISO())} à ${nowTime()}`);
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
    if (data.activities) setActivities(data.activities);
    if (data.categoryActivity) setCategoryActivity(data.categoryActivity);
    if (data.activityCapital) setActivityCapital(data.activityCapital);
    if (typeof data.monthlyObjective === "number") setMonthlyObjective(data.monthlyObjective);
    if (data.chargeOverrides) setChargeOverrides(data.chargeOverrides);
    if (typeof data.includeGrundfosVoiture === "boolean") setIncludeGrundfosVoiture(data.includeGrundfosVoiture);
    if (data.customDepSubcategories) setCustomDepSubcategories(data.customDepSubcategories);
    if (data.customRevSubcategories) setCustomRevSubcategories(data.customRevSubcategories);
  };

  if (!allLoaded) {
    return <div style={{ minHeight: "100vh", background: COLOR.bg, color: COLOR.inkMuted, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>Chargement…</div>;
  }

  // Écran de choix bloquant : rien n'est jamais sauvegardé tant qu'un choix explicite
  // n'a pas été fait ici. Corrige la cause de fond d'une perte de données passée : avant,
  // dès que le stockage local était vide, l'app se mettait à sauvegarder silencieusement
  // les données de démonstration comme si c'était les vraies. Sur demande explicite de
  // l'utilisateur (11/08/2026) : "résoudre ça une fois pour de bon".
  if (txStartedEmpty && !dataGateResolved) {
    return (
      <DataRecoveryGate
        onRestore={(data) => { restoreFromBackup(data); setDataGateResolved(true); }}
        onConnectSync={async (code) => {
          setSyncCode(code);
          const remote = await fetchRemoteState(code);
          if (remote?.data?.transactions?.length) {
            restoreFromBackup(remote.data);
            setDataGateResolved(true);
            return true;
          }
          return false;
        }}
        onStartFresh={() => setDataGateResolved(true)}
      />
    );
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
          <GlobalReminderBanner transactions={transactions} dismissedDate={dismissedReminderDate} setDismissedDate={setDismissedReminderDate} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} />
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
                      <FilterBar filters={filters} setFilters={setFilters} allMonths={allMonths} allCategories={allCategories} categoryOptions={groupedCatOptions} allAccounts={accounts.map((a) => a.name)} onReset={() => setFilters(defaultFilters)} />
                    </div>
                  )}
                </div>
              ) : (
                <FilterBar filters={filters} setFilters={setFilters} allMonths={allMonths} allCategories={allCategories} categoryOptions={groupedCatOptions} allAccounts={accounts.map((a) => a.name)} onReset={() => setFilters(defaultFilters)} />
              )}
            </div>
          )}

          {tab === "apercu" && <ApercuTab filtered={filtered} filters={filters} accounts={accounts} transactions={transactions} categoryGroups={resolvedGroups} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} monthlyObjective={monthlyObjective} />}
          {tab === "valeurnette" && <NetWorthTab accounts={accounts} transactions={transactions} filters={filters} />}
          {tab === "flux" && <FluxTab filtered={filtered} />}
          {tab === "comparatif" && <ComparatifTab transactions={transactions} categoryGroups={resolvedGroups} />}
          {tab === "comparateur" && <ComparateurTab transactions={transactions} categoryGroups={resolvedGroups} allMonths={allMonths} />}
          {tab === "topcategories" && <TopCategoriesTab transactions={transactions} setTransactions={setTransactionsTracked} categoryGroups={resolvedGroups} allMonths={allMonths} accounts={accounts} onNavigate={navigateTo} />}
          {tab === "categoryoverview" && <CategoryOverviewTab transactions={transactions} categoryGroups={resolvedGroups} allMonths={allMonths} navContext={navContext} />}
          {tab === "saisie" && <SaisieQuotidienneTab transactions={transactions} setTransactions={setTransactionsTracked} allCategories={allCategories} categoryGroups={resolvedGroups} accounts={accounts} monthlyObjective={monthlyObjective} setMonthlyObjective={setMonthlyObjectiveLogged} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} />}
          {tab === "mensuel" && <MensuelTab filtered={filtered} />}
          {tab === "journalier" && <JournalierTab filtered={filtered} />}
          {tab === "categories" && <CategoriesTab filtered={filtered} categoryGroups={categoryGroups} resolvedGroups={resolvedGroups} setCategoryGroups={setCategoryGroups} />}
          {tab === "gestioncategories" && <CategoryManagementTab
            transactions={transactions} setTransactions={setTransactionsTracked}
            customDepSubcategories={customDepSubcategories} setCustomDepSubcategories={setCustomDepSubcategories}
            customRevSubcategories={customRevSubcategories} setCustomRevSubcategories={setCustomRevSubcategories}
            categoryGroups={categoryGroups} setCategoryGroups={setCategoryGroups}
            categoryScope={categoryScope} setCategoryScope={setCategoryScopeLogged}
            categoryActivity={categoryActivity} setCategoryActivity={setCategoryActivityLogged}
            budgets={budgets} setBudgets={setBudgets}
          />}
          {tab === "groupes" && <GroupesTab filtered={filtered} />}
          {tab === "enveloppes" && <EnveloppesTab filtered={filtered} cap={envelopeCap} setCap={setEnvelopeCap} />}
          {tab === "budgets" && <BudgetsTab transactions={transactions} categoryGroups={resolvedGroups} budgets={budgets} setBudgets={setBudgets} allCategories={allCategories} />}
          {tab === "simulateur" && <SimulateurTab filtered={filtered} accounts={accounts} transactions={transactions} />}
          {tab === "objectif" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              <GoalsPanel goals={goals} setGoals={setGoals} accounts={accounts} transactions={transactions} />
              <ProjectionPanel accounts={accounts} transactions={transactions} />
              <CustomProjectionPanel transactions={transactions} accounts={accounts} allCategories={allCategories} />
            </div>
          )}
          {tab === "business" && <BusinessTab transactions={transactions} categoryGroups={resolvedGroups} categoryScope={categoryScope} setCategoryScope={setCategoryScopeLogged} allCategories={allCategories} />}
          {tab === "activites" && <ActivitiesTab transactions={transactions} setTransactions={setTransactionsTracked} activities={activities} setActivities={setActivitiesLogged} categoryActivity={categoryActivity} setCategoryActivity={setCategoryActivityLogged} activityCapital={activityCapital} setActivityCapital={setActivityCapitalLogged} allCategories={allCategories} categoryGroups={resolvedGroups} accounts={accounts} onNavigate={navigateTo} periodRange={[filters.from, filters.to]} />}
          {tab === "charges" && <ChargesTab transactions={transactions} chargeOverrides={chargeOverrides} setChargeOverrides={setChargeOverridesLogged} includeGrundfosVoiture={includeGrundfosVoiture} setIncludeGrundfosVoiture={setIncludeGrundfosVoitureLogged} onNavigate={navigateTo} periodRange={[filters.from, filters.to]} />}
          {tab === "diagnostic" && <DiagnosticTab transactions={transactions} accounts={accounts} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} setIncludeGrundfosVoiture={setIncludeGrundfosVoitureLogged} onNavigate={navigateTo} periodRange={[filters.from, filters.to]} />}
          {tab === "rapprochement" && <RapprochementTab transactions={transactions} setTransactions={setTransactionsTracked} accounts={accounts} />}
          {tab === "creances" && <CreancesTab loans={loans} setLoans={setLoans} />}
          {tab === "comptes" && <ComptesTab accounts={accounts} setAccounts={setAccounts} transactions={transactions} setTransactions={setTransactionsTracked} />}
          {tab === "payees" && <PayeesTab transactions={transactions} />}
          {tab === "recurrences" && <RecurrencesTab recurring={recurring} setRecurring={setRecurring} transactions={transactions} setTransactions={setTransactionsTracked} allCategories={allCategories} accounts={accounts} chargeOverrides={chargeOverrides} includeGrundfosVoiture={includeGrundfosVoiture} />}
          {tab === "journal" && <JournalTab filtered={filtered} allCategories={allCategories} categoryGroups={resolvedGroups} transactions={transactions} setTransactions={setTransactionsTracked} rules={rules} setRules={setRules} accounts={accounts} />}
          {tab === "export" && <ExportTab filtered={filtered} filters={filters} setFilters={setFilters} allMonths={allMonths} />}
          {tab === "sauvegarde" && (
            <SauvegardeTab
              getSnapshot={() => ({
                transactions, categoryGroups, categoryScope, rules, loans, envelopeCap, accounts, budgets, goals, recurring,
                activities, categoryActivity, activityCapital, monthlyObjective, chargeOverrides, includeGrundfosVoiture, customDepSubcategories, customRevSubcategories,
              })}
              restore={restoreFromBackup}
              undoSnapshotAt={preRestoreSnapshotAt}
              settingsLog={settingsLog}
              onUndoRestore={() => {
                if (!preRestoreSnapshot) return;
                const data = preRestoreSnapshot;
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
                if (data.activities) setActivities(data.activities);
                if (data.categoryActivity) setCategoryActivity(data.categoryActivity);
                if (data.activityCapital) setActivityCapital(data.activityCapital);
                if (typeof data.monthlyObjective === "number") setMonthlyObjective(data.monthlyObjective);
                if (data.chargeOverrides) setChargeOverrides(data.chargeOverrides);
                if (typeof data.includeGrundfosVoiture === "boolean") setIncludeGrundfosVoiture(data.includeGrundfosVoiture);
                if (data.customDepSubcategories) setCustomDepSubcategories(data.customDepSubcategories);
                if (data.customRevSubcategories) setCustomRevSubcategories(data.customRevSubcategories);
                setPreRestoreSnapshot(null);
                setPreRestoreSnapshotAt(null);
              }}
              syncCode={syncCode}
              setSyncCode={setSyncCode}
              syncStatus={syncStatus}
              lastSyncedAt={lastSyncedAt}
              realtimeConnected={realtimeConnected}
              onForceSync={() => { runSync(); }}
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
        <QuickAddFAB transactions={transactions} setTransactions={setTransactionsTracked} accounts={accounts} categoryGroups={resolvedGroups} isMobile={isMobile} />
      )}
    </div>
  );
}
