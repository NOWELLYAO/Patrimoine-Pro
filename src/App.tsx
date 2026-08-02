import React, { useState, useEffect, useMemo } from "react";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, ReferenceLine,
} from "recharts";
import {
  LayoutDashboard, CalendarRange, PiggyBank, Layers, BookOpen, TrendingUp,
  TrendingDown, Filter, X, Plus, Pencil, Trash2, Save, RotateCcw, Search,
  ArrowUpDown, Wallet, Target, AlertTriangle, Info, Check, Circle, ChevronRight,
  SlidersHorizontal, Workflow, CalendarDays, BarChart3, Briefcase, HandCoins, Clock,
  Users, Repeat, ClipboardList, UploadCloud, CheckSquare, Square,
  Download, Printer, Bell, Sparkles, Gauge, ArrowRight, Percent, Upload, Mail,
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
`;

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
const seedTransactions: Transaction[] = [{"category":"GRUNDFOS","type":"Dépense","amount":494562,"id":"t1","date":"2024-06-01"},{"category":"Voiture","type":"Dépense","amount":4000,"id":"t2","date":"2024-07-01"},{"category":"Cadeaux","type":"Dépense","amount":127700,"id":"t3","date":"2024-07-01"},{"category":"Santé","type":"Dépense","amount":912,"id":"t4","date":"2024-07-01"},{"category":"Générales","type":"Dépense","amount":2000,"id":"t5","date":"2024-07-01"},{"category":"Divertissement","type":"Dépense","amount":34500,"id":"t6","date":"2024-07-01"},{"category":"Utilitaires","type":"Dépense","amount":70500,"id":"t7","date":"2024-07-01"},{"category":"Aliments","type":"Dépense","amount":80000,"id":"t8","date":"2024-07-01"},{"category":"Personnel","type":"Dépense","amount":2000,"id":"t9","date":"2024-07-01"},{"category":"Shopping","type":"Dépense","amount":58657,"id":"t10","date":"2024-07-01"},{"category":"Enfants & Maman","type":"Dépense","amount":30300,"id":"t11","date":"2024-07-01"},{"category":"Épargne","type":"Dépense","amount":147351,"id":"t12","date":"2024-07-01"},{"category":"Invitation","type":"Dépense","amount":179000,"id":"t13","date":"2024-07-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t14","date":"2024-07-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t15","date":"2024-07-01"},{"category":"GRUNDFOS","type":"Dépense","amount":93200,"id":"t16","date":"2024-07-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t17","date":"2024-07-01"},{"category":"Logement","type":"Dépense","amount":550550,"id":"t18","date":"2024-08-01"},{"category":"Voiture","type":"Dépense","amount":10000,"id":"t19","date":"2024-08-01"},{"category":"Vêtements","type":"Dépense","amount":47000,"id":"t20","date":"2024-08-01"},{"category":"Cadeaux","type":"Dépense","amount":293800,"id":"t21","date":"2024-08-01"},{"category":"Santé","type":"Dépense","amount":22500,"id":"t22","date":"2024-08-01"},{"category":"Dette","type":"Dépense","amount":797415,"id":"t23","date":"2024-08-01"},{"category":"Générales","type":"Dépense","amount":8874,"id":"t24","date":"2024-08-01"},{"category":"Divertissement","type":"Dépense","amount":360600,"id":"t25","date":"2024-08-01"},{"category":"Aliments","type":"Dépense","amount":359740,"id":"t26","date":"2024-08-01"},{"category":"Personnel","type":"Dépense","amount":2000,"id":"t27","date":"2024-08-01"},{"category":"Abonnements","type":"Dépense","amount":12300,"id":"t28","date":"2024-08-01"},{"category":"Shopping","type":"Dépense","amount":41000,"id":"t29","date":"2024-08-01"},{"category":"Enfants & Maman","type":"Dépense","amount":57200,"id":"t30","date":"2024-08-01"},{"category":"Invitation","type":"Dépense","amount":127915,"id":"t31","date":"2024-08-01"},{"category":"Ajustement","type":"Dépense","amount":32542,"id":"t32","date":"2024-08-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t33","date":"2024-08-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t34","date":"2024-08-01"},{"category":"GRUNDFOS","type":"Dépense","amount":195000,"id":"t35","date":"2024-08-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t36","date":"2024-08-01"},{"category":"Logement","type":"Dépense","amount":550550,"id":"t37","date":"2024-09-01"},{"category":"Voiture","type":"Dépense","amount":176500,"id":"t38","date":"2024-09-01"},{"category":"Transport","type":"Dépense","amount":22000,"id":"t39","date":"2024-09-01"},{"category":"Éducation","type":"Dépense","amount":270000,"id":"t40","date":"2024-09-01"},{"category":"Cadeaux","type":"Dépense","amount":163550,"id":"t41","date":"2024-09-01"},{"category":"Santé","type":"Dépense","amount":500,"id":"t42","date":"2024-09-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t43","date":"2024-09-01"},{"category":"Générales","type":"Dépense","amount":12000,"id":"t44","date":"2024-09-01"},{"category":"Divertissement","type":"Dépense","amount":325500,"id":"t45","date":"2024-09-01"},{"category":"Utilitaires","type":"Dépense","amount":24500,"id":"t46","date":"2024-09-01"},{"category":"Aliments","type":"Dépense","amount":267231,"id":"t47","date":"2024-09-01"},{"category":"Shopping","type":"Dépense","amount":62500,"id":"t48","date":"2024-09-01"},{"category":"Enfants & Maman","type":"Dépense","amount":125300,"id":"t49","date":"2024-09-01"},{"category":"Invitation","type":"Dépense","amount":50000,"id":"t50","date":"2024-09-01"},{"category":"Ajustement","type":"Dépense","amount":20787,"id":"t51","date":"2024-09-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t52","date":"2024-09-01"},{"category":"Plan Éducation","type":"Dépense","amount":2457,"id":"t53","date":"2024-09-01"},{"category":"Payement Maison Bingerville","type":"Dépense","amount":3000000,"id":"t54","date":"2024-09-01"},{"category":"GRUNDFOS","type":"Dépense","amount":255954,"id":"t55","date":"2024-09-01"},{"category":"Achat Mazda","type":"Dépense","amount":6851100,"id":"t56","date":"2024-09-01"},{"category":"Logement","type":"Dépense","amount":551500,"id":"t57","date":"2024-10-01"},{"category":"Voiture","type":"Dépense","amount":105500,"id":"t58","date":"2024-10-01"},{"category":"Transport","type":"Dépense","amount":84360,"id":"t59","date":"2024-10-01"},{"category":"Cadeaux","type":"Dépense","amount":464080,"id":"t60","date":"2024-10-01"},{"category":"Santé","type":"Dépense","amount":13500,"id":"t61","date":"2024-10-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t62","date":"2024-10-01"},{"category":"Générales","type":"Dépense","amount":31404,"id":"t63","date":"2024-10-01"},{"category":"Divertissement","type":"Dépense","amount":340075,"id":"t64","date":"2024-10-01"},{"category":"Utilitaires","type":"Dépense","amount":33500,"id":"t65","date":"2024-10-01"},{"category":"Aliments","type":"Dépense","amount":124170,"id":"t66","date":"2024-10-01"},{"category":"Personnel","type":"Dépense","amount":6000,"id":"t67","date":"2024-10-01"},{"category":"Shopping","type":"Dépense","amount":140500,"id":"t68","date":"2024-10-01"},{"category":"Enfants & Maman","type":"Dépense","amount":118435,"id":"t69","date":"2024-10-01"},{"category":"Invitation","type":"Dépense","amount":125500,"id":"t70","date":"2024-10-01"},{"category":"Ajustement","type":"Dépense","amount":41475,"id":"t71","date":"2024-10-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t72","date":"2024-10-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t73","date":"2024-10-01"},{"category":"GRUNDFOS","type":"Dépense","amount":226230,"id":"t74","date":"2024-10-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t75","date":"2024-10-01"},{"category":"Achat Mazda","type":"Dépense","amount":456000,"id":"t76","date":"2024-10-01"},{"category":"Logement","type":"Dépense","amount":560000,"id":"t77","date":"2024-11-01"},{"category":"Voiture","type":"Dépense","amount":17500,"id":"t78","date":"2024-11-01"},{"category":"Vêtements","type":"Dépense","amount":86020,"id":"t79","date":"2024-11-01"},{"category":"Transport","type":"Dépense","amount":1000,"id":"t80","date":"2024-11-01"},{"category":"Cadeaux","type":"Dépense","amount":188350,"id":"t81","date":"2024-11-01"},{"category":"Santé","type":"Dépense","amount":14500,"id":"t82","date":"2024-11-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t83","date":"2024-11-01"},{"category":"Générales","type":"Dépense","amount":6000,"id":"t84","date":"2024-11-01"},{"category":"Divertissement","type":"Dépense","amount":298750,"id":"t85","date":"2024-11-01"},{"category":"Utilitaires","type":"Dépense","amount":31500,"id":"t86","date":"2024-11-01"},{"category":"Aliments","type":"Dépense","amount":212500,"id":"t87","date":"2024-11-01"},{"category":"Personnel","type":"Dépense","amount":212500,"id":"t88","date":"2024-11-01"},{"category":"Shopping","type":"Dépense","amount":96000,"id":"t89","date":"2024-11-01"},{"category":"Enfants & Maman","type":"Dépense","amount":125600,"id":"t90","date":"2024-11-01"},{"category":"Invitation","type":"Dépense","amount":240500,"id":"t91","date":"2024-11-01"},{"category":"Ajustement","type":"Dépense","amount":68147,"id":"t92","date":"2024-11-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":70000,"id":"t93","date":"2024-11-01"},{"category":"GRUNDFOS","type":"Dépense","amount":178500,"id":"t94","date":"2024-11-01"},{"category":"Pack Club","type":"Dépense","amount":10000,"id":"t95","date":"2024-11-01"},{"category":"Logement","type":"Dépense","amount":550500,"id":"t96","date":"2024-12-01"},{"category":"Voiture","type":"Dépense","amount":15100,"id":"t97","date":"2024-12-01"},{"category":"Vêtements","type":"Dépense","amount":54500,"id":"t98","date":"2024-12-01"},{"category":"Cadeaux","type":"Dépense","amount":517020,"id":"t99","date":"2024-12-01"},{"category":"Santé","type":"Dépense","amount":30000,"id":"t100","date":"2024-12-01"},{"category":"Dette","type":"Dépense","amount":740573,"id":"t101","date":"2024-12-01"},{"category":"Générales","type":"Dépense","amount":29200,"id":"t102","date":"2024-12-01"},{"category":"Divertissement","type":"Dépense","amount":340000,"id":"t103","date":"2024-12-01"},{"category":"Utilitaires","type":"Dépense","amount":20000,"id":"t104","date":"2024-12-01"},{"category":"Aliments","type":"Dépense","amount":479300,"id":"t105","date":"2024-12-01"},{"category":"Shopping","type":"Dépense","amount":203000,"id":"t106","date":"2024-12-01"},{"category":"Enfants & Maman","type":"Dépense","amount":191800,"id":"t107","date":"2024-12-01"},{"category":"Invitation","type":"Dépense","amount":21000,"id":"t108","date":"2024-12-01"},{"category":"Ajustement","type":"Dépense","amount":90161,"id":"t109","date":"2024-12-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t110","date":"2024-12-01"},{"category":"Plan Éducation","type":"Dépense","amount":64914,"id":"t111","date":"2024-12-01"},{"category":"GRUNDFOS","type":"Dépense","amount":320000,"id":"t112","date":"2024-12-01"},{"category":"Logement","type":"Dépense","amount":551000,"id":"t113","date":"2025-01-01"},{"category":"Voiture","type":"Dépense","amount":15000,"id":"t114","date":"2025-01-01"},{"category":"Des sports","type":"Dépense","amount":74000,"id":"t115","date":"2025-01-01"},{"category":"Vêtements","type":"Dépense","amount":182000,"id":"t116","date":"2025-01-01"},{"category":"Voyage","type":"Dépense","amount":30000,"id":"t117","date":"2025-01-01"},{"category":"Cadeaux","type":"Dépense","amount":638982,"id":"t118","date":"2025-01-01"},{"category":"Santé","type":"Dépense","amount":7500,"id":"t119","date":"2025-01-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t120","date":"2025-01-01"},{"category":"Générales","type":"Dépense","amount":201800,"id":"t121","date":"2025-01-01"},{"category":"Divertissement","type":"Dépense","amount":479500,"id":"t122","date":"2025-01-01"},{"category":"Utilitaires","type":"Dépense","amount":23000,"id":"t123","date":"2025-01-01"},{"category":"Aliments","type":"Dépense","amount":365100,"id":"t124","date":"2025-01-01"},{"category":"Personnel","type":"Dépense","amount":9500,"id":"t125","date":"2025-01-01"},{"category":"Abonnements","type":"Dépense","amount":3877,"id":"t126","date":"2025-01-01"},{"category":"Shopping","type":"Dépense","amount":166200,"id":"t127","date":"2025-01-01"},{"category":"Enfants & Maman","type":"Dépense","amount":134500,"id":"t128","date":"2025-01-01"},{"category":"Épargne","type":"Dépense","amount":310000,"id":"t129","date":"2025-01-01"},{"category":"Invitation","type":"Dépense","amount":70000,"id":"t130","date":"2025-01-01"},{"category":"Ajustement","type":"Dépense","amount":68682,"id":"t131","date":"2025-01-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t132","date":"2025-01-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t133","date":"2025-01-01"},{"category":"GRUNDFOS","type":"Dépense","amount":243200,"id":"t134","date":"2025-01-01"},{"category":"Logement","type":"Dépense","amount":550500,"id":"t135","date":"2025-02-01"},{"category":"Voiture","type":"Dépense","amount":54000,"id":"t136","date":"2025-02-01"},{"category":"Vêtements","type":"Dépense","amount":22000,"id":"t137","date":"2025-02-01"},{"category":"Cadeaux","type":"Dépense","amount":405850,"id":"t138","date":"2025-02-01"},{"category":"Santé","type":"Dépense","amount":10000,"id":"t139","date":"2025-02-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t140","date":"2025-02-01"},{"category":"Générales","type":"Dépense","amount":61570,"id":"t141","date":"2025-02-01"},{"category":"Divertissement","type":"Dépense","amount":388000,"id":"t142","date":"2025-02-01"},{"category":"Utilitaires","type":"Dépense","amount":3000,"id":"t143","date":"2025-02-01"},{"category":"Aliments","type":"Dépense","amount":249500,"id":"t144","date":"2025-02-01"},{"category":"Personnel","type":"Dépense","amount":4000,"id":"t145","date":"2025-02-01"},{"category":"Abonnements","type":"Dépense","amount":3000,"id":"t146","date":"2025-02-01"},{"category":"Shopping","type":"Dépense","amount":65500,"id":"t147","date":"2025-02-01"},{"category":"Enfants & Maman","type":"Dépense","amount":136600,"id":"t148","date":"2025-02-01"},{"category":"Épargne","type":"Dépense","amount":280000,"id":"t149","date":"2025-02-01"},{"category":"Invitation","type":"Dépense","amount":25000,"id":"t150","date":"2025-02-01"},{"category":"Ajustement","type":"Dépense","amount":57258,"id":"t151","date":"2025-02-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t152","date":"2025-02-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t153","date":"2025-02-01"},{"category":"GRUNDFOS","type":"Dépense","amount":316500,"id":"t154","date":"2025-02-01"},{"category":"Logement","type":"Dépense","amount":550500,"id":"t155","date":"2025-03-01"},{"category":"Voiture","type":"Dépense","amount":27000,"id":"t156","date":"2025-03-01"},{"category":"Vêtements","type":"Dépense","amount":15000,"id":"t157","date":"2025-03-01"},{"category":"Transport","type":"Dépense","amount":5200,"id":"t158","date":"2025-03-01"},{"category":"Cadeaux","type":"Dépense","amount":294150,"id":"t159","date":"2025-03-01"},{"category":"Santé","type":"Dépense","amount":13000,"id":"t160","date":"2025-03-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t161","date":"2025-03-01"},{"category":"Générales","type":"Dépense","amount":19030,"id":"t162","date":"2025-03-01"},{"category":"Divertissement","type":"Dépense","amount":556000,"id":"t163","date":"2025-03-01"},{"category":"Utilitaires","type":"Dépense","amount":59000,"id":"t164","date":"2025-03-01"},{"category":"Aliments","type":"Dépense","amount":224500,"id":"t165","date":"2025-03-01"},{"category":"Personnel","type":"Dépense","amount":5000,"id":"t166","date":"2025-03-01"},{"category":"Abonnements","type":"Dépense","amount":4000,"id":"t167","date":"2025-03-01"},{"category":"Shopping","type":"Dépense","amount":185000,"id":"t168","date":"2025-03-01"},{"category":"Enfants & Maman","type":"Dépense","amount":66300,"id":"t169","date":"2025-03-01"},{"category":"Épargne","type":"Dépense","amount":310000,"id":"t170","date":"2025-03-01"},{"category":"Ajustement","type":"Dépense","amount":64699,"id":"t171","date":"2025-03-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t172","date":"2025-03-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t173","date":"2025-03-01"},{"category":"Payement Maison Bingerville","type":"Dépense","amount":5000000,"id":"t174","date":"2025-03-01"},{"category":"GRUNDFOS","type":"Dépense","amount":260204,"id":"t175","date":"2025-03-01"},{"category":"Pack Club","type":"Dépense","amount":10000,"id":"t176","date":"2025-03-01"},{"category":"Logement","type":"Dépense","amount":550550,"id":"t177","date":"2025-04-01"},{"category":"Voiture","type":"Dépense","amount":85510,"id":"t178","date":"2025-04-01"},{"category":"Vêtements","type":"Dépense","amount":60000,"id":"t179","date":"2025-04-01"},{"category":"Transport","type":"Dépense","amount":3200,"id":"t180","date":"2025-04-01"},{"category":"Cadeaux","type":"Dépense","amount":402250,"id":"t181","date":"2025-04-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t182","date":"2025-04-01"},{"category":"Générales","type":"Dépense","amount":85115,"id":"t183","date":"2025-04-01"},{"category":"Divertissement","type":"Dépense","amount":379880,"id":"t184","date":"2025-04-01"},{"category":"Utilitaires","type":"Dépense","amount":5500,"id":"t185","date":"2025-04-01"},{"category":"Aliments","type":"Dépense","amount":219400,"id":"t186","date":"2025-04-01"},{"category":"Personnel","type":"Dépense","amount":5000,"id":"t187","date":"2025-04-01"},{"category":"Abonnements","type":"Dépense","amount":10000,"id":"t188","date":"2025-04-01"},{"category":"Shopping","type":"Dépense","amount":110834,"id":"t189","date":"2025-04-01"},{"category":"Enfants & Maman","type":"Dépense","amount":156200,"id":"t190","date":"2025-04-01"},{"category":"Invitation","type":"Dépense","amount":77000,"id":"t191","date":"2025-04-01"},{"category":"Ajustement","type":"Dépense","amount":17484,"id":"t192","date":"2025-04-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t193","date":"2025-04-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t194","date":"2025-04-01"},{"category":"GRUNDFOS","type":"Dépense","amount":929495,"id":"t195","date":"2025-04-01"},{"category":"General","type":"Dépense","amount":2682680,"id":"t196","date":"2025-04-01"},{"category":"Formation","type":"Dépense","amount":5171,"id":"t197","date":"2025-04-01"},{"category":"Logement","type":"Dépense","amount":551550,"id":"t198","date":"2025-05-01"},{"category":"Voiture","type":"Dépense","amount":15000,"id":"t199","date":"2025-05-01"},{"category":"Transport","type":"Dépense","amount":3000,"id":"t200","date":"2025-05-01"},{"category":"Cadeaux","type":"Dépense","amount":342465,"id":"t201","date":"2025-05-01"},{"category":"Santé","type":"Dépense","amount":24510,"id":"t202","date":"2025-05-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t203","date":"2025-05-01"},{"category":"Générales","type":"Dépense","amount":12000,"id":"t204","date":"2025-05-01"},{"category":"Divertissement","type":"Dépense","amount":270300,"id":"t205","date":"2025-05-01"},{"category":"Utilitaires","type":"Dépense","amount":2000,"id":"t206","date":"2025-05-01"},{"category":"Aliments","type":"Dépense","amount":185169,"id":"t207","date":"2025-05-01"},{"category":"Personnel","type":"Dépense","amount":349000,"id":"t208","date":"2025-05-01"},{"category":"Abonnements","type":"Dépense","amount":4000,"id":"t209","date":"2025-05-01"},{"category":"Shopping","type":"Dépense","amount":104550,"id":"t210","date":"2025-05-01"},{"category":"Enfants & Maman","type":"Dépense","amount":114010,"id":"t211","date":"2025-05-01"},{"category":"Invitation","type":"Dépense","amount":403000,"id":"t212","date":"2025-05-01"},{"category":"Ajustement","type":"Dépense","amount":41630,"id":"t213","date":"2025-05-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t214","date":"2025-05-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t215","date":"2025-05-01"},{"category":"GRUNDFOS","type":"Dépense","amount":183500,"id":"t216","date":"2025-05-01"},{"category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":600000,"id":"t217","date":"2025-05-01"},{"category":"Logement","type":"Dépense","amount":551550,"id":"t218","date":"2025-06-01"},{"category":"Voiture","type":"Dépense","amount":116500,"id":"t219","date":"2025-06-01"},{"category":"Cadeaux","type":"Dépense","amount":629060,"id":"t220","date":"2025-06-01"},{"category":"Santé","type":"Dépense","amount":8500,"id":"t221","date":"2025-06-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t222","date":"2025-06-01"},{"category":"Générales","type":"Dépense","amount":5090,"id":"t223","date":"2025-06-01"},{"category":"Divertissement","type":"Dépense","amount":306780,"id":"t224","date":"2025-06-01"},{"category":"Utilitaires","type":"Dépense","amount":2500,"id":"t225","date":"2025-06-01"},{"category":"Aliments","type":"Dépense","amount":119500,"id":"t226","date":"2025-06-01"},{"category":"Personnel","type":"Dépense","amount":2000,"id":"t227","date":"2025-06-01"},{"category":"Abonnements","type":"Dépense","amount":66866,"id":"t228","date":"2025-06-01"},{"category":"Shopping","type":"Dépense","amount":105400,"id":"t229","date":"2025-06-01"},{"category":"Enfants & Maman","type":"Dépense","amount":105600,"id":"t230","date":"2025-06-01"},{"category":"Invitation","type":"Dépense","amount":181000,"id":"t231","date":"2025-06-01"},{"category":"Ajustement","type":"Dépense","amount":44359,"id":"t232","date":"2025-06-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t233","date":"2025-06-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t234","date":"2025-06-01"},{"category":"GRUNDFOS","type":"Dépense","amount":400793,"id":"t235","date":"2025-06-01"},{"category":"Formation","type":"Dépense","amount":5000,"id":"t236","date":"2025-06-01"},{"category":"INVEST SGO","type":"Dépense","amount":57908,"id":"t237","date":"2025-06-01"},{"category":"Achat Terrain Port Bouet","type":"Dépense","amount":121200,"id":"t238","date":"2025-06-01"},{"category":"Logement","type":"Dépense","amount":600550,"id":"t239","date":"2025-07-01"},{"category":"Voiture","type":"Dépense","amount":58520,"id":"t240","date":"2025-07-01"},{"category":"Cadeaux","type":"Dépense","amount":163894,"id":"t241","date":"2025-07-01"},{"category":"Santé","type":"Dépense","amount":5500,"id":"t242","date":"2025-07-01"},{"category":"Dette","type":"Dépense","amount":631166,"id":"t243","date":"2025-07-01"},{"category":"Générales","type":"Dépense","amount":2000,"id":"t244","date":"2025-07-01"},{"category":"Divertissement","type":"Dépense","amount":258828,"id":"t245","date":"2025-07-01"},{"category":"Aliments","type":"Dépense","amount":118385,"id":"t246","date":"2025-07-01"},{"category":"Personnel","type":"Dépense","amount":5000,"id":"t247","date":"2025-07-01"},{"category":"Abonnements","type":"Dépense","amount":28200,"id":"t248","date":"2025-07-01"},{"category":"Shopping","type":"Dépense","amount":44275,"id":"t249","date":"2025-07-01"},{"category":"Enfants & Maman","type":"Dépense","amount":93820,"id":"t250","date":"2025-07-01"},{"category":"Invitation","type":"Dépense","amount":161500,"id":"t251","date":"2025-07-01"},{"category":"Ajustement","type":"Dépense","amount":17121,"id":"t252","date":"2025-07-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t253","date":"2025-07-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t254","date":"2025-07-01"},{"category":"Payement Maison Bingerville","type":"Dépense","amount":5300000,"id":"t255","date":"2025-07-01"},{"category":"GRUNDFOS","type":"Dépense","amount":135000,"id":"t256","date":"2025-07-01"},{"category":"Formation","type":"Dépense","amount":207500,"id":"t257","date":"2025-07-01"},{"category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":76313,"id":"t258","date":"2025-07-01"},{"category":"Logement","type":"Dépense","amount":576550,"id":"t259","date":"2025-08-01"},{"category":"Voiture","type":"Dépense","amount":12000,"id":"t260","date":"2025-08-01"},{"category":"Vêtements","type":"Dépense","amount":71000,"id":"t261","date":"2025-08-01"},{"category":"Cadeaux","type":"Dépense","amount":382115,"id":"t262","date":"2025-08-01"},{"category":"Santé","type":"Dépense","amount":15500,"id":"t263","date":"2025-08-01"},{"category":"Générales","type":"Dépense","amount":20070,"id":"t264","date":"2025-08-01"},{"category":"Divertissement","type":"Dépense","amount":567100,"id":"t265","date":"2025-08-01"},{"category":"Aliments","type":"Dépense","amount":136000,"id":"t266","date":"2025-08-01"},{"category":"Personnel","type":"Dépense","amount":6000,"id":"t267","date":"2025-08-01"},{"category":"Abonnements","type":"Dépense","amount":6000,"id":"t268","date":"2025-08-01"},{"category":"Shopping","type":"Dépense","amount":155000,"id":"t269","date":"2025-08-01"},{"category":"Enfants & Maman","type":"Dépense","amount":230850,"id":"t270","date":"2025-08-01"},{"category":"Invitation","type":"Dépense","amount":101000,"id":"t271","date":"2025-08-01"},{"category":"Ajustement","type":"Dépense","amount":105210,"id":"t272","date":"2025-08-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t273","date":"2025-08-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t274","date":"2025-08-01"},{"category":"GRUNDFOS","type":"Dépense","amount":688093,"id":"t275","date":"2025-08-01"},{"category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":10100,"id":"t276","date":"2025-08-01"},{"category":"INVEST SGO","type":"Dépense","amount":216951,"id":"t277","date":"2025-08-01"},{"category":"Logement","type":"Dépense","amount":100000,"id":"t278","date":"2025-09-01"},{"category":"Voiture","type":"Dépense","amount":165100,"id":"t279","date":"2025-09-01"},{"category":"Vêtements","type":"Dépense","amount":3000,"id":"t280","date":"2025-09-01"},{"category":"Cadeaux","type":"Dépense","amount":551680,"id":"t281","date":"2025-09-01"},{"category":"Santé","type":"Dépense","amount":8500,"id":"t282","date":"2025-09-01"},{"category":"Générales","type":"Dépense","amount":39050,"id":"t283","date":"2025-09-01"},{"category":"Divertissement","type":"Dépense","amount":316750,"id":"t284","date":"2025-09-01"},{"category":"Utilitaires","type":"Dépense","amount":1000,"id":"t285","date":"2025-09-01"},{"category":"Aliments","type":"Dépense","amount":168000,"id":"t286","date":"2025-09-01"},{"category":"Abonnements","type":"Dépense","amount":12777,"id":"t287","date":"2025-09-01"},{"category":"Shopping","type":"Dépense","amount":45452,"id":"t288","date":"2025-09-01"},{"category":"Enfants & Maman","type":"Dépense","amount":126100,"id":"t289","date":"2025-09-01"},{"category":"Invitation","type":"Dépense","amount":176500,"id":"t290","date":"2025-09-01"},{"category":"Ajustement","type":"Dépense","amount":12116,"id":"t291","date":"2025-09-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t292","date":"2025-09-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t293","date":"2025-09-01"},{"category":"GRUNDFOS","type":"Dépense","amount":484400,"id":"t294","date":"2025-09-01"},{"category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":585650,"id":"t295","date":"2025-09-01"},{"category":"INVEST SGO","type":"Dépense","amount":152712,"id":"t296","date":"2025-09-01"},{"category":"Déménagement","type":"Dépense","amount":3725750,"id":"t297","date":"2025-09-01"},{"category":"Logement","type":"Dépense","amount":101000,"id":"t298","date":"2025-10-01"},{"category":"Voiture","type":"Dépense","amount":63000,"id":"t299","date":"2025-10-01"},{"category":"Vêtements","type":"Dépense","amount":94530,"id":"t300","date":"2025-10-01"},{"category":"Cadeaux","type":"Dépense","amount":669485,"id":"t301","date":"2025-10-01"},{"category":"Santé","type":"Dépense","amount":8500,"id":"t302","date":"2025-10-01"},{"category":"Générales","type":"Dépense","amount":13300,"id":"t303","date":"2025-10-01"},{"category":"Divertissement","type":"Dépense","amount":526000,"id":"t304","date":"2025-10-01"},{"category":"Utilitaires","type":"Dépense","amount":2000,"id":"t305","date":"2025-10-01"},{"category":"Aliments","type":"Dépense","amount":150600,"id":"t306","date":"2025-10-01"},{"category":"Personnel","type":"Dépense","amount":59615,"id":"t307","date":"2025-10-01"},{"category":"Abonnements","type":"Dépense","amount":13900,"id":"t308","date":"2025-10-01"},{"category":"Shopping","type":"Dépense","amount":170758,"id":"t309","date":"2025-10-01"},{"category":"Enfants & Maman","type":"Dépense","amount":363000,"id":"t310","date":"2025-10-01"},{"category":"Invitation","type":"Dépense","amount":322500,"id":"t311","date":"2025-10-01"},{"category":"Ajustement","type":"Dépense","amount":191539,"id":"t312","date":"2025-10-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t313","date":"2025-10-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t314","date":"2025-10-01"},{"category":"GRUNDFOS","type":"Dépense","amount":148000,"id":"t315","date":"2025-10-01"},{"category":"Déménagement","type":"Dépense","amount":236000,"id":"t316","date":"2025-10-01"},{"category":"Logement","type":"Dépense","amount":480000,"id":"t317","date":"2025-11-01"},{"category":"Voiture","type":"Dépense","amount":254000,"id":"t318","date":"2025-11-01"},{"category":"Vêtements","type":"Dépense","amount":31000,"id":"t319","date":"2025-11-01"},{"category":"Voyage","type":"Dépense","amount":65000,"id":"t320","date":"2025-11-01"},{"category":"Cadeaux","type":"Dépense","amount":490380,"id":"t321","date":"2025-11-01"},{"category":"Santé","type":"Dépense","amount":3230,"id":"t322","date":"2025-11-01"},{"category":"Générales","type":"Dépense","amount":85100,"id":"t323","date":"2025-11-01"},{"category":"Divertissement","type":"Dépense","amount":365000,"id":"t324","date":"2025-11-01"},{"category":"Aliments","type":"Dépense","amount":101640,"id":"t325","date":"2025-11-01"},{"category":"Personnel","type":"Dépense","amount":6000,"id":"t326","date":"2025-11-01"},{"category":"Abonnements","type":"Dépense","amount":8502,"id":"t327","date":"2025-11-01"},{"category":"Shopping","type":"Dépense","amount":235600,"id":"t328","date":"2025-11-01"},{"category":"Enfants & Maman","type":"Dépense","amount":123100,"id":"t329","date":"2025-11-01"},{"category":"Invitation","type":"Dépense","amount":141500,"id":"t330","date":"2025-11-01"},{"category":"Ajustement","type":"Dépense","amount":64880,"id":"t331","date":"2025-11-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t332","date":"2025-11-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t333","date":"2025-11-01"},{"category":"GRUNDFOS","type":"Dépense","amount":1120125,"id":"t334","date":"2025-11-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t335","date":"2025-11-01"},{"category":"Securicompte","type":"Dépense","amount":110319,"id":"t336","date":"2025-11-01"},{"category":"Logement","type":"Dépense","amount":481000,"id":"t337","date":"2025-12-01"},{"category":"Voiture","type":"Dépense","amount":97000,"id":"t338","date":"2025-12-01"},{"category":"Vêtements","type":"Dépense","amount":22500,"id":"t339","date":"2025-12-01"},{"category":"Voyage","type":"Dépense","amount":426660,"id":"t340","date":"2025-12-01"},{"category":"Cadeaux","type":"Dépense","amount":1194770,"id":"t341","date":"2025-12-01"},{"category":"Santé","type":"Dépense","amount":34610,"id":"t342","date":"2025-12-01"},{"category":"Générales","type":"Dépense","amount":20400,"id":"t343","date":"2025-12-01"},{"category":"Divertissement","type":"Dépense","amount":574835,"id":"t344","date":"2025-12-01"},{"category":"Aliments","type":"Dépense","amount":198805,"id":"t345","date":"2025-12-01"},{"category":"Personnel","type":"Dépense","amount":379515,"id":"t346","date":"2025-12-01"},{"category":"Abonnements","type":"Dépense","amount":31385,"id":"t347","date":"2025-12-01"},{"category":"Shopping","type":"Dépense","amount":248804,"id":"t348","date":"2025-12-01"},{"category":"Enfants & Maman","type":"Dépense","amount":355800,"id":"t349","date":"2025-12-01"},{"category":"Invitation","type":"Dépense","amount":221000,"id":"t350","date":"2025-12-01"},{"category":"Ajustement","type":"Dépense","amount":63348,"id":"t351","date":"2025-12-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t352","date":"2025-12-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t353","date":"2025-12-01"},{"category":"Payement Maison Bingerville","type":"Dépense","amount":419000,"id":"t354","date":"2025-12-01"},{"category":"GRUNDFOS","type":"Dépense","amount":626445,"id":"t355","date":"2025-12-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t356","date":"2025-12-01"},{"category":"Securicompte","type":"Dépense","amount":912,"id":"t357","date":"2025-12-01"},{"category":"Logement","type":"Dépense","amount":480500,"id":"t358","date":"2026-01-01"},{"category":"Voiture","type":"Dépense","amount":62500,"id":"t359","date":"2026-01-01"},{"category":"Voyage","type":"Dépense","amount":28700,"id":"t360","date":"2026-01-01"},{"category":"Cadeaux","type":"Dépense","amount":313300,"id":"t361","date":"2026-01-01"},{"category":"Santé","type":"Dépense","amount":91480,"id":"t362","date":"2026-01-01"},{"category":"Dette","type":"Dépense","amount":12276,"id":"t363","date":"2026-01-01"},{"category":"Générales","type":"Dépense","amount":6500,"id":"t364","date":"2026-01-01"},{"category":"Divertissement","type":"Dépense","amount":324945,"id":"t365","date":"2026-01-01"},{"category":"Utilitaires","type":"Dépense","amount":22000,"id":"t366","date":"2026-01-01"},{"category":"Aliments","type":"Dépense","amount":113500,"id":"t367","date":"2026-01-01"},{"category":"Personnel","type":"Dépense","amount":5000,"id":"t368","date":"2026-01-01"},{"category":"Abonnements","type":"Dépense","amount":7235,"id":"t369","date":"2026-01-01"},{"category":"Shopping","type":"Dépense","amount":486694,"id":"t370","date":"2026-01-01"},{"category":"Enfants & Maman","type":"Dépense","amount":255550,"id":"t371","date":"2026-01-01"},{"category":"Invitation","type":"Dépense","amount":216030,"id":"t372","date":"2026-01-01"},{"category":"Ajustement","type":"Dépense","amount":76910,"id":"t373","date":"2026-01-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t374","date":"2026-01-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t375","date":"2026-01-01"},{"category":"GRUNDFOS","type":"Dépense","amount":366200,"id":"t376","date":"2026-01-01"},{"category":"Pack Club","type":"Dépense","amount":10000,"id":"t377","date":"2026-01-01"},{"category":"Création Entreprise ECO PUMP AFRIK","type":"Dépense","amount":50500,"id":"t378","date":"2026-01-01"},{"category":"Prêt","type":"Dépense","amount":10100,"id":"t379","date":"2026-01-01"},{"category":"Logement","type":"Dépense","amount":481000,"id":"t380","date":"2026-02-01"},{"category":"Voiture","type":"Dépense","amount":169450,"id":"t381","date":"2026-02-01"},{"category":"Vêtements","type":"Dépense","amount":17590,"id":"t382","date":"2026-02-01"},{"category":"Cadeaux","type":"Dépense","amount":628890,"id":"t383","date":"2026-02-01"},{"category":"Générales","type":"Dépense","amount":14800,"id":"t384","date":"2026-02-01"},{"category":"Divertissement","type":"Dépense","amount":271665,"id":"t385","date":"2026-02-01"},{"category":"Utilitaires","type":"Dépense","amount":15000,"id":"t386","date":"2026-02-01"},{"category":"Aliments","type":"Dépense","amount":130430,"id":"t387","date":"2026-02-01"},{"category":"Personnel","type":"Dépense","amount":20170,"id":"t388","date":"2026-02-01"},{"category":"Abonnements","type":"Dépense","amount":7835,"id":"t389","date":"2026-02-01"},{"category":"Shopping","type":"Dépense","amount":81225,"id":"t390","date":"2026-02-01"},{"category":"Enfants & Maman","type":"Dépense","amount":209850,"id":"t391","date":"2026-02-01"},{"category":"Invitation","type":"Dépense","amount":235500,"id":"t392","date":"2026-02-01"},{"category":"Ajustement","type":"Dépense","amount":75944,"id":"t393","date":"2026-02-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t394","date":"2026-02-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t395","date":"2026-02-01"},{"category":"GRUNDFOS","type":"Dépense","amount":954214,"id":"t396","date":"2026-02-01"},{"category":"Pack Club","type":"Dépense","amount":10000,"id":"t397","date":"2026-02-01"},{"category":"INVEST SGO","type":"Dépense","amount":50904,"id":"t398","date":"2026-02-01"},{"category":"Prêt","type":"Dépense","amount":505000,"id":"t399","date":"2026-02-01"},{"category":"Logement","type":"Dépense","amount":480000,"id":"t400","date":"2026-03-01"},{"category":"Voiture","type":"Dépense","amount":10000,"id":"t401","date":"2026-03-01"},{"category":"Vêtements","type":"Dépense","amount":26000,"id":"t402","date":"2026-03-01"},{"category":"Cadeaux","type":"Dépense","amount":380315,"id":"t403","date":"2026-03-01"},{"category":"Santé","type":"Dépense","amount":18400,"id":"t404","date":"2026-03-01"},{"category":"Générales","type":"Dépense","amount":9300,"id":"t405","date":"2026-03-01"},{"category":"Divertissement","type":"Dépense","amount":446245,"id":"t406","date":"2026-03-01"},{"category":"Utilitaires","type":"Dépense","amount":19000,"id":"t407","date":"2026-03-01"},{"category":"Aliments","type":"Dépense","amount":131315,"id":"t408","date":"2026-03-01"},{"category":"Personnel","type":"Dépense","amount":11000,"id":"t409","date":"2026-03-01"},{"category":"Abonnements","type":"Dépense","amount":22123,"id":"t410","date":"2026-03-01"},{"category":"Shopping","type":"Dépense","amount":86803,"id":"t411","date":"2026-03-01"},{"category":"Enfants & Maman","type":"Dépense","amount":184220,"id":"t412","date":"2026-03-01"},{"category":"Invitation","type":"Dépense","amount":241010,"id":"t413","date":"2026-03-01"},{"category":"Ajustement","type":"Dépense","amount":42785,"id":"t414","date":"2026-03-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t415","date":"2026-03-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t416","date":"2026-03-01"},{"category":"GRUNDFOS","type":"Dépense","amount":1265825,"id":"t417","date":"2026-03-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t418","date":"2026-03-01"},{"category":"INVEST SGO","type":"Dépense","amount":3029,"id":"t419","date":"2026-03-01"},{"category":"Déménagement","type":"Dépense","amount":3192430,"id":"t420","date":"2026-03-01"},{"category":"Securicompte","type":"Dépense","amount":912,"id":"t421","date":"2026-03-01"},{"category":"Prêt","type":"Dépense","amount":250000,"id":"t422","date":"2026-03-01"},{"category":"Logement","type":"Dépense","amount":100000,"id":"t423","date":"2026-04-01"},{"category":"Voiture","type":"Dépense","amount":374555,"id":"t424","date":"2026-04-01"},{"category":"Vêtements","type":"Dépense","amount":45000,"id":"t425","date":"2026-04-01"},{"category":"Voyage","type":"Dépense","amount":4000,"id":"t426","date":"2026-04-01"},{"category":"Cadeaux","type":"Dépense","amount":463650,"id":"t427","date":"2026-04-01"},{"category":"Santé","type":"Dépense","amount":22552,"id":"t428","date":"2026-04-01"},{"category":"Générales","type":"Dépense","amount":13580,"id":"t429","date":"2026-04-01"},{"category":"Divertissement","type":"Dépense","amount":173055,"id":"t430","date":"2026-04-01"},{"category":"Aliments","type":"Dépense","amount":161055,"id":"t431","date":"2026-04-01"},{"category":"Personnel","type":"Dépense","amount":14050,"id":"t432","date":"2026-04-01"},{"category":"Abonnements","type":"Dépense","amount":25402,"id":"t433","date":"2026-04-01"},{"category":"Shopping","type":"Dépense","amount":63150,"id":"t434","date":"2026-04-01"},{"category":"Enfants & Maman","type":"Dépense","amount":371850,"id":"t435","date":"2026-04-01"},{"category":"Invitation","type":"Dépense","amount":159000,"id":"t436","date":"2026-04-01"},{"category":"Ajustement","type":"Dépense","amount":58170,"id":"t437","date":"2026-04-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t438","date":"2026-04-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t439","date":"2026-04-01"},{"category":"GRUNDFOS","type":"Dépense","amount":250708,"id":"t440","date":"2026-04-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t441","date":"2026-04-01"},{"category":"Déménagement","type":"Dépense","amount":81000,"id":"t442","date":"2026-04-01"},{"category":"Securicompte","type":"Dépense","amount":912,"id":"t443","date":"2026-04-01"},{"category":"Prêt","type":"Dépense","amount":82290,"id":"t444","date":"2026-04-01"},{"category":"Voiture","type":"Dépense","amount":10000,"id":"t445","date":"2026-05-01"},{"category":"Voyage","type":"Dépense","amount":98200,"id":"t446","date":"2026-05-01"},{"category":"Cadeaux","type":"Dépense","amount":1055110,"id":"t447","date":"2026-05-01"},{"category":"Santé","type":"Dépense","amount":21130,"id":"t448","date":"2026-05-01"},{"category":"Générales","type":"Dépense","amount":5400,"id":"t449","date":"2026-05-01"},{"category":"Divertissement","type":"Dépense","amount":422950,"id":"t450","date":"2026-05-01"},{"category":"Utilitaires","type":"Dépense","amount":250000,"id":"t451","date":"2026-05-01"},{"category":"Aliments","type":"Dépense","amount":246470,"id":"t452","date":"2026-05-01"},{"category":"Personnel","type":"Dépense","amount":267000,"id":"t453","date":"2026-05-01"},{"category":"Abonnements","type":"Dépense","amount":62816,"id":"t454","date":"2026-05-01"},{"category":"Shopping","type":"Dépense","amount":283034,"id":"t455","date":"2026-05-01"},{"category":"Enfants & Maman","type":"Dépense","amount":215800,"id":"t456","date":"2026-05-01"},{"category":"Invitation","type":"Dépense","amount":259100,"id":"t457","date":"2026-05-01"},{"category":"Ajustement","type":"Dépense","amount":188136,"id":"t458","date":"2026-05-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t459","date":"2026-05-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t460","date":"2026-05-01"},{"category":"GRUNDFOS","type":"Dépense","amount":381175,"id":"t461","date":"2026-05-01"},{"category":"Pack Club","type":"Dépense","amount":9087,"id":"t462","date":"2026-05-01"},{"category":"Déménagement","type":"Dépense","amount":5573690,"id":"t463","date":"2026-05-01"},{"category":"Logement","type":"Dépense","amount":1201550,"id":"t464","date":"2026-06-01"},{"category":"Voiture","type":"Dépense","amount":6000,"id":"t465","date":"2026-06-01"},{"category":"Vêtements","type":"Dépense","amount":95800,"id":"t466","date":"2026-06-01"},{"category":"Éducation","type":"Dépense","amount":9300,"id":"t467","date":"2026-06-01"},{"category":"Cadeaux","type":"Dépense","amount":728664,"id":"t468","date":"2026-06-01"},{"category":"Santé","type":"Dépense","amount":10080,"id":"t469","date":"2026-06-01"},{"category":"Générales","type":"Dépense","amount":71884,"id":"t470","date":"2026-06-01"},{"category":"Divertissement","type":"Dépense","amount":480660,"id":"t471","date":"2026-06-01"},{"category":"Aliments","type":"Dépense","amount":224500,"id":"t472","date":"2026-06-01"},{"category":"Personnel","type":"Dépense","amount":4000,"id":"t473","date":"2026-06-01"},{"category":"Abonnements","type":"Dépense","amount":49738,"id":"t474","date":"2026-06-01"},{"category":"Shopping","type":"Dépense","amount":134400,"id":"t475","date":"2026-06-01"},{"category":"Enfants & Maman","type":"Dépense","amount":247800,"id":"t476","date":"2026-06-01"},{"category":"Invitation","type":"Dépense","amount":54500,"id":"t477","date":"2026-06-01"},{"category":"Ajustement","type":"Dépense","amount":93307,"id":"t478","date":"2026-06-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t479","date":"2026-06-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t480","date":"2026-06-01"},{"category":"GRUNDFOS","type":"Dépense","amount":522985,"id":"t481","date":"2026-06-01"},{"category":"INVEST SGO","type":"Dépense","amount":101725,"id":"t482","date":"2026-06-01"},{"category":"Vacance Nesher","type":"Dépense","amount":571330,"id":"t483","date":"2026-06-01"},{"category":"Logement","type":"Dépense","amount":100000,"id":"t484","date":"2026-07-01"},{"category":"Voiture","type":"Dépense","amount":153020,"id":"t485","date":"2026-07-01"},{"category":"Vêtements","type":"Dépense","amount":339642,"id":"t486","date":"2026-07-01"},{"category":"Éducation","type":"Dépense","amount":8900,"id":"t487","date":"2026-07-01"},{"category":"Cadeaux","type":"Dépense","amount":452987,"id":"t488","date":"2026-07-01"},{"category":"Santé","type":"Dépense","amount":25565,"id":"t489","date":"2026-07-01"},{"category":"Générales","type":"Dépense","amount":15500,"id":"t490","date":"2026-07-01"},{"category":"Divertissement","type":"Dépense","amount":410360,"id":"t491","date":"2026-07-01"},{"category":"Aliments","type":"Dépense","amount":127295,"id":"t492","date":"2026-07-01"},{"category":"Personnel","type":"Dépense","amount":18960,"id":"t493","date":"2026-07-01"},{"category":"Abonnements","type":"Dépense","amount":32694,"id":"t494","date":"2026-07-01"},{"category":"Shopping","type":"Dépense","amount":303913,"id":"t495","date":"2026-07-01"},{"category":"Enfants & Maman","type":"Dépense","amount":149270,"id":"t496","date":"2026-07-01"},{"category":"Invitation","type":"Dépense","amount":107500,"id":"t497","date":"2026-07-01"},{"category":"Ajustement","type":"Dépense","amount":86239,"id":"t498","date":"2026-07-01"},{"category":"Âge D'or Retraite","type":"Dépense","amount":40000,"id":"t499","date":"2026-07-01"},{"category":"Plan Éducation","type":"Dépense","amount":32457,"id":"t500","date":"2026-07-01"},{"category":"GRUNDFOS","type":"Dépense","amount":469655,"id":"t501","date":"2026-07-01"},{"category":"Pack Club","type":"Dépense","amount":10000,"id":"t502","date":"2026-07-01"},{"category":"Achat Terrain Port Bouet","type":"Dépense","amount":422100,"id":"t503","date":"2026-07-01"},{"category":"Prêt","type":"Dépense","amount":80333,"id":"t504","date":"2026-07-01"},{"category":"Vacance Nesher","type":"Dépense","amount":162390,"id":"t505","date":"2026-07-01"},{"category":"Cadeaux","type":"Dépense","amount":405,"id":"t506","date":"2026-08-01","note":"Femme"},{"category":"Générales","type":"Dépense","amount":3999,"id":"t507","date":"2026-08-01"},{"category":"Petty Cash","type":"Revenu","amount":494562,"id":"t508","date":"2024-06-01"},{"category":"Revenu général","type":"Revenu","amount":2630203,"id":"t509","date":"2024-07-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t510","date":"2024-07-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":730147,"id":"t511","date":"2024-07-01"},{"category":"Ajustement","type":"Revenu","amount":31000,"id":"t512","date":"2024-07-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t513","date":"2024-08-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":732311,"id":"t514","date":"2024-08-01"},{"category":"Vente Pompe","type":"Revenu","amount":1000500,"id":"t515","date":"2024-08-01"},{"category":"Ajustement","type":"Revenu","amount":278412,"id":"t516","date":"2024-08-01"},{"category":"Emprunt Bancaire","type":"Revenu","amount":6664569,"id":"t517","date":"2024-08-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t518","date":"2024-09-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":732311,"id":"t519","date":"2024-09-01"},{"category":"Ajustement","type":"Revenu","amount":7666,"id":"t520","date":"2024-09-01"},{"category":"Petty Cash","type":"Revenu","amount":2500000,"id":"t521","date":"2024-09-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t522","date":"2024-09-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t523","date":"2024-10-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1697111,"id":"t524","date":"2024-10-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t525","date":"2024-11-01"},{"category":"Un salaire","type":"Revenu","amount":2391204,"id":"t526","date":"2024-12-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t527","date":"2024-12-01"},{"category":"Vente Pompe","type":"Revenu","amount":721307,"id":"t528","date":"2024-12-01"},{"category":"Petty Cash","type":"Revenu","amount":2000000,"id":"t529","date":"2024-12-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t530","date":"2024-12-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t531","date":"2025-01-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":2238058,"id":"t532","date":"2025-01-01"},{"category":"Vente Pompe","type":"Revenu","amount":410000,"id":"t533","date":"2025-01-01"},{"category":"Ajustement","type":"Revenu","amount":20100,"id":"t534","date":"2025-01-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t535","date":"2025-02-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t536","date":"2025-02-01"},{"category":"Vente Pompe","type":"Revenu","amount":250000,"id":"t537","date":"2025-02-01"},{"category":"Ajustement","type":"Revenu","amount":1242,"id":"t538","date":"2025-02-01"},{"category":"Petty Cash","type":"Revenu","amount":2500000,"id":"t539","date":"2025-02-01"},{"category":"Un salaire","type":"Revenu","amount":1419055,"id":"t540","date":"2025-03-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t541","date":"2025-03-01"},{"category":"Vente Pompe","type":"Revenu","amount":1447000,"id":"t542","date":"2025-03-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t543","date":"2025-03-01"},{"category":"Revenu général","type":"Revenu","amount":0,"id":"t544","date":"2025-04-01"},{"category":"Un salaire","type":"Revenu","amount":2804390,"id":"t545","date":"2025-04-01"},{"category":"Épargne","type":"Revenu","amount":900000,"id":"t546","date":"2025-04-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t547","date":"2025-04-01"},{"category":"Vente Pompe","type":"Revenu","amount":230900,"id":"t548","date":"2025-04-01"},{"category":"Ajustement","type":"Revenu","amount":1280,"id":"t549","date":"2025-04-01"},{"category":"Petty Cash","type":"Revenu","amount":434075,"id":"t550","date":"2025-04-01"},{"category":"General","type":"Revenu","amount":2682680,"id":"t551","date":"2025-04-01"},{"category":"Revenu général","type":"Revenu","amount":1,"id":"t552","date":"2025-05-01"},{"category":"Un salaire","type":"Revenu","amount":1553752,"id":"t553","date":"2025-05-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t554","date":"2025-05-01"},{"category":"Vente Pompe","type":"Revenu","amount":162290,"id":"t555","date":"2025-05-01"},{"category":"Ajustement","type":"Revenu","amount":20538,"id":"t556","date":"2025-05-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t557","date":"2025-05-01"},{"category":"Revenu général","type":"Revenu","amount":0,"id":"t558","date":"2025-06-01"},{"category":"Un salaire","type":"Revenu","amount":1553752,"id":"t559","date":"2025-06-01"},{"category":"Bourse","type":"Revenu","amount":37300,"id":"t560","date":"2025-06-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t561","date":"2025-06-01"},{"category":"Vente Pompe","type":"Revenu","amount":913660,"id":"t562","date":"2025-06-01"},{"category":"Ajustement","type":"Revenu","amount":12480,"id":"t563","date":"2025-06-01"},{"category":"Petty Cash","type":"Revenu","amount":2500000,"id":"t564","date":"2025-06-01"},{"category":"Un salaire","type":"Revenu","amount":1553752,"id":"t565","date":"2025-07-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t566","date":"2025-07-01"},{"category":"Vente Pompe","type":"Revenu","amount":50000,"id":"t567","date":"2025-07-01"},{"category":"Ajustement","type":"Revenu","amount":3787,"id":"t568","date":"2025-07-01"},{"category":"Petty Cash","type":"Revenu","amount":2500000,"id":"t569","date":"2025-07-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t570","date":"2025-08-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t571","date":"2025-08-01"},{"category":"Ajustement","type":"Revenu","amount":28882,"id":"t572","date":"2025-08-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t573","date":"2025-08-01"},{"category":"General","type":"Revenu","amount":28000,"id":"t574","date":"2025-08-01"},{"category":"Revenu général","type":"Revenu","amount":1443037,"id":"t575","date":"2025-09-01"},{"category":"Allocation","type":"Revenu","amount":120000,"id":"t576","date":"2025-09-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t577","date":"2025-09-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t578","date":"2025-09-01"},{"category":"Vente Pompe","type":"Revenu","amount":150000,"id":"t579","date":"2025-09-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t580","date":"2025-10-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t581","date":"2025-10-01"},{"category":"Vente Pompe","type":"Revenu","amount":340000,"id":"t582","date":"2025-10-01"},{"category":"Ajustement","type":"Revenu","amount":11755,"id":"t583","date":"2025-10-01"},{"category":"Petty Cash","type":"Revenu","amount":2500000,"id":"t584","date":"2025-10-01"},{"category":"Prêt Orange","type":"Revenu","amount":63464,"id":"t585","date":"2025-10-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t586","date":"2025-11-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t587","date":"2025-11-01"},{"category":"Vente Pompe","type":"Revenu","amount":59500,"id":"t588","date":"2025-11-01"},{"category":"Ajustement","type":"Revenu","amount":57651,"id":"t589","date":"2025-11-01"},{"category":"Loyer","type":"Revenu","amount":395020,"id":"t590","date":"2025-11-01"},{"category":"ECO PUMP","type":"Revenu","amount":585000,"id":"t591","date":"2025-11-01"},{"category":"Revenu général","type":"Revenu","amount":2500000,"id":"t592","date":"2025-12-01"},{"category":"Un salaire","type":"Revenu","amount":2636156,"id":"t593","date":"2025-12-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t594","date":"2025-12-01"},{"category":"Vente Pompe","type":"Revenu","amount":200000,"id":"t595","date":"2025-12-01"},{"category":"Ajustement","type":"Revenu","amount":80,"id":"t596","date":"2025-12-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t597","date":"2025-12-01"},{"category":"Revenu général","type":"Revenu","amount":29305,"id":"t598","date":"2026-01-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t599","date":"2026-01-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t600","date":"2026-01-01"},{"category":"Vente Pompe","type":"Revenu","amount":1524500,"id":"t601","date":"2026-01-01"},{"category":"Ajustement","type":"Revenu","amount":362,"id":"t602","date":"2026-01-01"},{"category":"Petty Cash","type":"Revenu","amount":2500000,"id":"t603","date":"2026-01-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t604","date":"2026-02-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t605","date":"2026-02-01"},{"category":"Prêt Orange","type":"Revenu","amount":76984,"id":"t606","date":"2026-02-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t607","date":"2026-03-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t608","date":"2026-03-01"},{"category":"Vente Pompe","type":"Revenu","amount":150000,"id":"t609","date":"2026-03-01"},{"category":"Ajustement","type":"Revenu","amount":415179,"id":"t610","date":"2026-03-01"},{"category":"Petty Cash","type":"Revenu","amount":1085000,"id":"t611","date":"2026-03-01"},{"category":"Loyer","type":"Revenu","amount":1500000,"id":"t612","date":"2026-03-01"},{"category":"General","type":"Revenu","amount":200000,"id":"t613","date":"2026-03-01"},{"category":"Un salaire","type":"Revenu","amount":1555362,"id":"t614","date":"2026-04-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t615","date":"2026-04-01"},{"category":"Ajustement","type":"Revenu","amount":4249,"id":"t616","date":"2026-04-01"},{"category":"General","type":"Revenu","amount":150000,"id":"t617","date":"2026-04-01"},{"category":"Un salaire","type":"Revenu","amount":3340975,"id":"t618","date":"2026-05-01"},{"category":"Vente Pompe","type":"Revenu","amount":265000,"id":"t619","date":"2026-05-01"},{"category":"General","type":"Revenu","amount":85000,"id":"t620","date":"2026-05-01"},{"category":"Prêt Orange","type":"Revenu","amount":75050,"id":"t621","date":"2026-05-01"},{"category":"Revenu général","type":"Revenu","amount":15000,"id":"t622","date":"2026-06-01"},{"category":"Un salaire","type":"Revenu","amount":1629526,"id":"t623","date":"2026-06-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":2238058,"id":"t624","date":"2026-06-01"},{"category":"Ajustement","type":"Revenu","amount":1000,"id":"t625","date":"2026-06-01"},{"category":"Petty Cash","type":"Revenu","amount":2482797,"id":"t626","date":"2026-06-01"},{"category":"Loyer","type":"Revenu","amount":1490164,"id":"t627","date":"2026-06-01"},{"category":"General","type":"Revenu","amount":150000,"id":"t628","date":"2026-06-01"},{"category":"Un salaire","type":"Revenu","amount":1629526,"id":"t629","date":"2026-07-01"},{"category":"Revenus Location Mazda","type":"Revenu","amount":1119029,"id":"t630","date":"2026-07-01"},{"category":"Vente Pompe","type":"Revenu","amount":260000,"id":"t631","date":"2026-07-01"},{"category":"Ajustement","type":"Revenu","amount":900,"id":"t632","date":"2026-07-01"},{"category":"Petty Cash","type":"Revenu","amount":2482797,"id":"t633","date":"2026-07-01"},{"category":"Voiture","type":"Dépense","amount":2000,"date":"2026-08-01","note":"Lavage","id":"t634"},{"category":"Aliments","type":"Dépense","amount":4500,"date":"2026-08-01","note":"Le déjeuner","id":"t635"},{"category":"Divertissement","type":"Dépense","amount":7000,"date":"2026-08-01","note":"Alcool","id":"t636"},{"category":"Cadeaux","type":"Dépense","amount":2000,"date":"2026-08-01","note":"Pourboire","id":"t637"},{"category":"Aliments","type":"Dépense","amount":17000,"date":"2026-08-01","note":"Dîner","id":"t638"}];
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
  "Transport": ["Peage", "Souterrain", "Autobus", "Taxi", "Train"],
  "Utilitaires": ["Nettoyage", "Électricité", "Eau", "Gaz", "Chauffage", "Des ordures", "l'Internet", "Téléphones", "la télé", "Ordinateur HP"],
  "Voiture": ["Lavage", "Peinture-Retouche", "Installation GPS", "Carburant", "Visite Technique", "Entretien", "La lessive", "Parking", "Assurance"],
  "Voyage": ["Pourboire", "Peage", "Divertissement", "Aliment", "Shopping", "Vol", "Un hôtel", "Location de voiture"],
  "Vêtements": ["Chemises", "Chaussures", "Un pantalon", "Tops", "Des sacs", "Accessoires", "Draps"],
  "Éducation": ["Nesher", "Cours", "Livres", "Fournitures scolaires", "Prêt étudiant"],
  "Dette": ["PEL"],
  "Divertissement": ["Femme", "Residence", "Alcool", "BAP", "Films", "Boisson", "Anniversaire", "La musique", "Jeux", "Performance", "Fête", "Funérailles"],
  "Déménagement": ["Lits", "Installation Clim Chauffe Eau", "Nettoyage", "Micro-ondes Four", "Chauffe-eau", "Réfrigérateur", "Gaziniere", "Remplacement Gaziniere", "Deco & Senteur", "Electricien", "Étagère Cuisine", "Splits", "Autres"],
  "Aliments": ["Déjeuner", "Invitation", "Le déjeuner", "Dîner", "Les courses", "Dîner à l'extérieur"],
  "Cadeaux": ["Mardochee", "Olokpacha", "Femme", "Ruth", "Enfants Nesher", "Ndjore", "Pourboire", "Adrien", "Cotisations", "Metty", "Obed", "Anniversaire", "Noël", "Juste pour le fun", "Dot Jo"],
  "Création Entreprise ECO PUMP AFRIK": ["FNE", "Documents", "Dédouanement", "Timbre", "Application De Gestion", "Cachet", "Boîte Postale"],
  "Des sports": ["Gym", "Équipement", "Piscine"],
  "Enfants & Maman": ["Maman", "Nesher", "Hemra"],
  "Formation": ["Finelo Invest", "Emergent", "Piano & Guitare"],
  "GRUNDFOS": ["Appel", "Restaurant", "Location Prado", "Eau", "Enjoy", "Impression", "Impression Allowance", "Internet", "Ajustement Petty Cash", "Voyage", "FedEx", "Hotel", "Électricité", "Peage", "iPhone 16 Pro", "Divertissement", "Cachet", "Carburant", "Infraction", "AUTRES", "Hinoter"],
  "Générales": ["Cachet Grundfos", "Impression", "Abonnement IScanner", "Visite Maison", "Cachet OMÉGA", "Certificat De Perte SIB", "Carte Money Fusion", "Police", "Réparation iPhone 13", "Souris Sans Fil", "Carte Djamo", "Vol Djamo", "Badoo", "Yango Livraison", "Peage", "Réparation Robinet", "Création Entreprise ECO PUMP", "Yango Transport Pompe", "Payement Pour Terrain Port Bouet", "Livraison Pompe"],
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
// Les transactions sans compte assigné (données historiques importées) n'affectent aucun
// solde, puisqu'elles sont déjà reflétées dans le solde de départ au moment du suivi.
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
  { id: "a1", name: "SIB", kind: "Banque", openingBalance: 1578.50 },
  { id: "a2", name: "PETTY CASH", kind: "Banque", openingBalance: 6176159 },
  { id: "a3", name: "Dépôt LOYER", kind: "Banque", openingBalance: 184164 },
  { id: "a4", name: "Revenus MAZDA", kind: "Banque", openingBalance: 4162759.50 },
  { id: "a5", name: "SALAIRE", kind: "Banque", openingBalance: 1099559 },
  { id: "a6", name: "SGO", kind: "Banque", openingBalance: -463429 },
  { id: "a7", name: "PUMP", kind: "Espèces", openingBalance: 794500 },
];

const seedBudgets: CategoryBudget[] = [
  { id: "b1", category: "Cadeaux", amount: 300000, rollover: false },
  { id: "b2", category: "Divertissement", amount: 200000, rollover: false },
  { id: "b3", category: "Invitation", amount: 100000, rollover: false },
];

const seedGoals: Goal[] = [
  { id: "g1", name: "Valeur nette cible", target: 20000000, current: 11955291, date: "déc. 2027" },
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
  const [y, mm] = m.split("_");
  return `${MONTH_NAMES[parseInt(mm, 10) - 1]} ${y.slice(2)}`;
}
function monthSortKey(m: string) {
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
function dateToMonthKey(date: string) {
  const [y, m] = date.split("-");
  return `${parseInt(y, 10)}_${parseInt(m, 10)}`;
}
function monthKeyToFirstDate(monthKey: string) {
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
function categoriesForType(transactions: Transaction[], type: TxType): string[] {
  const used = new Set(transactions.filter((t) => t.type === type).map((t) => t.category));
  const known = Object.keys(type === "Dépense" ? depSubcategories : revSubcategories);
  known.forEach((c) => used.add(c));
  return Array.from(used).sort((a, b) => a.localeCompare(b, "fr"));
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
    (async () => {
      try {
        const res = await window.storage?.get(key, false);
        if (res) setState(JSON.parse(res.value));
      } catch {}
      setLoaded(true);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => {
    if (!loaded) return;
    window.storage?.set(key, JSON.stringify(state), false).catch(() => {});
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
interface Filters { from: string; to: string; type: string; group: string; category: string; search: string; scope: string; }

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
      <Select label="Catégorie" value={filters.category} onChange={(v) => patch({ category: v })} options={[{ value: "Toutes", label: "Toutes" }, ...allCategories.map((c) => ({ value: c, label: c }))]} />
      <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 160 }}>
        <label style={{ fontSize: 10.5, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.06em" }}>Recherche</label>
        <div style={{ position: "relative" }}>
          <Search size={13} color={COLOR.inkMuted} style={{ position: "absolute", left: 9, top: 9 }} />
          <input value={filters.search} onChange={(e) => patch({ search: e.target.value })} placeholder="ex: Cadeaux, GRUNDFOS…" style={{ width: "100%", background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink, padding: "8px 10px 8px 28px", fontSize: 12.5, fontFamily: "'Inter', sans-serif", boxSizing: "border-box" }} />
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
        <Panel title="Évolution de la valeur nette" subtitle="Relevé de solde — rapport séparé, non affecté par les filtres de type/groupe">
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
        </Panel>
      )}

      <Panel title="Revenus vs Dépenses" subtitle={`${byMonth.length} mois dans la période filtrée`}>
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
      </Panel>

      <Panel title="Répartition des dépenses par groupe" subtitle="Selon la classification des catégories">
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
      </Panel>
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
      <Panel title="Flux Revenus → Dépenses" subtitle="Comment le revenu de la période se répartit entre les groupes de dépenses et le solde">
        <FlowDiagram filtered={filtered} />
      </Panel>
      <Panel title="Calendrier d'intensité — dépenses non-productives" subtitle="Repérer les mois et saisons à risque">
        <HeatmapCalendar filtered={filtered} />
      </Panel>
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
      <Panel title="Revenus, dépenses et solde par mois" subtitle="Cliquer sur une colonne pour trier">
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
      </Panel>
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
      <Panel title="Détail par catégorie" subtitle={`${rows.length} catégorie(s) · cliquez sur une catégorie pour voir ses sous-catégories · reclassez le groupe via le menu`}
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
                <div onClick={() => toggleExpand(r.name)} style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }}>
                  <ChevronRight size={13} color={COLOR.inkMuted} style={{ flexShrink: 0, transform: isOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                  <div style={{ width: 178, fontSize: 12.5, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.name}>{r.name}</div>
                  <div style={{ flex: 1, background: COLOR.hairline, borderRadius: 4, height: 16, position: "relative" }}>
                    <div style={{ width: `${(r.value / maxVal) * 100}%`, height: "100%", borderRadius: 4, background: groupColor[r.group] || COLOR.inkMuted }} />
                  </div>
                  <div style={{ width: 95, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, flexShrink: 0 }}>{fmt(r.value)}</div>
                  <div style={{ width: 42, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: COLOR.inkMuted, flexShrink: 0 }}>{total ? ((r.value / total) * 100).toFixed(1) : "0"}%</div>
                  {r.type === "Dépense" ? (
                    <select value={resolvedGroups[r.name] || "Non classifié"} onClick={(e) => e.stopPropagation()} onChange={(e) => setCategoryGroups({ ...categoryGroups, [r.name]: e.target.value as Group })}
                      style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: groupColor[resolvedGroups[r.name] || "Non classifié"], padding: "5px 8px", fontSize: 11.5, fontFamily: "'Inter', sans-serif", flexShrink: 0, width: 128, cursor: "pointer" }}>
                      {GROUPS.map((g) => <option key={g} value={g}>{g}</option>)}
                    </select>
                  ) : <div style={{ width: 128, flexShrink: 0, fontSize: 11.5, color: COLOR.goldSoft, textAlign: "center" }}>Revenu</div>}
                </div>
                {isOpen && (
                  <div style={{ marginLeft: 25, marginTop: 8, marginBottom: 4, paddingLeft: 12, borderLeft: `2px solid ${COLOR.hairline}`, display: "flex", flexDirection: "column", gap: 5 }}>
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
// GROUPES TAB
// ============================================================
function GroupesTab({ filtered }: { filtered: any[] }) {
  const dep = filtered.filter((t) => t.type === "Dépense");
  const totalDep = dep.reduce((a, t) => a + t.amount, 0);
  const cards = GROUPS.map((g) => {
    const items = dep.filter((t) => t.group === g);
    const value = items.reduce((a, t) => a + t.amount, 0);
    const cats: Record<string, number> = {};
    items.forEach((t) => { cats[t.category] = (cats[t.category] || 0) + t.amount; });
    const top = Object.entries(cats).sort((a, b) => b[1] - a[1]).slice(0, 4);
    return { group: g, value, count: items.length, top, pct: totalDep ? (value / totalDep) * 100 : 0 };
  });
  return (
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
      <Panel title="Top 10 catégories — comparaison par année" subtitle="Totaux bruts par catégorie et par année">
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
              <input type="number" value={cap} onChange={(e) => setCap(Number(e.target.value) || 0)} style={{ background: COLOR.surfaceInput, border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.ink, padding: "6px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, width: 130 }} />
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

      <Panel title="Impact projeté sur la valeur nette" subtitle="Comparaison : trajectoire actuelle vs trajectoire avec les réductions appliquées chaque mois">
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
      </Panel>
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
    <Panel title="Projection de valeur nette" subtitle={`Basée sur la tendance des 6 derniers relevés — bande optimiste/pessimiste (±1 écart-type)`}
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
    </Panel>
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
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Cible (FCFA)</label><input type="number" style={{ ...inputStyle, width: 150 }} value={form.target} onChange={(e) => setForm({ ...form, target: Number(e.target.value) || 0 })} /></div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Déjà atteint (FCFA)</label><input type="number" style={{ ...inputStyle, width: 150 }} value={form.current} onChange={(e) => setForm({ ...form, current: Number(e.target.value) || 0 })} /></div>
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
                <input type="number" value={g.current} onChange={(e) => update(g.id, { current: Number(e.target.value) || 0 })} style={{ ...inputStyle, width: 140 }} />
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
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant</label><input type="number" style={{ ...inputStyle, width: 130 }} value={form.amount || ""} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) })} /></div>
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
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Solde de départ (FCFA)</label><input type="number" style={{ ...inputStyle, width: 160 }} value={form.openingBalance} onChange={(e) => setForm({ ...form, openingBalance: Number(e.target.value) || 0 })} /></div>
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
                    <input type="number" value={a.openingBalance} onChange={(e) => update(a.id, { openingBalance: Number(e.target.value) || 0 })} style={{ ...inputStyle, width: 150 }} />
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
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Limite mensuelle</label><input type="number" style={{ ...inputStyle, width: 150 }} value={form.amount} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) || 0 })} /></div>
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
                    <input type="number" style={{ ...inputStyle, width: 120 }} value={transferAmount || ""} onChange={(e) => setTransferAmount(Number(e.target.value) || 0)} placeholder="Montant" />
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
    setTransactions([...transactions, { id: uid(), date: r.nextDate, category: r.category, type: r.type, amount: r.amount, account: r.account, payee: r.payee }]);
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
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant</label><input type="number" style={{ ...inputStyle, width: 130 }} value={form.amount || ""} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) })} /></div>
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
function SauvegardeTab({ getSnapshot, restore }: { getSnapshot: () => any; restore: (data: any) => void }) {
  const [status, setStatus] = useState<string | null>(null);
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

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
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
  return { date: todayISO(), category: categoriesForType(transactions, "Dépense")[0] || "Cadeaux", type: "Dépense", amount: 0, account: accounts[0]?.name };
}

function JournalTab({ filtered, allCategories, categoryGroups, transactions, setTransactions, rules, setRules, accounts }: {
  filtered: any[]; allCategories: string[]; categoryGroups: Record<string, Group>;
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void;
  rules: CategorizationRule[]; setRules: (r: CategorizationRule[]) => void; accounts: Account[];
}) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState<Omit<Transaction, "id">>(emptyForm(transactions, accounts));
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<Omit<Transaction, "id"> | null>(null);
  const [page, setPage] = useState(0);
  const [showImport, setShowImport] = useState(false);
  const [importText, setImportText] = useState("");
  const [importPreview, setImportPreview] = useState<Transaction[]>([]);
  const [showRules, setShowRules] = useState(false);
  const [newRule, setNewRule] = useState({ keyword: "", group: "Non classifié" as Group });
  const [selected, setSelected] = useState<Set<string>>(new Set());
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
  const startEdit = (t: Transaction) => { setEditingId(t.id); setEditForm({ date: t.date, category: t.category, subcategory: t.subcategory, type: t.type, amount: t.amount, payee: t.payee, note: t.note, account: t.account }); };
  const saveEdit = () => { if (!editingId || !editForm) return; setTransactions(transactions.map((t) => (t.id === editingId ? { ...editForm, id: editingId } : t))); setEditingId(null); setEditForm(null); };
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
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}><label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant (FCFA)</label><input style={inputStyle} type="number" value={form.amount || ""} onChange={(e) => setForm({ ...form, amount: Number(e.target.value) })} /></div>
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
            <button onClick={bulkDelete} style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: `1px solid ${COLOR.clay}`, borderRadius: 6, color: COLOR.claySoft, padding: "6px 12px", fontSize: 11.5, cursor: "pointer" }}><Trash2 size={12} /> Supprimer la sélection</button>
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
                const isEditing = editingId === t.id;
                return (
                  <tr key={t.id}>
                    <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}` }}>
                      <button onClick={() => toggleSelect(t.id)} style={{ background: "transparent", border: "none", cursor: "pointer", display: "flex" }}>
                        {selected.has(t.id) ? <CheckSquare size={14} color={COLOR.goldSoft} /> : <Square size={14} color={COLOR.inkMuted} />}
                      </button>
                    </td>
                    {isEditing && editForm ? (
                      <>
                        <td style={{ padding: 6 }}><input type="date" style={inputStyle} value={editForm.date} onChange={(e) => setEditForm({ ...editForm, date: e.target.value })} /></td>
                        <td style={{ padding: 6 }}>
                          <select style={inputStyle} value={editForm.category} onChange={(e) => setEditForm({ ...editForm, category: e.target.value, subcategory: "" })}>
                            {categoriesForType(transactions, editForm.type).map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                          {getSubcategories(editForm.type, editForm.category).length > 0 && (
                            <select style={{ ...inputStyle, marginTop: 4 }} value={editForm.subcategory || ""} onChange={(e) => setEditForm({ ...editForm, subcategory: e.target.value })}>
                              <option value="">— sous-catégorie —</option>
                              {getSubcategories(editForm.type, editForm.category).map((s) => <option key={s} value={s}>{s}</option>)}
                            </select>
                          )}
                        </td>
                        <td style={{ padding: 6 }}><select style={inputStyle} value={editForm.type} onChange={(e) => { const ty = e.target.value as TxType; setEditForm({ ...editForm, type: ty, subcategory: "", category: categoriesForType(transactions, ty)[0] || editForm.category }); }}><option value="Dépense">Dépense</option><option value="Revenu">Revenu</option></select></td>
                        <td style={{ padding: 6 }}>
                          <select style={inputStyle} value={editForm.account || ""} onChange={(e) => setEditForm({ ...editForm, account: e.target.value })}>
                            {!accounts.length && <option value="">Aucun compte</option>}
                            {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                          </select>
                        </td>
                        <td style={{ padding: 6 }}><input style={{ ...inputStyle, textAlign: "right" }} type="number" value={editForm.amount} onChange={(e) => setEditForm({ ...editForm, amount: Number(e.target.value) })} /></td>
                        <td style={{ padding: 6, whiteSpace: "nowrap" }}><button onClick={saveEdit} style={iconBtnStyle(COLOR.emerald)}><Save size={13} /></button><button onClick={() => { setEditingId(null); setEditForm(null); }} style={iconBtnStyle(COLOR.inkMuted)}><X size={13} /></button></td>
                      </>
                    ) : (
                      <>
                        <td style={{ padding: "9px 10px", fontSize: 12.5, borderBottom: `1px solid ${COLOR.hairline}`, fontFamily: "'IBM Plex Mono', monospace" }}>{dateLabelFull(t.date)}</td>
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
                        <td style={{ padding: "9px 10px", borderBottom: `1px solid ${COLOR.hairline}`, whiteSpace: "nowrap" }}><button onClick={() => startEdit(t)} style={iconBtnStyle(COLOR.slateBlueSoft)}><Pencil size={13} /></button><button onClick={() => remove(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button></td>
                      </>
                    )}
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
    </div>
  );
}

// ============================================================
// RAPPORTS & EXPORT TAB
// ============================================================
function ExportTab({ filtered, filters }: { filtered: any[]; filters: Filters }) {
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
    const header = "Date,Mois,Catégorie,Sous-catégorie,Type,Groupe,Compte,Montant\n";
    const rows = filtered.map((t) => `${t.date},${monthLabel(t.month)},"${t.category}","${t.subcategory || ""}",${t.type},${t.group},"${t.account || ""}",${t.amount}`).join("\n");
    const blob = new Blob([header + rows], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `grand-livre_${filters.from}_${filters.to}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Panel title="Résumé automatique" subtitle="Généré à partir de la période filtrée" right={<Sparkles size={16} color={COLOR.goldSoft} />}>
        <p style={{ fontSize: 13.5, lineHeight: 1.7, color: COLOR.ink, margin: 0 }}>{summary}</p>
      </Panel>
      <Panel title="Exporter le rapport filtré" subtitle={`${filtered.length} transaction(s) dans la sélection actuelle`}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <button onClick={exportCSV} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(63,156,122,0.14)", border: `1px solid ${COLOR.emerald}`, borderRadius: 8, color: COLOR.emeraldSoft, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
            <Download size={15} /> Exporter en CSV (Excel)
          </button>
          <button onClick={() => window.print()} style={{ display: "flex", alignItems: "center", gap: 8, background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 8, color: COLOR.inkMuted, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
            <Printer size={15} /> Imprimer / PDF
          </button>
        </div>
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
  const [quickCategory, setQuickCategory] = useState(() => categoriesForType(transactions, "Dépense")[0] || "Aliments");
  const [quickSubcategory, setQuickSubcategory] = useState("");
  const [quickType, setQuickType] = useState<TxType>("Dépense");
  const [quickAmount, setQuickAmount] = useState<number | "">("");
  const [quickAccount, setQuickAccount] = useState(() => accounts[0]?.name || "");
  const [justAdded, setJustAdded] = useState(false);

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

  const submit = () => {
    if (!quickCategory || !quickAmount || Number(quickAmount) <= 0) return;
    setTransactions([...transactions, { id: uid(), date: quickDate, category: quickCategory, subcategory: quickSubcategory || undefined, type: quickType, amount: Number(quickAmount), account: quickAccount || undefined }]);
    setQuickAmount("");
    setJustAdded(true);
    setTimeout(() => setJustAdded(false), 1200);
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
        <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Date</label>
            <input type="date" value={quickDate} onChange={(e) => setQuickDate(e.target.value)} style={{ ...inputStyle, width: 150 }} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Type</label>
            <div style={{ display: "flex", gap: 6 }}>
              {(["Dépense", "Revenu"] as TxType[]).map((ty) => (
                <button key={ty} onClick={() => { setQuickType(ty); setQuickSubcategory(""); setQuickCategory(categoriesForType(transactions, ty)[0] || ""); }} style={{
                  padding: "8px 16px", borderRadius: 6, fontSize: 12.5, cursor: "pointer",
                  border: `1px solid ${quickType === ty ? (ty === "Revenu" ? COLOR.emerald : COLOR.clay) : COLOR.hairline}`,
                  background: quickType === ty ? (ty === "Revenu" ? "rgba(63,156,122,0.15)" : "rgba(193,84,63,0.15)") : "transparent",
                  color: quickType === ty ? (ty === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft) : COLOR.inkMuted,
                }}>{ty}</button>
              ))}
            </div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Catégorie</label>
            <select style={{ ...inputStyle, width: 190 }} value={quickCategory} onChange={(e) => { setQuickCategory(e.target.value); setQuickSubcategory(""); }}>
              {categoriesForType(transactions, quickType).map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          {getSubcategories(quickType, quickCategory).length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Sous-catégorie</label>
              <select style={{ ...inputStyle, width: 160 }} value={quickSubcategory} onChange={(e) => setQuickSubcategory(e.target.value)}>
                <option value="">—</option>
                {getSubcategories(quickType, quickCategory).map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          )}
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant (FCFA)</label>
            <input type="number" value={quickAmount} onChange={(e) => setQuickAmount(e.target.value === "" ? "" : Number(e.target.value))}
              onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
              style={{ ...inputStyle, width: 140 }} placeholder="0" autoFocus />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte</label>
            <select style={{ ...inputStyle, width: 150 }} value={quickAccount} onChange={(e) => setQuickAccount(e.target.value)}>
              {!accounts.length && <option value="">Aucun compte créé</option>}
              {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
            </select>
          </div>
          <button onClick={submit} style={{
            display: "flex", alignItems: "center", gap: 6, background: justAdded ? COLOR.emerald : "rgba(201,162,39,0.16)",
            border: `1px solid ${justAdded ? COLOR.emerald : COLOR.gold}`, borderRadius: 6,
            color: justAdded ? COLOR.bg : COLOR.goldSoft, padding: "8px 16px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", height: 34,
          }}>
            {justAdded ? <Check size={14} /> : <Plus size={14} />} {justAdded ? "Ajouté" : "Ajouter"}
          </button>
        </div>
      </Panel>

      <Panel title={`Entrées du ${dateLabelFull(quickDate)}`} subtitle={`Revenus ${fmt(sumFor((t) => t.date === quickDate).rev)} · Dépenses ${fmt(sumFor((t) => t.date === quickDate).dep)} · Solde ${fmt(sumFor((t) => t.date === quickDate).solde)}`}>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {quickDateEntries.map((t) => (
            <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: COLOR.surfaceRaised, borderRadius: 6 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: groupColor[t.group] || COLOR.inkMuted, display: "inline-block" }} />
                <span style={{ fontSize: 12.5 }}>{t.category}{t.subcategory && ` · ${t.subcategory}`}</span>
                <span style={{ fontSize: 11, color: COLOR.inkMuted }}>{t.type}</span>
                {t.account && <span style={{ fontSize: 10.5, color: COLOR.slateBlueSoft }}>{t.account}</span>}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: t.type === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft }}>{fmt(t.amount)}</span>
                <button onClick={() => remove(t.id)} style={iconBtnStyle(COLOR.claySoft)}><Trash2 size={13} /></button>
              </div>
            </div>
          ))}
          {!quickDateEntries.length && <EmptyState text="Aucune entrée pour cette date." />}
        </div>
      </Panel>
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
      <Panel title="Solde net par jour" subtitle={`${rows.length} jour(s) avec activité dans la période filtrée`}>
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
      </Panel>
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
function QuickAddFAB({ transactions, setTransactions, accounts, categoryGroups }: {
  transactions: Transaction[]; setTransactions: (t: Transaction[]) => void; accounts: Account[]; categoryGroups: Record<string, Group>;
}) {
  const [open, setOpen] = useState(false);
  const [date, setDate] = useState(todayISO());
  const [type, setType] = useState<TxType>("Dépense");
  const [category, setCategory] = useState(() => categoriesForType(transactions, "Dépense")[0] || "");
  const [subcategory, setSubcategory] = useState("");
  const [amount, setAmount] = useState<number | "">("");
  const [account, setAccount] = useState(() => accounts[0]?.name || "");
  const [justAdded, setJustAdded] = useState(false);

  const submit = () => {
    if (!category || !amount || Number(amount) <= 0) return;
    setTransactions([...transactions, { id: uid(), date, category, subcategory: subcategory || undefined, type, amount: Number(amount), account: account || undefined }]);
    setAmount("");
    setJustAdded(true);
    setTimeout(() => setJustAdded(false), 1000);
  };

  const changeType = (ty: TxType) => {
    setType(ty);
    setSubcategory("");
    setCategory(categoriesForType(transactions, ty)[0] || "");
  };

  const group = category ? (type === "Revenu" ? "Revenu" : (categoryGroups[category] || "Non classifié")) : "Revenu";

  return (
    <>
      {open && (
        <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 90, backdropFilter: "blur(2px)" }} />
      )}
      {open && (
        <div onClick={(e) => e.stopPropagation()} style={{
          position: "fixed", bottom: 96, right: 24, width: 340, maxWidth: "calc(100vw - 32px)", zIndex: 100,
          background: COLOR.surface, border: `1px solid ${COLOR.hairline}`, borderRadius: 14, padding: 20,
          boxShadow: "0 12px 40px rgba(0,0,0,0.5)",
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 17 }}>Saisie rapide</div>
            <button onClick={() => setOpen(false)} style={{ background: "transparent", border: "none", color: COLOR.inkMuted, cursor: "pointer", display: "flex" }}><X size={16} /></button>
          </div>

          <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
            {(["Dépense", "Revenu"] as TxType[]).map((ty) => (
              <button key={ty} onClick={() => changeType(ty)} style={{
                flex: 1, padding: "8px 0", borderRadius: 6, fontSize: 12.5, cursor: "pointer",
                border: `1px solid ${type === ty ? (ty === "Revenu" ? COLOR.emerald : COLOR.clay) : COLOR.hairline}`,
                background: type === ty ? (ty === "Revenu" ? "rgba(63,156,122,0.15)" : "rgba(193,84,63,0.15)") : "transparent",
                color: type === ty ? (ty === "Revenu" ? COLOR.emeraldSoft : COLOR.claySoft) : COLOR.inkMuted,
              }}>{ty}</button>
            ))}
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div>
              <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Date</label>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...inputStyle, marginTop: 4 }} />
            </div>
            <div>
              <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Catégorie</label>
              <select value={category} onChange={(e) => { setCategory(e.target.value); setSubcategory(""); }} style={{ ...inputStyle, marginTop: 4 }}>
                {categoriesForType(transactions, type).map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            {getSubcategories(type, category).length > 0 && (
              <div>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Sous-catégorie</label>
                <select value={subcategory} onChange={(e) => setSubcategory(e.target.value)} style={{ ...inputStyle, marginTop: 4 }}>
                  <option value="">—</option>
                  {getSubcategories(type, category).map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            )}
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ flex: 1 }}>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Montant (FCFA)</label>
                <input type="number" value={amount} onChange={(e) => setAmount(e.target.value === "" ? "" : Number(e.target.value))}
                  onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
                  style={{ ...inputStyle, marginTop: 4 }} placeholder="0" autoFocus />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ fontSize: 10.5, color: COLOR.inkMuted }}>Compte</label>
                <select value={account} onChange={(e) => setAccount(e.target.value)} style={{ ...inputStyle, marginTop: 4 }}>
                  {!accounts.length && <option value="">Aucun</option>}
                  {accounts.map((a) => <option key={a.id} value={a.name}>{a.name}</option>)}
                </select>
              </div>
            </div>
            <button onClick={submit} style={{
              display: "flex", alignItems: "center", justifyContent: "center", gap: 6, marginTop: 4,
              background: justAdded ? COLOR.emerald : "rgba(201,162,39,0.16)", border: `1px solid ${justAdded ? COLOR.emerald : COLOR.gold}`,
              borderRadius: 8, color: justAdded ? COLOR.bg : COLOR.goldSoft, padding: "10px 0", fontSize: 13, fontWeight: 600, cursor: "pointer",
            }}>
              {justAdded ? <Check size={15} /> : <Plus size={15} />} {justAdded ? "Ajouté" : "Ajouter"}
            </button>
          </div>
        </div>
      )}
      <button onClick={() => setOpen((o) => !o)} className="gl-noprint" style={{
        position: "fixed", bottom: 28, right: 24, zIndex: 100, width: 56, height: 56, borderRadius: "50%",
        background: open ? COLOR.hairline : COLOR.gold, border: "none", color: COLOR.bg, cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 6px 20px rgba(201,162,39,0.4)",
        transition: "transform 0.15s", transform: open ? "rotate(45deg)" : "none",
      }}>
        <Plus size={26} />
      </button>
    </>
  );
}

// ============================================================ END OF PART 6 — App component follows in part 7
// ============================================================
// MAIN APP
// ============================================================
type Tab = "saisie" | "apercu" | "flux" | "comparatif" | "mensuel" | "journalier" | "categories" | "groupes" | "enveloppes" | "budgets" | "simulateur" | "objectif" | "business" | "creances" | "comptes" | "payees" | "recurrences" | "journal" | "export" | "sauvegarde";

const NAV: { section: string; items: { id: Tab; label: string; icon: any }[] }[] = [
  { section: "Saisie rapide", items: [
    { id: "saisie", label: "Saisie du jour", icon: Clock },
  ]},
  { section: "Tableau de bord", items: [
    { id: "apercu", label: "Aperçu", icon: LayoutDashboard },
    { id: "flux", label: "Flux & Calendrier", icon: Workflow },
    { id: "comparatif", label: "Comparatif annuel", icon: BarChart3 },
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

  const allMonths = useMemo(() => {
    const s = new Set(transactions.map((t) => dateToMonthKey(t.date)));
    return Array.from(s).sort((a, b) => monthSortKey(a) - monthSortKey(b));
  }, [transactions]);

  const allCategories = useMemo(() => {
    const s = new Set(transactions.map((t) => t.category));
    return Array.from(s).sort();
  }, [transactions]);

  const defaultFilters: Filters = {
    from: allMonths[0] || "2024_6", to: allMonths[allMonths.length - 1] || "2026_8",
    type: "Tous", group: "Tous", category: "Toutes", search: "", scope: "Tous",
  };
  const [filters, setFilters] = useState<Filters>(defaultFilters);

  useEffect(() => {
    if (allMonths.length && (!filters.from || !filters.to)) setFilters((f) => ({ ...f, from: allMonths[0], to: allMonths[allMonths.length - 1] }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allMonths.length]);

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

  const txWithGroup = useMemo(() => transactions.map((t) => ({
    ...t,
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
      if (filters.search && !t.category.toLowerCase().includes(filters.search.toLowerCase())) return false;
      return true;
    });
  }, [txWithGroup, filters]);

  const allLoaded = txLoaded && groupsLoaded && scopeLoaded && rulesLoaded && loansLoaded && capLoaded && accountsLoaded && budgetsLoaded && goalsLoaded && recurringLoaded;
  if (!allLoaded) {
    return <div style={{ minHeight: "100vh", background: COLOR.bg, color: COLOR.inkMuted, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>Chargement…</div>;
  }

  const lastNW = (() => { const s = liveNetWorthSeries(accounts, transactions); return s[s.length - 1][1]; })();

  return (
    <div style={{ minHeight: "100vh", background: COLOR.bg, color: COLOR.ink, fontFamily: "'Inter', sans-serif", display: "flex" }}>
      <style>{fontImport}</style>

      {/* SIDEBAR */}
      <aside className="gl-noprint" style={{ width: sidebarOpen ? 226 : 0, flexShrink: 0, borderRight: `1px solid ${COLOR.hairline}`, transition: "width 0.2s", overflow: "hidden" }}>
        <div style={{ width: 226, padding: "24px 16px" }}>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, letterSpacing: "0.16em", color: COLOR.gold, textTransform: "uppercase", marginBottom: 6, paddingLeft: 8 }}>XOF</div>
          <h1 style={{ fontFamily: "'Fraunces', serif", fontWeight: 500, fontSize: 24, margin: "0 0 22px 0", paddingLeft: 8 }}>Grand Livre</h1>
          {NAV.map((section) => (
            <div key={section.section} style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 10, color: COLOR.inkMuted, textTransform: "uppercase", letterSpacing: "0.08em", padding: "0 8px 8px 8px" }}>{section.section}</div>
              {section.items.map((item) => {
                const Icon = item.icon; const active = tab === item.id;
                return (
                  <button key={item.id} onClick={() => setTab(item.id)} style={{
                    display: "flex", alignItems: "center", gap: 10, width: "100%", textAlign: "left",
                    background: active ? "rgba(201,162,39,0.1)" : "transparent", border: "none",
                    borderLeft: active ? `2px solid ${COLOR.gold}` : "2px solid transparent",
                    color: active ? COLOR.ink : COLOR.inkMuted, padding: "8px 8px 8px 10px", fontSize: 13, cursor: "pointer",
                    borderRadius: 4, marginBottom: 2, fontFamily: "'Inter', sans-serif",
                  }}>
                    <Icon size={14} /> {item.label}
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      </aside>

      {/* MAIN */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <header style={{ borderBottom: `1px solid ${COLOR.hairline}`, padding: "20px 32px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <button className="gl-noprint" onClick={() => setSidebarOpen((s) => !s)} style={{ background: "transparent", border: `1px solid ${COLOR.hairline}`, borderRadius: 6, color: COLOR.inkMuted, padding: 7, cursor: "pointer", display: "flex" }}>
                <Layers size={14} />
              </button>
              <div>
                <div style={{ fontSize: 12.5, color: COLOR.inkMuted }}>
                  {allMonths.length ? `${monthLabel(allMonths[0])} — ${monthLabel(allMonths[allMonths.length - 1])}` : ""} · {transactions.length} transactions
                </div>
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, letterSpacing: "0.12em", color: COLOR.inkMuted, textTransform: "uppercase" }}>Valeur nette (dernier relevé)</div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 26, fontWeight: 600, color: COLOR.goldSoft }}>{fmt(lastNW)}<span style={{ fontSize: 13, color: COLOR.inkMuted, marginLeft: 6 }}>FCFA</span></div>
            </div>
          </div>
        </header>

        <main className="gl-print-full" style={{ maxWidth: 1180, padding: "24px 32px 60px 32px" }}>
          {tab !== "saisie" && (
            <div className="gl-noprint" style={{ marginBottom: 20 }}>
              <FilterBar filters={filters} setFilters={setFilters} allMonths={allMonths} allCategories={allCategories} onReset={() => setFilters(defaultFilters)} />
            </div>
          )}

          {tab === "apercu" && <ApercuTab filtered={filtered} filters={filters} accounts={accounts} transactions={transactions} />}
          {tab === "flux" && <FluxTab filtered={filtered} />}
          {tab === "comparatif" && <ComparatifTab transactions={transactions} categoryGroups={resolvedGroups} />}
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
          {tab === "export" && <ExportTab filtered={filtered} filters={filters} />}
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
            />
          )}
        </main>

        <footer className="gl-noprint" style={{ borderTop: `1px solid ${COLOR.hairline}`, padding: "18px 32px", textAlign: "center", color: COLOR.inkMuted, fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace" }}>
          Grand Livre · {transactions.length} transactions · {loans.length} créance(s) suivie(s) · données stockées uniquement dans ce navigateur
        </footer>
      </div>
      {tab !== "saisie" && (
        <QuickAddFAB transactions={transactions} setTransactions={setTransactions} accounts={accounts} categoryGroups={resolvedGroups} />
      )}
    </div>
  );
}
