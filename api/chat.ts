// Fonction serveur Vercel — jamais exécutée dans le navigateur, donc la clé API
// (ANTHROPIC_API_KEY) n'est jamais exposée côté client. Appelée par AssistantTab dans
// App.tsx via fetch("/api/chat"). Sur demande explicite de l'utilisateur (14/08/2026).
//
// Configuration requise sur Vercel : Project Settings → Environment Variables →
// ANTHROPIC_API_KEY = ta clé obtenue sur https://console.anthropic.com
//
// Pour tester en local : `vercel dev` (pas `npm run dev`, qui ne fait tourner que
// Vite et n'exécute pas les fonctions du dossier /api).

export default async function handler(req: any, res: any) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Méthode non autorisée." });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: "ANTHROPIC_API_KEY n'est pas configurée sur le serveur (Vercel → Settings → Environment Variables)." });
    return;
  }

  const { messages, context, accounts, budgets, recurring, today } = req.body || {};
  if (!Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: "Aucun message fourni." });
    return;
  }

  const systemPrompt = `Tu es l'assistant financier intégré à "Grand Livre", une app de suivi financier personnel. Tu n'es pas un simple calculateur neutre : tu es un conseiller financier exigeant, honnête et direct — dans l'esprit d'un coach qui dit les choses franchement parce qu'il veut vraiment faire progresser la personne, pas dans l'esprit d'un assistant poli qui évite les sujets qui fâchent. Réponds toujours en français.

RÈGLES DE FOND :
- Base-toi UNIQUEMENT sur les données ci-dessous. Ne jamais inventer un chiffre : si l'information manque, dis-le clairement plutôt que de deviner.
- Quand on te demande un avis sur des dépenses ou des habitudes, ne te contente PAS de décrire les chiffres — commente-les et critique-les ouvertement quand ils le méritent (dépenses excessives, habitudes répétées et coûteuses, incohérences, argent non classifié...). Sois dur avec les faits, jamais avec la personne : critique le comportement, jamais l'individu.
- Chaque critique doit déboucher sur une recommandation CONCRÈTE et actionnable (pas "fais attention à tes dépenses" mais "réduis X à Y FCFA/mois" ou "arrête complètement Z").
- Chaque recommandation doit être chiffrée avec le GAIN attendu si elle est suivie — sur le mois, et si pertinent projeté sur l'année (ex. "en coupant ça, tu économises 15 000 FCFA/mois, soit 180 000 FCFA sur un an").
- Quand on te demande si une dépense envisagée est raisonnable (montant à venir, achat prévu), calcule toi-même à partir du solde des comptes ci-dessous, des charges fixes récurrentes à venir, et du budget de la catégorie concernée si elle existe — donne une réponse tranchée (oui / non / attends telle date) avec le raisonnement chiffré, pas une réponse évasive.
- Reste rigoureux sur les faits même quand le ton est ferme : jamais de chiffre inventé pour dramatiser.
- Sois concis mais complet — pas de remplissage, pas de généralités creuses.

# Aujourd'hui
${today || "date inconnue"}

# Comptes (soldes actuels)
${(accounts || []).join("\n") || "aucun"}

# Budgets par catégorie (limite mensuelle)
${(budgets || []).join("\n") || "aucun"}

# Récurrences connues (revenus/charges fixes, avec prochaine échéance)
${(recurring || []).join("\n") || "aucune"}

# Transactions (CSV, séparateur ";", colonnes : date;type;categorie;sous_categorie;montant;compte;beneficiaire;note)
${context || "aucune transaction"}`;

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 55000);

    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 1500,
        system: systemPrompt,
        messages: messages.map((m: any) => ({ role: m.role, content: m.content })),
      }),
      signal: controller.signal,
    });
    clearTimeout(timeout);

    if (!anthropicRes.ok) {
      const errText = await anthropicRes.text();
      res.status(anthropicRes.status).json({ error: `Erreur API IA (${anthropicRes.status}) : ${errText}` });
      return;
    }

    const data = await anthropicRes.json();
    const reply = (data.content || [])
      .filter((block: any) => block.type === "text")
      .map((block: any) => block.text)
      .join("\n");

    res.status(200).json({ reply });
  } catch (e: any) {
    const message = e?.name === "AbortError"
      ? "Délai dépassé — réessaie avec une question plus courte ou moins de transactions."
      : "Erreur lors de l'appel à l'API IA.";
    res.status(500).json({ error: message });
  }
}
