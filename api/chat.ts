// Fonction serveur Vercel — jamais exécutée dans le navigateur, donc la clé API
// (ANTHROPIC_API_KEY) n'est jamais exposée côté client. Appelée par AssistantTab dans
// App.tsx via fetch("/api/chat"). Sur demande explicite de l'utilisateur (14/08/2026).
//
// Configuration requise sur Vercel : Project Settings → Environment Variables →
// ANTHROPIC_API_KEY = ta clé obtenue sur https://console.anthropic.com
//
// Pour tester en local : `vercel dev` (pas `npm run dev`, qui ne fait tourner que
// Vite et n'exécute pas les fonctions du dossier /api).

// Déclaration minimale de `process` — corrigé le 23/08/2026 après une erreur de build
// Vercel ("Cannot find name 'process'", TS2580). `process` existe bel et bien à
// l'exécution (fonction Node.js côté serveur), le souci est purement une déclaration
// de type manquante, faute du paquet @types/node. Plutôt que de dépendre de son
// installation (hors de portée d'une simple édition de ce fichier), on déclare le strict
// nécessaire ici : `process.env`, la seule chose utilisée dans ce fichier.
declare const process: { env: Record<string, string | undefined> };

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

  const { messages, context, accounts, budgets, recurring, bourse, today, contextWindow } = req.body || {};
  if (!Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: "Aucun message fourni." });
    return;
  }

  const systemPrompt = `Tu es l'assistant financier intégré à "Grand Livre", une app de suivi financier personnel. Tu n'es pas un simple calculateur neutre : tu es un conseiller financier exigeant, honnête et direct — dans l'esprit d'un coach qui dit les choses franchement parce qu'il veut vraiment faire progresser la personne, pas dans l'esprit d'un assistant poli qui évite les sujets qui fâchent. Réponds toujours en français.

RÈGLES DE FOND :
- Base-toi UNIQUEMENT sur les données ci-dessous. Ne jamais inventer un chiffre : si l'information manque, dis-le clairement plutôt que de deviner.
- Important : les transactions fournies couvrent ${contextWindow || "une fenêtre limitée"}, pas forcément tout l'historique. Si la question porte sur une période visiblement hors de cette fenêtre, dis-le explicitement plutôt que de répondre comme si tu avais tout vu.
- Quand on te demande un avis sur des dépenses ou des habitudes, ne te contente PAS de décrire les chiffres — commente-les et critique-les ouvertement quand ils le méritent (dépenses excessives, habitudes répétées et coûteuses, incohérences, argent non classifié...). Sois dur avec les faits, jamais avec la personne : critique le comportement, jamais l'individu.
- Chaque critique doit déboucher sur une recommandation CONCRÈTE et actionnable (pas "fais attention à tes dépenses" mais "réduis X à Y FCFA/mois" ou "arrête complètement Z").
- Chaque recommandation doit être chiffrée avec le GAIN attendu si elle est suivie — sur le mois, et si pertinent projeté sur l'année (ex. "en coupant ça, tu économises 15 000 FCFA/mois, soit 180 000 FCFA sur un an").
- Quand on te demande si une dépense envisagée est raisonnable (montant à venir, achat prévu), calcule toi-même à partir du solde des comptes ci-dessous, des charges fixes récurrentes à venir, et du budget de la catégorie concernée si elle existe — donne une réponse tranchée (oui / non / attends telle date) avec le raisonnement chiffré, pas une réponse évasive.
- Reste rigoureux sur les faits même quand le ton est ferme : jamais de chiffre inventé pour dramatiser.
- Sois concis mais complet : vise 300 à 500 mots dans la plupart des cas, quitte à choisir les 3-4 points les plus importants plutôt que de tout couvrir. Si la question porte sur une longue période ou beaucoup de catégories, hiérarchise et développe seulement ce qui compte vraiment, plutôt que de vouloir tout détailler exhaustivement — une réponse longue et coupée en plein milieu est pire qu'une réponse plus courte mais complète.
- Réponds comme dans une VRAIE conversation, jamais comme un rapport ou un document. Pas de titres, pas de sections numérotées, pas de longue introduction qui reformule la question avant de répondre, pas de conclusion générale qui résume ce qui vient d'être dit. Va droit au fait dès la première phrase. Des puces courtes sont acceptables pour lister plusieurs points, mais la réponse dans son ensemble doit se lire comme quelqu'un qui parle, pas comme un livrable.

# Aujourd'hui
${today || "date inconnue"}

# Comptes (soldes actuels)
${(accounts || []).join("\n") || "aucun"}

# Budgets par catégorie (limite mensuelle)
${(budgets || []).join("\n") || "aucun"}

# Récurrences connues (revenus/charges fixes, avec prochaine échéance)
${(recurring || []).join("\n") || "aucune"}

# Portefeuille Bourse (FCP) — SÉPARÉ des comptes et de la Valeur nette ci-dessus, sauf mention contraire
${(bourse || []).join("\n") || "aucun fonds"}

# Transactions (CSV, séparateur ";", colonnes : date;type;categorie;sous_categorie;montant;compte;beneficiaire;note)
${context || "aucune transaction"}`;

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 55000);
    // Chronométrage + taille du contexte — sur demande explicite de l'utilisateur
    // (23/08/2026), après un timeout systématique sur une question pourtant modeste.
    // Visible dans Vercel → Deployments → Functions → /api/chat → Logs, pour
    // diagnostiquer précisément si un futur ralentissement vient du volume de données
    // envoyées (contextChars élevé) ou d'autre chose (réponse lente malgré peu de
    // données — pointerait vers un souci réseau ou côté API Anthropic, pas notre code).
    const startedAt = Date.now();
    console.log(`[chat] Requête reçue — contexte : ${context?.length || 0} caractères, ${messages.length} message(s), fenêtre : ${contextWindow || "?"}`);

    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        // Corrigé le 23/08/2026, deuxième passe : Claude Sonnet 5 a la réflexion
        // adaptative ACTIVÉE PAR DÉFAUT, et ses tokens de réflexion sont prélevés sur
        // le MÊME budget que max_tokens — avec 1800 (mon premier correctif), tout
        // partait en réflexion interne et il ne restait plus rien pour la vraie
        // réponse texte (stop_reason: max_tokens, aucun bloc "text"). Deux ajustements
        // pour corriger ça sans revenir au problème initial (réponses trop longues) :
        // - effort: "low" réduit la profondeur de réflexion adaptative — inutile
        //   d'avoir un raisonnement poussé pour une question de conversation courante.
        // - max_tokens remonté à 4096 pour laisser de la marge à la réflexion (même
        //   réduite) SANS jamais pouvoir manger tout le budget avant la réponse texte.
        max_tokens: 4096,
        output_config: { effort: "low" },
        system: systemPrompt,
        messages: messages.map((m: any) => ({ role: m.role, content: m.content })),
      }),
      signal: controller.signal,
    });
    clearTimeout(timeout);
    console.log(`[chat] Réponse Anthropic reçue en ${Date.now() - startedAt} ms (statut ${anthropicRes.status})`);

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

    if (!reply) {
      // Réponse HTTP 200 mais aucun bloc texte exploitable — arrivé au moins une fois
      // avec un très gros contexte (14/08/2026). On journalise le corps brut côté
      // serveur (visible dans Vercel → Deployments → Functions → /api/chat → Logs)
      // et on renvoie un diagnostic exploitable au lieu d'un silence total.
      console.error("Réponse Anthropic sans texte exploitable :", JSON.stringify(data));
      res.status(200).json({
        reply: "",
        debug: {
          stop_reason: data.stop_reason,
          usage: data.usage,
          content_types: (data.content || []).map((b: any) => b.type),
        },
      });
      return;
    }

    res.status(200).json({ reply });
  } catch (e: any) {
    const message = e?.name === "AbortError"
      ? "Délai dépassé — réessaie avec une question plus courte ou moins de transactions."
      : "Erreur lors de l'appel à l'API IA.";
    res.status(500).json({ error: message });
  }
}
