import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// --------------------------------------------------------------------------
// Mise à jour du service worker — entièrement automatique, sans jamais
// demander à l'utilisateur de désinscrire quoi que ce soit à la main.
//
// Historique du problème (corrigé le 13/08/2026, complète le correctif du
// 11/08/2026) : "autoUpdate" + onNeedRefresh() ne suffisaient pas en
// pratique, pour deux raisons qui se cumulent :
//   1) Le navigateur ne revérifie sw.js que toutes les ~24h par défaut (et
//      encore, seulement à la navigation) — sur un onglet resté ouvert ou
//      une PWA iOS relancée depuis l'arrière-plan, un déploiement pouvait
//      rester invisible pendant des heures.
//   2) sw.js pouvait être servi par Vercel avec une réponse mise en cache
//      HTTP par le navigateur, donc même une revérification "voyait"
//      encore l'ancien fichier.
//
// La solution ci-dessous (pattern officiel recommandé par vite-plugin-pwa
// pour les mises à jour périodiques) contourne les deux : on va chercher
// sw.js nous-mêmes avec cache: "no-store" à intervalle rapproché ET à
// chaque retour au premier plan, puis on force registration.update(). Si
// une nouvelle version est détectée (onNeedRefresh), on désinscrit tous les
// service workers, on vide tous les caches, puis on recharge — exactement
// l'action manuelle que l'utilisateur faisait dans DevTools, automatisée.
// --------------------------------------------------------------------------

const UPDATE_CHECK_INTERVAL_MS = 60 * 1000; // vérifie toutes les 60s pendant que l'app est ouverte

if ("serviceWorker" in navigator) {
  const nukeServiceWorkersAndReload = async () => {
    try {
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(regs.map((r) => r.unregister()));
    } catch {}
    try {
      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(keys.map((k) => caches.delete(k)));
      }
    } catch {}
    window.location.reload();
  };

  import("virtual:pwa-register").then(({ registerSW }) => {
    registerSW({
      immediate: true,
      onNeedRefresh() {
        nukeServiceWorkersAndReload();
      },
      onRegisteredSW(swUrl, registration) {
        if (!registration) return;

        const checkForUpdate = async () => {
          try {
            if (registration.installing) return; // une mise à jour est déjà en cours d'installation
            if ("onLine" in navigator && !navigator.onLine) return;

            // Contourne le cache HTTP du navigateur pour être sûr de voir le
            // VRAI sw.js actuellement déployé, pas une copie mise en cache.
            const resp = await fetch(swUrl, {
              cache: "no-store",
              headers: { "cache-control": "no-cache" },
            });
            if (resp?.status === 200) await registration.update();
          } catch {
            // Hors-ligne ou requête échouée — on retentera au prochain cycle.
          }
        };

        setInterval(checkForUpdate, UPDATE_CHECK_INTERVAL_MS);
        document.addEventListener("visibilitychange", () => {
          if (document.visibilityState === "visible") checkForUpdate();
        });
        window.addEventListener("focus", checkForUpdate);
        checkForUpdate();
      },
    });
  }).catch(() => {});
}
