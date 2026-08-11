import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// Corrigé le 11/08/2026 : "autoUpdate" met à jour le service worker en arrière-plan,
// mais une page déjà ouverte continue de tourner avec l'ANCIEN code JS en mémoire tant
// qu'elle n'est pas rechargée — ce qui a probablement empêché plusieurs correctifs
// récents de prendre effet malgré un redéploiement confirmé. On force maintenant un
// rechargement automatique dès qu'une nouvelle version est détectée.
if ("serviceWorker" in navigator) {
  import("virtual:pwa-register").then(({ registerSW }) => {
    registerSW({
      immediate: true,
      onNeedRefresh() { window.location.reload(); },
    });
  }).catch(() => {});
}
