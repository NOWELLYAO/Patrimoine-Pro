import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      injectRegister: false, // enregistrement manuel dans main.tsx, avec rechargement forcé sur mise à jour
      // Précache tout ce que Vite génère (JS, CSS, HTML) — l'app peut donc se lancer
      // hors-ligne même si le téléphone redémarre ou que l'onglet est fermé puis
      // rouvert sans connexion, sur demande explicite de l'utilisateur (11/08/2026).
      workbox: {
        globPatterns: ["**/*.{js,css,html,ico,png,svg,woff2}"],
        // Les appels réseau vers Supabase (sync) ne sont volontairement PAS mis en
        // cache ici : hors-ligne, ils échouent simplement et l'app continue avec le
        // localStorage, qui reste la source de vérité locale immédiate.
        navigateFallback: "index.html",
        cleanupOutdatedCaches: true,
        // Ajouté le 13/08/2026 : le nouveau service worker prend le contrôle de la
        // page dès son activation (clientsClaim), au lieu d'attendre la prochaine
        // navigation — indispensable pour que le rechargement forcé côté main.tsx
        // serve bien le NOUVEAU code juste après, et pas l'ancien encore actif.
        clientsClaim: true,
        skipWaiting: true,
      },
      manifest: {
        name: "Grand Livre",
        short_name: "Grand Livre",
        description: "Suivi financier personnel — fonctionne hors-ligne",
        theme_color: "#0d1410",
        background_color: "#0d1410",
        display: "standalone",
        start_url: "/",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
    }),
  ],
});
