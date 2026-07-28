import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { copyFileSync, existsSync } from 'fs';
import { fileURLToPath } from 'url';

const rootDir = fileURLToPath(new URL('.', import.meta.url));

// Tout est à plat à la racine du dépôt (pas de dossier public/), donc on copie
// uniquement les icônes nécessaires dans le résultat de build — pas tout le
// dossier racine (qui contiendrait aussi .git, dangereux à copier tel quel).
const iconFiles = [
  'apple-touch-icon.png',
  'apple-touch-icon-152.png',
  'apple-touch-icon-167.png',
  'icon-512.png',
  'icon-1024.png',
];

export default defineConfig({
  plugins: [
    react(),
    {
      name: 'copy-icons',
      closeBundle() {
        for (const f of iconFiles) {
          const src = rootDir + f;
          if (existsSync(src)) copyFileSync(src, rootDir + 'dist/' + f);
        }
      },
    },
  ],
  publicDir: false,
  build: {
    // Nécessaire pour les imports dynamiques (Supabase, jsPDF, html2canvas) : la cible
    // par défaut de Vite (~ES2020) ne supporte pas "top-level await", qu'esbuild génère
    // parfois lui-même dans les fragments de code créés pour ces imports dynamiques.
    target: 'esnext',
  },
});
