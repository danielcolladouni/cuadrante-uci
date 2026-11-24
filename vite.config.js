import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// BASE_PATH lo pone GitHub Actions ("/cuadrante-uci/") al publicar en GitHub Pages.
// En Vercel no existe, así que allí la web sigue sirviéndose desde la raíz ("/").
export default defineConfig({
  plugins: [react()],
  base: process.env.BASE_PATH || "/",
});
