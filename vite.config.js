import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// If you deploy to GitHub Pages under a repo subpath (e.g.
// https://username.github.io/catalog-compliance-checker/), set base to
// "/catalog-compliance-checker/". For Vercel/Netlify (custom domain or
// root-level project), leave it as "/".
export default defineConfig({
  plugins: [react()],
  base: "/",
});
