// vite.config.ts
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "src"),
    },
  },
  // Hard-code the endpoint here if you don't use .env files
  define: {
    "process.env.VITE_ANALYTICS_ENDPOINT": JSON.stringify("https://your-api-url-here.com"),
  },
  server: {
    port: 3000,
    strictPort: false,
  },
  // ... rest of your config
});