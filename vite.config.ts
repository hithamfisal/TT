// vite.config.ts
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";
import { handleDownloadWorkbook } from "./server/microsoftProxy.js";

function readJsonBody(req: any): Promise<any> {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk: Buffer) => {
      body += chunk.toString("utf8");
      if (body.length > 1024 * 1024) {
        reject(new Error("Request body is too large."));
      }
    });
    req.on("end", () => {
      if (!body.trim()) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

export default defineConfig({
  plugins: [
    react(),
    {
      name: "microsoft-workbook-proxy",
      configureServer(server) {
        server.middlewares.use("/api/download-workbook", async (req, res) => {
          if (req.method !== "POST") {
            res.statusCode = 405;
            res.end("Method Not Allowed");
            return;
          }

          try {
            const body = await readJsonBody(req);
            await handleDownloadWorkbook(
              { body },
              {
                setHeader: (name: string, value: string) =>
                  res.setHeader(name, value),
                status: (code: number) => {
                  res.statusCode = code;
                  return {
                    json: (payload: unknown) => {
                      res.setHeader("Content-Type", "application/json");
                      res.end(JSON.stringify(payload));
                    },
                  };
                },
                send: (payload: Buffer) => res.end(payload),
              }
            );
          } catch (error) {
            res.statusCode = 400;
            res.setHeader("Content-Type", "application/json");
            res.end(
              JSON.stringify({
                error: "Could not read proxy request.",
                details: error instanceof Error ? error.message : String(error),
              })
            );
          }
        });
      },
    },
  ],
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
