import { cloudflare } from "@cloudflare/vite-plugin";
import tailwindcss from "@tailwindcss/vite";
import { tanstackStart } from "@tanstack/react-start/plugin/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  server: {
    host: "0.0.0.0",
    port: 3001,
  },
  resolve: {
    tsconfigPaths: true,
  },
  plugins: [
    {
      name: "onlyoffice-resource-timing-headers",
      configureServer(server) {
        server.middlewares.use((request, response, next) => {
          if (
            request.url?.startsWith("/packages/onlyoffice/") ||
            request.url?.startsWith("/onlyoffice/runtime/")
          ) {
            const setHeader = response.setHeader.bind(response);
            const appendHeader = response.appendHeader.bind(response);
            response.removeHeader("Vary");
            response.setHeader = ((
              name: string,
              value: string | number | readonly string[],
            ) => {
              const lowerName = name.toLowerCase();
              if (lowerName === "vary") {
                const values = String(value)
                  .split(",")
                  .map((entry) => entry.trim())
                  .filter(
                    (entry) => entry && entry.toLowerCase() !== "origin",
                  );
                if (values.length === 0) {
                  response.removeHeader(name);
                  return response;
                }
                return setHeader(name, values.join(", "));
              }
              return setHeader(
                name,
                lowerName === "cache-control"
                  ? "public, max-age=31536000, immutable"
                  : value,
              );
            }) as typeof response.setHeader;
            response.appendHeader = ((
              name: string,
              value: string | readonly string[],
            ) => {
              if (name.toLowerCase() !== "vary") {
                return appendHeader(name, value);
              }
              const values = (Array.isArray(value) ? value : [value])
                .flatMap((entry) => String(entry).split(","))
                .map((entry) => entry.trim())
                .filter(
                  (entry) => entry && entry.toLowerCase() !== "origin",
                );
              if (values.length === 0) return response;
              return appendHeader(name, values);
            }) as typeof response.appendHeader;
            response.setHeader("Access-Control-Allow-Origin", "*");
            response.setHeader("Cross-Origin-Resource-Policy", "cross-origin");
            response.setHeader("Timing-Allow-Origin", "*");
            response.setHeader(
              "Cache-Control",
              "public, max-age=31536000, immutable",
            );
          }
          next();
        });
      },
    },
    tailwindcss(),
    cloudflare({ viteEnvironment: { name: "ssr" } }),
    tanstackStart({
      srcDirectory: "src",
      router: {
        routesDirectory: "app",
      },
    }),
    react(),
  ],
});
