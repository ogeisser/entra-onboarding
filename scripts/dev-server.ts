import { join } from "node:path";
import serveStatic from "serve-static-bun";
import { serve } from "bun";
import index from "@/index.html";
import taskpane from "@/taskpane.html";
import help from "@/help.html";
import auth from "@/auth.html";

const publicFiles = serveStatic("public", { stripFromPathname: "/public" });

const mkcertDir = join(process.env.HOME ?? "", ".vite-plugin-mkcert");

const server = serve({
  port: 3000,

  tls: {
    cert: Bun.file(join(mkcertDir, "cert.pem")),
    key: Bun.file(join(mkcertDir, "dev.pem")),
  },

  // HTML Files
  routes: {
    "/": index,
    "/index.html": index,
    "/taskpane.html": taskpane,
    "/help.html": help,
    "/auth.html": auth
  },

  // Public Files
  fetch(req) {
    if (new URL(req.url).pathname.startsWith("/public"))
        return publicFiles(req);
    return new Response("Not Found", { status: 404 });
  },

  development: { hmr: true, console: true },
});

console.log(`🚀 Server running at ${server.url}`);
