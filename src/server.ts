import { serve } from "@hono/node-server";
import { serveStatic } from "@hono/node-server/serve-static";
import { Hono } from "hono";
import open from "open";
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { createApi } from "./api.js";

export type StartServerOptions = {
  host: string;
  port: number;
  openBrowser: boolean;
};

function resolveWebDist(): string | null {
  const here = dirname(fileURLToPath(import.meta.url));
  const candidates = [
    resolve(here, "web"), // dist/web after build copy
    resolve(here, "../web"), // published layout: package/web next to dist
    resolve(process.cwd(), "web"),
  ];
  for (const dir of candidates) {
    if (existsSync(join(dir, "index.html"))) {
      return dir;
    }
  }
  return null;
}

function readPackageVersion(): string {
  try {
    const here = dirname(fileURLToPath(import.meta.url));
    const pkgPath = join(here, "../package.json");
    const pkg = JSON.parse(readFileSync(pkgPath, "utf-8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

export async function startServer(options: StartServerOptions): Promise<void> {
  const startedAt = new Date();
  const version = readPackageVersion();
  const app = new Hono();
  const webDist = resolveWebDist();

  app.route("/api", createApi({ startedAt, version }));

  if (webDist) {
    app.use(
      "/*",
      serveStatic({
        root: webDist,
      }),
    );
    app.get("*", (c) => {
      const html = readFileSync(join(webDist, "index.html"), "utf-8");
      return c.html(html);
    });
  } else {
    app.get("/", (c) =>
      c.html(`<!doctype html>
<html lang="th">
  <head><meta charset="utf-8" /><title>@peerapat.srs/docx-form</title></head>
  <body style="font-family: system-ui; padding: 2rem;">
    <h1>@peerapat.srs/docx-form API is running</h1>
    <p>ยังไม่มี Web UI — รัน <code>npm run build</code> แล้วเริ่มใหม่</p>
    <p><a href="/api/health">/api/health</a></p>
  </body>
</html>`),
    );
  }

  const url = `http://${options.host}:${options.port}`;

  await new Promise<void>((resolveListen) => {
    serve(
      {
        fetch: app.fetch,
        hostname: options.host,
        port: options.port,
      },
      (info) => {
        console.log(`docx-form listening on http://${info.address}:${info.port}`);
        if (!webDist) {
          console.log("Tip: run `npm run build` to include the Web UI");
        }
        resolveListen();
      },
    );
  });

  if (options.openBrowser) {
    await open(url);
  }
}
