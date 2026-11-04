#!/usr/bin/env node
import { Command } from "commander";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { startServer } from "./server.js";

function readVersion(): string {
  try {
    const here = dirname(fileURLToPath(import.meta.url));
    const pkgPath = join(here, "../package.json");
    const pkg = JSON.parse(readFileSync(pkgPath, "utf-8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

const program = new Command();
program
  .name("docx-form")
  .description("Local CLI that opens a Web UI to fill Word (.docx) templates")
  .version(readVersion());

program
  .command("start", { isDefault: true })
  .description("Start local server and open the Web UI")
  .option("-p, --port <number>", "Port to listen on", "3850")
  .option("--host <host>", "Host to bind", "127.0.0.1")
  .option("--no-open", "Do not open the browser")
  .action(async (opts: { port: string; host: string; open: boolean }) => {
    const isProduction =
      process.env.NODE_ENV === "production" || process.env.RENDER === "true";

    // Prefer env (e.g. Render PORT) over CLI defaults; keep 3850 locally.
    const portRaw = process.env.PORT?.trim() || opts.port;
    const port = Number(portRaw);
    if (!Number.isInteger(port) || port < 1 || port > 65535) {
      console.error(`Invalid port: ${portRaw}`);
      process.exit(1);
    }

    // HOST when set; otherwise 0.0.0.0 in production; else CLI default (127.0.0.1).
    const host =
      process.env.HOST?.trim() || (isProduction ? "0.0.0.0" : opts.host);

    // Never auto-open a browser on hosted / production deploys.
    const openBrowser = isProduction ? false : opts.open;

    await startServer({
      host,
      port,
      openBrowser,
    });
  });

program.parse();
