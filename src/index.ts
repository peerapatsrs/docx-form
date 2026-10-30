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
    const port = Number(opts.port);
    if (!Number.isInteger(port) || port < 1 || port > 65535) {
      console.error(`Invalid port: ${opts.port}`);
      process.exit(1);
    }
    await startServer({
      host: opts.host,
      port,
      openBrowser: opts.open,
    });
  });

program.parse();
