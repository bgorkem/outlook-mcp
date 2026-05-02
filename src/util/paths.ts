import { homedir } from "node:os";
import { mkdirSync } from "node:fs";
import { join } from "node:path";

export function resolveCacheDir(): string {
  const fromEnv = process.env.OUTLOOK_MCP_CACHE_DIR?.trim();
  const dir = fromEnv && fromEnv.length > 0 ? fromEnv : join(homedir(), ".outlook-mcp");
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  return dir;
}

export function cacheFilePath(): string {
  return join(resolveCacheDir(), "cache.json");
}

export function auditLogPath(): string {
  return join(resolveCacheDir(), "audit.log");
}
