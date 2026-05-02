import { appendFileSync, openSync, closeSync, fchmodSync, existsSync } from "node:fs";
import { createHash } from "node:crypto";
import { auditLogPath } from "../util/paths.js";

export interface AuditEntry {
  tool: string;
  args: Record<string, unknown>;
  result: "ok" | "error";
  durationMs: number;
  error?: string;
}

const SENSITIVE_BODY_KEYS = new Set(["body", "html", "content"]);

function redactArgs(args: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(args)) {
    if (SENSITIVE_BODY_KEYS.has(k) && typeof v === "string") {
      out[k] = {
        kind: "redacted-body",
        sha256: createHash("sha256").update(v).digest("hex").slice(0, 16),
        length: v.length,
      };
    } else {
      out[k] = v;
    }
  }
  return out;
}

function ensurePerms(path: string): void {
  if (!existsSync(path)) return;
  try {
    const fd = openSync(path, "r+");
    try {
      fchmodSync(fd, 0o600);
    } finally {
      closeSync(fd);
    }
  } catch {
    // best-effort
  }
}

export function appendAudit(entry: AuditEntry): void {
  const line = JSON.stringify({
    ts: new Date().toISOString(),
    tool: entry.tool,
    args: redactArgs(entry.args),
    result: entry.result,
    durationMs: entry.durationMs,
    ...(entry.error ? { error: entry.error } : {}),
  });
  const path = auditLogPath();
  try {
    appendFileSync(path, line + "\n", { mode: 0o600 });
    ensurePerms(path);
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    process.stderr.write(`[outlook-mcp] failed to write audit log: ${reason}\n`);
  }
}
