import { createHash } from "node:crypto";
import { writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { resolveAttachmentDir } from "./paths.js";

export interface SavedAttachment {
  path: string;
  size: number;
  sha256: string;
}

/** Reduce an attacker-controlled attachment name to a safe, single path segment. */
export function safeFileName(name: string): string {
  const leaf = basename(name.replace(/\\/g, "/"));
  const cleaned = leaf
    .replace(/[^A-Za-z0-9._ -]/g, "_")
    .replace(/^[. ]+/, "")
    .trim()
    .slice(-150);
  return cleaned.length > 0 ? cleaned : "attachment";
}

/**
 * Write attachment bytes under the attachment directory. The location is decided
 * here, never by the caller, so a tool call cannot write elsewhere on disk. The
 * id-derived prefix keeps same-named attachments from different items apart.
 */
export function saveAttachment(
  messageId: string,
  attachmentId: string,
  name: string,
  bytes: Buffer,
): SavedAttachment {
  const prefix = createHash("sha256")
    .update(`${messageId}\0${attachmentId}`)
    .digest("hex")
    .slice(0, 12);
  const path = join(resolveAttachmentDir(), `${prefix}-${safeFileName(name)}`);
  writeFileSync(path, bytes, { mode: 0o600 });
  return {
    path,
    size: bytes.length,
    sha256: createHash("sha256").update(bytes).digest("hex"),
  };
}
