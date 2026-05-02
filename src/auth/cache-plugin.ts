import { promises as fs, existsSync } from "node:fs";
import { dirname } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";

export class FileCachePlugin implements ICachePlugin {
  constructor(private readonly path: string) {}

  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!existsSync(this.path)) return;
    try {
      const data = await fs.readFile(this.path, "utf8");
      if (data.length > 0) ctx.tokenCache.deserialize(data);
    } catch (err) {
      const reason = err instanceof Error ? err.message : String(err);
      process.stderr.write(
        `[outlook-mcp] failed to read token cache at ${this.path}: ${reason}\n`,
      );
    }
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!ctx.cacheHasChanged) return;
    const dir = dirname(this.path);
    await fs.mkdir(dir, { recursive: true, mode: 0o700 });
    const tmp = `${this.path}.${process.pid}.${Date.now()}.tmp`;
    const serialized = ctx.tokenCache.serialize();
    await fs.writeFile(tmp, serialized, { mode: 0o600 });
    await fs.rename(tmp, this.path);
    try {
      await fs.chmod(this.path, 0o600);
    } catch {
      // best-effort
    }
  }
}
