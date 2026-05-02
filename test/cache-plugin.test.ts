import { describe, it, expect, beforeEach } from "vitest";
import { mkdtempSync, statSync, readFileSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { FileCachePlugin } from "../src/auth/cache-plugin.js";

function mockCtx(opts: { initial?: string; produce?: string; cacheHasChanged: boolean }) {
  let serialized = opts.initial ?? "";
  return {
    cacheHasChanged: opts.cacheHasChanged,
    tokenCache: {
      deserialize(data: string) {
        serialized = data;
      },
      serialize() {
        return opts.produce ?? serialized;
      },
    },
    get current(): string {
      return serialized;
    },
  };
}

describe("FileCachePlugin", () => {
  let dir: string;
  let path: string;

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), "outlook-mcp-cache-"));
    path = join(dir, "cache.json");
  });

  it("does nothing on beforeCacheAccess when file is missing", async () => {
    const plugin = new FileCachePlugin(path);
    const ctx = mockCtx({ cacheHasChanged: false });
    await plugin.beforeCacheAccess(ctx as any);
    expect(ctx.current).toBe("");
  });

  it("writes the cache when cacheHasChanged is true, with mode 0600", async () => {
    const plugin = new FileCachePlugin(path);
    const ctx = mockCtx({ produce: '{"hello":"world"}', cacheHasChanged: true });
    await plugin.afterCacheAccess(ctx as any);
    expect(existsSync(path)).toBe(true);
    expect(readFileSync(path, "utf8")).toBe('{"hello":"world"}');
    const mode = statSync(path).mode & 0o777;
    expect(mode).toBe(0o600);
  });

  it("does not write when cacheHasChanged is false", async () => {
    const plugin = new FileCachePlugin(path);
    const ctx = mockCtx({ produce: "ignored", cacheHasChanged: false });
    await plugin.afterCacheAccess(ctx as any);
    expect(existsSync(path)).toBe(false);
  });

  it("roundtrips: read what was previously written", async () => {
    const plugin = new FileCachePlugin(path);
    await plugin.afterCacheAccess(mockCtx({ produce: '{"a":1}', cacheHasChanged: true }) as any);

    const readCtx = mockCtx({ cacheHasChanged: false });
    await plugin.beforeCacheAccess(readCtx as any);
    expect(readCtx.current).toBe('{"a":1}');
  });
});
