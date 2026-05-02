import type { TokenProvider } from "../auth/token.js";

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
  ) {
    super(message);
  }
}

export interface RequestOptions {
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  query?: Record<string, string | number | boolean | undefined>;
  body?: unknown;
  headers?: Record<string, string>;
  expectJson?: boolean;
}

const BASE = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES = 3;

export class GraphClient {
  constructor(private readonly tokens: TokenProvider) {}

  async request<T = unknown>(path: string, opts: RequestOptions = {}): Promise<T> {
    const method = opts.method ?? "GET";
    const url = buildUrl(path, opts.query);
    const expectJson = opts.expectJson ?? method !== "DELETE";

    let attempt = 0;
    // eslint-disable-next-line no-constant-condition
    while (true) {
      const token = await this.tokens.getAccessToken();
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...(opts.headers ?? {}),
      };
      let body: BodyInit | undefined;
      if (opts.body !== undefined) {
        headers["Content-Type"] ??= "application/json";
        body = typeof opts.body === "string" ? opts.body : JSON.stringify(opts.body);
      }

      const res = await fetch(url, { method, headers, body });

      if (res.status === 429 || (res.status >= 500 && res.status <= 599)) {
        if (attempt < MAX_RETRIES) {
          attempt++;
          const retryAfter = parseRetryAfter(res.headers.get("Retry-After"));
          const backoff = retryAfter ?? Math.min(2_000 * 2 ** (attempt - 1), 8_000);
          await delay(backoff);
          continue;
        }
      }

      if (!res.ok) {
        const text = await res.text().catch(() => "");
        const parsed = safeJson(text);
        const msg =
          parsed?.error?.message ||
          (text ? text.slice(0, 500) : `HTTP ${res.status} ${res.statusText}`);
        throw new GraphError(msg, res.status, parsed?.error?.code);
      }

      if (!expectJson || res.status === 204) return undefined as T;
      const text = await res.text();
      if (text.length === 0) return undefined as T;
      return JSON.parse(text) as T;
    }
  }
}

function buildUrl(path: string, query?: RequestOptions["query"]): string {
  const url = new URL(path.startsWith("http") ? path : `${BASE}${path}`);
  if (query) {
    for (const [k, v] of Object.entries(query)) {
      if (v === undefined) continue;
      url.searchParams.set(k, String(v));
    }
  }
  return url.toString();
}

function parseRetryAfter(value: string | null): number | null {
  if (!value) return null;
  const seconds = Number(value);
  if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
  const date = Date.parse(value);
  if (Number.isFinite(date)) return Math.max(0, date - Date.now());
  return null;
}

function safeJson(text: string): { error?: { message?: string; code?: string } } | null {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
