/** POSTs JSON with a timeout; throws on non-2xx. Returns the response status. */
export async function postJson(url: string, body: unknown, timeoutMs = 30_000): Promise<number> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`POST ${res.status} ${res.statusText}: ${text.slice(0, 500)}`);
    }
    return res.status;
  } finally {
    clearTimeout(timer);
  }
}

export function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable ${name}`);
  }
  return value;
}
