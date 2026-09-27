// Minimaler Microsoft-Graph-Client (fetch). Fehler werden mit Graph-Fehlercode und -Text
// geworfen, damit die Prüfberichte aussagekräftig sind.

import { getToken } from "./auth";

const BASE = "https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(
    readonly status: number,
    readonly code: string,
    message: string,
  ) {
    super(`Graph ${status} ${code}: ${message}`);
  }
}

async function request(
  method: string,
  path: string,
  scopes: string[],
  body?: BodyInit,
  contentType?: string,
): Promise<Response> {
  const token = await getToken(scopes);
  const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
  if (contentType) headers["Content-Type"] = contentType;
  const response = await fetch(path.startsWith("http") ? path : BASE + path, { method, headers, body });
  if (!response.ok) {
    let code = "unknown";
    let message = response.statusText;
    try {
      const json = await response.json();
      code = json?.error?.code ?? code;
      message = json?.error?.message ?? message;
    } catch {
      // kein JSON-Fehlerkörper
    }
    throw new GraphError(response.status, code, message);
  }
  return response;
}

export async function graphGet<T = any>(path: string, scopes: string[]): Promise<T> {
  return (await request("GET", path, scopes)).json();
}

export async function graphGetBytes(path: string, scopes: string[]): Promise<Uint8Array> {
  return new Uint8Array(await (await request("GET", path, scopes)).arrayBuffer());
}

export async function graphPut<T = any>(
  path: string,
  scopes: string[],
  body: BodyInit,
  contentType = "application/octet-stream",
): Promise<T> {
  return (await request("PUT", path, scopes, body, contentType)).json();
}

/** Pfadsegmente für Graph-Pfadadressierung (`/root:/a/b:/...`) sicher kodieren. */
export function encodePath(path: string): string {
  return path
    .split("/")
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}
