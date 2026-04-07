export type M365GraphMethod = "GET" | "POST" | "PUT" | "DELETE";

export interface M365GraphRequestOptions {
  method?: M365GraphMethod;
  headers?: Record<string, string>;
  body?: unknown;
}

/**
 * Central Microsoft Graph transport helper.
 * Keeps HTTP behavior consistent across all domain clients.
 */
export async function requestM365Graph(
  path: string,
  accessToken: string,
  options: M365GraphRequestOptions = {},
): Promise<unknown> {
  const response = await fetch(`https://graph.microsoft.com/v1.0/${path}`, {
    method: options.method ?? "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...options.headers,
    },
    body:
      options.body === undefined ? undefined : JSON.stringify(options.body),
  });

  const data: unknown = await response.json().catch(async () => {
    const text = await response.text().catch(() => "");
    return text ? { raw: text } : undefined;
  });

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph ${path} request failed: ${response.status} ${response.statusText} - ${JSON.stringify(data)}`,
    );
  }

  return data;
}
