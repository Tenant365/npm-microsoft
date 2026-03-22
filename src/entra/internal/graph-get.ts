/**
 * Shared GET helper for Microsoft Graph v1.0 (Entra / directory resources).
 * @param extraHeaders e.g. `ConsistencyLevel: eventual` when using `$search`
 */
export async function entraGraphGet(
  path: string,
  accessToken: string,
  extraHeaders?: Record<string, string>,
): Promise<unknown> {
  const response = await fetch(`https://graph.microsoft.com/v1.0/${path}`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...extraHeaders,
    },
  });

  const data: unknown = await response.json().catch(async () => {
    const text = await response.text().catch(() => "");
    return { error: text };
  });

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph ${path} request failed: ${response.status} ${response.statusText} - ${JSON.stringify(data)}`,
    );
  }

  return data;
}
