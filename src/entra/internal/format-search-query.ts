/**
 * Builds the value for the `$search` query parameter for Microsoft Graph.
 * Clauses like `displayName:Panama` must appear as `"displayName:Panama"` in the request.
 * The result is URL-encoded for use after `$search=`.
 */
export function encodeGraphSearchParameter(search: string): string {
  const trimmed = search.trim();
  if (trimmed.length === 0) {
    return encodeURIComponent('""');
  }

  // Caller already passed a full Graph search literal including outer quotes.
  if (trimmed.startsWith('"') && trimmed.endsWith('"') && trimmed.length >= 2) {
    return encodeURIComponent(trimmed);
  }

  return encodeURIComponent(`"${trimmed}"`);
}
