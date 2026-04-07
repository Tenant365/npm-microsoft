import { requestM365Graph } from "../../common/graph/request";

/**
 * Shared GET helper for Entra directory modules.
 * @param extraHeaders e.g. `ConsistencyLevel: eventual` when using `$search`
 */
export async function entraGraphGet(
  path: string,
  accessToken: string,
  extraHeaders?: Record<string, string>,
): Promise<unknown> {
  return await requestM365Graph(path, accessToken, {
    method: "GET",
    headers: extraHeaders,
  });
}
