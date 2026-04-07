import { M365Authentication, MS365Scopes } from "../../core/auth";
import { requestM365Graph, M365GraphRequestOptions } from "./request";

/**
 * Enterprise-style base class for Microsoft Graph domain clients.
 * Ensures consistent token resolution and request handling.
 */
export abstract class M365GraphClientBase {
  protected constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(scope: string = MS365Scopes.DEFAULT): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(scope);
    return accessToken.token;
  }

  protected async graphRequest(
    path: string,
    options: M365GraphRequestOptions = {},
  ): Promise<unknown> {
    const token = await this.getAccessToken();
    return await requestM365Graph(path, token, options);
  }
}
