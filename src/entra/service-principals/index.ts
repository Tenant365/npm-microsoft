import { M365Authentication } from "../../core/auth";
import { MS365Scopes } from "../../core/auth";
import { entraGraphGet } from "../internal/graph-get";

export interface M365EntraServicePrincipal {
  id: string;
  appId?: string;
  displayName?: string;
  servicePrincipalType?: string;
  [key: string]: unknown;
}

export interface GraphEntraServicePrincipalsResponse {
  "@odata.context"?: string;
  value: M365EntraServicePrincipal[];
}

export const isGraphEntraServicePrincipalsResponse = (
  value: unknown,
): value is GraphEntraServicePrincipalsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphEntraServicePrincipalsResponse).value) &&
    typeof (value as GraphEntraServicePrincipalsResponse)["@odata.context"] ===
      "string"
  );
};

const isM365EntraServicePrincipal = (
  value: unknown,
): value is M365EntraServicePrincipal => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365EntraServicePrincipal).id === "string";
};

export class EntraServicePrincipalsClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async requestServicePrincipal(id: string): Promise<unknown> {
    return await entraGraphGet(
      `servicePrincipals/${id}`,
      await this.getAccessToken(),
    );
  }

  private async requestServicePrincipals(filter?: string): Promise<unknown> {
    return await entraGraphGet(
      `servicePrincipals${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
    );
  }

  public async getServicePrincipal(
    id: string,
  ): Promise<M365EntraServicePrincipal> {
    const data = await this.requestServicePrincipal(id);
    if (!isM365EntraServicePrincipal(data)) {
      throw new Error(
        `Microsoft Graph servicePrincipals/${id} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getServicePrincipalById(
    id: string,
  ): Promise<M365EntraServicePrincipal> {
    return await this.getServicePrincipal(id);
  }

  /**
   * @param queryParams OData query without leading `?`, e.g. `$filter=appId eq '…'`.
   */
  public async getAllServicePrincipals(
    queryParams?: string,
  ): Promise<M365EntraServicePrincipal[]> {
    const normalized =
      queryParams?.startsWith("?") === true
        ? queryParams.slice(1)
        : queryParams;
    const data = await this.requestServicePrincipals(normalized);
    if (!isGraphEntraServicePrincipalsResponse(data)) {
      throw new Error(
        "Microsoft Graph servicePrincipals response has an invalid format.",
      );
    }
    return data.value;
  }
}

export const createEntraServicePrincipalsClient = (
  authentication: M365Authentication,
): EntraServicePrincipalsClient => {
  return new EntraServicePrincipalsClient(authentication);
};
