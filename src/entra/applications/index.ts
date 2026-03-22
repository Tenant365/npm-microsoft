import { M365Authentication } from "../../core/auth";
import { MS365Scopes } from "../../core/auth";
import { entraGraphGet } from "../internal/graph-get";

export interface M365EntraApplication {
  id: string;
  appId?: string;
  displayName?: string;
  [key: string]: unknown;
}

export interface GraphEntraApplicationsResponse {
  "@odata.context"?: string;
  value: M365EntraApplication[];
}

export const isGraphEntraApplicationsResponse = (
  value: unknown,
): value is GraphEntraApplicationsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphEntraApplicationsResponse).value) &&
    typeof (value as GraphEntraApplicationsResponse)["@odata.context"] ===
      "string"
  );
};

const isM365EntraApplication = (
  value: unknown,
): value is M365EntraApplication => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365EntraApplication).id === "string";
};

export class EntraApplicationsClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async requestApplication(objectId: string): Promise<unknown> {
    return await entraGraphGet(
      `applications/${objectId}`,
      await this.getAccessToken(),
    );
  }

  private async requestApplications(filter?: string): Promise<unknown> {
    return await entraGraphGet(
      `applications${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
    );
  }

  /**
   * Application **object id** in the directory (not the `appId` / client id).
   */
  public async getApplication(objectId: string): Promise<M365EntraApplication> {
    const data = await this.requestApplication(objectId);
    if (!isM365EntraApplication(data)) {
      throw new Error(
        `Microsoft Graph applications/${objectId} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getApplicationById(
    objectId: string,
  ): Promise<M365EntraApplication> {
    return await this.getApplication(objectId);
  }

  /**
   * @param queryParams OData query without leading `?`, e.g. `$top=100` or `$filter=appId eq '…'`.
   */
  public async getAllApplications(
    queryParams?: string,
  ): Promise<M365EntraApplication[]> {
    const normalized =
      queryParams?.startsWith("?") === true
        ? queryParams.slice(1)
        : queryParams;
    const data = await this.requestApplications(normalized);
    if (!isGraphEntraApplicationsResponse(data)) {
      throw new Error(
        "Microsoft Graph applications response has an invalid format.",
      );
    }
    return data.value;
  }
}

export const createEntraApplicationsClient = (
  authentication: M365Authentication,
): EntraApplicationsClient => {
  return new EntraApplicationsClient(authentication);
};
