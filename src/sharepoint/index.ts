import { M365Authentication } from "../core/auth";
import { MS365Scopes } from "../core/auth";

export interface M365SharePointSite {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  [key: string]: unknown;
}

export interface M365SharePointSiteMetadata {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  [key: string]: unknown;
}

export interface GraphSharePointSitesResponse {
  "@odata.context"?: string;
  value: M365SharePointSite[];
}

export const isGraphSharePointSitesResponse = (
  value: unknown,
): value is GraphSharePointSitesResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphSharePointSitesResponse).value) &&
    typeof (value as GraphSharePointSitesResponse)["@odata.context"] ===
      "string"
  );
};

const isM365SharePointSite = (value: unknown): value is M365SharePointSite => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365SharePointSite).id === "string";
};

const isM365SharePointSiteMetadata = (
  value: unknown,
): value is M365SharePointSiteMetadata => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365SharePointSiteMetadata).id === "string";
};
export class SharePointClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async graphRequest(
    path: string,
    accessToken: string,
  ): Promise<unknown> {
    const response = await fetch(`https://graph.microsoft.com/v1.0/${path}`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
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

  private async requestSharePointSite(
    siteId: string,
    filter?: string,
  ): Promise<unknown> {
    return await this.graphRequest(
      `sites/${siteId}${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
    );
  }

  private async requestSharePointSites(filter?: string): Promise<unknown> {
    return await this.graphRequest(
      `sites${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
    );
  }

  private async requestSharePointSiteMetadata(
    siteId: string,
  ): Promise<unknown> {
    return await this.graphRequest(
      `sites/${siteId}/microsoft.graph.site?$select=id,name,displayName,webUrl`,
      await this.getAccessToken(),
    );
  }

  public async getSharePointSiteById(
    siteId: string,
  ): Promise<M365SharePointSite> {
    return await this.getSharePointSite(siteId);
  }

  public async getSharePointSite(siteId: string): Promise<M365SharePointSite> {
    const data = await this.requestSharePointSite(siteId);
    if (!isM365SharePointSite(data)) {
      throw new Error(
        `Microsoft Graph sites/${siteId} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getAllSharePointSites(
    search?: string,
  ): Promise<M365SharePointSite[]> {
    const data = await this.requestSharePointSites(
      search ? `?$search=${search}` : "",
    );
    if (!isGraphSharePointSitesResponse(data)) {
      throw new Error(`Microsoft Graph sites response has an invalid format.`);
    }
    return data.value;
  }

  public async getSharePointSitesBySearch(
    search?: string,
  ): Promise<M365SharePointSite[]> {
    return await this.getAllSharePointSites(search);
  }

  public async getSharePointAllSitesMetadata(): Promise<
    M365SharePointSiteMetadata[]
  > {
    const data = await this.requestSharePointSites();
    if (!isGraphSharePointSitesResponse(data)) {
      throw new Error(`Microsoft Graph sites response has an invalid format.`);
    }
    return data.value.map((site: M365SharePointSite) => ({
      id: site.id,
      name: site.name,
      displayName: site.displayName,
      webUrl: site.webUrl,
    }));
  }
}

export const createSharePointClient = (
  authentication: M365Authentication,
): SharePointClient => {
  return new SharePointClient(authentication);
};
