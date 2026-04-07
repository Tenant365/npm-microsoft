import { M365Authentication } from "../core/auth";
import { M365GraphClientBase } from "../common/graph/client-base";
import {
  isGraphSharePointColumnsResponse,
  isGraphSharePointListsResponse,
  isGraphSharePointSitesResponse,
  isM365SharePointColumn,
  isM365SharePointListInfo,
  isM365SharePointSite,
} from "./guards";
import {
  M365SharePointColumn,
  M365SharePointListInfo,
  M365SharePointSite,
  M365SharePointSiteMetadata,
  M365SharePointViewXmlBuilderOptions,
} from "./types";

export class SharePointClient extends M365GraphClientBase {
  public constructor(authentication: M365Authentication) {
    super(authentication);
  }

  private async requestSharePointSite(
    siteId: string,
    filter?: string,
  ): Promise<unknown> {
    return await this.graphRequest(`sites/${siteId}${filter ? `?${filter}` : ""}`);
  }

  private async requestSharePointSites(filter?: string): Promise<unknown> {
    return await this.graphRequest(`sites${filter ? `?${filter}` : ""}`);
  }

  private async requestSharePointLists(siteId: string): Promise<unknown> {
    return await this.graphRequest(
      `sites/${siteId}/lists?$select=id,displayName,list`,
    );
  }

  private async requestSharePointListColumns(
    siteId: string,
    listId: string,
  ): Promise<unknown> {
    return await this.graphRequest(`sites/${siteId}/lists/${listId}/columns`);
  }

  private async requestSharePointRest(
    absoluteUrl: string,
    body?: unknown,
  ): Promise<unknown> {
    const token = await this.getAccessToken();
    const response = await fetch(absoluteUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });

    const data: unknown = await response.json().catch(async () => {
      const text = await response.text().catch(() => "");
      return text ? { raw: text } : undefined;
    });

    if (!response.ok) {
      throw new Error(
        `SharePoint REST request failed: ${response.status} ${response.statusText} - ${JSON.stringify(data)}`,
      );
    }

    return data;
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
      search ? `$search=${search}` : "",
    );
    if (!isGraphSharePointSitesResponse(data)) {
      throw new Error("Microsoft Graph sites response has an invalid format.");
    }
    return data.value;
  }

  public async getSharePointSitesBySearch(
    search?: string,
  ): Promise<M365SharePointSite[]> {
    return await this.getAllSharePointSites(search);
  }

  public async searchSharePointSitesWithSelect(
    search: string,
  ): Promise<M365SharePointSite[]> {
    const data = await this.requestSharePointSites(
      `$search=${search}&$select=name,id,displayName,webUrl`,
    );
    if (!isGraphSharePointSitesResponse(data)) {
      throw new Error("Microsoft Graph sites response has an invalid format.");
    }
    return data.value;
  }

  public async getSharePointAllSitesMetadata(): Promise<
    M365SharePointSiteMetadata[]
  > {
    const data = await this.requestSharePointSites();
    if (!isGraphSharePointSitesResponse(data)) {
      throw new Error("Microsoft Graph sites response has an invalid format.");
    }
    return data.value.map((site: M365SharePointSite) => ({
      id: site.id,
      name: site.name,
      displayName: site.displayName,
      webUrl: site.webUrl,
    }));
  }

  public async getSharePointLists(
    siteId: string,
  ): Promise<M365SharePointListInfo[]> {
    const data = await this.requestSharePointLists(siteId);
    if (!isGraphSharePointListsResponse(data)) {
      throw new Error("Microsoft Graph lists response has an invalid format.");
    }
    const invalid = data.value.some((entry) => !isM365SharePointListInfo(entry));
    if (invalid) {
      throw new Error("Microsoft Graph lists payload contains invalid entries.");
    }
    return data.value;
  }

  public async getSharePointListColumns(
    siteId: string,
    listId: string,
  ): Promise<M365SharePointColumn[]> {
    const data = await this.requestSharePointListColumns(siteId, listId);
    if (!isGraphSharePointColumnsResponse(data)) {
      throw new Error("Microsoft Graph columns response has an invalid format.");
    }
    const invalid = data.value.some((entry) => !isM365SharePointColumn(entry));
    if (invalid) {
      throw new Error(
        "Microsoft Graph columns payload contains invalid entries.",
      );
    }
    return data.value;
  }

  public buildSharePointViewXml(
    options: M365SharePointViewXmlBuilderOptions,
  ): string {
    const fields = (options.viewFields ?? [])
      .map((fieldName) => `<FieldRef Name='${fieldName}' />`)
      .join("");
    const viewFields = fields ? `<ViewFields>${fields}</ViewFields>` : "";
    const query = `<Query>${options.whereClauseXml ?? ""}${options.orderByClauseXml ?? ""}</Query>`;
    const rowLimit =
      typeof options.rowLimit === "number"
        ? `<RowLimit>${options.rowLimit}</RowLimit>`
        : "";
    return `<View>${query}${viewFields}${rowLimit}</View>`;
  }

  public async setSharePointListViewXml(request: {
    siteWebUrl: string;
    listId: string;
    viewId: string;
    viewXml: string;
  }): Promise<unknown> {
    const url = `${request.siteWebUrl}/_api/Web/Lists(guid'${request.listId}')/Views(guid'${request.viewId}')/SetViewXml`;
    return await this.requestSharePointRest(url, { viewXml: request.viewXml });
  }
}

export const createSharePointClient = (
  authentication: M365Authentication,
): SharePointClient => {
  return new SharePointClient(authentication);
};
