import { M365Authentication, MS365Scopes } from "../core/auth";
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
  M365SharePointColumnCreateInput,
  M365SharePointColumn,
  M365SharePointListInfo,
  M365SharePointListViewInfo,
  M365SharePointListViewXmlRequest,
  M365SharePointSetViewXmlPayload,
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

  private async createSharePointListColumnRequest(
    siteId: string,
    listId: string,
    column: M365SharePointColumnCreateInput,
  ): Promise<unknown> {
    return await this.graphRequest(`sites/${siteId}/lists/${listId}/columns`, {
      method: "POST",
      body: column,
    });
  }

  private async requestSharePointRest(
    absoluteUrl: string,
    method: "GET" | "POST",
    body?: unknown,
  ): Promise<unknown> {
    const normalizedAbsoluteUrl = this.normalizeSharePointSiteWebUrl(absoluteUrl);
    const token = await this.getAccessToken(
      this.getSharePointRestScopeFromAbsoluteUrl(normalizedAbsoluteUrl),
    );
    const response = await fetch(normalizedAbsoluteUrl, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });

    const data: unknown = await response.json().catch(async () => {
      const text = await response.text().catch(() => "");
      return text ? this.safeJsonParse(text) ?? { raw: text } : undefined;
    });

    if (!response.ok) {
      const wwwAuthenticate = response.headers.get("www-authenticate") ?? "";
      throw new Error(
        `SharePoint REST request failed: ${response.status} ${response.statusText}${
          wwwAuthenticate ? ` - www-authenticate: ${wwwAuthenticate}` : ""
        } - ${JSON.stringify(data)}`,
      );
    }

    return data;
  }

  private safeJsonParse(value: string): unknown | undefined {
    try {
      return JSON.parse(value) as unknown;
    } catch {
      return undefined;
    }
  }

  private getSharePointRestScopeFromAbsoluteUrl(absoluteUrl: string): string {
    try {
      const parsed = new URL(absoluteUrl);
      return `${parsed.origin}/.default`;
    } catch {
      return MS365Scopes.DEFAULT;
    }
  }

  private buildSharePointViewXmlEndpoint(
    request: M365SharePointListViewXmlRequest,
  ): string {
    const normalizedSiteWebUrl = this.normalizeSharePointSiteWebUrl(
      request.siteWebUrl,
    );
    return `${normalizedSiteWebUrl}/_api/Web/Lists(guid'${request.listId}')/Views(guid'${request.viewId}')`;
  }

  private buildSharePointListViewsEndpoint(
    siteWebUrl: string,
    listId: string,
  ): string {
    const normalizedSiteWebUrl = this.normalizeSharePointSiteWebUrl(siteWebUrl);
    return `${normalizedSiteWebUrl}/_api/Web/Lists(guid'${listId}')/Views`;
  }

  private normalizeSharePointSiteWebUrl(siteWebUrl: string): string {
    try {
      const parsed = new URL(siteWebUrl);
      parsed.hostname = parsed.hostname.replace("-admin.sharepoint.com", ".sharepoint.com");
      return parsed.toString().replace(/\/$/, "");
    } catch {
      return siteWebUrl.replace("-admin.sharepoint.com", ".sharepoint.com");
    }
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
      search ? `search=${encodeURIComponent(search)}` : "",
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
      `search=${encodeURIComponent(search)}&$select=name,id,displayName,webUrl`,
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

  public async createSharePointListColumn(
    siteId: string,
    listId: string,
    column: M365SharePointColumnCreateInput,
  ): Promise<M365SharePointColumn> {
    const data = await this.createSharePointListColumnRequest(siteId, listId, column);
    if (!isM365SharePointColumn(data)) {
      throw new Error(
        `Microsoft Graph column create response has an invalid format: ${JSON.stringify(data)}`,
      );
    }
    return data;
  }

  public async createSharePointListColumns(
    siteId: string,
    listId: string,
    columns: M365SharePointColumnCreateInput[],
  ): Promise<M365SharePointColumn[]> {
    const created: M365SharePointColumn[] = [];
    for (const column of columns) {
      const createdColumn = await this.createSharePointListColumn(
        siteId,
        listId,
        column,
      );
      created.push(createdColumn);
    }
    return created;
  }

  public buildSharePointViewXml(
    options: M365SharePointViewXmlBuilderOptions,
  ): string {
    if (typeof options.baseViewXml === "string") {
      return this.buildSharePointViewXmlFromBase(
        options.baseViewXml,
        options.viewFields ?? [],
      );
    }

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

  public buildSharePointSetViewXmlPayload(
    options: M365SharePointViewXmlBuilderOptions,
  ): M365SharePointSetViewXmlPayload {
    return {
      viewXml: this.buildSharePointViewXml(options),
    };
  }

  public buildSharePointSetViewXmlPayloadFromViewXml(
    viewXml: string,
  ): M365SharePointSetViewXmlPayload {
    return { viewXml };
  }

  private buildSharePointViewXmlFromBase(
    baseViewXml: string,
    viewFields: string[],
  ): string {
    const fields = viewFields
      .map((fieldName) => `<FieldRef Name="${fieldName}"/>`)
      .join("");
    const newViewFields = `<ViewFields>${fields}</ViewFields>`;

    if (/<ViewFields>[\s\S]*?<\/ViewFields>/i.test(baseViewXml)) {
      return baseViewXml.replace(
        /<ViewFields>[\s\S]*?<\/ViewFields>/i,
        newViewFields,
      );
    }

    if (/<\/Query>/i.test(baseViewXml)) {
      return baseViewXml.replace(/<\/Query>/i, `</Query>${newViewFields}`);
    }

    return baseViewXml.replace(/<\/View>/i, `${newViewFields}</View>`);
  }

  public async getSharePointListViewXml(
    request: M365SharePointListViewXmlRequest,
  ): Promise<string> {
    const url = `${this.buildSharePointViewXmlEndpoint(request)}/ListViewXml`;
    const data = await this.requestSharePointRest(url, "GET");
    const maybeValue = (data as any)?.d?.ListViewXml;
    if (typeof maybeValue !== "string") {
      throw new Error(
        `SharePoint REST ListViewXml response has an invalid format: ${JSON.stringify(data)}`,
      );
    }
    return maybeValue;
  }

  public async getSharePointDefaultListViewXmlByTitle(
    siteWebUrl: string,
    listTitle: string,
    listId: string,
  ): Promise<string> {
    const defaultView = await this.getSharePointDefaultViewByListTitle(
      siteWebUrl,
      listTitle,
    );
    if (!defaultView?.Id) {
      throw new Error(
        `No default view found for SharePoint list '${listTitle}' at '${siteWebUrl}'.`,
      );
    }

    return await this.getSharePointListViewXml({
      siteWebUrl,
      listId,
      viewId: defaultView.Id,
    });
  }

  public async getSharePointDefaultViewByListId(
    siteWebUrl: string,
    listId: string,
  ): Promise<M365SharePointListViewInfo | null> {
    const url = `${this.buildSharePointListViewsEndpoint(siteWebUrl, listId)}?$filter=DefaultView eq true&$select=Id,Title`;
    const data = await this.requestSharePointRest(url, "GET");
    const maybeResults = (data as any)?.d?.results;
    if (!Array.isArray(maybeResults)) {
      throw new Error(
        `SharePoint REST default view response has an invalid format: ${JSON.stringify(data)}`,
      );
    }

    const first = maybeResults[0];
    if (first === undefined) {
      return null;
    }
    if (typeof first?.Id !== "string") {
      throw new Error(
        `SharePoint REST default view payload contains an invalid entry: ${JSON.stringify(first)}`,
      );
    }

    return first as M365SharePointListViewInfo;
  }

  public async getSharePointDefaultListViewXmlByListId(
    siteWebUrl: string,
    listId: string,
  ): Promise<string> {
    const defaultView = await this.getSharePointDefaultViewByListId(
      siteWebUrl,
      listId,
    );
    if (!defaultView?.Id) {
      throw new Error(
        `No default view found for SharePoint list '${listId}' at '${siteWebUrl}'.`,
      );
    }

    return await this.getSharePointListViewXml({
      siteWebUrl,
      listId,
      viewId: defaultView.Id,
    });
  }

  public async getSharePointDefaultViewByListTitle(
    siteWebUrl: string,
    listTitle: string,
  ): Promise<M365SharePointListViewInfo | null> {
    const normalizedSiteWebUrl = this.normalizeSharePointSiteWebUrl(siteWebUrl);
    const escapedListTitle = listTitle.replace(/'/g, "''");
    const url = `${normalizedSiteWebUrl}/_api/web/lists/getbytitle('${escapedListTitle}')/views?$filter=DefaultView eq true&$select=Id,Title`;
    const data = await this.requestSharePointRest(url, "GET");
    const maybeResults = (data as any)?.d?.results;
    if (!Array.isArray(maybeResults)) {
      throw new Error(
        `SharePoint REST default view response has an invalid format: ${JSON.stringify(data)}`,
      );
    }

    const first = maybeResults[0];
    if (first === undefined) {
      return null;
    }
    if (typeof first?.Id !== "string") {
      throw new Error(
        `SharePoint REST default view payload contains an invalid entry: ${JSON.stringify(first)}`,
      );
    }

    return first as M365SharePointListViewInfo;
  }

  public async getSharePointListViewsByTitle(
    siteWebUrl: string,
    listTitle: string,
  ): Promise<M365SharePointListViewInfo[]> {
    const normalizedSiteWebUrl = this.normalizeSharePointSiteWebUrl(siteWebUrl);
    const escapedListTitle = listTitle.replace(/'/g, "''");
    const url = `${normalizedSiteWebUrl}/_api/web/lists/getbytitle('${escapedListTitle}')/views?$select=Id,Title,DefaultView`;
    const data = await this.requestSharePointRest(url, "GET");
    const maybeResults = (data as any)?.d?.results;
    if (!Array.isArray(maybeResults)) {
      throw new Error(
        `SharePoint REST list views response has an invalid format: ${JSON.stringify(data)}`,
      );
    }

    const invalidEntry = maybeResults.find(
      (entry) => typeof entry?.Id !== "string",
    );
    if (invalidEntry !== undefined) {
      throw new Error(
        `SharePoint REST list views payload contains an invalid entry: ${JSON.stringify(invalidEntry)}`,
      );
    }

    return maybeResults as M365SharePointListViewInfo[];
  }

  public async setSharePointListViewXml(
    request: M365SharePointListViewXmlRequest & { viewXml: string },
  ): Promise<unknown> {
    const url = `${this.buildSharePointViewXmlEndpoint(request)}/SetViewXml`;
    return await this.requestSharePointRest(
      url,
      "POST",
      this.buildSharePointSetViewXmlPayloadFromViewXml(request.viewXml),
    );
  }
}

export const createSharePointClient = (
  authentication: M365Authentication,
): SharePointClient => {
  return new SharePointClient(authentication);
};
