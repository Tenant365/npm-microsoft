import { M365Authentication } from "../../core/auth";
import { MS365Scopes } from "../../core/auth";
import { encodeGraphSearchParameter } from "../internal/format-search-query";
import { entraGraphGet } from "../internal/graph-get";

export interface M365EntraGroup {
  id: string;
  displayName?: string;
  description?: string;
  mail?: string;
  mailNickname?: string;
  [key: string]: unknown;
}

export interface M365EntraGroupMetadata {
  id: string;
  displayName?: string;
  description?: string;
  mail?: string;
  mailNickname?: string;
  [key: string]: unknown;
}

export interface GraphEntraGroupsResponse {
  "@odata.context"?: string;
  value: M365EntraGroup[];
}

export const isGraphEntraGroupsResponse = (
  value: unknown,
): value is GraphEntraGroupsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphEntraGroupsResponse).value) &&
    typeof (value as GraphEntraGroupsResponse)["@odata.context"] === "string"
  );
};

const isM365EntraGroup = (value: unknown): value is M365EntraGroup => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365EntraGroup).id === "string";
};

export class EntraGroupsClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async requestGroup(groupId: string): Promise<unknown> {
    return await entraGraphGet(`groups/${groupId}`, await this.getAccessToken());
  }

  private async requestGroups(filter?: string): Promise<unknown> {
    const usesSearch = filter?.includes("$search=") === true;
    return await entraGraphGet(
      `groups${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
      usesSearch ? { ConsistencyLevel: "eventual" } : undefined,
    );
  }

  public async getGroupById(groupId: string): Promise<M365EntraGroup> {
    return await this.getGroup(groupId);
  }

  public async getGroup(groupId: string): Promise<M365EntraGroup> {
    const data = await this.requestGroup(groupId);
    if (!isM365EntraGroup(data)) {
      throw new Error(
        `Microsoft Graph groups/${groupId} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getAllGroups(search?: string): Promise<M365EntraGroup[]> {
    const data = await this.requestGroups(
      search ? `$search=${encodeGraphSearchParameter(search)}` : "",
    );
    if (!isGraphEntraGroupsResponse(data)) {
      throw new Error(
        "Microsoft Graph groups response has an invalid format.",
      );
    }
    return data.value;
  }

  public async getGroupsBySearch(search?: string): Promise<M365EntraGroup[]> {
    return await this.getAllGroups(search);
  }

  public async getAllGroupsMetadata(): Promise<M365EntraGroupMetadata[]> {
    const data = await this.requestGroups();
    if (!isGraphEntraGroupsResponse(data)) {
      throw new Error(
        "Microsoft Graph groups response has an invalid format.",
      );
    }
    return data.value.map((group: M365EntraGroup) => ({
      id: group.id,
      displayName: group.displayName,
      description: group.description,
      mail: group.mail,
      mailNickname: group.mailNickname,
    }));
  }
}

export const createEntraGroupsClient = (
  authentication: M365Authentication,
): EntraGroupsClient => {
  return new EntraGroupsClient(authentication);
};
