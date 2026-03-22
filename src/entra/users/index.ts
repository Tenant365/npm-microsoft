import { M365Authentication } from "../../core/auth";
import { MS365Scopes } from "../../core/auth";
import { encodeGraphSearchParameter } from "../internal/format-search-query";
import { entraGraphGet } from "../internal/graph-get";

export interface M365EntraUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  [key: string]: unknown;
}

export interface M365EntraUserMetadata {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  [key: string]: unknown;
}

export interface GraphEntraUsersResponse {
  "@odata.context"?: string;
  value: M365EntraUser[];
}

export const isGraphEntraUsersResponse = (
  value: unknown,
): value is GraphEntraUsersResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphEntraUsersResponse).value) &&
    typeof (value as GraphEntraUsersResponse)["@odata.context"] === "string"
  );
};

const isM365EntraUser = (value: unknown): value is M365EntraUser => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365EntraUser).id === "string";
};

export class EntraUsersClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async requestUser(userId: string): Promise<unknown> {
    return await entraGraphGet(`users/${userId}`, await this.getAccessToken());
  }

  private async requestUsers(filter?: string): Promise<unknown> {
    const usesSearch = filter?.includes("$search=") === true;
    return await entraGraphGet(
      `users${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
      usesSearch ? { ConsistencyLevel: "eventual" } : undefined,
    );
  }

  public async getUserById(userId: string): Promise<M365EntraUser> {
    return await this.getUser(userId);
  }

  public async getUser(userId: string): Promise<M365EntraUser> {
    const data = await this.requestUser(userId);
    if (!isM365EntraUser(data)) {
      throw new Error(
        `Microsoft Graph users/${userId} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getAllUsers(search?: string): Promise<M365EntraUser[]> {
    const data = await this.requestUsers(
      search ? `$search=${encodeGraphSearchParameter(search)}` : "",
    );
    if (!isGraphEntraUsersResponse(data)) {
      throw new Error("Microsoft Graph users response has an invalid format.");
    }
    return data.value;
  }

  public async getUsersBySearch(search?: string): Promise<M365EntraUser[]> {
    return await this.getAllUsers(search);
  }

  public async getAllUsersMetadata(): Promise<M365EntraUserMetadata[]> {
    const data = await this.requestUsers();
    if (!isGraphEntraUsersResponse(data)) {
      throw new Error("Microsoft Graph users response has an invalid format.");
    }
    return data.value.map((user: M365EntraUser) => ({
      id: user.id,
      displayName: user.displayName,
      mail: user.mail,
      userPrincipalName: user.userPrincipalName,
    }));
  }
}

export const createEntraUsersClient = (
  authentication: M365Authentication,
): EntraUsersClient => {
  return new EntraUsersClient(authentication);
};
