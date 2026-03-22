import { M365Authentication } from "../../core/auth";
import { MS365Scopes } from "../../core/auth";
import { entraGraphGet } from "../internal/graph-get";

export interface M365EntraDirectoryRole {
  id: string;
  displayName?: string;
  description?: string;
  roleTemplateId?: string;
  [key: string]: unknown;
}

export interface GraphEntraDirectoryRolesResponse {
  "@odata.context"?: string;
  value: M365EntraDirectoryRole[];
}

export const isGraphEntraDirectoryRolesResponse = (
  value: unknown,
): value is GraphEntraDirectoryRolesResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphEntraDirectoryRolesResponse).value) &&
    typeof (value as GraphEntraDirectoryRolesResponse)["@odata.context"] ===
      "string"
  );
};

const isM365EntraDirectoryRole = (
  value: unknown,
): value is M365EntraDirectoryRole => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365EntraDirectoryRole).id === "string";
};

export interface M365EntraDirectoryObject {
  id?: string;
  [key: string]: unknown;
}

export interface GraphEntraDirectoryObjectsResponse {
  "@odata.context"?: string;
  value: M365EntraDirectoryObject[];
}

const isGraphEntraDirectoryObjectsResponse = (
  value: unknown,
): value is GraphEntraDirectoryObjectsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphEntraDirectoryObjectsResponse).value) &&
    typeof (value as GraphEntraDirectoryObjectsResponse)["@odata.context"] ===
      "string"
  );
};

export class EntraDirectoryRolesClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async requestDirectoryRoles(): Promise<unknown> {
    return await entraGraphGet(
      "directoryRoles",
      await this.getAccessToken(),
    );
  }

  private async requestDirectoryRoleMembers(roleId: string): Promise<unknown> {
    return await entraGraphGet(
      `directoryRoles/${roleId}/members`,
      await this.getAccessToken(),
    );
  }

  public async getAllDirectoryRoles(): Promise<M365EntraDirectoryRole[]> {
    const data = await this.requestDirectoryRoles();
    if (!isGraphEntraDirectoryRolesResponse(data)) {
      throw new Error(
        "Microsoft Graph directoryRoles response has an invalid format.",
      );
    }
    return data.value;
  }

  public async getDirectoryRole(
    roleId: string,
  ): Promise<M365EntraDirectoryRole> {
    const data = await entraGraphGet(
      `directoryRoles/${roleId}`,
      await this.getAccessToken(),
    );
    if (!isM365EntraDirectoryRole(data)) {
      throw new Error(
        `Microsoft Graph directoryRoles/${roleId} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getDirectoryRoleMembers(
    roleId: string,
  ): Promise<M365EntraDirectoryObject[]> {
    const data = await this.requestDirectoryRoleMembers(roleId);
    if (!isGraphEntraDirectoryObjectsResponse(data)) {
      throw new Error(
        `Microsoft Graph directoryRoles/${roleId}/members response has an invalid format.`,
      );
    }
    return data.value;
  }
}

export const createEntraDirectoryRolesClient = (
  authentication: M365Authentication,
): EntraDirectoryRolesClient => {
  return new EntraDirectoryRolesClient(authentication);
};
