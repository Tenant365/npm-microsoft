import { M365Authentication } from "@/core";

export interface M365Team {
  id: string;
  displayName?: string;
  description?: string;
  [key: string]: unknown;
}

interface GraphTeamsResponse {
  value: M365Team[];
}

const isGraphTeamsResponse = (value: unknown): value is GraphTeamsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return Array.isArray((value as GraphTeamsResponse).value);
};

export class TeamsClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(scope?: string): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(scope);
    return accessToken.token;
  }

  public async getAllTeamsWithAccessToken(
    accessToken: string,
  ): Promise<M365Team[]> {
    const response = await fetch("https://graph.microsoft.com/v1.0/teams", {
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
        `Microsoft Graph teams request failed: ${response.status} ${response.statusText} - ${JSON.stringify(data)}`,
      );
    }

    if (!isGraphTeamsResponse(data)) {
      throw new Error("Microsoft Graph teams response has an invalid format.");
    }

    return data.value;
  }

  public async getAllTeams(scope?: string): Promise<M365Team[]> {
    const accessToken = await this.getAccessToken(scope);
    return this.getAllTeamsWithAccessToken(accessToken);
  }
}

export const createTeamsClient = (
  authentication: M365Authentication,
): TeamsClient => {
  return new TeamsClient(authentication);
};
