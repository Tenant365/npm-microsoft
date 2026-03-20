import { M365Authentication } from "../core/auth";
import { MS365Scopes } from "../core/auth";

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

const isM365Team = (value: unknown): value is M365Team => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365Team).id === "string";
};

export class TeamsClient {
  public constructor(private readonly authentication: M365Authentication) {}

  public async getAccessToken(): Promise<string> {
    const accessToken = await this.authentication.GetAccessToken(
      MS365Scopes.DEFAULT,
    );
    return accessToken.token;
  }

  private async graphRequest(path: string, accessToken: string): Promise<unknown> {
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

  private async requestTeam(
    teamId: string,
    accessToken: string,
  ): Promise<M365Team> {
    const data = await this.graphRequest(`teams/${teamId}`, accessToken);

    if (!isM365Team(data)) {
      throw new Error(
        `Microsoft Graph teams/${teamId} response has an invalid format.`,
      );
    }

    return data;
  }

  public async getTeam(teamId: string): Promise<M365Team> {
    const accessToken = await this.getAccessToken();
    return await this.requestTeam(teamId, accessToken);
  }

  public async getAllTeams(): Promise<M365Team[]> {
    const accessToken = await this.getAccessToken();
    const data = await this.graphRequest("teams", accessToken);

    if (!isGraphTeamsResponse(data)) {
      throw new Error("Microsoft Graph teams response has an invalid format.");
    }

    return data.value;
  }
}

export const createTeamsClient = (
  authentication: M365Authentication,
): TeamsClient => {
  return new TeamsClient(authentication);
};
