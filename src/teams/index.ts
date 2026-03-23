import { M365Authentication } from "../core/auth";
import { MS365Scopes } from "../core/auth";

export interface M365Team {
  id?: string;
  displayName?: string;
  description?: string;
  [key: string]: unknown;
}

export interface M365TeamMetadata {
  id: string;
  displayName?: string;
  description?: string;
  [key: string]: unknown;
}

/** Entry from `GET /teamsTemplates` (team creation template catalog). */
export interface M365TeamTemplate {
  id: string;
  [key: string]: unknown;
}

/** Owner or member to add when creating a team (Microsoft Graph requirement). */
export interface M365CreateTeamMemberInput {
  /** Azure AD object id of the user */
  userId: string;
  roles?: ("owner" | "member")[];
}

/** Visibility of the team */
export type M365TeamVisibility = "private" | "public";

/**
 * Input for {@link TeamsClient.createTeam}. Graph requires a template binding
 * and at least one owner in `members`.
 */
export interface M365CreateTeamInput {
  displayName: string;
  description?: string;
  /**
   * Teams template id for `teamsTemplates('...')`.
   * Default: `standard` (Microsoft 365 default team).
   * Ignored if {@link templateOdataBind} is set.
   */
  templateId?: string;
  /**
   * Full OData bind URL for the template (e.g. from {@link TeamsClient.getAllTeamTemplates}).
   * Use this if the default `standard` template returns 404 from the Teams templates backend.
   */
  templateOdataBind?: string;
  members: M365CreateTeamMemberInput[];
  visibility?: M365TeamVisibility;
  image?: string;
}

/** Result when Graph provisions a team asynchronously (HTTP 202). */
export interface M365CreateTeamProvisionAccepted {
  status: 202;
  operationLocation: string | null;
  contentLocation: string | null;
  body?: unknown;
}

export type M365CreateTeamResult = M365Team | M365CreateTeamProvisionAccepted;

export interface GraphTeamsResponse {
  "@odata.context"?: string;
  value: M365Team[];
}

export const isGraphTeamsResponse = (
  value: unknown,
): value is GraphTeamsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphTeamsResponse).value) &&
    typeof (value as GraphTeamsResponse)["@odata.context"] === "string"
  );
};

interface GraphTeamTemplatesResponse {
  value: M365TeamTemplate[];
}

const isGraphTeamTemplatesResponse = (
  value: unknown,
): value is GraphTeamTemplatesResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return Array.isArray((value as GraphTeamTemplatesResponse).value);
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

  private async graphRequest(
    path: string,
    accessToken: string,
    body?: string,
    method?: "GET" | "POST" | "PUT" | "DELETE",
    headers?: Record<string, string>,
  ): Promise<unknown> {
    const response = await fetch(`https://graph.microsoft.com/v1.0/${path}`, {
      method: method ?? "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...headers,
      },
      body: body ? JSON.stringify(body) : undefined,
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

  private async requestTeam(teamId: string): Promise<unknown> {
    return await this.graphRequest(
      `teams/${teamId}`,
      await this.getAccessToken(),
    );
  }

  private async requestTeams(filter?: string): Promise<unknown> {
    return await this.graphRequest(
      `teams${filter ? `?${filter}` : ""}`,
      await this.getAccessToken(),
    );
  }

  private async requestTeamTemplates(): Promise<unknown> {
    return await this.graphRequest(
      "teamsTemplates",
      await this.getAccessToken(),
    );
  }

  private async requestTeamCreate(
    input: M365CreateTeamInput,
  ): Promise<unknown> {
    return await this.graphRequest(
      "teams",
      await this.getAccessToken(),
      JSON.stringify(input),
      "POST",
      {
        "Content-Type": "application/json",
      },
    );
  }

  private async requestTeamImageUpload(
    teamId: string,
    image: string,
  ): Promise<unknown> {
    return await this.graphRequest(
      `teams/${teamId}/photo/$value`,
      await this.getAccessToken(),
      image,
      "PUT",
      {
        "Content-Type": "image/jpeg",
      },
    );
  }

  private async requestTeamImage(teamId: string): Promise<unknown> {
    return await this.graphRequest(
      `teams/${teamId}/photo/$value`,
      await this.getAccessToken(),
    );
  }

  /**
   * Lists team templates available for the tenant (`GET /v1.0/teamsTemplates`).
   * Use an `id` from this list in {@link M365CreateTeamInput.templateId} or build
   * {@link M365CreateTeamInput.templateOdataBind} if `standard` fails with 404.
   */
  public async getAllTeamTemplates(): Promise<M365TeamTemplate[]> {
    const data = await this.requestTeamTemplates();
    if (!isGraphTeamTemplatesResponse(data)) {
      throw new Error(
        "Microsoft Graph teamsTemplates response has an invalid format.",
      );
    }
    return data.value;
  }

  public async getTeamById(teamId: string): Promise<M365Team> {
    return await this.getTeam(teamId);
  }

  public async getTeam(teamId: string): Promise<M365Team> {
    const data = await this.requestTeam(teamId);
    if (!isM365Team(data)) {
      throw new Error(
        `Microsoft Graph teams/${teamId} response has an invalid format.`,
      );
    }
    return data;
  }

  public async getAllTeams(search?: string): Promise<M365Team[]> {
    const data = await this.requestTeams(search ? `$search=${search}` : "");
    if (!isGraphTeamsResponse(data)) {
      throw new Error("Microsoft Graph teams response has an invalid format.");
    }
    return data.value;
  }

  public async getTeamsBySearch(search?: string): Promise<M365Team[]> {
    return await this.getAllTeams(search);
  }

  public async getAllTeamsMetadata(): Promise<M365TeamMetadata[]> {
    const data = await this.requestTeams();
    if (!isGraphTeamsResponse(data)) {
      throw new Error("Microsoft Graph teams response has an invalid format.");
    }
    return data.value.map((team: M365Team) => {
      if (typeof team.id !== "string") {
        throw new Error("Microsoft Graph team in list is missing id.");
      }
      return {
        id: team.id,
        displayName: team.displayName,
        description: team.description,
      };
    });
  }

  /**
   * Creates a team via Microsoft Graph (`POST /v1.0/teams`).
   * Sends the required `template@odata.bind` and `members` payload.
   * Graph often returns **202 Accepted** while provisioning; then use
   * `operationLocation` / `contentLocation` headers to track completion.
   */
  public async createTeam(
    input: M365CreateTeamInput,
  ): Promise<M365CreateTeamResult> {
    if (!input.members?.length) {
      throw new Error(
        "createTeam requires at least one member (Graph requires an owner).",
      );
    }

    const accessToken = await this.getAccessToken();
    const templateId = input.templateId ?? "standard";
    const visibility = input.visibility ?? "private";
    const templateBind =
      input.templateOdataBind ??
      `https://graph.microsoft.com/v1.0/teamsTemplates('${templateId.replace(/'/g, "''")}')`;
    const graphBody = {
      "template@odata.bind": templateBind,
      displayName: input.displayName,
      ...(input.description !== undefined
        ? { description: input.description }
        : {}),
      members: input.members.map((m) => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: m.roles?.length ? m.roles : ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${m.userId}')`,
      })),
      visibility: visibility,
    };

    const data = await this.requestTeamCreate(
      graphBody as unknown as M365CreateTeamInput,
    );
    if (!isM365Team(data)) {
      throw new Error(
        `Microsoft Graph teams create response has an invalid format: ${JSON.stringify(data)}`,
      );
    }
    return data;
  }

  public async uploadTeamImage(
    teamId: string,
    image: string,
  ): Promise<unknown> {
    return await this.requestTeamImageUpload(teamId, image);
  }

  public async getTeamImage(teamId: string): Promise<unknown> {
    return await this.requestTeamImage(teamId);
  }
}

export const createTeamsClient = (
  authentication: M365Authentication,
): TeamsClient => {
  return new TeamsClient(authentication);
};
