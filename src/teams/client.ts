import { M365Authentication } from "../core/auth";
import { M365GraphClientBase } from "../common/graph/client-base";
import {
  isGraphTeamsResponse,
  isGraphTeamTemplatesResponse,
  isM365Team,
} from "./guards";
import {
  M365CreateTeamInput,
  M365CreateTeamResult,
  M365Team,
  M365TeamMetadata,
  M365TeamTemplate,
} from "./types";

export class TeamsClient extends M365GraphClientBase {
  public constructor(authentication: M365Authentication) {
    super(authentication);
  }

  private async executeGraphRequest(
    path: string,
    body?: unknown,
    method: "GET" | "POST" | "PUT" | "DELETE" = "GET",
    headers?: Record<string, string>,
  ): Promise<unknown> {
    return await super.graphRequest(path, {
      method,
      headers,
      body,
    });
  }

  private async requestTeam(teamId: string): Promise<unknown> {
    return await this.executeGraphRequest(`teams/${teamId}`);
  }

  private async requestTeams(filter?: string): Promise<unknown> {
    return await this.executeGraphRequest(`teams${filter ? `?${filter}` : ""}`);
  }

  private async requestTeamTemplates(): Promise<unknown> {
    return await this.executeGraphRequest("teamsTemplates");
  }

  private async requestTeamCreate(input: unknown): Promise<unknown> {
    return await this.executeGraphRequest(
      "teams",
      input,
      "POST",
      { "Content-Type": "application/json" },
    );
  }

  private async requestTeamImageUpload(
    teamId: string,
    image: string,
  ): Promise<unknown> {
    return await this.executeGraphRequest(
      `teams/${teamId}/photo/$value`,
      image,
      "PUT",
      { "Content-Type": "image/jpeg" },
    );
  }

  private async requestTeamImage(teamId: string): Promise<unknown> {
    return await this.executeGraphRequest(`teams/${teamId}/photo/$value`);
  }

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

  public async createTeam(
    input: M365CreateTeamInput,
  ): Promise<M365CreateTeamResult> {
    if (!input.members?.length) {
      throw new Error(
        "createTeam requires at least one member (Graph requires an owner).",
      );
    }

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
      visibility,
    };

    const data = await this.requestTeamCreate(graphBody);

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
