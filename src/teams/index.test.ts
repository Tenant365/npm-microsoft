import { describe, expect, it, vi, afterEach } from "vitest";
import { TeamsClient, isGraphTeamsResponse } from "./index";
import { MS365Scopes } from "../core/auth";

afterEach(() => {
  vi.restoreAllMocks();
});

describe("TeamsClient", () => {
  const auth = {
    GetAccessToken: vi
      .fn()
      .mockResolvedValue({ token: "jwt-token", expiresAt: new Date() }),
  };

  afterEach(() => {
    auth.GetAccessToken.mockClear();
  });

  it("returns token from authentication provider", async () => {
    const GetAccessToken = vi
      .fn()
      .mockResolvedValue({ token: "abc", expiresAt: new Date() });
    const client = new TeamsClient({ GetAccessToken } as any);

    const token = await client.getAccessToken();

    expect(token).toBe("abc");
    expect(GetAccessToken).toHaveBeenCalledWith(MS365Scopes.DEFAULT);
  });

  it("fetches teams with bearer token", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
        value: [{ id: "t1", displayName: "Team A" }],
      }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);

    const client = new TeamsClient(auth as any);
    const teams = await client.getAllTeams();

    expect(fetchMock).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/teams",
      expect.objectContaining({
        method: "GET",
        headers: { Authorization: "Bearer jwt-token" },
      }),
    );
    expect(teams).toEqual([{ id: "t1", displayName: "Team A" }]);
  });

  it("throws on graph http errors", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: false,
        status: 401,
        statusText: "Unauthorized",
        json: async () => ({ error: "invalid_token" }),
        text: async () => "",
      }),
    );

    const client = new TeamsClient(auth as any);

    await expect(client.getAllTeams()).rejects.toThrow(
      "Microsoft Graph teams request failed: 401 Unauthorized",
    );
  });

  it("isGraphTeamsResponse matches Graph list shape", () => {
    expect(
      isGraphTeamsResponse({
        "@odata.context": "ctx",
        value: [],
      }),
    ).toBe(true);
    expect(isGraphTeamsResponse({ value: [] })).toBe(false);
    expect(isGraphTeamsResponse({ items: [] })).toBe(false);
  });

  it("throws when graph payload is invalid", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({ items: [] }),
        text: async () => "",
      }),
    );

    const client = new TeamsClient(auth as any);

    await expect(client.getAllTeams()).rejects.toThrow(
      "Microsoft Graph teams response has an invalid format.",
    );
  });

  it("fetches a single team", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({ id: "team-1", displayName: "Team One" }),
        text: async () => "",
      }),
    );

    const client = new TeamsClient(auth as any);
    const team = await client.getTeam("team-1");

    expect(team).toEqual({ id: "team-1", displayName: "Team One" });
  });

  it("getTeamById delegates to getTeam", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({ id: "x", displayName: "X" }),
        text: async () => "",
      }),
    );
    const client = new TeamsClient(auth as any);
    await expect(client.getTeamById("x")).resolves.toEqual({
      id: "x",
      displayName: "X",
    });
  });

  it("getAllTeamsMetadata maps list items", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({
          "@odata.context": "ctx",
          value: [
            { id: "a", displayName: "A", description: "d", extra: 1 },
          ],
        }),
        text: async () => "",
      }),
    );
    const client = new TeamsClient(auth as any);
    const meta = await client.getAllTeamsMetadata();
    expect(meta).toEqual([
      { id: "a", displayName: "A", description: "d" },
    ]);
  });

  it("createTeam sends Graph template and members payload", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 202,
      statusText: "Accepted",
      headers: new Headers({
        Location: "https://graph.microsoft.com/v1.0/teams('op')/operations('1')",
        "Content-Location":
          "https://graph.microsoft.com/v1.0/teams('pending')",
      }),
      json: async () => ({}),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);

    const client = new TeamsClient(auth as any);
    const result = await client.createTeam({
      displayName: "New",
      description: "Desc",
      templateId: "standard",
      members: [{ userId: "user-guid-1", roles: ["owner"] }],
    });

    expect(fetchMock).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/teams",
      expect.objectContaining({
        method: "POST",
        headers: expect.objectContaining({
          Authorization: "Bearer jwt-token",
          "Content-Type": "application/json",
        }),
        body: JSON.stringify({
          "template@odata.bind":
            "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
          displayName: "New",
          description: "Desc",
          members: [
            {
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              roles: ["owner"],
              "user@odata.bind":
                "https://graph.microsoft.com/v1.0/users('user-guid-1')",
            },
          ],
        }),
      }),
    );

    expect(result).toMatchObject({ status: 202 });
  });

  it("createTeam rejects empty members", async () => {
    const client = new TeamsClient(auth as any);
    await expect(
      client.createTeam({ displayName: "X", members: [] }),
    ).rejects.toThrow("requires at least one member");
  });

  it("getAllTeamTemplates returns template catalog", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({
          value: [{ id: "standard" }, { id: "educationClass" }],
        }),
        text: async () => "",
      }),
    );
    const client = new TeamsClient(auth as any);
    const templates = await client.getAllTeamTemplates();
    expect(templates).toEqual([{ id: "standard" }, { id: "educationClass" }]);
  });

  it("createTeam uses templateOdataBind when provided", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 202,
      statusText: "Accepted",
      headers: new Headers({ Location: "https://example/op" }),
      json: async () => ({}),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);
    const client = new TeamsClient(auth as any);
    await client.createTeam({
      displayName: "T",
      templateOdataBind:
        "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')",
      members: [{ userId: "u1" }],
    });
    const body = JSON.parse(fetchMock.mock.calls[0][1].body);
    expect(body["template@odata.bind"]).toBe(
      "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')",
    );
  });
});
