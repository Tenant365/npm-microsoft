import { afterEach, describe, expect, it, vi } from "vitest";
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

  it("returns token from authentication provider", async () => {
    const client = new TeamsClient(auth as any);
    const token = await client.getAccessToken();
    expect(token).toBe("jwt-token");
    expect(auth.GetAccessToken).toHaveBeenCalledWith(MS365Scopes.DEFAULT);
  });

  it("validates graph list payload", () => {
    expect(isGraphTeamsResponse({ "@odata.context": "ctx", value: [] })).toBe(
      true,
    );
    expect(isGraphTeamsResponse({ value: [] })).toBe(false);
  });

  it("fetches all teams", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({
          "@odata.context": "ctx",
          value: [{ id: "t1", displayName: "Team" }],
        }),
        text: async () => "",
      }),
    );

    const client = new TeamsClient(auth as any);
    await expect(client.getAllTeams()).resolves.toEqual([
      { id: "t1", displayName: "Team" },
    ]);
  });

  it("fetches single team", async () => {
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
    await expect(client.getTeam("team-1")).resolves.toEqual({
      id: "team-1",
      displayName: "Team One",
    });
  });

  it("creates team with expected payload", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 201,
      statusText: "Created",
      json: async () => ({ id: "created-1", displayName: "New" }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);

    const client = new TeamsClient(auth as any);
    const result = await client.createTeam({
      displayName: "New",
      description: "Desc",
      members: [{ userId: "u-1", roles: ["owner"] }],
    });

    expect(result).toMatchObject({ id: "created-1" });
    expect(fetchMock).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/teams",
      expect.objectContaining({ method: "POST" }),
    );
  });

  it("accepts create-team success with no response body", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 202,
        statusText: "Accepted",
        json: async () => {
          throw new Error("no json");
        },
        text: async () => "",
      }),
    );

    const client = new TeamsClient(auth as any);
    const result = await client.createTeam({
      displayName: "Async Team",
      members: [{ userId: "u-1", roles: ["owner"] }],
    });

    expect(result).toMatchObject({
      status: 202,
      operationLocation: null,
      contentLocation: null,
    });
  });

  it("deletes a team", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 204,
      statusText: "No Content",
      json: async () => ({}),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);

    const client = new TeamsClient(auth as any);
    await client.deleteTeam("team-delete-1");

    expect(fetchMock).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/teams/team-delete-1",
      expect.objectContaining({ method: "DELETE" }),
    );
  });
});
