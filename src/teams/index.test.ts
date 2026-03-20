import { describe, expect, it, vi, afterEach } from "vitest";
import { TeamsClient } from "./index";
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
});
