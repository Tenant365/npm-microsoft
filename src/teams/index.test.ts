import { describe, expect, it, vi, afterEach } from "vitest";
import { TeamsClient } from "./index";

afterEach(() => {
  vi.restoreAllMocks();
});

describe("TeamsClient", () => {
  it("returns token from authentication provider", async () => {
    const GetAccessToken = vi
      .fn()
      .mockResolvedValue({ token: "abc", expiresAt: new Date() });
    const client = new TeamsClient({ GetAccessToken } as any);

    const token = await client.getAccessToken("scope-a");

    expect(token).toBe("abc");
    expect(GetAccessToken).toHaveBeenCalledWith("scope-a");
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

    const client = new TeamsClient({ GetAccessToken: vi.fn() } as any);
    const teams = await client.getAllTeamsWithAccessToken("jwt-token");

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

    const client = new TeamsClient({ GetAccessToken: vi.fn() } as any);

    await expect(client.getAllTeamsWithAccessToken("bad-token")).rejects.toThrow(
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

    const client = new TeamsClient({ GetAccessToken: vi.fn() } as any);

    await expect(client.getAllTeamsWithAccessToken("token")).rejects.toThrow(
      "Microsoft Graph teams response has an invalid format.",
    );
  });
});
