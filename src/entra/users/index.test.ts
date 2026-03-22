import { describe, expect, it, vi, afterEach } from "vitest";
import { EntraUsersClient, isGraphEntraUsersResponse } from "./index";
import { MS365Scopes } from "../../core/auth";

afterEach(() => {
  vi.restoreAllMocks();
});

describe("EntraUsersClient", () => {
  const auth = {
    GetAccessToken: vi
      .fn()
      .mockResolvedValue({ token: "jwt-token", expiresAt: new Date() }),
  };

  it("getAccessToken forwards scope", async () => {
    const client = new EntraUsersClient(auth as any);
    await client.getAccessToken();
    expect(auth.GetAccessToken).toHaveBeenCalledWith(MS365Scopes.DEFAULT);
  });

  it("isGraphEntraUsersResponse", () => {
    expect(
      isGraphEntraUsersResponse({
        "@odata.context": "ctx",
        value: [],
      }),
    ).toBe(true);
    expect(isGraphEntraUsersResponse({ value: [] })).toBe(false);
  });

  it("getUser returns single user", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({
          id: "u1",
          displayName: "User One",
        }),
        text: async () => "",
      }),
    );
    const client = new EntraUsersClient(auth as any);
    await expect(client.getUser("u1")).resolves.toEqual({
      id: "u1",
      displayName: "User One",
    });
  });

  it("getAllUsers parses list", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({
          "@odata.context": "ctx",
          value: [{ id: "a", displayName: "A" }],
        }),
        text: async () => "",
      }),
    );
    const client = new EntraUsersClient(auth as any);
    const users = await client.getAllUsers();
    expect(users).toEqual([{ id: "a", displayName: "A" }]);
  });

  it("getUsersBySearch encodes $search for Graph (quoted property:value)", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({
        "@odata.context": "ctx",
        value: [],
      }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);
    const client = new EntraUsersClient(auth as any);
    await client.getUsersBySearch("displayName:Panama");
    expect(fetchMock).toHaveBeenCalledWith(
      `https://graph.microsoft.com/v1.0/users?$search=${encodeURIComponent('"displayName:Panama"')}`,
      expect.objectContaining({
        headers: expect.objectContaining({
          ConsistencyLevel: "eventual",
        }),
      }),
    );
  });
});
