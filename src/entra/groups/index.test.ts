import { describe, expect, it, vi, afterEach } from "vitest";
import { EntraGroupsClient, isGraphEntraGroupsResponse } from "./index";

afterEach(() => {
  vi.restoreAllMocks();
});

describe("EntraGroupsClient", () => {
  const auth = {
    GetAccessToken: vi
      .fn()
      .mockResolvedValue({ token: "jwt-token", expiresAt: new Date() }),
  };

  it("isGraphEntraGroupsResponse", () => {
    expect(
      isGraphEntraGroupsResponse({
        "@odata.context": "ctx",
        value: [],
      }),
    ).toBe(true);
  });

  it("getGroup returns group", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({ id: "g1", displayName: "G" }),
        text: async () => "",
      }),
    );
    const client = new EntraGroupsClient(auth as any);
    await expect(client.getGroup("g1")).resolves.toEqual({
      id: "g1",
      displayName: "G",
    });
  });
});
