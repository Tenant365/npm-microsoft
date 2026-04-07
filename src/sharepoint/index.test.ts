import { afterEach, describe, expect, it, vi } from "vitest";
import { SharePointClient } from "./index";

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SharePointClient API integrations", () => {
  const auth = {
    GetAccessToken: vi
      .fn()
      .mockResolvedValue({ token: "jwt-token", expiresAt: new Date() }),
  };

  it("searches sites with fixed select projection", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({
        "@odata.context": "ctx",
        value: [{ id: "s1", displayName: "Controlx6Team" }],
      }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);

    const client = new SharePointClient(auth as any);
    await client.searchSharePointSitesWithSelect("Controlx6Team");

    expect(fetchMock).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/sites?search=Controlx6Team&$select=name,id,displayName,webUrl",
      expect.any(Object),
    );
  });

  it("loads list columns from Graph", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        statusText: "OK",
        json: async () => ({
          "@odata.context": "ctx",
          value: [{ displayName: "Title", name: "Title" }],
        }),
        text: async () => "",
      }),
    );
    const client = new SharePointClient(auth as any);
    const columns = await client.getSharePointListColumns("site-id", "list-id");
    expect(columns[0].displayName).toBe("Title");
  });

  it("builds and posts view xml to SharePoint REST endpoint", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({ d: { SetViewXml: true } }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);
    const client = new SharePointClient(auth as any);
    const xml = client.buildSharePointViewXml({
      viewFields: ["Title", "Modified"],
      rowLimit: 100,
      whereClauseXml: "<Where><IsNotNull><FieldRef Name='Title' /></IsNotNull></Where>",
    });
    await client.setSharePointListViewXml({
      siteWebUrl: "https://secnexdev.sharepoint.com/sites/Controlx11Team",
      listId: "f3d9da8b-39d1-4567-9cec-996894b2ed78",
      viewId: "8cf2f9a6-11a3-4eca-8a34-83b9fec192b2",
      viewXml: xml,
    });

    expect(fetchMock).toHaveBeenCalledWith(
      "https://secnexdev.sharepoint.com/sites/Controlx11Team/_api/Web/Lists(guid'f3d9da8b-39d1-4567-9cec-996894b2ed78')/Views(guid'8cf2f9a6-11a3-4eca-8a34-83b9fec192b2')/SetViewXml",
      expect.objectContaining({ method: "POST" }),
    );
  });

  it("loads ListViewXml from SharePoint REST endpoint", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({ d: { ListViewXml: "<View><Query /></View>" } }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);
    const client = new SharePointClient(auth as any);
    const viewXml = await client.getSharePointListViewXml({
      siteWebUrl: "https://secnexdev.sharepoint.com/sites/Controlx11Team",
      listId: "f3d9da8b-39d1-4567-9cec-996894b2ed78",
      viewId: "8cf2f9a6-11a3-4eca-8a34-83b9fec192b2",
    });

    expect(viewXml).toBe("<View><Query /></View>");
    expect(fetchMock).toHaveBeenCalledWith(
      "https://secnexdev.sharepoint.com/sites/Controlx11Team/_api/Web/Lists(guid'f3d9da8b-39d1-4567-9cec-996894b2ed78')/Views(guid'8cf2f9a6-11a3-4eca-8a34-83b9fec192b2')/ListViewXml",
      expect.objectContaining({ method: "GET" }),
    );
  });

  it("loads default view by list title from SharePoint REST endpoint", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({
        d: { results: [{ Id: "8cf2f9a6-11a3-4eca-8a34-83b9fec192b2", Title: "All Documents" }] },
      }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);
    const client = new SharePointClient(auth as any);
    const defaultView = await client.getSharePointDefaultViewByListTitle(
      "https://secnexdev.sharepoint.com/sites/Controlx11Team",
      "Documents",
    );

    expect(defaultView?.Id).toBe("8cf2f9a6-11a3-4eca-8a34-83b9fec192b2");
    expect(fetchMock).toHaveBeenCalledWith(
      "https://secnexdev.sharepoint.com/sites/Controlx11Team/_api/web/lists/getbytitle('Documents')/views?$filter=DefaultView eq true&$select=Id,Title",
      expect.objectContaining({ method: "GET" }),
    );
  });

  it("loads all views by list title from SharePoint REST endpoint", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: "OK",
      json: async () => ({
        d: {
          results: [
            { Id: "8cf2f9a6-11a3-4eca-8a34-83b9fec192b2", Title: "All Documents", DefaultView: true },
            { Id: "e258545f-c971-4be5-af00-6c5288546370", Title: "By Author", DefaultView: false },
          ],
        },
      }),
      text: async () => "",
    });
    vi.stubGlobal("fetch", fetchMock);
    const client = new SharePointClient(auth as any);
    const views = await client.getSharePointListViewsByTitle(
      "https://secnexdev.sharepoint.com/sites/Controlx11Team",
      "Documents",
    );

    expect(views).toHaveLength(2);
    expect(views[0]?.DefaultView).toBe(true);
    expect(fetchMock).toHaveBeenCalledWith(
      "https://secnexdev.sharepoint.com/sites/Controlx11Team/_api/web/lists/getbytitle('Documents')/views?$select=Id,Title,DefaultView",
      expect.objectContaining({ method: "GET" }),
    );
  });
});
