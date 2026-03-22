import { describe, expect, it } from "vitest";
import { encodeGraphSearchParameter } from "./format-search-query";

describe("encodeGraphSearchParameter", () => {
  it("wraps property:value in outer quotes and encodes", () => {
    expect(encodeGraphSearchParameter("displayName:Panama")).toBe(
      encodeURIComponent('"displayName:Panama"'),
    );
  });

  it("leaves already-quoted literals unchanged aside from encoding", () => {
    expect(encodeGraphSearchParameter('"displayName:Test"')).toBe(
      encodeURIComponent('"displayName:Test"'),
    );
  });

  it("quotes bare keywords", () => {
    expect(encodeGraphSearchParameter("ada")).toBe(encodeURIComponent('"ada"'));
  });
});
