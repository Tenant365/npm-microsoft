import {
  GraphSharePointColumnsResponse,
  GraphSharePointListsResponse,
  GraphSharePointSitesResponse,
  M365SharePointColumn,
  M365SharePointListInfo,
  M365SharePointSite,
  M365SharePointSiteMetadata,
} from "./types";

export const isGraphSharePointSitesResponse = (
  value: unknown,
): value is GraphSharePointSitesResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphSharePointSitesResponse).value) &&
    typeof (value as GraphSharePointSitesResponse)["@odata.context"] ===
      "string"
  );
};

export const isM365SharePointSite = (
  value: unknown,
): value is M365SharePointSite => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365SharePointSite).id === "string";
};

export const isM365SharePointSiteMetadata = (
  value: unknown,
): value is M365SharePointSiteMetadata => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365SharePointSiteMetadata).id === "string";
};

export const isGraphSharePointListsResponse = (
  value: unknown,
): value is GraphSharePointListsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphSharePointListsResponse).value) &&
    typeof (value as GraphSharePointListsResponse)["@odata.context"] ===
      "string"
  );
};

export const isM365SharePointListInfo = (
  value: unknown,
): value is M365SharePointListInfo => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365SharePointListInfo).id === "string";
};

export const isGraphSharePointColumnsResponse = (
  value: unknown,
): value is GraphSharePointColumnsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphSharePointColumnsResponse).value) &&
    typeof (value as GraphSharePointColumnsResponse)["@odata.context"] ===
      "string"
  );
};

export const isM365SharePointColumn = (
  value: unknown,
): value is M365SharePointColumn => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365SharePointColumn).displayName === "string";
};
