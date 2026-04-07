import {
  GraphTeamsResponse,
  GraphTeamTemplatesResponse,
  M365Team,
} from "./types";

export const isGraphTeamsResponse = (
  value: unknown,
): value is GraphTeamsResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return (
    Array.isArray((value as GraphTeamsResponse).value) &&
    typeof (value as GraphTeamsResponse)["@odata.context"] === "string"
  );
};

export const isGraphTeamTemplatesResponse = (
  value: unknown,
): value is GraphTeamTemplatesResponse => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return Array.isArray((value as GraphTeamTemplatesResponse).value);
};

export const isM365Team = (value: unknown): value is M365Team => {
  if (!value || typeof value !== "object") {
    return false;
  }

  return typeof (value as M365Team).id === "string";
};
