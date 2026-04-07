export interface M365Team {
  id?: string;
  displayName?: string;
  description?: string;
  [key: string]: unknown;
}

export interface M365TeamMetadata {
  id: string;
  displayName?: string;
  description?: string;
  [key: string]: unknown;
}

export interface M365TeamTemplate {
  id: string;
  [key: string]: unknown;
}

export interface M365CreateTeamMemberInput {
  userId: string;
  roles?: ("owner" | "member")[];
}

export type M365TeamVisibility = "private" | "public";

export interface M365CreateTeamInput {
  displayName: string;
  description?: string;
  templateId?: string;
  templateOdataBind?: string;
  members: M365CreateTeamMemberInput[];
  visibility?: M365TeamVisibility;
  image?: string;
}

export interface M365CreateTeamProvisionAccepted {
  status: 202;
  operationLocation: string | null;
  contentLocation: string | null;
  body?: unknown;
}

export type M365CreateTeamResult = M365Team | M365CreateTeamProvisionAccepted;

export interface GraphTeamsResponse {
  "@odata.context"?: string;
  value: M365Team[];
}

export interface GraphTeamTemplatesResponse {
  value: M365TeamTemplate[];
}
