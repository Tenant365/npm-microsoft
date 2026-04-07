export interface M365SharePointSite {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  [key: string]: unknown;
}

export interface M365SharePointSiteMetadata {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  [key: string]: unknown;
}

export interface GraphSharePointSitesResponse {
  "@odata.context"?: string;
  value: M365SharePointSite[];
}

export interface M365SharePointListInfo {
  id: string;
  displayName?: string;
  list?: {
    template?: string;
    hidden?: boolean;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

export interface GraphSharePointListsResponse {
  "@odata.context"?: string;
  value: M365SharePointListInfo[];
}

export interface M365SharePointColumn {
  id?: string;
  name?: string;
  displayName?: string;
  [key: string]: unknown;
}

export interface GraphSharePointColumnsResponse {
  "@odata.context"?: string;
  value: M365SharePointColumn[];
}

export interface M365SharePointViewXmlBuilderOptions {
  viewFields?: string[];
  rowLimit?: number;
  whereClauseXml?: string;
  orderByClauseXml?: string;
}

export interface M365SharePointListViewXmlRequest {
  siteWebUrl: string;
  listId: string;
  viewId: string;
}

export interface M365SharePointListViewInfo {
  Id: string;
  Title?: string;
  DefaultView?: boolean;
  [key: string]: unknown;
}
