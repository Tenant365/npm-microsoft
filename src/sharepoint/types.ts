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

export interface M365SharePointColumnBaseInput {
  description?: string;
  enforceUniqueValues?: boolean;
  hidden?: boolean;
  indexed?: boolean;
  name: string;
  displayName: string;
  required?: boolean;
}

export interface M365SharePointTextColumnInput
  extends M365SharePointColumnBaseInput {
  text: {
    allowMultipleLines?: boolean;
    appendChangesToExistingText?: boolean;
    linesForEditing?: number;
    maxLength?: number;
    textType?: "plain" | "richText";
  };
}

export interface M365SharePointChoiceColumnInput
  extends M365SharePointColumnBaseInput {
  choice: {
    allowTextEntry?: boolean;
    choices: readonly string[];
    displayAs?: "dropDownMenu" | "radioButtons" | "checkBoxes";
  };
}

export interface M365SharePointBooleanColumnInput
  extends M365SharePointColumnBaseInput {
  boolean: Record<string, never>;
}

export interface M365SharePointNumberColumnInput
  extends M365SharePointColumnBaseInput {
  number: {
    decimalPlaces?: "automatic" | "none" | number;
    displayAs?: "number" | "percentage";
    minimum?: number;
    maximum?: number;
  };
}

export interface M365SharePointDateTimeColumnInput
  extends M365SharePointColumnBaseInput {
  dateTime: {
    displayAs?: "default" | "friendly" | "standard";
    format?: "dateOnly" | "dateTime";
  };
}

export type M365SharePointColumnCreateInput =
  | M365SharePointTextColumnInput
  | M365SharePointChoiceColumnInput
  | M365SharePointBooleanColumnInput
  | M365SharePointNumberColumnInput
  | M365SharePointDateTimeColumnInput;

export interface M365SharePointViewXmlBuilderOptions {
  viewFields?: string[];
  rowLimit?: number;
  whereClauseXml?: string;
  orderByClauseXml?: string;
  baseViewXml?: string;
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

export interface M365SharePointSetViewXmlPayload {
  viewXml: string;
}
