import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISpDescrepencyProps {
  context: WebPartContext; // Added context property
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext?: boolean;
  userDisplayName: string;
}
