import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IGed365WebpartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  list_title:string;
  backgroundColor:string;
  textColor: string; 
  selectedColumns:string[];

}
