import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBudgetDashProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  selectedBudgetList: string;
  selectedKontiList: string;
  //selectedList: string;
}
