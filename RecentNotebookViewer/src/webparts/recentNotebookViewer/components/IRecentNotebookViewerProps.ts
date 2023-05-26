import { SPHttpClient } from '@microsoft/sp-http';

export interface IRecentNotebookViewerProps {
  description: string;
  imgSrc: string;
  siteURL: string;
  folderRelativeURL: string;
  spHttpClient: SPHttpClient;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
