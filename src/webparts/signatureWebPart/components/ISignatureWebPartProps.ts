import type { SPHttpClient } from '@microsoft/sp-http';

export interface ISignatureWebPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail?: string;
  siteUrl?: string;
  taskListName?: string;
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
}
