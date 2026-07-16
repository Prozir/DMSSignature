import type { SPHttpClient } from '@microsoft/sp-http';

export type DocumentDisplayMode = 'detailsList' | 'cards';

export interface ISignatureWebPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail?: string;
  siteUrl?: string;
  taskListName?: string;
  displayMode: DocumentDisplayMode;
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
}
