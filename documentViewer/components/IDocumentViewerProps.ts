import { SPHttpClient } from '@microsoft/sp-http';

export interface IDocumentViewerProps {
  description: string;
  documentUrl: string;
  documents: any[];
  documentsLength: number;
  siteUrl: string;
}
