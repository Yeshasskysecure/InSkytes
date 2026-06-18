import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMigrationProps {
  description: string;
  context: WebPartContext;
  projectId?: string;
  currentPage?: string;
  initialDocumentId?: number | null;
  initialSearchQuery?: string;
}

