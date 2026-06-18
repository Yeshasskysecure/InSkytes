import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentDetailPageProps {
  context: WebPartContext;
  documentId: number;
  listTitle?: string;
  listId?: string;
  onClose?: () => void;
  onMetricsUpdated?: () => void;
  backTo?: 'home' | 'library' | 'section'; // 'section' returns to the current page section overlay
  onBackToLibrary?: () => void; // Callback to go back to library (ViewAllDocumentsPage)
  onDownloadRecorded?: (documentId: number) => void;
  disableSocialActions?: boolean;
  hideTopDocumentActions?: boolean;
  backButtonLabel?: string;
  onHomeOpen?: () => void;
  onAllDocumentsOpen?: () => void;
  onBusinessUnitsOpen?: () => void;
  onBookmarksOpen?: () => void;
  onDocumentsOpen?: () => void;
  onContactOpen?: () => void;
  onAuditLogOpen?: () => void;
  onAnalyticsOpen?: () => void;
  showReviewerNav?: boolean;
  hideDocumentsNav?: boolean;
  userName?: string;
  userEmail?: string;
  userPhotoUrl?: string;
  onLogout?: () => void;
}

