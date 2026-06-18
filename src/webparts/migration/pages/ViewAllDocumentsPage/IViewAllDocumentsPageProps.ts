import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IViewAllDocumentsPageProps {
  context: WebPartContext;
  onClose: () => void;
  onViewDocument: (documentId: number, tags?: string[]) => void;
  onDownloadRecorded?: (documentId: number) => void;
  scope?: 'all' | 'kmReviewHub' | 'recentlyPublished' | 'myDoc' | 'bookmarks';
  isLearner?: boolean;
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
  onLogout?: () => void;
  userPhotoUrl?: string;
}

