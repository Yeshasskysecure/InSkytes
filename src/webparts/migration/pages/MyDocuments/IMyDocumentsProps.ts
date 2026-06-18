import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMyDocumentsProps {
  context: WebPartContext;
  onClose?: () => void;
  onViewDocument?: (documentId: number) => void;
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
  isLearner?: boolean;
}
