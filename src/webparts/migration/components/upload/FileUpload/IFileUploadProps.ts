import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFileUploadProps {
  onClose?: () => void;
  onUploaded?: (file: File) => void;
  /** SPFx WebPart context used to call SharePoint REST APIs */
  context: WebPartContext;
  targetItemId?: number;
  targetFileRef?: string;
  projectId?: string | null;
  sidebarOffset?: number;
  variant?: 'default' | 'kmArtifact' | 'replace';
  flowDensity?: 'default' | 'compact';
  showReviewerFields?: boolean;
  onAllDocumentsOpen?: () => void;
  onBusinessUnitsOpen?: () => void;
  onBookmarksOpen?: () => void;
  onDocumentsOpen?: () => void;
  onContactOpen?: () => void;
  onAuditLogOpen?: () => void;
  onAnalyticsOpen?: () => void;
  onHomeOpen?: () => void;
  showReviewerNav?: boolean;
  hideDocumentsNav?: boolean;
  userName?: string;
  userEmail?: string;
  userPhotoUrl?: string;
  onLogout?: () => void;
}
