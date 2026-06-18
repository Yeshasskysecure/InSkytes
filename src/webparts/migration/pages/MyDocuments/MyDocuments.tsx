import * as React from 'react';
import { IMyDocumentsProps } from './IMyDocumentsProps';
import { ViewAllDocumentsPage } from '../ViewAllDocumentsPage/ViewAllDocumentsPage';
import { NAV_PATHS, openAppPageInNewTab } from '../../services/permalinkService';

export const MyDocuments: React.FC<IMyDocumentsProps> = ({ 
  context, 
  onClose, 
  onViewDocument,
  onHomeOpen,
  onAllDocumentsOpen,
  onBusinessUnitsOpen,
  onBookmarksOpen,
  onDocumentsOpen,
  onContactOpen,
  onAuditLogOpen,
  onAnalyticsOpen,
  showReviewerNav,
  hideDocumentsNav,
  onLogout,
  userPhotoUrl,
  isLearner
}) => {
  const handleViewDocument = React.useCallback((documentId: number): void => {
    if (onViewDocument) {
      onViewDocument(documentId);
      return;
    }

    openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
  }, [onViewDocument]);

  const handleClose = React.useCallback((): void => {
    if (onClose) {
      onClose();
      return;
    }

    window.dispatchEvent(new PopStateEvent('popstate'));
  }, [onClose]);

  return (
    <ViewAllDocumentsPage
      context={context}
      scope="myDoc"
      onClose={handleClose}
      onViewDocument={handleViewDocument}
      onHomeOpen={onHomeOpen}
      onAllDocumentsOpen={onAllDocumentsOpen}
      onBusinessUnitsOpen={onBusinessUnitsOpen}
      onBookmarksOpen={onBookmarksOpen}
      onDocumentsOpen={onDocumentsOpen}
      onContactOpen={onContactOpen}
      onAuditLogOpen={onAuditLogOpen}
      onAnalyticsOpen={onAnalyticsOpen}
      showReviewerNav={showReviewerNav}
      hideDocumentsNav={hideDocumentsNav}
      onLogout={onLogout}
      userPhotoUrl={userPhotoUrl}
      isLearner={isLearner}
    />
  );
};

export default MyDocuments;
