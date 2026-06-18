import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IRecentlyPublishedSectionProps {
  context?: WebPartContext;
  sectionStyle?: React.CSSProperties;
  isLearner?: boolean;
  onViewAllOpen?: () => void;
  onViewDocument?: (documentId: number) => void;
}

