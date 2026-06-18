import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ConfigService } from '../../../services/ConfigService';

export interface IMainContentRouterProps {
  activePage: string;
  context?: WebPartContext;
  onSearchOpen?: () => void;
  onUploadOpen?: () => void;
  onCategorySelect?: (category: string) => void;
  hideUploadButton?: boolean;
  hideCategorySection?: boolean;
  isAdmin?: boolean;
  configService?: ConfigService | null;
}

