import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IRenameBusinessUnitDialogProps {
  context: WebPartContext;
  isOpen: boolean;
  onClose: () => void;
  variant?: 'modal' | 'inline';
}
