import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITaxonomyFieldOptions } from '../../../services/TaxonomyService';

export interface IMetadataFormProps {
  context: WebPartContext;
  onSubmit?: (data: Record<string, any>) => void;
  onClose?: () => void;
  initialValues?: Record<string, any>;
  taxonomyOptions: ITaxonomyFieldOptions;
  showReviewerFields?: boolean;
  hideActions?: boolean;
  density?: 'default' | 'compact';
  onDraftChange?: (data: Record<string, any>, isValid: boolean) => void;
}

