export interface IDocumentDataChangedDetail {
  documentIds?: number[];
  reason?: 'upload' | 'metadata' | 'bookmark' | 'unknown';
  timestamp: number;
}

const DOCUMENT_DATA_CHANGED_EVENT = 'iknowledge:document-data-changed';

export const emitDocumentDataChanged = (detail: Omit<IDocumentDataChangedDetail, 'timestamp'> = {}): void => {
  if (typeof window === 'undefined') {
    return;
  }

  window.dispatchEvent(
    new CustomEvent<IDocumentDataChangedDetail>(DOCUMENT_DATA_CHANGED_EVENT, {
      detail: {
        ...detail,
        timestamp: Date.now()
      }
    })
  );
};

export const subscribeToDocumentDataChanged = (
  handler: (detail: IDocumentDataChangedDetail) => void
): (() => void) => {
  if (typeof window === 'undefined') {
    return () => undefined;
  }

  const listener = (event: Event): void => {
    const customEvent = event as CustomEvent<IDocumentDataChangedDetail>;
    handler(customEvent.detail || { timestamp: Date.now() });
  };

  window.addEventListener(DOCUMENT_DATA_CHANGED_EVENT, listener as EventListener);

  return () => {
    window.removeEventListener(DOCUMENT_DATA_CHANGED_EVENT, listener as EventListener);
  };
};
