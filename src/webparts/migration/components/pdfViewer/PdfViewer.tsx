import * as React from 'react';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react';
import { Document, Page, pdfjs } from 'react-pdf';
import * as pdfWorkerModule from 'pdfjs-dist/build/pdf.worker.mjs';
import styles from './PdfViewer.module.scss';

export interface IPdfViewerProps {
  fileUrl: string;
  fileName?: string;
  canPrint?: boolean;
  fitToPageOnLoad?: boolean;
}

interface IPdfThumbnailListProps {
  numPages: number;
  pageNumber: number;
  rotation: number;
  onSelectPage: (pageNumber: number) => void;
}

interface ILazyPdfPageProps {
  pageNumber: number;
  activePageNumber: number;
  renderScale: number;
  scaleRatio: number;
  rotation: number;
  pageSize?: PdfPageSize;
  scrollRootRef: React.RefObject<HTMLElement>;
  setPageRef: (pageNumber: number, element: HTMLDivElement | null) => void;
  onPageClick: (event: React.MouseEvent<HTMLElement>, pageNumber: number) => void;
  onGoToPage: (pageNumber: number) => void;
  onPageLoad: (pageNumber: number, page: any) => void;
}

interface ILazyThumbnailPreviewProps {
  pageNumber: number;
  activePageNumber: number;
  rotation: number;
}

type PdfPageSize = {
  width: number;
  height: number;
};

const MIN_SCALE = 0.5;
const MAX_SCALE = 2.5;
const SCALE_STEP = 0.15;
const THUMBNAIL_WIDTH = 126;
const PINCH_ZOOM_SENSITIVITY = 0.0015;
const ESTIMATED_PAGE_WIDTH = 640;
const ESTIMATED_PAGE_HEIGHT = 880;
const PAGE_RENDER_MARGIN = '1100px 0px';
const THUMBNAIL_RENDER_MARGIN = '500px 0px';

(globalThis as any).pdfjsWorker = {
  WorkerMessageHandler: (pdfWorkerModule as any).WorkerMessageHandler
};
pdfjs.GlobalWorkerOptions.workerSrc = './pdf.worker.mjs';

const clampScale = (value: number): number =>
  Math.min(Math.max(Number(value.toFixed(2)), MIN_SCALE), MAX_SCALE);

const makeIconOnlyToolbarItem = (
  item: ICommandBarItemProps,
  label: string
): ICommandBarItemProps => ({
  ...item,
  text: '',
  ariaLabel: label,
  title: label,
  iconOnly: true
});

const PDF_DOCUMENT_OPTIONS = {
  disableAutoFetch: true,
  disableRange: false,
  disableStream: true,
  isEvalSupported: false,
  rangeChunkSize: 262144,
  useSystemFonts: true
};

const getPageShellSize = (
  pageSize: PdfPageSize | undefined,
  renderScale: number,
  rotation: number
): React.CSSProperties => {
  const isSideways = Math.abs(rotation % 180) === 90;
  const baseWidth = pageSize?.width || ESTIMATED_PAGE_WIDTH;
  const baseHeight = pageSize?.height || ESTIMATED_PAGE_HEIGHT;

  return {
    width: Math.round((isSideways ? baseHeight : baseWidth) * renderScale),
    minHeight: Math.round((isSideways ? baseWidth : baseHeight) * renderScale)
  };
};

const getPdfPageSize = (page: any): PdfPageSize | null => {
  const viewport = page?.getViewport ? page.getViewport({ scale: 1, rotation: 0 }) : null;
  const pageWidth = Number(viewport?.width || page?.width || 0);
  const pageHeight = Number(viewport?.height || page?.height || 0);

  if (pageWidth <= 0 || pageHeight <= 0) {
    return null;
  }

  return { width: pageWidth, height: pageHeight };
};

const LazyThumbnailPreview: React.FunctionComponent<ILazyThumbnailPreviewProps> = React.memo(({
  pageNumber,
  activePageNumber,
  rotation
}) => {
  const containerRef = React.useRef<HTMLSpanElement | null>(null);
  const [shouldRender, setShouldRender] = React.useState(pageNumber === activePageNumber);

  React.useEffect(() => {
    if (shouldRender || pageNumber === activePageNumber) {
      setShouldRender(true);
      return;
    }

    const element = containerRef.current;
    if (!element || typeof IntersectionObserver === 'undefined') {
      setShouldRender(true);
      return;
    }

    const observer = new IntersectionObserver((entries) => {
      if (entries.some((entry) => entry.isIntersecting)) {
        setShouldRender(true);
        observer.disconnect();
      }
    }, { root: null, rootMargin: THUMBNAIL_RENDER_MARGIN, threshold: 0.01 });

    observer.observe(element);
    return () => observer.disconnect();
  }, [activePageNumber, pageNumber, shouldRender]);

  return (
    <span ref={containerRef} className={styles.thumbnailPreview} aria-hidden="true">
      {shouldRender ? (
        <Page
          pageNumber={pageNumber}
          width={THUMBNAIL_WIDTH}
          rotate={rotation}
          renderAnnotationLayer={false}
          renderTextLayer={false}
          loading={null}
          error={null}
        />
      ) : (
        <span className={styles.thumbnailPlaceholder} />
      )}
    </span>
  );
});

LazyThumbnailPreview.displayName = 'LazyThumbnailPreview';

const PdfThumbnailList: React.FunctionComponent<IPdfThumbnailListProps> = React.memo(({
  numPages,
  pageNumber,
  rotation,
  onSelectPage
}) => (
  <aside className={styles.thumbnailPanel} aria-label="PDF thumbnails">
    <h3 className={styles.thumbnailPanelHeader}>Pages</h3>
    <div className={styles.thumbnailList}>
      {Array.from({ length: numPages }, (_, index) => index + 1).map((thumbnailPageNumber) => (
        <button
          key={`pdf-thumbnail-${thumbnailPageNumber}`}
          type="button"
          className={`${styles.thumbnailButton} ${thumbnailPageNumber === pageNumber ? styles.thumbnailButtonActive : ''}`}
          onClick={() => onSelectPage(thumbnailPageNumber)}
          aria-label={`Go to page ${thumbnailPageNumber}`}
        >
          <LazyThumbnailPreview
            pageNumber={thumbnailPageNumber}
            activePageNumber={pageNumber}
            rotation={rotation}
          />
          <span className={styles.thumbnailLabel}>Page {thumbnailPageNumber}</span>
        </button>
      ))}
    </div>
  </aside>
));

PdfThumbnailList.displayName = 'PdfThumbnailList';

const LazyPdfPage: React.FunctionComponent<ILazyPdfPageProps> = React.memo(({
  pageNumber,
  activePageNumber,
  renderScale,
  scaleRatio,
  rotation,
  pageSize,
  scrollRootRef,
  setPageRef,
  onPageClick,
  onGoToPage,
  onPageLoad
}) => {
  const localRef = React.useRef<HTMLDivElement | null>(null);
  const [shouldRender, setShouldRender] = React.useState(pageNumber <= 2 || pageNumber === activePageNumber);
  const shellSize = React.useMemo(
    () => getPageShellSize(pageSize, renderScale, rotation),
    [pageSize, renderScale, rotation]
  );

  const assignRef = React.useCallback((element: HTMLDivElement | null): void => {
    localRef.current = element;
    setPageRef(pageNumber, element);
  }, [pageNumber, setPageRef]);

  React.useEffect(() => {
    if (shouldRender || pageNumber === activePageNumber) {
      setShouldRender(true);
      return;
    }

    const element = localRef.current;
    if (!element || typeof IntersectionObserver === 'undefined') {
      setShouldRender(true);
      return;
    }

    const observer = new IntersectionObserver((entries) => {
      if (entries.some((entry) => entry.isIntersecting)) {
        setShouldRender(true);
        observer.disconnect();
      }
    }, {
      root: scrollRootRef.current,
      rootMargin: PAGE_RENDER_MARGIN,
      threshold: 0.01
    });

    observer.observe(element);
    return () => observer.disconnect();
  }, [activePageNumber, pageNumber, scrollRootRef, shouldRender]);

  return (
    <div
      ref={assignRef}
      className={styles.pageShell}
      style={{
        ...shellSize,
        transform: `scale(${scaleRatio})`,
        transformOrigin: 'top center'
      }}
      onClick={(event) => onPageClick(event, pageNumber)}
      role="button"
      tabIndex={0}
      aria-label={`Page ${pageNumber}`}
      onKeyDown={(event) => {
        if (event.key === 'ArrowLeft') {
          event.preventDefault();
          onGoToPage(pageNumber - 1);
        }
        if (event.key === 'ArrowRight') {
          event.preventDefault();
          onGoToPage(pageNumber + 1);
        }
      }}
    >
      {shouldRender ? (
        <Page
          pageNumber={pageNumber}
          scale={renderScale}
          rotate={rotation}
          renderAnnotationLayer={false}
          renderTextLayer={false}
          loading={<div className={styles.pagePlaceholder}>Loading page...</div>}
          error={<div className={styles.pagePlaceholder}>Unable to render this page.</div>}
          onLoadSuccess={(page) => onPageLoad(pageNumber, page)}
        />
      ) : (
        <div className={styles.pagePlaceholder}>Page {pageNumber}</div>
      )}
    </div>
  );
});

LazyPdfPage.displayName = 'LazyPdfPage';

export const PdfViewer: React.FunctionComponent<IPdfViewerProps> = ({
  fileUrl,
  fileName,
  canPrint = true,
  fitToPageOnLoad = false
}) => {
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const mainPanelRef = React.useRef<HTMLElement | null>(null);
  const pageRefs = React.useRef<Record<number, HTMLDivElement | null>>({});
  const gestureStartScaleRef = React.useRef<number>(1);
  const pendingScaleRef = React.useRef<number | null>(null);
  const zoomFrameRef = React.useRef<number | null>(null);
  const zoomCommitTimeoutRef = React.useRef<number | null>(null);
  const hasUserAdjustedZoomRef = React.useRef<boolean>(false);
  const printFrameRef = React.useRef<HTMLIFrameElement | null>(null);
  const isPrintFrameLoadedRef = React.useRef<boolean>(false);
  const [numPages, setNumPages] = React.useState<number>(0);
  const [pageNumber, setPageNumber] = React.useState<number>(1);
  const [scale, setScale] = React.useState<number>(1);
  const [renderScale, setRenderScale] = React.useState<number>(1);
  const [zoomInputValue, setZoomInputValue] = React.useState<string>('100');
  const [isEditingZoom, setIsEditingZoom] = React.useState<boolean>(false);
  const [rotation, setRotation] = React.useState<number>(0);
  const [showThumbnails, setShowThumbnails] = React.useState<boolean>(false);
  const [isFullscreen, setIsFullscreen] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');
  const [pageSizes, setPageSizes] = React.useState<Record<number, PdfPageSize>>({});
  const [firstPageSize, setFirstPageSize] = React.useState<PdfPageSize | null>(null);
  const [hasAppliedInitialFit, setHasAppliedInitialFit] = React.useState<boolean>(false);

  React.useEffect(() => {
    pageRefs.current = {};
    setNumPages(0);
    setPageNumber(1);
    setScale(1);
    setRenderScale(1);
    setRotation(0);
    hasUserAdjustedZoomRef.current = false;
    setLoadError('');
    setPageSizes({});
    setFirstPageSize(null);
    setHasAppliedInitialFit(false);
  }, [fileUrl]);

  const applyFitToPageScale = React.useCallback((pageSize: PdfPageSize): void => {
    const panel = mainPanelRef.current;
    if (!fitToPageOnLoad || !panel || pageSize.width <= 0 || pageSize.height <= 0) {
      return;
    }

    const panelStyle = window.getComputedStyle(panel);
    const horizontalPadding =
      parseFloat(panelStyle.paddingLeft || '0') +
      parseFloat(panelStyle.paddingRight || '0');
    const verticalPadding =
      parseFloat(panelStyle.paddingTop || '0') +
      parseFloat(panelStyle.paddingBottom || '0');
    const availableWidth = Math.max(panel.clientWidth - horizontalPadding - 8, 320);
    const availableHeight = Math.max(panel.clientHeight - verticalPadding - 8, 260);
    const nextScale = clampScale(Math.min(
      availableWidth / pageSize.width,
      availableHeight / pageSize.height
    ));
    setScale(nextScale);
    setRenderScale(nextScale);
  }, [fitToPageOnLoad]);

  const handlePageLoad = React.useCallback((loadedPageNumber: number, page: any): void => {
    const pageSize = getPdfPageSize(page);
    if (!pageSize) {
      return;
    }

    setPageSizes((currentSizes) => {
      const currentSize = currentSizes[loadedPageNumber];
      if (currentSize?.width === pageSize.width && currentSize.height === pageSize.height) {
        return currentSizes;
      }

      return {
        ...currentSizes,
        [loadedPageNumber]: pageSize
      };
    });

    if (loadedPageNumber !== 1) {
      return;
    }

    setFirstPageSize(pageSize);
    if (fitToPageOnLoad && !hasAppliedInitialFit) {
      applyFitToPageScale(pageSize);
      setHasAppliedInitialFit(true);
    }
  }, [applyFitToPageScale, fitToPageOnLoad, hasAppliedInitialFit]);

  React.useEffect(() => {
    if (!fitToPageOnLoad || !firstPageSize || !hasAppliedInitialFit || typeof ResizeObserver === 'undefined') {
      return;
    }

    const panel = mainPanelRef.current;
    if (!panel) {
      return;
    }

    const observer = new ResizeObserver(() => {
      if (hasUserAdjustedZoomRef.current) {
        return;
      }

      applyFitToPageScale(firstPageSize);
    });

    observer.observe(panel);
    return () => observer.disconnect();
  }, [applyFitToPageScale, firstPageSize, fitToPageOnLoad, hasAppliedInitialFit]);

  React.useEffect(() => {
    if (!canPrint || !fileUrl || numPages <= 0) {
      return;
    }

    let iframe: HTMLIFrameElement | null = null;
    let idleId: number | null = null;
    let timeoutId: number | null = null;

    const createPrintFrame = (): void => {
      iframe = document.createElement('iframe');
      iframe.style.position = 'fixed';
      iframe.style.right = '0';
      iframe.style.bottom = '0';
      iframe.style.width = '0';
      iframe.style.height = '0';
      iframe.style.border = '0';
      iframe.style.visibility = 'hidden';
      iframe.setAttribute('aria-hidden', 'true');
      iframe.src = fileUrl;

      isPrintFrameLoadedRef.current = false;
      iframe.onload = () => {
        isPrintFrameLoadedRef.current = true;
      };

      printFrameRef.current?.remove();
      printFrameRef.current = iframe;
      document.body.appendChild(iframe);
    };

    if ('requestIdleCallback' in window) {
      idleId = (window as any).requestIdleCallback(createPrintFrame, { timeout: 5000 });
    } else {
      timeoutId = window.setTimeout(createPrintFrame, 4000);
    }

    return () => {
      if (idleId !== null && 'cancelIdleCallback' in window) {
        (window as any).cancelIdleCallback(idleId);
      }
      if (timeoutId !== null) {
        window.clearTimeout(timeoutId);
      }
      iframe?.remove();
      if (iframe && printFrameRef.current === iframe) {
        printFrameRef.current = null;
        isPrintFrameLoadedRef.current = false;
      }
    };
  }, [canPrint, fileUrl, numPages]);

  React.useEffect(() => {
    const handleFullscreenChange = (): void => {
      setIsFullscreen(document.fullscreenElement === containerRef.current);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  const file = React.useMemo(() => ({
    url: fileUrl,
    withCredentials: true
  }), [fileUrl]);

  const pageNumbers = React.useMemo(
    () => Array.from({ length: numPages }, (_, index) => index + 1),
    [numPages]
  );

  const goToPage = React.useCallback((nextPageNumber: number): void => {
    const safePageNumber = Math.min(Math.max(nextPageNumber, 1), Math.max(numPages, 1));
    setPageNumber(safePageNumber);

    const pageElement = pageRefs.current[safePageNumber];
    const scrollRoot = mainPanelRef.current;
    if (pageElement && scrollRoot) {
      const pageBounds = pageElement.getBoundingClientRect();
      const scrollRootBounds = scrollRoot.getBoundingClientRect();
      scrollRoot.scrollTo({
        top: scrollRoot.scrollTop + pageBounds.top - scrollRootBounds.top,
        behavior: 'smooth'
      });
    }
  }, [numPages]);

  const setPageRef = React.useCallback((targetPageNumber: number, element: HTMLDivElement | null): void => {
    pageRefs.current[targetPageNumber] = element;
  }, []);

  const setZoomImmediately = React.useCallback((nextScale: number): void => {
    hasUserAdjustedZoomRef.current = true;
    const clampedScale = clampScale(nextScale);
    setScale(clampedScale);
    setRenderScale(clampedScale);
  }, []);

  React.useEffect(() => {
    if (!isEditingZoom) {
      setZoomInputValue(String(Math.round(scale * 100)));
    }
  }, [isEditingZoom, scale]);

  const commitZoomInput = React.useCallback((): void => {
    const numericValue = Number(zoomInputValue.replace('%', '').trim());
    if (Number.isFinite(numericValue)) {
      setZoomImmediately(numericValue / 100);
    } else {
      setZoomInputValue(String(Math.round(scale * 100)));
    }
    setIsEditingZoom(false);
  }, [scale, setZoomImmediately, zoomInputValue]);

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) {
      return;
    }

    const queueScaleUpdate = (nextScale: number): void => {
      hasUserAdjustedZoomRef.current = true;
      pendingScaleRef.current = clampScale(nextScale);

      if (zoomFrameRef.current !== null) {
        return;
      }

      zoomFrameRef.current = window.requestAnimationFrame(() => {
        zoomFrameRef.current = null;
        const pendingScale = pendingScaleRef.current;
        pendingScaleRef.current = null;

        if (pendingScale !== null) {
          setScale(pendingScale);

          if (zoomCommitTimeoutRef.current !== null) {
            window.clearTimeout(zoomCommitTimeoutRef.current);
          }

          zoomCommitTimeoutRef.current = window.setTimeout(() => {
            setRenderScale(pendingScale);
            zoomCommitTimeoutRef.current = null;
          }, 220);
        }
      });
    };

    const handleWheel = (event: WheelEvent): void => {
      const target = event.target as Node | null;
      if (!target || !container.contains(target)) {
        return;
      }

      if (!event.ctrlKey && !event.metaKey) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      const baseScale = pendingScaleRef.current ?? scale;
      queueScaleUpdate(baseScale - event.deltaY * PINCH_ZOOM_SENSITIVITY);
    };

    const handleGestureStart = (event: Event): void => {
      const target = event.target as Node | null;
      if (!target || !container.contains(target)) {
        return;
      }

      event.preventDefault();
      gestureStartScaleRef.current = scale;
    };

    const handleGestureChange = (event: Event): void => {
      const target = event.target as Node | null;
      if (!target || !container.contains(target)) {
        return;
      }

      event.preventDefault();
      const gestureScale = Number((event as any).scale || 1);
      queueScaleUpdate(gestureStartScaleRef.current * gestureScale);
    };

    window.addEventListener('wheel', handleWheel, { capture: true, passive: false });
    window.addEventListener('gesturestart', handleGestureStart, { capture: true, passive: false } as AddEventListenerOptions);
    window.addEventListener('gesturechange', handleGestureChange, { capture: true, passive: false } as AddEventListenerOptions);

    return () => {
      window.removeEventListener('wheel', handleWheel, true);
      window.removeEventListener('gesturestart', handleGestureStart, true);
      window.removeEventListener('gesturechange', handleGestureChange, true);
      if (zoomFrameRef.current !== null) {
        window.cancelAnimationFrame(zoomFrameRef.current);
        zoomFrameRef.current = null;
      }
      if (zoomCommitTimeoutRef.current !== null) {
        window.clearTimeout(zoomCommitTimeoutRef.current);
        zoomCommitTimeoutRef.current = null;
      }
    };
  }, [scale]);

  const handlePageAreaClick = React.useCallback((event: React.MouseEvent<HTMLElement>, clickedPageNumber: number): void => {
    if (numPages <= 1) {
      return;
    }

    const bounds = event.currentTarget.getBoundingClientRect();
    const isLeftSide = event.clientX - bounds.left < bounds.width / 2;
    goToPage(clickedPageNumber + (isLeftSide ? -1 : 1));
  }, [goToPage, numPages]);

  const syncPageFromScroll = React.useCallback((): void => {
    const panel = mainPanelRef.current;
    if (!panel || numPages <= 0) {
      return;
    }

    const panelBounds = panel.getBoundingClientRect();
    const panelCenterY = panelBounds.top + panelBounds.height / 2;
    let nearestPageNumber = pageNumber;
    let nearestDistance = Number.POSITIVE_INFINITY;

    pageNumbers.forEach((candidatePageNumber) => {
      const pageElement = pageRefs.current[candidatePageNumber];
      if (!pageElement) {
        return;
      }

      const bounds = pageElement.getBoundingClientRect();
      const pageCenterY = bounds.top + bounds.height / 2;
      const distance = Math.abs(pageCenterY - panelCenterY);

      if (distance < nearestDistance) {
        nearestDistance = distance;
        nearestPageNumber = candidatePageNumber;
      }
    });

    if (nearestPageNumber !== pageNumber) {
      setPageNumber(nearestPageNumber);
    }
  }, [numPages, pageNumber, pageNumbers]);

  const printPdf = React.useCallback((): void => {
    if (!canPrint || !fileUrl) {
      return;
    }

    const openPrintFallback = (): void => {
      const printWindow = window.open(fileUrl, '_blank', 'noopener,noreferrer');
      printWindow?.focus();
    };

    const printFrame = (iframe: HTMLIFrameElement): boolean => {
      try {
        iframe.contentWindow?.focus();
        iframe.contentWindow?.print();
        return true;
      } catch (error) {
        console.warn('Direct PDF print failed, opening PDF in a new tab:', error);
        return false;
      }
    };

    const existingPrintFrame = printFrameRef.current;
    if (existingPrintFrame && isPrintFrameLoadedRef.current) {
      if (!printFrame(existingPrintFrame)) {
        openPrintFallback();
      }
      return;
    }

    if (existingPrintFrame) {
      existingPrintFrame.onload = () => {
        isPrintFrameLoadedRef.current = true;
        if (!printFrame(existingPrintFrame)) {
          openPrintFallback();
        }
      };
      return;
    }

    openPrintFallback();
  }, [canPrint, fileUrl]);

  const toggleFullscreen = React.useCallback((): void => {
    const container = containerRef.current;
    if (!container) {
      return;
    }

    if (document.fullscreenElement === container) {
      void document.exitFullscreen?.();
      return;
    }

    void container.requestFullscreen?.();
  }, []);

  const leftToolbarItems = React.useMemo<ICommandBarItemProps[]>(() => [
    makeIconOnlyToolbarItem({
      key: 'thumbnails',
      iconProps: { iconName: 'BulletedList' },
      checked: showThumbnails,
      onClick: () => setShowThumbnails((value) => !value)
    }, 'Thumbnails'),
    makeIconOnlyToolbarItem({
      key: 'rotate',
      iconProps: { iconName: 'Rotate' },
      onClick: () => setRotation((value) => (value + 90) % 360)
    }, 'Rotate')
  ], [showThumbnails]);

  const centerToolbarItems = React.useMemo<ICommandBarItemProps[]>(() => [
    makeIconOnlyToolbarItem({
      key: 'zoomOut',
      iconProps: { iconName: 'ZoomOut' },
      disabled: scale <= MIN_SCALE,
      onClick: () => setZoomImmediately(scale - SCALE_STEP)
    }, 'Zoom out'),
    {
      key: 'zoomValue',
      onRender: () => (
        <label className={styles.toolbarZoomValue} title="Edit zoom percentage">
          <input
            className={styles.toolbarZoomInput}
            value={isEditingZoom ? zoomInputValue : String(Math.round(scale * 100))}
            inputMode="numeric"
            aria-label="Zoom percentage"
            onFocus={(event) => {
              setIsEditingZoom(true);
              setZoomInputValue(String(Math.round(scale * 100)));
              event.currentTarget.select();
            }}
            onChange={(event) => setZoomInputValue(event.currentTarget.value)}
            onBlur={commitZoomInput}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.currentTarget.blur();
              } else if (event.key === 'Escape') {
                setZoomInputValue(String(Math.round(scale * 100)));
                setIsEditingZoom(false);
                event.currentTarget.blur();
              }
            }}
          />
          <span aria-hidden="true">%</span>
        </label>
      )
    },
    makeIconOnlyToolbarItem({
      key: 'zoomIn',
      iconProps: { iconName: 'ZoomIn' },
      disabled: scale >= MAX_SCALE,
      onClick: () => setZoomImmediately(scale + SCALE_STEP)
    }, 'Zoom in'),
    {
      key: 'zoomPageSpacer',
      onRender: () => <span className={styles.toolbarGroupSpacer} aria-hidden="true" />
    },
    makeIconOnlyToolbarItem({
      key: 'previous',
      iconProps: { iconName: 'ChevronLeft' },
      disabled: pageNumber <= 1,
      onClick: () => goToPage(pageNumber - 1)
    }, 'Previous'),
    {
      key: 'pageCount',
      onRender: () => (
        <span className={styles.toolbarPageCount}>
          {numPages ? `${pageNumber} of ${numPages}` : 'Loading...'}
        </span>
      )
    },
    makeIconOnlyToolbarItem({
      key: 'next',
      iconProps: { iconName: 'ChevronRight' },
      disabled: !numPages || pageNumber >= numPages,
      onClick: () => goToPage(pageNumber + 1)
    }, 'Next')
  ], [commitZoomInput, goToPage, isEditingZoom, numPages, pageNumber, scale, setZoomImmediately, zoomInputValue]);

  const rightToolbarItems = React.useMemo<ICommandBarItemProps[]>(() => [
    makeIconOnlyToolbarItem({
      key: 'fullscreen',
      iconProps: { iconName: isFullscreen ? 'BackToWindow' : 'FullScreen' },
      onClick: toggleFullscreen
    }, isFullscreen ? 'Exit fullscreen' : 'Fullscreen'),
    makeIconOnlyToolbarItem({
      key: 'print',
      iconProps: { iconName: 'Print' },
      disabled: !canPrint,
      onClick: printPdf
    }, 'Print')
  ], [canPrint, isFullscreen, printPdf, toggleFullscreen]);

  const keepToolbarItemsVisible = React.useCallback((): undefined => undefined, []);

  return (
    <div
      ref={containerRef}
      className={`${styles.pdfViewer} ${isFullscreen ? styles.pdfViewerFullscreen : ''}`}
      data-testid="custom-pdf-viewer"
    >
      <div className={styles.toolbarShell}>
        <div className={styles.toolbarLayout}>
          <CommandBar
            className={styles.toolbarCommandBar}
            items={leftToolbarItems}
            onReduceData={keepToolbarItemsVisible}
            ariaLabel="PDF viewer left controls"
          />
          <CommandBar
            className={styles.toolbarCommandBar}
            items={centerToolbarItems}
            onReduceData={keepToolbarItemsVisible}
            ariaLabel="PDF viewer zoom and page controls"
          />
          <CommandBar
            className={styles.toolbarCommandBar}
            items={rightToolbarItems}
            onReduceData={keepToolbarItemsVisible}
            ariaLabel="PDF viewer right controls"
          />
        </div>
      </div>

      {loadError ? (
        <div className={styles.stateMessage}>{loadError}</div>
      ) : (
        <Document
          className={styles.documentRoot}
          file={file}
          options={PDF_DOCUMENT_OPTIONS}
          loading={<div className={`${styles.stateMessage} ${styles.documentLoadingState}`}>Loading PDF...</div>}
          error={<div className={`${styles.stateMessage} ${styles.documentLoadingState}`}>Unable to load PDF preview.</div>}
          onLoadSuccess={({ numPages: loadedPages }) => {
            setNumPages(loadedPages);
            setPageNumber((currentPage) => Math.min(Math.max(currentPage, 1), loadedPages));
            setLoadError('');
          }}
          onLoadError={(error) => {
            console.error('Unable to load PDF:', error);
            setLoadError('Unable to load PDF preview.');
          }}
        >
          <div className={styles.viewerBody}>
            {showThumbnails && numPages > 0 && (
              <PdfThumbnailList
                numPages={numPages}
                pageNumber={pageNumber}
                rotation={rotation}
                onSelectPage={goToPage}
              />
            )}

            <main
              ref={mainPanelRef}
              className={styles.mainPanel}
              aria-label={fileName || 'PDF document'}
              onScroll={syncPageFromScroll}
            >
              <div className={styles.pageCanvasShell}>
                {pageNumbers.map((visiblePageNumber) => (
                  <LazyPdfPage
                    key={`pdf-page-${visiblePageNumber}`}
                    pageNumber={visiblePageNumber}
                    activePageNumber={pageNumber}
                    renderScale={renderScale}
                    scaleRatio={scale / renderScale}
                    rotation={rotation}
                    pageSize={pageSizes[visiblePageNumber]}
                    scrollRootRef={mainPanelRef}
                    setPageRef={setPageRef}
                    onPageClick={handlePageAreaClick}
                    onGoToPage={goToPage}
                    onPageLoad={handlePageLoad}
                  />
                ))}
              </div>
            </main>
          </div>
        </Document>
      )}
    </div>
  );
};
