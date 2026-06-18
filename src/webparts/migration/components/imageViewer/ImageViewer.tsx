import * as React from 'react';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react';
import styles from '../pdfViewer/PdfViewer.module.scss';

export interface IImageViewerProps {
  fileUrl: string;
  fileName?: string;
}

const MIN_SCALE = 0.5;
const MAX_SCALE = 2.5;
const SCALE_STEP = 0.15;
const PINCH_ZOOM_SENSITIVITY = 0.0015;

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

export const ImageViewer: React.FunctionComponent<IImageViewerProps> = ({
  fileUrl,
  fileName
}) => {
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const mainPanelRef = React.useRef<HTMLElement | null>(null);
  const gestureStartScaleRef = React.useRef<number>(1);
  const [scale, setScale] = React.useState<number>(1);
  const [rotation, setRotation] = React.useState<number>(0);
  const [showThumbnails, setShowThumbnails] = React.useState<boolean>(false);
  const [isFullscreen, setIsFullscreen] = React.useState<boolean>(false);
  const [zoomInputValue, setZoomInputValue] = React.useState<string>('100');
  const [isEditingZoom, setIsEditingZoom] = React.useState<boolean>(false);
  const [naturalSize, setNaturalSize] = React.useState<{ width: number; height: number } | null>(null);
  const [panelSize, setPanelSize] = React.useState<{ width: number; height: number }>({ width: 0, height: 0 });

  React.useEffect(() => {
    setScale(1);
    setRotation(0);
    setShowThumbnails(false);
    setNaturalSize(null);
  }, [fileUrl]);

  React.useEffect(() => {
    const handleFullscreenChange = (): void => {
      setIsFullscreen(document.fullscreenElement === containerRef.current);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  const setZoomImmediately = React.useCallback((nextScale: number): void => {
    setScale(clampScale(nextScale));
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
    const panel = mainPanelRef.current;
    if (!panel || typeof ResizeObserver === 'undefined') {
      return;
    }

    const updatePanelSize = (): void => {
      setPanelSize({
        width: panel.clientWidth,
        height: panel.clientHeight
      });
    };

    updatePanelSize();
    const observer = new ResizeObserver(updatePanelSize);
    observer.observe(panel);
    return () => observer.disconnect();
  }, []);

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) {
      return;
    }

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
      setScale((currentScale) => clampScale(currentScale - event.deltaY * PINCH_ZOOM_SENSITIVITY));
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
      setScale(clampScale(gestureStartScaleRef.current * gestureScale));
    };

    window.addEventListener('wheel', handleWheel, { capture: true, passive: false });
    window.addEventListener('gesturestart', handleGestureStart, { capture: true, passive: false } as AddEventListenerOptions);
    window.addEventListener('gesturechange', handleGestureChange, { capture: true, passive: false } as AddEventListenerOptions);

    return () => {
      window.removeEventListener('wheel', handleWheel, true);
      window.removeEventListener('gesturestart', handleGestureStart, true);
      window.removeEventListener('gesturechange', handleGestureChange, true);
    };
  }, [scale]);

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

  const imageLayout = React.useMemo(() => {
    const sourceWidth = naturalSize?.width || 900;
    const sourceHeight = naturalSize?.height || 640;
    const availableWidth = Math.max(panelSize.width - 48, 320);
    const availableHeight = Math.max(panelSize.height - 48, 240);
    const fitScale = Math.min(availableWidth / sourceWidth, availableHeight / sourceHeight, 1);
    const displayWidth = Math.max(Math.round(sourceWidth * fitScale * scale), 1);
    const displayHeight = Math.max(Math.round(sourceHeight * fitScale * scale), 1);
    const isSideways = Math.abs(rotation % 180) === 90;

    return {
      imageWidth: displayWidth,
      imageHeight: displayHeight,
      shellWidth: isSideways ? displayHeight : displayWidth,
      shellHeight: isSideways ? displayWidth : displayHeight
    };
  }, [naturalSize, panelSize, rotation, scale]);

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
    }, 'Zoom in')
  ], [commitZoomInput, isEditingZoom, scale, setZoomImmediately, zoomInputValue]);

  const rightToolbarItems = React.useMemo<ICommandBarItemProps[]>(() => [
    makeIconOnlyToolbarItem({
      key: 'fullscreen',
      iconProps: { iconName: isFullscreen ? 'BackToWindow' : 'FullScreen' },
      onClick: toggleFullscreen
    }, isFullscreen ? 'Exit fullscreen' : 'Fullscreen')
  ], [isFullscreen, toggleFullscreen]);

  return (
    <div
      ref={containerRef}
      className={`${styles.pdfViewer} ${isFullscreen ? styles.pdfViewerFullscreen : ''}`}
      data-testid="custom-image-viewer"
    >
      <div className={styles.toolbarShell}>
        <div className={styles.toolbarLayout}>
          <CommandBar className={styles.toolbarCommandBar} items={leftToolbarItems} ariaLabel="Image viewer left controls" />
          <CommandBar className={styles.toolbarCommandBar} items={centerToolbarItems} ariaLabel="Image viewer zoom controls" />
          <CommandBar className={styles.toolbarCommandBar} items={rightToolbarItems} ariaLabel="Image viewer right controls" />
        </div>
      </div>
      <div className={styles.viewerBody}>
        {showThumbnails && (
          <aside className={styles.thumbnailPanel} aria-label="Image thumbnails">
            <h3 className={styles.thumbnailPanelHeader}>Pages</h3>
            <button type="button" className={`${styles.thumbnailButton} ${styles.thumbnailButtonActive}`}>
              <span className={styles.thumbnailPreview}>
                <img src={fileUrl} alt="" style={{ maxWidth: '100%', maxHeight: '142px', objectFit: 'contain' }} />
              </span>
              <span className={styles.thumbnailLabel}>Page 1</span>
            </button>
          </aside>
        )}
        <main ref={mainPanelRef} className={styles.mainPanel} aria-label={fileName || 'Image preview'}>
          <div className={styles.imageCanvasShell}>
            <div
              className={styles.imageSurface}
              style={{
                width: imageLayout.shellWidth,
                height: imageLayout.shellHeight
              }}
            >
              <img
                src={fileUrl}
                alt={fileName || 'Image preview'}
                className={styles.imagePreview}
                style={{
                  width: imageLayout.imageWidth,
                  height: imageLayout.imageHeight,
                  transform: `rotate(${rotation}deg)`
                }}
                onLoad={(event) => {
                  const image = event.currentTarget;
                  setNaturalSize({
                    width: image.naturalWidth || imageLayout.imageWidth,
                    height: image.naturalHeight || imageLayout.imageHeight
                  });
                }}
              />
            </div>
          </div>
        </main>
      </div>
    </div>
  );
};
