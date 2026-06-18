import * as React from 'react';
import { IAboutPageProps } from './IAboutPageProps';
import { IHomePageConfig } from '../../services/ConfigService';
import { CACHE_KEYS } from '../../config/appConfig';
import styles from './AboutPage.module.scss';

const categories = [
  { label: 'Case Studies', icon: 'case' },
  { label: 'Capabilities', icon: 'capability' },
  { label: 'Proposals', icon: 'proposal' },
  { label: 'Lessons Learned', icon: 'lessons' },
  { label: 'Knowledge Sharing Sessions', icon: 'ks2' },
];

const DEFAULT_HERO_CONFIG: IHomePageConfig = {
  heroMessage: 'The strength of the team is each individual member. The strength of each member is the team. Alone we can do so little; together we can do so much.',
  quoteAuthor: 'Helen Keller',
  imageUrl: ''
};

type CachedHeroConfig = IHomePageConfig;

type HeroImageSource = {
  src: string;
  shouldRevoke: boolean;
  fromCache: boolean;
};

const HERO_IMAGE_CACHE_NAME = 'ikn-home-hero-image-v1';

const isDefaultHeroConfig = (config: IHomePageConfig): boolean => {
  return (
    (config.heroMessage || '').trim() === DEFAULT_HERO_CONFIG.heroMessage &&
    (config.quoteAuthor || '').trim() === DEFAULT_HERO_CONFIG.quoteAuthor &&
    !(config.imageUrl || '').trim()
  );
};

const parseCachedHeroConfig = (cached: string | null): CachedHeroConfig | null => {
  if (!cached) {
    return null;
  }

  try {
    const config = { ...DEFAULT_HERO_CONFIG, ...JSON.parse(cached) };
    return isDefaultHeroConfig(config) ? null : config;
  } catch {
    return null;
  }
};

const readCachedHeroConfig = (): CachedHeroConfig | null => {
  try {
    return (
      parseCachedHeroConfig(window.sessionStorage.getItem(CACHE_KEYS.homePageConfig)) ||
      parseCachedHeroConfig(window.localStorage.getItem(CACHE_KEYS.homePageConfig))
    );
  } catch {
    return null;
  }
};

const cacheHeroConfig = (config: CachedHeroConfig): void => {
  if (isDefaultHeroConfig(config)) {
    return;
  }

  const storageConfig: CachedHeroConfig = {
    id: config.id,
    heroMessage: config.heroMessage,
    quoteAuthor: config.quoteAuthor,
    imageUrl: config.imageUrl
  };

  try {
    window.sessionStorage.setItem(CACHE_KEYS.homePageConfig, JSON.stringify(storageConfig));
  } catch {
    // Ignore session storage failures.
  }

  try {
    window.localStorage.setItem(CACHE_KEYS.homePageConfig, JSON.stringify(storageConfig));
  } catch {
    // Ignore local storage failures.
  }
};

const waitForImageDecode = async (imageUrl: string): Promise<void> => {
  await new Promise<void>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve();
    image.onerror = () => reject(new Error('Unable to load hero image'));
    image.src = imageUrl;
  });
};

const pruneHeroImageCache = async (currentImageUrl: string): Promise<void> => {
  if (!('caches' in window)) {
    return;
  }

  try {
    const cache = await window.caches.open(HERO_IMAGE_CACHE_NAME);
    const requests = await cache.keys();
    await Promise.all(
      requests.map(request => request.url === currentImageUrl ? Promise.resolve(false) : cache.delete(request))
    );
  } catch {
    // Browser cache cleanup is best-effort only.
  }
};

const cacheHeroImageBlob = async (imageUrl: string, blob: Blob): Promise<void> => {
  if (!imageUrl || !('caches' in window)) {
    return;
  }

  try {
    const cache = await window.caches.open(HERO_IMAGE_CACHE_NAME);
    await cache.put(new Request(imageUrl, { credentials: 'include' }), new Response(blob));
    await pruneHeroImageCache(imageUrl);
  } catch {
    // If Cache API is unavailable or blocked, the normal image URL still works.
  }
};

const loadHeroImageSource = async (imageUrl: string): Promise<HeroImageSource> => {
  if ('caches' in window) {
    try {
      const cache = await window.caches.open(HERO_IMAGE_CACHE_NAME);
      const cachedResponse = await cache.match(new Request(imageUrl, { credentials: 'include' }));
      if (cachedResponse) {
        const cachedObjectUrl = URL.createObjectURL(await cachedResponse.blob());
        await waitForImageDecode(cachedObjectUrl);
        return { src: cachedObjectUrl, shouldRevoke: true, fromCache: true };
      }

      const response = await fetch(imageUrl, { credentials: 'include', cache: 'force-cache' });
      if (response.ok) {
        const responseClone = response.clone();
        const objectUrl = URL.createObjectURL(await response.blob());
        await waitForImageDecode(objectUrl);
        await cache.put(new Request(imageUrl, { credentials: 'include' }), responseClone);
        void pruneHeroImageCache(imageUrl);
        return { src: objectUrl, shouldRevoke: true, fromCache: false };
      }
    } catch {
      // Fall back to the SharePoint URL below.
    }
  }

  await waitForImageDecode(imageUrl);
  return { src: imageUrl, shouldRevoke: false, fromCache: false };
};

export const AboutCategorySection: React.FunctionComponent<Pick<IAboutPageProps, 'onCategorySelect'>> = (props) => {
  const handleCategoryClick = (category: string): void => {
    props.onCategorySelect?.(category);
  };

  return (
    <section className={styles.categorySection} aria-labelledby="browse-category-title">
      <h2 id="browse-category-title" className={styles.categoryTitle}>Browse by category</h2>
      <div className={styles.categoryGrid}>
        {categories.map((category) => (
          <button
            key={category.label}
            type="button"
            className={styles.categoryCard}
            onClick={() => handleCategoryClick(category.label)}
          >
            <span className={styles.categoryIcon} aria-hidden="true">
              {category.icon === 'capability' ? (
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="3" />
                  <path d="M12 2v3M12 19v3M4.22 4.22l2.12 2.12M17.66 17.66l2.12 2.12M2 12h3M19 12h3M4.22 19.78l2.12-2.12M17.66 6.34l2.12-2.12" />
                </svg>
              ) : category.icon === 'lessons' ? (
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z" />
                  <path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z" />
                </svg>
              ) : category.icon === 'ks2' ? (
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
                  <circle cx="9" cy="7" r="4" />
                  <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
                  <path d="M16 3.13a4 4 0 0 1 0 7.75" />
                </svg>
              ) : category.icon === 'template' ? (
                <svg width="38" height="38" viewBox="0 0 24 24" fill="none">
                  <rect x="4" y="5" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="2" />
                  <path d="m12 8 1.4 2.8 3.1.4-2.2 2.1.5 3-2.8-1.4-2.8 1.4.5-3-2.2-2.1 3.1-.4L12 8Z" fill="currentColor" />
                </svg>
              ) : category.icon === 'brochure' ? (
                <svg width="38" height="38" viewBox="0 0 24 24" fill="none">
                  <path d="M5 5.5A2.5 2.5 0 0 1 7.5 3H19v17H7.5A2.5 2.5 0 0 1 5 17.5v-12Z" stroke="currentColor" strokeWidth="2" />
                  <path d="M8 7h7M8 11h7M8 15h4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
              ) : category.icon === 'ebook' ? (
                <svg width="38" height="38" viewBox="0 0 24 24" fill="none">
                  <path d="M5 4h10a3 3 0 0 1 3 3v13H8a3 3 0 0 1-3-3V4Z" stroke="currentColor" strokeWidth="2" />
                  <path d="M18 7h1a2 2 0 0 1 2 2v11h-3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                  <path d="M8 8h6M8 12h5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
              ) : category.icon === 'case' ? (
                <svg width="38" height="38" viewBox="0 0 24 24" fill="none">
                  <path d="M9 6V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v1" stroke="currentColor" strokeWidth="2" />
                  <rect x="4" y="6" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="2" />
                  <circle cx="15.5" cy="14.5" r="2.5" stroke="currentColor" strokeWidth="2" />
                  <path d="m17.5 16.5 2 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
              ) : (
                <svg width="38" height="38" viewBox="0 0 24 24" fill="none">
                  <path d="M9 5h6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                  <path d="M9 5a3 3 0 0 1 6 0" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                  <rect x="5" y="6" width="14" height="15" rx="2" stroke="currentColor" strokeWidth="2" />
                  <path d="m9 15.6 5.3-5.3 1.4 1.4-5.3 5.3H9v-1.4Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
                </svg>
              )}
            </span>
            <span>{category.label}</span>
          </button>
        ))}
      </div>
    </section>
  );
};

export const AboutPage: React.FunctionComponent<IAboutPageProps> = (props) => {
  const cachedHeroConfig = React.useMemo(() => readCachedHeroConfig(), []);
  const [heroConfig, setHeroConfig] = React.useState<IHomePageConfig>(() => cachedHeroConfig || DEFAULT_HERO_CONFIG);
  const [isHeroConfigReady, setIsHeroConfigReady] = React.useState<boolean>(() => Boolean(cachedHeroConfig));
  const [heroImageSrc, setHeroImageSrc] = React.useState<string>('');
  const [isHeroImageReady, setIsHeroImageReady] = React.useState<boolean>(() => !cachedHeroConfig?.imageUrl);
  const [shouldFadeHeroImage, setShouldFadeHeroImage] = React.useState(false);
  const heroImageObjectUrlRef = React.useRef<string>('');
  const [isEditing, setIsEditing] = React.useState(false);
  const [editMessage, setEditMessage] = React.useState('');
  const [editAuthor, setEditAuthor] = React.useState('');
  const [editImageFile, setEditImageFile] = React.useState<File | null>(null);
  const [editImagePreview, setEditImagePreview] = React.useState('');
  const [isSaving, setIsSaving] = React.useState(false);

  const setResolvedHeroImageSrc = React.useCallback((src: string, shouldRevoke: boolean): void => {
    if (heroImageObjectUrlRef.current && heroImageObjectUrlRef.current !== src) {
      URL.revokeObjectURL(heroImageObjectUrlRef.current);
    }
    heroImageObjectUrlRef.current = shouldRevoke ? src : '';
    setHeroImageSrc(src);
  }, []);

  React.useEffect(() => {
    return () => {
      if (heroImageObjectUrlRef.current) {
        URL.revokeObjectURL(heroImageObjectUrlRef.current);
      }
    };
  }, []);

  React.useEffect(() => {
    if (!props.configService) return;
    void props.configService.fetchHomePageConfig()
      .then(config => {
        cacheHeroConfig(config);
        setHeroConfig(config);
        setIsHeroConfigReady(true);
      })
      .catch(() => {
        setIsHeroConfigReady(Boolean(readCachedHeroConfig()));
      });
  }, [props.configService]);

  React.useEffect(() => {
    const imageUrl = heroConfig.imageUrl;

    if (!imageUrl) {
      setResolvedHeroImageSrc('', false);
      setShouldFadeHeroImage(false);
      setIsHeroImageReady(true);
      return;
    }

    setResolvedHeroImageSrc('', false);
    setShouldFadeHeroImage(false);
    setIsHeroImageReady(false);
    let isDisposed = false;

    void loadHeroImageSource(imageUrl)
      .then((imageSource) => {
        if (!isDisposed) {
          setResolvedHeroImageSrc(imageSource.src, imageSource.shouldRevoke);
          setShouldFadeHeroImage(!imageSource.fromCache);
          setIsHeroImageReady(true);
        } else if (imageSource.shouldRevoke) {
          URL.revokeObjectURL(imageSource.src);
        }
      })
      .catch(() => {
        if (!isDisposed) {
          setResolvedHeroImageSrc('', false);
          setShouldFadeHeroImage(false);
          setIsHeroImageReady(true);
        }
      });

    return () => {
      isDisposed = true;
    };
  }, [heroConfig.imageUrl, setResolvedHeroImageSrc]);

  return (
    <div className={styles.aboutPage}>
      <div className={styles.contentContainer}>
        <section className={styles.heroPanel} aria-label="Welcome">
          <div className={styles.heroContent}>
            <h1 className={styles.heroTitle}>Welcome to iKnowledgeNext</h1>
            <p className={styles.heroQuote}>
              {isHeroConfigReady ? (
                <>
                  {heroConfig.heroMessage}
                  {' - '}
                  <span className={styles.heroQuoteAuthor}>{heroConfig.quoteAuthor}</span>
                </>
              ) : null}
            </p>
          </div>
          {(heroImageSrc || (isHeroConfigReady && heroConfig.imageUrl)) ? (
            <div className={styles.heroPortrait} aria-hidden="true">
              <span className={styles.heroImageFrame}>
                {heroImageSrc && isHeroImageReady ? (
                  <img
                    src={heroImageSrc}
                    alt="Hero"
                    className={`${styles.heroImage} ${shouldFadeHeroImage ? styles.heroImageFadeIn : ''}`}
                    loading="eager"
                    decoding="async"
                    draggable={false}
                  />
                ) : null}
              </span>
              {props.isAdmin && (
                <button
                  className={styles.editIconBtn}
                  onClick={() => {
                    setEditMessage(heroConfig.heroMessage);
                    setEditAuthor(heroConfig.quoteAuthor);
                    setEditImagePreview(heroConfig.imageUrl);
                    setIsEditing(true);
                  }}
                  aria-label="Edit hero section"
                  title="Edit"
                  type="button"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
                  </svg>
                </button>
              )}
            </div>
          ) : null}
        </section>

        {isEditing && (
          <div className={styles.modalOverlay} onClick={(e) => {
            if (e.target === e.currentTarget) setIsEditing(false);
          }}>
            <div className={styles.modalBox}>
              <div className={styles.modalHeader}>
                <h3 className={styles.modalTitle}>Edit Home Page</h3>
                <button
                  className={styles.modalCloseBtn}
                  onClick={() => setIsEditing(false)}
                  aria-label="Close"
                  type="button"
                >
                  ✕
                </button>
              </div>

              <div className={styles.modalBody}>
                <label className={styles.modalLabel}>Welcome Message</label>
                <textarea
                  value={editMessage}
                  onChange={e => setEditMessage(e.target.value)}
                  className={styles.modalTextarea}
                  rows={4}
                  placeholder="Enter welcome message"
                />

                <label className={styles.modalLabel}>Quote Author</label>
                <input
                  type="text"
                  value={editAuthor}
                  onChange={e => setEditAuthor(e.target.value)}
                  className={styles.modalInput}
                  placeholder="e.g. Helen Keller"
                />

                <label className={styles.modalLabel}>Profile Photo</label>
                <div className={styles.imageUploadRow}>
                  {editImagePreview && (
                    <img
                      src={editImagePreview}
                      alt="Preview"
                      className={styles.modalImagePreview}
                    />
                  )}
                  <label className={styles.chooseFileBtn}>
                    {editImageFile ? editImageFile.name : 'Upload Image'}
                    <input
                      type="file"
                      accept="image/*"
                      style={{ display: 'none' }}
                      onChange={e => {
                        const file = e.target.files?.[0];
                        if (file) {
                          setEditImageFile(file);
                          setEditImagePreview(URL.createObjectURL(file));
                        }
                      }}
                    />
                  </label>
                </div>
              </div>

              <div className={styles.modalFooter}>
                <button
                  className={styles.modalCancelBtn}
                  onClick={() => setIsEditing(false)}
                  disabled={isSaving}
                  type="button"
                >
                  Cancel
                </button>
                <button
                  className={styles.modalSaveBtn}
                  disabled={isSaving}
                  onClick={async () => {
                    if (!props.configService) return;
                    setIsSaving(true);
                    try {
                      let imageUrl = heroConfig.imageUrl;
                      if (editImageFile) {
                        imageUrl = await props.configService.uploadHeroImage(editImageFile);
                      }
                      const success = await props.configService.saveHomePageConfig({
                        id: heroConfig.id,
                        heroMessage: editMessage,
                        quoteAuthor: editAuthor,
                        imageUrl
                      });
                      if (success) {
                        const nextHeroConfig: CachedHeroConfig = {
                          ...heroConfig,
                          heroMessage: editMessage,
                          quoteAuthor: editAuthor,
                          imageUrl
                        };
                        setHeroConfig(nextHeroConfig);
                        setIsHeroConfigReady(true);
                        if (editImageFile) {
                          await cacheHeroImageBlob(imageUrl, editImageFile);
                          const uploadedImageObjectUrl = URL.createObjectURL(editImageFile);
                          await waitForImageDecode(uploadedImageObjectUrl);
                          setResolvedHeroImageSrc(uploadedImageObjectUrl, true);
                          setShouldFadeHeroImage(false);
                          setIsHeroImageReady(true);
                        }
                        cacheHeroConfig(nextHeroConfig);
                        setIsEditing(false);
                        setEditImageFile(null);
                      }
                    } finally {
                      setIsSaving(false);
                    }
                  }}
                  type="button"
                >
                  {isSaving ? 'Saving...' : 'Save Changes'}
                </button>
              </div>
            </div>
          </div>
        )}

        <div className={`${styles.actionCards} ${props.hideUploadButton ? styles.actionCardsSearchOnly : ''}`} aria-label="Primary actions">
          <button type="button" className={`${styles.actionCard} ${props.hideUploadButton ? styles.actionCardSearchOnly : ''}`} onClick={props.onSearchOpen}>
            <span className={styles.actionIcon} aria-hidden="true">
              <svg width="30" height="30" viewBox="0 0 24 24" fill="none">
                <circle cx="11" cy="11" r="6" stroke="currentColor" strokeWidth="2" />
                <path d="m16 16 4 4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              </svg>
            </span>
            <span className={styles.actionText}>
              <span className={styles.actionTitle}>Search Knowledge</span>
              <span className={styles.actionSubtitle}>Find knowledge that matters</span>
            </span>
          </button>

          {!props.hideUploadButton && (
            <button type="button" className={styles.actionCard} onClick={props.onUploadOpen}>
              <span className={styles.actionIcon} aria-hidden="true">
                <svg width="30" height="30" viewBox="0 0 24 24" fill="none">
                  <path d="M12 16V5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                  <path d="m7 10 5-5 5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M5 19h14" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
              </span>
              <span className={styles.actionText}>
                <span className={styles.actionTitle}>Contribute Knowledge</span>
                <span className={styles.actionSubtitle}>Contribute effortlessly with AI-powered tagging</span>
              </span>
            </button>
          )}
        </div>

        {!props.hideCategorySection && <AboutCategorySection onCategorySelect={props.onCategorySelect} />}

      </div>
    </div>
  );
};



