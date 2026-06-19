import * as React from 'react';
import styles from './IKShellChrome.module.scss';
import { NAV_PATHS, pushPageUrl } from '../../services/permalinkService';
import { SITE_URL } from '../../config/appConfig';
import dmSansFontData from '../../assets/fonts/dmSansFontData';
import { openOutlookCompose } from '../../utils/contactActions';

const dmSansFontFace = `
@font-face {
  font-family: 'DM Sans';
  font-style: normal;
  font-display: swap;
  font-weight: 100 1000;
  src: url('data:font/woff2;base64,${dmSansFontData}') format('woff2-variations');
}
`;

const SHOW_BUSINESS_UNITS_NAV = false;
const FOOTER_QUICK_TOUR_URL = `${SITE_URL}/SitePages/Assets.aspx`;
const FOOTER_FAQ_URL = `${SITE_URL}/SitePages/Assets.aspx`;

export type IKShellHeaderNavKey =
  | 'businessUnits'
  | 'bookmarks'
  | 'documents'
  | 'contact'
  | 'adminPanel'
  | 'kmReviewHub'
  | 'kmHarvestHub'
  | 'kmLibrary'
  | 'auditLog'
  | 'analytics';

export interface IIKShellHeaderProps {
  userName?: string;
  userEmail?: string;
  userPhotoUrl?: string;
  compact?: boolean;
  onLogout?: () => void;
  onHomeOpen?: () => void;
  onAllDocumentsOpen?: () => void;
  onKmReviewHubOpen?: () => void;
  onKmHarvestHubOpen?: () => void;
  onKmLibraryOpen?: () => void;
  onBusinessUnitsOpen?: () => void;
  onBookmarksOpen?: () => void;
  onDocumentsOpen?: () => void;
  onContactOpen?: () => void;
  onAuditLogOpen?: () => void;
  onAnalyticsOpen?: () => void;
  showReviewerNav?: boolean;
  hideDocumentsNav?: boolean;
  activeNavKey?: IKShellHeaderNavKey;
}

export interface IIKShellFooterProps {
  onBackHome?: () => void;
  backHomeLabel?: string;
  isHomeFooter?: boolean;
  compact?: boolean;
  className?: string;
}

const navigateTo = (path: string): void => {
  pushPageUrl(path);
};

const openAdminLink = (url: string): void => {
  window.open(url, '_blank', 'noopener,noreferrer');
};

const ADMIN_PANEL_NAV_KEYS: IKShellHeaderNavKey[] = ['adminPanel', 'kmReviewHub', 'kmHarvestHub', 'kmLibrary', 'auditLog', 'analytics'];

const renderNavItem = (
  label: string,
  onClick?: () => void,
  fallbackPath?: string,
  isActive?: boolean,
  itemStyle?: React.CSSProperties
): JSX.Element => {
  const handleClick = onClick || (fallbackPath ? () => navigateTo(fallbackPath) : undefined);

  if (!handleClick) {
    return <span className={isActive ? styles.navItemActive : undefined} style={itemStyle}>{label}</span>;
  }
  
  return (
    <button
      type="button"
      className={isActive ? styles.navItemActive : undefined}
      aria-current={isActive ? 'page' : undefined}
      onClick={handleClick}
      style={itemStyle}
    >
      {label}
    </button>
  );
};

const getInitials = (name: string): string => {
  const parts = name.trim().split(' ').filter(Boolean);
  if (parts.length >= 2) return `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase();
  if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
  return 'U';
};

export const IKShellHeader: React.FC<IIKShellHeaderProps> = (props) => {
  const [isProfileOpen, setIsProfileOpen] = React.useState(false);
  const [showMoreMenu, setShowMoreMenu] = React.useState(false);
  const [isMobileNavOpen, setIsMobileNavOpen] = React.useState(false);
  const profileRef = React.useRef<HTMLDivElement>(null);
  const moreMenuRef = React.useRef<HTMLDivElement>(null);
  const mobileNavRef = React.useRef<HTMLDivElement>(null);
  const displayName = props.userName || 'User';
  const email = props.userEmail || '-';
  const initials = getInitials(displayName);
  const hasPhoto = !!props.userPhotoUrl;
  const isAdminPanelActive = !!props.activeNavKey && ADMIN_PANEL_NAV_KEYS.indexOf(props.activeNavKey) !== -1;
  const navStyle: React.CSSProperties = props.compact
    ? { gap: 'clamp(19.2px, 2vw, 38.4px)', fontSize: '11.2px', lineHeight: '16px' }
    : { gap: 'clamp(24px, 2.5vw, 48px)', fontSize: '14px', lineHeight: '20px' };
  const navItemStyle: React.CSSProperties = props.compact
    ? { fontSize: '11.2px', lineHeight: '16px', maxWidth: '128px' }
    : { fontSize: '14px', lineHeight: '20px', maxWidth: '160px' };
  const adminTriggerStyle: React.CSSProperties = props.compact
    ? { gap: '3.2px', fontSize: '11.2px', lineHeight: '16px' }
    : { gap: '4px', fontSize: '14px', lineHeight: '20px' };
  const moreMenuItemStyle: React.CSSProperties = props.compact
    ? { fontSize: '11.2px', lineHeight: '16px', padding: '8px 11.2px' }
    : { fontSize: '14px', lineHeight: '20px', padding: '10px 14px' };
  const mobileNavPanelStyle: React.CSSProperties = props.compact
    ? { width: 'min(224px, calc(100vw - 16px))', maxWidth: 'calc(100vw - 16px)', padding: '6.4px', gap: '1.6px' }
    : { width: 'min(280px, calc(100vw - 16px))', maxWidth: 'calc(100vw - 16px)', padding: '8px', gap: '2px' };
  const mobileNavItemStyle: React.CSSProperties = props.compact
    ? { fontSize: '11.2px', lineHeight: '1.35', minHeight: '35px', padding: '8.8px 9.6px' }
    : { fontSize: '14px', lineHeight: '1.35', minHeight: '44px', padding: '11px 12px' };
  const mobileNavSectionLabelStyle: React.CSSProperties = props.compact
    ? { fontSize: '8.8px', lineHeight: '11.2px', minHeight: '22.4px', padding: '4.8px 9.6px 3.2px' }
    : { fontSize: '11px', lineHeight: '14px', minHeight: '26px', padding: '6px 12px 4px' };

  React.useEffect(() => {
    const handleClickOutside = (e: MouseEvent): void => {
      if (moreMenuRef.current && !moreMenuRef.current.contains(e.target as Node)) {
        setShowMoreMenu(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  React.useEffect(() => {
    if (!isMobileNavOpen) {
      return;
    }

    const handleDocumentMouseDown = (event: MouseEvent): void => {
      if (mobileNavRef.current && !mobileNavRef.current.contains(event.target as Node)) {
        setIsMobileNavOpen(false);
      }
    };

    document.addEventListener('mousedown', handleDocumentMouseDown);
    return () => document.removeEventListener('mousedown', handleDocumentMouseDown);
  }, [isMobileNavOpen]);

  React.useEffect(() => {
    if (!isProfileOpen) {
      return;
    }

    const handleDocumentMouseDown = (event: MouseEvent): void => {
      if (profileRef.current && !profileRef.current.contains(event.target as Node)) {
        setIsProfileOpen(false);
      }
    };

    document.addEventListener('mousedown', handleDocumentMouseDown);
    return () => document.removeEventListener('mousedown', handleDocumentMouseDown);
  }, [isProfileOpen]);

  const handleLogout = (): void => {
    setIsProfileOpen(false);
    setIsMobileNavOpen(false);
    if (props.onLogout) {
      props.onLogout();
      return;
    }

    window.location.href = '/_layouts/15/SignOut.aspx';
  };

  const runMobileNavAction = (onClick?: () => void, fallbackPath?: string): void => {
    setIsMobileNavOpen(false);
    setShowMoreMenu(false);

    if (onClick) {
      onClick();
      return;
    }

    if (fallbackPath) {
      navigateTo(fallbackPath);
    }
  };

  const runMobileAdminAction = (
    navKey: IKShellHeaderNavKey,
    onClick?: () => void,
    fallbackPath?: string
  ): void => {
    if (navKey === 'kmReviewHub' && !onClick && props.onAllDocumentsOpen) {
      runMobileNavAction(props.onAllDocumentsOpen);
      return;
    }

    if (navKey === 'kmHarvestHub' && !onClick) {
      runMobileNavAction(() => window.dispatchEvent(new CustomEvent('ikn:open-km-harvest-hub')));
      return;
    }

    runMobileNavAction(onClick, fallbackPath);
  };

  return (
    <header className={`${styles.topBar} ${props.compact ? styles.topBarCompact : ''}`}>
      <style>{dmSansFontFace}</style>
      <button
        type="button"
        className={styles.brand}
        onClick={props.onHomeOpen || (() => navigateTo(NAV_PATHS.home))}
        aria-label="Go to InSkytes home"
      >
        <img 
          src={require('../../assets/indegene_logo.png')} 
          alt="Indegene" 
          className={styles.logoImage} 
        />
        <span className={styles.brandDivider}>|</span>
        <img
          src={require('../../assets/iknowledgenext_logo.png')}
          alt="InSkytes"
          className={styles.productLogoImage}
        />
      </button>
      <div className={styles.headerActions}>
        <nav className={styles.nav} style={navStyle} aria-label="InSkytes navigation">
          {SHOW_BUSINESS_UNITS_NAV && renderNavItem('Business Units', props.onBusinessUnitsOpen, NAV_PATHS.businessUnits, props.activeNavKey === 'businessUnits', navItemStyle)}
          {renderNavItem('My Bookmarks', props.onBookmarksOpen, NAV_PATHS.bookmark, props.activeNavKey === 'bookmarks', navItemStyle)}
          {!props.hideDocumentsNav && renderNavItem('My Documents', props.onDocumentsOpen, NAV_PATHS.myDocuments, props.activeNavKey === 'documents', navItemStyle)}
          {renderNavItem('Contact Us', props.onContactOpen || openOutlookCompose, undefined, props.activeNavKey === 'contact', navItemStyle)}
          {props.showReviewerNav && (
            <div className={styles.moreMenuWrapper} ref={moreMenuRef}>
              <button
                className={`${styles.moreMenuTrigger} ${isAdminPanelActive ? styles.moreMenuTriggerActive : ''}`}
                onClick={() => setShowMoreMenu(prev => !prev)}
                aria-expanded={showMoreMenu}
                aria-haspopup="true"
                aria-current={isAdminPanelActive ? 'page' : undefined}
                type="button"
                style={adminTriggerStyle}
              >
                Admin Panel
                <svg
                  className={`${styles.moreMenuChevron} ${showMoreMenu ? styles.moreMenuChevronOpen : ''}`}
                  width="12" height="12" viewBox="0 0 24 24"
                  fill="none" stroke="currentColor" strokeWidth="2.5"
                  strokeLinecap="round" strokeLinejoin="round"
                >
                  <polyline points="6 9 12 15 18 9" />
                </svg>
              </button>

              {showMoreMenu && (
                <div className={styles.moreMenuDropdown} role="menu">
                  <button
                    className={`${styles.moreMenuItem} ${props.activeNavKey === 'kmReviewHub' ? styles.moreMenuItemActive : ''}`}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      if (props.onKmReviewHubOpen) props.onKmReviewHubOpen();
                      else if (props.onAllDocumentsOpen) props.onAllDocumentsOpen();
                    }}
                  >
                    KM Review Hub
                  </button>
                  <button
                    className={`${styles.moreMenuItem} ${props.activeNavKey === 'kmHarvestHub' ? styles.moreMenuItemActive : ''}`}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      if (props.onKmHarvestHubOpen) props.onKmHarvestHubOpen();
                      else window.dispatchEvent(new CustomEvent('ikn:open-km-harvest-hub'));
                    }}
                  >
                    KM Harvest Hub
                  </button>
                  <button
                    className={`${styles.moreMenuItem} ${props.activeNavKey === 'kmLibrary' ? styles.moreMenuItemActive : ''}`}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      if (props.onKmLibraryOpen) props.onKmLibraryOpen();
                      else navigateTo(NAV_PATHS.kmLibrary);
                    }}
                  >
                    KM Library
                  </button>
                  <div className={styles.moreMenuDivider} />
                  <button
                    className={`${styles.moreMenuItem} ${props.activeNavKey === 'auditLog' ? styles.moreMenuItemActive : ''}`}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      if (props.onAuditLogOpen) props.onAuditLogOpen();
                    }}
                  >
                    Audit Log
                  </button>
                  <button
                    className={`${styles.moreMenuItem} ${props.activeNavKey === 'analytics' ? styles.moreMenuItemActive : ''}`}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      if (props.onAnalyticsOpen) props.onAnalyticsOpen();
                    }}
                  >
                    Analytics
                  </button>
                  <div className={styles.moreMenuDivider} />
                  <button
                    className={styles.moreMenuItem}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      openAdminLink(`${SITE_URL}/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter`);
                    }}
                  >
                    Term Store
                  </button>
                  <button
                    className={styles.moreMenuItem}
                    role="menuitem"
                    type="button"
                    style={moreMenuItemStyle}
                    onClick={() => {
                      setShowMoreMenu(false);
                      openAdminLink(`${SITE_URL}/_layouts/15/settings.aspx`);
                    }}
                  >
                    Site Settings
                  </button>
                </div>
              )}
            </div>
          )}
        </nav>
        <div className={styles.mobileNavWrapper} ref={mobileNavRef}>
          <button
            type="button"
            className={styles.mobileMenuButton}
            aria-label="Open navigation menu"
            aria-expanded={isMobileNavOpen}
            aria-haspopup="menu"
            onClick={() => {
              setIsMobileNavOpen((value) => !value);
              setIsProfileOpen(false);
              setShowMoreMenu(false);
            }}
          >
            <span />
            <span />
            <span />
          </button>
          {isMobileNavOpen && (
            <div className={styles.mobileNavPanel} style={mobileNavPanelStyle} role="menu" aria-label="Navigation menu">
              {SHOW_BUSINESS_UNITS_NAV && (
                <button
                  type="button"
                  role="menuitem"
                  className={`${styles.mobileNavItem} ${props.activeNavKey === 'businessUnits' ? styles.mobileNavItemActive : ''}`}
                  style={mobileNavItemStyle}
                  onClick={() => runMobileNavAction(props.onBusinessUnitsOpen, NAV_PATHS.businessUnits)}
                >
                  Business Units
                </button>
              )}
              <button
                type="button"
                role="menuitem"
                className={`${styles.mobileNavItem} ${props.activeNavKey === 'bookmarks' ? styles.mobileNavItemActive : ''}`}
                style={mobileNavItemStyle}
                onClick={() => runMobileNavAction(props.onBookmarksOpen, NAV_PATHS.bookmark)}
              >
                My Bookmarks
              </button>
              {!props.hideDocumentsNav && (
                <button
                  type="button"
                  role="menuitem"
                  className={`${styles.mobileNavItem} ${props.activeNavKey === 'documents' ? styles.mobileNavItemActive : ''}`}
                  style={mobileNavItemStyle}
                  onClick={() => runMobileNavAction(props.onDocumentsOpen, NAV_PATHS.myDocuments)}
                >
                  My Documents
                </button>
              )}
              <button
                type="button"
                role="menuitem"
                className={`${styles.mobileNavItem} ${props.activeNavKey === 'contact' ? styles.mobileNavItemActive : ''}`}
                style={mobileNavItemStyle}
                onClick={() => runMobileNavAction(props.onContactOpen || openOutlookCompose)}
              >
                Contact Us
              </button>
              {props.showReviewerNav && (
                <>
                  <div className={styles.mobileNavDivider} />
                  <div className={styles.mobileNavSectionLabel} style={mobileNavSectionLabelStyle}>Admin Panel</div>
                  <button
                    type="button"
                    role="menuitem"
                    className={`${styles.mobileNavItem} ${props.activeNavKey === 'kmReviewHub' ? styles.mobileNavItemActive : ''}`}
                    style={mobileNavItemStyle}
                    onClick={() => runMobileAdminAction('kmReviewHub', props.onKmReviewHubOpen)}
                  >
                    KM Review Hub
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    className={`${styles.mobileNavItem} ${props.activeNavKey === 'kmHarvestHub' ? styles.mobileNavItemActive : ''}`}
                    style={mobileNavItemStyle}
                    onClick={() => runMobileAdminAction('kmHarvestHub', props.onKmHarvestHubOpen)}
                  >
                    KM Harvest Hub
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    className={`${styles.mobileNavItem} ${props.activeNavKey === 'kmLibrary' ? styles.mobileNavItemActive : ''}`}
                    style={mobileNavItemStyle}
                    onClick={() => runMobileAdminAction('kmLibrary', props.onKmLibraryOpen, NAV_PATHS.kmLibrary)}
                  >
                    KM Library
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    className={`${styles.mobileNavItem} ${props.activeNavKey === 'auditLog' ? styles.mobileNavItemActive : ''}`}
                    style={mobileNavItemStyle}
                    onClick={() => runMobileAdminAction('auditLog', props.onAuditLogOpen)}
                  >
                    Audit Log
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    className={`${styles.mobileNavItem} ${props.activeNavKey === 'analytics' ? styles.mobileNavItemActive : ''}`}
                    style={mobileNavItemStyle}
                    onClick={() => runMobileAdminAction('analytics', props.onAnalyticsOpen)}
                  >
                    Analytics
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    className={styles.mobileNavItem}
                    style={mobileNavItemStyle}
                    onClick={() => {
                      setIsMobileNavOpen(false);
                      openAdminLink(`${SITE_URL}/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter`);
                    }}
                  >
                    Term Store
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    className={styles.mobileNavItem}
                    style={mobileNavItemStyle}
                    onClick={() => {
                      setIsMobileNavOpen(false);
                      openAdminLink(`${SITE_URL}/_layouts/15/settings.aspx`);
                    }}
                  >
                    Site Settings
                  </button>
                </>
              )}
            </div>
          )}
        </div>
        <div
          className={styles.profileMenu}
          ref={profileRef}
        >
          <button
            type="button"
            className={styles.userIconButton}
            onClick={() => {
              setIsProfileOpen((value) => !value);
              setIsMobileNavOpen(false);
            }}
            aria-label="Open profile menu"
            aria-haspopup="menu"
            aria-expanded={isProfileOpen}
          >
            {hasPhoto ? (
              <img
                src={props.userPhotoUrl}
                alt={displayName}
                className={styles.userPhoto}
              />
            ) : (
              <span className={styles.userInitials}>{initials}</span>
            )}
          </button>
          {isProfileOpen && (
            <div className={styles.profileDropdown} role="menu">
              <button
                type="button"
                className={styles.profileClose}
                onClick={() => setIsProfileOpen(false)}
                aria-label="Close profile menu"
              >
                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <path d="M7 7l10 10M17 7L7 17" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
                </svg>
              </button>
              <div className={styles.profileIdentity}>
                <span className={styles.profileAvatar} aria-hidden="true">
                  {hasPhoto ? (
                    <img
                      src={props.userPhotoUrl}
                      alt={displayName}
                      className={styles.profileAvatarPhoto}
                    />
                  ) : (
                    <span className={styles.profileAvatarInitials}>{initials}</span>
                  )}
                </span>
                <div className={styles.profileName}>{displayName}</div>
              </div>
              <div className={styles.profileEmail}>{email}</div>
              <button type="button" className={styles.profileLogout} onClick={handleLogout} role="menuitem">
                Logout
              </button>
            </div>
          )}
        </div>
      </div>
    </header>
  );
};

export const IKShellFooter: React.FC<IIKShellFooterProps> = (props) => {
  const footerStyle: React.CSSProperties = props.compact
    ? {
        gap: '22.4px',
        fontSize: '11.2px',
        lineHeight: '16px',
        minHeight: '57.6px',
        height: '57.6px',
        padding: '0 68.8px',
        justifyContent: 'flex-end',
        flexWrap: 'nowrap'
      }
    : { gap: '28px', fontSize: '14px', lineHeight: '20px' };
  const footerItemStyle: React.CSSProperties = props.compact
    ? { fontSize: '11.2px', lineHeight: '16px' }
    : { fontSize: '14px', lineHeight: '20px' };
  const footerButtonStyle: React.CSSProperties = props.compact
    ? { minHeight: '24px', padding: '2.4px 14.4px', borderRadius: '2.4px', fontSize: '11.2px', lineHeight: '16px' }
    : { minHeight: '30px', padding: '3px 18px', fontSize: '14px', lineHeight: '20px' };

  return (
  <footer className={`${styles.footer} ${props.compact ? styles.footerCompact : ''} ${props.isHomeFooter ? styles.footerHome : ''} ${props.className || ''}`} style={footerStyle}>
    <a
      href={SITE_URL}
      target="_blank"
      rel="noopener noreferrer"
      className={styles.footerLink}
      style={footerItemStyle}
    >
      KM HelpDesk
    </a>
    <a
      href={FOOTER_QUICK_TOUR_URL}
      target="_blank"
      rel="noopener noreferrer"
      className={styles.footerLink}
      style={footerItemStyle}
    >
      Quick Tour
    </a>
    <a
      href={FOOTER_FAQ_URL}
      target="_blank"
      rel="noopener noreferrer"
      className={styles.footerLink}
      style={footerItemStyle}
    >
      FAQ
    </a>
    {props.onBackHome && (
      <button type="button" className={styles.footerButton} style={footerButtonStyle} onClick={props.onBackHome}>
        {props.backHomeLabel || 'Back to homepage'}
      </button>
    )}
  </footer>
  );
};
