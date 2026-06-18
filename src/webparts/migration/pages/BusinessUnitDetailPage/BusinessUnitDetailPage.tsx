import * as React from 'react';
import styles from './BusinessUnitDetailPage.module.scss';
import { FileUpload } from '../../components/upload/FileUpload/FileUpload';
import { IKShellFooter, IKShellHeader } from '../../components/shell/IKShellChrome';
import { IBusinessUnitDetailPageProps } from './IBusinessUnitDetailPageProps';
import { NAV_PATHS, pushPageUrl } from '../../services/permalinkService';

const buTabs = [
  'About',
  'Service Lines',
  'Knowledge Base',
  'Onboarding',
  'Quick Links',
  'Key Contacts'
];

const knowledgeBaseItems = [
  {
    label: 'Case Studies',
    icon: 'clipboard'
  },
  {
    label: 'Capability Decks',
    icon: 'sparkle'
  },
  {
    label: 'Proposals',
    icon: 'shield'
  },
  {
    label: 'Templates',
    icon: 'search'
  },
  {
    label: 'Lessons Learned',
    icon: 'book'
  }
];

const quickLinkItems = [
  {
    label: 'Internet page',
    icon: 'page'
  },
  {
    label: 'IdeaStorm',
    icon: 'page'
  },
  {
    label: 'BU-specific portals',
    icon: 'portal'
  },
  {
    label: 'RFIs',
    icon: 'page'
  }
];

const keyContacts = [
  {
    name: 'Sarah Johnson',
    role: 'Senior Product Manager',
    badge: 'Admin',
    employeeId: '#EMP01',
    department: 'Managerial',
    workType: 'Fulltime',
    email: 'sarah.j@indegene.com',
    phone: '+1 (555) 123-4567',
    joined: 'Joined at 15 Jan, 2022',
    variant: 'light'
  },
  {
    name: 'Michael Chen',
    role: 'Lead Software Engineer',
    badge: 'Contributor',
    employeeId: '#EMP02',
    department: 'Technical',
    workType: 'Fulltime',
    email: 'michael.c@indegene.com',
    phone: '+1 (555) 234-5678',
    joined: 'Joined at 03 Mar, 2021',
    variant: 'dark'
  },
  {
    name: 'Priya Patel',
    role: 'UX Designer',
    badge: 'Contributor',
    employeeId: '#EMP03',
    department: 'Creative',
    workType: 'Fulltime',
    email: 'priya.p@indegene.com',
    phone: '+1 (555) 345-6789',
    joined: 'Joined at 20 Jul, 2022',
    variant: 'accent'
  }
];

const renderKnowledgeIcon = (icon: string): JSX.Element => {
  switch (icon) {
    case 'sparkle':
      return (
        <svg viewBox="0 0 72 72" aria-hidden="true">
          <path d="M14 16h44v31H32L19 60V47h-5V16Z" />
          <path d="M39 23l4 9 9 4-9 4-4 9-4-9-9-4 9-4 4-9Z" />
        </svg>
      );
    case 'shield':
      return (
        <svg viewBox="0 0 72 72" aria-hidden="true">
          <path d="M36 8 58 19v15c0 15-9 25-22 31C23 59 14 49 14 34V19L36 8Z" />
          <path d="m26 36 7 7 15-17" />
        </svg>
      );
    case 'search':
      return (
        <svg viewBox="0 0 72 72" aria-hidden="true">
          <path d="M18 17h34v34H18V17Z" />
          <path d="M27 17v-6h16v6" />
          <path d="M41 41a11 11 0 1 0 0-22 11 11 0 0 0 0 22Z" />
          <path d="m49 49 10 10" />
        </svg>
      );
    case 'book':
      return (
        <svg viewBox="0 0 72 72" aria-hidden="true">
          <path d="M14 15h18c7 0 12 5 12 12v31c-3-3-7-5-12-5H14V15Z" />
          <path d="M58 15H40c-7 0-12 5-12 12v31c3-3 7-5 12-5h18V15Z" />
          <path d="M24 28h10" />
          <path d="M24 38h8" />
          <path d="M46 28h7" />
          <path d="M46 38h6" />
        </svg>
      );
    default:
      return (
        <svg viewBox="0 0 72 72" aria-hidden="true">
          <path d="M18 17h36v44H18V17Z" />
          <path d="M28 17v-6h16v6" />
          <path d="m30 45 18-18" />
          <path d="m44 27 5 5" />
        </svg>
      );
  }
};

const renderQuickLinkIcon = (icon: string): JSX.Element => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    {icon === 'portal' ? (
      <>
        <rect x="6" y="4" width="12" height="16" rx="1.8" />
        <path d="M9 8h6M9 11h6M9 14h2M13 14h2" />
      </>
    ) : (
      <>
        <path d="M7 4h7l3 3v13H7V4Z" />
        <path d="M14 4v4h4" />
        <path d="M10 12h5M10 15h5" />
      </>
    )}
  </svg>
);

const renderContactAvatar = (variant: string): JSX.Element => (
  <span className={`${styles.contactAvatar} ${variant === 'dark' ? styles.contactAvatarDark : variant === 'accent' ? styles.contactAvatarAccent : ''}`}>
    <span className={styles.contactAvatarHead} />
    <span className={styles.contactAvatarBody} />
  </span>
);

const BusinessUnitDetailPage: React.FC<IBusinessUnitDetailPageProps> = ({
  selectedBU,
  context,
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
  isLearner = false
}) => {
  const [activeTab, setActiveTab] = React.useState('About');
  const [showUploader, setShowUploader] = React.useState(false);
  const businessUnitName = selectedBU || 'Business Unit Name';

  return (
    <div className={styles.buPage}>
      <div className={styles.pageFrame}>
        <div className={styles.buHeader}>
          <IKShellHeader
            userName={context?.pageContext.user.displayName}
            userEmail={context?.pageContext.user.email || context?.pageContext.user.loginName}
            userPhotoUrl={userPhotoUrl}
            onLogout={onLogout}
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
          />
        </div>

        <main className={styles.main}>
          <section className={styles.hero}>
            <div>
              <h1>{businessUnitName}</h1>
            </div>
            {!isLearner && context && (
              <button type="button" className={styles.uploadButton} onClick={() => setShowUploader(true)}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <path d="M12 16V4m0 0 5 5m-5-5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M5 20h14" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
                Upload
              </button>
            )}
          </section>

          <section className={styles.content}>
            <nav className={styles.tabs} aria-label="Business unit sections">
              {buTabs.map((tab) => (
                <button
                  key={tab}
                  type="button"
                  className={`${styles.tab} ${activeTab === tab ? styles.activeTab : ''}`}
                  onClick={() => setActiveTab(tab)}
                >
                  {tab}
                </button>
              ))}
            </nav>

            {activeTab === 'Key Contacts' ? (
              <div className={styles.keyContactsSection}>
                <h2>Business Heads</h2>
                <div className={styles.contactsGrid}>
                  {keyContacts.map((contact) => (
                    <article key={contact.employeeId} className={styles.contactCard}>
                      <div className={styles.contactTopRow}>
                        <span className={styles.activeBadge}>Active</span>
                        <button type="button" className={styles.contactMenu} aria-label={`More actions for ${contact.name}`}>⋮</button>
                      </div>
                      <div className={styles.contactIdentity}>
                        {renderContactAvatar(contact.variant)}
                        <span className={`${styles.contactRoleBadge} ${contact.badge === 'Admin' ? styles.contactRoleAdmin : ''}`}>
                          {contact.badge}
                        </span>
                        <h3>{contact.name}</h3>
                        <p>{contact.role}</p>
                      </div>
                      <div className={styles.contactInfoPanel}>
                        <div><span>#</span>{contact.employeeId}</div>
                        <div><span>▣</span>{contact.department}<span className={styles.inlineMeta}>◷</span>{contact.workType}</div>
                        <div><span>✉</span>{contact.email}</div>
                        <a href={`tel:${contact.phone.replace(/[^\d+]/g, '')}`}>{contact.phone}</a>
                      </div>
                      <div className={styles.contactFooter}>
                        <span>{contact.joined}</span>
                        <button type="button">View details ›</button>
                      </div>
                    </article>
                  ))}
                </div>
              </div>
            ) : activeTab === 'Quick Links' ? (
              <div className={styles.quickLinksGrid}>
                {quickLinkItems.map((item) => (
                  <button key={item.label} type="button" className={styles.quickLinkCard}>
                    <span className={styles.quickLinkIcon}>{renderQuickLinkIcon(item.icon)}</span>
                    <span className={styles.quickLinkLabel}>{item.label}</span>
                    <span className={styles.quickLinkArrow} aria-hidden="true">→</span>
                  </button>
                ))}
              </div>
            ) : activeTab === 'Onboarding' ? (
              <div className={styles.emptyTabPanel} aria-label="Onboarding content" />
            ) : activeTab === 'Knowledge Base' ? (
              <div className={styles.knowledgeGrid}>
                {knowledgeBaseItems.map((item) => (
                  <button key={item.label} type="button" className={styles.knowledgeCard}>
                    <span className={styles.knowledgeIcon}>{renderKnowledgeIcon(item.icon)}</span>
                    <span>{item.label}</span>
                  </button>
                ))}
              </div>
            ) : activeTab === 'Service Lines' ? (
              <div className={styles.serviceLineContent}>
                <h2>Service Line name</h2>
                <div className={styles.serviceLineCopy}>
                  <p>
                    Our expertise spans across commercial operations, medical affairs, regulatory compliance, and
                    clinical development. We combine deep domain knowledge with cutting-edge technology to deliver
                    solutions that make a real difference in healthcare delivery and patient care.
                  </p>
                  <p>
                    Indegene is a leading healthcare organization dedicated to transforming the way healthcare
                    companies operate through innovative technology solutions and strategic partnerships. With over
                    two decades of experience, we partner with the world's leading life sciences companies to
                    accelerate innovation and drive better patient outcomes.
                  </p>
                </div>
              </div>
            ) : (
              <div className={styles.aboutGrid}>
                <div className={styles.copy}>
                  <blockquote className={styles.quote}>
                    <span>
                      The strength of the team is each individual member. The strength of each member is the team.
                      Alone we can do so little; together we can do so much.
                    </span>
                    <strong>- Helen Keller</strong>
                  </blockquote>

                  <div className={styles.bodyCopy}>
                    <p>
                      Our expertise spans across commercial operations, medical affairs, regulatory compliance, and
                      clinical development. We combine deep domain knowledge with cutting-edge technology to deliver
                      solutions that make a real difference in healthcare delivery and patient care.
                    </p>
                    <p>
                      Indegene is a leading healthcare organization dedicated to transforming the way healthcare
                      companies operate through innovative technology solutions and strategic partnerships. With over
                      two decades of experience, we partner with the world's leading life sciences companies to
                      accelerate innovation and drive better patient outcomes.
                    </p>
                    <p>
                      At Indegene, we believe in the power of collaboration and knowledge sharing. This Knowledge Hub
                      serves as a central repository for our collective expertise, best practices, case studies, and
                      insights that drive excellence across our organization and partnerships.
                    </p>
                  </div>
                </div>

                <div className={styles.leaderPortrait} aria-label={`${businessUnitName} leader profile`}>
                  <div className={styles.portraitFace}>
                    <span className={styles.portraitHair} />
                    <span className={styles.portraitHead} />
                    <span className={styles.portraitBody} />
                  </div>
                </div>
              </div>
            )}
          </section>
        </main>

        <div className={styles.buFooter}>
          <IKShellFooter onBackHome={onHomeOpen || (() => { pushPageUrl(NAV_PATHS.home); })} />
        </div>
      </div>

      {context && showUploader && (
        <FileUpload
          onClose={() => setShowUploader(false)}
          context={context}
          userName={context.pageContext.user.displayName}
          userEmail={context.pageContext.user.email || context.pageContext.user.loginName}
          userPhotoUrl={userPhotoUrl}
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
        />
      )}
    </div>
  );
};

export { BusinessUnitDetailPage };
export default BusinessUnitDetailPage;
