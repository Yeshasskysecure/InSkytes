/**
 * Central application configuration.
 * Change values here to deploy to a new tenant, site, or library.
 * Do NOT hardcode these values anywhere else in the codebase.
 */

// Site
export const SITE_URL = 'https://skysecuretech.sharepoint.com/sites/InSkytes';
export const SITE_RELATIVE_URL = '/sites/InSkytes';

// Site Pages
export const PAGE_URLS = {
  home: '/sites/InSkytes/SitePages/Home.aspx',
  businessUnits: '/sites/InSkytes/SitePages/Business-Units.aspx',
  categories: '/sites/InSkytes/SitePages/Categories.aspx',
  search: '/sites/InSkytes/SitePages/Search.aspx',
  assets: '/sites/InSkytes/SitePages/Assets.aspx',
  bookmark: '/sites/InSkytes/SitePages/Bookmark.aspx',
  myDocuments: '/sites/InSkytes/SitePages/MyDocuments.aspx',
  analytics: '/sites/InSkytes/SitePages/Analytics.aspx',
  auditLog: '/sites/InSkytes/SitePages/Audit-Log.aspx',
  kmReviewHub: '/sites/InSkytes/SitePages/KM-Review-Hub.aspx',
  kmHarvestHub: '/sites/InSkytes/SitePages/KM-Harvest-Hub.aspx',
  kmLibrary: '/sites/InSkytes/SitePages/KM-Library.aspx',
};

export interface IPowerBIPageConfig {
  name: string;
  pageName: string;
  sectionName?: string;
  reportId?: string;
  groupId?: string;
}

export const IFRAME_URLS: {
  kmLibrary: string;
  powerBI: {
    baseEmbedUrl: string;
    reportId: string;
    groupId: string;
    autoAuth: string;
    ctid: string;
    pages: IPowerBIPageConfig[];
  };
} = {
  kmLibrary: 'https://skysecuretech.sharepoint.com/sites/InSkytes/KM%20Review%20Hub/Forms/AllItems.aspx',

  // Power BI Analytics
  powerBI: {
    baseEmbedUrl: 'https://app.powerbi.com/reportEmbed',
    reportId: '',
    groupId: '',
    autoAuth: 'true',
    ctid: '',
    pages: [],
  },
};

export const KM_REVIEW_HUB_DRIVE_ID = 'b!6vc8nWEL2UqHgEM2Jexqb7BiCJCRWNlHmsKDr8aR48jmBJ5sRvZbT6YQlOLusHmY';

// Library Names
export const LIBRARY_NAMES = {
  kmDataHub: 'KM Review Hub',
  kmHarvestHub: 'KM Harvest Hub',
};

// List Names
export const LIST_NAMES = {
  documentViews: 'DocumentViews',
  documentLikes: 'DocumentLikes',
  documentComments: 'DocumentComments',
  documentDownloads: 'DocumentDownloads',
  documentMetrics: 'DocumentMetrics',
  documentShares: 'DocumentShares',
  documentBookmarks: 'DocumentBookmarks',
  documentFollows: 'DocumentFollows',
  userInteractions: 'UserInteractions',
  auditLog: 'Audit Log',
  notificationQueue: 'KM_NotificationQueue',
  syncConfig: 'KM_SyncConfig',
  searchQueryHistoryLog: 'Search_Query_History_Log',
  faqs: 'FAQs',
  quickLinks: 'QuickLinks',
  homePageConfig: 'HomePageConfig',
};

// List Paths
export const LIST_PATHS = {
  kmHarvestHub: '/sites/InSkytes/Lists/KM%20Harvest%20Hub/AllItems.aspx',
};

// Column Internal Names
export const COLUMN_NAMES = {
  // ── existing columns (keep these) ──────────────────
  status: 'Status',
  author: 'Author0',
  authorId: 'AuthorId',
  documentType: 'Document_x0020_Type',
  documentTypeId: 'Document_x0020_TypeId',
  published: 'Published',
  businessUnit: 'BU',
  department: 'Department_x0020__x002f__x0020_Sub_x0020_Department',
  client: 'Client',
  geography: 'Geography',
  therapyArea: 'Therapy_x0020_Area',
  diseaseArea: 'Disease_x0020_Area',
  url: 'URL',
  fileRef: 'FileRef',
  fileLeafRef: 'FileLeafRef',
  created: 'Created',
  modified: 'Modified',
  edited: 'Edited',
  version: 'OData__UIVersionString',
  documentId: 'DocumentId',
  userId: 'UserId',
  userName: 'UserName',
  comment: 'Comment',
  tag: 'Tag',

  // ── NEW: KM Review Hub columns ─────────────────────
  title: 'Title',
  description: '_ExtendedDescription',
  contentRefreshDate: 'ContentRefreshDate',
  views: 'Views',
  likes: 'Likes',
  comments: 'Comments',
  downloads: 'Downloads',
  follow: 'Follow',
  share: 'Share',
  bookmark: 'Bookmark',
  sensitiveTerms: 'SensitiveTerms',
  reviewerComments: 'ReviewerComments',
  projectId: 'ProjectId',
  versionFileName: 'VersionFileName',
  versionFileType: 'VersionFileType',
  modifiedBy: 'Editor',
  editedBy: 'Editor0',
  docIcon: 'DocIcon',
};

// Term Store GUIDs
export const TERM_SET_IDS = {
  buDepartment: '2cc9b6f4-ba29-438c-9f4a-877418912b85',
  client: 'f205a202-e08c-4424-b39f-34e1ccddc708',
  diseaseArea: '98942a12-8bd1-450c-8f64-e7e0ec55e75c',
  documentType: 'f638dc73-d5d6-49ee-8a8f-59ff6b59ae66',
  geography: '3a16d60d-5af6-44aa-a47d-0834e971a12e',
  therapyArea: 'f213c083-a349-4be0-8db1-a6544f41b4c8',
};

// Search Managed Properties
export const SEARCH_PROPERTIES = {
  status: 'RefinableString00',
  documentType: 'RefinableString03',  // ← clean label e.g. "Case Study"
  published: 'RefinableDate00',        // ← date sort
  businessUnit: 'RefinableString01',
  department: 'RefinableString02',
  client: 'RefinableString04',
  geography: 'RefinableString05',
  diseaseArea: 'RefinableString06',
  therapyArea: 'RefinableString07',
  listItemId: 'ListItemID',
  path: 'Path',
  title: 'Title',
  write: 'Write',
  rank: 'Rank',
};

// Role / Group GUIDs
export const GROUP_IDS = {
  admins: 'df3deb63-a6d5-42f8-9fda-0dfddde87fac',
  approvers: 'df3deb63-a6d5-42f8-9fda-0dfddde87fac',
  contributors: 'df3deb63-a6d5-42f8-9fda-0dfddde87fac',
  learners: 'df3deb63-a6d5-42f8-9fda-0dfddde87fac',
  kmsUsers: 'df3deb63-a6d5-42f8-9fda-0dfddde87fac',
  //contributors: 'd62780b8-47bb-4121-bb8e-90b9b3225832',
  //learners: '53d337d6-9b9b-4818-98ae-8cbf0cd92a0f',
};

// Role / Group Names (Azure AD display names)
export const GROUP_NAMES = {
  admins: 'InSkytes Owners',
  approvers: 'InSkytes Owners',
  contributors: 'InSkytes Owners',
  learners: 'InSkytes Owners',
  kmsUsers: 'InSkytes Owners',
  //contributors: 'KM-Approver',
  //learners: 'KM-Learner',
};

// List GUIDs
export const LIST_GUIDS = {
  kmDataHub: '6c9e04e6-f646-4f5b-a610-94e2eeb07998',
  kmHarvestHub: '90f2a0f9-85db-4a16-94dc-5bd94bd34cee',
  documentComments: '188171e3-5da5-44fe-aaf7-5c673b1084d8',
  documentMetrics: '3b829fad-27a2-4e09-ac9c-4009096a1b4f',
  userInteractions: '73d35d04-7367-4d88-8a7a-b578bf1a573f',
  auditLog: 'cb2e6f43-7e54-4e08-b149-bf97d21fbf5d',
  documentReviewFlags: 'f5e64a1b-5e88-41be-82be-309934c0c31b',
  homePageConfig: '41c93959-14c3-4f95-9126-1f6ac39df257',
  notificationQueue: '',
  syncConfig: '',
};

// Site Pages Library (for Share dialog)
export const SITE_PAGES = {
  sitePagesLibraryGuid: '3be0eb5c-4b89-4856-b7b1-49668191e3ce',
  assetPageId: 5,
  assetPageName: 'Assets',
};

// sessionStorage Cache Keys
export const CACHE_KEYS = {
  fieldMap: 'kmFieldMap_v1',
  taxonomy: 'kmTaxonomy_v1',
  currentUser: 'kmCurrentUser_v1',
  fields: 'kmFields_v1',
  buTerms: 'kmBUTerms_v1',
  metricsKnowledgeHub: 'kmMetrics_knowledgeHub_v1',
  metricsRecentlyPublished: 'kmMetrics_recentlyPublished_v1',
  metricsSynced: 'kmMetricsSynced_v1',
  kmsUsers: 'kmKmsUsers_v1',
  kmsUsersGroupId: 'kmUploadForm:kmsUsersGroupId:v1',
  uploadSuccessRestore: 'ikn:upload-success-restore',
  sharedSearchHistory: 'kms-shared-search-history',
  auditSync: 'kmAuditSynced_v1',
  homePageConfig: 'kmHomePageConfig_v1',
};

// Pagination / Batch Sizes
export const PAGE_SIZES = {
  documentsPerPage: 30,
  searchPageSize: 50,
  searchResultsPerPage: 30,
  maxListFetch: 5000,
  metricsBatchSize: 10,
  updateBatchSize: 20,
};

// TTL Values (milliseconds)
export const TTL_MS = {
  metricsCache: 5 * 60 * 1000,
  searchCache: 60 * 60 * 1000,
  whosWhoCache: 10 * 60 * 1000,
  auditSyncInterval: 5 * 60 * 1000,
};

// Audit Log Action Types
export const AUDIT_ACTIONS = {
  underReview: 'Under Review',
  active: 'Active',
  archive: 'Archive',
  reject: 'Reject',
};

// SharePoint List Item Types
export const LIST_ITEM_TYPES = {
  auditLog: 'SP.Data.Audit_x0020_LogListItem',
};
