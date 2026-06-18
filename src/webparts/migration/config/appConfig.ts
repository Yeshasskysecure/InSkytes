/**
 * Central application configuration.
 * Change values here to deploy to a new tenant, site, or library.
 * Do NOT hardcode these values anywhere else in the codebase.
 */

// Site
export const SITE_URL = 'https://indegene123.sharepoint.com/sites/iKnowledgeNext';
export const SITE_RELATIVE_URL = '/sites/iKnowledgeNext';

// Site Pages
export const PAGE_URLS = {
  home: '/sites/iKnowledgeNext/SitePages/Home.aspx',
  businessUnits: '/sites/iKnowledgeNext/SitePages/Business-Units.aspx',
  categories: '/sites/iKnowledgeNext/SitePages/Categories.aspx',
  search: '/sites/iKnowledgeNext/SitePages/Search.aspx',
  assets: '/sites/iKnowledgeNext/SitePages/Assets.aspx',
  bookmark: '/sites/iKnowledgeNext/SitePages/Bookmark.aspx',
  myDocuments: '/sites/iKnowledgeNext/SitePages/MyDocuments.aspx',
  analytics: '/sites/iKnowledgeNext/SitePages/Analytics.aspx',
  auditLog: '/sites/iKnowledgeNext/SitePages/Audit-Log.aspx',
  kmReviewHub: '/sites/iKnowledgeNext/SitePages/KM-Review-Hub.aspx',
  kmHarvestHub: '/sites/iKnowledgeNext/SitePages/KM-Harvest-Hub.aspx',
  kmLibrary: '/sites/iKnowledgeNext/SitePages/KM-Library.aspx',
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
  kmLibrary: 'https://indegene123.sharepoint.com/sites/iKnowledgeNext/KM%20Review%20Hub/Forms/AllItems.aspx',

  // Power BI Analytics
  powerBI: {
    baseEmbedUrl: 'https://app.powerbi.com/reportEmbed',
    reportId: 'ef1d7266-0188-4440-8176-505d522dc687',
    groupId: '0dbbc058-a22e-4b5e-88bc-c8b7018241c6',
    autoAuth: 'true',
    ctid: '6d787ab7-f295-424f-b8cc-e3116a0f8520',
    pages: [
      { name: 'Platform Adoption', pageName: '2e0979c73e51ee01c462', sectionName: 'section-1' },
      { name: 'Platform Adoption 2', pageName: 'b3eb8e010386ff7390a1', sectionName: 'section-1' },
      { name: 'Search Effectiveness', pageName: 'b2ca777165d080a93cd5', sectionName: 'section-2' },
      { name: 'Content Health', pageName: 'a7b74412800b96542abd', sectionName: 'section-3' },
    ],
  },
};

export const KM_REVIEW_HUB_DRIVE_ID = 'b!olUm0jdVrEG6BCl3djVAYSeFuWe55TRKmdC5U7PAAYq1CBr4lunfSKVgJmSl6XPV';

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
  kmHarvestHub: '/sites/iKnowledgeNext/Lists/KM%20Harvest%20Hub/AllItems.aspx',
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
  description: 'Description',
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
  buDepartment: '03d64d66-372c-4826-8992-8d9a23148ee7',
  client: 'fa42f63d-139b-48e1-9243-54e8fd1318c3',
  diseaseArea: '23b86b53-691d-48a9-a017-c2404312d7d9',
  documentType: 'e98e96b5-a9f8-4932-a244-ba17d8211a38',
  geography: '944bbb14-a0e1-4155-8e56-15b0babe8747',
  therapyArea: 'd28756c2-1eec-4220-b737-15a3c8e78d64',
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
  admins: '7dc962e9-b8cf-49f6-ae35-0db943c8beef',
  approvers: '7d821944-49be-4ca9-ba89-a474f20de039',
  contributors: 'db81f939-f757-44ed-a4f3-fc1e6467975d',
  learners: '14894a06-471b-4756-9ab6-2dfc05ec453e',
  kmsUsers: 'd18283a5-6370-414a-a066-265f986bb73d',
  //contributors: 'd62780b8-47bb-4121-bb8e-90b9b3225832',
  //learners: '53d337d6-9b9b-4818-98ae-8cbf0cd92a0f',
};

// Role / Group Names (Azure AD display names)
export const GROUP_NAMES = {
  admins: 'SG_KM-Admin',
  approvers: 'SG_KM-Approver',
  contributors: 'SG_ KM-Contributor',
  learners: 'SG_ KM-Learner',
  kmsUsers: 'SG_ Indegene-All-Employees',
  //contributors: 'KM-Approver',
  //learners: 'KM-Learner',
};

// List GUIDs
export const LIST_GUIDS = {
  kmDataHub: 'f81a08b5-e996-48df-a560-2664a5e973d5',
  documentComments: 'ab0480d7-1a5c-469a-94d4-e182e9e53155',
  documentMetrics: 'e153d995-8f4b-4266-85ce-8727bd88e504',
  userInteractions: '050f7f27-9b11-41df-86ea-13a4b4d3e0d6',
  auditLog: '4ff06037-fce1-4a94-ab05-3c113d005a7a',
  notificationQueue: '973cd570-50dc-416f-b059-a0256e21e2c1',
  syncConfig: '2b199ed8-d13c-434a-a7a9-57edc3f56682',
};

// Site Pages Library (for Share dialog)
export const SITE_PAGES = {
  sitePagesLibraryGuid: '8d812734-facd-4a16-9d6f-ecc560e8b59d',
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
