// src/webparts/migration/models/SearchResult.ts

export interface ISearchResult {
    id: string;
    title: string;
    content: string;
    url?: string;
    score?: number;
    // Reranked or semantic score
    semanticScore?: number;
    rank?: number;
    // Metadata fields
    author?: string;
    publishedDate?: string;
    description?: string;
    fileType?: string;
    status?: string;
    // Enhanced Metadata
    likesCount?: number;
    commentsCount?: number;
    viewsCount?: number;
    downloadsCount?: number;
    businessUnit?: string;
    documentType?: string;
    department?: string;
    geography?: string;
    therapyArea?: string;
    diseaseArea?: string;
    client?: string;
    abstract?: string;
    listTitle?: string;
    listId?: string;
    parentDocumentId?: string;
    chunkIndex?: number;
    fileName?: string;
    contentRefreshDate?: string;
    downloads?: number;
    follow?: boolean;
    share?: boolean;
    bookmark?: boolean;
    sensitiveTerms?: string[];
    reviewerComments?: string;
    projectId?: string;
    versionFileName?: string;
    versionFileType?: string;
    modifiedBy?: {
        title: string;
        email: string;
        id: number;
    };
    docIcon?: string;
    searchRankDebug?: {
        titleMatchCount: number;
        metadataMatchCount: number;
        abstractMatchCount: number;
        contentMatchCount: number;
        exactPhraseMatch: boolean;
        fullSentenceMatch: boolean;
        notTermExcluded: boolean;
        semanticScore: number;
        finalRankScore: number;
    };
}
