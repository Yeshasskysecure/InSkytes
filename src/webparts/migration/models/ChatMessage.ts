// src/webparts/migration/models/ChatMessage.ts

import { ISearchResult } from './SearchResult';

export interface IChatMessage {
    id?: string;
    sender: 'user' | 'bot';
    text: string;
    timestamp: Date;
    // Citations are the source documents used for the RAG answer
    citations?: ISearchResult[];
    // For UI: title/url if it's a simple redirection
    title?: string;
    url?: string;
}
