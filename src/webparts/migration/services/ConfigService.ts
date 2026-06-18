// src/webparts/migration/services/ConfigService.ts

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LIST_NAMES } from '../config/appConfig';

export interface IFaqItem {
    Title: string;
    Answer: string;
    Category?: string;
}

export interface IQuickLinkItem {
    Title: string;
    Url: string;
    Description?: string;
}

export interface IHomePageConfig {
    id?: number;
    heroMessage: string;
    quoteAuthor: string;
    imageUrl: string;
}

export class ConfigService {
    private context: WebPartContext;
    private cache: {
        faqs?: IFaqItem[];
        quickLinks?: IQuickLinkItem[];
        lastFetched?: number;
    } = {};

    // Cache expiry: 10 minutes (Improved for production freshness)
    private readonly CACHE_EXPIRY_MS = 10 * 60 * 1000;

    constructor(context: WebPartContext) {
        this.context = context;
    }

    /**
     * Checks if the cache is still valid.
     */
    private isCacheValid(): boolean {
        if (!this.cache.lastFetched) return false;
        return (Date.now() - this.cache.lastFetched) < this.CACHE_EXPIRY_MS;
    }

    /**
     * Clears cache to force fresh fetch.
     */
    public forceRefresh(): void {
        this.cache = {};
        console.log("ConfigService: Cache cleared manually.");
    }

    /**
     * Ensures the FAQs list exists in SharePoint.
     */
    public async ensureFaqsListExists(): Promise<boolean> {
        try {
            const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.faqs}')?$select=Id`;
            const response = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            
            if (response.ok) return true;

            // 🚩 ROBUST PROVISIONING: Use simple metadata and handle 400 errors gracefully
            const createResponse = await this.context.spHttpClient.post(
                `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-Type': 'application/json;odata=nometadata'
                    },
                    body: JSON.stringify({
                        'AllowContentTypes': true,
                        'BaseTemplate': 100,
                        'ContentTypesEnabled': true,
                        'Description': 'Repository for Chatbot FAQs',
                        'Title': LIST_NAMES.faqs
                    })
                }
            );

            if (!createResponse.ok) {
                const errText = await createResponse.text();
                console.error("Failed to create FAQs list:", errText);
                return false;
            }

            console.log("ConfigService: FAQs list provisioned. Adding fields...");
            
            const fields = [
                { title: 'Answer', type: 3 }, // Note
                { title: 'Category', type: 2 } // Text
            ];

            for (const field of fields) {
                await this.context.spHttpClient.post(
                    `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.faqs}')/fields`,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-Type': 'application/json;odata=nometadata'
                        },
                        body: JSON.stringify({
                            'Title': field.title,
                            'FieldTypeKind': field.type,
                            'Required': false
                        })
                    }
                );
            }

            console.log("ConfigService: FAQs list fields added.");
            return true;
        } catch (error) {
            console.error("Error provisioning FAQs list:", error);
            return false;
        }
    }

    /**
     * Fetches FAQs from a SharePoint list.
     */
    public async fetchFAQs(): Promise<IFaqItem[]> {
        if (this.isCacheValid() && this.cache.faqs) {
            return this.cache.faqs;
        }

        try {
            // Attempt to ensure list exists first
            await this.ensureFaqsListExists();

            // 🚩 RESILIENCE: Remove $select to avoid 400 errors if field names differ
            const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.faqs}')/items`;
            const response: SPHttpClientResponse = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            
            if (!response.ok) {
                if (response.status === 404) {
                    console.warn("FAQs list still not found after attempt (404).");
                } else {
                    console.warn(`Failed to fetch FAQs list. Status: ${response.status}`);
                }
                return [];
            }

            const data = await response.json();
            this.cache.faqs = data.value || [];
            this.cache.lastFetched = Date.now();
            return this.cache.faqs || [];
        } catch (error) {
            console.error("Error fetching FAQs:", error);
            return [];
        }
    }

    /**
     * Fetches QuickLinks from a SharePoint list.
     */
    public async fetchQuickLinks(): Promise<IQuickLinkItem[]> {
        if (this.isCacheValid() && this.cache.quickLinks) {
            return this.cache.quickLinks;
        }

        try {
            const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.quickLinks}')/items?$select=Title,Url,Description`;
            const response: SPHttpClientResponse = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            
            if (!response.ok) {
                console.warn("Failed to fetch QuickLinks list.");
                return [];
            }

            const data = await response.json();
            this.cache.quickLinks = data.value || [];
            this.cache.lastFetched = Date.now();
            return this.cache.quickLinks || [];
        } catch (error) {
            console.error("Error fetching QuickLinks:", error);
            return [];
        }
    }

    public async fetchHomePageConfig(): Promise<IHomePageConfig> {
        const DEFAULT_CONFIG: IHomePageConfig = {
            heroMessage: 'The strength of the team is each individual member. The strength of each member is the team. Alone we can do so little; together we can do so much.',
            quoteAuthor: 'Helen Keller',
            imageUrl: ''
        };

        try {
            const listUrl =
                `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.homePageConfig}')/items` +
                `?$select=Id,HeroMessage,QuoteAuthor,ImageUrl&$filter=Title eq 'heroConfig'&$top=1`;
            const response = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            if (!response.ok) return DEFAULT_CONFIG;
            const data = await response.json();
            const item = (data.value || [])[0];
            if (!item) return DEFAULT_CONFIG;
            return {
                id: item.Id,
                heroMessage: item.HeroMessage || DEFAULT_CONFIG.heroMessage,
                quoteAuthor: item.QuoteAuthor || DEFAULT_CONFIG.quoteAuthor,
                imageUrl: item.ImageUrl || ''
            };
        } catch {
            return DEFAULT_CONFIG;
        }
    }

    public async saveHomePageConfig(
        config: IHomePageConfig
    ): Promise<boolean> {
        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const itemId = config.id;
            if (!itemId) return false;

            const url =
                `${webUrl}/_api/web/lists/getbytitle('${LIST_NAMES.homePageConfig}')/items(${itemId})`;

            const response = await this.context.spHttpClient.post(url,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        Accept: 'application/json;odata=nometadata',
                        'Content-Type': 'application/json;odata=nometadata',
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE'
                    },
                    body: JSON.stringify({
                        HeroMessage: config.heroMessage,
                        QuoteAuthor: config.quoteAuthor,
                        ImageUrl: config.imageUrl
                    })
                }
            );
            if (!response.ok && response.status !== 204) {
                const errText = await response.text().catch(() => '');
                console.error('saveHomePageConfig failed:', response.status, errText);
                return false;
            }
            return response.ok || response.status === 204;
        } catch {
            return false;
        }
    }

    private async createOptimizedHeroImage(file: File): Promise<Blob> {
        if (!file.type || file.type.indexOf('image/') !== 0) {
            return file;
        }

        const imageUrl = URL.createObjectURL(file);

        try {
            const image = await new Promise<HTMLImageElement>((resolve, reject) => {
                const img = new Image();
                img.onload = () => resolve(img);
                img.onerror = () => reject(new Error('Unable to load hero image'));
                img.src = imageUrl;
            });

            const outputSize = 768;
            const canvas = document.createElement('canvas');
            canvas.width = outputSize;
            canvas.height = outputSize;

            const context = canvas.getContext('2d');
            if (!context) {
                return file;
            }

            context.imageSmoothingEnabled = true;
            context.imageSmoothingQuality = 'high';

            context.fillStyle = '#ffffff';
            context.fillRect(0, 0, outputSize, outputSize);

            const scale = Math.min(
                outputSize / image.naturalWidth,
                outputSize / image.naturalHeight
            );
            const drawWidth = image.naturalWidth * scale;
            const drawHeight = image.naturalHeight * scale;
            const drawX = (outputSize - drawWidth) / 2;
            const drawY = (outputSize - drawHeight) / 2;

            context.drawImage(
                image,
                drawX,
                drawY,
                drawWidth,
                drawHeight
            );

            return await new Promise<Blob>((resolve) => {
                canvas.toBlob(
                    blob => resolve(blob || file),
                    file.type === 'image/png' ? 'image/png' : 'image/jpeg',
                    0.92
                );
            });
        } catch {
            return file;
        } finally {
            URL.revokeObjectURL(imageUrl);
        }
    }

    public async uploadHeroImage(file: File): Promise<string> {
        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const optimizedImage = await this.createOptimizedHeroImage(file);
            const fileName = `hero-image-${Date.now()}.${file.type === 'image/png' ? 'png' : 'jpg'}`;
            const siteRelativeUrl = this.context.pageContext.site.serverRelativeUrl;
            const uploadUrl =
                `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${siteRelativeUrl}/SiteAssets')/Files/add(url='${fileName}',overwrite=true)`;

            const response = await this.context.spHttpClient.post(
                uploadUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: { Accept: 'application/json;odata=nometadata' },
                    body: optimizedImage
                }
            );
            if (!response.ok) return '';
            await response.json().catch(() => ({}));
            return `${webUrl}/SiteAssets/${fileName}`;
        } catch {
            return '';
        }
    }

    public findFaqMatch(query: string, faqs: IFaqItem[]): IFaqItem | null {
        if (!query || faqs.length === 0) return null;
        
        const q = query.toLowerCase().trim();
        
        // 🚩 BYPASS: If asking about documents, lists, or people specifically, do not use FAQ overrides
        const docKeywords = ['document', 'list', 'bcp', 'active', 'show', 'tell me about', 'file', 'who', 'whom', 'whose', 'personnel', 'team', 'contact', 'employee'];
        if (docKeywords.some(k => q.includes(k))) return null;

        const queryTokens = q.split(/\s+/).filter(t => t.length > 2);
        
        // 1. Try Exact Match (Highest Priority)
        const exact = faqs.find(f => f.Title.toLowerCase().trim() === q);
        if (exact) return exact;

        // 2. Try High-Confidence Keyword Match
        let bestMatch: IFaqItem | null = null;
        let maxOverlap = 0;

        for (const faq of faqs) {
            const title = faq.Title.toLowerCase();
            let overlap = 0;
            
            for (const token of queryTokens) {
                if (title.includes(token)) overlap++;
            }

            // 🚩 REQUIRE HIGHER CONFIDENCE: Match at least 80% of query tokens
            const matchConfidence = overlap / queryTokens.length;
            if (overlap > maxOverlap && matchConfidence >= 0.8) {
                maxOverlap = overlap;
                bestMatch = faq;
            }
        }

        return bestMatch;
    }
}
