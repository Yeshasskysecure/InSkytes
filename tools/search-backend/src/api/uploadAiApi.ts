import { requireValue, SearchEnv, toNumber } from '../runtime/env';
import { getChatCompletionResult } from '../runtime/openAiClient';

export interface UploadAiContext {
  correlationId?: string;
  userHash?: string;
  log?: (event: Record<string, unknown>) => void;
}

export interface UploadedMediaFile {
  buffer: Buffer;
  filename: string;
  contentType?: string;
}

interface MetadataExtraction {
  title?: string;
  documentType?: string;
  bu?: string;
  department?: string;
  subDepartment?: string;
  geography?: string;
  region?: string;
  client?: string;
  description?: string;
  abstract?: string;
  diseaseArea?: string;
  therapyArea?: string;
  sensitiveTerms?: string;
  emails?: string;
  phones?: string;
  ids?: string;
  pricing?: string;
}

interface SensitiveInfoCategories {
  clientOrganizations?: string[];
  clientBranding?: string[];
  clientPersonnelDetails?: string[];
  commercialInformation?: string[];
  regulatoryFinancialInformation?: string[];
}

interface MediaAnalysis {
  abstract: string;
  businessUnit: string;
  department: string;
  documentType: string;
  client: string;
  geography: string;
  therapyArea: string;
  diseaseArea: string;
  title: string;
  sensitiveTerms?: string;
}

const MISSING_VALUES = new Set([
  '',
  'n/a',
  'na',
  'not applicable',
  'not available',
  'none',
  'null',
  'undefined',
  'unknown'
]);

const SPECIAL_CLIENT_VALUES = new Set([
  'not applicable',
  'others',
  'multi-client'
]);

const truncate = (value: string, maxLength: number): string =>
  value.length > maxLength ? value.slice(0, maxLength) : value;

const TITLE_MAX_CHARS = 100;
const DESCRIPTION_MAX_CHARS = 500;

const normalizeString = (value: unknown): string =>
  String(value || '').replace(/\s+/g, ' ').trim();

const cleanGeneratedText = (value: unknown, maxLength?: number): string => {
  let cleaned = normalizeString(value)
    .replace(/[$%@*+?!#&^~`|\\<>]/g, '')
    .replace(/\.(pdf|docx|pptx|xlsx|mp4|mp3|wav|mov|avi|mkv|webm|m4a|aac|flac|ogg)$/gi, '')
    .replace(/[\s_-]?v\d+(\.\d+)*[\s_-]?/gi, ' ')
    .replace(/[\s_-]?(final|draft|copy|revised|updated|new|old|backup|temp)[\s_-]?/gi, ' ')
    .replace(/[\s_-]?\d{4}[-*]?\d{2}[-*]?\d{2}[\s_-]?/gi, ' ')
    .replace(/[\s_-]#?\d{2,6}[\s_-]?/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (maxLength && cleaned.length > maxLength) {
    cleaned = cleaned.slice(0, maxLength).trim();
  }

  return cleaned;
};

const truncateAtWord = (value: string, maxLength: number): string => {
  const text = normalizeString(value);
  if (text.length <= maxLength) return text;

  const truncated = text.slice(0, maxLength).trimEnd();
  const lastSpace = truncated.lastIndexOf(' ');
  return (lastSpace >= Math.floor(maxLength * 0.65)
    ? truncated.slice(0, lastSpace)
    : truncated).trim();
};

const cleanGeneratedTitle = (value: unknown): string => {
  const cleaned = normalizeString(value)
    .replace(/[–—−]/g, '-')
    .replace(/\.(pdf|docx|pptx|xlsx|mp4|mp3|wav|mov|avi|mkv|webm|m4a|aac|flac|ogg)$/gi, '')
    .replace(/[\s_-]?v\d+(\.\d+)*[\s_-]?/gi, ' ')
    .replace(/[\s_-]?(final|draft|copy|revised|updated|new|old|backup|temp)[\s_-]?/gi, ' ')
    .replace(/[_/\\:;,.()[\]{}'"!?%@#*&+=$^~`|<>]/g, ' ')
    .replace(/[^A-Za-z0-9 -]/g, ' ')
    .replace(/-{2,}/g, '-')
    .replace(/\s+-\s+/g, ' - ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  return truncateAtWord(cleaned, TITLE_MAX_CHARS);
};

const cleanGeneratedDescription = (value: unknown): string =>
  truncateAtWord(
    cleanGeneratedText(value)
      .replace(/^\s*(this document contains information about|this presentation discusses|please refer to|the purpose of this document is)\s*/i, '')
      .replace(/\s{2,}/g, ' ')
      .trim(),
    DESCRIPTION_MAX_CHARS
  );

const cleanSensitiveText = (value: unknown, maxLength?: number): string => {
  let cleaned = normalizeString(value)
    .replace(/[<>`|\\]/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (maxLength && cleaned.length > maxLength) {
    cleaned = cleaned.slice(0, maxLength).trim();
  }

  return cleaned;
};

const normalizeOptionalClassification = (value: unknown): string => {
  const text = cleanGeneratedText(value);
  return MISSING_VALUES.has(text.toLowerCase()) ? 'Not Applicable' : text;
};

const normalizeClientValue = (value: unknown): string => {
  const text = cleanGeneratedText(value);
  if (!text || MISSING_VALUES.has(text.toLowerCase())) return 'Not Applicable';

  const values = text
    .split(/[,;\n]/)
    .map((item) => cleanGeneratedText(item))
    .filter(Boolean);

  const concrete = values.filter((item) => !SPECIAL_CLIENT_VALUES.has(item.toLowerCase()));
  const unique = Array.from(new Set(concrete.map((item) => item.trim()))).filter(Boolean);

  if (unique.length > 1) return 'Multi-Client';
  if (unique.length === 1) return unique[0];
  return values.some((item) => item.toLowerCase() === 'others') ? 'Others' : 'Not Applicable';
};

const parseJsonFromModel = (content: string): Record<string, unknown> => {
  const fenced = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  const jsonText = fenced ? fenced[1] : content;
  return JSON.parse(jsonText.trim()) as Record<string, unknown>;
};

const toStringList = (value: unknown): string[] => {
  if (Array.isArray(value)) {
    return value
      .reduce((items: string[], item) => items.concat(toStringList(item)), [])
      .map((item) => item.trim())
      .filter(Boolean);
  }

  return String(value || '')
    .split(/[\n;]+|,(?=\s*[A-Za-z@$€£₹])|(?<=[A-Za-z])\s*,\s*(?=\d)/)
    .map((item) => cleanSensitiveText(item))
    .filter((item) => !!item && !MISSING_VALUES.has(item.toLowerCase()));
};

const dedupeValues = (values: string[]): string[] => {
  const seen = new Set<string>();
  const deduped: string[] = [];

  values.forEach((value) => {
    const cleaned = cleanSensitiveText(value);
    const key = cleaned.toLowerCase().replace(/[$€£₹]/g, '').replace(/\s+/g, ' ').trim();
    if (!cleaned || seen.has(key)) return;
    seen.add(key);
    deduped.push(cleaned);
  });

  return deduped;
};

const extractSensitiveInfoCategories = (metadata: Record<string, unknown>): SensitiveInfoCategories => {
  const source = (metadata.sensitiveInfo || metadata.sensitiveInformation || metadata.sensitiveData || {}) as Record<string, unknown>;
  return {
    clientOrganizations: dedupeValues(toStringList(source.clientOrganizations)),
    clientBranding: dedupeValues(toStringList(source.clientBranding)),
    clientPersonnelDetails: dedupeValues(toStringList(source.clientPersonnelDetails)),
    commercialInformation: dedupeValues(toStringList(source.commercialInformation)),
    regulatoryFinancialInformation: dedupeValues(toStringList(source.regulatoryFinancialInformation))
  };
};

const buildSensitiveTermsValue = (metadata: Record<string, unknown>): string => {
  const sensitiveInfo = extractSensitiveInfoCategories(metadata);
  const categoryValues: string[] = [
    sensitiveInfo.clientOrganizations || [],
    sensitiveInfo.clientBranding || [],
    sensitiveInfo.clientPersonnelDetails || [],
    sensitiveInfo.commercialInformation || [],
    sensitiveInfo.regulatoryFinancialInformation || []
  ].flat();

  return dedupeValues([
    ...categoryValues,
    ...toStringList(metadata.sensitiveTerms),
    ...toStringList(metadata.emails),
    ...toStringList(metadata.phones),
    ...toStringList(metadata.ids),
    ...toStringList(metadata.pricing)
  ]).join(', ');
};

const buildSensitiveInformationRules = (clientTermsSection = ''): string => {
  const clientDictionaryInstruction = clientTermsSection
    ? `Known client dictionary from the live iKnowledgeNext client taxonomy. Treat exact matches as strong client evidence, but still apply the ownership/confidence rules below:
${clientTermsSection}`
    : `No client dictionary was provided. Use only clear evidence in the content.`;

  return `Sensitive information detection rules:
- Review the full supplied text, including body text, tables, lists, headers/footers, captions, hyperlinks, embedded object text, comments, tracked-change text, filenames/attachment names if present in text, watermarks, and branding text that appears in the extracted content.
- Report only CLIENT-SENSITIVE or THIRD-PARTY SENSITIVE information.
- Internal organization exclusions: Indegene, Indegene Limited, Indegene Inc., Indegene Lifesystems Private Limited, iKnowledgeNext, iKnowledgeNEXT, Knowledge Hub, Addressable Health, BioPharm Communications, Cult Health, DT Associates, MJL Advertising, Trilogy, Trilogy Writing & Consulting, Trilogy Writing Consulting, Warn and Co, @indegene.com.
- Exclude internal business taxonomy/medical classification terms when they are not external organizations, especially "Disease area" and "Therapy area".
- A client organization is an external organization not in the internal exclusion list.
- Resolve abbreviations, tickers, and short forms only when the context makes the external organization unambiguous, for example Pfizer/PFE, Roche, AstraZeneca/AZ. Unknown or ambiguous names must not be reported.
- Before reporting any value, classify ownership as Internal, External, or Unknown/Ambiguous. Report only External values with Medium or High confidence.
- Client personnel details: report a person only if the same section or within three nearby lines contains an external client email, phone, title, or confirmed external organization association.
- Commercial information: report monetary values, budgets, rates, deal values, pricing, forecasts, or contract terms only when associated with a named external organization or clear external proposal/contract/commercial context.
- Regulatory and financial information: report only when it belongs to, references, or is submitted for a named external organization or third-party product/program.
- Do not mask values. Do not add explanations, scores, or reasoning.
- Dedupe exact and near-duplicate values.

${clientDictionaryInstruction}

Return sensitiveInfo as category arrays:
- clientOrganizations
- clientBranding
- clientPersonnelDetails
- commercialInformation
- regulatoryFinancialInformation

Also return sensitiveTerms as one comma-separated union of those sensitiveInfo values plus true external emails, phones, IDs, and pricing values.`;
};

const buildDocumentMetadataPrompt = (documentText: string, docTypeTermsSection = '', clientTermsSection = ''): string => {
  const docTypeInstruction = docTypeTermsSection
    ? `documentType - MANDATORY - Choose the single BEST matching document type from this exact list. Return only the exact term name.

Available document types:
${docTypeTermsSection}`
    : `documentType - MANDATORY - Identify the best matching document type from the content.`;

  return `Extract Knowledge Hub metadata from the document text.

Rules:
- Use only information present in the document text.
- Do not invent clients, business units, departments, therapy areas, disease areas, IDs, emails, or pricing.
- For fields not found, use "Not Applicable".
- Title must be based on the complete understanding of the document. Generate a concise meaningful title with a maximum of ${TITLE_MAX_CHARS} characters.
- Title must clearly describe the primary topic, purpose, and business context. Include important searchable keywords where natural.
- Use professional enterprise language and title case where appropriate.
- Do not use generic titles such as Presentation, Document, Final Version, Updated Deck, Revised, or file names.
- Do not include version numbers, dates, author names, confidentiality markings, document IDs, or unnecessary abbreviations unless critical to understanding the content.
- If available, prioritize the business process, capability, therapy area, permitted client context, business objective, or key outcome over document type.
- Title characters allowed: letters, numbers, spaces, and hyphens only. Do not use punctuation or special characters.
- Description must be a professional one paragraph knowledge asset summary with a maximum of ${DESCRIPTION_MAX_CHARS} characters.
- Description must explain the purpose, primary topics, processes, methodologies, insights, recommendations, outcomes, intended audience or business functions where identifiable, and business value or operational impact.
- Description must be standalone, business-friendly, complete sentences, and include relevant searchable business, industry, service, therapy area, and disease area terms where supported by the content.
- Do not repeat or closely paraphrase the title in the description.
- Do not include file names, version numbers, revision history, dates, document IDs, author names, approval information, metadata, confidentiality statements, legal disclaimers, headers, footers, page numbers, references, navigation text, bullets, numbering, markdown, or vague boilerplate such as "This document contains information about", "This presentation discusses", "Please refer to", or "The purpose of this document is".
- If one specific client is present, return that client.
- If multiple specific clients are present, return "Multi-Client".
- If client exists but is not identifiable, return "Others".
- If no client is present, return "Not Applicable".
- geography should be one geography if clear, "General" if multiple, or "Not Applicable" if absent.
- therapyArea should be one therapy area if clear, "Cross-Therapy" if multiple, or "Not Applicable" if absent.
- diseaseArea should be one disease area if clear, "Cross-Disease" if multiple, or "Not Applicable" if absent.

Fields:
- title
- documentType
- bu
- department
- subDepartment
- geography
- client
- description
- diseaseArea
- therapyArea
- sensitiveTerms
- emails
- phones
- ids
- pricing

${docTypeInstruction}

${buildSensitiveInformationRules(clientTermsSection)}

Document text:
${documentText}

Return only valid JSON in this exact shape:
{
  "title": "",
  "documentType": "",
  "bu": "",
  "department": "",
  "subDepartment": "",
  "geography": "",
  "client": "",
  "description": "",
  "diseaseArea": "",
  "therapyArea": "",
  "sensitiveTerms": "",
  "emails": "",
  "phones": "",
  "ids": "",
  "pricing": "",
  "sensitiveInfo": {
    "clientOrganizations": [],
    "clientBranding": [],
    "clientPersonnelDetails": [],
    "commercialInformation": [],
    "regulatoryFinancialInformation": []
  }
}`;
};

const buildMediaAnalysisPrompt = (fileName: string, transcriptText: string, docTypeTermsSection = '', clientTermsSection = ''): string => {
  const docTypeInstruction = docTypeTermsSection
    ? `documentType: Choose the single BEST matching document type from this list based on the transcript content. Return only the exact term name.

Available document types:
${docTypeTermsSection}`
    : `documentType: Identify the media document type. Use labels like "Training", "Webinar", "Podcast", "Demo", "Presentation", or "Interview".`;

  return `Analyze the following transcript from a media file.

Rules:
- Only extract information actually present in the transcript.
- Do not hallucinate or invent information not in the transcript.
- For fields not found, use "Not Applicable".
- Remove special characters $ % @ * + ? ! # & ^ ~ | < > from all generated text.
- Do not include file extensions, version numbers, dates, IDs, author names, confidentiality markings, or unnecessary abbreviations in title.
- Do not mention "transcript", "audio", "video", "recording", "speaker", "the speaker says", "this video", or "this audio" in the abstract.
- Title must be based on the full transcript content, not just the file name. Generate a concise meaningful title with a maximum of ${TITLE_MAX_CHARS} characters.
- Title must clearly describe the primary topic, purpose, and business context. Use professional enterprise language and title case where appropriate.
- Title characters allowed: letters, numbers, spaces, and hyphens only. Do not use punctuation or special characters.
- Abstract must be a professional one paragraph knowledge asset summary with a maximum of ${DESCRIPTION_MAX_CHARS} characters.
- Abstract must explain the purpose, primary topics, processes, methodologies, insights, recommendations, outcomes, intended audience or business functions where identifiable, and business value or operational impact.
- Do not repeat or closely paraphrase the title in the abstract. Do not include file names, metadata, bullets, numbering, markdown, or vague boilerplate.

File name, for fallback context only:
${fileName}

Fields:
- title: clean descriptive title based on the transcript content, max ${TITLE_MAX_CHARS} characters.
- abstract: professional one paragraph summary of the content, max ${DESCRIPTION_MAX_CHARS} characters.
- businessUnit: business unit or organizational area mentioned or implied. Never empty.
- department: department, team, or function mentioned or implied. Never empty.
- ${docTypeInstruction}
- client: one client name, "Multi-Client", "Others", or "Not Applicable".
- geography: one geography, "General", or "Not Applicable".
- therapyArea: one therapy area, "Cross-Therapy", or "Not Applicable".
- diseaseArea: one disease area, "Cross-Disease", or "Not Applicable".
- sensitiveTerms: comma-separated external client-sensitive values detected under the sensitive information rules, or "".

${buildSensitiveInformationRules(clientTermsSection)}

Transcript to analyze:
${transcriptText}

Return only valid JSON in this exact shape:
{
  "title": "",
  "abstract": "",
  "businessUnit": "",
  "department": "",
  "documentType": "",
  "client": "",
  "geography": "",
  "therapyArea": "",
  "diseaseArea": "",
  "sensitiveTerms": "",
  "sensitiveInfo": {
    "clientOrganizations": [],
    "clientBranding": [],
    "clientPersonnelDetails": [],
    "commercialInformation": [],
    "regulatoryFinancialInformation": []
  }
}`;
};

const sanitizeDocumentMetadata = (metadata: Record<string, unknown>): MetadataExtraction => ({
  title: cleanGeneratedTitle(metadata.title),
  documentType: normalizeOptionalClassification(metadata.documentType),
  bu: normalizeOptionalClassification(metadata.bu),
  department: normalizeOptionalClassification(metadata.department),
  subDepartment: normalizeOptionalClassification(metadata.subDepartment),
  geography: normalizeOptionalClassification(metadata.geography || metadata.region),
  region: normalizeOptionalClassification(metadata.region || metadata.geography),
  client: normalizeClientValue(metadata.client),
  description: cleanGeneratedDescription(metadata.description || metadata.abstract),
  abstract: cleanGeneratedDescription(metadata.abstract || metadata.description),
  diseaseArea: normalizeOptionalClassification(metadata.diseaseArea),
  therapyArea: normalizeOptionalClassification(metadata.therapyArea),
  sensitiveTerms: cleanSensitiveText(buildSensitiveTermsValue(metadata)),
  emails: cleanSensitiveText(metadata.emails),
  phones: cleanSensitiveText(metadata.phones),
  ids: cleanSensitiveText(metadata.ids),
  pricing: cleanSensitiveText(metadata.pricing)
});

const sanitizeMediaAnalysis = (metadata: Record<string, unknown>): MediaAnalysis => ({
  title: cleanGeneratedTitle(metadata.title),
  abstract: cleanGeneratedDescription(metadata.abstract),
  businessUnit: normalizeOptionalClassification(metadata.businessUnit || metadata.bu),
  department: normalizeOptionalClassification(metadata.department),
  documentType: normalizeOptionalClassification(metadata.documentType),
  client: normalizeClientValue(metadata.client),
  geography: normalizeOptionalClassification(metadata.geography || metadata.region),
  therapyArea: normalizeOptionalClassification(metadata.therapyArea),
  diseaseArea: normalizeOptionalClassification(metadata.diseaseArea),
  sensitiveTerms: cleanSensitiveText(buildSensitiveTermsValue(metadata))
});

const buildWhisperUrl = (env: SearchEnv): string => {
  const rawEndpoint = String(env.AZURE_OPENAI_WHISPER_ENDPOINT || '').trim();
  const endpoint = (rawEndpoint || requireValue(env, 'AZURE_OPENAI_ENDPOINT')).replace(/\/$/, '');
  if (/\/audio\/(transcriptions|translations)\?/i.test(endpoint)) {
    return endpoint;
  }

  const deployment = env.AZURE_OPENAI_WHISPER_DEPLOYMENT || 'whisper';
  const mode = String(env.AZURE_OPENAI_WHISPER_MODE || 'transcriptions').toLowerCase() === 'translations'
    ? 'translations'
    : 'transcriptions';
  const apiVersion = env.AZURE_OPENAI_WHISPER_API_VERSION || env.AZURE_OPENAI_API_VERSION || '2024-06-01';
  return `${endpoint}/openai/deployments/${encodeURIComponent(deployment)}/audio/${mode}?api-version=${encodeURIComponent(apiVersion)}`;
};

export const extractUploadMetadata = async (
  env: SearchEnv,
  body: Record<string, unknown>,
  context: UploadAiContext = {}
): Promise<MetadataExtraction> => {
  const documentText = truncate(String(body.documentText || ''), toNumber(env.SEARCH_UPLOAD_METADATA_MAX_CHARS, 120000));
  if (!documentText.trim()) {
    throw new Error('documentText is required.');
  }

  const startedAt = Date.now();
  const prompt = buildDocumentMetadataPrompt(
    documentText,
    String(body.docTypeTermsSection || ''),
    String(body.clientTermsSection || '')
  );
  const completion = await getChatCompletionResult(
    env,
    [
      { role: 'system', content: 'You extract Knowledge Hub upload metadata. Return only valid JSON.' },
      { role: 'user', content: prompt }
    ],
    { maxTokens: toNumber(env.SEARCH_UPLOAD_METADATA_MAX_TOKENS, 1200), temperature: 0.1 }
  );

  const metadata = sanitizeDocumentMetadata(parseJsonFromModel(completion.content));
  context.log?.({
    event: 'upload-ai-metadata-completed',
    durationMs: Date.now() - startedAt,
    deployment: completion.deployment,
    promptTokens: completion.usage?.prompt_tokens,
    completionTokens: completion.usage?.completion_tokens
  });
  return metadata;
};

export const analyzeMediaTranscript = async (
  env: SearchEnv,
  body: Record<string, unknown>,
  context: UploadAiContext = {}
): Promise<MediaAnalysis> => {
  const transcriptText = truncate(String(body.transcriptText || ''), toNumber(env.SEARCH_UPLOAD_TRANSCRIPT_MAX_CHARS, 120000));
  if (!transcriptText.trim()) {
    throw new Error('transcriptText is required.');
  }

  const startedAt = Date.now();
  const prompt = buildMediaAnalysisPrompt(
    String(body.fileName || 'media file'),
    transcriptText,
    String(body.docTypeTermsSection || ''),
    String(body.clientTermsSection || '')
  );
  const completion = await getChatCompletionResult(
    env,
    [
      { role: 'system', content: 'You extract metadata from media transcripts. Return only valid JSON.' },
      { role: 'user', content: prompt }
    ],
    { maxTokens: toNumber(env.SEARCH_UPLOAD_MEDIA_ANALYSIS_MAX_TOKENS, 900), temperature: 0.1 }
  );

  const metadata = sanitizeMediaAnalysis(parseJsonFromModel(completion.content));
  context.log?.({
    event: 'upload-ai-media-analysis-completed',
    durationMs: Date.now() - startedAt,
    deployment: completion.deployment,
    promptTokens: completion.usage?.prompt_tokens,
    completionTokens: completion.usage?.completion_tokens
  });
  return metadata;
};

export const transcribeUploadedMedia = async (
  env: SearchEnv,
  file: UploadedMediaFile,
  context: UploadAiContext = {}
): Promise<{ transcript: Array<{ time: number; text: string }>; transcriptText: string }> => {
  const apiKey = env.AZURE_OPENAI_WHISPER_API_KEY || env.AZURE_OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error('Missing required env value: AZURE_OPENAI_WHISPER_API_KEY');
  }

  const startedAt = Date.now();
  const deployment = env.AZURE_OPENAI_WHISPER_DEPLOYMENT || 'whisper';
  const formData = new FormData();
  const blob = new Blob([file.buffer as unknown as BlobPart], {
    type: file.contentType || 'application/octet-stream'
  });
  formData.append('file', blob, file.filename || 'media');
  formData.append('model', deployment);
  formData.append('response_format', 'verbose_json');

  const response = await fetch(buildWhisperUrl(env), {
    method: 'POST',
    headers: {
      'api-key': apiKey
    },
    body: formData
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Azure OpenAI Whisper failed ${response.status}: ${text.slice(0, 700)}`);
  }

  const data = await response.json() as {
    text?: string;
    segments?: Array<{ start?: number; text?: string }>;
  };
  const transcript = Array.isArray(data.segments)
    ? data.segments.map((segment) => ({
      time: Math.floor(Number(segment.start || 0)),
      text: normalizeString(segment.text)
    })).filter((segment) => segment.text)
    : [{ time: 0, text: normalizeString(data.text) }].filter((segment) => segment.text);
  const transcriptText = transcript.length > 0
    ? transcript.map((segment) => segment.text).join(' ')
    : normalizeString(data.text);

  context.log?.({
    event: 'upload-ai-transcription-completed',
    durationMs: Date.now() - startedAt,
    fileName: file.filename,
    fileBytes: file.buffer.length,
    segmentCount: transcript.length
  });

  return { transcript, transcriptText };
};
