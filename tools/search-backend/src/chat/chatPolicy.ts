import { normalizeDomainRetrievalPhrase } from '../search/domainNormalizations';

export type ChatIntentKind =
  | 'empty'
  | 'greeting'
  | 'capability'
  | 'unsupported_action'
  | 'off_topic'
  | 'ambiguous_followup'
  | 'people'
  | 'document'
  | 'metadata_listing'
  | 'mixed';

export interface ChatIntent {
  kind: ChatIntentKind;
  shouldRetrieveDocuments: boolean;
  shouldRetrievePeople: boolean;
  shouldUseDocumentMetadata: boolean;
  retrievalQuery?: string;
  conversationContext?: string;
  directAnswer?: string;
}

const scopedRedirect =
  'I can still help you look for the right Active Knowledge Hub document or iKnowledgeNext topic.';

const outOfScopeAnswer = (value: string): string => {
  const normalized = normalize(value);
  if (/\b(pm|prime minister|president|minister|government|parliament)\b/.test(normalized)) {
    return 'That is outside the Knowledge Hub scope I can use here. I can help with Indegene, iKnowledgeNext, or Active Knowledge Hub document questions instead.';
  }

  if (/^\s*(?:who\s+is|who's|whose|tell me about)\s+[a-z][a-z.'-]+\s+[a-z][a-z.'-]+/.test(normalized)) {
    return 'I cannot answer personal-profile, employee-directory, or general biography questions in KM Assistant right now. I can help with Indegene, iKnowledgeNext, or Active Knowledge Hub document questions instead.';
  }

  if (/\b(ipl|cricket|football|sports?|team|match|score|league|t20|odi|test match)\b/.test(normalized)) {
    return 'That looks like a sports or general-world question, so I cannot answer it from KM Assistant. I can help with Indegene, iKnowledgeNext, or Active Knowledge Hub documents if you share a Knowledge Hub topic, client, BU, department, therapy area, disease area, or document type.';
  }

  if (/\b(actor|actress|celebrity|player|politician|prime minister|president|minister|movie|song|lyrics)\b/.test(normalized)) {
    return 'That is outside the Knowledge Hub scope I can use here. I can help with Indegene, iKnowledgeNext, or Active Knowledge Hub document questions instead.';
  }

  return 'That topic is outside the current Knowledge Hub scope I can use. I can help with Indegene, iKnowledgeNext, or Active Knowledge Hub documents by topic, client, BU, department, therapy area, disease area, author, or document type.';
};

// Client requested legacy people-directory answers be removed from KM Assistant for now.
// The old intent shape is retained so the feature can be re-enabled later without rebuilding the chat contract.
const CHAT_PEOPLE_DIRECTORY_ENABLED = false;

const normalize = (value: string): string =>
  value
    .toLowerCase()
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201c\u201d]/g, '"')
    .replace(/\s+/g, ' ')
    .trim();

const hasAny = (value: string, patterns: RegExp[]): boolean =>
  patterns.some((pattern) => pattern.test(value));

const FOLLOW_UP_PATTERNS = [
  /^(no|nope|nah|not that|actually|i mean|sorry|ok but|okay but|but)\b/,
  /^(ok|okay|cool|fine|alright)[,\s]+(?:now|then|so)?\b.*\b(it|this|that|them|these|those|same|under|source|sources|result|results|above|previous|earlier)\b/,
  /^(what about|and|also|that one|those|them|it|this|same|more|explain more|tell me more|any other|other documents|more documents|supporting this)\b/,
  /^(?:now\s+)?(?:only|just|filter|which|what|any)\b.*\b(?:case\s+stud(?:y|ies)|proposals?|rfps?|rfis?|sops?|polic(?:y|ies)|guidelines?|capabilit(?:y|ies)|decks?|brochures?|white\s*papers?|blogs?|videos?|recordings?)\b/,
  /^(who is he|who is she|what is that|summarize it|summarize that|what does it say|what does that say)$/i,
  /^(?:some|top|few|best|strongest|most relevant)[,\s]+.*\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|assets?|materials?|links?|references?)\b/,
  /\b(summarize|explain|describe|compare)\b.*\b(best matching|top result|first result|that document|that deck|that certificate|that source|retrieved source|previous source)\b/,
  /\b(only|just)\b.*\b(retrieved source|source actually says|that source says|the source says|actual source)\b/,
  /\b(current|latest|newest|most recent)\b.*\b(source|sources|result|results|above|those|them)\b/,
  /^(what|who|where|which|when|why|how|can you|can u|could you|could u|please|pls|tell me|show me|find|list|share)\b.*\b(it|this|that|them|these|those|he|she|they|his|her|their|same)\b/,
  /^(thanks|thank you|ok thanks|okay thanks)[,\s]+.*\b(more|other|additional)\b.*\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|sources?|references?)\b/,
  /\b(pull|get|fetch|show|list)\b.*\b(more|other|additional)\b.*\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|sources?|references?)\b/,
  /\b(more|other|additional|supporting|similar|related|before|previous|earlier)\b.*\b(this|that|these|those|same|one|topic|document|team|thing)\b/
];

const isFollowUpQuestion = (value: string): boolean => hasAny(value, FOLLOW_UP_PATTERNS);

const isCorrectionFollowUp = (value: string): boolean =>
  /^(no|nope|nah|not that|actually|i mean|sorry)\b/.test(value);

const isAffirmativeFollowUp = (value: string): boolean =>
  /^(yes|yeah|yep|yup|please|pls|sure|go ahead|do it|give|give me|show|show me)\b/.test(value);

export const isCredentialDisclosureQuery = (rawValue: string): boolean => {
  const value = normalize(rawValue);
  const credentialTerm = /\b(api\s*keys?|admin\s*keys?|access\s*keys?|secret|secrets|client\s*secrets?|password|passwords|passwd|token|tokens|bearer|connection\s*strings?|private\s*keys?|credential|credentials|key\s*vault)\b/;
  const disclosureVerb = /\b(share|show|tell|give|send|reveal|print|display|expose|provide|what is|what are|where is|where are|copy|paste)\b/;
  return credentialTerm.test(value) && disclosureVerb.test(value);
};

const isMedicalAdviceQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(cure|treat|treatment|diagnose|diagnosis|medicine|medication|drug|dose|dosage|symptom|therapy)\b.*\b(cancer|disease|patient|illness|infection|pain)\b/,
    /\b(cancer|disease|patient|illness|infection|pain|symptoms?)\b.*\b(medicine|medication|drug|dose|dosage|take|use|treatment|therapy)\b/,
    /\bwhat\s+(medicine|medication|drug|dose|dosage)\b.*\b(take|use)\b/,
    /\b(should|can|may|do)\s+i\b.*\b(take|use|consume|start|stop)\b.*\b(drug|medicine|medication|paracetamol|tablet|pill|dose|dosage)\b/,
    /\b(how|can you|what should i)\b.*\b(cure|treat|diagnose|medicate)\b/,
    /\bhow to cure\b/
  ]);

const isPersonalOrTransactionalQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(lend|loan|laon|borrow|send|transfer|give)\b.*\b(money|cash|rupees?|rs|inr|dollars?|salary)\b/,
    /\b(lend|loan|borrow|send|transfer|give)\b.*\b\d+\s*(k|rs|inr|rupees?|dollars?)\b/,
    /\b(give|provide|offer|need|want)\b.*\b(loan|laon)\b/,
    /\b(i love\s*you|love me|marry me|date me|do you love|relationship with me)\b/,
    /\b(are you human|do you have feelings|do you feel)\b/
  ]);

const isUnsupportedKnowledgeHubItemAction = (value: string): boolean =>
  hasAny(value, [
    /\b(create|delete|approve|reject|archive|upload|download|edit|update|rename|move)\b.*\b(document|documents|doc|docs|file|files|record|records|item|items|version)\b/,
    /\b(delete|approve|reject|archive|upload|download|edit|update|rename|move)\b.*\b(agentic|holy book|payment|sop|policy|certificate|deck|manual|plan|file|doc)\b/,
    /\b(download|send|email|forward)\b.*\b(recording|media|video|audio|mp4|mp3|everyone|all users|all employees|all colleagues)\b/,
    /\b(download|send|email|forward)\b.*\b(all|every|active|these|those|the)?\s*(document|documents|doc|docs|file|files|record|records|item|items|source|sources|reference|references)\b/,
    /\bshare\b.*\b(with|to)\b.*\b(manager|someone|everybody|everyone|all users|all employees|all colleagues|team|teams|client|clients|email|mail)\b/,
    /\b(reset password|raise ticket|book meeting|send email|grant access|change permission)\b/
  ]);

const isUnsafeAdultQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(nsfw|porn|pornography|explicit sexual|sex video|nude|nudes|naked|erotic|escort|hookup|sext|sexting)\b/,
    /\b(write|create|generate|describe|explain|show|share|send)\b.*\b(sexual|porn|pornographic|erotic|nude|naked|nsfw)\b/
  ]);

const isBulkSensitiveDirectoryQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(list|show|give|export|download|print|send|email|share)\b.*\b(all|every|entire|full)\b.*\b(employee|employees|people|person|staff|user|users)\b.*\b(email|emails|contact|contacts|phone|phones)\b/,
    /\b(list|show|give|export|download|print|send|email|share)\b.*\b(email|emails|contact|contacts|phone|phones)\b.*\b(all|every|entire|full)\b.*\b(directory|employee|employees|people|staff|users)\b/,
    /\b(every|all)\b.*\b(employee|employees|people|staff|users)\b.*\b(email|emails|contact|contacts|phone|phones)\b/
  ]);

const isExternalScrapingOrHarvestingQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(scrape|scraping|crawl|crawler|harvest|harvesting|extract|collect|dump)\b.*\b(linkedin|facebook|instagram|twitter|x\.com|profiles?|contacts?|emails?|phone numbers?)\b/,
    /\b(write|create|generate|build|give me)\b.*\b(script|code|bot|crawler|scraper)\b.*\b(linkedin|profiles?|contacts?|emails?)\b/,
    /\b(linkedin|profiles?)\b.*\b(scrape|scraping|crawler|scraper|harvest|extract|collect)\b/
  ]);

export const shouldSuppressDocumentSearchQuery = (rawValue: string): boolean => {
  const value = normalize(rawValue);
  if (!value) return false;

  if (isCredentialDisclosureQuery(value) || isPromptDisclosureQuery(value) || isBulkSensitiveDirectoryQuery(value) || isExternalScrapingOrHarvestingQuery(value) || isUnsafeAdultQuery(value)) {
    return true;
  }

  if (hasAny(value, [
    /\b(drop|truncate|wipe|erase|destroy|delete|reset|purge|clear|disable|shutdown|stop|kill)\b.*\b(index|indexes|indices|database|db|table|tables|search|service|backend|server|system|everything|all data|records?)\b/,
    /\b(grant|change|elevate|bypass|disable)\b.*\b(access|permission|permissions|security|auth|authentication|authorization|admin|policy|policies)\b/
  ])) {
    return true;
  }

  const asksForDocument = hasAny(value, [/\b(document|documents|doc|docs|file|files|knowledge hub|find|search|show|list|summarize)\b/]);
  if (isMedicalAdviceQuery(value) && !asksForDocument) {
    return true;
  }

  if (isPersonalOrTransactionalQuery(value)) {
    return true;
  }

  return hasAny(value, [
    /\b(download|send|email|forward)\b.*\b(all|every|active|documents|files|records|items|everyone|all users|all employees|all colleagues)\b/,
    /\b(create|delete|approve|reject|archive|upload|download|edit|update|rename|move)\b.*\b(document|file|record|item|version)\b/
  ]);
};

const isSmallTalkQuery = (value: string): boolean =>
  hasAny(value, [
    /^(hi|hello|hey|yo|good morning|good afternoon|good evening|good night|thanks|thank you|ok thanks|okay thanks)\b[!.?, ]*$/,
    /^hi+\b[!.?, ]*$/,
    /^(hi|hello|hey|yo|good morning|good afternoon|good evening|good night)\b.*\b(how are you|how are u|how r u|what'?s up|how is it going|how are things)\b[!.?, ]*$/,
    /^(hi|hello|hey|yo|good morning|good afternoon|good evening|good night)\b.*\bhow\s+(are|r|am)\s+(you|u|i)\b[!.?, ]*$/,
    /^(how are you|how are u|how r u|how am i|how am i doing|what'?s up|whatsup|what is up|how is it going|how are things|are you there)\b[!.?, ]*$/,
    /^(ok|okay|cool|great|nice|fine|alright)\b[!.?, ]*$/
  ]);

const isCapabilityQuestion = (value: string): boolean =>
  hasAny(value, [
    /\b(what can you do|who are you|what are you|what is your name|what's your name|your name|your purpose|how can you help)\b/,
    /\b(what can and cant (?:you|u) do|what can (?:you|u) and cant (?:you|u) do|what can(?:not|'?t)? (?:you|u) do)\b/,
    /\b(what should i use (?:this|the)?\s*(?:chatbot|bot|assistant) for|what can i use (?:this|the)?\s*(?:chatbot|bot|assistant) for|how should i use (?:this|the)?\s*(?:chatbot|bot|assistant))\b/,
    /\b(what should i ask|what can i ask|what questions can i ask|what kind of questions can i ask)\b/,
    /\b(what do you support|what can i ask|what model|which model|model are you using|gpt\s*5|gpt-?5)\b/,
    /\b(how does|how do)\b.*\b(search|chatbot|bot|assistant|km assistant)\b.*\b(work|works|working)\b/
  ]);

const capabilityDirectAnswer = (): string =>
  [
    'I can help you find and understand Active Knowledge Hub content.',
    '',
    'You can ask me to find documents by topic, client, BU, department, therapy area, disease area, document type, author, or keywords. I can also summarize or explain matching Knowledge Hub documents when the available sources contain enough evidence.',
    '',
    'I cannot send, share, edit, delete, approve, download, or expose restricted information on your behalf. If you ask to "share" or "show" documents, I will treat that as a request to find and list the relevant Knowledge Hub sources unless you explicitly ask me to send them to someone.'
  ].join('\n');

const isTextRewriteUtilityQuestion = (value: string): boolean =>
  hasAny(value, [
    /\b(fix|correct|improve|polish|rewrite|rephrase)\b.*\b(grammar|sentence|wording|text|copy|paragraph)\b/,
    /\b(grammar|grammatical|spelling)\b.*\b(fix|correct|check|improve)\b/
  ]);

const isUserIdentityQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(who am i|what is my name|what's my name|do you know me)\b/,
    /\b(my manager|name of my manager|who is my manager|who do i report to)\b/
  ]);

const isPromptDisclosureQuery = (value: string): boolean =>
  hasAny(value, [
    /\b(system|hidden|developer|internal)\s+(prompt|instruction|instructions|message|messages)\b/,
    /\b(ignore|override|bypass|forget)\b.*\b(instruction|instructions|rules|policy|policies|system prompt)\b/,
    /\b(print|show|reveal|display|tell me|share)\b.*\b(prompt|instructions|rules|policy)\b/
  ]);

const hasKnowledgeScopeHint = (value: string): boolean =>
  hasAny(value, [
    /\b(indegene|iknowledge|knowledge hub|km assistant|km|document|documents|docs?|file|files|asset|assets|library)\b/,
    /\b(client|bu|business unit|department|therapy area|disease area|document type|author|proposal|sop|case stud(?:y|ies)|rfp|rfi|training|deck|decks|capability|capabilities|esg)\b/,
    /\b(pharma|life sciences|healthcare|medical|clinical|regulatory|commercial|market access|payer|payer engagement|heor|prma|health economics|outcomes research|patient|oncology|cancer|therapy|disease|dermo|cosmetic|cosmetics|pharmacovigilance|pv|icsr)\b/,
    /\b(hematology|rbc|rbcs|red blood cells?|wbc|wbcs|white blood cells?|blood|anesthesia|anaesthesia|anesthetic|anaesthetic|sedation|perioperative|guideline|guidelines|protocol|protocols|surgery|surgical)\b/,
    /\b(ai|agentic|analytics|matomo|internal traffic|business continuity|certificate|certification|medical device|device directive)\b/
  ]);

const isBroadPersonOrWorldQuestion = (value: string): boolean =>
  hasAny(value, [
    /^\s*(who|how)\s+(is|are|was|were)\b/,
    /^\s*(tell me about|do you know|can you explain)\b.*\b(he|she|they|person|player|actor|politician|celebrity|prime minister|president)\b/
  ]) && !hasKnowledgeScopeHint(value);

const isGeneralWorldKnowledgeQuestion = (value: string): boolean =>
  hasAny(value, [
    /^\s*(who|what|which|when|where)\s+(is|are|was|were|do|does|did)\b/,
    /^\s*(tell me about|explain|describe|define|can you share more about|share more about)\b/,
    /\b(pm|prime minister|president|minister|government|parliament|country|capital|ipl|cricket|football|fifa|nba|world cup|tournament|championship|match|won|winner|movie|actor|singer|song|lyrics|recipe|weather|stock|score)\b/
  ]) && !hasKnowledgeScopeHint(value);

const isClearlyGeneralOffTopicQuery = (value: string): boolean =>
  isBroadPersonOrWorldQuestion(value) || isGeneralWorldKnowledgeQuestion(value) || hasAny(value, [
    /\b(what(?:'s| is)?|tell me)\b.*\b(time|date|weather|temperature|score|stock price|exchange rate)\b/,
    /\b(who|which team)\b.*\b(won|winner|champion|championship|world cup|fifa|nba|ipl|match|tournament)\b/,
    /\b(who|what|how)\b.*\b(prime minister|president|celebrity|fictional character|anime character|sports person|cricketer|footballer|public figure)\b/,
    /\b(hi|hello|hey)\b.*\b(prime minister|president|celebrity|public figure)\b/,
    /\b(joke|story|poem|song|lyrics|recipe|translate|calculate|coding question|programming question)\b/,
    /\b(nsfw|porn|pornography|explicit sexual|sex video|nude|nudes|naked|erotic|escort|hookup|sext|sexting)\b/
  ]);

const hasClearlyGeneralWorldTopic = (value: string): boolean =>
  hasAny(value, [
    /\b(time|date|weather|temperature|score|stock price|exchange rate|capital city|recipe|movie|song|lyrics)\b/,
    /\b(pm|prime minister|president|minister|government|parliament|country|ipl|cricket|football|fifa|nba|world cup|tournament|championship|match|won|winner|actor|actress|celebrity|singer|fictional character|anime character)\b/
  ]);

const extractKnowledgeScopedQuestion = (value: string): string => {
  const normalizedValue = normalize(value.replace(/[/?]+$/g, ''));
  const clauses = normalizedValue
    .split(/\s+\b(?:and|also|plus|but)\b\s+|[;]+/i)
    .map((clause) => clause.trim())
    .filter(Boolean);

  if (clauses.length < 2) {
    return '';
  }

  const scopedClauses = clauses.filter((clause) => hasKnowledgeScopeHint(clause));
  const outOfScopeClauses = clauses.filter((clause) => isClearlyGeneralOffTopicQuery(clause));

  if (scopedClauses.length === 0 || outOfScopeClauses.length === 0) {
    return '';
  }

  return scopedClauses.join(' ').trim();
};

const likelyKnowledgeDocumentTopic = (value: string): boolean =>
  hasAny(value, [
    /\b(agentic|agent|ai|bcp|business continuity|certificate|certification)\b/,
    /\b(policy|sop|procedure|manual|deck|training|approach plan|feasibility|proposal|proposals|capability|capabilities|asset|assets|assest|assests|reference|playbook|guideline)\b/,
    /\b(cancer|oncology|tumou?r|breast|gastric|ovary|ovarian|dermo|cosmetic|cosmetics|medical device|device directive|regulatory|market access|payer|heor|prma|health economics|outcomes research|asean|agreement|analytics|matomo|internal traffic|pharmacovigilance|icsr)\b/,
    /\b(hematology|rbc|rbcs|red blood cells?|wbc|wbcs|white blood cells?|blood|anesthesia|anaesthesia|anesthetic|anaesthetic|sedation|perioperative|clinical guidelines?|protocols?)\b/
  ]);

const looksLikeNamedKnowledgeAssetQuestion = (value: string): boolean => {
  const normalized = normalize(value);
  if (!/\b(what|which|tell|describe|explain|summarize|overview)\b/i.test(normalized)) return false;
  if (!/\b(about|called|titled|named|document|asset|file|deck|article|report|orchestration)\b/i.test(normalized)) return false;
  const retrievalTopic = cleanDocumentRetrievalQuery(normalized);
  if (tokenCount(retrievalTopic) < 2) return false;
  return !hasAny(retrievalTopic, [
    /\b(time|date|weather|temperature|score|stock price|exchange rate)\b/,
    /\b(prime minister|president|celebrity|fictional character|anime character|sports person|cricketer|footballer|public figure)\b/
  ]);
};

const looksLikeKnowledgeExplanationQuestion = (value: string): boolean => {
  if (!hasAny(value, [
    /^\s*(what|what's|explain|describe|define)\b/i,
    /^\s*(help me understand|tell me about|can you explain|can u explain)\b/i,
    /\b(what does|what do)\b.*\b(mean|cover|include|contain|say)\b/i
  ])) {
    return false;
  }

  if (hasClearlyGeneralWorldTopic(value)) {
    return false;
  }

  const retrievalTopic = cleanDocumentRetrievalQuery(value);
  return tokenCount(retrievalTopic) >= 2 || likelyKnowledgeDocumentTopic(retrievalTopic) || hasKnowledgeScopeHint(retrievalTopic);
};

const isLowInformationSearchNoise = (value: string): boolean => {
  if (hasKnowledgeScopeHint(value) || likelyKnowledgeDocumentTopic(value) || isExplicitDocumentRequest(value) || isExplicitMetadataListingRequest(value)) {
    return false;
  }

  const tokens = value
    .replace(/[^a-z0-9\s]/gi, ' ')
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean);

  return tokens.length > 0
    && tokens.length <= 3
    && tokens.every((token) => token.length <= 4)
    && !hasAny(value, [/^(hi|hello|hey|thanks|thank you|ok|okay)\b/i]);
};

const isExplicitDocumentRequest = (value: string): boolean =>
  hasAny(value, [
    /\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|source|sources|retrieved source|policy|sop|certificate|certification|training|procedure|manual|deck|proposal|proposals|capability|capabilities|assets?|assests?|ppt|pptx|pdf|docx|xlsx|csv|mp4|mp3|audio|media|video|recording|tracker|spreadsheet)\b/,
    /\b(show|find|list|search|share|pull|get|fetch|have|hbave|has)\b.*\b(cmmi|iso|certificate|certificates|doc(?:u?e?m?e?n?t?s?)?|documents?|files?|decks|proposals?|capabilit(?:y|ies)|assets?|assests?|audio|media|videos|trackers|spreadsheets)\b/
  ]);

const isExplicitMetadataListingRequest = (value: string): boolean =>
  hasAny(value, [
    /\b(show|find|list|share|pull|get|fetch|which|all|any|what|current|latest|newest|most recent|created|published|modified|last\s+\d+\s+(days?|weeks?|months?|years?)|total|count|number|no\.?)\b.*\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|source|sources|result|results|certificate|certificates|deck|decks|proposals?|capabilit(?:y|ies)|assets?|assests?|audio|media|videos|trackers|spreadsheets)\b/,
    /\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|source|sources|result|results|certificate|certificates|deck|decks|proposals?|capabilit(?:y|ies)|assets?|assests?|audio|media|videos|trackers|spreadsheets)\b.*\b(created|published|modified|current|latest|newest|most recent|last\s+\d+\s+(days?|weeks?|months?|years?)|total|count|number|no\.?)\b/,
    /\b(bu|business unit|department|sub department|client|region|geography|therapy area|disease area|document type|author|published|modified|version)\b/
  ]);

const isExplicitPeopleFollowUp = (value: string): boolean =>
  hasAny(value, [
    /\b(who|who's|whose|team|teams|lead|leader|head|heads|manager|contact|email|person|people|employee|employees|owner|owns|role|reports to|reporting|under him|under her|under that|sub-teams|subteams)\b/,
    /\bwhat\s+(?:he|she|they|[a-z][a-z.'-]+(?:\s+[a-z][a-z.'-]+){1,4})\s+do(?:es)?\b/
  ]);

const authorDocumentFallbackName = (value: string): string => {
  const normalized = normalize(value);
  const asksForDocumentsByPerson = hasAny(normalized, [
    /\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|assets?|files?)\b.*\b(author|authors|authored|created|uploaded|owner|by\s+(?:her|him|them|that person))\b/,
    /\b(author|authors|authored|created|uploaded|owner|by\s+(?:her|him|them|that person))\b.*\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|assets?|files?)\b/
  ]);

  if (!asksForDocumentsByPerson) return '';

  const nameToken = String.raw`[a-z](?:[a-z.'-]+)?`;
  const explicitAuthorMatch = normalized.match(new RegExp(String.raw`\b(?:author|authors|authored|created|uploaded|owner|by)\s+(?:by\s+)?(${nameToken}(?:\s+${nameToken}){0,4})\b`, 'i'))?.[1] || '';
  const explicitAuthor = /^(?:her|him|them|that person)$/i.test(explicitAuthorMatch.trim()) ? '' : explicitAuthorMatch;
  const whoSubject = normalized.match(new RegExp(String.raw`\bwho\s+is\s+(${nameToken}(?:\s+${nameToken}){1,4})(?:\s*,|\s+if\b|\?|$)`, 'i'))?.[1];
  const candidate = (explicitAuthor || whoSubject || '')
    .replace(/\b(?:if|cannot|can't|find|that|give|docs?|documents?|authored|created|uploaded|by|her|him|them)\b.*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();

  return tokenCount(candidate) >= 2 ? candidate : '';
};

const isLikelyPersonNameQuery = (value: string): boolean => {
  if (isSmallTalkQuery(value) || likelyPastedDocumentText(value) || hasKnowledgeScopeHint(value) || likelyKnowledgeDocumentTopic(value)) {
    return false;
  }

  if (hasAny(value, [
    /\b(what|which|where|when|why|how)\s+(is|are|was|were|does|do)\b/,
    /\b(member state|member states|government|governments|republic|kingdom|directive|agreement|analytics|traffic)\b/
  ])) {
    return false;
  }

  const ignored = new Set([
    'who', 'what', 'where', 'which', 'when', 'why', 'how', 'about', 'is', 'are', 'the',
    'and', 'or', 'dept', 'department', 'come', 'under', 'does', 'do', 'tell', 'me',
    'hi', 'hello', 'hey', 'yo', 'u', 'am', 'i', 'whatsup', 'please', 'pls'
  ]);
  const tokens = value
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean)
    .filter((token) => !ignored.has(token));

  return tokens.length === 2
    && /^\s*(?:who\s+is|who's|whose|tell me about)\b/.test(value)
    && !isExplicitDocumentRequest(value)
    && !likelyKnowledgeDocumentTopic(value);
};

const tokenCount = (value: string): number =>
  value.split(/\s+/).map((token) => token.trim()).filter(Boolean).length;

const likelyPastedDocumentText = (value: string): boolean =>
  tokenCount(value) >= 20 ||
  value.length >= 120 ||
  (/[\r\n]/.test(value) && tokenCount(value) >= 12) ||
  /\b(contents|definitions|purpose and scope|termination|confidentiality|intellectual property|governing law)\b/.test(value);

const hasSpecificQuotedTopic = (value: string): boolean =>
  /"[^"]{8,}"|'[^']{8,}'/.test(value);

const normalizeDomainTermsForRetrieval = (value: string): string =>
  normalizeDomainRetrievalPhrase(value);

const normalizeRetrievalText = (value: string): string =>
  normalizeDomainTermsForRetrieval(value)
    .trim()
    .replace(/\s+/g, ' ');

const extractQuotedTopic = (value: string): string => {
  const match = String(value || '').match(/"([^"]{2,120})"|'([^']{2,120})'/);
  return normalizeRetrievalText(match?.[1] || match?.[2] || '');
};

const removeRetrievalFiller = (value: string): string =>
  value
    .replace(/[?!.]+$/g, ' ')
    .replace(/\beven if\b.*$/gi, ' ')
    .replace(/\bnot\s+tagged\b/gi, ' ')
    .replace(/\b(?:can|could|would)\s+(?:you|u)\b/gi, ' ')
    .replace(/\b(?:can|could|would|should)\b/gi, ' ')
    .replace(/\bfor\s+me\b/gi, ' ')
    .replace(/^(?:what|which)\s+about\b/gi, ' ')
    .replace(/\bthis\s+is\s+(?:a\s+)?(?:doc|document|deck|file|asset)\b/gi, ' ')
    .replace(/\b(?:please|pls|kindly|me|them)\b/gi, ' ')
    .replace(/\b(?:share|show|find|fetch|pull|get|give|list|tell)\b(?:\s+me)?\b/gi, ' ')
    .replace(/\b(?:summarize|summary|brief|explain|describe)\b(?:\s+it)?\b/gi, ' ')
    .replace(/\b(?:all|any|some|more|every|each|about|related|relevant|matching|available|current|and)\b/gi, ' ')
    .replace(/\b(?:the|a|an|of|to|for|on|in)\b/gi, ' ')
    .replace(/\b(?:assets?|assests?|materials?|documents?|docuemnts?|docuemtns?|docs?|files?|links?|references?|sources?|list)\b/gi, ' ')
    .replace(/\b(?:from|under|within|inside)\s+(?:the\s+)?/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const cleanDocumentRetrievalQuery = (rawValue: string): string => {
  const original = normalizeRetrievalText(rawValue);
  if (!original || likelyPastedDocumentText(original) || hasSpecificQuotedTopic(original)) {
    return extractQuotedTopic(original) || original;
  }

  let value = removeRetrievalFiller(original);

  if (/\bpv\b/i.test(original) && /\b(doc|docs|documents?|files?|samples?|case processing|quality|safety|icsr|qms|training|sop|deck)\b/i.test(original)) {
    value = normalizeRetrievalText(value.replace(/\bpv\b/gi, 'pharmacovigilance'));
  }

  const fromMatch = original.match(/\b(?:from|under|within|inside)\s+(?:the\s+)?([^?.,;]{2,120})/i);
  if (fromMatch) {
    const scopedTopic = fromMatch[1]
      .replace(/\b(?:assets?|assests?|materials?|documents?|docs?|files?|links?|references?|sources?)\b/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const leadingQuestion = removeRetrievalFiller(original.slice(0, fromMatch.index));
    const shouldPreserveLeadingTopic =
      /\b(what|why|how|explain|summarize|describe|definition|meaning|used|use|purpose|patterns?|workflows?)\b/i.test(original.slice(0, fromMatch.index))
      && tokenCount(leadingQuestion) >= 2;

    value = shouldPreserveLeadingTopic
      ? normalizeRetrievalText(`${leadingQuestion} ${scopedTopic || ''}`)
      : scopedTopic || value;
  }

  const topicMatch = original.match(/\b(?:about|on|related to|for)\s+(?:the\s+)?([^?.,;]{2,120})/i);
  if (topicMatch && value.length < 3) {
    value = topicMatch[1]
      .replace(/\b(?:assets?|assests?|materials?|documents?|docs?|files?|links?|references?|sources?)\b/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  return normalizeRetrievalText(value.length >= 2 ? value : original);
};

const standaloneFollowUpTopicValue = (rawValue: string): string => {
  const normalized = normalize(rawValue);
  const match = normalized.match(/^(?:what about|and|also)\s+([a-z0-9][a-z0-9\s.'-]{1,80})[?!.]*$/i);
  const topic = cleanDocumentRetrievalQuery(match?.[1] || '');
  if (!topic) return '';
  const scopedShortAcronym = tokenCount(topic) === 1
    && /^[a-z0-9]{2,10}$/i.test(topic)
    && (hasKnowledgeScopeHint(topic) || likelyKnowledgeDocumentTopic(topic));
  if (tokenCount(topic) < 2 && !scopedShortAcronym) return '';
  if (/\b(it|this|that|them|these|those|same|above|previous|earlier)\b/i.test(topic)) return '';
  return (hasKnowledgeScopeHint(topic) || likelyKnowledgeDocumentTopic(topic) || tokenCount(topic) >= 3)
    ? topic
    : '';
};

const isMoreDocumentsFollowUp = (value: string): boolean =>
  /\b(more|similar|related|other|additional)\b/i.test(value)
  && /\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|assets?|materials?|sources?|references?)\b/i.test(value);

const isDocumentTypeFilterFollowUp = (value: string): boolean =>
  /\b(?:now|only|just|filter|which|what|any|these|those|among|from)\b/i.test(value)
  && /\b(?:case\s+stud(?:y|ies)|proposals?|rfps?|rfis?|sops?|polic(?:y|ies)|guidelines?|capabilit(?:y|ies)|decks?|brochures?|white\s*papers?|blogs?|videos?|recordings?)\b/i.test(value)
  && tokenCount(value) <= 10;

const isPreviousResultReferenceRequest = (value: string): boolean =>
  !isMoreDocumentsFollowUp(value)
  && !isDocumentTypeFilterFollowUp(value)
  && /\b(summarize|summary|brief|briefly|explain|describe|overview|more|details|tell me more|walk me through)\b/i.test(value)
  && /\b(first|second|third|fourth|fifth|1st|2nd|3rd|4th|5th|one|result|source|reference|document|doc|asset|found)\b/i.test(value);

const lastSubstantiveUserMessage = (history?: Array<{ role: string; content: string }>): string => {
  const messages = (history || [])
    .map((message) => ({
      role: message.role === 'assistant' ? 'assistant' : 'user',
      content: String(message.content || '').trim()
    }))
    .filter((message) => message.content);

  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    const normalized = normalize(message.content);
    if (message.role !== 'user'
      || isSmallTalkQuery(normalized)
      || isCredentialDisclosureQuery(normalized)
      || isPromptDisclosureQuery(normalized)
      || isPreviousResultReferenceRequest(normalized)
      || isMoreDocumentsFollowUp(normalized)
      || isDocumentTypeFilterFollowUp(normalized)
      || (isClearlyGeneralOffTopicQuery(normalized) && !isExplicitDocumentRequest(normalized) && !isExplicitMetadataListingRequest(normalized))) {
      continue;
    }

    const standaloneTopic = standaloneFollowUpTopicValue(normalized);
    if (standaloneTopic) {
      return standaloneTopic.slice(0, 700);
    }

    if (!isFollowUpQuestion(normalized)) {
      return cleanDocumentRetrievalQuery(message.content).slice(0, 700);
    }
  }

  return '';
};

const lastAssistantMessage = (history?: Array<{ role: string; content: string }>): string => {
  const messages = (history || [])
    .map((message) => ({
      role: message.role === 'assistant' ? 'assistant' : 'user',
      content: String(message.content || '').trim()
    }))
    .filter((message) => message.content);

  for (let index = messages.length - 1; index >= 0; index -= 1) {
    if (messages[index].role === 'assistant') {
      return messages[index].content.slice(0, 900);
    }
  }

  return '';
};

const requestedPreviousSourceOrdinal = (value: string): number | null => {
  const normalized = normalize(value);
  if (/\b(first|1st|one)\b/.test(normalized)) return 1;
  if (/\b(second|2nd|two)\b/.test(normalized)) return 2;
  if (/\b(third|3rd|three)\b/.test(normalized)) return 3;
  if (/\b(fourth|4th|four)\b/.test(normalized)) return 4;
  if (/\b(fifth|5th|five)\b/.test(normalized)) return 5;
  return null;
};

const recentAssistantSourceTitle = (
  rawQuestion: string,
  history?: Array<{ role: string; content: string }>
): string => {
  if (!/\b(source|sources|reference|references|document|doc|file|asset|one|that|this|first|second|third|fourth|fifth)\b/i.test(rawQuestion)) {
    return '';
  }

  const ordinal = requestedPreviousSourceOrdinal(rawQuestion) || 1;
  const assistantMessage = lastAssistantMessage(history);
  const sourceBlock = assistantMessage.split(/sources used:/i)[1] || '';
  const sourceTitles = sourceBlock
    .split(/\n+/)
    .map((line) => line.replace(/^\s*\d+[\).]\s*/, '').replace(/\s+-\s+https?:\/\/\S+$/i, '').trim())
    .filter(Boolean)
    .slice(0, 8);

  if (sourceTitles.length > 0) {
    return sourceTitles[ordinal - 1] || sourceTitles[0] || '';
  }

  if (/\b(?:could not find|couldn’t find|don't have|don’t have|not seeing|no matching|no documents|no results|try searching|try refining|author filter)\b/i.test(assistantMessage)) {
    return '';
  }

  const foundTitlePatterns = [
    /\bI\s+found\s+(?:"([^"]{4,180})"|'([^']{4,180})')/i,
    /\bI\s+found\s+(?:a\s+)?(?:closely\s+matching\s+)?(?:Knowledge Hub\s+)?(?:document|asset|file)\s*:?\s*(?:"([^"]{4,180})"|'([^']{4,180})')/i,
    /\b(?:document|asset|file)\s+(?:is|called|titled)\s+(?:"([^"]{4,180})"|'([^']{4,180})')/i
  ];
  for (const pattern of foundTitlePatterns) {
    const match = assistantMessage.match(pattern);
    const title = (match?.[1] || match?.[2] || match?.[3] || match?.[4] || '').trim();
    if (title) {
      return title;
    }
  }

  const numberedAnswerTitles = assistantMessage
    .split(/\n+/)
    .map((line) => {
      const match = line.match(/^\s*\d+[\).]\s+(?:\*\*)?(.+?)(?:\*\*)?(?:\s+[-–]\s+|\s*$)/);
      return (match?.[1] || '')
        .replace(/\[[^\]]+\]\(([^)]+)\)/g, '$1')
        .replace(/^["']|["']$/g, '')
        .trim();
    })
    .filter((title) =>
      title.length >= 4
      && title.length <= 180
      && !/\b(type|status|client|author|bu|business unit|why it may match|preview)\s*:/i.test(title)
    )
    .slice(0, 8);

  return numberedAnswerTitles[ordinal - 1] || numberedAnswerTitles[0] || '';
};

const isPreviousSourceContentFollowUp = (value: string): boolean =>
  !isMoreDocumentsFollowUp(value)
  && !isDocumentTypeFilterFollowUp(value)
  && /\b(summarize|summary|explain|describe|overview|more|details|tell me more|brief|briefly|walk me through)\b/i.test(value)
  && /\b(first|second|third|fourth|fifth|1st|2nd|3rd|4th|5th|one|that|this|source|reference|document|doc|asset|found)\b/i.test(value);

const recentConversationContext = (history?: Array<{ role: string; content: string }>): string => {
  const messages = (history || [])
    .map((message) => ({
      role: message.role === 'assistant' ? 'assistant' : 'user',
      content: String(message.content || '').trim()
    }))
    .filter((message) => message.content)
    .slice(-6);

  return messages
    .map((message) => `${message.role}: ${message.content}`)
    .join('\n')
    .slice(0, 1600);
};

const compactContextualRetrievalQuery = (
  rawQuestion: string,
  history?: Array<{ role: string; content: string }>
): string => {
  const isCorrection = isCorrectionFollowUp(normalize(rawQuestion));
  const cleanedQuestion = rawQuestion
    .trim()
    .replace(/^(no|nope|nah|not that|actually|i mean|sorry)\b[\s,.:;-]*/i, '')
    .replace(/^the\s+/i, '')
    .replace(/\bholybook\b/gi, 'holy book')
    .replace(/\bia agent\b/gi, 'ai agent');
  const previousTopic = lastSubstantiveUserMessage(history);
  const shouldUsePreviousTopic = !(isCorrection && tokenCount(cleanedQuestion) >= 3);
  const previousAnswer = isCorrection
    ? ''
    : lastAssistantMessage(history)
      .replace(/\s*\[(?:\d+|[1-9]\d*\s*[-–]\s*[1-9]\d*)(?:\s*,\s*(?:\d+|[1-9]\d*\s*[-–]\s*[1-9]\d*))*\]/g, '')
      .replace(/\s+/g, ' ')
      .slice(0, 450);
  const previousSourceTitle = recentAssistantSourceTitle(rawQuestion, history);
  const previousAssistantRejectedScope =
    /\b(?:cannot expose hidden prompts|private instructions|internal configuration|outside\s+the\s+(?:current\s+)?knowledge hub scope|outside\s+km assistant'?s scope|cannot answer personal-profile|cannot answer employee-directory)\b/i
      .test(previousAnswer);
  const genericPronounDocumentRequest = /\b(find|search|show|list|share|get|fetch|pull)\b/i.test(cleanedQuestion)
    && /\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|assets?|materials?|links?|references?)\b/i.test(cleanedQuestion)
    && /\b(it|this|that|them|these|those|same)\b/i.test(cleanedQuestion);
  const genericShortListingFollowUp = /\b(?:some|top|few|best|strongest|most relevant|relevant)\b/i.test(cleanedQuestion)
    && /\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|assets?|materials?|links?|references?)\b/i.test(cleanedQuestion)
    && tokenCount(cleanedQuestion) <= 8;
  const filterFollowUp = isDocumentTypeFilterFollowUp(cleanedQuestion);
  const usesPreviousTopicAsPrimary = (genericPronounDocumentRequest || genericShortListingFollowUp || filterFollowUp) && previousTopic && shouldUsePreviousTopic;
  const primaryQuestion = usesPreviousTopicAsPrimary
    ? normalizeRetrievalText(`${previousTopic} ${filterFollowUp ? cleanedQuestion : ''}`)
    : cleanedQuestion || rawQuestion.trim();

  return [
    normalizeRetrievalText(primaryQuestion),
    previousSourceTitle ? `Previous referenced source title: ${previousSourceTitle}` : '',
    previousTopic && shouldUsePreviousTopic && primaryQuestion !== previousTopic ? `Previous topic: ${previousTopic}` : '',
    previousAnswer && !previousAssistantRejectedScope && !usesPreviousTopicAsPrimary ? `Previous answer summary: ${previousAnswer}` : ''
  ].filter(Boolean).join('\n');
};

export const classifyChatIntent = (
  rawQuestion: string,
  history?: Array<{ role: string; content: string }>
): ChatIntent => {
  const question = normalize(rawQuestion);
  const conversationContext = recentConversationContext(history);
  const contextText = normalize(conversationContext);
  const pastedDocumentLike = likelyPastedDocumentText(question);
  const knowledgeScopedQuestion = extractKnowledgeScopedQuestion(rawQuestion);
  const quotedSpecificTopic = hasSpecificQuotedTopic(rawQuestion);
  const followUpSubject = question.replace(/^(what about|and|also)\b[\s,.:;-]*/i, '').trim();
  const standaloneFollowUpTopic = quotedSpecificTopic
    || (/^(what about|and|also)\b/.test(question)
      && tokenCount(followUpSubject) >= 2
      && !/\b(it|this|that|them|these|those|same|above|previous|earlier)\b/.test(followUpSubject)
      && (hasKnowledgeScopeHint(followUpSubject) || likelyKnowledgeDocumentTopic(followUpSubject) || isExplicitDocumentRequest(followUpSubject) || isExplicitMetadataListingRequest(followUpSubject)));

  if (!question) {
    return {
      kind: 'empty',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'Ask me about an Active Knowledge Hub document, topic, client, BU, department, therapy area, disease area, or document type and I will look it up.'
    };
  }

  if (isCredentialDisclosureQuery(question)) {
    return {
      kind: 'unsupported_action',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot help reveal or retrieve credentials such as keys, secrets, passwords, tokens, or connection strings. If you need the approved process, I can look for access-request or credential-rotation documents in Knowledge Hub.'
    };
  }

  if (isPromptDisclosureQuery(question)) {
    return {
      kind: 'unsupported_action',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot expose hidden prompts, private instructions, or internal configuration. I can answer from Active Knowledge Hub sources that are meant to be searched.'
    };
  }

  if (isSmallTalkQuery(question)) {
    return {
      kind: 'greeting',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      conversationContext
    };
  }

  if (isCapabilityQuestion(question)) {
    return {
      kind: 'capability',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      conversationContext,
      directAnswer: capabilityDirectAnswer()
    };
  }

  if (isTextRewriteUtilityQuestion(question) && !hasKnowledgeScopeHint(question) && !isExplicitDocumentRequest(question)) {
    return {
      kind: 'capability',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'Yes. Paste the text you want improved and I can help rewrite or correct the grammar. For factual answers about Knowledge Hub content, I will use Active Knowledge Hub sources and cite the matching documents.'
    };
  }

  if (isUnsupportedKnowledgeHubItemAction(question)) {
    return {
      kind: 'unsupported_action',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      conversationContext,
      directAnswer: 'I cannot send, email, forward, download, edit, delete, approve, or change Knowledge Hub items on your behalf. If you want to find documents, ask me to show or list them by topic, client, BU, department, therapy area, disease area, document type, or title.'
    };
  }

  if (isUserIdentityQuery(question)) {
    return {
      kind: 'capability',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot identify personal profiles or employee-directory details in KM Assistant right now. I can still help search Active Knowledge Hub documents by topic, author name, client, BU, department, therapy area, disease area, or document type.'
    };
  }

  if (isBulkSensitiveDirectoryQuery(question)) {
    return {
      kind: 'unsupported_action',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot bulk-list or export employee contact details. I can help search Active Knowledge Hub documents by author, topic, client, BU, department, therapy area, disease area, or document type.'
    };
  }

  if (isExternalScrapingOrHarvestingQuery(question)) {
    return {
      kind: 'unsupported_action',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot help scrape, harvest, or collect external profile/contact data. I can help find approved Knowledge Hub documents about compliant data, analytics, or web processes if that is what you need.'
    };
  }

  const fallbackAuthorName = authorDocumentFallbackName(question);
  if (fallbackAuthorName) {
    return {
      kind: 'metadata_listing',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: true,
      retrievalQuery: `author ${fallbackAuthorName}`
    };
  }

  if (isMedicalAdviceQuery(question) && !hasAny(question, [/\b(document|documents|file|files|knowledge hub|find|search|show|list|summarize)\b/])) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot give medical advice, diagnose, or recommend treatment. I can help find or summarize Knowledge Hub material on oncology, clinical, regulatory, or medical topics, but personal health decisions need a qualified medical professional.'
    };
  }

  if (isPersonalOrTransactionalQuery(question)) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: `I cannot help with personal transactions or relationship requests. ${scopedRedirect}`
    };
  }

  const previousSourceTitle = recentAssistantSourceTitle(rawQuestion, history);
  if (previousSourceTitle && isPreviousSourceContentFollowUp(rawQuestion)) {
    return {
      kind: 'document',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      retrievalQuery: previousSourceTitle,
      conversationContext
    };
  }

  if (!previousSourceTitle && conversationContext && isPreviousSourceContentFollowUp(rawQuestion)) {
    return {
      kind: 'ambiguous_followup',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I do not have a Knowledge Hub result selected yet. Search for a document or share the title, and I can summarize it.'
    };
  }

  if (knowledgeScopedQuestion) {
    return {
      kind: 'document',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: isExplicitMetadataListingRequest(knowledgeScopedQuestion),
      retrievalQuery: cleanDocumentRetrievalQuery(knowledgeScopedQuestion),
      conversationContext
    };
  }

  const shortFollowUpTopic = question.match(/^(?:what about|and|also)\s+([a-z0-9][a-z0-9\s.'-]{1,40})[?!.]*$/i)?.[1]?.trim() || '';
  const previousTurnWasOutOfScope = hasAny(contextText, [
    /\boutside\s+the\s+(?:current\s+)?knowledge hub scope\b/,
    /\boutside\s+km assistant'?s scope\b/,
    /\boutside\s+the\s+scope\b/
  ]);
  if (conversationContext
    && previousTurnWasOutOfScope
    && shortFollowUpTopic
    && tokenCount(shortFollowUpTopic) <= 3
    && /^[a-z0-9. '-]{2,12}$/i.test(shortFollowUpTopic)
    && !isExplicitDocumentRequest(question)) {
    return {
      kind: 'ambiguous_followup',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: `Can you clarify what "${shortFollowUpTopic}" means in the Knowledge Hub context? For example, is it a client, project, acronym, medical topic, or document title?`
    };
  }

  if (isUnsafeAdultQuery(question)) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot help with explicit adult or NSFW requests. I can help with appropriate Knowledge Hub topics, including life sciences, healthcare, pharma, regulatory, commercial, and Indegene-related documents.'
    };
  }

  if (isLowInformationSearchNoise(question)) {
    return {
      kind: 'ambiguous_followup',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: `Can you clarify what "${rawQuestion.trim()}" means in the Knowledge Hub context? For example, is it a client, project, acronym, topic, or document title?`
    };
  }

  if (looksLikeKnowledgeExplanationQuestion(question)) {
    return {
      kind: 'document',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      retrievalQuery: cleanDocumentRetrievalQuery(rawQuestion),
      conversationContext
    };
  }

  if (isClearlyGeneralOffTopicQuery(question)
    && !isExplicitDocumentRequest(question)
    && !likelyKnowledgeDocumentTopic(question)
    && !looksLikeNamedKnowledgeAssetQuestion(question)) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      conversationContext,
      directAnswer: outOfScopeAnswer(rawQuestion)
    };
  }

  if (isAffirmativeFollowUp(question) && conversationContext) {
    const contextualQuestion = compactContextualRetrievalQuery(rawQuestion, history);
    const peopleIntentFromContext = hasAny(contextText, [
      /\b(who|whose|team|lead|leader|head|manager|contact|email|person|people(?!'s)|employee|employees|owner|owns|role|reports to|reporting)\b/,
      /\b(directory)\b/
    ]);
    const metadataIntentFromContext = hasAny(contextText, [
      /\b(bu|business unit|department|sub department|client|region|geography|therapy area|disease area|document type|author|published|modified|version)\b/,
      /\b(list|show|find|which|all)\b.*\b(documents|files|policies|sops|certificates|decks)\b/
    ]);

    return {
      kind: metadataIntentFromContext ? 'metadata_listing' : 'document',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: CHAT_PEOPLE_DIRECTORY_ENABLED && peopleIntentFromContext,
      shouldUseDocumentMetadata: metadataIntentFromContext,
      retrievalQuery: normalizeRetrievalText(contextualQuestion),
      conversationContext
    };
  }

  if (!standaloneFollowUpTopic && isFollowUpQuestion(question) && (!pastedDocumentLike || tokenCount(question) <= 24)) {
    const shortNewFollowUpTopic = shortFollowUpTopic;
    if (shortNewFollowUpTopic
      && tokenCount(shortNewFollowUpTopic) <= 4
      && !hasKnowledgeScopeHint(shortNewFollowUpTopic)
      && !likelyKnowledgeDocumentTopic(shortNewFollowUpTopic)
      && !isExplicitDocumentRequest(shortNewFollowUpTopic)) {
      return {
        kind: 'ambiguous_followup',
        shouldRetrieveDocuments: false,
        shouldRetrievePeople: false,
        shouldUseDocumentMetadata: false,
        directAnswer: `Can you clarify what "${shortNewFollowUpTopic}" means in the Knowledge Hub context? For example, is it a client, project, topic, acronym, or document title?`
      };
    }

    if (conversationContext) {
      const contextualQuestion = compactContextualRetrievalQuery(rawQuestion, history);
      const peopleIntentFromContext = hasAny(contextText, [
        /\b(who|whose|team|lead|leader|head|manager|contact|email|person|people(?!'s)|employee|employees|owner|owns|role|reports to|reporting)\b/,
        /\b(directory)\b/
      ]);
      const metadataIntentFromContext = hasAny(contextText, [
        /\b(bu|business unit|department|sub department|client|region|geography|therapy area|disease area|document type|author|published|modified|version)\b/,
        /\b(list|show|find|which|all)\b.*\b(documents|files|policies|sops|certificates|decks)\b/
      ]);
      const explicitDocumentRequest = isExplicitDocumentRequest(question);
      const explicitMetadataRequest = isExplicitMetadataListingRequest(question);
      const explicitPeopleRequest = isExplicitPeopleFollowUp(question);
      const followUpPeopleIntent = !explicitDocumentRequest && (explicitPeopleRequest || peopleIntentFromContext);
      const followUpMetadataIntent = explicitDocumentRequest
        ? explicitMetadataRequest
        : metadataIntentFromContext;

      return {
        kind: followUpMetadataIntent ? 'metadata_listing' : 'document',
        shouldRetrieveDocuments: true,
        shouldRetrievePeople: CHAT_PEOPLE_DIRECTORY_ENABLED && followUpPeopleIntent,
        shouldUseDocumentMetadata: followUpMetadataIntent,
        retrievalQuery: normalizeRetrievalText(contextualQuestion),
        conversationContext
      };
    }

    const explicitPeopleRequestWithoutContext = isExplicitPeopleFollowUp(question) || isLikelyPersonNameQuery(question);
    const explicitDocumentRequestWithoutContext = isExplicitDocumentRequest(question);
    const explicitMetadataRequestWithoutContext = isExplicitMetadataListingRequest(question);
    if (explicitPeopleRequestWithoutContext
      && !explicitDocumentRequestWithoutContext
      && !explicitMetadataRequestWithoutContext
      && !CHAT_PEOPLE_DIRECTORY_ENABLED) {
      return {
        kind: 'off_topic',
        shouldRetrieveDocuments: false,
        shouldRetrievePeople: false,
        shouldUseDocumentMetadata: false,
        directAnswer: 'I cannot answer employee-directory or personal-profile questions in KM Assistant right now. I can still help search Active Knowledge Hub documents by author name, topic, client, BU, department, therapy area, disease area, or document type.'
      };
    }
    if (explicitPeopleRequestWithoutContext && (explicitDocumentRequestWithoutContext || explicitMetadataRequestWithoutContext)) {
      return {
        kind: explicitMetadataRequestWithoutContext ? 'metadata_listing' : 'document',
        shouldRetrieveDocuments: true,
        shouldRetrievePeople: false,
        shouldUseDocumentMetadata: explicitMetadataRequestWithoutContext,
        retrievalQuery: normalizeRetrievalText(rawQuestion)
      };
    }

    if (isExplicitPeopleFollowUp(question) || isLikelyPersonNameQuery(question)) {
      return {
        kind: 'off_topic',
        shouldRetrieveDocuments: false,
        shouldRetrievePeople: false,
        shouldUseDocumentMetadata: false,
        directAnswer: 'I cannot answer employee-directory or personal-profile questions in KM Assistant right now. If you need documents connected to a person or team, ask for the Knowledge Hub documents by author, topic, BU, or department.'
      };
    }

    const explicitSearchRequestWithoutContext = hasAny(question, [
      /\b(find|search|show|list|give me|look for)\b/,
      /\b(document|documents|doc|docs|file|files|source|sources|policy|sop|certificate|deck|capability|capabilities|assets?|assests?|ppt|pptx|pdf|docx|xlsx|csv|mp4|mp3|media|video|recording|tracker|spreadsheet)\b/
    ]) || likelyKnowledgeDocumentTopic(question);

    if (explicitSearchRequestWithoutContext && tokenCount(question) >= 5) {
      const peopleIntentFromQuestion = (isExplicitPeopleFollowUp(question) || isLikelyPersonNameQuery(question)) && !isExplicitDocumentRequest(question);
      if (peopleIntentFromQuestion && !CHAT_PEOPLE_DIRECTORY_ENABLED) {
        return {
          kind: 'off_topic',
          shouldRetrieveDocuments: false,
          shouldRetrievePeople: false,
          shouldUseDocumentMetadata: false,
          directAnswer: 'I cannot answer employee-directory or personal-profile questions in KM Assistant right now. I can still help search Active Knowledge Hub documents by author name, topic, client, BU, department, therapy area, disease area, or document type.'
        };
      }
      const metadataIntentFromQuestion = isExplicitMetadataListingRequest(question);
      return {
        kind: metadataIntentFromQuestion ? 'metadata_listing' : 'document',
        shouldRetrieveDocuments: true,
        shouldRetrievePeople: CHAT_PEOPLE_DIRECTORY_ENABLED && peopleIntentFromQuestion,
        shouldUseDocumentMetadata: metadataIntentFromQuestion,
        retrievalQuery: cleanDocumentRetrievalQuery(rawQuestion)
      };
    }

    return {
      kind: 'ambiguous_followup',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      conversationContext
    };
  }

  const peopleIntent = hasAny(question, [
    /\b(who|who's|whose|team|lead|leader|head|manager|contact|email|person|people(?!'s)|employee|employees|owner|owns|role|reports to|reporting)\b/,
    /\bwhat\s+(?:he|she|they|[a-z][a-z.'-]+(?:\s+[a-z][a-z.'-]+){1,4})\s+do(?:es)?\b/,
    /\b(directory)\b/
  ]) || isLikelyPersonNameQuery(question);
  const documentIntent = hasAny(question, [
    /\b(document|documents|file|files|policy|sop|certificate|training|procedure|manual|deck|capability|capabilities|assets?|assests?|ppt|pdf|summarize|summary|explain|content|contains|where does it say)\b/,
    /\b(find|search|show|list|give me|related to|about)\b/,
    /\b(what|which)\b.*\b(cover|covers|covered|contain|contains|include|includes|about)\b/,
    /\b(tell me about|describe|overview of|details of|walk me through)\b/,
    /\b(do you have|any|more|other)\b.*\b(document|documents|file|files|policies|sops|decks|capabilities|assets?|assests?|materials)\b/
  ]) || isExplicitDocumentRequest(question) || likelyKnowledgeDocumentTopic(question) || looksLikeNamedKnowledgeAssetQuestion(question);
  const metadataIntent = isExplicitMetadataListingRequest(question);

  if (peopleIntent && !isExplicitDocumentRequest(question) && !metadataIntent && !CHAT_PEOPLE_DIRECTORY_ENABLED) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot answer employee-directory or personal-profile questions in KM Assistant right now. I can still help search Active Knowledge Hub documents by author name, topic, client, BU, department, therapy area, disease area, or document type.'
    };
  }

  if (peopleIntent && (documentIntent || metadataIntent)) {
    return {
      kind: metadataIntent ? 'metadata_listing' : 'document',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: metadataIntent,
      retrievalQuery: cleanDocumentRetrievalQuery(rawQuestion)
    };
  }

  if (peopleIntent) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      directAnswer: 'I cannot answer employee-directory or personal-profile questions in KM Assistant right now. I can still help search Active Knowledge Hub documents by author name, topic, client, BU, department, therapy area, disease area, or document type.'
    };
  }

  if (metadataIntent) {
    return {
      kind: 'metadata_listing',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: true,
      retrievalQuery: cleanDocumentRetrievalQuery(rawQuestion)
    };
  }

  if (documentIntent || pastedDocumentLike) {
    return {
      kind: 'document',
      shouldRetrieveDocuments: true,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      retrievalQuery: pastedDocumentLike ? normalizeRetrievalText(rawQuestion) : cleanDocumentRetrievalQuery(rawQuestion)
    };
  }

  if (isClearlyGeneralOffTopicQuery(question)) {
    return {
      kind: 'off_topic',
      shouldRetrieveDocuments: false,
      shouldRetrievePeople: false,
      shouldUseDocumentMetadata: false,
      conversationContext,
      directAnswer: outOfScopeAnswer(rawQuestion)
    };
  }

  return {
    kind: 'document',
    shouldRetrieveDocuments: true,
    shouldRetrievePeople: false,
    shouldUseDocumentMetadata: false,
    retrievalQuery: cleanDocumentRetrievalQuery(rawQuestion)
  };
};
