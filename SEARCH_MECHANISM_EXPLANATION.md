# How Search Works for Different Words/Queries

This document explains how the chatbot processes different types of queries and performs semantic search on SharePoint documents.

## Overall Query Processing Flow

When a user sends a message, the chatbot follows this processing order:

```
User Query
    ↓
1. Quick Links Check (getLinkForQuery)
    ↓ (if not found)
2. Internal Queries Check (time, date, day)
    ↓ (if not found)
3. "All Documents" Query Check
    ↓ (if not found)
4. General Queries Check (hi, help, who are you, etc.)
    ↓ (if not found)
5. Context Map Lookup (getAnswerFromContext)
    ↓
6. Dynamic AI Answer Generation (getDynamicAnswerFromWebsite)
    ↓
7. Out-of-Domain Check (if decline message, STOP HERE)
    ↓ (if in-domain)
8. Semantic Search on KMArtifacts (performSemanticSearchOnKMArtifacts)
    ↓
9. Consolidate Answer + Document Links
    ↓
10. Send to User
```

## 1. Quick Links Check

**Function**: `getLinkForQuery(query: string)`

**Purpose**: Checks if the query matches predefined quick links (external Indegene website links).

**How it works**:
- Matches query against `QUICK_LINKS_MAP` keys
- Returns external website links (e.g., careers, case studies, contact)

**Example**: Query "careers" → Returns careers.indegene.com link

---

## 2. Internal Queries Check

**Function**: `handleInternalQueries(query: string)`

**Purpose**: Handles time, date, and day queries.

**How it works**:
- Checks for keywords: "time", "date", "day", "today", "now"
- Returns current time/date/day information

**Example**: Query "what time is it" → Returns current time

---

## 3. "All Documents" Query Check

**Function**: `getAllDocumentLinks()`

**Purpose**: Fetches ALL documents from KMArtifacts library when user asks for all document links.

**How it works**:
- Detects queries like "all documents", "all document links", "links for all documents"
- Fetches all published documents from SharePoint KMArtifacts library
- Formats them as clickable markdown links

**Example**: Query "give me links for all documents" → Returns list of all document links

---

## 4. General Queries Check

**Function**: `handleGeneralQueries(query: string)`

**Purpose**: Handles small-talk and meta-queries.

**How it works**:
- Matches patterns like "hi", "hello", "who are you", "what can you do", "help", "quick links"
- Returns predefined responses
- **Important**: Has specific regex to avoid catching "what is km" as "what can you do"

**Example**: Query "hi" → Returns greeting message

---

## 5. Context Map Lookup

**Function**: `getAnswerFromContext(query: string)`

**Purpose**: Retrieves predefined answers from `CONTEXT_MAP` for specific queries.

**How it works**:
1. **Normalizes query**: Converts to lowercase, removes extra spaces
2. **Direct match**: Checks if normalized query exists in `CONTEXT_MAP`
3. **Plural/singular variations**: For policy queries, tries both forms
4. **Hard overrides**: Special handling for:
   - "skysecure" keyword
   - CSR policy queries
   - Privacy policy queries
   - Retention/Archival policy queries
5. **Token-based matching**: 
   - Splits query into tokens
   - Scores each `CONTEXT_MAP` key based on:
     - Token overlap
     - Phrase containment
     - All tokens present (subset match)
     - Domain-specific boosts (employees, founders, CEO, headquarters)
6. **Returns best match**: Returns the answer from the highest-scoring key

**Example**: Query "what is poc" → Matches "poc" key in `CONTEXT_MAP` → Returns "Proof of Concept" definition

---

## 6. Dynamic AI Answer Generation

**Function**: `getDynamicAnswerFromWebsite(query: string, contextMapAnswer?: string)`

**Purpose**: Uses Azure OpenAI to generate comprehensive answers.

**How it works**:
1. Takes the query and optional context map answer
2. Builds a system prompt with:
   - Company overview (Indegene information)
   - Internal terminology (POC = Proof of Concept, etc.)
   - Domain boundaries (what to answer vs. decline)
3. Calls Azure OpenAI API to generate answer
4. Returns AI-generated response

**Example**: Query "what is poc" → AI generates comprehensive answer about Proof of Concept

---

## 7. Out-of-Domain Check

**Function**: `isOutOfDomainQuery(query: string)` + decline message detection

**Purpose**: Detects queries outside Indegene/healthcare domain and stops processing.

**How it works**:
- Checks for keywords related to:
  - Weather, sports, entertainment
  - Personal finance, cooking
  - General knowledge (unrelated to healthcare)
  - Political figures (unrelated to healthcare)
  - Philosophical questions
  - Harmful queries
- Also checks if AI answer contains decline phrases
- **If out-of-domain**: Sends decline message and **STOPS** (no semantic search)

**Example**: Query "what is the president of india" → Returns decline message, stops processing

---

## 8. Semantic Search on KMArtifacts

**Function**: `performSemanticSearchOnKMArtifacts(query: string)`

**Purpose**: Searches SharePoint KMArtifacts documents for relevant content.

### Step-by-Step Process:

#### Step 8.1: Fetch Documents from SharePoint
- Fetches all published documents from `KMArtifacts` library
- Gets metadata: Title, Abstract, FileName, FileRef, BusinessUnit, DocumentType, etc.

#### Step 8.2: Extract Document Content
- Processes documents in batches of 10
- For supported file types (PDF, DOCX, PPTX, MHTML, MHT, SVG):
  - Downloads file from SharePoint
  - Parses content using `DocumentParser`
  - Caches parsed content for performance
- Creates array of documents with content: `allDocsWithContent`

#### Step 8.3: Perform Semantic Search
- Calls `openAIServiceRef.current.semanticSearchLocalDocuments(query, allDocsWithContent, 20)`
- This uses OpenAI embeddings to find semantically similar documents
- Returns top 20 results with similarity scores
- **Similarity scores**:
  - Keyword matches: similarity >= 5 (high)
  - Semantic matches: similarity 0.0 - 1.0 (varies)

#### Step 8.4: Filter Results by Query Type

**A. For Queries with Known Terms** (like "poc", "km", "prma", "heor", "cpc", "regulatory", "oncology", "pharmacovigilance"):

1. **Explicit Filename/Title Matching FIRST**:
   - Extracts query terms (e.g., "poc" from "what is poc")
   - Filters semantic results to find documents where:
     - Query term appears in filename (e.g., "POC_Cost_Planning_Update_UMmySheets")
     - Query term appears in title
     - Query term appears in description
   - Sorts by similarity (descending)
   - **Priority**: These filename matches are used FIRST if found

2. **Similarity Filtering**:
   - Filters results with similarity >= 0.3 (standard threshold)
   - If no results, tries lower threshold (>= 0.2) for known terms

3. **Final Selection**:
   - If filename matches found → Use top 10 filename matches
   - Else if similarity matches found → Use similarity-filtered results
   - Else if known term with lower threshold → Use top 5 lower threshold results

**B. For Regular Queries** (no known terms):

1. **Similarity Filtering**:
   - Filters results with similarity >= 0.3
   - Uses these filtered results

2. **Fallback Filename Matching**:
   - If no similarity matches, checks for filename/title matches
   - Extracts query terms
   - Finds documents where terms appear in filename, title, or description
   - Sorts by similarity, takes top 5

#### Step 8.5: Generate Answer Text and Document Links

1. **Answer Text Generation**:
   - **Priority 1**: Use document's Abstract/Description field (if exists and not empty)
   - **Priority 2**: Extract summary from document content (first 1000 chars, find good stopping point)
   - **Priority 3**: Generate generic summary: "This document provides comprehensive information about [query]..."

2. **Document Link Generation**:
   - **Top Document Link**:
     - Gets `serverRelativeUrl` from `fileRef` or `serverRelativeUrl` field
     - Normalizes URL (removes duplicate site paths)
     - For Office files (DOCX, PPTX, XLSX): Uses WopiFrame URL
     - For other files: Uses direct SharePoint URL
     - Formats as markdown: `**[Document Title](URL)**`
   
   - **Additional Document Links** (if more than 1 result):
     - Formats as numbered list with abstracts
     - Format: `2. **[Title](URL)**\n   [Abstract preview]...`

3. **Return Value**:
   - Returns `{ answerText: string, documentLinks: string }`
   - `answerText`: Abstract or summary of top document
   - `documentLinks`: Markdown-formatted clickable links

---

## 9. Consolidation

**Function**: `handleSendMessage()` (consolidation section)

**Purpose**: Combines dynamic AI answer with document links into a single message.

**How it works**:
1. Gets `dynamicAnswer` from AI/Context Map
2. Gets `documentLinks` from semantic search
3. **Unconditionally appends** `documentLinks` to `dynamicAnswer` if links exist
4. Adds proper spacing (`\n\n`) between answer and links
5. Multiple safety checks ensure links are not lost
6. Final check before sending ensures links are present

**Example**:
```
Dynamic Answer: "POC stands for Proof of Concept..."
Document Links: "\n\n**[POC_Cost_Planning_Update_UMmySheets](URL)**\n\n📚 Additional Relevant Documents:\n\n2. **[Another Doc](URL)**..."

Final Answer: "POC stands for Proof of Concept...\n\n**[POC_Cost_Planning_Update_UMmySheets](URL)**\n\n📚 Additional Relevant Documents:\n\n2. **[Another Doc](URL)**..."
```

---

## Key Differences: Known Terms vs. Regular Queries

### Known Terms (poc, km, prma, heor, cpc, regulatory, oncology, pharmacovigilance)

1. **Explicit Filename Matching FIRST**: 
   - Checks filename/title/description for query terms BEFORE similarity filtering
   - Ensures documents like "POC_Cost_Planning_Update_UMmySheets" are found even if semantic similarity is low

2. **Lower Similarity Threshold**:
   - Standard: similarity >= 0.3
   - Known terms fallback: similarity >= 0.2

3. **Priority Order**:
   - Filename matches → Similarity matches → Lower threshold matches

### Regular Queries

1. **Similarity-First Approach**:
   - Filters by similarity >= 0.3 first
   - Only uses filename matching as fallback if no similarity matches

2. **Standard Threshold**:
   - Uses similarity >= 0.3 consistently

---

## Example Flows

### Example 1: "what is poc"

```
1. Quick Links Check → Not found
2. Internal Queries → Not found
3. All Documents → Not found
4. General Queries → Not found
5. Context Map → Found "poc" key → Returns "Proof of Concept" definition
6. AI Answer → Generates comprehensive answer about POC
7. Out-of-Domain Check → Not out of domain
8. Semantic Search:
   - Fetches all documents
   - Performs semantic search (top 20 results)
   - Detects "poc" as known term
   - Does explicit filename matching FIRST
   - Finds "POC_Cost_Planning_Update_UMmySheets" (filename contains "poc")
   - Filters by similarity (if needed, uses >= 0.2 threshold)
   - Generates document links
9. Consolidation → Appends document links to AI answer
10. Sends to user
```

### Example 2: "what is breast cancer"

```
1-4. Same checks → Not found
5. Context Map → Not found (no "breast cancer" key)
6. AI Answer → Generates comprehensive answer about breast cancer
7. Out-of-Domain Check → Not out of domain (healthcare topic)
8. Semantic Search:
   - Fetches all documents
   - Performs semantic search (top 20 results)
   - No known terms detected
   - Filters by similarity >= 0.3
   - Finds documents with high semantic similarity to "breast cancer"
   - Generates document links
9. Consolidation → Appends document links to AI answer
10. Sends to user
```

### Example 3: "what is the president of india"

```
1-4. Same checks → Not found
5. Context Map → Not found
6. AI Answer → Generates answer (may include decline message)
7. Out-of-Domain Check → DETECTED as out of domain
   - Query contains "president of india" (political figure, unrelated to healthcare)
   - OR AI answer contains decline phrases
   - **STOPS HERE** → Sends decline message, NO semantic search
```

---

## Debugging Tips

The code includes extensive logging. Check browser console for:

- `=== SEMANTIC SEARCH RESULT (IMMEDIATE) ===`: Shows document links immediately after semantic search
- `=== FINAL FILTERED RESULTS ===`: Shows which documents passed filtering
- `=== DOCUMENT LINKS CHECK ===`: Shows if document links exist before appending
- `=== FINAL ABSOLUTE CHECK BEFORE SENDING ===`: Final check before sending message
- `=== MESSAGE BEING SENT TO CHAT ===`: Shows final message content

These logs help trace where document links might be lost in the process.

