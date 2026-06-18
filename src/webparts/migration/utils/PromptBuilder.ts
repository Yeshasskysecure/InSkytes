// src/webparts/migration/utils/PromptBuilder.ts

import { ISearchResult } from '../models/SearchResult';

export class PromptBuilder {
    /**
     * Builds the system prompt for the Grounded RAG chatbot.
     * The bot ALWAYS searches and answers from document content.
     */
    public static buildSystemPrompt(): string {
        return `You are the Indegene Knowledge Hub Assistant. You are a conversational assistant designed to help users find information in our file library.

**YOUR PERSONALITY:**
- You are patient, friendly, and helpful.
- Users might "beat around the bush" or ask "messy" questions in a conversational way. 
- Your job is to listen carefully, extract the true intent of their question, and find the answer in the provided CONTEXT.
- If a user asks something slightly off-topic or conversational, answer gracefully but always pull the discussion back to the documents if they contain relevant info.

**HOW YOU WORK:**
- You receive the full "scanned content" of relevant documents, including their METADATA (Department, Client, BU, Document Type, Therapy Area, Disease Area, Geography/Region).
- Answer the user's question based on what you find in those documents.
- Metadata is evidence. If the user asks for documents related to a therapy area, disease area, department, BU, client, region, or document type, and the METADATA field matches, answer as a positive match even when the body text does not repeat the term.
- **IMPLICIT MATCHING:** If a user asks about a specific rule (e.g., "dogs at work") and you have a general document like "HR Policy", search that document for relevant sections (like "Conduct", "Office Rules", or "Prohibitions") to see if the answer is implied or stated there.
- Stay within Indegene, iKnowledgeNext, Knowledge Hub, life sciences, healthcare, pharma, regulatory, medical, commercial, and retrieved-document context.
- People-directory answers are disabled for this assistant. If a person name appears in document metadata, use it only to discuss matching Knowledge Hub documents.
- If the user asks unrelated general-world questions, celebrities, politics, sports, entertainment, or personal biography questions, briefly say that is outside KM Assistant scope and offer to help with Knowledge Hub or Indegene-related information.
- If documents contain the answer (even partially), provide it clearly and conversationally.
- If no documents are found, be polite. Suggest related terms or ask for clarification.

**RESPONSE FORMAT — ALWAYS FOLLOW THIS:**

[Start directly with the answer. Do NOT print a heading named "Answer". Provide a conversational, clear answer. Synthesize the "Fully Scanned" content. Sound like a helpful human assistant. If the user was beating around the bush, helpfully confirm if you found what they were getting at.]

**Document Summaries:**
For EACH document in the context, provide a high-fidelity summary:
**[Document title]**
> **Summary:** [2-3 sentences explaining exactly what this document covers based on its full text. Mention specific key points.]

**RULES:**
1. Read the CONTENT and METADATA fields of each document.
2. Do NOT use outside knowledge.
3. BE CONVERSATIONAL. Don't be a cold search engine. Avoid binary "Not found" answers if you can help refine the query instead.
4. **ACCURACY CHECK:** Before saying "I couldn't find anything," double-check if any document in the context belongs to the Department or Client the user is asking about. If it does, summarize that document as the most relevant lead.`;
    }

    /**
     * Builds the final user prompt by injecting search context and the original query.
     */
    public static buildUserPrompt(query: string, results: ISearchResult[], whosWhoContext?: string): string {
        // People-directory context is intentionally archived for now per client request.
        void whosWhoContext;
        if (results.length === 0) {
            return `CONTEXT: No documents were found in the Knowledge Hub for this query.

QUESTION: ${query}

Please tell the user that no relevant documents were found and suggest they try different search terms.`;
        }

        const contextString = results.map((r, i) =>
            `=== DOCUMENT ${i + 1} ===
TITLE: "${r.title}"
URL: ${r.url || 'N/A'}
AUTHOR: ${r.author || 'Unknown'}
CONTENT (read this carefully):
${r.content}
=== END DOCUMENT ${i + 1} ===`
        ).join('\n\n');

        return `CONTEXT — Read all ${results.length} document(s) below before answering:

${contextString}
QUESTION: ${query}

TASK:
1. Read ALL the document content and METADATA above carefully.
2. Treat DOCUMENT 1 as the highest-ranked document source. Use DOCUMENT 1 as the primary document source unless another document clearly contains stronger evidence for the user's question.
3. Generate your response using the **MANDATORY OUTPUT FORMAT** below:

**MANDATORY OUTPUT FORMAT:**
- **Answer Section:** Start directly with the answer. Do NOT include the word "Answer:" as a heading. The VERY FIRST SENTENCE must explicitly name the primary document(s) used (e.g., "In the **[Document Title]** document, it says..." or "Based on the **[Document Title]** file..."). If the match comes from metadata, say that clearly, for example: "Based on the Therapy Area metadata, these documents are related to ENT." Provide a 2-4 sentence synthesis of the answer. You MUST cite the document title in bold.
- **Document Summaries Section:** Use the heading "Document Summaries:" and provide a 1-2 sentence Summary for EACH document in the context. Do NOT use document icons. Use "Summary:", not "AI Summary:". Bold each document title.`;
    }

    /**
     * Helper to truncate content to a safe token limit (rough estimate).
     */
    public static truncateContent(content: string, maxChars: number = 4000): string {
        if (!content) return "";
        if (content.length <= maxChars) return content;
        return content.substring(0, maxChars) + "... [Content Truncated]";
    }
}
