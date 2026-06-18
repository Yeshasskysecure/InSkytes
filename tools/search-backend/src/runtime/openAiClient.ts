import { requireValue, SearchEnv, toNumber } from './env';

interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface ChatCompletionUsage {
  prompt_tokens?: number;
  completion_tokens?: number;
  total_tokens?: number;
}

export interface ChatCompletionResult {
  content: string;
  deployment: string;
  usage?: ChatCompletionUsage;
}

export interface ChatCompletionOptions {
  maxTokens?: number;
  temperature?: number;
  responseFormat?: Record<string, unknown>;
  parallelToolCalls?: boolean;
}

const parseRetryAfterMs = (headers: Headers, bodyText: string): number | undefined => {
  const retryAfter = headers.get('retry-after');
  const retryAfterMsHeader = Number(headers.get('retry-after-ms'));
  if (Number.isFinite(retryAfterMsHeader) && retryAfterMsHeader > 0) return retryAfterMsHeader;
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds > 0) return seconds * 1000;
    const retryDate = Date.parse(retryAfter);
    if (Number.isFinite(retryDate)) return Math.max(0, retryDate - Date.now());
  }
  const retryAfterText = /retry after\s+(\d+)\s+seconds/i.exec(bodyText);
  return retryAfterText ? Number(retryAfterText[1]) * 1000 : undefined;
};

export const embedTexts = async (env: SearchEnv, inputs: string[]): Promise<number[][]> => {
  if (inputs.length === 0) return [];

  const endpoint = requireValue(env, 'AZURE_OPENAI_ENDPOINT').replace(/\/$/, '');
  const deployment = requireValue(env, 'AZURE_OPENAI_EMBEDDING_DEPLOYMENT');
  const apiVersion = requireValue(env, 'AZURE_OPENAI_API_VERSION');
  const apiKey = requireValue(env, 'AZURE_OPENAI_API_KEY');
  const dimensions = toNumber(env.AZURE_OPENAI_EMBEDDING_DIMENSIONS, 3072);
  const useLegacy = String(env.AZURE_OPENAI_USE_LEGACY_API || 'true').toLowerCase() === 'true';
  const url = useLegacy
    ? `${endpoint}/openai/deployments/${encodeURIComponent(deployment)}/embeddings?api-version=${apiVersion}`
    : `${endpoint}/openai/v1/embeddings`;
  const body = useLegacy
    ? { input: inputs, dimensions }
    : { input: inputs, model: deployment, dimensions };
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const text = await response.text();
    const error = new Error(`Azure OpenAI embeddings failed ${response.status}: ${text.slice(0, 700)}`) as Error & { retryAfterMs?: number; nonRetryable?: boolean };
    error.retryAfterMs = parseRetryAfterMs(response.headers, text);
    error.nonRetryable = response.status === 400 && /maximum input length|invalid 'input/i.test(text);
    throw error;
  }

  const json = await response.json();
  return (json.data || []).map((item: { embedding: number[] }) => item.embedding);
};

export const embedTextsInBatches = async (env: SearchEnv, inputs: string[]): Promise<number[][]> => {
  const batchSize = toNumber(env.AZURE_OPENAI_EMBEDDING_BATCH_SIZE, 100);
  const results: number[][] = [];

  for (let index = 0; index < inputs.length; index += batchSize) {
    const batch = inputs.slice(index, index + batchSize);
    results.push(...await embedTexts(env, batch));
  }

  return results;
};

export const getChatCompletionResult = async (
  env: SearchEnv,
  messages: ChatMessage[],
  options: ChatCompletionOptions = {}
): Promise<ChatCompletionResult> => {
  const endpoint = requireValue(env, 'AZURE_OPENAI_ENDPOINT').replace(/\/$/, '');
  const deployment = requireValue(env, 'AZURE_OPENAI_CHAT_DEPLOYMENT');
  const apiVersion = requireValue(env, 'AZURE_OPENAI_API_VERSION');
  const apiKey = requireValue(env, 'AZURE_OPENAI_API_KEY');
  const useLegacy = String(env.AZURE_OPENAI_USE_LEGACY_API || 'true').toLowerCase() === 'true';
  const useCompletionTokenLimit = String(env.AZURE_OPENAI_USE_MAX_COMPLETION_TOKENS || 'true').toLowerCase() !== 'false';
  const url = useLegacy
    ? `${endpoint}/openai/deployments/${encodeURIComponent(deployment)}/chat/completions?api-version=${apiVersion}`
    : `${endpoint}/openai/v1/chat/completions`;
  const tokenLimit = { [useCompletionTokenLimit ? 'max_completion_tokens' : 'max_tokens']: options.maxTokens ?? 900 };
  const chatOptions: Record<string, unknown> = {
    temperature: options.temperature ?? 0.1,
    ...tokenLimit
  };

  if (options.responseFormat) {
    chatOptions.response_format = options.responseFormat;
  }

  if (typeof options.parallelToolCalls === 'boolean') {
    chatOptions.parallel_tool_calls = options.parallelToolCalls;
  }

  const body = useLegacy
    ? { messages, ...chatOptions }
    : { model: deployment, messages, ...chatOptions };
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const text = await response.text();
    const error = new Error(`Azure OpenAI chat failed ${response.status}: ${text.slice(0, 700)}`) as Error & { retryAfterMs?: number };
    error.retryAfterMs = parseRetryAfterMs(response.headers, text);
    throw error;
  }

  const json = await response.json();
  const content = json.choices && json.choices[0] && json.choices[0].message
    ? String(json.choices[0].message.content || '')
    : '';

  return {
    content,
    deployment,
    usage: json.usage
  };
};

export const getChatCompletion = async (
  env: SearchEnv,
  messages: ChatMessage[],
  options: ChatCompletionOptions = {}
): Promise<string> => {
  const result = await getChatCompletionResult(env, messages, options);
  return result.content;
};
