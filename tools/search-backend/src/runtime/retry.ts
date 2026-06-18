export interface RetryOptions {
  retries?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  label?: string;
}

export const sleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms));

export const withRetry = async <T>(
  operation: () => Promise<T>,
  options: RetryOptions = {}
): Promise<T> => {
  const retries = Number.isFinite(Number(options.retries)) ? Number(options.retries) : 3;
  const baseDelayMs = Number.isFinite(Number(options.baseDelayMs)) ? Number(options.baseDelayMs) : 1500;
  const maxDelayMs = Number.isFinite(Number(options.maxDelayMs)) ? Number(options.maxDelayMs) : 120000;
  const label = options.label || 'operation';
  let lastError: unknown;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      if ((error as { nonRetryable?: boolean } | undefined)?.nonRetryable) break;
      if (attempt >= retries) break;
      const retryAfterMs = Number((error as { retryAfterMs?: number } | undefined)?.retryAfterMs);
      const exponentialDelay = Math.min(baseDelayMs * Math.pow(2, attempt), maxDelayMs);
      const jitterMs = Math.floor(Math.random() * Math.min(1000, Math.max(100, exponentialDelay * 0.1)));
      const delay = Number.isFinite(retryAfterMs) && retryAfterMs > 0
        ? retryAfterMs
        : exponentialDelay + jitterMs;
      const message = error instanceof Error ? error.message : String(error);
      console.warn(`${label} failed on attempt ${attempt + 1}/${retries + 1}: ${message}. Retrying in ${delay}ms.`);
      await sleep(delay);
    }
  }

  throw lastError;
};
