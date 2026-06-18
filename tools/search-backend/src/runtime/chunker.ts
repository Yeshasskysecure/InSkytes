export interface TextChunk {
  text: string;
  ordinal: number;
}

export interface ChunkOptions {
  targetTokens?: number;
  overlapTokens?: number;
  maxChars?: number;
}

const normalizeWhitespace = (value: string): string =>
  String(value || '').replace(/\s+/g, ' ').trim();

const splitLongWord = (word: string, maxChars: number): string[] => {
  if (word.length <= maxChars) return [word];

  const parts: string[] = [];
  for (let index = 0; index < word.length; index += maxChars) {
    parts.push(word.slice(index, index + maxChars));
  }
  return parts;
};

export const chunkText = (text: string, options: ChunkOptions = {}): string[] => {
  const targetTokens = options.targetTokens || 750;
  const overlapTokens = options.overlapTokens || 120;
  const maxChars = Math.max(1000, options.maxChars || 5000);
  const words = normalizeWhitespace(text)
    .split(' ')
    .filter(Boolean)
    .flatMap((word) => splitLongWord(word, maxChars));

  if (words.length === 0) return [];

  const chunks: string[] = [];

  for (let start = 0; start < words.length;) {
    const selected: string[] = [];
    let charCount = 0;
    let index = start;

    while (index < words.length && selected.length < targetTokens) {
      const next = words[index];
      const nextLength = next.length + (selected.length > 0 ? 1 : 0);
      if (selected.length > 0 && charCount + nextLength > maxChars) break;
      selected.push(next);
      charCount += nextLength;
      index += 1;
    }

    if (selected.length === 0) {
      selected.push(words[start]);
      index = start + 1;
    }

    chunks.push(selected.join(' '));
    if (index >= words.length) break;
    start = Math.max(start + 1, index - overlapTokens);
  }

  return chunks;
};
