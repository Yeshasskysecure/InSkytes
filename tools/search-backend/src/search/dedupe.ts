export interface ChunkSearchHit {
  documentId: string;
  title: string;
  webUrl: string;
  rerankerScore?: number;
  searchScore?: number;
}

export const dedupeChunkHitsByDocument = <T extends ChunkSearchHit>(hits: T[]): T[] => {
  const bestByDocument = new Map<string, T>();

  hits.forEach((hit) => {
    const existing = bestByDocument.get(hit.documentId);
    if (!existing) {
      bestByDocument.set(hit.documentId, hit);
      return;
    }

    const currentScore = hit.rerankerScore ?? hit.searchScore ?? 0;
    const existingScore = existing.rerankerScore ?? existing.searchScore ?? 0;
    if (currentScore > existingScore) {
      bestByDocument.set(hit.documentId, hit);
    }
  });

  return Array.from(bestByDocument.values());
};
