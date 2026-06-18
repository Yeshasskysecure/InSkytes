/**
 * Find the best match from allowed values using fuzzy matching
 * Compatible with ES5 (no Array.find or String.includes)
 */
export function findBestMatch(value: string, allowedValues: string[]): string {
  if (!value || value.trim() === '') {
    return '';
  }

  const normalizedValue = value.trim().toLowerCase();
  
  // Exact match (case-insensitive)
  for (let i = 0; i < allowedValues.length; i++) {
    const allowed = allowedValues[i];
    if (allowed.toLowerCase() === normalizedValue) {
      return allowed;
    }
  }

  // Partial match (contains)
  for (let i = 0; i < allowedValues.length; i++) {
    const allowed = allowedValues[i];
    const allowedLower = allowed.toLowerCase();
    if (allowedLower.indexOf(normalizedValue) !== -1 || normalizedValue.indexOf(allowedLower) !== -1) {
      return allowed;
    }
  }

  // Fuzzy match - check if any words match
  const valueWords = normalizedValue.split(/\s+/);
  for (let i = 0; i < allowedValues.length; i++) {
    const allowed = allowedValues[i];
    const allowedWords = allowed.toLowerCase().split(/\s+/);
    
    for (let j = 0; j < valueWords.length; j++) {
      const word = valueWords[j];
      for (let k = 0; k < allowedWords.length; k++) {
        const allowedWord = allowedWords[k];
        if (allowedWord.indexOf(word) !== -1 || word.indexOf(allowedWord) !== -1) {
          return allowed;
        }
      }
    }
  }

  // No match found
  return '';
}

