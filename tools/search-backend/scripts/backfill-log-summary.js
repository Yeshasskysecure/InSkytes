const fs = require('fs');
const path = require('path');

const inputPath = process.argv[2];

if (!inputPath) {
  console.error('Usage: node tools/search-backend/scripts/backfill-log-summary.js <run-log.ndjson>');
  process.exit(1);
}

const absolutePath = path.resolve(inputPath);
if (!fs.existsSync(absolutePath)) {
  console.error(`Run log not found: ${absolutePath}`);
  process.exit(1);
}

const rows = fs.readFileSync(absolutePath, 'utf8')
  .split(/\r?\n/)
  .filter(Boolean)
  .map((line) => {
    try {
      return JSON.parse(line);
    } catch {
      return { action: 'unparseable', raw: line };
    }
  });

const byAction = {};
const byMethod = {};
const metadataOnlyReasons = {};
const failures = [];
const chunksByFile = [];
const durations = [];
let firstAt = '';
let lastAt = '';

rows.forEach((row) => {
  byAction[row.action || '(missing)'] = (byAction[row.action || '(missing)'] || 0) + 1;
  if (row.at) {
    firstAt = firstAt || row.at;
    lastAt = row.at;
  }
  if (row.method) byMethod[row.method] = (byMethod[row.method] || 0) + 1;
  if (row.action === 'metadata_only') {
    metadataOnlyReasons[row.reason || '(missing)'] = (metadataOnlyReasons[row.reason || '(missing)'] || 0) + 1;
  }
  if (row.action === 'failed') failures.push(row);
  if (row.action === 'indexed') {
    chunksByFile.push({
      fileName: row.fileName,
      chunks: row.chunks || 0,
      durationMs: row.durationMs || 0
    });
  }
  if (Number.isFinite(Number(row.durationMs))) durations.push(Number(row.durationMs));
});

chunksByFile.sort((a, b) => b.chunks - a.chunks);
durations.sort((a, b) => a - b);

const percentile = (values, p) => {
  if (values.length === 0) return 0;
  const index = Math.min(values.length - 1, Math.max(0, Math.ceil((p / 100) * values.length) - 1));
  return values[index];
};

const output = {
  logPath: absolutePath,
  firstAt,
  lastAt,
  totalEvents: rows.length,
  byAction,
  extractionMethods: byMethod,
  metadataOnlyReasons,
  failures: failures.map((row) => ({
    fileName: row.fileName,
    reason: row.reason
  })),
  indexedFiles: chunksByFile.length,
  topChunkedFiles: chunksByFile.slice(0, 10),
  durationMs: {
    p50: percentile(durations, 50),
    p90: percentile(durations, 90),
    p95: percentile(durations, 95),
    max: durations[durations.length - 1] || 0
  }
};

console.log(JSON.stringify(output, null, 2));
