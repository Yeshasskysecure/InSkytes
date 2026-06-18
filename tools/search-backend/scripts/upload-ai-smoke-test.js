const fs = require('fs');
const path = require('path');
const { loadEnv } = require('./lib/env');
const { extractDocumentText } = require('./lib/extractors');
const {
  analyzeMediaTranscript,
  extractUploadMetadata,
  transcribeUploadedMedia
} = require('../dist/api/uploadAiApi');

const envPath = process.argv[2] || 'config/search.env';
const reportPath = process.argv[3] || 'C:\\tmp\\iknowledge-upload-ai-smoke-report.json';

const mediaFiles = [
  'C:\\Users\\shweta.kumari\\Downloads\\CultHealth.mp4',
  'C:\\Users\\shweta.kumari\\Downloads\\Rajesh_Nair (1).mp4',
  'C:\\Users\\shweta.kumari\\Downloads\\20200626113327-747_ABInBevCEOonAdaptingintheFaceofCrisis.mp3',
  'C:\\Users\\shweta.kumari\\Downloads\\Demo - HR agent-20260226_154541-Meeting Recording.mp4'
];

const documentFiles = [
  'C:\\Users\\shweta.kumari\\Downloads\\AI_Meeting_Platform_Feasibility_Reference.docx',
  'C:\\Users\\shweta.kumari\\Downloads\\Agentic_AI_Holy_Book.pdf',
  'C:\\Users\\shweta.kumari\\Downloads\\Payment_Tracking_Agent_Approach_Plan.pdf'
];

const mimeTypeFor = (filePath) => {
  const extension = path.extname(filePath).toLowerCase();
  if (extension === '.pdf') return 'application/pdf';
  if (extension === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (extension === '.mp3') return 'audio/mpeg';
  if (extension === '.mp4') return 'video/mp4';
  return 'application/octet-stream';
};

const safeError = (error) => ({
  message: error && error.message ? error.message : String(error),
  name: error && error.name ? error.name : 'Error'
});

const previewText = (value, max = 350) =>
  String(value || '').replace(/\s+/g, ' ').trim().slice(0, max);

const summarizeMetadata = (metadata) => ({
  title: metadata.title || '',
  description: metadata.description || metadata.abstract || '',
  documentType: metadata.documentType || '',
  bu: metadata.bu || metadata.businessUnit || '',
  department: metadata.department || '',
  client: metadata.client || '',
  geography: metadata.geography || metadata.region || '',
  therapyArea: metadata.therapyArea || '',
  diseaseArea: metadata.diseaseArea || ''
});

const logEvents = [];
const context = {
  correlationId: `upload-ai-smoke-${Date.now()}`,
  userHash: 'local-smoke-test',
  log: (event) => logEvents.push({
    ...event,
    at: new Date().toISOString()
  })
};

const runDocumentTest = async (env, filePath) => {
  const startedAt = Date.now();
  const fileName = path.basename(filePath);
  try {
    if (!fs.existsSync(filePath)) {
      return { fileName, path: filePath, ok: false, stage: 'exists', error: { message: 'File not found.' } };
    }

    const buffer = fs.readFileSync(filePath);
    const extraction = await extractDocumentText(env, fileName, mimeTypeFor(filePath), buffer);
    if (!extraction.text || extraction.text.trim().length < 50) {
      return {
        fileName,
        path: filePath,
        ok: false,
        stage: 'extract',
        method: extraction.method,
        reason: extraction.reason || 'No meaningful text extracted.',
        durationMs: Date.now() - startedAt
      };
    }

    const metadata = await extractUploadMetadata(env, {
      documentText: extraction.text,
      docTypeTermsSection: ''
    }, context);

    return {
      fileName,
      path: filePath,
      ok: true,
      stage: 'complete',
      sizeBytes: buffer.length,
      extractionMethod: extraction.method,
      extractedTextChars: extraction.text.length,
      extractedTextPreview: previewText(extraction.text),
      metadata: summarizeMetadata(metadata),
      durationMs: Date.now() - startedAt
    };
  } catch (error) {
    return {
      fileName,
      path: filePath,
      ok: false,
      stage: 'exception',
      error: safeError(error),
      durationMs: Date.now() - startedAt
    };
  }
};

const runMediaTest = async (env, filePath, maxUploadBytes) => {
  const startedAt = Date.now();
  const fileName = path.basename(filePath);
  try {
    if (!fs.existsSync(filePath)) {
      return { fileName, path: filePath, ok: false, stage: 'exists', error: { message: 'File not found.' } };
    }

    const stat = fs.statSync(filePath);
    if (stat.size > maxUploadBytes) {
      return {
        fileName,
        path: filePath,
        ok: false,
        skipped: true,
        stage: 'size-gate',
        sizeBytes: stat.size,
        maxUploadBytes,
        reason: 'File is larger than the direct backend Whisper request cap. In the browser upload flow it must use client-side audio chunking before calling this backend endpoint.'
      };
    }

    const buffer = fs.readFileSync(filePath);
    const transcription = await transcribeUploadedMedia(env, {
      buffer,
      filename: fileName,
      contentType: mimeTypeFor(filePath)
    }, context);

    const analysis = await analyzeMediaTranscript(env, {
      fileName,
      transcriptText: transcription.transcriptText,
      docTypeTermsSection: ''
    }, context);

    return {
      fileName,
      path: filePath,
      ok: true,
      stage: 'complete',
      sizeBytes: buffer.length,
      transcriptChars: transcription.transcriptText.length,
      segmentCount: transcription.transcript.length,
      transcriptPreview: previewText(transcription.transcriptText),
      metadata: summarizeMetadata(analysis),
      durationMs: Date.now() - startedAt
    };
  } catch (error) {
    return {
      fileName,
      path: filePath,
      ok: false,
      stage: 'exception',
      error: safeError(error),
      durationMs: Date.now() - startedAt
    };
  }
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const maxUploadBytes = Number(env.SEARCH_UPLOAD_MAX_MEDIA_BYTES || 25 * 1024 * 1024);
  const startedAt = new Date().toISOString();

  const report = {
    startedAt,
    envPath: absolutePath,
    reportPath,
    maxUploadBytes,
    documents: [],
    media: [],
    events: logEvents,
    summary: {
      documentsPassed: 0,
      documentsFailed: 0,
      mediaPassed: 0,
      mediaFailed: 0,
      mediaSkipped: 0
    }
  };

  for (const filePath of documentFiles) {
    console.log(`Testing document: ${filePath}`);
    const result = await runDocumentTest(env, filePath);
    report.documents.push(result);
    if (result.ok) report.summary.documentsPassed += 1;
    else report.summary.documentsFailed += 1;
    console.log(`  ${result.ok ? 'PASS' : 'FAIL'} ${result.fileName} (${result.stage})`);
  }

  for (const filePath of mediaFiles) {
    console.log(`Testing media: ${filePath}`);
    const result = await runMediaTest(env, filePath, maxUploadBytes);
    report.media.push(result);
    if (result.ok) report.summary.mediaPassed += 1;
    else if (result.skipped) report.summary.mediaSkipped += 1;
    else report.summary.mediaFailed += 1;
    console.log(`  ${result.ok ? 'PASS' : result.skipped ? 'SKIP' : 'FAIL'} ${result.fileName} (${result.stage})`);
  }

  report.completedAt = new Date().toISOString();
  fs.mkdirSync(path.dirname(reportPath), { recursive: true });
  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2), 'utf8');
  console.log(`Report written: ${reportPath}`);
  console.log(JSON.stringify(report.summary, null, 2));

  if (report.summary.documentsFailed > 0 || report.summary.mediaFailed > 0) {
    process.exitCode = 1;
  }
};

run().catch((error) => {
  console.error(error.message || String(error));
  process.exit(1);
});
