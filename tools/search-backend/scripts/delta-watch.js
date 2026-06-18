const { spawn } = require('child_process');
const path = require('path');
const { loadEnv, toNumber } = require('./lib/env');

const envPath = process.argv[2] || 'config/search.env';
const intervalSecondsArg = process.argv[3];
const { env, absolutePath } = loadEnv(envPath);
const intervalSeconds = toNumber(intervalSecondsArg || env.SYNC_DELTA_WATCH_INTERVAL_SECONDS, 180);
let stopRequested = false;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const runDeltaCatchup = () => new Promise((resolve) => {
  const scriptPath = path.join('tools', 'search-backend', 'scripts', 'delta-sync.js');
  const child = spawn(process.execPath, [scriptPath, envPath, 'delta', 'all'], {
    stdio: 'inherit',
    shell: false
  });

  child.on('exit', (code) => {
    resolve(code || 0);
  });
});

const requestStop = (signal) => {
  stopRequested = true;
  console.log(`\n${signal} received. Delta watch will stop after the current cycle.`);
};

process.once('SIGINT', () => requestStop('SIGINT'));
process.once('SIGTERM', () => requestStop('SIGTERM'));

const run = async () => {
  console.log(`Using env: ${absolutePath}`);
  console.log(`Delta watch interval: ${intervalSeconds}s`);
  console.log('Press Ctrl+C once to stop after the current cycle.');

  while (!stopRequested) {
    const startedAt = Date.now();
    const code = await runDeltaCatchup();
    const durationMs = Date.now() - startedAt;
    console.log(JSON.stringify({
      at: new Date().toISOString(),
      action: 'delta_watch_cycle_complete',
      exitCode: code,
      durationMs
    }));

    if (stopRequested) break;
    await sleep(intervalSeconds * 1000);
  }
};

run().catch((error) => {
  console.error(error.message || String(error));
  process.exit(1);
});
