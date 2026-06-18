import fs = require('fs');
import path = require('path');

export type SearchEnv = Record<string, string>;

export interface LoadedEnv {
  env: SearchEnv;
  absolutePath: string;
}

export const parseEnvContent = (content: string): SearchEnv => {
  const values: SearchEnv = {};
  content.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) return;
    const index = trimmed.indexOf('=');
    if (index === -1) return;
    values[trimmed.slice(0, index).trim()] = trimmed.slice(index + 1).trim();
  });
  return values;
};

export const loadEnv = (envPath?: string): LoadedEnv => {
  if (String(envPath || '').toLowerCase() === 'process') {
    const processEnv = (process as unknown as { env: Record<string, string | undefined> }).env;
    return {
      env: Object.fromEntries(
        Object.entries(processEnv)
          .filter((entry): entry is [string, string] => typeof entry[1] === 'string')
      ),
      absolutePath: 'process.env'
    };
  }

  const absolutePath = path.resolve(envPath || path.join('config', 'search.env'));
  if (!fs.existsSync(absolutePath)) {
    throw new Error(`Env file not found: ${absolutePath}`);
  }

  return {
    env: parseEnvContent(fs.readFileSync(absolutePath, 'utf8')),
    absolutePath
  };
};

export const requireValue = (env: SearchEnv, key: string): string => {
  const value = env[key];
  if (!value) {
    throw new Error(`Missing required env value: ${key}`);
  }
  return value;
};

export const toNumber = (value: unknown, fallback: number): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
};

export const safeSearchKey = (key: unknown): string =>
  String(key || '').replace(/[^A-Za-z0-9_\-=]/g, '_');
