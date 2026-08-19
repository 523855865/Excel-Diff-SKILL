#!/usr/bin/env node
import { once } from 'node:events';
import { realpathSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { CompareError, comparePartitioned } from './compare.js';
import { PartitionError } from './partitions.js';
import { InputError } from './read-xlsx.js';
import { createReportWriter } from './report.js';
import { SpecError, loadSpec } from './spec.js';

class UsageError extends Error {
  constructor(message) {
    super(message);
    this.code = 'USAGE';
  }
}

function parseArgs(args) {
  if (args[0] !== 'compare') {
    throw new UsageError('usage: excel-diff compare --spec <path> [--progress] [--keep-temp]');
  }
  let path;
  let progress = false;
  let keepTemp = false;
  for (let index = 1; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === '--spec' && path === undefined) {
      path = args[index + 1];
      if (!path || path.startsWith('--')) throw new UsageError('usage: excel-diff compare --spec <path> [--progress] [--keep-temp]');
      index += 1;
    } else if (argument === '--progress' && !progress) progress = true;
    else if (argument === '--keep-temp' && !keepTemp) keepTemp = true;
    else throw new UsageError('usage: excel-diff compare --spec <path> [--progress] [--keep-temp]');
  }
  if (path === undefined) throw new UsageError('usage: excel-diff compare --spec <path> [--progress] [--keep-temp]');
  return { path, progress, keepTemp };
}

export function failure(error) {
  const comparisonError = error instanceof InputError || error instanceof CompareError || error instanceof PartitionError;
  const diskFull = error?.code === 'ENOSPC' || error?.code === 'EDQUOT';
  const known = error instanceof UsageError || error instanceof SpecError || comparisonError || diskFull;
  const resourceError = diskFull || (comparisonError && typeof error?.code === 'string'
    && (error.code.endsWith('_LIMIT_EXCEEDED') || error.code === 'HOT_KEY_TOO_LARGE'));
  const exitCode = resourceError ? 5 : comparisonError ? 4 : known ? 2 : 6;
  const output = {
    status: 'FAILED',
    code: diskFull ? 'DISK_FULL' : known ? error.code : 'INTERNAL_ERROR',
    message: diskFull ? 'output storage is full' : known ? error.message : 'unexpected error'
  };
  const tempDirectory = error?.tempDirectory ?? error?.cause?.tempDirectory;
  if (typeof tempDirectory === 'string') output.tempDirectory = tempDirectory;
  if (process.env.EXCEL_DIFF_DEBUG === '1' && error.stack) output.stack = error.stack;
  return { exitCode, output };
}

async function writeJsonLine(stream, value) {
  if (!stream.write(`${JSON.stringify(value)}\n`)) await once(stream, 'drain');
}

export async function main(args, dependencies = {}) {
  const options = parseArgs(args);
  const spec = await loadSpec(options.path);
  const report = await (dependencies.createReportWriter ?? createReportWriter)(spec);
  let latestProgress;
  let lastProgressRows = 0;
  let result;
  const emitProgress = async (progress) => writeJsonLine(process.stderr, {
    bytesWritten: progress.bytesWritten,
    currentFile: progress.currentFile,
    rowsScanned: progress.rowsScanned,
    type: 'PROGRESS'
  });
  try {
    result = await comparePartitioned(spec, report, {
      keepTemp: options.keepTemp,
      onProgress: options.progress ? async (progress) => {
        latestProgress = progress;
        if (progress.rowsScanned - lastProgressRows >= 1000) {
          lastProgressRows = progress.rowsScanned;
          await emitProgress(progress);
        }
      } : undefined
    });
    if (options.progress && latestProgress?.rowsScanned !== lastProgressRows) await emitProgress(latestProgress ?? {
      bytesWritten: 0,
      currentFile: spec.baseline,
      rowsScanned: result.summary.totalRowsScanned
    });
    const completed = await report.complete(result.summary);
    await writeJsonLine(process.stdout, {
      ...completed.summary,
      directory: completed.directory,
      ...(result.tempDirectory ? { tempDirectory: result.tempDirectory } : {})
    });
  } catch (error) {
    if (result?.tempDirectory && error && (typeof error === 'object' || typeof error === 'function')
      && error.tempDirectory == null) {
      try { error.tempDirectory = result.tempDirectory; } catch {}
    }
    await report.abort(error);
    throw error;
  }
}

let invokedDirectly = false;
try {
  invokedDirectly = Boolean(process.argv[1]) && realpathSync(process.argv[1]) === fileURLToPath(import.meta.url);
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
if (invokedDirectly) {
  main(process.argv.slice(2)).catch((error) => {
    const { exitCode, output } = failure(error);
    process.stderr.write(`${JSON.stringify(output)}\n`);
    process.exitCode = exitCode;
  });
}
