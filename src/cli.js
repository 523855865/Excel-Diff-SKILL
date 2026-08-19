#!/usr/bin/env node
import { once } from 'node:events';

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

function failure(error) {
  const comparisonError = error instanceof InputError || error instanceof CompareError || error instanceof PartitionError;
  const known = error instanceof UsageError || error instanceof SpecError || comparisonError;
  const exitCode = comparisonError ? 4 : known ? 2 : 6;
  const output = {
    status: 'FAILED',
    code: known ? error.code : 'INTERNAL_ERROR',
    message: known ? error.message : 'unexpected error'
  };
  if (process.env.EXCEL_DIFF_DEBUG === '1' && error.stack) output.stack = error.stack;
  return { exitCode, output };
}

async function writeJsonLine(stream, value) {
  if (!stream.write(`${JSON.stringify(value)}\n`)) await once(stream, 'drain');
}

async function main(args) {
  const options = parseArgs(args);
  const spec = await loadSpec(options.path);
  const report = await createReportWriter(spec);
  let latestProgress;
  let lastProgressRows = 0;
  const emitProgress = async (progress) => writeJsonLine(process.stderr, {
    bytesWritten: progress.bytesWritten,
    currentFile: progress.currentFile,
    rowsScanned: progress.rowsScanned,
    type: 'PROGRESS'
  });
  try {
    const result = await comparePartitioned(spec, report, {
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
    await report.abort(error);
    throw error;
  }
}

main(process.argv.slice(2)).catch((error) => {
  const { exitCode, output } = failure(error);
  process.stderr.write(`${JSON.stringify(output)}\n`);
  process.exitCode = exitCode;
});
