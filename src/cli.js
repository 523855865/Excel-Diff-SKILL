#!/usr/bin/env node
import { CompareError, compare } from './compare.js';
import { PartitionError } from './partitions.js';
import { InputError } from './read-xlsx.js';
import { writeReport } from './report.js';
import { SpecError, loadSpec } from './spec.js';

class UsageError extends Error {
  constructor(message) {
    super(message);
    this.code = 'USAGE';
  }
}

function specPath(args) {
  if (args.length !== 3 || args[0] !== 'compare' || args[1] !== '--spec' || args[2] === '') {
    throw new UsageError('usage: excel-diff compare --spec <path>');
  }
  return args[2];
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

async function main(args) {
  const spec = await loadSpec(specPath(args));
  const result = await compare(spec);
  const report = await writeReport(spec, result);
  process.stdout.write(`${JSON.stringify({ ...report.summary, directory: report.directory })}\n`);
}

main(process.argv.slice(2)).catch((error) => {
  const { exitCode, output } = failure(error);
  process.stderr.write(`${JSON.stringify(output)}\n`);
  process.exitCode = exitCode;
});
