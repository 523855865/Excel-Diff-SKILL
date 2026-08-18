import Ajv2020 from 'ajv/dist/2020.js';
import { readFile } from 'node:fs/promises';
import { dirname, extname, resolve } from 'node:path';

const schema = JSON.parse(await readFile(new URL('../schemas/compare-spec.schema.json', import.meta.url), 'utf8'));
const validate = new Ajv2020({ allErrors: true, useDefaults: true }).compile(schema);

export class SpecError extends Error {
  constructor(message) {
    super(message);
    this.name = 'SpecError';
    this.code = 'SPEC_INVALID';
  }
}

function invalid(message) {
  return new SpecError(message);
}

function validateFilters(filters) {
  for (const [index, filter] of filters.entries()) {
    if (filter.operator === 'isNull' || filter.operator === 'isNotNull') {
      if (Object.hasOwn(filter, 'value') || Object.hasOwn(filter, 'values')) {
        throw invalid(`filters/${index}: ${filter.operator} must not include value or values`);
      }
      continue;
    }
    const values = filter.values ?? filter.value;
    if (filter.operator === 'in' || filter.operator === 'notIn') {
      if (!Array.isArray(values) || values.length === 0) throw invalid(`filters/${index}: ${filter.operator} requires at least one value`);
      continue;
    }
    if (filter.operator === 'between') {
      if (!Array.isArray(values) || values.length !== 2) throw invalid(`filters/${index}: between requires exactly two values`);
      continue;
    }
    if (!Object.hasOwn(filter, 'value')) throw invalid(`filters/${index}: ${filter.operator} requires value`);
  }
}

export async function loadSpec(specPath) {
  const absoluteSpecPath = resolve(specPath);
  let raw;
  try {
    raw = JSON.parse(await readFile(absoluteSpecPath, 'utf8'));
  } catch (error) {
    throw invalid(`cannot read or parse spec: ${error.message}`);
  }

  const spec = structuredClone(raw);
  if (!validate(spec)) {
    throw invalid(validate.errors.map(({ instancePath, message }) => `${instancePath || '/'} ${message}`).join('; '));
  }

  const specDirectory = dirname(absoluteSpecPath);
  const ids = new Set();
  const paths = new Set();
  for (const file of spec.files) {
    if (ids.has(file.id)) throw invalid(`duplicate file ID: ${file.id}`);
    ids.add(file.id);
    file.path = resolve(specDirectory, file.path);
    if (paths.has(file.path)) throw invalid(`duplicate input path: ${file.path}`);
    paths.add(file.path);
    if (extname(file.path).toLowerCase() !== '.xlsx') throw invalid(`input file must use .xlsx: ${file.path}`);
  }

  if (!ids.has(spec.baseline)) throw invalid(`baseline does not reference a file ID: ${spec.baseline}`);

  spec.output.directory = resolve(specDirectory, spec.output.directory);
  if (paths.has(spec.output.directory)) throw invalid(`output directory must not equal an input file: ${spec.output.directory}`);

  validateFilters(spec.filters);
  return spec;
}
