import { isDeepStrictEqual } from 'node:util';

export function normalizeValue(value, rule = {}) {
  if (value == null || (rule.emptyEqualsNull && value === '')) return ['blank', null];
  if (value instanceof Date) return ['date', value.toISOString()];
  if (typeof value === 'string') {
    let text = rule.trim ? value.trim() : value;
    if (rule.caseSensitive === false) text = text.toLowerCase();
    return ['string', text];
  }
  if (typeof value === 'number') return ['number', value];
  if (typeof value === 'boolean') return ['boolean', value];
  if (typeof value === 'object' && typeof value.error === 'string') return ['error', value.error];
  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return normalizeValue(value.richText.map(({ text = '' }) => text).join(''), rule);
  }
  if (typeof value === 'object' && typeof value.formula === 'string') {
    if (rule.formulaMode === 'formula') return ['formula', value.formula];
    const result = normalizeValue(value.result, rule);
    if (rule.formulaMode === 'cached-result') return result;
    return ['formula', [value.formula, result]];
  }
  return ['string', String(value)];
}

export function equalValues(left, right, rule = {}) {
  if (left[0] === 'number' && right[0] === 'number' && Object.hasOwn(rule, 'numericTolerance')) {
    return Math.abs(left[1] - right[1]) <= rule.numericTolerance;
  }
  return isDeepStrictEqual(left, right);
}

export function encodeKey(values) {
  return JSON.stringify(values);
}

function normalizeFilterValue(value, cellType, rule) {
  if (cellType === 'date' && typeof value === 'string') {
    const date = new Date(value);
    if (!Number.isNaN(date.valueOf())) return normalizeValue(date, rule);
  }
  return normalizeValue(value, rule);
}

function comparable(left, right) {
  return left[0] === right[0] && ['number', 'string', 'date'].includes(left[0]);
}

export function matchesFilter(value, filter, rule = {}) {
  const cell = value;
  const literal = (item) => normalizeFilterValue(item, cell[0], rule);
  const values = Object.hasOwn(filter, 'value') ? filter.value : filter.values;

  switch (filter.operator) {
    case 'eq': return equalValues(cell, literal(values), rule);
    case 'ne': return !equalValues(cell, literal(values), rule);
    case 'gt': return comparable(cell, literal(values)) && cell[1] > literal(values)[1];
    case 'gte': return comparable(cell, literal(values)) && cell[1] >= literal(values)[1];
    case 'lt': return comparable(cell, literal(values)) && cell[1] < literal(values)[1];
    case 'lte': return comparable(cell, literal(values)) && cell[1] <= literal(values)[1];
    case 'in': return values.some((item) => equalValues(cell, literal(item), rule));
    case 'notIn': return values.every((item) => !equalValues(cell, literal(item), rule));
    case 'contains': return cell[0] === 'string' && literal(values)[0] === 'string' && cell[1].includes(literal(values)[1]);
    case 'startsWith': return cell[0] === 'string' && literal(values)[0] === 'string' && cell[1].startsWith(literal(values)[1]);
    case 'endsWith': return cell[0] === 'string' && literal(values)[0] === 'string' && cell[1].endsWith(literal(values)[1]);
    case 'isNull': return cell[0] === 'blank';
    case 'isNotNull': return cell[0] !== 'blank';
    case 'between': {
      const [lower, upper] = values.map(literal);
      return comparable(cell, lower) && comparable(cell, upper) && lower[1] <= cell[1] && cell[1] <= upper[1];
    }
    default: return false;
  }
}

export function printValue([type, value]) {
  if (type === 'blank') return '';
  return type === 'formula' || (value !== null && typeof value === 'object') ? JSON.stringify(value) : String(value);
}
