import { isDeepStrictEqual } from 'node:util';

export function normalizeValue(value, rule = {}) {
  if (value == null || (rule.emptyEqualsNull && value === '')) return ['blank', null];
  if (value instanceof Date) return ['date', value.toISOString()];
  if (typeof value === 'string') {
    let text = rule.trim ? value.trim() : value;
    if (rule.caseSensitive === false) text = text.toLowerCase();
    return ['string', text];
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError('number must be finite');
    return ['number', Object.is(value, -0) ? 0 : value];
  }
  if (typeof value === 'boolean') return ['boolean', value];
  if (typeof value === 'object' && typeof value.error === 'string') return ['error', value.error];
  if (typeof value === 'object' && Array.isArray(value.richText)) {
    return normalizeValue(value.richText.map(({ text = '' }) => text).join(''), rule);
  }
  if (typeof value === 'object' && typeof value.text === 'string' && typeof value.hyperlink === 'string'
    && (value.tooltip == null || typeof value.tooltip === 'string')) {
    return ['hyperlink', {
      text: normalizeValue(value.text, rule),
      target: value.hyperlink,
      tooltip: value.tooltip ?? null
    }];
  }
  const formula = typeof value === 'object' && (typeof value.formula === 'string'
    ? value.formula
    : typeof value.sharedFormula === 'string' ? `shared:${value.sharedFormula}` : null);
  if (formula !== null) {
    if (rule.formulaMode === 'formula') return ['formula', formula];
    const result = normalizeValue(value.result, rule);
    if (rule.formulaMode === 'cached-result') return result;
    return ['formula', [formula, result]];
  }
  if (typeof value === 'object') throw new TypeError('unsupported cell value');
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

function dateLiteral(value) {
  const dateOnly = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  const datetime = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(Z|[+-](\d{2}):(\d{2}))$/.exec(value);
  const parts = dateOnly || datetime;
  const invalid = () => { throw new TypeError(`invalid ISO date literal: ${value}`); };
  if (!parts) invalid();

  const [year, month, day] = parts.slice(1, 4).map(Number);
  const leapYear = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const days = [31, leapYear ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (month < 1 || month > 12 || day < 1 || day > days[month - 1]) invalid();
  if (datetime) {
    const hour = Number(datetime[4]);
    const minute = Number(datetime[5]);
    const second = datetime[6] === undefined ? 0 : Number(datetime[6]);
    const offsetHour = datetime[9] === undefined ? 0 : Number(datetime[9]);
    const offsetMinute = datetime[10] === undefined ? 0 : Number(datetime[10]);
    if (hour > 23 || minute > 59 || second > 59 || offsetHour > 23 || offsetMinute > 59) invalid();
  }
  const date = new Date(dateOnly ? `${value}T00:00:00.000Z` : value);
  if (Number.isNaN(date.valueOf())) invalid();
  return date.toISOString();
}

function normalizeFilterValue(value, cellType, rule) {
  if (cellType === 'date' && typeof value === 'string') return ['date', dateLiteral(value)];
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
    case 'gt': {
      const operand = literal(values);
      return comparable(cell, operand) && cell[1] > operand[1];
    }
    case 'gte': {
      const operand = literal(values);
      return comparable(cell, operand) && cell[1] >= operand[1];
    }
    case 'lt': {
      const operand = literal(values);
      return comparable(cell, operand) && cell[1] < operand[1];
    }
    case 'lte': {
      const operand = literal(values);
      return comparable(cell, operand) && cell[1] <= operand[1];
    }
    case 'in': return values.some((item) => equalValues(cell, literal(item), rule));
    case 'notIn': return values.every((item) => !equalValues(cell, literal(item), rule));
    case 'contains': {
      const operand = literal(values);
      return cell[0] === 'string' && operand[0] === 'string' && cell[1].includes(operand[1]);
    }
    case 'startsWith': {
      const operand = literal(values);
      return cell[0] === 'string' && operand[0] === 'string' && cell[1].startsWith(operand[1]);
    }
    case 'endsWith': {
      const operand = literal(values);
      return cell[0] === 'string' && operand[0] === 'string' && cell[1].endsWith(operand[1]);
    }
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
