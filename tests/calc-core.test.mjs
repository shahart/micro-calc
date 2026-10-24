import test from 'node:test';
import assert from 'node:assert/strict';
import { Calculator, displayValue, parseFormula, parseRawConstant } from '../docs/dist/calc-core.js';
import { handleGridEntryKey, parseSheet, serializeSheet, shiftColumns, shiftRows } from '../docs/dist/sheet-ops.js';

function calc(entries, rows = 32, cols = 8) {
  const cells = Object.fromEntries(Object.entries(entries).map(([key, input]) => [key, { input }]));
  return new Calculator({ rows, cols, cells });
}

function numeric(result) {
  assert.ok(['number', 'long'].includes(result.type), `expected numeric value, got ${result.type}`);
  return result.value;
}

test('parser respects arithmetic precedence and right-associative powers', () => {
  const calculator = calc({ '0,0': '=2+3*4', '0,1': '=2^3^2', '0,2': '=-(2+3)*4' });
  assert.equal(numeric(calculator.valueOfCell(0, 0)), 14);
  assert.equal(numeric(calculator.valueOfCell(0, 1)), 512);
  assert.equal(numeric(calculator.valueOfCell(0, 2)), -20);
});

test('references, ranges, and absolute references evaluate against sheet cells', () => {
  const calculator = calc({
    '0,0': '3', '1,0': '5', '2,0': '7',
    '0,1': '=SUM(A1:A3)', '0,2': '=$A$1*2'
  });
  assert.equal(numeric(calculator.valueOfCell(0, 1)), 15);
  assert.equal(numeric(calculator.valueOfCell(0, 2)), 6);
});

test('aggregates count numeric cells and calculate average, sample deviation, and sum of squares', () => {
  const calculator = calc({
    '0,0': '2', '1,0': '4', '2,0': '6', '3,0': 'word',
    '0,1': '=COUNT(A1:A4)', '1,1': '=AVERAGE(A1:A3)',
    '2,1': '=STDEV(A1:A3)', '3,1': '=SUMSQ(A1:A3)'
  });
  assert.equal(numeric(calculator.valueOfCell(0, 1)), 3);
  assert.equal(numeric(calculator.valueOfCell(1, 1)), 4);
  assert.ok(Math.abs(numeric(calculator.valueOfCell(2, 1)) - 2) < 1e-12);
  assert.equal(numeric(calculator.valueOfCell(3, 1)), 56);
});

test('PMT uses MicroCalc argument order and returns the periodic payment', () => {
  const rate = 0.041 / 12;
  const periods = 240;
  const principal = 130000;
  const expected = principal / ((1 - 1 / (1 + rate) ** periods) / rate);
  const calculator = calc({ '0,0': `=PMT(${rate},${periods},${principal})` });
  assert.ok(Math.abs(numeric(calculator.valueOfCell(0, 0)) - expected) < 1e-8);
});

test('logical functions and string concatenation preserve value types', () => {
  const calculator = calc({
    '0,0': '=IF(AND(2<3,OR(FALSE,TRUE)),"ok","no")',
    '0,1': '="Micro"&"Calc"'
  });
  assert.equal(calculator.valueOfCell(0, 0).value, 'ok');
  assert.equal(calculator.valueOfCell(0, 1).value, 'MicroCalc');
});

test('date literals parse as day/month/year and format consistently', () => {
  const value = parseFormula('=#2/1/2024#', 32, 8).value;
  assert.equal(displayValue(value), '02/01/24');
  const dateCell = parseRawConstant('2/1/2024');
  assert.equal(displayValue(dateCell), '02/01/24');
});

test('date formatted cells still accept date input and preserve their format', () => {
  const sheet = { rows: 1, cols: 1, cells: { '0,0': { input: '31/12/2026', format: { type: 'date', showDate: true, showTime: false, monthFirst: false, year4: true } } } };
  const value = new Calculator(sheet).valueOfCell(0, 0);
  assert.equal(value.type, 'date');
  assert.equal(displayValue(value, sheet.cells['0,0'].format), '31/12/2026');
});

test('circular references throw and division by zero returns a spreadsheet error', () => {
  const circular = calc({ '0,0': '=B1', '0,1': '=A1' });
  assert.throws(() => circular.valueOfCell(0, 0), /Circular reference/);
  const divide = calc({ '0,0': '=1/0' });
  assert.deepEqual(divide.valueOfCell(0, 0), { type: 'error', value: '#DIV/0!' });
});

test('first printable key from the grid is inserted once into the formula field', () => {
  const calls = { prevented: 0, focused: 0, selection: null };
  const event = { key: '=', metaKey: false, ctrlKey: false, altKey: false, preventDefault() { calls.prevented++; } };
  const input = { value: '', focus() { calls.focused++; }, setSelectionRange(start, end) { calls.selection = [start, end]; } };
  assert.equal(handleGridEntryKey(event, input), true);
  assert.equal(input.value, '=');
  assert.deepEqual(calls, { prevented: 1, focused: 1, selection: [1, 1] });
  assert.equal(handleGridEntryKey({ ...event, key: '2', ctrlKey: true }, input), false);
  assert.equal(input.value, '=');
});

test('inserting and deleting rows moves cells and adjusts formula references', () => {
  const sheet = { rows: 5, cols: 3, cells: {
    '2,1': { input: '=A2+$B$1' },
    '3,0': { input: 'kept' }
  }, rowHeights: [29, 30, 31, 32, 33] };
  shiftRows(sheet, 1, 1);
  assert.equal(sheet.cells['3,1'].input, '=A3+$B$1');
  assert.equal(sheet.cells['4,0'].input, 'kept');
  assert.deepEqual(sheet.rowHeights, [29, 29, 30, 31, 32]);

  shiftRows(sheet, 1, -1);
  assert.equal(sheet.cells['2,1'].input, '=A2+$B$1');
  assert.equal(sheet.cells['3,0'].input, 'kept');

  const deletedReference = { rows: 3, cols: 2, cells: { '0,0': { input: '=A2' } }, rowHeights: [29, 29, 29] };
  shiftRows(deletedReference, 1, -1);
  assert.equal(deletedReference.cells['0,0'].input, '=#REF!');
});

test('inserting and deleting columns shifts cells and references', () => {
  const sheet = { rows: 2, cols: 4, cells: {
    '0,2': { input: '=A1+C1' },
    '1,3': { input: 'tail' }
  }, colWidths: [100, 110, 120, 130] };
  shiftColumns(sheet, 1, 1);
  assert.equal(sheet.cells['0,3'].input, '=A1+D1');
  assert.equal(sheet.cells['1,4'], undefined);
  assert.deepEqual(sheet.colWidths, [100, 128, 110, 120]);

  shiftColumns(sheet, 1, -1);
  assert.equal(sheet.cells['0,2'].input, '=A1+C1');
});

test('JSON backup export and import round-trip sheet contents and normalize dimensions', () => {
  const original = {
    version: 1, name: 'Budget', rows: 2, cols: 2,
    cells: { '0,0': { input: '42', format: { type: 'number', digits: 2 } }, '1,1': { input: '=A1*2', format: { type: 'general' } } },
    colWidths: [144, 160], rowHeights: [31, 33]
  };
  const exported = serializeSheet(original, true);
  const imported = parseSheet(exported);
  assert.deepEqual(imported.cells, original.cells);
  assert.equal(imported.name, 'Budget');
  assert.deepEqual(imported.colWidths, [144, 160]);
  assert.deepEqual(imported.rowHeights, [31, 33]);
  assert.throws(() => parseSheet('{"rows":0,"cols":2,"cells":{}}'), /Invalid MicroCalc backup/);
  const normalized = parseSheet(JSON.stringify({ rows: 1, cols: 1, cells: { '0,0': { input: 7, format: null }, bad: null } }));
  assert.deepEqual(normalized.cells, { '0,0': { input: '', format: { type: 'general' } } });
});
