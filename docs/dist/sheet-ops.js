import { columnName } from './calc-core.js';
function columnIndex(letters) {
    let value = 0;
    for (const letter of letters.toUpperCase())
        value = value * 26 + letter.charCodeAt(0) - 64;
    return value - 1;
}
export function transformReferences(formula, transform) {
    if (!formula.startsWith('='))
        return formula;
    let output = '', inString = false;
    for (let i = 0; i < formula.length;) {
        if (formula[i] === '"') {
            output += formula[i++];
            if (inString && formula[i] === '"') {
                output += formula[i++];
                continue;
            }
            inString = !inString;
            continue;
        }
        if (!inString) {
            const match = formula.slice(i).match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)/i);
            if (match) {
                const ref = { row: Number(match[4]) - 1, col: columnIndex(match[2]), absCol: Boolean(match[1]), absRow: Boolean(match[3]) };
                const next = transform({ ...ref });
                output += next ? `${match[1]}${columnName(next.col)}${match[3]}${next.row + 1}` : '#REF!';
                i += match[0].length;
                continue;
            }
        }
        output += formula[i++];
    }
    return output;
}
function shiftDimension(values, at, count, fallback, size) {
    const next = [...values];
    if (count > 0)
        next.splice(at, 0, ...Array(count).fill(fallback));
    else
        next.splice(at, -count);
    next.length = size;
    while (next.length < size)
        next.push(fallback);
    return next;
}
export function shiftRows(sheet, at, count) {
    const next = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
        const [row, col] = key.split(',').map(Number);
        const shiftedRow = row >= at ? row + count : row;
        if (shiftedRow < 0 || shiftedRow >= sheet.rows)
            continue;
        const copied = { ...cell };
        copied.input = transformReferences(copied.input, (ref) => {
            if (count > 0 && ref.row >= at)
                ref.row += count;
            else if (count < 0 && ref.row >= at && ref.row < at - count)
                return null;
            else if (count < 0 && ref.row >= at - count)
                ref.row += count;
            return ref;
        });
        next[`${shiftedRow},${col}`] = copied;
    }
    sheet.cells = next;
    sheet.rowHeights = shiftDimension(sheet.rowHeights, at, count, 29, sheet.rows);
}
export function shiftColumns(sheet, at, count) {
    const next = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
        const [row, col] = key.split(',').map(Number);
        const shiftedCol = col >= at ? col + count : col;
        if (shiftedCol < 0 || shiftedCol >= sheet.cols)
            continue;
        const copied = { ...cell };
        copied.input = transformReferences(copied.input, (ref) => {
            if (count > 0 && ref.col >= at)
                ref.col += count;
            else if (count < 0 && ref.col >= at && ref.col < at - count)
                return null;
            else if (count < 0 && ref.col >= at - count)
                ref.col += count;
            return ref;
        });
        next[`${row},${shiftedCol}`] = copied;
    }
    sheet.cells = next;
    sheet.colWidths = shiftDimension(sheet.colWidths, at, count, 128, sheet.cols);
}
export function handleGridEntryKey(event, input) {
    if (event.metaKey || event.ctrlKey || event.altKey || event.key.length !== 1)
        return false;
    event.preventDefault();
    input.focus();
    input.value = event.key;
    input.setSelectionRange(1, 1);
    return true;
}
export function serializeSheet(sheet, pretty = false) {
    return JSON.stringify(sheet, null, pretty ? 2 : 0);
}
export function parseSheet(raw, { maxRows = 500, maxCols = 100 } = {}) {
    const value = JSON.parse(raw);
    if (!value || !Number.isInteger(value.rows) || !Number.isInteger(value.cols) || value.rows < 1 || value.cols < 1 || value.rows > maxRows || value.cols > maxCols || !value.cells || typeof value.cells !== 'object' || Array.isArray(value.cells)) {
        throw new Error('Invalid MicroCalc backup');
    }
    const rows = value.rows, cols = value.cols;
    const cells = {};
    for (const [key, rawCell] of Object.entries(value.cells)) {
        const cell = rawCell;
        if (!/^\d+,\d+$/.test(key) || !cell || typeof cell !== 'object' || Array.isArray(cell))
            continue;
        const format = cell.format && typeof cell.format === 'object' && !Array.isArray(cell.format) ? cell.format : {};
        cells[key] = {
            ...cell,
            input: typeof cell.input === 'string' ? cell.input : '',
            format: { ...format, type: ['general', 'number', 'date'].includes(format.type) ? format.type : 'general' }
        };
    }
    return {
        version: 1,
        name: String(value.name || 'untitled').slice(0, 30),
        rows,
        cols,
        cells,
        colWidths: Array.from({ length: cols }, (_, i) => clampInt(value.colWidths?.[i], 48, 480, 128)),
        rowHeights: Array.from({ length: rows }, (_, i) => clampInt(value.rowHeights?.[i], 20, 120, 29))
    };
}
function clampInt(value, min, max, fallback) {
    const n = Number.parseInt(String(value), 10);
    return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback;
}
