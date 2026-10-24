const EMPTY = Object.freeze({ type: 'empty', value: null });
const FUNCTIONS = new Set(['EXP', 'LN', 'SQRT', 'ABS', 'POWER', 'SUM', 'IF', 'AND', 'OR', 'RAND', 'MAX', 'MIN', 'PMT', 'SUMSQ', 'COUNT', 'AVERAGE', 'STDEV', 'PI', 'E']);
export class FormulaError extends Error {
}
export { EMPTY };
export function num(value, kind = 'number') { return { type: kind, value: Number(value) }; }
export function bool(value) { return { type: 'boolean', value: Boolean(value) }; }
export function str(value) { return { type: 'string', value: String(value) }; }
export function date(value) {
    const result = value instanceof Date ? value : typeof value === 'string' ? new Date(value) : new Date(value);
    return { type: 'date', value: result };
}
export function error(message = '#ERR') { return { type: 'error', value: message }; }
function isNumber(value) { return value.type === 'number' || value.type === 'long'; }
function numeric(value) {
    if (isNumber(value))
        return value.value;
    if (value.type === 'boolean')
        return value.value ? 1 : 0;
    if (value.type === 'date')
        return value.value.getTime();
    if (value.type === 'empty')
        return 0;
    throw new FormulaError('Wrong type of function argument');
}
function clampInt(value, min, max, fallback) {
    const n = Number.parseInt(value, 10);
    return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback;
}
export function columnName(index) {
    let n = index + 1;
    let name = '';
    while (n > 0) {
        const remainder = (n - 1) % 26;
        name = String.fromCharCode(65 + remainder) + name;
        n = Math.floor((n - 1) / 26);
    }
    return name;
}
export function address(row, col) { return `${columnName(col)}${row + 1}`; }
function columnIndex(letters) {
    let value = 0;
    for (const letter of letters.toUpperCase())
        value = value * 26 + letter.charCodeAt(0) - 64;
    return value - 1;
}
export function parseAddress(text, rows, cols) {
    const match = text.match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)$/i);
    if (!match)
        throw new FormulaError(`Illegal cell address [${text}]`);
    const row = Number(match[4]) - 1;
    const col = columnIndex(match[2]);
    if (row < 0 || col < 0 || row >= rows || col >= cols)
        throw new FormulaError('Reference outside the sheet');
    return { row, col, absRow: Boolean(match[3]), absCol: Boolean(match[1]) };
}
export function displayValue(value, format = { type: 'general' }) {
    if (!value || value.type === 'empty')
        return '';
    if (value.type === 'error')
        return value.value;
    if (value.type === 'string')
        return value.value;
    if (value.type === 'boolean')
        return value.value ? 'TRUE' : 'FALSE';
    if (value.type === 'date') {
        const d = value.value;
        const parts = [];
        if (format.type !== 'date' || format.showDate !== false) {
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = format.year4 ? d.getFullYear() : String(d.getFullYear()).slice(-2);
            const separator = format.separator || '/';
            parts.push(format.monthFirst ? `${month}${separator}${day}${separator}${year}` : `${day}${separator}${month}${separator}${year}`);
        }
        if (format.type === 'date' && format.showTime)
            parts.push(`${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`);
        return parts.join(' ');
    }
    if (!Number.isFinite(value.value))
        return '#ERR';
    const digits = format.type === 'number' ? clampInt(format.digits, 0, 8, 2) : 5;
    return new Intl.NumberFormat('en-US', {
        useGrouping: format.type === 'number' && Boolean(format.separator),
        minimumFractionDigits: 0,
        maximumFractionDigits: digits
    }).format(Number(value.value));
}
class Parser {
    source;
    position;
    rows;
    cols;
    constructor(source, rows, cols) { this.source = source; this.position = source.startsWith('=') ? 1 : 0; this.rows = rows; this.cols = cols; }
    skip() { while (/\s/.test(this.source[this.position] || '') && this.position < this.source.length)
        this.position++; }
    take(token) { this.skip(); if (this.source.slice(this.position, this.position + token.length) === token) {
        this.position += token.length;
        return true;
    } return false; }
    parse() {
        const ast = this.parseConcat();
        this.skip();
        if (this.position !== this.source.length)
            throw new FormulaError('Bad expression');
        return ast;
    }
    parseConcat() {
        let node = this.parseCompare();
        while (this.take('&'))
            node = { kind: 'binary', op: '&', left: node, right: this.parseCompare() };
        return node;
    }
    parseCompare() {
        let node = this.parseAdd();
        while (true) {
            const op = ['!=', '>=', '<=', '=', '>', '<'].find((candidate) => this.take(candidate));
            if (!op)
                return node;
            node = { kind: 'binary', op, left: node, right: this.parseAdd() };
        }
    }
    parseAdd() {
        let node = this.parseMultiply();
        while (true) {
            if (this.take('+'))
                node = { kind: 'binary', op: '+', left: node, right: this.parseMultiply() };
            else if (this.take('-'))
                node = { kind: 'binary', op: '-', left: node, right: this.parseMultiply() };
            else
                return node;
        }
    }
    parseMultiply() {
        let node = this.parsePower();
        while (true) {
            if (this.take('*'))
                node = { kind: 'binary', op: '*', left: node, right: this.parsePower() };
            else if (this.take('/'))
                node = { kind: 'binary', op: '/', left: node, right: this.parsePower() };
            else
                return node;
        }
    }
    parsePower() {
        const left = this.parseUnary();
        if (this.take('^'))
            return { kind: 'binary', op: '^', left, right: this.parsePower() };
        return left;
    }
    parseUnary() {
        if (this.take('-'))
            return { kind: 'unary', op: '-', value: this.parsePrimary() };
        return this.parsePrimary();
    }
    parsePrimary() {
        this.skip();
        if (this.take('(')) {
            const node = this.parseConcat();
            if (!this.take(')'))
                throw new FormulaError("Missing ')'");
            return { kind: 'group', value: node };
        }
        if (this.take('#')) {
            const end = this.source.indexOf('#', this.position);
            if (end < 0)
                throw new FormulaError("Expected trailing '#' in date constant");
            const value = parseDateLiteral(this.source.slice(this.position, end));
            this.position = end + 1;
            return { kind: 'literal', value };
        }
        const quote = this.source[this.position];
        if (quote === '"') {
            this.position++;
            let value = '';
            while (this.position < this.source.length) {
                if (this.source[this.position] === '"') {
                    if (this.source[this.position + 1] === '"') {
                        value += '"';
                        this.position += 2;
                        continue;
                    }
                    this.position++;
                    return { kind: 'literal', value: str(value) };
                }
                value += this.source[this.position++];
            }
            throw new FormulaError('Unclosed string');
        }
        const rest = this.source.slice(this.position);
        const numberMatch = rest.match(/^(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[Ll])?%?/);
        if (numberMatch) {
            this.position += numberMatch[0].length;
            const raw = numberMatch[0];
            const isLong = /[Ll]$/.test(raw);
            const isPercent = raw.endsWith('%');
            const value = Number(raw.replace(/[Ll%]/g, ''));
            if (!Number.isFinite(value))
                throw new FormulaError('Bad number');
            return { kind: 'literal', value: num(isPercent ? value / 100 : value, isLong && !isPercent ? 'long' : 'number') };
        }
        const tokenMatch = rest.match(/^(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|[A-Z_][A-Z0-9_]*)/i);
        if (!tokenMatch)
            throw new FormulaError('Bad formula: expected a value');
        this.position += tokenMatch[0].length;
        const token = tokenMatch[0];
        const upper = token.toUpperCase();
        if (this.take('(')) {
            const args = [];
            if (!this.take(')')) {
                do {
                    args.push(this.parseConcat());
                } while (this.take(','));
                if (!this.take(')'))
                    throw new FormulaError('Missing ) in function');
            }
            if (!FUNCTIONS.has(upper))
                throw new FormulaError(`Unknown function <${upper}>`);
            return { kind: 'function', name: upper, args };
        }
        if (upper === 'TRUE')
            return { kind: 'literal', value: bool(true) };
        if (upper === 'FALSE')
            return { kind: 'literal', value: bool(false) };
        if (token.includes(':')) {
            const [start, end] = token.split(':');
            return { kind: 'range', start: parseAddress(start, this.rows, this.cols), end: parseAddress(end, this.rows, this.cols) };
        }
        if (/^\$?[A-Z]{1,3}\$?\d+$/i.test(token))
            return { kind: 'ref', ref: parseAddress(token, this.rows, this.cols) };
        throw new FormulaError(`Unknown name <${token}>`);
    }
}
export function parseFormula(source, rows, cols) { return new Parser(source, rows, cols).parse(); }
function parseDateLiteral(raw) {
    const match = raw.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
    if (!match)
        throw new FormulaError('Bad date format (must be DD/MM/YYYY)');
    const [, day, month, year, hour = '0', minute = '0', second = '0'] = match;
    const d = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
    if (d.getFullYear() !== Number(year) || d.getMonth() !== Number(month) - 1 || d.getDate() !== Number(day))
        throw new FormulaError('Bad date format (must be DD/MM/YYYY)');
    return date(d);
}
export class Calculator {
    sheet;
    constructor(sheet) { this.sheet = sheet; }
    getCell(row, col) { return this.sheet.cells[`${row},${col}`] || { input: '', format: { type: 'general' } }; }
    valueOfCell(row, col, stack = new Set()) {
        if (row < 0 || col < 0 || row >= this.sheet.rows || col >= this.sheet.cols)
            throw new FormulaError('Reference outside the sheet');
        const key = `${row},${col}`;
        const cell = this.getCell(row, col);
        if (!cell.input)
            return EMPTY;
        if (!cell.input.startsWith('='))
            return parseRawConstant(cell.input, this.sheet.rows, this.sheet.cols);
        if (stack.has(key))
            throw new FormulaError('Circular reference');
        stack.add(key);
        try {
            return this.evaluate(parseFormula(cell.input, this.sheet.rows, this.sheet.cols), stack);
        }
        finally {
            stack.delete(key);
        }
    }
    evalValue(node, stack) {
        if (node.kind === 'literal')
            return node.value;
        if (node.kind === 'group')
            return this.evaluate(node.value, stack);
        if (node.kind === 'ref')
            return this.valueOfCell(node.ref.row, node.ref.col, stack);
        if (node.kind === 'range')
            return node;
        return this.evaluate(node, stack);
    }
    evalRange(range, stack) {
        const cells = [];
        const top = Math.min(range.start.row, range.end.row);
        const bottom = Math.max(range.start.row, range.end.row);
        const left = Math.min(range.start.col, range.end.col);
        const right = Math.max(range.start.col, range.end.col);
        if (top < 0 || left < 0 || bottom >= this.sheet.rows || right >= this.sheet.cols)
            throw new FormulaError('Reference outside the sheet');
        for (let row = top; row <= bottom; row++)
            for (let col = left; col <= right; col++)
                cells.push(this.valueOfCell(row, col, stack));
        return cells;
    }
    scalarValue(value, stack) {
        if (!value || value.kind !== 'range')
            return value;
        return this.valueOfCell(value.start.row, value.start.col, stack);
    }
    evaluate(node, stack = new Set()) {
        if (node.kind === 'literal')
            return node.value;
        if (node.kind === 'group')
            return this.evaluate(node.value, stack);
        if (node.kind === 'ref')
            return this.valueOfCell(node.ref.row, node.ref.col, stack);
        if (node.kind === 'range')
            return node;
        if (node.kind === 'unary') {
            const value = this.scalarValue(this.evalValue(node.value, stack), stack);
            if (!isNumber(value))
                throw new FormulaError('Unary minus with unsupported data type');
            return num(-value.value, value.type);
        }
        if (node.kind === 'binary')
            return this.applyOperator(node.op, this.scalarValue(this.evalValue(node.left, stack), stack), this.scalarValue(this.evalValue(node.right, stack), stack));
        if (node.kind === 'function')
            return this.applyFunction(node.name, node.args.map((arg) => this.evalValue(arg, stack)), stack);
        throw new FormulaError('Internal error');
    }
    applyOperator(op, left, right) {
        if (left.type === 'error' || right.type === 'error')
            return error();
        if (op === '&')
            return str(this.toText(left) + this.toText(right));
        if (op === '=' || op === '!=') {
            const equal = this.sameValue(left, right);
            return bool(op === '=' ? equal : !equal);
        }
        const rank = { empty: 0, boolean: 1, date: 2, long: 3, number: 4, string: 5 };
        const type = rank[left.type] > rank[right.type] ? left.type : right.type;
        if (type === 'string') {
            const a = this.toText(left), b = this.toText(right);
            if (op === '+')
                return str(a + b);
            if (['<', '>', '<=', '>='].includes(op)) {
                const cmp = a.length - b.length;
                return bool(op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0);
            }
            return error();
        }
        if (type === 'date') {
            if (left.type === 'date' && right.type === 'date' && op === '-')
                return num((left.value - right.value) / 86400000);
            if (left.type === 'date' && isNumber(right) && (op === '+' || op === '-'))
                return date(new Date(left.value.getTime() + (op === '+' ? 1 : -1) * right.value * 86400000));
            if (['<', '>', '<=', '>='].includes(op) && left.type === 'date' && right.type === 'date')
                return this.compare(left.value.getTime(), right.value.getTime(), op);
            return error('Illegal date operation');
        }
        let a, b;
        try {
            a = numeric(left);
            b = numeric(right);
        }
        catch {
            return error();
        }
        if (op === '<' || op === '>' || op === '<=' || op === '>=')
            return this.compare(a, b, op);
        if (op === '+')
            return num(a + b, left.type === 'long' && right.type === 'long' ? 'long' : 'number');
        if (op === '-')
            return num(a - b, left.type === 'long' && right.type === 'long' ? 'long' : 'number');
        if (op === '*')
            return num(a * b, left.type === 'long' && right.type === 'long' ? 'long' : 'number');
        if (op === '/')
            return b === 0 ? error('#DIV/0!') : num(a / b);
        if (op === '^')
            return num(Math.pow(a, b));
        return error();
    }
    compare(a, b, op) { return bool(op === '<' ? a < b : op === '>' ? a > b : op === '<=' ? a <= b : a >= b); }
    sameValue(a, b) {
        if (a.type === 'date' && b.type === 'date')
            return a.value.getTime() === b.value.getTime();
        if (isNumber(a) && isNumber(b))
            return a.value === b.value;
        return a.type === b.type && a.value === b.value;
    }
    toText(value) {
        if (value.type === 'empty')
            return '';
        if (value.type === 'string')
            return value.value;
        return displayValue(value);
    }
    functionArgs(name, args, count) {
        if (args.length !== count)
            throw new FormulaError(`Wrong number of arguments of function ${name}`);
    }
    flattenArgs(args, stack) { return args.flatMap((arg) => arg && arg.kind === 'range' ? this.evalRange(arg, stack) : [arg]); }
    applyFunction(name, args, stack) {
        if (name === 'PI')
            return num(Math.PI);
        if (name === 'E')
            return num(Math.E);
        if (name === 'IF') {
            this.functionArgs(name, args, 3);
            const condition = this.scalarValue(args[0], stack);
            if (condition.type !== 'boolean')
                throw new FormulaError('Wrong type of function argument');
            return this.scalarValue(condition.value ? args[1] : args[2], stack);
        }
        if (name === 'AND' || name === 'OR') {
            this.functionArgs(name, args, 2);
            const values = args.map((arg) => this.scalarValue(arg, stack));
            if (values.some((arg) => arg.type !== 'boolean'))
                throw new FormulaError('Wrong type of function argument');
            return bool(name === 'AND' ? values.every((arg) => arg.value) : values.some((arg) => arg.value));
        }
        if (name === 'POWER' || name === 'PMT') {
            this.functionArgs(name, args, 3 - (name === 'POWER' ? 1 : 0));
            const numbers = args.map((arg) => numeric(this.scalarValue(arg, stack)));
            if (name === 'POWER')
                return num(Math.pow(numbers[0], numbers[1]));
            const [rate, periods, principal] = numbers;
            const factor = (1 - 1 / Math.pow(1 + rate, periods)) / rate;
            return num(principal / factor);
        }
        if (['EXP', 'LN', 'SQRT', 'ABS', 'RAND'].includes(name)) {
            this.functionArgs(name, args, 1);
            const value = numeric(this.scalarValue(args[0], stack));
            if (name === 'EXP')
                return num(Math.exp(value));
            if (name === 'LN')
                return value <= 0 ? error('#NUM!') : num(Math.log(value));
            if (name === 'SQRT')
                return value < 0 ? error('#NUM!') : num(Math.sqrt(value));
            if (name === 'ABS')
                return num(Math.abs(value));
            return num(Math.floor(Math.random() * Math.max(0, Math.floor(value))));
        }
        const values = this.flattenArgs(args, stack).filter((value) => value.type !== 'empty' && value.type !== 'string');
        if (values.some((value) => value.type === 'error'))
            return error();
        const nums = values.map(numeric);
        if (name === 'COUNT')
            return num(nums.length);
        if (name === 'SUM')
            return num(nums.reduce((sum, value) => sum + value, 0));
        if (name === 'SUMSQ')
            return num(nums.reduce((sum, value) => sum + value * value, 0));
        if (name === 'MAX')
            return num(nums.length ? Math.max(...nums) : -Number.MAX_VALUE);
        if (name === 'MIN')
            return num(nums.length ? Math.min(...nums) : Number.MAX_VALUE);
        if (name === 'AVERAGE')
            return nums.length ? num(nums.reduce((sum, value) => sum + value, 0) / nums.length) : error('#DIV/0!');
        if (name === 'STDEV') {
            if (nums.length < 2)
                return error('#DIV/0!');
            const mean = nums.reduce((sum, value) => sum + value, 0) / nums.length;
            const variance = nums.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (nums.length - 1);
            return num(Math.sqrt(variance));
        }
        throw new FormulaError(`Unknown function <${name}>`);
    }
}
export function parseRawConstant(raw, rows = 32, cols = 8) {
    const value = raw.trim();
    if (!value)
        return EMPTY;
    if (/^'.*$/.test(value))
        return str(value.slice(1));
    if (/^(TRUE|FALSE)$/i.test(value))
        return bool(value.toUpperCase() === 'TRUE');
    if (/^\d{1,2}\/\d{1,2}\/\d{4}(?:\s+\d{1,2}:\d{1,2}(?::\d{1,2})?)?$/.test(value))
        return parseDateLiteral(value);
    try {
        return new Parser(value, rows, cols).parse().value;
    }
    catch {
        return str(raw);
    }
}
