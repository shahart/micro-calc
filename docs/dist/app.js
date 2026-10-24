import { Calculator, address, columnName, displayValue, parseFormula } from './calc-core.js';
import { handleGridEntryKey, parseSheet, serializeSheet, shiftColumns, shiftRows } from './sheet-ops.js';
(() => {
    'use strict';
    const DEFAULT_ROWS = 32;
    const DEFAULT_COLS = 8;
    const MAX_ROWS = 500;
    const MAX_COLS = 100;
    const STORAGE_PREFIX = 'microcalc.sheet.';
    const $ = (selector) => document.querySelector(selector);
    const grid = $('#sheet-grid');
    const gridWrap = $('#grid-wrap');
    const formulaInput = $('#formula-input');
    const toast = $('#toast');
    let toastTimer;
    let sheet = newSheet('untitled', DEFAULT_ROWS, DEFAULT_COLS);
    let selected = { row: 0, col: 0 };
    let clipboard = null;
    let pendingFormula = '';
    let loadingFormula = false;
    let statusFormulaMode = false;
    function newSheet(name, rows, cols) {
        return { version: 1, name, rows, cols, cells: {}, colWidths: Array(cols).fill(128), rowHeights: Array(rows).fill(29) };
    }
    function cellKey(row, col) { return `${row},${col}`; }
    function getCell(row, col) { return sheet.cells[cellKey(row, col)] || { input: '', format: { type: 'general' } }; }
    function setCell(row, col, value) {
        const key = cellKey(row, col);
        if (!value.input && value.format.type === 'general')
            delete sheet.cells[key];
        else
            sheet.cells[key] = value;
    }
    function currentValue(row = selected.row, col = selected.col) {
        try {
            const value = new Calculator(sheet).valueOfCell(row, col);
            return value && typeof value.type === 'string' ? value : { type: 'error', value: '#ERR' };
        }
        catch (err) {
            return { type: 'error', value: err.message || '#ERR' };
        }
    }
    function clampInt(value, min, max, fallback) {
        const n = Number.parseInt(value, 10);
        return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : fallback;
    }
    function normalizeCell(row, col) {
        const cell = getCell(row, col);
        if (!cell.format)
            cell.format = { type: 'general' };
        return cell;
    }
    function renderGrid() {
        const table = document.createDocumentFragment();
        const colgroup = document.createElement('colgroup');
        const rowHeadingCol = document.createElement('col');
        rowHeadingCol.style.width = '45px';
        colgroup.append(rowHeadingCol);
        for (let col = 0; col < sheet.cols; col++) {
            const c = document.createElement('col');
            c.style.width = `${sheet.colWidths[col] || 128}px`;
            colgroup.append(c);
        }
        table.append(colgroup);
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        const corner = document.createElement('th');
        corner.className = 'corner';
        corner.textContent = '·';
        headerRow.append(corner);
        for (let col = 0; col < sheet.cols; col++) {
            const th = document.createElement('th');
            th.textContent = columnName(col);
            if (selected.col === col)
                th.classList.add('active-heading');
            headerRow.append(th);
        }
        thead.append(headerRow);
        table.append(thead);
        const tbody = document.createElement('tbody');
        for (let row = 0; row < sheet.rows; row++) {
            const tr = document.createElement('tr');
            tr.style.height = `${sheet.rowHeights[row] || 29}px`;
            const th = document.createElement('th');
            th.className = 'row-heading';
            th.textContent = String(row + 1);
            if (selected.row === row)
                th.classList.add('active-heading');
            tr.append(th);
            for (let col = 0; col < sheet.cols; col++) {
                const cellData = normalizeCell(row, col);
                const value = currentValue(row, col);
                const td = document.createElement('td');
                td.dataset.row = String(row);
                td.dataset.col = String(col);
                td.setAttribute('role', 'gridcell');
                td.setAttribute('aria-label', `${address(row, col)} ${displayValue(value, cellData.format)}`);
                td.title = `${address(row, col)}${cellData.input ? ` · ${cellData.input}` : ''}`;
                td.textContent = displayValue(value, cellData.format);
                if (row === selected.row && col === selected.col)
                    td.classList.add('selected');
                if (cellData.input.startsWith('='))
                    td.classList.add('formula-cell');
                if (value.type === 'string')
                    td.classList.add('text-cell');
                if (value.type === 'error')
                    td.classList.add('error-cell');
                tr.append(td);
            }
            tbody.append(tr);
        }
        table.append(tbody);
        grid.replaceChildren(table);
        syncSelectionUI();
    }
    function syncSelectionUI() {
        const cell = normalizeCell(selected.row, selected.col);
        $('#name-box').textContent = address(selected.row, selected.col);
        $('#sheet-title').textContent = sheet.name;
        if (document.activeElement !== formulaInput || loadingFormula)
            formulaInput.value = cell.input;
        pendingFormula = cell.input;
        $('#status-cell').textContent = address(selected.row, selected.col);
        const value = currentValue();
        $('#status-value').textContent = statusFormulaMode ? `· ${cell.input || '(empty)'}` : `· ${displayValue(value, cell.format) || '(empty)'}`;
        $('#status-message').textContent = value.type === 'error' ? value.value : 'Ready';
        $('#cell-format').value = cell.format.type || 'general';
        $('#digits-control').hidden = cell.format.type !== 'number';
        $('#separator-control').hidden = cell.format.type !== 'number';
        $('#date-control').hidden = cell.format.type !== 'date';
        $('#time-control').hidden = cell.format.type !== 'date';
        $('#date-order-control').hidden = cell.format.type !== 'date';
        $('#date-year-control').hidden = cell.format.type !== 'date';
        $('#date-separator-control').hidden = cell.format.type !== 'date';
        $('#decimal-digits').value = cell.format.digits ?? 2;
        $('#thousands-separator').checked = Boolean(cell.format.separator);
        $('#show-date').checked = cell.format.showDate !== false;
        $('#show-time').checked = Boolean(cell.format.showTime);
        $('#date-order').value = cell.format.monthFirst ? 'month-first' : 'day-first';
        $('#date-year').value = cell.format.year4 ? '4' : '2';
        $('#date-separator').value = cell.format.separator || '/';
        $('#column-width').value = sheet.colWidths[selected.col] || 128;
        updateSavedSheets();
    }
    function selectCell(row, col, shouldFocus = false) {
        selected = { row: Math.max(0, Math.min(sheet.rows - 1, row)), col: Math.max(0, Math.min(sheet.cols - 1, col)) };
        renderGrid();
        if (shouldFocus)
            gridWrap.focus({ preventScroll: true });
        const active = grid.querySelector(`[data-row="${selected.row}"][data-col="${selected.col}"]`);
        active?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    function commitFormula(value = formulaInput.value) {
        const cell = normalizeCell(selected.row, selected.col);
        cell.input = value;
        setCell(selected.row, selected.col, cell);
        saveAutosave();
        renderGrid();
        toastMessage(currentValue().type === 'error' ? currentValue().value : 'Cell updated');
    }
    function updateSavedSheets() {
        const select = $('#saved-sheets');
        const current = select.value;
        const names = listSavedSheets();
        select.replaceChildren(new Option('Open sheet…', ''));
        for (const name of names)
            select.add(new Option(name, name));
        select.value = names.includes(current) ? current : '';
    }
    function listSavedSheets() {
        const names = [];
        for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key?.startsWith(STORAGE_PREFIX) && key !== `${STORAGE_PREFIX}__current`)
                names.push(key.slice(STORAGE_PREFIX.length));
        }
        return names.sort((a, b) => a.localeCompare(b));
    }
    function saveAutosave() {
        try {
            localStorage.setItem(`${STORAGE_PREFIX}${sheet.name}`, serializeSheet(sheet));
            localStorage.setItem(`${STORAGE_PREFIX}__current`, sheet.name);
            $('#save-state').innerHTML = '<span></span> Saved locally';
        }
        catch {
            $('#save-state').innerHTML = '<span style="background:#d08a3c"></span> Storage full';
        }
    }
    function loadSaved(name) {
        try {
            const raw = localStorage.getItem(`${STORAGE_PREFIX}${name}`);
            if (!raw)
                throw new Error('Saved sheet not found');
            sheet = parseSheet(raw, { maxRows: MAX_ROWS, maxCols: MAX_COLS });
            selected = { row: 0, col: 0 };
            renderGrid();
            toastMessage(`Opened “${sheet.name}”`);
        }
        catch (err) {
            toastMessage(err.message || 'Could not open sheet');
        }
    }
    function toastMessage(message) {
        toast.textContent = message;
        toast.classList.add('visible');
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => toast.classList.remove('visible'), 2000);
    }
    function columnIndex(letters) {
        let value = 0;
        for (const letter of letters.toUpperCase())
            value = value * 26 + letter.charCodeAt(0) - 64;
        return value - 1;
    }
    function adjustFormulaRefs(formula, deltaRow, deltaCol) {
        if (!formula.startsWith('='))
            return formula;
        let output = '';
        let inString = false;
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
                    const row = Number(match[4]) - 1;
                    const col = columnIndex(match[2]);
                    const nextRow = match[3] ? row : row + deltaRow;
                    const nextCol = match[1] ? col : col + deltaCol;
                    output += nextRow < 0 || nextCol < 0 ? '#REF!' : `${match[1]}${columnName(nextCol)}${match[3]}${nextRow + 1}`;
                    i += match[0].length;
                    continue;
                }
            }
            output += formula[i++];
        }
        return output;
    }
    function copySelected() { clipboard = { ...normalizeCell(selected.row, selected.col), row: selected.row, col: selected.col }; toastMessage(`Copied ${address(selected.row, selected.col)}`); }
    function pasteSelected() {
        if (!clipboard)
            return toastMessage('Copy a cell first');
        const row = selected.row, col = selected.col;
        const data = { ...clipboard, format: { ...clipboard.format } };
        data.input = adjustFormulaRefs(data.input, row - clipboard.row, col - clipboard.col);
        delete data.row;
        delete data.col;
        setCell(row, col, data);
        saveAutosave();
        renderGrid();
        toastMessage(`Pasted into ${address(row, col)}`);
    }
    function clearSelected() { setCell(selected.row, selected.col, { input: '', format: { type: 'general' } }); saveAutosave(); renderGrid(); toastMessage('Cell cleared'); }
    function syAddress(ref, baseRow, baseCol) {
        const row = ref.absRow ? String(ref.row + 1) : ref.row === baseRow ? '' : `[${ref.row - baseRow}]`;
        const col = ref.absCol ? String(ref.col + 1) : ref.col === baseCol ? '' : `[${ref.col - baseCol}]`;
        return `R${row}C${col}`;
    }
    function sylkExpr(node, row, col) {
        if (node.kind === 'literal') {
            const value = node.value;
            if (value.type === 'string')
                return `"${value.value.replaceAll('"', '""')}"`;
            if (value.type === 'boolean')
                return value.value ? 'TRUE' : 'FALSE';
            if (value.type === 'date')
                return `"${displayValue(value)}"`;
            return String(value.value);
        }
        if (node.kind === 'group')
            return `(${sylkExpr(node.value, row, col)})`;
        if (node.kind === 'ref')
            return syAddress(node.ref, row, col);
        if (node.kind === 'range')
            return `${syAddress(node.start, row, col)}:${syAddress(node.end, row, col)}`;
        if (node.kind === 'unary')
            return `-${sylkExpr(node.value, row, col)}`;
        if (node.kind === 'binary')
            return `${sylkExpr(node.left, row, col)}${node.op}${sylkExpr(node.right, row, col)}`;
        if (node.kind === 'function')
            return `${node.name}(${node.args.map((arg) => sylkExpr(arg, row, col)).join(',')})`;
        return '';
    }
    function exportSylk() {
        const lines = ['ID;PMICROCALC;N;E', 'P;PGeneral', 'O;L;D;V0;K47;G100 0.001', `B;Y${sheet.rows};X${sheet.cols}`];
        for (let row = 0; row < sheet.rows; row++)
            for (let col = 0; col < sheet.cols; col++) {
                const cell = getCell(row, col);
                if (!cell.input)
                    continue;
                const value = currentValue(row, col);
                let rendered = value.type === 'date' ? String((value.value.getTime() / 86400000) + 25569) : displayValue(value, { type: 'general' });
                if (value.type === 'string')
                    rendered = `"${value.value.replaceAll('"', '""')}"`;
                let line = `C;Y${row + 1};X${col + 1};K${rendered}`;
                if (cell.input.startsWith('=')) {
                    try {
                        line += `;E${sylkExpr(parseFormula(cell.input, sheet.rows, sheet.cols), row, col)}`;
                    }
                    catch {
                        line += `;E${cell.input}`;
                    }
                }
                lines.push(line);
            }
        lines.push('E');
        download(`${safeFilename(sheet.name)}.slk`, `${lines.join('\r\n')}\r\n`, 'text/plain;charset=utf-8');
        toastMessage('SYLK file downloaded');
    }
    function safeFilename(name) { return name.trim().replace(/[^a-z0-9._-]+/gi, '_') || 'microcalc'; }
    function download(filename, content, mime) {
        const blob = new Blob([content], { type: mime });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = filename;
        anchor.click();
        setTimeout(() => URL.revokeObjectURL(url), 1500);
    }
    $('#new-sheet').addEventListener('click', () => {
        const name = window.prompt('Name for the new sheet', 'untitled');
        if (name === null)
            return;
        sheet = newSheet(name.trim().slice(0, 30) || 'untitled', DEFAULT_ROWS, DEFAULT_COLS);
        selected = { row: 0, col: 0 };
        saveAutosave();
        renderGrid();
        toastMessage('New 32 × 8 sheet created');
    });
    $('#save-sheet').addEventListener('click', () => {
        const name = window.prompt('Save sheet as', sheet.name);
        if (name === null)
            return;
        const oldName = sheet.name;
        sheet.name = name.trim().slice(0, 30) || 'untitled';
        if (oldName !== sheet.name && localStorage.getItem(`${STORAGE_PREFIX}${oldName}`))
            localStorage.removeItem(`${STORAGE_PREFIX}${oldName}`);
        saveAutosave();
        renderGrid();
        toastMessage(`Saved “${sheet.name}”`);
    });
    $('#saved-sheets').addEventListener('change', (event) => { if (event.target.value)
        loadSaved(event.target.value); });
    $('#delete-sheet').addEventListener('click', () => {
        if (!window.confirm(`Delete saved sheet “${sheet.name}”?`))
            return;
        localStorage.removeItem(`${STORAGE_PREFIX}${sheet.name}`);
        sheet = newSheet('untitled', DEFAULT_ROWS, DEFAULT_COLS);
        selected = { row: 0, col: 0 };
        saveAutosave();
        renderGrid();
        toastMessage('Sheet deleted');
    });
    $('#export-sylk').addEventListener('click', exportSylk);
    $('#export-json').addEventListener('click', () => {
        download(`${safeFilename(sheet.name)}.json`, serializeSheet(sheet, true), 'application/json;charset=utf-8');
        toastMessage('JSON backup downloaded');
    });
    $('#import-json').addEventListener('click', () => $('#json-file').click());
    $('#json-file').addEventListener('change', async (event) => {
        const file = event.target.files?.[0];
        if (!file)
            return;
        try {
            sheet = parseSheet(await file.text(), { maxRows: MAX_ROWS, maxCols: MAX_COLS });
            selected = { row: 0, col: 0 };
            saveAutosave();
            renderGrid();
            toastMessage(`Imported “${sheet.name}”`);
        }
        catch (err) {
            toastMessage(err.message || 'Could not import JSON');
        }
        event.target.value = '';
    });
    $('#copy-cell').addEventListener('click', copySelected);
    $('#paste-cell').addEventListener('click', pasteSelected);
    $('#clear-cell').addEventListener('click', clearSelected);
    $('#apply-formula').addEventListener('click', () => commitFormula());
    formulaInput.addEventListener('input', () => { pendingFormula = formulaInput.value; });
    formulaInput.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
            event.preventDefault();
            commitFormula();
            gridWrap.focus({ preventScroll: true });
        }
        if (event.key === 'Escape') {
            event.preventDefault();
            formulaInput.value = getCell(selected.row, selected.col).input;
            gridWrap.focus({ preventScroll: true });
        }
    });
    grid.addEventListener('click', (event) => {
        const td = event.target.closest('td[data-row]');
        if (!td)
            return;
        selectCell(Number(td.dataset.row), Number(td.dataset.col));
        gridWrap.focus({ preventScroll: true });
    });
    grid.addEventListener('dblclick', (event) => {
        const td = event.target.closest('td[data-row]');
        if (!td)
            return;
        selectCell(Number(td.dataset.row), Number(td.dataset.col));
        formulaInput.focus();
        formulaInput.select();
    });
    gridWrap.addEventListener('keydown', (event) => {
        const keyMoves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, event.shiftKey ? -1 : 1] };
        if (keyMoves[event.key]) {
            event.preventDefault();
            selectCell(selected.row + keyMoves[event.key][0], selected.col + keyMoves[event.key][1], true);
            return;
        }
        if (event.key === 'F2') {
            event.preventDefault();
            formulaInput.focus();
            formulaInput.select();
            return;
        }
        if (event.key === 'Delete' || event.key === 'Backspace') {
            event.preventDefault();
            clearSelected();
            return;
        }
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'c') {
            event.preventDefault();
            copySelected();
            return;
        }
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'v') {
            event.preventDefault();
            pasteSelected();
            return;
        }
        if (!event.metaKey && !event.ctrlKey && !event.altKey && event.key.length === 1) {
            // The key event started on the grid, but focus moves to the input here.
            // Cancel its default action so the browser does not also insert that key.
            handleGridEntryKey(event, formulaInput);
            pendingFormula = formulaInput.value;
        }
    });
    $('#cell-format').addEventListener('change', (event) => {
        const cell = normalizeCell(selected.row, selected.col);
        cell.format = { ...cell.format, type: event.target.value };
        if (event.target.value === 'date') {
            if (cell.format.showDate === undefined)
                cell.format.showDate = true;
            if (cell.format.showTime === undefined)
                cell.format.showTime = false;
            if (cell.format.monthFirst === undefined)
                cell.format.monthFirst = false;
            if (cell.format.year4 === undefined)
                cell.format.year4 = false;
            if (cell.format.separator === undefined)
                cell.format.separator = '/';
        }
        setCell(selected.row, selected.col, cell);
        saveAutosave();
        renderGrid();
    });
    for (const [selector, key, convert] of [
        ['#decimal-digits', 'digits', (v) => clampInt(v, 0, 8, 2)],
        ['#thousands-separator', 'separator', (v) => Boolean(v)],
        ['#show-date', 'showDate', (v) => Boolean(v)],
        ['#show-time', 'showTime', (v) => Boolean(v)]
    ]) {
        $(selector).addEventListener('change', (event) => {
            const cell = normalizeCell(selected.row, selected.col);
            cell.format[key] = convert(event.target.type === 'checkbox' ? event.target.checked : event.target.value);
            setCell(selected.row, selected.col, cell);
            saveAutosave();
            renderGrid();
        });
    }
    for (const [selector, key, convert] of [
        ['#date-order', 'monthFirst', (value) => value === 'month-first'],
        ['#date-year', 'year4', (value) => value === '4'],
        ['#date-separator', 'separator', (value) => value]
    ]) {
        $(selector).addEventListener('change', (event) => {
            const cell = normalizeCell(selected.row, selected.col);
            cell.format[key] = convert(event.target.value);
            setCell(selected.row, selected.col, cell);
            saveAutosave();
            renderGrid();
        });
    }
    $('#column-width').addEventListener('change', (event) => {
        sheet.colWidths[selected.col] = clampInt(event.target.value, 48, 480, 128);
        saveAutosave();
        renderGrid();
    });
    $('#resize-sheet').addEventListener('click', () => {
        const rows = clampInt(window.prompt('Number of rows', String(sheet.rows)), 1, MAX_ROWS, sheet.rows);
        const cols = clampInt(window.prompt('Number of columns', String(sheet.cols)), 1, MAX_COLS, sheet.cols);
        if (rows === sheet.rows && cols === sheet.cols)
            return;
        const next = {};
        for (const [key, value] of Object.entries(sheet.cells)) {
            const [row, col] = key.split(',').map(Number);
            if (row < rows && col < cols)
                next[key] = value;
        }
        sheet.cells = next;
        sheet.rowHeights = Array.from({ length: rows }, (_, i) => sheet.rowHeights[i] || 29);
        sheet.colWidths = Array.from({ length: cols }, (_, i) => sheet.colWidths[i] || 128);
        sheet.rows = rows;
        sheet.cols = cols;
        selected = { row: Math.min(selected.row, rows - 1), col: Math.min(selected.col, cols - 1) };
        saveAutosave();
        renderGrid();
        toastMessage(`Sheet resized to ${rows} × ${cols}`);
    });
    $('#insert-row').addEventListener('click', () => { if (sheet.rows >= MAX_ROWS)
        return toastMessage('Maximum row count reached'); shiftRows(sheet, selected.row, 1); sheet.rows++; sheet.rowHeights.push(29); saveAutosave(); renderGrid(); });
    $('#delete-row').addEventListener('click', () => { if (sheet.rows <= 1)
        return; shiftRows(sheet, selected.row, -1); sheet.rows--; sheet.rowHeights.pop(); selected.row = Math.min(selected.row, sheet.rows - 1); saveAutosave(); renderGrid(); });
    $('#insert-column').addEventListener('click', () => { if (sheet.cols >= MAX_COLS)
        return toastMessage('Maximum column count reached'); shiftColumns(sheet, selected.col, 1); sheet.cols++; sheet.colWidths.push(128); saveAutosave(); renderGrid(); });
    $('#delete-column').addEventListener('click', () => { if (sheet.cols <= 1)
        return; shiftColumns(sheet, selected.col, -1); sheet.cols--; sheet.colWidths.pop(); selected.col = Math.min(selected.col, sheet.cols - 1); saveAutosave(); renderGrid(); });
    $('#show-cell-info').addEventListener('change', (event) => { statusFormulaMode = event.target.checked; syncSelectionUI(); });
    $('#help-button').addEventListener('click', () => $('#help-dialog').showModal());
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && $('#help-dialog').open)
            return;
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') {
            event.preventDefault();
            saveAutosave();
            toastMessage('Saved locally');
        }
    });
    try {
        const currentName = localStorage.getItem(`${STORAGE_PREFIX}__current`);
        const saved = currentName && localStorage.getItem(`${STORAGE_PREFIX}${currentName}`);
        if (saved)
            sheet = parseSheet(saved, { maxRows: MAX_ROWS, maxCols: MAX_COLS });
    }
    catch {
        sheet = newSheet('untitled', DEFAULT_ROWS, DEFAULT_COLS);
    }
    renderGrid();
    saveAutosave();
})();
