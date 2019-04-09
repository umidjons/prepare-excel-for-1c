const XLSX = require('xlsx');
const {mapping} = require('./config');

function getRowsFactory(sheet, stopValue, startIndex = 6) {
    return function () {
        const getRow = getRowFactory(sheet);
        let rowNumber = startIndex;
        const rows = [];
        do {
            const row = getRow(rowNumber);
            if (rowNumber > 1000) {
                break;
            }
            if (row) {
                rows.push(row);
                if (row[0] === stopValue || rowNumber > 1000) {
                    break;
                }
            }
            rowNumber++;
        } while (true);
        return rows;
    }
}

function getRowFactory(sheet) {
    return function (rowNumber) {
        const keys = ['B', 'C', 'D', 'G', 'H', 'J'];
        const row = [];
        for (const key of keys) {
            const addr = `${key}${rowNumber}`;
            const cell = sheet[addr];
            if (cell === undefined) {
                return null;
            }
            row.push(cell.v);
        }
        return row;
    }
}

function filterNonNumeric(rows) {
    return rows.filter(row => row && row[0] && /^\d+$/.test(row[0].trim()));
}

function getGroup(num) {
    for (const [group, values] of Object.entries(mapping)) {
        if (values.includes(num)) {
            return group;
        }
    }
    console.warn(`Number ${num} does not belong to any group!`);
    return null;
}

function calcTotals(rows) {
    const totals = {};
    for (const row of rows) {
        const [num, title, , , , total] = row;
        const group = getGroup(num);
        // console.log(`group=${group} num=${num} title=${title} total=${total}`);
        if (group) {
            if (!totals[group]) {
                totals[group] = {group, title, total: 0};
            }
            totals[group].total += total;
        }
    }
    return totals;
}

function saveAs(filename, totals, sheetName = 'Result') {
    const book = XLSX.utils.book_new();
    const rows = [];
    for (const row of Object.values(totals)) {
        rows.push([row.group, row.title, row.total]);
    }

    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = [{wpx: 70}, {wpx: 120}, {wpx: 100}];
    XLSX.utils.book_append_sheet(book, sheet, sheetName);
    XLSX.writeFile(book, filename);
}

module.exports = {
    getRowsFactory,
    filterNonNumeric,
    getGroup,
    calcTotals,
    saveAs
};