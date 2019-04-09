const Excel = require('exceljs');
const config = require('./config');

const INPUT_FILE = './in/PlanReport_new.xlsx';
const INPUT_START_ROW = 6;

const OUTPUT_FILE = './out/PlanReport_new.xlsx';
const OUTPUT_START_ROW = 5;
const OUTPUT_START_COL = 15;

function getGroup(num) {
    for (const [group, values] of Object.entries(config.mapping)) {
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
        const [num, title, total] = row;
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

function printData(sheet, totals, startRow, startCol) {
    function setBorder(cell) {
        cell.border = {
            top: {style: 'thin'},
            left: {style: 'thin'},
            bottom: {style: 'thin'},
            right: {style: 'thin'}
        };
        return cell;
    }

    function setWidth(col, width) {
        sheet.getColumn(col).width = width;
    }

    function set(row, col, value, width) {
        const cell = sheet.getCell(row, col);
        setBorder(cell);
        setWidth(col, width);
        cell.value = value;
    }

    const rows = Object.values(totals);
    rows.forEach((row, index) => {
        set(startRow + index, startCol, row.group, 8);
        set(startRow + index, startCol + 1, row.title, 40);
        set(startRow + index, startCol + 2, row.total, 15);
    });
}

(async function main() {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(INPUT_FILE);
    const sheet = workbook.getWorksheet(1);
    const rows = [];
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber < INPUT_START_ROW) {
            return;
        }
        const [, , key, title, , , , , , , total] = row.values;
        if (key && /^\d+$/.test(key.trim())) {
            rows.push([key, title, total]);
        }
    });
    const totals = calcTotals(rows);

    printData(sheet, totals, OUTPUT_START_ROW, OUTPUT_START_COL);
    await workbook.xlsx.writeFile(OUTPUT_FILE);
})();

