'use strict';

const Excel = require('exceljs');
const config = require('./config');

const INPUT_FILE = './in/PlanReport_new.xlsx';
const INPUT_START_ROW = 6;

const OUTPUT_FILE = './out/PlanReport_new.xlsx';
const OUTPUT_START_ROW = 5;
const OUTPUT_START_COL = 15;

function getGroup(num) {
    for (const group in config.mapping) {
        const values = config.mapping[group];
        if (values.indexOf(num) > -1) {
            return group;
        }
    }
    console.warn(`Number ${num} does not belong to any group!`);
    return null;
}

function calcTotals(rows) {
    const totals = {};
    for (let row of rows) {
        let num = row[0];
        let title = row[1];
        let total = row[2];
        let group = getGroup(num);
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

    const rows = [];
    for (const key in totals) {
        rows.push(totals[key]);
    }
    rows.forEach((row, index) => {
        set(startRow + index, startCol, row.group, 8);
        set(startRow + index, startCol + 1, row.title, 40);
        set(startRow + index, startCol + 2, row.total, 15);
    });
}

function main() {
    const workbook = new Excel.Workbook();
    workbook.xlsx.readFile(INPUT_FILE)
        .then(workbook => {
            const sheet = workbook.getWorksheet(1);
            const rows = [];
            sheet.eachRow((row, rowNumber) => {
                if (rowNumber < INPUT_START_ROW) {
                    return;
                }
                const key = row.values[2];
                const title = row.values[3];
                const total = row.values[10];
                if (key && /^\d+$/.test(key.trim())) {
                    rows.push([key, title, total]);
                }
            });
            const totals = calcTotals(rows);

            printData(sheet, totals, OUTPUT_START_ROW, OUTPUT_START_COL);
            workbook.xlsx.writeFile(OUTPUT_FILE)
                .then(() => console.log('Done.'))
                .catch(err => console.log(err));
        })
        .catch(err => console.log(err));
}

main();

