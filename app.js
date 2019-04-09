const XLSX = require('xlsx');
const utils = require('./excel-utils');
const config = require('./config');

const book = XLSX.readFile('./in/PlanReport_new.xls');
const sheetName = book.SheetNames[0];
const sheet = book.Sheets[sheetName];

const rawRows = utils.getRowsFactory(sheet, config.stopWord)();
const rows = utils.filterNonNumeric(rawRows);
const totals = utils.calcTotals(rows);
console.log('Totals:', totals);
utils.saveAs('./out/PlanReport_new.xls', totals);

