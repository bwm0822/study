const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'docs', 'test.xlsx');
const workbook = XLSX.readFile(filePath);

console.log('可用的 Sheets:');
workbook.SheetNames.forEach((name, i) => {
  console.log(`  ${i + 1}. ${name}`);
});
