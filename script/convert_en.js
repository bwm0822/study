const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 讀取Excel文件
const filePath = path.join(__dirname, '..', 'docs', 'english.xlsx');
const workbook = XLSX.readFile(filePath);

// 要轉換的sheet列表
const sheetsToConvert = [
  '進階文法',
  'GEPT',
  '全民英檢(上)',
  '文法進階篇',
  '補充'
];

// 輸出結果
const output = {};
let totalRows = 0;

sheetsToConvert.forEach(sheetName => {
  if (!workbook.SheetNames.includes(sheetName)) {
    console.warn(`⚠ Sheet "${sheetName}" 找不到！`);
    return;
  }

  const worksheet = workbook.Sheets[sheetName];

  // 轉換為JSON
  const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  // 過濾：移除第一列為'#'的行，以及空行
  const jsonData = rawData
    .filter(row => {
      // 取得第一個欄位的值
      const firstField = Object.values(row)[0];

      // 跳過第一列為 '#' 的行
      if (firstField === '#') {
        return false;
      }

      // 過濾掉只有 __EMPTY 的列
      const hasContent = Object.entries(row).some(([key, val]) => {
        return !key.includes('__EMPTY') && val !== '' && val !== null && val !== undefined;
      });
      return hasContent;
    })
    .map(row => {
      // 先添加 sheet 欄位
      const cleaned = { sheet: sheetName };
      // 再添加其他欄位
      Object.entries(row).forEach(([key, val]) => {
        if (!key.includes('__EMPTY')) {
          cleaned[key] = val;
        }
      });
      return cleaned;
    });

  output[sheetName] = jsonData;
  totalRows += jsonData.length;

  console.log(`✓ Sheet: ${sheetName} - ${jsonData.length} 列`);
});

// 輸出到文件，格式為 {sheetName: [...]}
const outputPath = path.join(__dirname, '..', 'json', 'english.json');
fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf-8');

console.log(`\n✓ 轉換成功！`);
console.log(`總計: ${totalRows} 列`);
console.log(`輸出文件: ${outputPath}`);

// 顯示轉換的sheet列表
console.log('\n轉換的Sheets:');
Object.entries(output).forEach(([name, data]) => {
  console.log(`  - ${name}: ${data.length} 列`);
});
