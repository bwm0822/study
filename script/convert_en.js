const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 讀取Excel文件
const filePath = path.join(__dirname, '..', 'docs', 'english.xlsx');
const workbook = XLSX.readFile(filePath);

// 要轉換的sheet列表
const sheetsToConvert = [
  '家-進階文法',
  '家-GEPT',
  '英-全民英檢(上)',
  '英-文法進階篇',
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

// 轉換「發音規則」sheet：非表格式資料，依「分類」與「規則」結構整理
// 排除「發音」欄位（注音式發音提示），只保留 母音/說明/音標/範例/備註
function parsePhoneticsRulesSheet(workbook) {
  const sheetName = '發音規則';
  if (!workbook.SheetNames.includes(sheetName)) {
    console.warn(`⚠ Sheet "${sheetName}" 找不到！`);
    return [];
  }

  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  const sections = [];
  let currentSection = null;
  let currentPattern = '';
  let introNotes = [];

  rows.forEach(rawRow => {
    // 欄位順序：母音, 說明, 音標, 發音(排除), 範例, 備註
    const [pattern, desc, phonetic, , examples, note] = [0, 1, 2, 3, 4, 5]
      .map(i => String(rawRow[i] ?? '').replace(/\r?\n/g, ' ').trim());

    const hasAnyContent = [pattern, desc, phonetic, examples, note].some(c => c !== '');
    if (!hasAnyContent) return; // 空白分隔列

    // 欄位標題列（母音,說明,音標,發音,範例,備註）
    if (pattern === '母音' && desc === '說明') return;

    const hasPhonetic = phonetic.includes('[');

    if (!hasPhonetic) {
      if (!currentSection) {
        // 尚未進入任何分類前的說明文字
        introNotes.push(pattern);
      } else if ([desc, phonetic, examples, note].every(c => c === '')) {
        // 只有第一欄有內容 -> 新分類標題
        currentSection = { 分類: pattern, 規則: [] };
        sections.push(currentSection);
        currentPattern = '';
      }
      return;
    }

    // 資料列
    if (!currentSection) {
      currentSection = { 分類: '母音', 說明: introNotes, 規則: [] };
      sections.push(currentSection);
    }
    if (pattern) currentPattern = pattern;

    const rule = { 母音: currentPattern, 音標: phonetic };
    if (desc) rule.說明 = desc;
    if (examples) rule.範例 = examples.split('、').map(s => s.trim()).filter(Boolean);
    if (note) rule.備註 = note;
    currentSection.規則.push(rule);
  });

  console.log(`✓ Sheet: ${sheetName} - ${sections.length} 分類`);
  return sections;
}

// 「發音規則」獨立輸出成自己的 json，不併入 english.json
const rulesOutputPath = path.join(__dirname, '..', 'json', 'pronunciation-rules.json');
const pronunciationRules = parsePhoneticsRulesSheet(workbook);
fs.writeFileSync(rulesOutputPath, JSON.stringify(pronunciationRules, null, 2), 'utf-8');
console.log(`✓ 輸出文件: ${rulesOutputPath}`);

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
