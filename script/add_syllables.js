// 為 docs/english.xlsx 的英文單字/片語新增「音節」欄位，預設值跟「單字」欄位一樣（不拆）。
// 之後要標記音節時，直接在 Excel 裡手動把對應儲存格改成用「·」分隔的拆法即可。
//
// 重要：依專案慣例，直接寫入儲存格，不使用 XLSX.utils.aoa_to_sheet() 重建整張表（會破壞空白表頭儲存格）。
const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'docs', 'english.xlsx');
const workbook = XLSX.readFile(filePath);
const sheetsToConvert = ['家-進階文法','家-GEPT','英-全民英檢(上)','英-文法進階篇','補充-1','補充-2'];

let totalWritten = 0;

sheetsToConvert.forEach(sheetName => {
  if (!workbook.SheetNames.includes(sheetName)) { console.warn(`⚠ Sheet "${sheetName}" 找不到！`); return; }
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref']);

  let headerRow = -1, wordCol = -1, syllableCol = -1;
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v === '單字') { headerRow = r; wordCol = c; }
      if (cell && cell.v === '音節') { syllableCol = c; }
    }
    if (headerRow !== -1 && wordCol !== -1) break;
  }
  if (headerRow === -1 || wordCol === -1) { console.warn(`⚠ Sheet "${sheetName}" 找不到「單字」欄位！`); return; }

  const newCol = syllableCol !== -1 ? syllableCol : range.e.c + 1;
  ws[XLSX.utils.encode_cell({ r: headerRow, c: newCol })] = { t: 's', v: '音節' };

  let sheetWritten = 0;
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const wordCell = ws[XLSX.utils.encode_cell({ r, c: wordCol })];
    const word = wordCell && typeof wordCell.v === 'string' ? wordCell.v : '';
    const cellAddr = XLSX.utils.encode_cell({ r, c: newCol });
    if (!word.trim()) { delete ws[cellAddr]; continue; }
    ws[cellAddr] = { t: 's', v: word };
    sheetWritten++;
  }

  if (newCol > range.e.c) {
    range.e.c = newCol;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
  console.log(`✓ Sheet: ${sheetName} - 寫入 ${sheetWritten} 列的「音節」欄位（第 ${XLSX.utils.encode_col(newCol)} 欄，預設值 = 單字）`);
  totalWritten += sheetWritten;
});

XLSX.writeFile(workbook, filePath);
console.log(`\n✓ 完成，共寫入 ${totalWritten} 列。已儲存至 ${filePath}`);
