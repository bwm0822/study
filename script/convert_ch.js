const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 讀取Excel文件
const chengYuPath = path.join(__dirname, '..', 'docs', '成語.xlsx');
const guShiPath = path.join(__dirname, '..', 'docs', '文章.xlsx');
const gaiCuoPath = path.join(__dirname, '..', 'docs', '改錯.xlsx');
const faYinPath = path.join(__dirname, '..', 'docs', '特殊發音.xlsx');

const output = {
  成語: {},
  文章: {},
  改錯: {},
  發音: {}
};

let chengYuCount = 0;
let guShiCount = 0;
let gaiCuoCount = 0;

// 處理成語Excel文件
if (fs.existsSync(chengYuPath)) {
  const workbook = XLSX.readFile(chengYuPath);

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];

    // 讀取資料
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    // 轉換資料格式 - 使用陣列存儲
    const sheetData = data
      .filter(row => row.成語 || Object.values(row).some(v => v))
      .map(row => ({
        sheet: sheetName,
        tag: row.tag || '',
        成語: row.成語 || '',
        注音: row.注音 || '',
        解釋: row.解釋 || '',
        造句: row.造句 || ''
      }));

    if (sheetData.length > 0) {
      output.成語[sheetName] = sheetData;
      chengYuCount += sheetData.length;
      console.log(`✓ Sheet: ${sheetName}`);
      console.log(`  - 成語: ${sheetData.length} 條`);
    }
  });
} else {
  console.warn(`⚠ 成語.xlsx 找不到！`);
}

// 處理文章Excel文件
if (fs.existsSync(guShiPath)) {
  const workbook = XLSX.readFile(guShiPath);

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];

    // 讀取資料
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    // 轉換資料格式 - 使用陣列存儲
    const sheetData = data
      .filter(row => row.標題 || Object.values(row).some(v => v))
      .map(row => ({
        sheet: sheetName,
        tag: row.tag || '',
        標題: row.標題 || '',
        文章: row.文章 || ''
      }));

    if (sheetData.length > 0) {
      output.文章[sheetName] = sheetData;
      guShiCount += sheetData.length;
      console.log(`✓ Sheet: ${sheetName}`);
      console.log(`  - 文章: ${sheetData.length} 條`);
    }
  });
} else {
  console.warn(`⚠ 文章.xlsx 找不到！`);
}

// 處理改錯Excel文件
if (fs.existsSync(gaiCuoPath)) {
  const workbook = XLSX.readFile(gaiCuoPath);

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];

    // 讀取資料
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    // 轉換資料格式 - 使用陣列存儲
    const sheetData = data
      .filter(row => row.題目 || Object.values(row).some(v => v))
      .map(row => ({
        sheet: sheetName,
        tag: row.tag || '',
        題目: row.題目 || '',
        答案: row.答案 || '',
        成語: row.成語 || '',
        解釋: row.解釋 || ''
      }));

    if (sheetData.length > 0) {
      output.改錯[sheetName] = sheetData;
      gaiCuoCount += sheetData.length;
      console.log(`✓ Sheet: ${sheetName}`);
      console.log(`  - 改錯: ${sheetData.length} 題`);
    }
  });
} else {
  console.warn(`⚠ 改錯.xlsx 找不到！`);
}

// 處理特殊發音Excel文件
if (fs.existsSync(faYinPath)) {
  const workbook = XLSX.readFile(faYinPath);

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];

    // 讀取資料
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    // 轉換資料格式 - 第一列為文字，第二列為發音
    data.forEach(row => {
      const keys = Object.keys(row);
      if (keys.length >= 2) {
        const textKey = row[keys[0]]; // 第一列：文字
        const pronounce = row[keys[1]]; // 第二列：發音

        if (textKey && pronounce) {
          output.發音[textKey] = pronounce;
        }
      }
    });

    if (Object.keys(output.發音).length > 0) {
      console.log(`✓ Sheet: ${sheetName}`);
      console.log(`  - 特殊發音: ${Object.keys(output.發音).length} 項`);
    }
  });
} else {
  console.warn(`⚠ 特殊發音.xlsx 找不到！`);
}

// 輸出到文件
const outputPath = path.join(__dirname, '..', 'json', 'chinese.json');
fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf-8');

console.log(`\n✓ 轉換成功！`);
console.log(`\n統計資訊:`);
console.log(`  - 成語: ${chengYuCount} 條`);
console.log(`  - 文章: ${guShiCount} 條`);
console.log(`  - 改錯: ${gaiCuoCount} 題`);
console.log(`  - 特殊發音: ${Object.keys(output.發音).length} 項`);
console.log(`  - 合計: ${chengYuCount + guShiCount + gaiCuoCount + Object.keys(output.發音).length} 項`);
console.log(`\n輸出文件: ${outputPath}`);
console.log(`\n輸出結構:`);
console.log(`{`);
console.log(`  "成語": {`);
Object.keys(output.成語).forEach(sheetName => {
  console.log(`    "${sheetName}": [${output.成語[sheetName].length} 項]`);
});
console.log(`  },`);
console.log(`  "文章": {`);
Object.keys(output.文章).forEach(sheetName => {
  console.log(`    "${sheetName}": [${output.文章[sheetName].length} 項]`);
});
console.log(`  },`);
console.log(`  "改錯": {`);
Object.keys(output.改錯).forEach(sheetName => {
  console.log(`    "${sheetName}": [${output.改錯[sheetName].length} 題]`);
});
console.log(`  }`);
console.log(`}`);
