const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const scriptDir = __dirname;
const convertScripts = [
  'convert_en.js',
  'convert_ch.js',
];

console.log('🚀 開始執行轉換腳本...\n');

let successCount = 0;
let failCount = 0;

convertScripts.forEach((script, index) => {
  const scriptPath = path.join(scriptDir, script);

  // 檢查文件是否存在
  if (!fs.existsSync(scriptPath)) {
    console.error(`❌ 文件不存在: ${script}`);
    failCount++;
    return;
  }

  console.log(`[${index + 1}/${convertScripts.length}] 執行 ${script}...`);
  console.log('─'.repeat(50));

  try {
    // 執行腳本
    execSync(`node "${scriptPath}"`, {
      stdio: 'inherit',
      cwd: scriptDir
    });

    console.log(`✓ ${script} 執行成功`);
    console.log('');
    successCount++;
  } catch (error) {
    console.error(`❌ ${script} 執行失敗`);
    console.error(`錯誤: ${error.message}`);
    console.log('');
    failCount++;
  }
});

console.log('═'.repeat(50));
console.log('📊 執行總結:');
console.log(`  ✓ 成功: ${successCount}/${convertScripts.length}`);
console.log(`  ❌ 失敗: ${failCount}/${convertScripts.length}`);

if (failCount === 0) {
  console.log('\n✅ 所有轉換完成！');
  process.exit(0);
} else {
  console.log('\n⚠️ 部分轉換失敗，請檢查錯誤信息。');
  process.exit(1);
}
