#!/usr/bin/env node
/**
 * 簡單加解密工具
 * 加密: 每個字元 + 125
 * 解密: 每個字元 - 125
 */

const readline = require('readline');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// 加密：每個字元 + 125，然後轉換成十六進制（避免不可見字符）
function encrypt(plaintext) {
    const encrypted = plaintext
        .split('')
        .map(char => {
            const code = char.charCodeAt(0) + 125;
            return code.toString(16).padStart(4, '0');  // 轉成 16 進位，補足 4 位
        })
        .join('');
    return encrypted;
}

// 解密：從十六進制解析，然後每個字元 - 125
function decrypt(ciphertext) {
    const plaintext = [];
    for (let i = 0; i < ciphertext.length; i += 4) {
        const hexCode = ciphertext.substr(i, 4);
        const code = parseInt(hexCode, 16) - 125;
        plaintext.push(String.fromCharCode(code));
    }
    return plaintext.join('');
}

function main() {
    console.log('\n=== GitHub Token 加解密工具 (簡單版) ===\n');
    console.log('加密方式: 每個字元 + 125');
    console.log('解密方式: 每個字元 - 125\n');

    rl.question('請選擇操作 (1=加密, 2=解密): ', (choice) => {
        if (choice === '1') {
            encryptMode();
        } else if (choice === '2') {
            decryptMode();
        } else {
            console.log('❌ 無效選擇');
            rl.close();
        }
    });
}

function encryptMode() {
    rl.question('請輸入 Token: ', (token) => {
        try {
            const encrypted = encrypt(token);
            console.log('\n✅ 加密成功！\n');
            console.log('加密後的 Token:');
            console.log('─'.repeat(70));
            console.log(encrypted);
            console.log('─'.repeat(70));
            console.log('\n💡 步驟:');
            console.log('1. 複製上面的加密值');
            console.log('2. 在 index.html 中找到 ENCRYPTED_TOKEN');
            console.log('3. 把加密值貼進去 (用單引號包起來)\n');
        } catch (error) {
            console.log('❌ 加密失敗:', error.message);
        }

        rl.close();
    });
}

function decryptMode() {
    rl.question('請輸入加密的 Token: ', (encrypted) => {
        try {
            const decrypted = decrypt(encrypted);
            console.log('\n✅ 解密成功！\n');
            console.log('原始 Token:');
            console.log('─'.repeat(70));
            console.log(decrypted);
            console.log('─'.repeat(70) + '\n');
        } catch (error) {
            console.log('❌ 解密失敗:', error.message);
        }

        rl.close();
    });
}

main();
