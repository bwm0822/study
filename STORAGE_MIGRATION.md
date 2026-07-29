# LocalStorage 版本控制與遷移機制

## 系統架構概覽

```
初始化流程
    ↓
initializeLocalStorage()
    ↓
├─ 檢查 _storageVersion
├─ 判斷是否需要遷移
├─ 執行對應遷移方案
└─ 設置新版本號

完成後執行
    ↓
cleanupStoragePrefixes()
    ↓
清理過渡期殘留的舊前綴
```

---

## 版本控制配置

### STORAGE_CONFIG 常數

```javascript
const STORAGE_CONFIG = {
    VERSION: '1.0',           // 當前版本號
    MARKER: '_storageVersion' // 版本標記鍵名
};
```

### 初始化結構

```javascript
const INITIAL_STORAGE = {
    setting: {},              // 應用設定
    english: {                // 英文數據
        favorite: {},         // 收藏單字
        quiz: {},             // 測驗歷史（正確率）
        customWords: []       // 自訂單字
    },
    chinese: {}               // 中文數據（預留）
};
```

---

## 版本判斷邏輯

### 判斷流程

```javascript
function initializeLocalStorage() {
    const currentVersion = getStorageVersion();
    
    if (!currentVersion) {
        // 無版本號 → 舊版本系統
        migrateFromLegacy();
    } else if (currentVersion !== STORAGE_CONFIG.VERSION) {
        // 版本不符 → 執行升級
        migrateVersion(currentVersion, STORAGE_CONFIG.VERSION);
    } else {
        // 版本相同 → 保持不變
        console.log('版本驗證通過');
    }
    
    setStorageVersion(STORAGE_CONFIG.VERSION);
}
```

### 判斷規則

| 情況 | 檢測方式 | 處理方案 |
|------|--------|--------|
| 首次使用 | `_storageVersion` 不存在 | 調用 `migrateFromLegacy()` |
| 版本升級 | 版本號不匹配 | 調用 `migrateVersion()` |
| 正常運行 | 版本號相同 | 跳過遷移 |

---

## 遷移方案

### 1. 舊版本遷移 (migrateFromLegacy)

**適用場景**：無 `_storageVersion` 標記的系統

**轉換規則**：

| 舊格式 | 新位置 | 轉換方式 |
|-------|-------|--------|
| `favorite_進階文法/U12/` | `english.favorite["進階文法/U12/"]` | 移除 `favorite_` 前綴 |
| `quizHistory_進階文法/U12/` | `english.quiz["進階文法/U12/"]` | 移除 `quizHistory_` 前綴 |
| `customWords` (JSON) | `english.customWords` (array) | 保持陣列格式 |

**執行步驟**：

```javascript
function migrateFromLegacy() {
    1. 初始化新結構 (favorite: {}, quiz: {}, customWords: [])
    2. 掃描所有 localStorage 鍵值
    3. 根據前綴分類遷移：
       - favorite_* → 移除前綴後存入 english.favorite
       - quizHistory_* → 移除前綴後存入 english.quiz
       - customWords → 保持格式存入 english.customWords
    4. 清除舊數據
    5. 建立新結構 (setting, english, chinese)
    6. 記錄遷移統計 (收藏數、測驗數、自訂單字數)
}
```

**日誌輸出**：

```
[Storage] 首次初始化，執行舊版遷移
[Storage] 舊版遷移完成：
{
  favorites: 24,
  quizzes: 18,
  customWords: 5
}
```

### 2. 版本升級 (migrateVersion)

**適用場景**：版本號存在但不符當前版本

**當前狀態**：

```javascript
function migrateVersion(fromVersion, toVersion) {
    console.log(`[Storage] 版本升級邏輯 (${fromVersion} → ${toVersion}) 待實現`);
}
```

**未來使用**：當升級至 v2.0 或更高時，在此實現增量遷移邏輯。

---

## 數據清理 (cleanupStoragePrefixes)

**目的**：清理遷移過程中殘留的帶前綴鍵值

**執行時機**：每次初始化後自動執行

**清理邏輯**：

```javascript
function cleanupStoragePrefixes() {
    掃描 english.favorite 中的所有鍵值
    ├─ 若鍵名以 'favorite_' 開頭
    │  ├─ 移除前綴
    │  └─ 保存清理後的鍵值
    └─ 若未變化
       └─ 跳過保存
}
```

**日誌輸出**：

```
[Storage] 清理favorite前綴完成
```

---

## 存儲操作包裝函數

### Favorite 操作

```javascript
// 讀取
getFavoriteFromStorage(favoriteKey)
// 返回值: 'true' 或 null

// 保存
saveFavoriteToStorage(favoriteKey, value)
// 自動保存至 english.favorite[favoriteKey]

// 刪除
deleteFavoriteFromStorage(favoriteKey)
// 自動刪除 english.favorite[favoriteKey]
```

### Quiz 歷史操作

```javascript
// 讀取
getQuizHistoryFromStorage(quizKey)
// 返回值: 數值 (如 3) 或 null

// 保存
saveQuizHistoryToStorage(quizKey, value)
// 自動保存至 english.quiz[quizKey]

// 刪除
deleteQuizHistoryFromStorage(quizKey)
// 自動刪除 english.quiz[quizKey]
```

### Custom Words 操作

```javascript
// 讀取
getCustomWords()
// 返回值: 陣列 []

// 保存
saveCustomWords(words)
// 自動保存至 english.customWords
```

---

## 向後相容性

### 讀取策略（Try-Fallback）

各包裝函數採用「嘗試新格式→回退舊格式」策略：

```javascript
function getFavoriteFromStorage(favoriteKey) {
    try {
        // 1. 嘗試新格式
        const data = JSON.parse(localStorage.getItem('english'));
        if (data.favorite && data.favorite[favoriteKey]) {
            return data.favorite[favoriteKey];
        }
    } catch (e) {}
    
    // 2. 回退舊格式
    return localStorage.getItem(`favorite_${favoriteKey}`);
}
```

**優勢**：
- 新舊數據無縫銜接
- 舊數據自動轉換
- 無需用戶干預

---

## 存儲結構範例

### 舊版本

```javascript
localStorage = {
    favorite_進階文法/U12/fill in: "true",
    favorite_進階文法/U12/position: "true",
    quizHistory_進階文法/U12/: "3",
    quizHistory_基礎文法/U1/: "2",
    customWords: "[{word: 'apple', ...}, ...]",
    _migrationDone: "true"
}
```

### 新版本 (v1.0)

```javascript
localStorage = {
    _storageVersion: "1.0",
    setting: "{}",
    english: {
        "favorite": {
            "進階文法/U12/fill in": "true",
            "進階文法/U12/position": "true"
        },
        "quiz": {
            "進階文法/U12/": "3",
            "基礎文法/U1/": "2"
        },
        "customWords": [
            {word: 'apple', ...},
            {...}
        ]
    },
    chinese: "{}"
}
```

---

## 調試方法

### 查看當前版本

```javascript
// 控制台執行
console.log('Version:', localStorage.getItem('_storageVersion'));
console.log('Storage:', {
    setting: JSON.parse(localStorage.getItem('setting')),
    english: JSON.parse(localStorage.getItem('english')),
    chinese: JSON.parse(localStorage.getItem('chinese'))
});
```

### 強制重新初始化

```javascript
// 控制台執行（清除所有數據）
localStorage.clear();
location.reload();
```

### 檢查特定數據

```javascript
// 檢查收藏
const eng = JSON.parse(localStorage.getItem('english'));
console.log('Favorites:', eng.favorite);
console.log('Quizzes:', eng.quiz);
console.log('Custom Words:', eng.customWords);
```

---

## 未來升級指南

### 升級至 v2.0

1. 修改版本號：
   ```javascript
   STORAGE_CONFIG.VERSION = '2.0'
   ```

2. 更新初始化結構（如有新字段）：
   ```javascript
   INITIAL_STORAGE.english.newField = {...}
   ```

3. 實現版本升級邏輯：
   ```javascript
   function migrateVersion(fromVersion, toVersion) {
       if (fromVersion === '1.0' && toVersion === '2.0') {
           // v1.0 → v2.0 升級邏輯
       }
   }
   ```

4. 系統會自動偵測並執行升級

---

## 常見問題

### Q: 為什麼還看到 `favorite_` 前綴？

**A**: 初次遷移後需手動刷新，`cleanupStoragePrefixes()` 會自動清理。

### Q: 如何確認遷移成功？

**A**: 打開控制台，執行：
```javascript
const eng = JSON.parse(localStorage.getItem('english'));
console.log('Cleanup check:', {
    hasPrefixes: Object.keys(eng.favorite).some(k => k.startsWith('favorite_')),
    cleanKeys: Object.keys(eng.favorite)
});
```

### Q: 舊數據會遺失嗎？

**A**: 否。遷移完整保留所有數據，僅調整結構和鍵名。

---

## 相關檔案

- `english.html` - 英文學習頁面（包含遷移邏輯）
- `chinese.html` - 中文學習頁面（包含遷移邏輯）
- `index.html` - 主頁面

---

**最後更新**：2026-07-22  
**版本**：1.0  
**狀態**：生產就緒
