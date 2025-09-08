import pandas as pd
from pathlib import Path

# 來源與輸出資料夾
src_dir = Path("../doc")
out_dir = Path("../content")
out_dir.mkdir(parents=True, exist_ok=True)

# 要轉換的 Excel 檔
files = ["irregular.xlsx", "phrase.xlsx", "pronouns.xlsx", "voca.xlsx"]

for fname in files:
    src_path = src_dir / fname
    out_path = out_dir / (src_path.stem + ".txt")
    
    try:
        # 讀第一個工作表
        df = pd.read_excel(src_path, sheet_name=0, dtype=str)
        # 填補 NaN → 空字串
        df = df.fillna("")
        # 輸出成 tab 分隔 txt
        df.to_csv(out_path, sep="\t", index=False, encoding="utf-8")
        print(f"✅ {src_path} -> {out_path}")
    except Exception as e:
        print(f"❌ 轉換失敗 {src_path}: {e}")
