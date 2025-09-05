# merge.py
from __future__ import annotations
import sys
from pathlib import Path
from typing import Optional
from openpyxl import load_workbook
from pydub import AudioSegment
from pydub.utils import which

# ===== 路徑（相對於本檔） =====
BASE = Path(__file__).resolve().parent
EXCEL_PATH = (BASE / "../doc/table.xlsx").resolve()
SHEET_NAME = "kk"
HEADER_ROW = 5  # 第 5 列為欄名（1-based）

AUDIO_SRC_DIR = (BASE / "../tmp/audio").resolve()
AUDIO_OUT_DIR = (BASE / "../audio").resolve()
XLSX_OUT_DIR  = (BASE / "output").resolve()
TXT_OUT_DIR   = (BASE / "../content").resolve()

OUT_MP3  = AUDIO_OUT_DIR / "kk.mp3"
OUT_XLSX = XLSX_OUT_DIR / "kk.xlsx"
OUT_TXT  = TXT_OUT_DIR / "kk.txt"

# ===== 檢查 & 準備 =====
if not EXCEL_PATH.exists():
    sys.exit(f"找不到 Excel：{EXCEL_PATH}")
for d in (AUDIO_OUT_DIR, XLSX_OUT_DIR, TXT_OUT_DIR):
    d.mkdir(parents=True, exist_ok=True)

if not which("ffmpeg"):
    print("⚠️ 找不到 ffmpeg（pydub 需要）。請先安裝並加入 PATH。")
if not AUDIO_SRC_DIR.exists():
    print(f"⚠️ 找不到音檔來源資料夾：{AUDIO_SRC_DIR}（仍會嘗試相對/絕對路徑）")

# ===== 讀取整張 kk（保留 1~4 列與所有欄位）=====
wb = load_workbook(EXCEL_PATH, data_only=True)
if SHEET_NAME not in wb.sheetnames:
    sys.exit(f"找不到工作表：{SHEET_NAME}")
ws = wb[SHEET_NAME]

max_row = ws.max_row
max_col = ws.max_column
if HEADER_ROW > max_row:
    sys.exit(f"表格列數不足，沒有第 {HEADER_ROW} 列可當標頭")

# ===== 取得欄位索引（1-based）=====
def find_col(header_name: str, ci: bool = False) -> Optional[int]:
    target = header_name.strip()
    t_l = target.lower()
    for c in range(1, max_col + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if isinstance(v, str):
            vv = v.strip()
            if (ci and vv.lower() == t_l) or (not ci and vv == target):
                return c
    return None

COL_PHON = find_col("音標", ci=False)
COL_MP3  = find_col("mp3",  ci=True)
COL_S    = find_col("start",ci=True)
COL_E    = find_col("end",  ci=True)
missing = [name for name, idx in [("音標", COL_PHON), ("mp3", COL_MP3), ("start", COL_S), ("end", COL_E)] if idx is None]
if missing:
    sys.exit(f"欄位缺少：{missing}（請確認第 {HEADER_ROW} 列有 音標 / mp3 / start / end）")

# ===== 解析 mp3 實體路徑 =====
def resolve_audio_path(name: str) -> Optional[Path]:
    if not name:
        return None
    name = str(name).strip()
    candidates = [name] if name.lower().endswith(".mp3") else [name + ".mp3"]

    # 1) ../tmp/audio/<name>
    for nm in candidates:
        p = (AUDIO_SRC_DIR / nm).resolve()
        if p.exists():
            return p
    # 2) 視為相對/絕對路徑
    for nm in [name] + candidates:
        p = Path(nm).expanduser().resolve()
        if p.exists():
            return p
    # 3) 遞迴搜尋來源資料夾
    if AUDIO_SRC_DIR.exists():
        stem = Path(candidates[0]).stem.lower()
        for path in AUDIO_SRC_DIR.rglob("*.mp3"):
            if path.name.lower() == candidates[0].lower() or path.stem.lower() == stem:
                return path.resolve()
    return None

# ===== 擷取 + 合併 + 更新工作表（僅改 mp3/start/end）=====
merged: Optional[AudioSegment] = None
cur_ms = 0
updated_rows = 0

print("▶ 開始擷取與合併…")
for r in range(HEADER_ROW + 1, max_row + 1):
    v_mp3 = ws.cell(row=r, column=COL_MP3).value
    v_s   = ws.cell(row=r, column=COL_S).value
    v_e   = ws.cell(row=r, column=COL_E).value

    # 轉型與檢查
    try:
        mp3name = "" if v_mp3 is None else str(v_mp3).strip()
        s = float(v_s) if v_s not in (None, "") else None
        e = float(v_e) if v_e not in (None, "") else None
    except Exception:
        continue
    if not mp3name or s is None or e is None or e <= s:
        continue

    apath = resolve_audio_path(mp3name)
    if not apath:
        print(f"  ✗ 列 {r}: 找不到音檔 {mp3name}")
        continue

    try:
        seg = AudioSegment.from_file(apath)   # 需要 ffmpeg
        s_ms = max(0, min(int(s * 1000), len(seg)))
        e_ms = max(0, min(int(e * 1000), len(seg)))
        if e_ms <= s_ms:
            print(f"  ✗ 列 {r}: 時間區間異常 start={s}, end={e}")
            continue

        clip = seg[s_ms:e_ms]

        # 合併後 kk.mp3 的時間（秒，三位小數）
        merged_s = round(cur_ms / 1000.0, 3)
        merged_e = round((cur_ms + len(clip)) / 1000.0, 3)

        merged = clip if merged is None else (merged + clip)
        cur_ms += len(clip)

        # ★ 只更新 mp3/start/end，其他欄位與 1~4 列保持原樣
        ws.cell(row=r, column=COL_MP3, value="kk.mp3")
        ws.cell(row=r, column=COL_S,   value=merged_s)
        ws.cell(row=r, column=COL_E,   value=merged_e)
        updated_rows += 1

        print(f"  ✓ 列 {r}: {apath.name} → [{merged_s}, {merged_e}]")

    except Exception as e2:
        print(f"  ✗ 列 {r}: 處理失敗 {mp3name} → {e2}")

# ===== 匯出 kk.mp3 =====
if merged is None:
    print("❌ 未產生 kk.mp3：沒有成功可合併的片段。")
else:
    merged.export(OUT_MP3, format="mp3", bitrate="192k")
    print(f"✅ 已輸出 MP3：{OUT_MP3}")

# ===== 匯出 kk.xlsx（完整保留 1~4 列與所有欄位）=====
wb.save(OUT_XLSX)
print(f"✅ 已輸出 Excel：{OUT_XLSX}（已更新 {updated_rows} 列的 mp3/start/end）")

# ===== 匯出 kk.txt（TSV；內容與 kk.xlsx 相同、保留所有欄位與 1~4 列）=====
# 逐列逐欄輸出；None→空字串；以 \t 分隔；每列以 \n 結尾
def cell_to_text(v) -> str:
    if v is None:
        return ""
    s = str(v)
    return s.replace("\t", " ").replace("\r\n", " ").replace("\n", " ")

with open(OUT_TXT, "w", encoding="utf-8", newline="\n") as f:
    for r in range(1, ws.max_row + 1):
        row_vals = [cell_to_text(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        f.write("\t".join(row_vals) + "\n")

print(f"✅ 已輸出 TXT：{OUT_TXT}（與 kk.xlsx 同欄位，TSV 格式）")




