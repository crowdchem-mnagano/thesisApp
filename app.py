import streamlit as st
import pandas as pd
import json
import os
import shutil
import re
from copy import deepcopy
from io import BytesIO
import zipfile

# ==========================================
# Streamlit UI 部分 / Streamlit UI section
# ==========================================
st.set_page_config(page_title="Excel→JSON tool", layout="wide")
st.title("Excel → JSON tool / Excel → JSON ツール")

st.markdown("""
### 📘 このアプリの機能 / About this tool
このアプリでは以下の処理を行います：  
This app performs the following steps:

1. **JSONテンプレートをアップロード**  
   Upload the JSON template file.  

2. **Excelデータをアップロード**  
   Excel形式は必ず以下の構成で統一してください：  
   Please ensure your Excel follows this fixed format:  
   - 1行目 / Row 1 → 材料カテゴリ (Category: Resin, Hardener, etc.)  
   - 2行目 / Row 2 → 正式名 (Formal name: IUPAC or trade name)  
   - 3行目 / Row 3 → プレースホルダ (%A1%, %B1%, %P3%, etc.)  
   - 4行目 / Row 4 → 略称 (Abbreviation: optional, not used here)  
   - 5行目以降 / Row 5 onward → データ (Numeric or text data)

3. **1行ごとにテンプレートを複製して全ての%…%を置換**  
   Each row replaces all placeholders (%…%) in the JSON template.  

4. **アップロードしたJSON名ベースでファイル出力**  
   Output JSON files named based on the uploaded template.  

5. **ZIPファイルとして一括ダウンロード可能**  
   Download all converted JSON files as a ZIP archive.  
""")

# ==========================================
# ファイルアップロード / File upload
# ==========================================
json_file = st.file_uploader("📄 JSONテンプレートをアップロード / Upload JSON template", type=["json"])
excel_file = st.file_uploader("📊 Excelファイルをアップロード / Upload Excel file", type=["xlsx", "xls"])

output_dir = "output_json"
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir, exist_ok=True)

# ==========================================
# Excelフォーマット検証関数 / Excel validation
# ==========================================
def validate_excel(raw):
    errors = []
    if len(raw) < 5:
        errors.append("❌ 行数が不足しています（最低5行必要） / Not enough rows (minimum 5 required).")
    if len(raw) >= 3:
        placeholder_row = raw.iloc[2].tolist()
        invalid = [f"列{idx+1}" for idx, val in enumerate(placeholder_row)
                   if not re.match(r"^%[A-Za-z0-9_]+%$", str(val).strip())]
        if invalid:
            errors.append(f"❌ 3行目の{', '.join(invalid)} に不正なプレースホルダがあります / Invalid placeholders in row 3: {', '.join(invalid)}.")
    if len(raw) >= 3 and len(raw.iloc[1]) != len(raw.iloc[2]):
        errors.append("❌ 2行目と3行目の列数が一致していません / Row 2 and row 3 column counts differ.")
    if errors:
        return False, "\n".join(errors)
    return True, "✅ Excel構造は正常です / Excel structure validated successfully."

# ==========================================
# プロパティ置換関数 / Property replacement function
# ==========================================
def fill_properties(props, row):
    """Excel 3 行目の %xx% を列名として直接置換"""
    if not isinstance(props, list):
        return
    for prop in props:
        v = prop.get("value")
        if isinstance(v, str) and v in row:
            val = row[v]
            if pd.isna(val) or str(val).strip().lower() in ["", "none"]:
                prop["value"] = ""
            else:
                prop["value"] = str(val)

# ==========================================
# 実行ボタン / Run conversion
# ==========================================
if st.button("🚀 変換を実行 / Run conversion", type="primary"):
    if json_file is None or excel_file is None:
        st.error("⚠ JSONテンプレートとExcelファイルを両方アップロードしてください / Please upload both JSON and Excel files.")
    else:
        try:
            # === JSON読み込み ===
            json_template = json.load(json_file)
            json_filename = os.path.splitext(os.path.basename(json_file.name))[0]

            # === Excel読み込み ===
            raw = pd.read_excel(excel_file, header=None, dtype=str).fillna("")
            ok, msg = validate_excel(raw)
            if not ok:
                st.error(msg)
                st.stop()
            else:
                st.success(msg)

            # === データ抽出（3行目 %xx% を列名に） ===
            labels = [str(x).strip() for x in raw.iloc[2]]  # 3行目（%A1%, %B1%, …）
            data = raw.iloc[4:].reset_index(drop=True)
            data.columns = labels

            st.info(f"Excelに {len(data)} 行のデータが見つかりました / Found {len(data)} data rows.")

            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []

            # === 各行ごとの処理 ===
            for idx, row in data.iterrows():
                d = deepcopy(json_template)

                # --- materials置換 ---
                new_materials = []
                for m in d["examples"][0]["processes"][0]["materials"]:
                    amount = m.get("amount")
                    if isinstance(amount, str) and amount in row:
                        val = row[amount]
                        if str(val).strip().lower() in ["", "none", "0", "0.0"]:
                            continue
                        m["amount"] = str(val)
                    elif not amount or (isinstance(amount, str) and amount.startswith("%")):
                        continue
                    new_materials.append(m)
                d["examples"][0]["processes"][0]["materials"] = new_materials

                # --- 全物性（複数%対応）を一括置換 ---
                for proc in d["examples"][0]["processes"]:
                    fill_properties(proc.get("properties", []), row)
                for mat in d.get("materials", []):
                    fill_properties(mat.get("properties", []), row)

                # --- 未置換 %...% を削除（安全処理） ---
                j_str = json.dumps(d, ensure_ascii=False)
                j_str = re.sub(r"%[A-Za-z0-9_]+%", "", j_str)
                d = json.loads(j_str)

                # --- 保存 ---
                output_path = os.path.join(output_dir, f"{json_filename}_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(data))
                status_text.text(f"{idx+1}/{len(data)} 件処理完了 / {idx+1}/{len(data)} rows processed")

            # === ZIP化してダウンロード ===
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))
            zip_buffer.seek(0)

            st.success("✅ 変換が完了しました！ / ✅ Conversion completed successfully!")
            st.download_button(
                "出力結果をダウンロード (ZIP) / Download results (ZIP)",
                data=zip_buffer,
                file_name=f"{json_filename}_output.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"エラーが発生しました: {e} / An error occurred: {e}")
