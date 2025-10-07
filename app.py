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
   - 3行目 / Row 3 → プレースホルダ (%A1%, %B1%, …)  
   - 4行目 / Row 4 → 略称 (Abbreviation: short labels used in data rows)  
   - 5行目以降 / Row 5 onward → データ (Numeric or text data)

3. **各行ごとにテンプレートを置換**  
   Replace placeholders in the JSON template row by row  
   (apply material deletion and property replacement rules).  

4. **アップロードしたJSON名ベースでファイル出力**  
   Output JSON files named based on the uploaded template.  

5. **ZIPファイルとして一括ダウンロード可能**  
   Download all converted JSON files as a ZIP archive.  

---

⚠️ **注意 / Notes**  
- Excelフォーマットが異なる場合はエラーで停止します。  
  If your Excel structure does not follow the required format, processing will stop with an error message.  
- 各セルの空欄・0・"none" は自動的に削除されます。  
  Blank cells, "0", or "none" are automatically excluded.  
- ExcelはUTF-8互換フォーマットで保存してください。  
  Save Excel files in UTF-8 compatible format (no special characters).  

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
    """
    Excel構造と内容を検証する / Validate Excel structure and content.
    問題があれば (False, エラーメッセージ文字列) を返す。
    Return (False, message) if any issues are found.
    """
    errors = []

    # --- 行数チェック / Row count check ---
    if len(raw) < 5:
        errors.append("❌ 行数が不足しています（最低5行必要：カテゴリ・正式名・%記号・略称・データ） / Not enough rows. Minimum 5 required: category, formal, placeholder, abbreviation, data.")

    # --- プレースホルダ行(%XX%)の検証 / Placeholder format check ---
    if len(raw) >= 3:
        placeholder_row = raw.iloc[2].tolist()
        invalid = [f"列{idx+1}" for idx, val in enumerate(placeholder_row)
                   if not str(val).startswith("%") or not str(val).endswith("%")]
        if invalid:
            errors.append(f"❌ 3行目の{', '.join(invalid)} に不正なプレースホルダがあります（'%A1%' のような形式が必要） / Invalid placeholders found in row 3 ({', '.join(invalid)}). They must follow the format like '%A1%'. ")

    # --- 略称行の空欄チェック / Abbreviation check ---
    if len(raw) >= 4:
        abbr_row = raw.iloc[3].tolist()
        empty_abbr = [f"列{idx+1}" for idx, val in enumerate(abbr_row)
                      if str(val).strip() == "" or str(val).lower() == "nan"]
        if empty_abbr:
            errors.append(f"⚠️ 4行目の{', '.join(empty_abbr)} が空欄です（略称が必要） / Empty abbreviations found in row 4 ({', '.join(empty_abbr)}). Each column needs an abbreviation label.")

    # --- 正式名行とプレースホルダ行の列数一致 / Column count consistency ---
    if len(raw) >= 3 and len(raw.iloc[1]) != len(raw.iloc[2]):
        errors.append("❌ 2行目（正式名）と3行目（プレースホルダ）の列数が一致していません / The number of columns in row 2 (formal names) and row 3 (placeholders) do not match.")

    # --- データ行（5行目以降）の空行チェック / Empty data rows ---
    if len(raw) >= 5:
        data = raw.iloc[4:].fillna("")
        for r_idx, row in data.iterrows():
            if all(str(x).strip() == "" for x in row):
                errors.append(f"⚠️ データが空の行があります（Excel {r_idx + 5} 行目） / Empty data row detected (Excel row {r_idx + 5}).")

    if errors:
        return False, "\n".join(errors)
    else:
        return True, "✅ Excel構造は正常です / Excel structure validated successfully."


# ==========================================
# プロパティ置換関数 / Property replacement function
# ==========================================
def fill_properties(props, row, mapping):
    """物性: 0は残す / 空欄(None, '', 'none')は空文字に"""
    if not isinstance(props, list):
        return
    for prop in props:
        v = prop.get("value")
        if isinstance(v, str) and v in mapping:
            col = mapping[v]
            val = row[col] if col in row else None
            if pd.isna(val) or str(val).strip().lower() in ["", "none"]:
                prop["value"] = ""
            else:
                prop["value"] = str(val)


# ==========================================
# 処理実行ボタン / Execute process button
# ==========================================
if st.button("🚀 変換を実行 / Run conversion", type="primary"):
    if json_file is None or excel_file is None:
        st.error("⚠ JSONテンプレートとExcelファイルを両方アップロードしてください / Please upload both JSON and Excel files.")
    else:
        try:
            # JSONテンプレート読み込み / Load JSON template
            json_template = json.load(json_file)
            json_filename = os.path.splitext(os.path.basename(json_file.name))[0]

            # Excel読み込み / Read Excel
            raw = pd.read_excel(excel_file, header=None, dtype=str)
            raw = raw.fillna("")

            # === 構造検証 / Validate structure ===
            ok, msg = validate_excel(raw)
            if not ok:
                st.error("Excelフォーマットに問題があります / Excel format issues detected:")
                st.error(msg)
                st.stop()
            else:
                st.success(msg)

            # === データ抽出 / Extract data ===
            formals = [str(x).strip() for x in raw.iloc[1]]  # 正式名 / Formal names
            labels = [str(x).strip() for x in raw.iloc[2]]   # プレースホルダ / Placeholders
            abbrs  = [str(x).strip() for x in raw.iloc[3]]   # 略称 / Abbreviations
            data   = raw.iloc[4:].reset_index(drop=True)     # データ本体 / Main data
            data.columns = abbrs

            # プレースホルダ→略称対応表 / Mapping: placeholder → abbreviation
            mapping = {lab: abbr for lab, abbr in zip(labels, abbrs) if lab and abbr}

            st.info(f"Excelに {len(data)} 行のデータが見つかりました / Found {len(data)} rows of data in Excel.")

            # === 各行ごとの処理 / Process each row ===
            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []

            for idx, row in data.iterrows():
                d = deepcopy(json_template)

                # --- materials（最初のprocess） / First process materials ---
                new_materials = []
                for m in d["examples"][0]["processes"][0]["materials"]:
                    amount = m.get("amount")
                    if isinstance(amount, str) and amount in mapping:
                        col = mapping[amount]
                        val = row[col] if col in row else ""
                        v = str(val).strip()
                        if v in ["", "none", "0", "0.0"]:
                            continue  # 未入力・0は削除 / Skip empty or zero values
                        m["amount"] = v
                    else:
                        if not amount or (isinstance(amount, str) and amount.startswith("%")):
                            continue
                    new_materials.append(m)
                d["examples"][0]["processes"][0]["materials"] = new_materials

                # --- properties（プロセス内） / Properties inside process ---
                for proc in d["examples"][0]["processes"]:
                    fill_properties(proc.get("properties", []), row, mapping)
                # --- ルート直下 materials[*].properties も置換 / Root-level materials properties ---
                for mat in d.get("materials", []):
                    fill_properties(mat.get("properties", []), row, mapping)

                # --- 未置換プレースホルダ削除 / Remove unreplaced placeholders ---
                j_str = json.dumps(d, ensure_ascii=False)
                j_str = re.sub(r"%[A-Za-z0-9]+%", "", j_str)
                d = json.loads(j_str)

                # --- 保存 / Save file ---
                output_path = os.path.join(output_dir, f"{json_filename}_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(data))
                status_text.text(f"{idx+1}/{len(data)} 件処理完了 / {idx+1}/{len(data)} rows processed")

            # ZIP化してダウンロード / Zip and download
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
