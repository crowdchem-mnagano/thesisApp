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
# Streamlit UI
# ==========================================
st.set_page_config(page_title="Excel→JSON tool", layout="wide")
st.title("Excel → JSON tool / Excel → JSON Tool")

st.markdown("""
### このアプリの機能 / About this tool

このアプリは、ExcelファイルとJSONテンプレートを使用して、自動でJSONデータを生成します。  
This app automatically generates JSON files using an Excel data file and a JSON template.

#### 処理手順 / Workflow:
1. JSONテンプレートをアップロード  
   Upload a JSON template.  
2. Excelファイルをアップロード（以下の形式に従う）  
   Upload an Excel file with the following structure:
   - 1行目: カテゴリ / Row 1: Category  
   - 2行目: 正式名 / Row 2: Formal name  
   - 3行目: プレースホルダ（必ず `%A1%` のような形式） / Row 3: Placeholders (must be in `%A1%` format)  
   - 4行目: 略称（任意） / Row 4: Abbreviations (optional)  
   - 5行目以降: データ / Row 5 onward: Data values  

#### 置換ルール / Replacement Rules:
| 条件 / Condition | 動作 / Action |
|------------------|---------------|
| Excelに同じキーがある / Key exists in Excel | 該当値に置換 / Replace normally |
| Excelにキーがない / Key not found in Excel | 警告を表示 / Show warning |
| `"value"` または `"amount"` が空欄・NaN・"none" | `{}` 削除 / Delete the entire object |
| `"unit"`, `"name"`, `"memo"` が空欄 | 無視（削除しない） / Keep as is (not deleted) |
| JSON内に `%…%` が残っている | エラーで停止 / Stop with error |
| 3行目のセルが `%…%` 形式でない | エラーで停止 / Stop if placeholders are invalid |
""")

# ==========================================
# ファイルアップロード / File Upload
# ==========================================
json_file = st.file_uploader("JSONテンプレートをアップロード / Upload JSON template", type=["json"])
excel_file = st.file_uploader("Excelファイルをアップロード / Upload Excel file", type=["xlsx", "xls"])

output_dir = "output_json"
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir, exist_ok=True)

# ==========================================
# Excelフォーマット検証 / Excel Structure Validation
# ==========================================
def validate_excel(raw):
    errors = []
    if len(raw) < 5:
        errors.append("行数が不足しています（最低5行必要） / Not enough rows (minimum 5 required).")
    if len(raw) >= 3:
        placeholder_row = raw.iloc[2].tolist()
        invalid = [f"列 {idx+1}" for idx, val in enumerate(placeholder_row)
                   if not re.fullmatch(r"%[A-Za-z0-9_]+%", str(val).strip())]
        if invalid:
            errors.append(f"3行目の{', '.join(invalid)} に不正なプレースホルダがあります / Invalid placeholders in row 3: {', '.join(invalid)}.")
    if len(raw) >= 3 and len(raw.iloc[1]) != len(raw.iloc[2]):
        errors.append("2行目と3行目の列数が一致していません / Row 2 and row 3 column counts differ.")
    if errors:
        return False, "\n".join(errors)
    return True, "Excel構造は正常です / Excel structure validated successfully."

# ==========================================
# JSON全体を再帰的に探索して置換 / Recursive JSON Replacement
# ==========================================
def replace_placeholders_recursively(obj, row, unmatched_keys):
    """
    JSON全体を再帰的に探索して、%…% をExcel値で置換します。
    Recursively traverse JSON to replace all %...% placeholders using Excel values.

    "value" または "amount" が空欄・NaN・none の場合のみ、そのオブジェクトを削除（CrowdChem仕様）。
    Only delete objects where "value" or "amount" is empty, NaN, or "none".
    "unit", "name", "memo" が空欄の場合は削除しません。
    "unit", "name", and "memo" are kept even if empty.
    """
    if isinstance(obj, dict):
        new_dict = {}
        for key, value in obj.items():
            replaced = replace_placeholders_recursively(value, row, unmatched_keys)

            # --- プレースホルダ置換 / Replace placeholder ---
            if isinstance(replaced, str) and re.fullmatch(r"%[A-Za-z0-9_]+%", replaced):
                placeholder = replaced
                if placeholder in row:
                    val = row[placeholder]
                    if pd.isna(val):
                        replaced = ""
                    else:
                        replaced = str(val)
                else:
                    unmatched_keys.add(placeholder)
                    replaced = replaced  # keep as is for later warning

            # --- 空欄削除ロジック / Deletion rule for empty values ---
            if key in ["value", "amount"] and (pd.isna(replaced) or str(replaced).strip().lower() in ["", "none"]):
                return None  # delete the whole object
            else:
                new_dict[key] = replaced

        return new_dict if new_dict else None

    elif isinstance(obj, list):
        new_list = []
        for item in obj:
            replaced_item = replace_placeholders_recursively(item, row, unmatched_keys)
            if replaced_item not in [None, {}, []]:
                new_list.append(replaced_item)
        return new_list if new_list else None

    else:
        return obj

# ==========================================
# 実行ボタン / Execute Conversion
# ==========================================
if st.button("変換を実行 / Run Conversion", type="primary"):
    if json_file is None or excel_file is None:
        st.error("JSONテンプレートとExcelファイルを両方アップロードしてください / Please upload both JSON and Excel files.")
    else:
        try:
            # === JSONテンプレート読み込み / Load JSON template ===
            json_template = json.load(json_file)
            json_filename = os.path.splitext(os.path.basename(json_file.name))[0]

            # === Excel読み込み / Read Excel ===
            raw = pd.read_excel(excel_file, header=None, dtype=str).fillna("")

            # === 構造検証 / Validate Excel structure ===
            ok, msg = validate_excel(raw)
            if not ok:
                st.error(msg)
                st.stop()
            else:
                st.success(msg)

            # === プレースホルダ（3行目）をそのまま使用 / Use row 3 placeholders directly ===
            labels = [str(x).strip() for x in raw.iloc[2]]
            data = raw.iloc[4:].reset_index(drop=True)
            data.columns = labels

            st.info(f"Excelに {len(data)} 行のデータが見つかりました / Found {len(data)} data rows.")

            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []
            unmatched_keys_global = set()

            # === 行単位で処理 / Process each row ===
            for idx, row in data.iterrows():
                d = deepcopy(json_template)
                unmatched_keys = set()

                # JSON全体で置換 / Replace placeholders in entire JSON
                d = replace_placeholders_recursively(d, row, unmatched_keys)

                # 未一致プレースホルダ警告 / Warn unmatched placeholders
                if unmatched_keys:
                    unmatched_keys_global |= unmatched_keys
                    st.warning(f"未一致プレースホルダ（行 {idx+1}） / Unmatched placeholders (row {idx+1}): {', '.join(sorted(unmatched_keys))}")

                # 未置換プレースホルダ検出 / Detect unreplaced placeholders
                j_str = json.dumps(d, ensure_ascii=False)
                leftovers = re.findall(r"%[A-Za-z0-9_]+%", j_str)
                if leftovers:
                    st.error(f"未置換プレースホルダがあります（行 {idx+1}） / Unreplaced placeholders found (row {idx+1}): {', '.join(sorted(set(leftovers)))}")
                    st.stop()

                # JSONファイル保存 / Save JSON file
                output_path = os.path.join(output_dir, f"{json_filename}_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(data))
                status_text.text(f"{idx+1}/{len(data)} 件処理完了 / {idx+1}/{len(data)} rows processed")

            # === ZIP化 / Create ZIP archive ===
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))
            zip_buffer.seek(0)

            if unmatched_keys_global:
                st.warning(f"以下のプレースホルダはExcelに存在しませんでした / Some placeholders were not found in Excel: {', '.join(sorted(unmatched_keys_global))}")

            st.success("変換が完了しました / Conversion completed successfully.")
            st.download_button(
                "出力結果をダウンロード (ZIP) / Download results (ZIP)",
                data=zip_buffer,
                file_name=f"{json_filename}_output.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"エラーが発生しました / An error occurred: {e}")
