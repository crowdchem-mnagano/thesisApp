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
st.title("Excel → JSON tool ver3.0")

st.markdown("""
### このアプリの機能 / About this tool

このアプリは、ExcelファイルとJSONテンプレートを使用して、自動で複数のJSONファイルを生成します。  
This app automatically generates multiple JSON files using an Excel data file and a JSON template.

#### 処理手順 / Workflow:
1. **JSONテンプレートをアップロード**  
   Upload a JSON template file (with placeholders like `%A1%`).
2. **Excelファイルをアップロード**  
   Upload an Excel file formatted as follows:
   - **1行目:** カテゴリ / Category  
   - **2行目:** 正式名（日本語名など） / Formal name  
   - **3行目:** プレースホルダ（必ず `%A1%` のような形式） / Placeholders (must be in `%A1%` format)  
   - **4行目:** 論文などに記載の略称（任意） / Abbreviation (optional)  
   - **5行目以降:** 各データ行 / Data rows (each will produce one JSON file)

#### 置換および整形ルール / Replacement & Formatting Rules:
| 条件 / Condition | 動作 / Behavior |
|------------------|----------------|
| Excelに同じキーがある / Key exists in Excel | Excelの値で置換 / Replace normally |
| Excelにキーがない / Key not found in Excel | 警告を表示（置換は行わない） / Show warning (keep placeholder) |
| `"value"` または `"amount"` が空欄・NaN・"none" | **削除せず `"異常値"` としてマーク** / **Marked as `"異常値"`, not deleted** |
| それ以外のキー (`unit`, `name`, `memo`, `smiles`, `properties`, `conditions` など) | 空でも削除しない (`null`禁止) / Keep even if empty (no `null`) |
| リスト要素が空 / Empty list | `[]` を出力 / Output as `[]` |
| 辞書要素が空 / Empty dict | `{}` を出力 / Output as `{}` |
| JSON内に `%…%` が残っている / Unreplaced placeholders remain | エラーで停止 / Stop with error |
| 3行目のセルが `%…%` 形式でない / Invalid placeholder format | エラーで停止 / Stop with error |

#### 出力仕様 / Output:
- 各データ行から1つのJSONファイルを生成します。  
- 出力ファイルはZIPにまとめてダウンロード可能です。  
- 未置換プレースホルダやExcelに存在しないキーは警告として表示されます。
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
def validate_excel(raw: pd.DataFrame):
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
def replace_placeholders_recursively(obj, row: pd.Series, unmatched_keys: set):
    """
    JSON全体を再帰的に探索して、%…% をExcel値で置換します。
    - "value" または "amount" が空欄・NaN・"none" の場合のみ、そのオブジェクト（現在の辞書）を削除（Noneを返す）。
    - それ以外は削除しない。リストは空でも []、辞書は空でも {} を返す。
    - 子が None（削除）になった場合、親の辞書には key を追加しない（nullを残さない）。
    """
    # --- 文字列（プレースホルダ） ---
    if isinstance(obj, str) and re.fullmatch(r"%[A-Za-z0-9_]+%", obj):
        placeholder = obj
        if placeholder in row:
            val = row[placeholder]
            if pd.isna(val):
                return "異常値"  # 空欄は空文字に
            return str(val)
        else:
            unmatched_keys.add(placeholder)
            return obj  # そのまま残し、後で検知

    # --- 配列 ---
    if isinstance(obj, list):
        new_list = []
        for item in obj:
            replaced_item = replace_placeholders_recursively(item, row, unmatched_keys)
            # 子が削除（None）の場合はスキップ。空辞書{}や空配列[]は採用しない（要素として意味が薄い場合が多いため）
            if replaced_item is "異常値":
                continue
            # if replaced_item in ({}, []):
            #     # リスト要素としての空{}や[]は実用上ノイズになりやすいので除外
            #     continue
            new_list.append(replaced_item)
        # 空でも [] を返す（null禁止）
        return new_list

    # --- 辞書 ---
    if isinstance(obj, dict):
        new_dict = {}
        # まず子を処理
        for key, value in obj.items():
            replaced = replace_placeholders_recursively(value, row, unmatched_keys)

            # "value"/"amount" の削除ルール（このキーの値が空/noneなら、**このオブジェクト自体を削除**）
            # if key in ["value", "amount"]:
                # 空・NaN・"none" は削除トリガ
                # if (replaced is None) or (isinstance(replaced, str) and replaced.strip().lower() in ["", "none"]):
                #     return None

            # 子が削除された（None）なら親辞書に key を追加しない（key:null を避ける）
            if replaced == "異常値":
                print("該当")
                continue

            # それ以外は通常どおり採用。空文字 "" / 空辞書 {} / 空配列 [] も許容（構造保持）
            new_dict[key] = replaced

        # 空でも {} を返す（null禁止）
        return new_dict

    # --- それ以外（数値・bool・None等） ---
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
            raw = pd.read_excel(excel_file, header=None, dtype=str)

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
            # 空欄セルは NaN ではなく空文字に（仕様に合わせる）
            data = data.fillna("")
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
