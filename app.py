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
st.title("Excel → JSON tool / Excel → JSON ツール")

st.markdown("""
### 📘 このアプリの機能 / About this tool
このアプリでは以下の処理を行います：  
This app performs the following steps:

1. **JSONテンプレートをアップロード**  
2. **Excelデータをアップロード（固定構造）**  
   - 1行目: カテゴリ  
   - 2行目: 正式名  
   - 3行目: プレースホルダ (%A1%, %B1%, …)  
   - 4行目: 略称（任意）  
   - 5行目以降: データ（数値や文字列）  
3. **置換実行時の動作**  
   | 状況 | 動作 |
   |------|------|
   | Excel に同じキーがある | 正常置換 |
   | Excel にキーがない | 🔶 warning に出す |
   | `"value"` が空欄/NaN/"none" | ⚠️ `{}` ごと削除（CrowdChem仕様） |
   | `"unit"`, `"name"`, `"memo"` が空欄 | 無視（削除しない） |
   | JSON 内に `%…%` が残った | 🔴 error（%xx%が置換されませんでした） |
""")

# ==========================================
# ファイルアップロード
# ==========================================
json_file = st.file_uploader("📄 JSONテンプレートをアップロード", type=["json"])
excel_file = st.file_uploader("📊 Excelファイルをアップロード", type=["xlsx", "xls"])

output_dir = "output_json"
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir, exist_ok=True)

# ==========================================
# Excelフォーマット検証
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
# JSON全体を再帰的に探索して置換
# ==========================================
def replace_placeholders_recursively(obj, row, unmatched_keys):
    """
    JSON全体を再帰的に探索して、%…% をExcel値で置換。
    "value" が空欄・NaN・none の場合のみ {} ごと削除（CrowdChem仕様）。
    unitやname,memoが空でも削除しない。
    """
    if isinstance(obj, dict):
        new_dict = {}
        for key, value in obj.items():
            replaced = replace_placeholders_recursively(value, row, unmatched_keys)

            # --- プレースホルダ置換 ---
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
                    replaced = replaced  # 残す（後で未一致警告）

            # --- 空欄削除ロジック（"value"キー限定） ---
            if key == "value" and (pd.isna(replaced) or str(replaced).strip().lower() in ["", "none"]):
                return None  # ⚠️ CrowdChem仕様：{} ごと削除
            else:
                new_dict[key] = replaced

        # 空dictは削除
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
# 実行ボタン
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

            # === Excelデータ準備 ===
            labels = [str(x).strip() for x in raw.iloc[2]]
            labels = [("%" + x.strip("%") + "%") if not str(x).startswith("%") else str(x) for x in labels]
            data = raw.iloc[4:].reset_index(drop=True)
            data.columns = labels

            st.info(f"Excelに {len(data)} 行のデータが見つかりました / Found {len(data)} data rows.")

            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []
            unmatched_keys_global = set()

            # === 行単位処理 ===
            for idx, row in data.iterrows():
                d = deepcopy(json_template)
                unmatched_keys = set()

                # ✅ JSON全体で置換
                d = replace_placeholders_recursively(d, row, unmatched_keys)

                # ⚠ 未一致キー警告
                if unmatched_keys:
                    unmatched_keys_global |= unmatched_keys
                    st.warning(f"⚠ 未一致プレースホルダ（行 {idx+1}）: {', '.join(sorted(unmatched_keys))}")

                # ❌ 未置換プレースホルダ検出
                j_str = json.dumps(d, ensure_ascii=False)
                leftovers = re.findall(r"%[A-Za-z0-9_]+%", j_str)
                if leftovers:
                    st.error(f"❌ 未置換プレースホルダがあります（行 {idx+1}）: {', '.join(sorted(set(leftovers)))}")
                    st.stop()

                # 保存
                output_path = os.path.join(output_dir, f"{json_filename}_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(data))
                status_text.text(f"{idx+1}/{len(data)} 件処理完了 / {idx+1}/{len(data)} rows processed")

            # === ZIP化 ===
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))
            zip_buffer.seek(0)

            if unmatched_keys_global:
                st.warning(f"⚠ 以下のプレースホルダはExcelに存在しませんでした: {', '.join(sorted(unmatched_keys_global))}")

            st.success("✅ 変換が完了しました！ / ✅ Conversion completed successfully!")
            st.download_button(
                "出力結果をダウンロード (ZIP) / Download results (ZIP)",
                data=zip_buffer,
                file_name=f"{json_filename}_output.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"エラーが発生しました: {e} / An error occurred: {e}")
