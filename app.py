import streamlit as st
import pandas as pd
import json
import os
import shutil
from copy import deepcopy
from io import BytesIO
from tqdm import tqdm

# ==========================================
# Streamlit UI 部分
# ==========================================
st.set_page_config(page_title="Excel→JSON変換ツール", layout="wide")
st.title("📊 Excel → JSON 変換ツール")

st.markdown("""
このアプリでは以下の処理を行います：
1. JSONテンプレートをアップロード  
2. Excelデータをアップロード  
3. Excel列名とJSONの置換キーの対応表をGUIで編集  
4. 各行ごとにテンプレートを置換して JSON 出力  
5. ZIP にまとめてダウンロード可能  
""")

# ==========================================
# ファイルアップロード
# ==========================================
json_file = st.file_uploader("JSONテンプレートをアップロード", type=["json"])
excel_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx", "xls"])

output_dir = "output_json"
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir, exist_ok=True)

# ==========================================
# 対応表設定
# ==========================================
mapping_dict = {}

if json_file is not None and excel_file is not None:
    try:
        # JSONテンプレートを先に読み込む
        json_template = json.load(json_file)
        template_str = json.dumps(json_template, ensure_ascii=False)

        # Excelを読み込む
        df = pd.read_excel(excel_file)
        st.info(f"Excelに {len(df)} 行のデータが見つかりました。")

        st.markdown("### 📝 Excel列とJSONキーの対応表を設定してください")
        for col in df.columns:
            # 候補として %列名% がJSONに含まれていれば初期値にする
            default_placeholder = f"%{col}%" if f"%{col}%" in template_str else ""
            placeholder = st.text_input(
                f"Excel列 '{col}' を置換する JSONキー",
                value=default_placeholder,
                key=f"map_{col}"
            )
            if placeholder:
                mapping_dict[col] = placeholder

    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")

# ==========================================
# 処理実行ボタン
# ==========================================
if st.button("🚀 変換を実行", type="primary"):
    if json_file is None or excel_file is None:
        st.error("⚠ JSONテンプレートとExcelファイルを両方アップロードしてください")
    elif not mapping_dict:
        st.error("⚠ 少なくとも1つはExcel列とJSONキーの対応を設定してください")
    else:
        try:
            # JSONテンプレート再読み込み（file_uploaderは一度読むとポインタが進むので注意）
            json_file.seek(0)
            json_template = json.load(json_file)

            # Excel再読み込み
            excel_file.seek(0)
            df = pd.read_excel(excel_file)

            # 進捗バー
            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []

            for idx, row in tqdm(df.iterrows(), total=len(df)):
                json_data = deepcopy(json_template)

                # ===== 対応表に基づく置換 =====
                for col, placeholder in mapping_dict.items():
                    if placeholder in str(json_data):
                        json_data = json.loads(
                            json.dumps(json_data).replace(placeholder, str(row[col]))
                        )

                # 出力
                output_path = os.path.join(output_dir, f"output_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(json_data, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(df))
                status_text.text(f"{idx+1}/{len(df)} 件処理完了")

            # ZIP化してダウンロード
            import zipfile
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))
            zip_buffer.seek(0)

            st.success("✅ 変換が完了しました！")
            st.download_button(
                "📥 出力結果をダウンロード (ZIP)",
                data=zip_buffer,
                file_name="output_json.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
