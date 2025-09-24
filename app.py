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
2. Excelデータをアップロード（形式は統一：1行目=正式名, 2行目=プレースホルダ, 4行目以降=データ）  
3. 各行ごとにテンプレートを置換して JSON 出力  
4. ZIP にまとめてダウンロード可能  
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
# 処理実行ボタン
# ==========================================
if st.button("🚀 変換を実行", type="primary"):
    if json_file is None or excel_file is None:
        st.error("⚠ JSONテンプレートとExcelファイルを両方アップロードしてください")
    else:
        try:
            # JSONテンプレート読み込み
            json_template = json.load(json_file)

            # Excel読み込み（header=None で行指定）
            raw = pd.read_excel(excel_file, header=None)

            # 1行目: 正式名
            formals = [str(x).strip() for x in raw.iloc[1]]
            # 2行目: プレースホルダ
            labels = [str(x).strip() for x in raw.iloc[2]]
            # プレースホルダ→正式名の対応表
            mapping = {lab: formal for lab, formal in zip(labels, formals) if lab and formal}

            # 4行目以降: データ本体
            data = raw.iloc[4:].reset_index(drop=True)
            data.columns = formals

            st.info(f"Excelに {len(data)} 行のデータが見つかりました。")

            # 進捗バー
            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []

            # データ行ごとに処理
            for idx, row in tqdm(data.iterrows(), total=len(data)):
                d = deepcopy(json_template)
                j_str = json.dumps(d, ensure_ascii=False)

                # プレースホルダをデータで置換
                for ph, col in mapping.items():
                    if col in row and pd.notna(row[col]):
                        j_str = j_str.replace(ph, str(row[col]))
                    else:
                        j_str = j_str.replace(ph, "")

                # JSONに戻す
                d = json.loads(j_str)

                # 出力
                output_path = os.path.join(output_dir, f"output_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(data))
                status_text.text(f"{idx+1}/{len(data)} 件処理完了")

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
