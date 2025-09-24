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
# Streamlit UI 部分
# ==========================================
st.set_page_config(page_title="Excel→JSON変換ツール", layout="wide")
st.title("📊 Excel → JSON 変換ツール（詳細処理版）")

st.markdown("""
このアプリでは以下の処理を行います：
1. JSONテンプレートをアップロード  
2. Excelデータをアップロード（形式は統一：1行目=正式名, 2行目=プレースホルダ, 4行目以降=データ）  
3. 各行ごとにテンプレートを置換（材料の削除ルールや物性置換ルールも適用）  
4. アップロードしたJSON名ベースでファイル出力  
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
# プロパティ置換関数
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
# 処理実行ボタン
# ==========================================
if st.button("🚀 変換を実行", type="primary"):
    if json_file is None or excel_file is None:
        st.error("⚠ JSONテンプレートとExcelファイルを両方アップロードしてください")
    else:
        try:
            # JSONテンプレート読み込み
            json_template = json.load(json_file)

            # アップロードしたJSONファイル名（拡張子なし）を取得
            json_filename = os.path.splitext(os.path.basename(json_file.name))[0]

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
            for idx, row in data.iterrows():
                d = deepcopy(json_template)

                # --- materials（最初のprocess） ---
                new_materials = []
                for m in d["examples"][0]["processes"][0]["materials"]:
                    amount = m.get("amount")
                    if isinstance(amount, str) and amount in mapping:
                        col = mapping[amount]
                        val = row[col] if col in row else None
                        if pd.isna(val) or str(val).strip().lower() in ["", "none"]:
                            continue  # {}ごと削除
                        m["amount"] = str(val)
                    else:
                        # プレースホルダや空欄は削除
                        if not amount or (isinstance(amount, str) and amount.startswith("%")):
                            continue
                    new_materials.append(m)
                d["examples"][0]["processes"][0]["materials"] = new_materials

                # --- properties（プロセス内） ---
                for proc in d["examples"][0]["processes"]:
                    fill_properties(proc.get("properties", []), row, mapping)

                # --- ルート直下 materials[*].properties も置換 ---
                for mat in d.get("materials", []):
                    fill_properties(mat.get("properties", []), row, mapping)

                # --- 未置換プレースホルダ確認 ---
                j_str = json.dumps(d, ensure_ascii=False)
                leftovers = re.findall(r"%[A-Za-z0-9]+%", j_str)
                if leftovers:
                    st.warning(f"未置換のプレースホルダがあります（idx={idx}）: {', '.join(sorted(set(leftovers)))}")
                    # 未置換を空文字に強制置換
                    j_str = re.sub(r"%[A-Za-z0-9]+%", "", j_str)
                    d = json.loads(j_str)

                # --- 保存 ---
                output_path = os.path.join(output_dir, f"{json_filename}_{idx}.json")
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                generated_files.append(output_path)

                progress_bar.progress((idx + 1) / len(data))
                status_text.text(f"{idx+1}/{len(data)} 件処理完了")

            # ZIP化してダウンロード
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))
            zip_buffer.seek(0)

            st.success("✅ 変換が完了しました！")
            st.download_button(
                "📥 出力結果をダウンロード (ZIP)",
                data=zip_buffer,
                file_name=f"{json_filename}_output.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
