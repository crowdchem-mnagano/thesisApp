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
st.title("Excel → JSON tool")

st.markdown("""
このアプリでは以下の処理を行います：  
This app performs the following steps:
1. JSONテンプレートをアップロード / Upload JSON template  
2. Excelデータをアップロード（形式は統一：1行目=カテゴリ、2行目=正式名、3行目=テンプレート記号(%XX%)、4行目=略称、5行目以降=データ）  
3. 各行ごとにテンプレートを置換（材料の削除ルールや物性置換ルールも適用）  
4. アップロードしたJSON名ベースでファイル出力  
5. ZIP にまとめてダウンロード可能  
""")

# ==========================================
# ファイルアップロード / File upload
# ==========================================
json_file = st.file_uploader("JSONテンプレートをアップロード / Upload JSON template", type=["json"])
excel_file = st.file_uploader("Excelファイルをアップロード / Upload Excel file", type=["xlsx", "xls"])

output_dir = "output_json"
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir, exist_ok=True)


# ==========================================
# Excelフォーマット検証関数 / Excel validation
# ==========================================
def validate_excel(raw):
    """
    Excel構造と内容を検証する。
    問題があれば (False, エラーメッセージ文字列) を返す。
    """
    errors = []

    # --- 行数チェック ---
    if len(raw) < 5:
        errors.append("❌ 行数が不足しています（最低5行必要：カテゴリ・正式名・%記号・略称・データ）")

    # --- プレースホルダ行(%XX%)の検証 ---
    if len(raw) >= 3:
        placeholder_row = raw.iloc[2].tolist()
        invalid = [f"列{idx+1}" for idx, val in enumerate(placeholder_row)
                   if not str(val).startswith("%") or not str(val).endswith("%")]
        if invalid:
            errors.append(f"❌ 3行目の{', '.join(invalid)} に不正なプレースホルダがあります（'%A1%' のような形式が必要）")

    # --- 略称行の空欄チェック ---
    if len(raw) >= 4:
        abbr_row = raw.iloc[3].tolist()
        empty_abbr = [f"列{idx+1}" for idx, val in enumerate(abbr_row)
                      if str(val).strip() == "" or str(val).lower() == "nan"]
        if empty_abbr:
            errors.append(f"⚠️ 4行目の{', '.join(empty_abbr)} が空欄です（略称が必要）")

    # --- 正式名行とプレースホルダ行の列数一致 ---
    if len(raw) >= 3 and len(raw.iloc[1]) != len(raw.iloc[2]):
        errors.append("❌ 2行目（正式名）と3行目（プレースホルダ）の列数が一致していません。")

    # --- データ行（5行目以降）の空行チェック ---
    if len(raw) >= 5:
        data = raw.iloc[4:].fillna("")
        for r_idx, row in data.iterrows():
            if all(str(x).strip() == "" for x in row):
                errors.append(f"⚠️ データが空の行があります（Excel {r_idx + 5} 行目）")

    if errors:
        return False, "\n".join(errors)
    else:
        return True, "✅ Excel構造は正常です。"


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
        st.error("⚠ JSONテンプレートとExcelファイルを両方アップロードしてください")
    else:
        try:
            # JSONテンプレート読み込み
            json_template = json.load(json_file)
            json_filename = os.path.splitext(os.path.basename(json_file.name))[0]

            # Excel読み込み
            raw = pd.read_excel(excel_file, header=None, dtype=str)
            raw = raw.fillna("")

            # === 構造検証 ===
            ok, msg = validate_excel(raw)
            if not ok:
                st.error("Excelフォーマットに問題があります：")
                st.error(msg)
                st.stop()
            else:
                st.success(msg)

            # === データ抽出 ===
            formals = [str(x).strip() for x in raw.iloc[1]]  # 正式名
            labels = [str(x).strip() for x in raw.iloc[2]]   # プレースホルダ
            abbrs  = [str(x).strip() for x in raw.iloc[3]]   # 略称
            data   = raw.iloc[4:].reset_index(drop=True)     # データ本体
            data.columns = abbrs

            # プレースホルダ→略称対応表
            mapping = {lab: abbr for lab, abbr in zip(labels, abbrs) if lab and abbr}

            st.info(f"Excelに {len(data)} 行のデータが見つかりました。")

            # === 各行ごとの処理 ===
            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []

            for idx, row in data.iterrows():
                d = deepcopy(json_template)

                # --- materials（最初のprocess） ---
                new_materials = []
                for m in d["examples"][0]["processes"][0]["materials"]:
                    amount = m.get("amount")
                    if isinstance(amount, str) and amount in mapping:
                        col = mapping[amount]
                        val = row[col] if col in row else ""
                        v = str(val).strip()
                        if v in ["", "none", "0", "0.0"]:
                            continue  # 未入力・0は削除
                        m["amount"] = v
                    else:
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

                # --- 未置換プレースホルダ削除 ---
                j_str = json.dumps(d, ensure_ascii=False)
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
                "出力結果をダウンロード (ZIP)",
                data=zip_buffer,
                file_name=f"{json_filename}_output.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
