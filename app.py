import streamlit as st
import pandas as pd
import pdfplumber
import io

# --- PDFからデータを抽出する関数 ---
def extract_tables_from_pdf(pdf_file, keyword):
    """アップロードされたPDFファイルからキーワードを含む表を抽出し、DataFrameを返す"""
    rows_for_excel = []
    found = False

    if not keyword:
        st.error("❗ キーワードが入力されていません。")
        return None

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""

                if keyword in text:
                    found = True
                    tables = page.extract_tables()
                    for table_index, table in enumerate(tables):
                        if not table:
                            continue
                        
                        # テーブルの前にページとテーブル番号の情報を追加
                        rows_for_excel.append([f"--- ページ {page_number} / テーブル {table_index + 1} ---"])
                        
                        for row in table:
                            # セル内の改行をスペースに置換し、Noneは空文字に
                            cleaned_row = ["" if item is None else str(item).replace('\n', ' ') for item in row]
                            rows_for_excel.append(cleaned_row)
                        
                        rows_for_excel.append([]) # テーブル間に空行を挿入
    
    except Exception as e:
        st.error(f"PDFの処理中にエラーが発生しました: {e}")
        return None

    if not found:
        st.warning(f"⚠️ キーワード「{keyword}」を含む表が見つかりませんでした。")
        return None

    return pd.DataFrame(rows_for_excel)

# --- StreamlitのUI部分 ---

# 1. アプリのタイトルと説明
st.set_page_config(page_title="PDF表データ抽出ツール", layout="centered")
st.title("📄 PDF表データ抽出ツール")
st.write("PDFファイルから、指定したキーワードが含まれるページの表を抽出し、Excel形式でダウンロードします。")

# 2. UI要素の配置
uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")
keyword = st.text_input("検索キーワードを入力してください", placeholder="例: 発行済株式")

# 3. 実行ボタンと処理
if st.button("抽出開始 ▶️", type="primary"):
    if uploaded_file:
        with st.spinner("PDFを解析中..."):
            df_result = extract_tables_from_pdf(uploaded_file, keyword)

        if df_result is not None:
            st.success("✅ 抽出が完了しました！")
            
            # DataFrameをExcelファイル形式のバイナリデータに変換
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, header=False, sheet_name='抽出結果')
            
            # ダウンロードボタンを表示
            st.download_button(
                label="📥 Excelファイルをダウンロード",
                data=output.getvalue(),
                file_name=f"{keyword}_抽出結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 結果のプレビュー
            st.dataframe(df_result)
    else:
        st.error("❗ PDFファイルをアップロードしてください。")