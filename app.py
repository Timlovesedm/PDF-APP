import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

# --- PDFからデータを抽出するメインの関数 ---
def extract_tables_from_multiple_pdfs(pdf_files, keyword, start_page, end_page):
    """
    複数のPDFファイルからキーワードを含む表を抽出し、一つのDataFrameにまとめる
    ページ範囲指定にも対応
    """
    all_rows = []
    
    if not keyword:
        st.error("❗ キーワードが入力されていません。")
        return None

    for pdf_file in pdf_files:
        all_rows.append([f"ファイル名: {pdf_file.name}"])
        all_rows.append([])

        found_in_file = False
        last_page_number = -1

        try:
            with pdfplumber.open(pdf_file) as pdf:
                # ページ範囲の決定
                # 指定がなければ全ページを対象
                start_index = start_page - 1 if start_page else 0
                end_index = end_page if end_page else len(pdf.pages)
                
                # ページのインデックスが0から始まるため調整
                target_pages = pdf.pages[start_index:end_index]

                for page in target_pages:
                    page_number = page.page_number
                    text = page.extract_text() or ""

                    if keyword in text:
                        found_in_file = True
                        
                        if last_page_number != -1 and last_page_number != page_number:
                            all_rows.append([])

                        tables = page.extract_tables()
                        for table_index, table in enumerate(tables):
                            if not table:
                                continue
                            
                            all_rows.append([f"--- ページ {page_number} / テーブル {table_index + 1} ---"])
                            
                            for row in table:
                                cleaned_row = ["" if item is None else str(item).replace('\n', ' ') for item in row]
                                all_rows.append(cleaned_row)
                            
                            all_rows.append([])
                        
                        last_page_number = page_number
        
        except Exception as e:
            st.error(f"ファイル「{pdf_file.name}」の処理中にエラーが発生しました: {e}")
            continue

        if not found_in_file:
            st.warning(f"ファイル「{pdf_file.name}」の指定範囲では、キーワード「{keyword}」を含む表が見つかりませんでした。")

        if len(pdf_files) > 1:
            all_rows.append(['---' * 20])
            all_rows.append([])

    if not all_rows:
        return None

    return pd.DataFrame(all_rows)

# --- StreamlitのUI部分 ---

st.set_page_config(page_title="PDF表データ一括抽出ツール", layout="wide")
st.title("📄 PDF表データ一括抽出ツール")
st.write("複数のPDFファイルから、指定したキーワードとページ範囲に合致する表をまとめて抽出し、Excel形式でダウンロードします。")

# --- 入力部分 ---
uploaded_files = st.file_uploader(
    "PDFファイルをアップロードしてください（複数選択可）", 
    type="pdf", 
    accept_multiple_files=True
)
keyword = st.text_input("検索キーワードを入力してください", placeholder="例: 発行済株式")

# ページ範囲指定のUIを追加
st.write("ページ範囲を指定（空欄の場合は全ページ対象）")
col1, col2 = st.columns(2)
with col1:
    start_page_input = st.text_input("開始ページ", placeholder="例: 5")
with col2:
    end_page_input = st.text_input("終了ページ", placeholder="例: 10")

# --- 実行ボタンと処理 ---
if st.button("抽出開始 ▶️", type="primary"):
    # 入力されたページ番号を整数に変換（空欄や数字以外はNoneにする）
    start_page = int(start_page_input) if start_page_input.isdigit() else None
    end_page = int(end_page_input) if end_page_input.isdigit() else None
    
    if uploaded_files:
        with st.spinner("PDFを解析中..."):
            df_result = extract_tables_from_multiple_pdfs(uploaded_files, keyword, start_page, end_page)

        if df_result is not None and not df_result.empty:
            st.success("✅ 抽出が完了しました！")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, header=False, sheet_name='抽出結果')
            
            st.download_button(
                label="📥 Excelファイルをダウンロード",
                data=output.getvalue(),
                file_name=f"{keyword}_抽出結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.dataframe(df_result)
    else:
        st.error("❗ PDFファイルをアップロードしてください。")
