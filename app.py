import streamlit as st
import pandas as pd
import pdfplumber
import io

# --- PDFからデータを抽出するメインの関数 ---
def extract_tables_from_multiple_pdfs(pdf_files, keyword):
    """
    複数のPDFファイルからキーワードを含む表を抽出し、一つのDataFrameにまとめる
    """
    all_rows = []
    
    if not keyword:
        st.error("❗ キーワードが入力されていません。")
        return None

    # アップロードされた各ファイルをループ処理
    for pdf_file in pdf_files:
        # ファイル名を最初のセルに追加
        all_rows.append([f"ファイル名: {pdf_file.name}"])
        all_rows.append([]) # ファイル名の後に空行を挿入

        found_in_file = False
        last_page_number = -1

        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page_number, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text() or ""

                    if keyword in text:
                        found_in_file = True
                        
                        # --- ページが変わった時の処理 ---
                        # 同じファイル内で、最初のページではない場合
                        if last_page_number != -1 and last_page_number != page_number:
                            all_rows.append([]) # ページ間に一行空ける

                        tables = page.extract_tables()
                        for table_index, table in enumerate(tables):
                            if not table:
                                continue
                            
                            # テーブルの前にページとテーブル番号の情報を追加
                            all_rows.append([f"--- ページ {page_number} / テーブル {table_index + 1} ---"])
                            
                            for row in table:
                                cleaned_row = ["" if item is None else str(item).replace('\n', ' ') for item in row]
                                all_rows.append(cleaned_row)
                            
                            all_rows.append([]) # テーブル間に空行を挿入
                        
                        last_page_number = page_number
        
        except Exception as e:
            st.error(f"ファイル「{pdf_file.name}」の処理中にエラーが発生しました: {e}")
            continue # エラーが発生しても次のファイルの処理を続ける

        if not found_in_file:
            st.warning(f"ファイル「{pdf_file.name}」では、キーワード「{keyword}」を含む表が見つかりませんでした。")

        # --- ファイルが変わる時の処理 ---
        # 複数のファイルを処理する場合、ファイル間に黒い罫線のような区切りを入れる
        # （Excel上では空行と記号で表現）
        if len(pdf_files) > 1:
            all_rows.append(['---' * 20]) # 区切り線
            all_rows.append([])


    if not all_rows:
        return None

    return pd.DataFrame(all_rows)

# --- StreamlitのUI部分 ---

st.set_page_config(page_title="PDF表データ一括抽出ツール", layout="centered")
st.title("📄 PDF表データ一括抽出ツール")
st.write("")

# 複数ファイルのアップロードに対応
uploaded_files = st.file_uploader(
    "PDFファイルをアップロードしてください（複数選択可）", 
    type="pdf", 
    accept_multiple_files=True # この行が重要
)
keyword = st.text_input("検索キーワードを入力してください", placeholder="例: 発行済株式")

if st.button("抽出開始 ▶️", type="primary"):
    if uploaded_files:
        with st.spinner("PDFを解析中..."):
            df_result = extract_tables_from_multiple_pdfs(uploaded_files, keyword)

        if df_result is not None and not df_result.empty:
            st.success("✅ 抽出が完了しました！")
            
            output = io.BytesIO()
            # Excelに変換する際の処理
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, header=False, sheet_name='抽出結果')
                # ここでExcelの書式設定も可能ですが、まずはシンプルに
            
            st.download_button(
                label="📥 Excelファイルをダウンロード",
                data=output.getvalue(),
                file_name=f"{keyword}_一括抽出結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.dataframe(df_result)
    else:
        st.error("❗ PDFファイルをアップロードしてください。")
