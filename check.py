import streamlit as st
import pandas as pd

def main():
    st.title("列名確認ツール")
    st.write("来店データとMETA広告データのExcelファイルをアップロードしてください。")

    store_file = st.file_uploader("来店データファイル (Excel)", type="xlsx")
    ad_file = st.file_uploader("META広告データファイル (Excel)", type="xlsx")

    if store_file:
        store_df = pd.read_excel(store_file)
        st.write("来店データの列名一覧: ", store_df.columns.tolist())

    if ad_file:
        xls = pd.ExcelFile(ad_file)
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=34)  # 必要に応じてheader調整
                st.write(f"{sheet_name} シートの列名一覧: ", df.columns.tolist())
            except Exception as e:
                st.warning(f"{sheet_name} シートは読み込めませんでした。理由: {e}")

if __name__ == "__main__":
    main()
