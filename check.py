import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def safe_float(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def process_data(store_df, ad_sheets):
    # 列名クリーンアップ
    store_df.columns = [str(col).strip() for col in store_df.columns]
    
    # 日付列確認
    if "日付" not in store_df.columns:
        st.error("来店データに「日付」列が見つかりません。列名を確認してください。")
        st.write("検出した列名: ", store_df.columns.tolist())
        return None

    store_df["日付"] = pd.to_datetime(store_df["日付"], errors="coerce")
    ad_dfs = []

    for sheet_name, sheet_df in ad_sheets.items():
        sheet_df.columns = [str(col).strip() for col in sheet_df.columns]
        st.write(f"{sheet_name} シートの列名一覧: ", sheet_df.columns.tolist())

        found = False
        for col in sheet_df.columns:
            if any(key in str(col) for key in ["日", "日付", "年月", "週次"]):
