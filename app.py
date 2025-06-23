import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def process_data(store_df, ad_sheets):
    store_df["日付"] = pd.to_datetime(store_df["日付"], errors="coerce")
    ad_dfs = []

    for sheet_name, sheet_df in ad_sheets.items():
        sheet_df.columns = [str(col).strip() for col in sheet_df.columns]
        st.write(f"{sheet_name} シートの列名一覧: ", sheet_df.columns.tolist())

        found = False
        for col in sheet_df.columns:
            if any(key in str(col) for key in ["日", "日付", "年月", "週次"]):
                sheet_df["日付"] = pd.to_datetime(sheet_df[col], errors="coerce")
                ad_dfs.append(sheet_df)
                st.success(f"{sheet_name} シートから日付列 '{col}' を使用しました。")
                found = True
                break
        if not found:
            st.warning(f"{sheet_name} シートに日付列が見つからなかったためスキップしました。")

    if not ad_dfs:
        st.error("日付列を含む広告データがありません。処理を終了します。")
        return None

    ad_df = pd.concat(ad_dfs, ignore_index=True)
    merged = pd.merge(ad_df, store_df, on="日付", how="outer")

    # 安全な指標計算
    merged["ROAS"] = merged.apply(
        lambda x: x["売上（円）"] / x["Cost"] if pd.notnull(x["売上（円）"]) and pd.notnull(x["Cost"]) and x["Cost"] != 0 else None, axis=1
    )
    merged["CPA"] = merged.apply(
        lambda x: x["Cost"] / x["CV"] if pd.notnull(x["Cost"]) and pd.notnull(x["CV"]) and x["CV"] != 0 else None, axis=1
    )
    merged["LTV"] = merged.apply(
        lambda x: x["売上（円）"] / x["CV"] if pd.notnull(x["売上（円）"]) and pd.notnull(x["CV"]) and x["CV"] != 0 else None, axis=1
    )
    merged["ROI"] = merged.apply(
        lambda x: (x["売上（円）"] - x["Cost"]) / x["Cost"] if pd.notnull(x["売上（円）"]) and pd.notnull(x["Cost"]) and x["Cost"] != 0 else None, axis=1
    )

    # ベンチマーク比較
    BENCHMARKS = {"ROAS": 1.2, "CPA": 3000, "LTV": 6000, "ROI": 0.1}
    comments = []
    roas_avg = merged["ROAS"].mean(skipna=True)
    cpa_avg = merged["CPA"].mean(skipna=True)
    ltv_avg = merged["LTV"].mean(skipna=True)
    roi_avg = merged["ROI"].mean(skipna=True)

    if roas_avg < BENCHMARKS["ROAS"]:
        comments.append("ROASが業界平均を下回っています。ターゲティングや訴求強化を推奨します。")
    else:
        comments.append("ROASは業界平均以上です。現状の施策を維持・拡大を検討ください。")

    if cpa_avg > BENCHMARKS["CPA"]:
        comments.append("CPAが高めです。クリエイティブやLP改善を推奨します。")
    else:
        comments.append("CPAは業界平均以下で良好です。現状維持で効率化を。")

    if ltv_avg < BENCHMARKS["LTV"]:
        comments.append("LTVが低めです。リピート促進やクロスセルを強化しましょう。")
    else:
        comments.append("LTVは良好です。維持施策を継続しましょう。")

    if roi_avg < BENCHMARKS["ROI"]:
        comments.append("ROIが低く、投資回収が不十分です。抜本的な施策見直しを推奨します。")
    else:
        comments.append("ROIは業界平均以上です。現状施策を拡大可能です。")

    # Excel 出力
    wb = Workbook()
    ws = wb.active
    ws.title = "KPIレポート"
    for row in dataframe_to_rows(merged, index=False, header=True):
        ws.append(row)

    cws = wb.create_sheet("改善コメント")
    for c in comments:
        cws.append([c])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Webマーケ分析アプリ (割り算安全処理版)")
    st.write("来店データとMETA広告データをアップロードしてください。")

    store_file = st.file_uploader("来店データファイル (Excel)", type="xlsx")
    ad_file = st.file_uploader("META広告データファイル (Excel)", type="xlsx")

    if store_file and ad_file:
        store_df = pd.read_excel(store_file)
        ad_sheets = {}
        xls = pd.ExcelFile(ad_file)
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=34)
                ad_sheets[sheet_name] = df
                st.success(f"{sheet_name} シートを正常に読み込みました。")
            except Exception as e:
                st.warning(f"{sheet_name} シートをスキップしました。理由: {e}")

        st.success("データを読み込みました。KPIを計算中です…")
        excel_output = process_data(store_df, ad_sheets)

        if excel_output:
            st.download_button("分析レポートをダウンロード", data=excel_output, file_name="マーケ分析レポート.xlsx")

if __name__ == "__main__":
    main()
