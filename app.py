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

def generate_segment_comments(df, segment_col):
    comments = []
    grouped = df.groupby(segment_col).mean(numeric_only=True)
    for seg, row in grouped.iterrows():
        roas = safe_float(row.get("ROAS"))
        cpa = safe_float(row.get("CPA"))
        ltv = safe_float(row.get("LTV"))
        roi = safe_float(row.get("ROI"))

        comment = f"【{segment_col}: {seg}】\n"
        if roas is not None and roas < 1.2:
            comment += "- ROAS低め。訴求軸・ターゲティングを見直し、A/Bテストを強化。\n"
        elif roas is not None:
            comment += "- ROAS良好。投資増を検討可能。\n"
        
        if cpa is not None and cpa > 3000:
            comment += "- CPA高め。CV導線（LP・CTA・フォーム）改善を推奨。\n"
        elif cpa is not None:
            comment += "- CPA良好。スケール検討可能。\n"
        
        if ltv is not None and ltv < 6000:
            comment += "- LTV低め。リピート施策・単価UP施策強化を。\n"
        elif ltv is not None:
            comment += "- LTV良好。優良顧客拡大を狙いましょう。\n"
        
        if roi is not None and roi < 0.1:
            comment += "- ROI低め。広告構造見直し・無駄停止を検討。\n"
        elif roi is not None:
            comment += "- ROI良好。オーガニック連携施策検討を。\n"
        
        comments.append(comment)
    return comments

def process_data(store_df, ad_sheets):
    store_df.columns = [str(col).strip() for col in store_df.columns]
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

    # KPI計算
    def calc_roas(x):
        cost = safe_float(x.get("Cost"))
        sales = safe_float(x.get("売上（円）"))
        return sales / cost if cost and cost != 0 else None

    def calc_cpa(x):
        cost = safe_float(x.get("Cost"))
        cv = safe_float(x.get("CV"))
        return cost / cv if cost and cv and cv != 0 else None

    def calc_ltv(x):
        sales = safe_float(x.get("売上（円）"))
        cv = safe_float(x.get("CV"))
        return sales / cv if sales and cv and cv != 0 else None

    def calc_roi(x):
        cost = safe_float(x.get("Cost"))
        sales = safe_float(x.get("売上（円）"))
        return (sales - cost) / cost if cost and cost != 0 else None

    merged["ROAS"] = merged.apply(calc_roas, axis=1)
    merged["CPA"] = merged.apply(calc_cpa, axis=1)
    merged["LTV"] = merged.apply(calc_ltv, axis=1)
    merged["ROI"] = merged.apply(calc_roi, axis=1)

    # 全体コメント
    comments = []
    roas_avg = merged["ROAS"].mean(skipna=True)
    cpa_avg = merged["CPA"].mean(skipna=True)
    ltv_avg = merged["LTV"].mean(skipna=True)
    roi_avg = merged["ROI"].mean(skipna=True)

    if roas_avg < 1.2:
        comments.append("全体ROAS低め。訴求軸・ターゲティング見直し、A/Bテスト強化を。")
    else:
        comments.append("全体ROAS良好。投資増を検討可能。")

    if cpa_avg > 3000:
        comments.append("全体CPA高め。CV導線（LP・CTA・フォーム）改善を推奨。")
    else:
        comments.append("全体CPA良好。スケール検討可能。")

    if ltv_avg < 6000:
        comments.append("全体LTV低め。リピート施策・単価UP施策強化を。")
    else:
        comments.append("全体LTV良好。優良顧客拡大を狙いましょう。")

    if roi_avg < 0.1:
        comments.append("全体ROI低め。広告構造見直し・無駄停止を検討。")
    else:
        comments.append("全体ROI良好。オーガニック連携施策検討を。")

    # セグメント別コメント
    if "媒体" in merged.columns:
        comments += generate_segment_comments(merged, "媒体")
    if "キャンペーン" in merged.columns:
        comments += generate_segment_comments(merged, "キャンペーン")

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
    st.title("Webマーケ分析アプリ (強化コメント＋セグメント別提案版)")
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
