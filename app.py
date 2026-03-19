import datetime as dt
import io

import streamlit as st

from repair_pdf_convert import convert_pdf_to_results, create_excel_bytes


st.set_page_config(page_title="修理見積PDF → Excel", layout="centered")
st.title("修理見積PDF → Excel 変換")

st.write("メーカーPDFから「ディーラー様用」ページだけを抽出し、必要項目をExcelに出力します。")

uploaded = st.file_uploader("修理見積PDFをアップロード", type=["pdf"])

if uploaded is not None:
    with st.spinner("PDFを解析しています…"):
        pdf_bytes = uploaded.getvalue()
        results, total_pages = convert_pdf_to_results(io.BytesIO(pdf_bytes))
else:
    st.stop()

if not results:
    st.warning("ディーラー様用ページが見つかりませんでした。PDFの内容を確認してください。")
    st.stop()

st.success(f"解析完了: {total_pages}ページ / 抽出 {len(results)} 件")

st.subheader("抽出結果（先頭のみ）")
preview = results[: min(5, len(results))]
st.json(preview)

excel_bytes = create_excel_bytes(results)
today = dt.datetime.now().strftime("%Y%m%d")
out_name = f"修理見積_{today}.xlsx"

st.download_button(
    label="Excelをダウンロード",
    data=excel_bytes,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

