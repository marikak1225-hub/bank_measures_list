import io
from datetime import datetime
import openpyxl
import streamlit as st


st.set_page_config(page_title="転記用生成ツール", layout="wide")

FMT_PATH = "転記用FMT.xlsx"  # ←転記用フォーマット

st.title("転記用エクセル生成ツール")
st.caption("施策一覧から転記用形式のExcelを生成します")

# ✅ アップは施策一覧
input_file = st.file_uploader("施策一覧をアップロード", type=["xlsx"])


def build_workbook(input_bytes, fmt_bytes):
    # ✅ 入力 = 施策一覧
    src_wb = openpyxl.load_workbook(io.BytesIO(input_bytes), data_only=True)

    # ✅ 出力 = 転記用FMT
    out_wb = openpyxl.load_workbook(io.BytesIO(fmt_bytes))

    template_ws = out_wb.worksheets[0]
    out_wb.remove(template_ws)

    sheet_count = 0
    record_count = 0

    for src_ws in src_wb.worksheets:

        records_by_media = {}

        # ===== 施策一覧を読み取る =====
        for row in range(4, src_ws.max_row + 1):
            start = src_ws.cell(row, 2).value
            end = src_ws.cell(row, 1).value
            name = src_ws.cell(row, 3).value
            media = src_ws.title  # ←シート名を媒体扱い

            if not start or not end or not name:
                continue

            if media not in records_by_media:
                records_by_media[media] = []

            records_by_media[media].append({
                "start": start,
                "end": end,
                "name": name
            })

        # ===== 媒体ごとにシート作成 =====
        for media, records in records_by_media.items():

            ws = out_wb.copy_worksheet(template_ws)
            ws.title = media[:31]

            row_idx = 2  # 転記用の開始行想定

            for rec in records:
                ws.cell(row_idx, 1).value = rec["start"]
                ws.cell(row_idx, 2).value = rec["end"]
                ws.cell(row_idx, 3).value = rec["name"]

                row_idx += 1
                record_count += 1

            sheet_count += 1

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    return output.getvalue(), sheet_count, record_count


if input_file:
    try:
        with open(FMT_PATH, "rb") as f:
            fmt_bytes = f.read()

        output_bytes, sheet_count, record_count = build_workbook(
            input_file.getvalue(),
            fmt_bytes
        )

        st.success(
            f"""
作成完了！

媒体シート数：{sheet_count}
件数：{record_count}
"""
        )

        st.download_button(
            "ダウンロード",
            data=output_bytes,
            file_name="転記用_媒体別.xlsx"
        )

    except Exception as e:
        st.error(f"エラー：{e}")

else:
    st.info("施策一覧をアップロードしてください")
