import io
from copy import copy
from datetime import datetime, date, timedelta

import openpyxl
import streamlit as st


st.set_page_config(page_title="施策一覧 転記ツール", layout="wide")

SRC_PATH = "転記用.xlsx"

st.title("施策一覧 転記ツール")
st.caption("媒体ごとにシート分割して出力します")

fmt_file = st.file_uploader("施策一覧FMTをアップロード", type=["xlsx"])


def to_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def is_active(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip() not in ("", "0", "-", "FALSE", "false")
    if isinstance(value, (int, float)):
        return value != 0
    return False


def copy_row_format(ws, src_row, dst_row):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height

    for col in range(1, ws.max_column + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)

        if src.has_style:
            dst._style = copy(src._style)

        dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)


def continuous_ranges(date_list):
    if not date_list:
        return []

    dates = sorted(set(date_list))
    ranges = []

    start = dates[0]
    prev = dates[0]

    for d in dates[1:]:
        if d == prev + timedelta(days=1):
            prev = d
        else:
            ranges.append((start, prev))
            start = d
            prev = d

    ranges.append((start, prev))
    return ranges


def build_workbook(src_bytes, fmt_bytes):
    src_wb = openpyxl.load_workbook(io.BytesIO(src_bytes), data_only=False)
    out_wb = openpyxl.load_workbook(io.BytesIO(fmt_bytes), data_only=False)

    template_ws = out_wb.worksheets[0]

    # もとのシート削除（コピー用にだけ使う）
    out_wb.remove(template_ws)

    record_count = 0
    sheet_count = 0

    for src_ws in src_wb.worksheets:
        # ✅ テンプレコピーして新シート作成
        ws = out_wb.copy_worksheet(template_ws)
        ws.title = src_ws.title[:31]  # Excelシート名制限

        section_template_row = 4
        detail_template_row = 5
        out_row = 4

        records = []

        # データ収集
        for col in range(7, src_ws.max_column + 1):
            campaign_name = src_ws.cell(1, col).value

            if campaign_name in (None, ""):
                continue

            active_dates = []

            for r in range(2, src_ws.max_row + 1):
                current_date = to_date(src_ws.cell(r, 1).value)
                value = src_ws.cell(r, col).value

                if current_date and is_active(value):
                    active_dates.append(current_date)

            for start_date, end_date in continuous_ranges(active_dates):
                records.append({
                    "start": start_date,
                    "end": end_date,
                    "name": campaign_name
                })

        # データなければスキップ（空シート防止）
        if not records:
            out_wb.remove(ws)
            continue

        sheet_count += 1

        # 日付ソート
        records.sort(key=lambda x: x["start"])

        # 出力
        for rec in records:
            if out_row > ws.max_row:
                ws.insert_rows(out_row)

            copy_row_format(ws, detail_template_row, out_row)

            # ✅ FMTに合わせて逆（終了→開始）
            ws.cell(out_row, 1).value = rec["end"]
            ws.cell(out_row, 2).value = rec["start"]
            ws.cell(out_row, 3).value = rec["name"]

            ws.cell(out_row, 1).number_format = "yyyy/m/d"
            ws.cell(out_row, 2).number_format = "yyyy/m/d"

            out_row += 1
            record_count += 1

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    return output.getvalue(), sheet_count, record_count


if fmt_file:
    try:
        with open(SRC_PATH, "rb") as f:
            src_bytes = f.read()

        output_bytes, sheet_count, record_count = build_workbook(
            src_bytes,
            fmt_file.getvalue()
        )

        st.success(
            f"""
作成完了！

媒体シート数：{sheet_count}
出力件数：{record_count:,}件
"""
        )

        st.download_button(
            label="ダウンロード",
            data=output_bytes,
            file_name="施策一覧_媒体別.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"エラー：{e}")

else:
    st.info("施策一覧FMTをアップロードしてください")
