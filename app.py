import io
from copy import copy
from datetime import datetime, date, timedelta

import openpyxl
import pandas as pd
import streamlit as st


st.set_page_config(page_title="施策一覧 転記ツール", layout="wide")

FMT_PATH = "FMT/施策一覧.xlsx"

st.title("施策一覧 転記ツール")
st.caption("転記用.xlsx から、施策一覧FMTの書式・数式・条件付き書式を維持したファイルを作成します。")

transfer_file = st.file_uploader("転記用.xlsx をアップロード", type=["xlsx"])


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


def build_preview_rows(src_wb):
    rows = []

    for src_ws in src_wb.worksheets:
        rows.append({
            "媒体": src_ws.title,
            "開始日": "",
            "終了日": "",
            "枠名": "【媒体見出し】",
        })

        for col in range(5, src_ws.max_column + 1):
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
                rows.append({
                    "媒体": src_ws.title,
                    "開始日": start_date.strftime("%Y/%-m/%-d") if hasattr(start_date, "strftime") else start_date,
                    "終了日": end_date.strftime("%Y/%-m/%-d") if hasattr(end_date, "strftime") else end_date,
                    "枠名": campaign_name,
                })

    return pd.DataFrame(rows)


def build_workbook(transfer_bytes, fmt_bytes):
    src_wb = openpyxl.load_workbook(io.BytesIO(transfer_bytes), data_only=False)
    out_wb = openpyxl.load_workbook(io.BytesIO(fmt_bytes), data_only=False)

    ws = out_wb.worksheets[0]

    all_dates = []
    for src_ws in src_wb.worksheets:
        for r in range(2, src_ws.max_row + 1):
            d = to_date(src_ws.cell(r, 1).value)
            if d:
                all_dates.append(d)

    if not all_dates:
        raise ValueError("転記用ファイルのA列に日付が見つかりませんでした。")

    target_year = min(all_dates).year
    target_month = min(all_dates).month
    yyyymm = f"{target_year}{target_month:02d}"

    ws.title = f"{target_year}年{target_month}月"

    first_day = date(target_year, target_month, 1)
    if target_month == 12:
        next_month = date(target_year + 1, 1, 1)
    else:
        next_month = date(target_year, target_month + 1, 1)

    month_dates = []
    d = first_day
    while d < next_month:
        month_dates.append(d)
        d += timedelta(days=1)

    for i, d in enumerate(month_dates, start=4):
        ws.cell(2, i).value = datetime(d.year, d.month, d.day)
        ws.cell(2, i).number_format = "yyyy/m/d"

        if ws.cell(3, i).value in (None, ""):
            ws.cell(3, i).value = f'=TEXT({ws.cell(2, i).coordinate},"aaa")'

    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    section_template_row = 4
    detail_template_row = 5
    out_row = 4

    for src_ws in src_wb.worksheets:
        if out_row > ws.max_row:
            ws.insert_rows(out_row)

        copy_row_format(ws, section_template_row, out_row)
        ws.cell(out_row, 1).value = src_ws.title
        out_row += 1

        for col in range(5, src_ws.max_column + 1):
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
                if out_row > ws.max_row:
                    ws.insert_rows(out_row)

                copy_row_format(ws, detail_template_row, out_row)

                ws.cell(out_row, 1).value = datetime(start_date.year, start_date.month, start_date.day)
                ws.cell(out_row, 2).value = datetime(end_date.year, end_date.month, end_date.day)
                ws.cell(out_row, 3).value = campaign_name

                ws.cell(out_row, 1).number_format = "yyyy/m/d"
                ws.cell(out_row, 2).number_format = "yyyy/m/d"

                out_row += 1

    if out_row <= ws.max_row:
        ws.delete_rows(out_row, ws.max_row - out_row + 1)

    ws.freeze_panes = "D4"

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    preview_df = build_preview_rows(src_wb)

    return yyyymm, output.getvalue(), preview_df


if transfer_file:
    try:
        with open(FMT_PATH, "rb") as f:
            fmt_bytes = f.read()

        yyyymm, output_bytes, preview_df = build_workbook(
            transfer_file.getvalue(),
            fmt_bytes
        )

        st.success(f"作成完了：施策一覧_{yyyymm}.xlsx")

        st.subheader("プレビュー")
        st.caption("出力ファイルに書き込まれる開始日・終了日・枠名の一覧です。")
        st.dataframe(preview_df, use_container_width=True, height=500)

        st.download_button(
            label="施策一覧ファイルをダウンロード",
            data=output_bytes,
            file_name=f"施策一覧_{yyyymm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except FileNotFoundError:
        st.error("FMT/施策一覧.xlsx が見つかりません。GitHub上に FMT フォルダを作って、その中に施策一覧.xlsx を置いてください。")

    except Exception as e:
        st.error(f"エラーが発生しました：{e}")

else:
    st.info("転記用.xlsx をアップロードしてください。")