import io
from datetime import datetime, date, timedelta
import openpyxl
import streamlit as st

st.set_page_config(page_title="転記用生成ツール", layout="wide")

FMT_PATH = "転記用FMT.xlsx"

st.title("転記用エクセル生成ツール")

input_file = st.file_uploader("施策一覧をアップロード", type=["xlsx"])


# ===== 日付変換 =====
def to_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.strptime(str(val), "%Y/%m/%d").date()
    except:
        return None


# ===== 超柔軟パース =====
def parse_input(ws):
    media_blocks = {}
    current_media = None

    for row in range(1, ws.max_row + 1):
        row_values = [ws.cell(row, col).value for col in range(1, 10)]

        # === 媒体行検出 ===
        # 「1列目に文字があり、他が空」は媒体とみなす
        if isinstance(row_values[0], str) and all(v is None for v in row_values[1:4]):
            current_media = row_values[0].strip()

            if current_media not in media_blocks:
                media_blocks[current_media] = []
            continue

        # === データ行検出 ===
        if current_media:
            dates = []
            name = None

            # 行の中から日付2つと文字1つを探す
            for val in row_values:
                d = to_date(val)
                if d:
                    dates.append(d)
                elif isinstance(val, str) and val.strip():
                    name = val.strip()

            if len(dates) >= 2 and name:
                start = min(dates)
                end = max(dates)

                media_blocks[current_media].append({
                    "start": start,
                    "end": end,
                    "name": name
                })

    return media_blocks


# ===== メイン =====
def build_workbook(input_bytes, fmt_bytes, selected_sheet):
    src_wb = openpyxl.load_workbook(io.BytesIO(input_bytes))
    src_ws = src_wb[selected_sheet]

    out_wb = openpyxl.load_workbook(io.BytesIO(fmt_bytes))
    template_ws = out_wb.worksheets[0]
    out_wb.remove(template_ws)

    media_blocks = parse_input(src_ws)

    # ✅ デバッグ（1回だけ表示）
    st.write("DEBUG:", media_blocks)

    total_days = 0
    sheet_count = 0

    for media, records in media_blocks.items():

        if not records:
            continue

        ws = out_wb.copy_worksheet(template_ws)
        ws.title = media[:31]
        sheet_count += 1

        campaigns = sorted(set(r["name"] for r in records))

        for i, name in enumerate(campaigns):
            ws.cell(row=1, column=7 + i).value = name

        min_date = min(r["start"] for r in records)
        max_date = max(r["end"] for r in records)

        current = min_date
        row = 2

        while current <= max_date:
            ws.cell(row=row, column=1).value = current
            ws.cell(row=row, column=1).number_format = "yyyy/mm/dd"

            for i, name in enumerate(campaigns):
                flag = 0
                for r in records:
                    if r["name"] == name and r["start"] <= current <= r["end"]:
                        flag = 1
                        break

                ws.cell(row=row, column=7 + i).value = flag

            current += timedelta(days=1)
            row += 1
            total_days += 1

    # ✅ 保険
    if sheet_count == 0:
        ws = out_wb.create_sheet("NoData")
        ws.cell(1, 1).value = "データを取得できませんでした"

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    return output.getvalue(), sheet_count, total_days


# ===== UI =====
if input_file:
    wb = openpyxl.load_workbook(input_file)
    sheet_names = wb.sheetnames

    selected_sheet = st.selectbox("対象シートを選択", sheet_names)

    if st.button("実行"):
        with open(FMT_PATH, "rb") as f:
            fmt_bytes = f.read()

        output, sheet_count, total_days = build_workbook(
            input_file.getvalue(),
            fmt_bytes,
            selected_sheet
        )

        st.success(f"媒体数：{sheet_count} / 日数：{total_days}")
        st.balloons()

        st.download_button(
            "ダウンロード",
            data=output,
            file_name="転記用_媒体別.xlsx"
        )

else:
    st.info("施策一覧をアップしてください")
