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


# ===== 入力解析 =====
def parse_input(ws):
    media_blocks = {}
    current_media = None

    for row in range(1, ws.max_row + 1):
        a = ws.cell(row, 1).value
        b = ws.cell(row, 2).value
        c = ws.cell(row, 3).value
        d = ws.cell(row, 4).value

        # --- 媒体判定 ---
        if isinstance(a, str) and a.strip() != "" and not to_date(a):
            current_media = a.strip()
            if current_media not in media_blocks:
                media_blocks[current_media] = []
            continue

        # --- 空行スキップ ---
        if not any([a, b, c, d]):
            continue

        # --- データ行 ---
        if current_media:
            start = to_date(b)
            end = to_date(c)
            name = d

            if start and end and name:
                media_blocks[current_media].append({
                    "start": start,
                    "end": end,
                    "name": str(name).strip()
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

    total_days = 0
    sheet_count = 0

    for media, records in media_blocks.items():

        # ✅ 空データ防止
        if not records:
            continue

        ws = out_wb.copy_worksheet(template_ws)
        ws.title = media[:31]
        sheet_count += 1

        # ===== 施策一覧 =====
        campaigns = sorted(set(r["name"] for r in records))

        for i, name in enumerate(campaigns):
            ws.cell(row=1, column=7 + i).value = name

        # ===== 日付範囲 =====
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

    # ✅ シート0防止（超重要）
    if sheet_count == 0:
        ws = out_wb.create_sheet(title="NoData")
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
        try:
            with open(FMT_PATH, "rb") as f:
                fmt_bytes = f.read()

            output, sheet_count, total_days = build_workbook(
                input_file.getvalue(),
                fmt_bytes,
                selected_sheet
            )

            st.success(
                f"""
✅ 完成！

媒体数：{sheet_count}
日数：{total_days}
"""
            )

            st.balloons()

            st.download_button(
                "ダウンロード",
                data=output,
                file_name="転記用_媒体別.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"エラー：{e}")

else:
    st.info("施策一覧をアップしてください")
