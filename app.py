import io
from datetime import datetime, timedelta
import openpyxl
import streamlit as st

st.set_page_config(page_title="転記用生成ツール", layout="wide")

FMT_PATH = "転記用FMT.xlsx"

st.title("転記用エクセル生成ツール")

input_file = st.file_uploader("施策一覧をアップロード", type=["xlsx"])

def parse_input(ws):
    media_blocks = {}
    current_media = None

    for row in range(1, ws.max_row + 1):
        a = ws.cell(row, 1).value
        b = ws.cell(row, 2).value
        c = ws.cell(row, 3).value
        d = ws.cell(row, 4).value

        # 媒体判定
        if isinstance(a, str) and a not in ["開始日", None]:
            current_media = a
            if current_media not in media_blocks:
                media_blocks[current_media] = []

        # 施策行
        elif current_media and isinstance(b, datetime) and isinstance(d, datetime):
            media_blocks[current_media].append({
                "start": b.date(),
                "end": d.date(),
                "name": c
            })

    return media_blocks


def build_workbook(input_bytes, fmt_bytes, selected_sheet):
    src_wb = openpyxl.load_workbook(io.BytesIO(input_bytes))
    src_ws = src_wb[selected_sheet]

    out_wb = openpyxl.load_workbook(io.BytesIO(fmt_bytes))
    template_ws = out_wb.worksheets[0]
    out_wb.remove(template_ws)

    media_blocks = parse_input(src_ws)

    total_count = 0

    for media, records in media_blocks.items():

        ws = out_wb.copy_worksheet(template_ws)
        ws.title = media[:31]

        # === 施策ユニーク抽出 ===
        campaigns = sorted(list(set(r["name"] for r in records)))

        # === G列以降ヘッダ ===
        for i, name in enumerate(campaigns):
            ws.cell(row=1, column=7 + i).value = name

        # === 日付範囲 ===
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
            total_count += 1

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    return output.getvalue(), total_count


if input_file:
    wb = openpyxl.load_workbook(input_file)
    sheet_names = wb.sheetnames

    selected_sheet = st.selectbox("対象シートを選択", sheet_names)

    if st.button("実行"):
        try:
            with open(FMT_PATH, "rb") as f:
                fmt_bytes = f.read()

            output, count = build_workbook(
                input_file.getvalue(),
                fmt_bytes,
                selected_sheet
            )

            st.success(f"✅ 作成完了！ {count}日分生成")

            # 🎈復活
            st.balloons()

            st.download_button(
                "ダウンロード",
                data=output,
                file_name="転記用_媒体別.xlsx"
            )

        except Exception as e:
            st.error(f"エラー：{e}")

else:
    st.info("施策一覧をアップロードしてください")
