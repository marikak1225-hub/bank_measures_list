import io
from datetime import datetime, date, timedelta
import openpyxl
import streamlit as st
import pandas as pd
import statsmodels.api as sm
from openpyxl.styles import Border, Side

st.set_page_config(page_title="転記用生成ツール", layout="wide")

FMT_PATH = "転記用FMT.xlsx"

# ==============================
# 罫線
# ==============================
thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# ==============================
# 日付変換
# ==============================
def to_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.strptime(str(val), "%Y/%m/%d").date()
    except:
        return None


# ==============================
# 転記ツール（完全そのまま）
# ==============================
def parse_input(ws):
    media_blocks = {}
    current_media = None

    for row in range(1, ws.max_row + 1):
        a = ws.cell(row, 1).value
        b = ws.cell(row, 2).value
        c = ws.cell(row, 3).value

        a_date = to_date(a)
        b_date = to_date(b)

        if isinstance(a, str) and not a_date:
            current_media = a.strip()
            if current_media not in media_blocks:
                media_blocks[current_media] = []
            continue

        if current_media and a_date and b_date and c:
            media_blocks[current_media].append({
                "start": a_date,
                "end": b_date,
                "name": str(c).strip()
            })

    return media_blocks


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

    if sheet_count == 0:
        ws = out_wb.create_sheet("NoData")
        ws.cell(1, 1).value = "データを取得できませんでした"

    output = io.BytesIO()
    out_wb.save(output)
    output.seek(0)

    return output.getvalue(), sheet_count, total_days


# ==============================
# 回帰分析（OK出てたシンプル版）
# ==============================
def run_regression_all_sheets(input_bytes):
    xls = pd.ExcelFile(io.BytesIO(input_bytes))

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")

    sheet_count = 0

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)

        if df.shape[1] < 7:
            continue

        # Y = D列
        y = df.iloc[:, 3]

        # X = G列以降
        X = df.iloc[:, 6:]

        # シンプル処理（←OK出てたやつ）
        data = pd.concat([y, X], axis=1)
        data = data.apply(pd.to_numeric, errors='coerce').dropna()

        if len(data) < 3:
            continue

        y_clean = data.iloc[:, 0]
        X_clean = data.iloc[:, 1:]

        X_clean = sm.add_constant(X_clean)

        model = sm.OLS(y_clean, X_clean).fit()

        # ===== 回帰統計 =====
        stats = pd.DataFrame({
            "項目": ["重相関 R", "決定係数 R2", "補正 R2", "標準誤差", "観測数"],
            "値": [
                model.rsquared**0.5,
                model.rsquared,
                model.rsquared_adj,
                (model.mse_resid)**0.5,
                int(model.nobs)
            ]
        })

        # ===== 係数 table =====
        coef = pd.DataFrame({
            "係数": model.params,
            "標準誤差": model.bse,
            "t": model.tvalues,
            "P-値": model.pvalues,
            "下限95%": model.conf_int()[0],
            "上限95%": model.conf_int()[1]
        })

        sheet_name = sheet[:31]

        stats.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False)
        coef.to_excel(writer, sheet_name=sheet_name, startrow=10)

        sheet_count += 1

    # 保険（エラー回避だけ）
    if sheet_count == 0:
        pd.DataFrame({"message": ["有効なデータなし"]}).to_excel(writer, sheet_name="NoData", index=False)

    writer.close()
    output.seek(0)

    return output.getvalue()


# ==============================
# UI
# ==============================
tab1, tab2 = st.tabs(["転記用生成", "回帰分析"])


# ===== タブ1 =====
with tab1:
    st.title("転記用エクセル生成ツール")

    input_file = st.file_uploader("施策一覧をアップロード", type=["xlsx"])

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

            st.success(f"✅ 完成！ 媒体:{sheet_count} 日数:{total_days}")
            st.balloons()

            st.download_button(
                "ダウンロード",
                data=output,
                file_name=f"転記用_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("施策一覧をアップしてください")


# ===== タブ2 =====
with tab2:
    st.title("回帰分析ツール")

    reg_file = st.file_uploader("回帰分析用ファイルをアップロード", type=["xlsx"], key="reg")

    if reg_file:
        if st.button("回帰分析実行"):

            result = run_regression_all_sheets(reg_file.getvalue())
            today_str = datetime.now().strftime("%Y%m%d")

            st.success("✅ 回帰分析完了！")
            st.balloons()

            st.download_button(
                "ダウンロード",
                data=result,
                file_name=f"回帰分析結果_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
