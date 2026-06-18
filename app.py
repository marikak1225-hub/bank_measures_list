import io
from datetime import datetime, date, timedelta
import openpyxl
import streamlit as st
import pandas as pd
import statsmodels.api as sm
from openpyxl.styles import Border, Side, Font

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


def build_workbook(input_bytes, fmt_bytes, selected_sheets):
    src_wb = openpyxl.load_workbook(io.BytesIO(input_bytes), data_only=True)

    out_wb = openpyxl.load_workbook(io.BytesIO(fmt_bytes))
    template_ws = out_wb.worksheets[0]
    out_wb.remove(template_ws)

    # 複数シート分を、媒体名ごとにマージする
    media_blocks = {}

    for selected_sheet in selected_sheets:
        src_ws = src_wb[selected_sheet]
        sheet_media_blocks = parse_input(src_ws)

        for media, records in sheet_media_blocks.items():
            if media not in media_blocks:
                media_blocks[media] = []
            media_blocks[media].extend(records)

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
# 回帰分析（見た目完全寄せ版）
# ==============================
def write_regression(ws, model):

    ws["A1"] = "概要"

    # ===== 回帰統計 =====
    ws["A3"] = "回帰統計"
    stats = [
        ("重相関 R", model.rsquared**0.5),
        ("決定係数 R2", model.rsquared),
        ("補正 R2", model.rsquared_adj),
        ("標準誤差", (model.mse_resid)**0.5),
        ("観測数", int(model.nobs))
    ]

    row = 4
    for name, val in stats:
        ws.cell(row, 1).value = name
        ws.cell(row, 2).value = float(val)
        ws.cell(row, 1).border = thin
        ws.cell(row, 2).border = thin
        row += 1

    # ===== 分散分析 =====
    ws["A10"] = "分散分析表"

    headers = ["", "自由度", "変動", "分散", "F"]
    for col, h in enumerate(headers, 1):
        ws.cell(11, col, h).border = thin

    y = model.model.endog
    y_mean = y.mean()

    ss_total = ((y - y_mean) ** 2).sum()
    ss_resid = (model.resid ** 2).sum()
    ss_reg = ss_total - ss_resid

    df_reg = model.df_model
    df_resid = model.df_resid

    ms_reg = ss_reg / df_reg if df_reg else 0
    ms_resid = ss_resid / df_resid if df_resid else 0
    f_val = ms_reg / ms_resid if ms_resid else 0

    anova = [
        ("回帰", df_reg, ss_reg, ms_reg, f_val),
        ("残差", df_resid, ss_resid, ms_resid, ""),
        ("合計", df_reg + df_resid, ss_total, "", "")
    ]

    row = 12
    for r in anova:
        for col, v in enumerate(r, 1):
            ws.cell(row, col, v).border = thin
        row += 1

    # ===== 係数 =====
    headers = ["", "係数", "標準誤差", "t", "P-値", "下限95%", "上限95%"]
    start = 16

    for col, h in enumerate(headers, 1):
        ws.cell(start, col, h).border = thin

    params = model.params
    bse = model.bse
    tvals = model.tvalues
    pvals = model.pvalues
    conf = model.conf_int()

    row = start + 1

    for i in range(len(params)):
        name = params.index[i]
        label = "切片" if name == "const" else str(name)

        values = [
            label,
            params.iloc[i],
            bse.iloc[i],
            tvals.iloc[i],
            pvals.iloc[i],
            conf.iloc[i, 0],
            conf.iloc[i, 1]
        ]

        for col, v in enumerate(values, 1):
            ws.cell(row, col, v).border = thin

        row += 1


def run_regression_all_sheets(input_bytes):
    xls = pd.ExcelFile(io.BytesIO(input_bytes))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    created = 0

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)

        if df.shape[1] < 7:
            continue

        y = df.iloc[:, 3]
        X = df.iloc[:, 6:]

        data = pd.concat([y, X], axis=1)
        data = data.apply(pd.to_numeric, errors='coerce').dropna()

        if len(data) < 3:
            continue

        y_clean = data.iloc[:, 0]
        X_clean = data.iloc[:, 1:]

        X_clean = sm.add_constant(X_clean)
        model = sm.OLS(y_clean, X_clean).fit()

        ws = wb.create_sheet(sheet[:31])
        write_regression(ws, model)

        created += 1

    if created == 0:
        ws = wb.create_sheet("NoData")
        ws["A1"] = "有効データなし"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# ==============================
# UI
# ==============================
tab1, tab2 = st.tabs(["転記用生成", "回帰分析"])


with tab1:
    st.title("転記用エクセル生成ツール")

    input_file = st.file_uploader("施策一覧をアップロード", type=["xlsx"])

    if input_file:
        wb = openpyxl.load_workbook(input_file)
        sheet_names = wb.sheetnames

        selected_sheets = st.multiselect(
            "対象シートを選択（複数選択可）",
            sheet_names,
            default=[sheet_names[0]] if sheet_names else []
        )

        if st.button("実行"):
            if not selected_sheets:
                st.warning("対象シートを1つ以上選択してください")
            else:
                with open(FMT_PATH, "rb") as f:
                    fmt_bytes = f.read()

                output, sheet_count, total_days = build_workbook(
                    input_file.getvalue(),
                    fmt_bytes,
                    selected_sheets
                )

                st.success(f"✅ 完成！ 媒体:{sheet_count} 日数:{total_days}")
                st.balloons()

                today_str = datetime.now().strftime("%Y%m%d")
                st.download_button(
                    "ダウンロード",
                    data=output,
                    file_name=f"転記用_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


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
