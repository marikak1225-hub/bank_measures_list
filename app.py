import streamlit as st
import pandas as pd
import statsmodels.api as sm
from openpyxl.styles import Border, Side
from openpyxl.styles import Border, Side, Font

st.set_page_config(page_title="転記用生成ツール", layout="wide")

@@ -127,29 +127,115 @@ def build_workbook(input_bytes, fmt_bytes, selected_sheet):


# ==============================
# 回帰分析（OK出てたシンプル版）
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

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_count = 0
    created = 0

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

@@ -160,43 +246,19 @@ def run_regression_all_sheets(input_bytes):
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
        ws = wb.create_sheet(sheet[:31])
        write_regression(ws, model)

        sheet_count += 1
        created += 1

    # 保険（エラー回避だけ）
    if sheet_count == 0:
        pd.DataFrame({"message": ["有効なデータなし"]}).to_excel(writer, sheet_name="NoData", index=False)
    if created == 0:
        ws = wb.create_sheet("NoData")
        ws["A1"] = "有効データなし"

    writer.close()
    output = io.BytesIO()
    wb.save(output)
output.seek(0)

return output.getvalue()
@@ -208,7 +270,6 @@ def run_regression_all_sheets(input_bytes):
tab1, tab2 = st.tabs(["転記用生成", "回帰分析"])


# ===== タブ1 =====
with tab1:
st.title("転記用エクセル生成ツール")

@@ -239,11 +300,8 @@ def run_regression_all_sheets(input_bytes):
file_name=f"転記用_{selected_sheet}.xlsx",
mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
    else:
        st.info("施策一覧をアップしてください")


# ===== タブ2 =====
with tab2:
st.title("回帰分析ツール")

@@ -264,3 +322,4 @@ def run_regression_all_sheets(input_bytes):
file_name=f"回帰分析結果_{today_str}.xlsx",
mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
            
