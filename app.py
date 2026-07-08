# -*- coding: utf-8 -*-
"""
超慧科技｜每日機台庫存容量看板

計算邏輯：
- 發料日 / 發料日期 / 預計發料日 = 進，庫存 +
- 入庫日 / 入庫日期 / 出機日 / 出機日期 / 排程入庫日 = 出，庫存 -
- 每日庫存 = 前一日期末庫存 + 當日進 - 當日出

啟動方式：
1. 安裝套件：pip install streamlit pandas openpyxl altair numpy
2. 執行：streamlit run app.py
"""

from __future__ import annotations

from io import BytesIO
from typing import Optional, Tuple, List

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

APP_TITLE = "超慧科技｜每日機台庫存容量看板"
DEFAULT_CAPACITY = 82


# =========================================================
# 資料處理工具
# =========================================================
def clean_col_name(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().replace("\n", "").replace(" ", "")


def normalize_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
    """支援 Excel 日期序號、文字日期、datetime，並排除 TBD/空白。"""
    s = series.copy()
    s = s.replace({"TBD": np.nan, "tbd": np.nan, "待定": np.nan, "": np.nan, "-": np.nan, "nan": np.nan})

    numeric = pd.to_numeric(s, errors="coerce")
    out = pd.to_datetime(s, errors="coerce")

    # Excel serial date 常見範圍約 30000~60000
    mask = numeric.between(30000, 60000)
    if mask.any():
        out.loc[mask] = pd.to_datetime(numeric.loc[mask], unit="D", origin="1899-12-30", errors="coerce")

    return out.dt.normalize()


def find_best_sheet_and_header(uploaded_file) -> Tuple[str, int, List[str]]:
    """自動找出最像資料表的 Sheet 與表頭列。"""
    xls = pd.ExcelFile(uploaded_file)
    best = None
    keywords = ["發料", "製令", "入庫", "出機"]

    for sheet in xls.sheet_names:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, dtype=object)
        scan_rows = min(len(raw), 50)
        for i in range(scan_rows):
            row_values = [clean_col_name(v) for v in raw.iloc[i].tolist()]
            text = "|".join(row_values)
            score = sum(1 for k in keywords if k in text)
            if score >= 2:
                non_empty_cols = [v for v in row_values if v]
                candidate = (score, sheet, i, non_empty_cols)
                if best is None or candidate[0] > best[0]:
                    best = candidate

    if best is None:
        raise ValueError("找不到含有『發料、製令、入庫/出機』欄位的表頭列，請確認 Excel 格式。")

    _, sheet, header_row, cols = best
    return sheet, header_row, cols


def read_source_table(uploaded_file) -> pd.DataFrame:
    sheet, header_row, _ = find_best_sheet_and_header(uploaded_file)
    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row, dtype=object)
    df.columns = [clean_col_name(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith("Unnamed")]].copy()
    df = df.dropna(how="all").copy()
    df["來源Sheet"] = sheet
    return df


def find_column(df: pd.DataFrame, candidates: List[str], required: bool = True) -> Optional[str]:
    cols = list(df.columns)

    # 完全符合優先
    for cand in candidates:
        for col in cols:
            if cand == col:
                return col

    # 包含文字次之
    for cand in candidates:
        for col in cols:
            if cand in col:
                return col

    if required:
        raise ValueError(f"找不到必要欄位：{candidates}")
    return None


def prepare_events(df: pd.DataFrame, qty_mode: str, qty_col: Optional[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    issue_col = find_column(df, ["發料日期", "發料日", "預計發料日"])
    out_col = find_column(df, ["入庫日", "入庫日期", "出機日", "出機日期", "排程入庫日"])
    order_col = find_column(df, ["製令", "工單", "製令單號", "訂單", "訂單號"], required=False)
    place_col = find_column(df, ["組立地點", "組裝地點", "地點", "廠區"], required=False)
    model_col = find_column(df, ["機型", "機種", "品名", "產品"], required=False)

    work = df.copy()
    work["發料日_進"] = excel_serial_to_datetime(work[issue_col])
    work["入庫日_出"] = excel_serial_to_datetime(work[out_col])

    if qty_mode == "固定每筆=1台" or not qty_col:
        work["計算台數"] = 1
    else:
        work["計算台數"] = pd.to_numeric(work[qty_col], errors="coerce").fillna(1)
        work.loc[work["計算台數"] <= 0, "計算台數"] = 1

    work["製令"] = work[order_col].map(normalize_text) if order_col else ""
    work["組立地點"] = work[place_col].map(normalize_text) if place_col else "未指定"
    work["機型"] = work[model_col].map(normalize_text) if model_col else "未指定"
    work = work[(work["發料日_進"].notna()) | (work["入庫日_出"].notna())].copy()

    events = []
    for _, r in work.iterrows():
        qty = float(r["計算台數"])
        base = {
            "製令": r.get("製令", ""),
            "組立地點": r.get("組立地點", "未指定"),
            "機型": r.get("機型", "未指定"),
            "計算台數": qty,
            "原始發料日欄位": issue_col,
            "原始入庫/出機欄位": out_col,
        }
        if pd.notna(r["發料日_進"]):
            events.append({**base, "日期": r["發料日_進"], "類型": "進", "進": qty, "出": 0.0, "異動": qty})
        if pd.notna(r["入庫日_出"]):
            events.append({**base, "日期": r["入庫日_出"], "類型": "出", "進": 0.0, "出": qty, "異動": -qty})

    event_df = pd.DataFrame(events)
    if event_df.empty:
        raise ValueError("沒有可計算的日期資料，請確認發料日與入庫日/出機日。")

    event_df = event_df.sort_values(["日期", "類型", "製令"]).reset_index(drop=True)
    return work, event_df


def status_by_usage(usage: float) -> str:
    if pd.isna(usage):
        return "無容量"
    if usage <= 0.80:
        return "綠燈"
    if usage <= 0.90:
        return "黃燈"
    if usage <= 0.95:
        return "橘燈"
    if usage <= 1.00:
        return "紅燈"
    return "超載"


def action_by_status(status: str) -> str:
    return {
        "綠燈": "正常運作",
        "黃燈": "提前規劃移機或入庫",
        "橘燈": "空間高負載，需主管確認",
        "紅燈": "接近滿載，需安排出機/調整場地",
        "超載": "超過容量，需立即處理",
    }.get(status, "請確認")


def build_daily_summary(event_df: pd.DataFrame, capacity: int, initial_stock: int) -> pd.DataFrame:
    min_date = event_df["日期"].min()
    max_date = event_df["日期"].max()
    date_index = pd.date_range(min_date, max_date, freq="D")

    grouped = event_df.groupby("日期", as_index=True).agg({"進": "sum", "出": "sum", "異動": "sum"})
    daily = pd.DataFrame(index=date_index).join(grouped, how="left").fillna(0).reset_index()
    daily = daily.rename(columns={"index": "日期"})
    daily["期初庫存"] = initial_stock + daily["異動"].cumsum().shift(1).fillna(0)
    daily["期末庫存"] = initial_stock + daily["異動"].cumsum()
    daily["容量上限"] = capacity
    daily["使用率"] = np.where(capacity > 0, daily["期末庫存"] / capacity, np.nan)
    daily["管理燈號"] = daily["使用率"].apply(status_by_usage)
    daily["管理說明"] = daily["管理燈號"].apply(action_by_status)

    for col in ["進", "出", "異動", "期初庫存", "期末庫存"]:
        daily[col] = daily[col].round(0).astype(int)

    return daily


def build_monthly_summary(event_df: pd.DataFrame) -> pd.DataFrame:
    monthly = event_df.copy()
    monthly["月份"] = monthly["日期"].dt.strftime("%Y-%m")
    return monthly.groupby("月份", as_index=False).agg({"進": "sum", "出": "sum", "異動": "sum"})


def to_excel_bytes(daily: pd.DataFrame, events: pd.DataFrame, detail: pd.DataFrame, capacity: int, initial_stock: int) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kpi = pd.DataFrame({
            "項目": ["容量上限", "初始庫存", "最大庫存", "最大使用率", "超載天數", "計算起日", "計算迄日"],
            "數值": [
                capacity,
                initial_stock,
                int(daily["期末庫存"].max()),
                f"{daily['使用率'].max():.1%}",
                int((daily["期末庫存"] > capacity).sum()),
                daily["日期"].min().strftime("%Y-%m-%d"),
                daily["日期"].max().strftime("%Y-%m-%d"),
            ],
        })
        kpi.to_excel(writer, sheet_name="KPI摘要", index=False)
        daily.to_excel(writer, sheet_name="每日庫存彙總", index=False)
        events.to_excel(writer, sheet_name="進出事件明細", index=False)
        detail.to_excel(writer, sheet_name="原始資料整理", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 32)

    return output.getvalue()


# =========================================================
# Streamlit 畫面
# =========================================================
st.set_page_config(page_title=APP_TITLE, page_icon="🏭", layout="wide")

st.markdown(
    """
    <style>
    .main {background: linear-gradient(180deg, #f7fbff 0%, #ffffff 45%);} 
    .block-container {padding-top: 1.8rem; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #d9e6f2;
        padding: 16px;
        border-radius: 14px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title(APP_TITLE)
st.caption("發料日＝進（+）｜入庫日/出機日＝出（-）｜每日累計機台庫存與容量使用率")

with st.sidebar:
    st.header("參數設定")
    capacity = st.number_input("容量上限（台）", min_value=1, value=DEFAULT_CAPACITY, step=1)
    initial_stock = st.number_input("計算起日前初始庫存（台）", min_value=0, value=0, step=1)
    uploaded = st.file_uploader("上傳機台容量管理 Excel", type=["xlsx", "xls"])
    st.divider()
    st.info("建議先使用『固定每筆=1台』。若原始資料有明確台數欄位，再改用指定台數欄位。")

if uploaded is None:
    st.warning("請先上傳 Excel 檔案。")
    st.stop()

try:
    df = read_source_table(uploaded)

    numeric_candidates = []
    for c in df.columns:
        if any(k in c for k in ["數量", "台數", "QTY", "Qty", "qty"]):
            if "容量" not in c and "使用率" not in c:
                numeric_candidates.append(c)

    with st.sidebar:
        qty_mode = st.radio("每筆計算台數", ["固定每筆=1台", "使用指定台數欄位"], index=0)
        qty_col = None
        if qty_mode == "使用指定台數欄位":
            qty_col = st.selectbox("選擇台數欄位", numeric_candidates or list(df.columns))

    detail_df, events_df = prepare_events(df, qty_mode, qty_col)
    daily_df = build_daily_summary(events_df, int(capacity), int(initial_stock))
    monthly_df = build_monthly_summary(events_df)

    max_stock = int(daily_df["期末庫存"].max())
    max_usage = float(daily_df["使用率"].max())
    overload_days = int((daily_df["期末庫存"] > capacity).sum())
    final_stock = int(daily_df.iloc[-1]["期末庫存"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("最大庫存", f"{max_stock} 台")
    c2.metric("最大使用率", f"{max_usage:.1%}")
    c3.metric("超載天數", f"{overload_days} 天")
    c4.metric("最終庫存", f"{final_stock} 台")

    st.subheader("每日庫存趨勢")
    line = alt.Chart(daily_df).mark_line(point=True).encode(
        x=alt.X("日期:T", title="日期"),
        y=alt.Y("期末庫存:Q", title="每日期末庫存"),
        tooltip=[
            alt.Tooltip("日期:T", title="日期"),
            alt.Tooltip("進:Q", title="進"),
            alt.Tooltip("出:Q", title="出"),
            alt.Tooltip("異動:Q", title="異動"),
            alt.Tooltip("期末庫存:Q", title="期末庫存"),
            alt.Tooltip("使用率:Q", title="使用率", format=".1%"),
            alt.Tooltip("管理燈號:N", title="燈號"),
        ],
    ).properties(height=360)

    cap_line = alt.Chart(pd.DataFrame({"容量上限": [capacity]})).mark_rule(strokeDash=[6, 4]).encode(y="容量上限:Q")
    st.altair_chart(line + cap_line, use_container_width=True)

    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("每日庫存彙總")
        show_daily = daily_df.copy()
        show_daily["日期"] = show_daily["日期"].dt.strftime("%Y-%m-%d")
        st.dataframe(show_daily.style.format({"使用率": "{:.1%}"}), use_container_width=True, height=360)

    with right:
        st.subheader("月別進出彙總")
        st.dataframe(monthly_df, use_container_width=True, height=220)

        bar = alt.Chart(monthly_df).transform_fold(
            ["進", "出"], as_=["類別", "台數"]
        ).mark_bar().encode(
            x=alt.X("月份:N", title="月份"),
            y=alt.Y("台數:Q", title="台數"),
            color=alt.Color("類別:N", title="類別"),
            tooltip=["月份:N", "類別:N", "台數:Q"],
        ).properties(height=260)
        st.altair_chart(bar, use_container_width=True)

    st.subheader("進出事件明細")
    show_events = events_df.copy()
    show_events["日期"] = show_events["日期"].dt.strftime("%Y-%m-%d")
    st.dataframe(show_events, use_container_width=True, height=300)

    output_bytes = to_excel_bytes(daily_df, events_df, detail_df, int(capacity), int(initial_stock))
    st.download_button(
        "下載每日庫存彙整 Excel",
        data=output_bytes,
        file_name="每日機台庫存容量彙整.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.error("計算失敗，請確認 Excel 欄位是否包含發料日/發料日期與入庫日/出機日。")
    st.exception(e)
