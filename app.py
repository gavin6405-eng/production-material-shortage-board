from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st

APP_TITLE = "超慧科技製令缺料看板"
APP_SUBTITLE = "SPT Manufacturing Shortage Intelligence System"
DEFAULT_FILE = Path(__file__).parent / "data" / "sample_data.xlsx"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------
# 科技風 UI
# ------------------------------
st.markdown(
    """
<style>
:root {
  --bg:#061322; --panel:#092039; --panel2:#0c2542; --line:#0f86ad;
  --cyan:#20d6ee; --text:#f1f6ff; --muted:#91b5d8; --green:#19dc8b;
}
.stApp {background: radial-gradient(circle at 85% 5%, #171b50 0%, #07182a 34%, #05111f 72%); color:var(--text);}
[data-testid="stSidebar"] {background:#03111f; border-right:1px solid #133451;}
[data-testid="stSidebar"] * {color:#eaf6ff;}
.block-container {padding-top:1.05rem; padding-bottom:2rem; max-width:1600px;}
.spt-brand {border:1px solid #0c668c; border-radius:14px; padding:15px 13px; margin:8px 0 14px;
 background:linear-gradient(135deg,#08273d,#071a2c); box-shadow:inset 0 0 22px rgba(31,213,236,.05);}
.spt-brand b {font-size:14px; color:#79e9ff;} .spt-brand small {color:#60c5ed;}
.hero {border:1px solid #146b91; border-radius:18px; padding:22px 26px; margin:4px 0 20px;
 background:linear-gradient(100deg,rgba(5,48,67,.9),rgba(25,25,82,.72)); position:relative; overflow:hidden;}
.hero:after {content:"";position:absolute;inset:0;opacity:.14;background-image:linear-gradient(#3cb9db 1px,transparent 1px),linear-gradient(90deg,#3cb9db 1px,transparent 1px);background-size:25px 25px;pointer-events:none;}
.hero-inner {position:relative;z-index:1;display:flex;gap:22px;align-items:center;}
.hero-icon {width:165px;min-width:165px;height:80px;border:8px solid #f8fbff;border-radius:12px;background:linear-gradient(145deg,#292334,#0b2037);display:flex;align-items:center;justify-content:center;box-shadow:0 0 20px rgba(44,216,240,.18);}
.hero-icon span {text-align:center;font-weight:800;font-size:17px;line-height:1.25;}
.hero h1 {font-size:31px;margin:0 0 5px;color:white;letter-spacing:1px}.hero p{margin:0;color:#8dcfff;font-size:13px}
.section-title{font-size:23px;font-weight:850;margin:12px 0 2px;color:#f5f8ff}.section-note{font-size:13px;color:#8bb3d7;margin-bottom:18px}
.info-box{background:#08213a;border:1px solid #0f6889;border-left:4px solid #20d6ee;border-radius:12px;padding:14px 16px;margin:10px 0 16px;color:#d9edff;line-height:1.8}
.kpi {background:linear-gradient(150deg,#0a2841,#081b31);border:1px solid #175377;border-radius:13px;padding:14px 16px;min-height:104px;box-shadow:inset 0 0 24px rgba(25,214,238,.03)}
.kpi .label{font-size:12px;color:#8db7d9}.kpi .value{font-size:29px;font-weight:900;margin-top:5px;color:#fff}.kpi .sub{font-size:11px;color:#55d7ec;margin-top:2px}
div[data-testid="stDataFrame"] {border:1px solid #183c5b;border-radius:10px;overflow:hidden;}
.stButton>button,.stDownloadButton>button {border:1px solid #1fb6d4;border-radius:9px;background:#0a2941;color:white;font-weight:700;}
.stButton>button:hover,.stDownloadButton>button:hover {border-color:#63e9ff;color:#63e9ff;}
[data-testid="stFileUploader"] {border:1px dashed #137a9d;border-radius:12px;background:rgba(13,37,59,.68);padding:5px 10px;}
hr {border-color:#183c5b}
</style>
""",
    unsafe_allow_html=True,
)


def sidebar() -> str:
    with st.sidebar:
        st.markdown(
            '<div class="spt-brand"><b>SPT Production System</b><br><small>Manufacturing Planning Intelligence</small></div>',
            unsafe_allow_html=True,
        )
        options = [
            "00. 首頁儀表板",
            "01. 製令缺料整合",
            "02. 專案缺料總覽",
            "03. 材料品號明細",
            "04. 急料與逾期",
            "05. 資料匯入與匯出",
            "06. 專案名稱設定",
            "07. 系統說明",
        ]
        page = st.radio("功能選單", options, label_visibility="collapsed")
        st.markdown("---")
        st.caption("SPT Material & Shortage System")
        st.caption("Version 2026.07 Modular UI")
    return page


def hero() -> None:
    st.markdown(
        f"""
<div class="hero"><div class="hero-inner">
  <div class="hero-icon"><span>製令缺料<br>專案整合看板</span></div>
  <div><h1>{APP_TITLE}</h1><p>{APP_SUBTITLE}</p></div>
</div></div>
""",
        unsafe_allow_html=True,
    )


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def derive_project_code(note: object) -> str:
    """由備註擷取專案主碼，例如 22M0026-01 → 22M0026。"""
    text = normalize_text(note).upper()
    if not text:
        return "未分類"
    patterns = [
        r"(?<![A-Z0-9])(\d{2}[A-Z]{1,3}\d{4})(?:[-_/]\d+)?(?![A-Z0-9])",
        r"(?<![A-Z0-9])(\d{2}[A-Z]\d{3,5})(?:[-_/]\d+)?(?![A-Z0-9])",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    # 無法套用規則時，保留連字號前的內容，避免資料遺失
    fallback = re.split(r"[-_/]", text, maxsplit=1)[0].strip()
    return fallback or "未分類"


def find_column(columns: Iterable[str], aliases: list[str]) -> str | None:
    normalized = {str(c).strip().replace(" ", ""): c for c in columns}
    for alias in aliases:
        key = alias.strip().replace(" ", "")
        if key in normalized:
            return normalized[key]
    for c in columns:
        compact = str(c).strip().replace(" ", "")
        if any(alias.replace(" ", "") in compact for alias in aliases):
            return c
    return None


@st.cache_data(show_spinner=False)
def read_excel_file(file_bytes: bytes, sheet_name: str | int = 0, header_row: int = 0) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)


def load_source(uploaded_file, selected_sheet: str | int = 0, header_row: int = 0) -> tuple[pd.DataFrame, str]:
    if uploaded_file is not None:
        raw = uploaded_file.getvalue()
        return read_excel_file(raw, selected_sheet, header_row), uploaded_file.name
    if DEFAULT_FILE.exists():
        raw = DEFAULT_FILE.read_bytes()
        return read_excel_file(raw, selected_sheet, header_row), DEFAULT_FILE.name
    return pd.DataFrame(), "尚未上傳"


def prepare_data(df: pd.DataFrame, project_mapping: dict[str, str] | None = None) -> tuple[pd.DataFrame, dict[str, str]]:
    if df.empty:
        empty_columns = [
            "材料品號", "品名", "規格", "製令編號", "急料", "欠料數量",
            "現有庫存", "逾期未入", "欠料包裝數量", "現有包裝庫存",
            "備註", "逾期包裝未入", "專案代碼", "專案名稱",
            "急料判定", "逾期判定"
        ]
        return pd.DataFrame(columns=empty_columns), {}
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    aliases = {
        "材料品號": ["材料品號", "料號", "物料編號", "品號"],
        "品名": ["品名", "材料名稱", "物料名稱"],
        "規格": ["規格", "規格型號", "型號"],
        "製令編號": ["製令編號", "製令", "工單編號", "工單"],
        "急料": ["急料", "急件", "緊急"],
        "欠料數量": ["欠料數量", "缺料數量", "短缺數量", "欠數"],
        "現有庫存": ["現有庫存", "庫存數量", "庫存"],
        "逾期未入": ["逾期未入", "逾期數量", "未入數量"],
        "欠料包裝數量": ["欠料包裝數量", "缺包裝數量"],
        "現有包裝庫存": ["現有包裝庫存", "包裝庫存"],
        "備註": ["備註", "註記", "專案", "專案編號"],
        "逾期包裝未入": ["逾期包裝未入", "包裝逾期未入"],
    }
    detected: dict[str, str] = {}
    for canonical, names in aliases.items():
        found = find_column(df.columns, names)
        if found is not None:
            detected[canonical] = found

    required = ["材料品號", "備註"]
    missing = [x for x in required if x not in detected]
    if missing:
        raise ValueError(f"找不到必要欄位：{', '.join(missing)}。請確認 Excel 至少包含『材料品號』與『備註』。")

    # 建立標準欄位，未提供的選用欄位以空值補齊
    result = pd.DataFrame(index=df.index)
    for col in aliases:
        result[col] = df[detected[col]] if col in detected else ""

    text_cols = ["材料品號", "品名", "規格", "製令編號", "急料", "備註"]
    for col in text_cols:
        result[col] = result[col].map(normalize_text)
    numeric_cols = ["欠料數量", "現有庫存", "逾期未入", "欠料包裝數量", "現有包裝庫存", "逾期包裝未入"]
    for col in numeric_cols:
        result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0)

    result = result[result["材料品號"] != ""].copy()
    result["專案代碼"] = result["備註"].map(derive_project_code)
    mapping = project_mapping or {}
    result["專案名稱"] = result["專案代碼"].map(mapping).fillna(result["專案代碼"])
    result["急料判定"] = result["急料"].str.upper().isin(["Y", "YES", "是", "急", "1", "TRUE"])
    result["逾期判定"] = (result["逾期未入"] > 0) | (result["逾期包裝未入"] > 0)
    return result.reset_index(drop=True), detected


def project_summary(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame()
    s = (
        data.groupby(["專案代碼", "專案名稱"], dropna=False)
        .agg(
            材料品項數=("材料品號", "nunique"),
            製令數=("製令編號", lambda x: x[x != ""].nunique()),
            欠料總數=("欠料數量", "sum"),
            現有庫存=("現有庫存", "sum"),
            逾期未入=("逾期未入", "sum"),
            急料筆數=("急料判定", "sum"),
            原始筆數=("材料品號", "size"),
        )
        .reset_index()
    )
    return s.sort_values(["急料筆數", "逾期未入", "欠料總數"], ascending=False).reset_index(drop=True)


def combine_unique(series: pd.Series, limit: int = 30) -> str:
    vals = [normalize_text(v) for v in series if normalize_text(v)]
    vals = list(dict.fromkeys(vals))
    if len(vals) > limit:
        return "、".join(vals[:limit]) + f"…(+{len(vals)-limit})"
    return "、".join(vals)


def integrated_materials(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame()
    group_cols = ["專案代碼", "專案名稱", "材料品號", "品名", "規格"]
    output = (
        data.groupby(group_cols, dropna=False)
        .agg(
            製令數=("製令編號", lambda x: x[x != ""].nunique()),
            製令清單=("製令編號", combine_unique),
            欠料數量=("欠料數量", "sum"),
            現有庫存=("現有庫存", "max"),
            逾期未入=("逾期未入", "sum"),
            欠料包裝數量=("欠料包裝數量", "sum"),
            現有包裝庫存=("現有包裝庫存", "max"),
            逾期包裝未入=("逾期包裝未入", "sum"),
            急料=("急料判定", "max"),
            原始備註=("備註", combine_unique),
            原始筆數=("材料品號", "size"),
        )
        .reset_index()
    )
    output["風險等級"] = "一般"
    output.loc[output["逾期未入"] > 0, "風險等級"] = "逾期"
    output.loc[output["急料"], "風險等級"] = "急料"
    output.loc[output["急料"] & (output["逾期未入"] > 0), "風險等級"] = "急料且逾期"
    return output.sort_values(["急料", "逾期未入", "欠料數量"], ascending=False).reset_index(drop=True)


def kpi(label: str, value: str, sub: str = "") -> None:
    st.markdown(
        f'<div class="kpi"><div class="label">{label}</div><div class="value">{value}</div><div class="sub">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def make_excel(raw: pd.DataFrame, integrated: pd.DataFrame, summary: pd.DataFrame, mapping: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="專案總覽", index=False)
        integrated.to_excel(writer, sheet_name="專案料號整合", index=False)
        raw.to_excel(writer, sheet_name="標準化明細", index=False)
        mapping.to_excel(writer, sheet_name="專案名稱對照", index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="0B5C78")
            for col in sheet.columns:
                letter = col[0].column_letter
                width = min(max(len(str(c.value or "")) for c in col) + 2, 38)
                sheet.column_dimensions[letter].width = max(width, 10)
    return output.getvalue()


page = sidebar()
hero()

# 共用資料來源
with st.expander("📥 資料來源與讀取設定", expanded=page in ["01. 製令缺料整合", "05. 資料匯入與匯出"]):
    uploaded = st.file_uploader("上傳製令缺料 Excel", type=["xlsx", "xls"], key="main_upload")
    c1, c2 = st.columns([3, 1])
    with c1:
        header_row = st.number_input("標題列（Excel 列號）", min_value=1, max_value=30, value=1, step=1)
    with c2:
        st.caption("系統預設使用第 1 個工作表")

try:
    base_df, source_name = load_source(uploaded, 0, int(header_row) - 1)
except Exception as exc:
    st.error(f"Excel 讀取失敗：{exc}")
    st.stop()

if base_df.empty:
    st.warning("目前尚未讀取到缺料資料。請展開上方『資料來源與讀取設定』並上傳製令缺料 Excel。")
    st.info("Excel 至少需要包含『材料品號』與『備註』欄位；備註中的 22M0026-01 會自動整合為專案代碼 22M0026。")
    st.stop()

# 先建立代碼，再提供名稱對照編輯
try:
    preliminary, detected_cols = prepare_data(base_df)
except ValueError as exc:
    st.error(str(exc))
    st.info("目前讀取到的欄位：" + "、".join(map(str, base_df.columns)))
    st.stop()

codes = sorted(preliminary["專案代碼"].dropna().unique().tolist())
if "project_map_df" not in st.session_state or set(st.session_state.project_map_df.get("專案代碼", [])) != set(codes):
    old = {}
    if "project_map_df" in st.session_state:
        old = dict(zip(st.session_state.project_map_df["專案代碼"], st.session_state.project_map_df["專案名稱"]))
    st.session_state.project_map_df = pd.DataFrame({"專案代碼": codes, "專案名稱": [old.get(c, c) for c in codes]})

mapping_dict = dict(zip(st.session_state.project_map_df["專案代碼"], st.session_state.project_map_df["專案名稱"]))
data, detected_cols = prepare_data(base_df, mapping_dict)
summary = project_summary(data)
integrated = integrated_materials(data)

# 全域篩選
if not data.empty and page not in ["06. 專案名稱設定", "07. 系統說明"]:
    with st.expander("🔎 全域篩選", expanded=False):
        f1, f2, f3 = st.columns([2, 2, 1])
        project_options = sorted(data["專案名稱"].unique().tolist())
        selected_projects = f1.multiselect("專案名稱", project_options, placeholder="全部專案")
        material_keyword = f2.text_input("材料品號／品名關鍵字")
        only_risk = f3.checkbox("只看急料或逾期")
    filtered = data.copy()
    if selected_projects:
        filtered = filtered[filtered["專案名稱"].isin(selected_projects)]
    if material_keyword:
        mask = (
            filtered["材料品號"].str.contains(material_keyword, case=False, na=False)
            | filtered["品名"].str.contains(material_keyword, case=False, na=False)
            | filtered["規格"].str.contains(material_keyword, case=False, na=False)
        )
        filtered = filtered[mask]
    if only_risk:
        filtered = filtered[filtered["急料判定"] | filtered["逾期判定"]]
    filtered_summary = project_summary(filtered)
    filtered_integrated = integrated_materials(filtered)
else:
    filtered, filtered_summary, filtered_integrated = data, summary, integrated

if page == "00. 首頁儀表板":
    st.markdown('<div class="section-title">00. 首頁儀表板</div><div class="section-note">快速掌握專案、材料品號、急料與逾期風險。</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="info-box"><b>目前資料：</b>{source_name}<br><b>專案判定：</b>由「備註」自動擷取主碼，例如 22M0026-01 → 22M0026；同專案、同材料品號會合併呈現。</div>', unsafe_allow_html=True)
    cols = st.columns(5)
    with cols[0]: kpi("專案數", f"{data['專案代碼'].nunique():,}", "依備註主碼")
    with cols[1]: kpi("材料品項", f"{data['材料品號'].nunique():,}", "不重複品號")
    with cols[2]: kpi("欠料總數", f"{data['欠料數量'].sum():,.0f}", "全部缺料需求")
    with cols[3]: kpi("急料筆數", f"{data['急料判定'].sum():,.0f}", "Y／是／急")
    with cols[4]: kpi("逾期筆數", f"{data['逾期判定'].sum():,.0f}", "逾期未入 > 0")

    a, b = st.columns([1.25, 1])
    with a:
        st.subheader("專案欠料排行")
        top = summary.head(15)
        if not top.empty:
            fig = px.bar(top.sort_values("欠料總數"), x="欠料總數", y="專案名稱", orientation="h", hover_data=["材料品項數", "急料筆數", "逾期未入"])
            fig.update_layout(height=430, margin=dict(l=10, r=10, t=20, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#dcecff")
            st.plotly_chart(fig, use_container_width=True)
    with b:
        st.subheader("高風險專案")
        st.dataframe(summary.head(15), use_container_width=True, hide_index=True, height=430)

elif page == "01. 製令缺料整合":
    st.markdown('<div class="section-title">01. 製令缺料整合</div><div class="section-note">將備註中的專案主碼抽出，並把相同專案、相同材料品號整合成一列。</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box"><b>整合邏輯</b><br>備註 22M0026-01、22M0026-02 → 專案代碼 22M0026<br>專案代碼＋材料品號＋品名＋規格相同 → 欠料數量加總，製令與原始備註合併顯示。</div>', unsafe_allow_html=True)
    st.success(f"已讀取：{source_name}｜原始資料 {len(data):,} 筆｜整合後 {len(filtered_integrated):,} 筆｜辨識專案 {data['專案代碼'].nunique():,} 個")
    st.dataframe(
        filtered_integrated,
        use_container_width=True,
        hide_index=True,
        height=610,
        column_config={
            "急料": st.column_config.CheckboxColumn("急料"),
            "欠料數量": st.column_config.NumberColumn(format="%.2f"),
            "逾期未入": st.column_config.NumberColumn(format="%.2f"),
        },
    )

elif page == "02. 專案缺料總覽":
    st.markdown('<div class="section-title">02. 專案缺料總覽</div><div class="section-note">以專案為核心彙總材料品項、製令數、欠料、急料及逾期。</div>', unsafe_allow_html=True)
    st.dataframe(filtered_summary, use_container_width=True, hide_index=True, height=580)
    if not filtered_summary.empty:
        fig = px.scatter(filtered_summary, x="材料品項數", y="欠料總數", size="逾期未入", color="急料筆數", hover_name="專案名稱")
        fig.update_layout(height=430, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#dcecff")
        st.plotly_chart(fig, use_container_width=True)

elif page == "03. 材料品號明細":
    st.markdown('<div class="section-title">03. 材料品號明細</div><div class="section-note">查詢某一材料品號分布在哪些專案與製令。</div>', unsafe_allow_html=True)
    keyword = st.text_input("輸入材料品號、品名或規格", placeholder="例如 #051-00830-00")
    detail = data.copy()
    if keyword:
        detail = detail[
            detail["材料品號"].str.contains(keyword, case=False, na=False)
            | detail["品名"].str.contains(keyword, case=False, na=False)
            | detail["規格"].str.contains(keyword, case=False, na=False)
        ]
    st.caption(f"符合 {len(detail):,} 筆")
    st.dataframe(detail, use_container_width=True, hide_index=True, height=630)

elif page == "04. 急料與逾期":
    st.markdown('<div class="section-title">04. 急料與逾期</div><div class="section-note">集中顯示需要優先追蹤的急料與逾期未入項目。</div>', unsafe_allow_html=True)
    risk = filtered_integrated[(filtered_integrated["急料"]) | (filtered_integrated["逾期未入"] > 0) | (filtered_integrated["逾期包裝未入"] > 0)]
    r1, r2, r3 = st.columns(3)
    with r1: kpi("風險材料品項", f"{len(risk):,}", "整合後筆數")
    with r2: kpi("急料品項", f"{risk['急料'].sum():,}", "需要優先處理")
    with r3: kpi("逾期未入", f"{risk['逾期未入'].sum():,.0f}", "逾期數量合計")
    st.dataframe(risk, use_container_width=True, hide_index=True, height=620)

elif page == "05. 資料匯入與匯出":
    st.markdown('<div class="section-title">05. 資料匯入與匯出</div><div class="section-note">確認欄位辨識結果，並下載整合後 Excel。</div>', unsafe_allow_html=True)
    st.success(f"已讀取工作表資料：{source_name}｜資料筆數：{len(data):,}")
    with st.expander("查看系統辨識到的欄位", expanded=True):
        st.json(detected_cols)
    export_bytes = make_excel(data, integrated, summary, st.session_state.project_map_df)
    st.download_button("⬇️ 下載專案整合缺料表.xlsx", data=export_bytes, file_name="專案整合缺料表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.subheader("原始資料預覽")
    st.dataframe(base_df.head(100), use_container_width=True, hide_index=True)

elif page == "06. 專案名稱設定":
    st.markdown('<div class="section-title">06. 專案名稱設定</div><div class="section-note">專案代碼會自動建立；可在此將代碼改為更容易理解的專案名稱。</div>', unsafe_allow_html=True)
    st.info("例如：專案代碼 22M0026，可把專案名稱改成『Rorze 新增設備專案』。修改後其他頁面會同步更新。")
    edited = st.data_editor(
        st.session_state.project_map_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=["專案代碼"],
        key="mapping_editor",
    )
    if st.button("套用專案名稱", use_container_width=True):
        edited["專案名稱"] = edited["專案名稱"].fillna(edited["專案代碼"]).astype(str).str.strip()
        edited.loc[edited["專案名稱"] == "", "專案名稱"] = edited["專案代碼"]
        st.session_state.project_map_df = edited
        st.success("專案名稱已套用，請切換至其他頁面查看。")
        st.rerun()

elif page == "07. 系統說明":
    st.markdown('<div class="section-title">07. 系統說明</div><div class="section-note">部署方式與資料規則。</div>', unsafe_allow_html=True)
    st.markdown(
        """
<div class="info-box">
<b>必要欄位</b>：材料品號、備註。<br>
<b>建議欄位</b>：品名、規格、製令編號、急料、欠料數量、現有庫存、逾期未入。<br>
<b>專案規則</b>：系統從備註擷取如 22M0026、26M0090、25MN0014 等主碼，尾端 -01、-02 視為同專案。<br>
<b>執行方式</b>：安裝 requirements.txt 後執行 <code>streamlit run app.py</code>。<br>
<b>Streamlit Cloud</b>：將 app.py、requirements.txt、data 資料夾上傳 GitHub，Main file path 選 app.py。
</div>
""",
        unsafe_allow_html=True,
    )
