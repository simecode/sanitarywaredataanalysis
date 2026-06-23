import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import os
import re
import numpy as np
import warnings
warnings.filterwarnings('ignore')

MAPPING_FILE_NAME = "区域映射表.xlsx"

st.set_page_config(page_title="卫浴行业数据观察智库", layout="wide")
st.title("📊 卫浴与泛家居进出口多维洞察大屏")

# ================= 侧边栏 =================
with st.sidebar:
    st.header("⚙️ 引擎配置")
    analysis_mode = st.radio("选择分析维度", ["年度全景统计", "月度/前N月动态"])
    if analysis_mode == "月度/前N月动态":
        n_months = st.slider("统计前N个月", 1, 12, 4)
    else:
        n_months = None

    st.markdown("---")
    st.header("🔑 AI 智库接入")
    openrouter_key = st.text_input("OpenRouter API Key", type="password")
    ai_model = st.selectbox("推理模型", ["openai/gpt-oss-120b:free", "deepseek/deepseek-chat:free"])

    st.markdown("---")
    st.header("📂 数据上传")
    raw_files = st.file_uploader("海关原始数据（支持多选Excel）", type=['xlsx'], accept_multiple_files=True)
    mapping_upload = st.file_uploader("区域映射表（可选）", type=['xlsx'])


# ================= 字段识别配置 =================

# 系统需要的字段及候选列名（中英文混合）
FIELD_SPECS = {
    "贸易伙伴名称": {
        "label": "目的地国家/地区 *",
        "required": True,
        "candidates": ["贸易伙伴名称","贸易伙伴名称称","贸易伙伴名","国家/地区","国家地区","目的地","country","Country","COUNTRY",
                       "Partner","partner","Trade Partner","destination","Destination","Country/Region"],
    },
    "金额_美元": {
        "label": "出口金额（美元）*",
        "required": True,
        "candidates": ["金额_美元","美元","出口金额_美元","金额（美元）","金额(美元)","USD","Amount_USD","Value_USD",
                       "Export Value","export_value","金额","Amount","Value","usd_value","usd_amount"],
    },
    "注册地名称": {
        "label": "注册地/省份",
        "required": False,
        "candidates": ["注册地名称","注册地","省份","出口地区","企业注册地","Province","province","Region","region",
                       "Registered Region","Origin Province"],
    },
    "贸易类型": {
        "label": "贸易类型（出口/进口）",
        "required": False,
        "candidates": ["贸易类型","进出口类型","贸易方式","Trade Type","trade_type","Type","type","Direction"],
    },
    "数量_统一": {
        "label": "数量/重量（用于计算单价）",
        "required": False,
        "candidates": ["数量","第一法定数量","统计数量","法定数量","净重（千克）","总重量（千克）","重量",
                       "Quantity","quantity","Weight","weight","Net Weight","net_weight","Qty","qty"],
    },
    "数据年月": {
        "label": "年月（月度模式用）",
        "required": False,
        "candidates": ["数据年月","年月","统计月份","月份","报告期","Year Month","YearMonth","Date","date","Period"],
    },
    "统计年份": {
        "label": "统计年份",
        "required": False,
        "candidates": ["统计年份","年份","Year","year","YEAR","fiscal_year","统计年度"],
    },
}

def guess_column(df_cols, candidates):
    """从候选列名中找最优匹配（精确优先，模糊兜底）"""
    cols_lower = {c.lower(): c for c in df_cols}
    for cand in candidates:
        if cand in df_cols:
            return cand
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    # 模糊匹配：候选关键词 in 列名
    for keyword in ["美元","usd","amount","value","partner","country","province","region","qty","weight","year","month"]:
        for col in df_cols:
            if keyword in col.lower():
                for cand in candidates:
                    if keyword in cand.lower():
                        return col
    return None

def extract_year_from_df(df, year_col, filename=""):
    """从年份列或文件名中提取四位年份"""
    if year_col and year_col in df.columns:
        vals = df[year_col].dropna().astype(str).str.strip()
        valid = vals.str.extract(r'(20\d{2})')[0].dropna()
        if not valid.empty:
            return valid.mode()[0]
    m = re.search(r'(20\d{2})', filename)
    if m:
        return m.group(1)
    return None


# ================= 主流程 =================

if not raw_files:
    st.info("👈 请在左侧上传海关原始数据，系统将自动识别字段并引导你完成配置。")
    st.stop()

# --- Step 1: 读取所有文件表头，取并集 ---
all_columns_union = set()
file_cols_map = {}  # filename -> list of columns
file_bytes_map = {}  # filename -> bytes (for re-read after mapping)

for f in raw_files:
    fname = f.name
    try:
        preview = pd.read_excel(f, nrows=3)
        cols = list(preview.columns)
        file_cols_map[fname] = cols
        all_columns_union.update(cols)
        f.seek(0)
        file_bytes_map[fname] = f.read()
    except Exception as e:
        st.warning(f"⚠️ {fname} 读取失败：{e}")

if not file_cols_map:
    st.error("所有文件读取失败，请检查文件格式。")
    st.stop()

all_cols_list = sorted(all_columns_union)

# --- Step 2: 字段映射 UI ---
st.markdown("## 第一步：确认字段映射")
st.caption("系统已自动识别建议列名，请检查并按需修改。未找到的字段选「— 不使用 —」。")

# 用第一个文件的列名做猜测基准（最常见情况）
ref_cols = list(list(file_cols_map.values())[0])

col_mapping = {}  # field_key -> selected_col or None
options_with_none = ["— 不使用 —"] + all_cols_list

mapping_cols = st.columns(2)
field_keys = list(FIELD_SPECS.keys())

for i, field_key in enumerate(field_keys):
    spec = FIELD_SPECS[field_key]
    guess = guess_column(ref_cols, spec["candidates"])
    default_idx = (all_cols_list.index(guess) + 1) if (guess and guess in all_cols_list) else 0

    with mapping_cols[i % 2]:
        chosen = st.selectbox(
            label=spec["label"],
            options=options_with_none,
            index=default_idx,
            key=f"map_{field_key}"
        )
        col_mapping[field_key] = None if chosen == "— 不使用 —" else chosen

# 检查必填字段
missing_required = [FIELD_SPECS[k]["label"] for k, v in col_mapping.items()
                    if FIELD_SPECS[k]["required"] and v is None]
if missing_required:
    st.error(f"以下必填字段未映射：{', '.join(missing_required)}")

# 数据预览
with st.expander("🔍 预览第一个文件（前5行）"):
    first_fname = list(file_cols_map.keys())[0]
    prev_df = pd.read_excel(io.BytesIO(file_bytes_map[first_fname]), nrows=5)
    st.dataframe(prev_df, use_container_width=True)

st.markdown("---")

# --- Step 3: 贸易类型值配置 ---
trade_col = col_mapping.get("贸易类型")
export_keyword = "出口"
if trade_col:
    st.markdown("### 贸易类型筛选配置")
    # 采样实际值
    first_df_sample = pd.read_excel(io.BytesIO(file_bytes_map[list(file_bytes_map.keys())[0]]), nrows=200)
    if trade_col in first_df_sample.columns:
        sample_vals = first_df_sample[trade_col].dropna().astype(str).unique().tolist()
        st.caption(f"该列实际值示例：{sample_vals[:8]}")
        export_keyword = st.text_input(
            "「出口」的标识值（与上方示例对应）",
            value="出口" if "出口" in sample_vals else (sample_vals[0] if sample_vals else "出口"),
            help="如数据是英文，可填 Export 或 E 等实际值"
        )
    st.markdown("---")

# --- Step 4: 开始分析 ---
if missing_required:
    st.stop()

if st.button("🚀 开始分析", type="primary"):

    # ---- 区域映射表 ----
    region_dict = {}
    if mapping_upload:
        try:
            mapping_upload.seek(0)
            map_df = pd.read_excel(mapping_upload, sheet_name="区域映射")
            region_dict = dict(zip(map_df["原始名称"].str.strip(), map_df["子区域"].str.strip().fillna("其他")))
        except:
            pass
    if not region_dict and os.path.exists(MAPPING_FILE_NAME):
        try:
            map_df = pd.read_excel(MAPPING_FILE_NAME, sheet_name="区域映射")
            region_dict = dict(zip(map_df["原始名称"].str.strip(), map_df["子区域"].str.strip().fillna("其他")))
        except:
            pass

    # ---- 读取并标准化所有文件 ----
    all_data = []
    parse_log = []

    for fname, file_bytes in file_bytes_map.items():
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))

            # 按映射重命名列
            rename_map = {}
            for field_key, src_col in col_mapping.items():
                if src_col and src_col in df.columns and src_col != field_key:
                    rename_map[src_col] = field_key
            df = df.rename(columns=rename_map)

            # 提取年份
            year_str = extract_year_from_df(df, "统计年份" if "统计年份" in df.columns else None, fname)
            if not year_str:
                parse_log.append(f"⚠️ {fname}：无法识别年份，跳过")
                continue
            df["统计年份"] = year_str

            # 月度筛选
            if analysis_mode == "月度/前N月动态" and n_months and "数据年月" in df.columns:
                df["数据年月"] = df["数据年月"].astype(str).str.strip()
                suffixes = tuple(str(i).zfill(2) for i in range(1, n_months + 1))
                before = len(df)
                df = df[df["数据年月"].str[-2:].isin(suffixes)]
                parse_log.append(f"✅ {fname}（{year_str}）：前{n_months}月筛选 {before}→{len(df)}行")
            else:
                parse_log.append(f"✅ {fname}（{year_str}）：{len(df)}行")

            # 贸易类型过滤
            if "贸易类型" in df.columns:
                df = df[df["贸易类型"].astype(str).str.strip() == export_keyword]
            # 若无贸易类型列，默认全部为出口
            df["贸易类型"] = "出口"

            # 数量/重量统一
            if "数量_统一" in df.columns:
                df["数量_统一"] = pd.to_numeric(df["数量_统一"], errors="coerce").fillna(0)
            else:
                df["数量_统一"] = 0

            # 注册地兜底
            if "注册地名称" not in df.columns:
                df["注册地名称"] = "未知"

            # 清洗
            df = df.dropna(subset=["贸易伙伴名称", "金额_美元"])
            df["金额_美元"] = pd.to_numeric(df["金额_美元"], errors="coerce").fillna(0)
            df["贸易伙伴名称"] = df["贸易伙伴名称"].astype(str).str.strip()
            df["所属区域"] = df["贸易伙伴名称"].map(region_dict).fillna("其他")

            keep = [c for c in ["贸易伙伴名称","注册地名称","贸易类型","金额_美元","数量_统一","统计年份","所属区域","数据年月"] if c in df.columns]
            all_data.append(df[keep])

        except Exception as e:
            parse_log.append(f"❌ {fname}：{str(e)[:80]}")

    with st.expander("📋 文件解析日志"):
        for log in parse_log:
            st.text(log)

    if not all_data:
        st.error("没有成功解析任何数据，请检查字段映射。")
        st.stop()

    export_df = pd.concat(all_data, ignore_index=True)
    years = sorted(export_df["统计年份"].unique())
    latest_year = years[-1]

    # ---- 计算函数 ----
    def calc_metrics(df, group_cols):
        res = df.groupby(group_cols + ["统计年份"], as_index=False).agg({"金额_美元": "sum", "数量_统一": "sum"})
        res = res.sort_values(group_cols + ["统计年份"])
        res["出口单价（美元/单位）"] = np.where(res["数量_统一"] > 0, (res["金额_美元"] / res["数量_统一"]).round(4), np.nan)
        if group_cols:
            res["上期金额"] = res.groupby(group_cols)["金额_美元"].shift(1)
            res["上期单价"] = res.groupby(group_cols)["出口单价（美元/单位）"].shift(1)
        else:
            res["上期金额"] = res["金额_美元"].shift(1)
            res["上期单价"] = res["出口单价（美元/单位）"].shift(1)
        res["金额同比%"] = ((res["金额_美元"] - res["上期金额"]) / res["上期金额"].replace(0, np.nan) * 100).round(2)
        res["单价同比%"] = ((res["出口单价（美元/单位）"] - res["上期单价"]) / res["上期单价"].replace(0, np.nan) * 100).round(2)
        yearly_total = res.groupby("统计年份")["金额_美元"].transform("sum")
        res["金额份额%"] = (res["金额_美元"] / yearly_total * 100).round(2)
        return res

    annual_sum      = calc_metrics(export_df, [])
    partner_sum     = calc_metrics(export_df, ["贸易伙伴名称", "所属区域"])
    latest_partner  = partner_sum[partner_sum["统计年份"] == latest_year].sort_values("金额_美元", ascending=False)
    partner_top20   = latest_partner.head(20).copy()
    partner_top20["排名"] = range(1, len(partner_top20) + 1)
    partner_top20["累计份额%"] = partner_top20["金额份额%"].cumsum().round(2)

    has_price = export_df["数量_统一"].sum() > 0
    if has_price:
        high_val = latest_partner.head(40).sort_values("出口单价（美元/单位）", ascending=False).head(10).copy()
        high_val["高附加值排名"] = range(1, len(high_val) + 1)
    else:
        high_val = pd.DataFrame()

    province_sum   = calc_metrics(export_df, ["注册地名称"])
    prov_top10     = province_sum[province_sum["统计年份"] == latest_year].sort_values("金额_美元", ascending=False).head(10)
    region_sum     = calc_metrics(export_df, ["所属区域"])
    latest_region  = region_sum[region_sum["统计年份"] == latest_year]

    # ---- 展示 ----
    st.success(f"✅ 解析完成！年份范围：{' / '.join(years)}，最新年份：**{latest_year}**")
    st.markdown(f"### 🌐 {latest_year} 全球市场格局透视")

    latest_annual = annual_sum[annual_sum["统计年份"] == latest_year]
    ca, cb, cc = st.columns(3)
    with ca:
        st.metric(f"{latest_year} 出口总额（美元）", f"${latest_annual['金额_美元'].sum():,.0f}")
    with cb:
        yoy = latest_annual["金额同比%"].values[0] if not latest_annual.empty else None
        st.metric("同比增长", f"{yoy:.2f}%" if (yoy is not None and not np.isnan(float(yoy))) else "—")
    with cc:
        n_countries = export_df[export_df["统计年份"] == latest_year]["贸易伙伴名称"].nunique()
        st.metric("出口目的地数", n_countries)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"🏆 {latest_year} 前十大出口目的地")
        hover_cols = {c: True for c in ["金额份额%","金额同比%","累计份额%"] if c in partner_top20.columns}
        fig1 = px.bar(partner_top20.head(10), x="贸易伙伴名称", y="金额_美元",
                      color="所属区域", hover_data=hover_cols, text_auto='.2s')
        fig1.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(fig1, use_container_width=True)

        if not high_val.empty:
            st.subheader(f"💎 {latest_year} 高附加值市场 TOP10")
            fig3 = px.scatter(high_val, x="金额_美元", y="出口单价（美元/单位）",
                              size="金额_美元", color="所属区域", hover_name="贸易伙伴名称", size_max=40)
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("未映射数量/重量字段，无法渲染单价视图。")

    with col2:
        st.subheader(f"🏭 {latest_year} 核心出口省份 TOP10")
        fig2 = px.pie(prov_top10, names="注册地名称", values="金额_美元", hole=0.4)
        fig2.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig2, use_container_width=True)

        st.subheader(f"🌍 {latest_year} 大区同比波动")
        if "金额同比%" in latest_region.columns:
            fig4 = px.bar(latest_region, x="所属区域", y="金额同比%", color="所属区域", text_auto='.2f')
            fig4.add_hline(y=0, line_dash="dash", line_color="gray")
            st.plotly_chart(fig4, use_container_width=True)

    if len(years) > 1:
        st.markdown("---")
        st.subheader("📈 历年出口总额趋势")
        fig_trend = px.line(annual_sum, x="统计年份", y="金额_美元", markers=True, text="金额_美元")
        fig_trend.update_traces(texttemplate='$%{text:,.0f}', textposition='top center')
        st.plotly_chart(fig_trend, use_container_width=True)

    st.markdown("---")
    st.subheader(f"📋 {latest_year} 各国 TOP20 明细")
    show_cols = [c for c in ["排名","贸易伙伴名称","所属区域","金额_美元","金额份额%","累计份额%","金额同比%","出口单价（美元/单位）","单价同比%"] if c in partner_top20.columns]
    st.dataframe(partner_top20[show_cols].reset_index(drop=True), use_container_width=True)

    # ---- Excel 导出 ----
    st.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="清洗后完整数据", index=False)
        annual_sum.to_excel(writer, sheet_name="年度总览_同比", index=False)
        partner_sum.to_excel(writer, sheet_name="贸易伙伴_明细", index=False)
        partner_top20.to_excel(writer, sheet_name=f"{latest_year}各国TOP20", index=False)
        if not high_val.empty:
            high_val.to_excel(writer, sheet_name=f"{latest_year}高附加值TOP10", index=False)
        province_sum.to_excel(writer, sheet_name="注册地_明细", index=False)
        prov_top10.to_excel(writer, sheet_name="省份TOP10", index=False)
        region_sum.to_excel(writer, sheet_name="区域汇总_同比", index=False)

    st.download_button("📥 下载全维数据底稿", data=output.getvalue(),
                       file_name=f"多维洞察底稿_{latest_year}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ---- AI 撰稿 ----
    st.header("✍️ AI 智库洞察报告")
    if openrouter_key:
        ctx  = f"数据年份：{', '.join(years)}\n"
        ctx += f"年度异动：\n{annual_sum.tail(3).to_string()}\n\n"
        ctx += f"核心市场TOP5：\n{partner_top20[['贸易伙伴名称','金额_美元','金额份额%','金额同比%']].head(5).to_string()}\n\n"
        if not high_val.empty:
            ctx += f"高附加值市场：\n{high_val[['贸易伙伴名称','出口单价（美元/单位）','金额_美元']].head(3).to_string()}"

        system_prompt = """你是一位资深的泛家居与卫浴行业市场观察者，主理专业媒体矩阵。
根据海关出口多维数据撰写约1500字智库分析文章。
要求：1.保持"行业观察者"的深刻与专业；2.将数据变化与产业带位移、欧美去库存周期、高附加值市场突围建立逻辑关联；3.含标题、核心摘要、三个层级清晰的段落。"""

        with st.spinner("AI 推演中..."):
            try:
                resp = requests.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers={"Authorization": f"Bearer {openrouter_key}", "HTTP-Referer": "https://github.com/"},
                    json={"model": ai_model, "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"核心数据：\n{ctx}\n请开始撰写："}
                    ]},
                    timeout=90
                )
                if resp.status_code == 200:
                    st.success("✨ 报告生成完毕！")
                    st.text_area("媒体发布底稿", resp.json()['choices'][0]['message']['content'], height=600)
                else:
                    st.error(f"调用失败：{resp.status_code} - {resp.text}")
            except Exception as e:
                st.error(f"请求失败：{e}")
    else:
        st.warning("请在左侧输入 API Key 以启动 AI 撰稿。")
