import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import os
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ================= 核心配置区 =================
MAPPING_FILE_NAME = "区域映射表.xlsx"

# ================= 页面配置 =================
st.set_page_config(page_title="卫浴行业数据观察智库", layout="wide")
st.title("📊 卫浴与泛家居进出口多维洞察大屏")

# ================= 侧边栏配置 =================
with st.sidebar:
    st.header("⚙️ 引擎配置")
    analysis_mode = st.radio("选择分析维度", ["年度全景统计", "月度/前N月动态（如前三月）"])

    if analysis_mode == "月度/前N月动态（如前三月）":
        n_months = st.slider("选择统计前N个月", min_value=1, max_value=12, value=4)
    else:
        n_months = None

    st.markdown("---")
    st.header("🔑 AI 智库模型接入")
    openrouter_key = st.text_input("OpenRouter API Key", type="password", help="在此输入你的专属密钥（以 sk-or-v1 开头）")
    ai_model = st.selectbox("选择推理模型", ["openai/gpt-oss-120b:free", "deepseek/deepseek-chat:free"])

    st.markdown("---")
    st.header("📂 原始数据上传")
    st.info("💡 提示：区域映射表已在系统后台内置，直接上传海关明细表即可。")
    raw_files = st.file_uploader("请拖入海关原始数据 (支持多选Excel)", type=['xlsx'], accept_multiple_files=True)

    st.markdown("---")
    st.header("🗺️ 区域映射表（可选）")
    mapping_file = st.file_uploader("上传区域映射表（可选，覆盖内置表）", type=['xlsx'])


# ================= 智能字段识别工具函数 =================

def detect_year(df, filename=""):
    """
    智能提取年份：优先从「统计年份」列读取真实年份值，
    其次从文件名中提取，避免把日期字符串的前4位当年份。
    """
    if "统计年份" in df.columns:
        raw = df["统计年份"].dropna().astype(str).str.strip()
        # 尝试找到合法的4位年份（2000-2099）
        valid = raw.str.extract(r'(20\d{2})')[0].dropna()
        if not valid.empty:
            return valid.iloc[0]

    # 从文件名中提取年份
    import re
    m = re.search(r'(20\d{2})', filename)
    if m:
        return m.group(1)

    return None


def detect_amount_col(df):
    """智能识别金额列（美元）"""
    candidates = ["金额_美元", "美元", "出口金额_美元", "金额（美元）", "金额(美元)", "USD金额"]
    for c in candidates:
        if c in df.columns:
            return c
    # 模糊匹配
    for c in df.columns:
        if "美元" in c and "金额" in c:
            return c
        if c.upper() in ["USD", "AMOUNT_USD"]:
            return c
    return None


def detect_qty_col(df):
    """智能识别数量列"""
    candidates = ["数量", "第一法定数量", "统计数量", "法定数量", "净重（千克）", "总重量（千克）", "重量"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def detect_weight_col(df):
    """智能识别重量/净重列"""
    candidates = ["总重量（千克）", "净重（千克）", "重量（千克）", "净重", "总重量", "重量"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def detect_partner_col(df):
    """智能识别贸易伙伴名称列（参考水龙头脚本的兼容写法）"""
    candidates = ["贸易伙伴名称", "贸易伙伴名称称", "贸易伙伴名", "国家/地区", "国家地区", "目的地"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def detect_province_col(df):
    """智能识别注册地/省份列"""
    candidates = ["注册地名称", "注册地", "省份", "出口地区", "企业注册地"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def detect_trade_type_col(df):
    """智能识别贸易类型列"""
    candidates = ["贸易类型", "进出口类型", "贸易方式"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def detect_month_col(df):
    """智能识别月份/年月列"""
    candidates = ["数据年月", "年月", "统计月份", "月份", "报告期"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def detect_product_code_col(df):
    """智能识别商品编码列"""
    candidates = ["商品编码", "HS编码", "HS Code", "产品编码", "编码"]
    for c in candidates:
        if c in df.columns:
            return c
    return None


# ================= 核心多维处理引擎（不缓存，避免年份错乱）=================

def process_data(raw_files, mode, mapping_file_obj=None, n_months=None):
    # ---- 1. 读取区域映射表 ----
    region_dict = {}
    if mapping_file_obj is not None:
        try:
            map_df = pd.read_excel(mapping_file_obj, sheet_name="区域映射")
            region_dict = dict(zip(
                map_df["原始名称"].str.strip(),
                map_df["子区域"].str.strip().fillna("其他")
            ))
            st.sidebar.success("✅ 已使用上传的区域映射表")
        except Exception as e:
            st.sidebar.warning(f"⚠️ 上传映射表解析失败：{e}，尝试内置映射表")

    if not region_dict and os.path.exists(MAPPING_FILE_NAME):
        try:
            map_df = pd.read_excel(MAPPING_FILE_NAME, sheet_name="区域映射")
            region_dict = dict(zip(
                map_df["原始名称"].str.strip(),
                map_df["子区域"].str.strip().fillna("其他")
            ))
        except:
            pass

    # ---- 2. 逐文件智能读取 ----
    all_data = []
    parse_log = []

    for file in raw_files:
        filename = file.name if hasattr(file, 'name') else str(file)
        try:
            df = pd.read_excel(file)

            # —— 智能字段识别 ——
            partner_col   = detect_partner_col(df)
            amount_col    = detect_amount_col(df)
            qty_col       = detect_qty_col(df)
            weight_col    = detect_weight_col(df)
            province_col  = detect_province_col(df)
            trade_col     = detect_trade_type_col(df)
            month_col     = detect_month_col(df)
            code_col      = detect_product_code_col(df)

            if partner_col is None or amount_col is None:
                parse_log.append(f"⚠️ {filename}：未识别到「贸易伙伴名称」或「金额_美元」列，已跳过")
                continue

            # —— 统一列名 ——
            rename_map = {}
            if partner_col != "贸易伙伴名称":
                rename_map[partner_col] = "贸易伙伴名称"
            if amount_col != "金额_美元":
                rename_map[amount_col] = "金额_美元"
            if province_col and province_col != "注册地名称":
                rename_map[province_col] = "注册地名称"
            if trade_col and trade_col != "贸易类型":
                rename_map[trade_col] = "贸易类型"
            if month_col and month_col != "数据年月":
                rename_map[month_col] = "数据年月"
            if code_col and code_col != "商品编码":
                rename_map[code_col] = "商品编码"
            if qty_col and qty_col not in ["数量", "金额_美元"]:
                rename_map[qty_col] = "数量_原始"
            if weight_col and weight_col not in list(rename_map.keys()) + ["总重量（千克）"]:
                rename_map[weight_col] = "总重量（千克）"

            df = df.rename(columns=rename_map)

            # —— 年份提取（核心修复：从列值中提取真实年份）——
            year_str = detect_year(df, filename)
            if not year_str:
                parse_log.append(f"⚠️ {filename}：无法识别年份，已跳过")
                continue
            df["统计年份"] = year_str

            # —— 月度模式：筛选前N月 ——
            if mode == "月度/前N月动态（如前三月）" and n_months:
                if "数据年月" in df.columns:
                    df["数据年月"] = df["数据年月"].astype(str).str.strip()
                    month_suffixes = tuple(str(i).zfill(2) for i in range(1, n_months + 1))
                    before = len(df)
                    df = df[df["数据年月"].str[-2:].isin(month_suffixes)]
                    parse_log.append(f"✅ {filename}（{year_str}年）：月度筛选前{n_months}月，{before}行 → {len(df)}行")
                else:
                    parse_log.append(f"⚠️ {filename}：未找到月份列，按全量数据处理")

            # —— 贸易类型处理：若无此列，默认全部视为「出口」 ——
            if "贸易类型" not in df.columns:
                df["贸易类型"] = "出口"

            # —— 数量/重量统一 ——
            if "数量_原始" in df.columns:
                df["数量_统一"] = pd.to_numeric(df["数量_原始"], errors="coerce").fillna(0)
            elif "总重量（千克）" in df.columns:
                df["数量_统一"] = pd.to_numeric(df["总重量（千克）"], errors="coerce").fillna(0)
            else:
                df["数量_统一"] = 0

            if "总重量（千克）" not in df.columns:
                df["总重量（千克）"] = df["数量_统一"]

            # —— 注册地兜底 ——
            if "注册地名称" not in df.columns:
                df["注册地名称"] = "未知"

            # —— 数据清洗 ——
            df = df.dropna(subset=["贸易伙伴名称", "金额_美元"])
            df["金额_美元"] = pd.to_numeric(df["金额_美元"], errors="coerce").fillna(0)
            df["贸易伙伴名称"] = df["贸易伙伴名称"].astype(str).str.strip()
            df["所属区域"] = df["贸易伙伴名称"].map(region_dict).fillna("其他")

            # —— 选取保留列 ——
            keep = ["商品编码", "贸易伙伴名称", "注册地名称", "贸易类型",
                    "金额_美元", "数量_统一", "总重量（千克）", "统计年份", "所属区域"]
            if "数据年月" in df.columns:
                keep.append("数据年月")
            df = df[[c for c in keep if c in df.columns]]

            all_data.append(df)
            if mode != "月度/前N月动态（如前三月）":
                parse_log.append(f"✅ {filename}（{year_str}年）：{len(df)}行")

        except Exception as e:
            parse_log.append(f"❌ {filename} 处理失败：{str(e)[:80]}")

    # 显示解析日志
    with st.expander("📋 文件解析日志", expanded=False):
        for log in parse_log:
            st.text(log)

    if not all_data:
        st.error("❌ 没有成功解析任何数据文件，请检查文件格式和列名。")
        return None

    all_df = pd.concat(all_data, ignore_index=True)

    # 仅保留出口数据用于分析
    export_df = all_df[all_df["贸易类型"] == "出口"].copy()

    if export_df.empty:
        # 若贸易类型字段全为"出口"（兜底填充），则用全量
        export_df = all_df.copy()

    # 获取统计年度列表（从数据中读取，不硬编码）
    years = sorted(export_df["统计年份"].unique())
    latest_year = years[-1] if years else "未知"

    # ================= 核心计算函数 =================
    def calc_metrics(df, group_cols):
        agg_dict = {"金额_美元": "sum", "数量_统一": "sum", "总重量（千克）": "sum"}
        agg_dict = {k: v for k, v in agg_dict.items() if k in df.columns}

        res = df.groupby(group_cols + ["统计年份"], as_index=False).agg(agg_dict)
        res = res.sort_values(group_cols + ["统计年份"])

        # 单价（优先用重量，其次用数量）
        weight_col_name = "总重量（千克）" if "总重量（千克）" in res.columns else "数量_统一"
        if weight_col_name in res.columns:
            res["出口单价（美元/千克）"] = np.where(
                res[weight_col_name] > 0,
                (res["金额_美元"] / res[weight_col_name]).round(4),
                np.nan
            )

        # 同比
        if group_cols:
            res["上期金额"] = res.groupby(group_cols)["金额_美元"].shift(1)
            if "出口单价（美元/千克）" in res.columns:
                res["上期单价"] = res.groupby(group_cols)["出口单价（美元/千克）"].shift(1)
        else:
            res["上期金额"] = res["金额_美元"].shift(1)
            if "出口单价（美元/千克）" in res.columns:
                res["上期单价"] = res["出口单价（美元/千克）"].shift(1)

        res["金额同比%"] = ((res["金额_美元"] - res["上期金额"]) / res["上期金额"].replace(0, np.nan) * 100).round(2)
        if "上期单价" in res.columns:
            res["单价同比%"] = ((res["出口单价（美元/千克）"] - res["上期单价"]) / res["上期单价"].replace(0, np.nan) * 100).round(2)

        # 大盘份额
        yearly_total = res.groupby("统计年份")["金额_美元"].transform("sum")
        res["金额份额%"] = (res["金额_美元"] / yearly_total * 100).round(2)

        return res

    # --- 各维度汇总 ---
    annual_summary   = calc_metrics(export_df, [])
    partner_summary  = calc_metrics(export_df, ["贸易伙伴名称", "所属区域"])
    latest_partner   = partner_summary[partner_summary["统计年份"] == latest_year].sort_values("金额_美元", ascending=False)

    partner_top20 = latest_partner.head(20).copy()
    partner_top20["排名"] = range(1, len(partner_top20) + 1)
    # 累计份额
    partner_top20["累计金额份额%"] = partner_top20["金额份额%"].cumsum().round(2)

    # 高附加值市场（金额前40中单价最高的10个）
    if "出口单价（美元/千克）" in latest_partner.columns and latest_partner["出口单价（美元/千克）"].notna().any():
        top40 = latest_partner.head(40)
        high_value_top10 = top40.sort_values("出口单价（美元/千克）", ascending=False).head(10).copy()
        high_value_top10["高附加值排名"] = range(1, len(high_value_top10) + 1)
    else:
        high_value_top10 = pd.DataFrame()

    province_summary = calc_metrics(export_df, ["注册地名称"])
    latest_province  = province_summary[province_summary["统计年份"] == latest_year].sort_values("金额_美元", ascending=False)
    province_top10   = latest_province.head(10).copy()

    region_summary = calc_metrics(export_df, ["所属区域"])

    return {
        "all_df": all_df,
        "export_df": export_df,
        "years": years,
        "latest_year": latest_year,
        "annual_summary": annual_summary,
        "partner_summary": partner_summary,
        "partner_top20": partner_top20,
        "high_value_top10": high_value_top10,
        "province_summary": province_summary,
        "province_top10": province_top10,
        "region_summary": region_summary,
    }


# ================= 业务执行流 =================
if raw_files:
    if st.button("🚀 开始多维解码与报告生成"):
        with st.spinner("正在智能解析字段并重构高阶多维数据模型..."):
            data_pack = process_data(raw_files, analysis_mode, mapping_file, n_months)

        if data_pack is None:
            st.stop()

        latest_year  = data_pack["latest_year"]
        years        = data_pack["years"]
        p_top20      = data_pack["partner_top20"]
        prov_top10   = data_pack["province_top10"]
        reg_sum      = data_pack["region_summary"]
        latest_region = reg_sum[reg_sum["统计年份"] == latest_year]
        high_val     = data_pack["high_value_top10"]
        annual_sum   = data_pack["annual_summary"]

        # ---- 顶部核心指标 ----
        mode_label = f"前{n_months}月" if analysis_mode == "月度/前N月动态（如前三月）" and n_months else "年度"
        st.success(f"✅ 解析完成！数据年份范围：{' / '.join(years)}，最新年份：**{latest_year}**")
        st.markdown(f"### 🌐 {latest_year} {mode_label}全球市场格局透视")

        # 年度总量核心指标卡
        if not annual_sum.empty:
            latest_annual = annual_sum[annual_sum["统计年份"] == latest_year]
            prev_annual   = annual_sum[annual_sum["统计年份"] == (years[-2] if len(years) > 1 else latest_year)]
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                total_amt = latest_annual["金额_美元"].sum() if not latest_annual.empty else 0
                st.metric(f"{latest_year}出口总额（美元）", f"${total_amt:,.0f}")
            with col_b:
                yoy = latest_annual["金额同比%"].values[0] if not latest_annual.empty and "金额同比%" in latest_annual.columns else None
                st.metric("同比增长", f"{yoy:.2f}%" if yoy is not None and not np.isnan(yoy) else "—")
            with col_c:
                country_cnt = data_pack["export_df"][data_pack["export_df"]["统计年份"] == latest_year]["贸易伙伴名称"].nunique()
                st.metric("出口目的地国家/地区数", country_cnt)

        st.markdown("---")

        # ---- 主图表大屏 ----
        col1, col2 = st.columns(2)

        with col1:
            st.subheader(f"🏆 {latest_year} 前十大出口目的地（金额+市占率）")
            top10_display = p_top20.head(10)
            fig1 = px.bar(
                top10_display, x="贸易伙伴名称", y="金额_美元",
                color="所属区域",
                hover_data={c: True for c in ["金额份额%", "金额同比%", "累计金额份额%"] if c in top10_display.columns},
                text_auto='.2s',
                labels={"金额_美元": "出口金额（美元）", "贸易伙伴名称": "目的地"}
            )
            fig1.update_layout(xaxis_tickangle=-30)
            st.plotly_chart(fig1, use_container_width=True)

            # 高附加值市场散点图
            st.subheader(f"💎 {latest_year} 高附加值市场 TOP10（单价视角）")
            if not high_val.empty and "出口单价（美元/千克）" in high_val.columns:
                fig3 = px.scatter(
                    high_val, x="金额_美元", y="出口单价（美元/千克）",
                    size="金额_美元", color="所属区域",
                    hover_name="贸易伙伴名称", size_max=40,
                    labels={"金额_美元": "出口金额（美元）", "出口单价（美元/千克）": "单价（美元/千克）"}
                )
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("源数据中缺乏有效的数量/重量字段，无法渲染单价视图。")

        with col2:
            st.subheader(f"🏭 {latest_year} 核心出口省份份额（TOP10）")
            fig2 = px.pie(prov_top10, names="注册地名称", values="金额_美元", hole=0.4,
                          labels={"注册地名称": "省份/注册地", "金额_美元": "出口金额（美元）"})
            fig2.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader(f"🌍 {latest_year} 全球大区同比波动")
            if not latest_region.empty and "金额同比%" in latest_region.columns:
                fig4 = px.bar(
                    latest_region, x="所属区域", y="金额同比%",
                    color="所属区域", text_auto='.2f',
                    labels={"金额同比%": "同比增长率（%）", "所属区域": "大区"}
                )
                fig4.add_hline(y=0, line_dash="dash", line_color="gray")
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.info("大区同比数据不足（需至少两年数据）。")

        # ---- 年度趋势折线图 ----
        if len(years) > 1:
            st.markdown("---")
            st.subheader("📈 历年出口总额趋势")
            trend_df = data_pack["annual_summary"].copy()
            fig_trend = px.line(trend_df, x="统计年份", y="金额_美元", markers=True,
                                labels={"金额_美元": "出口总额（美元）", "统计年份": "年份"},
                                text="金额_美元")
            fig_trend.update_traces(texttemplate='$%{text:,.0f}', textposition='top center')
            st.plotly_chart(fig_trend, use_container_width=True)

        # ---- TOP20 详细数据表 ----
        st.markdown("---")
        st.subheader(f"📋 {latest_year} 各国 TOP20 出口明细（含份额）")
        display_cols = [c for c in ["排名", "贸易伙伴名称", "所属区域", "金额_美元",
                                    "金额份额%", "累计金额份额%", "金额同比%",
                                    "出口单价（美元/千克）", "单价同比%"] if c in p_top20.columns]
        st.dataframe(p_top20[display_cols].reset_index(drop=True), use_container_width=True)

        # ================= Excel 底稿导出 =================
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data_pack["all_df"].to_excel(writer, sheet_name="清洗后完整数据", index=False)
            annual_sum.to_excel(writer, sheet_name="年度总览_整体同比", index=False)
            data_pack["partner_summary"].to_excel(writer, sheet_name="贸易伙伴_明细(含单价同比)", index=False)
            p_top20.head(10).to_excel(writer, sheet_name="前十大出口目的地", index=False)
            p_top20.to_excel(writer, sheet_name=f"{latest_year}各国TOP20(金额+份额)", index=False)
            if not high_val.empty:
                high_val.to_excel(writer, sheet_name=f"{latest_year}高附加值市场TOP10", index=False)
            data_pack["province_summary"].to_excel(writer, sheet_name="注册地_明细", index=False)
            prov_top10.to_excel(writer, sheet_name="前十大出口省份_份额", index=False)
            reg_sum.to_excel(writer, sheet_name="区域汇总_同比", index=False)

        col_dl1, col_dl2 = st.columns([1, 2])
        with col_dl1:
            st.download_button(
                label="📥 一键下载全维数据底稿",
                data=output.getvalue(),
                file_name=f"多维市场洞察底稿_{latest_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # ================= AI 智库撰稿模块 =================
        st.header("✍️ 审美智库 AI 洞察报告生成")

        if openrouter_key:
            context_str  = f"数据年份范围：{', '.join(years)}\n"
            context_str += f"年度总大盘异动：\n{annual_sum.tail(3).to_string()}\n\n"
            context_str += f"核心市场(TOP5)市占率与增幅：\n{p_top20[['贸易伙伴名称', '金额_美元', '金额份额%', '金额同比%']].head(5).to_string()}\n\n"
            if not high_val.empty and "出口单价（美元/千克）" in high_val.columns:
                context_str += f"高附加值蓝海市场：\n{high_val[['贸易伙伴名称', '出口单价（美元/千克）', '金额_美元']].head(3).to_string()}"

            system_prompt = """
你是一位资深的泛家居与卫浴行业市场观察者，主理着一家专业媒体矩阵。
请根据提供的海关出口多维结构数据（包含市场份额、同比增幅、高单价蓝海市场），撰写一篇约 1500 字的高级智库分析文章。
要求：
1. 语调需保持"行业和市场观察者"的深刻与专业，面向市场层，不讲空泛套话。
2. 融入"栖居的美学"理念，并将冰冷的数据变化与宏观产业带位移（如产能出海、欧美去库存周期、高附加值市场突围）建立逻辑关联。
3. 文章需有标题、核心摘要、以及三个层级清晰的数据解读段落。
            """

            with st.spinner("AI 引擎正在结合多维数据进行深度思考推演..."):
                headers = {
                    "Authorization": f"Bearer {openrouter_key}",
                    "HTTP-Referer": "https://github.com/",
                    "X-Title": "Aesthetic Think Tank Data Observer"
                }
                payload = {
                    "model": ai_model,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"以下是提取的最核心高维异动数据：\n{context_str}\n请开始撰写智库观察："}
                    ]
                }

                try:
                    response = requests.post(
                        "https://openrouter.ai/api/v1/chat/completions",
                        headers=headers, json=payload, timeout=60
                    )
                    if response.status_code == 200:
                        report = response.json()['choices'][0]['message']['content']
                        st.success("✨ 高维洞察报告生成完毕！")
                        st.text_area("媒体发布底稿", report, height=600)
                    else:
                        st.error(f"调用异常。错误代码: {response.status_code} - {response.text}")
                except Exception as e:
                    st.error(f"网络请求失败: {str(e)}")
        else:
            st.warning("⚠️ 请在左侧边栏输入 API Key 以启动智库撰稿引擎。")
else:
    st.info("👈 请在左侧边栏上传「海关原始数据」，系统将自动识别字段、计算单价、份额、同比并产出对标底稿。")
