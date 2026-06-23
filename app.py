import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import os
import numpy as np

# ================= 核心配置区 =================
MAPPING_FILE_NAME = "区域映射表.xlsx"

# ================= 页面配置 =================
st.set_page_config(page_title="卫浴行业数据观察智库", layout="wide")
st.title("📊 卫浴与泛家居进出口多维洞察大屏")

# ================= 侧边栏配置 =================
with st.sidebar:
    st.header("⚙️ 引擎配置")
    analysis_mode = st.radio("选择分析维度", ["年度全景统计", "月度/前N月动态（如前三月）"])
    
    # 1. 恢复前台手动输入 API Key
    st.markdown("---")
    st.header("🔑 AI 智库模型接入")
    openrouter_key = st.text_input("OpenRouter API Key", type="password", help="在此输入你的专属密钥（以 sk-or-v1 开头）")
    ai_model = st.selectbox("选择推理模型", ["openai/gpt-oss-120b:free", "deepseek/deepseek-chat:free"])
    
    st.markdown("---")
    st.header("📂 原始数据上传")
    st.info("💡 提示：区域映射表已在系统后台内置，直接上传海关明细表即可。")
    raw_files = st.file_uploader("请拖入海关原始数据 (支持多选Excel)", type=['xlsx'], accept_multiple_files=True)

# ================= 核心多维处理引擎 =================
@st.cache_data
def process_data(raw_files, mode):
    # 1. 读取内置映射表
    region_dict = {}
    if os.path.exists(MAPPING_FILE_NAME):
        try:
            map_df = pd.read_excel(MAPPING_FILE_NAME, sheet_name="区域映射")
            region_dict = dict(zip(map_df["原始名称"], map_df["子区域"]))
        except:
            pass

    # 2. 读取原始数据（智能识别数量字段以计算单价）
    all_data = []
    for file in raw_files:
        df = pd.read_excel(file)
        
        # 兼容不同海关数据源的数量命名
        qty_col = "数量"
        if "第一法定数量" in df.columns:
            qty_col = "第一法定数量"
        elif "统计数量" in df.columns:
            qty_col = "统计数量"
            
        base_cols = ["商品编码", "商品名称", "贸易伙伴编码", "贸易伙伴名称",
                     "注册地编码", "注册地名称", "贸易类型", "金额_美元", "统计年份"]
                     
        # 提取存在的列
        keep_cols = [c for c in base_cols if c in df.columns]
        if qty_col in df.columns:
            keep_cols.append(qty_col)
            
        df = df[keep_cols]
        df = df.dropna(subset=["贸易伙伴名称", "金额_美元"])
        df["金额_美元"] = pd.to_numeric(df["金额_美元"], errors='coerce').fillna(0)
        
        if qty_col in df.columns:
            df["数量_统一"] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
        else:
            df["数量_统一"] = 0 # 若无数量字段则置零，避免报错
            
        df["贸易伙伴名称"] = df["贸易伙伴名称"].str.strip()
        df["统计年份"] = df["统计年份"].astype(str).str[:4]
        df["所属区域"] = df["贸易伙伴名称"].map(region_dict).fillna("其他")
        all_data.append(df)
        
    all_df = pd.concat(all_data, ignore_index=True)
    export_df = all_df[all_df["贸易类型"] == "出口"].copy()
    
    # 获取统计年度列表
    years = sorted(export_df["统计年份"].unique())
    latest_year = years[-1] if years else "最新年份"

    # ================= 核心计算函数 (含份额、单价、同比) =================
    def calc_metrics(df, group_cols):
        res = df.groupby(group_cols + ["统计年份"], as_index=False).agg({"金额_美元": "sum", "数量_统一": "sum"})
        res = res.sort_values(group_cols + ["统计年份"])
        
        # 计算单价 (若无数量或数量为0则返回空)
        res["单价_美元"] = np.where(res["数量_统一"] > 0, res["金额_美元"] / res["数量_统一"], np.nan)
        
        # 计算同比
        group_base = group_cols if len(group_cols) > 0 else ["统计年份"] # 规避无分组报错
        res["上期金额"] = res.groupby(group_cols)["金额_美元"].shift(1) if group_cols else res["金额_美元"].shift(1)
        res["上期单价"] = res.groupby(group_cols)["单价_美元"].shift(1) if group_cols else res["单价_美元"].shift(1)
        
        res["金额同比%"] = (res["金额_美元"] - res["上期金额"]) / res["上期金额"].replace(0, np.nan) * 100
        res["单价同比%"] = (res["单价_美元"] - res["上期单价"]) / res["上期单价"].replace(0, np.nan) * 100
        
        # 计算大盘份额 (该维度当年总金额的占比)
        yearly_total = res.groupby("统计年份")["金额_美元"].transform("sum")
        res["金额份额%"] = res["金额_美元"] / yearly_total * 100
        
        return res

    # 1. 年度总览_整体同比
    annual_summary = calc_metrics(export_df, [])
    
    # 2. 贸易伙伴 (含单价、份额、同比)
    partner_summary = calc_metrics(export_df, ["贸易伙伴名称", "所属区域"])
    latest_partner = partner_summary[partner_summary["统计年份"] == latest_year].sort_values("金额_美元", ascending=False)
    
    # 3. 各国TOP20 (金额+份额)
    partner_top20 = latest_partner.head(20).copy()
    partner_top20["排名"] = range(1, len(partner_top20) + 1)
    
    # 4. 高附加值市场 TOP10 (在金额前40的国家中，筛选单价最高的10个)
    if "单价_美元" in partner_summary.columns and partner_summary["单价_美元"].notna().any():
        top40_markets = latest_partner.head(40)
        high_value_top10 = top40_markets.sort_values("单价_美元", ascending=False).head(10).copy()
        high_value_top10["高附加值排名"] = range(1, len(high_value_top10) + 1)
    else:
        high_value_top10 = pd.DataFrame()

    # 5. 注册地 (省份)
    province_summary = calc_metrics(export_df, ["注册地名称", "所属区域"])
    latest_province = province_summary[province_summary["统计年份"] == latest_year].sort_values("金额_美元", ascending=False)
    province_top10 = latest_province.head(10).copy()
    
    # 6. 区域汇总
    region_summary = calc_metrics(export_df, ["所属区域"])
    
    return {
        "all_df": all_df,
        "latest_year": latest_year,
        "annual_summary": annual_summary,
        "partner_summary": partner_summary,
        "partner_top20": partner_top20,
        "high_value_top10": high_value_top10,
        "province_summary": province_summary,
        "province_top10": province_top10,
        "region_summary": region_summary
    }

# ================= 业务执行流 =================
if raw_files:
    if st.button("🚀 开始多维解码与报告生成"):
        with st.spinner("正在重构高阶多维数据模型并绘制大屏..."):
            data_pack = process_data(raw_files, analysis_mode)
            
            latest_year = data_pack["latest_year"]
            p_top20 = data_pack["partner_top20"]
            prov_top10 = data_pack["province_top10"]
            reg_sum = data_pack["region_summary"]
            latest_region = reg_sum[reg_sum["统计年份"] == latest_year]
            high_val = data_pack["high_value_top10"]
            
            # ================= 可视化大屏排版 (高阶维度) =================
            st.markdown(f"### 🌐 {latest_year} 全球市场与产区格局透视 (含份额与单价)")
            col1, col2 = st.columns(2)
            
            with col1:
                # 图1：目的地 Top10 金额与份额
                st.subheader("🏆 前十大出口目的地 (金额与市占率)")
                fig1 = px.bar(p_top20.head(10), x="贸易伙伴名称", y="金额_美元", color="所属区域", 
                              hover_data=["金额份额%", "金额同比%"], text_auto='.2s')
                st.plotly_chart(fig1, use_container_width=True)
                
                # 图3：高附加值市场
                st.subheader("💎 高附加值市场 TOP10 (单价视角)")
                if not high_val.empty:
                    fig3 = px.scatter(high_val, x="金额_美元", y="单价_美元", size="金额_美元", 
                                      color="所属区域", hover_name="贸易伙伴名称", size_max=40)
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("源数据中缺乏有效的数量字段，无法渲染单价视图。")

            with col2:
                # 图2：国内产区 Top10
                st.subheader("🏭 核心出口省份份额占比 (TOP10)")
                fig2 = px.pie(prov_top10, names="注册地名称", values="金额_美元", hole=0.4)
                fig2.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig2, use_container_width=True)

                # 图4：全球大区格局
                st.subheader("🌍 全球大区动态与波动率")
                fig4 = px.bar(latest_region, x="所属区域", y="金额同比%", color="所属区域", text_auto='.2f')
                fig4.add_hline(y=0, line_dash="dash", line_color="gray")
                st.plotly_chart(fig4, use_container_width=True)

            # ================= 严格对齐截图的 Excel 底稿导出 =================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # 按照你截图的命名习惯写入 Sheet
                data_pack["all_df"].to_excel(writer, sheet_name="清洗后完整数据", index=False)
                data_pack["annual_summary"].to_excel(writer, sheet_name="年度总览_整体同比", index=False)
                data_pack["partner_summary"].to_excel(writer, sheet_name="贸易伙伴_明细(含单价同比)", index=False)
                p_top20.head(10).to_excel(writer, sheet_name="前十大出口目的地", index=False)
                p_top20.to_excel(writer, sheet_name=f"{latest_year}各国TOP20(金额+份额)", index=False)
                if not high_val.empty:
                    high_val.to_excel(writer, sheet_name=f"{latest_year}高附加值市场TOP10", index=False)
                data_pack["province_summary"].to_excel(writer, sheet_name="注册地_明细", index=False)
                prov_top10.to_excel(writer, sheet_name="前十大出口省份_份额", index=False)
                reg_sum.to_excel(writer, sheet_name="区域汇总_同比", index=False)

            st.markdown("---")
            col_dl1, col_dl2 = st.columns([1, 2])
            with col_dl1:
                st.download_button(
                    label="📥 一键下载全维数据底稿 (对标截图结构)",
                    data=output.getvalue(),
                    file_name=f"多维市场洞察底稿_{latest_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # ================= 审美智库 AI 撰稿模块 =================
            st.header("✍️ 审美智库 AI 洞察报告生成")
            
            if openrouter_key:
                # 为 AI 提供高阶数据：包含份额、单价、高附加值信息
                context_str = f"年度总大盘异动：\n{data_pack['annual_summary'].tail(2).to_string()}\n\n"
                context_str += f"核心市场(TOP5)市占率与增幅：\n{p_top20[['贸易伙伴名称', '金额_美元', '金额份额%', '金额同比%']].head(5).to_string()}\n\n"
                if not high_val.empty:
                    context_str += f"高附加值蓝海市场：\n{high_val[['贸易伙伴名称', '单价_美元', '金额_美元']].head(3).to_string()}"

                system_prompt = """
                你是一位资深的泛家居与卫浴行业市场观察者，主理着一家专业媒体矩阵。
                请根据提供的海关出口多维结构数据（包含市场份额、同比增幅、高单价蓝海市场），撰写一篇约 1500 字的高级智库分析文章。
                要求：
                1. 语调需保持“行业和市场观察者”的深刻与专业，面向市场层，不讲空泛套话。
                2. 融入“栖居的美学”理念，并将冰冷的数据变化与宏观产业带位移（如产能出海、欧美去库存周期、高附加值市场突围）建立逻辑关联。
                3. 文章需有标题、核心摘要、以及三个层级清晰的数据解读段落。
                """
                
                with st.spinner("AI 引擎正在结合多维数据进行深度思考推演..."):
                    headers = {
                        "Authorization": f"Bearer {openrouter_key}",
                        "HTTP-Referer": "https://github.com/",
                        "X-Title": "Aesthetic Think Tank Data Observer"
                    }
                    data = {
                        "model": ai_model, 
                        "messages": [
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"以下是提取的最核心高维异动数据：\n{context_str}\n请开始撰写智库观察："}
                        ]
                    }
                    
                    try:
                        response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=data)
                        if response.status_code == 200:
                            report = response.json()['choices'][0]['message']['content']
                            st.success("✨ 高维洞察报告生成完毕！")
                            st.text_area("媒体发布底稿 (可直接提取至公众号或产业论坛使用)", report, height=600)
                        else:
                            st.error(f"调用异常。错误代码: {response.status_code} - {response.text}")
                    except Exception as e:
                        st.error(f"网络请求失败: {str(e)}")
            else:
                st.warning("⚠️ 请在左侧边栏输入 API Key 以启动智库撰稿引擎。")
else:
    st.info("👈 请在左侧边栏上传「海关原始数据」，系统将自动计算单价、份额、同比并产出对标底稿。")
