import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import os

# ================= 核心配置区 =================
# 1. 在这里填入你的 OpenRouter API Key (保留引号)
BUILTIN_API_KEY = "sk-or-v1-这里替换成你真实的API_KEY"

# 2. 内置的区域映射表文件名
MAPPING_FILE_NAME = "区域映射表.xlsx"

# ================= 页面配置 =================
st.set_page_config(page_title="卫浴行业数据观察智库", layout="wide")
st.title("📊 卫浴与泛家居进出口数据洞察大屏")

# ================= 侧边栏配置 =================
with st.sidebar:
    st.header("⚙️ 分析引擎配置")
    analysis_mode = st.radio("选择分析维度", ["年度全景统计", "月度/前N月动态"])
    
    st.markdown("---")
    st.header("📂 原始数据上传")
    st.info("💡 提示：区域映射表已在系统后台内置，无需重复上传。")
    raw_files = st.file_uploader("请拖入海关原始数据 (支持多选Excel)", type=['xlsx'], accept_multiple_files=True)

# ================= 核心处理函数 =================
@st.cache_data
def process_data(raw_files, mode):
    # 1. 自动读取内置的映射表
    region_dict = {}
    if os.path.exists(MAPPING_FILE_NAME):
        try:
            map_df = pd.read_excel(MAPPING_FILE_NAME, sheet_name="区域映射")
            region_dict = dict(zip(map_df["原始名称"], map_df["子区域"]))
        except Exception as e:
            st.toast(f"读取映射表出错: {str(e)}")
    else:
        st.warning(f"⚠️ 未找到 {MAPPING_FILE_NAME} 文件，区域分类将默认为'其他'。")

    # 2. 读取原始数据
    all_data = []
    for file in raw_files:
        df = pd.read_excel(file)
        cols = ["商品编码", "商品名称", "贸易伙伴编码", "贸易伙伴名称",
                "注册地编码", "注册地名称", "贸易类型", "金额_美元", "统计年份"]
        df = df[[c for c in cols if c in df.columns]]
        df = df.dropna(subset=["贸易伙伴名称", "金额_美元"])
        df["金额_美元"] = pd.to_numeric(df["金额_美元"], errors='coerce').fillna(0)
        df["贸易伙伴名称"] = df["贸易伙伴名称"].str.strip()
        df["统计年份"] = df["统计年份"].astype(str).str[:4]
        df["所属区域"] = df["贸易伙伴名称"].map(region_dict).fillna("其他")
        all_data.append(df)
        
    all_df = pd.concat(all_data, ignore_index=True)
    
    # 3. 统计逻辑 A：目的地视角 (贸易伙伴)
    partner_summary = all_df.groupby(["统计年份", "贸易伙伴名称", "所属区域", "贸易类型"], as_index=False)["金额_美元"].sum()
    partner_summary = partner_summary.sort_values(["贸易伙伴名称", "统计年份"])
    partner_summary["上期金额_美元"] = partner_summary.groupby("贸易伙伴名称")["金额_美元"].shift(1)
    partner_summary["同比变化"] = (partner_summary["金额_美元"] - partner_summary["上期金额_美元"]) / partner_summary["上期金额_美元"].replace(0, 1)
    export_partner = partner_summary[partner_summary["贸易类型"] == "出口"].copy()

    # 4. 统计逻辑 B：国内产区视角 (注册地)
    province_summary = all_df.groupby(["统计年份", "注册地名称", "所属区域", "贸易类型"], as_index=False)["金额_美元"].sum()
    province_summary = province_summary.sort_values(["注册地名称", "统计年份"])
    province_summary["上期金额_美元"] = province_summary.groupby("注册地名称")["金额_美元"].shift(1)
    province_summary["同比变化"] = (province_summary["金额_美元"] - province_summary["上期金额_美元"]) / province_summary["上期金额_美元"].replace(0, 1)
    export_province = province_summary[province_summary["贸易类型"] == "出口"].copy()

    # 5. 统计逻辑 C：全球大区视角 (所属区域)
    region_summary = all_df.groupby(["统计年份", "所属区域", "贸易类型"], as_index=False)["金额_美元"].sum()
    region_summary = region_summary.sort_values(["所属区域", "统计年份"])
    region_summary["上期金额_美元"] = region_summary.groupby("所属区域")["金额_美元"].shift(1)
    region_summary["同比变化"] = (region_summary["金额_美元"] - region_summary["上期金额_美元"]) / region_summary["上期金额_美元"].replace(0, 1)
    export_region = region_summary[region_summary["贸易类型"] == "出口"].copy()
    
    return all_df, export_partner, export_province, export_region

# ================= 业务执行流 =================
if raw_files:
    if st.button("🚀 开始多维解码与报告生成"):
        with st.spinner("正在重构多维数据模型并绘制大屏..."):
            all_df, export_partner, export_province, export_region = process_data(raw_files, analysis_mode)
            
            latest_year = export_partner['统计年份'].max()
            
            # 提取最新一期各项Top数据
            latest_partner = export_partner[export_partner['统计年份'] == latest_year].sort_values("金额_美元", ascending=False).head(10)
            latest_province = export_province[export_province['统计年份'] == latest_year].groupby("注册地名称", as_index=False)["金额_美元"].sum().sort_values("金额_美元", ascending=False).head(10)
            latest_region = export_region[export_region['统计年份'] == latest_year].groupby("所属区域", as_index=False)["金额_美元"].sum().sort_values("金额_美元", ascending=False)
            
            # ================= 可视化大屏排版 =================
            st.markdown("### 🌐 全球市场与国内产区格局透视")
            col1, col2 = st.columns(2)
            
            with col1:
                # 图1：目的地 Top10
                st.subheader(f"🏆 {latest_year}年 前十大出口目的地")
                fig1 = px.bar(latest_partner, x="贸易伙伴名称", y="金额_美元", color="所属区域", text_auto='.2s')
                st.plotly_chart(fig1, use_container_width=True)
                
                # 图3：大区市场份额
                st.subheader("🌍 全球大区市场份额占比")
                fig3 = px.pie(latest_region, names="所属区域", values="金额_美元", hole=0.4)
                fig3.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig3, use_container_width=True)

            with col2:
                # 图2：国内产区 Top10
                st.subheader(f"🏭 {latest_year}年 前十大出口省份(注册地)")
                fig2 = px.bar(latest_province, x="注册地名称", y="金额_美元", text_auto='.2s', color_discrete_sequence=['#4B8BBE'])
                st.plotly_chart(fig2, use_container_width=True)

                # 图4：核心市场同比增跌幅
                st.subheader("📈 核心海外市场动态与波动率")
                latest_partner['同比%'] = latest_partner['同比变化'] * 100
                fig4 = px.scatter(latest_partner, x="金额_美元", y="同比%", color="所属区域", size="金额_美元", hover_name="贸易伙伴名称")
                # 添加一条0轴参考线
                fig4.add_hline(y=0, line_dash="dash", line_color="gray")
                st.plotly_chart(fig4, use_container_width=True)

            # ================= 底稿数据导出 =================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                export_partner.to_excel(writer, sheet_name="出口目的地_明细与同比", index=False)
                export_province.to_excel(writer, sheet_name="出口省份注册地_明细", index=False)
                export_region.to_excel(writer, sheet_name="全球大区_汇总与同比", index=False)
                all_df.to_excel(writer, sheet_name="清洗后完整底层数据", index=False)
            
            st.download_button(
                label="📥 下载多维清洗统计数据包 (含产区与大区结构)",
                data=output.getvalue(),
                file_name=f"行业多维市场洞察底稿_{latest_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ================= AI 撰稿模块 =================
            st.markdown("---")
            st.header("✍️ 审美智库 AI 洞察报告生成")
            
            # 为 AI 构建更丰满的上下文
            data_context = f"""
            【1. 前五大出口目的地】：
            {latest_partner[['贸易伙伴名称', '所属区域', '金额_美元', '同比%']].head(5).to_string()}
            
            【2. 前五大国内出口省份】：
            {latest_province.head(5).to_string()}
            
            【3. 全球大区金额排序】：
            {latest_region.head(5).to_string()}
            """
            
            if BUILTIN_API_KEY and not BUILTIN_API_KEY.startswith("sk-or-v1-这里替换"):
                system_prompt = """
                你是一位资深的行业和市场观察者，面对专业的泛家居和卫浴市场进行深度解读。
                请根据以下最新的海关出口多维数据（含目的地、国内产区、全球大区），撰写一篇约 1500 字的深度市场观察文章。
                重点要求：
                1. 语调需保持思想领导力，客观、深刻，带有高级审美智库的沉淀感。
                2. 结合“栖居的美学”理念，以及宏观层面的供应链位移（如中美贸易摩擦带来的China+1战略、国内主产区如广东/福建的产能外溢、东南亚替代效应等）进行数据拆解。
                3. 分析全球大区格局的演变以及国内出海力量的地缘分布。
                请直接输出文章正文，结构清晰，适合在高端行业论坛或核心媒体首发。
                """
                
                with st.spinner("AI 正在深度思考并撰写多维行业观察报告..."):
                    headers = {
                        "Authorization": f"Bearer {BUILTIN_API_KEY}",
                        "HTTP-Referer": "https://github.com/",
                        "X-Title": "Sanitary Ware Data Observer"
                    }
                    data = {
                        "model": "openai/gpt-oss-120b:free", 
                        "messages": [
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"以下是最新核心出海异动数据：\n{data_context}\n请开始撰写分析："}
                        ]
                    }
                    
                    try:
                        response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=data)
                        if response.status_code == 200:
                            report = response.json()['choices'][0]['message']['content']
                            st.success("✨ 多维洞察报告生成完毕！")
                            st.text_area("文章初稿 (可直接提取至公众号或论坛使用)", report, height=600)
                        else:
                            st.error(f"AI 调用异常，请检查 API Key 状态。错误代码: {response.status_code} - {response.text}")
                    except Exception as e:
                        st.error(f"网络请求失败: {str(e)}")
            else:
                st.error("⚠️ 开发者提示：请先在 app.py 代码顶部的 BUILTIN_API_KEY 变量中填入真实的 OpenRouter Key。")
else:
    st.info("👈 请在左侧边栏上传「海关原始数据」以启动智库多维引擎。")
