import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import os

# ================= 核心配置区 =================
# 1. 在这里填入你的 OpenRouter API Key (保留引号)
BUILTIN_API_KEY = "openai/gpt-oss-20b:free"

# 2. 内置的区域映射表文件名 (请确保该文件已上传至 GitHub 仓库)
MAPPING_FILE_NAME = "区域映射表.xlsx"

# ================= 页面配置 =================
st.set_page_config(page_title="卫浴行业数据观察智库", layout="wide")
st.title("📊 卫浴与泛家居进出口数据洞察")

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
        st.warning(f"⚠️ 未在后台找到内置的 {MAPPING_FILE_NAME} 文件，区域分类将默认为'其他'。请记得将其上传至 GitHub。")

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
        # 匹配区域
        df["所属区域"] = df["贸易伙伴名称"].map(region_dict).fillna("其他")
        all_data.append(df)
        
    all_df = pd.concat(all_data, ignore_index=True)
    
    # 3. 核心统计逻辑 (目的地视角)
    partner_summary = all_df.groupby(["统计年份", "贸易伙伴名称", "所属区域", "贸易类型"], as_index=False)["金额_美元"].sum()
    partner_summary = partner_summary.sort_values(["贸易伙伴名称", "统计年份"])
    partner_summary["上期金额_美元"] = partner_summary.groupby("贸易伙伴名称")["金额_美元"].shift(1)
    partner_summary["同比变化"] = (partner_summary["金额_美元"] - partner_summary["上期金额_美元"]) / partner_summary["上期金额_美元"].replace(0, 1)
    
    export_partner = partner_summary[partner_summary["贸易类型"] == "出口"].copy()
    
    return all_df, export_partner

# ================= 业务执行流 =================
if raw_files:
    if st.button("🚀 开始数据解码与报告生成"):
        with st.spinner("正在清洗海关数据并绘制洞察图表..."):
            all_df, export_partner = process_data(raw_files, analysis_mode)
            
            # 获取最新年份数据做展示
            latest_year = export_partner['统计年份'].max()
            latest_data = export_partner[export_partner['统计年份'] == latest_year].sort_values("金额_美元", ascending=False).head(10)
            
            # 布局展示
            col1, col2 = st.columns(2)
            with col1:
                st.subheader(f"🏆 {latest_year}年 前十大出口目的地")
                fig1 = px.bar(latest_data, x="贸易伙伴名称", y="金额_美元", color="所属区域", text_auto='.2s')
                st.plotly_chart(fig1, use_container_width=True)
                
            with col2:
                st.subheader("📈 核心市场同比增跌幅")
                latest_data['同比%'] = latest_data['同比变化'] * 100
                fig2 = px.scatter(latest_data, x="金额_美元", y="同比%", color="所属区域", size="金额_美元", hover_name="贸易伙伴名称")
                st.plotly_chart(fig2, use_container_width=True)

            # 导出数据准备
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                export_partner.to_excel(writer, sheet_name="出口目的地明细", index=False)
                all_df.to_excel(writer, sheet_name="清洗后底层数据", index=False)
            
            st.download_button(
                label="📥 下载完整清洗统计数据包 (Excel)",
                data=output.getvalue(),
                file_name=f"行业市场洞察数据底稿_{latest_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ================= AI 撰稿模块 =================
            st.markdown("---")
            st.header("✍️ 审美智库 AI 洞察报告生成")
            
            # 提取核心数据喂给 AI
            data_context = latest_data[['贸易伙伴名称', '所属区域', '金额_美元', '同比%']].to_string()
            
            if BUILTIN_API_KEY and not BUILTIN_API_KEY.startswith("sk-or-v1-这里替换"):
                # 针对多品类卫浴市场的通用深度提示词
                system_prompt = """
                你是一位资深的行业和市场观察者，面对专业的泛家居和卫浴市场进行深度解读。
                请根据以下最新的海关出口核心数据，撰写一篇约 1500 字的深度市场观察文章。
                重点要求：
                1. 语调需保持思想领导力，客观、深刻，摒弃传统销售说辞，带有高级审美智库（如《ELLE家居廊》）的沉淀感。
                2. 结合“栖居的美学”理念，以及宏观层面的供应链转移（如中美贸易摩擦带来的China+1战略、原材料波动影响、海外去库存周期等）进行数据拆解。
                3. 分析前十大出口国的格局变化，挖掘数据背后的全球产线重构逻辑。
                请直接输出文章正文，结构清晰，适合在高端行业论坛或媒体首发。
                """
                
                with st.spinner("AI 正在深度思考并撰写行业观察报告..."):
                    headers = {
                        "Authorization": f"Bearer {BUILTIN_API_KEY}",
                        "HTTP-Referer": "https://github.com/",
                        "X-Title": "Sanitary Ware Data Observer"
                    }
                    data = {
                        "model": "openai/gpt-oss-120b:free", # 已替换为你指定的内置模型
                        "messages": [
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"以下是最新核心出口数据：\n{data_context}\n请开始撰写分析："}
                        ]
                    }
                    
                    try:
                        response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=data)
                        if response.status_code == 200:
                            report = response.json()['choices'][0]['message']['content']
                            st.success("✨ 报告生成完毕！")
                            st.text_area("文章初稿 (可直接复制修改)", report, height=600)
                        else:
                            st.error(f"AI 调用异常，请检查 API Key 状态。错误代码: {response.status_code} - {response.text}")
                    except Exception as e:
                        st.error(f"网络请求失败: {str(e)}")
            else:
                st.error("⚠️ 开发者提示：请先在 app.py 代码顶部的 BUILTIN_API_KEY 变量中填入真实的 OpenRouter Key。")
else:
    st.info("👈 请在左侧边栏上传「海关原始数据」以启动智库引擎。")
