import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import io

st.set_page_config(page_title="卫浴数据分析引擎", layout="wide")
st.title("📊 卫浴进出口多维分析引擎")

# 1. 自动提取列名工具
def auto_detect_col(df, target_keywords):
    for col in df.columns:
        for kw in target_keywords:
            if kw in str(col): return col
    return None

# 2. 核心清洗引擎
def process_data(uploaded_files):
    all_dfs = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        
        # 自动识别关键列
        amt_col = auto_detect_col(df, ["金额", "美元", "USD"])
        qty_col = auto_detect_col(df, ["数量", "重量", "净重", "第一法定数量"])
        year_col = auto_detect_col(df, ["统计年份", "年份", "Year"])
        
        # 标准化重命名
        rename_map = {amt_col: "金额_美元", qty_col: "数量_统一", year_col: "统计年份"}
        df.rename(columns=rename_map, inplace=True)
        
        # 数据类型强转
        df["金额_美元"] = pd.to_numeric(df["金额_美元"], errors='coerce').fillna(0)
        df["数量_统一"] = pd.to_numeric(df["数量_统一"], errors='coerce').fillna(0)
        df["统计年份"] = df["统计年份"].astype(str).str[:4] # 强制提取4位年份
        
        all_dfs.append(df)
        
    return pd.concat(all_dfs, ignore_index=True)

# 侧边栏上传
raw_files = st.file_uploader("上传海关Excel", accept_multiple_files=True)

if raw_files:
    df = process_data(raw_files)
    
    # 动态获取年份
    available_years = sorted(df["统计年份"].unique(), reverse=True)
    selected_year = st.selectbox("选择分析年份", available_years)
    
    df_filtered = df[df["统计年份"] == selected_year]
    
    # 计算逻辑：单价 = 金额 / 数量
    df_filtered["单价"] = np.where(df_filtered["数量_统一"] > 0, df_filtered["金额_美元"] / df_filtered["数量_统一"], 0)
    
    # 维度：前十大出口目的地 (按金额)
    top10_partner = df_filtered.groupby("贸易伙伴名称")["金额_美元"].sum().nlargest(10).reset_index()
    
    # 维度：前十大高附加值市场 (单价最高)
    top10_price = df_filtered.groupby("贸易伙伴名称")["单价"].mean().nlargest(10).reset_index()

    # 可视化排版
    col1, col2 = st.columns(2)
    with col1:
        st.subheader(f"{selected_year} 出口金额 TOP10")
        st.bar_chart(top10_partner.set_index("贸易伙伴名称"))
    with col2:
        st.subheader(f"{selected_year} 高单价市场 TOP10")
        st.bar_chart(top10_price.set_index("贸易伙伴名称"))

    # 下载功能
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="全量明细", index=False)
    st.download_button("下载处理结果", data=buffer, file_name="分析结果.xlsx")
