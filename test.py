import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="LPCalc", layout="wide")
st.title("LPCalc - 最小版本")

# 先不导入任何自定义模块
st.write("如果能看到这个页面，说明基础导入正常")

# 逐步测试导入
try:
    st.write("正在测试导入 calc_engine...")
    from calc_engine import get_correction_factors
    st.success("✅ calc_engine 导入成功")
except Exception as e:
    st.error(f"calc_engine 导入失败: {e}")