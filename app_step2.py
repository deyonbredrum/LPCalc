import streamlit as st
import pandas as pd
from datetime import datetime

# ===== 导入所有模块 =====
from calc_engine import (
    calc_lightning_protection, calc_electronic_protection,
    get_correction_factors, get_building_attributes,
    get_c1_options, get_c2_options, get_c3_options, get_c4_options,
    get_c5_options, get_c6_options, get_surrounding_options,
    get_power_cable_options, get_signal_cable_options,
    get_power_cable_coef, get_signal_cable_coef,
    get_lightning_level, BUILDING_ATTR
)

from data_manager import (
    load_data_from_excel, save_data_to_excel,
    add_city, delete_city, get_all_cities,
    get_cities_by_province, get_td_by_city
)

from excel_export import export_excel_report, export_excel_separate

from params import PROVINCES, C1_OPTIONS, C2_OPTIONS, C3_OPTIONS, C4_OPTIONS, C5_OPTIONS, C6_OPTIONS

st.set_page_config(page_title="LPCalc", layout="wide")

st.title("LPCalc - 步骤2")
st.write("✅ 所有模块导入成功，正在测试侧边栏...")

# ===== 测试侧边栏 =====
with st.sidebar:
    st.header("📐 输入参数")

    # ---- 分类1：建筑物尺寸 ----
    st.subheader("📏 建筑物尺寸")
    L = st.number_input("建筑物长 L (m)", min_value=0.1, value=117.0, step=1.0)
    W = st.number_input("建筑物宽 W (m)", min_value=0.1, value=36.0, step=1.0)
    H = st.number_input("建筑物高 H (m)", min_value=0.1, value=6.5, step=0.5)

    st.success("✅ 侧边栏加载成功！")

st.write("如果看到这个页面，说明侧边栏也没有问题。")