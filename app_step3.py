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

st.title("LPCalc - 步骤3")
st.write("✅ 测试防雷等级计算 Tab")

# ===== 侧边栏 =====
with st.sidebar:
    st.header("📐 输入参数")

    st.subheader("📏 建筑物尺寸")
    L = st.number_input("建筑物长 L (m)", min_value=0.1, value=117.0, step=1.0)
    W = st.number_input("建筑物宽 W (m)", min_value=0.1, value=36.0, step=1.0)
    H = st.number_input("建筑物高 H (m)", min_value=0.1, value=6.5, step=0.5)

    st.subheader("📍 地理位置")
    td_data = load_data_from_excel()
    provinces = sorted(td_data.keys())
    selected_province = st.selectbox("选择省份", provinces)
    cities_in_province = get_cities_by_province(td_data, selected_province)
    if cities_in_province:
        city_sel = st.selectbox("选择城市", cities_in_province)
    else:
        city_sel = st.text_input("输入城市名称", value="自定义城市")
    city = city_sel

    st.subheader("🏗️ 周边与校正")
    surrounding_type = st.selectbox("周边情况", get_surrounding_options())
    Lz = 0.0
    if surrounding_type in [
        "周边2D范围内有等高或比它低的其他建筑物，但不在所考虑建筑物以外100m的保护范围内",
        "周边2D范围内有比它高的其他建筑物",
    ]:
        Lz = st.number_input("Lz (m)", min_value=0.0, value=260.0, step=10.0)
    building_condition = st.selectbox("建筑物状况（校正系数）", get_correction_factors())

    st.subheader("🏷️ 建筑物属性")
    building_attr = st.selectbox("建筑物属性（防雷等级判定）", get_building_attributes())

# ===== Tab 测试：只保留防雷等级计算 =====
tab1, tab2 = st.tabs(["📊 防雷等级计算", "🛡️ 电子信息防护等级"])

with tab1:
    st.header("📊 防雷等级计算")
    st.caption("依据 GB 50057-2010")

    # 快速测试：显示当前参数
    st.write(f"建筑物尺寸: {L}m × {W}m × {H}m")
    st.write(f"城市: {city}")
    st.write(f"周边情况: {surrounding_type}")
    st.write(f"建筑物属性: {building_attr}")

    if st.button("计算防雷等级", key="calc_lp", type="primary"):
        with st.spinner("计算中..."):
            result = calc_lightning_protection(
                L=L, W=W, H=H,
                city=city,
                building_condition=building_condition,
                surrounding_type=surrounding_type,
                Lz=Lz
            )

            attr_type = BUILDING_ATTR.get(building_attr, "normal")
            level_num = get_lightning_level(result["N"], attr_type)
            level_map = {1: "一类", 2: "二类", 3: "三类"}
            level_text = level_map.get(level_num, "未达到设防要求")

            st.success(f"✅ 计算完成！年预计雷击次数 N = {result['N']:.6f} 次/a")
            st.info(f"🏷️ 防雷等级：{level_text}")

with tab2:
    st.header("🛡️ 电子信息防护等级（测试中）")
    st.info("此功能将在下一步添加")