import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="LPCalc", layout="wide")
st.title("LPCalc - 步骤2")

# 测试所有导入
try:
    from calc_engine import (
        calc_lightning_protection, calc_electronic_protection,
        get_correction_factors, get_building_attributes,
        get_c1_options, get_c2_options, get_c3_options, get_c4_options,
        get_c5_options, get_c6_options, get_surrounding_options,
        get_power_cable_options, get_signal_cable_options,
        get_power_cable_coef, get_signal_cable_coef,
        get_lightning_level, BUILDING_ATTR
    )
    st.success("✅ calc_engine 全部导入成功")
except Exception as e:
    st.error(f"❌ calc_engine 导入失败: {e}")

try:
    from data_manager import (
        load_data_from_excel, save_data_to_excel,
        add_city, delete_city, get_all_cities,
        get_cities_by_province, get_td_by_city
    )
    st.success("✅ data_manager 导入成功")
except Exception as e:
    st.error(f"❌ data_manager 导入失败: {e}")

try:
    from excel_export import export_excel_report, export_excel_separate
    st.success("✅ excel_export 导入成功")
except Exception as e:
    st.error(f"❌ excel_export 导入失败: {e}")

from params import PROVINCES, C1_OPTIONS, C2_OPTIONS, C3_OPTIONS, C4_OPTIONS, C5_OPTIONS, C6_OPTIONS
st.success("✅ params 导入成功")