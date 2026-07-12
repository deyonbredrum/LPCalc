# -*- coding: utf-8 -*-
"""
LPCalc - 建筑物防雷及电子信息防护等级计算
GB 50057-2010 + GB 50343-2012
"""

import streamlit as st
import pandas as pd
from datetime import datetime

# 计算引擎
from calc_engine import (
    calc_lightning_protection, calc_electronic_protection,
    get_correction_factors, get_building_attributes,
    get_c1_options, get_c2_options, get_c3_options, get_c4_options,
    get_c5_options, get_c6_options, get_surrounding_options,
    get_power_cable_options, get_signal_cable_options,
    get_power_cable_coef, get_signal_cable_coef,
    get_lightning_level, BUILDING_ATTR
)

# 数据管理
from data_manager import (
    load_data_from_excel, save_data_to_excel,
    add_city, delete_city, get_all_cities,
    get_cities_by_province, get_td_by_city
)

# Word 导出（注释掉，改为在导出按钮中动态导入）
# from word_export import export_calculation_report

# 参数
from params import PROVINCES, C1_OPTIONS, C2_OPTIONS, C3_OPTIONS, C4_OPTIONS, C5_OPTIONS, C6_OPTIONS

st.set_page_config(page_title="LPCalc - 防雷等级计算", layout="wide")

st.title("⚡ LPCalc - 建筑物防雷及电子信息防护等级计算")
st.caption("LPCalc v1.0 | 依据 GB 50057-2010 + GB 50343-2012")

# ==================== 侧边栏 ====================
with st.sidebar:
    st.header("📐 输入参数")

    # ---- 分类1：建筑物尺寸 ----
    st.subheader("📏 建筑物尺寸")
    L = st.number_input("建筑物长 L (m)", min_value=0.1, value=117.0, step=1.0)
    W = st.number_input("建筑物宽 W (m)", min_value=0.1, value=36.0, step=1.0)
    H = st.number_input("建筑物高 H (m)", min_value=0.1, value=6.5, step=0.5)

    # ---- 分类2：地理位置 ----
    st.subheader("📍 地理位置")

    # 从 Excel 加载数据
    td_data = load_data_from_excel()
    provinces = sorted(td_data.keys())

    # 选择省份
    selected_province = st.selectbox("选择省份", provinces)

    # 根据省份显示城市
    cities_in_province = get_cities_by_province(td_data, selected_province)
    if cities_in_province:
        default_city_index = 0
        if "上海" in cities_in_province:
            default_city_index = cities_in_province.index("上海")
        city_sel = st.selectbox("选择城市", cities_in_province, index=default_city_index)
    else:
        city_sel = st.text_input("输入城市名称", value="自定义城市")

    # 显示当前城市的年雷暴日
    current_td = get_td_by_city(td_data, city_sel) if city_sel else None
    if current_td is not None:
        st.caption(f"📍 {city_sel} 年雷暴日：{current_td:.1f} d/a")
    else:
        st.caption("⚠️ 该城市数据未录入，请手动输入年雷暴日")

    # 手动输入选项
    use_manual_city = st.checkbox("手动输入年雷暴日")

    if use_manual_city:
        manual_td = st.number_input("年雷暴日 Td (d/a)", min_value=0.1, value=30.0, step=0.1)
        manual_city = city_sel if city_sel else "自定义城市"
        city = manual_city
    else:
        manual_city = city_sel if city_sel else "上海"
        manual_td = current_td if current_td is not None else 30.0
        city = city_sel if city_sel else "上海"

    # ---- 分类3：周边与校正 ----
    st.subheader("🏗️ 周边与校正")
    surrounding_type = st.selectbox("周边情况", get_surrounding_options())

    Lz = 0.0
    if surrounding_type in [
        "周边2D范围内有等高或比它低的其他建筑物，但不在所考虑建筑物以外100m的保护范围内",
        "周边2D范围内有比它高的其他建筑物",
    ]:
        Lz = st.number_input("与所考虑建筑物边长平行长度总和 Lz (m)",
                             min_value=0.0, value=260.0, step=10.0)

    building_condition = st.selectbox("建筑物状况（校正系数）", get_correction_factors())

    # ---- 分类4：建筑物属性 ----
    st.subheader("🏷️ 建筑物属性")
    building_attr = st.selectbox("建筑物属性（防雷等级判定）", get_building_attributes())

    # ---- 分类5：电子防护参数 ----
    st.subheader("🔌 电子防护参数")
    soil_resistivity = st.number_input("土壤电阻率 ρs (Ω·m)",
                                        min_value=0.0, value=500.0, step=50.0)
    L_cable = st.number_input("线路长度 L (m)",
                               min_value=0.0, value=200.0, step=10.0,
                               help="所考虑建筑物至网络的第一个分支点或相邻建筑物的长度")

# ==================== Tab 1: 防雷等级 ====================
tab1, tab2, tab3, tab4 = st.tabs(["📊 防雷等级计算", "🛡️ 电子信息防护等级", "📁 年雷暴日数据管理", "📤 导出设置"])

with tab1:
    st.header("📊 防雷等级计算")
    st.caption("依据 GB 50057-2010《建筑物防雷设计规范》")

    col_info, col_result = st.columns([1, 1])

    with col_info:
        st.subheader("📋 当前输入参数")
        df_input = pd.DataFrame([
            ["建筑物长 L", f"{L} m"],
            ["建筑物宽 W", f"{W} m"],
            ["建筑物高 H", f"{H} m"],
            ["城市", city],
            ["周边情况", surrounding_type],
            ["Lz（平行长度总和）", f"{Lz} m" if Lz > 0 else "不适用"],
            ["建筑物状况", building_condition],
            ["建筑物属性", building_attr],
        ], columns=["参数", "值"])
        st.dataframe(df_input, width="stretch", hide_index=True)

    with col_result:
        if st.button("计算防雷等级", key="calc_lp", type="primary", width="stretch"):
            with st.spinner("计算中..."):
                if use_manual_city:
                    # 使用手动输入的年雷暴日
                    # 临时修改数据字典
                    td_data_local = load_data_from_excel()
                    td_data_local[selected_province][manual_city] = manual_td
                    # 重新赋值 city
                    city = manual_city
                else:
                    city = city_sel

                result = calc_lightning_protection(
                    L=L, W=W, H=H,
                    city=city,
                    building_condition=building_condition,
                    surrounding_type=surrounding_type,
                    Lz=Lz
                )

                st.session_state.last_lp_result = result

                attr_type = BUILDING_ATTR.get(building_attr, "normal")
                level_num = get_lightning_level(result["N"], attr_type)
                level_map = {1: "一类", 2: "二类", 3: "三类"}
                level_text = level_map.get(level_num, "未达到设防要求")

                st.subheader("📋 计算过程")
                df_calc = pd.DataFrame([
                    ["年雷暴日 Td", f"{result['Td']:.2f} d/a"],
                    ["雷击大地密度 Ng", f"{result['Ng']:.4f} 次/(km²·a)"],
                    ["扩大宽度 D", f"{result['D']:.3f} m"],
                    ["等效面积 Ae", f"{result['Ae']:.6f} km²"],
                    ["校正系数 k", f"{result['k']:.2f}"],
                    ["年预计雷击次数 N", f"{result['N']:.6f} 次/a"],
                ], columns=["参数", "值"])
                st.dataframe(df_calc, width="stretch", hide_index=True)

                st.subheader("🏷️ 判定结果")
                if level_num:
                    st.success(f"✅ 防雷等级：**{level_text}**")
                else:
                    st.warning("⚠️ 可不设防雷装置")

                with st.expander("📖 规范依据（GB 50057-2010）"):
                    st.markdown(f"""
                    - 年雷暴日 Td = {result['Td']:.2f} d/a（{city}）
                    - 雷击大地密度 Ng = 0.1 × Td = {result['Ng']:.4f} 次/(km²·a)
                    - 扩大宽度 D = √(H(200-H)) = {result['D']:.3f} m
                    - 等效面积 Ae = {result['Ae']:.6f} km²
                    - 校正系数 k = {result['k']:.2f}
                    - 年预计雷击次数 N = k×Ng×Ae = {result['N']:.6f} 次/a
                    - 防雷等级：**{level_text}**
                    """)

with tab2:
    st.header("🛡️ 电子信息系统雷电防护等级计算")
    st.caption("依据 GB 50343-2012《建筑物电子信息系统防雷技术规范》")

    st.info("ℹ️ 复用防雷等级计算中的 N1 和 Ng 结果")

    col_param1, col_param2 = st.columns(2)

    with col_param1:
        st.subheader("📥 入户设施参数")

        power_cable_display = st.selectbox("电源线缆入户方式", get_power_cable_options())
        power_config = get_power_cable_coef(power_cable_display)
        power_cable_type = power_config["type"]
        power_cable_coef = power_config["coef"]

        if power_cable_type == "fiber":
            st.caption(f"✅ {power_cable_display} → Ae1' = 0")
        elif power_cable_type == "overhead":
            st.caption(f"✅ {power_cable_display} → Ae1' = {power_cable_coef:.1f} × L × 10⁻⁶")
            st.caption(f"   L = {L_cable} m")
        else:
            ds_display = min(soil_resistivity, 500) if soil_resistivity > 0 else 500
            st.caption(f"✅ {power_cable_display} → Ae1' = {power_cable_coef:.1f} × ds × L × 10⁻⁶")
            st.caption(f"   ds = min(ρs, 500) = {ds_display:.1f} m, L = {L_cable} m")

        signal_cable_display = st.selectbox("信号线缆入户方式", get_signal_cable_options())
        signal_config = get_signal_cable_coef(signal_cable_display)
        signal_cable_type = signal_config["type"]
        signal_cable_coef = signal_config["coef"]

        if signal_cable_type == "fiber":
            st.caption(f"✅ {signal_cable_display} → Ae2' = 0")
        elif signal_cable_type == "overhead":
            st.caption(f"✅ {signal_cable_display} → Ae2' = {signal_cable_coef:.1f} × L × 10⁻⁶")
            st.caption(f"   L = {L_cable} m")
        else:
            ds_display = min(soil_resistivity, 500) if soil_resistivity > 0 else 500
            st.caption(f"✅ {signal_cable_display} → Ae2' = {signal_cable_coef:.1f} × ds × L × 10⁻⁶")
            st.caption(f"   ds = min(ρs, 500) = {ds_display:.1f} m, L = {L_cable} m")

    with col_param2:
        st.subheader("📊 C1~C6 因子选择")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            c1_type = st.selectbox("C1 材料结构", get_c1_options())
            c1_val = C1_OPTIONS[c1_type]
            st.caption(f"值：{c1_val:.2f}")

        with col_c2:
            c2_type = st.selectbox("C2 重要程度", get_c2_options())
            c2_val = C2_OPTIONS[c2_type]
            st.caption(f"值：{c2_val:.2f}")

        col_c3, col_c4 = st.columns(2)
        with col_c3:
            c3_type = st.selectbox("C3 耐冲击能力", get_c3_options())
            c3_val = C3_OPTIONS[c3_type]
            st.caption(f"值：{c3_val:.2f}")

        with col_c4:
            c4_type = st.selectbox("C4 雷电防护区", get_c4_options())
            c4_val = C4_OPTIONS[c4_type]
            st.caption(f"值：{c4_val:.2f}")

        col_c5, col_c6 = st.columns(2)
        with col_c5:
            c5_type = st.selectbox("C5 事故后果", get_c5_options())
            c5_val = C5_OPTIONS[c5_type]
            st.caption(f"值：{c5_val:.2f}")

        with col_c6:
            c6_type = st.selectbox("C6 区域雷暴等级", get_c6_options())
            c6_val = C6_OPTIONS[c6_type]
            st.caption(f"值：{c6_val:.2f}")

    if st.button("计算防护等级", key="calc_ep", type="primary", width="stretch"):
        with st.spinner("计算中..."):
            try:
                if use_manual_city:
                    city = manual_city
                else:
                    city = city_sel

                lp_result = calc_lightning_protection(
                    L=L, W=W, H=H,
                    city=city,
                    building_condition=building_condition,
                    surrounding_type=surrounding_type,
                    Lz=Lz
                )

                N1 = lp_result["N"]
                Ng = lp_result["Ng"]

                st.session_state.last_lp_result = lp_result

                ep_result = calc_electronic_protection(
                    N1=N1,
                    Ng=Ng,
                    L_cable=L_cable,
                    soil_resistivity=soil_resistivity,
                    power_cable_type=power_cable_type,
                    power_cable_coef=power_cable_coef,
                    signal_cable_type=signal_cable_type,
                    signal_cable_coef=signal_cable_coef,
                    C1_val=c1_val,
                    C2_val=c2_val,
                    C3_val=c3_val,
                    C4_val=c4_val,
                    C5_val=c5_val,
                    C6_val=c6_val,
                    C1_type=c1_type,
                    C2_type=c2_type,
                    C3_type=c3_type,
                    C4_type=c4_type,
                    C5_type=c5_type,
                    C6_type=c6_type,
                    power_cable_display=power_cable_display,
                    signal_cable_display=signal_cable_display,
                )

                st.session_state.last_ep_result = ep_result

                col1, col2 = st.columns(2)

                with col1:
                    st.subheader("📋 计算过程")
                    df_ep = pd.DataFrame([
                        ["N1（建筑物年预计雷击次数）", f"{ep_result['N1']:.6f} 次/a"],
                        ["Ae1'（电源线缆截收面积）", f"{ep_result['Ae1']:.6f} km²"],
                        ["Ae2'（信号线缆截收面积）", f"{ep_result['Ae2']:.6f} km²"],
                        ["N2（入户设施年预计雷击次数）", f"{ep_result['N2']:.6f} 次/a"],
                        ["N（总年预计雷击次数）", f"{ep_result['N']:.6f} 次/a"],
                        ["C1", f"{ep_result['C1']:.2f} ({ep_result['C1_type']})"],
                        ["C2", f"{ep_result['C2']:.2f} ({ep_result['C2_type']})"],
                        ["C3", f"{ep_result['C3']:.2f} ({ep_result['C3_type']})"],
                        ["C4", f"{ep_result['C4']:.2f} ({ep_result['C4_type']})"],
                        ["C5", f"{ep_result['C5']:.2f} ({ep_result['C5_type']})"],
                        ["C6", f"{ep_result['C6']:.2f} ({ep_result['C6_type']})"],
                        ["C = C1+C2+C3+C4+C5+C6", f"{ep_result['C']:.2f}"],
                        ["Nc = 0.58/C", f"{ep_result['Nc']:.6f} 次/a"],
                        ["E = 1 - Nc/N", f"{ep_result['E']:.4f}"],
                    ], columns=["参数", "值"])
                    st.dataframe(df_ep, width="stretch", hide_index=True)

                with col2:
                    st.subheader("🏷️ 判定结果")
                    st.metric("总年预计雷击次数 N", f"{ep_result['N']:.6f} 次/a")
                    st.metric("拦截效率 E", f"{ep_result['E']:.4f}")

                    level = ep_result['level']
                    if "A" in level:
                        st.success(f"✅ **{level}**（最高防护要求）")
                    elif "B" in level:
                        st.success(f"✅ **{level}**（较高防护要求）")
                    elif "C" in level:
                        st.warning(f"⚠️ **{level}**（中等防护要求）")
                    elif "D" in level:
                        st.info(f"ℹ️ **{level}**（基本防护要求）")
                    else:
                        st.info(f"ℹ️ {level}")

                with st.expander("📖 规范依据（GB 50343-2012）"):
                    st.markdown(f"""
                    **计算依据：GB 50343-2012 附录A**

                    - 电源线缆截收面积 Ae1' = {ep_result['Ae1']:.6f} km²
                    - 信号线缆截收面积 Ae2' = {ep_result['Ae2']:.6f} km²
                    - N2 = Ng × (Ae1' + Ae2') = {ep_result['Ng']:.4f} × ({ep_result['Ae1']:.6f} + {ep_result['Ae2']:.6f}) = {ep_result['N2']:.6f} 次/a
                    - N = N1 + N2 = {ep_result['N']:.6f} 次/a
                    - C = C1+C2+C3+C4+C5+C6 = {ep_result['C']:.2f}
                    - Nc = 0.58/C = {ep_result['Nc']:.6f} 次/a
                    - E = 1 - Nc/N = {ep_result['E']:.4f}
                    - 防护等级：**{level}**
                    """)

            except Exception as e:
                st.error(f"计算失败：{e}")
                st.code(f"错误详情：{e}", language="python")

with tab3:
    st.header("📁 年雷暴日数据管理")
    st.caption("管理各省份城市的年雷暴日数据，数据保存在 thunderstorm_data.xlsx 文件中")

    td_data = load_data_from_excel()

    col_manage, col_preview = st.columns([1, 1])

    with col_manage:
        st.subheader("➕ 添加/修改城市")

        provinces_list = sorted(td_data.keys())
        selected_prov = st.selectbox("选择省份", provinces_list, key="manage_province")

        city_name = st.text_input("城市名称", key="manage_city")
        td_value = st.number_input("年雷暴日 (d/a)", min_value=0.0, value=30.0, step=0.1, key="manage_td")

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("添加/更新", key="add_city_btn", width="stretch"):
                if city_name:
                    td_data = add_city(selected_prov, city_name, td_value, td_data)
                    st.success(f"✅ {selected_prov} → {city_name} = {td_value} d/a 已保存")
                    st.rerun()
                else:
                    st.warning("请输入城市名称")

        with col_btn2:
            if st.button("删除城市", key="del_city_btn", width="stretch"):
                if city_name:
                    td_data = delete_city(selected_prov, city_name, td_data)
                    st.success(f"✅ {city_name} 已删除")
                    st.rerun()
                else:
                    st.warning("请输入要删除的城市名称")

        st.subheader("➕ 新增省份")
        new_prov = st.text_input("新省份名称", key="new_province")
        if st.button("新增省份", key="add_prov_btn", width="stretch"):
            if new_prov and new_prov not in td_data:
                td_data[new_prov] = {}
                save_data_to_excel(td_data)
                st.success(f"✅ 省份 {new_prov} 已添加")
                st.rerun()
            elif new_prov in td_data:
                st.warning("该省份已存在")
            else:
                st.warning("请输入省份名称")

    with col_preview:
        st.subheader("📊 数据预览")
        rows = []
        for prov, cities in td_data.items():
            for cty, tdv in cities.items():
                rows.append({"省份": prov, "城市": cty, "年雷暴日 (d/a)": tdv})

        if rows:
            df_display = pd.DataFrame(rows)
            st.dataframe(df_display, width="stretch", hide_index=True)
            st.caption(f"共 {len(rows)} 个城市")
        else:
            st.info("暂无数据")

with tab4:
    st.header("📤 导出 Word 计算书")
    st.caption("将当前计算结果导出为 Word 文档")

    if 'last_lp_result' in st.session_state:
        st.subheader("📋 导出选项")

        col_opt1, col_opt2 = st.columns(2)

        with col_opt1:
            export_content = st.radio(
                "选择导出内容",
                options=["仅防雷等级", "仅电子防护等级", "两者都导出"],
                index=2,
                help="选择要导出的计算内容"
            )

        with col_opt2:
            if export_content == "两者都导出":
                export_mode = st.radio(
                    "文件合并方式",
                    options=["合并到一个文件", "分开两个文件"],
                    index=0,
                    help="合并：防雷+电子防护在同一份计算书中；分开：各生成独立文件"
                )
            else:
                export_mode = "合并到一个文件"
                st.info("ℹ️ 仅导出单项内容，将生成单个文件")

        st.subheader("📄 文件名")
        now_str = datetime.now().strftime('%Y%m%d_%H%M%S')

        if export_content == "两者都导出" and export_mode == "分开两个文件":
            col_name1, col_name2 = st.columns(2)
            with col_name1:
                lp_filename = st.text_input("防雷等级计算书文件名",
                                            value=f"防雷等级计算书_{now_str}.docx")
            with col_name2:
                ep_filename = st.text_input("电子防护等级计算书文件名",
                                            value=f"电子防护等级计算书_{now_str}.docx")
        else:
            default_filename = f"防雷计算书_{now_str}.docx"
            if export_content == "仅电子防护等级":
                default_filename = f"电子防护等级计算书_{now_str}.docx"
            filename = st.text_input("文件名", value=default_filename)

        if st.button("📥 导出 Word 计算书", key="export_word", type="primary", width="stretch"):
            with st.spinner("正在生成计算书..."):
                try:
                    from word_export import export_calculation_report

                    building_params = {
                        "建筑物长 L": f"{L} m",
                        "建筑物宽 W": f"{W} m",
                        "建筑物高 H": f"{H} m",
                        "城市": city,
                        "周边情况": surrounding_type,
                        "建筑物状况": building_condition,
                        "建筑物属性": building_attr,
                        "土壤电阻率": f"{soil_resistivity} Ω·m",
                        "线路长度": f"{L_cable} m",
                    }

                    cable_params = {
                        "电源线缆入户方式": power_cable_display if 'power_cable_display' in locals() else "未选择",
                        "信号线缆入户方式": signal_cable_display if 'signal_cable_display' in locals() else "未选择",
                    }

                    c_factors = {
                        "C1": {"type": c1_type, "val": c1_val},
                        "C2": {"type": c2_type, "val": c2_val},
                        "C3": {"type": c3_type, "val": c3_val},
                        "C4": {"type": c4_type, "val": c4_val},
                        "C5": {"type": c5_type, "val": c5_val},
                        "C6": {"type": c6_type, "val": c6_val},
                    }

                    attr_type = BUILDING_ATTR.get(building_attr, "normal")
                    level_num = get_lightning_level(st.session_state.last_lp_result["N"], attr_type)
                    level_map = {1: "一类", 2: "二类", 3: "三类"}
                    level_text = level_map.get(level_num, "未达到设防要求")

                    if export_content == "仅防雷等级":
                        mode = 'combined'
                        has_ep = False
                        ep_result = None
                        ep_level_text = None
                    elif export_content == "仅电子防护等级":
                        if 'last_ep_result' not in st.session_state:
                            st.error("❌ 请先在「电子信息防护等级」页面中计算一次")
                            st.stop()
                        mode = 'combined'
                        has_ep = True
                        ep_result = st.session_state.last_ep_result
                        ep_level_text = ep_result.get("level", "未计算")
                    else:
                        has_ep = True
                        ep_result = st.session_state.last_ep_result if 'last_ep_result' in st.session_state else None
                        ep_level_text = ep_result.get("level", "未计算") if ep_result else "未计算"
                        mode = 'separate' if export_mode == "分开两个文件" else 'combined'

                    result = export_calculation_report(
                        lp_result=st.session_state.last_lp_result,
                        ep_result=ep_result,
                        building_params=building_params,
                        cable_params=cable_params,
                        c_factors=c_factors,
                        level_text=level_text,
                        ep_level_text=ep_level_text,
                        export_mode=mode,
                        has_ep=has_ep,
                        building_attr=building_attr,
                        attr_type=attr_type
                    )

                    if mode == 'separate':
                        lp_buffer, ep_buffer = result
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            st.download_button(
                                label="📥 下载防雷等级计算书",
                                data=lp_buffer,
                                file_name=lp_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                width="stretch"
                            )
                        with col_dl2:
                            st.download_button(
                                label="📥 下载电子防护等级计算书",
                                data=ep_buffer,
                                file_name=ep_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                width="stretch"
                            )
                        st.success("✅ 两份计算书已生成，请分别点击下载")
                    else:
                        st.download_button(
                            label="📥 下载 Word 计算书",
                            data=result,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            width="stretch"
                        )
                        st.success("✅ 计算书生成完成，请点击下载")

                except Exception as e:
                    st.error(f"导出失败：{e}")
                    st.code(f"错误详情：{e}", language="python")

    else:
        st.info("ℹ️ 请先在「防雷等级计算」或「电子信息防护等级」页面中计算一次，再导出计算书")
        st.markdown("""
        **操作步骤：**
        1. 在左侧输入参数
        2. 点击「计算防雷等级」或「计算防护等级」
        3. 返回此页面导出
        """)

# ==================== 页脚 ====================
st.divider()
st.caption("⚡ LPCalc v1.0 | 计算结果仅供参考，正式设计请以专业审核为准")