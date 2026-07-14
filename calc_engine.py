# -*- coding: utf-8 -*-
"""
GB 50057-2010 + GB 50343-2012 计算引擎
"""

import math
from params import (
    CORRECTION_FACTOR, BUILDING_ATTR,
    get_lightning_level, get_electronic_level,
    C1_OPTIONS, C2_OPTIONS, C3_OPTIONS, C4_OPTIONS, C5_OPTIONS, C6_OPTIONS,
    SURROUNDING_OPTIONS
)


def calc_equivalent_area(L, W, H, surrounding_type, Lz=0):
    """计算等效面积 Ae (km²)"""
    D = math.sqrt(H * (200 - H))

    if surrounding_type == "周边2D范围内无其他建筑":
        Ae = L * W + 2 * (L + W) * D + math.pi * D ** 2
    elif surrounding_type == "周边2D范围内有等高或比它低的其他建筑物，但不在所考虑建筑物以外100m的保护范围内":
        Ae = L * W + 2 * (L + W) * D + math.pi * D ** 2 - (Lz * D) / 2
    elif surrounding_type == "周边2D范围内有等高或比它低的其他建筑物":
        Ae = L * W + (L + W) * D + (math.pi * D ** 2) / 4
    elif surrounding_type == "周边2D范围内有比它高的其他建筑物":
        Ae = L * W + 2 * (L + W) * D + math.pi * D ** 2 - Lz * D
    elif surrounding_type == "周边2D范围内都有比它高的其他建筑物":
        Ae = L * W
    else:
        Ae = L * W + 2 * (L + W) * D + math.pi * D ** 2

    if Ae < 0:
        Ae = 0
    return Ae * 1e-6


def calc_lightning_protection(L, W, H, city, building_condition,
                              surrounding_type="周边2D范围内无其他建筑", Lz=0):
    """计算防雷等级"""
    from params import PROVINCE_CITY_TD

    # 从字典中获取年雷暴日
    Td = None
    for province, cities in PROVINCE_CITY_TD.items():
        if city in cities:
            Td = cities[city]
            break

    if Td is None:
        Td = 30.0  # 默认值

    Ng = 0.1 * Td
    Ae = calc_equivalent_area(L, W, H, surrounding_type, Lz)
    k = CORRECTION_FACTOR.get(building_condition, 1.0)
    N = k * Ng * Ae
    D = math.sqrt(H * (200 - H))

    return {
        "Td": Td,
        "Ng": Ng,
        "Ae": Ae,
        "k": k,
        "N": N,
        "D": D,
        "city": city,
        "L": L,
        "W": W,
        "H": H,
        "surrounding_type": surrounding_type,
        "Lz": Lz,
        "building_condition": building_condition,
    }


def calc_ae_cable(L, ds, cable_type, coef):
    """
    计算入户设施截收面积 (GB 50343-2012 附录A)

    参数:
        L: 线路长度 (m)
        ds: 埋地线缆等效宽度 (m)，取 min(土壤电阻率, 500)
        cable_type: "overhead" 架空, "buried" 埋地, "fiber" 光纤
        coef: 系数 (2000, 500, 2, 0.1, 0)

    返回:
        Ae: 截收面积 (km²)
    """
    if cable_type == "fiber":
        return 0.0

    if cable_type == "overhead":
        # 架空线缆: Ae = coef × L × 10⁻⁶
        Ae = coef * L * 1e-6
    else:  # buried
        # 埋地线缆: Ae = coef × ds × L × 10⁻⁶
        Ae = coef * ds * L * 1e-6

    return Ae


def calc_electronic_protection(N1, Ng, L_cable, soil_resistivity,
                               power_cable_type, power_cable_coef,
                               signal_cable_type, signal_cable_coef,
                               C1_val, C2_val, C3_val, C4_val, C5_val, C6_val,
                               C1_type, C2_type, C3_type, C4_type, C5_type, C6_type,
                               power_cable_display, signal_cable_display):
    """
    计算电子信息系统雷电防护等级 (GB 50343-2012)
    """
    # 计算 ds = min(土壤电阻率, 500)
    ds = min(soil_resistivity, 500) if soil_resistivity > 0 else 1.0

    # 1. 计算 Ae1' (电源线缆截收面积)
    Ae1 = calc_ae_cable(L_cable, ds, power_cable_type, power_cable_coef)

    # 2. 计算 Ae2' (信号线缆截收面积)
    Ae2 = calc_ae_cable(L_cable, ds, signal_cable_type, signal_cable_coef)

    # 3. N2 = Ng * (Ae1 + Ae2)
    N2 = Ng * (Ae1 + Ae2)

    # 4. N = N1 + N2
    N = N1 + N2

    # 5. C = C1 + C2 + C3 + C4 + C5 + C6
    C = C1_val + C2_val + C3_val + C4_val + C5_val + C6_val

    # 6. Nc = 0.58 / C
    Nc = 0.58 / C if C > 0 else 999

    # 7. E = 1 - Nc / N
    E = 1 - Nc / N if N > 0 else 0

    # 8. 防护等级
    level = get_electronic_level(E)

    return {
        "N1": N1,
        "N2": N2,
        "N": N,
        "C": C,
        "Nc": Nc,
        "E": E,
        "level": level,
        "L_cable": L_cable,
        "soil_resistivity": soil_resistivity,
        "ds": ds,
        "Ae1": Ae1,
        "Ae2": Ae2,
        "power_cable_type": power_cable_type,
        "signal_cable_type": signal_cable_type,
        "power_cable_display": power_cable_display,
        "signal_cable_display": signal_cable_display,
        "power_cable_coef": power_cable_coef,
        "signal_cable_coef": signal_cable_coef,
        "C1": C1_val,
        "C2": C2_val,
        "C3": C3_val,
        "C4": C4_val,
        "C5": C5_val,
        "C6": C6_val,
        "C1_type": C1_type,
        "C2_type": C2_type,
        "C3_type": C3_type,
        "C4_type": C4_type,
        "C5_type": C5_type,
        "C6_type": C6_type,
        "Ng": Ng,
    }


def get_available_cities():
    from data_manager import load_data_from_excel, get_all_cities
    td_data = load_data_from_excel()
    return get_all_cities(td_data)


def get_correction_factors():
    return list(CORRECTION_FACTOR.keys())


def get_building_attributes():
    return list(BUILDING_ATTR.keys())


def get_c1_options():
    return list(C1_OPTIONS.keys())


def get_c2_options():
    return list(C2_OPTIONS.keys())


def get_c3_options():
    return list(C3_OPTIONS.keys())


def get_c4_options():
    return list(C4_OPTIONS.keys())


def get_c5_options():
    return list(C5_OPTIONS.keys())


def get_c6_options():
    return list(C6_OPTIONS.keys())


def get_surrounding_options():
    return SURROUNDING_OPTIONS


def get_power_cable_options():
    from params import POWER_CABLE_OPTIONS
    return list(POWER_CABLE_OPTIONS.keys())


def get_signal_cable_options():
    from params import SIGNAL_CABLE_OPTIONS
    return list(SIGNAL_CABLE_OPTIONS.keys())


def get_power_cable_coef(display_name):
    from params import POWER_CABLE_OPTIONS
    return POWER_CABLE_OPTIONS[display_name]


def get_signal_cable_coef(display_name):
    from params import SIGNAL_CABLE_OPTIONS
    return SIGNAL_CABLE_OPTIONS[display_name]