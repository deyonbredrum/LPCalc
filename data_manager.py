# -*- coding: utf-8 -*-
"""
年雷暴日数据管理模块
支持 Excel 文件读写和页面编辑
"""

import os
import pandas as pd
from params import PROVINCE_CITY_TD

DATA_FILE = "thunderstorm_data.xlsx"


def load_data_from_excel():
    """
    从 Excel 文件加载年雷暴日数据
    如果文件不存在，则从 params.py 中的默认数据创建
    """
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_excel(DATA_FILE, dtype={"省份": str, "城市": str, "年雷暴日": float})
            # 转换为字典格式
            data = {}
            for _, row in df.iterrows():
                province = row["省份"]
                city = row["城市"]
                td = row["年雷暴日"]
                if province not in data:
                    data[province] = {}
                data[province][city] = td
            return data
        except Exception as e:
            print(f"读取 Excel 文件失败: {e}")
            return PROVINCE_CITY_TD.copy()
    else:
        # 创建默认文件
        save_data_to_excel(PROVINCE_CITY_TD)
        return PROVINCE_CITY_TD.copy()


def save_data_to_excel(data):
    """
    将年雷暴日数据保存到 Excel 文件
    """
    rows = []
    for province, cities in data.items():
        for city, td in cities.items():
            rows.append({"省份": province, "城市": city, "年雷暴日": td})

    df = pd.DataFrame(rows)
    df.to_excel(DATA_FILE, index=False, sheet_name="年雷暴日数据")
    return df


def add_city(province, city, td, data):
    """添加或更新城市数据"""
    if province not in data:
        data[province] = {}
    data[province][city] = td
    save_data_to_excel(data)
    return data


def delete_city(province, city, data):
    """删除城市数据"""
    if province in data and city in data[province]:
        del data[province][city]
        # 如果省份下没有城市了，删除省份
        if not data[province]:
            del data[province]
        save_data_to_excel(data)
    return data


def get_all_cities(data):
    """获取所有城市列表"""
    cities = []
    for province, cities_dict in data.items():
        for city in cities_dict.keys():
            cities.append(city)
    return sorted(cities)


def get_cities_by_province(data, province):
    """获取某省份的所有城市"""
    return sorted(data.get(province, {}).keys())


def get_td_by_city(data, city):
    """根据城市名获取年雷暴日"""
    for province, cities in data.items():
        if city in cities:
            return cities[city]
    return None