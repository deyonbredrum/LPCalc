# -*- coding: utf-8 -*-
"""
年雷暴日数据管理模块
直接操作字典，无需 Excel 文件
"""

from params import PROVINCE_CITY_TD


def load_data():
    """直接返回字典数据"""
    return PROVINCE_CITY_TD


def save_data(data):
    """更新字典数据"""
    PROVINCE_CITY_TD.clear()
    PROVINCE_CITY_TD.update(data)
    # 不写入文件，只更新内存


def add_city(province, city, td):
    """添加或更新城市"""
    if province not in PROVINCE_CITY_TD:
        PROVINCE_CITY_TD[province] = {}
    PROVINCE_CITY_TD[province][city] = td


def delete_city(province, city):
    """删除城市"""
    if province in PROVINCE_CITY_TD and city in PROVINCE_CITY_TD[province]:
        del PROVINCE_CITY_TD[province][city]
        if not PROVINCE_CITY_TD[province]:
            del PROVINCE_CITY_TD[province]


def get_cities_by_province(province):
    """获取某省份的城市列表"""
    return sorted(PROVINCE_CITY_TD.get(province, {}).keys())


def get_td_by_city(city):
    """获取城市的年雷暴日"""
    for province, cities in PROVINCE_CITY_TD.items():
        if city in cities:
            return cities[city]
    return None


def get_all_cities():
    """获取所有城市"""
    cities = []
    for province, cities_dict in PROVINCE_CITY_TD.items():
        for city in cities_dict.keys():
            cities.append(city)
    return sorted(cities)


def get_all_provinces():
    """获取所有省份"""
    return sorted(PROVINCE_CITY_TD.keys())