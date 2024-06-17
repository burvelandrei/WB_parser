import requests
import json
import openpyxl
import time
import pandas as pd
from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem

# изменить на количество страниц которые хотите распарсить в подкатегории
count_page = 20


def rotation_user_agent():
    """функция возвращает рандомный user agent"""
    software_names = [SoftwareName.CHROME.value]
    operating_systems = [OperatingSystem.WINDOWS.value, OperatingSystem.LINUX.value]
    user_agent_rotator = UserAgent(software_names=software_names, operating_systems=operating_systems, limit=100)
    user_agent = user_agent_rotator.get_random_user_agent()
    return user_agent


def get_catalogs_wb() -> dict:
    """функция для получения от вб json со всеми данными о категориях и их подкатегориях"""
    """получаем полный каталог Wildberries"""
    url = "https://static-basket-01.wbbasket.ru/vol0/data/main-menu-ru-ru-v2.json"
    headers = {
        "Accept": "*/*",
        "User-Agent": f'{rotation_user_agent()}'
    }
    response = requests.get(url, headers=headers).json()
    return response


def get_data_category(catalogs_wb: dict) -> list:
    """фунция для парсинга каталогов и точка входа в скрипт"""
    print("Парсинг данных начат")
    # создание таблицы
    create_excel()
    for category in catalogs_wb:
        # если категория имеет подкатегории, то опраялем её на дополнительный парсинг по подкатегориям
        if isinstance(category, dict) and "childs" in category:
            category_name = f"{category['name']}"
            all_podcategory = get_data_podcategory(category["childs"], category_name)
            write_excel(all_podcategory, category_name)
            print(f"Категория {category_name} обработана")
        # если категория не имеет подкатегорий, то обрабатываем контент в данной категории
        elif isinstance(category, dict):
            if "shard" in category and "query" in category:
                category_name = f"{category['name']}"
                shard = category["shard"]
                query = category["query"]
                write_excel(get_content(shard, query), category_name, level = 1)
    # удаляем первый базовый лист
    delete_first_sheet()
    print("Парсинг данных окончен")


def get_data_podcategory(podcategorys: dict, category_name: str, level = 1) -> list:
    """функция для парсинга подкатегорий"""
    all_podcategory = []
    # если в подкатегории нет своих подкатегорий, то отправляем на парсинг контента данной подкатегории
    if isinstance(podcategorys, dict) and "childs" not in podcategorys:
        all_podcategory.append(
            {
                "id": f"{podcategorys['id']}",
                "name": f"{podcategorys['name']}",
                "level": level
            }
        )
        if "shard" in podcategorys and "query" in podcategorys:
            shard = podcategorys["shard"]
            query = podcategorys["query"]
            all_podcategory.extend(get_content(shard, query, level))
    # если подкатегория имеет свои подкатегории, то отправляем подкатегорию в дополнительный парсинг её подкатегорий
    elif isinstance(podcategorys, dict):
        all_podcategory.append(
            {
                "id": f"{podcategorys['id']}",
                "name": f"{podcategorys['name']}",
                "level": level
            }
        )
        level += 1
        all_podcategory.extend(
            get_data_podcategory(podcategorys["childs"], category_name, level)
        )
    else:
        for child in podcategorys:
            all_podcategory.extend(get_data_podcategory(child, category_name, level))
    return all_podcategory


def get_data_from_json(json_file: dict, level: int):
    """извлекаем из json интересующие нас данные"""
    products = []
    for data in json_file['data']['products']:
        id = data.get("id")
        name = data.get("name")
        products.append({
            'id': id,
            'name': name,
            'level': level})
    return products


def get_content(shard: str, query: str, level: int):
    """функция для получения контента по переданной категории или подкатегории"""
    data_list = []
    level += 1
    for page in range(1, count_page + 1):
        try:
            headers = {'Accept': "*/*", "User-Agent": f'{rotation_user_agent()}'}
            url = f'https://catalog.wb.ru/catalog/{shard}/catalog?appType=1&curr=rub' \
                f'&dest=-1257786' \
                f'&locale=ru' \
                f'&page={page}' \
                f'&sort=popular&spp=0' \
                f'&{query}'
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                data = response.json()
                products = get_data_from_json(data, level)
                if len(products) > 0:
                    data_list.extend(products)
        except Exception as e:
            print(e)
            continue
    print("Подкатегория обработана")
    return data_list


def create_excel():
    """функция для создания таблицы Эксель"""
    wb = openpyxl.Workbook()
    wb.save("parser.xlsx")


def delete_first_sheet():
    """функция для удаления стартового листа"""
    wb = openpyxl.load_workbook("parser.xlsx")
    wb.remove(wb["Sheet"])
    wb.save("parser.xlsx")


def write_excel(data: list, category_name: str):
    """функция для записи данных в Эксель таблицу в нужный лист"""
    df = pd.DataFrame(data)
    writer = pd.ExcelWriter("parser.xlsx", mode="a")
    df.to_excel(writer, sheet_name=category_name, index=False)
    writer.close()


if __name__ == "__main__":
    get_data_category(get_catalogs_wb())
