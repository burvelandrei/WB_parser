import requests
import json
import openpyxl
import pandas as pd


def get_catalogs_wb() -> dict:
    """получаем полный каталог Wildberries"""
    url = "https://static-basket-01.wbbasket.ru/vol0/data/main-menu-ru-ru-v2.json"
    headers = {
        "Accept": "*/*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    }
    response = requests.get(url, headers=headers).json()
    with open("wb_catalogs_data.json", "w", encoding="UTF-8") as file:
        json.dump(response, file, indent=2, ensure_ascii=False)
    return response


def get_data_category(catalogs_wb: dict) -> list:
    print("Парсинг данных начат")
    create_excel()
    for category in catalogs_wb:
        if isinstance(category, dict) and "childs" in category:
            category_name = f"{category['name']}"
            all_podcategory = get_data_podcategory(category["childs"], category_name)
            write_excel(all_podcategory, category_name)
            print("Категория обработана")
        elif isinstance(category, dict):
            if "shard" in category and "query" in category:
                category_name = f"{category['name']}"
                shard = category["shard"]
                query = category["query"]
                write_excel(get_content(shard, query), category_name)
    delete_first_sheet()
    print("Парсинг данных окончен")


def get_data_podcategory(podcategorys: dict, category_name: str) -> list:
    all_podcategory = []
    if isinstance(podcategorys, dict) and "childs" not in podcategorys:
        all_podcategory.append(
            {
                "id": f"{podcategorys['id']}",
                "name": f"{podcategorys['name']}",
            }
        )
        if "shard" in podcategorys and "query" in podcategorys:
            shard = podcategorys["shard"]
            query = podcategorys["query"]
            # all_podcategory.extend(get_content(shard, query))
    elif isinstance(podcategorys, dict):
        all_podcategory.append(
            {
                "id": f"{podcategorys['id']}",
                "name": f"{podcategorys['name']}",
            }
        )
        all_podcategory.extend(
            get_data_podcategory(podcategorys["childs"], category_name)
        )
    else:
        for child in podcategorys:
            all_podcategory.extend(get_data_podcategory(child, category_name))
    return all_podcategory


def get_data_from_json(json_file):
    """извлекаем из json интересующие нас данные"""
    products = []
    for data in json_file['data']['products']:
        id = data.get("id")
        name = data.get("name")
        products.append({
            'id': id,
            'name': name,})
    return products


def get_content(shard, query):
    # вайлдбериз отдает только 100 страниц
    headers = {'Accept': "*/*", 'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    data_list = []
    for page in range(1):
        try:
            url = f'https://catalog.wb.ru/catalog/{shard}/catalog?appType=1&curr=rub&dest=-1075831,-77677,-398551,12358499' \
                f'&locale=ru&page={page}'\
                f'&reg=0&regions=64,83,4,38,80,33,70,82,86,30,69,1,48,22,66,31,40&sort=popular&spp=0&{query}'
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                data = response.json()
                print(f'Добавлено позиций: {len(get_data_from_json(data))}')
                products = get_data_from_json(data)
                if len(products) > 0:
                    data_list.extend(products)
        except:
            continue
    return data_list

def create_excel():
    wb = openpyxl.Workbook()
    wb.save("parser.xlsx")


def delete_first_sheet():
    wb = openpyxl.load_workbook("parser.xlsx")
    wb.remove(wb["Sheet"])
    wb.save("parser.xlsx")


def write_excel(data: list, category_name: str):
    df = pd.DataFrame(data)
    writer = pd.ExcelWriter("parser.xlsx", mode="a")
    df.to_excel(writer, sheet_name=category_name, index=False)
    writer.close()


if __name__ == "__main__":
    get_data_category(get_catalogs_wb())
