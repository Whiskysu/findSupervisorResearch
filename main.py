import pandas as pd
import webbrowser
from pypinyin import pinyin, Style
import urllib.parse
import time


def load_data(excel_file):
    # 从Excel文件加载数据
    df = pd.read_excel(excel_file)
    return df


def name_to_pinyin(name):
    # 将中文名字转换为拼音
    py = pinyin(name, style=Style.NORMAL)
    return " ".join([item[0] for item in py])


def search_baidu(name, school):
    # 使用百度搜索
    query = urllib.parse.quote(f"{name} {school}")
    url = f"https://www.baidu.com/s?wd={query}"
    webbrowser.open(url)


def search_google(name, school):
    # 使用Google搜索
    query = urllib.parse.quote(f"{name} {school}")
    url = f"https://www.google.com/search?q={query}"
    webbrowser.open(url)


def search_ieee(name, school_cn, school_en):
    # 使用IEEE搜索
    name_pinyin = name_to_pinyin(name)
    name_parts = name_pinyin.split()

    if len(name_parts) == 2:
        last_name, first_name = name_parts
    elif len(name_parts) >= 3:
        last_name, first_name = name_parts[0], "".join(name_parts[1:])
    else:
        print(f"警告：姓名 '{name}' 格式不正确")
        return
    query = urllib.parse.quote(
        f'("Authors":{first_name} {last_name}) AND ("Author Affiliations":{school_en})'
    )
    url = f"https://ieeexplore.ieee.org/search/searchresult.jsp?action=search&newsearch=true&matchBoolean=true&queryText={query}"
    webbrowser.open(url)


def search_scopus(name, school_cn, school_en):
    # 使用Scopus搜索
    name_pinyin = name_to_pinyin(name)
    name_parts = name_pinyin.split()

    if len(name_parts) == 2:
        last_name, first_name = name_parts
    elif len(name_parts) >= 3:
        last_name, first_name = name_parts[0], "".join(name_parts[1:])
    else:
        print(f"警告：姓名 '{name}' 格式不正确")
        return

    school_encoded = school_en.replace(" ", "+")
    url = f"https://www.scopus.com/results/authorNamesList.uri?st1={last_name}&st2={first_name}&institute={school_encoded}&origin=searchauthorlookup"
    webbrowser.open(url)


def search_with_delay(name, school_cn, school_en, delay):
    search_functions = [search_baidu, search_google, search_ieee, search_scopus]
    # search_functions = [search_scopus]

    for func in search_functions:
        if func in [search_ieee, search_scopus]:
            func(name, school_cn, school_en)
        else:
            func(name, school_cn)
        time.sleep(delay)  # 在每次搜索之间添加延迟


def main():
    excel_file = "supervisorInfo.xlsx"
    df = load_data(excel_file)

    delay = 1.0231

    for index, row in df.iterrows():
        name = row["姓名"]
        school_cn = row["学校"]
        school_en = row["学校英文名"]

        print(f"正在搜索：{name} - {school_cn}")
        search_with_delay(name, school_cn, school_en, delay)
        print(f"完成搜索：{name} - {school_cn}")
        print("---")


if __name__ == "__main__":
    main()
