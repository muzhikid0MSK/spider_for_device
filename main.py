import time

import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlrd
from shutil import copyfile
import os
import docx
from termcolor import colored

doc = docx.Document()


def run_a_query(id_, country_name, device_name):
    q = device_name + "在" + country_name + "的所属基地。如果你知晓，请直接简要告知位置后再补充信息。比如xx基地。xx信息"
    req = "https://gptgo.ai/?q= " + q + "&hl=zh&hlgpt=default"
    # C:\Users\meishengke\PycharmProjects\Spider\geckodriver.exe
    browser = webdriver.Firefox(executable_path="C:\\Users\\meish\\PycharmProjects\\Spider\\geckodriver.exe")
    browser.get(req)
    try:

        elem = WebDriverWait(browser, 30).until(
            EC.visibility_of_element_located((By.ID, "downloadchat"))
        )
        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser')
        word_content = ""
        for result in soup.find_all('div'):
            if 'id' in result.attrs and result.attrs['id'] == 'ai-result':
                print(result.text)
                word_content = q, "\n", result.text, "\n"
    except:
        print(colored("Something Wrong at question " + q, "red"))
        browser.close()
    finally:
        browser.close()
        return id_, word_content


def write_to_word(file_path, word_result):
    if os.path.exists(file_path):
        os.remove(file_path)
    for result in word_result:
        doc.add_paragraph(result)
    doc.save(file_path)


def get_device_names(file_path):
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_index(0)
    ncols = worksheet.ncols
    nrows = worksheet.nrows
    first_row = worksheet.row(0)
    typ_col = 0
    device_name_col = 0
    device_names = []

    for c in range(0, ncols):
        if first_row[c].value == "军种/其它":
            typ_col = c
        if first_row[c].value == "武器中文名称":
            device_name_col = c
    for r in range(0, nrows):
        if worksheet.cell_value(r, typ_col) == "空军" or worksheet.cell_value(r, typ_col) == "海军":
            # get device name
            device_names.append(worksheet.cell_value(r, device_name_col))
    return worksheet.cell_value(1, 0), device_names


# No use
def copy_file(file_path: str):
    new_file_path = file_path.replace(".xlsx", "") + "_result.xlsx"
    if os.path.exists(new_file_path):
        return

    copyfile(file_path, new_file_path)


# No use
def result_filter(raw_result: str):
    passive_words = ["抱歉", "Error", "error", "无法", "咨询", "相关", "作为"]
    for word in passive_words:
        if word in raw_result:
            return False, raw_result
    return True, raw_result.find("所属基地是")


def main():
    import threading
    from concurrent.futures import ThreadPoolExecutor
    import multiprocessing
    file_path = "泰国武器装备.xlsx"
    copy_file(file_path)
    country_name, device_names = get_device_names(file_path)
    cnt = 0

    with ThreadPoolExecutor(multiprocessing.cpu_count() - 1) as pool:
        word_result = [""] * len(device_names)

        def call_back(future):
            id_ = future.result()[0]
            content = future.result()[1]
            word_result[id_] = content

        for i in range(0, len(device_names)):
            device_name = device_names[i]
            f = pool.submit(run_a_query, i, country_name, device_name)
            f.add_done_callback(call_back)
        # for device_name in device_names:
        #     f = pool.submit(run_a_query, country_name,device_name)
        #     f.add_done_callback(call_back)

        pool.shutdown()
        write_to_word(country_name + ".docx", word_result)


if __name__ == "__main__":
    main()
