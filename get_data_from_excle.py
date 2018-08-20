# -*- coding: utf-8 -*-

import xlrd
import re
import json
import logging
import requests


logger = logging.getLogger()
logger.info('start worker')

def download_img(url, file_name):
    r = requests.get(url, stream=True)
    if r.status_code == 200:
        with open(file_name, 'wb') as f:
            f.write(r.content)


SHEET_NAME = "*****"
DIRECTORY_PATH = "****/*******"
EXCEL_FILE = "****.xlsx"

# 以下3行はカスタマイズ
COL_SHOP_NAME = 1
COL_CROOZ_PRODUCT_NAME = 6
COL_IMG_URL = 10

xl_bk = xlrd.open_workbook(EXCEL_FILE)
xl_sh = xl_bk.sheet_by_index(0)
sheet = xl_bk.sheet_by_name(SHEET_NAME)

# シートの全行数を取得
totalRowsNum = sheet.nrows
logging.info("totalRowsNum : " + str(sheet.nrows))

# 一行ずつ対象列のデータのみ取得
for i in range(1, totalRowsNum):

    logging.info("---------------------------")
    logging.info("Row " + str(i) + " : ")
    row = sheet.row_values(i)

    logging.info("Shop name : " + row[COL_SHOP_NAME])
    logging.info("crooz_product_code : " + row[COL_CROOZ_PRODUCT_NAME])
    logging.info("Image url : " + row[COL_IMG_URL])
    tmp_data = row[COL_IMG_URL].split("/")
    tmp_col = len(tmp_data)

    # example file_name = yamatwo_YMTW0000008_1.jpg
    file_name = DIRECTORY_PATH + row[COL_SHOP_NAME] + "_" + row[COL_CROOZ_PRODUCT_NAME] + "_" + tmp_data[tmp_col - 1]
    logging.info("File name : " + file_name)
    download_img(row[COL_IMG_URL], file_name)
    logging.info("---------------------------")
