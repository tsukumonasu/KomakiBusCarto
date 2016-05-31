# -*- coding: utf-8 -*-

import xlrd
import re

def read_geocoding_dict():

    book = xlrd.open_workbook('./books/bus_idokeido.xls')
    sheet = book.sheet_by_index(0)
    dict = {}

    for row in range(1, sheet.nrows):
        dict[sheet.cell(row, 1).value] = (float(sheet.cell(row, 2).value), float(sheet.cell(row, 3).value))

    return dict

def read_patrol_dict():

    book = xlrd.open_workbook('.//books/jyunkai-bus-zikokuhyou-20160401.xls')
    pattern = re.compile("^[0-9]")

    dict = {}
    for sheet in book.sheets():

        course_name = sheet.cell(0, 0).value
        row_first = 2

        list = []
        # なぜ２行目から始めたのか。。。
        if (course_name == '') :
            course_name = sheet.cell(1, 0).value
            row_first = 3

        for row_index in range(row_first, sheet.nrows):
            for col_index in range(1, sheet.ncols):

                cell = sheet.cell(row_index, col_index)

                # float型なら追加
                if (isinstance(cell.value, float)):
                    list.append((sheet.cell(row_index, 0) ,sheet.cell(row_index, col_index)))

        dict[course_name] = list


def make_geopat_list(geocoding_dict, patrol_dict):
    return

if __name__ == "__main__":
    geocoding_dict = read_geocoding_dict()
    patrol_dict = read_patrol_dict()

    geopat_list = make_geopat_list(geocoding_dict, patrol_dict)



