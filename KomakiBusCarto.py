# -*- coding: utf-8 -*-

import xlrd

def make_dic():

    book = xlrd.open_workbook('./books/bus_idokeido.xls')
    sheet = book.sheet_by_index(0)
    dict = {}

    for row in range(1, sheet.nrows):
        dict[sheet.cell(row, 1).value] = (float(sheet.cell(row, 2).value), float(sheet.cell(row, 3).value))

    return dict

def make_patrol():

    book = xlrd.open_workbook('.//books/jyunkai-bus-zikokuhyou-20160401.xls')

    dict = {}
    for sheet in book.sheets():

        course_name = sheet.cell(0, 0).value
        row_index = 2

        list = []
        # なぜ２行目から始めたのか。。。
        if (course_name == '') :
            course_name = sheet.cell(1, 0).value
            row_index = 3

        for row in range(row_index, sheet.nrows):
            print(sheet.cell(row_index, 0).value)

            list.append([sheet.cell(row_index, 0).value, sheet.cell(row_index, 1)])



        #print(course_name)
        #print('%s %sx%s' % (sheet.name, sheet.nrows, sheet.ncols))

if __name__ == "__main__":

    # dict = make_dic()
    # for key in dict:
    #     print(key, dict[key])
    make_patrol()



