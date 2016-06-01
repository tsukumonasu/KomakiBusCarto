# -*- coding: utf-8 -*-

import xlrd
import xlwt
import datetime



def read_geocoding_dict():

    book = xlrd.open_workbook('./books/bus_idokeido.xls')
    sheet = book.sheet_by_index(0)
    dict = {}

    for row in range(1, sheet.nrows):
        dict[sheet.cell(row, 1).value] = (float(sheet.cell(row, 2).value), float(sheet.cell(row, 3).value))

    return dict



def read_patrol_dict():

    book = xlrd.open_workbook('.//books/jyunkai-bus-zikokuhyou-20160401.xls')

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
                    list.append((sheet.cell(row_index, 0) ,sheet.cell(row_index, col_index), course_name))

        dict[course_name] = list
    return dict



def make_geopat_list(geocoding_dict, patrol_dict):

    list = []

    for root_name, time_list in patrol_dict.items():
        for time in time_list:
            # たまにないやつがあるけどまぁいっか。。。
            if geocoding_dict.__contains__(time[0].value):
                list.append((geocoding_dict[time[0].value] + time))
                #print(root_name, geocoding_dict[time[0].value])
    return list

if __name__ == "__main__":
    geocoding_dict = read_geocoding_dict()
    patrol_dict = read_patrol_dict()
    geopat_list = make_geopat_list(geocoding_dict, patrol_dict)

    book = xlwt.Workbook()
    sheet = book.add_sheet("NewSheet_1")
    xf = xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm')

    sheet.write(0, 0, "root")
    sheet.write(0, 1, "stop")
    sheet.write(0, 2, "time")
    sheet.write(0, 3, "longitude")
    sheet.write(0, 4, "latitude")

    row_index = 1
    for geopat in geopat_list:
        sheet.write(row_index, 0, geopat[4])
        sheet.write(row_index, 1, geopat[2].value)
        sheet.write(row_index, 3, geopat[0])
        sheet.write(row_index, 4, geopat[1])

        # 日付
        py_date = xlrd.xldate_as_tuple(geopat[3].value, 0)
        sheet.write(row_index, 2, datetime.datetime(2016, 4, 1, py_date[3], py_date[4]), xf)
        row_index += 1

    book.save('komaki_bus_time.xls')




