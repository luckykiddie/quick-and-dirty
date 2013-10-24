#-*- coding: utf-8 -*-

import xlrd
import xlwt
import re

# 检查是否满足报考条件
def check(row_value):
    zy = row_value[11]
    if not checkZY(zy):
        return False

    xw = row_value[13]
    if not checkXW(xw):
        return False

    if checkSpecial(row_value):
        return False

    return True

# 检查是否满足专业要求
def checkZY(value):
    pat = re.compile(u'不限|限制|生物工程|化工')
    if re.search(pat, value):
        return True

    return False

# 检查是否满足学位要求
def checkXW(value):
    pat = re.compile(u'学士|不|无')
    if re.search(pat, value):
        return True

    return False

# 减产是否需要满足特殊要求
def checkSpecial(row_value):
    pat = re.compile(u'是')
    for i in range(16, 19):
        value = row_value[i]
        if re.search(pat, value):
            return True

    return False

# 根据条件筛选出职位
def filterTitle():
    data = xlrd.open_workbook('gjgwy.xls')
    output = xlwt.Workbook(encoding='utf-8')

    for sheet in data.sheets():
        output_sheet = output.add_sheet(sheet.name)
        output_row = 1
        for row in range(sheet.nrows):
            row_value = sheet.row_values(row)
            if len(row_value) < 11:
                continue

            choosed = True
            if row != 2 and not check(row_value):
                choosed = False

            if choosed == True:
                for col in range(sheet.ncols):
                    output_sheet.row(output_row).write(col, sheet.cell(row,col).value)

                output_sheet.flush_row_data()
                output_row += 1

    output.save('output.xls')

if __name__ == '__main__':
    filterTitle()
