# Excel中数据导入Oracle
# JOCY
# to connect oracle
import cx_Oracle
# to read excel files
import xlrd
import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
# coding: utf-8


# 打开excel
excel = xlrd.open_workbook('C:\\Users\\jixiang\\Desktop\\aaa.xlsx')
# 通过索引顺序获取
table = excel.sheet_by_index(0)

# 获取整行的值（数组）
data = table.row_values(0)
# get values by cols in array
data2 = table.col_values(0)
# 获取行数
rows = table.nrows
# 获取列数
cols = table.ncols

# print(table.cell(3, 3).value)

# 建立数据库连接
conn = cx_Oracle.connect('examine/root@localhost/orcl')

c = conn.cursor()

# for i in range(cols):
# print(table.col_values(0))
# 中文，字符插入报错，从oracle直接插入可以插入字符及中文，从Python插入不对
count = 3000
# line = line.strip().split(',')
# 使用单引号，可传字符串 转义
for i in range(rows):
    x = c.execute('insert into DM_USERINFO values(\'%s\',\'%s\',3,4,5)' % (count, data2[i]))
    count += 1
# x = c.execute('insert into DM_USERINFO (ID, NAME, DEPART, USERTYPE, CREATOR) values(%s,%s,%s,%s,%s)'
#              % (table.cell(4, 2).value, "直接输入",
#                 table.cell(3, 2).value, table.cell(3, 3).value,
#                 table.cell(3, 1).value))
conn.commit()
c.close()
conn.close()

