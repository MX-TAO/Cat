import json
import xlrd
import xlwt
import argparse
import sys
import os

parser = argparse.ArgumentParser()
parser.add_argument('-n', '--name', type=str, default='')
parser.add_argument('-v', '--version', type=str, default='')
parser.add_argument('-d', '--date', type=str, default='')

args = parser.parse_args()

if len(args.name) < 1:
	if len(args.version) < 1 or len(args.date) < 1:
		path = '2019猫咪档案_20190825_upd9.1.xlsx'
	else:
		path = '%s猫咪档案_%s_upd%s.xlsx' % (args.date[:4], args.date, args.version)
else:
	path = args.name

workbook = xlrd.open_workbook(path)

data = workbook.sheets()[0]

rowNum = data.nrows  # sheet行数
colNum = data.ncols  # sheet列数

# 获取所有单元格的内容
data_list = []
for i in range(12, rowNum):
	rowlist = []
	for j in range(colNum):
		if data.cell(i, j).ctype == 3:
			dt = xlrd.xldate.xldate_as_tuple(data.cell_value(i, j), 0)
			rowlist.append('%04d-%02d-%02d' % dt[0:3])
		else:
			rowlist.append(data.cell_value(i, j))
	data_list.append(rowlist)
# 输出所有单元格的内容

rowNum -= 12

"""
for i in range(rowNum):
	for j in range(colNum):
		print(data_list[i][j], '\t', end="")
	print()
"""

labels = [
	[1, 'id', lambda x:int(x)],
	[2, 'name', lambda x:'【还没有名字】' if len(x) < 1 else x],
	[3, 'nickname', lambda x:x],
	[4, 'fur_color', lambda x:x],
	[5, 'site', lambda x:x],
	[6, 'gender', lambda x:'♂' if x == 1 else '♀' if x == 0 else '未知'],
	[7, 'state', lambda x:'不明' if len(x) < 1 else x],
	[8, 'is_sterilized', lambda x:'已绝育' if x == 1 else '未绝育' if x == 0 else '未知/可能不适宜绝育'],
	[9, 'date_of_sterilized', lambda x:x],
	[10, 'birth', lambda x:x],
	[11, 'appearance', lambda x:x],
	[12, 'courage', lambda x:int(x) if type(x)==int else int(-1)],
	[13, 'first_seen', lambda x:x],
]

data_json = []

for i in range(rowNum):
	json_line = {}

	if len(data_list[i][4]) < 1: # 毛色都不知道那就是没有【x
		continue

	for j in labels:
		json_line[j[1]] = j[2](data_list[i][j[0]])

	data_json.append(json_line)
	print(json_line)

with open('cat_info.jsonl', 'w', encoding='utf-8') as f:
	for line in data_json:
		f.write(json.dumps(line))
		f.write('\n')

with open('cat_info.json', 'w', encoding='utf-8') as f:
	f.write(json.dumps(data_json, indent=2))