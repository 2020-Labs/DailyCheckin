'''
description:   统计未提交打卡名单
version: 0.4
'''

# README
import json
import time

'''
## 使用说明
- 第一步：准备工作
  - 收集原始数据并导出为csv格式
  - 确认姓名列的索引值
- 第二步：执行脚本

  ```python
  daliy_check.py -f xxx.xls  -i 7
  ```
- 命令说明：  
  ```
  Usage: [-f] [--help|--file=]
      -f|--file=
         xls文件路径，如:d:\2020.xls
  ```
'''
import sys
import getopt
import csv
import os
import pandas as pd

# 检查的提交的名单
MEMBERS = [
    '尚利兵', '王磊', '郭艳雯', '马魁', '王小超', '史发龙', '宋智峰', '刘炳星', '柯艳婷', '王柯权',
    '张超', '张建', '宋洁玉', '杨才林', '闫少波',
    '张胜利', '段恒飞', '何楠', '段东峰', '刘志领', '林洋', '王蕾', '满建超', '屈成栋', '王超迪', '陈民华',
    '田佳佳', '王阿雷', '李帅', '王伟', '成二磊', '麻洋',
    '赵瑾', '常焕利', '苏磊', '何娟娟',
    '尹超', '付自成', '张永刚', '惠文博', '李增辉', '吴佳琪', '潘建勇', '王帅'
]

# 表示只统计Network组的人员
NETWORK = False

if NETWORK:
    MEMBERS = ['尚利兵', '王磊', '郭艳雯', '马魁', '王小超', '史发龙', '宋智峰', '刘炳星', '柯艳婷', '王柯权']
# 数据来源
XLS_FILE = None

# 表格里姓名栏的索引值，第1列是0
NAME_COL_INDEX = 7

KEY_VALUE_MAPS = {
    '3': {
        '9': '西安',
        '*': '西安'
    },
    '4': {
        '7': '西安黄区',
        '*': '西安黄区'

    },
    '5': {
        '*': '15'
    },
    '6': {
        '1': '离开过',
        '2': '没有离开'
    },
    '7': {
        '1': '具备远程办公条件可远程办公',
        '2': '无网络、电脑不具备远程办公条件',
        '3': '工作岗位性质无法远程办公',
        '4': '身体不适无法远程办公',
    },
    '8': {
        '1': '正常',
        '2': '感冒咳嗽',
        '3': '隔离中',
        '4': '已确诊',
    },
    '9': {
        '1': '正常',
        '2': '感冒咳嗽',
        '3': '隔离中',
        '4': '已确诊',
        '5': '本人独居',
    },
    '10': {
        '1': '没有',
        '2': '有，但戴了口罩',
        '3': '有，但没有戴口罩',
    },
    '12': {
        '1': '是',
        '2': '否',
        '-2': '未填',
    },
    '13': {
        '-3': '(跳过)'
    },
    '14': {
        '-3': '(跳过)'
    },
    '15': {
        '1': '是',
        '2': '否'
    },
    '16': {
        '1': '自驾',
        '2': '高铁/火车',
        '3': '大巴',
        '4': '飞机',
        '5': '拼车（顺风车）',
        '-2': '(空)',
    },
    '18': {
        '1': '自驾',
        '2': '拼同事顺风车',
        '3': '公司班车',
        '4': '步行',
        '5': '公共交通',
        '6': '滴滴/打车'
    },
    '20': {
        '1': '是',
        '2': '否'
    },
    '21': {
        '1': '是',
        '2': '否'
    },
    '22': {
        '1': '远程办公',
        '2': '现场办公',
        '-3': '(跳过)'
    }

}


def do_check():
    # 收集表格里的人员名单
    cols, rows = read_excel()

    summited_list = []
    for row in rows:
        summited_list.append(row[NAME_COL_INDEX].strip())

    # 检查统计未提交的名单
    unsummit_list = []
    for m in MEMBERS:
        if not m in summited_list:
            unsummit_list.append(m)

    # 输出结果
    if len(unsummit_list) > 0:
        print('=' * 50)
        print('## 输出结果：')
        print(' > 未提交名单:')
        print('  ', unsummit_list)
    else:
        print('#' * 50)
        print('恭喜！全部提交！！！！')
        print('#' * 50)


def get_value(key1, key2):
    value = ''
    value = KEY_VALUE_MAPS[key1].get(str(key2))
    if not value:
        value = KEY_VALUE_MAPS[key1].get('*')
    return value


def convert_text():
    output_rows = []
    headline = False
    head_row = []

    with open(XLS_FILE, 'r') as csvfile:
        reader = csv.reader(csvfile)

        for row in reader:
            # print(row)
            if not headline:
                head_row = row[6:]
                headline = True
                continue

            for k in KEY_VALUE_MAPS:
                display_text = get_value(k, row[int(k) + 5])
                if display_text:
                    row[int(k) + 5] = display_text

            new_row = row[6:]
            output_rows.append(row[6:])

    print('检查填写有误的数据：')
    error_rows1=[]
    error_rows2=[]

    for new_row in output_rows:
        if new_row[5] == '没有离开' and new_row[10].find('西安市') < 0:
            #print(new_row)
            error_rows1.append(new_row)

        if new_row[11] == '否' and new_row[18].strip() == '(空)':
            #print(new_row)
            error_rows2.append(new_row)


    QUESTION_ID_IDX = 0
    QUESTION_NAME_IDX =1
    QUESTION_6_IDX = 5
    QUESTION_10_IDX = 19

    if len(error_rows1) > 0:
        print('问卷中问题6、问题11填写错误的名单：')
        for r in error_rows1:
            print('  {0}      {1}'.format(r[QUESTION_ID_IDX], r[QUESTION_NAME_IDX]))
            #print('   Q: {0}'.format(head_row[QUESTION_6_IDX]))
            #print('   A: {0}'.format(r[QUESTION_6_IDX]))
            #print('   Q: {0}'.format(head_row[QUESTION_10_IDX]))
            #print('   A: {0}'.format(r[QUESTION_10_IDX]))

    print('')
    if len(error_rows2) > 0:
        print('问卷中问题12、问题19填写错误的名单：')
        for r in error_rows2:
            print('  {0}      {1}'.format(r[QUESTION_ID_IDX], r[QUESTION_NAME_IDX]))
            #print('   Q: {0}'.format(head_row[11]))
            #print('   A: {0}'.format(r[11]))
            #print('   Q: {0}'.format(head_row[18].replace('\n','')))
            #print('   A: {0}'.format(r[18]))
    print('检查完毕')

    #output to csv file
    output_file = XLS_FILE.replace('.csv', '_output.csv')
    with open(output_file, 'w', newline="") as file:
        writer = csv.writer(file)
        writer.writerow(head_row)
        for r in output_rows:
            writer.writerow(r)


def read_excel():
    df = pd.read_excel(XLS_FILE);
    columns = []
    for col in df.columns:
        columns.append(col.replace('\n',''))

    values = []
    for v in df.values:
        values.append(v)

    return columns,values


def output_result(cols,rows):
    print('输出Excel')
    print('列头共有{0}列'.format(len(cols)))
    print('数据行共有{0}列'.format(len(rows[0])))

    data = {}
    for idx in range(len(cols)):
        col = cols[idx]
        data[col] = []
        for val in rows:
            data[col].append(val[idx])

    df = pd.DataFrame(data, index=range(1, len(rows) + 1))

    output_file = os.path.join(os.path.dirname(XLS_FILE), 'OPPO合作伙伴每日健康打卡4.0v2_西安地区_{0}.xlsx'.format(time.strftime('%Y%m%d', time.localtime())))
    writer = pd.ExcelWriter(output_file,  engine='xlsxwriter')

    df.to_excel(writer, index=False)

    #设置Excel表格格式
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    header_format = workbook.add_format({
        'valign': 'top',  # 垂直对齐方式
        'align': 'left',  # 水平对齐方式
        'fg_color': '#B4C6E7',  # 单元格背景颜色
        'border': 1,  # 单元格边框宽度
        'font_size': 10,
        'font_name': '微软雅黑'
        })

    cell_format = workbook.add_format({
        'align': 'left',
        'border': 1,
        'font_size': 10,
        'font_name': '微软雅黑'
    })

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    for index, value in df.iterrows():
        worksheet.set_row(index, 16)

    worksheet.set_column(0, len(df.columns.values), cell_format=cell_format)

    #设定第1行的高度
    worksheet.set_row(0, 20)

    #设定各列的宽度
    #设置默认列宽
    worksheet.set_column("B:V", 14)
    worksheet.set_column("A:A", 13)
    worksheet.set_column("F:F", 20)
    worksheet.set_column("G:G", 23)
    worksheet.set_column("K:K", 23)
    worksheet.set_column("M:N", 20)
    worksheet.set_column("P:R", 20)
    worksheet.set_column("S:S", 40)
    worksheet.set_column("T:T", 16)
    worksheet.set_column("U:V", 14)

    #固定1行1列
    worksheet.freeze_panes(1, 1)
    try:
        writer.save()
    except Exception as e:
        print('保存失败，错误信息：', e)
        print('请确认文件是否已打开。')
        return

    print('成功写入文件: {0}'.format(output_file))
    print('共 {0} 条数据。'.format(len(rows)))


def convrt_text_v2():
    cols, rows = read_excel()

    #去重 , 多次提交的数据以最后一次为准
    new_rows = {}
    for row in rows:
        key = row[7]
        new_rows[key] = row

    output_rows = []
    for row in new_rows.values():
        for k in KEY_VALUE_MAPS:
            display_text = get_value(k, row[int(k) + 5])
            if display_text:
                row[int(k) + 5] = display_text

        new_row = row[6:]
        output_rows.append(row[6:])

    check_rows(output_rows)

    output_result(cols[6:], output_rows)


def check_rows(output_rows):
    QUESTION_ID_IDX = 0
    QUESTION_NAME_IDX =1
    QUESTION_6_IDX = 5
    QUESTION_11_IDX = 10
    QUESTION_12_IDX = 11
    QUESTION_19_IDX = 19
    print('检查填写有误的数据：')
    error_rows1=[]
    error_rows2=[]
    error_rows3=[]
    for new_row in output_rows:
        if new_row[5] == '没有离开' and new_row[10].find('西安市') < 0:
            #print(new_row)
            error_rows1.append(new_row)

        if new_row[11] == '否' and new_row[18].strip() == '(空)':
            #print(new_row)
            error_rows2.append(new_row)
        if not len(new_row[0]) == 8:
            error_rows3.append(new_row)

    if len(error_rows3) > 0:
        print('问卷中问题1写错误的名单：（工号填写有误）')
        for r in error_rows3:
            print('  "{0}"      {1}'.format(r[QUESTION_ID_IDX], r[QUESTION_NAME_IDX]))
    print('')
    if len(error_rows1) > 0:
        print('问卷中问题6、问题11填写错误的名单：')
        for r in error_rows1:
            print('  {0}      {1}  6_是否离开西安: {2}   11_当前所在区域: {3}'.format(
                r[QUESTION_ID_IDX],
                r[QUESTION_NAME_IDX],
                r[QUESTION_6_IDX],
                r[QUESTION_11_IDX]
            ))

    print('')
    if len(error_rows2) > 0:
        print('问卷中问题12、问题19填写错误的名单：')
        for r in error_rows2:
            print('  {0}      {1} 12_是否确定返程日期: {2} ， 19_备注: {2}'.format(
                r[QUESTION_ID_IDX],
                r[QUESTION_NAME_IDX],
                r[QUESTION_12_IDX],
                r[QUESTION_19_IDX]))
            #print('   Q: {0}'.format(head_row[11]))
            #print('   A: {0}'.format(r[11]))
            #print('   Q: {0}'.format(head_row[18].replace('\n','')))
            #print('   A: {0}'.format(r[18]))
    print('检查完毕')
    if len(error_rows1) == 0 and len(error_rows2) == 0:
        print('提交的数据全部正确')


def usage():
    print('Usage: [-f] [--help|--file=]');
    print('  -f|--file=')
    print('     Excel文件路径，如:d:\\2020.xls')


def check_arg():
    global XLS_FILE, NAME_COL_INDEX
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hf:i:", ["help", "file=", "index="]);

        # check all param
        for opt, arg in opts:
            if opt in ("-h", "--help"):
                usage();
                sys.exit(1);
            elif opt in ("-f", "--file"):
                XLS_FILE = arg
            elif opt in ("-i", "--index"):
                NAME_COL_INDEX = arg
            else:
                print("%s ==> %s" % (opt, arg));

    except getopt.GetoptError:
        print("getopt error!");
        usage();
        return False

    if not XLS_FILE:
        usage();
        print('-_-' * 20)
        print('错误：-f|--file, 请指定CSV文件路径！！')
        return False

    if not os.path.exists(XLS_FILE):
        print('文件:{0}不存在！！'.format(XLS_FILE))
        return False

    try:
        NAME_COL_INDEX = int(NAME_COL_INDEX)
    except ValueError:
        print('-_-' * 20)
        print('错误：-i|--index,参数不是整数')
        return False

    if NAME_COL_INDEX < 0 or NAME_COL_INDEX > 65535:
        print('错误：-i|--index,参数无数值，必须是0~65535之间有整数')
        return False

    return True


if __name__ == '__main__':
    if check_arg():
        print('-' * 50)
        print('XLS_FILE= ', XLS_FILE)
        print('-' * 50)
        do_check()
        convrt_text_v2()
