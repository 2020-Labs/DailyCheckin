'''
description:   统计未提交打卡名单
version: 0.1
'''

# README
'''
## 使用说明
- 第一步：准备工作
  - 收集原始数据并导出为csv格式
  - 确认姓名列的索引值
- 第二步：执行脚本

  ```python
  daliy_check.py -f xxx.csv  -i 7
  ```
- 命令说明：  
  ```
  Usage: [-f|-i] [--help|--file=|--index=]
      -f|--file=
         csv文件路径，如:d:\2020.csv
      -i|--index=
         姓名一栏的索引值，从0算起
  ```
'''
import sys
import getopt
import csv
import os

# 检查的提交的名单
MEMBERS = [
    '尚利兵', '王磊', '郭艳雯', '马魁', '王小超', '史发龙', '宋智峰', '刘炳星', '柯艳婷', '王柯权',
    '张超', '张建', '宋洁玉', '杨才林', '闫少波',
    '张胜利', '段恒飞', '何楠', '段东峰', '刘志领', '林洋', '王蕾', '满建超', '屈成栋', '王超迪', '陈民华',
    '田佳佳', '王阿雷', '李帅', '王伟', '成二磊', '麻洋',
    '赵瑾', '常焕利', '苏磊', '何娟娟',
    '尹超', '付自成', '张永刚', '彭飞', '惠文博', '李增辉', '吴佳琪', '潘建勇', '王帅'
]

# 表示只统计Network组的人员
NETWORK = False

if NETWORK:
    MEMBERS = ['尚利兵', '王磊', '郭艳雯', '马魁', '王小超', '史发龙', '宋智峰', '刘炳星', '柯艳婷', '王柯权']
# 数据来源
XLS_FILE = None

# 表格里姓名栏的索引值，第1列是0
NAME_COL_INDEX = 0

KEY_VALUE_MAPS = {
    '3': {
        '9': '西安'
    },
    '4': {
        '7': '西安黄区'
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
        '-2': '未填',
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
    with open(XLS_FILE, 'r') as csvfile:
        reader = csv.reader(csvfile)
        summited_list = []

        for row in reader:
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
    value = '';
    value = KEY_VALUE_MAPS[key1].get(key2)
    return value


def convert_text():
    output_rows = []
    with open(XLS_FILE, 'r') as csvfile:
        reader = csv.reader(csvfile)

        for row in reader:
            # print(row)

            for k in KEY_VALUE_MAPS:
                display_text = get_value(k, row[int(k) + 5])
                if display_text:
                    row[int(k) + 5] = display_text
            output_rows.append(row[6:])


    #output to csv file
    output_file = XLS_FILE.replace('.csv', '_output.csv')
    with open(output_file, 'w', newline="") as file:
        writer = csv.writer(file)
        for r in output_rows:
            print(r)
            writer.writerow(r)


def usage():
    print('Usage: [-f|-i] [--help|--file=|--index=]');
    print('  -f|--file=')
    print('     csv文件路径，如:d:\\2020.csv')
    print('  -i|--index=')
    print('     姓名一栏的索引值，从0算起')


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
        print('NAME_COL_INDEX =', NAME_COL_INDEX)
        print('-' * 50)
        # do_check()
        convert_text()
