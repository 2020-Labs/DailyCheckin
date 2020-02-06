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

#表示只统计Network组的人员
NETWORK = False

if NETWORK:
    MEMBERS = ['尚利兵', '王磊', '郭艳雯', '马魁', '王小超', '史发龙', '宋智峰', '刘炳星', '柯艳婷', '王柯权']
# 数据来源
XLS_FILE = None

# 表格里姓名栏的索引值，第1列是0
NAME_COL_INDEX = 0


def do_check():
    #收集表格里的人员名单
    with open(XLS_FILE, 'r') as csvfile:
        reader = csv.reader(csvfile)
        summited_list = []
        for row in reader:
            summited_list.append(row[NAME_COL_INDEX].strip())

    #检查统计未提交的名单
    unsummit_list = []
    for m in MEMBERS:
        if not m in summited_list:
            unsummit_list.append(m)

    #输出结果
    if len(unsummit_list) > 0:
        print('=' * 50)
        print('未提交名单:')
        print(unsummit_list)
    else:
        print('#' * 50)
        print('恭喜！全部提交！！！！')
        print('#' * 50)


def usage():
    print('Usage: [-f|-i] [--help|--file=|--index=]');
    print('  -f|--file=')
    print('     csv文件路径，如:d:\\2020.csv')
    print('  -i|--index=')
    print('     姓名一栏的索引值，从0算起')

def check_arg():
    global  XLS_FILE,NAME_COL_INDEX
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
        do_check()
