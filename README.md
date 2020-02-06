## 使用说明
- 第一步：准备工作
  - 收集原始数据并导出为csv格式
  - 确认姓名列的索引值
- 第二步：执行脚本
  
    ```python
    daily_checkin.py -f xxx.csv  -i 7
    or
    daily_checkin.py --file=xxx.csv  --index=7
    ```
- 命令说明：  
  ```shell
  Usage: [-f|-i] [--help|--file=|--index=]
      -f|--file=
         csv文件路径，如:d:\2020.csv
      -i|--index=
         姓名一栏的索引值，从0算起
  ```
