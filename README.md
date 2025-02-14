# excel_solo简易版

随手写的一个python处理excel文件的脚本，放在这里下次用的时候二开；使用条件比较苛刻，因为是给特定排列的excel处理数据

# 代码

tuxinghua.py中为实现图形化界面
othersoloexcel.py中实现主要功能

# 功能

注释里写的很清楚，简单来说你输入要搜索的文字（a），则整个表格中搜索a，然后从a的下一行开始一行一行的遍历，将遍历的先存到汇总.xlsx中，然后再写一个方法来遍历这个汇总.xlsx，使用groupby根据第一列对数据进行分组，然后对第2、3、4、5列进行求和然后输出到汇总.xlsx中

# 运行

python3 tuxinghua.py
