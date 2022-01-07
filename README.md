# Book Recommendation Program
`book-list.xlsx`：图书列表文件。

​	若筛选标记为 0 表示不参与抽奖。

`main.py`：图书抽奖程序

​	程序运行命令：`python main.py`

​	快速从筛选标记非0的图书中抽取一本书可使用命令：` python main.py 0`

​	抽取一本技术书：`python main.py 1`

​	抽取一本非技术书：`python main.py 2`

## 运行
程序依赖以下第三方模块：
* openpyxl
* pyperclip

可以运行如下语句安装：

* `pip install openpyxl`
* `pip install pyperclip`

