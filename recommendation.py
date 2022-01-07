from os import read
from openpyxl import Workbook, load_workbook
import random


# 一个过分简单的测试函数
def test(workBook: Workbook)->bool:
    sheet = workBook.active
    a1 = sheet['A1']
    return bool(a1)

def fetchData(storeData: list, workBook: Workbook)->None:
    """获取数据，并存储在函数外的列表中

    Args:
        storeData (list): 用于存储数据的列表
        workBook (Workbook): 用于获取数据的WorkBook
    """

    sheet = workBook.active
    attrList = []
    for cell in sheet[1]:
        attrList.append(cell.value)
    # print(attrList)
    for row in sheet.iter_rows(min_row=2):
        newBook = {}
        for index, cell in enumerate(row):
            newBook[attrList[index]] = cell.value
        storeData.append(newBook)


def getInfo(book: dict) -> str:
    """获取书本的介绍性信息

    Args:
        book (dict): 以字典形式存储的书本

    Returns:
        str: 返回格式化的介绍信息
    """
    # 不予显示的属性
    notDisplayAttrNames = ['筛选标记']

    result = ''
    for key,value in book.items():
        if str(key) in notDisplayAttrNames:
            continue
        else:
            newLine = str(key)+' '+str(value)+'\n'
            result+=newLine
    return result



def randomMode(data: list) -> None:
    """全随机模式：从整个图书列表中随机抽取一本书并打印到控制台。

    Args:
        data (list): 图书列表
    """
    print('='*20, '随机模式', '='*20)
    choosenBook = random.choice(data)
    result = getInfo(choosenBook)
    print(result)


def main() -> None:
    OPENFILENAME = 'book-list.xlsx'
    readBook = load_workbook(OPENFILENAME)
    bookList = []
    if not test(readBook):
        # 过分简单的异常处理
        print('运行错误')
        return

    fetchData(bookList, readBook)
    randomMode(bookList)


if __name__ == '__main__':
    main()
