# -*- coding: utf-8 -*-

# 导入MySQL驱动:
import mysql.connector
import base64
import xlwt
from datetime import datetime

# 从数据库查询数据
def get_data(sql):
    # 注意把password设为你的root口令:
    conn = mysql.connector.connect(user='root', password='root', database='xforeign1')
    # 运行查询:
    cursor = conn.cursor()
    cursor.execute(sql)
    values = cursor.fetchall()
    # 关闭Cursor和Connection:
    cursor.close()
    conn.close()
    return values

#写入到excel
def write_data_to_excel(name,sql):
    # 将sql作为参数传递调用get_data并将结果赋值给result,(result为一个嵌套元组)
    result = get_data(sql)
    # 实例化一个Workbook()对象(即excel文件)
    wbk = xlwt.Workbook()
    # 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
    sheet = wbk.add_sheet('Sheet1',cell_overwrite_ok=True)
    # 遍历result中的没个元素。
    for i in xrange(len(result)):
        #对result的每个子元素作遍历，
        for j in xrange(len(result[i])):
            # 解密
            if result[i][j]:
                write_value = base64.b64decode(result[i][j]).decode('gbk')
                #将每一行的每个元素按行号i,列号j,写入到excel中。
                sheet.write(i,j,write_value)
    
    # 获取当前日期，得到一个datetime对象如：(2016, 8, 9, 23, 12, 23, 424000)
    today = datetime.today()
    # 将获取到的datetime对象仅取日期如：2016-8-9
    today_date = datetime.date(today)
    # 以传递的name+当前日期作为excel名称保存。
    wbk.save(name+str(today_date)+'.xls')


# 如果该文件不是被import,则执行下面代码。
if __name__ == '__main__':
    #定义一个字典，key为对应的数据类型也用作excel命名，value为查询语句
    db_dict = {'castResult':'select title,content from base_user'}
    # 遍历字典每个元素的key和value。
    for k,v in db_dict.items():
        # 用字典的每个key和value调用write_data_to_excel函数。
        write_data_to_excel(k,v)
