# Create By: Aaron
# Date : 07/15/2020
import xlwt
import pymysql
class MYSQL:
  def __init__(self):
    pass
  def __del__(self):
    self._cursor.close()
    self._connect.close()
  def connectDB(self):
    """
    连接数据库
    connect Database
    :return:
    """
    try:
      self._connect = pymysql.Connect(
        host='127.0.0.1',
        # 請在此處放入你的主機IP或者域名， 此處為本地地址
        # Please put your host IP or domain name here， here is localhost Address
        port=3306,
        # 數據庫默認端口
        # Database default port
        user='root',
        # 默認root用戶
        # Default root user
        passwd='123456',
        # 請修改為你的數據庫密碼， 不是服務器的密碼
        # Please change to your database password, not the server password
        db='table_name',
        # db變量存放你需要導出成Excel的數據庫名
        # The db variable stores the database name you need to export to Excel
        charset='utf8'
        # 字符編碼默認為"utf8"
        # The character encoding defaults to "utf8"
      )
      return 0
    except:
      return -1
  def export(self, table_name, output_path):
    self._cursor = self._connect.cursor()

    count = self._cursor.execute('select * from ' + table_name )
    print(count)

    # 重置游標的位置
    # Reset cursor position
    self._cursor.scroll(0, mode='absolute')

    # 搜取所有的結果
    # Search all results
    results = self._cursor.fetchall()

    # 獲取MySQL裡面的字段名稱
    # Get the name of the data field in MYSQL
    fields = self._cursor.description
    workbook = xlwt.Workbook()

    # 注意: 在add_sheet時, 置參數cell_overwrite_ok=True, 可以覆蓋原單元格中的數據。
    # Note: When add_sheet, set the parameter cell_overwrite_ok=True to overwrite the data in the original cell.

    # cell_overwrite_ok默认为False, 覆蓋的話, 會拋出異常.
    # cell_overwrite_ok defaults to False, if overwritten, an exception will be thrown.

    sheet = workbook.add_sheet('table_'+table_name, cell_overwrite_ok=True)

    # 寫上字段信息
    # Write field information
    for field in range(0, len(fields)):
      sheet.write(0, field, fields[field][0])

    # 獲取並寫入信息
    # Get and write data segment information
    row = 1
    col = 0
    for row in range(1,len(results)+1):
      for col in range(0, len(fields)):
        sheet.write(row, col, u'%s' % results[row-1][col])
    workbook.save(output_path)
if __name__ == '__main__':
  mysql = MYSQL()
  flag = mysql.connectDB()
  if flag == -1:
    print('數據庫連接失敗--Database Connect Faild')
  else:
    print('數據庫連接成功--Database Connect Successful')
    mysql.export("`table_name`", '/Users/Aaron/Desktop/xxx.xls')
    # 此處的第一參數請你要導出的數據庫表名， 後面的參數放置導出數據庫表的位置， 此處用的是macOS， Windows， Linux請修改其他的文件路徑。
    # The first parameter here is the name of the database table you want to export. The following parameters place the location of the exported database table. Here, macOS, Windows, and Linux are used. Please modify other file paths.
