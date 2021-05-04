import sys
import os
sys.path.append(os.path.dirname(os.path.realpath(__file__)))
import pdfplumber
import pandas as pd


import openpyxl

from os import listdir
from os.path import isfile, isdir, join


def main(mypath):
    try:
        excelFile = 'Exception'
        # 指定要列出所有檔案的目錄
        # mypath = "D:/temp/test"

        # 取得所有檔案與子目錄名稱
        files = listdir(mypath)

        # 以迴圈處理
        for f in files:
            # 產生檔案的絕對路徑
            fullpath = join(mypath, f)
            # 判斷 fullpath 是檔案還是目錄
            if isfile(fullpath):
                if f.lower().endswith('.pdf'):
                    # print("檔案：", fullpath)

                    excelFile = fullpath.replace('.pdf', '.xlsx').replace('.PDF', '.xlsx')
                    if os.path.exists(excelFile):
                        os.remove(excelFile)
                        # os.unlink(excelFile)

                    with pdfplumber.open(fullpath) as pdf:
                        writer = pd.ExcelWriter(excelFile, engine='openpyxl')
                        if os.path.exists(excelFile):
                            book = openpyxl.load_workbook(excelFile)
                            writer.book = book

                        for page in range(len(pdf.pages)):
                            # print(page)
                            pdf_page = pdf.pages[page]

                            # 解析表格
                            tables = pdf_page.extract_tables()
                            # print(len(tables))
                            count = 0
                            for table in tables:
                                # print(table)
                                df = pd.DataFrame(table[1:], columns=table[0])

                                df.to_excel(writer, sheet_name='P' + str(page) + '_T' + str(count))
                                writer.save()
                                writer.close()
                                count = count + 1
                    
        excelFile = 'Done'
                    
    finally:
        return excelFile    

if __name__ == '__main__':
    main(sys.argv[1])   # 或是任何你想執行的函式
