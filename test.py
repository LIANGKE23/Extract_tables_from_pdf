import pdfplumber
import xlwt
import os

def GetPDFName(fileName):
    if os.path.splitext(fileName)[1] == '.pdf':
        fileName_without_suffix = os.path.splitext(fileName)[0]
    return fileName_without_suffix

if __name__=="__main__":
     PDF_folder_Path = 'pdfthing/'
     CSV_folder_Path = 'csvthing/'
     files_list = os.listdir(PDF_folder_Path)
     number_of_objects = len(files_list)
     for file_name in files_list:
          workbook = xlwt.Workbook()  # 定义workbook
          sheet = workbook.add_sheet('Sheet1')  # 添加sheet
          i = 0  # Excel起始位置
          object_name = GetPDFName(file_name)
          with pdfplumber.open(PDF_folder_Path + file_name) as pdf:
               print('\n')
               print('Loading data')
               print('\n')
               for page in pdf.pages:
                    # print(page.extract_text())
                    for table in page.extract_tables():
                         # print(table)
                         for row in table:
                              k = 0
                              for j in range(len(row)):
                                   a=str(row[j])
                                   if a == '':
                                        k = k + 1
                                   if row[j] == None:
                                        k = k + 1
                                   if (row[j] != None)&(a != ''):
                                        sheet.write(i, j-k, row[j])
                              i += 1
               pdf.close()
               workbook.save(CSV_folder_Path+object_name+ '.xls')
               print('\n')
               print('DONE!')
               print('保存位置：'+ CSV_folder_Path)
               print('\n')



