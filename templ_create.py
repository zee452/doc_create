# ==========================формирование  шаблона из пустого бланка=====================================================
# имена переменных читаются из БД PostgreSQL и пишутся в файл \\temlates<file_name>
import os
from sys import argv, exit
import docx
import psycopg2
import openpyxl
import re

try:
  prname, dir_name = argv
  varvel = [] # данные переменной из БД [переменная,значение,тип ]
  strnow = '' # текущий день и время
  vv =''
  pred = ''
  Const_st = []
#==========================читаем из БД значение переменной=============================================================
  # def get_var_val(s,di,bpr,bp,entp):# считать значение переменной из БД s - переменная
  #    ss = f"'{s}'"
  #    sa = 'select docp_p,docp_v,docp_t,docp_r,docp_c from docp where docp_p='+ss+' and doc_id ='+di+' and bpr_id ='+bpr+ \
  #         ' and bp_id =' + bp +' and entp_id ='+entp
  #    cursor.execute(sa)
  #    return cursor.fetchone()
  def find_files(dir_path):# находим все файлы в папке бланки
     f = []
     for (dirpath, dirnames, filenames) in os.walk(dir_path):
         f.extend(filenames)
         break
     return f

  def write_blank(s, r, c, t, doc):
      ss = f"'{s}'"
      ts = f"'{t}'"
      docs = f"'{os.path.basename(doc)}'"
      sa = 'insert into docb (docb_p,docb_s,docb_r,docb_c,doc_name) values ('+ss+ ',' + ts + ',' + str(r) + ',' + \
            str(c)+','+docs+') on conflict do nothing'
      cursor.execute(sa)
#============================создание документа  по шаблону=============================================================
  def doc_cr(file_name):
      global vv
#------------------------------обработка файлов txt---------------------------------------------------------------------
# невозможна
#-------------------------------python-docx----------------------------------------------------------------------------
      if ext[1] in ['.odf','.docx']:
         doc = docx.Document(file_name)
         # if len(doc.paragraphs) > 1:
         #    for par in doc.paragraphs:
         #        write_blank()
         n = len(doc.tables)
         if n > 0:
            for tab in doc.tables:
                r = -1
                for ro in tab.rows:
                   r +=1
                   c = -1
                   pred = ''
                   mm = len(Const_st)
                   for cell in ro.cells:
                       c += 1
                       s =  cell.text
                       # if len(s) > 1 and s != pred and (re.search(r"[а-яА-Я]",s) != None) and \
                       #    all(i not in s.lower() for i in Const_st):
                       if  s != pred and (re.search(r"[а-яА-Я]",s) != None):

                          if cell.paragraphs[0].runs[0].font.superscript:
                             t = 'H'
                             write_blank(s, r, c, t, file_name)
                          elif cell.paragraphs[0].runs[0].font.subscript:
                             t = 'D'
                             write_blank(s, r, c, t, file_name)
                          elif any(' '+i+' ' in ' '+s.lower()+' ' for i in Const_st):
                             t = 'R'
                             write_blank(s, r, c, t, file_name)
                          # paragraphs = cell.paragraphs
                          # for paragraph in paragraphs:
                          #     for run in paragraph.runs:
                          #         font = run.font

                       pred = s

#-------------------------------------xlrd------------------------------------------------------------------------------
      # if ext[1] =='.xls':
      #    workbook = xlrd.open_workbook(file_name)
      #    sheets_name = workbook.sheet_names()
      #    for names in sheets_name:
      #        worksheet = workbook.sheet_by_name(names)
      #        num_rows = worksheet.nrows
      #        num_cells = worksheet.ncols
      #        for row in range(num_rows):
      #            for sel in range(num_cells):
      #                val = worksheet.cell_value(row, sel)
      #                if val != None:
      #                  if (len(val) > 3):
      #                     var_rep(val)
      #        workbook.save(name)
#------------------------------------openpyxl---------------------------------------------------------------------------
      if ext[1] =='.xlsx':
         workbook = openpyxl.load_workbook(file_name,data_only=True)
         sheet = workbook.active
         for row in range(sheet.max_row):
             row += 1
             for sel in range(sheet.max_column):
                 sel += 1
                 val = sheet.cell(row, sel).value
                 if val != None:
                   if (len(val) > 3):
                      if var_rep(val) == 0:
                         sheet.cell(row,sel).value = sheet.cell(row,sel).value.replace(varvel[0],varvel[1])
         workbook.save(file_name)
 #======================основная программа=============================================================================
  print('Идет запись данных в файл. Ждите...')
  conn = psycopg2.connect(host='localhost', database='BP', user='postgres', password='rfn15')
  # Получаем объект курсора для выполнения SQL-запросов
  cursor = conn.cursor()
  conn.autocommit = True
  nn = dir_name[0:dir_name.rfind('\\') + 1]+'const.txt'
  with open(nn, "r",encoding='UTF8',) as rf:
      lines = rf.readlines()
  for line in  lines:
      Const_st.append(line.strip())
  for file_name in find_files(dir_name):
      file_name = dir_name+'\\'+file_name
      ext = os.path.splitext(file_name)
      doc_cr(file_name)
  cursor.close()
  conn.close()
  exit(0)

except FileNotFoundError:
       print('Path not found-' + dir_name)
       exit(-1)
except psycopg2.Error:
       print ('ошибка БД')
       exit(-1)