from xlrd import open_workbook
import xlwt
from xlutils.copy import copy
from random import randint

def style(num):
    if num%2:
        return xlwt.easyxf('align:vert centre, horiz center;pattern: pattern solid, \
                                fore-colour white;border: left thin,right thin,top thin,bottom thin')
                                
    return xlwt.easyxf('align:vert centre, horiz center;pattern: pattern solid, \
                            fore-colour gray25;border: left thin,right thin,top thin,bottom thin')


nama_penilai= input('Nama penilai : ')
nama_kelompok = input('Nama kelompok : ')
kelas = input('Nama kelas : ')
fakultas = input('Nama Fakultas : ')

nama_presentan= input('Nama mahasiswa. dipisah dengan koma : ').split(',')
nama_presentan = nama_presentan if nama_presentan != [''] else [None] *7

npm_presentan= input('npm mahasiswa. dipisah dengan koma : ').split(',')
npm_presentan = npm_presentan if npm_presentan != [''] else [None] *7
file_name= input('Lokasi dan nama file (ex: c:/folder/149.xls) : ') or '149.xls'

rb = open_workbook(file_name, formatting_info=True)
wb = copy(rb)

s = wb.get_sheet(0)
s.write(5, 2, nama_penilai)
s.write(5, 5, nama_kelompok)
s.write(5, 7, kelas)
s.write(5, 9, fakultas)

for i, mahasiswa in enumerate(nama_presentan, start=7):
    s.write(i, 1, nama_presentan[i-7],style(1))
    s.write(i, 2, npm_presentan[i-7],style(2))
    s.write(i, 10, xlwt.Formula(f'((E{i+1}*20)+(F{i+1}*10)+(G{i+1}*20)+(H{i+1}*15)+(I{i+1}*15)+(J{i+1}*20))/5'),style(10))
    for j in range(3,10):
        s.write(i, j, randint(1,5), style(j))

wb.save(file_name)
