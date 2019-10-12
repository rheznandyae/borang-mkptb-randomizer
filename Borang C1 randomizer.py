'''
Made by Mustafa Zaki Assagaf
This program will randomize borang C1 score for MPKT-B UI Class
inspired by Francis Wibosono
'''
from xlrd import open_workbook
import xlwt
from xlutils.copy import copy
from random import randint

nama_penilai, nama_kelompok, kelas, fakultas, file_name = ['']*5
nama_presentan, npm_presentan, min_val, max_val = [], [], 3, 5


def style(num):
    if num % 2:
        return xlwt.easyxf('align:vert centre, horiz center;pattern: pattern solid, \
            fore-colour white;border: left thin,right thin,top thin,bottom thin')                           
    return xlwt.easyxf('align:vert centre, horiz center;pattern: pattern solid, \
        fore-colour gray25;border: left thin,right thin,top thin,bottom thin')


def fill():
    global nama_penilai; global nama_kelompok; global kelas; global fakultas; global nama_presentan
    global npm_presentan; global file_name; global min_val; global max_val


    nama_penilai = input('Nama penilai : ')
    nama_kelompok = input('Nama kelompok : ')
    kelas = input('Nama kelas : ')
    fakultas = input('Nama Fakultas : ')

    min_val = int(input('Nilai terendah : ')) or 3
    max_val = int(input('Nilai tertinggi : ')) or 5

    nama_presentan = input('Nama mahasiswa. Dipisah dengan koma : ').split(',')
    nama_presentan = nama_presentan if nama_presentan != [''] else [None] * 7

    npm_presentan = input('npm mahasiswa. Dipisah dengan koma : ').split(',')
    npm_presentan = npm_presentan if npm_presentan != [''] else [None] * 7

    file_name = input('Lokasi dan nama file (ex: c:/folder/149.xls) : ')
    file_name = file_name if file_name is not '' else '149.xls'


def default():
    global nama_penilai; global nama_kelompok; global kelas; global fakultas; global nama_presentan
    global npm_presentan; global file_name; global min_val; global max_val

    with open('default.txt', 'r') as file:
        data = file.read().split('\n')

        nama_penilai, nama_kelompok, kelas, fakultas, min_val, \
        max_val, nama_presentan, npm_presentan, file_name = data

    nama_presentan = nama_presentan.split(',')
    npm_presentan = npm_presentan.split(',')
    min_val, max_val=int(min_val), int(max_val)


def openFile():
    global wb
    rb = open_workbook(file_name, formatting_info=True)
    wb = copy(rb)


def write():
    global nama_penilai; global nama_kelompok; global kelas; global fakultas; global nama_presentan
    global npm_presentan; global file_name; global min_val; global max_val

    s = wb.get_sheet(0)
    s.write(5, 2, nama_penilai)
    s.write(5, 5, nama_kelompok)
    s.write(5, 7, kelas)
    s.write(5, 9, fakultas)

    for i, mahasiswa in enumerate(nama_presentan, start=7):
        s.write(i, 1, nama_presentan[i-7], style(1))
        s.write(i, 2, npm_presentan[i-7], style(2))
        s.write(i, 10, xlwt.Formula(f'((E{i+1}*20)+(F{i+1}*10)+(G{i+1}*20)+\
            (H{i+1}*15)+(I{i+1}*15)+(J{i+1}*20))/5'), style(10))
        for j in range(3, 10):
            s.write(i, j, randint(min_val, max_val), style(j))


def save():
    wb.save(file_name)


def main():
    mode = input('Masukan mode input (kosongkan jika ingin isi manual) : ')
    if mode: 
        default()
    else: 
        fill()

    openFile()
    write()
    save()


if __name__ == '__main__':
    main()
