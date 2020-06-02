# coding=u8

"""

Create_at: 2020/3/8 9:46
Update_at: 2020/3/8 9:46
Create_by:
           zhuwei
@python: 
@description:

"""

# 导入文件夹
import os
from tkinter import messagebox
import xlwt
from tkinter import *

from chardet import detect

def get_code_type(file):
    with open(file, 'rb+') as f:
        content = f.read()
        return detect(content)['encoding']


def get_column(file, start_row, target_column, num, column_name):
    """"获取一个文件中的指定列，并返回list，第一个元素为column_name"""
    data = [column_name]
    with open(file, 'r', encoding=get_code_type(file)) as f:
        lines = f.readlines()[start_row - 1: start_row + num - 1]
        for row, line in enumerate(lines, start_row):
            data.append(float(line.split()[target_column - 1]))

    return data




def handle(source_path, target_path, start_row, target_column, num):

    Excel = xlwt.Workbook(encoding='utf-8', style_compression=0)
    table = Excel.add_sheet('data', cell_overwrite_ok=True)

    # 获取物质名称
    col = 0
    for path in os.listdir(source_path.get()):
        try:
            if path.split('.')[-1] == 'D':
                file = source_path.get() + '\\' + path + '\\' + '定量报告.txt'
                row = 0
                with open(file, 'r', encoding=get_code_type(file)) as f:
                    names = ['Mater\\Date']
                    lines = f.readlines()[int(start_row.get()) - 1: int(start_row.get()) + int(num.get()) - 1]
                    for r, line in enumerate(lines, start=int(start_row.get())):
                        try:
                            names.append(line.split()[1])
                        except Exception as e:
                            # messagebox.showinfo(title='错误信息',
                            #                     message='文件{0}的第{1}行数据有误,请核对参数和文件'.format(path, r))
                            names.append('')
                            continue

                    for na in names:
                        table.write(row, col, na)
                        row += 1
        except Exception as e:
            messagebox.showinfo(title='错误信息', message='文件{0}的数据格式有误，请核对参数和文件'.format(path))
            continue
        else:
            if path.split('.')[-1] == 'D':
                break

    col = 1
    for path in os.listdir(source_path.get()):
        try:
            if path.split('.')[-1] == 'D':
                file = source_path.get() + '\\' + path + '\\' + '定量报告.txt'
                row = 0
                with open(file, 'r', encoding=get_code_type(file)) as f:
                    data = []
                    column_name = file.split('\\')[-2][:10]
                    data.append(column_name)
                    lines = f.readlines()[(int(start_row.get())-1): int(start_row.get()) + int(num.get()) - 1]
                    for r, line in enumerate(lines, start=int(start_row.get())):
                        try:
                            data.append(float(line.split()[int(target_column.get())-1]))
                        except Exception as e:
                            # messagebox.showinfo(title='错误信息',
                            #                     message='文件{0}的第{1}行数据有误,请核对参数和文件'.format(path, r))
                            data.append('')
                            continue

                    for d in data:
                        table.write(row, col,d)
                        row += 1
                    col += 1
        except Exception as e:
            messagebox.showinfo(title='错误信息', message='{0}的数据格式有误, 请核对参数和文件'.format(path))
            col += 1
            continue
    Excel.save(target_path.get())  # Excel表保存为world.xls
    messagebox.showinfo(title='处理结果', message='成功处理{0}个文件'.format(col-1))
    # source_path.delete(0, END)
    # target_path.delete(0, END)
    # start_row.delete(0, END)
    # target_column.delete(0, END)
    # num.delete(0, END)



root = Tk()
root.iconbitmap(r'.\\favicon.ico')
root.title('FXZ-数据分析  @CreatedByZhu.')
root.geometry('600x285')

la2 = Label(root, text='请输入结果文件名:', justify=LEFT)
la2.grid(row=2, column=0, sticky=W)
desktop = StringVar(value=os.path.join(os.path.expanduser('~'), 'Desktop\\物质定量结果.xls'))
target_path = Entry(root, textvariable=desktop, width=35)
target_path.grid()

la1 = Label(root, text='请输入源文件夹:', justify=LEFT)
la1.grid(column=0, sticky=W)
source_path = Entry(root, width=35)
source_path.grid()


la3 = Label(root, text='请输入采集起始行:', justify=LEFT)
la3.grid(column=0, sticky=W)
start_row = Entry(root, textvariable=IntVar(value=26), width=35)
start_row.grid()

la5 = Label(root, text='请输入物质数量:', justify=LEFT)
la5.grid(column=0, sticky=W)
num = Entry(root, textvariable=IntVar(value=115), width=35)
num.grid()

la4 = Label(root, text='请输入采集列:', justify=LEFT)
la4.grid(column=0, sticky=W)
target_column = Entry(root, textvariable=IntVar(value=6), width=35)
target_column.grid()

btn1 = Button(root, text='确定', command=lambda: handle(source_path, target_path, start_row, target_column, num))
btn1.place(relx=0.65, rely=0.65, relwidth=0.3, relheight=0.1)

root.mainloop()

