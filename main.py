import tkinter as tk
from tkinter import messagebox, filedialog, PhotoImage
import openpyxl

root = tk.Tk()
root.iconphoto(False, PhotoImage(file='dog.png'))
root.title("考试信息查询")
root.geometry("800x400")


title_label = tk.Label(root, text="考试人员信息查询", font=("微软雅黑", 20), fg="#333333")
title_label.pack(pady=20)

input_frame = tk.Frame(root)
input_frame.pack(pady=10)

input_frame = tk.Frame(root)
input_frame.pack(pady=10)

query_frame = tk.Frame(input_frame)
query_frame.pack()

name_label = tk.Label(query_frame, text="姓名：", font=("微软雅黑", 14))
name_label.pack(side=tk.LEFT, padx=5)
name_input = tk.Entry(query_frame, font=("微软雅黑", 14), width=10)
name_input.pack(side=tk.LEFT)

idn_label = tk.Label(query_frame, text="身份证号：", font=("微软雅黑", 14))
idn_label.pack(side=tk.LEFT, padx=5)
idn_input = tk.Entry(query_frame, font=("微软雅黑", 14), width=25)
idn_input.pack(side=tk.LEFT)

exam_label = tk.Label(query_frame, text="考试类别：", font=("微软雅黑", 14))
exam_label.pack(side=tk.LEFT, padx=5)
exam_input = tk.Entry(query_frame, font=("微软雅黑", 14), width=5)
exam_input.pack(side=tk.LEFT)


def search():
    global worksheet

    name = name_input.get()
    idn = idn_input.get()
    exam = exam_input.get()

    try:
        worksheet
    except NameError:
        messagebox.showwarning("提示", "请先导入文件")
        return

    right = 0

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        idn2 = str(row[1])
        a = (row[0] == name)
        b = (idn2 == idn)
        c = (row[2] == exam)
        # print(a)
        # print(type(row[1]))
        # print(type(idn))
        # print(c)

        v = a and b and c

        if v:
            right = right + 1
        else:
            right = right + 0

    if right != 0:
        right = str(right)
        messagebox.showinfo("提示", "查询成功,查询到" + right + "人")
    else:
        messagebox.showinfo("提示", "查无此人！")


def import_excel():
    global file_path, workbook, worksheet

    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])

    if file_path == "":
        messagebox.showwarning("提示", "请选择文件")
        return

    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active

    messagebox.showinfo("提示", "导入成功")


import_button = tk.Button(root, text="导入文件", font=("微软雅黑", 14), bg="#007ACC", fg="#FFFFFF",
                          command=import_excel)
import_button.pack(pady=20)

search_button = tk.Button(root, text="查询", font=("微软雅黑", 14), bg="#007ACC", fg="#FFFFFF", command=search)
search_button.pack(pady=20)

root.mainloop()
