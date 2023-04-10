# 考试信息查询工具
该程序是一个基于Python的GUI应用程序。它能够帮助用户查询考试人员的信息，使用了Tkinter和Openpyxl两个Python库，并提供了一个简单易用的界面。

界面说明
程序的窗口尺寸为800x400，带有一个标题栏和一个画布。画布上是一个标签和三个输入框，用于输入查询所需的姓名、身份证号和考试类别。此外，还有两个按钮，分别用于导入Excel文件和查询考试人员信息。

## 功能说明
### 导入Excel文件
通过点击“导入文件”按钮，可以选择一个Excel文件。该文件应该包含要查询的考试人员信息，其中包括姓名、身份证号和考试类别。如果选择的文件不符合要求，会弹出一个消息框进行提示。

### 查询考试人员信息
在输入框中输入查询所需信息，然后单击“查询”按钮。程序将读取已经导入的Excel文件并查找符合条件的考试人员。如果找到了匹配的条目，则会弹出一个消息框显示查询结果；否则，会弹出一个提示框告诉用户“查无此人”。

### 注意事项
如果用户单击“查询”按钮而没有先导入Excel文件，则会弹出一个消息框进行提示。此外，我们强烈建议用户在使用本程序之前备份他们的Excel文件，以避免意外修改或删除数据。

## 代码说明
### 导入所需包
```python
import tkinter as tk
from tkinter import messagebox, filedialog, PhotoImage
import openpyxl
```
在代码中导入了三个Python库：tkinter、openpyxl和PhotoImage。其中，tkinter提供了GUI界面的创建和操作功能，openpyxl用于读取Excel文件中的数据，PhotoImage则用于设置图标。

### 创建窗口
```python
root = tk.Tk()
root.iconphoto(False, PhotoImage(file='dog.png'))
root.title("考试信息查询")
root.geometry("800x400")
```
使用Tkinter创建了一个窗口，设置了窗口图标和标题，并将窗口大小设置为800x400。

### 创建标签及输入框
```python
title_label = tk.Label(root, text="考试人员信息查询", font=("微软雅黑", 20), fg="#333333")
...
```
在窗口中创建了一个标签，用于显示程序名称。
```python
query_frame = tk.Frame(input_frame)
...
```
创建一个Frame，并在其中创建三个输入框，分别用于输入姓名、身份证号和考试类别。

### 实现查询功能
```python
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
```
当用户单击“查询”按钮时，程序会读取输入框中的值，然后在已导入的Excel文件中查找符合条件的考试人员。如果找到了匹配的条目，则会弹出一个消息框显示查询结果；否则，会弹出一个提示框告诉用户“查无此人”。

### 实现导入Excel文件功能
```python
def import_excel():
    global file_path, workbook, worksheet

    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])

    if file_path == "":
        messagebox.showwarning("提示", "请选择文件")
        return

    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active

    messagebox.showinfo("提示", "导入成功")
```
当用户单击“导入文件”按钮时，程序会让用户选择要导入的Excel文件，并加载工作表。如果文件不符合要求，会弹出一个消息框进行提示。

### 创建按钮并绑定事件
```python
import_button = tk.Button(root, text="导入文件", font=("微软雅黑", 14), bg="#007ACC", fg="#FFFFFF",
                          command=import_excel)
import_button.pack(pady=20)

search_button = tk.Button(root, text="查询", font=("微软雅黑", 14), bg="#007ACC", fg="#FFFFFF", command=search)
search_button.pack(pady=20)
```
最后，创建导入文件和查询按钮，并分别绑定事件触发函数。