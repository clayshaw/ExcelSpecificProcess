import os
from tkinter import *
import tkinter.messagebox

def rename_files(folder_path, prefix):
    # 遍歷指定資料夾下的所有檔案和資料夾
    for root, dirs, files in os.walk(folder_path):
        for old_name in files:
            # 取得檔案的副檔名 + 檔名
            name, extension = os.path.splitext(old_name)
            # print(name + '   ' + extension)
            # 組合新的檔案名稱
            new_name = str(name).upper() + extension
            # 重新命名檔案
            os.rename(os.path.join(root, old_name), os.path.join(root, new_name))

def inspect(folder_path):
    prefix = ''
    if os.path.exists(folder_path):
        rename_files(folder_path, prefix)
        anwser = tkinter.messagebox.askokcancel('Success 已完成處理')
    else:
        anwser = tkinter.messagebox.askokcancel('Warning 並非正確的檔案位置，請輸入正確檔案位置')




root = Tk()
root.geometry('460x240')
root.title("Excel")


lab1 = Label(root,text='請輸入資料夾位置(絕對路徑)')
lab1.place(relx=0.3,rely=0.1)
lab2 = Label(root,text='範例:  C:\\Users\\path\\folder')
lab2.place(relx=0.3,rely=0.2)

inp1 = Entry(root)
inp1.place(relx=0.15,rely=0.4,relwidth=0.7,relheight=0.1)

btn = Button(root,text='執行',command=lambda: inspect(inp1.get()))
btn.place(relx = 0.4,rely=0.6,relwidth=0.2,relheight=0.1)

root.mainloop()

