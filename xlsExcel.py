import xlrd
import xlwt
import math
from tkinter import *
import tkinter.messagebox
import os

def cmpSame(s1,s2):
    
    if(len(s1)==len(s2)):
        return False
    elif(math.fabs(len(s1)-len(s2))==1):
        if(len(s1)>len(s2) and s1[:-1]==s2 and s1[-1]=='I'):
            return True
        elif(s1==s2[:-1] and s2[-1]=='I'):
            return True
        else: 
            return False
    else:
        return False
    

def mainfunc(s):
    Rworkbook = xlrd.open_workbook(s)
    Wworkbook = xlwt.Workbook(encoding="utf-8",style_compression=0)

    Rworksheet = Rworkbook.sheet_by_index(0)
    Rnrows = Rworksheet.nrows
    Rncols = Rworksheet.ncols

    Wsheet = Wworkbook.add_sheet('run', cell_overwrite_ok=True)

    for i in range(0,Rnrows):
        Wsheet.write(i,0,Rworksheet.row_values(i)[0])
        Wsheet.write(i,1,Rworksheet.row_values(i)[1])
        Wsheet.write(i,2,Rworksheet.row_values(i)[2])

    iscon = False
    cnt = 1
    Wsheet.write(0,5,'新的列標籤')
    Wsheet.write(0,6,'價錢')
    Wsheet.write(0,7,'數量')
    for i in range(0,Rnrows-1):
        
        if(iscon):
            iscon = False
            continue
        if(cmpSame(Rworksheet.row_values(i)[0],Rworksheet.row_values(i+1)[0]) == True):
            txtName = str(Rworksheet.row_values(i)[0]+'+'+Rworksheet.row_values(i+1)[0])
            txtMoney = Rworksheet.row_values(i)[1]+Rworksheet.row_values(i+1)[1]
            txtNum = Rworksheet.row_values(i)[2]+Rworksheet.row_values(i+1)[2]
            Wsheet.write(cnt,5,txtName)
            Wsheet.write(cnt,6,int(txtMoney))
            Wsheet.write(cnt,7,int(txtNum))
            iscon = True
        else:
            Wsheet.write(cnt,5,Rworksheet.row_values(i)[0])
            Wsheet.write(cnt,6,Rworksheet.row_values(i)[1])
            Wsheet.write(cnt,7,Rworksheet.row_values(i)[2])
        cnt+=1
    Wworkbook.save(s)
            


#Test File
# mainfunc('C:\\Users\\user\\Desktop\\workspace\\python\\excel\\data.xls')


def inspect(val):
    filena, extension = os.path.splitext(val)
    if(os.path.exists(val) and (extension == '.xls')):
        mainfunc(val)
        anwser = tkinter.messagebox.askokcancel('Success',val+'已完成處理')
        return False
    else:
        anwser = tkinter.messagebox.askokcancel('Warning',val+'並非正確的檔案位置，請輸入正確檔案位置及檔名')
        return True


root = Tk()
root.geometry('460x240')
root.title("Excel")


lab1 = Label(root,text='請輸入excel檔案位置(絕對路徑)於下方表格')
lab1.place(relx=0.25,rely=0.1)
lab2 = Label(root,text='範例:  C:\\Users\\excel\\data.xls')
lab2.place(relx=0.3,rely=0.2)
lab3 = Label(root,text='!!!使用前務必將需整理資料之名稱排序!!!')
lab3.place(relx=0.25,rely=0.3)

inp1 = Entry(root)
inp1.place(relx=0.15,rely=0.4,relwidth=0.7,relheight=0.1)

btn = Button(root,text='執行',command=lambda: inspect(inp1.get()))
btn.place(relx = 0.4,rely=0.6,relwidth=0.2,relheight=0.1)

root.mainloop()