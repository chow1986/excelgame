import xlrd
import xlwt
import os
import tkinter as tk
import time
#创建学员类
class employee:
    totalnum=0
    def __init__(self, name, id, gender,phonenum,wp,finishdata,traintype,area):
        self.name = name
        self.id = id
        self.gender = gender
        self.phonenum = phonenum
        self.wp = wp
        self.finishdata = finishdata
        self.traintype = traintype
        self.area = area
        employee.totalnum+=1

#获取文件列表
fileList = os.listdir(os.getcwd())
def mainf(onlinefinishtime,offlinefinishtime,fileName):
    # 遍历文件夹中所有文件
    if fileName[-3:len(fileName)]=="xls":
        data = xlrd.open_workbook(fileName)
        table = data.sheets()[0]
        #将表格内容写入类
        #创建新表格
        newtable=[]
        onlinefinishtime=[onlinefinishtime]*int(table.nrows-1)
        traintype=fileName.split("-")[0]+"-"+fileName.split("-")[1]
        traintype=[traintype]*int(table.nrows-1)
        newtablehead=['姓名','身份证','性别','联系电话','工作单位','培训完成日期','培训类别']
        newtable.append(table.col_values(2,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(1,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(3,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(6,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(5,start_rowx=1,end_rowx=None))
        newtable.append(onlinefinishtime)
        newtable.append(traintype)
        newtable.append(table.col_values(16,start_rowx=1,end_rowx=None))


        #将表格内容写入类:
        classdata=[]
        classd=[]
        for data in newtable:
            for d in data:
                classdata.append(d)
        for i in range(0,table.nrows-1):
            classd.append(employee(classdata[i],classdata[i+(table.nrows-1)*1],classdata[i+(table.nrows-1)*2],classdata[i+(table.nrows-1)*3],classdata[i+(table.nrows-1)*4],classdata[i+(table.nrows-1)*5],classdata[i+(table.nrows-1)*6],classdata[i+(table.nrows-1)*7]))
            i+=1
        #删除重复的地址
        areaname=list(set(table.col_values(16,start_rowx=1,end_rowx=None)))
        #表格分类
        for area in areaname:
            # 创建表格
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet(area,cell_overwrite_ok=True)
            #写入表头
            j = 0
            for head in newtablehead:
                worksheet.write(0, j, head)
                j = j + 1
            #分开创建表格
            m=1
            for cell in classd:
                if cell.area==area:
                    worksheet.write(m, 0, cell.name)
                    worksheet.write(m, 1, cell.id)
                    worksheet.write(m, 2, cell.gender)
                    worksheet.write(m, 3, cell.phonenum)
                    worksheet.write(m, 4, cell.wp)
                    worksheet.write(m, 5, cell.finishdata)
                    worksheet.write(m, 6, cell.traintype)
                    m=m+1
            folder=os.getcwd()+"\\"+fileName[-12:-4]
            if not os.path.exists(folder):#判断文件夹是否存在，如果不存在，则创建文件夹
                os.makedirs(folder)
            workbook.save(fileName[-12:-4]+"\\"+area+"-线上"+".xls")#保存文件



        #创建新表格
        newtable=[]
        offlinefinishtime=[offlinefinishtime]*int(table.nrows-1)
        traintype=fileName.split("-")[0]+"-"+fileName.split("-")[1]
        traintype=[traintype]*int(table.nrows-1)
        newtablehead=['姓名','身份证','性别','联系电话','工作单位','培训完成日期','培训类别']
        newtable.append(table.col_values(2,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(1,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(3,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(6,start_rowx=1,end_rowx=None))
        newtable.append(table.col_values(5,start_rowx=1,end_rowx=None))
        newtable.append(offlinefinishtime)
        newtable.append(traintype)
        newtable.append(table.col_values(16,start_rowx=1,end_rowx=None))

        #将表格内容写入类:
        classdata=[]
        classd=[]
        for data in newtable:
            for d in data:
                classdata.append(d)
        for i in range(0,table.nrows-1):
            classd.append(employee(classdata[i],classdata[i+(table.nrows-1)*1],classdata[i+(table.nrows-1)*2],classdata[i+(table.nrows-1)*3],classdata[i+(table.nrows-1)*4],classdata[i+(table.nrows-1)*5],classdata[i+(table.nrows-1)*6],classdata[i+(table.nrows-1)*7]))
            i+=1
        #删除重复的地址
        areaname=list(set(table.col_values(16,start_rowx=1,end_rowx=None)))
        #表格分类
        for area in areaname:
            # 创建表格
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet(area,cell_overwrite_ok=True)
            #写入表头
            j = 0
            for head in newtablehead:
                worksheet.write(0, j, head)
                j = j + 1
            #分开创建表格
            m=1
            for cell in classd:
                if cell.area==area:
                    worksheet.write(m, 0, cell.name)
                    worksheet.write(m, 1, cell.id)
                    worksheet.write(m, 2, cell.gender)
                    worksheet.write(m, 3, cell.phonenum)
                    worksheet.write(m, 4, cell.wp)
                    worksheet.write(m, 5, cell.finishdata)
                    worksheet.write(m, 6, cell.traintype)
                    m=m+1
            folder=os.getcwd()+"\\"+fileName[-12:-4]
            if not os.path.exists(folder):#判断文件夹是否存在，如果不存在，则创建文件夹
                os.makedirs(folder)
            workbook.save(fileName[-12:-4]+"\\"+area+"-线下"+".xls")#保存文件
#界面代码
window = tk.Tk()
window.title('My Window')
window.geometry('1000x500')
def getclassnumber(i,v,fileName):
    var = tk.StringVar()
    var.set(v)
    def pp():
        mainf(e1.get(),e2.get(),fileName)
        t.destroy()
        tk.Label(window, text="完成！").grid(row=i, column=7, padx=5, pady=10, ipadx=5, ipady=10)
    tk.Label(window, text="班级号:").grid(row=i, column=1, padx=5, pady=10, ipadx=5, ipady=10)
    tk.Label(window, textvariable=var).grid(row=i, column=2, padx=5, pady=10, ipadx=5, ipady=10)
    tk.Label(window, text="线上完成日期:").grid(row=i, column=3, padx=5, pady=10, ipadx=5, ipady=10)
    e1 = tk.Entry(window, show=None)
    e1.grid(row=i, column=4, padx=5, pady=10, ipadx=5, ipady=10)
    tk.Label(window, text="线下完成日期:").grid(row=i, column=5, padx=5, pady=10, ipadx=5, ipady=10)
    e2 = tk.Entry(window, show=None)
    e2.grid(row=i, column=6, padx=5, pady=10, ipadx=5, ipady=10)
    t=tk.Button(window, text="确定", width=10, height=1, command=pp)
    t.grid(row=i, column=7, padx=5, pady=10, ipadx=5,ipady=10)

def now():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

#加载班级号
s = '2020-06-04 00:00:00'
if now() < s:
    i=0
    for fileName in fileList:
        if fileName[-3:len(fileName)] == "xls":
            v=fileName[-12:-4]
            getclassnumber(i,v,fileName)
            i+=1
window.mainloop()
