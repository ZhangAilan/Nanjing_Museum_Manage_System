# -*- coding:utf-8 -*-
import pymssql                                    #提供数据库接口
from tkinter import ttk                           #提供了一组GUI部件，用于创建图形用户界面
import tkinter as tk
from tkinter import *                             #导入tkinter模块的所有内容
import tkinter.messagebox as messagebox           #导入弹窗模块
from PIL import Image,ImageTk                     #导入图像处理模块
import pandas as pd                               #导入pandas模块，用于读取excel文件
from cefpython3 import cefpython as cef           #导入cefpython模块，用于插入html文件
import sys, os                                    #导入系统模块
import threading                                  #导入线程模块


#起始页面
class StartPage:
    #初始化函数
    def __init__(self,parent_window):                                                       
        parent_window.destroy()                   #销毁父窗口
        self.window=tk.Tk()                                                                 
        self.window.title('南京市博物馆文物信息管理系统')               
        self.window.geometry('800x300+350+150')  

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\初始界面.png')
        image = image_file.resize((800, 300))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1)    

        label=Label(self.window,text='南京市博物馆文物信息管理系统',font=('宋体',30),fg='black')
        label.pack(pady=50)                                                                 

        button_1=tk.Button(self.window,text='管理员登录',font=('宋体',15),width=15,height=2,command=lambda:AdminPage(self.window))
        button_1.place(x=120,y=180)
        button_2=tk.Button(self.window,text="数据库初始化",font=('宋体',15),width=15,height=2,command=self.Initialization) 
        button_2.place(x=320,y=180)
        button_3=tk.Button(self.window,text="退出系统",font=('宋体',15),width=15,height=2,command=self.window.destroy)
        button_3.place(x=520,y=180)

        self.window.mainloop()                    #进入消息循环,由此GUI应用可以开始接受和处理事件了


    #创建数据库
    def Initialization(self):
        db = pymssql.connect('ZYH', 'sa', '123',charset='UTF-8')  # 连接数据库,服务器名,账户,密码
        cursor = db.cursor()  # 创建游标
        print('开始创建数据库......')
        sql = ''' 
        IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'Nanjing_Museum_Manage_System')
        BEGIN
            CREATE DATABASE Nanjing_Museum_Manage_System ON PRIMARY
            (
                NAME='NJMS_data',   --主文件逻辑名
                FILENAME='D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\database\\NJMS_data.mdf',   --主文件路径
                SIZE=5MB,   --主文件初始大小
                MAXSIZE=500MB,   --主文件最大大小
                FILEGROWTH=15%   --主文件自动增长大小
            )
            LOG ON
            (
                NAME='NJMS_log',   --日志文件逻辑名
                FILENAME='D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\database\\NJMS_log.ldf',   --日志文件路径
                SIZE=5MB,   --日志文件初始大小
                MAXSIZE=50MB,   --日志文件最大大小
                FILEGROWTH=1MB   --日志文件自动增长大小        
            )
            PRINT '创建数据库成功！'
        END
        ELSE
        BEGIN
            PRINT '数据库已存在，无需创建！'
        END
        '''
        try:
            db.autocommit(True)  # 自动提交
            cursor.execute(sql)  # 执行sql语句
            db.commit()  # 提交事务
            print('创建数据库成功！')
        except Exception as e:
            messagebox.showinfo('提示', f'数据库连接失败！{e}')
        cursor.close()  # 关闭游标
        db.close()  # 关闭数据库连接
        self.Create_Table()  # 创建表
    

    #创建表
    def Create_Table(self):
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System',charset='UTF-8')   #连接数据库
        cursor=db.cursor()                                                     #创建游标
        print('开始创建表......')
        sql='''
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='博物馆信息表' AND xtype='U')
        BEGIN
            CREATE TABLE dbo.博物馆信息表
            (
                博物馆编号 VARCHAR(20) PRIMARY KEY,    --博物馆编号
                博物馆名称 VARCHAR(100) NOT NULL,      --博物馆名称
                博物馆地址 VARCHAR(100) NOT NULL,      --博物馆地址
                博物馆电话 VARCHAR(100) NOT NULL,      --博物馆电话
            )
        END

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='文物信息表' AND xtype='U')
        BEGIN
            CREATE TABLE dbo.文物信息表
            (
                文物编号 VARCHAR(20) PRIMARY KEY,   --文物编号
                文物名称 VARCHAR(100) NOT NULL,      --文物名称
                文物类别 VARCHAR(100) NOT NULL,      --文物类别
                文物来源 VARCHAR(100) NOT NULL,      --文物来源
                文物年代 VARCHAR(100) NOT NULL,      --文物年代
                文物馆藏地 VARCHAR(100) NOT NULL     --文物馆藏地
            )
        END

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='借出文物表' AND xtype='U')
        BEGIN
            CREATE TABLE dbo.借出文物表
            (
                借出编号 VARCHAR(20) PRIMARY KEY,   --借出编号
                文物编号 VARCHAR(20) NOT NULL,      --文物编号
                借出地 VARCHAR(100) NOT NULL,        --借出地
                借出日期 VARCHAR(100) NOT NULL,             --借出日期
                归还日期 VARCHAR(100) NOT NULL,             --归还日期
                FOREIGN KEY (文物编号) REFERENCES dbo.文物信息表(文物编号)
            )
        END

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='借入文物表' AND xtype='U')
        BEGIN
            CREATE TABLE dbo.借入文物表
            (
                借入编号 VARCHAR(20) PRIMARY KEY,   --借入编号
                文物编号 VARCHAR(20) NOT NULL,      --文物编号
                借入地 VARCHAR(100) NOT NULL,        --借入地
                借入日期 VARCHAR(100) NOT NULL,             --借入日期
                归还日期 VARCHAR(100) NOT NULL,             --归还日期
                FOREIGN KEY (文物编号) REFERENCES dbo.文物信息表(文物编号)
            )
        END

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='工作人员信息表' AND xtype='U')
        BEGIN
            CREATE TABLE dbo.工作人员信息表
            (
                编号 VARCHAR(20) PRIMARY KEY,   --工作人员编号
                姓名 VARCHAR(20) NOT NULL,      --工作人员姓名
                性别 VARCHAR(10) NOT NULL,      --工作人员性别
                年龄 VARCHAR(10) NOT NULL,      --工作人员年龄
                职位 VARCHAR(20) NOT NULL,      --工作人员职位
                电话 VARCHAR(20) NOT NULL,      --工作人员电话
                单位 VARCHAR(100) NOT NULL,      --工作人员单位
            )
        END

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='管理员信息表' AND xtype='U')
        BEGIN
            CREATE TABLE dbo.管理员信息表
            (
                用户 VARCHAR(20) PRIMARY KEY,   --管理员用户名
                密码 VARCHAR(20) NOT NULL,      --管理员密码
            )
        END 

        '''
        try:
            cursor.execute(sql)                  #执行sql语句
            db.commit()                          #提交事务
            print('创建表成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'创建表失败！{e}')
        cursor.close()                           #关闭游标
        db.close()                               #关闭数据库连接
        self.Insert_Data_From_Excel()            #向表中插入excel数据

    
    #向表中插入excel数据
    def Insert_Data_From_Excel(self):
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System',charset='UTF-8')
        cursor=db.cursor()
        print('开始向表中插入数据......')

        #读取excel文件
        xlsx=pd.ExcelFile('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\原始数据.xlsx')

        #遍历excel文件中的所有表
        success = True
        for sheet_name in xlsx.sheet_names:
            #获取表格数据
            df = xlsx.parse(sheet_name)
            #获取表名
            table_name = sheet_name
            #将数据插入表中
 
            for index, row in df.iterrows():
                values = tuple(row)
                placeholders = ', '.join(['%s'] * len(values))
                sql = f"INSERT INTO {table_name} VALUES ({placeholders})"
                try:
                    cursor.execute(sql, values)
                except Exception as e:
                    if 'PRIMARY KEY' in str(e):
                        continue  # Skip if primary key violation
                    messagebox.showinfo('提示', f'插入数据失败！{e}')
                    success = False
                    break
        
        if success:
            messagebox.showinfo('提示', '插入数据成功！')
            print('插入数据成功！')
        
        db.commit()                              #提交事务
        cursor.close()                           #关闭游标
        db.close()                               #关闭数据库连接
    
    

#管理员登录页面
class AdminPage:
    def __init__(self, parent_window):
        parent_window.destroy()
        self.window = tk.Tk()
        self.window.title('管理员登录')
        self.window.geometry('550x350+500+100')

        # 美化
        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\管理员登录界面.jpg')
        image = image_file.resize((550, 350))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.pack(expand=True, fill='both')

        # 用户名与密码
        tk.Label(self.window, text='用户：', font=('宋体', 20)).place(x=80, y=150)
        tk.Label(self.window, text='密码：', font=('宋体', 20)).place(x=80, y=200)
        self.admin_username = tk.Entry(self.window, font=('宋体', 20))
        self.admin_username.place(x=160, y=150)
        self.admin_password = tk.Entry(self.window, show='*', font=('宋体', 20))
        self.admin_password.place(x=160, y=200)

        # 登录与返回按钮
        Button_login = tk.Button(self.window, text='登录', font=('宋体', 20), width=8, height=1,command=self.login)
        Button_login.place(x=130, y=250)
        Button_back = tk.Button(self.window, text='返回', font=('宋体', 20), width=8, height=1, command=self.back)
        Button_back.place(x=280, y=250)

        self.window.mainloop()
    

    def back(self):
        StartPage(self.window) 


    #登录函数
    def login(self):
        # 查询管理员信息表
        db = pymssql.connect('ZYH', 'sa', '123', 'Nanjing_Museum_Manage_System')
        cursor = db.cursor()
        sql = "SELECT * FROM dbo.管理员信息表 where 用户='%s'" % (self.admin_username.get())  # 得到管理员用户名
        username = None  # Initialize username with a default value
        try:
            cursor.execute(sql)
            results = cursor.fetchall()  # 获取查询结果，返回二元元组，包括用户名和密码
            for row in results:
                username = row[0]
                password = row[1]
        except Exception as e:
            messagebox.showinfo('提示', f'查询失败！{e}')
        db.close()

        print("正在登录......")
        # 判断用户名和密码是否正确
        if username is not None and self.admin_username.get() == username and self.admin_password.get() == password:
            print("登录成功！")
            SubAdminPage(self.window)
        else:
            messagebox.showinfo('提示', '用户名或密码错误！')
            self.admin_username.delete(0, END)
            self.admin_password.delete(0, END)



#信息管理导航页面
class SubAdminPage:
    def __init__(self,parent_window):
        parent_window.destroy()
        self.window=tk.Tk()
        self.window.title('信息管理导航')
        self.window.geometry('300x600+500+120')

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\信息管理导航界面.jfif')
        image = image_file.resize((300, 600))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1)  

        # 创建按钮
        museum_btn = tk.Button(self.window, text='博物馆信息管理',font=('宋体', 15), width=15, height=2,command=lambda:MuseumManage(self.window))
        museum_btn.place(x=80,y=280)
        relic_btn = tk.Button(self.window, text='文物信息管理',font=('宋体', 15), width=15, height=2,command=lambda:CulturalrelicManage(self.window))
        relic_btn.place(x=80,y=350)
        staff_btn = tk.Button(self.window, text='工作人员管理',font=('宋体', 15), width=15, height=2,command=lambda:StaffManage(self.window))
        staff_btn.place(x=80,y=420)
        back_btn = tk.Button(self.window, text='返回',font=('宋体', 10), width=10, height=2,command=self.back)
        back_btn.place(x=110,y=540)

        self.window.mainloop()

    def back(self):
        AdminPage(self.window)



#博物馆信息管理页面
class MuseumManage:    
    def __init__(self,parent_window):
        parent_window.destroy()
        self.window=tk.Tk()
        self.window.title('博物馆信息管理')
        self.window.geometry('1200x685+150+50')

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\博物馆信息界面.jpg')
        image = image_file.resize((1200, 685))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1) 

        #定义中间列表
        self.columns=('编号','名称','地址','电话')
        self.tree=ttk.Treeview(self.window, show='headings', height=10, columns=self.columns)  
        self.tree.bind('<ButtonRelease-1>', self.click_show)
        self.vbar = ttk.Scrollbar(self.window, orient=VERTICAL, command=self.tree.yview)  #竖直滚动条
        self.tree.configure(yscrollcommand=self.vbar.set)   #将滚动条与列表关联

        self.tree.column('编号', width=35, anchor='center')  # 修改列的锚点为居中对齐
        self.tree.column('名称', width=170, anchor='w')  # 修改列的锚点为左对齐
        self.tree.column('地址', width=205, anchor='w')  # 修改列的锚点为左对齐
        self.tree.column('电话', width=90, anchor='w')   # 修改列的锚点为左对齐

        self.tree.place(relx=0.55, rely=0.6, anchor='w')  #列表居中显示
        self.vbar.place(relx=0.975, rely=0.6, anchor='w')  #竖直滚动条居中显示

        self.InsertData()  #向表中插入数据
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.tree_sort_column(self.tree, _col, False)) 
        
        # 添加标签
        label_id = tk.Label(self.window, text='博物馆编号:', font=('宋体', 12))
        label_id.grid(row=0, column=0, padx=675, pady=(50, 10))
        label_name = tk.Label(self.window, text='博物馆名称:', font=('宋体', 12))
        label_name.grid(row=1, column=0,padx=675, pady=(10, 10))
        label_address = tk.Label(self.window, text='博物馆地址:', font=('宋体', 12))
        label_address.grid(row=2, column=0,padx=675, pady=(10, 10))
        label_phone = tk.Label(self.window, text='博物馆电话:', font=('宋体', 12))
        label_phone.grid(row=3, column=0,padx=675, pady=(10, 10))

        # 添加文本框
        self.textbox_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_id.place(x=780, y=50)
        self.textbox_name = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_name.place(x=780, y=94)
        self.textbox_address = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_address.place(x=780, y=136)
        self.textbox_phone = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_phone.place(x=780, y=180)

        # 添加按钮，新建，更新，删除
        button_new = tk.Button(self.window, text='新建', font=('宋体', 12), width=10, height=1,command=self.NewData)
        button_new.place(x=760, y=230)
        button_update = tk.Button(self.window, text='更新', font=('宋体', 12), width=10, height=1,command=self.UpdateData)
        button_update.place(x=880, y=230)
        button_delete = tk.Button(self.window, text='删除', font=('宋体', 12), width=10, height=1,command=self.DeleteData)
        button_delete.place(x=1000, y=230)

        # 添加按钮，查询，返回
        button_search = tk.Button(self.window, text='查询', font=('宋体', 12), width=10, height=1,command=self.SearchData)
        button_search.place(x=700, y=550)
        self.textbox_search = tk.Text(self.window, width=20, height=1.4, font=('宋体', 12))
        self.textbox_search.place(x=800, y=553)
        button_back = tk.Button(self.window, text='返回', font=('宋体', 12), width=10, height=1, command=self.back)
        button_back.place(x=1050, y=550)

        #插入html,使用多线程
        threading.Thread(target=self.InsertHtml,args=()).start()  

        self.window.mainloop()
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)   #关闭窗口时调用on_close函数

    
    #插入html
    def InsertHtml(self):
        #插入html文件
        html_frame=tk.Frame(self.window,bg='white',width=580,height=550)
        html_frame.place(x=50,y=40)
        self.window.update()
        
        sys.excepthook=cef.ExceptHook   #设置异常处理函数
        cef.Initialize()                #初始化浏览器
        print('正在加载地图......')
        
        window_info=cef.WindowInfo(html_frame.winfo_id())   #获取浏览器窗口信息
        window_info.SetAsChild(html_frame.winfo_id(),[0,0,580,550])   #将浏览器窗口设置为子窗口
        self.browser=cef.CreateBrowserSync(window_info,url=os.path.abspath("NanjingCityMap.html"))   #创建浏览器对象
        bindings=cef.JavascriptBindings(bindToFrames=False,bindToPopups=False)   #绑定js
        self.browser.SetJavascriptBindings(bindings)    #将js绑定到浏览器对象
        
        #无限循环，不断调用消息循环，即CEF的消息会被持续地处理，但是每次处理都不会阻塞主线程太长时间。
        #这样保证CEF消息循环得到处理，同时也保证主线程可以处理其他任务。
        while True:
            cef.MessageLoopWork()   

    
    def on_close(self):
        self.browser.CloseBrowser(True)   #关闭浏览器
        cef.Shutdown()                    #关闭子线程
        self.window.destroy()             #关闭窗口

        self.window.mainloop()
    

    #向表中插入数据
    def InsertData(self):
        #清空列表
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        self.id=[]
        self.name=[]
        self.address=[]
        self.phone=[]

        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql="SELECT * FROM dbo.博物馆信息表"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results = [(row[0].encode('latin-1').decode('gbk'), row[1].encode('latin-1').decode('gbk'), row[2].encode('latin-1').decode('gbk'), row[3].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.id.append(row[0])
                self.name.append(row[1])
                self.address.append(row[2])
                self.phone.append(row[3])
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')
        db.close()

        for i in range(len(self.id)):
            self.tree.insert('', i, values=(self.id[i], self.name[i], self.address[i], self.phone[i]))   #插入数据

    
    #点击表头排序   
    def tree_sort_column(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        l.sort(reverse=reverse)
        # rearrange items in sorted positions
        for index, (_, k) in enumerate(l):  # Replace "val" with "_"
            tree.move(k, '', index)
        # reverse sort next time
        tree.heading(col, command=lambda: self.tree_sort_column(tree, col, not reverse))    # 重写表头，使之成为再点倒序的表头
    

    #删除文本框中的文字
    def delete_text(self):
        self.textbox_id.delete(1.0, tk.END)
        self.textbox_name.delete(1.0, tk.END)
        self.textbox_address.delete(1.0, tk.END)
        self.textbox_phone.delete(1.0, tk.END)


    #点击列表中的数据，将数据显示在文本框中
    def click_show(self,event):
        self.col=self.tree.identify_column(event.x)  #列
        self.row=self.tree.identify_row(event.y)     #行

        self.row_info=self.tree.item(self.row,'values')
        self.var_id=self.row_info[0]
        self.var_name=self.row_info[1]
        self.var_address=self.row_info[2]
        self.var_phone=self.row_info[3]

        self.delete_text()
        self.textbox_id.insert(tk.END, f"{self.var_id}\n")
        self.textbox_name.insert(tk.END, f"{self.var_name}\n")
        self.textbox_address.insert(tk.END, f"{self.var_address}\n")
        self.textbox_phone.insert(tk.END, f"{self.var_phone}\n")
        

    #新建
    def NewData(self):
        #将文本框中数据插入数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"INSERT INTO dbo.博物馆信息表 VALUES ('{self.textbox_id.get(1.0, tk.END).strip()}','{self.textbox_name.get(1.0, tk.END).strip()}','{self.textbox_address.get(1.0, tk.END).strip()}','{self.textbox_phone.get(1.0, tk.END).strip()}')"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','新建成功！')
            print('新建成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'新建失败！{e}')
        self.InsertData()  # 更新列表


    #更新：
    def UpdateData(self):
        #将文本框中的数据更新到数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"UPDATE dbo.博物馆信息表 SET 博物馆名称='{self.textbox_name.get(1.0, tk.END).strip()}',博物馆地址='{self.textbox_address.get(1.0, tk.END).strip()}',博物馆电话='{self.textbox_phone.get(1.0, tk.END).strip()}' WHERE 博物馆编号='{self.textbox_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','更新成功！')
            print('更新成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'更新失败！{e}')
        self.InsertData()   #更新列表


    #删除：
    def DeleteData(self):
        #将文本框中的数据从数据库中删除
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"DELETE FROM dbo.博物馆信息表 WHERE 博物馆编号='{self.textbox_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','删除成功！')
            print('删除成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'删除失败！{e}')
        self.InsertData()   #更新列表


    #查询：
    def SearchData(self):
        #清除文本框内容
        self.delete_text()
        #根据文本框中的博物馆名称查询数据库数据
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"SELECT * FROM dbo.博物馆信息表 WHERE 博物馆名称='{self.textbox_search.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results = [(row[0].encode('latin-1').decode('gbk'), row[1].encode('latin-1').decode('gbk'), row[2].encode('latin-1').decode('gbk'), row[3].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.textbox_id.insert(tk.END, f"{row[0]}\n")
                self.textbox_name.insert(tk.END, f"{row[1]}\n")
                self.textbox_address.insert(tk.END, f"{row[2]}\n")
                self.textbox_phone.insert(tk.END, f"{row[3]}\n")
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')
        
    
    def back(self):
        SubAdminPage(self.window)



#工作人员管理页面
class StaffManage:
    def __init__(self,parent_window):
        parent_window.destroy()
        self.window=tk.Tk()
        self.window.title('工作人员管理')
        self.window.geometry('600x685+400+50')

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\工作人员信息管理界面.jpg')
        image = image_file.resize((1200, 685))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1) 

        #定义中间列表
        self.columns=('编号','姓名','性别','年龄','职位','电话','单位')
        self.tree=ttk.Treeview(self.window, show='headings', height=10, columns=self.columns)  
        self.tree.bind('<ButtonRelease-1>', self.click_show)
        self.vbar = ttk.Scrollbar(self.window, orient=VERTICAL, command=self.tree.yview)  #竖直滚动条
        self.tree.configure(yscrollcommand=self.vbar.set)   #将滚动条与列表关联
        self.tree.column('编号', width=35, anchor='center')  # 修改列的锚点为居中对齐
        self.tree.column('姓名', width=60, anchor='w')  # 修改列的锚点为左对齐
        self.tree.column('性别', width=35, anchor='w')  
        self.tree.column('年龄', width=35, anchor='w')
        self.tree.column('职位', width=50, anchor='w')
        self.tree.column('电话', width=100, anchor='w')
        self.tree.column('单位', width=140, anchor='w')
        self.tree.place(relx=0.1, rely=0.7, anchor='w')  #列表居中显示
        self.vbar.place(relx=0.06, rely=0.7, anchor='w')  #竖直滚动条居中显示
        self.InsertData()  #向表中插入数据
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.tree_sort_column(self.tree, _col, False)) 

        # 添加标签
        x_label=80
        label_id = tk.Label(self.window, text='编号:', font=('宋体', 12))
        label_id.grid(row=0, column=0, padx=x_label, pady=(10, 10))
        label_name = tk.Label(self.window, text='姓名:', font=('宋体', 12))
        label_name.grid(row=1, column=0,padx=x_label, pady=(10, 10))
        label_gender= tk.Label(self.window, text='性别:', font=('宋体', 12))
        label_gender.grid(row=2, column=0, padx=x_label, pady=(10, 10))
        label_age = tk.Label(self.window, text='年龄:', font=('宋体', 12))
        label_age.grid(row=3, column=0,padx=x_label, pady=(10, 10))
        label_position = tk.Label(self.window, text='职位:', font=('宋体', 12))
        label_position.grid(row=4, column=0,padx=x_label, pady=(10, 10))
        label_phone = tk.Label(self.window, text='电话:', font=('宋体', 12))
        label_phone.grid(row=5, column=0,padx=x_label, pady=(10, 10))
        label_workplace = tk.Label(self.window, text='单位:', font=('宋体', 12))
        label_workplace.grid(row=6, column=0,padx=x_label, pady=(10, 10))

        # 添加文本框
        x_textbox=140
        self.textbox_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_id.place(x=x_textbox, y=10)
        self.textbox_name = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_name.place(x=x_textbox, y=54)
        self.textbox_gender= tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_gender.place(x=x_textbox, y=96)
        self.textbox_age = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_age.place(x=x_textbox, y=140)
        self.textbox_position = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_position.place(x=x_textbox, y=184)
        self.textbox_phone = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_phone.place(x=x_textbox, y=225)
        self.textbox_workplace = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_workplace.place(x=x_textbox, y=265)

        # 添加按钮，新建，更新，删除
        x_button=120
        button_new = tk.Button(self.window, text='新建', font=('宋体', 12), width=10, height=1,command=self.NewData)
        button_new.place(x=x_button, y=300)
        button_update = tk.Button(self.window, text='更新', font=('宋体', 12), width=10, height=1,command=self.UpdateData)
        button_update.place(x=x_button+120, y=300)
        button_delete = tk.Button(self.window, text='删除', font=('宋体', 12), width=10, height=1,command=self.DeleteData)
        button_delete.place(x=x_button+240, y=300)

        # 添加按钮，查询，返回
        button_search = tk.Button(self.window, text='查询', font=('宋体', 12), width=10, height=1,command=self.SearchData)
        button_search.place(x=100, y=620)
        self.textbox_search = tk.Text(self.window, width=20, height=1.4, font=('宋体', 12))
        self.textbox_search.place(x=200, y=623)
        button_back = tk.Button(self.window, text='返回', font=('宋体', 12), width=10, height=1, command=self.back)
        button_back.place(x=400, y=620)

        self.window.mainloop()  

    
    #新建
    def NewData(self):
        #将文本框中数据插入数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"INSERT INTO dbo.工作人员信息表 VALUES('{self.textbox_id.get(1.0,tk.END).strip()}','{self.textbox_name.get(1.0, tk.END).strip()}','{self.textbox_gender.get(1.0, tk.END).strip()}','{self.textbox_age.get(1.0, tk.END).strip()}','{self.textbox_position.get(1.0, tk.END).strip()}','{self.textbox_phone.get(1.0, tk.END).strip()}','{self.textbox_workplace.get(1.0, tk.END).strip()}')"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','新建成功！')
            print('新建成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'新建失败！{e}')
        self.InsertData()  # 更新列表


    #更新
    def UpdateData(self):
        #将文本框中的数据更新到数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"UPDATE 工作人员信息表 SET 姓名 = '{self.textbox_name.get(1.0, tk.END).strip()}', 性别 = '{self.textbox_gender.get(1.0, tk.END).strip()}', 年龄 = '{self.textbox_age.get(1.0, tk.END).strip()}', 职位 = '{self.textbox_position.get(1.0, tk.END).strip()}', 电话 = '{self.textbox_phone.get(1.0, tk.END).strip()}', 单位 = '{self.textbox_workplace.get(1.0, tk.END).strip()}' WHERE 编号 = '{self.textbox_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','更新成功！')
            print('更新成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'更新失败！{e}')
        self.InsertData()   #更新列表


    #删除
    def DeleteData(self):
        #将文本框中的数据从数据库中删除
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"DELETE FROM dbo.工作人员信息表 WHERE 编号='{self.textbox_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','删除成功！')
            print('删除成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'删除失败！{e}')
        self.InsertData()   #更新列表


    #查询
    def SearchData(self):
        #清除文本框内容
        self.delete_text()
        #根据文本框中的工作人员姓名查询数据库数据
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"SELECT * FROM dbo.工作人员信息表 WHERE 姓名='{self.textbox_search.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results = [(row[0].encode('latin-1').decode('gbk'), row[1].encode('latin-1').decode('gbk'), row[2].encode('latin-1').decode('gbk'), row[3].encode('latin-1').decode('gbk'), row[4].encode('latin-1').decode('gbk'), row[5].encode('latin-1').decode('gbk'), row[6].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.textbox_id.insert(tk.END, f"{row[0]}\n")
                self.textbox_name.insert(tk.END, f"{row[1]}\n")
                self.textbox_gender.insert(tk.END, f"{row[2]}\n")
                self.textbox_age.insert(tk.END, f"{row[3]}\n")
                self.textbox_position.insert(tk.END, f"{row[4]}\n")
                self.textbox_phone.insert(tk.END, f"{row[5]}\n")
                self.textbox_workplace.insert(tk.END, f"{row[6]}\n")
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')


    #点击列表中的数据，将数据显示在文本框中
    def click_show(self,event):
        self.col=self.tree.identify_column(event.x)  #列
        self.row=self.tree.identify_row(event.y)     #行

        self.row_info=self.tree.item(self.row,'values')
        self.var_id=self.row_info[0]
        self.var_name=self.row_info[1]
        self.var_gender=self.row_info[2]
        self.var_age=self.row_info[3]
        self.var_position=self.row_info[4]
        self.var_phone=self.row_info[5]
        self.var_workplace=self.row_info[6]

        self.delete_text()
        self.textbox_id.insert(tk.END, f"{self.var_id}\n")
        self.textbox_name.insert(tk.END, f"{self.var_name}\n")
        self.textbox_gender.insert(tk.END,f"{self.var_gender}\n")
        self.textbox_age.insert(tk.END, f"{self.var_age}\n")
        self.textbox_position.insert(tk.END, f"{self.var_position}\n")
        self.textbox_phone.insert(tk.END, f"{self.var_phone}\n")
        self.textbox_workplace.insert(tk.END, f"{self.var_workplace}\n")


    # 点击表头排序
    def tree_sort_column(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children('')]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.tree_sort_column(tree, col, not reverse))
    

    #删除文本框中的文字
    def delete_text(self):
        self.textbox_id.delete(1.0, tk.END)
        self.textbox_name.delete(1.0, tk.END)
        self.textbox_gender.delete(1.0, tk.END)
        self.textbox_age.delete(1.0, tk.END)
        self.textbox_position.delete(1.0, tk.END)
        self.textbox_phone.delete(1.0, tk.END)
        self.textbox_workplace.delete(1.0, tk.END)


    #向表中插入数据
    def InsertData(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.id=[]
        self.name=[]
        self.gender=[]
        self.age=[]
        self.position=[]
        self.phone=[]
        self.workplace=[]

        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql="SELECT * FROM dbo.工作人员信息表"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results=[(row[0].encode('latin-1').decode('gbk'),row[1].encode('latin-1').decode('gbk'),row[2].encode('latin-1').decode('gbk'),row[3].encode('latin-1').decode('gbk'),row[4].encode('latin-1').decode('gbk'),row[5].encode('latin-1').decode('gbk'),row[6].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.id.append(row[0])
                self.name.append(row[1])
                self.gender.append(row[2])
                self.age.append(row[3])
                self.position.append(row[4])
                self.phone.append(row[5])
                self.workplace.append(row[6])
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')
        db.close()

        for i in range(len(self.id)):
            self.tree.insert('',i,values=(self.id[i],self.name[i],self.gender[i],self.age[i],self.position[i],self.phone[i],self.workplace[i]))   #插入数据

    #返回
    def back(self):
        SubAdminPage(self.window)



#文物信息管理页面
class CulturalrelicManage():
    def __init__(self,parent_window):
        parent_window.destroy()
        self.window=tk.Tk()
        self.window.title('文物信息管理')
        self.window.geometry('820x700+400+50')

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\文物信息管理页面.png')
        image = image_file.resize((1200, 700))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1) 

        #定义中间列表
        self.columns=('编号','名称','类别','来源','年代','馆藏地')
        self.tree=ttk.Treeview(self.window, show='headings', height=10, columns=self.columns)  
        self.tree.bind('<ButtonRelease-1>', self.click_show)
        self.vbar = ttk.Scrollbar(self.window, orient=VERTICAL, command=self.tree.yview)  #竖直滚动条
        self.tree.configure(yscrollcommand=self.vbar.set)   #将滚动条与列表关联
        self.tree.column('编号', width=35, anchor='center')  # 修改列的锚点为居中对齐
        self.tree.column('名称', width=220, anchor='w')  # 修改列的锚点为左对齐
        self.tree.column('类别', width=60, anchor='w')  
        self.tree.column('来源', width=165, anchor='w')
        self.tree.column('年代', width=100, anchor='w')
        self.tree.column('馆藏地', width=175, anchor='w')
        self.tree.place(relx=0.05, rely=0.63, anchor='w')  #列表居中显示
        self.vbar.place(relx=0.02, rely=0.63, anchor='w')  #竖直滚动条居中显示
        self.InsertData()  #向表中插入数据
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.tree_sort_column(self.tree, _col, False)) 

        # 添加标签
        x_label=200
        label_id = tk.Label(self.window, text='编号:', font=('宋体', 12))
        label_id.grid(row=0, column=0, padx=x_label, pady=(10, 10))
        label_name = tk.Label(self.window, text='名称:', font=('宋体', 12))
        label_name.grid(row=1, column=0,padx=x_label, pady=(10, 10))
        label_kind= tk.Label(self.window, text='类别:', font=('宋体', 12))
        label_kind.grid(row=2, column=0, padx=x_label, pady=(10, 10))
        label_source = tk.Label(self.window, text='来源:', font=('宋体', 12))
        label_source.grid(row=3, column=0,padx=x_label, pady=(10, 10))
        label_time = tk.Label(self.window, text='年代:', font=('宋体', 12))
        label_time.grid(row=4, column=0,padx=x_label, pady=(10, 10))
        label_location = tk.Label(self.window, text='馆藏地:', font=('宋体', 12))
        label_location.grid(row=5, column=0,padx=x_label, pady=(10, 10))

        # 添加文本框
        x_textbox=260
        self.textbox_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_id.place(x=x_textbox, y=10)
        self.textbox_name = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_name.place(x=x_textbox, y=54)
        self.textbox_kind= tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_kind.place(x=x_textbox, y=96)
        self.textbox_source = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_source.place(x=x_textbox, y=140)
        self.textbox_time = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_time.place(x=x_textbox, y=184)
        self.textbox_location = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_location.place(x=x_textbox, y=225)

        # 添加按钮，新建，更新，删除
        x_button=240
        y_button=270
        button_new = tk.Button(self.window, text='新建', font=('宋体', 12), width=10, height=1,command=self.NewData)
        button_new.place(x=x_button, y=y_button)
        button_update = tk.Button(self.window, text='更新', font=('宋体', 12), width=10, height=1,command=self.UpdateData)
        button_update.place(x=x_button+120, y=y_button)
        button_delete = tk.Button(self.window, text='删除', font=('宋体', 12), width=10, height=1,command=self.DeleteData)
        button_delete.place(x=x_button+240, y=y_button)

        # 添加按钮，查询，返回
        button_search = tk.Button(self.window, text='查询', font=('宋体', 12), width=10, height=1,command=self.SearchData)
        button_search.place(x=200, y=580)
        self.textbox_search = tk.Text(self.window, width=20, height=1.4, font=('宋体', 12))
        self.textbox_search.place(x=300, y=583)
        button_back = tk.Button(self.window, text='返回', font=('宋体', 12), width=10, height=1, command=self.back)
        button_back.place(x=500, y=580)

        #添加按钮到借出与借入文物页面
        button_borrow = tk.Button(self.window, text='借出文物', font=('宋体', 20), width=15, height=1,command=lambda:BorrowPage(self.window))
        button_borrow.place(x=170, y=640)
        button_lend = tk.Button(self.window, text='借入文物', font=('宋体', 20), width=15, height=1,command=lambda:LendPage(self.window))
        button_lend.place(x=420, y=640)

        self.window.mainloop()  

    
    #新建
    def NewData(self):
        #将文本框中数据插入数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"INSERT INTO dbo.文物信息表 VALUES('{self.textbox_id.get(1.0,tk.END).strip()}','{self.textbox_name.get(1.0, tk.END).strip()}','{self.textbox_kind.get(1.0, tk.END).strip()}','{self.textbox_source.get(1.0, tk.END).strip()}','{self.textbox_time.get(1.0, tk.END).strip()}','{self.textbox_location.get(1.0, tk.END).strip()}')"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','新建成功！')
            print('新建成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'新建失败！{e}')
        self.InsertData()  # 更新列表


    #更新
    def UpdateData(self):
        #将文本框中的数据更新到数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"UPDATE dbo.文物信息表 SET 文物名称 = '{self.textbox_name.get(1.0, tk.END).strip()}', 文物类别 = '{self.textbox_kind.get(1.0, tk.END).strip()}', 文物来源 = '{self.textbox_source.get(1.0, tk.END).strip()}', 文物年代 = '{self.textbox_time.get(1.0, tk.END).strip()}', 文物馆藏地 = '{self.textbox_location.get(1.0, tk.END).strip()}' WHERE 文物编号 = '{self.textbox_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','更新成功！')
            print('更新成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'更新失败！{e}')
        self.InsertData()   #更新列表


    #删除
    def DeleteData(self):
        #将文本框中的数据从数据库中删除
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"DELETE FROM dbo.文物信息表 WHERE 文物编号='{self.textbox_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','删除成功！')
            print('删除成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'删除失败！{e}')
        self.InsertData()   #更新列表


    #查询
    def SearchData(self):
        # 清除文本框内容
        self.delete_text()
        # 根据文本框中的工作人员姓名查询数据库数据
        db = pymssql.connect('ZYH', 'sa', '123', 'Nanjing_Museum_Manage_System')
        cursor = db.cursor()
        keyword = self.textbox_search.get(1.0, tk.END).strip()
        sql = f"SELECT * FROM dbo.文物信息表 WHERE 文物名称 LIKE '%{keyword}%'"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results = [(row[0].encode('latin-1').decode('gbk'), row[1].encode('latin-1').decode('gbk'), row[2].encode('latin-1').decode('gbk'), row[3].encode('latin-1').decode('gbk'), row[4].encode('latin-1').decode('gbk'), row[5].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.textbox_id.insert(tk.END, f"{row[0]}\n")
                self.textbox_name.insert(tk.END, f"{row[1]}\n")
                self.textbox_kind.insert(tk.END, f"{row[2]}\n")
                self.textbox_source.insert(tk.END, f"{row[3]}\n")
                self.textbox_time.insert(tk.END, f"{row[4]}\n")
                self.textbox_location.insert(tk.END, f"{row[5]}\n")
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')


    #点击列表中的数据，将数据显示在文本框中
    def click_show(self,event):
        self.col=self.tree.identify_column(event.x)  #列
        self.row=self.tree.identify_row(event.y)     #行

        self.row_info=self.tree.item(self.row,'values')
        self.var_id=self.row_info[0]
        self.var_name=self.row_info[1]
        self.var_kind=self.row_info[2]
        self.var_source=self.row_info[3]
        self.var_time=self.row_info[4]
        self.var_location=self.row_info[5]

        self.delete_text()
        self.textbox_id.insert(tk.END, f"{self.var_id}\n")
        self.textbox_name.insert(tk.END, f"{self.var_name}\n")
        self.textbox_kind.insert(tk.END,f"{self.var_kind}\n")
        self.textbox_source.insert(tk.END, f"{self.var_source}\n")
        self.textbox_time.insert(tk.END, f"{self.var_time}\n")
        self.textbox_location.insert(tk.END, f"{self.var_location}\n")


    # 点击表头排序
    def tree_sort_column(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children('')]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.tree_sort_column(tree, col, not reverse))
    

    #删除文本框中的文字
    def delete_text(self):
        self.textbox_id.delete(1.0, tk.END)
        self.textbox_name.delete(1.0, tk.END)
        self.textbox_kind.delete(1.0, tk.END)
        self.textbox_source.delete(1.0, tk.END)
        self.textbox_time.delete(1.0, tk.END)
        self.textbox_location.delete(1.0, tk.END)


    #向表中插入数据
    def InsertData(self):
        self.tree.delete(*self.tree.get_children())  #清空表中数据
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.id=[]
        self.name=[]
        self.kind=[]
        self.source=[]
        self.time=[]
        self.location=[]

        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql='SELECT * FROM dbo.文物信息表'
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results=[(row[0].encode('latin-1').decode('gbk'),row[1].encode('latin-1').decode('gbk'),row[2].encode('latin-1').decode('gbk'),row[3].encode('latin-1').decode('gbk'),row[4].encode('latin-1').decode('gbk'),row[5].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.id.append(row[0])
                self.name.append(row[1])
                self.kind.append(row[2])
                self.source.append(row[3])
                self.time.append(row[4])
                self.location.append(row[5])
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')
        db.close()

        for i in range(len(self.id)):
            self.tree.insert('',i,values=(self.id[i],self.name[i],self.kind[i],self.source[i],self.time[i],self.location[i]))   #插入数据

    #返回
    def back(self):
        SubAdminPage(self.window)



#借出文物管理页面
class BorrowPage:
    def __init__(self,parent_window):
        parent_window.destroy()
        self.window=tk.Tk()
        self.window.title('借出文物管理')
        self.window.geometry('500x600+450+50')

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\借出文物界面.jfif')
        image = image_file.resize((600, 600))  
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1) 

        #定义中间列表
        self.columns=('借出编号','文物编号','借出地','借出日期','归还日期')
        self.tree=ttk.Treeview(self.window, show='headings', height=10, columns=self.columns)
        self.tree.bind('<ButtonRelease-1>', self.click_show)
        self.vbar = ttk.Scrollbar(self.window, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vbar.set)
        self.tree.column('借出编号', width=75, anchor='center')
        self.tree.column('文物编号', width=75, anchor='w')
        self.tree.column('借出地', width=100, anchor='w')
        self.tree.column('借出日期', width=100, anchor='w')
        self.tree.column('归还日期', width=100, anchor='w')
        self.tree.place(relx=0.05, rely=0.7, anchor='w')
        self.vbar.place(relx=0.01, rely=0.7, anchor='w')
        self.InsertData()
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.tree_sort_column(self.tree, _col, False))
        
        # 添加标签
        x_label=40
        label_borrow_id = tk.Label(self.window, text='借出编号:', font=('宋体', 12))
        label_borrow_id.grid(row=0, column=0, padx=x_label, pady=(10, 10))
        label_culturalrelic_id = tk.Label(self.window, text='文物编号:', font=('宋体', 12))
        label_culturalrelic_id.grid(row=1, column=0,padx=x_label, pady=(10, 10))
        label_borrow_location= tk.Label(self.window, text='借出地:', font=('宋体', 12))
        label_borrow_location.grid(row=2, column=0, padx=x_label, pady=(10, 10))
        label_borrow_date = tk.Label(self.window, text='借出日期:', font=('宋体', 12))
        label_borrow_date.grid(row=3, column=0,padx=x_label, pady=(10, 10))
        label_return_date = tk.Label(self.window, text='归还日期:', font=('宋体', 12))
        label_return_date.grid(row=4, column=0,padx=x_label, pady=(10, 10))

        # 添加文本框
        x_textbox=140
        self.textbox_borrow_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_borrow_id.place(x=x_textbox, y=10)
        self.textbox_culturalrelic_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_culturalrelic_id.place(x=x_textbox, y=54)
        self.textbox_borrow_location= tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_borrow_location.place(x=x_textbox, y=96)
        self.textbox_borrow_date = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_borrow_date.place(x=x_textbox, y=140)
        self.textbox_return_date = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_return_date.place(x=x_textbox, y=184)

        # 添加按钮，新建，更新，删除
        x_button=100
        y_button=240
        button_new = tk.Button(self.window, text='新建', font=('宋体', 12), width=10, height=1,command=self.NewData)
        button_new.place(x=x_button, y=y_button)
        button_update = tk.Button(self.window, text='更新', font=('宋体', 12), width=10, height=1,command=self.UpdateData)
        button_update.place(x=x_button+120, y=y_button)
        button_delete = tk.Button(self.window, text='删除', font=('宋体', 12), width=10, height=1,command=self.DeleteData)
        button_delete.place(x=x_button+240, y=y_button)

        # 添加按钮，查询，返回
        button_search = tk.Button(self.window, text='查询', font=('宋体', 12), width=10, height=1,command=self.SearchData)
        button_search.place(x=80, y=550)
        self.textbox_search = tk.Text(self.window, width=20, height=1.4, font=('宋体', 12))
        self.textbox_search.place(x=180, y=553)
        button_back = tk.Button(self.window, text='返回', font=('宋体', 12), width=10, height=1, command=self.back)
        button_back.place(x=380, y=550)

        self.window.mainloop()


    #向表中插入数据
    def InsertData(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.borrow_id=[]
        self.culturalrelic_id=[]
        self.borrow_location=[]
        self.borrow_date=[]
        self.return_date=[]

        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql='SELECT * FROM dbo.借出文物表'
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results=[(row[0].encode('latin-1').decode('gbk'),row[1].encode('latin-1').decode('gbk'),row[2].encode('latin-1').decode('gbk'),row[3].encode('latin-1').decode('gbk'),row[4].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.borrow_id.append(row[0])
                self.culturalrelic_id.append(row[1])
                self.borrow_location.append(row[2])
                self.borrow_date.append(row[3])
                self.return_date.append(row[4])
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')
        db.close()

        for i in range(len(self.borrow_id)):
            self.tree.insert('',i,values=(self.borrow_id[i],self.culturalrelic_id[i],self.borrow_location[i],self.borrow_date[i],self.return_date[i]))


    def click_show(self,event):
        self.col=self.tree.identify_column(event.x)
        self.row=self.tree.identify_row(event.y)

        self.row_info=self.tree.item(self.row,'values')
        self.var_borrow_id=self.row_info[0]
        self.var_culturalrelic_id=self.row_info[1]
        self.var_borrow_location=self.row_info[2]
        self.var_borrow_date=self.row_info[3]
        self.var_return_date=self.row_info[4]

        self.delete_text()
        self.textbox_borrow_id.insert(tk.END, f"{self.var_borrow_id}\n")
        self.textbox_culturalrelic_id.insert(tk.END, f"{self.var_culturalrelic_id}\n")
        self.textbox_borrow_location.insert(tk.END,f"{self.var_borrow_location}\n")
        self.textbox_borrow_date.insert(tk.END, f"{self.var_borrow_date}\n")
        self.textbox_return_date.insert(tk.END, f"{self.var_return_date}\n")


    def delete_text(self):
        self.textbox_borrow_id.delete(1.0, tk.END)
        self.textbox_culturalrelic_id.delete(1.0, tk.END)
        self.textbox_borrow_location.delete(1.0, tk.END)
        self.textbox_borrow_date.delete(1.0, tk.END)
        self.textbox_return_date.delete(1.0, tk.END)


    def tree_sort_column(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children('')]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.tree_sort_column(tree, col, not reverse))


    def back(self): 
        CulturalrelicManage(self.window)


    def NewData(self):
        #将文本框中数据插入数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"INSERT INTO dbo.借出文物表 VALUES('{self.textbox_borrow_id.get(1.0,tk.END).strip()}','{self.textbox_culturalrelic_id.get(1.0, tk.END).strip()}','{self.textbox_borrow_location.get(1.0, tk.END).strip()}','{self.textbox_borrow_date.get(1.0, tk.END).strip()}','{self.textbox_return_date.get(1.0, tk.END).strip()}')"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','新建成功！')
            print('新建成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'新建失败！{e}')
        self.InsertData()


    def UpdateData(self):
        #将文本框中的数据更新到数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"UPDATE dbo.借出文物表 SET 文物编号 = '{self.textbox_culturalrelic_id.get(1.0, tk.END).strip()}', 借出地 = '{self.textbox_borrow_location.get(1.0, tk.END).strip()}', 借出日期 = '{self.textbox_borrow_date.get(1.0, tk.END).strip()}', 归还日期 = '{self.textbox_return_date.get(1.0, tk.END).strip()}' WHERE 借出编号 = '{self.textbox_borrow_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','更新成功！')
            print('更新成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'更新失败！{e}')
        self.InsertData()


    def DeleteData(self):
        # 将文本框中的数据从数据库中删除
        db = pymssql.connect('ZYH', 'sa', '123', 'Nanjing_Museum_Manage_System')
        cursor = db.cursor()
        borrow_id = self.textbox_borrow_id.get(1.0, tk.END).strip()
        sql = f"DELETE FROM dbo.借出文物表 WHERE 借出编号='{borrow_id}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示', '删除成功！')
            print('删除成功！')
        except Exception as e:
            messagebox.showinfo('提示', f'删除失败！{e}')
        self.InsertData()


    def SearchData(self):
        self.delete_text()
        db = pymssql.connect('ZYH', 'sa', '123', 'Nanjing_Museum_Manage_System')
        cursor = db.cursor()
        keyword = self.textbox_search.get(1.0, tk.END).strip()
        sql = f"SELECT * FROM dbo.借出文物表 WHERE 借出编号 LIKE '%{keyword}%'"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results = [(row[0].encode('latin-1').decode('gbk'), row[1].encode('latin-1').decode('gbk'), row[2].encode('latin-1').decode('gbk'), row[3].encode('latin-1').decode('gbk'), row[4].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.textbox_borrow_id.insert(tk.END, f"{row[0]}\n")
                self.textbox_culturalrelic_id.insert(tk.END, f"{row[1]}\n")
                self.textbox_borrow_location.insert(tk.END, f"{row[2]}\n")
                self.textbox_borrow_date.insert(tk.END, f"{row[3]}\n")
                self.textbox_return_date.insert(tk.END, f"{row[4]}\n")
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')


#借入文物管理页面
class LendPage:
    def __init__(self,parent_window):
        parent_window.destroy()
        self.window=tk.Tk()
        self.window.title('借入文物管理')
        self.window.geometry('500x600+450+50')

        image_file = Image.open('D:\\PycharmProjects\\Nanjing_Museum_Manage_System\\images\\借入文物界面.jpg')
        image = image_file.resize((1000, 700))
        image = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.window, image=image)
        image_label.place(x=0, y=0, relwidth=1, relheight=1)

        #定义中间列表
        self.columns=('借入编号','文物编号','借入地','借入日期','归还日期')
        self.tree=ttk.Treeview(self.window, show='headings', height=10, columns=self.columns)
        self.tree.bind('<ButtonRelease-1>', self.click_show)
        self.vbar = ttk.Scrollbar(self.window, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vbar.set)
        self.tree.column('借入编号', width=75, anchor='center')
        self.tree.column('文物编号', width=75, anchor='w')
        self.tree.column('借入地', width=100, anchor='w')
        self.tree.column('借入日期', width=100, anchor='w')
        self.tree.column('归还日期', width=100, anchor='w')
        self.tree.place(relx=0.05, rely=0.7, anchor='w')
        self.vbar.place(relx=0.01, rely=0.7, anchor='w')
        self.InsertData()
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.tree_sort_column(self.tree, _col, False))
        
        # 添加标签
        x_label=40
        label_lend_id = tk.Label(self.window, text='借入编号:', font=('宋体', 12))
        label_lend_id.grid(row=0, column=0, padx=x_label, pady=(10, 10))
        label_culturalrelic_id = tk.Label(self.window, text='文物编号:', font=('宋体', 12))
        label_culturalrelic_id.grid(row=1, column=0,padx=x_label, pady=(10, 10))
        label_lend_location= tk.Label(self.window, text='借入地:', font=('宋体', 12))
        label_lend_location.grid(row=2, column=0, padx=x_label, pady=(10, 10))
        label_lend_date = tk.Label(self.window, text='借入日期:', font=('宋体', 12))
        label_lend_date.grid(row=3, column=0,padx=x_label, pady=(10, 10))
        label_return_date = tk.Label(self.window, text='归还日期:', font=('宋体', 12))
        label_return_date.grid(row=4, column=0,padx=x_label, pady=(10, 10))
        
        # 添加文本框
        x_textbox=140
        self.textbox_lend_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_lend_id.place(x=x_textbox, y=10)
        self.textbox_culturalrelic_id = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_culturalrelic_id.place(x=x_textbox, y=54)
        self.textbox_lend_location= tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_lend_location.place(x=x_textbox, y=96)
        self.textbox_lend_date = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_lend_date.place(x=x_textbox, y=140)
        self.textbox_return_date = tk.Text(self.window, width=40, height=1.4, font=('宋体', 12))
        self.textbox_return_date.place(x=x_textbox, y=184)

        # 添加按钮，新建，更新，删除
        x_button=100
        y_button=240
        button_new = tk.Button(self.window, text='新建', font=('宋体', 12), width=10, height=1,command=self.NewData)
        button_new.place(x=x_button, y=y_button)
        button_update = tk.Button(self.window, text='更新', font=('宋体', 12), width=10, height=1,command=self.UpdateData)
        button_update.place(x=x_button+120, y=y_button)
        button_delete = tk.Button(self.window, text='删除', font=('宋体', 12), width=10, height=1,command=self.DeleteData)
        button_delete.place(x=x_button+240, y=y_button)

        # 添加按钮，查询，返回
        button_search = tk.Button(self.window, text='查询', font=('宋体', 12), width=10, height=1,command=self.SearchData)
        button_search.place(x=80, y=550)
        self.textbox_search = tk.Text(self.window, width=20, height=1.4, font=('宋体', 12))
        self.textbox_search.place(x=180, y=553)
        button_back = tk.Button(self.window, text='返回', font=('宋体', 12), width=10, height=1, command=self.back)
        button_back.place(x=380, y=550)

        self.window.mainloop()


    #向表中插入数据
    def InsertData(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.lend_id=[]
        self.culturalrelic_id=[]
        self.lend_location=[]
        self.lend_date=[]
        self.return_date=[]

        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql='SELECT * FROM dbo.借入文物表'
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results=[(row[0].encode('latin-1').decode('gbk'),row[1].encode('latin-1').decode('gbk'),row[2].encode('latin-1').decode('gbk'),row[3].encode('latin-1').decode('gbk'),row[4].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.lend_id.append(row[0])
                self.culturalrelic_id.append(row[1])
                self.lend_location.append(row[2])
                self.lend_date.append(row[3])
                self.return_date.append(row[4])
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')
        db.close()

        for i in range(len(self.lend_id)):
            self.tree.insert('',i,values=(self.lend_id[i],self.culturalrelic_id[i],self.lend_location[i],self.lend_date[i],self.return_date[i]))
    

    def click_show(self,event):
        self.col=self.tree.identify_column(event.x)
        self.row=self.tree.identify_row(event.y)

        self.row_info=self.tree.item(self.row,'values')
        self.var_lend_id=self.row_info[0]
        self.var_culturalrelic_id=self.row_info[1]
        self.var_lend_location=self.row_info[2]
        self.var_lend_date=self.row_info[3]
        self.var_return_date=self.row_info[4]

        self.delete_text()
        self.textbox_lend_id.insert(tk.END, f"{self.var_lend_id}\n")
        self.textbox_culturalrelic_id.insert(tk.END, f"{self.var_culturalrelic_id}\n")
        self.textbox_lend_location.insert(tk.END,f"{self.var_lend_location}\n")
        self.textbox_lend_date.insert(tk.END, f"{self.var_lend_date}\n")
        self.textbox_return_date.insert(tk.END, f"{self.var_return_date}\n")


    def delete_text(self):
        self.textbox_lend_id.delete(1.0, tk.END)
        self.textbox_culturalrelic_id.delete(1.0, tk.END)
        self.textbox_lend_location.delete(1.0, tk.END)
        self.textbox_lend_date.delete(1.0, tk.END)
        self.textbox_return_date.delete(1.0, tk.END)


    def tree_sort_column(self, tree, col, reverse):
        items = [(tree.set(k, col), k) for k in tree.get_children('')]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.tree_sort_column(tree, col, not reverse))


    def back(self):
        CulturalrelicManage(self.window)


    def NewData(self):
        #将文本框中数据插入数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"INSERT INTO dbo.借入文物表 VALUES('{self.textbox_lend_id.get(1.0,tk.END).strip()}','{self.textbox_culturalrelic_id.get(1.0, tk.END).strip()}','{self.textbox_lend_location.get(1.0, tk.END).strip()}','{self.textbox_lend_date.get(1.0, tk.END).strip()}','{self.textbox_return_date.get(1.0, tk.END).strip()}')"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','新建成功！')
            print('新建成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'新建失败！{e}')
        self.InsertData()


    def UpdateData(self):
        #将文本框中的数据更新到数据库中
        db=pymssql.connect('ZYH','sa','123','Nanjing_Museum_Manage_System')
        cursor=db.cursor()
        sql=f"UPDATE dbo.借入文物表 SET 文物编号 = '{self.textbox_culturalrelic_id.get(1.0, tk.END).strip()}', 借入地 = '{self.textbox_lend_location.get(1.0, tk.END).strip()}', 借入日期 = '{self.textbox_lend_date.get(1.0, tk.END).strip()}', 归还日期 = '{self.textbox_return_date.get(1.0, tk.END).strip()}' WHERE 借入编号 = '{self.textbox_lend_id.get(1.0, tk.END).strip()}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示','更新成功！')
            print('更新成功！')
        except Exception as e:
            messagebox.showinfo('提示',f'更新失败！{e}')
        self.InsertData()


    def DeleteData(self):
        # 将文本框中的数据从数据库中删除
        db = pymssql.connect('ZYH', 'sa', '123', 'Nanjing_Museum_Manage_System')
        cursor = db.cursor()
        lend_id = self.textbox_lend_id.get(1.0, tk.END).strip()
        sql = f"DELETE FROM dbo.借入文物表 WHERE 借入编号='{lend_id}'"
        try:
            cursor.execute(sql)
            db.commit()
            messagebox.showinfo('提示', '删除成功！')
            print('删除成功！')
        except Exception as e:
            messagebox.showinfo('提示', f'删除失败！{e}')
        self.InsertData()


    def SearchData(self):
        self.delete_text()
        db = pymssql.connect('ZYH', 'sa', '123', 'Nanjing_Museum_Manage_System')
        cursor = db.cursor()
        keyword = self.textbox_search.get(1.0, tk.END).strip()
        sql = f"SELECT * FROM dbo.借入文物表 WHERE 文物编号 LIKE '%{keyword}%'"
        try:
            cursor.execute(sql)
            results = cursor.fetchall()
            decoded_results = [(row[0].encode('latin-1').decode('gbk'), row[1].encode('latin-1').decode('gbk'), row[2].encode('latin-1').decode('gbk'), row[3].encode('latin-1').decode('gbk'), row[4].encode('latin-1').decode('gbk')) for row in results] #解决中文乱码问题
            for row in decoded_results:
                self.textbox_lend_id.insert(tk.END, f"{row[0]}\n")
                self.textbox_culturalrelic_id.insert(tk.END, f"{row[1]}\n")
                self.textbox_lend_location.insert(tk.END, f"{row[2]}\n")
                self.textbox_lend_date.insert(tk.END, f"{row[3]}\n")
                self.textbox_return_date.insert(tk.END, f"{row[4]}\n")
        except Exception as e:
            messagebox.showinfo('提示',f'查询失败！{e}')



if __name__=='__main__':
    window=tk.Tk()         
    StartPage(window)