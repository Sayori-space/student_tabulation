from tkinter import *
from tkinter.messagebox import *
from tkinter import filedialog, messagebox
from manger_cotraller import *
from tkinter import ttk
from random import randint
from PIL import ImageTk

class Application(Frame):
    """一个经典的GUI程序的类的写法"""

    def __init__(self, master=None):
        super().__init__(master)        # super()代表的是父类的定义，而不是父类对象
        self.master = master
        self.pack()
        self.createWidget()
        self.target_excel = ""
        self.manager = student_model_contraller(self.target_excel)
        self.tem_list = []
        self.cbx_list = [[[], [], [], [], []],
                        [[], [], [], [], []],
                        [[], [], [], [], []],
                        [[], [], [], [], []]]

    def createWidget(self):
        """创建组件"""
        self.fubu_value = IntVar()
        self.singal_value = IntVar()
        self.double_value = IntVar()
        self.formal_value = IntVar()
        self.back_value = IntVar()

        # 图片部分
        global photo_baimao
        photo_baimao = ImageTk.PhotoImage(file="images/black.gif")
        self.label_baimao = Label(self,image = photo_baimao)
        self.label_baimao.place(relx=0.5, rely=0.5)


        menubar = Menu(root)
        menu_file = Menu(menubar)
        menu_choose = Menu(menubar)
        menubar.add_cascade(label="文件(F)", menu=menu_file)
        menubar.add_cascade(label="筛选(E)", menu=menu_choose)


        menu_file.add_command(label="导入Excel文件",command=self.import_excel)
        menu_file.add_separator()  # 添加分割线
        menu_file.add_command(label="开始生成",command=self.add_name)
        menu_file.add_separator()  # 添加分割线
        menu_file.add_command(label="导出Excel文件",command=self.export_file)

        menu_choose.add_checkbutton(label="去除副部",variable=self.fubu_value,
                                    onvalue=1, offvalue=0)
        menu_choose.add_separator()
        menu_choose.add_checkbutton(label="去除单周",variable=self.singal_value,
                                    onvalue=1, offvalue=0)
        menu_choose.add_separator()
        menu_choose.add_checkbutton(label="去除双周",variable=self.double_value,
                                    onvalue=1,offvalue=0)
        menu_choose.add_separator()
        menu_choose.add_checkbutton(label="去除前八周",variable=self.formal_value,
                                    onvalue=1,offvalue=0)
        menu_choose.add_separator()
        menu_choose.add_checkbutton(label="去除后八周",variable=self.back_value,
                                    onvalue=1,offvalue=0)

        root['menu'] = menubar

        # 星期表头部分
        week_list = ["一","二","三","四","五"]
        for i in week_list:
            week_index = week_list.index(i)
            self.label = Label(root,text="星期" + i)
            self.label.place(relx=0.15+week_index*0.1, rely=0.2)

        # 课节表头部分
        class_list = ['一',"二","三","四"]
        for i in class_list:
            class_index = class_list.index(i)
            self.label = Label(root,text="第"+i+"节")
            self.label.place(relx=0.05,rely = 0.3+class_index*0.1)





    def add_name(self):
        if self.target_excel:
            self.manager.add_init_data()
            self.manager.sperate_name()
            if self.fubu_value.get() == 1:
                self.manager.select_student(3)
            if self.singal_value.get() == 1:
                self.manager.select_student(1)
            if self.double_value.get() == 1:
                self.manager.select_student(2)
            if self.formal_value.get() == 1:
                self.manager.select_student(4)
            if self.back_value.get() == 1:
                self.manager.select_student(5)
            target_list = self.manager.get_students_list()
            for i in range(len(target_list)):
                for j in range(len(target_list[i])):
                    if target_list[i][j]:
                        name = target_list[i][j][random.randint(0,len(target_list[i][j])-1)]
                        self.cbx = ttk.Combobox(root,width=10)
                        self.cbx['values'] = target_list[i][j]
                        index = target_list[i][j].index(name)
                        self.cbx.current(index)
                        self.cbx.bind("<<ComboboxSelected>>", self.cbx_select)
                        self.cbx.place(relx=0.15+j*0.1,rely=0.3+i*0.1)
                        self.cbx_list[i][j].append(self.cbx)
                        target_list[i][j].clear()
                        target_list[i][j].append(name)
                    else:
                        self.cbx = ttk.Combobox(root,width=10)
                        self.cbx.bind("<<ComboboxSelected>>", self.cbx_select)
                        self.cbx.place(relx=0.15 + j * 0.1, rely=0.3 + i * 0.1)
                        self.cbx_list[i][j].append(self.cbx)
                        continue
        else:
            messagebox.showerror("错误", "请先导入Excel文件")




    def import_excel(self):
        self.target_excel = filedialog.askopenfilename(title="上传Excel文件", initialdir="D:")
        self.manager = student_model_contraller(self.target_excel)
        show = Label(root, width=40, height=3, bg="grey")
        show.place(relx=0, rely=0)
        show["text"] = self.target_excel

    def cbx_select(self, *args):
        self.manager = student_model_contraller(self.target_excel)
        final_list = self.manager.get_students_list()
        for i in range(len(self.cbx_list)):
            for j in range(len(self.cbx_list[i])):
                final_list[i][j].clear()
                name = self.cbx_list[i][j][0].get()
                final_list[i][j].append(name)

    def export_file(self):
        print(self.manager.get_students_list())
        self.manager.export_excel()

    def append_name(self,current_list:list,target_list:list):
        for i in range(len(current_list)):
            for j in range(len(current_list[i])):
                if current_list[i][j]:
                    target_list[i][j].append(current_list[i][j])
                else:
                    continue




if __name__ == '__main__':
    root = Tk()
    root.geometry("800x500+0+0")
    root.title("半糖随机排表系统")
    app = Application(master=root)
    root.mainloop()