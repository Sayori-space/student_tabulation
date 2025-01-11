import xlrd
import xlwt
import random




class student_model_contraller:
    def __init__(self, excel_file :str):
        self.excel_file = excel_file
        self.__students_list = [[[], [], [], [], []],
                                [[], [], [], [], []],
                                [[], [], [], [], []],
                                [[], [], [], [], []]]
    def get_students_list(self):
        return self.__students_list

    def add_init_data(self):
        '''
        将原始数据添加到student-list中
        :return:
        '''
        book = xlrd.open_workbook(self.excel_file)
        sheets = book.sheets()
        for sheet in sheets[1:]:
            current_sheet = book.sheet_by_name(sheet.name)
            for x in range(2, current_sheet.nrows):
                for y in range(1, current_sheet.ncols):
                    self.__students_list[x-2][y-1].append(current_sheet.cell_value(x, y))

    def sperate_name(self):
        '''
        对添加完原视数据的student_list进行处理，去掉数据中的换行
        :return:
        '''
        temp_list = []
        for i in range(len(self.__students_list)):
            for j in range(len(self.__students_list[i])):
                for name in self.__students_list[i][j]:
                    name_list = name.splitlines()
                    temp_list.extend(name_list)
                else:
                    self.__students_list[i][j].clear()
                    self.__students_list[i][j].extend(temp_list)
                    temp_list.clear()

    def select_student(self,select_mode=0):
        '''
        筛选出不符合条件的数据
        :param select_mode:
        :return:
        '''
        if select_mode == 0:
            return
        elif select_mode == 1:# 单周
            self.sperate("（单","(单")
        elif select_mode == 2:# 双周
            self.sperate("（双","(双")
        elif select_mode == 3:# 副部
            pass
        elif select_mode == 4:# 前八周
            self.sperate("1-")
        elif select_mode == 5:# 后八周
            self.sperate("9-")

    def sperate(self, *key: str):
        for i in range(len(self.__students_list)):
            for j in range(len(self.__students_list[i])):
                for name in reversed(self.__students_list[i][j]):
                    for key in key:
                        if key in name:
                            self.__students_list[i][j].remove(name)

    def random_choose(self,sort_list: list):
        '''
        从每一个位置中随机抽取一个学生
        :param sort_list: 
        :return: 
        '''
        sort_list = self.sort_list_lenth(self.__students_list)
        repeated_list = []
        for i in sort_list:
            if self.__students_list[i[0]][i[1]]:
                random_name_num = random.randint(0,len(self.__students_list[i[0]][i[1]]) - 1)
                random_name = self.__students_list[i[0]][i[1]][random_name_num]
                repeated_list.append(random_name)
                if random_name not in repeated_list:
                    self.__students_list[i[0]][i[1]].clear()
                    self.__students_list[i[0]][i[1]].append(random_name)
                elif random_name in repeated_list:
                    if len(self.__students_list[i[0]][i[1]]) > 1:
                        set_repeat = set(repeated_list)
                        set_student_list = set(self.__students_list[i[0]][i[1]])
                        set_target = set_student_list - set_repeat
                        new_random_name = list(set_target)[random.randint(0,len(set_target)-1)]
                        self.__students_list[i[0]][i[1]].clear()
                        self.__students_list[i[0]][i[1]].append(new_random_name)
                    else:
                        continue
            else:
                continue

    def quick_sort(self, arr):
        if len(arr) <= 1:
            return arr
        pivot = arr[len(arr) // 2]
        left = [x for x in arr if x < pivot]
        middle = [x for x in arr if x == pivot]
        right = [x for x in arr if x > pivot]
        return self.quick_sort(left) + middle + self.quick_sort(right)

    def sort_list_lenth(self,lst: list) -> list:
        block_list = []
        key_list = []
        x_list = []
        y_list = []
        y = 1
        for i in range(len(lst)):
            for j in range(len(lst[i])):
                block_list.append(len(lst[i][j]))
                key_list.append([len(lst[i][j]), (i, j)])
        num_list = self.quick_sort(block_list)
        for i in range(len(num_list)):
            for j in range(len(num_list)):
                if num_list[i] == key_list[j][0] and j not in y_list:
                    x_list.append(key_list[j][1])
                    y_list.append(j)
        return x_list

    def export_excel(self, save_path):
        week_list = ["一","二","三","四","五"]
        day_list = ["一","二","三","四"]
        wb = xlwt.Workbook()
        ws = wb.add_sheet('学生会值班表')
        for i in week_list:
            week_index = week_list.index(i)
            ws.write(2, week_index+1, "星期"+i)

        for i in day_list:
            day_index = day_list.index(i)
            ws.write(day_index+3, 0, "第"+i+"节")

        for i in range(len(self.__students_list)):
            for j in range(len(self.__students_list[i])):
                ws.write(i+3, j+1, self.__students_list[i][j])

        ws.write_merge(0, 1, 0, 5, '学生值班表')
        wb.save(save_path)






















