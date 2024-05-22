import pandas as pd
from datetime import time

class WorkTime:
    def __init__(self, df, line):
        self.df = df
        self.line = line
        # 抽出日期列(日期形式需要转换)
        self.raw_dates_column = self.df.columns[3:]

        self.return_date_column()
        self.return_shifts_column()
        self.extract_work_time()
        self.return_worktime_info_dict()
        self.return_worktime_for_day_and_night()

    @classmethod
    def read_data(cls, file_path, line):
        df = pd.read_excel(file_path, header = [0, 1])
        instance = WorkTime(df, line)
        return instance


    ##### 把用数字表示的日期转换成可理解的日期文本形式
    def change_to_date(self, num):
        return pd.to_datetime(num, origin = '1899-12-30', unit = 'D').date()


    ##### 返回日期列表，存储排班表上每一天的日期跟对应的星期几
    def return_date_column(self):
        dates_column = []
        for i in self.raw_dates_column:
            dates_column.append((self.change_to_date(i[0]), i[1]))
            # print(i[0]) -> 45350 / 45353 / 45356
            # print(i[1]) -> '三' / '四' / '日'
        self.dates_column = dates_column
        # print(self.dates_column) -> [(datetime.date(2024, 2, 28), '三'), (datetime.date(2024, 2, 29), '四'), ....]
    

    ##### 班别列表，存储产线上所有的班别
    def return_shifts_column(self):
        shifts_column = []
        for index, row in self.df.iterrows():
            if row[('日期\n机组', 'Unnamed: 1_level_1')] == self.line:
                shifts_column.append(row[('日期\n机组', '班组')])
                for _, next_row in self.df.iloc[index+1:].iterrows():
                    if pd.isna(next_row[('日期\n机组', 'Unnamed: 1_level_1')]):
                        shifts_column.append(next_row['日期\n机组', '班组'])
                    else:
                        break
                break
        self.shifts_column = shifts_column
        # print(f'shifts_column: {shifts_column}') -> ['甲班', '乙班']


    def extract_work_time(self):
        line_worktime = {}
        for index, row in self.df.iterrows():
            if row[('日期\n机组', 'Unnamed: 1_level_1')] == self.line:
                line_worktime[row[('日期\n机组', '班组')]] = []
                for col in self.raw_dates_column:
                    line_worktime[row[('日期\n机组', '班组')]].append(row[col])
                for _, next_row in self.df.iloc[index + 1:].iterrows():
                    if pd.isna(next_row[('日期\n机组', 'Unnamed: 1_level_1')]):
                        line_worktime[next_row['日期\n机组', '班组']] = []
                        for col in self.raw_dates_column:
                            line_worktime[next_row[('日期\n机组', '班组')]].append(next_row[col])
                    else:
                        break
                break
        self.line_worktime = line_worktime
        # print(line_worktime) -> {
        #                           '甲班': ['08:00-18:00', '08:00-18:00', '08:00-18:00', 
        #                                    '08:00-18:00', nan, '18:30-04:30', '18:30-04:30'], 
        # 
        #                           '乙班': ['18:30-04:30', '18:30-04:30', '18:30-04:30', '18:30-04:30', 
        #                                    nan, '08:00-18:00', '08:00-18:00']
        #                         }
    

    def return_worktime_info_dict(self):
        worktime_info_dict = {date: {shift: None for shift in self.shifts_column} for date in self.dates_column}
        for index, date in enumerate(worktime_info_dict.keys()):
            for shift in worktime_info_dict[date]:
                if not pd.isna(self.line_worktime[shift][index]):
                    time_list = self.line_worktime[shift][index].split('-')
                    worktime_info_dict[date][shift] = [self.time_from_string(i) for i in time_list]
                else:
                    worktime_info_dict[date][shift] = []
        self.worktime_info_dict = worktime_info_dict


    def time_from_string(self, time_str):
        hours, minutes = map(int, time_str.split(':'))
        return time(hours, minutes)
    

    def return_worktime_for_day_and_night(self):
        day_night = ['早班', '晚班']
        worktime_for_day_and_night = {date: {shift: [] for shift in day_night} for date in self.dates_column}
        for key, item in self.worktime_info_dict.items():
            time = [j[0] for j in item.values() if j]
            if time:
                minimum_time = min([j[0] for j in item.values() if j])
            
            for i, j in item.items():
                if j:
                    if j[0] == minimum_time:
                        worktime_for_day_and_night[key]['早班'] = item[i]
                    else:
                        worktime_for_day_and_night[key]['晚班'] = item[i]


        self.worktime_for_day_and_night = worktime_for_day_and_night


    