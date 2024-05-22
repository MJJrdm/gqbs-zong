import pandas as pd
import datetime


class PlanTime:
    def __init__(self, plan_data, worktime):
        self.plan_data = plan_data
        self.worktime = worktime
    
    ##### 得到唯一的不同组以及对应方案的生产预计花费时间
    def get_plan_group_and_production_time(self):
        unique_groups = [] 
        groups_and_production_time = {}
        for _, row in self.plan_data.sorted_df.iterrows():
            if row[('分组', '分组子列')] not in unique_groups:
                unique_groups.append(row[('分组', '分组子列')])
                groups_and_production_time[row[('分组', '分组子列')]] = row[('生产时间', '生产时间子列')]
        return unique_groups, groups_and_production_time
    

    def compute_plan_start_and_end_time(self):
        start_working_time = self.worktime.worktime_for_day_and_night[list(self.worktime.worktime_for_day_and_night.keys())[0]]["早班"][0]
        start_time_list = []
        end_time_list = []
        (unique_groups, groups_and_production_time) = self.get_plan_group_and_production_time()
        # unique_groups -> [3, 5, 7, 1, 6, 11]
        # groups_and_production_time -> {3: 30, 5: 0, 7: 0, 1: 21, 6: 24, 11: 21}
        for index, group in enumerate(unique_groups):
            duration = datetime.timedelta(minutes = groups_and_production_time[group])
            if index == 0:
                start = datetime.timedelta(hours = start_working_time.hour)
                end = start + duration
            else:
                start = end_time_list[index - 1]
                end = start + duration

            start_time_list.append(start)
            end_time_list.append(end)

            groups_and_production_time[group] = str(start) + '-' + str(end)
            
            if index == 10:
                break

        return start_time_list, end_time_list
        # start_time_list -> [datetime.timedelta(seconds=28800), datetime.timedelta(seconds=30600), ...]
        # end_time_list -->  [datetime.timedelta(seconds=30600), datetime.timedelta(seconds=30600), ...]


# # a = datetime.time(8, 0, 0)
# delta = datetime.timedelta(minutes=1440)
# # result_time = ( datetime.timedelta(hours = a.hour) + delta)
# # print(result_time)  


# worktime_file_path = "E:/富士康/工作/排产项目/广汽宝钢/排产/数据/排班/排班表.xlsx"
# worktime = WorkTime.read_data(worktime_file_path, '横切')
# a = worktime.worktime_for_day_and_night[list(worktime.worktime_for_day_and_night.keys())[0]]["早班"][0]
# print(delta + datetime.timedelta(hours = a.hour))