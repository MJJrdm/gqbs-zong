import datetime
from work_time_class import WorkTime


# a = datetime.time(8, 0, 0)
delta = datetime.timedelta(minutes=1440)
# result_time = ( datetime.timedelta(hours = a.hour) + delta)
# print(result_time)  


worktime_file_path = "E:/富士康/工作/排产项目/广汽宝钢/排产/数据/排班/排班表.xlsx"
worktime = WorkTime.read_data(worktime_file_path, '横切')
a = worktime.worktime_for_day_and_night[list(worktime.worktime_for_day_and_night.keys())[0]]["早班"][0]
print(delta + datetime.timedelta(hours = a.hour))
