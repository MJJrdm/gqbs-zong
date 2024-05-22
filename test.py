from work_time_class import WorkTime
from data_class import PlanData
from processing_class import ProcessingPlan
from plan_time import PlanTime
import pandas as pd


plan_file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/测试数据/0423排刀算法输出结果（第一批数据）.xlsx'
plan_data = PlanData.read_data(plan_file_path)
ProcessingPlan.add_production_time(plan_data)
ProcessingPlan.adjust_change_knife_time(plan_data)
worktime_file_path = "E:/富士康/工作/排产项目/广汽宝钢/排产/数据/排班/排班表.xlsx"
worktime = WorkTime.read_data(worktime_file_path, '横切')
plan_time = PlanTime(plan_data, worktime)
unique_groups, groups_and_production_time = plan_time.get_plan_group_and_production_time()
print(unique_groups)

start_time_list, end_time_list = plan_time.compute_plan_start_and_end_time()

for i, j in zip(start_time_list, end_time_list):
    print(i, '-', j)