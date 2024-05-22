from data_class import PlanData
from processing_class import ProcessingPlan


if __name__ == '__main__':
    file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/测试数据/0423排刀算法输出结果（第一批数据）.xlsx'
    # file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/测试数据/0423排刀算法输出结果（第二批数据）.xlsx'
    data = PlanData.read_data(file_path)
    ProcessingPlan.add_production_time(data)
    ProcessingPlan.adjust_time(data)
    data.combine_cells()
    data.merge_column_names()
    total_time = ProcessingPlan.compute_total_time(data)
    print(f'总生产时间: {total_time}')
    output_file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/输出结果/20240423_1_2.xlsx'
    # output_file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/输出结果/20240423_2_2.xlsx'
    data.output(output_file_path)
    
