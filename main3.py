import pandas as pd
from helper import *
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows



def main():
    file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/测试数据/03-14算法输出结果（第一批数据）.xlsx'
    # file_path = 'E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/测试数据/03-23算法输出结果（第二批数据）.xlsx'
    df = pd.read_excel(file_path, header = [0, 1])
    
    # 添加几列后续需要用到的信息列
    fill_columns(df)

    ##### 进行对dataframe的排序
    sorted_df = df.sort_values(by = [('新纳期', '新纳期子列'), # 纳期为第一优先级
                                     ('刀数', '刀数'),  # 刀数多的排前面进行生产
                                     ('加工信息', '加工方式'), 
                                     ('分组', '分组子列')], 
                                     ascending=[True, False, True, True]).reset_index(drop = True)
    # 调整计划生产时间
    adjust_time(sorted_df)

    col_names_dict = col_names_to_dict(sorted_df.columns)
    wb, ws = to_wb_ws(sorted_df)

    ##### 合并单元格
    combine_cells(sorted_df, ws, col_names_dict)
    merge_column_names(ws, col_names_dict)

    # 计算总生产时间
    total_time = compute_total_time(col_names_dict, ws)
    print(f'总生产时间: {total_time}')

    
    # wb.save('E:/富士康/工作/排产项目/广汽宝钢/排产/纵切线/代码/输出结果/20240402_output1.xlsx')


if __name__ == '__main__':
    main()
