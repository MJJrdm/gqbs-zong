import math
from data_class import PlanData

class ProcessingPlan:

    @classmethod
    ##### 根据分条数返回加工效率
    def return_efficiency(cls, num, type = ''):
        assert num >= 0, "错误：分条数小于0"
        if type == '外板' or type == '精整分流':
            return 6
        else:
            if num <= 3:
                return 15
            elif num <= 5:
                return 12
            else:
                return 9.6
        
        
    @classmethod
    ##### 调整生产时间，如果相邻的计划不属于同一个方案但裁切方式相同，那么靠后的计划时间减去3分钟的换刀时间
    def adjust_time(cls, data):
        for index, row in data.sorted_df.iterrows():
            if index == 0:
                previous_row_group = row[('分组', '分组子列')]
                previous_row_time = row[('生产时间', '生产时间子列')]
                previous_row_cut = row[('加工信息', '加工方式')]
            else:
                if row[('分组', '分组子列')] == previous_row_group:
                    data.sorted_df.at[index, ('生产时间', '生产时间子列')] = previous_row_time
                else:
                    if row[('加工信息', '加工方式')] == previous_row_cut and row[('加工信息', '加工方式')] != '/':
                        data.sorted_df.at[index, ('生产时间', '生产时间子列')] = row[('生产时间', '生产时间子列')] - 3

            
            previous_row_group = row[('分组', '分组子列')] 
            previous_row_time = row[('生产时间', '生产时间子列')]
            previous_row_cut = row[('加工信息', '加工方式')]


    @classmethod
    ##### 除了生产准备时间，加上实际生产时间
    def add_production_time(cls, data):
        skip = []
        total_weight = 0
        for index, row in data.sorted_df.iterrows():
            if index in skip:
                continue
            cut_num = row[('方案总刀数', '方案总刀数')]
            if cut_num == 0:
                continue
            current_row_group = row[('分组', '分组子列')]
            total_weight += row[('加工信息', '净重(kg)')]
            for next_index, next_row in data.sorted_df.iloc[index+1:].iterrows():
                if next_row[('分组', '分组子列')] == current_row_group:
                    total_weight += next_row[('加工信息', '净重(kg)')]
                    skip.append(next_index)
                else:
                    break
            efficiency = ProcessingPlan.return_efficiency(cut_num)
            production_time = total_weight / 100 / efficiency * 6
            data.sorted_df.at[index, ('生产时间', '生产时间子列')] += production_time
            data.sorted_df.at[index, ('生产时间', '生产时间子列')] = math.ceil(data.sorted_df.at[index, ('生产时间', '生产时间子列')])
            total_weight = 0


    @classmethod
    ##### 计算总的生产时间
    def compute_total_time(cls, data):
        column_values = []
        column_index, _ = PlanData.start_end_col_num('生产时间', data.col_names_dict) 
        for cell in data.ws.iter_rows(min_col = column_index, max_col = column_index, values_only = True):
            column_values.append(cell[0]) 
        column_values = column_values[2:]
        column_values = [i for i in column_values if i is not None]
        return sum(column_values)
    

