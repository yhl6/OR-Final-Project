import pandas as pd
import xlsxwriter
from gurobipy import *
# pd.set_option('display.max_columns', None, 'display.max_rows', None)


# Set and Parameters
N_list = [40, 80, 120]  # 醫護人員數量上限列表
H_list = [25, 15, 10]  # 設站低標列表
S = 8  # 每站醫護人員數量
U = 300  # 每站可快篩人數上限
Days = range(0, 16)  # 日期集合
Districts = range(0, 12)  # 行政區集合
Villages = range(0, 456)  # 里的集合 j
Candidates = range(0, 131)  # 候選點集合 k
# Data
df = pd.ExcelFile('Data.xlsx')
P = pd.read_excel(df, sheet_name='各里每日需篩檢人數（有考慮陽性率）', index_col=0, skiprows=0)
R = pd.read_excel(df, sheet_name='風險係數', index_col=0, skiprows=0)
Z = pd.read_excel(df, sheet_name='Z', index_col=0, skiprows=0)
# for i in Candidates:
#     print(Z.iat[455, i])
Station_info = pd.read_excel(df, sheet_name='station_info', index_col=0, skiprows=0)
Dist_Group = Station_info.groupby('town')
Dist_Name = Dist_Group.groups.keys()
# print(Dist_Group.groups)
# for key, value in Dist_Group.groups.items():
#     print(key, value)
for n in range(3):
    Y = pd.read_excel(df, sheet_name=f'設站低標{H_list[n]}', index_col=0, skiprows=0)
    # print(Y.columns)
    writer = pd.ExcelWriter(f'Result_設站低標{H_list[n]}.xlsx', engine='xlsxwriter')
    # print(Y.head())
    for t in Days:
        MCP = Model()
        MCP.Params.LogToConsole = 0  # 最佳化時不要印出log
        # Variables
        # x_k 若第k個候選點有設置快篩站則為1，反之為0
        x = MCP.addVars(Candidates, lb=0, vtype=GRB.BINARY)
        # f_jk 第k個候選點上的快篩站服務到j里的比例
        f = MCP.addVars(Villages, Candidates, lb=0, vtype=GRB.CONTINUOUS)

        # Objective Value
        MCP.setObjective(quicksum(P.iat[j, t + 3] * R.iat[j, t + 3] * quicksum(f[j, k] for k in Candidates) for j in Villages), GRB.MAXIMIZE)

        # Constraints
        # 所有快篩站的總醫護人數不能超過某個上限(醫療資源有限)
        MCP.addConstr(S * quicksum(x.select('*')) <= N_list[n])

        # 所有候選點上的快篩站服務到j里的比例和不超過1
        MCP.addConstrs((quicksum(f.select(j, '*')) <= 1) for j in Villages)

        # f_{jk}定義
        MCP.addConstrs((Z.iat[j, k] * x[k] >= f[j, k]) for j in Villages for k in Candidates)

        # 快篩站服務人數不超過上限
        MCP.addConstrs((quicksum(f[j, k] * P.iat[j, t + 3] for j in Villages) <= x[k] * U) for k in Candidates)

        # 一區當日新增確診人數超過標準時至少要有一個快篩站
        for key, value in Dist_Group.groups.items():
            MCP.addConstr(quicksum(x[k - 1] for k in value) >= Y.at[key, 20210514 + t])

        MCP.optimize()
        print(f'{t + 1}: {MCP.objVal}')
        f_sol_df = pd.DataFrame(index=[i + 1 for i in Villages], columns=[i + 1 for i in Candidates])
        x_sol_df = pd.DataFrame(index=[i + 1 for i in Villages], columns=[f'{20210514 + t}'])
        for j in Villages:
            for k in Candidates:
                f_sol_df.iat[j, k] = f[j, k].x
        for k in Candidates:
            x_sol_df.iloc[k] = x[k].x
        f_sol_df.to_excel(writer, sheet_name=f'f_{20210514 + t}')
        x_sol_df.to_excel(writer, sheet_name=f'x_{20210514 + t}')
    writer.save()