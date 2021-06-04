import numpy as np
import pandas as pd
import xlsxwriter
from gurobipy import *
pd.set_option('display.max_columns', None, 'display.max_rows', None)

# Set and Parameters
N = 300  # 醫護人員數量上限列表
S = 15  # 每站醫護人員數量
U = 300  # 每站可快篩人數上限
Days = range(0, 16)  # 日期集合
Districts = range(0, 12)  # 行政區集合
Villages = range(0, 456)  # 里的集合 j
Candidates = range(0, 129)  # 候選點集合 k
# Levels = range(1, 6)  # 快篩站等級集合

# Data
df = pd.ExcelFile('Data.xlsx')
P = pd.read_excel(df, sheet_name='各里每日需篩檢人數（有考慮陽性率）', index_col=0, skiprows=0)
R = pd.read_excel(df, sheet_name='風險係數', index_col=0, skiprows=0)
Z = pd.read_excel(df, sheet_name='Z', index_col=0, skiprows=0)
# print(P.head())
# print(R.head())
# print(Z.head())
# print(P.iat[0, 3])
# print(R.iat[0, 3])
# print(Z.iat[0, 1])
writer = pd.ExcelWriter('Result有乘R.xlsx', engine='xlsxwriter')
for t in Days:
    MCP = Model()
    MCP.Params.LogToConsole = 0
    # Variables
    # x_k 若第k個候選點有設置快篩站則為1，反之為0
    x = MCP.addVars(Candidates, lb=0, vtype=GRB.BINARY)
    # f_jk 第k個候選點上的快篩站服務到j里的比例
    f = MCP.addVars(Villages, Candidates, lb=0, vtype=GRB.CONTINUOUS)
    # # y_k 若第t天第i行政區是需要設置快篩站的熱區則為1，反之為0
    # y = MCP.addVars(Days, lb=0, vtype=GRB.BINARY)
    # # c_j 若j里有被k區第l個候選點所設的q級快篩站所服務到則為1，反之為
    # c = MCP.addVars(Villages, lb=0, vtype=GRB.BINARY)

    # Objective Value
    MCP.setObjective(quicksum(P.iat[j, t + 3] * R.iat[j, t + 3] * quicksum(f[j, k] for k in Candidates) for j in Villages), GRB.MAXIMIZE)

    # Constraints
    # 所有快篩站的總醫護人數不能超過某個上限(醫療資源有限)
    MCP.addConstr(S * quicksum(x.select('*')) <= N)

    # 所有候選點上的快篩站服務到j里的比例和不超過1
    MCP.addConstrs((quicksum(f.select(j, '*')) <= 1) for j in Villages)

    # f_{jk}定義
    MCP.addConstrs((Z.iat[j, k + 1] * x[k] >= f[j, k]) for j in Villages for k in Candidates)

    # 快篩站服務人數不超過上限
    MCP.addConstrs((quicksum(f[j, k] * P.iat[j, t + 3] for j in Villages) <= x[k] * U) for k in Candidates)

    MCP.optimize()
    print(f'{t + 1}: {MCP.objVal}')
    f_sol_df = pd.DataFrame(index=range(1, 457), columns=range(1, 130))
    x_sol_df = pd.DataFrame(index=range(1, 130), columns=[20210514 + t])
    # print(sol_df)
    for j in Villages:
        for k in Candidates:
            f_sol_df.iat[j, k] = f[j, k].x
    for k in Candidates:
        x_sol_df.iloc[k] = x[k].x
    f_sol_df.to_excel(writer, sheet_name=f'f_{20210514 + t}')
    x_sol_df.to_excel(writer, sheet_name=f'x_{20210514 + t}')
writer.save()