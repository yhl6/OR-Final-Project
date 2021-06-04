import numpy as np
import pandas as pd
from gurobipy import *

MCP = Model()

# Set and Parameters
N = 1000  # 醫護人員數量上限
H = 300  # 熱區標準
Days = range(1, 17)  # 日期集合
Districts = range(1, 13)  # 行政區集合
Villages = range(1, 457)  # 里的集合
Candidates = range(1, 130)  # 候選點集合
Levels = range(1, 6)  # 快篩站等級集合

# Variables
# x_lq 若第k個行政區的第l個候選點有設置等級q的快篩站則為1，反之為0
x = MCP.addVars(Candidates, Levels, lb=0, vtype=GRB.BINARY)
# y_k 若第t天第i行政區是需要設置快篩站的熱區則為1，反之為0
y = MCP.addVars(Days, lb=0, vtype=GRB.BINARY)
# c_j 若j里有被k區第l個候選點所設的q級快篩站所服務到則為1，反之為
c = MCP.addVars(Villages, lb=0, vtype=GRB.BINARY)

# Objective Value



