
from pathlib import Path as _WorkspacePath
import os as _workspace_os
RESEARCH_WORKSPACE = _WorkspacePath(_workspace_os.environ.get("RESEARCH_WORKSPACE", _workspace_os.getcwd())).resolve()

import pandas as pd

df = pd.read_excel(str(RESEARCH_WORKSPACE / '乡村振兴面板数据_2015-2022.xlsx'), sheet_name="面板数据")
print("Columns:", df.columns.tolist())
print("\nShape:", df.shape)
print("\nFirst few rows:")
print(df.head())
print("\nData types:")
print(df.dtypes)
