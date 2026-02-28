from pathlib import Path
import pandas as pd
import sys

print("PYTHON:", sys.executable)

# 从当前脚本位置开始一路向上找，直到找到含有 data 目录的项目根
p = Path(__file__).resolve()
project_root = None
for parent in [p] + list(p.parents):
    if (parent / "data").exists():
        project_root = parent
        break

if project_root is None:
    raise FileNotFoundError("Cannot find project root with a 'data' folder")

DATA_PATH = project_root / "data" / "store_ops_testdata_50rows.xlsx"

print("Reading from:", DATA_PATH)

df = pd.read_excel(DATA_PATH)
print("Rows:", len(df))
print(df.head(3))
