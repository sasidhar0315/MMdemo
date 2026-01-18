import os
import pandas as pd
import random
from datetime import datetime, timedelta

# Number of rows
ROWS = 5000

employee_file = "employee_data.xlsx"
dept_file = "department.xlsx"

# If an employee file exists, read it; otherwise create one
if os.path.exists(employee_file):
    df = pd.read_excel(employee_file)
else:
    data = []
    for i in range(1, ROWS + 1):
        record = {
            "ID": i,
            "Name": f"User_{i}",
            "Age": random.randint(20, 60),
            "Salary": random.randint(30000, 120000),
            "Department": random.choice(["DevOps", "Backend", "Frontend", "QA", "HR"]),
            "Join_Date": datetime.now() - timedelta(days=random.randint(1, 3000))
        }
        data.append(record)
    df = pd.DataFrame(data)

# Read HOD details and merge (requires department.xlsx with Department and HOD columns)
if not os.path.exists(dept_file):
    raise SystemExit(f"{dept_file} not found in working directory. Place department.xlsx alongside this script.")

dept_df = pd.read_excel(dept_file)
required_cols = {"Department", "HOD"}
if not required_cols.issubset(set(dept_df.columns)):
    raise SystemExit(f"{dept_file} must contain columns: {', '.join(required_cols)}")

df = df.merge(dept_df, on="Department", how="left")
missing_hod = int(df["HOD"].isna().sum())

# Overwrite the same employee file with HOD included
df.to_excel(employee_file, index=False)

print(f"✅ Excel file updated: {employee_file}")
print(f"📊 Total Rows Written: {len(df)}")
print(f"ℹ️ Missing HOD entries: {missing_hod}")
