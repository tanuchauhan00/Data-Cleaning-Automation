import pandas as pd
import numpy as np
import random

random.seed(42)
np.random.seed(42)

regions = ["North", "South", "East", "West", "north", "SOUTH", " East "]
products = ["Widget A", "Widget B", "Widget C", "Widget D", "widget a", "Widget B "]
reps = ["J. Smith", "A. Khan", "M. Chen", "R. Patel", "L. Garcia"]

rows = []
start_date = pd.Timestamp("2026-01-01")
for i in range(1, 201):
    date_fmt = random.choice(["%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y"])
    date = (start_date + pd.Timedelta(days=random.randint(0, 150))).strftime(date_fmt)
    row = {
        "OrderID": 1000 + i,
        "Date": date,
        "Region": random.choice(regions),
        "Product": random.choice(products),
        "SalesRep": random.choice(reps),
        "Units": random.choice([np.nan, 1, 2, 3, 5, 10, -1]),
        "UnitPrice": random.choice([np.nan, 9.99, 19.99, 29.99, 49.99]),
        "Revenue": np.nan,
    }
    rows.append(row)

df = pd.DataFrame(rows)

dupes = df.sample(15, random_state=1)
df = pd.concat([df, dupes], ignore_index=True)

for _ in range(5):
    df.loc[len(df)] = [np.nan] * len(df.columns)

df = df.sample(frac=1, random_state=2).reset_index(drop=True)
df.to_csv("/home/claude/project/data/raw_sales_data.csv", index=False)
print(df.shape)
