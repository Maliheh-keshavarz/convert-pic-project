import pandas as pd

df = pd.read_excel(r"C:\Users\malih\OneDrive\Documents\GitLab project\project for image\convert pic\p2\convert-pic-project\data.xlsx", header=1, dtype=str)
print(df.columns.tolist())
print(df.head())

