from FileData import FileData
import pandas as pd
import numpy as np


file_data = FileData("F:\\Python Projects\\170622_MDS.txt")

print(file_data.df)
file_data.df.fillna(0)
print(file_data.df)

df = pd.DataFrame([[np.nan, 2, np.nan, 0],
                   [3, 4, np.nan, 1],
                   [np.nan, np.nan, np.nan, 5],
                   [np.nan, 3, np.nan, 4]],
                  columns=list('ABCD'))
print(df)
df.fillna(0, inplace = True)
print(df)
