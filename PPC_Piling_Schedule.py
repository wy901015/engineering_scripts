import pandas as pd 
import datetime as dt
import glob
import os

df_piling_schedule = pd.DataFrame()
df_point_reaction = pd.DataFrame()
df_point_coordinate = pd.DataFrame()
files = glob.glob("MSC_P2 v1.0.xls")
output_files = "PPC Piling Schedule"

for each in files:
    df_point_reaction = pd.read_excel(each, sheetname="Nodal Reactions", skiprows=1)
    df_point_coordinate = pd.read_excel(each, sheetname="Obj Geom - Point Coordinates", skiprows=1)

# print(df_point_reaction.dtypes)
df = df_point_reaction[df_point_reaction.Point.str.startswith("P", na=False)].reset_index(drop=True)
df2 = df.loc[:, ["Point", "OutputCase", "Fz"]]
df3 = df2.pivot(index = "Point", columns = "OutputCase", values = "Fz")
print(df3)