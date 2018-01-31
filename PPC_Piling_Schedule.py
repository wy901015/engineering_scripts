import pandas as pd 
import os
import glob

# User input
files = glob.glob("MSC_P2 v1.0.xls")
output_files = "PPC Piling Schedule"
requiredCases = ["DL", "SDL", "LL", "WX", "WY", "WU", "WV", "EX", "EY"]

for each in files:
    df_point_reaction = pd.read_excel(each, sheetname="Nodal Reactions", skiprows=1, converters={"OutputCase": str})

def returnNotMatches(a, b):
    a = set(a)
    b = set(b)
    return list(b - a)

def ppc_piling_schedule(): 
  
    df = df_point_reaction[df_point_reaction.Point.str.startswith("P", na=False)].reset_index(drop=True)
    outputCases = df.OutputCase.unique()
    unmatchesValue = returnNotMatches(outputCases, requiredCases)
    point = df.Point.unique()
    df_case = pd.DataFrame({"Point" : point})

    for case in requiredCases:
        if case in unmatchesValue:
            df_na = pd.DataFrame(data={"Point" : point, case: 0}, columns=["Point", case])
            df_case = df_case.merge(df_na, on="Point", how="right")
            print("missing " + case)
        else:
            df2 = df[df["OutputCase"] == case].loc[:, ["Point", "Fz"]].rename(columns = {"Fz": case}).reset_index(drop=True)
            df_case = df_case.merge(df2, on="Point", how="right")
            print(case)

    return df_case



print(ppc_piling_schedule())

