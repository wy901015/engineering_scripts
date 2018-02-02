import pandas as pd
import numpy as np
import os
import glob

# Sys configuration
files = glob.glob("MSC_P2 v1.0.xls")
output_files = "D:/Python/engineering_scripts/output/PPC Piling Schedule.xlsx"

#User input
requiredCases = ["DL", "SDL", "LL", "WX", "WY", "WU", "WV", "EX", "EY", "UPLIFT"]
pileCapacity = 3500
groupFactor = 0.85
nsf = {147.1: 1037, "else": 747, 51.355: 591}
targetUtilization = 0.5
slsWithoutWind = {
    "D+L+NSF" : {"D": 1, "L": 1, "E": 0, "W": 0, "NSF": 0.85, "U": 0}
}
slsWithWind = {
    "D+L+E" : {"D": 1, "L": 1, "E": 1, "W" : 0, "NSF": 0, "U" : 0},
    "D+L+W+E" : {"D": 1, "L": 1, "W": 1, "E": 1, "NSF": 0, "U": 0},
    "D+0.4L+W+E" : {"D": 1, "L": 0.4, "W": 1, "E": 1, "NSF": 0, "U": 0},
    "0.67D-W-E-U" : {"D": 0.67, "L": 0.4, "W": -1, "E": -1, "NSF": 0, "U": -1}
}

for each in files:
    df_point_reaction = pd.read_excel(each, sheetname="Nodal Reactions", skiprows=1, converters={"OutputCase": str})
    df_point_coordinates = pd.read_excel(each, sheetname="Obj Geom - Point Coordinates", skiprows=1)

def returnNotMatches(a, b):
    a = set(a)
    b = set(b)
    return list(b - a)

df = df_point_reaction[df_point_reaction.Point.str.startswith("P", na=False)].reset_index(drop=True)
df2 = df_point_coordinates[df_point_coordinates.Point.str.startswith("P", na=False) & ~df_point_coordinates.Point.str.contains("_", na=False)].reset_index(drop=True).loc[:, ["Point", "GlobalX", "GlobalY"]]

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
        df_new = df[df["OutputCase"] == case].loc[:, ["Point", "Fz"]].rename(columns = {"Fz": case}).reset_index(drop=True)
        df_case = df_case.merge(df_new, on="Point", how="right")
        print(case)

df_case['DDL'] = df_case["DL"] + df_case['SDL']
df_schedule = df_case.merge(df2, on="Point", how="right")

df_schedule["NSF"] = np.where(df_schedule.GlobalY < 51.355, nsf[51.355], np.where(df_schedule.GlobalY > 147.1, nsf[147.1], nsf["else"]))
df_schedule["NSF"] = df_schedule["NSF"].astype(float, errors="raise")
df_schedule["PileCP"] = pileCapacity * groupFactor
df_schedule["PileCPW"] = df_schedule.PileCP * 1.25

df_wind_max = df_schedule[["WX", "WY", "WU", "WV"]].max(axis = 1)
df_seismic_max = df_schedule[["EX", "EY"]].max(axis = 1)

for slsCase, value in slsWithoutWind.items():
    print ("Working load cases include: " + slsCase)
    df_schedule[slsCase] = df_schedule.DDL * slsWithoutWind[slsCase]["D"] + df_schedule.LL * slsWithoutWind[slsCase]["L"] + df_wind_max * slsWithoutWind[slsCase]["W"] + df_seismic_max * slsWithoutWind[slsCase]["E"] + df_schedule.NSF * slsWithoutWind[slsCase]["NSF"] + df_schedule.UPLIFT * slsWithoutWind[slsCase]["U"] 
    df_schedule["Ratio " + slsCase] = np.where(df_schedule["PileCP"] > df_schedule[slsCase], df_schedule[slsCase]/df_schedule["PileCP"], -1).astype(float)
    df_schedule[slsCase] = df_schedule[slsCase].astype(float, errors="raise")

for slsCase, value in slsWithWind.items():
    print ("Working load cases include: " + slsCase)
    df_schedule[slsCase] = df_schedule.DDL * slsWithWind[slsCase]["D"] + df_schedule.LL * slsWithWind[slsCase]["L"] + df_wind_max * slsWithWind[slsCase]["W"] + df_seismic_max * slsWithWind[slsCase]["E"] + df_schedule.NSF * slsWithWind[slsCase]["NSF"] + df_schedule.UPLIFT * slsWithWind[slsCase]["U"] 
    df_schedule["Ratio " + slsCase] = np.where(df_schedule["PileCPW"] > df_schedule[slsCase], df_schedule[slsCase]/df_schedule["PileCPW"], -1).astype(float)
    df_schedule[slsCase] = df_schedule[slsCase].astype(float, errors="raise")

df_ratio = df_schedule.filter(regex="Ratio")
df_schedule["Ratio"] = df_ratio.max(axis=1)
df_schedule_out = df_schedule.loc[:, ["Point", "NSF", "DDL", "LL", "WX", "WY", "WU", "WV", "EX", "EY", "UPLIFT"] + list(slsWithWind.keys()) + ["PileCP", "PileCPW", "Ratio"]]
df_schedule_out.to_excel(output_files, sheet_name="PPC Piling Schedule", index=False)

# Print out to dxf
df_to_dxf = df_schedule.loc[df_schedule["Ratio"] < targetUtilization , ["Point", "GlobalX", "GlobalY"]]
df_to_dxf["Input"] = "-TEXT"
df_to_dxf["Text Height"] = 300
df_to_dxf["Text Angle"] = 0
df_to_dxf["Coordinates"] = (df_to_dxf["GlobalX"] * 1000).astype(str) + "," + (df_to_dxf["GlobalY"] * 1000).astype(str)
df_to_dxf = df_to_dxf.merge(df_schedule[["Point", "Ratio"]], on="Point")
df_to_dxf = df_to_dxf.loc[:, ["Input", "Coordinates", "Text Height", "Text Angle", "Ratio"]]
df_to_dxf.to_csv('D:/Python/engineering_scripts/output/to dxf.csv', header=False, index=False, sep=" ")

