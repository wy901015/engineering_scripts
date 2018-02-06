#PPC Piling Schedule Module
import pandas as pd
import numpy as np
import os
import glob
import ppc

# Sys configuration
files = glob.glob("MSC_P2 v1.0.xls")


#User input
required_cases = ["DL", "SDL", "LL", "WX", "WY", "WU", "WV", "EX", "EY", "UPLIFT"]
pile_capacity = 3500
group_factor = 0.85
nsf = {147.1: 1037, "else": 747, 51.355: 591}
target_util = 0.5
sls_without_wind = {
    "D+L+NSF" : {"D": 1, "L": 1, "E": 0, "W": 0, "NSF": 0.85, "U": 0}
}
sls_with_wind = {
    "D+L+E" : {"D": 1, "L": 1, "E": 1, "W" : 0, "NSF": 0, "U" : 0},
    "D+L+W+E" : {"D": 1, "L": 1, "W": 1, "E": 1, "NSF": 0, "U": 0},
    "D+0.4L+W+E" : {"D": 1, "L": 0.4, "W": 1, "E": 1, "NSF": 0, "U": 0},
    "0.67D-W-E-U" : {"D": 0.67, "L": 0.4, "W": -1, "E": -1, "NSF": 0, "U": -1}
}

for each in files:
    df_point_reaction = pd.read_excel(each, sheetname="Nodal Reactions", skiprows=1, converters={"OutputCase": str})
    df_point_coordinates = pd.read_excel(each, sheetname="Obj Geom - Point Coordinates", skiprows=1)


def return_not_matches(a, b):
    a = set(a)
    b = set(b)
    return list(b - a)

# Print out ppc piling schedule   
def return_ppc_piling_schedule():   
    df = df_point_reaction[df_point_reaction.Point.str.startswith("P", na=False)].reset_index(drop=True)
    df2 = df_point_coordinates[df_point_coordinates.Point.str.startswith("P", na=False) & ~df_point_coordinates.Point.str.contains("_", na=False)].reset_index(drop=True).loc[:, ["Point", "GlobalX", "GlobalY"]]
    df_cap = ppc.return_ppc_pile_cap(files)

    output_cases = df.OutputCase.unique()
    unmatch_value = return_not_matches(output_cases, required_cases)
    point = df.Point.unique()
    df_case = pd.DataFrame({"Point" : point})

    for case in required_cases:
        if case in unmatch_value:
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
    df_schedule["PileCP"] = pile_capacity * group_factor
    df_schedule["PileCPW"] = df_schedule.PileCP * 1.25

    df_wind_max = df_schedule[["WX", "WY", "WU", "WV"]].max(axis = 1)
    df_seismic_max = df_schedule[["EX", "EY"]].max(axis = 1)

    for sls_case, value in sls_without_wind.items():
        print ("Working load cases include: " + sls_case)
        df_schedule[sls_case] = df_schedule.DDL * sls_without_wind[sls_case]["D"] + df_schedule.LL * sls_without_wind[sls_case]["L"] + df_wind_max * sls_without_wind[sls_case]["W"] + df_seismic_max * sls_without_wind[sls_case]["E"] + df_schedule.NSF * sls_without_wind[sls_case]["NSF"] + df_schedule.UPLIFT * sls_without_wind[sls_case]["U"] 
        df_schedule["Ratio " + sls_case] = np.where(df_schedule["PileCP"] > df_schedule[sls_case], df_schedule[sls_case]/df_schedule["PileCP"], 999).astype(float)
        df_schedule[sls_case] = df_schedule[sls_case].astype(float, errors="raise")

    for sls_case, value in sls_with_wind.items():
        print ("Working load cases include: " + sls_case)
        df_schedule[sls_case] = df_schedule.DDL * sls_with_wind[sls_case]["D"] + df_schedule.LL * sls_with_wind[sls_case]["L"] + df_wind_max * sls_with_wind[sls_case]["W"] + df_seismic_max * sls_with_wind[sls_case]["E"] + df_schedule.NSF * sls_with_wind[sls_case]["NSF"] + df_schedule.UPLIFT * sls_with_wind[sls_case]["U"] 
        df_schedule["Ratio " + sls_case] = np.where(df_schedule["PileCPW"] > df_schedule[sls_case], df_schedule[sls_case]/df_schedule["PileCPW"], 999).astype(float)
        df_schedule[sls_case] = df_schedule[sls_case].astype(float, errors="raise")

    df_ratio = df_schedule.filter(regex="Ratio")
    df_schedule["Ratio"] = df_ratio.max(axis=1)
    cap = df_schedule.merge(df_cap[["Point", "Pile Cap"]], on="Point")
    count = pd.DataFrame(cap["Pile Cap"].value_counts().reset_index())
    count.columns = ["Pile Cap", "Count"]
    df_schedule_with_cap_count = cap.merge(count, on="Pile Cap")
    df_schedule_out = df_schedule_with_cap_count.loc[:, ["Point", "Pile Cap", "Count", "NSF", "DDL", "LL", "WX", "WY", "WU", "WV", "EX", "EY", "UPLIFT"] + list(sls_without_wind.keys()) + list(sls_with_wind.keys()) + ["PileCP", "PileCPW", "Ratio"]]

    return df_schedule_out

# Print out to dxf
def return_bad_piles_in_dxf():
    df_to_dxf = df_schedule.loc[df_schedule["Ratio"] < target_util , ["Point", "GlobalX", "GlobalY"]]
    df_to_dxf["Input"] = "-TEXT"
    df_to_dxf["Text Height"] = 300
    df_to_dxf["Text Angle"] = 0
    df_to_dxf["Coordinates"] = (df_to_dxf["GlobalX"] * 1000).astype(str) + "," + (df_to_dxf["GlobalY"] * 1000).astype(str)
    df_to_dxf = df_to_dxf.merge(df_schedule[["Point", "Ratio"]], on="Point")
    df_to_dxf = df_to_dxf.loc[:, ["Input", "Coordinates", "Text Height", "Text Angle", "Ratio"]]
    df_to_dxf.to_csv('D:/Python/engineering_scripts/output/to dxf.csv', header=False, index=False, sep=" ")
