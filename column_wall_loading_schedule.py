import pandas as pd, datetime as dt
import glob, os
import numpy as np

files = glob.glob("MSC2 Tower v3-column wall loading data.xlsx")


df_pier_force = pd.DataFrame()
df_Pier_section = pd.DataFrame()
df_reaction_force = pd.DataFrame()
df_wall_connect = pd.DataFrame()
df_area_data = pd.DataFrame()
df_column_data = pd.DataFrame()
df_story_data = pd.DataFrame()
df_frame_data = pd.DataFrame()
df_section_assgin = pd.DataFrame()
df_DL = pd.DataFrame()
df_SDL = pd.DataFrame()
df_LL = pd.DataFrame()
df_WX = pd.DataFrame()
df_WY = pd.DataFrame()
df_WU = pd.DataFrame()
df_WV = pd.DataFrame()
df_EX = pd.DataFrame()
df_EY = pd.DataFrame()
df_UP = pd.DataFrame()

def column_wall_loading_schedule():
    
    
    for each in files:
        #sheets = pd.ExcelFile(each).sheet_names

        df_pier_force = pd.read_excel(each, sheetname = 'Pier_Forces',converters={'Story':str})
        df_Pier_section = pd.read_excel(each, sheetname = 'Pier_Section_Properties',converters={'Story':str})
        df_reaction_force = pd.read_excel(each, sheetname = 'Support_Reactions',converters={'Story':str})
        df_wall_connect = pd.read_excel(each, sheetname = 'Wall_Connectivity_Data')
        df_area_data = pd.read_excel(each, sheetname = 'Area_Piers_Spandrels',converters={'Story':str})
        df_column_data = pd.read_excel(each, sheetname = 'Column_Connectivity_Data',converters={'Story':str})
        df_story_data = pd.read_excel(each, sheetname = 'Story_Data',converters={'Story':str})
        df_frame_data = pd.read_excel(each, sheetname = 'Frame_Assignments_Summary',converters={'Story':str}) 
        df_section_assgin = pd.read_excel(each, sheetname = 'Frame_Section_Assignments',converters={'Story':str})

    df_reaction_point = df_reaction_force.loc[:,['Story','Point']]
    df_story_data['Story2'] = df_story_data.Story.shift(1)
    df_story_data2 = df_story_data.loc[:,['Story','Story2']]
    df_re=pd.merge(df_reaction_point,df_story_data2,on=['Story'],how='left')
    df_re.rename(columns = {'Story':'Story_below','Story2':'Story'}, inplace = True)
    
    df_frame_data_all = pd.merge(df_frame_data,df_section_assgin,on=['Story','Line','LineType','AnalysisSect'],how='left')
    
    

    df_pier_data = df_frame_data_all.loc[:,['Story','Line','LineType','Pier','SectionType']].dropna(axis=0)
    df_pier_data.rename(columns = {'Line':'Wall'}, inplace = True)

    df2=pd.melt(df_wall_connect,id_vars=['Wall'], value_vars=['Point1', 'Point2'])
    df2_column=pd.melt(df_column_data,id_vars=['Column'], value_vars=['IEndPt', 'JEndPt'])

    df2=df2.sort_values(['Wall','value'], ascending=[False,True]).reset_index(drop=True)
    df2.rename(columns = {'value':'Point'}, inplace = True)
    df2.Point = df2.Point.astype(str)
    df2_wall=df2.loc[:,['Wall','Point']]

    df2_column=df_column_data.loc[:,['Column','IEndPt']]
    df2_column.rename(columns = {'Column':'Wall','IEndPt':'Point'}, inplace = True)
    df2_column.Point = df2_column.Point.astype(str)

    df3 = df2_wall.append(df2_column).reset_index(drop=True)
    df4=pd.merge(df_re,df3,on=['Point'],how='left')
    df41=df4.loc[:,['Story','Wall']].dropna(axis=0).drop_duplicates().reset_index(drop=True)
    
    df_pier_size = df_Pier_section.loc[:,['Story','Pier','AxisAngle','WidthBot','ThickBot']]

    df_area_data=df_area_data.dropna(axis=1)
    df_area_data.rename(columns = {'Area':'Wall'}, inplace = True)
    df_pier_all=df_pier_data.append(df_area_data)
    df_pier = pd.merge(df41,df_pier_all,on=['Story','Wall'],how='left')
    df6 = pd.merge(df_pier,df_pier_force,on=['Story','Pier'],how='left')
    df_pier_size = df_Pier_section.loc[:,['Story','Pier','AxisAngle','WidthBot','ThickBot']]
    df_pier_size.rename(columns = {'WidthBot':'D','ThickBot':'B'}, inplace = True)
    df_pier_size.B = df_pier_size.B*1000
    df_pier_size.D = df_pier_size.D*1000

    df7 = pd.merge(df6,df_pier_size,on=['Story','Pier'],how='left')
    df7.loc[df7.SectionType == 'Circle','B'] = '-'
    df_out=df7.drop_duplicates(subset=['Story','Pier','Load','Loc']).reset_index(drop=True).drop(['SectionType'], axis=1)
     
    return df_out

def loading_schedule():

    df_bp0 = df_out[df_out.Loc == 'Bottom'].loc[:,['Story','Pier','Load','P','V2','V3','M2','M3','AxisAngle']].reset_index(drop=True)
    df_bp0['AxisAngle90'] = df_bp0.AxisAngle+90
    df_bp0.P = df_bp0.P*-1
    df_bp0['Vx'] = df_bp0.V2*np.cos(df_bp0.AxisAngle*np.pi/180)-df_bp0.V3*np.sin(df_bp0.AxisAngle*np.pi/180)
    df_bp0['Vy'] = df_bp0.V2*np.sin(df_bp0.AxisAngle*np.pi/180)+df_bp0.V3*np.cos(df_bp0.AxisAngle*np.pi/180)
    df_bp0['Mx'] = df_bp0.M2*np.cos(df_bp0.AxisAngle*np.pi/180)-df_bp0.M3*np.sin(df_bp0.AxisAngle*np.pi/180)
    df_bp0['My'] = df_bp0.M2*np.sin(df_bp0.AxisAngle*np.pi/180)+df_bp0.M3*np.cos(df_bp0.AxisAngle*np.pi/180)
    df_bp = df_bp0.loc[:,['Story','Pier','Load','P','Vx','Vy','Mx','My']]
    df_pier_size = df_out[df_out.Loc == 'Bottom'].loc[:,['Story','Pier','B','D']].drop_duplicates().reset_index(drop=True)
    df_pier_angle = df_out[df_out.Loc == 'Bottom'].loc[:,['Story','Pier','AxisAngle']].drop_duplicates().reset_index(drop=True)

    df_DL = df_bp[df_bp.Load == 'DDL'].drop(['Load'], axis=1)
    df_SDL = df_bp[df_bp.Load == 'DSDL'].drop(['Load'], axis=1)
    df_LL = df_bp[df_bp.Load == 'DLL'].drop(['Load'], axis=1)
    df_WX = df_bp[df_bp.Load == 'DWX'].drop(['Load'], axis=1)
    df_WY = df_bp[df_bp.Load == 'DWY'].drop(['Load'], axis=1)
    df_WU = df_bp[df_bp.Load == 'DWU'].drop(['Load'], axis=1)
    df_WV = df_bp[df_bp.Load == 'DWV'].drop(['Load'], axis=1)
    df_EX = df_bp[df_bp.Load == 'DEX'].drop(['Load'], axis=1)
    df_EY = df_bp[df_bp.Load == 'DEY'].drop(['Load'], axis=1)
    df_SOIL = df_bp[df_bp.Load == 'SOIL' ].drop(['Load'], axis=1)
    df_UP = df_bp[df_bp.Load == 'UPLIFT'].drop(['Load'], axis=1)
    df_T1 = df_bp[df_bp.Load == 'T1' ].drop(['Load'], axis=1)
    df_T2 = df_bp[df_bp.Load == 'T2' ].drop(['Load'], axis=1)

    df_SDL.rename(columns = {'P':'SDL-P','Vx':'SDL-Vx','Vy':'SDL-Vy','Mx':'SDL-Mx','My':'SDL-My'}, inplace = True)
    df_LL.rename(columns = {'P':'LL-P','Vx':'LL-Vx','Vy':'LL-Vy','Mx':'LL-Mx','My':'LL-My'}, inplace = True)

    df_pp= pd.merge(df_pier_size,df_DL,on=['Story','Pier'],how='left')
    DF_DD = pd.merge(df_pp,df_SDL,on=['Story','Pier'],how='left')
    DF_DDLL = pd.merge(DF_DD,df_LL,on=['Story','Pier'],how='left')
    DF_DDLL['DL'] = DF_DDLL.P+DF_DDLL['SDL-P']
    DF_DDLL['DL+LL'] = DF_DDLL.DL+DF_DDLL['LL-P']

    DF_DDLL['DL-MX'] = DF_DDLL.Mx+DF_DDLL['SDL-Mx']
    DF_DDLL['DLLL-MX'] = DF_DDLL['DL-MX']+DF_DDLL['LL-Mx']

    DF_DDLL['DL-MY'] = DF_DDLL.My+DF_DDLL['SDL-My']
    DF_DDLL['DLLL-MY'] = DF_DDLL['DL-MY']+DF_DDLL['LL-My']

    df_dead = DF_DDLL.loc[:,['Story','Pier','B','D','DL','DL+LL','DL-MX','DLLL-MX','DL-MY','DLLL-MY']]
    df1 = pd.merge(df_dead,df_WX,on=['Story','Pier'],how='left')
    df2 = pd.merge(df1,df_WY,on=['Story','Pier'],how='left')
    df3 = pd.merge(df2,df_WU,on=['Story','Pier'],how='left')
    df4 = pd.merge(df3,df_WV,on=['Story','Pier'],how='left')
    df5 = pd.merge(df4,df_EX,on=['Story','Pier'],how='left')
    df6 = pd.merge(df5,df_EY,on=['Story','Pier'],how='left')
    df61 = pd.merge(df6,df_SOIL,on=['Story','Pier'],how='left')
    df7 = pd.merge(df61,df_UP,on=['Story','Pier'],how='left')
    df8 = pd.merge(df7,df_T1,on=['Story','Pier'],how='left')
    df9 = pd.merge(df7,df_T2,on=['Story','Pier'],how='left')
    df_all = df9
    df_all_out = pd.merge(df_all,df_pier_angle,on=['Story','Pier'],how='left')
    df_all_out = df_all_out.drop(['Story'],axis =1)
    
    return df_all_out

def out(df_out,out_file):
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_out.to_excel(writer, sheet_name='Sorted')
    writer.save()

if __name__ == '__main__':
    df_out = column_wall_loading_schedule()
    out(df_out,'Pier_force-Sorted.xlsx')
    df_all_out = loading_schedule()
    out(df_all_out,'Column wall loading schedule - TOWER2.xlsx')
    print('Done')
    