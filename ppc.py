import pandas as pd, datetime as dt
import glob, os
import sys

def return_ppc_pile_cap(files):
    
    # output_file = 'PPC Pile.xlsx'
    df_frame_force = pd.DataFrame()
    df_frame_coor = pd.DataFrame()
    df_slab = pd.DataFrame()

    for each in files:
        #sheets = pd.ExcelFile(each).sheet_names
        sheets = pd.ExcelFile(each).sheet_names
        df_frame_force = pd.read_excel(each, sheetname = 'Obj Geom - Areas 01 - General', skiprows=1)
        df_frame_coor = pd.read_excel(each, sheetname = 'Obj Geom - Point Coordinates', skiprows=1)
        df_slab = pd.read_excel(each, sheetname = 'Slab Property Assignments', skiprows=1)
        
    def is_point_inside_rect(pilecap,xx,yy):
        bp = pd.DataFrame()
        bp = pilecap.reset_index()
        bp['x'] = xx.iloc[0]-bp.X_min
        bp['y'] = bp.Y_max-yy.iloc[0]
        bp['wd']=bp.Width*bp.Depth
        bp['xy']=bp.x*bp.y
        bp['ff']=0
        bp.loc[(bp.x > 0) & (bp.y > 0) & (bp.y < bp.Depth) & (bp.x < bp.Width), 'ff'] = 1
        # aa=bp.Area[bp.ff==1]
        return bp.Area[bp.ff==1].iloc[0]

    df=df_frame_force.loc[df_frame_force.index[1:],['Area','Point1','Point2','Point3','Point4']]
    df2=pd.melt(df,id_vars=['Area'], value_vars=['Point1', 'Point2', 'Point3', 'Point4'])
    df2=df2.sort_values(['Area','value'], ascending=[False,True]).reset_index(drop=True)
    df2.rename(columns = {'value':'Point'}, inplace = True)
    df3 = df_frame_coor.loc[1:, ["Point", "GlobalX", "GlobalY"]]
    df4 = pd.merge(df2, df3, how='left', on=['Point'])
    df_slab2=df_slab.loc[1:,:]
    df_slab3=df_slab2[(df_slab2.SlabProp != 'None') & (df_slab2.SlabProp != 'S450') & (df_slab2.SlabProp != 'S2500') & (df_slab2.SlabProp != 'S3000')]
    df5 = pd.merge(df4, df_slab3, how='inner', on=['Area'])

    bp_xmin=df5.groupby(['Area']).GlobalX.min()
    bp_xmax=df5.groupby(['Area']).GlobalX.max()
    bp_ymin=df5.groupby(['Area']).GlobalY.min()
    bp_ymax=df5.groupby(['Area']).GlobalY.max()

    df_pilecap = pd.DataFrame()
    df_pilecap['X_min'] = bp_xmin
    df_pilecap['X_max'] = bp_xmax
    df_pilecap['Y_min'] = bp_ymin
    df_pilecap['Y_max'] = bp_ymax
    df_pilecap['Width'] = df_pilecap.X_max - df_pilecap.X_min
    df_pilecap['Depth'] = df_pilecap.Y_max - df_pilecap.Y_min
    df_pilecap['X_center'] = df_pilecap.X_min + df_pilecap.Width/2
    df_pilecap['Y_center'] = df_pilecap.Y_min + df_pilecap.Depth/2

    ii=0
    ppc = pd.DataFrame()
    ppc_name = []
    ppc=df3[df3.Point.str.startswith('P', na=False) & ~df3.Point.str.contains('_', na=False)].reset_index(drop=True)
    for name in ppc.Point:
        x = ppc.GlobalX[ppc.Point==name]
        y = ppc.GlobalY[ppc.Point==name]
        aa11=is_point_inside_rect(df_pilecap,x,y)
        ppc_name.append(aa11)
        ii+=1

    ppc["Pile Cap"] = ppc_name

    return ppc

    # writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # # Convert the dataframe to an XlsxWriter Excel object.

    # ppc.index +=1

    # ppc.to_excel(writer, sheet_name='PPC')

    # df_pilecap_OUT = df_pilecap.reset_index()
    # df_pilecap_OUT.index +=1
    # df_pilecap_OUT.to_excel(writer, sheet_name='Pile Cap')

    # writer.save()

    # print("Complete! Please check " + output_file)
