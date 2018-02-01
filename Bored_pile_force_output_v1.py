import numpy as np
import pandas as pd, datetime as dt
import glob, os

# for Axial and lateral loading
files = glob.glob("Safe_output-lateral v0.2.xlsx")
output_file = 'Bored Pile-lateral-0.2.xlsx'
output_file_to_Prokon = 'Prokon_input.xlsx'
df_force = pd.DataFrame()


for each in files:

	df_force = pd.read_excel(each, sheetname = 'Nodal Reactions',skiprows=1)
	df_axial = pd.read_excel(each, sheetname = 'Axial')      

df_all = df_force.ix[1:,['Node','OutputCase','Fx','Fy','Fz','Mx','My','Mz']]
df_bp_all=df_all[df_all.Node.str.contains('BP')].reset_index(drop=True)



df_bp_all['Pile'], df_bp_all['Loc']=zip(*df_bp_all['Node'].map(lambda x: x.split('-')))

df_bp=df_bp_all.groupby(['Pile','OutputCase']).sum().reset_index()

df_bp_lateral = df_bp_all[df_bp_all.Mx != 0]
df_bp_lateral = pd.merge(df_bp_lateral,df_axial,on=['Pile'], how = 'left')

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

df_DL = df_bp[df_bp.OutputCase == 'DL']
df_SDL = df_bp[df_bp.OutputCase == 'SDL']
df_LL = df_bp[df_bp.OutputCase == 'LL']
df_WX = df_bp[df_bp.OutputCase == 'WX']
df_WY = df_bp[df_bp.OutputCase == 'WY']
df_WU = df_bp[df_bp.OutputCase == 'WU']
df_WV = df_bp[df_bp.OutputCase == 'WV']
df_EX = df_bp[df_bp.OutputCase == 'EX']
df_EY = df_bp[df_bp.OutputCase == 'EY']
df_UP = df_bp[df_bp.OutputCase == 'UPLIFT']
df_T1 = df_bp[df_bp.OutputCase == '+T']
df_T2 = df_bp[df_bp.OutputCase == '-T']


df_DL.rename(columns = {'Fx':'DL-Fx','Fy':'DL-Fy','Fz':'DL-Fz'}, inplace = True)
df_SDL.rename(columns = {'Fx':'SDL-Fx','Fy':'SDL-Fy','Fz':'SDL-Fz'}, inplace = True)
df_LL.rename(columns = {'Fx':'LL-Fx','Fy':'LL-Fy','Fz':'LL-Fz'}, inplace = True)
df_WX.rename(columns = {'Fx':'WX-Fx','Fy':'WX-Fy','Fz':'WX-Fz'}, inplace = True)
df_WY.rename(columns = {'Fx':'WY-Fx','Fy':'WY-Fy','Fz':'WY-Fz'}, inplace = True)
df_WU.rename(columns = {'Fx':'WU-Fx','Fy':'WU-Fy','Fz':'WU-Fz'}, inplace = True)
df_WV.rename(columns = {'Fx':'WV-Fx','Fy':'WV-Fy','Fz':'WV-Fz'}, inplace = True)
df_EX.rename(columns = {'Fx':'EX-Fx','Fy':'EX-Fy','Fz':'EX-Fz'}, inplace = True)
df_EY.rename(columns = {'Fx':'EY-Fx','Fy':'EY-Fy','Fz':'EY-Fz'}, inplace = True)
df_UP.rename(columns = {'Fx':'UP-Fx','Fy':'UP-Fy','Fz':'UP-Fz'}, inplace = True)
df_T1.rename(columns = {'Fx':'T1-Fx','Fy':'T1-Fy','Fz':'T1-Fz'}, inplace = True)
df_T2.rename(columns = {'Fx':'T2-Fx','Fy':'T2-Fy','Fz':'T2-Fz'}, inplace = True)

df_WX['WX']= df_WX[['WX-Fx','WX-Fy']].abs().max(axis=1)
df_WY['WY']= df_WY[['WY-Fx','WY-Fy']].abs().max(axis=1)
df_WU['WU']= df_WU[['WU-Fx','WU-Fy']].abs().max(axis=1)
df_WV['WV']= df_WV[['WV-Fx','WV-Fy']].abs().max(axis=1)
df_EX['EX']= df_EX[['EX-Fx','EX-Fy']].abs().max(axis=1)
df_EY['EY']= df_EY[['EY-Fx','EY-Fy']].abs().max(axis=1)

df2 = pd.merge(df_DL[['Pile','DL-Fz']], df_SDL[['Pile','SDL-Fz']], how='left', on=['Pile'])
df2 ['DL-total'] = df2['DL-Fz']+df2['SDL-Fz']

df2 = pd.merge(df2, df_LL[['Pile','LL-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_WX[['Pile','WX-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_WY[['Pile','WY-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_WU[['Pile','WU-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_WV[['Pile','WV-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_UP[['Pile','UP-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_EX[['Pile','EX-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_EY[['Pile','EY-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_T1[['Pile','T1-Fz']], how='left', on=['Pile'])
df2 = pd.merge(df2, df_T2[['Pile','T2-Fz']], how='left', on=['Pile'])


df3 = pd.merge(df_WX[['Pile','WX',]], df_WY[['Pile','WY']], how='left', on=['Pile'])
df3 = pd.merge(df3, df_WU[['Pile','WU']], how='left', on=['Pile'])
df3 = pd.merge(df3, df_WV[['Pile','WV']], how='left', on=['Pile'])
df3 = pd.merge(df3, df_EX[['Pile','EX']], how='left', on=['Pile'])
df3 = pd.merge(df3, df_EY[['Pile','EY']], how='left', on=['Pile'])

#for lateral loading output
df_bp_out = df_bp.ix[:,['Pile','OutputCase','Fx','Fy','Fz','Mx','My','Mz']]
df_bp_out.index +=1

writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.

df2.index +=1

df2.to_excel(writer, sheet_name='Axial Load(BP)')

df3.index +=1
df3.to_excel(writer, sheet_name='Lateral Load (BP)')

df_bp2_out.to_excel(writer, sheet_name='Lateral Load Sum')

df_bp_lateral.rename(columns = {'Node':'Node1','Pile':'Node'}, inplace = True)
df_bp_lateral=df_bp_lateral.sort_values(['Node1','OutputCase'], ascending=[True,True]).reset_index(drop=True)

writer = pd.ExcelWriter(output_file_to_Prokon, engine='xlsxwriter')

df_bp_lateral.to_excel(writer, sheet_name='Sorted')
writer.save()
