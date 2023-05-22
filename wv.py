import pandas as pd
df = pd.read_excel('delay.xlsx')
column_names = df.columns
print(column_names)
stage_delay_df = df[(df['Parameter']=='stage_delay')]
file_path = 'test.xlsx'
stage_delay_df.to_excel(file_path, index=False)


dfnew = pd.read_excel('test.xlsx')
sorted_df = dfnew.sort_values(by=['Temp','ProcessCorner'])


def highlight_rows(row):
    if 'INV' in row['Reference_Key']:  
        return ['background-color: yellow'] * len(row)
    else:
        return ['background-color: white'] * len(row)

def write_data(dataFrame):
    print("hello")
#for nominal tt corner
filtered_df = sorted_df[(sorted_df['PexCorner'] =='Nominal') & (sorted_df['Temp'] == 25) & (sorted_df['ProcessCorner']=='TT')]
#target weighted value calculation
tar_val=filtered_df['TargetValue']
tar_val_list=tar_val.to_list()
w_val=tar_val_list[0]*0.4
for val in tar_val_list[1:]:
    w_val=w_val+0.1*val

row_numbers = filtered_df.index.to_list()
sorted_df['target_wval']=None
sorted_df.loc[row_numbers[0],'target_wval']=w_val
freq_loc=row_numbers[1]
sorted_df.loc[freq_loc,'target_wval']=1/w_val

#reference weighted value calculation
ref_val=filtered_df['ReferenceValue']
ref_val_list=ref_val.to_list()
w_val=ref_val_list[0]*0.4
for val in ref_val_list[1:]:
    w_val=w_val+0.1*val

row_numbers = filtered_df.index.to_list()
sorted_df['reference_wval']=None
sorted_df.loc[row_numbers[0],'reference_wval']=w_val
freq_loc=row_numbers[1]
sorted_df.loc[freq_loc,'reference_wval']=1/w_val


#for funcmin -40 FF
filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmin') & (sorted_df['Temp'] == -40) & (sorted_df['ProcessCorner']=='FFG')]
#target weighted value calculation
tar_val=filtered_df['TargetValue']
tar_val_list=tar_val.to_list()
w_val=tar_val_list[0]*0.4
for val in tar_val_list[1:]:
    w_val=w_val+0.1*val

row_numbers = filtered_df.index.to_list()
sorted_df.loc[row_numbers[0],'target_wval']=w_val
freq_loc=row_numbers[1]
sorted_df.loc[freq_loc,'target_wval']=1/w_val

#reference weighted value calculation
ref_val=filtered_df['ReferenceValue']
ref_val_list=ref_val.to_list()
w_val=ref_val_list[0]*0.4
for val in ref_val_list[1:]:
    w_val=w_val+0.1*val

row_numbers = filtered_df.index.to_list()
sorted_df.loc[row_numbers[0],'reference_wval']=w_val
freq_loc=row_numbers[1]
sorted_df.loc[freq_loc,'reference_wval']=1/w_val




styled_df = sorted_df.style.apply(highlight_rows, axis=1)
file_path = 'output.xlsx'
styled_df.to_excel(file_path, index=False)
