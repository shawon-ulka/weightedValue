import pandas as pd
df = pd.read_excel('delay.xlsx')
stage_delay_df = df[(df['Parameter']=='stage_delay')]
iddq_stage_df =  df[(df['Parameter']=='iddq_stage')]
sorted_df = stage_delay_df.sort_values(by=['Temp','ProcessCorner'])


def highlight_rows(row):
    if 'INV' in row['Reference_Key']:  
        return ['background-color: yellow'] * len(row)
    else:
        return ['background-color: white'] * len(row)

sorted_df['target_wval']=None
sorted_df['reference_wval']=None
sorted_df['percentage_change']=None

def write_data(filtered_df):
    #target weighted value calculation
    tar_val=filtered_df['TargetValue']
    tar_val_list=tar_val.to_list()
    tar_w_val=tar_val_list[0]*0.4
    for val in tar_val_list[1:]:
        tar_w_val=tar_w_val+0.1*val
    tar_freq=1/tar_w_val
    row_numbers = filtered_df.index.to_list()
    sorted_df.loc[row_numbers[0],'target_wval']=tar_w_val
    freq_loc=row_numbers[1]
    sorted_df.loc[freq_loc,'target_wval']=tar_freq

    #reference weighted value calculation
    ref_val=filtered_df['ReferenceValue']
    ref_val_list=ref_val.to_list()
    ref_w_val=ref_val_list[0]*0.4
    
    for val in ref_val_list[1:]:
        ref_w_val=ref_w_val+0.1*val

    ref_freq=1/ref_w_val
    sorted_df.loc[row_numbers[0],'reference_wval']=ref_w_val
    sorted_df.loc[freq_loc,'reference_wval']=ref_freq

    per_delay_change=((tar_w_val-ref_w_val)/ref_w_val)*100
    per_freq_change=((tar_freq-ref_freq)/ref_freq)*100

    sorted_df.loc[row_numbers[0],'percentage_change']=per_delay_change
    sorted_df.loc[freq_loc,'percentage_change']=per_freq_change


#for nominal tt corner
filtered_df = sorted_df[(sorted_df['PexCorner'] =='Nominal') & (sorted_df['Temp'] == 25) & (sorted_df['ProcessCorner']=='TT')]
write_data(filtered_df)
#for funcmin -40 FF
filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmin') & (sorted_df['Temp'] == -40) & (sorted_df['ProcessCorner']=='FFG')]
write_data(filtered_df)
#for funcmin 125 FF
filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmin') & (sorted_df['Temp'] == 125) & (sorted_df['ProcessCorner']=='FFG')]
write_data(filtered_df)
# #for FuncCmax -40 SS
filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmax') & (sorted_df['Temp'] == -40) & (sorted_df['ProcessCorner']=='SSG')]
write_data(filtered_df)
# #for FuncCmax 125 SS
filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmax') & (sorted_df['Temp'] == 125) & (sorted_df['ProcessCorner']=='SSG')]
write_data(filtered_df)

styled_df = sorted_df.style.apply(highlight_rows, axis=1)
file_path = 'output.xlsx'
styled_df.to_excel(file_path,sheet_name="stage_delay", index=False)
