import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook
df = pd.read_excel('delay.xlsx')
stage_delay_df = df[(df['Parameter']=='stage_delay')]
iddq_stage_df =  df[(df['Parameter']=='iddq_stage')]
excel_file = 'output.xlsx'
import xlsxwriter

df_output=[]

def test(param):
    fil_df = df[(df['Parameter']==param)]
    sorted_df = fil_df.sort_values(by=['Temp','ProcessCorner'])
    sorted_df['target_wval']=None
    sorted_df['reference_wval']=None
    sorted_df['percentage_change']=None

    def highlight_rows(row):
        if 'INV' in row['Reference_Key']:  
            return ['background-color: yellow'] * len(row)
        else:
            return ['background-color: white'] * len(row)


    def write_data(filtered_df):
        #target weighted value calculation
        tar_val=filtered_df['TargetValue']
        tar_val_list=tar_val.to_list()
        tar_w_val=tar_val_list[0]*0.4
        for val in tar_val_list[1:]:
            tar_w_val=tar_w_val+0.1*val

        row_numbers = filtered_df.index.to_list()
        sorted_df.loc[row_numbers[0],'target_wval']=tar_w_val
        if param=='stage_delay':
            tar_freq=1/tar_w_val
            freq_loc=row_numbers[1]
            sorted_df.loc[freq_loc,'target_wval']=tar_freq

        #reference weighted value calculation
        ref_val=filtered_df['ReferenceValue']
        ref_val_list=ref_val.to_list()
        ref_w_val=ref_val_list[0]*0.4
        
        for val in ref_val_list[1:]:
            ref_w_val=ref_w_val+0.1*val

        sorted_df.loc[row_numbers[0],'reference_wval']=ref_w_val
        if param=='stage_delay':
            ref_freq=1/ref_w_val
            sorted_df.loc[freq_loc,'reference_wval']=ref_freq

        per_delay_change=((tar_w_val-ref_w_val)/ref_w_val)*100
        sorted_df.loc[row_numbers[0],'percentage_change']=per_delay_change
        if param=='stage_delay':
            per_freq_change=((tar_freq-ref_freq)/ref_freq)*100
            sorted_df.loc[freq_loc,'percentage_change']=per_freq_change


    filtered_df = sorted_df[(sorted_df['PexCorner'] =='Nominal') & (sorted_df['Temp'] == 25) & (sorted_df['ProcessCorner']=='TT')]
    write_data(filtered_df)
    filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmin') & (sorted_df['Temp'] == -40) & (sorted_df['ProcessCorner']=='FFG')]
    write_data(filtered_df)
    filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmin') & (sorted_df['Temp'] == 125) & (sorted_df['ProcessCorner']=='FFG')]
    write_data(filtered_df)
    filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmax') & (sorted_df['Temp'] == -40) & (sorted_df['ProcessCorner']=='SSG')]
    write_data(filtered_df)
    filtered_df = sorted_df[(sorted_df['PexCorner'] =='FuncCmax') & (sorted_df['Temp'] == 125) & (sorted_df['ProcessCorner']=='SSG')]
    write_data(filtered_df)

    styled_df= sorted_df.style.apply(highlight_rows, axis=1)
    df_output.append(styled_df)

    


test("stage_delay")
test("iddq_stage")

writer = pd.ExcelWriter(excel_file, engine='xlsxwriter', mode='w')
df_output[0].to_excel(writer, sheet_name='stage_delay', index=False)
df_output[1].to_excel(writer, sheet_name='iddq_stage', index=False)
writer.close()

