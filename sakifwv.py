import pandas as pd
import xlsxwriter
readFile='LVT_C24_IP007325.xlsx'
writeFile="Output_LVT_C24_IP007325.xlsx"

def checktargetInv(dct,valType):
    if "inv"in dct['IP_Name'].lower():
        return dct[valType]*.4
    else:
        return dct[valType]*.1

df=pd.read_excel(readFile)
# comparison =df[1:]
filtered=df[["IP_Name","Tech","VT","Track","Parameter","PexCorner" ,"PexTemp","Unit","TargetValue","ReferenceValue"]]
filtered=filtered.sort_values(by=['Parameter',"PexTemp","PexCorner","VT"])
delay=filtered[(filtered["Parameter"]=="stage_delay")]
leakage=filtered[(filtered["Parameter"]=="iddq_stage")]
delay["wieghted_Target"]=delay.apply(lambda x :checktargetInv(x,"TargetValue"),axis=1)
delay["wieghted_Reference"]=delay.apply(lambda x :checktargetInv(x,"ReferenceValue"),axis=1)
leakage["wieghted_Target"]=leakage.apply(lambda x :checktargetInv(x,"TargetValue"),axis=1)
leakage["wieghted_Reference"]=leakage.apply(lambda x :checktargetInv(x,"ReferenceValue"),axis=1)

#leakage & delay grouping 
leakage_sum=leakage.groupby(['PexCorner','PexTemp']).sum()
leakage_sum['newChange']=(leakage_sum['wieghted_Target']-leakage_sum['wieghted_Reference'])*100/leakage_sum['wieghted_Reference']
delay_sum=delay.groupby(['PexCorner','PexTemp']).sum()
delay_sum['Ref_freq']=1/delay_sum['wieghted_Reference']
delay_sum['targ_freq']=1/delay_sum['wieghted_Target']
delay_sum['freq_change']=(delay_sum['targ_freq']-delay_sum['Ref_freq'])*100/delay_sum['Ref_freq']
leakage_sum=leakage.groupby(['PexCorner','PexTemp']).sum()
delay_sum['newChange']=(delay_sum['wieghted_Target']-delay_sum['wieghted_Reference'])*100/delay_sum['wieghted_Reference']
delay_sum.drop(delay_sum.loc[:,'Track':'ReferenceValue'].columns,axis=1,inplace=True)
delay_sum=delay_sum.transpose()
leakage_sum['newChange']=(leakage_sum['wieghted_Target']-leakage_sum['wieghted_Reference'])*100/leakage_sum['wieghted_Reference']
leakage_sum['newChange']=(leakage_sum['wieghted_Target']-leakage_sum['wieghted_Reference'])*100/leakage_sum['wieghted_Reference']
leakage_sum=leakage_sum.transpose()

#write file
writer=pd.ExcelWriter(writeFile)
leakage.to_excel(writer,sheet_name="leakage")
delay.to_excel(writer,sheet_name="delay")
leakage_sum.to_excel(writer,sheet_name="leakage_avg")
delay_sum.to_excel(writer,sheet_name="Delay_avg")
writer.close()