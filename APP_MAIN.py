import pandas as pd
import numpy as np


# Importing data 

lt_df = pd.read_excel('D:/RMC_APP/DATA/LT_DB.xlsx')
ht_df = pd.read_excel('D:/RMC_APP/DATA/Ht_DB.xlsx')
Instrument_df = pd.read_excel('D:/RMC_APP/DATA/Instrument_Cable.xlsx')
FL_Y_df = pd.read_excel('D:/RMC_APP/DATA/Flexi_Cable_FL-Y.xlsx')
FL_YY_df = pd.read_excel('D:/RMC_APP/DATA/Flexi_Cable.xlsx')
IS_OS_df  = pd.read_excel('D:/RMC_APP/DATA/IS_OS.xlsx')


lt_df = lt_df.iloc[:,1:]
ht_df = ht_df.iloc[:,1:]
Instrument_df = Instrument_df.iloc[:,1:]
FL_Y_df = FL_Y_df.iloc[:,1:]
FL_YY_df = FL_YY_df.iloc[:,1:]



input_df = pd.read_excel('C:/Users/PC-71/Desktop/Input_Data.xlsx')

Table_Name = [lt_df,ht_df,Instrument_df,FL_Y_df,FL_YY_df,IS_OS_df]

lt_df_dic = {}
ht_df_dic = {}
Instrument_df_dic = {}
FL_Y_df_dic = {}
FL_YY_df_dic ={}
IS_OS_df_dic = {}

Table_List = tuple([lt_df_dic,ht_df_dic,Instrument_df_dic,FL_Y_df_dic,FL_YY_df_dic,IS_OS_df_dic])

for i in range(0,len(Table_Name)) :
    colums = list(Table_Name[i].columns)
    for col in colums:
        Table_List[i][col] = [0]



Instrument_Output = pd.DataFrame(Instrument_df_dic)
lt_output= pd.DataFrame(lt_df_dic)
ht_output= pd.DataFrame(ht_df_dic)
FL_Y_output= pd.DataFrame(FL_Y_df_dic)
FL_YY_output= pd.DataFrame(FL_YY_df_dic)
IS_OS_output = pd.DataFrame(IS_OS_df_dic)


def type_cable (elemet):
    lis_ele = elemet.split(' ')
    size = len(lis_ele)
    if size ==1:
        if lis_ele[0][-1] == "P" or lis_ele[0][-1] == "T" :
            return "OS"
        elif lis_ele[0] == "FL-Y":
            return "FL-Y"
        elif lis_ele[0] == "FL-YY":
            return "FL-YY"
        else:
            return "LT"

    elif size == 3:
        if lis_ele[-1] =="HT":
            return "HT"
    else:
        if lis_ele[-1] == 'OS':
            return "OS"
        elif lis_ele[-1] == "IS&OS":
          return "IS&OS"



def core_find (element):

    if element *1 == element :
        return "0"
    else :
        return"OS"



cable_type = []

for i in range (0,input_df.shape[0]):

    Type_1 = core_find(input_df.iloc[i,0])
    Type_2 = type_cable(input_df.iloc[i,-1])


    if Type_1 == "OS" or Type_2 == "OS":
        cable_type.append('OS')

    elif Type_2 == "IS&OS":
      cable_type.append('IS&OS')

    elif Type_2 == "LT":
        cable_type.append('LT')
    elif Type_2 == "FL-Y":
        cable_type.append('FL-Y')
    elif Type_2 == "FL-YY":
        cable_type.append('FL-YY')
    elif Type_2 == "HT":
        cable_type.append('HT')



    if Type_1 == "OS" or Type_2 == "OS":
        Core = input_df.loc[i,'Core']
        Cond = input_df.loc[i,'Cond']

        filter_process =  Instrument_df[Instrument_df['Core'] == Core]
        filter_process = filter_process[filter_process['Cond'] == Cond]


        Instrument_Output = pd.concat([Instrument_Output,filter_process], axis= 0)

    elif Type_2 == "IS&OS":
      Core = input_df.loc[i,'Core']
      Cond = input_df.loc[i,'Cond']

      Type_Insulation = input_df.loc[i,'Type'].split(' ')[0]

      if Type_Insulation[0] == "2":
        Type_Insulation = "XLPE"

      elif Type_Insulation[0] == "A" and Type_Insulation[1] == "2":
        Type_Insulation = "XLPE"

      else:
        Type_Insulation = "PVC"

      filter_process =  IS_OS_df[IS_OS_df['Core'] == Core]
      filter_process = filter_process[filter_process['Cond'] == Cond]
      filter_process = filter_process[filter_process['TYPE INSULATION'] == Type_Insulation]

      IS_OS_output = pd.concat([IS_OS_output,filter_process], axis= 0)

    elif Type_2 == "HT":

        breakup_lis = str(input_df.loc[i,'Type']).split(' ')
        Core = float(input_df.loc[i,'Core'])
        Cond = float(input_df.loc[i,'Cond'])
        Type = breakup_lis[0]
        KV = breakup_lis[1]
        lis = ht_df.KV_Type.unique().tolist()

        if KV in lis:
            break
        else:
            KV = str(breakup_lis[1].split('KV')[0])
            KV = float(KV)
            if KV >0 and KV <= 8:
                KV = '6.35/11'
            elif KV > 8 and KV < 12:
                KV = '11/11'
            elif KV >= 12  and KV <= 22 :
                KV = '12.7/22'
            else:
                KV = '19/33'

        filter_process =  ht_df[ht_df['Core'] == Core]
        filter_process = filter_process[filter_process['Cond'] == Cond]
        filter_process = filter_process[filter_process['Type'] == Type]
        filter_process = filter_process[filter_process['Type'] == Type]
        filter_process = filter_process[filter_process['KV_Type'] == KV]


        ht_output = pd.concat([ht_output,filter_process], axis= 0)

    elif Type_2 == "LT":

        Core = float(input_df.loc[i,"Core"])
        Cond = float(input_df.loc[i,"Cond"])
        Type = input_df.loc[i,"Type"]
        # Filter Process

        filter_process = lt_df[lt_df['Core'] == Core]
        filter_process= filter_process[filter_process['Cond'] == Cond ]
        filter_process= filter_process[filter_process['Type'] == Type]

        # In our data some ['Type'] has " STD" is added in the last for that we put this condition

        if filter_process.shape[0] == 0:
            filter_process = lt_df[lt_df['Core'] == Core]
            filter_process= filter_process[filter_process['Cond'] == Cond ]
            filter_process= filter_process[filter_process['Type'] == Type + " STD"]


            lt_output = pd.concat([lt_output, filter_process], axis=0)

        elif filter_process.shape[0] == 1:
            lt_output = pd.concat([lt_output, filter_process], axis=0)

        # In our data some have multipal data so we put first row by this condtion

        else:
            filter_process = filter_process.iloc[0:1,:]
            lt_output = pd.concat([lt_output, filter_process], axis=0)

    elif Type_2 == "FL-Y":
        Core = float(input_df.loc[i,"Core"])
        Cond = float(input_df.loc[i,"Cond"])
        Type = input_df.loc[i,"Type"]

        filter_process =  FL_Y_df[FL_Y_df['CORE'] == Core]
        filter_process = filter_process[filter_process['Cond'] == Cond]
        filter_process = filter_process[filter_process['TYPE'] == Type]

        FL_Y_output = pd.concat([FL_Y_output,filter_process], axis= 0)

    elif Type_2 == "FL-YY" :
        Core = float(input_df.loc[i,"Core"])
        Cond = float(input_df.loc[i,"Cond"])
        Type = input_df.loc[i,"Type"]

        filter_process =  FL_YY_df[FL_YY_df['CORE'] == Core]
        filter_process = filter_process[filter_process['Cond'] == Cond]
        filter_process = filter_process[filter_process['TYPE'] == Type]

        FL_YY_output = pd.concat([FL_YY_output,filter_process], axis= 0)


cable_type = list(set(cable_type))

for i in  cable_type:
    if i == "LT":
      lt_output = lt_output.iloc[1:,:]
      lt_output.to_excel('C:/Users/PC-71/Desktop/OUTPUT/lt_output.xlsx', index=False)
    elif i == "HT":
      ht_output = ht_output.iloc[1:,:]
      ht_output.to_excel('C:/Users/PC-71/Desktop/OUTPUT/ht_output.xlsx', index=False)
    elif i == 'OS':
      Instrument_Output = Instrument_Output.iloc[1:,:]
      Instrument_Output.to_excel('C:/Users/PC-71/Desktop/OUTPUT/Instrument_Output.xlsx', index=False)
    elif i == 'FL-Y':
      FL_Y_output = FL_Y_output.iloc[1:,:]
      FL_Y_output.to_excel('C:/Users/PC-71/Desktop/OUTPUT/FL_Y_output.xlsx', index=False)
    elif i == "FL-YY":
      FL_YY_output = FL_YY_output.iloc[1:,:]
      FL_YY_output.to_excel('C:/Users/PC-71/Desktop/OUTPUT/FL_YY_output.xlsx', index=False)
    elif i == "IS&OS":
      IS_OS_output= IS_OS_output.iloc[1:,:]
      IS_OS_output.to_excel('C:/Users/PC-71/Desktop/OUTPUT/IS_OS_output.xlsx', index=False)

print('ALL DONE')