import pandas as pd 
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import os,sys
import openpyxl


'''
def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

os.chdir(sys.path[0])


template_path = get_resource_path("report_template_two.docx")
excel_path = get_resource_path("store_data.xlsx")
'''
template_path = input("Please enter the path to the Template file (store_data.xlsx): ")
excel_path = input("Please enter the path to the Excel file (store_data.xlsx): ")


doc = DocxTemplate(template_path)
df = pd.read_excel(excel_path)

row=df.loc[1]
value=df.at[0,'Solar kWhr Savings']
#print(value)
val2=df.at[1, 'Solar kWhr Savings']
#print(val2)

#print("the value var is of type:", type(value))
#print(row)
#print(df.columns[0])
#print(df.shape)
#print(df)

date=df.columns[0]
total_kWh=('{:,}'.format(value)) 
bldg_kWh=('{:,}'.format(val2)) 
perg_kWh=('{:,}'.format(df.at[2,'Solar kWhr Savings'])) 

totp=df.at[0,'% of Expected']
total_perc=str(int(totp*100))+'%'
bldp=df.at[1,'% of Expected']
bldg_perc=str(int(bldp*100))+'%'
pergp=df.at[2,'% of Expected']
perg_perc=str(round(pergp*100))+'%'
#print(total_perc+","+bldg_perc+","+perg_per

tdl=df.at[0,'Monthly Dollar Savings']
bdl=df.at[1,'Monthly Dollar Savings']
pdl=df.at[2,'Monthly Dollar Savings']
t_dlr="$"+str(('{:,}'.format(tdl)))
b_dlr="$"+str(('{:,}'.format(bdl)))
p_dlr="$"+str(('{:,}'.format(pdl)))

tpe=df.at[0,'% Electricity Provided by Solar']
bpe=df.at[1,'% Electricity Provided by Solar']
ppe=df.at[2,'% Electricity Provided by Solar']
t_esol=str(round(tpe*100,2))+"%"
b_esol=str(round(bpe*100,2))+"%"
p_esol=str(round(ppe*100,2))+"%"

tco=df.at[0,'CO2 Savings Metric Tons']
bco=df.at[1,'CO2 Savings Metric Tons']
pco=df.at[2,'CO2 Savings Metric Tons']



context={
    'month':date,
    'totalkwhr': total_kWh, 
    'mainbldkwhr': bldg_kWh, 
    'pergolakwhr': perg_kWh,
    'perc_tot': total_perc, 
    'perc_main': bldg_perc, 
    'perc_perg': perg_perc,
    'dollar_tot': t_dlr,
    'dollar_bld': b_dlr,
    'dollar_perg': p_dlr,
    'perc_sol_tot': t_esol, 
    'perc_sol_bldg': b_esol, 
    'perc_sol_perg': p_esol,
    'CO_tot': tco, 
    'CO_bldg': bco, 
    'CO_perg': pco,
}
doc.render(context)

outputpath= input("Please enter the path to the location where you wish to store new file: ")
doc.save(outputpath)






"""
context={ 
        'month':date,
        'totalkwhr': total_kWh, 
        'mainbldkwhr': bldg_kWh, 
        'pergolakwhr': perg_kWh,
        'perc_tot': total_perc, 
        'perc_main': bldg_perc, 
        'perc_perg': perg_perc,
        'dollar_tot': t_dlr,
        'dollar_bld': b_dlr,
        'dollar_perg': p_dlr,
        'perc_sol_tot': t_esol, 
        'perc_sol_bldg': b_esol, 
        'perc_sol_perg': p_esol,
        'CO_tot': t_co, 
        'CO_bldg': b_co, 
        'CO_perg': p_co, 
    }

    doc.render(context)
    doc.save("exceltoreport2.docx")
"""