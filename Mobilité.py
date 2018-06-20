import pandas as pd
from Datatransform import *
import datetime as dt
from datetime import datetime


def comte_change(source,key,champ_change,beg_champ =None):
    if isinstance(source ,pd.DataFrame) is False:
        raise TypeError
    result = pd.DataFrame(columns=[key, "result",beg_champ])
    ref_key = ""
    ref_change = ""
    pos = 0

    for i, val in enumerate(source[key]):

        if ref_key == val and ref_change != source.at[i, champ_change]:
            result.at[pos,key]= val
            result.at[pos,"result"] = 1
            result.at[pos,beg_champ] = source.at[i,beg_champ]
            ref_change = source.at[i, champ_change]
            pos += 1

        if ref_key == "":
            ref_key = val

        if ref_key != val:
            ref_key = val
            ref_change = source.at[i, champ_change]

    return result

def convert_date_time(source, champ):
    source[champ] = source[champ].astype('str')
    source[champ] = source[champ].apply(lambda x: x.replace(".", "/"))
    source[champ] = source[champ].apply(lambda x: datetime.strptime(x, '%d/%m/%Y'))
    #source[champ] = pd.to_datetime(source[champ], errors='coerce')

source = pd.DataFrame()
res_int = pd.DataFrame()
faits = pd.DataFrame()
mobility_exit = pd.DataFrame(columns=["Mat.","Date déb.","Mobility_exit"])
"Chargement des données"
source = pd.read_csv(r'D:\Users\sgasmi\Desktop\PA0001.csv', delimiter=";", low_memory=False)


"Création de la clé unique"
source["Clé"] = source["Date déb."].astype('str')
source["Clé"] = source["Clé"].apply(lambda x: x.replace(".",""))
source["Clé"] = source["Mat."].astype('str') + source["Clé"]
source["Clé"] = source["Clé"].astype('int64')

"conversion date de début en temps"
convert_date_time(source, "Date déb.")

"Préparation des données"
source = source.sort_values(["Mat.","CSté","Date déb."])
mobility = comte_change(source,"Mat.","CSté","Date déb.")
#convert_date_time(mobility,"Date déb.")

"Mobilité sorties"
"""pos = 0
for i, val in enumerate(mobility):
    mobility_exit.at[pos, "Mat."] = mobility.at[pos, "Mat."]
    mobility_exit.at[pos, "Date déb."] = mobility.at[pos, "Date déb."] - 1"""

mobility_exit["Mat."] = mobility["Mat."]
mobility_exit["Date déb."] = mobility["Date déb."] + dt.timedelta(days=-1)
mobility_exit["Mobility_exit"] = mobility["result"]



"Alimentation de la table de faits"
faits["Mat."] = mobility["Mat."]
faits["Date déb."] = mobility["Date déb."]
faits["Mobility Entry"] = mobility["result"]
faits = faits.merge(mobility_exit,  how='outer')
print(mobility)

#source = pd.merge(source, mobility, on=["Mat."], how ='outer')
#source = source.merge(mobility, left_on=["Mat.","Date déb."], right_on=["Mat.","Date déb."], how='outer')
writer = pd.ExcelWriter(r'D:\Users\sgasmi\Desktop\mob.xlsx', engine='xlsxwriter')
faits.to_excel(writer, sheet_name="res")
writer.save()

#print(analyse.source)
#analyse.Entre_groupe("Entrées_Groupe")
#analyse.Entre_soc("Entrée_société")
#analyse.exportexcel(r'D:\Users\sgasmi\Desktop\Mobilité.xlsx','res')


