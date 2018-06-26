import pandas as pd
from Datatransform import *
import datetime as dt
from datetime import datetime


def comte_change(source,key,champ_change,beg_champ,name_res=None):
    if isinstance(source ,pd.DataFrame) is False:
        raise TypeError
    if name_res is None:
        name_res ='result'

    result = pd.DataFrame(columns=[key,beg_champ,name_res])
    ref_key = ""
    ref_change = ""
    pos = 0

    for i, val in enumerate(source[key]):

        if ref_key == val and ref_change != source.at[i, champ_change]:
            result.at[pos,key]= val
            result.at[pos,name_res] = 1
            result.at[pos,beg_champ] = source.at[i,beg_champ]
            ref_change = source.at[i, champ_change]
            pos += 1

        if ref_key == "":
            ref_key = val
            ref_change = source.at[i, champ_change]

        if ref_key != val:
            ref_key = val
            ref_change = source.at[i, champ_change]

        result[beg_champ] = pd.to_datetime(result[beg_champ])
    #chang_fonct["Date déb."] = pd.to_datetime(chang_fonct["Date déb."])
    return result

def convert_date_time(source, champ):
    source[champ] = source[champ].astype('str')
    source[champ] = source[champ].apply(lambda x: x.replace(".", "/"))
    source[champ] = source[champ].apply(lambda x: datetime.strptime(x, '%d/%m/%Y'))
    source[champ] = pd.to_datetime(source[champ], errors='coerce')

"Dataframes stockants les résultats"
source = pd.DataFrame()
res_int = pd.DataFrame()
faits = pd.DataFrame()
ent_group = pd.DataFrame()
mobility_exit = pd.DataFrame(columns=["Mat."])
chang_fonct = pd.DataFrame()

"Chargement des données"
source = pd.read_csv(r'D:\Users\sgasmi\Desktop\PA0001.csv', delimiter=";", low_memory=False)

source_civil = pd.read_csv(r'D:\Users\sgasmi\Desktop\Civil.csv', delimiter=";", low_memory=False)

"Création de la clé unique"
source["Clé"] = source["Date déb."].astype('str')
source["Clé"] = source["Clé"].apply(lambda x: x.replace(".",""))
source["Clé"] = source["Mat."].astype('str') + source["Clé"]
source["Clé"] = source["Clé"].astype('int64')

source_civil['Clé_civil'] = source_civil["Date déb."].astype('str')
source_civil['Clé_civil'] = source_civil["Clé_civil"].apply(lambda x: x.replace(".",""))
source_civil['Clé_civil'] = source_civil["Mat."].astype('str') + source_civil['Clé_civil']
source_civil['Clé_civil'] = source_civil['Clé_civil'].astype('int64')

"conversion date de début en temps"
convert_date_time(source, "Date déb.")
convert_date_time(source, "Fin valid.")
date_fin = datetime.strptime("01/01/2030",'%d/%m/%Y')
source['Fin valid.'] = source["Fin valid."].apply(lambda x: date_fin if str(x) == "NaT" else x)
convert_date_time(source_civil,"Date déb.")
convert_date_time(source_civil,"Fin valid.")
source_civil['Fin valid.'] = source_civil['Fin valid.'].apply(lambda x: date_fin if str(x) == "NaT" else x)

"Préparation des données"
source = source.sort_values(["Mat.","CSté","Date déb."])
mobility = comte_change(source,"Mat.","CSté","Date déb.")
mobility_exit["Mat."] = mobility["Mat."]
mobility_exit["Date déb."] = mobility["Date déb."] + dt.timedelta(days=-1)
mobility_exit["Mobility_exit"] = mobility["result"]
ent_group = source.sort_values('Date déb.').groupby(['Mat.'], as_index=False).min()
ent_group["Entrée_Groupe"] =1
ent_group = ent_group[["Mat.", "Date déb.", "Entrée_Groupe"]]
chang_fonct = comte_change(source,"Mat.","Fonction","Date déb.","Chang_fonct")


"Alimentation de la table de faits"
faits["Mat."] = mobility["Mat."]
faits["Date déb."] = mobility["Date déb."]
faits["Mobility Entry"] = mobility["result"]
faits = faits.merge(mobility_exit,on=["Mat.","Date déb."],  how='outer')
faits = faits.merge(ent_group,on=["Mat.","Date déb."], how='outer')
faits = faits.merge(chang_fonct, on=["Mat.","Date déb."], how='outer')

"Intégration de la clé technique dans la table de faits"
def int_ext_key(faits,source,cle_name):
    fusion =  pd.merge(faits,source[['Mat.', cle_name, 'Date déb.','Fin valid.']], on='Mat.', how='left')
    fusion["Date_fait"] = fusion["Date déb._x"]
    fusion["Début_dim"] = fusion["Date déb._y"]
    fusion["Fin_dim"] = fusion["Fin valid."]
    fusion = fusion.query('Date_fait >= Début_dim and Date_fait <= Fin_dim')
    fusion["Date déb."] = fusion["Date_fait"]
    fusion = fusion[["Mat.", "Date déb.", cle_name]]
    fusion["Date déb."] = pd.to_datetime(fusion["Date déb."])
    return faits.merge(fusion, on=["Mat.", "Date déb."], how="left")

faits = int_ext_key(faits,source,"Clé")
faits = int_ext_key(faits, source_civil,"Clé_civil" )

writer = pd.ExcelWriter(r'D:\Users\sgasmi\Desktop\mob.xlsx', engine='xlsxwriter')
source_civil.to_excel(writer, sheet_name="res")
writer.save()

print("ok")


