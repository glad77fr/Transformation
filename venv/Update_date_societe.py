import pandas as pd
import numpy as np
from datetime import datetime

class Imp_soc:

    def __init__(self):
        self.measure = pd.DataFrame
        self.affectation = pd.DataFrame
        self.result = pd.DataFrame(columns=["Matricule", "Date_debut", "Type_cas"])

    def chargement(self, fichier, chemin, feuille=""):

        try:
            if chemin[len(chemin) - 4:len(chemin)] == "xlsx" and feuille != "":
                if fichier == "mesure":
                    Excel = pd.ExcelFile(chemin)
                    self.measure = Excel.parse(feuille)
                if fichier == "affectation":
                    Excel = pd.ExcelFile(chemin)
                    self.affectation = Excel.parse(feuille)
        except:
            print("Fichier ou feuil inexistant")

    def exportexcel(self, repertoire, onglet):
        """Méthode permettant d'extraire le résultat sous forme d'un fichier excel"""
        writer = pd.ExcelWriter(repertoire, engine='xlsxwriter')
        self.result.to_excel(writer, sheet_name=onglet)
        writer.save()

    def transformation(self):

        res_aff = pd.DataFrame() # contient le distinct mat/soc/date debut
        res_aff = self.affectation[['Mat.', 'CSté', 'Date déb.']].drop_duplicates() #Recup des champs utiles dans affectation
        res_aff['Date déb.'] = pd.to_datetime(res_aff['Date déb.'], errors='coerce')
        res_soc = pd.DataFrame() #contient que les val embauche, reembauche
        list_mat = pd.Series

        list_mat = self.measure['Mat.'].drop_duplicates() # liste unique des matricules

        res_aff = res_aff.sort_values(by=['Mat.', 'CSté', 'Date déb.'], ascending=False)

        res_aff = res_aff.sort_values('Date déb.').groupby(['Mat.', 'CSté'], as_index=False).min()

        res_soc = self.measure

        for i, cel in enumerate(list_mat): # alimentation des resultats
            self.result.at[i, 'Matricule'] = cel

        for i, mat in enumerate(self.result["Matricule"]):
            dateref = "vide"
            date_cible =""
            if mat in res_aff["Mat."].values:
                for y, val in enumerate(res_aff.itertuples()):
                    if self.result.at[i,"Matricule"] == res_aff.at[y,"Mat."]:

                        if dateref != "vide":
                            if (dateref - res_aff.at[y, "Date déb."]).days > 0:
                                #date_cible = datetime.strptime(str(res_aff.at[y,"Date déb."]),'%Y-%m-%d %H:%M:%S')
                                date_cible = datetime.strptime(str(dateref), '%Y-%m-%d %H:%M:%S')
                        if dateref == "vide":
                            dateref = datetime.strptime(str(res_aff.at[y,"Date déb."]),'%Y-%m-%d %H:%M:%S')

            self.result.at[i,'Date_debut'] = date_cible
            self.result.at[i,'Type_cas'] = "2"


        for i, mat in enumerate(self.result["Matricule"]):
            for y, val in enumerate(res_soc.itertuples()):

                if self.result.at[i,'Matricule'] == res_soc.at[y,"Mat."]:
                    if self.result.at[i,'Date_debut'] == "" and res_soc.at[y, 'Mes.'] =="00":
                        self.result.at[i,'Date_debut'] =res_soc.at[y,'Date déb.']
                        break


analyse = Imp_soc()
analyse.chargement("mesure", r"D:\Users\sgasmi\Desktop\PA0000_2304 (mesure).xlsx", "PA0000_2304")
analyse.chargement("affectation", r"D:\Users\sgasmi\Desktop\PA0001 (affectation).xlsx", "PA0001")
analyse.transformation()

analyse.exportexcel(r"D:\Users\sgasmi\Desktop\res.xlsx","Resultat")