import pandas as pd
import numpy as np

class Imp_soc:

    def __init__(self):
        self.measure = pd.DataFrame
        self.affectation = pd.DataFrame
        self.result = pd.DataFrame(columns=["Matricule", "Date_debut"])

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

    def transformation(self):

        res_aff = pd.DataFrame() # contient le distinct mat/soc/date debut
        res_soc = pd.DataFrame() #contient que les val embauche, reembauche
        list_mat = pd.Series

        list_mat = self.affectation['Mat.'].drop_duplicates()

        res_aff = self.affectation[['Mat.', 'CSté', 'Date déb.']].drop_duplicates()
        res_aff['Date déb.']= pd.to_datetime(res_aff['Date déb.'], errors='coerce')
        res_soc = self.measure.loc[self.measure['Mes.'].isin([10, 0])]

        res_aff = res_aff.sort_values(by=['Mat.','CSté', 'Date déb.'], ascending=False)

        g1 = res_aff.sort_values('Date déb.').groupby(['Mat.','CSté'], as_index=False).min()

        print(g1)
        """for i, cel in enumerate(self.result):
            date_cible = self.source.at[i,'Date déb.']
            for y, val in enumerate(res_aff):
                if self.source.at[i, 'Mat.'] == res_aff.at[y, 'Mat.']:
                    a ="""


analyse = Imp_soc()
analyse.chargement("mesure", r"D:\Users\sgasmi\Desktop\PA0000_2304 (mesure).xlsx", "PA0000_2304")
analyse.chargement("affectation", r"D:\Users\sgasmi\Desktop\PA0001 (affectation).xlsx", "PA0001")
analyse.transformation()