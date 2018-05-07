import pandas as pd
import numpy as np
from datetime import datetime

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

    def exportexcel(self, repertoire, onglet):
        """Méthode permettant d'extraire le résultat sous forme d'un fichier excel"""
        writer = pd.ExcelWriter(repertoire, engine='xlsxwriter')
        self.result.to_excel(writer, sheet_name=onglet)
        writer.save()

    def transformation(self):

        res_aff = pd.DataFrame() # contient le distinct mat/soc/date debut
        a = self.affectation
        val = ["Z","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
        a['Test'] = a['Domaine'].apply(lambda x: "ko" if x[1]in val else 'ok')
        b = a[a.Test == "ok"]

       #a = a.query('Domaine=="A"')
        #val = ['F2',"F1"]
        #a = a[a['Domaine'].str.match(val)]

        res_aff = b[['Mat.', 'CSté', 'Date déb.']].drop_duplicates()


        #res_aff = self.affectation[['Mat.', 'CSté', 'Date déb.']].drop_duplicates() #Recup des champs utiles dans affectation
        res_aff['Date déb.'] = pd.to_datetime(res_aff['Date déb.'], errors='coerce')
        res_soc = pd.DataFrame() #contient que les val embauche, reembauche
        list_mat = pd.Series

        list_mat = self.measure['Mat.'].drop_duplicates() # liste unique des matricules

        res_aff = res_aff.sort_values(by=['Mat.', 'CSté', 'Date déb.'], ascending=False)

        res_aff = res_aff.sort_values('Date déb.').groupby(['Mat.', 'CSté'], as_index=False).min()
        res_aff = res_aff.sort_values('Date déb.').groupby(['Mat.'], as_index=False).max()

        writer = pd.ExcelWriter(r"D:\Users\sgasmi\Desktop\affectation_distinct.xlsx", engine='xlsxwriter')
        res_aff.to_excel(writer, sheet_name="Res")
        writer.save()
        print("Ok")
        res_soc = self.measure




analyse = Imp_soc()
analyse.chargement("affectation", r"D:\Users\sgasmi\Desktop\Copie de exportv12589.xlsx", "Sheet1")
analyse.chargement("mesure", r"D:\Users\sgasmi\Desktop\PA0000_2304 (mesure).xlsx", "PA0000_2304")
#analyse.chargement("affectation", r"C:\Users\Sabri.GASMI\Desktop\Copie de PA0001 (affectation) - Copie.xlsx", "PA0001")
analyse.transformation()

#analyse.exportexcel(r"D:\Users\sgasmi\Desktop\res.xlsx","Resultat")
analyse.exportexcel(r"D:\Users\sgasmi\Desktop\leres.xlsx","Resultat")