import pandas as pd
from datetime import datetime
from xlsxwriter import *
import numpy as np

class Extract:

    def __init__(self):
        self.source = pd.DataFrame() #source de données
        self.result = pd.DataFrame() #resultat

    def chargement_excel(self, chemin, feuille=""):
        if chemin[len(chemin) - 4:len(chemin)] == "xlsx" and feuille != "":
            Excel = pd.ExcelFile(chemin)
            self.source = Excel.parse(feuille)

    def chargement_csv(self,chemin):
        if chemin[len(chemin) - 3:len(chemin)] == "csv":
            self.source = pd.read_csv(chemin)

    def export_excel(self, repertoire, onglet):
        """Méthode permettant d'extraire le résultat sous forme d'un fichier excel"""
        writer = pd.ExcelWriter(repertoire, engine='xlsxwriter')
        self.source.to_excel(writer, sheet_name=onglet)
        writer.save()

    def extract_line(self, date_cible):
        self.source.sort_values(By="Matricule")
        #for i, cel in enumerate(self.source.itertuples()):


Mon_extract = Extract()
Mon_extract.chargement_csv(r"D:\Users\sgasmi\Desktop\export.csv")
print(Mon_extract.source)
#Mon_extract.chargement(r"D:\Users\sgasmi\Desktop\Export.xlsx","Sheet1")
