import pandas as pd
from Datatransform import Changedata


analyse = Changedata()
analyse.chargement(r'D:\Users\sgasmi\Desktop\Données maquette\Datamobilité - Copie', 'monres')

print("ok")
#print(pd.unique(analyse.source["Matricule"]))


def chargement(self, chemin, feuille=""):
    if chemin[len(chemin) - 4:len(chemin)] == "xlsx" and feuille != "":
        Excel = pd.ExcelFile(chemin)
        self.source = Excel.parse(feuille)

chargement()