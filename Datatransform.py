import pandas as pd
from xlsxwriter import *
from datetime import date

class Changedata:
    def __init__(self):
        self.source = pd.DataFrame()

    def chargement(self, chemin, feuille=""):
        if chemin[len(chemin)-4:len(chemin)] =="xlsx" and feuille != "":
            Excel = pd.ExcelFile(chemin)
            self.source = Excel.parse(feuille)

    def colonne(self):
            print(self.source.columns)

    def remplacement_valeurs(self, champ, dic):
        self.source.replace(dic, inplace=True)

    def tranche(self, champ_source, champ_cible,val_text, signe, val1, signe2=None, val2=None):
        if champ_source != champ_cible:
            test = champ_cible in self.source.columns
            if test is False:
                self.source[champ_cible] = self.source[champ_source]
                print("ok")

        if signe != "" and val1 != "" and signe2 is None:
            if signe == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x < val1 else x)
            if signe == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x <= val1 else x)
            if signe == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x > val1 else x)
            if signe == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x >= val1 else x)

        if signe !="" and signe2 is not None and val1 != "" and val2 is not None:
            if signe =="inf" and signe2 =="inf":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x < val1 and x < val2 else x)
            if signe =="inf" and signe2 =="inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x < val1 and x <= val2 else x)
            if signe =="inf" and signe2 =="sup":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x < val1 and x > val2 else x)
            if signe =="inf" and signe2 =="sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x < val1 and x >= val2 else x)
            if signe =="inf ou égal" and signe2 =="inf":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x <= val1 and x < val2 else x)
            if signe =="inf ou égal" and signe2 =="inf ou égal":
                self.source[champ_cible] = self.source[champ_cible][champ_cible].apply(lambda x: val_text if isinstance(x, int) and x <= val1 and x <= val2 else x)
            if signe =="inf ou égal" and signe2 =="sup":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x <= val1 and x > val2 else x)
            if signe =="inf ou égal" and signe2 =="sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x <= val1 and x >= val2 else x)
            if signe =="sup" and signe2 =="inf":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x > val1 and x < val2 else x)
            if signe =="sup" and signe2 =="inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x > val1 and x <= val2 else x)
            if signe =="sup" and signe2 =="sup":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x > val1 and x > val2 else x)
            if signe =="sup" and signe2 =="sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x > val1 and x >= val2 else x)
            if signe =="sup ou égal" and signe2 =="inf":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x >= val1 and x < val2 else x)
            if signe =="sup ou égal" and signe2 =="inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x >= val1 and x <= val2 else x)
            if signe =="sup ou égal" and signe2 =="sup":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x >= val1 and x > val2 else x)
            if signe =="sup ou égal" and signe2 =="sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(lambda x: val_text if isinstance(x, int) and x >= val1 and x >= val2 else x)

    def exportexcel(self, repertoire, onglet):
        """Méthode permettant d'extraire le résultat sous forme d'un fichier excel"""
        writer = pd.ExcelWriter(repertoire, engine='xlsxwriter')
        self.source.to_excel(writer, sheet_name=onglet)
        writer.save()

    def convert_date(self,champs, nom_champ_cible):
        "Methode convertissant le champ en date"
        self.source[champs] = pd.to_datetime(self.source[nom_champ_cible])

    def delta_time(self, date_deb, date_fin, champ_cible, nb_decimal = None):
        self.source[champ_cible] = (date_fin - self.source[date_deb]) #/ 365.25
        self.source[champ_cible] = round(self.source[champ_cible].dt.days / 365.25, nb_decimal)
analyse = Changedata()

analyse.chargement(r'D:\Users\sgasmi\Desktop\mydata2.xlsx','mydata')
analyse.convert_date('Date de naissance','Date de naissance')
analyse.delta_time('Date de naissance', date.today(), "Nv Age", 2)
analyse.tranche("Age","LOL", ">30 ans", "inf",30)
analyse.tranche("Age","LOL", "[30-40[", "sup ou égal", 30, "inf", 40)
analyse.tranche("Age","LOL", "[40-50[", "sup ou égal", 40, "inf", 50)
analyse.tranche("Age","LOL", "[50-99[", "sup ou égal",50)


analyse.exportexcel(r'D:\Users\sgasmi\Desktop\monresultat.xlsx',"Data")


