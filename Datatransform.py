import pandas as pd
from xlsxwriter import *
from datetime import date
from datetime import datetime
import numpy as np

class Changedata:

    def __init__(self):
        self.source = pd.DataFrame()

    def chargement(self, chemin, feuille=""):
        if chemin[len(chemin) - 4:len(chemin)] == "xlsx" and feuille != "":
            Excel = pd.ExcelFile(chemin)
            self.source = Excel.parse(feuille)

    def colonne(self):
        print(self.source.columns)

    def remplacement_valeurs(self, champ, dic):
        self.source.replace(dic, inplace=True)

    def tranche(self, champ_source, champ_cible, val_text, signe, val1, signe2=None, val2=None):
        if champ_source != champ_cible:
            test = champ_cible in self.source.columns
            if test is False:
                self.source[champ_cible] = self.source[champ_source]


        if signe != "" and val1 != "" and signe2 is None:
            if signe == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x < val1 else x)
            if signe == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x <= val1 else x)
            if signe == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x > val1 else x)
            if signe == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x >= val1 else x)

        if signe != "" and signe2 is not None and val1 != "" and val2 is not None:
            if signe == "inf" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x < val1 and x < val2 else x)
            if signe == "inf" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x < val1 and x <= val2 else x)
            if signe == "inf" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x < val1 and x > val2 else x)
            if signe == "inf" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x < val1 and x >= val2 else x)
            if signe == "inf ou égal" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x <= val1 and x < val2 else x)
            if signe == "inf ou égal" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible][champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x <= val1 and x <= val2 else x)
            if signe == "inf ou égal" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x <= val1 and x > val2 else x)
            if signe == "inf ou égal" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x <= val1 and x >= val2 else x)
            if signe == "sup" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x > val1 and x < val2 else x)
            if signe == "sup" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x > val1 and x <= val2 else x)
            if signe == "sup" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x > val1 and x > val2 else x)
            if signe == "sup" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x > val1 and x >= val2 else x)
            if signe == "sup ou égal" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x >= val1 and x < val2 else x)
            if signe == "sup ou égal" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x >= val1 and x <= val2 else x)
            if signe == "sup ou égal" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x >= val1 and x > val2 else x)
            if signe == "sup ou égal" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, int) and x >= val1 and x >= val2 else x)

    def exportexcel(self, repertoire, onglet):
        """Méthode permettant d'extraire le résultat sous forme d'un fichier excel"""
        writer = pd.ExcelWriter(repertoire, engine='xlsxwriter')
        self.source.to_excel(writer, sheet_name=onglet)
        writer.save()

    def convert_date(self, champ):
        "Methode convertissant le champ en date"
        self.source[champ] = pd.to_datetime(self.source[champ], errors='coerce')

    def delta_time(self, date_deb, date_fin, champ_cible, nb_decimal=None):
        if isinstance(date_fin,str):
            date_fin = datetime.strptime(date_fin,'%d/%m/%Y')
        self.source[champ_cible] = (date_fin - self.source[date_deb])  # / 365.25
        self.source[champ_cible] = round(self.source[champ_cible].dt.days / 365.25, nb_decimal)
        self.source[champ_cible] = self.source[champ_cible].values.astype(np.int64)

analyse = Changedata()

analyse.chargement(r'D:\Users\sgasmi\Desktop\mydata2.xlsx', 'mydata')
#analyse.convert_date('Date de naissance', 'Date de naissance')

analyse.convert_date('Date de naissance')
analyse.convert_date("Date d'entrée gr société mère")
analyse.convert_date("Date d'entrée groupe")
analyse.convert_date("Date d'entrée société")
analyse.convert_date("Date d'entrée poste")

#analyse.delta_time('Date de naissance', date.today(), "Nv Age", 2)

analyse.delta_time('Date de naissance', '28/2/2018', "Age fin Février", 2)
analyse.delta_time("Date d'entrée gr société mère", '28/2/2018', "Ancienneté Vinci fin Février", 2)
analyse.delta_time("Date d'entrée groupe", '28/2/2018', "Ancienneté Eurovia fin Février", 2)
analyse.delta_time("Date d'entrée société", '28/2/2018', "Ancienneté société fin Février", 2)
analyse.delta_time("Date d'entrée poste", '28/2/2018', "Ancienneté poste fin Février", 2)

#analyse.delta_time('AA', '28/2/2018', "Ancienneté Vinci fin Février", 2)

analyse.tranche("Age", "Tranche d'âge", "<30ans", "inf", 30)
analyse.tranche("Age", "Tranche d'âge", "[30-40[", "sup ou égal", 30, "inf", 40)
analyse.tranche("Age", "Tranche d'âge", "[40-50[", "sup ou égal", 40, "inf", 50)
analyse.tranche("Age", "Tranche d'âge", "[50-99[", "sup ou égal", 50)

analyse.tranche("Ancienneté Eurovia fin Février", "Tranche ancienneté Eurovia", "<2ans", "inf", 2)
analyse.tranche("Ancienneté Eurovia fin Février", "Tranche ancienneté Eurovia", "[2-5[", "sup ou égal", 2, "inf", 5)
analyse.tranche("Ancienneté Eurovia fin Février", "Tranche ancienneté Eurovia", "[5-10[", "sup ou égal", 5, "inf", 10)
analyse.tranche("Ancienneté Eurovia fin Février", "Tranche ancienneté Eurovia", "[10-15[", "sup ou égal", 10, "inf", 15)
analyse.tranche("Ancienneté Eurovia fin Février", "Tranche ancienneté Eurovia", "[15-20[", "sup ou égal", 15, "inf", 20)
analyse.tranche("Ancienneté Eurovia fin Février", "Tranche ancienneté Eurovia", "[20-99[", "sup ou égal", 20)

print(analyse.source.dtypes)
analyse.exportexcel(r'D:\Users\sgasmi\Desktop\monresultat.xlsx', "Data")
