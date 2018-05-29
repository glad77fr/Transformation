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
                    lambda x: val_text if isinstance(x, float) and x < val1 else x)
            if signe == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x <= val1 else x)
            if signe == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x > val1 else x)
            if signe == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x >= val1 else x)

        if signe != "" and signe2 is not None and val1 != "" and val2 is not None:
            if signe == "inf" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x < val1 and x < val2 else x)
            if signe == "inf" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x < val1 and x <= val2 else x)
            if signe == "inf" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x < val1 and x > val2 else x)
            if signe == "inf" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x < val1 and x >= val2 else x)
            if signe == "inf ou égal" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x <= val1 and x < val2 else x)
            if signe == "inf ou égal" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible][champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x <= val1 and x <= val2 else x)
            if signe == "inf ou égal" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x <= val1 and x > val2 else x)
            if signe == "inf ou égal" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x <= val1 and x >= val2 else x)
            if signe == "sup" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x > val1 and x < val2 else x)
            if signe == "sup" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x > val1 and x <= val2 else "toto")
            if signe == "sup" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x > val1 and x > val2 else x)
            if signe == "sup" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x > val1 and x >= val2 else x)
            if signe == "sup ou égal" and signe2 == "inf":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x >= val1 and x < val2 else x)
            if signe == "sup ou égal" and signe2 == "inf ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x >= val1 and x <= val2 else x)
            if signe == "sup ou égal" and signe2 == "sup":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x >= val1 and x > val2 else x)
            if signe == "sup ou égal" and signe2 == "sup ou égal":
                self.source[champ_cible] = self.source[champ_cible].apply(
                    lambda x: val_text if isinstance(x, float) and x >= val1 and x >= val2 else x)

    def exportexcel(self, repertoire, onglet):
        """Méthode permettant d'extraire le résultat sous forme d'un fichier excel"""
        writer = pd.ExcelWriter(repertoire, engine='xlsxwriter')
        self.source.to_excel(writer, sheet_name=onglet)
        writer.save()

    def convert_date(self, champ, nv_champ= None):
        "Methode convertissant le champ en date"
        if nv_champ is None:
            self.source[champ] = pd.to_datetime(self.source[champ], errors='coerce')
        else:
            self.source[nv_champ] = pd.to_datetime(self.source[champ], errors='coerce')

    def delta_time(self, date_deb, date_fin, champ_cible, nb_decimal=None):
        if isinstance(date_fin,str):
            date_fin = datetime.strptime(date_fin,'%d/%m/%Y')

        #self.source[champ_cible] = self.source[champ_cible].values.astype(np.int64)
        self.source[champ_cible] = (date_fin - self.source[date_deb])
        self.source[champ_cible] = round(self.source[champ_cible].dt.days / 365.25, nb_decimal)

    def effectif_inscrit(self,champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] in [1, 3]) and (self.source.at[i, "CtSAL"] in [1, 8])\
                    and (self.source.at[i, "Sté LC"] not in [9108, 2429]):
                self.source.at[i, champ_cible] = 1

    def effectif_physique_actif(self,champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] == 3) and (self.source.at[i, "CtSAL"] in [1, 8]) \
                    and (self.source.at[i, "Sté LC"] not in [9108, 2429]):
                self.source.at[i, champ_cible] = 1

    def effectif_ETP_theorique(self,champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] in [1, 3]) and (self.source.at[i, "CtSAL"] in [1, 8]) \
                    and (self.source.at[i, "Sté LC"] not in [9108, 2429]):
                self.source.at[i, champ_cible] = self.source.at[i, "Equivalent temps plein"]

    def effectif_retraite(self, champ_cible, date_cible, champ_age):
        if isinstance(date_cible,str):
            date_cible = datetime.strptime(date_cible,'%d/%m/%Y')
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] in [1, 3]) and (self.source.at[i, "CtSAL"] in [1, 8]) \
                    and (self.source.at[i, "Sté LC"] not in [9108, 2429]) and pd.isnull(self.source.at[i,"Date de naissance"]) is False:

                        if int(self.source.at[i, "Date de naissance"].year) < 1956 and self.source.at[i, champ_age] >= 65: #60 pour la france
                            self.source.at[i, champ_cible] = 1

                        if int(self.source.at[i, "Date de naissance"].year) >= 1956 and self.source.at[i, champ_age] >= 65: #62 pour la france
                            self.source.at[i,champ_cible] = 1

    def effectif_interim(self, champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] in [1, 3]) and self.source.at[i, "CtSAL"] == 7 \
                    and (self.source.at[i, "Sté LC"] not in [9108, 2429]):
                self.source.at[i, champ_cible] = 1

    def headcount(self, champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] in [1, 3]) and str(self.source.at[i, "DPer"])[0] == "G" and \
                    self.source.at[i, "Tranche de décompte"] not in [99]:
                self.source.at[i, champ_cible] = 1

    def FTE(self,champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if (self.source.at[i, "Clé statut d'activité"] in [1, 3]) and str(self.source.at[i, "DPer"])[0] == "G" and \
                    self.source.at[i, "Tranche de décompte"] not in [99]:
                self.source.at[i, champ_cible] = self.source.at[i, "Hres ouvrées/semaine"]/self.source.at[i, "Full Time Equivalent"]

analyse = Changedata()

analyse.chargement(r'D:\Users\sgasmi\Desktop\Données maquette\Effectif fin février.xlsx', 'Base')
#analyse.convert_date('Date de naissance', 'Date de naissance')

analyse.convert_date('Date de naissance')
analyse.convert_date("Date d'entrée gr société mère")
analyse.convert_date("Date d'entrée groupe")
analyse.convert_date("Date d'entrée société")
analyse.convert_date("Date d'entrée poste")

#analyse.delta_time('Date de naissance', date.today(), "Nv Age", 2)
""""
analyse.delta_time('Date de naissance', '28/2/2018', "Age salarié", 2)
analyse.delta_time("Date d'entrée gr société mère", '28/2/2018', "Ancienneté Vinci", 2)
analyse.delta_time("Date d'entrée groupe", '28/2/2018', "Ancienneté Eurovia", 2)
analyse.delta_time("Date d'entrée société", '28/2/2018', "Ancienneté société", 2)
analyse.delta_time("Date d'entrée poste", '28/2/2018', "Ancienneté poste", 2)

#analyse.delta_time('AA', '28/2/2018', "Ancienneté Vinci fin Février", 2)

analyse.tranche("Age salarié", "Tranche d'âge", "<30ans", "inf", 30)
analyse.tranche("Age salarié", "Tranche d'âge", "[30-40[", "sup ou égal", 30, "inf", 40)
analyse.tranche("Age salarié", "Tranche d'âge", "[40-50[", "sup ou égal", 40, "inf", 50)
analyse.tranche("Age salarié", "Tranche d'âge", "[50-99[", "sup ou égal", 50)

analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "<2ans", "inf", 2)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "[2-5[", "sup ou égal", 2, "inf", 5)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "[5-10[", "sup ou égal", 5, "inf", 10)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "[10-15[", "sup ou égal", 10, "inf", 15)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "[15-20[", "sup ou égal", 15, "inf", 20)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "[20-99[", "sup ou égal", 20)

analyse.effectif_inscrit("Effectif inscrit")
analyse.effectif_physique_actif("Effectif inscrit actif")
analyse.effectif_ETP_theorique("Effectif ETP")
analyse.effectif_retraite("Prediction retraite", '28/02/2018', "Age salarié")
analyse.effectif_interim("Interimaire")
#print(analyse.source.dtypes)
"""

analyse.source["Sexe"].value_counts().head(10).plot.bar()
analyse.delta_time('Date de naissance', '31/3/2018', "Age salarié", 2)
analyse.delta_time("Date d'entrée gr société mère", '31/3/2018', "Ancienneté Vinci", 2)
analyse.delta_time("Date d'entrée groupe", '31/3/2018', "Ancienneté Eurovia", 2)
analyse.delta_time("Date d'entrée société", '31/3/2018', "Ancienneté société", 2)
analyse.delta_time("Date d'entrée poste", '31/3/2018', "Ancienneté poste", 2)

#analyse.delta_time('AA', '28/2/2018', "Ancienneté Vinci fin Février", 2)

analyse.tranche("Age salarié", "Tranche d'âge", "< or = 25 years", "inf", 26)
analyse.tranche("Age salarié", "Tranche d'âge", "26 to 30 years", "sup ou égal", 26, "inf", 31)
analyse.tranche("Age salarié", "Tranche d'âge", "31 to 35 years", "sup ou égal", 31, "inf", 36)
analyse.tranche("Age salarié", "Tranche d'âge", "36 to 40 years", "sup ou égal", 36, "inf", 41)
analyse.tranche("Age salarié", "Tranche d'âge", "41 to 45 years", "sup ou égal", 41, "inf", 46)
analyse.tranche("Age salarié", "Tranche d'âge", "46 to 50 years", "sup ou égal", 46, "inf", 51)
analyse.tranche("Age salarié", "Tranche d'âge", "51 to 55 years", "sup ou égal", 51, "inf", 56)
analyse.tranche("Age salarié", "Tranche d'âge", "56 to 60 years", "sup ou égal", 56, "inf", 61)
analyse.tranche("Age salarié", "Tranche d'âge", " + or = 61 years", "sup ou égal", 61)


analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "<1", "inf", 1)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "1 to 5 years", "sup ou égal", 1, "inf", 6)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "6 to 10 years", "sup ou égal", 6, "inf", 11)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "11 to 15 years", "sup ou égal", 11, "inf", 16)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "16 to 20 years", "sup ou égal", 16, "inf", 21)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "21 to 25 years", "sup ou égal", 21, "inf", 26)
analyse.tranche("Ancienneté Eurovia", "Tranche ancienneté Eurovia", "+ or = 26 years", "sup ou égal", 26)

analyse.remplacement_valeurs("Type de contrat.1",{"Perm full time" : "Headcount with open-ended contract","Perm part time" : "Headcount with open-ended contract",\
"Fix term full T" :"Headcount with term contract", "Fix term part T" : "Headcount with term contract", "Casual" : "Headcount with term contract",\
"Seasonal" : "Headcount with term contract","Apprentice Full" : "Headcount with open-ended contract"})

analyse.effectif_inscrit("Effectif inscrit")
analyse.effectif_physique_actif("Effectif inscrit actif")
analyse.effectif_ETP_theorique("Effectif ETP")
analyse.effectif_retraite("Prediction retraite", '31/3/2018', "Age salarié")
analyse.effectif_interim("Interimaire")
analyse.headcount("Headcount")
analyse.FTE("FTE")

analyse.exportexcel(r'D:\Users\sgasmi\Desktop\monresultat.xlsx', "Base")
