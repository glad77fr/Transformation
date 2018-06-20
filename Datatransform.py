import pandas as pd
from xlsxwriter import *
from datetime import date
from datetime import datetime
import numpy as np

class Changedata:

    def __init__(self):
        self.source = pd.DataFrame()

    def chargement(self, chemin, feuille=None, delimiter="None",header_positions = "None"):
        if chemin[len(chemin) - 4:len(chemin)] == "xlsx" and feuille != "":
            Excel = pd.ExcelFile(chemin)
            self.source = Excel.parse(feuille)

        if chemin[len(chemin) - 3:len(chemin)] == "csv":
            self.source = pd.read_csv(chemin,delimiter=delimiter,low_memory=False,header  = header_positions)


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

    def Entre_groupe(self,champ_cible):
        for i, cel in enumerate(self.source.itertuples()):
            if self.source.at[i, "Clé catégorie de mesure"] in [00,10]:
                self.source.at[i, champ_cible] = 1

    def Entre_soc(self,champ_cible):
        for i,cel in enumerate(self.source.itertuples()):
           if self.source.at[i, "Clé catégorie de mesure"] in [25,26,27]:
                self.source.at[i, champ_cible] = 1

    def Sortie_group(self,champ_cible):
        for i,cel in enumerate(self.source.itertuples()):
            if self.source.at[i, "Clé catégorie de mesure"] in [90]:
                self.source.at[i, champ_cible] = 1
