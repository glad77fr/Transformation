
analyse = Changedata()

analyse.chargement(r'D:\Users\sgasmi\Desktop\Données maquette\Effectif fin février.xlsx', 'Base')
#analyse.convert_date('Date de naissance', 'Date de naissance')

analyse.convert_date('Date de naissance')
analyse.convert_date("Date d'entrée gr société mère")
analyse.convert_date("Date d'entrée groupe")
analyse.convert_date("Date d'entrée société")
analyse.convert_date("Date d'entrée poste")

#analyse.delta_time('Date de naissance', date.today(), "Nv Age", 2)

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