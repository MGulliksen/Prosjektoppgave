# -*- coding: utf-8 -*-
"""
I denne filen vil du finne totalt 6 koder som løser de 6 
deloppgavene i prosjektoppgaven. 
Det er tatt høyde for at oppgavene kjøres fra samme sted, slik at dersom f.eks. import numpy as np er skrevet inn 
i en oppgave er den ikke lagt inn på ny dersom det er behov for det senere. 

Created on Tue Apr  8 21:29:38 2025

@author: Marianne Gulliksen ma_gulliksen@hotmail.com
"""

""" Del a: Viser ett program som leser inn en bestemt fil og lagrer dataene
fra kolonne A, B, C og D til hver sin array med forhåndsdefinierte navn."""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

data = pd.read_excel("support_uke_24.xlsx")

u_dag = data ["Ukedag"].values
kl_slett = data ["Klokkeslett"].values
varighet = data ["Varighet"].values
score = data ["Tilfredshet"].values

"""Brukt for kontroll av plassering; plt.plot(u_dag, kl_slett, varighet, score)"""



""" Del b: Viser et program som finner antall henvendeseler for hver av de 5 
ukenedagene. Resultatet visualiseres ved bruk av et søylediagram 
(stoplediagram)."""

"""Teller antall hendvenelser pr. ukedag i filen"""
henvendelser_pr_dag = data["Ukedag"].value_counts()

"""Setter opp et søylediagram for å vise resultatet"""
henvendelser_pr_dag.plot(kind= "bar")
plt.title("Antall henvendelser pr. ukedag")
plt.xlabel("Ukedag")
plt.ylabel("Antall henvendelser")
plt.show()



""" Del c: Viser et program som finner minste og lengste samtaletid som er 
loggført for uke24. Svaret skrives deretter til skjerm med forklarende tekst. """

"""Finner den minste og den lengste samtaletiden"""
minst_varighet = data["Varighet"].min()
lengst_varighet = data["Varighet"].max()

"""Print resultatet"""
print(f"Den minste samtaletiden som er loggført for uke 24 var på {minst_varighet:.4} minutter.")
print(f"Den lengste samtaletiden som er loggført for uke 24 var på {lengst_varighet:.4} minutter.")
print()




""" Del d: KREVENDE: Skriv et program som regner ut gjennomsnittlig 
samtaletid basert på alle henvendelser i uke 24. """

"""Denne oppgaven er merket med KREVENDE. Her drar jeg nytte av at noe av det jeg gjør 
mest i min nåværende jobb er å bearbeide hovedbøker og annen dokumentasjon i Excel fra regnskapsførere, slik at 
det er enklere å jobbe med. Løsningen her er å konvertere tiden slik at den vises som desimaltegn i excelfilen. Når dette er
korrigert i opprinnelig fil, så fortsetter vi med oppgaven"""

"""Beregner gjennomsnittlig varighet"""
gjennomsnitt_samtaletid = data["Varighet"].mean()

"""Print resultatet"""
print(f"Gjennomsnittlig samtaletid i uke 24 er {gjennomsnitt_samtaletid:.2f} minutter.")
print()




""" Del e: I denne oppgaven så finner vi antall henvendelser fordelt på spesefike 
tidsrom gjennom døgnet i uke 24, og viser disse i et kakediagram"""

"""Konverterer klokkeslett, slik at de kan leses inn til kakediagrammet"""
data["Klokkeslett"] = pd.to_datetime(data["Klokkeslett"], format='%H:%M:%S', errors='coerce')

"""Oppdeling i 2-timers bolker"""
tidsintervaller = [8, 10, 12, 14, 16]
tidsmerker = ["08-10", "10-12", "12-14", "14-16"]

"""Kategoriserer klokkeslettene i sine intervaller"""
data["2timersbolk"] = pd.cut(data["Klokkeslett"].dt.hour, bins=tidsintervaller, labels=tidsmerker, right=False)

"""Teller antall hendelser for hver 2-timers bolk"""
henvendelser_per_bolk = data["2timersbolk"].value_counts(sort=False)

"""Setter opp kakediagram for å vise resultatet"""
henvendelser_per_bolk.plot(kind="pie", autopct="%1.0f%%", colors=["blue", "red", "green", "pink"])
plt.title("Henvendelser per 2-timers bolk (Uke 24)")
plt.show()
    


""" Del f: I denne oppgaven har vi sett på resultatet av kundetilfredshetsundersøkelsen, og 
har fordelt svarene i egne kundegrupper; positiv, nøytral og negativ. På bagrunn av dette regner vi ut NPS. 
I tillegg er det valgt å vise faktisk prosentfordeling på de forskjellige gruppene."""

"""Fjerner tomme celler i kolonnen for Tilfredshet"""
data = data.dropna(subset=["Tilfredshet"])

"""Oppdeling av kundetilfredshet"""
negative_kunder = data[data["Tilfredshet"] <= 6].shape[0]
nøytrale_kunder = data[(data["Tilfredshet"] == 7) | (data["Tilfredshet"] == 8)].shape[0]
positive_kunder = data[data["Tilfredshet"] >= 9].shape[0]

"""Totalt antall kunder """
total_kunder = data.shape[0]

"""Oppdeling i % for de forskjellige kundegruppene"""
prosent_negative = (negative_kunder / total_kunder) * 100
prosent_nøytrale = (nøytrale_kunder / total_kunder) * 100
prosent_positive = (positive_kunder / total_kunder) * 100

"""Utregning av NPS"""
NPS = prosent_positive - prosent_negative

"""Printer svaret til skjerm. Her vises Net Promotor Score, i tillegg til 
fordelingen mellom de enkelte kundegruppene"""

print(f"Net Promoter Score (NPS): {NPS:.2f}")
print()
print("Fordelt pr. kundegruppe:")
print(f"Prosent negative kunder: {prosent_negative:.2f}%")
print(f"Prosent nøytrale kunder: {prosent_nøytrale:.2f}%")
print(f"Prosent positive kunder: {prosent_positive:.2f}%")

