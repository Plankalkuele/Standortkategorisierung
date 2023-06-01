# -*- coding: utf-8 -*-
"""
Created on Sun May 28 22:06:08 2023

@author: H. Pfaff
"""
import openpyxl
import requests
import json
import re
import math

# Dateikonstanten
ZIELE = "Ziele.xlsx"
QUELLEN = "Quellen.xlsx"
ERGEBNIS = "Ergebnisse.xlsx"

# Kategorisierungskonstanten
KAT_I_MAX = 100
KAT_II_MAX = 200
KAT_III_MAX = 300

def adresskoordinaten_schreiben(worksheet, adressen):
    for i, adresse in enumerate(adressen, 1):
        ws.cell(row = i, column= 1, value = adresse)
        koordinaten = koordinaten_abrufen(i, adresse)
        ws.cell(row = i, column= 2, value = float(koordinaten[0]))
        ws.cell(row = i, column= 3, value = float(koordinaten[1]))

def adressen_abrufen(dateiname):
    workbook = openpyxl.load_workbook(dateiname)
    ws = workbook.active
    adress_liste = []
    for i, row in enumerate(ws.rows, 1):
        adress_liste.append(row[0].value)
    return adress_liste

def koordinaten_abrufen(i, adresse):
    split_adresse = re.split(',', adresse)
    plz_ort = split_adresse[-1]
    suchparameter = {"q": plz_ort}
    antwort = requests.get("https://geocode.maps.co/search", suchparameter)
    geodaten = json.loads(antwort.text)
    breite = geodaten[0]['lat']
    laenge = geodaten[0]['lon']
    return [breite, laenge]
    
def entfernung_berechnen(breite1, laenge1, breite2, laenge2):
    R = 6371
    dLat = math.radians(breite2 - breite1)
    dLon = math.radians(laenge2 - laenge1)
    a = math.sin(dLat/2) * math.sin(dLat/2) + \
        math.cos(math.radians(breite1)) * math.cos(math.radians(breite2)) * \
        math.sin(dLon/2) * math.sin(dLon/2)
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    d = R * c
    return d

def min_entfernungen_schreiben():
    for i, row in enumerate(wb["Ziele_Koordinaten"].rows, 1):
        liste_entfernungen = []
        for j, row in enumerate(wb["Quellen_Koordinaten"].rows, 1):
            entfernung = entfernung_berechnen(float(wb["Ziele_Koordinaten"].cell(row = i, column= 2).value), 
                                              float(wb["Ziele_Koordinaten"].cell(row = i, column= 3).value), 
                                              float(wb["Quellen_Koordinaten"].cell(row = j, column= 2).value), 
                                              float(wb["Quellen_Koordinaten"].cell(row = j, column= 3).value))
            liste_entfernungen.append(entfernung)
            min_entfernung = min(liste_entfernungen)
            min_index = [a for a, b in enumerate(liste_entfernungen) if b == min_entfernung]
        wb["Ziele_Koordinaten"].cell(row = i, column= 4, value = min_entfernung)
        wb["Ziele_Koordinaten"].cell(row = i, column= 5, value = "[" + str(min_index[0] + 1) + "]")
        wb["Ziele_Koordinaten"].cell(row = i, column= 6, value = kategorisieren(min_entfernung))
        
def kategorisieren(entfernung):
    if entfernung < KAT_I_MAX:
        return "I"
    elif entfernung < KAT_II_MAX:
        return "II"
    elif entfernung < KAT_III_MAX:
        return "III"
    else: 
        return "IV"
    
def formatieren_erweitert(blatt):
    formatieren(blatt)
    blatt.cell(row = 1, column= 5, value = "Entfernung in km")
    blatt.cell(row = 1, column= 6, value = "Quellenindex")
    blatt.cell(row = 1, column= 7, value = "Kategorie")
    
def formatieren(blatt):
    blatt.insert_rows(1,1)
    blatt.insert_cols(1,1)
    blatt.cell(row = 1, column= 1, value = "Index")
    blatt.cell(row = 1, column= 2, value = "Adresse")
    blatt.cell(row = 1, column= 3, value = "Breitengrad")
    blatt.cell(row = 1, column= 4, value = "Längengrad")
    for i, row in enumerate(blatt.rows, 1):
        blatt.cell(row = i + 1, column= 1, value = "[" + str(i) + "]")
    blatt.delete_rows(len(blatt['A']), 1)
    
wb = openpyxl.Workbook()
ws = wb.active

ws.title = "Ziele_Koordinaten"
adresskoordinaten_schreiben(ws, adressen_abrufen(ZIELE))

wb.create_sheet("Quellen_Koordinaten")
ws = wb["Quellen_Koordinaten"]
adresskoordinaten_schreiben(ws, adressen_abrufen(QUELLEN))


min_entfernungen_schreiben()
formatieren_erweitert(wb["Ziele_Koordinaten"])
formatieren(wb["Quellen_Koordinaten"])

wb.save(ERGEBNIS)
