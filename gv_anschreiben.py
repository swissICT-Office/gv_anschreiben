import subprocess
import requests
import openpyxl
import json
import pandas as pd



auswertung_dict = {
    "GV Anschreiben": 309
}

class Auswertungen:
    def __init__(self):
        self.gv_id = 309 # GV ID der Auswertung in Sewobe
        self.roles = ["Hauptkontakt", "CIO", "CEO", "HR-Manager"]

    @staticmethod
    def get_session_token():
        username = "ASWISSICT"
        password = "Adressliste24!"
        url = 'https://swissict.sewobe.de/applikation/restlogin/api/REST_LOGIN'

        # Data to be sent with the POST request (if any)
        data = {
            "USERNAME_REST": username,
            "PASSWORT_REST": password
        }

        login = requests.post(url, data=data)
        session = json.loads(login.text)["SESSION"]
        return session


    def get_adresse_seite(session_token, m):
        """Ort, PLZ und Strasse einer Einzelperson heraussuchen"""
        base_url = "https://swissict.sewobe.de/applikation/"
        get_adressen = f"adressen/api/GET_ADRESSEN&SEITE={m}&SESSION={session_token}"
        get_adressen_url = base_url + get_adressen
        adressen = requests.get(get_adressen_url)

        if json.loads(adressen.text):
            print(json.loads(adressen.text)["LISTE"])

    @staticmethod
    def get_funktion_prio(funktion):
        """Ermittelt die höchste Priorität aus einem FUNKTION-String."""
        if pd.isna(funktion):
            return None

        funktion = str(funktion)

        priority_order = [
            ("Hauptkontakt", 1),
            ("CIO", 2),
            ("CEO", 3),
            ("HR-Manager", 4),
        ]

        for role, prio in priority_order:
            if role in funktion:
                return prio

        return None

    @staticmethod
    def get_auswertung(session_token, id, file_name):
        """Ort, PLZ und Strasse einer Einzelperson heraussuchen"""
        base_url = "https://swissict.sewobe.de/applikation/"
        get_adressen = f"auswertungen/api/SUCHE_INDIV&AUSWERTUNG_ID={id}&SESSION={session_token}"
        get_adressen_url = base_url + get_adressen
        adressen = requests.get(get_adressen_url)

        if json.loads(adressen.text):
            data = json.loads(adressen.text)["ERGEBNIS"]
            df = pd.DataFrame(data).T
            df_flattened = pd.json_normalize(df["DATENSATZ"])

            df_combined = pd.concat(
                [df.drop(columns=["DATENSATZ"]), df_flattened],
                axis=1
            ).dropna(how="all")

            df_combined["AUSTRITT"] = df_combined["AUSTRITT"].replace("0000-00-00", "")

            tmp = df_combined.copy()

            # Priorität anhand aller enthaltenen Rollen in FUNKTION
            tmp["prio"] = tmp["FUNKTION"].apply(Auswertungen.get_funktion_prio)

            # Fallback auf [Zentrale] in ANREDE, aber nur wenn keine passende Rolle gefunden wurde
            tmp.loc[
                tmp["prio"].isna() & tmp["ANREDE"].eq("[Zentrale]"),
                "prio"
            ] = 5

            # Hinweis-Spalte vorbereiten
            tmp["HINWEIS"] = ""

            # Nach NR und Priorität sortieren
            # NaN kommt zuletzt, damit bei vorhandener Priorität zuerst diese gewählt wird
            tmp = tmp.sort_values(["NR", "prio"], na_position="last").copy()

            # Besten Kontakt pro NR auswählen:
            # - zuerst Rollen
            # - dann [Zentrale]
            # - sonst automatisch erster Kontakt der Gruppe
            result = tmp.drop_duplicates(subset="NR", keep="first").copy()

            result["BRIEFANREDE"] = result["ANREDE"].apply(
                lambda x: "Sehr geehrter" if x == "Herr"
                else ("Sehr geehrte" if x == "Frau" else "Sehr geehrte Damen und Herren")
            )

            # Falls weder Rolle noch [Zentral] gefunden wurde, Hinweis setzen
            result.loc[
                result["prio"].isna() & (result["ANREDE"] != "[Zentrale]"),
                "HINWEIS"
            ] = "Zentrale fehlt"

            result = result.drop(columns="prio")

            result.to_excel(file_name, index=False)
            return result


session_token = Auswertungen.get_session_token()

for key in auswertung_dict:
    print(key)
    Auswertungen.get_auswertung(session_token, auswertung_dict[key], f"{key}.xlsx")

