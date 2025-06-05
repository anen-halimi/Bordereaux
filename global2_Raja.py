import pandas as pd
import unicodedata
import os
from datetime import datetime

# === Fonctions utilitaires ===

def supprimer_accents(texte):
    if pd.isnull(texte):
        return texte
    texte = str(texte).upper()
    return ''.join(c for c in unicodedata.normalize('NFD', texte) if unicodedata.category(c) != 'Mn')

def extraire_valeurs_numeriques(texte):
    if pd.isnull(texte):
        return ''
    return ''.join(c for c in str(texte) if c.isdigit() or c in ['.', '-'])

def extraire_unite(texte):
    if pd.isnull(texte):
        return ''
    return ''.join(c for c in str(texte) if c.isalpha())

def extraire_bogie(nom_mesure, nom_essai):
    texte_mesure = supprimer_accents(str(nom_mesure)).upper()
    texte_essai = supprimer_accents(str(nom_essai)).upper()

    motifs_bogies = sorted(
        [f"{prefix}{i}" for prefix in ["BME", "BMI", "BPI"] for i in range(1, 100)],
        key=lambda x: -len(x)
    )
    motifs_voitures = [f"V{str(i).zfill(2)}" for i in range(1, 100)]
    tous_motifs = motifs_bogies + motifs_voitures

    def detecter(txt):
        return next((motif for motif in tous_motifs if motif in txt), None)

    trouvee_mesure = detecter(texte_mesure)
    trouvee_essai = detecter(texte_essai)

    return trouvee_mesure or trouvee_essai or "Rame"

def detecter_groupe_essai(nom):
    nom = supprimer_accents(str(nom)).strip()
    if "FONCTION AE" in nom:
        return "Fonction AE"
    elif "TPS" in nom and "FU" in nom and "MP(TT-F)" in nom:
        return "TPS S/D : FU par MP(TT-F)"
    elif "TPS" in nom and "BP" in nom and "PNEU" in nom:
        return "TPS S/D : FU pneu par BP(URG)"
    elif "TPS" in nom and "MDS" in nom and "MP(TT-F)" in nom:
        return "TPS S/D : MDS par MP(TT-F)"
    elif "TPS" in nom and "FU" in nom and "MEU" in nom:
        return "TPS S/D : FU électropneu par MEU"
    elif "BP" in nom and "URG" in nom:
        return "BP(URG)"
    elif "DET.SH" in nom or "ESSAI SH" in nom:
        return "Détendeur SH"
    elif "FEM" in nom and "RB" not in nom and "FP" not in nom and "ETANCH" not in nom:
        return "Détendeur FEM"
    elif "DETENDEUR" in nom and "FP" in nom:
        return "Détendeur FP"
    elif "ESSAIS" in nom and "RB" in nom:
        return "Essai RB MA(RA)FEM"
    elif "ETANCH" in nom and "CP" in nom and "CG" in nom:
        return "Etanchéité CP-CG"
    elif "ETANCH" in nom and ("RA FEM" in nom or "RA-FEM" in nom or "RA ET RA-FEM" in nom or "RA" in nom):
        return "Etanchéité des RA et RA-FEM"
    elif "IBU" in nom and "CAPTEUR" in nom:
        return "IBU(capteurs)"
    elif "IBU" in nom and "BME" in nom:
        return "IBU(BME)"
    elif "IBU" in nom and "BPI" in nom:
        return "IBU(BPI)"
    elif "IBU" in nom and "BMI" in nom:
        return "IBU(BMI)"
    elif "MA(PRD" in nom or "MA (PRD" in nom:
        return "MA(PRD)"
    elif "MA(URG" in nom and "CG" in nom and "CP" not in nom:
        return "MA(URG)CG"
    elif "MA (URG" in nom and "CG" in nom and "CP" not in nom:
        return "MA(URG)CG"
    elif "MA(URG" in nom and "CP" in nom:
        return "MA(URG)CP"
    elif "MA (URG" in nom and "CP" in nom:
        return "MA(URG)CP"
    elif "OPERATION" in nom and "LIBERATOIRES" in nom:
        return "opérations liberatoires"
    elif "PREPA" in nom and "ESSAIS" in nom:
        return "PREPA DES ESSAIS"
    elif "RM" in nom and "MINITROL" in nom:
        return "RM Minitrol"
    else:
        return nom



def detecter_groupe_mesure(nom):
    nom = supprimer_accents(str(nom))

    if "TPS" in nom and "BP(URG)G" in nom:
        return "Temps purge gauche"
    if "TEMPS" in nom and "BP(URG)G" in nom:
        return "Temps purge gauche"
    if "TPS" in nom and "BP(URG)D" in nom:
        return "Temps purge droit"
    if "TEMPS" in nom and "BP(URG)D" in nom:
        return "Temps purge droit"
    if "TPS" in nom and "PURGE CG" in nom:
        return "Temps purge CG"
    if "TPS" in nom and "ALIM CG" in nom:
        return "Temps alimentation CG"
    if "PRESSION" in nom and "GIME" in nom:
        return "pression régime CG"
    if "TEMPS" in nom and "SER" in nom and "DES" not in nom:
        return "Temps serrage"
    if "TEMPS" in nom and "DES" in nom:
        return "Temps desserrage"
    if "TANCH" in nom and "CG" in nom:
        return "Etanchéité CG"
    if "TANCH" in nom and "CP" in nom:
        return "Etanchéité CP"
    if "TANCH" in nom and "RA" in nom and "FEM" not in nom:
        return "Etanchéité RA"
    if "TANCH" in nom and "FEM" in nom:
        return "Etanchéité RA FEM"
    if "PR." in nom and "FIS" in nom:
        return "Pression détendeur FP"
    if "PR." in nom and "FEM" in nom and "RA" not in nom:
        return "Pression détendeur FEM"
    if "PR." in nom and "SH" in nom:
        return "Pression détendeur SH"
    if "MESUR" in nom and "REX" in nom:
        return "Mesure avant réglage"
    if "REMONT" in nom and "CF" in nom and "TPS" not in nom:
        return "Montée pression CF"
    if "PRESS" in nom and "CF" in nom and "TPS" not in nom:
        return "Pression CF"
    if "PURG" in nom and "CF" in nom and "TPS" not in nom:
        return "Purge complète CF"
    if "TPS" in nom and "CF" in nom and "MONT" in nom:
        return "Temps de montée pression CF"
    if "TPS" in nom and "CF" in nom and "PURG" in nom:
        return "Temps de purge pression CF"
    if "PR" in nom and "CROI" in nom and "DE" in nom:
        return "Pression décroissante"
    if "PR" in nom and "CROI" in nom and "DE" not in nom:
        return "Pression croissante"
    if "FU" in nom and "CG" in nom:
        return "Pression CG après FU"
    if "PRISE" in nom and ("REX" in nom or "FP" in nom):
        return "Mesure avant réglage"
    if "CHUT" in nom and "FEM" in nom:
        return "Chute pression FEM"
    if "MONT" in nom and "FEM" in nom:
        return "Remontée pression FEM"
    if "CG" in nom and "0,5" in nom and "CHUT" in nom:
        return "Chute de Pression CG 1ère Dep"
    if "CG" in nom and "TPS" in nom and "<" in nom and "RE" not in nom and "PURG" not in nom:
        return "Temps 1ère Dep CG"
    if "RE/CG" in nom and "0.05" in nom:
        return "Ecart RE/CG"
    if "CG" in nom and "4.50" in nom:
        return "Pression CG 1ère dep"
    if "RE" in nom and "1" in nom and "TPS" in nom and "CG" not in nom:
        return "Temps 1ère Dep RE"
    if "CFF" in nom and "DIS" in nom and "3.80" in nom and "RE" not in nom:
        return "Pression CFF DIS maxi"
    if "TPS" in nom and "PURG" in nom and "CG" in nom:
        return "Temps de purge CG"
    if "CFF-DIS" in nom and "0,10" in nom:
        return "CFF-DIS à 0b"
    if "INFO" in nom and "REG" in nom:
        return "Info DE RG IBU"
    if "MESU" in nom and "RG" in nom:
        return "Mesure avant réglage"
    if "MESU" in nom and "CF" in nom:
        return "Mesure avant réglage"
    if "DE" in nom and "REG" in nom and "3.80" in nom:
        return "Pression détendeur RG-IBU"

    pesee_vals = ["= 0", "2.94", "3.14", "3.38", "3.41", "3.49", "3.70", "3.88", "3.94", "4.35", "4.57", "4.63", "4.80", "5.06",
                  "5.20", "5.27", "5.31", "5.42", "5.52", "5.62", "5.67", "5.83", "5.88", "6.42", "6.62", "6.67", "6.83", "6.88"]
    for val in pesee_vals:
        if "CF" in nom and "PESEE" in nom and val in nom and "MESU" not in nom:
            return f"CF à pesée {val} Bar"

    return nom




# === Paramètres ===
chemin_fichier = "/home/anen/2025 1/2025/outputs/Résumé_Global_Complet.xlsx"

# === Lecture fichier ===
df = pd.read_excel(chemin_fichier)

# Nettoyage valeurs
df["Valeur de la mesure"] = df["Valeur de la mesure"].astype(str).str.strip().str.replace(",", ".").str.replace("\u00A0", " ")

# Extraire num, unité
df["Texte extrait"] = df["Valeur de la mesure"].apply(extraire_valeurs_numeriques)
df["Valeur numérique"] = pd.to_numeric(df["Texte extrait"], errors='coerce')
df["Unité"] = df["Valeur de la mesure"].apply(extraire_unite)

df["Valeur Bar"] = df.apply(lambda x: x["Valeur numérique"] if str(x["Unité"]).lower() == "bar" else 0, axis=1)
df["Valeur Seconde"] = df.apply(lambda x: x["Valeur numérique"] if str(x["Unité"]).lower() == "s" else 0, axis=1)

# Série, module
series_possibles = ["6 caisses", "7 caisses", "8 caisses", "10 caisses", "10 caisses ONO"]
modules_valides = ["MSAJ5", "MSAJ6", "MSAJ7"]
df["Nom du programme original"] = df["Nom du programme"]
df["Série"] = df["Nom du programme original"].apply(lambda x: next((s for s in series_possibles if s in str(x)), None))
df["Nom du module"] = df["Nom du programme original"].apply(lambda x: "_".join([m for m in modules_valides if m in str(x)]))

# Site, ville
codes_sites = {
    "AS": "Amiens", "BX": "Bordeaux", "COE": "Corbeil-essonnes", "LL": "Le Landy",
    "MG": "Montrouge", "AMC": "Chatillon - Montrouge", "TE": "Toulouse",
    "VSX": "Vénissieux", "SO": "Sotteville", "NB": "Nantes", "ORL": "Orléans",
    "NSR": "Nice St Roch", "MB": "Marseille", "VSG": "Villeneuve"
}
df["Site"] = df["Nom du programme original"].apply(lambda x: next((code for code in codes_sites if code in str(x)), None))
df["Ville"] = df["Site"].map(codes_sites)

# Date fichier et essai
df["Date dernière modif fichier"] = datetime.fromtimestamp(os.path.getmtime(chemin_fichier))
df["DateEssai_Vraie"] = pd.to_datetime(df["Date d'essai"], errors='coerce', dayfirst=True)
df["Année"] = df["DateEssai_Vraie"].dt.year

# Groupes
df["Nom groupe"] = df["Nom de l'essai"].apply(detecter_groupe_essai)
df["Groupe Mesure"] = df["Nom de la mesure"].apply(detecter_groupe_mesure)

# ✅ Ajout colonne Bogie
df["Bogie"] = df.apply(lambda row: extraire_bogie(row["Nom de la mesure"], row["Nom de l'essai"]), axis=1)

# === Exports ===
df.to_excel("/home/anen/résultat_complet.xlsx", index=False)

# Par année
dossier_export = "/home/anen"
os.makedirs(dossier_export, exist_ok=True)
for annee, df_annee in df.groupby("Année"):
    if pd.notnull(annee):
        fichier = os.path.join(dossier_export, f"data_{int(annee)}.xlsx")
        df_annee.to_excel(fichier, index=False)

print("✅ Fichier global et fichiers par année générés avec succès.")

