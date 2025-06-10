def enrichir_resume_global(path_fichier_excel, log_func=print):
        import pandas as pd
        import os
        import unicodedata
        from datetime import datetime

        def log(msg):
            log_func(msg)

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
            return nom

        log("📥 Chargement du fichier Excel...")
        df = pd.read_excel(path_fichier_excel)

        log("🧹 Nettoyage et transformation des valeurs de mesure...")
        df["Valeur de la mesure"] = (
            df["Valeur de la mesure"]
            .astype(str)
            .str.strip()
            .str.replace(",", ".", regex=False)
            .str.replace("\\u00A0", " ", regex=False)
        )

        df["Texte extrait"] = df["Valeur de la mesure"].apply(extraire_valeurs_numeriques)
        df["Valeur numérique"] = pd.to_numeric(df["Texte extrait"], errors='coerce')
        df["Unité"] = df["Valeur de la mesure"].apply(extraire_unite)

        log("🔢 Calcul des colonnes 'Valeur Bar' et 'Valeur Seconde'...")
        df["Valeur Bar"] = df.apply(lambda x: x["Valeur numérique"] if str(x["Unité"]).lower() == "bar" else 0, axis=1)
        df["Valeur Seconde"] = df.apply(lambda x: x["Valeur numérique"] if str(x["Unité"]).lower() == "s" else 0, axis=1)

        log("📌 Extraction des métadonnées depuis 'Nom du programme'...")
        df["Nom du programme original"] = df["Nom du programme"]
        series_possibles = ["6 caisses", "7 caisses", "8 caisses", "10 caisses", "10 caisses ONO"]
        modules_valides = ["MSAJ5", "MSAJ6", "MSAJ7"]
        df["Série"] = df["Nom du programme original"].apply(lambda x: next((s for s in series_possibles if s in str(x)), None))
        df["Nom du module"] = df["Nom du programme original"].apply(lambda x: "_".join([m for m in modules_valides if m in str(x)]))

        log("🌍 Détection de la ville à partir du site...")
        codes_sites = {
            "AS": "Amiens", "BX": "Bordeaux", "COE": "Corbeil-essonnes", "LL": "Le Landy",
            "MG": "Montrouge", "AMC": "Chatillon - Montrouge", "TE": "Toulouse",
            "VSX": "Vénissieux", "SO": "Sotteville", "NB": "Nantes", "ORL": "Orléans",
            "NSR": "Nice St Roch", "MB": "Marseille", "VSG": "Villeneuve"
        }
        df["Site"] = df["Nom du programme original"].apply(lambda x: next((code for code in codes_sites if code in str(x)), None))
        df["Ville"] = df["Site"].map(codes_sites)

        log("📆 Calcul de la date de dernière modification et de l’année d’essai...")
        df["Date dernière modif fichier"] = datetime.fromtimestamp(os.path.getmtime(path_fichier_excel))
        df["DateEssai_Vraie"] = pd.to_datetime(df["Date d'essai"], errors='coerce', dayfirst=True)
        df["Année"] = df["DateEssai_Vraie"].dt.year

        log("📊 Application des règles 'Nom groupe' et 'Groupe mesure'...")
        df["Nom groupe"] = df["Nom de l'essai"].apply(detecter_groupe_essai)
        df["Groupe Mesure"] = df["Nom de la mesure"].apply(detecter_groupe_mesure)

        log("🚄 Extraction du type de bogie...")
        df["Bogie"] = df.apply(lambda row: extraire_bogie(row["Nom de la mesure"], row["Nom de l'essai"]), axis=1)

        log("💾 Enregistrement du fichier global...")
        dossier_export = os.path.dirname(path_fichier_excel)
        df.to_excel(os.path.join(dossier_export, "résultat_complet.xlsx"), index=False)

        log("📂 Génération des fichiers par année...")
        for annee, df_annee in df.groupby("Année"):
            if pd.notnull(annee):
                df_annee.to_excel(os.path.join(dossier_export, f"data_{int(annee)}.xlsx"), index=False)
                log(f"✅ Fichier exporté pour l’année {int(annee)}")

        log("✅ Enrichissement terminé avec succès.")

def enrichir_global(path_fichier_excel, log_func=print):
    import pandas as pd
    import os
    df = pd.read_excel(path_fichier_excel)
    # … toutes tes transformations (mêmes étapes que tu fais déjà)
    # jusqu’à...
    log_func("💾 Création de résultat_complet.xlsx")
    df.to_excel(os.path.join(os.path.dirname(path_fichier_excel), "résultat_complet.xlsx"), index=False)
    log_func("✅ Fichier global généré avec succès.")

def enrichir_par_annee(path_fichier_excel, log_func=print):
    import pandas as pd
    import os
    df = pd.read_excel(path_fichier_excel)
    df["DateEssai_Vraie"] = pd.to_datetime(df["Date d'essai"], errors='coerce', dayfirst=True)
    df["Année"] = df["DateEssai_Vraie"].dt.year
    log_func("📂 Génération des fichiers par année...")
    for annee, df_annee in df.groupby("Année"):
        if pd.notnull(annee):
            nom_fichier = f"data_{int(annee)}.xlsx"
            df_annee.to_excel(os.path.join(os.path.dirname(path_fichier_excel), nom_fichier), index=False)
            log_func(f"[Télécharger fichier {annee}](/download_excel/{nom_fichier})")

