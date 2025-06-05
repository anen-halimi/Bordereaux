# Extraction et Fusion Automatique de Résumés PDF vers Excel
---
## 📋 Description du projet
Ce projet automatise l'extraction, l'organisation et la fusion de données provenant de rapports PDF structurés, notamment des essais techniques avec tableaux associés. Il permet de générer pour chaque PDF un fichier Excel détaillé avec plusieurs feuilles, puis de regrouper tous les résultats dans un fichier résumé global unique.
---
## 🚀 Fonctionnalités principales
- Extraction automatique des titres, sous-titres (ex. "Essai n°X"), et tableaux associés depuis des fichiers PDF multiples.
- Gestion avancée des titres avec reconnaissance OCR et détection couleur (ex. titres violets, "ANNEXES").
- Création d’un fichier Excel par PDF avec plusieurs feuilles :
  - **Résultat Ordonné** : titres + tableaux dans l'ordre visuel PDF.
  - **Résultat Fusionné** : fusion des tableaux par essai/titre.
  - **Résumé Global** : synthèse structurée des mesures clés avec colonnes dédiées.
- Traitement batch d’un dossier complet de fichiers PDF.
- Fusion finale de tous les résumés globaux en un fichier Excel unique.
---
## 🛠️ Prérequis et installation
### 1. Environnement Python
- Python 3.8 ou supérieur recommandé.
### 2. Modules Python nécessaires
Installer via pip :
\`\`\`bash
pip install pandas openpyxl pdfplumber pytesseract pdf2image opencv-python numpy
\`\`\`
### 3. Logiciels externes à installer
- **Tesseract OCR**
  - Ubuntu :
    \`\`\`bash
    sudo apt-get install tesseract-ocr tesseract-ocr-fra
    \`\`\`
  - Windows / MacOS : Télécharger ici
- **Poppler-utils** (nécessaire pour pdf2image) :
  - Ubuntu :
    \`\`\`bash
    sudo apt-get install poppler-utils
    \`\`\`
  - Windows / MacOS : voir guide d'installation Poppler
### 4. Configuration dans le script
Vérifie le chemin vers l’exécutable Tesseract dans le script Python (`pytesseract.pytesseract_cmd`), par exemple :
\`\`\`python
pytesseract.pytesseract_cmd = "/usr/bin/tesseract"
\`\`\`
---
## 📂 Structure du projet
\`\`\`
project_root/
├── main.py             # Script principal
├── requirements.txt    # Dépendances Python
├── README.md           # Ce fichier
└── pdfs/               # Dossier contenant les fichiers PDF à traiter
    └── outputs/        # Sorties Excel générées automatiquement
\`\`\`
---
## 📖 Mode d’emploi
### Lancement du traitement batch
\`\`\`bash
python main.py
\`\`\`
- Une fenêtre s’ouvre pour sélectionner le dossier contenant les PDF.
- Le script traite chaque PDF en générant un fichier Excel dans un sous-dossier `outputs/`.
- En fin d’exécution, un fichier `Résumé_Global_Complet.xlsx` fusionne tous les résumés.
---
## Résultats
Pour chaque PDF traité, tu obtiens :
- `NomDuPDF_fusion_finale_TOPDOWN_OK.xlsx` contenant :
  - Feuille **Résultat Ordonné**
  - Feuille **Résultat Fusionné**
  - Feuille **Résumé Global**
Un fichier global `Résumé_Global_Complet.xlsx` compile tous les résumés.
---
## 🔍 Détails techniques
### Détection des titres et sous-titres
- Analyse couleur HSV pour repérer titres violets.
- OCR via Tesseract pour détecter "ANNEXES" et sous-titres "Essai n°X".
- Utilisation de pdfplumber pour extraction précise des mots/titres.
### Extraction des tableaux
- Identification des tableaux par pdfplumber.
- Tri vertical des blocs pour associer chaque tableau à son titre.
- Fusion des tableaux entre deux titres.
### Extraction des données synthétiques
- Extraction et nettoyage des colonnes clés des essais.
- Séparation en colonnes : numéro d’essai, nom de l’essai, nombre de réalisation, nom de la mesure, valeur.
### Fusion finale
- Concaténation des fichiers Excel intermédiaires.
- Ajout d’une colonne source pour tracer l’origine des données.
---
## 🧰 Fichiers importants et fonctions
- `extract_and_order_blocks(pdf_path, output_excel)`  
  Extraction titres + tableaux, sauvegarde "Résultat Ordonné".
- `create_fused_results_sheet(extracted_excel_path)`  
  Fusion tableaux entre titres, création "Résultat Fusionné".
- `extract_essais_and_mesures_from_titles(ws)`  
  Extraction mesures depuis "Résultat Fusionné".
- `create_global_summary(extracted_excel_path)`  
  Création de la feuille "Résumé Global" avec colonnes structurées.
- `process_all_pdfs_in_folder(pdf_folder)`  
  Traitement batch dossier + fusion globale.
---
## 💡 Conseils d’utilisation
- Veille à ce que les PDF soient lisibles et non corrompus.
- Pour des résultats optimaux, utilise des PDF bien structurés avec des titres et tableaux réguliers.
- Ajuste les paramètres OCR (langue, mode psm) dans le script si nécessaire.
- L’interface tkinter facilite la sélection du dossier mais peut être modifiée selon besoin.
---
## 📞 Support
Si tu rencontres un problème, merci de fournir :
- Exemple PDF source (si possible).
- Capture d’écran des erreurs.
- Log d’exécution du script.
