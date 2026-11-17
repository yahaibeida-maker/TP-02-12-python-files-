
# ============================================================
# APPLICATION 02 : PANDAS AU SERVICE DE L'AGRONOMIE
# TP 01 : DECOUVERTE DU DATASET AVEC PANDAS
# ============================================================

import pandas as pd

# 1. Chargement du fichier Excel
file_path = "OPEN SOURCE DATA FILE.xlsx"

df = pd.read_excel(
    file_path,
    sheet_name="DATA"
)

# 2. Informations generales
print("\n========== INFORMATIONS GENERALES ==========")

print("Nombre d'observations :", df.shape[0])
print("Nombre de variables :", df.shape[1])

print("\nNoms des variables :")
print(df.columns.tolist())

# 3. Afficher les premieres observations
print("\n========== PREMIERES OBSERVATIONS ==========")
print(df.head())

# 4. Afficher les dernieres observations
print("\n========== DERNIERES OBSERVATIONS ==========")
print(df.tail())

# 5. Informations sur les variables
print("\n========== STRUCTURE DU DATASET ==========")
df.info()

# 6. Statistiques descriptives
print("\n========== STATISTIQUES DESCRIPTIVES ==========")
print(df.describe())

# 7. Dimensions du tableau
print("\nDimensions du DataFrame :", df.shape)

print("\n============================================")
print("ANALYSE PANDAS TERMINEE AVEC SUCCES")
print("============================================")


