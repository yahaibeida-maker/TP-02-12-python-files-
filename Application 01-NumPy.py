
# ============================================================
# APPLICATION 01 : NUMPY AU SERVICE DE L'AGRONOMIE
# Analyse d'un jeu de données environnementales
# Université 20 Août 1955 de Skikda
# Réalisation : Python sous Spyder (Anaconda)
# ============================================================

# 1. IMPORTATION DES BIBLIOTHEQUES
import numpy as np
from openpyxl import load_workbook

# ============================================================
# 2. CHARGEMENT DU FICHIER EXCEL
# ============================================================

file_path = "OPEN SOURCE DATA FILE.xlsx"

# Ouvrir le fichier Excel
workbook = load_workbook(
    file_path,
    data_only=True,
    read_only=True
)

# Sélectionner la feuille DATA
sheet = workbook["DATA"]

# Lire les noms des variables
variables = [
    cell.value for cell in sheet[1]
]

# Lire les observations
observations = list(
    sheet.iter_rows(
        min_row=2,
        values_only=True
    )
)

# Convertir les données en tableau NumPy
X = np.array(observations, dtype=float)

# ============================================================
# 3. DECOUVERTE DU DATASET
# ============================================================

print("\n========== DATASET INFORMATION ==========")

print("Nombre d'observations :", X.shape[0])
print("Nombre de variables :", X.shape[1])

print("\nNoms des variables :")
print(variables)

print("\nDimensions du tableau :", X.shape)

# Afficher les 5 premières observations
print("\nLes 5 premières observations :")
print(X[:5, :])

# ============================================================
# 4. CALCUL DE LA MOYENNE
# ============================================================

mean_values = np.mean(X, axis=0)

print("\n========== MOYENNE ==========")

for name, value in zip(variables, mean_values):
    print(f"{name} : {value:.3f}")

# ============================================================
# 5. CALCUL DE L'ECART-TYPE
# ============================================================

std_values = np.std(X, axis=0)

print("\n========== ECART-TYPE ==========")

for name, value in zip(variables, std_values):
    print(f"{name} : {value:.3f}")

# ============================================================
# 6. VALEURS MINIMALES ET MAXIMALES
# ============================================================

min_values = np.min(X, axis=0)
max_values = np.max(X, axis=0)

print("\n========== MINIMUM ET MAXIMUM ==========")

for i, name in enumerate(variables):

    print(
        f"{name} : Min = {min_values[i]:.3f} | "
        f"Max = {max_values[i]:.3f}"
    )

# ============================================================
# 7. MEDIANE
# ============================================================

median_values = np.median(X, axis=0)

print("\n========== MEDIANE ==========")

for name, value in zip(variables, median_values):
    print(f"{name} : {value:.3f}")

# ============================================================
# 8. MATRICE DE CORRELATION
# ============================================================

correlation = np.corrcoef(X, rowvar=False)

print("\n========== MATRICE DE CORRELATION ==========")

print(variables)
print(np.round(correlation, 3))

# ============================================================
# 9. FIN DU PROGRAMME
# ============================================================

print("\n==========================================")
print("ANALYSE NUMPY TERMINEE AVEC SUCCES")
print("==========================================")

workbook.close()


