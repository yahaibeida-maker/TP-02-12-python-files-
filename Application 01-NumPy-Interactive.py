
# ============================================================
# APPLICATION 01 : NUMPY AU SERVICE DE L'AGRONOMIE
# Interface graphique avec Tkinter
# Python - Anaconda - Spyder
# ============================================================

import numpy as np
from openpyxl import load_workbook

import tkinter as tk
from tkinter import ttk

# ============================================================
# 1. CHARGEMENT DES DONNEES EXCEL
# ============================================================

file_path = "OPEN SOURCE DATA FILE.xlsx"

wb = load_workbook(
    file_path,
    data_only=True,
    read_only=True
)

sheet = wb["DATA"]

variables = [
    cell.value for cell in sheet[1]
]

rows = list(
    sheet.iter_rows(
        min_row=2,
        max_col=7,
        values_only=True
    )
)

# Suppression des lignes vides
rows = [
    row for row in rows
    if all(v is not None for v in row)
]

X = np.array(rows, dtype=float)

wb.close()

# ============================================================
# 2. CALCULS NUMPY
# ============================================================

mean_values = np.mean(X, axis=0)
std_values = np.std(X, axis=0)

min_values = np.min(X, axis=0)
max_values = np.max(X, axis=0)

median_values = np.median(X, axis=0)

correlation = np.corrcoef(
    X, rowvar=False
)

# ============================================================
# 3. CREATION DE L'INTERFACE
# ============================================================

root = tk.Tk()

root.title("APPLICATION 01 - NUMPY")
root.geometry("1250x780")
root.configure(bg="#EAF0F7")

# Style général
style = ttk.Style()
style.theme_use("clam")

style.configure(
    "TNotebook.Tab",
    font=("Arial", 12, "bold"),
    padding=[18, 10]
)

style.configure(
    "Treeview",
    font=("Arial", 11),
    rowheight=30
)

style.configure(
    "Treeview.Heading",
    font=("Arial", 11, "bold")
)

# ============================================================
# 4. TITRE PRINCIPAL
# ============================================================

header = tk.Frame(
    root,
    bg="#123B68",
    height=100
)

header.pack(fill="x")

tk.Label(
    header,
    text="APPLICATION 01 : NUMPY",
    font=("Arial", 25, "bold"),
    fg="white",
    bg="#123B68"
).pack(pady=(15, 5))

tk.Label(
    header,
    text="Analyse des données environnementales et agronomiques",
    font=("Arial", 13),
    fg="#DCEBFF",
    bg="#123B68"
).pack()

# ============================================================
# 5. INFORMATIONS GENERALES
# ============================================================

info = tk.Frame(
    root,
    bg="#EAF0F7",
    pady=12
)

info.pack(fill="x")

infos = [
    ("OBSERVATIONS", X.shape[0]),
    ("VARIABLES", X.shape[1]),
    ("VALEURS MANQUANTES", np.isnan(X).sum())
]

for label, value in infos:

    card = tk.Frame(
        info,
        bg="white",
        padx=35,
        pady=10
    )

    card.pack(side="left", expand=True, padx=10)

    tk.Label(
        card,
        text=label,
        font=("Arial", 10, "bold"),
        bg="white",
        fg="#52657A"
    ).pack()

    tk.Label(
        card,
        text=str(value),
        font=("Arial", 22, "bold"),
        bg="white",
        fg="#087E8B"
    ).pack()

# ============================================================
# 6. CREATION DES ONGLETS
# ============================================================

notebook = ttk.Notebook(root)
notebook.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=10
)

tab1 = ttk.Frame(notebook)
tab2 = ttk.Frame(notebook)
tab3 = ttk.Frame(notebook)
tab4 = ttk.Frame(notebook)

notebook.add(tab1, text="01 - Dataset")
notebook.add(tab2, text="02 - Observations")
notebook.add(tab3, text="03 - Statistiques")
notebook.add(tab4, text="04 - Corrélation")

# ============================================================
# 7. FONCTION POUR CREER UN TABLEAU
# ============================================================

def create_table(parent, columns, data):

    frame = ttk.Frame(parent)
    frame.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=20
    )

    tree = ttk.Treeview(
        frame,
        columns=columns,
        show="headings"
    )

    for col in columns:

        tree.heading(
            col,
            text=col
        )

        tree.column(
            col,
            width=135,
            anchor="center"
        )

    scrollbar_y = ttk.Scrollbar(
        frame,
        orient="vertical",
        command=tree.yview
    )

    scrollbar_x = ttk.Scrollbar(
        frame,
        orient="horizontal",
        command=tree.xview
    )

    tree.configure(
        yscrollcommand=scrollbar_y.set,
        xscrollcommand=scrollbar_x.set
    )

    tree.pack(
        side="top",
        fill="both",
        expand=True
    )

    scrollbar_y.pack(
        side="right",
        fill="y"
    )

    scrollbar_x.pack(
        side="bottom",
        fill="x"
    )

    for row in data:
        tree.insert(
            "",
            "end",
            values=row
        )

    return tree

# ============================================================
# 8. ONGLET 01 : DATASET
# ============================================================

tk.Label(
    tab1,
    text="INFORMATIONS SUR LE JEU DE DONNEES",
    font=("Arial", 18, "bold"),
    fg="#123B68"
).pack(pady=25)

dataset_info = [
    ("Nombre d'observations", X.shape[0]),
    ("Nombre de variables", X.shape[1]),
    ("Dimensions", str(X.shape)),
    ("Type des données", str(X.dtype)),
    ("Bibliothèque utilisée", "NumPy")
]

create_table(
    tab1,
    ["Information", "Résultat"],
    dataset_info
)

# ============================================================
# 9. ONGLET 02 : PREMIERES OBSERVATIONS
# ============================================================

tk.Label(
    tab2,
    text="LES 10 PREMIERES OBSERVATIONS",
    font=("Arial", 18, "bold"),
    fg="#087E8B"
).pack(pady=20)

first_rows = [
    [round(float(v), 3) for v in row]
    for row in X[:10]
]

create_table(
    tab2,
    variables,
    first_rows
)

# ============================================================
# 10. ONGLET 03 : STATISTIQUES DESCRIPTIVES
# ============================================================

tk.Label(
    tab3,
    text="STATISTIQUES DESCRIPTIVES",
    font=("Arial", 18, "bold"),
    fg="#087E8B"
).pack(pady=20)

statistics = []

for i, name in enumerate(variables):

    statistics.append([
        name,
        round(mean_values[i], 3),
        round(std_values[i], 3),
        round(median_values[i], 3),
        round(min_values[i], 3),
        round(max_values[i], 3)
    ])

create_table(
    tab3,
    [
        "Variable",
        "Moyenne",
        "Ecart-type",
        "Médiane",
        "Minimum",
        "Maximum"
    ],
    statistics
)

# ============================================================
# 11. ONGLET 04 : MATRICE DE CORRELATION
# ============================================================

tk.Label(
    tab4,
    text="MATRICE DE CORRELATION - PEARSON",
    font=("Arial", 18, "bold"),
    fg="#087E8B"
).pack(pady=20)

corr_data = [
    [name] + [
        round(float(v), 3)
        for v in correlation[i]
    ]
    for i, name in enumerate(variables)
]

create_table(
    tab4,
    ["Variable"] + variables,
    corr_data
)

# ============================================================
# 12. PIED DE PAGE
# ============================================================

footer = tk.Label(
    root,
    text="Python | Anaconda | Spyder | NumPy",
    font=("Arial", 11, "bold"),
    bg="#123B68",
    fg="white",
    pady=10
)

footer.pack(fill="x")

# ============================================================
# 13. LANCEMENT
# ============================================================

root.mainloop()
