# ================================================================
# APPLICATION 02 : PANDAS AU SERVICE DE L'AGRONOMIE
# VERSION PRO : INTERFACE GRAPHIQUE TKINTER
# Python + Anaconda + Spyder + Pandas
#
# Fonctionnalites :
# 1. Tableau de bord et informations du dataset
# 2. Exploration des observations
# 3. Qualite et nettoyage des donnees
# 4. Statistiques descriptives
# 5. Selection et filtrage des variables
# 6. Groupement et agregation
# 7. Exportation Excel / CSV
# ================================================================

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import numpy as np
from openpyxl import load_workbook

# ================================================================
# 1. CONFIGURATION GENERALE
# ================================================================

DEFAULT_FILE = "OPEN SOURCE DATA FILE.xlsx"
SHEET_NAME = "DATA"

BLUE = "#123B68"
GREEN = "#087E55"
LIGHT = "#EAF0F7"
WHITE = "#FFFFFF"
ORANGE = "#E58B18"
PURPLE = "#7044A4"
TEXT = "#26384A"

root = tk.Tk()
root.title("PANDAS PRO - Analyse des donnees agronomiques")
root.geometry("1280x860")
root.minsize(1050, 700)
root.configure(bg=LIGHT)

style = ttk.Style(root)
style.theme_use("clam")
style.configure("Treeview", font=("Arial", 10), rowheight=28)
style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
style.configure(
    "Custom.Horizontal.TProgressbar",
    troughcolor="#DCE6F0",
    background=GREEN,
    thickness=18
)

# ================================================================
# 2. ECRAN DE CHARGEMENT
# ================================================================

loading_frame = tk.Frame(root, bg=LIGHT)
loading_frame.pack(fill="both", expand=True)

tk.Label(
    loading_frame,
    text="APPLICATION 02 : PANDAS PRO",
    font=("Arial", 26, "bold"),
    bg=BLUE,
    fg=WHITE,
    pady=25
).pack(fill="x")

loading_box = tk.Frame(loading_frame, bg=WHITE, padx=35, pady=30)
loading_box.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.78)

tk.Label(
    loading_box,
    text="CHARGEMENT DU JEU DE DONNEES",
    font=("Arial", 20, "bold"),
    bg=WHITE,
    fg=BLUE
).pack(pady=(5, 15))

loading_status = tk.Label(
    loading_box,
    text="Initialisation...",
    font=("Arial", 12),
    bg=WHITE,
    fg=TEXT
)
loading_status.pack(pady=8)

progress = ttk.Progressbar(
    loading_box,
    style="Custom.Horizontal.TProgressbar",
    orient="horizontal",
    mode="determinate",
    maximum=100,
    value=0
)
progress.pack(fill="x", pady=12)

progress_percent = tk.Label(
    loading_box,
    text="0 %",
    font=("Arial", 13, "bold"),
    bg=WHITE,
    fg=GREEN
)
progress_percent.pack()

def set_progress(value, status):
    progress["value"] = value
    progress_percent.config(text=f"{int(value)} %")
    loading_status.config(text=status)
    root.update_idletasks()

root.update_idletasks()

# ================================================================
# 3. SELECTION ET IMPORTATION EXCEL AVEC PANDAS
# ================================================================

candidate_paths = [
    os.path.abspath(DEFAULT_FILE),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_FILE)
    if "__file__" in globals() else ""
]

FILE_PATH = next(
    (p for p in candidate_paths if p and os.path.isfile(p)),
    None
)

if FILE_PATH is None:
    root.withdraw()
    FILE_PATH = filedialog.askopenfilename(
        parent=root,
        title="Selectionner le fichier Excel",
        filetypes=[("Fichiers Excel", "*.xlsx *.xls *.xlsm")]
    )
    root.deiconify()

if not FILE_PATH:
    root.destroy()
    raise SystemExit("Aucun fichier selectionne.")

try:
    set_progress(15, "Lecture du fichier Excel avec Pandas...")

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=SHEET_NAME
    )

    set_progress(55, "Preparation du DataFrame...")

    # Supprimer uniquement les lignes entierement vides.
    initial_rows = len(df)
    df = df.dropna(how="all").copy()
    empty_rows_removed = initial_rows - len(df)

    # Supprimer les colonnes entierement vides.
    empty_columns = df.columns[df.isna().all()].tolist()
    if empty_columns:
        df = df.drop(columns=empty_columns)

    # Nettoyage des noms de colonnes.
    df.columns = [str(col).strip() for col in df.columns]

    if df.empty or len(df.columns) == 0:
        raise ValueError("Le fichier ne contient aucune donnee exploitable.")

    set_progress(100, "DataFrame charge avec succes.")

except Exception as error:
    messagebox.showerror(
        "Erreur de chargement",
        f"Impossible de lire le fichier Excel.\n\n{error}",
        parent=root
    )
    root.destroy()
    raise SystemExit

# ================================================================
# 4. INFORMATIONS ET FONCTIONS COMMUNES
# ================================================================

def create_window(title, geometry="1100x700"):
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry(geometry)
    win.configure(bg=LIGHT)
    return win


def create_title(parent, text):
    tk.Label(
        parent,
        text=text,
        font=("Arial", 18, "bold"),
        bg=LIGHT,
        fg=BLUE
    ).pack(pady=18)


def create_table(parent, columns, data, widths=None):
    """Affiche un tableau avec defilement horizontal et vertical."""
    frame = tk.Frame(parent, bg=WHITE)
    frame.pack(fill="both", expand=True, padx=18, pady=12)

    tree = ttk.Treeview(frame, columns=columns, show="headings")

    for i, col in enumerate(columns):
        tree.heading(col, text=str(col))
        width = widths[i] if widths and i < len(widths) else 130
        tree.column(col, width=width, minwidth=85, anchor="center")

    scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)

    tree.configure(
        yscrollcommand=scroll_y.set,
        xscrollcommand=scroll_x.set
    )

    tree.grid(row=0, column=0, sticky="nsew")
    scroll_y.grid(row=0, column=1, sticky="ns")
    scroll_x.grid(row=1, column=0, sticky="ew")

    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    for row in data:
        tree.insert("", "end", values=[
            "" if pd.isna(value) else value for value in row
        ])

    return tree


def format_value(value):
    if pd.isna(value):
        return "NaN"
    if isinstance(value, (float, np.floating)):
        return round(float(value), 4)
    if isinstance(value, (np.integer,)):
        return int(value)
    return value


# ================================================================
# 5. FENETRE 01 : INFORMATIONS DU DATASET
# ================================================================

def show_dataset():
    win = create_window("Informations du dataset", "950x650")
    create_title(win, "INFORMATIONS GENERALES DU DATASET")

    data = [
        ("Fichier source", os.path.basename(FILE_PATH)),
        ("Feuille Excel", SHEET_NAME),
        ("Emplacement", os.path.abspath(FILE_PATH)),
        ("Nombre de lignes", df.shape[0]),
        ("Nombre de colonnes", df.shape[1]),
        ("Dimensions", str(df.shape)),
        ("Types de donnees", ", ".join(sorted(set(df.dtypes.astype(str))))),
        ("Lignes vides supprimees", empty_rows_removed),
        ("Valeurs manquantes", int(df.isna().sum().sum())),
        ("Lignes dupliquees", int(df.duplicated().sum())),
        ("Bibliotheque", "Pandas")
    ]

    create_table(
        win,
        ["Information", "Resultat"],
        data,
        widths=[250, 620]
    )


# ================================================================
# 6. FENETRE 02 : EXPLORATION DES OBSERVATIONS
# ================================================================

def show_observations():
    win = create_window("Exploration des observations", "1250x720")
    create_title(win, "EXPLORATION DU DATAFRAME")

    control = tk.Frame(win, bg=WHITE, padx=15, pady=10)
    control.pack(fill="x", padx=18)

    tk.Label(
        control,
        text="Nombre de lignes a afficher :",
        font=("Arial", 11, "bold"),
        bg=WHITE
    ).pack(side="left", padx=5)

    limit_var = tk.StringVar(value="20")
    limit_combo = ttk.Combobox(
        control,
        textvariable=limit_var,
        values=["5", "10", "20", "50", "100", "Toutes"],
        state="readonly",
        width=12
    )
    limit_combo.pack(side="left", padx=8)

    table_holder = tk.Frame(win, bg=LIGHT)
    table_holder.pack(fill="both", expand=True)

    def refresh(event=None):
        for child in table_holder.winfo_children():
            child.destroy()

        if limit_var.get() == "Toutes":
            view = df
        else:
            view = df.head(int(limit_var.get()))

        data = [
            [format_value(v) for v in row]
            for row in view.itertuples(index=False, name=None)
        ]

        create_table(
            table_holder,
            list(df.columns),
            data
        )

    limit_combo.bind("<<ComboboxSelected>>", refresh)
    refresh()


# ================================================================
# 7. FENETRE 03 : QUALITE DES DONNEES
# ================================================================

def show_data_quality():
    win = create_window("Qualite et nettoyage des donnees", "1150x700")
    create_title(win, "DIAGNOSTIC DE LA QUALITE DES DONNEES")

    missing = df.isna().sum()
    missing_percent = (missing / len(df) * 100) if len(df) else missing

    duplicated = int(df.duplicated().sum())

    data = []
    for col in df.columns:
        data.append([
            col,
            int(df[col].isna().sum()),
            round(float(missing_percent[col]), 2),
            int(df[col].nunique(dropna=True)),
            str(df[col].dtype)
        ])

    create_table(
        win,
        ["Variable", "Valeurs manquantes", "Manquants (%)",
         "Valeurs uniques", "Type"],
        data
    )

    summary = tk.Label(
        win,
        text=f"Lignes dupliquees : {duplicated}    |    "
             f"Valeurs manquantes au total : {int(missing.sum())}",
        font=("Arial", 12, "bold"),
        bg=WHITE,
        fg=BLUE,
        pady=12
    )
    summary.pack(fill="x", padx=18, pady=(0, 15))

    def remove_duplicates():
        nonlocal_dummy = None
        global df
        before = len(df)
        df = df.drop_duplicates().copy()
        removed = before - len(df)
        messagebox.showinfo(
            "Nettoyage termine",
            f"{removed} ligne(s) dupliquee(s) supprimee(s).\n"
            "Les autres valeurs n'ont pas ete modifiees.",
            parent=win
        )
        win.destroy()

    tk.Button(
        win,
        text="SUPPRIMER LES LIGNES DUPLIQUEES",
        command=remove_duplicates,
        font=("Arial", 11, "bold"),
        bg=ORANGE,
        fg=WHITE,
        relief="flat",
        padx=15,
        pady=10
    ).pack(pady=8)


# ================================================================
# 8. FENETRE 04 : STATISTIQUES DESCRIPTIVES
# ================================================================

def show_statistics():
    win = create_window("Statistiques descriptives", "1250x650")
    create_title(win, "STATISTIQUES DESCRIPTIVES")

    numeric_df = df.select_dtypes(include=np.number)

    if numeric_df.empty:
        messagebox.showwarning(
            "Aucune variable numerique",
            "Le DataFrame ne contient pas de colonne numerique.",
            parent=win
        )
        return

    description = numeric_df.describe().T
    data = []

    for variable, row in description.iterrows():
        data.append([
            variable,
            format_value(row.get("count")),
            format_value(row.get("mean")),
            format_value(row.get("std")),
            format_value(row.get("min")),
            format_value(row.get("25%")),
            format_value(row.get("50%")),
            format_value(row.get("75%")),
            format_value(row.get("max"))
        ])

    create_table(
        win,
        ["Variable", "N", "Moyenne", "Ecart-type", "Min",
         "Q1 (25%)", "Mediane", "Q3 (75%)", "Max"],
        data
    )


# ================================================================
# 9. FENETRE 05 : SELECTION ET FILTRAGE
# ================================================================

def show_filter():
    win = create_window("Selection et filtrage des donnees", "1150x720")
    create_title(win, "SELECTION ET FILTRAGE DES DONNEES")

    numeric_columns = list(
        df.select_dtypes(include=np.number).columns
    )

    if not numeric_columns:
        messagebox.showwarning(
            "Filtrage indisponible",
            "Aucune variable numerique n'a ete detectee.",
            parent=win
        )
        return

    control = tk.Frame(win, bg=WHITE, padx=15, pady=15)
    control.pack(fill="x", padx=18)

    tk.Label(
        control, text="Variable :", font=("Arial", 11, "bold"),
        bg=WHITE
    ).grid(row=0, column=0, padx=5, pady=5)

    col_var = tk.StringVar(value=str(numeric_columns[0]))
    col_combo = ttk.Combobox(
        control, textvariable=col_var,
        values=[str(c) for c in numeric_columns],
        state="readonly", width=15
    )
    col_combo.grid(row=0, column=1, padx=5)

    tk.Label(
        control, text="Operateur :", font=("Arial", 11, "bold"),
        bg=WHITE
    ).grid(row=0, column=2, padx=5)

    op_var = tk.StringVar(value=">")
    op_combo = ttk.Combobox(
        control, textvariable=op_var,
        values=[">", ">=", "<", "<=", "==", "!="],
        state="readonly", width=8
    )
    op_combo.grid(row=0, column=3, padx=5)

    tk.Label(
        control, text="Valeur :", font=("Arial", 11, "bold"),
        bg=WHITE
    ).grid(row=0, column=4, padx=5)

    value_entry = tk.Entry(control, width=12, font=("Arial", 11))
    value_entry.grid(row=0, column=5, padx=5)
    value_entry.insert(0, "10")

    result_label = tk.Label(
        win, text="Aucun filtre applique.",
        font=("Arial", 11, "bold"), bg=LIGHT, fg=BLUE
    )
    result_label.pack(pady=8)

    holder = tk.Frame(win, bg=LIGHT)
    holder.pack(fill="both", expand=True)

    def apply_filter():
        try:
            column = col_var.get()
            threshold = float(value_entry.get())
            series = df[column]

            operations = {
                ">": series > threshold,
                ">=": series >= threshold,
                "<": series < threshold,
                "<=": series <= threshold,
                "==": series == threshold,
                "!=": series != threshold
            }

            filtered = df.loc[operations[op_var.get()]]

            for child in holder.winfo_children():
                child.destroy()

            data = [
                [format_value(v) for v in row]
                for row in filtered.head(200).itertuples(
                    index=False, name=None
                )
            ]

            create_table(holder, list(df.columns), data)

            result_label.config(
                text=f"{len(filtered)} observation(s) correspondent "
                     f"a la condition. Affichage limite aux 200 premieres."
            )

        except ValueError:
            messagebox.showerror(
                "Valeur invalide",
                "Veuillez saisir une valeur numerique valide.",
                parent=win
            )

    tk.Button(
        control,
        text="APPLIQUER LE FILTRE",
        command=apply_filter,
        bg=GREEN,
        fg=WHITE,
        font=("Arial", 10, "bold"),
        relief="flat",
        padx=12,
        pady=7
    ).grid(row=0, column=6, padx=10)


# ================================================================
# 10. FENETRE 06 : GROUPEMENT ET AGREGATION
# ================================================================

def show_groupby():
    win = create_window("Groupement et agregation", "1100x700")
    create_title(win, "GROUPEMENT ET AGREGATION AVEC GROUPBY")

    numeric_columns = list(
        df.select_dtypes(include=np.number).columns
    )
    categorical_columns = list(
        df.select_dtypes(exclude=np.number).columns
    )

    if not categorical_columns:
        tk.Label(
            win,
            text="Aucune variable categorielle detectee.\n"
                 "Le groupement necessite une colonne de categorie.",
            font=("Arial", 13),
            bg=LIGHT,
            fg=TEXT
        ).pack(pady=30)
        return

    if not numeric_columns:
        tk.Label(
            win,
            text="Aucune variable numerique disponible pour l'agregation.",
            font=("Arial", 13),
            bg=LIGHT,
            fg=TEXT
        ).pack(pady=30)
        return

    control = tk.Frame(win, bg=WHITE, padx=15, pady=15)
    control.pack(fill="x", padx=18)

    tk.Label(control, text="Grouper par :", bg=WHITE).grid(
        row=0, column=0, padx=5
    )
    group_var = tk.StringVar(value=str(categorical_columns[0]))
    group_combo = ttk.Combobox(
        control, textvariable=group_var,
        values=[str(c) for c in categorical_columns],
        state="readonly", width=18
    )
    group_combo.grid(row=0, column=1, padx=5)

    tk.Label(control, text="Variable numerique :", bg=WHITE).grid(
        row=0, column=2, padx=5
    )
    value_var = tk.StringVar(value=str(numeric_columns[0]))
    value_combo = ttk.Combobox(
        control, textvariable=value_var,
        values=[str(c) for c in numeric_columns],
        state="readonly", width=18
    )
    value_combo.grid(row=0, column=3, padx=5)

    tk.Label(control, text="Agregation :", bg=WHITE).grid(
        row=0, column=4, padx=5
    )
    agg_var = tk.StringVar(value="mean")
    agg_combo = ttk.Combobox(
        control, textvariable=agg_var,
        values=["mean", "median", "sum", "min", "max", "count"],
        state="readonly", width=12
    )
    agg_combo.grid(row=0, column=5, padx=5)

    holder = tk.Frame(win, bg=LIGHT)
    holder.pack(fill="both", expand=True)

    def run_groupby():
        try:
            result = (
                df.groupby(group_var.get())[value_var.get()]
                .agg(agg_var.get())
                .reset_index()
            )

            for child in holder.winfo_children():
                child.destroy()

            data = [
                [format_value(v) for v in row]
                for row in result.itertuples(index=False, name=None)
            ]

            create_table(holder, list(result.columns), data)

        except Exception as error:
            messagebox.showerror(
                "Erreur GroupBy", str(error), parent=win
            )

    tk.Button(
        control,
        text="CALCULER",
        command=run_groupby,
        bg=PURPLE,
        fg=WHITE,
        font=("Arial", 10, "bold"),
        relief="flat",
        padx=12,
        pady=7
    ).grid(row=0, column=6, padx=10)


# ================================================================
# 11. EXPORTATION EXCEL ET CSV
# ================================================================

def export_data():
    choice = messagebox.askyesnocancel(
        "Format d'exportation",
        "Oui : Excel (.xlsx)\nNon : CSV (.csv)\nAnnuler : abandonner",
        parent=root
    )

    if choice is None:
        return

    if choice:
        filename = filedialog.asksaveasfilename(
            parent=root,
            defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")],
            initialfile="Pandas_Analysis_Results.xlsx"
        )
        if not filename:
            return

        try:
            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Dataset", index=False)
                df.describe(include="all").to_excel(
                    writer, sheet_name="Description"
                )
                df.isna().sum().to_frame(
                    "Missing values"
                ).to_excel(writer, sheet_name="Data_Quality")

            messagebox.showinfo(
                "Exportation reussie",
                f"Fichier Excel enregistre :\n{filename}",
                parent=root
            )
        except Exception as error:
            messagebox.showerror(
                "Erreur d'exportation", str(error), parent=root
            )

    else:
        filename = filedialog.asksaveasfilename(
            parent=root,
            defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv")],
            initialfile="Pandas_Dataset.csv"
        )
        if not filename:
            return

        try:
            df.to_csv(filename, index=False, encoding="utf-8-sig")
            messagebox.showinfo(
                "Exportation reussie",
                f"Fichier CSV enregistre :\n{filename}",
                parent=root
            )
        except Exception as error:
            messagebox.showerror(
                "Erreur d'exportation", str(error), parent=root
            )


# ================================================================
# 12. CONSTRUCTION DE L'INTERFACE PRINCIPALE
# ================================================================

loading_frame.destroy()

header = tk.Frame(root, bg=BLUE, height=115)
header.pack(fill="x")

tk.Label(
    header,
    text="APPLICATION 02 : PANDAS PRO",
    font=("Arial", 26, "bold"),
    bg=BLUE,
    fg=WHITE
).pack(pady=(18, 5))

tk.Label(
    header,
    text="Manipulation et analyse des donnees agronomiques",
    font=("Arial", 13),
    bg=BLUE,
    fg="#DCEBFF"
).pack()

# Panneau fichier source
file_panel = tk.Frame(root, bg=WHITE, padx=20, pady=10)
file_panel.pack(fill="x", padx=24, pady=10)

tk.Label(
    file_panel,
    text="FICHIER EXCEL UTILISE",
    font=("Arial", 11, "bold"),
    fg=BLUE,
    bg=WHITE
).pack(anchor="w")

tk.Label(
    file_panel,
    text=os.path.basename(FILE_PATH),
    font=("Arial", 13, "bold"),
    fg=GREEN,
    bg=WHITE
).pack(anchor="w", pady=2)

tk.Label(
    file_panel,
    text="Emplacement : " + os.path.abspath(FILE_PATH),
    font=("Arial", 9),
    fg="#52657A",
    bg=WHITE,
    wraplength=1180,
    justify="left"
).pack(anchor="w")

progress_ok = ttk.Progressbar(
    file_panel,
    style="Custom.Horizontal.TProgressbar",
    maximum=100,
    value=100
)
progress_ok.pack(fill="x", pady=(9, 3))

tk.Label(
    file_panel,
    text="✓ DATAFRAME CHARGE AVEC SUCCES",
    font=("Arial", 10, "bold"),
    fg=GREEN,
    bg=WHITE
).pack(anchor="w")

# Cartes d'information
info = tk.Frame(root, bg=LIGHT, pady=8)
info.pack(fill="x")

cards = [
    ("OBSERVATIONS", str(df.shape[0]), GREEN),
    ("VARIABLES", str(df.shape[1]), "#2874C7"),
    ("VALEURS MANQUANTES", str(int(df.isna().sum().sum())), PURPLE)
]

for title, value, color in cards:
    card = tk.Frame(info, bg=WHITE, padx=30, pady=9)
    card.pack(side="left", expand=True, padx=12)

    tk.Label(
        card, text=title,
        font=("Arial", 9, "bold"),
        bg=WHITE, fg="#53657A"
    ).pack()

    tk.Label(
        card, text=value,
        font=("Arial", 22, "bold"),
        bg=WHITE, fg=color
    ).pack()

tk.Label(
    root,
    text="CHOISIR UNE ANALYSE",
    font=("Arial", 18, "bold"),
    bg=LIGHT,
    fg=BLUE
).pack(pady=10)


def create_button(parent, text, command, color):
    return tk.Button(
        parent,
        text=text,
        command=command,
        font=("Arial", 12, "bold"),
        bg=color,
        fg=WHITE,
        activebackground=color,
        activeforeground=WHITE,
        relief="flat",
        cursor="hand2",
        height=2,
        width=36,
        bd=0
    )


menu = tk.Frame(root, bg=LIGHT)
menu.pack(expand=True, pady=5)

buttons = [
    ("▣  01 - Informations du dataset", show_dataset, BLUE),
    ("▤  02 - Explorer les observations", show_observations, GREEN),
    ("▥  03 - Qualite des donnees", show_data_quality, "#2874C7"),
    ("▦  04 - Statistiques descriptives", show_statistics, ORANGE),
    ("⌕  05 - Selection et filtrage", show_filter, PURPLE),
    ("▧  06 - Groupement et agregation", show_groupby, "#087E55")
]

for i, (label, command, color) in enumerate(buttons):
    row = i // 2
    col = i % 2
    create_button(menu, label, command, color).grid(
        row=row, column=col, padx=14, pady=9, sticky="nsew"
    )

menu.grid_columnconfigure(0, weight=1)
menu.grid_columnconfigure(1, weight=1)

export_button = tk.Button(
    root,
    text="EXPORTER LES RESULTATS : EXCEL / CSV",
    command=export_data,
    font=("Arial", 14, "bold"),
    bg=GREEN,
    fg=WHITE,
    activebackground="#066644",
    activeforeground=WHITE,
    relief="flat",
    cursor="hand2",
    padx=22,
    pady=11
)
export_button.pack(pady=12)

tk.Label(
    root,
    text="PYTHON  |  ANACONDA  |  SPYDER  |  PANDAS",
    font=("Arial", 10, "bold"),
    bg=BLUE,
    fg=WHITE,
    pady=9
).pack(fill="x")

# ================================================================
# 13. LANCEMENT DE L'APPLICATION
# ================================================================

root.mainloop()

