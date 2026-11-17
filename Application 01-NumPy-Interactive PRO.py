
# ================================================================
# APPLICATION 01 : NUMPY AU SERVICE DE L'AGRONOMIE
# VERSION 02 : INTERFACE GRAPHIQUE AVANCEE
# Python + Anaconda + Spyder + NumPy + Tkinter
# ================================================================

import numpy as np
from openpyxl import load_workbook, Workbook

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.figure import Figure

from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)

# ================================================================
# 1. CHARGEMENT DES DONNEES
# ================================================================

FILE_PATH = "OPEN SOURCE DATA FILE.xlsx"

try:
    excel = load_workbook(
        FILE_PATH,
        data_only=True,
        read_only=True
    )

    sheet = excel["DATA"]

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

    rows = [
        row for row in rows
        if all(value is not None for value in row)
    ]

    X = np.array(rows, dtype=float)

    excel.close()

    # Elimination des éventuelles lignes non finies
    X = X[np.isfinite(X).all(axis=1)]

except Exception as error:
    raise RuntimeError(
        f"Erreur de chargement du fichier : {error}"
    )

# ================================================================
# 2. CALCULS NUMPY
# ================================================================

mean_values = np.mean(X, axis=0)
std_values = np.std(X, axis=0)

min_values = np.min(X, axis=0)
max_values = np.max(X, axis=0)

median_values = np.median(X, axis=0)

correlation = np.corrcoef(
    X,
    rowvar=False
)

# ================================================================
# 3. CONFIGURATION DE L'INTERFACE
# ================================================================

root = tk.Tk()

root.title(
    "NUMPY PRO - Analyse des données environnementales"
)

root.geometry("1200x800")
root.minsize(950, 650)

root.configure(bg="#EAF0F7")

# Couleurs
BLUE = "#123B68"
GREEN = "#087E55"
LIGHT = "#EAF0F7"
WHITE = "#FFFFFF"
ORANGE = "#E58B18"
PURPLE = "#7044A4"

# Style
style = ttk.Style()
style.theme_use("clam")

style.configure(
    "Treeview",
    font=("Arial", 11),
    rowheight=30
)

style.configure(
    "Treeview.Heading",
    font=("Arial", 11, "bold")
)

style.configure(
    "TCombobox",
    font=("Arial", 12)
)

# ================================================================
# 4. FONCTIONS GRAPHIQUES
# ================================================================

def create_window(title, geometry="1000x650"):

    window = tk.Toplevel(root)

    window.title(title)
    window.geometry(geometry)
    window.configure(bg=LIGHT)

    return window


def create_title(parent, text):

    tk.Label(
        parent,
        text=text,
        font=("Arial", 19, "bold"),
        bg=LIGHT,
        fg=BLUE
    ).pack(pady=20)


def create_table(parent, columns, data):

    frame = tk.Frame(parent, bg=WHITE)

    frame.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=15
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

    scroll_y = ttk.Scrollbar(
        frame,
        orient="vertical",
        command=tree.yview
    )

    tree.configure(
        yscrollcommand=scroll_y.set
    )

    tree.pack(
        side="left",
        fill="both",
        expand=True
    )

    scroll_y.pack(
        side="right",
        fill="y"
    )

    for row in data:
        tree.insert(
            "",
            "end",
            values=row
        )

    return tree


# ================================================================
# 5. FENETRE 01 : INFORMATIONS DU DATASET
# ================================================================

def show_dataset():

    window = create_window(
        "Informations du dataset",
        "800x550"
    )

    create_title(
        window,
        "INFORMATIONS GENERALES DU DATASET"
    )

    data = [
        ("Nombre d'observations", X.shape[0]),
        ("Nombre de variables", X.shape[1]),
        ("Dimensions", str(X.shape)),
        ("Type des données", str(X.dtype)),
        ("Valeurs manquantes", np.isnan(X).sum()),
        ("Bibliothèque", "NumPy"),
        ("Fichier", FILE_PATH)
    ]

    create_table(
        window,
        ["Information", "Résultat"],
        data
    )


# ================================================================
# 6. FENETRE 02 : OBSERVATIONS
# ================================================================

def show_observations():

    window = create_window(
        "Exploration des observations",
        "1150x650"
    )

    create_title(
        window,
        "EXPLORATION DES DONNEES"
    )

    data = [
        [round(float(v), 3) for v in row]
        for row in X[:50]
    ]

    create_table(
        window,
        variables,
        data
    )


# ================================================================
# 7. FENETRE 03 : STATISTIQUES DESCRIPTIVES
# ================================================================

def show_statistics():

    window = create_window(
        "Statistiques descriptives",
        "1150x600"
    )

    create_title(
        window,
        "STATISTIQUES DESCRIPTIVES"
    )

    data = []

    for i, name in enumerate(variables):

        data.append([
            name,
            round(mean_values[i], 3),
            round(std_values[i], 3),
            round(median_values[i], 3),
            round(min_values[i], 3),
            round(max_values[i], 3)
        ])

    create_table(
        window,
        [
            "Variable",
            "Moyenne",
            "Ecart-type",
            "Médiane",
            "Minimum",
            "Maximum"
        ],
        data
    )


# ================================================================
# 8. FENETRE 04 : HISTOGRAMMES INTERACTIFS
# ================================================================

def show_histogram():

    window = create_window(
        "Analyse des distributions",
        "1050x750"
    )

    create_title(
        window,
        "ANALYSE DE LA DISTRIBUTION DES VARIABLES"
    )

    # Panneau de sélection
    control = tk.Frame(
        window,
        bg=WHITE,
        padx=15,
        pady=15
    )

    control.pack(
        fill="x",
        padx=20
    )

    tk.Label(
        control,
        text="Sélectionner une variable :",
        font=("Arial", 12, "bold"),
        bg=WHITE
    ).pack(side="left", padx=10)

    selected = tk.StringVar(
        value=variables[0]
    )

    combo = ttk.Combobox(
        control,
        textvariable=selected,
        values=variables,
        state="readonly",
        width=15
    )

    combo.pack(side="left", padx=10)

    # Figure Matplotlib
    fig, ax = plt.subplots(
        figsize=(9, 5),
        dpi=100
    )

    fig.tight_layout(pad=3)

    canvas = FigureCanvasTkAgg(
        fig,
        master=window
    )

    canvas.get_tk_widget().pack(
        fill="both",
        expand=True,
        padx=20,
        pady=10
    )

    toolbar_frame = tk.Frame(window)
    toolbar_frame.pack(fill="x")

    toolbar = NavigationToolbar2Tk(
        canvas,
        toolbar_frame
    )

    toolbar.update()

    # Actualisation du graphique
    def update_histogram(event=None):

        name = selected.get()

        index = variables.index(name)

        values = X[:, index]

        ax.clear()

        ax.hist(
            values,
            bins=20,
            edgecolor="black",
            alpha=0.8
        )

        ax.axvline(
            np.mean(values),
            linestyle="--",
            linewidth=2,
            color="red",
            label=f"Moyenne = {np.mean(values):.3f}"
        )

        ax.axvline(
            np.median(values),
            linestyle=":",
            linewidth=2,
            color="green",
            label=f"Médiane = {np.median(values):.3f}"
        )

        ax.set_title(
            f"Distribution de {name}",
            fontsize=16,
            fontweight="bold"
        )

        ax.set_xlabel(name, fontsize=13)
        ax.set_ylabel("Fréquence", fontsize=13)

        ax.legend()
        ax.grid(alpha=0.25)

        fig.tight_layout()

        canvas.draw_idle()

    combo.bind(
        "<<ComboboxSelected>>",
        update_histogram
    )

    update_histogram()


# ================================================================
# 9. FENETRE 05 : HEATMAP DE CORRELATION
# ================================================================

def show_heatmap():

    window = create_window(
        "Matrice de corrélation",
        "1000x800"
    )

    create_title(
        window,
        "MATRICE DE CORRELATION DE PEARSON"
    )

    fig, ax = plt.subplots(
        figsize=(8, 6),
        dpi=110
    )

    # Heatmap avec Matplotlib
    im = ax.imshow(
        correlation,
        cmap="coolwarm",
        vmin=-1,
        vmax=1
    )

    ax.set_xticks(range(len(variables)))
    ax.set_yticks(range(len(variables)))

    ax.set_xticklabels(variables, fontsize=12)
    ax.set_yticklabels(variables, fontsize=12)

    # Affichage des coefficients
    for i in range(len(variables)):

        for j in range(len(variables)):

            value = correlation[i, j]

            ax.text(
                j,
                i,
                f"{value:.2f}",
                ha="center",
                va="center",
                fontsize=11,
                fontweight="bold",
                color="black"
            )

    fig.colorbar(
        im,
        ax=ax,
        label="Coefficient de corrélation"
    )

    ax.set_title(
        "Correlation Heatmap",
        fontsize=17,
        fontweight="bold",
        pad=15
    )

    fig.tight_layout()

    canvas = FigureCanvasTkAgg(
        fig,
        master=window
    )

    canvas.draw()

    canvas.get_tk_widget().pack(
        fill="both",
        expand=True,
        padx=20,
        pady=10
    )

    toolbar_frame = tk.Frame(window)
    toolbar_frame.pack(fill="x")

    toolbar = NavigationToolbar2Tk(
        canvas,
        toolbar_frame
    )

    toolbar.update()


# ================================================================
# 10. EXPORTATION DES RESULTATS VERS EXCEL
# ================================================================

def export_excel():

    filename = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[
            ("Fichier Excel", "*.xlsx")
        ],
        initialfile="NumPy_Analysis_Results.xlsx"
    )

    if not filename:
        return

    try:

        output = Workbook()

        # Feuille 1 : Données originales
        ws1 = output.active
        ws1.title = "Dataset"

        ws1.append(variables)

        for row in X:
            ws1.append([
                float(v) for v in row
            ])

        # Feuille 2 : Statistiques
        ws2 = output.create_sheet(
            "Statistics"
        )

        ws2.append([
            "Variable",
            "Mean",
            "Std",
            "Median",
            "Minimum",
            "Maximum"
        ])

        for i, name in enumerate(variables):

            ws2.append([
                name,
                float(mean_values[i]),
                float(std_values[i]),
                float(median_values[i]),
                float(min_values[i]),
                float(max_values[i])
            ])

        # Feuille 3 : Corrélation
        ws3 = output.create_sheet(
            "Correlation"
        )

        ws3.append(["Variable"] + variables)

        for i, name in enumerate(variables):

            ws3.append(
                [name] +
                [
                    float(v)
                    for v in correlation[i]
                ]
            )

        # Enregistrement
        output.save(filename)

        messagebox.showinfo(
            "Exportation réussie",
            "Les résultats ont été exportés avec succès !\n\n"
            + filename
        )

    except Exception as error:

        messagebox.showerror(
            "Erreur",
            str(error)
        )


# ================================================================
# 11. INTERFACE PRINCIPALE
# ================================================================

header = tk.Frame(
    root,
    bg=BLUE,
    height=115
)

header.pack(fill="x")

tk.Label(
    header,
    text="APPLICATION 01 : NUMPY PRO",
    font=("Arial", 26, "bold"),
    bg=BLUE,
    fg=WHITE
).pack(pady=(18, 5))

tk.Label(
    header,
    text="Analyse des données environnementales et agronomiques",
    font=("Arial", 13),
    bg=BLUE,
    fg="#DCEBFF"
).pack()

# ================================================================
# 12. INFORMATIONS DU DATASET
# ================================================================

info = tk.Frame(
    root,
    bg=LIGHT,
    pady=15
)

info.pack(fill="x")

cards = [
    ("OBSERVATIONS", str(X.shape[0]), GREEN),
    ("VARIABLES", str(X.shape[1]), "#2874C7"),
    ("DONNEES VALIDES", str(X.size), PURPLE)
]

for title, value, color in cards:

    card = tk.Frame(
        info,
        bg=WHITE,
        padx=35,
        pady=12
    )

    card.pack(
        side="left",
        expand=True,
        padx=15
    )

    tk.Label(
        card,
        text=title,
        font=("Arial", 10, "bold"),
        bg=WHITE,
        fg="#53657A"
    ).pack()

    tk.Label(
        card,
        text=value,
        font=("Arial", 24, "bold"),
        bg=WHITE,
        fg=color
    ).pack()


# ================================================================
# 13. MENU PRINCIPAL
# ================================================================

tk.Label(
    root,
    text="CHOISIR UNE ANALYSE",
    font=("Arial", 18, "bold"),
    bg=LIGHT,
    fg=BLUE
).pack(pady=15)


def create_button(parent, text, command, color):

    return tk.Button(
        parent,
        text=text,
        command=command,
        font=("Arial", 13, "bold"),
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        relief="flat",
        cursor="hand2",
        height=2,
        width=27
    )


# Panneau des boutons
menu = tk.Frame(
    root,
    bg=LIGHT
)

menu.pack(
    expand=True,
    pady=10
)

buttons = [
    ("01 - Informations du dataset",
     show_dataset, BLUE),

    ("02 - Explorer les observations",
     show_observations, GREEN),

    ("03 - Statistiques descriptives",
     show_statistics, "#2874C7"),

    ("04 - Histogrammes interactifs",
     show_histogram, ORANGE),

    ("05 - Heatmap de corrélation",
     show_heatmap, PURPLE)
]

for i, (text, command, color) in enumerate(buttons):

    row = i // 2
    col = i % 2

    button = create_button(
        menu,
        text,
        command,
        color
    )

    button.grid(
        row=row,
        column=col,
        padx=15,
        pady=10
    )

# Bouton exportation
export_button = tk.Button(
    root,
    text="EXPORTER LES RESULTATS VERS EXCEL",
    command=export_excel,
    font=("Arial", 15, "bold"),
    bg="#087E55",
    fg="white",
    relief="flat",
    cursor="hand2",
    padx=20,
    pady=12
)

export_button.pack(pady=20)

# ================================================================
# 14. PIED DE PAGE
# ================================================================

tk.Label(
    root,
    text="PYTHON  |  ANACONDA  |  SPYDER  |  NUMPY",
    font=("Arial", 11, "bold"),
    bg=BLUE,
    fg=WHITE,
    pady=10
).pack(fill="x")

# ================================================================
# 15. LANCEMENT DE L'APPLICATION
# ================================================================

root.mainloop()
