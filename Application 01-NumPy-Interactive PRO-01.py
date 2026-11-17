# ================================================================
# APPLICATION 01 : NUMPY AU SERVICE DE L'AGRONOMIE
# VERSION 03 : INTERFACE GRAPHIQUE AMELIOREE
# Python + Anaconda + Spyder + NumPy + Tkinter
# ================================================================

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
from openpyxl import load_workbook, Workbook

# Figure directe : evite le conflit TkAgg / Qt dans Spyder.
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)

# ================================================================
# 1. CONFIGURATION
# ================================================================

DEFAULT_FILE = "OPEN SOURCE DATA FILE.xlsx"
SHEET_NAME = "DATA"

BLUE = "#123B68"
GREEN = "#087E55"
LIGHT = "#EAF0F7"
WHITE = "#FFFFFF"
ORANGE = "#E58B18"
PURPLE = "#7044A4"

root = tk.Tk()
root.title("NUMPY PRO - Analyse des donnees environnementales")
root.geometry("1250x850")
root.minsize(1000, 700)
root.configure(bg=LIGHT)

style = ttk.Style(root)
style.theme_use("clam")
style.configure("Treeview", font=("Arial", 11), rowheight=30)
style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
style.configure(
    "Custom.Horizontal.TProgressbar",
    troughcolor="#DCE6F0",
    background=GREEN,
    thickness=18
)

# ================================================================
# 2. ECRAN ET BARRE DE PROGRESSION DU CHARGEMENT
# ================================================================

loading_frame = tk.Frame(root, bg=LIGHT)
loading_frame.pack(fill="both", expand=True)

tk.Label(
    loading_frame,
    text="APPLICATION 01 : NUMPY PRO",
    font=("Arial", 26, "bold"),
    bg=BLUE,
    fg=WHITE,
    pady=25
).pack(fill="x")

loading_box = tk.Frame(
    loading_frame, bg=WHITE, padx=35, pady=30
)
loading_box.place(
    relx=0.5, rely=0.5, anchor="center", relwidth=0.78
)

tk.Label(
    loading_box,
    text="CHARGEMENT DES DONNEES",
    font=("Arial", 20, "bold"),
    bg=WHITE,
    fg=BLUE
).pack(pady=(5, 15))

loading_status = tk.Label(
    loading_box,
    text="Initialisation...",
    font=("Arial", 12),
    bg=WHITE
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
    value = max(0, min(100, int(value)))
    progress["value"] = value
    progress_percent.config(text=f"{value} %")
    loading_status.config(text=status)
    root.update_idletasks()

root.update_idletasks()

# ================================================================
# 3. SELECTION ET LECTURE DU FICHIER EXCEL
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
        title="Selectionner le fichier Excel des donnees",
        filetypes=[("Fichiers Excel", "*.xlsx *.xlsm")]
    )
    root.deiconify()

if not FILE_PATH:
    root.destroy()
    raise SystemExit("Aucun fichier Excel selectionne.")

try:
    set_progress(5, "Ouverture du fichier Excel...")

    excel = load_workbook(
        FILE_PATH, data_only=True, read_only=True
    )

    if SHEET_NAME not in excel.sheetnames:
        raise ValueError(
            f"Feuille '{SHEET_NAME}' introuvable. "
            f"Feuilles disponibles : {excel.sheetnames}"
        )

    sheet = excel[SHEET_NAME]

    header_values = next(
        sheet.iter_rows(
            min_row=1, max_row=1, values_only=True
        )
    )

    valid_columns = [
        i for i, value in enumerate(header_values)
        if value is not None
    ]
    variables = [header_values[i] for i in valid_columns]

    if not variables:
        raise ValueError("Aucune variable detectee dans la premiere ligne.")

    set_progress(10, "Lecture des observations...")

    total_rows = max(sheet.max_row - 1, 1)
    rows = []
    skipped_empty = 0
    skipped_invalid = 0

    for row_number, row in enumerate(
        sheet.iter_rows(
            min_row=2,
            max_col=max(valid_columns) + 1,
            values_only=True
        ),
        start=1
    ):
        selected = tuple(row[i] for i in valid_columns)

        if all(value is None for value in selected):
            skipped_empty += 1
            continue

        try:
            numeric_row = [float(value) for value in selected]
            if not np.isfinite(numeric_row).all():
                skipped_invalid += 1
                continue
            rows.append(numeric_row)
        except (TypeError, ValueError):
            skipped_invalid += 1
            continue

        if row_number % 20 == 0 or row_number == total_rows:
            percent = 10 + 75 * row_number / total_rows
            set_progress(
                percent,
                f"Lecture des observations : {row_number}/{total_rows}"
            )

    excel.close()

    if not rows:
        raise ValueError("Aucune observation numerique valide trouvee.")

    X = np.asarray(rows, dtype=float)

    if X.ndim != 2 or X.shape[1] != len(variables):
        raise ValueError("La structure du tableau de donnees est invalide.")

    set_progress(90, "Calcul des statistiques avec NumPy...")

    # ============================================================
    # 4. CALCULS NUMPY
    # ============================================================

    mean_values = np.mean(X, axis=0)
    std_values = np.std(X, axis=0, ddof=0)
    min_values = np.min(X, axis=0)
    max_values = np.max(X, axis=0)
    median_values = np.median(X, axis=0)
    correlation = np.corrcoef(X, rowvar=False)

    set_progress(100, "Donnees chargees et analysees avec succes.")

except Exception as error:
    try:
        excel.close()
    except Exception:
        pass
    messagebox.showerror(
        "Erreur de chargement",
        f"Impossible de charger ou d'analyser le fichier.\n\n{error}",
        parent=root
    )
    root.destroy()
    raise SystemExit

# ================================================================
# 5. FONCTIONS GENERALES
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


def create_table(parent, columns, data, widths=None):
    frame = tk.Frame(parent, bg=WHITE)
    frame.pack(fill="both", expand=True, padx=20, pady=15)

    tree = ttk.Treeview(
        frame, columns=columns, show="headings"
    )

    for index, col in enumerate(columns):
        tree.heading(col, text=col)
        width = widths[index] if widths else 135
        tree.column(
            col, width=width, minwidth=90, anchor="center"
        )

    scroll_y = ttk.Scrollbar(
        frame, orient="vertical", command=tree.yview
    )
    scroll_x = ttk.Scrollbar(
        frame, orient="horizontal", command=tree.xview
    )

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
        tree.insert("", "end", values=row)

    return tree

# ================================================================
# 6. FENETRE 01 : INFORMATIONS DU DATASET
# ================================================================

def show_dataset():
    window = create_window("Informations du dataset", "900x600")
    create_title(window, "INFORMATIONS GENERALES DU DATASET")

    data = [
        ("Fichier source", os.path.basename(FILE_PATH)),
        ("Emplacement", os.path.abspath(FILE_PATH)),
        ("Feuille Excel", SHEET_NAME),
        ("Nombre d'observations", X.shape[0]),
        ("Nombre de variables", X.shape[1]),
        ("Dimensions", str(X.shape)),
        ("Type des donnees", str(X.dtype)),
        ("Valeurs manquantes", int(np.isnan(X).sum())),
        ("Lignes vides ignorees", skipped_empty),
        ("Lignes invalides ignorees", skipped_invalid),
        ("Bibliotheque de calcul", "NumPy")
    ]

    create_table(
        window, ["Information", "Resultat"], data,
        widths=[230, 600]
    )

# ================================================================
# 7. FENETRE 02 : OBSERVATIONS
# ================================================================

def show_observations():
    window = create_window(
        "Exploration des observations", "1200x700"
    )
    create_title(
        window, "EXPLORATION DES 50 PREMIERES OBSERVATIONS"
    )

    data = [
        [round(float(v), 3) for v in row]
        for row in X[:50]
    ]
    create_table(window, variables, data)

# ================================================================
# 8. FENETRE 03 : STATISTIQUES DESCRIPTIVES
# ================================================================

def show_statistics():
    window = create_window(
        "Statistiques descriptives", "1200x600"
    )
    create_title(window, "STATISTIQUES DESCRIPTIVES")

    data = []
    for i, name in enumerate(variables):
        data.append([
            name,
            round(float(mean_values[i]), 3),
            round(float(std_values[i]), 3),
            round(float(median_values[i]), 3),
            round(float(min_values[i]), 3),
            round(float(max_values[i]), 3)
        ])

    create_table(
        window,
        [
            "Variable", "Moyenne", "Ecart-type",
            "Mediane", "Minimum", "Maximum"
        ],
        data
    )

# ================================================================
# 9. FENETRE 04 : HISTOGRAMMES INTERACTIFS
# ================================================================

def show_histogram():
    window = create_window(
        "Analyse des distributions", "1100x800"
    )
    create_title(
        window, "ANALYSE DE LA DISTRIBUTION DES VARIABLES"
    )

    control = tk.Frame(
        window, bg=WHITE, padx=15, pady=15
    )
    control.pack(fill="x", padx=20)

    tk.Label(
        control,
        text="Selectionner une variable :",
        font=("Arial", 12, "bold"),
        bg=WHITE
    ).pack(side="left", padx=10)

    selected = tk.StringVar(value=str(variables[0]))

    combo = ttk.Combobox(
        control,
        textvariable=selected,
        values=[str(v) for v in variables],
        state="readonly",
        width=18
    )
    combo.pack(side="left", padx=10)

    fig = Figure(figsize=(9, 5), dpi=100)
    ax = fig.add_subplot(111)

    canvas = FigureCanvasTkAgg(fig, master=window)
    canvas.get_tk_widget().pack(
        fill="both", expand=True, padx=20, pady=10
    )

    toolbar_frame = tk.Frame(window)
    toolbar_frame.pack(fill="x")

    toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
    toolbar.update()

    def update_histogram(event=None):
        name = selected.get()
        index = [str(v) for v in variables].index(name)
        values = X[:, index]

        ax.clear()
        ax.hist(
            values, bins=20, edgecolor="black", alpha=0.80
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
            label=f"Mediane = {np.median(values):.3f}"
        )

        ax.set_title(
            f"Distribution de {name}",
            fontsize=16,
            fontweight="bold"
        )
        ax.set_xlabel(name, fontsize=13)
        ax.set_ylabel("Frequence", fontsize=13)
        ax.legend()
        ax.grid(alpha=0.25)

        fig.tight_layout()
        canvas.draw_idle()

    combo.bind("<<ComboboxSelected>>", update_histogram)
    update_histogram()

# ================================================================
# 10. FENETRE 05 : HEATMAP DE CORRELATION
# ================================================================

def show_heatmap():
    window = create_window(
        "Matrice de correlation", "1050x850"
    )
    create_title(
        window, "MATRICE DE CORRELATION DE PEARSON"
    )

    fig = Figure(figsize=(8, 6), dpi=110)
    ax = fig.add_subplot(111)

    im = ax.imshow(
        correlation, cmap="coolwarm", vmin=-1, vmax=1
    )

    ax.set_xticks(range(len(variables)))
    ax.set_yticks(range(len(variables)))
    ax.set_xticklabels(variables, fontsize=12)
    ax.set_yticklabels(variables, fontsize=12)

    for i in range(len(variables)):
        for j in range(len(variables)):
            ax.text(
                j, i, f"{correlation[i, j]:.2f}",
                ha="center", va="center",
                fontsize=11, fontweight="bold",
                color="black"
            )

    fig.colorbar(
        im, ax=ax, label="Coefficient de correlation"
    )
    ax.set_title(
        "Correlation Heatmap",
        fontsize=17,
        fontweight="bold",
        pad=15
    )
    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=window)
    canvas.draw()
    canvas.get_tk_widget().pack(
        fill="both", expand=True, padx=20, pady=10
    )

    toolbar_frame = tk.Frame(window)
    toolbar_frame.pack(fill="x")
    toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
    toolbar.update()

# ================================================================
# 11. EXPORTATION DES RESULTATS VERS EXCEL
# ================================================================

def export_excel():
    filename = filedialog.asksaveasfilename(
        parent=root,
        defaultextension=".xlsx",
        filetypes=[("Fichier Excel", "*.xlsx")],
        initialfile="NumPy_Analysis_Results.xlsx"
    )

    if not filename:
        return

    try:
        output = Workbook()

        # Feuille 1 : Dataset
        ws1 = output.active
        ws1.title = "Dataset"
        ws1.append([str(v) for v in variables])

        for row in X:
            ws1.append([float(v) for v in row])

        # Feuille 2 : Statistiques
        ws2 = output.create_sheet("Statistics")
        ws2.append([
            "Variable", "Mean", "Std", "Median",
            "Minimum", "Maximum"
        ])

        for i, name in enumerate(variables):
            ws2.append([
                str(name),
                float(mean_values[i]),
                float(std_values[i]),
                float(median_values[i]),
                float(min_values[i]),
                float(max_values[i])
            ])

        # Feuille 3 : Correlation
        ws3 = output.create_sheet("Correlation")
        ws3.append(["Variable"] + [str(v) for v in variables])

        for i, name in enumerate(variables):
            ws3.append(
                [str(name)] +
                [float(v) for v in correlation[i]]
            )

        output.save(filename)

        messagebox.showinfo(
            "Exportation reussie",
            "Les resultats ont ete exportes avec succes !\n\n"
            + filename,
            parent=root
        )

    except Exception as error:
        messagebox.showerror(
            "Erreur d'exportation", str(error), parent=root
        )

# ================================================================
# 12. BANDEAU PRINCIPAL
# ================================================================

loading_frame.destroy()

header = tk.Frame(root, bg=BLUE, height=115)
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
    text="Analyse des donnees environnementales et agronomiques",
    font=("Arial", 13),
    bg=BLUE,
    fg="#DCEBFF"
).pack()

# ================================================================
# 13. PANNEAU DU FICHIER EXCEL ET ETAT DU CHARGEMENT
# ================================================================

file_panel = tk.Frame(
    root, bg=WHITE, padx=20, pady=12
)
file_panel.pack(fill="x", padx=25, pady=12)

tk.Label(
    file_panel,
    text="FICHIER EXCEL UTILISE",
    font=("Arial", 12, "bold"),
    fg=BLUE,
    bg=WHITE
).pack(anchor="w")

tk.Label(
    file_panel,
    text=os.path.basename(FILE_PATH),
    font=("Arial", 14, "bold"),
    fg=GREEN,
    bg=WHITE
).pack(anchor="w", pady=3)

tk.Label(
    file_panel,
    text="Emplacement : " + os.path.abspath(FILE_PATH),
    font=("Arial", 10),
    fg="#52657A",
    bg=WHITE,
    wraplength=1150,
    justify="left"
).pack(anchor="w")

progress_frame = tk.Frame(file_panel, bg=WHITE)
progress_frame.pack(fill="x", pady=(12, 3))

progress_main = ttk.Progressbar(
    progress_frame,
    style="Custom.Horizontal.TProgressbar",
    maximum=100,
    value=100
)
progress_main.pack(
    side="left", fill="x", expand=True, padx=(0, 15)
)

tk.Label(
    progress_frame,
    text="100 %",
    font=("Arial", 12, "bold"),
    fg=GREEN,
    bg=WHITE
).pack(side="right")

tk.Label(
    file_panel,
    text="✓ DONNEES CHARGEES AVEC SUCCES",
    font=("Arial", 11, "bold"),
    fg=GREEN,
    bg=WHITE
).pack(anchor="w", pady=4)

# ================================================================
# 14. CARTES D'INFORMATIONS
# ================================================================

info = tk.Frame(root, bg=LIGHT, pady=12)
info.pack(fill="x")

cards = [
    ("OBSERVATIONS", str(X.shape[0]), GREEN),
    ("VARIABLES", str(X.shape[1]), "#2874C7"),
    ("DONNEES VALIDES", str(X.size), PURPLE)
]

for title, value, color in cards:
    card = tk.Frame(
        info, bg=WHITE, padx=35, pady=12
    )
    card.pack(side="left", expand=True, padx=15)

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
# 15. MENU PRINCIPAL : BOUTONS AGRANDIS ET SYMBOLES
# ================================================================

tk.Label(
    root,
    text="CHOISIR UNE ANALYSE",
    font=("Arial", 18, "bold"),
    bg=LIGHT,
    fg=BLUE
).pack(pady=12)


def create_button(parent, text, command, color):
    return tk.Button(
        parent,
        text=text,
        command=command,
        font=("Arial", 14, "bold"),
        bg=color,
        fg=WHITE,
        activebackground=color,
        activeforeground=WHITE,
        relief="flat",
        cursor="hand2",
        height=3,
        width=36,
        bd=0
    )


menu = tk.Frame(root, bg=LIGHT)
menu.pack(expand=True, pady=8)

buttons = [
    ("▣  01 - Informations du dataset", show_dataset, BLUE),
    ("▤  02 - Explorer les observations", show_observations, GREEN),
    ("▥  03 - Statistiques descriptives", show_statistics, "#2874C7"),
    ("▦  04 - Histogrammes interactifs", show_histogram, ORANGE),
    ("▧  05 - Heatmap de correlation", show_heatmap, PURPLE)
]

for i, (text, command, color) in enumerate(buttons):
    row = i // 2
    col = i % 2

    button = create_button(menu, text, command, color)
    button.grid(
        row=row, column=col,
        padx=18, pady=12, sticky="nsew"
    )

menu.grid_columnconfigure(0, weight=1)
menu.grid_columnconfigure(1, weight=1)

# ================================================================
# 16. BOUTON D'EXPORTATION
# ================================================================

export_button = tk.Button(
    root,
    text="EXPORTER LES RESULTATS VERS EXCEL",
    command=export_excel,
    font=("Arial", 15, "bold"),
    bg=GREEN,
    fg=WHITE,
    activebackground="#066644",
    activeforeground=WHITE,
    relief="flat",
    cursor="hand2",
    padx=25,
    pady=13
)
export_button.pack(pady=15)

# ================================================================
# 17. PIED DE PAGE
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
# 18. LANCEMENT
# ================================================================

root.mainloop()
