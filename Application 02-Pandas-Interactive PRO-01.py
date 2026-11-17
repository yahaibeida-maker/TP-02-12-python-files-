# -*- coding: utf-8 -*-
"""
APPLICATION 03 : MATPLOTLIB PRO
Visualisation scientifique avec Python, Tkinter et Matplotlib.
Compatible avec Spyder / Anaconda : ne pas forcer de backend Matplotlib.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


# ============================================================
# CONFIGURATION
# ============================================================
APP_TITLE = "APPLICATION 03 : MATPLOTLIB PRO"
DEFAULT_FILE = "OPEN SOURCE DATA FILE.xlsx"
SHEET_NAME = "DATA"

BG = "#F2F5F9"
NAVY = "#17365D"
BLUE = "#2471A3"
GREEN = "#238B57"
ORANGE = "#D97706"
PURPLE = "#7551A1"
WHITE = "#FFFFFF"


# ============================================================
# APPLICATION
# ============================================================
class MatplotlibProApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1450x900")
        self.root.minsize(1120, 720)
        self.root.configure(bg=BG)

        self.df = None
        self.numeric_cols = []
        self.canvas = None
        self.toolbar = None
        self.current_figure = None

        self._build_interface()
        self._load_startup_file()

    # --------------------------------------------------------
    # INTERFACE
    # --------------------------------------------------------
    def _build_interface(self):
        header = tk.Frame(self.root, bg=NAVY, height=92)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header, text="APPLICATION 03 : MATPLOTLIB PRO",
            bg=NAVY, fg="white", font=("Times New Roman", 25, "bold")
        ).pack(anchor="w", padx=24, pady=(13, 0))

        tk.Label(
            header,
            text="VISUALISATION SCIENTIFIQUE DES DONNÉES AVEC PYTHON",
            bg=NAVY, fg="#DCE6F1", font=("Times New Roman", 14, "bold")
        ).pack(anchor="w", padx=27, pady=(2, 8))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=12, pady=12)

        self.sidebar = tk.Frame(body, bg=WHITE, width=330, highlightbackground="#D5DDE7",
                                highlightthickness=1)
        self.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.sidebar.pack_propagate(False)

        self.plot_area = tk.Frame(body, bg=WHITE, highlightbackground="#D5DDE7",
                                  highlightthickness=1)
        self.plot_area.pack(side="right", fill="both", expand=True)

        self._build_sidebar()
        self._build_plot_area()

        footer = tk.Frame(self.root, bg=NAVY, height=28)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(
            footer, text="MASTER 01  |  LOGICIELS LIBRES ET OPEN SOURCE 2026–2027",
            bg=NAVY, fg="white", font=("Arial", 9, "bold")
        ).pack(pady=5)

    def _section(self, parent, title):
        tk.Label(parent, text=title, bg=WHITE, fg=NAVY,
                 font=("Arial", 11, "bold")).pack(anchor="w", padx=14, pady=(12, 5))

    def _build_sidebar(self):
        self._section(self.sidebar, "1. SOURCE DES DONNÉES")
        file_row = tk.Frame(self.sidebar, bg=WHITE)
        file_row.pack(fill="x", padx=12)
        self.file_label = tk.Label(file_row, text="Aucun fichier chargé", bg="#EEF3F8",
                                   fg="#34495E", anchor="w", justify="left",
                                   wraplength=220, font=("Arial", 8))
        self.file_label.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 5))
        ttk.Button(file_row, text="Parcourir", command=self.choose_file).pack(side="right")

        self.dataset_info = tk.Label(
            self.sidebar, text="En attente du fichier Excel…", bg=WHITE, fg="#526273",
            justify="left", anchor="w", font=("Arial", 9), wraplength=290
        )
        self.dataset_info.pack(fill="x", padx=14, pady=(7, 0))

        ttk.Separator(self.sidebar).pack(fill="x", padx=12, pady=8)
        self._section(self.sidebar, "2. PARAMÈTRES DU GRAPHIQUE")

        tk.Label(self.sidebar, text="Type de visualisation", bg=WHITE,
                 font=("Arial", 9, "bold")).pack(anchor="w", padx=14)
        self.chart_type = tk.StringVar(value="Courbe")
        self.chart_combo = ttk.Combobox(
            self.sidebar, textvariable=self.chart_type, state="readonly",
            values=[
                "Courbe", "Nuage de points", "Histogramme", "Boxplot",
                "Diagramme en barres", "Heatmap de corrélation", "Tableau 2 × 2"
            ]
        )
        self.chart_combo.pack(fill="x", padx=14, pady=(3, 8))
        self.chart_combo.bind("<<ComboboxSelected>>", self._chart_type_changed)

        tk.Label(self.sidebar, text="Variable X / variable à représenter",
                 bg=WHITE, font=("Arial", 9, "bold")).pack(anchor="w", padx=14)
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(self.sidebar, textvariable=self.x_var, state="readonly")
        self.x_combo.pack(fill="x", padx=14, pady=(3, 8))

        tk.Label(self.sidebar, text="Variable Y (pour courbe / scatter)",
                 bg=WHITE, font=("Arial", 9, "bold")).pack(anchor="w", padx=14)
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(self.sidebar, textvariable=self.y_var, state="readonly")
        self.y_combo.pack(fill="x", padx=14, pady=(3, 8))

        tk.Label(self.sidebar, text="Titre de la figure", bg=WHITE,
                 font=("Arial", 9, "bold")).pack(anchor="w", padx=14)
        self.title_entry = ttk.Entry(self.sidebar)
        self.title_entry.pack(fill="x", padx=14, pady=(3, 8))
        self.title_entry.insert(0, "Visualisation des données")

        style_frame = tk.Frame(self.sidebar, bg=WHITE)
        style_frame.pack(fill="x", padx=14, pady=2)
        tk.Label(style_frame, text="Taille des points", bg=WHITE,
                 font=("Arial", 9)).grid(row=0, column=0, sticky="w")
        self.point_size = tk.IntVar(value=45)
        ttk.Spinbox(style_frame, from_=10, to=250, increment=5,
                    textvariable=self.point_size, width=8).grid(row=0, column=1, padx=5)
        tk.Label(style_frame, text="DPI export", bg=WHITE,
                 font=("Arial", 9)).grid(row=1, column=0, sticky="w", pady=5)
        self.dpi_value = tk.IntVar(value=300)
        ttk.Spinbox(style_frame, from_=72, to=1200, increment=50,
                    textvariable=self.dpi_value, width=8).grid(row=1, column=1, padx=5)

        ttk.Button(self.sidebar, text="GÉNÉRER LE GRAPHIQUE",
                   command=self.generate_plot).pack(fill="x", padx=14, pady=(12, 6), ipady=7)

        ttk.Button(self.sidebar, text="EXPORTER LA FIGURE",
                   command=self.export_figure).pack(fill="x", padx=14, pady=4, ipady=5)

        ttk.Button(self.sidebar, text="RÉINITIALISER",
                   command=self.reset_plot).pack(fill="x", padx=14, pady=4, ipady=5)

        ttk.Separator(self.sidebar).pack(fill="x", padx=12, pady=8)
        self._section(self.sidebar, "3. AIDE PÉDAGOGIQUE")
        self.help_text = tk.Label(
            self.sidebar,
            text=("• Courbe : évolution d'une variable.\n"
                  "• Scatter : relation entre deux variables.\n"
                  "• Histogramme : distribution des valeurs.\n"
                  "• Boxplot : médiane, dispersion et valeurs atypiques.\n"
                  "• Heatmap : corrélations entre variables.\n"
                  "• Tableau 2 × 2 : quatre graphiques synthétiques."),
            bg=WHITE, fg="#34495E", justify="left", anchor="w",
            font=("Arial", 9), wraplength=290
        )
        self.help_text.pack(fill="x", padx=14, pady=(0, 12))

    def _build_plot_area(self):
        top = tk.Frame(self.plot_area, bg=WHITE)
        top.pack(fill="x", padx=14, pady=10)
        tk.Label(top, text="ESPACE DE VISUALISATION",
                 bg=WHITE, fg=NAVY, font=("Arial", 14, "bold")).pack(side="left")
        self.status_label = tk.Label(top, text="Prêt", bg=WHITE, fg=GREEN,
                                     font=("Arial", 9, "bold"))
        self.status_label.pack(side="right")

        self.canvas_frame = tk.Frame(self.plot_area, bg=WHITE)
        self.canvas_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.welcome = tk.Label(
            self.canvas_frame,
            text=("MATPLOTLIB PRO\n\n"
                  "Chargez votre fichier Excel, choisissez un type de graphique\n"
                  "et cliquez sur « GÉNÉRER LE GRAPHIQUE »."),
            bg=WHITE, fg="#607080", justify="center",
            font=("Times New Roman", 19, "bold")
        )
        self.welcome.pack(expand=True)

    # --------------------------------------------------------
    # DATA LOADING
    # --------------------------------------------------------
    def _load_startup_file(self):
        candidates = [
            os.path.join(os.getcwd(), DEFAULT_FILE),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_FILE)
        ]
        for path in candidates:
            if os.path.isfile(path):
                self.load_file(path)
                return

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
        )
        if path:
            self.load_file(path)

    def load_file(self, path):
        try:
            try:
                data = pd.read_excel(path, sheet_name=SHEET_NAME)
            except ValueError:
                # Si la feuille DATA n'existe pas, proposer la première feuille.
                data = pd.read_excel(path, sheet_name=0)
                messagebox.showwarning(
                    "Feuille Excel",
                    f"La feuille '{SHEET_NAME}' n'a pas été trouvée.\n"
                    "La première feuille du classeur a été chargée."
                )

            data = data.dropna(how="all").dropna(axis=1, how="all")
            data.columns = [str(c).strip() for c in data.columns]
            if data.empty:
                raise ValueError("Le fichier ne contient aucune observation exploitable.")

            self.df = data
            self.numeric_cols = data.select_dtypes(include=np.number).columns.tolist()
            if not self.numeric_cols:
                raise ValueError("Aucune colonne numérique n'a été détectée.")

            self.file_label.config(text=os.path.basename(path))
            self.dataset_info.config(
                text=(f"Fichier : {os.path.basename(path)}\n"
                      f"Observations : {len(data):,}\n"
                      f"Variables : {len(data.columns)}\n"
                      f"Variables numériques : {len(self.numeric_cols)}")
            )
            self.x_combo["values"] = self.numeric_cols
            self.y_combo["values"] = self.numeric_cols
            self.x_var.set(self.numeric_cols[0])
            self.y_var.set(self.numeric_cols[1] if len(self.numeric_cols) > 1 else self.numeric_cols[0])
            self.status_label.config(text="Données chargées", fg=GREEN)
            self.generate_plot()
        except Exception as exc:
            messagebox.showerror("Erreur de chargement", str(exc))
            self.status_label.config(text="Erreur de chargement", fg="#C0392B")

    # --------------------------------------------------------
    # PLOT GENERATION
    # --------------------------------------------------------
    def _chart_type_changed(self, _event=None):
        chart = self.chart_type.get()
        if chart in ("Heatmap de corrélation", "Tableau 2 × 2"):
            self.x_combo.configure(state="disabled")
            self.y_combo.configure(state="disabled")
        else:
            self.x_combo.configure(state="readonly")
            self.y_combo.configure(state="readonly")

    def _clear_canvas(self):
        if self.toolbar is not None:
            self.toolbar.destroy()
            self.toolbar = None
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

    def _new_figure(self, rows=1, cols=1, figsize=(10, 6)):
        self._clear_canvas()
        fig = Figure(figsize=figsize, dpi=100, constrained_layout=True)
        axes = fig.subplots(rows, cols)
        self.current_figure = fig
        return fig, axes

    def generate_plot(self):
        if self.df is None:
            messagebox.showwarning("Données manquantes", "Veuillez charger un fichier Excel.")
            return

        chart = self.chart_type.get()
        title = self.title_entry.get().strip() or "Visualisation des données"
        x_name = self.x_var.get()
        y_name = self.y_var.get()

        try:
            if chart == "Courbe":
                if not x_name or not y_name:
                    raise ValueError("Sélectionnez les variables X et Y.")
                fig, ax = self._new_figure()
                ax.plot(np.arange(len(self.df)), self.df[y_name], linewidth=1.8,
                        marker="o", markersize=3.5, label=y_name)
                ax.set_xlabel("Indice de l'observation")
                ax.set_ylabel(y_name)
                ax.set_title(title, fontsize=15, fontweight="bold")
                ax.grid(True, alpha=0.3)
                ax.legend()

            elif chart == "Nuage de points":
                if not x_name or not y_name:
                    raise ValueError("Sélectionnez les variables X et Y.")
                fig, ax = self._new_figure()
                ax.scatter(self.df[x_name], self.df[y_name],
                           s=self.point_size.get(), alpha=0.72,
                           edgecolors="black", linewidths=0.35)
                ax.set_xlabel(x_name)
                ax.set_ylabel(y_name)
                ax.set_title(title, fontsize=15, fontweight="bold")
                ax.grid(True, alpha=0.3)

            elif chart == "Histogramme":
                if not x_name:
                    raise ValueError("Sélectionnez une variable.")
                values = self.df[x_name].dropna()
                fig, ax = self._new_figure()
                ax.hist(values, bins="auto", edgecolor="black", alpha=0.8)
                ax.set_xlabel(x_name)
                ax.set_ylabel("Fréquence")
                ax.set_title(title, fontsize=15, fontweight="bold")
                ax.grid(axis="y", alpha=0.25)

            elif chart == "Boxplot":
                if not x_name:
                    raise ValueError("Sélectionnez une variable.")
                fig, ax = self._new_figure()
                values = self.df[x_name].dropna()
                ax.boxplot(values, patch_artist=True, tick_labels=[x_name],
                           showmeans=True)
                ax.set_ylabel(x_name)
                ax.set_title(title, fontsize=15, fontweight="bold")
                ax.grid(axis="y", alpha=0.25)

            elif chart == "Diagramme en barres":
                if not x_name:
                    raise ValueError("Sélectionnez une variable.")
                values = self.df[x_name].dropna()
                counts = values.value_counts().sort_index()
                if len(counts) > 40:
                    counts = values.round(1).value_counts().sort_index()
                fig, ax = self._new_figure(figsize=(11, 6))
                ax.bar([str(v) for v in counts.index], counts.values,
                       edgecolor="black", alpha=0.82)
                ax.set_xlabel(x_name)
                ax.set_ylabel("Effectif")
                ax.set_title(title, fontsize=15, fontweight="bold")
                ax.tick_params(axis="x", labelrotation=45)
                ax.grid(axis="y", alpha=0.25)

            elif chart == "Heatmap de corrélation":
                corr = self.df[self.numeric_cols].corr()
                fig, ax = self._new_figure(figsize=(10, 8))
                image = ax.imshow(corr, cmap="coolwarm", vmin=-1, vmax=1, aspect="auto")
                ax.set_xticks(range(len(corr.columns)))
                ax.set_yticks(range(len(corr.columns)))
                ax.set_xticklabels(corr.columns, rotation=45, ha="right")
                ax.set_yticklabels(corr.columns)
                ax.set_title(title, fontsize=15, fontweight="bold")
                for i in range(len(corr.columns)):
                    for j in range(len(corr.columns)):
                        ax.text(j, i, f"{corr.iloc[i, j]:.2f}",
                                ha="center", va="center", fontsize=9)
                fig.colorbar(image, ax=ax, fraction=0.046, pad=0.04, label="Corrélation")

            elif chart == "Tableau 2 × 2":
                fig, axes = self._new_figure(rows=2, cols=2, figsize=(12, 8))
                axes = np.asarray(axes).ravel()
                for ax, col in zip(axes, self.numeric_cols[:4]):
                    values = self.df[col].dropna()
                    ax.hist(values, bins="auto", alpha=0.8, edgecolor="black")
                    ax.set_title(f"Distribution de {col}", fontweight="bold")
                    ax.set_xlabel(col)
                    ax.set_ylabel("Fréquence")
                    ax.grid(axis="y", alpha=0.2)
                for ax in axes[len(self.numeric_cols[:4]):]:
                    ax.set_visible(False)
                fig.suptitle(title, fontsize=16, fontweight="bold")

            self._display_figure(fig)
            self.status_label.config(text=f"Graphique généré : {chart}", fg=GREEN)

        except Exception as exc:
            messagebox.showerror("Erreur de visualisation", str(exc))
            self.status_label.config(text="Erreur de génération", fg="#C0392B")

    def _display_figure(self, fig):
        self._clear_canvas()
        self.current_figure = fig
        self.canvas = FigureCanvasTkAgg(fig, master=self.canvas_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar_frame = tk.Frame(self.canvas_frame, bg=WHITE)
        toolbar_frame.pack(fill="x")
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()

    # --------------------------------------------------------
    # EXPORT / RESET
    # --------------------------------------------------------
    def export_figure(self):
        if self.current_figure is None:
            messagebox.showwarning("Exportation", "Générez d'abord une figure.")
            return
        path = filedialog.asksaveasfilename(
            title="Exporter la figure",
            defaultextension=".png",
            filetypes=[
                ("Image PNG", "*.png"),
                ("Image JPEG", "*.jpg"),
                ("Document PDF", "*.pdf"),
                ("Image SVG", "*.svg")
            ]
        )
        if not path:
            return
        try:
            dpi = max(72, int(self.dpi_value.get()))
            self.current_figure.savefig(path, dpi=dpi, bbox_inches="tight",
                                        facecolor="white")
            messagebox.showinfo("Exportation réussie",
                                f"La figure a été exportée :\n{path}")
        except Exception as exc:
            messagebox.showerror("Erreur d'exportation", str(exc))

    def reset_plot(self):
        self.chart_type.set("Courbe")
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, "Visualisation des données")
        if self.numeric_cols:
            self.x_var.set(self.numeric_cols[0])
            self.y_var.set(self.numeric_cols[1] if len(self.numeric_cols) > 1 else self.numeric_cols[0])
        self._chart_type_changed()
        if self.df is not None:
            self.generate_plot()


# ============================================================
# LANCEMENT
# ============================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = MatplotlibProApp(root)
    root.mainloop()

