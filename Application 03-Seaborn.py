# -*- coding: utf-8 -*-
"""
APPLICATION 04 : SEABORN PRO V01
Première application pédagogique de visualisation statistique avec Seaborn.
Interface Tkinter - Anaconda / Spyder.
Données : OPEN SOURCE DATA FILE.xlsx, feuille DATA (ou première feuille).
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


APP_TITLE = "APPLICATION 04 : SEABORN PRO V01"
DEFAULT_FILE = "OPEN SOURCE DATA FILE.xlsx"
SHEET_NAME = "DATA"

BG = "#F2F5F9"
NAVY = "#17365D"
WHITE = "#FFFFFF"


class SeabornProApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1550x930")
        self.root.minsize(1180, 740)
        self.root.configure(bg=BG)

        self.df = None
        self.numeric_cols = []
        self.categorical_cols = []
        self.current_figure = None
        self.canvas = None
        self.toolbar = None

        sns.set_theme(style="whitegrid", context="notebook", font="DejaVu Sans")
        self._build_interface()
        self._load_startup_file()

    # ---------------------------------------------------------
    # INTERFACE
    # ---------------------------------------------------------
    def _build_interface(self):
        header = tk.Frame(self.root, bg=NAVY, height=88)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text=APP_TITLE, bg=NAVY, fg="white",
                 font=("Times New Roman", 25, "bold")).pack(anchor="w", padx=22, pady=(10, 0))
        tk.Label(header, text="VISUALISATION STATISTIQUE DES DONNÉES AVEC SEABORN",
                 bg=NAVY, fg="#DCE6F1",
                 font=("Times New Roman", 14, "bold")).pack(anchor="w", padx=26, pady=(1, 8))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=10, pady=10)

        self.sidebar = tk.Frame(body, bg=WHITE, width=350,
                                highlightbackground="#D5DDE7", highlightthickness=1)
        self.sidebar.pack(side="left", fill="y", padx=(0, 9))
        self.sidebar.pack_propagate(False)

        right = tk.Frame(body, bg=BG)
        right.pack(side="right", fill="both", expand=True)

        self.plot_area = tk.Frame(right, bg=WHITE,
                                  highlightbackground="#D5DDE7", highlightthickness=1)
        self.plot_area.pack(fill="both", expand=True)

        self.info_panel = tk.Frame(right, bg=WHITE,
                                   highlightbackground="#D5DDE7", highlightthickness=1)
        self.info_panel.pack(fill="x", pady=(8, 0))

        self._build_sidebar()
        self._build_plot_area()

        footer = tk.Frame(self.root, bg=NAVY, height=27)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(footer,
                 text="MASTER 01  |  LOGICIELS LIBRES ET OPEN SOURCE 2026–2027",
                 bg=NAVY, fg="white", font=("Arial", 9, "bold")).pack(pady=4)

    def _section(self, parent, text):
        tk.Label(parent, text=text, bg=WHITE, fg=NAVY,
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=12, pady=(9, 4))

    def _small_label(self, parent, text):
        tk.Label(parent, text=text, bg=WHITE,
                 font=("Arial", 8, "bold")).pack(anchor="w", padx=12, pady=(3, 1))

    def _build_sidebar(self):
        self._section(self.sidebar, "1. SOURCE DES DONNÉES")
        row = tk.Frame(self.sidebar, bg=WHITE)
        row.pack(fill="x", padx=11)

        self.file_label = tk.Label(row, text="Aucun fichier chargé", bg="#EEF3F8",
                                   fg="#34495E", anchor="w", justify="left",
                                   wraplength=220, font=("Arial", 8))
        self.file_label.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 5))
        ttk.Button(row, text="Parcourir", command=self.choose_file).pack(side="right")

        self.dataset_info = tk.Label(self.sidebar, text="En attente du fichier Excel…",
                                     bg=WHITE, fg="#526273", justify="left",
                                     anchor="w", font=("Arial", 8), wraplength=320)
        self.dataset_info.pack(fill="x", padx=12, pady=(5, 0))

        ttk.Separator(self.sidebar).pack(fill="x", padx=10, pady=6)
        self._section(self.sidebar, "2. TYPE DE GRAPHIQUE")

        self.chart_var = tk.StringVar(value="Histogramme + KDE")
        self.chart_combo = ttk.Combobox(
            self.sidebar, textvariable=self.chart_var, state="readonly",
            values=[
                "Histogramme + KDE",
                "KDE plot",
                "Boxplot",
                "Violin plot",
                "Strip plot",
                "Swarm plot",
                "Scatter plot",
                "Pairplot",
                "Heatmap de corrélation",
                "Jointplot",
                "Barplot",
                "Countplot"
            ]
        )
        self.chart_combo.pack(fill="x", padx=12, pady=(0, 5))
        self.chart_combo.bind("<<ComboboxSelected>>", self._chart_changed)

        self._section(self.sidebar, "3. SÉLECTION DES VARIABLES")
        self._small_label(self.sidebar, "Variable numérique X")
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(self.sidebar, textvariable=self.x_var, state="readonly")
        self.x_combo.pack(fill="x", padx=12, pady=(0, 5))

        self._small_label(self.sidebar, "Variable numérique Y")
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(self.sidebar, textvariable=self.y_var, state="readonly")
        self.y_combo.pack(fill="x", padx=12, pady=(0, 5))

        self._small_label(self.sidebar, "Variable catégorielle (hue / groupe)")
        self.hue_var = tk.StringVar(value="Aucune")
        self.hue_combo = ttk.Combobox(self.sidebar, textvariable=self.hue_var, state="readonly")
        self.hue_combo.pack(fill="x", padx=12, pady=(0, 5))

        self._section(self.sidebar, "4. PARAMÈTRES")
        row = tk.Frame(self.sidebar, bg=WHITE)
        row.pack(fill="x", padx=12)
        tk.Label(row, text="Palette", bg=WHITE, font=("Arial", 8)).grid(row=0, column=0, sticky="w")
        self.palette_var = tk.StringVar(value="deep")
        ttk.Combobox(row, textvariable=self.palette_var, state="readonly",
                     values=["deep", "muted", "bright", "pastel", "dark", "colorblind",
                             "viridis", "magma", "coolwarm"], width=14).grid(
                                 row=0, column=1, sticky="ew", padx=4, pady=2)
        tk.Label(row, text="Nombre de bins", bg=WHITE, font=("Arial", 8)).grid(
            row=1, column=0, sticky="w")
        self.bins_var = tk.IntVar(value=20)
        ttk.Spinbox(row, from_=5, to=100, increment=5,
                    textvariable=self.bins_var, width=7).grid(
                        row=1, column=1, sticky="w", padx=4, pady=2)
        tk.Label(row, text="Taille figure", bg=WHITE, font=("Arial", 8)).grid(
            row=2, column=0, sticky="w")
        self.figsize_var = tk.StringVar(value="Grande (12 × 7)")
        ttk.Combobox(row, textvariable=self.figsize_var, state="readonly",
                     values=["Moyenne (9 × 5)", "Grande (12 × 7)", "Très grande (15 × 9)"],
                     width=14).grid(row=2, column=1, sticky="ew", padx=4, pady=2)
        tk.Label(row, text="DPI export", bg=WHITE, font=("Arial", 8)).grid(
            row=3, column=0, sticky="w")
        self.dpi_var = tk.IntVar(value=300)
        ttk.Spinbox(row, from_=72, to=1200, increment=50,
                    textvariable=self.dpi_var, width=7).grid(
                        row=3, column=1, sticky="w", padx=4, pady=2)
        row.columnconfigure(1, weight=1)

        self.kde_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.sidebar, text="Afficher la courbe KDE (histogramme)",
                        variable=self.kde_var).pack(anchor="w", padx=12, pady=(4, 0))
        self.reg_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.sidebar, text="Afficher la régression (scatter)",
                        variable=self.reg_var).pack(anchor="w", padx=12)

        self._section(self.sidebar, "5. TITRE ET COMMANDES")
        self.title_entry = ttk.Entry(self.sidebar)
        self.title_entry.pack(fill="x", padx=12, pady=(0, 6))
        self.title_entry.insert(0, "Analyse statistique des variables")

        buttons = tk.Frame(self.sidebar, bg=WHITE)
        buttons.pack(fill="x", padx=10, pady=4)
        ttk.Button(buttons, text="GÉNÉRER", command=self.generate_plot).pack(
            side="left", fill="x", expand=True, padx=2, ipady=5)
        ttk.Button(buttons, text="EXPORTER", command=self.export_figure).pack(
            side="left", fill="x", expand=True, padx=2, ipady=5)
        ttk.Button(buttons, text="RESET", command=self.reset).pack(
            side="left", fill="x", expand=True, padx=2, ipady=5)

        self.help_label = tk.Label(
            self.sidebar,
            text="Aide : choisissez un graphique, puis les variables nécessaires.",
            bg="#EEF3F8", fg="#34495E", justify="left", anchor="w",
            font=("Arial", 8), wraplength=320, padx=8, pady=6)
        self.help_label.pack(fill="x", padx=10, pady=(5, 8))

    def _build_plot_area(self):
        top = tk.Frame(self.plot_area, bg=WHITE)
        top.pack(fill="x", padx=12, pady=8)
        tk.Label(top, text="ESPACE DE VISUALISATION SEABORN",
                 bg=WHITE, fg=NAVY, font=("Arial", 13, "bold")).pack(side="left")
        self.status = tk.Label(top, text="Prêt", bg=WHITE, fg="#238B57",
                               font=("Arial", 8, "bold"))
        self.status.pack(side="right")

        self.canvas_frame = tk.Frame(self.plot_area, bg=WHITE)
        self.canvas_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.welcome = tk.Label(
            self.canvas_frame,
            text="SEABORN PRO V01\n\nChargez votre fichier de données,\nchoisissez une visualisation statistique\net cliquez sur GÉNÉRER.",
            bg=WHITE, fg="#607080", justify="center",
            font=("Times New Roman", 19, "bold"))
        self.welcome.pack(expand=True)

        tk.Label(self.info_panel, text="INFORMATIONS ET STATISTIQUES",
                 bg=WHITE, fg=NAVY, font=("Arial", 10, "bold")).pack(
                     anchor="w", padx=10, pady=(5, 0))
        self.info_label = tk.Label(
            self.info_panel, text="Les informations apparaîtront après la génération.",
            bg=WHITE, fg="#34495E", font=("Arial", 9),
            justify="left", anchor="w", wraplength=1100)
        self.info_label.pack(fill="x", padx=10, pady=5)

    # ---------------------------------------------------------
    # DATA LOADING
    # ---------------------------------------------------------
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
            title="Sélectionner le fichier de données",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"),
                       ("Tous les fichiers", "*.*")])
        if path:
            self.load_file(path)

    def load_file(self, path):
        try:
            if path.lower().endswith(".csv"):
                data = pd.read_csv(path)
            else:
                try:
                    data = pd.read_excel(path, sheet_name=SHEET_NAME)
                except ValueError:
                    data = pd.read_excel(path, sheet_name=0)
                    messagebox.showwarning(
                        "Feuille Excel",
                        f"La feuille '{SHEET_NAME}' est introuvable.\n"
                        "La première feuille du classeur a été chargée.")
            data = data.dropna(how="all").dropna(axis=1, how="all")
            data.columns = [str(c).strip() for c in data.columns]
            if data.empty:
                raise ValueError("Aucune observation exploitable dans le fichier.")

            self.df = data
            self.numeric_cols = data.select_dtypes(include=np.number).columns.tolist()
            self.categorical_cols = data.select_dtypes(
                include=["object", "category", "bool"]).columns.tolist()
            if not self.numeric_cols:
                raise ValueError("Aucune variable numérique détectée.")

            self.file_label.config(text=os.path.basename(path))
            self.dataset_info.config(
                text=f"Fichier : {os.path.basename(path)} | N = {len(data):,} | "
                     f"Variables = {len(data.columns)} | Numériques = {len(self.numeric_cols)} | "
                     f"Catégorielles = {len(self.categorical_cols)}")

            self.x_combo["values"] = self.numeric_cols
            self.y_combo["values"] = self.numeric_cols
            self.x_var.set(self.numeric_cols[0])
            self.y_var.set(self.numeric_cols[1] if len(self.numeric_cols) > 1 else self.numeric_cols[0])

            hue_values = ["Aucune"] + self.categorical_cols
            self.hue_combo["values"] = hue_values
            self.hue_var.set("Aucune")
            self.status.config(text="Données chargées", fg="#238B57")
            self.generate_plot()
        except Exception as exc:
            messagebox.showerror("Erreur de chargement", str(exc))
            self.status.config(text="Erreur de chargement", fg="#C0392B")

    # ---------------------------------------------------------
    # PLOT CONTROLS
    # ---------------------------------------------------------
    def _chart_changed(self, _event=None):
        chart = self.chart_var.get()
        self.x_combo.configure(state="readonly")
        self.y_combo.configure(state="readonly")
        if chart in ("Heatmap de corrélation", "Pairplot"):
            self.x_combo.configure(state="disabled")
            self.y_combo.configure(state="disabled")
        if chart in ("Histogramme + KDE", "KDE plot", "Boxplot", "Violin plot",
                     "Strip plot", "Swarm plot", "Barplot", "Countplot"):
            self.y_combo.configure(state="disabled")
        descriptions = {
            "Histogramme + KDE": "Distribution d'une variable : histogramme et estimation de densité KDE.",
            "KDE plot": "Courbe de densité estimée : visualise la forme de la distribution.",
            "Boxplot": "Résumé de distribution : médiane, quartiles, dispersion et valeurs atypiques.",
            "Violin plot": "Combine un boxplot et une estimation de densité pour comparer les distributions.",
            "Strip plot": "Représente les observations individuelles par groupe.",
            "Swarm plot": "Répartit les points sans chevauchement lorsque cela est possible.",
            "Scatter plot": "Explore la relation entre deux variables numériques.",
            "Pairplot": "Matrice de relations bivariées entre plusieurs variables numériques.",
            "Heatmap de corrélation": "Carte thermique des coefficients de corrélation.",
            "Jointplot": "Relation entre deux variables et distributions marginales.",
            "Barplot": "Compare une statistique agrégée (moyenne par défaut) entre groupes.",
            "Countplot": "Compte le nombre d'observations dans chaque catégorie."
        }
        self.help_label.config(text=descriptions.get(chart, "Choisissez les variables nécessaires."))

    def _figsize(self):
        return {
            "Moyenne (9 × 5)": (9, 5),
            "Grande (12 × 7)": (12, 7),
            "Très grande (15 × 9)": (15, 9)
        }.get(self.figsize_var.get(), (12, 7))

    def _hue(self):
        value = self.hue_var.get()
        return value if value and value != "Aucune" else None

    def _make_figure(self):
        self._clear_canvas()
        fig = Figure(figsize=self._figsize(), dpi=100, constrained_layout=True)
        ax = fig.add_subplot(111)
        return fig, ax

    def _clear_canvas(self):
        if self.toolbar is not None:
            self.toolbar.destroy()
            self.toolbar = None
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

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

    # ---------------------------------------------------------
    # GENERATE SEABORN VISUALIZATION
    # ---------------------------------------------------------
    def generate_plot(self):
        if self.df is None:
            messagebox.showwarning("Données manquantes", "Veuillez charger un fichier Excel ou CSV.")
            return

        chart = self.chart_var.get()
        x_name, y_name = self.x_var.get(), self.y_var.get()
        hue = self._hue()
        title = self.title_entry.get().strip() or "Analyse statistique"
        palette = self.palette_var.get()
        bins = max(5, int(self.bins_var.get()))

        try:
            fig, ax = self._make_figure()
            stats_message = f"Type : {chart} | Nombre d'observations du jeu de données : {len(self.df)}"

            if chart == "Histogramme + KDE":
                sns.histplot(data=self.df, x=x_name, hue=hue, bins=bins,
                             kde=self.kde_var.get(), palette=palette if hue else None,
                             ax=ax, element="bars", stat="count")
                s = pd.to_numeric(self.df[x_name], errors="coerce").dropna()
                stats_message += f" | {x_name}: moyenne={s.mean():.4g}, médiane={s.median():.4g}, écart-type={s.std():.4g}"

            elif chart == "KDE plot":
                sns.kdeplot(data=self.df, x=x_name, hue=hue,
                            fill=True if hue is None else False,
                            common_norm=False, palette=palette if hue else None,
                            ax=ax, warn_singular=False)
                s = pd.to_numeric(self.df[x_name], errors="coerce").dropna()
                stats_message += f" | {x_name}: N={len(s)}, moyenne={s.mean():.4g}, médiane={s.median():.4g}"

            elif chart in ("Boxplot", "Violin plot", "Strip plot", "Swarm plot"):
                if hue is None:
                    # Une variable de groupe n'est pas obligatoire : représentation univariée.
                    if chart == "Boxplot":
                        sns.boxplot(data=self.df, y=x_name, color=sns.color_palette(palette)[0], ax=ax)
                    elif chart == "Violin plot":
                        sns.violinplot(data=self.df, y=x_name, color=sns.color_palette(palette)[0], ax=ax)
                    elif chart == "Strip plot":
                        sns.stripplot(data=self.df, y=x_name, color=sns.color_palette(palette)[0],
                                      jitter=True, alpha=0.65, ax=ax)
                    else:
                        sns.swarmplot(data=self.df, y=x_name, color=sns.color_palette(palette)[0],
                                      size=3, ax=ax)
                else:
                    if hue not in self.df.columns:
                        raise ValueError("La variable de groupe sélectionnée n'existe pas.")
                    if chart == "Boxplot":
                        sns.boxplot(data=self.df, x=hue, y=x_name, palette=palette, ax=ax, hue=hue, legend=False)
                    elif chart == "Violin plot":
                        sns.violinplot(data=self.df, x=hue, y=x_name, palette=palette, ax=ax, hue=hue, legend=False)
                    elif chart == "Strip plot":
                        sns.stripplot(data=self.df, x=hue, y=x_name, hue=hue, palette=palette,
                                      dodge=False, alpha=0.65, ax=ax, legend=False)
                    else:
                        sns.swarmplot(data=self.df, x=hue, y=x_name, hue=hue, palette=palette,
                                      dodge=False, size=3, ax=ax, legend=False)
                s = pd.to_numeric(self.df[x_name], errors="coerce").dropna()
                stats_message += f" | Variable {x_name}: médiane={s.median():.4g}, Q1={s.quantile(.25):.4g}, Q3={s.quantile(.75):.4g}"

            elif chart == "Scatter plot":
                sns.scatterplot(data=self.df, x=x_name, y=y_name, hue=hue,
                                palette=palette if hue else None, s=65, alpha=0.8, ax=ax)
                if self.reg_var.get():
                    sns.regplot(data=self.df, x=x_name, y=y_name, scatter=False,
                                color="#C0392B", line_kws={"linewidth": 2}, ax=ax)
                pair = self.df[[x_name, y_name]].apply(pd.to_numeric, errors="coerce").dropna()
                r = pair[x_name].corr(pair[y_name])
                stats_message += f" | N={len(pair)} | Corrélation de Pearson r={r:.4f}"

            elif chart == "Pairplot":
                cols = self.numeric_cols[:6]
                if len(cols) < 2:
                    raise ValueError("Pairplot exige au moins deux variables numériques.")
                data = self.df[cols + ([hue] if hue else [])].dropna()
                if len(data) > 1000:
                    data = data.sample(1000, random_state=42)
                grid = sns.pairplot(data=data, vars=cols, hue=hue,
                                    palette=palette if hue else None,
                                    diag_kind="hist", corner=True,
                                    plot_kws={"alpha": 0.65, "s": 25})
                grid.fig.suptitle(title, y=1.02, fontsize=14, fontweight="bold")
                self._display_figure(grid.fig)
                self.info_label.config(text=stats_message +
                                       f" | Variables affichées : {', '.join(cols)} | "
                                       f"Échantillon affiché : {len(data)}")
                self.status.config(text=f"Graphique généré : {chart}", fg="#238B57")
                return

            elif chart == "Heatmap de corrélation":
                corr = self.df[self.numeric_cols].corr()
                sns.heatmap(corr, annot=True, fmt=".2f", cmap=palette,
                            vmin=-1, vmax=1, center=0, square=True,
                            linewidths=0.5, cbar_kws={"label": "Corrélation"},
                            annot_kws={"size": 9}, ax=ax)
                stats_message += f" | Variables analysées : {len(corr.columns)}"

            elif chart == "Jointplot":
                # Jointplot construit sa propre figure Seaborn.
                plt_data = self.df[[x_name, y_name]].apply(pd.to_numeric, errors="coerce").dropna()
                if len(plt_data) < 3:
                    raise ValueError("Jointplot exige au moins trois paires de valeurs valides.")
                if len(plt_data) > 3000:
                    plt_data = plt_data.sample(3000, random_state=42)
                # Fermer la figure de base créée pour les autres types.
                import matplotlib.pyplot as plt
                plt.close(fig)
                grid = sns.jointplot(data=plt_data, x=x_name, y=y_name,
                                     kind="scatter", height=self._figsize()[1],
                                     marginal_kws={"bins": bins, "fill": True},
                                     joint_kws={"alpha": 0.65, "s": 30})
                grid.fig.suptitle(title, y=1.02, fontsize=14, fontweight="bold")
                r = plt_data[x_name].corr(plt_data[y_name])
                self._display_figure(grid.fig)
                self.info_label.config(text=stats_message + f" | Paires valides : {len(plt_data)} | r = {r:.4f}")
                self.status.config(text=f"Graphique généré : {chart}", fg="#238B57")
                return

            elif chart in ("Barplot", "Countplot"):
                if hue is None:
                    raise ValueError("Ce graphique nécessite une variable catégorielle dans le champ hue / groupe.")
                if chart == "Barplot":
                    sns.barplot(data=self.df, x=hue, y=x_name, hue=hue,
                                palette=palette, errorbar=("ci", 95),
                                ax=ax, legend=False)
                    stats_message += f" | Barplot : moyenne de {x_name} selon {hue} (IC 95 %)"
                else:
                    sns.countplot(data=self.df, x=hue, hue=hue,
                                  palette=palette, ax=ax, legend=False)
                    stats_message += f" | Countplot : effectifs par catégorie de {hue}"

            ax.set_title(title, fontsize=14, fontweight="bold")
            ax.tick_params(axis="x", labelrotation=35)
            fig.tight_layout()
            self._display_figure(fig)
            self.info_label.config(text=stats_message)
            self.status.config(text=f"Graphique généré : {chart}", fg="#238B57")

        except Exception as exc:
            messagebox.showerror("Erreur de visualisation", str(exc))
            self.status.config(text="Erreur de génération", fg="#C0392B")

    # ---------------------------------------------------------
    # EXPORT / RESET
    # ---------------------------------------------------------
    def export_figure(self):
        if self.current_figure is None:
            messagebox.showwarning("Exportation", "Générez d'abord un graphique.")
            return
        path = filedialog.asksaveasfilename(
            title="Exporter le graphique", defaultextension=".png",
            filetypes=[("Image PNG", "*.png"), ("Image JPEG", "*.jpg"),
                       ("Document PDF", "*.pdf"), ("Image SVG", "*.svg")])
        if not path:
            return
        try:
            dpi = max(72, int(self.dpi_var.get()))
            self.current_figure.savefig(path, dpi=dpi, bbox_inches="tight",
                                        facecolor="white")
            messagebox.showinfo("Exportation réussie", f"Graphique enregistré :\n{path}")
        except Exception as exc:
            messagebox.showerror("Erreur d'exportation", str(exc))

    def reset(self):
        self.chart_var.set("Histogramme + KDE")
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, "Analyse statistique des variables")
        self.palette_var.set("deep")
        self.bins_var.set(20)
        self.figsize_var.set("Grande (12 × 7)")
        self.dpi_var.set(300)
        self.kde_var.set(True)
        self.reg_var.set(True)
        self.hue_var.set("Aucune")
        self._chart_changed()
        if self.df is not None:
            self.generate_plot()


if __name__ == "__main__":
    root = tk.Tk()
    app = SeabornProApp(root)
    root.mainloop()
