# -*- coding: utf-8 -*-
"""
APPLICATION 03 : MATPLOTLIB PRO V02
Visualisation scientifique avec Tkinter + Matplotlib + Pandas + NumPy.
Aucun recours à Seaborn. Compatible Spyder / Anaconda.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


APP_TITLE = "APPLICATION 03 : MATPLOTLIB PRO V02"
DEFAULT_FILE = "OPEN SOURCE DATA FILE.xlsx"
SHEET_NAME = "DATA"

BG = "#F2F5F9"
NAVY = "#17365D"
BLUE = "#2471A3"
GREEN = "#238B57"
WHITE = "#FFFFFF"


class MatplotlibProApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1580x960")
        self.root.minsize(1200, 760)
        self.root.configure(bg=BG)

        self.df = None
        self.numeric_cols = []
        self.current_figure = None
        self.canvas = None
        self.toolbar = None
        self.plot_color = "#2471A3"
        self.regression_var = tk.BooleanVar(value=True)
        self.annotation_var = tk.BooleanVar(value=False)
        self.density_var = tk.BooleanVar(value=True)

        self._build_interface()
        self._load_startup_file()

    # ------------------------- UI -------------------------
    def _build_interface(self):
        header = tk.Frame(self.root, bg=NAVY, height=88)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=APP_TITLE, bg=NAVY, fg="white",
                 font=("Times New Roman", 25, "bold")).pack(anchor="w", padx=22, pady=(10, 0))
        tk.Label(header, text="VISUALISATION SCIENTIFIQUE • ANALYSE GRAPHIQUE • EXPORTATION",
                 bg=NAVY, fg="#DCE6F1", font=("Times New Roman", 13, "bold")).pack(
                     anchor="w", padx=26, pady=(1, 8))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=10, pady=10)

        self.sidebar = tk.Frame(body, bg=WHITE, width=355, highlightbackground="#D5DDE7",
                                highlightthickness=1)
        self.sidebar.pack(side="left", fill="y", padx=(0, 9))
        self.sidebar.pack_propagate(False)

        right = tk.Frame(body, bg=BG)
        right.pack(side="right", fill="both", expand=True)
        self.plot_area = tk.Frame(right, bg=WHITE, highlightbackground="#D5DDE7",
                                  highlightthickness=1)
        self.plot_area.pack(fill="both", expand=True)
        self._build_sidebar()
        self._build_plot_area(right)

        footer = tk.Frame(self.root, bg=NAVY, height=27)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(footer, text="MASTER 01  |  LOGICIELS LIBRES ET OPEN SOURCE 2026–2027",
                 bg=NAVY, fg="white", font=("Arial", 9, "bold")).pack(pady=4)

    def _section(self, parent, title):
        tk.Label(parent, text=title, bg=WHITE, fg=NAVY,
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=12, pady=(9, 4))

    def _build_sidebar(self):
        self._section(self.sidebar, "1. SOURCE DES DONNÉES")
        row = tk.Frame(self.sidebar, bg=WHITE)
        row.pack(fill="x", padx=11)
        self.file_label = tk.Label(row, text="Aucun fichier chargé", bg="#EEF3F8",
                                   fg="#34495E", anchor="w", justify="left",
                                   wraplength=225, font=("Arial", 8))
        self.file_label.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 5))
        ttk.Button(row, text="Parcourir", command=self.choose_file).pack(side="right")
        self.dataset_info = tk.Label(self.sidebar, text="En attente du fichier Excel…",
                                     bg=WHITE, fg="#526273", justify="left",
                                     anchor="w", font=("Arial", 8), wraplength=325)
        self.dataset_info.pack(fill="x", padx=12, pady=(5, 0))

        ttk.Separator(self.sidebar).pack(fill="x", padx=10, pady=6)
        self._section(self.sidebar, "2. TYPE DE VISUALISATION")
        self.chart_type = tk.StringVar(value="Courbe")
        self.chart_combo = ttk.Combobox(
            self.sidebar, textvariable=self.chart_type, state="readonly",
            values=["Courbe", "Nuage de points", "Observé vs Prédit",
                    "Histogramme", "Histogramme + densité", "Boxplot",
                    "Diagramme en barres", "Heatmap de corrélation",
                    "Tableau 2 × 2", "Comparaison multi-variables"]
        )
        self.chart_combo.pack(fill="x", padx=12, pady=(0, 6))
        self.chart_combo.bind("<<ComboboxSelected>>", self._chart_changed)

        self._section(self.sidebar, "3. VARIABLES")
        self._small_label("Variable X / variable à représenter")
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(self.sidebar, textvariable=self.x_var, state="readonly")
        self.x_combo.pack(fill="x", padx=12, pady=(0, 5))
        self._small_label("Variable Y / deuxième série")
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(self.sidebar, textvariable=self.y_var, state="readonly")
        self.y_combo.pack(fill="x", padx=12, pady=(0, 5))

        self._small_label("Sélection multiple (comparaison)")
        self.multi_list = tk.Listbox(self.sidebar, selectmode="extended", height=4,
                                     exportselection=False, font=("Arial", 8),
                                     activestyle="dotbox")
        self.multi_list.pack(fill="x", padx=12, pady=(0, 5))

        self._section(self.sidebar, "4. PERSONNALISATION")
        grid = tk.Frame(self.sidebar, bg=WHITE)
        grid.pack(fill="x", padx=12)
        tk.Label(grid, text="Couleur", bg=WHITE, font=("Arial", 8)).grid(row=0, column=0, sticky="w")
        self.color_var = tk.StringVar(value="Bleu")
        self.color_combo = ttk.Combobox(grid, textvariable=self.color_var, state="readonly",
                                        values=["Bleu", "Rouge", "Vert", "Orange", "Violet", "Noir", "Personnalisée"],
                                        width=13)
        self.color_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
        self.color_combo.bind("<<ComboboxSelected>>", self._choose_color)
        tk.Label(grid, text="Marqueur", bg=WHITE, font=("Arial", 8)).grid(row=1, column=0, sticky="w")
        self.marker_var = tk.StringVar(value="o")
        ttk.Combobox(grid, textvariable=self.marker_var, state="readonly",
                     values=["o", "s", "^", "D", "*", "x", "+", "."],
                     width=13).grid(row=1, column=1, sticky="ew", padx=4, pady=2)
        tk.Label(grid, text="Taille police", bg=WHITE, font=("Arial", 8)).grid(row=2, column=0, sticky="w")
        self.font_size = tk.IntVar(value=12)
        ttk.Spinbox(grid, from_=8, to=30, textvariable=self.font_size, width=7).grid(
            row=2, column=1, sticky="w", padx=4, pady=2)
        tk.Label(grid, text="Taille points", bg=WHITE, font=("Arial", 8)).grid(row=3, column=0, sticky="w")
        self.point_size = tk.IntVar(value=48)
        ttk.Spinbox(grid, from_=10, to=300, increment=5, textvariable=self.point_size,
                    width=7).grid(row=3, column=1, sticky="w", padx=4, pady=2)
        tk.Label(grid, text="DPI export", bg=WHITE, font=("Arial", 8)).grid(row=4, column=0, sticky="w")
        self.dpi_value = tk.IntVar(value=300)
        ttk.Spinbox(grid, from_=72, to=1200, increment=50, textvariable=self.dpi_value,
                    width=7).grid(row=4, column=1, sticky="w", padx=4, pady=2)
        grid.columnconfigure(1, weight=1)

        self._section(self.sidebar, "5. OPTIONS D'ANALYSE")
        ttk.Checkbutton(self.sidebar, text="Ajouter régression linéaire et R² (Scatter)",
                        variable=self.regression_var).pack(anchor="w", padx=12)
        ttk.Checkbutton(self.sidebar, text="Annoter les points extrêmes",
                        variable=self.annotation_var).pack(anchor="w", padx=12)
        ttk.Checkbutton(self.sidebar, text="Afficher densité lissée (histogramme)",
                        variable=self.density_var).pack(anchor="w", padx=12)

        self._section(self.sidebar, "6. TITRE ET EXPORTATION")
        self.title_entry = ttk.Entry(self.sidebar)
        self.title_entry.pack(fill="x", padx=12, pady=(0, 5))
        self.title_entry.insert(0, "Visualisation scientifique des données")
        buttons = tk.Frame(self.sidebar, bg=WHITE)
        buttons.pack(fill="x", padx=10, pady=4)
        ttk.Button(buttons, text="GÉNÉRER", command=self.generate_plot).pack(
            side="left", fill="x", expand=True, padx=2, ipady=5)
        ttk.Button(buttons, text="EXPORTER", command=self.export_figure).pack(
            side="left", fill="x", expand=True, padx=2, ipady=5)
        ttk.Button(buttons, text="RESET", command=self.reset_plot).pack(
            side="left", fill="x", expand=True, padx=2, ipady=5)

        self.help_label = tk.Label(
            self.sidebar,
            text="Aide : sélectionnez un graphique, ses variables et ses options, puis cliquez sur GÉNÉRER.",
            bg="#EEF3F8", fg="#34495E", justify="left", anchor="w",
            font=("Arial", 8), wraplength=320, padx=8, pady=6
        )
        self.help_label.pack(fill="x", padx=10, pady=(5, 8))

    def _small_label(self, text):
        tk.Label(self.sidebar, text=text, bg=WHITE, font=("Arial", 8, "bold")).pack(
            anchor="w", padx=12, pady=(2, 1))

    def _build_plot_area(self, parent):
        top = tk.Frame(self.plot_area, bg=WHITE)
        top.pack(fill="x", padx=12, pady=8)
        tk.Label(top, text="ESPACE DE VISUALISATION", bg=WHITE, fg=NAVY,
                 font=("Arial", 13, "bold")).pack(side="left")
        self.status_label = tk.Label(top, text="Prêt", bg=WHITE, fg=GREEN,
                                     font=("Arial", 8, "bold"))
        self.status_label.pack(side="right")

        self.canvas_frame = tk.Frame(self.plot_area, bg=WHITE)
        self.canvas_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        self.info_panel = tk.Frame(parent, bg=WHITE, highlightbackground="#D5DDE7",
                                   highlightthickness=1)
        self.info_panel.pack(fill="x", pady=(8, 0))
        tk.Label(self.info_panel, text="INFORMATIONS ET STATISTIQUES",
                 bg=WHITE, fg=NAVY, font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(5, 0))
        self.stats_text = tk.Label(self.info_panel, text="Les statistiques apparaîtront après la génération.",
                                   bg=WHITE, fg="#34495E", font=("Arial", 9),
                                   justify="left", anchor="w", wraplength=1100)
        self.stats_text.pack(fill="x", padx=10, pady=5)

        self.welcome = tk.Label(
            self.canvas_frame,
            text="MATPLOTLIB PRO V02\n\nChargez un fichier Excel, configurez le graphique\npuis cliquez sur GÉNÉRER.",
            bg=WHITE, fg="#607080", justify="center",
            font=("Times New Roman", 19, "bold"))
        self.welcome.pack(expand=True)

    # ----------------------- DATA --------------------------
    def _load_startup_file(self):
        for path in (os.path.join(os.getcwd(), DEFAULT_FILE),
                     os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_FILE)):
            if os.path.isfile(path):
                self.load_file(path)
                return

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier de données",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Fichiers CSV", "*.csv"),
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
                    messagebox.showwarning("Feuille Excel",
                        f"Feuille '{SHEET_NAME}' introuvable : première feuille chargée.")
            data = data.dropna(how="all").dropna(axis=1, how="all")
            data.columns = [str(c).strip() for c in data.columns]
            if data.empty:
                raise ValueError("Le fichier ne contient aucune observation.")
            self.df = data
            self.numeric_cols = data.select_dtypes(include=np.number).columns.tolist()
            if not self.numeric_cols:
                raise ValueError("Aucune variable numérique n'a été détectée.")
            self.file_label.config(text=os.path.basename(path))
            self.dataset_info.config(
                text=f"Fichier : {os.path.basename(path)} | Observations : {len(data):,} | "
                     f"Variables : {len(data.columns)} | Numériques : {len(self.numeric_cols)}")
            self.x_combo["values"] = self.numeric_cols
            self.y_combo["values"] = self.numeric_cols
            self.x_var.set(self.numeric_cols[0])
            self.y_var.set(self.numeric_cols[1] if len(self.numeric_cols) > 1 else self.numeric_cols[0])
            self.multi_list.delete(0, "end")
            for col in self.numeric_cols:
                self.multi_list.insert("end", col)
            self.status_label.config(text="Données chargées", fg=GREEN)
            self.generate_plot()
        except Exception as exc:
            messagebox.showerror("Erreur de chargement", str(exc))
            self.status_label.config(text="Erreur de chargement", fg="#C0392B")

    # ---------------------- OPTIONS ------------------------
    def _chart_changed(self, _event=None):
        chart = self.chart_type.get()
        multi = chart in ("Heatmap de corrélation", "Tableau 2 × 2", "Comparaison multi-variables")
        state = "disabled" if multi else "readonly"
        self.x_combo.configure(state=state)
        self.y_combo.configure(state=state)
        help_map = {
            "Observé vs Prédit": "Sélectionnez Observé en X et Prédit en Y. Le graphique calcule les indicateurs de comparaison.",
            "Nuage de points": "Le nuage de points explore la relation entre deux variables. La droite de régression et R² sont optionnels.",
            "Histogramme + densité": "La densité est une approximation lissée calculée à partir de l'histogramme avec NumPy.",
            "Comparaison multi-variables": "Sélectionnez plusieurs variables dans la liste (Ctrl ou Maj)."
        }
        self.help_label.config(text=help_map.get(chart, "Configurez les paramètres puis cliquez sur GÉNÉRER."))

    def _choose_color(self, _event=None):
        colors = {"Bleu": "#2471A3", "Rouge": "#C0392B", "Vert": "#238B57",
                  "Orange": "#D97706", "Violet": "#7551A1", "Noir": "#222222"}
        if self.color_var.get() == "Personnalisée":
            chosen = colorchooser.askcolor(title="Choisir une couleur")[1]
            if chosen:
                self.plot_color = chosen
        else:
            self.plot_color = colors.get(self.color_var.get(), "#2471A3")

    # --------------------- PLOTTING -------------------------
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

    def _valid_pair(self, x, y):
        pair = self.df[[x, y]].apply(pd.to_numeric, errors="coerce").dropna()
        if len(pair) < 2:
            raise ValueError("Il faut au moins deux paires de valeurs numériques valides.")
        return pair[x].to_numpy(), pair[y].to_numpy()

    def _apply_axes_style(self, ax):
        fs = int(self.font_size.get())
        ax.tick_params(axis="both", labelsize=max(8, fs - 1))
        ax.grid(True, alpha=0.25)
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontname("Times New Roman")

    def _linear_fit(self, x, y):
        slope, intercept = np.polyfit(x, y, 1)
        pred = slope * x + intercept
        ss_res = np.sum((y - pred) ** 2)
        ss_tot = np.sum((y - np.mean(y)) ** 2)
        r2 = 1 - ss_res / ss_tot if ss_tot > 0 else np.nan
        order = np.argsort(x)
        return slope, intercept, r2, x[order], pred[order]

    def _update_stats(self, variable=None, x=None, y=None, chart=None, n=None):
        parts = [f"Type : {chart}", f"Nombre d'observations représentées : {n if n is not None else len(self.df)}"]
        if variable and variable in self.df:
            s = pd.to_numeric(self.df[variable], errors="coerce").dropna()
            if len(s):
                parts.append(f"{variable} — moyenne : {s.mean():.4g} | médiane : {s.median():.4g} | "
                             f"écart-type : {s.std():.4g} | min : {s.min():.4g} | max : {s.max():.4g}")
        if x is not None and y is not None:
            r = np.corrcoef(x, y)[0, 1] if len(x) > 1 else np.nan
            parts.append(f"Corrélation de Pearson : r = {r:.4f}")
        self.stats_text.config(text="     |     ".join(parts))

    def generate_plot(self):
        if self.df is None:
            messagebox.showwarning("Données manquantes", "Veuillez charger un fichier.")
            return
        chart = self.chart_type.get()
        title = self.title_entry.get().strip() or "Visualisation scientifique"
        x_name, y_name = self.x_var.get(), self.y_var.get()
        color, marker = self.plot_color, self.marker_var.get()
        fs = int(self.font_size.get())

        try:
            if chart in ("Courbe", "Nuage de points", "Observé vs Prédit"):
                if not x_name or not y_name:
                    raise ValueError("Sélectionnez les variables X et Y.")
                x, y = self._valid_pair(x_name, y_name)
                fig, ax = self._new_figure()
                if chart == "Courbe":
                    ax.plot(np.arange(len(y)), y, color=color, marker=marker,
                            markersize=4, linewidth=1.8, label=y_name)
                    ax.set_xlabel("Indice de l'observation")
                    ax.set_ylabel(y_name)
                    self._update_stats(y_name, n=len(y), chart=chart)
                elif chart == "Nuage de points":
                    ax.scatter(x, y, s=int(self.point_size.get()), color=color,
                               marker=marker, alpha=0.75, edgecolors="black",
                               linewidths=0.35, label="Observations")
                    self._update_stats(x=x, y=y, n=len(x), chart=chart)
                    if self.regression_var.get():
                        if len(np.unique(x)) < 2:
                            raise ValueError("La variable X doit contenir au moins deux valeurs distinctes.")
                        slope, intercept, r2, xx, yy = self._linear_fit(x, y)
                        ax.plot(xx, yy, color="#C0392B", linewidth=2,
                                label=f"Régression linéaire (R² = {r2:.3f})")
                        ax.legend()
                    if self.annotation_var.get():
                        for idx in [int(np.argmin(y)), int(np.argmax(y))]:
                            ax.annotate(f"({x[idx]:.3g}, {y[idx]:.3g})",
                                        (x[idx], y[idx]), xytext=(6, 7),
                                        textcoords="offset points", fontsize=fs-2)
                    ax.set_xlabel(x_name)
                    ax.set_ylabel(y_name)
                else:
                    ax.scatter(x, y, s=int(self.point_size.get()), color=color,
                               alpha=0.75, edgecolors="black", linewidths=0.35,
                               label="Paires observées")
                    lo = min(float(np.min(x)), float(np.min(y)))
                    hi = max(float(np.max(x)), float(np.max(y)))
                    ax.plot([lo, hi], [lo, hi], "--", color="#C0392B",
                            linewidth=2, label="Ligne 1:1")
                    err = y - x
                    rmse = float(np.sqrt(np.mean(err ** 2)))
                    mae = float(np.mean(np.abs(err)))
                    bias = float(np.mean(err))
                    r2 = float(np.corrcoef(x, y)[0, 1] ** 2) if len(x) > 1 else np.nan
                    self.stats_text.config(text=f"Observé = {x_name} | Prédit = {y_name} | N = {len(x)} | "
                        f"RMSE = {rmse:.4g} | MAE = {mae:.4g} | Biais = {bias:.4g} | R² (corrélation²) = {r2:.4f}")
                    ax.set_xlabel(f"Observé — {x_name}")
                    ax.set_ylabel(f"Prédit — {y_name}")
                    ax.legend()
                ax.set_title(title, fontsize=fs+2, fontweight="bold")
                self._apply_axes_style(ax)

            elif chart in ("Histogramme", "Histogramme + densité"):
                values = pd.to_numeric(self.df[x_name], errors="coerce").dropna().to_numpy()
                if len(values) == 0:
                    raise ValueError("Aucune valeur valide dans la variable sélectionnée.")
                fig, ax = self._new_figure()
                counts, bins, _ = ax.hist(values, bins="auto", density=(chart == "Histogramme + densité"),
                                          color=color, edgecolor="black", alpha=0.72)
                if chart == "Histogramme + densité" and self.density_var.get() and len(values) > 2:
                    centers = (bins[:-1] + bins[1:]) / 2
                    kernel = np.ones(5) / 5
                    smooth = np.convolve(counts, kernel, mode="same")
                    ax.plot(centers, smooth, color="#C0392B", linewidth=2.2,
                            marker="o", markersize=3, label="Densité lissée (approximation)")
                    ax.legend()
                ax.set_xlabel(x_name)
                ax.set_ylabel("Densité" if chart == "Histogramme + densité" else "Fréquence")
                ax.set_title(title, fontsize=fs+2, fontweight="bold")
                self._update_stats(x_name, n=len(values), chart=chart)
                self._apply_axes_style(ax)

            elif chart == "Boxplot":
                values = pd.to_numeric(self.df[x_name], errors="coerce").dropna().to_numpy()
                if len(values) == 0:
                    raise ValueError("Aucune valeur valide dans la variable sélectionnée.")
                fig, ax = self._new_figure()
                ax.boxplot(values, patch_artist=True, tick_labels=[x_name],
                           showmeans=True, boxprops={"facecolor": color, "alpha": 0.65})
                ax.set_ylabel(x_name)
                ax.set_title(title, fontsize=fs+2, fontweight="bold")
                self._update_stats(x_name, n=len(values), chart=chart)
                self._apply_axes_style(ax)

            elif chart == "Diagramme en barres":
                values = pd.to_numeric(self.df[x_name], errors="coerce").dropna()
                counts = values.value_counts().sort_index()
                if len(counts) > 40:
                    counts = values.round(1).value_counts().sort_index()
                fig, ax = self._new_figure(figsize=(11, 6))
                ax.bar([str(v) for v in counts.index], counts.values,
                       color=color, edgecolor="black", alpha=0.8)
                ax.set_xlabel(x_name)
                ax.set_ylabel("Effectif")
                ax.set_title(title, fontsize=fs+2, fontweight="bold")
                ax.tick_params(axis="x", labelrotation=45)
                self._update_stats(x_name, n=len(values), chart=chart)
                self._apply_axes_style(ax)

            elif chart == "Heatmap de corrélation":
                corr = self.df[self.numeric_cols].corr()
                fig, ax = self._new_figure(figsize=(10, 8))
                image = ax.imshow(corr, cmap="coolwarm", vmin=-1, vmax=1, aspect="auto")
                ax.set_xticks(range(len(corr.columns)))
                ax.set_yticks(range(len(corr.columns)))
                ax.set_xticklabels(corr.columns, rotation=45, ha="right", fontsize=fs-1)
                ax.set_yticklabels(corr.columns, fontsize=fs-1)
                ax.set_title(title, fontsize=fs+2, fontweight="bold")
                for i in range(len(corr.columns)):
                    for j in range(len(corr.columns)):
                        value = corr.iloc[i, j]
                        ax.text(j, i, f"{value:.2f}", ha="center", va="center",
                                fontsize=max(7, fs-3))
                fig.colorbar(image, ax=ax, fraction=0.046, pad=0.04, label="Corrélation")
                self._update_stats(chart=chart, n=len(self.df))

            elif chart == "Tableau 2 × 2":
                fig, axes = self._new_figure(2, 2, (12, 8))
                axes = np.asarray(axes).ravel()
                for ax, col in zip(axes, self.numeric_cols[:4]):
                    vals = pd.to_numeric(self.df[col], errors="coerce").dropna()
                    ax.hist(vals, bins="auto", color=color, alpha=0.75, edgecolor="black")
                    ax.set_title(f"Distribution de {col}", fontsize=fs, fontweight="bold")
                    ax.set_xlabel(col)
                    ax.set_ylabel("Fréquence")
                    self._apply_axes_style(ax)
                for ax in axes[min(4, len(self.numeric_cols)):]:
                    ax.set_visible(False)
                fig.suptitle(title, fontsize=fs+3, fontweight="bold")
                self._update_stats(chart=chart, n=len(self.df))

            elif chart == "Comparaison multi-variables":
                selected = [self.multi_list.get(i) for i in self.multi_list.curselection()]
                if len(selected) < 2:
                    raise ValueError("Sélectionnez au moins deux variables dans la liste multiple.")
                fig, ax = self._new_figure(figsize=(11, 6))
                for col in selected:
                    vals = pd.to_numeric(self.df[col], errors="coerce")
                    # Standardisation descriptive pour comparer des variables d'unités différentes.
                    sd = vals.std()
                    if pd.isna(sd) or sd == 0:
                        continue
                    z = (vals - vals.mean()) / sd
                    ax.plot(np.arange(len(z)), z, linewidth=1.5, label=col)
                ax.set_xlabel("Indice de l'observation")
                ax.set_ylabel("Valeur standardisée (z-score)")
                ax.set_title(title, fontsize=fs+2, fontweight="bold")
                ax.legend(ncol=2)
                self._apply_axes_style(ax)
                self._update_stats(chart=f"{chart} — variables : {', '.join(selected)}",
                                   n=len(self.df))

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

    # ----------------------- EXPORT ------------------------
    def export_figure(self):
        if self.current_figure is None:
            messagebox.showwarning("Exportation", "Générez d'abord une figure.")
            return
        path = filedialog.asksaveasfilename(
            title="Exporter la figure", defaultextension=".png",
            filetypes=[("Image PNG", "*.png"), ("Image JPEG", "*.jpg"),
                       ("Document PDF", "*.pdf"), ("Image SVG", "*.svg")])
        if not path:
            return
        try:
            dpi = max(72, int(self.dpi_value.get()))
            self.current_figure.savefig(path, dpi=dpi, bbox_inches="tight",
                                        facecolor="white")
            messagebox.showinfo("Exportation réussie", f"Figure enregistrée :\n{path}")
        except Exception as exc:
            messagebox.showerror("Erreur d'exportation", str(exc))

    def reset_plot(self):
        self.chart_type.set("Courbe")
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, "Visualisation scientifique des données")
        self.color_var.set("Bleu")
        self.plot_color = "#2471A3"
        self.marker_var.set("o")
        self.font_size.set(12)
        self.point_size.set(48)
        self.dpi_value.set(300)
        self.regression_var.set(True)
        self.annotation_var.set(False)
        self.density_var.set(True)
        self._chart_changed()
        if self.df is not None:
            self.generate_plot()


if __name__ == "__main__":
    root = tk.Tk()
    app = MatplotlibProApp(root)
    root.mainloop()

