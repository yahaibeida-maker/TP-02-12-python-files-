# -*- coding: utf-8 -*-
"""
TP 01 - CREATION D'UN TABLEAU DE BORD INTERACTIF AVEC PYTHON
Master 01 | Logiciels libres et Open Source | 2026-2027

Application pedagogique avec :
- Importation Excel (.xlsx) ou CSV
- Jeu de demonstration synthetique au demarrage
- Selection des colonnes Date et variables numeriques
- Filtrage par periode
- KPI dynamiques : observations, moyenne, maximum, valeurs manquantes
- Courbe temporelle, histogramme, boxplot, nuage de points
- Matrice de correlation Pearson / Spearman
- Interface web locale avec Dash

Installation :
    conda install -c conda-forge pandas numpy plotly dash openpyxl

Execution :
    Lancer ce fichier dans Spyder puis ouvrir http://127.0.0.1:8050
"""

import base64
import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from dash import Dash, dcc, html, dash_table, Input, Output, State, no_update


# ============================================================
# 1. DONNEES DE DEMONSTRATION
# ============================================================
def create_demo_data():
    """Jeu synthetique pour permettre de tester le TP sans fichier externe."""
    rng = np.random.default_rng(21)
    dates = pd.date_range("2025-01-01", periods=180, freq="D")

    tw = 17 + 4 * np.sin(np.linspace(0, 3 * np.pi, len(dates))) + rng.normal(0, 1, len(dates))
    cod = 20 + 2.0 * tw + rng.normal(0, 4, len(dates))
    bod = 2.5 + 0.30 * cod + rng.normal(0, 1.1, len(dates))
    do = 13.0 - 0.22 * tw - 0.07 * bod + rng.normal(0, 0.45, len(dates))
    tn = 1.2 + 0.035 * cod + rng.normal(0, 0.20, len(dates))
    tp = 0.10 + 0.014 * bod + rng.normal(0, 0.035, len(dates))
    ss = 12 + 1.7 * bod + rng.normal(0, 3.0, len(dates))

    # Quelques valeurs manquantes pour les exercices de qualite des donnees.
    for col, indices in {"DO": [12, 54, 119], "COD": [31, 88], "SS": [42, 140]}.items():
        data = {"DO": do, "COD": cod, "SS": ss}[col]
        data[indices] = np.nan

    return pd.DataFrame({
        "Date": dates, "Tw": tw, "SS": ss, "TN": tn,
        "TP": tp, "COD": cod, "BOD": bod, "DO": do
    })


# ============================================================
# 2. IMPORTATION ET PREPARATION
# ============================================================
def read_uploaded_file(contents, filename):
    if not contents or not filename:
        raise ValueError("Veuillez selectionner un fichier.")

    try:
        _, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)
    except Exception as exc:
        raise ValueError("Le fichier transmis est illisible.") from exc

    extension = Path(filename).suffix.lower()
    if extension in (".xlsx", ".xls"):
        data = pd.read_excel(io.BytesIO(decoded), sheet_name=0)
    elif extension == ".csv":
        try:
            data = pd.read_csv(io.BytesIO(decoded), sep=None, engine="python")
        except UnicodeDecodeError:
            data = pd.read_csv(io.BytesIO(decoded), sep=None, engine="python",
                               encoding="latin-1")
    else:
        raise ValueError("Format non supporte. Utilisez .xlsx, .xls ou .csv.")

    if data.empty:
        raise ValueError("Le fichier ne contient aucune ligne.")
    data.columns = [str(c).strip() for c in data.columns]
    return data


def to_json(data):
    frame = data.copy()
    for col in frame.columns:
        if pd.api.types.is_datetime64_any_dtype(frame[col]):
            frame[col] = frame[col].dt.strftime("%Y-%m-%dT%H:%M:%S")
    return frame.to_json(orient="split", date_format="iso")


def from_json(data):
    if not data:
        return pd.DataFrame()
    return pd.read_json(io.StringIO(data), orient="split")


def detect_date_column(data):
    for col in data.columns:
        if any(word in str(col).lower() for word in ("date", "time", "jour")):
            parsed = pd.to_datetime(data[col], errors="coerce", dayfirst=True)
            if parsed.notna().mean() >= 0.5:
                return col
    for col in data.columns:
        parsed = pd.to_datetime(data[col], errors="coerce", dayfirst=True)
        if parsed.notna().mean() >= 0.8:
            return col
    return None


def numeric_columns(data, date_col=None):
    cols = []
    for col in data.columns:
        if col == date_col:
            continue
        values = pd.to_numeric(
            data[col].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        )
        if values.notna().sum() >= 2:
            cols.append(col)
    return cols


def prepare_data(raw, date_col, value_col=None):
    if date_col not in raw.columns:
        raise ValueError("Selectionnez une colonne Date valide.")
    result = raw.copy()
    result["_Date"] = pd.to_datetime(result[date_col], errors="coerce", dayfirst=True)
    result = result[result["_Date"].notna()].copy()
    if result.empty:
        raise ValueError("Aucune date valide n'a ete trouvee.")
    result = result.sort_values("_Date")
    return result


# ============================================================
# 3. STYLE ET COMPOSANTS
# ============================================================
NAVY = "#17365D"
BLUE = "#2471A3"
GREEN = "#238B57"
RED = "#C0392B"
PURPLE = "#7551A1"
BG = "#F2F5F9"

def make_kpi(title, value, accent=BLUE, subtitle=""):
    return html.Div([
        html.Div(title, style={"fontSize": "13px", "fontWeight": "bold",
                               "color": "#526273"}),
        html.Div(value, style={"fontSize": "27px", "fontWeight": "bold",
                              "color": accent, "marginTop": "7px"}),
        html.Div(subtitle, style={"fontSize": "11px", "color": "#718096"})
    ], style={"background": "white", "border": "1px solid #D5DDE7",
              "borderLeft": f"5px solid {accent}", "borderRadius": "8px",
              "padding": "13px", "flex": "1", "minWidth": "145px"})


def empty_fig(message):
    fig = go.Figure()
    fig.add_annotation(text=message, x=0.5, y=0.5, xref="paper", yref="paper",
                       showarrow=False, font={"size": 15})
    fig.update_layout(template="plotly_white", height=360)
    return fig


def card(children):
    return html.Div(children, style={"background": "white", "padding": "14px",
                                     "border": "1px solid #D5DDE7",
                                     "borderRadius": "8px"})


# ============================================================
# 4. APPLICATION DASH
# ============================================================
initial_data = create_demo_data()
app = Dash(__name__)
app.title = "TP 01 - Dashboard interactif"

app.layout = html.Div([
    dcc.Store(id="dataset-store", data=to_json(initial_data)),
    dcc.Store(id="source-store", data="Jeu de demonstration synthetique"),

    html.Div([
        html.H1("TP 01 : TABLEAU DE BORD INTERACTIF",
                style={"fontFamily": "Times New Roman", "fontSize": "30px",
                       "margin": "0", "fontWeight": "bold"}),
        html.Div("Python | Pandas | Plotly | Dash | Master 01",
                 style={"fontSize": "16px", "marginTop": "7px"})
    ], style={"background": NAVY, "color": "white", "padding": "22px 26px"}),

    html.Div([
        html.H3("1. Importation des donnees", style={"color": NAVY}),
        dcc.Upload(
            id="upload",
            children=html.Div(["Glissez un fichier Excel/CSV ici ou ",
                               html.A("cliquez pour parcourir")]),
            multiple=False,
            style={"width": "100%", "height": "65px", "lineHeight": "65px",
                   "borderWidth": "2px", "borderStyle": "dashed",
                   "borderColor": BLUE, "borderRadius": "8px",
                   "textAlign": "center", "background": "#F7FAFD",
                   "cursor": "pointer"}
        ),
        html.Div(id="upload-message",
                 children="Jeu de demonstration charge. Vous pouvez importer vos propres donnees.",
                 style={"marginTop": "8px", "fontSize": "13px", "color": GREEN})
    ], style={"padding": "15px", "background": "white", "margin": "14px",
              "borderRadius": "8px", "border": "1px solid #D5DDE7"}),

    html.Div([
        html.Div([
            html.Label("Colonne Date", style={"fontWeight": "bold"}),
            dcc.Dropdown(id="date-col", options=[{"label": "Date", "value": "Date"}],
                         value="Date", clearable=False)
        ], style={"flex": "1", "minWidth": "190px"}),
        html.Div([
            html.Label("Variable principale", style={"fontWeight": "bold"}),
            dcc.Dropdown(id="main-var", options=[
                {"label": c, "value": c} for c in initial_data.columns if c != "Date"
            ], value="DO", clearable=False)
        ], style={"flex": "1", "minWidth": "190px"}),
        html.Div([
            html.Label("Variable X (nuage de points)", style={"fontWeight": "bold"}),
            dcc.Dropdown(id="x-var", options=[
                {"label": c, "value": c} for c in initial_data.columns if c != "Date"
            ], value="COD", clearable=False)
        ], style={"flex": "1", "minWidth": "190px"}),
        html.Div([
            html.Label("Variable Y (nuage de points)", style={"fontWeight": "bold"}),
            dcc.Dropdown(id="y-var", options=[
                {"label": c, "value": c} for c in initial_data.columns if c != "Date"
            ], value="BOD", clearable=False)
        ], style={"flex": "1", "minWidth": "190px"})
    ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap",
              "padding": "15px", "background": "white", "margin": "14px",
              "border": "1px solid #D5DDE7", "borderRadius": "8px"}),

    card([
        html.Label("Periode d'analyse", style={"fontWeight": "bold"}),
        dcc.DatePickerRange(id="date-range", display_format="DD/MM/YYYY",
                            minimum_nights=0,
                            start_date=initial_data["Date"].min().date(),
                            end_date=initial_data["Date"].max().date())
    ]),

    html.H3("2. Indicateurs cles de performance (KPI)",
            style={"color": NAVY, "margin": "18px 14px 10px"}),
    html.Div(id="kpi-row", style={"display": "flex", "gap": "10px",
                                  "flexWrap": "wrap", "margin": "0 14px"}),

    html.H3("3. Visualisations interactives",
            style={"color": NAVY, "margin": "20px 14px 10px"}),
    html.Div([
        dcc.Graph(id="time-fig", config={"displaylogo": False}),
        dcc.Graph(id="hist-fig", config={"displaylogo": False}),
        dcc.Graph(id="box-fig", config={"displaylogo": False}),
        dcc.Graph(id="scatter-fig", config={"displaylogo": False})
    ], style={"display": "grid", "gridTemplateColumns": "repeat(2, minmax(0, 1fr))",
              "gap": "12px", "margin": "0 14px"}),

    card([
        html.H3("4. Matrice de correlation", style={"color": NAVY}),
        html.Div([
            html.Div([
                html.Label("Variables a comparer", style={"fontWeight": "bold"}),
                dcc.Dropdown(id="corr-vars", multi=True,
                             value=["Tw", "SS", "TN", "TP", "COD", "BOD", "DO"],
                             options=[{"label": c, "value": c}
                                      for c in ["Tw", "SS", "TN", "TP", "COD", "BOD", "DO"]])
            ], style={"flex": "3", "minWidth": "250px"}),
            html.Div([
                html.Label("Methode", style={"fontWeight": "bold"}),
                dcc.Dropdown(id="corr-method", value="pearson", clearable=False,
                             options=[{"label": "Pearson", "value": "pearson"},
                                      {"label": "Spearman", "value": "spearman"}])
            ], style={"flex": "1", "minWidth": "160px"})
        ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap"}),
        dcc.Graph(id="corr-fig", config={"displaylogo": False})
    ]),

    card([
        html.H3("5. Apercu des donnees", style={"color": NAVY}),
        dash_table.DataTable(
            id="data-table", page_size=8, sort_action="native",
            style_table={"overflowX": "auto"},
            style_header={"background": NAVY, "color": "white",
                          "fontWeight": "bold"},
            style_cell={"padding": "8px", "textAlign": "center",
                        "fontSize": "12px"},
            style_data_conditional=[{"if": {"row_index": "odd"},
                                     "backgroundColor": "#F2F5F9"}]
        )
    ]),

    html.Div("TP pedagogique | Application locale developpee avec Dash et Plotly",
             style={"textAlign": "center", "padding": "16px",
                    "color": "#667788", "fontSize": "12px"})
], style={"background": BG, "minHeight": "100vh", "padding": "8px",
          "fontFamily": "Arial, sans-serif", "color": "#263238"})


# ============================================================
# 5. CALLBACK : CHARGEMENT DU FICHIER
# ============================================================
@app.callback(
    Output("dataset-store", "data"),
    Output("source-store", "data"),
    Output("date-col", "options"),
    Output("date-col", "value"),
    Output("main-var", "options"),
    Output("main-var", "value"),
    Output("x-var", "options"),
    Output("x-var", "value"),
    Output("y-var", "options"),
    Output("y-var", "value"),
    Output("corr-vars", "options"),
    Output("corr-vars", "value"),
    Output("upload-message", "children"),
    Output("upload-message", "style"),
    Output("date-range", "min_date_allowed"),
    Output("date-range", "max_date_allowed"),
    Output("date-range", "start_date"),
    Output("date-range", "end_date"),
    Input("upload", "contents"),
    State("upload", "filename"),
    prevent_initial_call=True
)
def upload_dataset(contents, filename):
    if not contents:
        return (no_update,) * 18
    try:
        raw = read_uploaded_file(contents, filename)
        date_col = detect_date_column(raw)
        if date_col is None:
            raise ValueError("Aucune colonne de date reconnue dans le fichier.")

        nums = numeric_columns(raw, date_col)
        if not nums:
            raise ValueError("Aucune variable numerique exploitable n'a ete trouvee.")

        dates = pd.to_datetime(raw[date_col], errors="coerce", dayfirst=True)
        valid_dates = dates.dropna()
        if valid_dates.empty:
            raise ValueError("La colonne date ne contient aucune date valide.")

        options = [{"label": str(c), "value": str(c)} for c in nums]
        date_options = [{"label": str(c), "value": str(c)} for c in raw.columns]
        default_y = nums[1] if len(nums) > 1 else nums[0]
        start, end = valid_dates.min().date(), valid_dates.max().date()
        msg = f"Fichier charge : {filename} | {len(raw)} lignes | {len(raw.columns)} colonnes."
        return (to_json(raw), filename, date_options, date_col,
                options, nums[0], options, nums[0], options, default_y,
                options, nums[:min(7, len(nums))], msg,
                {"marginTop": "8px", "fontSize": "13px", "color": GREEN},
                start, end, start, end)
    except Exception as exc:
        return ((no_update,) * 12 + (f"Erreur : {exc}",
                {"marginTop": "8px", "fontSize": "13px", "color": RED}) + (no_update,) * 4)


# ============================================================
# 6. CALLBACK PRINCIPAL : KPI ET GRAPHIQUES
# ============================================================
@app.callback(
    Output("kpi-row", "children"),
    Output("time-fig", "figure"),
    Output("hist-fig", "figure"),
    Output("box-fig", "figure"),
    Output("scatter-fig", "figure"),
    Output("data-table", "data"),
    Output("data-table", "columns"),
    Input("dataset-store", "data"),
    Input("date-col", "value"),
    Input("main-var", "value"),
    Input("x-var", "value"),
    Input("y-var", "value"),
    Input("date-range", "start_date"),
    Input("date-range", "end_date")
)
def update_dashboard(json_data, date_col, main_var, x_var, y_var,
                     start_date, end_date):
    try:
        raw = from_json(json_data)
        if date_col not in raw.columns:
            raise ValueError("Colonne date invalide.")

        raw["_parsed_date"] = pd.to_datetime(raw[date_col], errors="coerce", dayfirst=True)
        data = raw[raw["_parsed_date"].notna()].copy()
        if start_date:
            data = data[data["_parsed_date"] >= pd.to_datetime(start_date)]
        if end_date:
            data = data[data["_parsed_date"] < pd.to_datetime(end_date) + pd.Timedelta(days=1)]

        if data.empty:
            raise ValueError("Aucune observation dans la periode selectionnee.")
        if main_var not in data.columns:
            raise ValueError("Variable principale invalide.")

        data[main_var] = pd.to_numeric(
            data[main_var].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        )
        valid = data.dropna(subset=[main_var]).copy()
        if valid.empty:
            raise ValueError("La variable selectionnee ne contient pas de valeurs numeriques.")

        values = valid[main_var]
        missing = int(data[main_var].isna().sum())
        kpis = [
            make_kpi("Observations", f"{len(data):,}".replace(",", " "), BLUE),
            make_kpi(f"Moyenne - {main_var}", f"{values.mean():.3f}", GREEN),
            make_kpi(f"Maximum - {main_var}", f"{values.max():.3f}", RED),
            make_kpi("Valeurs manquantes", str(missing), PURPLE),
            make_kpi("Variables numeriques", str(len(numeric_columns(raw, date_col))), "#D97706")
        ]

        time_fig = px.line(valid.sort_values("_parsed_date"),
                           x="_parsed_date", y=main_var, markers=True,
                           title=f"Evolution temporelle de {main_var}")
        time_fig.update_traces(line={"width": 2.5, "color": BLUE},
                               marker={"size": 6})
        time_fig.update_layout(template="plotly_white", height=360,
                               xaxis_title="Date", yaxis_title=main_var,
                               hovermode="x unified")

        hist_fig = px.histogram(valid, x=main_var, nbins=20,
                                title=f"Distribution de {main_var}",
                                color_discrete_sequence=["#55A6A6"])
        hist_fig.add_vline(x=values.mean(), line_dash="dash", line_color=RED,
                           annotation_text=f"Moyenne = {values.mean():.2f}")
        hist_fig.update_layout(template="plotly_white", height=360,
                               yaxis_title="Effectif")

        box_fig = px.box(valid, y=main_var, points="outliers",
                         title=f"Dispersion de {main_var}",
                         color_discrete_sequence=[PURPLE])
        box_fig.update_layout(template="plotly_white", height=360)

        if x_var in data.columns and y_var in data.columns:
            pair = data[[x_var, y_var]].copy()
            for col in [x_var, y_var]:
                pair[col] = pd.to_numeric(pair[col].astype(str).str.replace(",", ".", regex=False),
                                           errors="coerce")
            pair = pair.dropna()
            if not pair.empty:
                scatter_fig = px.scatter(pair, x=x_var, y=y_var,
                                         title=f"Relation entre {x_var} et {y_var}",
                                         trendline="ols" if len(pair) >= 3 else None,
                                         opacity=0.75)
                scatter_fig.update_traces(marker={"size": 9})
                scatter_fig.update_layout(template="plotly_white", height=360)
            else:
                scatter_fig = empty_fig("Aucune paire de valeurs numeriques valide.")
        else:
            scatter_fig = empty_fig("Selectionnez deux variables numeriques.")

        table_frame = data.drop(columns=["_parsed_date"]).head(100)
        table_frame = table_frame.astype(object).where(pd.notna(table_frame), None)
        table_data = table_frame.to_dict("records")
        table_columns = [{"name": str(c), "id": str(c)} for c in table_frame.columns]

        return kpis, time_fig, hist_fig, box_fig, scatter_fig, table_data, table_columns

    except Exception as exc:
        empty = empty_fig(str(exc))
        return [make_kpi("Statut", "A verifier", RED)], empty, empty, empty, empty, [], []


# ============================================================
# 7. CALLBACK : MATRICE DE CORRELATION
# ============================================================
@app.callback(
    Output("corr-vars", "options"),
    Output("corr-vars", "value"),
    Input("dataset-store", "data"),
    State("date-col", "value")
)
def update_correlation_options(json_data, date_col):
    try:
        raw = from_json(json_data)
        cols = numeric_columns(raw, date_col)
        opts = [{"label": str(c), "value": str(c)} for c in cols]
        return opts, cols[:min(7, len(cols))]
    except Exception:
        return [], []


@app.callback(
    Output("corr-fig", "figure"),
    Input("dataset-store", "data"),
    Input("date-col", "value"),
    Input("date-range", "start_date"),
    Input("date-range", "end_date"),
    Input("corr-vars", "value"),
    Input("corr-method", "value")
)
def update_correlation(json_data, date_col, start_date, end_date,
                        selected_vars, method):
    try:
        raw = from_json(json_data)
        dates = pd.to_datetime(raw[date_col], errors="coerce", dayfirst=True)
        mask = dates.notna()
        if start_date:
            mask &= dates >= pd.to_datetime(start_date)
        if end_date:
            mask &= dates < pd.to_datetime(end_date) + pd.Timedelta(days=1)
        frame = raw.loc[mask]
        cols = [c for c in (selected_vars or []) if c in frame.columns]
        numeric = pd.DataFrame()
        for col in cols:
            numeric[col] = pd.to_numeric(
                frame[col].astype(str).str.replace(",", ".", regex=False),
                errors="coerce"
            )
        numeric = numeric.dropna(axis=1, how="all")
        if numeric.shape[1] < 2:
            return empty_fig("Selectionnez au moins deux variables numeriques.")
        corr = numeric.corr(method=method)
        fig = go.Figure(go.Heatmap(
            z=corr.values, x=corr.columns, y=corr.index,
            zmin=-1, zmax=1, colorscale="RdBu", reversescale=True,
            text=np.round(corr.values, 2), texttemplate="%{text:.2f}",
            colorbar={"title": "Coefficient"},
            hovertemplate="%{y} / %{x}<br>Correlation : %{z:.3f}<extra></extra>"
        ))
        fig.update_layout(title=f"Matrice de correlation - {method.capitalize()}",
                          template="plotly_white", height=520)
        return fig
    except Exception as exc:
        return empty_fig(f"Correlation impossible : {exc}")


# ============================================================
# 8. LANCEMENT DU SERVEUR LOCAL
# ============================================================
if __name__ == "__main__":
    print("-" * 65)
    print("TP 01 - TABLEAU DE BORD INTERACTIF")
    print("Ouvrir : http://127.0.0.1:8050")
    print("Arreter le serveur : Ctrl + C dans la console Spyder")
    print("-" * 65)
    app.run(debug=False, host="127.0.0.1", port=8050)
