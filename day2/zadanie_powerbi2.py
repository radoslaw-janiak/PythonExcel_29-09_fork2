import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

from dash import Dash, html, dcc, dash_table, Input, Output, callback, no_update
import plotly.express as px

FILE = "sprzedaz.xlsx"  # plik przygotowany wcześniej (Transakcje, Produkty)

# -------------------------
# Ładowanie i przygotowanie danych
# -------------------------
def load_data(path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    trans = pd.read_excel(xls, "Transakcje", parse_dates=["Data"])
    prod = pd.read_excel(xls, "Produkty")
    df = trans.merge(prod, on="SKU", how="left")
    df["Przychod"] = df["Ilosc"] * df["Cena"]
    # Guard: w razie braków
    df["Kategoria"] = df["Kategoria"].fillna("Brak")
    df["Region"] = df["Region"].fillna("Brak")
    # Dodatkowe pola czasu
    df["Rok"] = df["Data"].dt.year
    df["Miesiac"] = df["Data"].dt.to_period("M").dt.to_timestamp()
    return df

if not Path(FILE).exists():
    raise SystemExit(
        f"Nie znaleziono pliku {FILE}. Upewnij się, że jest w tym samym folderze."
    )

df0 = load_data(FILE)

# Wartości do filtrów
min_date, max_date = df0["Data"].min().date(), df0["Data"].max().date()
regiony = sorted(df0["Region"].dropna().unique().tolist())
kategorie = sorted(df0["Kategoria"].dropna().unique().tolist())

def fmt_pln(x: float) -> str:
    try:
        return f"{x:,.2f} zł".replace(",", " ").replace(".", ",")
    except Exception:
        return str(x)

# -------------------------
# Aplikacja Dash
# -------------------------
app = Dash(__name__, title="Dashboard sprzedaży (Dash)")
app.layout = html.Div(
    style={"padding": "18px", "fontFamily": "Inter, system-ui, -apple-system, Segoe UI"},
    children=[
        html.H1("Dashboard sprzedaży — Python + Dash", style={"marginBottom": "8px"}),
        html.Div("Źródło: sprzedaz.xlsx (arkusze: Transakcje, Produkty)"),

        # Filtry (jak slicery)
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1.2fr 1fr 1fr", "gap": "12px", "marginTop": "16px"},
            children=[
                html.Div(
                    children=[
                        html.Label("Zakres dat"),
                        dcc.DatePickerRange(
                            id="f-date",
                            min_date_allowed=min_date,
                            max_date_allowed=max_date,
                            start_date=min_date,
                            end_date=max_date,
                            display_format="YYYY-MM-DD",
                            clearable=True
                        ),
                    ]
                ),
                html.Div(
                    children=[
                        html.Label("Region"),
                        dcc.Dropdown(
                            id="f-region",
                            options=[{"label": r, "value": r} for r in regiony],
                            multi=True,
                            placeholder="Wybierz region(y)",
                        ),
                    ]
                ),
                html.Div(
                    children=[
                        html.Label("Kategoria"),
                        dcc.Dropdown(
                            id="f-kategoria",
                            options=[{"label": k, "value": k} for k in kategorie],
                            multi=True,
                            placeholder="Wybierz kategorię(e)",
                        ),
                    ]
                ),
            ],
        ),

        # KPI
        html.Div(
            id="kpi-cards",
            style={
                "display": "grid",
                "gridTemplateColumns": "repeat(3, 1fr)",
                "gap": "12px",
                "marginTop": "16px",
            },
        ),

        # Wykresy
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1.5fr 1fr", "gap": "12px", "marginTop": "12px"},
            children=[
                dcc.Graph(id="g-line"),
                dcc.Graph(id="g-bar-cat"),
            ],
        ),
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "12px", "marginTop": "12px"},
            children=[
                dcc.Graph(id="g-bar-region"),
                dcc.Graph(id="g-top-products"),
            ],
        ),

        # Tabela przestawna
        html.H3("Tabela przestawna (Region × Kategoria, suma przychodu)", style={"marginTop": "18px"}),
        dash_table.DataTable(
            id="pivot",
            page_size=12,
            style_table={"overflowX": "auto"},
            style_header={"fontWeight": "bold"},
            style_cell={"padding": "6px", "fontSize": "14px"},
            export_format="xlsx",
            export_headers="display",
        ),
        html.Div(
            "Wskazówka: kliknij nagłówek kolumny, aby sortować; użyj pola eksportu, aby zapisać tabelę.",
            style={"marginTop": "6px", "color": "#444"}
        ),

        html.Hr(),
        html.Div(
            "Podoba Ci się ta wersja? Mogę dorobić miary, porównania okresów i dodatkowe wykresy na życzenie.",
            style={"color": "#666"}
        ),
    ],
)

# -------------------------
# Callbacks
# -------------------------
@callback(
    Output("kpi-cards", "children"),
    Output("g-line", "figure"),
    Output("g-bar-cat", "figure"),
    Output("g-bar-region", "figure"),
    Output("g-top-products", "figure"),
    Output("pivot", "columns"),
    Output("pivot", "data"),
    Input("f-date", "start_date"),
    Input("f-date", "end_date"),
    Input("f-region", "value"),
    Input("f-kategoria", "value"),
)
def update_dashboard(start_date, end_date, sel_regions, sel_cats):
    df = df0.copy()

    # Filtr dat
    if start_date:
        df = df[df["Data"] >= pd.to_datetime(start_date)]
    if end_date:
        df = df[df["Data"] <= pd.to_datetime(end_date)]

    # Filtry kategorii i regionów
    if sel_regions and len(sel_regions) > 0:
        df = df[df["Region"].isin(sel_regions)]
    if sel_cats and len(sel_cats) > 0:
        df = df[df["Kategoria"].isin(sel_cats)]

    if df.empty:
        # Puste figi + puste pivot
        empty_fig = px.scatter(title="Brak danych dla wybranych filtrów")
        kpi = [
            card("Przychód", "0,00 zł"),
            card("Sztuki", "0"),
            card("Śr. cena", "0,00 zł"),
        ]
        return kpi, empty_fig, empty_fig, empty_fig, empty_fig, [], []

    # KPI
    total_rev = float(df["Przychod"].sum())
    total_qty = int(df["Ilosc"].sum())
    avg_price = float(df["Przychod"].sum() / df["Ilosc"].sum()) if total_qty else 0.0

    kpi = [
        card("Przychód", fmt_pln(total_rev)),
        card("Sztuki", f"{total_qty:,}".replace(",", " ")),
        card("Śr. cena", fmt_pln(avg_price)),
    ]

    # Wykres liniowy po miesiącach (sumy przychodu)
    ts = df.groupby("Miesiac", as_index=False)["Przychod"].sum().sort_values("Miesiac")
    fig_line = px.line(ts, x="Miesiac", y="Przychod", title="Przychód w czasie (miesięcznie)")
    fig_line.update_layout(yaxis_tickformat=",")  # sep tysięcy

    # Słupki: kategoria
    by_cat = df.groupby("Kategoria", as_index=False)["Przychod"].sum().sort_values("Przychod", ascending=False)
    fig_cat = px.bar(by_cat, x="Kategoria", y="Przychod", title="Przychód wg kategorii")
    fig_cat.update_layout(yaxis_tickformat=",")

    # Słupki: region
    by_reg = df.groupby("Region", as_index=False)["Przychod"].sum().sort_values("Przychod", ascending=False)
    fig_reg = px.bar(by_reg, x="Region", y="Przychod", title="Przychód wg regionu")
    fig_reg.update_layout(yaxis_tickformat=",")

    # Top produkty (po przychodzie)
    by_prod = (
        df.groupby(["SKU", "Nazwa"], as_index=False)["Przychod"].sum()
        .sort_values("Przychod", ascending=False)
        .head(15)
    )
    fig_top = px.bar(by_prod, x="Przychod", y="Nazwa", orientation="h", title="Top produkty (przychód)")
    fig_top.update_layout(xaxis_tickformat=",")

    # Pivot: Region × Kategoria (suma przychodu)
    pivot = pd.pivot_table(
        df, values="Przychod", index="Region", columns="Kategoria", aggfunc="sum", fill_value=0.0, margins=True
    )
    pivot = pivot.reset_index().rename(columns={"Region": "Region/Kategoria"})
    # Formatowanie prezentacyjne (tekstowe) w DataTable
    pivot_fmt = pivot.copy()
    for col in pivot_fmt.columns:
        if col != "Region/Kategoria":
            pivot_fmt[col] = pivot_fmt[col].map(fmt_pln)

    columns = [{"name": c, "id": c} for c in pivot_fmt.columns]
    data = pivot_fmt.to_dict("records")

    return kpi, fig_line, fig_cat, fig_reg, fig_top, columns, data


def card(title: str, value: str):
    return html.Div(
        style={
            "border": "1px solid #eaeaea",
            "borderRadius": "16px",
            "padding": "14px 16px",
            "boxShadow": "0 1px 3px rgba(0,0,0,0.06)",
            "background": "white",
        },
        children=[
            html.Div(title, style={"fontSize": "14px", "color": "#666"}),
            html.Div(value, style={"fontSize": "22px", "fontWeight": 700, "marginTop": "4px"}),
        ],
    )


if __name__ == "__main__":
    app.run(debug=True)