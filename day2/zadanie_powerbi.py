import pandas as pd

# 1) Wczytanie danych z Excela (wiele arkuszy)
xls = pd.ExcelFile("sprzedaz.xlsx")
sprzedaz = pd.read_excel(xls, "Transakcje")      # kolumny: Data, SKU, Ilosc, Cena, Region
produkty = pd.read_excel(xls, "Produkty")        # kolumny: SKU, Kategoria

# 2) Transformacje jak w Power Query
# - merge (join) po SKU
df = sprzedaz.merge(produkty, on="SKU", how="left")

# - czyszczenie: usuwanie wierszy bez SKU lub z ujemną ilością
df = df.dropna(subset=["SKU"])
df = df[df["Ilosc"] > 0]

# - ewentualny unpivot/pivot (tu pivot do tabeli przestawnej po Region/Kategoria)
df["Przychod"] = df["Ilosc"] * df["Cena"]

# 3) „Miary” (KPI) jak DAX, ale w pandas
kpi_global = {
    "przychod": df["Przychod"].sum(),
    "sztuki": df["Ilosc"].sum(),
    "srednia_cena": (df["Przychod"].sum() / df["Ilosc"].sum())
}

# 4) Tabela przestawna (odpowiednik Matrix w Power BI)
pivot = pd.pivot_table(
    df,
    values="Przychod",
    index=["Region"],
    columns=["Kategoria"],
    aggfunc="sum",
    margins=True,
    margins_name="Suma"
).round(2)

# 5) Zapis wyników do Excela (na wielu arkuszach)
with pd.ExcelWriter("raport_python.xlsx", engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Dane_po_transformacjach", index=False)
    pivot.to_excel(writer, sheet_name="Pivot_Przychod")
    pd.DataFrame([kpi_global]).to_excel(writer, sheet_name="KPI", index=False)

print("Gotowe: raport_python.xlsx")