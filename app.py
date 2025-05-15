from flask import Flask, request, render_template, send_file
import pandas as pd
from scipy.stats import chi2_contingency
from openpyxl.styles import Alignment, Font, Border, Side
import os

app = Flask(__name__)

def przetworz_wielokrotny_wybor(df, kolumna):
    df[kolumna] = df[kolumna].apply(lambda x: str(x).split('; ') if isinstance(x, str) else [x])
    return df

def stworz_tabele(df, p1, p2, pytania):
    df = przetworz_wielokrotny_wybor(df, p1)
    df = przetworz_wielokrotny_wybor(df, p2)
    df_expanded = df.explode(p1).explode(p2)
    tabela = pd.crosstab(df_expanded[p1], df_expanded[p2])

    tabela['Suma_wierszy'] = tabela.sum(axis=1)
    tabela.loc['Suma_kolumn'] = tabela.sum(axis=0)

    tabela_bez_sum = tabela.drop(columns=['Suma_wierszy'], index=['Suma_kolumn'])
    chi2, p, dof, _ = chi2_contingency(tabela_bez_sum)

    with pd.ExcelWriter("processed.xlsx", engine="openpyxl") as writer:
        tabela.to_excel(writer, sheet_name="Wyniki")
        df_chi = pd.DataFrame([{
            'Statystyka chi²': chi2,
            'P-value': p,
            'Stopnie swobody': dof
        }])
        df_chi.to_excel(writer, sheet_name="Wyniki", startrow=tabela.shape[0] + 2)

@app.route("/", methods=["GET", "POST"])
def index():
    pytania = []
    plik_gotowy = False

    if request.method == "POST":
        if "plik" in request.files:
            f = request.files["plik"]
            f.save("upload.xlsx")
            try:
                df = pd.read_excel("upload.xlsx", sheet_name="dane")
                pytania = df.columns.tolist()
                return render_template("index.html", pytania=pytania)
            except Exception as e:
                return f"Blad: {e}"
        elif "pytanie1" in request.form and "pytanie2" in request.form:
            p1 = request.form["pytanie1"]
            p2 = request.form["pytanie2"]
            df = pd.read_excel("upload.xlsx", sheet_name="dane")
            pytania = df.columns.tolist()
            stworz_tabele(df, p1, p2, pytania)
            plik_gotowy = True
            return render_template("index.html", pytania=pytania, plik_gotowy=plik_gotowy)

    return render_template("index.html")

@app.route("/pobierz")
def pobierz():
    return send_file("processed.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
