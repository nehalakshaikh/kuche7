from flask import Flask, render_template, request, send_file
import pandas as pd
import numpy as np
import math
import os

app = Flask(__name__)

def handle(df):
    a = ''
    handle = ''
    edge = ''
    bhl = []
    bhc = []
    thl = []
    thc = []
    lhl = []
    lhc = []
    noegl = []
    noegc = []
    ehl = []
    ehc = []
    
    for i in range(len(df)):  # Corrected this line
        if isinstance(df.iloc[i, 0], str):
            a = ('base' if df.iloc[i, 0].lower() == 'base unit'
                 else 'tall' if df.iloc[i, 0].lower() == 'tall unit'
                 else 'wall' if df.iloc[i, 0].lower() == 'wall unit'
                 else 'loft' if df.iloc[i, 0].lower() == 'loft unit'
                 else 'filler' if (df.iloc[i, 0].lower()).startswith('fillers')
                 else a)

        if isinstance(df.iloc[i, 14], str):
            if (isinstance(df.iloc[i, 14], str)) & (handle == ''):
                handle = df.iloc[i, 14]

        if a == 'base':
            if (isinstance(df.iloc[i, 15], float)) & ~(math.isnan(df.iloc[i, 15])):

                if (df.iloc[i, 15]-2) in bhl:
                    bhc[bhl.index(df.iloc[i, 15]-2)] += df.iloc[i, 16]
                else:
                    bhl.append(df.iloc[i, 15]-2)
                    bhc.append(df.iloc[i, 16])

        if a == 'tall':
            if (isinstance(df.iloc[i, 15], float)) & ~(math.isnan(df.iloc[i, 15])):

                if (df.iloc[i, 15]-2) in thl:
                    thc[thl.index(df.iloc[i, 15]-2)] += df.iloc[i, 16]
                else:
                    thl.append(df.iloc[i, 15]-2)
                    thc.append(df.iloc[i, 16])

        if a == 'loft':
            if (isinstance(df.iloc[i, 15], float)) & ~(math.isnan(df.iloc[i, 15])):

                if (df.iloc[i, 15]-2) in lhl:
                    lhc[lhl.index(df.iloc[i, 15]-2)] += df.iloc[i, 16]
                else:
                    lhl.append(df.iloc[i, 15]-2)
                    lhc.append(df.iloc[i, 16])

        if a == 'wall':
            if df.iloc[i, 14] == handle:
                if (isinstance(df.iloc[i, 15], float)) & ~(math.isnan(df.iloc[i, 15])):

                    if (df.iloc[i, 15]-2) in lhl:
                        noegc[noegl.index(df.iloc[i, 15]-2)] += df.iloc[i, 16]
                    else:
                        noegl.append(df.iloc[i, 15]-2)
                        noegc.append(df.iloc[i, 16])
            else:
                if isinstance(df.iloc[i, 14], str):
                    if (isinstance(df.iloc[i, 14], str)) & (edge == ''):
                        edge = df.iloc[i, 14]

                if (isinstance(df.iloc[i, 15], float)) & ~(math.isnan(df.iloc[i, 15])):

                    if (df.iloc[i, 15]-35) in ehl:
                        ehc[ehl.index(df.iloc[i, 15]-35)] += df.iloc[i, 16]
                    else:
                        ehl.append(df.iloc[i, 15]-35)
                        ehc.append(df.iloc[i, 16])

    data = {
        'Handle Type': handle,
        'Base Handle': bhl,
        'BQTY': bhc,
        'Tall Handle': thl,
        'TQTY': thc,
        'Loft Handle': lhl,
        'LTY': lhc,
        'Wall(no edge)': noegl,
        'WNQTY': noegc,
        'Edge handle Type': edge,
        'Edge handle': ehl,
        'EQTY': ehc
    }
    dataframe = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in data.items()]))
    output_path = os.path.join('static', 'h.xlsx')
    dataframe.to_excel(output_path)
    return output_path

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            df = pd.read_excel(file,skiprows=4,skipfooter=16)
            output_path = handle(df)
            return send_file(output_path, as_attachment=True)
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
