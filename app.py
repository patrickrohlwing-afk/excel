from flask import Flask, request, render_template_string, jsonify
import pandas as pd
import os

app = Flask(__name__)

DATA_FILE = "data.xlsx"

HTML = """
<!DOCTYPE html>
<html>
<head>
<title>Excel Tool</title>

<link href="https://unpkg.com/tabulator-tables@5.5.0/dist/css/tabulator.min.css" rel="stylesheet">
<script src="https://unpkg.com/tabulator-tables@5.5.0/dist/js/tabulator.min.js"></script>

<style>
body {
    font-family: Arial;
    background: #1e1e2f;
    color: white;
    padding: 20px;
}

input, button {
    padding: 10px;
    margin: 5px 0;
    border-radius: 8px;
    border: none;
}

button {
    background: #4CAF50;
    color: white;
    cursor: pointer;
}

#table {
    margin-top: 20px;
}
</style>
</head>

<body>

<h2>📊 Excel Tool (Stabile Version)</h2>

<form method="post">
<input name="search" placeholder="11000/1">
<input name="value" placeholder="Wert">
<button type="submit">Eintragen</button>
</form>

<button onclick="saveData()">💾 Tabelle speichern</button>

<div id="table"></div>

<script>
var table;

function loadTable(data){
    table = new Tabulator("#table", {
        data: data,
        layout: "fitColumns",
        reactiveData:true,
        columns: Object.keys(data[0] || {}).map(key => ({
            title: key,
            field: key,
            editor: "input"
        }))
    });
}

fetch("/data")
.then(res => res.json())
.then(data => loadTable(data));

function saveData(){
    fetch("/save", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body: JSON.stringify(table.getData())
    }).then(()=>alert("Gespeichert ✔"));
}
</script>

</body>
</html>
"""

# 🔹 HAUPTSEITE
@app.route("/", methods=["GET", "POST"])
def index():

    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(columns=["C", "D", "E"])
        df.to_excel(DATA_FILE, index=False)

    if request.method == "POST":
        search = request.form.get("search")
        value = request.form.get("value")

        df = pd.read_excel(DATA_FILE, dtype=str)

        if search and "/" in search:
            a, b = search.split("/")

            col_c = df.columns[0]
            col_d = df.columns[1]
            col_e = df.columns[2]

            mask = (df[col_c].astype(str) == a) & (df[col_d].astype(str) == b)

            if mask.any():
                df.loc[mask, col_e] = value

        df.to_excel(DATA_FILE, index=False)

    return render_template_string(HTML)

# 🔹 DATEN LADEN
@app.route("/data")
def get_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype=str)
        return df.fillna("").to_json(orient="records")
    return "[]"

# 🔹 SPEICHERN
@app.route("/save", methods=["POST"])
def save():
    data = request.get_json()
    df = pd.DataFrame(data)
    df.to_excel(DATA_FILE, index=False)
    return jsonify({"status": "ok"})

# 🔹 START
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
