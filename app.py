from flask import Flask, request, render_template_string, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DATA_FILE = os.path.join(UPLOAD_FOLDER, "data.xlsx")

HTML = """
<!DOCTYPE html>
<html>
<head>
<title>Excel Tool Pro</title>

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
    margin: 5px;
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

<h2>📊 Excel Tool Pro (Cloud Version)</h2>

<form method="post" enctype="multipart/form-data">
<input type="file" name="file">
<button type="submit">📤 Hochladen</button>
</form>

<button onclick="saveData()">💾 Speichern</button>
<a href="/download"><button>📥 Excel Download</button></a>

<div id="table"></div>

<script>
var table;

// Tabelle laden
function loadTable(data){
    table = new Tabulator("#table", {
        data: data,
        layout: "fitColumns",
        reactiveData:true,
        autoColumns:true,
    });
}

// Daten laden
fetch("/data")
.then(res => res.json())
.then(data => loadTable(data));

// speichern
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

    # Datei hochladen
    if request.method == "POST":
        f = request.files.get("file")

        if f and f.filename:
            f.save(DATA_FILE)

    # Datei erstellen falls nicht vorhanden
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame([
            {"C": "11000", "D": "1", "E": "Startwert"}
        ])
        df.to_excel(DATA_FILE, index=False)

    return render_template_string(HTML)


# 🔹 DATEN LADEN
@app.route("/data")
def get_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype=str)
        return jsonify(df.fillna("").to_dict(orient="records"))
    return jsonify([])


# 🔹 SPEICHERN
@app.route("/save", methods=["POST"])
def save():
    data = request.get_json()
    df = pd.DataFrame(data)
    df.to_excel(DATA_FILE, index=False)
    return jsonify({"status": "ok"})


# 🔹 DOWNLOAD
@app.route("/download")
def download():
    if os.path.exists(DATA_FILE):
        return send_file(DATA_FILE, as_attachment=True)
    return "Keine Datei vorhanden"


# 🔹 START
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
