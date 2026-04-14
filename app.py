from flask import (
    Flask,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.utils import secure_filename
import os
import uuid
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

HTML = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
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

    .toolbar {
      display: flex;
      gap: 10px;
      margin: 10px 0;
      flex-wrap: wrap;
    }

    input, button {
      padding: 10px;
      border-radius: 8px;
      border: none;
    }

    button {
      background: #4CAF50;
      color: white;
      cursor: pointer;
    }

    .danger { background: #c0392b; }
    .secondary { background: #555; }

    #table {
      margin-top: 20px;
      background: white;
      border-radius: 8px;
    }

    .msg {
      margin-top: 10px;
      color: #b2ffb2;
    }
  </style>
</head>

<body>

<h2>📊 Excel Tool Pro</h2>

<form method="post" enctype="multipart/form-data" class="toolbar">
  <input type="file" name="file" accept=".xlsx,.xls" required>
  <button type="submit">📤 Hochladen</button>
  <span>Aktive Datei: <strong>{{ active_filename }}</strong></span>
</form>

<div class="toolbar">
  <button onclick="addRow()">➕ Zeile</button>
  <button class="danger" onclick="deleteSelected()">🗑️ Löschen</button>
  <button onclick="saveData()">💾 Speichern</button>
  <button class="secondary" onclick="download()">📥 Download</button>
</div>

<div id="msg" class="msg"></div>
<div id="table"></div>

<script>
let table;

function msg(text, err=false){
  const el = document.getElementById("msg");
  el.style.color = err ? "#ffb3b3" : "#b2ffb2";
  el.innerText = text;
}

function build(data){
  const cols = data.length ? Object.keys(data[0]) : ["C","D","E"];

  table = new Tabulator("#table", {
    data:data,
    layout:"fitDataStretch",
    selectableRows:true,
    reactiveData:true,
    height:"60vh",
    columns: cols.map(c => ({
      title:c,
      field:c,
      editor:"input"
    }))
  });
}

async function load(){
  const res = await fetch("/data");
  const data = await res.json();
  build(data);
}

function addRow(){
  const row = {};
  (table.getData()[0] || {"C":"","D":"","E":""});
  Object.keys(table.getData()[0] || {"C":"","D":"","E":""})
    .forEach(k => row[k]="");
  table.addRow(row);
}

function deleteSelected(){
  table.getSelectedRows().forEach(r=>r.delete());
}

async function saveData(){
  const res = await fetch("/save", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify(table.getData())
  });
  if(res.ok) msg("Gespeichert ✔");
  else msg("Fehler beim Speichern", true);
}

function download(){
  window.location.href="/download";
}

load();
</script>

</body>
</html>
"""

# 🔹 Hilfsfunktionen
def allowed_file(filename):
    return os.path.splitext(filename.lower())[1] in ALLOWED_EXTENSIONS

def create_default_excel(path):
    df = pd.DataFrame([{"C": "11000", "D": "1", "E": "Startwert"}])
    df.to_excel(path, index=False)

def get_active_file():
    file_path = session.get("file_path")
    original_name = session.get("original_name", "data.xlsx")

    if file_path and os.path.exists(file_path):
        return file_path, original_name

    default = os.path.join(UPLOAD_FOLDER, "default.xlsx")
    if not os.path.exists(default):
        create_default_excel(default)

    session["file_path"] = default
    session["original_name"] = "data.xlsx"

    return default, "data.xlsx"


# 🔹 ROUTES

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")

        if f and f.filename:
            name = secure_filename(f.filename)

            if not allowed_file(name):
                return "Nur Excel erlaubt", 400

            ext = os.path.splitext(name)[1]
            new_name = f"{uuid.uuid4().hex}{ext}"
            path = os.path.join(UPLOAD_FOLDER, new_name)

            f.save(path)

            session["file_path"] = path
            session["original_name"] = name

        return redirect(url_for("index"))

    _, name = get_active_file()
    return render_template_string(HTML, active_filename=name)


@app.route("/data")
def data():
    file_path, _ = get_active_file()
    df = pd.read_excel(file_path, dtype=str)
    return jsonify(df.fillna("").to_dict(orient="records"))


@app.route("/save", methods=["POST"])
def save():
    file_path, _ = get_active_file()
    data = request.get_json()

    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

    return jsonify({"status":"ok"})


@app.route("/download")
def download():
    file_path, name = get_active_file()
    return send_file(file_path, as_attachment=True, download_name=name)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
app.run(host="0.0.0.0", port=port)
