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
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Excel Tool Pro</title>

  <link href="https://unpkg.com/tabulator-tables@5.5.0/dist/css/tabulator.min.css" rel="stylesheet">
  <script src="https://unpkg.com/tabulator-tables@5.5.0/dist/js/tabulator.min.js"></script>

  <style>
    body {
      font-family: Arial, sans-serif;
      background: #1e1e2f;
      color: #fff;
      padding: 20px;
      margin: 0;
    }

    .wrap {
      max-width: 1200px;
      margin: 0 auto;
    }

    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      margin: 14px 0;
    }

    input, button {
      padding: 10px;
      border-radius: 8px;
      border: none;
      font-size: 14px;
    }

    input[type="file"] {
      background: #2d2d42;
      color: #fff;
    }

    button {
      background: #4caf50;
      color: white;
      cursor: pointer;
    }

    button:hover {
      filter: brightness(1.1);
    }

    .secondary {
      background: #555;
    }

    .danger {
      background: #c0392b;
    }

    .msg {
      margin-top: 8px;
      color: #b2ffb2;
      min-height: 20px;
    }

    #table {
      margin-top: 18px;
      background: white;
      border-radius: 8px;
      overflow: hidden;
    }

    .small {
      opacity: 0.85;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h2>📊 Excel Tool Pro</h2>
    <p class="small">Datei hochladen → in Tabelle bearbeiten → speichern → als Excel herunterladen.</p>

    <form method="post" enctype="multipart/form-data" class="toolbar">
      <input type="file" name="file" accept=".xlsx,.xls" required>
      <button type="submit">📤 Hochladen</button>
      <span class="small">Aktive Datei: <strong>{{ active_filename }}</strong></span>
    </form>

    <div class="toolbar">
      <button type="button" onclick="addRow()">➕ Zeile hinzufügen</button>
      <button type="button" class="danger" onclick="deleteSelectedRows()">🗑️ Markierte Zeilen löschen</button>
      <button type="button" onclick="saveData()">💾 Speichern</button>
      <button type="button" class="secondary" onclick="downloadExcel()">📥 Excel Download</button>
    </div>

    <div id="msg" class="msg"></div>
    <div id="table"></div>
  </div>

  <script>
    let table = null;

    function setMessage(text, isError=false) {
      const msg = document.getElementById("msg");
      msg.style.color = isError ? "#ffb3b3" : "#b2ffb2";
      msg.textContent = text;
    }

    function buildTable(data) {
      const cols = data.length ? Object.keys(data[0]) : ["A", "B", "C"];
      const tabColumns = cols.map(c => ({
        title: c,
        field: c,
        editor: "input",
      }));

      table = new Tabulator("#table", {
        data: data,
        layout: "fitDataStretch",
        selectableRows: true,
        reactiveData: true,
        columns: tabColumns,
        height: "65vh",
      });
    }

    async function loadData() {
      try {
        const res = await fetch("/data");
        if (!res.ok) throw new Error("Daten konnten nicht geladen werden");
        const data = await res.json();
        buildTable(data);
      } catch (err) {
        setMessage(err.message, true);
      }
    }

    function addRow() {
      if (!table) return;
      const first = table.getData()[0] || {"A": "", "B": "", "C": ""};
      const empty = {};
      Object.keys(first).forEach(k => empty[k] = "");
      table.addRow(empty, false);
    }

    function deleteSelectedRows() {
      if (!table) return;
      const selected = table.getSelectedRows();
      if (!selected.length) {
        setMessage("Bitte zuerst mindestens eine Zeile markieren.", true);
        return;
      }
      selected.forEach(r => r.delete());
      setMessage(selected.length + " Zeile(n) gelöscht.");
    }

    async function saveData() {
      if (!table) return;
      try {
        const res = await fetch("/save", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify(table.getData())
        });
        const payload = await res.json();
        if (!res.ok) throw new Error(payload.error || "Fehler beim Speichern");
        setMessage("Gespeichert ✔");
      } catch (err) {
        setMessage(err.message, true);
      }
    }

    function downloadExcel() {
      window.location.href = "/download";
    }

    loadData();
  </script>
</body>
</html>
"""


def allowed_file(filename: str) -> bool:
    return os.path.splitext(filename.lower())[1] in ALLOWED_EXTENSIONS


def create_default_excel(path: str) -> None:
    df = pd.DataFrame([{"C": "11000", "D": "1", "E": "Startwert"}])
    df.to_excel(path, index=False)


def get_active_file() -> tuple[str, str]:
    file_path = session.get("file_path")
    original_name = session.get("original_name", "data.xlsx")

    if file_path and os.path.exists(file_path):
        return file_path, original_name

    default_path = os.path.join(UPLOAD_FOLDER, "default_data.xlsx")
    if not os.path.exists(default_path):
        create_default_excel(default_path)

    session["file_path"] = default_path
    session["original_name"] = "data.xlsx"
    return default_path, "data.xlsx"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded = request.files.get("file")

        if uploaded and uploaded.filename:
            original = secure_filename(uploaded.filename)
            if not allowed_file(original):
                return "Nur Excel-Dateien (.xlsx/.xls) sind erlaubt.", 400

            ext = os.path.splitext(original)[1].lower() or ".xlsx"
            unique_name = f"{uuid.uuid4().hex}{ext}"
            file_path = os.path.join(UPLOAD_FOLDER, unique_name)
            uploaded.save(file_path)

            session["file_path"] = file_path
            session["original_name"] = original

        return redirect(url_for("index"))

    _, active_name = get_active_file()
    return render_template_string(HTML, active_filename=active_name)


@app.route("/data")
def get_data():
    file_path, _ = get_active_file()
    df = pd.read_excel(file_path, dtype=str)
    return jsonify(df.fillna("").to_dict(orient="records"))


@app.route("/save", methods=["POST"])
def save():
    file_path, _ = get_active_file()
    data = request.get_json(silent=True)
    if not isinstance(data, list):
        return jsonify({"error": "Ungültige Daten"}), 400

    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)
    return jsonify({"status": "ok"})


@app.route("/download")
def download():
    file_path, original_name = get_active_file()
    if not os.path.exists(file_path):
        return "Keine Datei vorhanden", 404

    return send_file(
        file_path,
        as_attachment=True,
        download_name=original_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
