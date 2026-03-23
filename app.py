#!/usr/bin/env python3
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from io import BytesIO
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
ALLOWED_EXTENSIONS = {".pdf"}
MAX_UPLOAD_SIZE = 25 * 1024 * 1024

SCRIPT_BY_TYPE = {
    "awb": BASE_DIR / "extract_awb.py",
    "srm": BASE_DIR / "extract_srm.py",
    "cdm": BASE_DIR / "extract_cdm.py",
}

LABEL_BY_TYPE = {
    "awb": "Attijari / AWB",
    "srm": "SRM",
    "cdm": "Crédit du Maroc / CDM",
}

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "bank-statement-extractor")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def output_name(filename: str) -> str:
    stem = Path(filename).stem or "releve"
    return f"{stem}_extrait.xlsx"


def clean_details(text: str) -> str:
    compact = " ".join(text.split())
    if len(compact) <= 320:
        return compact
    return compact[:317] + "..."


@app.route("/")
def index() -> str:
    return render_template("index.html", type_labels=LABEL_BY_TYPE)


@app.route("/logo-srm.jpeg")
def logo_srm():
    return send_file(BASE_DIR / "logo.jpeg", mimetype="image/jpeg")


@app.route("/extraire", methods=["POST"])
def extraire():
    selected_type = request.form.get("statement_type", "").strip().lower()
    uploaded_file = request.files.get("statement_file")

    if selected_type not in SCRIPT_BY_TYPE:
        flash("Veuillez choisir un type de relevé valide.", "erreur")
        return redirect(url_for("index"))

    if uploaded_file is None or uploaded_file.filename == "":
        flash("Veuillez déposer un fichier PDF avant de lancer l'extraction.", "erreur")
        return redirect(url_for("index"))

    if not allowed_file(uploaded_file.filename):
        flash("Seuls les fichiers PDF sont acceptés.", "erreur")
        return redirect(url_for("index"))

    workdir = Path(tempfile.mkdtemp(prefix="releves_"))
    try:
        source_name = secure_filename(uploaded_file.filename) or "releve.pdf"
        input_pdf = workdir / source_name
        output_xlsx = workdir / output_name(source_name)
        uploaded_file.save(input_pdf)

        command = [
            sys.executable,
            str(SCRIPT_BY_TYPE[selected_type]),
            str(input_pdf),
            "-o",
            str(output_xlsx),
        ]
        try:
            result = subprocess.run(
                command,
                cwd=BASE_DIR,
                capture_output=True,
                text=True,
            )
        except FileNotFoundError:
            flash(
                "Le serveur ne dispose pas de l'outil pdftotext. "
                "Installez Poppler sur l'hébergement pour activer l'extraction.",
                "erreur",
            )
            return redirect(url_for("index"))

        if result.returncode != 0:
            details = clean_details(
                result.stderr.strip() or result.stdout.strip() or "Erreur d'extraction inconnue."
            )
            flash(
                "L'extraction a échoué. Vérifiez le type choisi et le format du relevé. "
                f"Détail technique : {details}",
                "erreur",
            )
            return redirect(url_for("index"))

        if not output_xlsx.exists():
            flash("Le fichier Excel n'a pas été généré.", "erreur")
            return redirect(url_for("index"))

        file_bytes = output_xlsx.read_bytes()
        return send_file(
            BytesIO(file_bytes),
            as_attachment=True,
            download_name=output_xlsx.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.route("/sante")
def sante() -> dict[str, str]:
    return {"statut": "ok"}


@app.errorhandler(413)
def fichier_trop_lourd(_error):
    flash("Le fichier est trop volumineux. La taille maximale autorisée est de 25 Mo.", "erreur")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
