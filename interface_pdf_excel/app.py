from flask import Flask, render_template, request, Response, redirect, url_for

import os
import shutil
from werkzeug.utils import secure_filename
import threading
import queue
import time

from scripts.traitement_pdf import process_all_pdfs_in_folder_with_logs_sequentiel
from scripts.traitement_pdf_parallel import process_all_pdfs_in_folder_with_logs_parallel
from scripts.enrichissement_excel import enrichir_resume_global

app = Flask(__name__)
UPLOAD_FOLDER = "temp_upload"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 🎯 File de log partagée
log_queue = queue.Queue()

def log_streamer():
    """Générateur SSE pour le flux de logs."""
    buffer = ""
    while True:
        try:
            message = log_queue.get(timeout=1)
            buffer += message
            if buffer.endswith("\n") or buffer.endswith("\r"):
                yield f"data: {buffer.strip()}\n\n"
                buffer = ""
        except queue.Empty:
            continue


def make_logger():
    def log(text):
        print(text)
        if not text.endswith("\n"):
            text += "\n"
        log_queue.put(text)
    return log


@app.route("/", methods=["GET"])
def index():
    msg = request.args.get("msg")
    message = {
        "sequentiel": "⏳ Traitement séquentiel en cours...",
        "parallel": "⏳ Traitement parallèle en cours...",
        "enrich": "⏳ Enrichissement en cours..."
    }.get(msg, "")
    return render_template("index.html", message=message)


@app.route("/stream_logs")
def stream_logs():
    return Response(log_streamer(), mimetype="text/event-stream")

@app.route("/process_sequentiel", methods=["POST"])
def process_sequentiel_route():
    files = request.files.getlist("pdf_folder")
    if not files:
        return render_template("index.html", message="❌ Aucun fichier PDF sélectionné.")

    folder_path = os.path.join(UPLOAD_FOLDER, "seq")
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
    os.makedirs(folder_path)

    for f in files:
        filename = secure_filename(f.filename)
        f.save(os.path.join(folder_path, filename))

    def run():
        logger = make_logger()
        try:
            process_all_pdfs_in_folder_with_logs_sequentiel(folder_path, log_func=logger)
            logger("✅ Traitement séquentiel terminé avec succès.")
        except Exception as e:
            logger(f"❌ Erreur : {e}")

    threading.Thread(target=run).start()
    return render_template("index.html", message="⏳ Traitement séquentiel en cours...")



@app.route("/process_parallel", methods=["POST"])
def process_parallel_route():
    files = request.files.getlist("pdf_folder")
    if not files:
        return render_template("index.html", message="❌ Aucun fichier PDF sélectionné.")

    folder_path = os.path.join(UPLOAD_FOLDER, "par")
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
    os.makedirs(folder_path)

    for f in files:
        filename = secure_filename(f.filename)
        f.save(os.path.join(folder_path, filename))

    def run():
        logger = make_logger()
        try:
            process_all_pdfs_in_folder_with_logs_parallel(folder_path, log_func=logger)
            logger("✅ Traitement parallèle terminé avec succès.")
        except Exception as e:
            logger(f"❌ Erreur : {e}")

    threading.Thread(target=run).start()
    return render_template("index.html", message="⏳ Traitement séquentiel en cours...")



from flask import send_from_directory


@app.route("/enrichir_excel", methods=["POST"])
def enrichir_excel_route():
    file = request.files.get("global_excel")
    if not file:
        return render_template("index.html", message="❌ Aucun fichier Excel sélectionné.")

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # Définir le nom du fichier de sortie
    result_filename = "résultat_complet.xlsx"
    result_filepath = os.path.join(UPLOAD_FOLDER, result_filename)

    def run():
        logger = make_logger()
        try:
            enrichir_resume_global(filepath, log_func=logger)
            logger(f"✅ Enrichissement terminé avec succès.")
            logger(f"[Télécharger le fichier enrichi](/download_excel/{result_filename})")
        except Exception as e:
            logger(f"❌ Erreur enrichissement : {e}")

    threading.Thread(target=run).start()

    return render_template("index.html", message="⏳ Enrichissement en cours...")


@app.route("/enrichir_global_excel", methods=["POST"])
def enrichir_global_route():
    file = request.files.get("global_excel")
    if not file:
        return render_template("index.html", message="❌ Aucun fichier Excel sélectionné.")

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    result_filename = "résultat_complet.xlsx"

    def run():
        logger = make_logger()
        try:
            from scripts.enrichissement_excel import enrichir_global
            enrichir_global(filepath, log_func=logger)
            logger("✅ Enrichissement (global) terminé.")
            logger(f"[Télécharger résultat global](/download_excel/{result_filename})")
        except Exception as e:
            logger(f"❌ Erreur enrichissement : {e}")

    threading.Thread(target=run).start()
    return render_template("index.html", message="⏳ Enrichissement global en cours...")

@app.route("/enrichir_par_annee_excel", methods=["POST"])
def enrichir_par_annee_route():
    file = request.files.get("global_excel")
    if not file:
        return render_template("index.html", message="❌ Aucun fichier Excel sélectionné.")

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    def run():
        logger = make_logger()
        try:
            from scripts.enrichissement_excel import enrichir_par_annee
            enrichir_par_annee(filepath, log_func=logger)
            logger("✅ Enrichissement (par année) terminé.")
        except Exception as e:
            logger(f"❌ Erreur enrichissement : {e}")

    threading.Thread(target=run).start()
    return render_template("index.html", message="⏳ Enrichissement par année en cours...")


@app.route("/download_excel/<path:filename>")
def download_excel(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)



if __name__ == "__main__":
    app.run(debug=True, threaded=True)

