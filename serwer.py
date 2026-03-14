import os
import sys
import uuid
import threading
import pyodbc
from pathlib import Path
from flask import Flask, render_template, g, request, jsonify, send_file
from dotenv import load_dotenv

load_dotenv("A:/Antropic/db.env")

DB_CONN_STR = (
    f"DRIVER={{SQL Server}};SERVER={os.getenv('DB_SERVER')};"
    f"DATABASE={os.getenv('DB_NAME')};UID={os.getenv('DB_USER')};PWD={os.getenv('DB_PASSWORD')}"
)

# Ścieżka do modułu generuj_pdf
OFERTA_DIR = Path("A:/Antropic/pyOfertaPDF1/pyOfertaPDF1")
sys.path.insert(0, str(OFERTA_DIR))

from generuj_pdf import generuj_oferte, _load_openai_key

# Ładowanie klucza AI przy starcie
try:
    _load_openai_key("wwgg^&*^J)H*)(W^_)")
    print("[OK] Klucz OpenAI załadowany")
except Exception as e:
    print(f"[WARN] Tryb bez AI: {e}")

app = Flask(__name__)

# Słownik zadań: task_id -> {status, pdf_path, error}
tasks: dict[str, dict] = {}

# ================================================================
#  SŁOWNIKI SZABLONÓW / TŁA / JĘZYKÓW
# ================================================================

szablony_opcje = {
    "2x3 poziomo 6prod/str": "szablon5.html",
    "KP Karta produktu 2x3 poziomo 6prod/str": "szablon805.html",
    "KP Karta produktu 4pr/str, Opis AI skrócony": "szablon970.html",
    "KP Karta produktu 1pr/str 4 zdjęcia, opis pełny(AI skraca)": "szablon911.html",
    "KP Karta produktu 10pr/str, lista kompaktowa 1zdjęcie": "szablon930.html",
    "KP Karta produktu 12pr/str, lista kompaktowa 4zdjęcia": "szablon931.html",
    "5x4 gazetka+tło 20prod/str": "szablon940.html",
    "4x3 gazetka 12prod/str": "szablon4.html",
    "4x3 gazetka 12prod/str AI (opis uzupełniony do 250zn)": "szablon401.html",
    "1 produkt na stronę (pełny opis)": "szablon1.html",
    "1 produkt na stronę duże zdjęcie i 3małe+ (pełny opis)": "szablon11.html",
    "3 produkty na stronę (bez opisu)": "szablon2.html",
    "5 produktów na stronę 1 zdjęcie, opis krótki": "szablon6.html",
    "5 produktów na stronę 1 zdjęcie, opis długi AI 600znaków": "szablon7.html",
    "16 pr/str 1/zdj lista kompaktowa": "szablon3.html",
    "14 pr/str 4/zdj lista kompaktowa": "szablon32.html",
    "PROMO 10prod 1foto cena przed i po rabacie": "szablon350.html",
    "PROMO 10prod 1foto cena przed i po rabacie+wartość oferty": "szablon351.html",
}

tla_opcje_pion = {
    "PRO Technik": "templates/backgrounds/szablon_protechnik.jpg",
    "XL Tools": "templates/backgrounds/szablon_xltools.jpg",
    "XL Green": "templates/backgrounds/szablon_xlgreen.jpg",
    "XL Moto": "templates/backgrounds/szablon_xlmoto.jpg",
    "IRMA": "templates/backgrounds/szablon_irma.jpg",
    "MetalKraft": "templates/backgrounds/szablon_metalkraft.jpg",
    "Tornado": "templates/backgrounds/szablon_tornado.jpg",
    "Magneto": "templates/backgrounds/szablon_magneto.jpg",
}

tla_opcje_poziom = {
    "PRO Technik (poziomo)": "templates/backgrounds_horizontal/szablon_protechnik-poziomo.jpg",
    "IRMA (poziomo)": "templates/backgrounds_horizontal/szablon_IRMA-poziomo.jpg",
}

jezyki_opcje = [
    "polski", "angielski", "niemiecki", "francuski", "hiszpański", "włoski", "portugalski",
    "niderlandzki", "czeski", "słowacki", "węgierski", "rumuński", "bułgarski",
    "chorwacki", "grecki", "słoweński", "litewski", "łotewski", "estoński", "szwedzki",
    "duński", "norweski", "fiński", "islandzki", "ukraiński", "serbski", "macedoński",
    "czarnogórski", "bośniacki", "irlandzki", "luksemburski", "maltański",
]

# ================================================================
#  DB
# ================================================================

def get_db():
    if "db" not in g:
        g.db = pyodbc.connect(DB_CONN_STR)
    return g.db


@app.teardown_appcontext
def close_db(exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


# ================================================================
#  ROUTES
# ================================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generuj-oferte", methods=["GET"])
def generuj_oferte_get():
    tla_wszystkie = {**tla_opcje_pion, **tla_opcje_poziom}
    return render_template(
        "generuj_oferte.html",
        szablony=szablony_opcje,
        tla_pion=tla_opcje_pion,
        tla_poziom=tla_opcje_poziom,
        jezyki=jezyki_opcje,
    )


@app.route("/api/generuj", methods=["POST"])
def api_generuj():
    data = request.json
    numer = (data.get("numer") or "").strip()
    if not numer:
        return jsonify({"error": "Brak numeru oferty"}), 400

    szablon_key = data.get("szablon", list(szablony_opcje.keys())[0])
    plik_szablonu = szablony_opcje.get(szablon_key, "szablon5.html")
    jezyk = data.get("jezyk", "polski")
    sortuj_po_nazwie = data.get("sortuj") == "alfa"

    tlo_key = data.get("tlo", "").strip()
    plik_tla = tla_opcje_pion.get(tlo_key) or tla_opcje_poziom.get(tlo_key) or None
    rozszerz_ramki = bool(data.get("rozszerz_ramki", False))

    task_id = str(uuid.uuid4())
    tasks[task_id] = {"status": "running", "pdf_path": None, "error": None}

    def worker():
        prev_dir = os.getcwd()
        try:
            os.chdir(OFERTA_DIR)
            pdf_path = generuj_oferte(
                numer_oferty=numer,
                ile_zdjec=4,
                szablon=plik_szablonu,
                jezyk=jezyk,
                tlo=plik_tla,
                sortuj_po_nazwie=sortuj_po_nazwie,
                open_after=False,
                rozszerz_ramki=rozszerz_ramki,
            )
            # Absolutna ścieżka — przed przywróceniem katalogu!
            tasks[task_id]["pdf_path"] = str(Path(pdf_path).resolve())
            tasks[task_id]["status"] = "done"
        except Exception as e:
            tasks[task_id]["error"] = str(e)
            tasks[task_id]["status"] = "error"
        finally:
            os.chdir(prev_dir)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"task_id": task_id})


@app.route("/api/status/<task_id>")
def api_status(task_id):
    task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "Nieznane zadanie"}), 404
    return jsonify(task)


@app.route("/api/pobierz/<task_id>")
def api_pobierz(task_id):
    task = tasks.get(task_id)
    if not task or task["status"] != "done":
        return jsonify({"error": "PDF niedostępny"}), 404
    pdf_path = task["pdf_path"]
    return send_file(pdf_path, mimetype="application/pdf", as_attachment=False)


@app.route("/mailing")
def mailing():
    return render_template("placeholder.html", tytul="Mailing")


@app.route("/opisy-produktow")
def opisy_produktow():
    return render_template("placeholder.html", tytul="Opisy produktów")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
