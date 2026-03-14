import os
import sys
import json
import uuid
import platform
import threading
import pyodbc
from pathlib import Path
from flask import Flask, render_template, g, request, jsonify, send_file, abort
from dotenv import load_dotenv

load_dotenv("A:/Antropic/db.env")

DB_CONN_STR = (
    f"DRIVER={{SQL Server}};SERVER={os.getenv('DB_SERVER')};"
    f"DATABASE={os.getenv('DB_NAME')};UID={os.getenv('DB_USER')};PWD={os.getenv('DB_PASSWORD')}"
)

# Ścieżka do modułu generuj_pdf
OFERTA_DIR = Path("A:/Antropic/pyOfertaPDF1/pyOfertaPDF1")
OFERTY_ROOT = Path("A:/Antropic/gotowe oferty")
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
    sortuj = data.get("sortuj", "dokument")
    sortuj_po_nazwie = sortuj == "alfa"

    tlo_key = data.get("tlo", "").strip()
    plik_tla = tla_opcje_pion.get(tlo_key) or tla_opcje_poziom.get(tlo_key) or None
    rozszerz_ramki = bool(data.get("rozszerz_ramki", False))

    meta = {
        "szablon_nazwa": szablon_key,
        "tlo_nazwa": tlo_key,
    }

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
                meta=meta,
            )
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
    return send_file(task["pdf_path"], mimetype="application/pdf", as_attachment=False)


# ================================================================
#  MOJE OFERTY
# ================================================================

def _scan_oferty() -> list[dict]:
    """Skanuje OFERTY_ROOT i zwraca listę ofert z metadanymi (JSON lub parsowanie nazwy)."""
    oferty = []
    if not OFERTY_ROOT.exists():
        return oferty
    for pdf in OFERTY_ROOT.rglob("*.pdf"):
        host = pdf.parent.name
        sidecar = pdf.with_suffix(".json")
        if sidecar.exists():
            try:
                m = json.loads(sidecar.read_text(encoding="utf-8"))
            except Exception:
                m = {}
        else:
            m = _parse_filename(pdf.stem)
        m["_filename"] = pdf.name
        m["_host"] = host
        m["_size_kb"] = round(pdf.stat().st_size / 1024)
        # timestamp do sortowania
        ts = m.get("wygenerowano", "")
        if not ts:
            ts = m.get("_ts", "")
        m["_ts_sort"] = ts
        oferty.append(m)
    oferty.sort(key=lambda x: x["_ts_sort"], reverse=True)
    return oferty


def _parse_filename(stem: str) -> dict:
    """Parsuje nazwę pliku PDF gdy brak sidecara.
    Format: {safe_num}_{jezyk}_{safe_tpl}_{YYYYMMDD}_{HHMM}
    Przykład: PRO 1_TM_2025_polski_szablon5_20260314_1741
    """
    from datetime import datetime as _dt
    parts = stem.rsplit("_", 4)  # max 4 cięcia od prawej → 5 części
    ts = ""
    jezyk = ""
    szablon = ""
    numer = stem
    if len(parts) == 5:
        numer_raw, jezyk, szablon, date_part, time_part = parts
        numer = numer_raw.replace("_", "/")
        try:
            ts = _dt.strptime(date_part + time_part, "%Y%m%d%H%M").isoformat(timespec="seconds")
        except Exception:
            pass
    return {
        "numer": numer,
        "szablon_nazwa": szablon,
        "jezyk": jezyk,
        "tlo_nazwa": "",
        "ilosc_produktow": None,
        "wygenerowano": ts,
        "_ts": ts,
    }


@app.route("/moje-oferty")
def moje_oferty():
    host = platform.node()
    oferty = _scan_oferty()
    return render_template("moje_oferty.html", oferty=oferty, host=host)


@app.route("/api/oferty")
def api_oferty():
    return jsonify(_scan_oferty())


@app.route("/oferty/pdf/<host>/<filename>")
def oferty_pdf(host, filename):
    # Zapobiega path traversal
    if ".." in host or ".." in filename:
        abort(400)
    pdf_path = OFERTY_ROOT / host / filename
    if not pdf_path.is_file():
        abort(404)
    return send_file(str(pdf_path), mimetype="application/pdf", as_attachment=False)


# ================================================================
#  PLACEHOLDERY
# ================================================================

@app.route("/mailing")
def mailing():
    return render_template("placeholder.html", tytul="Mailing")


@app.route("/opisy-produktow")
def opisy_produktow():
    return render_template("placeholder.html", tytul="Opisy produktów")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
