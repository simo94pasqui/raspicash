"""RaspiCash - gestione delle spese di casa (Flask + SQLite)."""
import fcntl
import io
import json
import os
import re
import secrets
import sqlite3
import tempfile
from contextlib import closing, contextmanager
from datetime import date, datetime, timedelta
from decimal import Decimal
from functools import wraps
from glob import glob

from flask import (Flask, abort, flash, g, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from openpyxl import Workbook
from openpyxl.styles import Font
from werkzeug.security import check_password_hash, generate_password_hash

import importer
import receipt
from importer import calc_fuel

VERSION = "1.0.0"
SCHEMA_VERSION = 1

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
DB = os.path.join(DATA, "raspicash.db")
KEYFILE = os.path.join(DATA, "secret.key")
IMP_DIR = os.path.join(DATA, "import_tmp")      # File Excel caricati e non ancora confermati
MAX_IMPORT = 5 * 1024 * 1024
MAX_SCAN = 15 * 1024 * 1024
os.makedirs(DATA, exist_ok=True)


@contextmanager
def startup_lock():
    """Gunicorn avvia più worker insieme: chiave e database li prepara uno alla volta."""
    with open(os.path.join(DATA, ".startup.lock"), "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        yield


def load_secret_key():
    """Chiave di sessione: al primo avvio viene creata, leggibile solo dal proprietario."""
    with startup_lock():
        key = ""
        if os.path.exists(KEYFILE):
            with open(KEYFILE) as f:
                key = f.read().strip()
        if len(key) < 32:
            key = secrets.token_hex(32)
            fd = os.open(KEYFILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
            with os.fdopen(fd, "w") as f:
                f.write(key)
        return key


app = Flask(__name__)
app.secret_key = load_secret_key()
app.permanent_session_lifetime = timedelta(days=365)
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

MONTHS = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio",
          "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
PALETTE = ["#c2185b", "#1e88e5", "#43a047", "#fb8c00", "#8e24aa", "#00acc1",
           "#6d4c41", "#e53935", "#7cb342", "#3949ab", "#fdd835", "#546e7a"]

SCHEMA = """
create table if not exists users(
  id integer primary key, name text unique not null, pw_hash text not null, is_admin integer default 0);
create table if not exists categories(
  id integer primary key, name text unique not null, archived integer default 0,
  is_fuel integer not null default 0, is_car integer not null default 0, is_trip integer not null default 0);
create table if not exists expenses(
  id integer primary key, date text not null, description text not null,
  category_id integer not null references categories(id),
  amount integer not null default 0, voucher integer not null default 0,
  private integer not null default 0, owner_id integer not null references users(id),
  created_at text, updated_by integer references users(id), updated_at text);
create index if not exists ix_exp_date on expenses(date);
create table if not exists cars(
  id integer primary key, name text unique not null, archived integer default 0);
create table if not exists user_car(user_id integer primary key, car_id integer);
create table if not exists fuel_logs(
  expense_id integer primary key, car_id integer not null references cars(id),
  fill_no integer, odometer integer, liters integer, price_l integer,
  trip_km integer, claimed_km integer);
create table if not exists car_expenses(
  expense_id integer primary key references expenses(id), car_id integer not null references cars(id));
create table if not exists people(
  id integer primary key, name text unique not null, archived integer default 0);
create table if not exists loans(
  expense_id integer primary key references expenses(id), person_id integer not null references people(id),
  amount integer not null);
create table if not exists trips(
  id integer primary key, name text unique not null, is_work integer not null default 0, archived integer default 0);
create table if not exists trip_expenses(
  expense_id integer primary key references expenses(id), trip_id integer not null references trips(id),
  company integer not null default 0);
"""


def init_db():
    with startup_lock(), closing(sqlite3.connect(DB)) as c:
        if c.execute("pragma user_version").fetchone()[0] > SCHEMA_VERSION:
            raise RuntimeError("Il database è di una versione più recente di RaspiCash")
        c.executescript(SCHEMA)
        for name, flag in (("Benzina", "is_fuel"), ("Auto", "is_car"), ("Viaggi", "is_trip")):
            if not c.execute(f"select 1 from categories where {flag}=1").fetchone():
                c.execute(f"insert into categories(name,{flag}) values(?,1)", (name,))
        if not c.execute("select 1 from cars").fetchone():
            c.execute("insert into cars(name) values('La mia auto')")
        c.execute(f"pragma user_version={SCHEMA_VERSION}")
        c.commit()


init_db()


def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_):
    d = g.pop("db", None)
    if d:
        d.close()


# --- Utilità ---
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def to_cents(s):
    s = (s or "").strip().replace("€", "").replace(" ", "")
    if not s:
        return 0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return int((Decimal(s) * 100).to_integral_value())


def fmt_in(c):
    return "" if not c else f"{c / 100:.2f}".replace(".", ",")


def to_num(s, scale=1):
    """Numero opzionale come intero scalato (None se vuoto). Accetta virgola o punto."""
    s = (s or "").strip().replace(" ", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return int((Decimal(s) * scale).to_integral_value())


def fmt_dec(v, scale, maxdec):
    if not v:
        return ""
    t = f"{v / scale:.{maxdec}f}".rstrip("0").rstrip(".")
    return t.replace(".", ",")


@app.template_filter("dec")
def dec(v, n=2):
    return f"{v:,.{n}f}".replace(",", "X").replace(".", ",").replace("X", ".")


@app.template_filter("eur")
def eur(c):
    return f"{c / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@app.context_processor
def ctx():
    return dict(this_year=date.today().year, months=MONTHS, version=VERSION)


def years():
    ys = {r[0] for r in db().execute("select distinct substr(date,1,4) from expenses")}
    ys.add(str(date.today().year))
    return sorted(ys, reverse=True)


# --- Accesso ---
@app.before_request
def gate():
    g.user = None
    if session.get("uid"):
        g.user = db().execute("select * from users where id=?", (session["uid"],)).fetchone()
    if request.endpoint in ("static", "login", "setup", "restore_setup", "manifest"):
        return
    if db().execute("select count(*) from users").fetchone()[0] == 0:
        return redirect(url_for("setup"))
    if not g.user:
        return redirect(url_for("login"))


def admin_required(f):
    @wraps(f)
    def w(*a, **k):
        if not g.user["is_admin"]:
            abort(403)
        return f(*a, **k)
    return w


@app.route("/setup", methods=["GET", "POST"])
def setup():
    if db().execute("select count(*) from users").fetchone()[0]:
        return redirect(url_for("login"))
    if request.method == "POST" and request.form["name"].strip() and len(request.form["password"]) >= 4:
        db().execute("insert into users(name,pw_hash,is_admin) values(?,?,1)",
                     (request.form["name"].strip(), generate_password_hash(request.form["password"])))
        db().commit()
        flash("Utente creato: ora accedi")
        return redirect(url_for("login"))
    return render_template("login.html", setup=True)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = db().execute("select * from users where lower(name)=lower(?)",
                         (request.form["name"].strip(),)).fetchone()
        if u and check_password_hash(u["pw_hash"], request.form["password"]):
            session.clear()
            session.permanent = bool(request.form.get("remember"))
            session["uid"] = u["id"]
            return redirect(url_for("index"))
        flash("Nome o password errati")
    return render_template("login.html", setup=False)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/manifest.json")
def manifest():
    return jsonify(name="RaspiCash", short_name="RaspiCash", start_url="/", display="standalone",
                   background_color="#faf6f7", theme_color="#c2185b",
                   icons=[{"src": "/static/icon.svg", "sizes": "any", "type": "image/svg+xml"}])


# --- Spese ---
BASE_Q = """select e.*, c.name cat, o.name owner, u.name updater, pe.name person, ln.person_id person_id,
  ln.amount loan_part,
  tr.name trip, te.trip_id trip_id, coalesce(te.company,0) company
  from expenses e
  join categories c on c.id=e.category_id join users o on o.id=e.owner_id
  left join users u on u.id=e.updated_by
  left join loans ln on ln.expense_id=e.id left join people pe on pe.id=ln.person_id
  left join trip_expenses te on te.expense_id=e.id left join trips tr on tr.id=te.trip_id"""


def net_total(r):
    """Importo che conta nei conti: il totale meno la parte prestata ad altri
    (se ha pagato l'altra persona la spesa resta intera).
    """
    return r["amount"] + r["voucher"] - max(r["loan_part"] or 0, 0)


# Parte prestata (richiede un join su loans con alias ln): si toglie prima dall'importo, poi dal buono
LENT = "(case when ln.amount>0 then ln.amount else 0 end)"
LENT_A = f"min({LENT}, max(e.amount,0))"


def active_cats(extra_id=None):
    return db().execute("select * from categories where archived=0 or id=? order by lower(name)",
                        (extra_id,)).fetchall()


def parse_expense():
    f = request.form
    try:
        d = date.fromisoformat(f["date"]).isoformat()
        desc = f["description"].strip()
        cat = int(f["category_id"])
        a, v = to_cents(f.get("amount")), to_cents(f.get("voucher"))
    except (ValueError, KeyError, ArithmeticError):
        return None
    if not db().execute("select 1 from categories where id=?", (cat,)).fetchone():
        return None
    if not desc and is_fuel(cat):
        desc = "Rifornimento"
    if not desc:
        return None
    if is_fuel(cat) or is_car(cat):   # Benzina e auto non sono mai private
        return d, desc, cat, a, v, 0
    return d, desc, cat, a, v, 1 if f.get("private") else 0


def get_expense(id):
    e = db().execute(BASE_Q + " where e.id=? and (e.private=0 or e.owner_id=?)",
                     (id, g.user["id"])).fetchone()
    if not e:
        abort(404)
    return e


@app.route("/")
def index():
    e = dict(date=request.args.get("d") or date.today().isoformat(), description="",
             category_id=request.args.get("cat", type=int) or 0, amount="", voucher="", private=0)
    recent = db().execute(BASE_Q + " where e.private=0 or e.owner_id=? order by e.id desc limit 5",
                          (g.user["id"],)).fetchall()
    return render_template("index.html", e=e, cats=active_cats(), recent=recent, **fuel_ctx())


@app.post("/add")
def add():
    p = parse_expense()
    if not p:
        flash("Dati non validi: controlla descrizione e importi")
        return redirect(url_for("index"))
    d, desc, cat, a, v, priv = p
    fuel = None
    if is_fuel(cat):
        r = parse_fuel(a, v)
        if not r:
            flash("Rifornimento: scegli l'auto e indica almeno l'importo o i litri")
            return redirect(url_for("index", cat=cat, d=d))
        fuel, a = r
    car = None
    if is_car(cat):
        car = parse_car()
        if not car:
            flash("Spesa auto: scegli l'auto")
            return redirect(url_for("index", cat=cat, d=d))
    loan = None
    if request.form.get("subloan") and not (is_fuel(cat) or is_car(cat)):
        loan = parse_loan(a + v)
        if not loan:
            flash("Prestiti: scegli la persona e un importo del prestito maggiore di zero, non oltre il totale")
            return redirect(url_for("index", cat=cat, d=d))
    trip = None
    if is_trip(cat):
        trip = parse_trip()
        if not trip:
            flash("Viaggio: scegli il viaggio (o creane uno nuovo)")
            return redirect(url_for("index", cat=cat, d=d))
    cur = db().execute("insert into expenses(date,description,category_id,amount,voucher,private,owner_id,created_at)"
                       " values(?,?,?,?,?,?,?,?)", (d, desc, cat, a, v, priv, g.user["id"], now()))
    if fuel:
        save_fuel(cur.lastrowid, fuel)
    if car:
        save_car_link(cur.lastrowid, car)
    if loan:
        save_loan(cur.lastrowid, loan)
    if trip:
        save_trip(cur.lastrowid, trip)
    db().commit()
    flash("Spesa salvata ✓")
    return redirect(url_for("index", d=d))


def _suggest_category(store):
    """Categoria più usata per le spese con quel negozio nella descrizione (escluse le speciali)."""
    key = re.sub(r"[^\w ]", "", store).strip().lower()
    if len(key) < 3:
        return None
    r = db().execute("select e.category_id, count(*) n from expenses e join categories c on c.id=e.category_id "
                     "where (e.private=0 or e.owner_id=?) and c.archived=0 and c.is_fuel=0 and c.is_car=0 "
                     "and c.is_trip=0 and lower(e.description) like ? group by e.category_id order by n desc limit 1",
                     (g.user["id"], "%" + key + "%")).fetchone()
    return r[0] if r else None


@app.post("/scan")
def scan():
    """Legge la foto di uno scontrino e propone importo, negozio, data e categoria."""
    f = request.files.get("image")
    if not f:
        return jsonify(error="Nessuna immagine ricevuta"), 400
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "scontrino.jpg")
        f.save(path)
        if os.path.getsize(path) > MAX_SCAN:
            return jsonify(error="Immagine troppo grande"), 413
        try:
            text = receipt.ocr_image(path)
        except receipt.OcrError as e:
            return jsonify(error=str(e)), 503
    r = receipt.parse(text)
    total, voucher = r["total"], r["voucher"]
    return jsonify(amount=receipt.cents(total - voucher) if total else "",
                   voucher=receipt.cents(voucher) if voucher else "",
                   store=r["store"] or "", date=r["date"] or "", note=r["note"],
                   category_id=_suggest_category(r["store"]) if r["store"] else None, text=text[:4000])


@app.route("/history")
def history():
    year = request.args.get("year", type=int) or date.today().year
    month = request.args.get("month", type=int) or 0
    cat = request.args.get("cat", type=int) or 0
    show = request.args.get("show", "shared")
    q = BASE_Q + " where substr(e.date,1,4)=? and (e.private=0 or e.owner_id=?)"
    args = [str(year), g.user["id"]]
    if month:
        q += " and cast(substr(e.date,6,2) as int)=?"
        args.append(month)
    if cat:
        q += " and e.category_id=?"
        args.append(cat)
    if show == "shared":
        q += " and e.private=0"
    elif show == "mine":
        q += " and e.private=1"
    rows = db().execute(q + " order by e.date desc, e.id desc", args).fetchall()
    tot = sum(net_total(r) for r in rows if not r["company"])
    ncomp = sum(1 for r in rows if r["company"])
    allcats = db().execute("select * from categories order by lower(name)").fetchall()
    return render_template("history.html", rows=rows, tot=tot, ncomp=ncomp, year=year, month=month, cat=cat,
                           show=show, years=years(), allcats=allcats)


@app.route("/edit/<int:id>", methods=["GET", "POST"])
def edit(id):
    row = get_expense(id)
    if request.method == "POST":
        p = parse_expense()
        if p:
            d, desc, cat, a, v, priv = p
            if row["owner_id"] != g.user["id"]:
                priv = row["private"]  # La privacy la cambia solo il proprietario
            if is_fuel(cat) or is_car(cat):
                priv = 0
            fuel = car = loan = trip = None
            ok = True
            if is_fuel(cat):
                r = parse_fuel(a, v)
                if r:
                    fuel, a = r
                else:
                    ok = False
                    flash("Rifornimento: scegli l'auto e indica almeno l'importo o i litri")
            elif is_car(cat):
                car = parse_car()
                if not car:
                    ok = False
                    flash("Spesa auto: scegli l'auto")
            elif is_trip(cat):
                trip = parse_trip()
                if not trip:
                    ok = False
                    flash("Viaggio: scegli il viaggio (o creane uno nuovo)")
            if ok and request.form.get("subloan") and not (is_fuel(cat) or is_car(cat)):
                loan = parse_loan(a + v)
                if not loan:
                    ok = False
                    flash("Prestiti: scegli la persona e un importo del prestito maggiore di zero, non oltre il totale")
            if ok:
                db().execute("update expenses set date=?,description=?,category_id=?,amount=?,voucher=?,"
                             "private=?,updated_by=?,updated_at=? where id=?",
                             (d, desc, cat, a, v, priv, g.user["id"], now(), id))
                if fuel:
                    save_fuel(id, fuel)
                else:
                    db().execute("delete from fuel_logs where expense_id=?", (id,))
                if car:
                    save_car_link(id, car)
                else:
                    db().execute("delete from car_expenses where expense_id=?", (id,))
                if loan:
                    save_loan(id, loan)
                else:
                    db().execute("delete from loans where expense_id=?", (id,))
                if trip:
                    save_trip(id, trip)
                else:
                    db().execute("delete from trip_expenses where expense_id=?", (id,))
                db().commit()
                flash("Modifiche salvate ✓")
                return redirect(url_for("history", year=d[:4]))
        else:
            flash("Dati non validi")
    e = dict(row)
    e["amount"], e["voucher"] = fmt_in(row["amount"]), fmt_in(row["voucher"])
    ln = tp = None
    if row["trip_id"]:
        tp = dict(trip_id=row["trip_id"], paid_by="company" if row["company"] else "me")
    if row["person_id"]:
        ln = dict(person_id=row["person_id"], dir="given" if row["loan_part"] >= 0 else "received",
                  amount=fmt_in(abs(row["loan_part"])), sub=True)
    fl = db().execute("select * from fuel_logs where expense_id=?", (id,)).fetchone()
    f = None
    ce = db().execute("select car_id from car_expenses where expense_id=?", (id,)).fetchone()
    if ce:
        f = dict(car_id=ce["car_id"], fill_no="", odometer="", liters="", price="", trip_km="", claimed_km="")
    if fl:
        f = dict(car_id=fl["car_id"], fill_no=fl["fill_no"] or "", odometer=fl["odometer"] or "",
                 liters=fmt_dec(fl["liters"], 1000, 3), price=fmt_dec(fl["price_l"], 1000, 3),
                 trip_km=fl["trip_km"] or "", claimed_km=fl["claimed_km"] or "")
    return render_template("edit.html", e=e, cats=active_cats(row["category_id"]), **fuel_ctx(f, ln, tp))


@app.post("/delete/<int:id>")
def delete(id):
    row = get_expense(id)
    db().execute("delete from fuel_logs where expense_id=?", (row["id"],))
    db().execute("delete from car_expenses where expense_id=?", (row["id"],))
    db().execute("delete from loans where expense_id=?", (row["id"],))
    db().execute("delete from trip_expenses where expense_id=?", (row["id"],))
    db().execute("delete from expenses where id=?", (row["id"],))
    db().commit()
    flash("Spesa eliminata")
    return redirect(url_for("history", year=row["date"][:4]))


# --- Categorie ---
@app.route("/categories")
def categories():
    cats = db().execute("select c.*, (select count(*) from expenses where category_id=c.id) n "
                        "from categories c order by lower(name)").fetchall()
    return render_template("categories.html", cats=cats)


def _cat_write(sql, args, ok):
    try:
        db().execute(sql, args)
        db().commit()
        flash(ok)
    except sqlite3.IntegrityError:
        flash("Esiste già una categoria con questo nome")
    return redirect(url_for("categories"))


@app.post("/categories/add")
def cat_add():
    n = request.form["name"].strip()
    return _cat_write("insert into categories(name) values(?)", (n,), "Categoria aggiunta") if n else redirect(url_for("categories"))


@app.post("/categories/<int:id>/rename")
def cat_rename(id):
    n = request.form["name"].strip()
    return _cat_write("update categories set name=? where id=?", (n, id), "Categoria rinominata") if n else redirect(url_for("categories"))


@app.post("/categories/<int:id>/delete")
def cat_delete(id):
    if db().execute("select 1 from categories where id=? and (is_fuel=1 or is_car=1 or is_trip=1)", (id,)).fetchone():
        flash("Le categorie speciali (benzina, auto, viaggi) non si possono eliminare")
        return redirect(url_for("categories"))
    if db().execute("select count(*) from expenses where category_id=?", (id,)).fetchone()[0]:
        return _cat_write("update categories set archived=1 where id=?", (id,), "Categoria archiviata (è usata nello storico)")
    return _cat_write("delete from categories where id=?", (id,), "Categoria eliminata")


@app.post("/categories/<int:id>/restore")
def cat_restore(id):
    return _cat_write("update categories set archived=0 where id=?", (id,), "Categoria ripristinata")


# --- Auto e consumi ---
def is_fuel(cat_id):
    return bool(db().execute("select 1 from categories where id=? and is_fuel=1", (cat_id,)).fetchone())


def fuel_cat_ids():
    return sorted(r[0] for r in db().execute("select id from categories where is_fuel=1"))


def is_car(cat_id):
    return bool(db().execute("select 1 from categories where id=? and is_car=1", (cat_id,)).fetchone())


def car_cat_ids():
    return sorted(r[0] for r in db().execute("select id from categories where is_car=1"))


def parse_car():
    """Auto scelta nel modulo; None se manca o non esiste."""
    try:
        car = int(request.form["car_id"])
    except (ValueError, KeyError):
        return None
    return car if db().execute("select 1 from cars where id=?", (car,)).fetchone() else None


def save_car_link(eid, car_id):
    db().execute("insert or replace into car_expenses(expense_id,car_id) values(?,?)", (eid, car_id))


def car_expense_rows(uid, car_id, year=None):
    """Spese generiche dell'auto (assicurazione, tagliando...) visibili all'utente."""
    q = (BASE_Q + " join car_expenses x on x.expense_id=e.id"
         " where x.car_id=? and (e.private=0 or e.owner_id=?)")
    args = [car_id, uid]
    if year:
        q += " and substr(e.date,1,4)=?"
        args.append(str(year))
    return db().execute(q + " order by e.date desc, e.id desc", args).fetchall()


def active_cars(extra_id=None):
    return db().execute("select * from cars where archived=0 or id=? order by lower(name)", (extra_id,)).fetchall()


def current_car():
    """Auto in uso dall'utente; se non ne ha scelta una (o è archiviata) la prima attiva."""
    r = db().execute("select c.* from user_car u join cars c on c.id=u.car_id "
                     "where u.user_id=? and c.archived=0", (g.user["id"],)).fetchone()
    return (r or db().execute("select * from cars where archived=0 order by id limit 1").fetchone()
            or db().execute("select * from cars order by id limit 1").fetchone())


def next_numbers():
    return {r["car_id"]: (r["m"] or r["n"]) + 1 for r in db().execute(
        "select car_id, max(fill_no) m, count(*) n from fuel_logs group by car_id")}


def fuel_ctx(f=None, ln=None, tp=None):
    """Dati di rifornimento, prestito e viaggio per il modulo di inserimento e modifica."""
    cur = current_car()
    nexts = next_numbers()
    if f is None:
        f = dict(car_id=cur["id"], fill_no="", odometer="", liters="", price="", trip_km="", claimed_km="")
    if not f.get("fill_no"):
        f["fill_no"] = nexts.get(f["car_id"], 1)
    if ln is None:
        ln = dict(person_id=0, dir="given", amount="", sub=False)
    if tp is None:
        tp = dict(trip_id=request.args.get("trip", type=int) or 0, paid_by="me")
    return dict(f=f, cars=active_cars(f["car_id"]), nexts=nexts, fuel_ids=fuel_cat_ids(), car_ids=car_cat_ids(),
                ln=ln, people=active_people(ln["person_id"]),
                tp=tp, trips=active_trips(tp["trip_id"]), trip_ids=trip_cat_ids())


def parse_fuel(a, v):
    """Campi del rifornimento dal modulo. Ritorna (dati, importo in centesimi) oppure None se non validi.
    Serve almeno l'importo o i litri.
    """
    fm = request.form
    try:
        car = int(fm["car_id"])
        liters = to_num(fm.get("liters"), 1000)
        price = to_num(fm.get("price_l"), 1000)
        no, odo = to_num(fm.get("fill_no")), to_num(fm.get("odometer"))
        trip, claimed = to_num(fm.get("trip_km")), to_num(fm.get("claimed_km"))
    except (ValueError, KeyError, ArithmeticError):
        return None
    if not db().execute("select 1 from cars where id=?", (car,)).fetchone():
        return None
    if any(x is not None and x < 0 for x in (liters, price, no, odo, trip, claimed)):
        return None
    r = calc_fuel(a + v, liters, price)
    if not r:
        return None
    total, liters, price = r
    a = total - v
    if no is None:
        no = next_numbers().get(car, 1)
    return dict(car=car, no=no, odo=odo or None, liters=liters, price=price or None,
                trip=trip or None, claimed=claimed or None), a


def save_fuel(eid, x):
    db().execute("insert or replace into fuel_logs(expense_id,car_id,fill_no,odometer,liters,price_l,trip_km,claimed_km)"
                 " values(?,?,?,?,?,?,?,?)",
                 (eid, x["car"], x["no"], x["odo"], x["liters"], x["price"], x["trip"], x["claimed"]))


def scope_where(scope, uid):
    if scope == "shared":
        return "e.private=0", ()
    if scope == "mine":
        return "e.private=1 and e.owner_id=?", (uid,)
    return "(e.private=0 or e.owner_id=?)", (uid,)


def fuel_data(uid, car_id=None, year=None, scope="all"):
    """Rifornimenti visibili all'utente con km/L e L/100 km.
    Se mancano i km del pieno precedente usa la differenza dei km totali rispetto
    al rifornimento precedente della stessa auto.
    """
    auto = {}
    seq = db().execute("select expense_id, car_id, odometer from fuel_logs where odometer is not null "
                       "order by car_id, odometer, expense_id").fetchall()
    for a, b in zip(seq, seq[1:]):
        if a["car_id"] == b["car_id"] and b["odometer"] > a["odometer"]:
            auto[b["expense_id"]] = b["odometer"] - a["odometer"]
    w, p = scope_where(scope, uid)
    q = ("select f.*, e.date, e.description, e.amount, e.voucher, e.private, o.name owner, k.name car "
         "from fuel_logs f join expenses e on e.id=f.expense_id join users o on o.id=e.owner_id "
         "join cars k on k.id=f.car_id where " + w)
    args = list(p)
    if car_id:
        q += " and f.car_id=?"
        args.append(car_id)
    if year:
        q += " and substr(e.date,1,4)=?"
        args.append(str(year))
    rows = []
    for r in db().execute(q + " order by e.date desc, e.id desc", args):
        x = dict(r)
        L = r["liters"] / 1000 if r["liters"] else None
        total = r["amount"] + r["voucher"]
        trip = r["trip_km"] or auto.get(r["expense_id"])
        x.update(liters_l=L, total=total, trip=trip, trip_auto=bool(trip and not r["trip_km"]),
                 price=r["price_l"] / 1000 if r["price_l"] else (total / 100 / L if total and L else None),
                 kml=trip / L if trip and L else None, l100=L / trip * 100 if trip and L else None,
                 diff=trip - r["claimed_km"] if trip and r["claimed_km"] else None)
        x["diff_pct"] = x["diff"] * 100 / r["claimed_km"] if x["diff"] is not None else None
        rows.append(x)
    return rows


def fuel_stats(rows):
    c = [r for r in rows if r["trip"]]                    # Con i km percorsi
    cl = [r for r in c if r["liters_l"]]                  # Con anche i litri: contano per i consumi
    tl, tkl, tk = sum(r["liters_l"] for r in cl), sum(r["trip"] for r in cl), sum(r["trip"] for r in c)
    return dict(n=len(rows), n_cons=len(cl), tot_l=sum(r["liters_l"] or 0 for r in rows),
                tot_spend=sum(r["total"] for r in rows), km=tk,
                avg_kml=tkl / tl if tl else None, avg_l100=tl / tkl * 100 if tkl else None,
                cost100=sum(r["total"] for r in c) / tk * 100 if tk else None,
                best=min((r["l100"] for r in cl), default=None), worst=max((r["l100"] for r in cl), default=None))


@app.route("/fuel")
def fuel():
    cars = db().execute("select * from cars order by lower(name)").fetchall()
    cur = current_car()
    car = next((c for c in cars if c["id"] == request.args.get("car", type=int)), cur)
    year = request.args.get("year", type=int) or 0
    rows = fuel_data(g.user["id"], car["id"], year or None)
    pts = [r for r in reversed(rows) if r["l100"]][-24:]
    bars = []
    if len(pts) >= 2:
        lo, hi = min(r["l100"] for r in pts) * 0.8, max(r["l100"] for r in pts)
        bars = [(round(15 + 85 * (r["l100"] - lo) / ((hi - lo) or 1)), r["l100"], r["date"]) for r in pts]
    others = car_expense_rows(g.user["id"], car["id"], year or None)
    st = fuel_stats(rows)
    oth_tot = sum(r["amount"] + r["voucher"] for r in others)
    return render_template("fuel.html", cars=[c for c in cars if not c["archived"] or c["id"] == car["id"]],
                           car=car, in_use=cur["id"], rows=rows, st=st, bars=bars,
                           year=year, years=years(), fuel_cat=(fuel_cat_ids() or [0])[0],
                           car_cat=(car_cat_ids() or [0])[0], others=others, oth_tot=oth_tot,
                           car_total=st["tot_spend"] + oth_tot)


@app.route("/cars")
def cars_page():
    cars = db().execute("select c.*, (select count(*) from fuel_logs where car_id=c.id)"
                        " + (select count(*) from car_expenses where car_id=c.id) n "
                        "from cars c order by lower(name)").fetchall()
    return render_template("cars.html", cars=cars, in_use=current_car()["id"])


def _car_write(sql, args, ok):
    try:
        db().execute(sql, args)
        db().commit()
        flash(ok)
    except sqlite3.IntegrityError:
        flash("Esiste già un'auto con questo nome")
    return redirect(url_for("cars_page"))


@app.post("/cars/add")
def car_add():
    n = request.form["name"].strip()
    return _car_write("insert into cars(name) values(?)", (n,), "Auto aggiunta") if n else redirect(url_for("cars_page"))


@app.post("/cars/<int:id>/rename")
def car_rename(id):
    n = request.form["name"].strip()
    return _car_write("update cars set name=? where id=?", (n, id), "Auto rinominata") if n else redirect(url_for("cars_page"))


@app.post("/cars/<int:id>/delete")
def car_delete(id):
    if db().execute("select count(*) from cars where archived=0 and id!=?", (id,)).fetchone()[0] == 0:
        flash("Deve restare almeno un'auto")
        return redirect(url_for("cars_page"))
    if db().execute("select (select count(*) from fuel_logs where car_id=?) + "
                    "(select count(*) from car_expenses where car_id=?)", (id, id)).fetchone()[0]:
        return _car_write("update cars set archived=1 where id=?", (id,), "Auto archiviata (ha spese nello storico)")
    return _car_write("delete from cars where id=?", (id,), "Auto eliminata")


@app.post("/cars/<int:id>/restore")
def car_restore(id):
    return _car_write("update cars set archived=0 where id=?", (id,), "Auto ripristinata")


@app.post("/cars/use")
def car_use():
    cid = request.form.get("car_id", type=int)
    if cid and db().execute("select 1 from cars where id=? and archived=0", (cid,)).fetchone():
        db().execute("insert or replace into user_car(user_id,car_id) values(?,?)", (g.user["id"], cid))
        db().commit()
        flash("Ora stai usando questa auto")
    return redirect(url_for("fuel", car=cid) if cid else url_for("fuel"))


# --- Prestiti e persone ---
# Il prestito è un'opzione di una spesa normale: loans.amount è positivo se ho dato io,
# negativo se ha pagato l'altra persona. La somma per persona è il suo saldo
# (positivo: ti deve; negativo: gli devi).
def active_people(extra_id=None):
    return db().execute("select * from people where archived=0 or id=? order by lower(name)", (extra_id,)).fetchall()


def parse_loan(total):
    """Persona, direzione e importo (positivo, in centesimi) dal modulo.
    None se i dati non sono validi o l'importo supera il totale della spesa."""
    fm = request.form
    try:
        amount = to_cents(fm.get("loan_amount"))
    except (ValueError, ArithmeticError):
        return None
    if not 0 < amount <= total:
        return None
    pid, new = fm.get("person_id", ""), None
    if pid == "new":
        new = (fm.get("new_person") or "").strip()[:40]
        if not new:
            return None
        pid = None
    else:
        try:
            pid = int(pid)
        except ValueError:
            return None
        if not db().execute("select 1 from people where id=?", (pid,)).fetchone():
            return None
    return dict(person=pid, new=new, amount=-amount if fm.get("loan_dir") == "received" else amount)


def save_loan(eid, x):
    pid = x["person"]
    if pid is None:   # Persona nuova: se ne esiste già una con lo stesso nome la riusa
        r = db().execute("select id from people where lower(name)=lower(?)", (x["new"],)).fetchone()
        if r:
            pid = r[0]
            db().execute("update people set archived=0 where id=?", (pid,))
        else:
            pid = db().execute("insert into people(name) values(?)", (x["new"],)).lastrowid
    db().execute("insert or replace into loans(expense_id,person_id,amount) values(?,?,?)",
                 (eid, pid, x["amount"]))


def loan_balances(uid):
    """Saldo per persona sulle voci visibili all'utente (bal > 0: la persona gli deve)."""
    return [dict(r) for r in db().execute(
        "select p.id, p.name, p.archived, coalesce(sum(l.amount),0) bal, count(e.id) n from people p "
        "left join loans l on l.person_id=p.id "
        "left join expenses e on e.id=l.expense_id and (e.private=0 or e.owner_id=?) "
        "group by p.id order by lower(p.name)", (uid,))]


def loan_totals(bal):
    credit = sum(p["bal"] for p in bal if p["bal"] > 0)
    debt = -sum(p["bal"] for p in bal if p["bal"] < 0)
    return dict(credit=credit, debt=debt, net=credit - debt)


@app.route("/loans")
def loans():
    uid = g.user["id"]
    bal = loan_balances(uid)
    shown = [p for p in bal if not p["archived"] or p["bal"]]
    sel = next((p for p in bal if p["id"] == request.args.get("person", type=int)), None)
    rows = []
    if sel:
        run = 0
        for r in db().execute(BASE_Q + " where ln.person_id=? and (e.private=0 or e.owner_id=?) "
                              "order by e.date, e.id", (sel["id"], uid)):
            run += r["loan_part"]
            rows.append(dict(r, run=run))
        rows.reverse()
    return render_template("loans.html", people=shown, sel=sel, rows=rows, tot=loan_totals(bal))


@app.route("/people")
def people_page():
    ps = db().execute("select p.*, (select count(*) from loans where person_id=p.id) n "
                      "from people p order by lower(name)").fetchall()
    return render_template("people.html", people=ps)


def _person_write(sql, args, ok):
    try:
        db().execute(sql, args)
        db().commit()
        flash(ok)
    except sqlite3.IntegrityError:
        flash("Esiste già una persona con questo nome")
    return redirect(url_for("people_page"))


@app.post("/people/add")
def person_add():
    n = request.form["name"].strip()[:40]
    if n and db().execute("select 1 from people where lower(name)=lower(?)", (n,)).fetchone():
        flash("Esiste già una persona con questo nome")
        return redirect(url_for("people_page"))
    return _person_write("insert into people(name) values(?)", (n,), "Persona aggiunta") if n else redirect(url_for("people_page"))


@app.post("/people/<int:id>/rename")
def person_rename(id):
    n = request.form["name"].strip()[:40]
    return _person_write("update people set name=? where id=?", (n, id), "Persona rinominata") if n else redirect(url_for("people_page"))


@app.post("/people/<int:id>/delete")
def person_delete(id):
    if db().execute("select count(*) from loans where person_id=?", (id,)).fetchone()[0]:
        return _person_write("update people set archived=1 where id=?", (id,), "Persona archiviata (ha movimenti nello storico)")
    return _person_write("delete from people where id=?", (id,), "Persona eliminata")


@app.post("/people/<int:id>/restore")
def person_restore(id):
    return _person_write("update people set archived=0 where id=?", (id,), "Persona ripristinata")


# --- Viaggi ---
# Le spese di un viaggio hanno la categoria speciale Viaggi. In un viaggio di lavoro una spesa
# può essere pagata dall'azienda (company=1): resta nel viaggio ma non entra nei totali né nel riepilogo.
def is_trip(cat_id):
    return bool(db().execute("select 1 from categories where id=? and is_trip=1", (cat_id,)).fetchone())


def trip_cat_ids():
    return sorted(r[0] for r in db().execute("select id from categories where is_trip=1"))


def active_trips(extra_id=None):
    return db().execute("select * from trips where archived=0 or id=? order by lower(name)", (extra_id,)).fetchall()


def parse_trip():
    """Viaggio (esistente o nuovo) e chi ha pagato, dal modulo. None se non valido."""
    fm = request.form
    company = fm.get("paid_by") == "company"
    tid = fm.get("trip_id", "")
    if tid == "new":
        name = (fm.get("new_trip") or "").strip()[:60]
        if not name:
            return None
        return dict(trip=None, new=name, work=1 if fm.get("new_trip_work") else 0, company=company)
    try:
        tid = int(tid)
    except ValueError:
        return None
    if not db().execute("select 1 from trips where id=?", (tid,)).fetchone():
        return None
    return dict(trip=tid, new=None, work=None, company=company)


def save_trip(eid, x):
    tid = x["trip"]
    if tid is None:   # Viaggio nuovo: se ne esiste già uno con lo stesso nome lo riusa
        r = db().execute("select id from trips where lower(name)=lower(?)", (x["new"],)).fetchone()
        if r:
            tid = r[0]
            db().execute("update trips set archived=0 where id=?", (tid,))
        else:
            tid = db().execute("insert into trips(name,is_work) values(?,?)", (x["new"], x["work"])).lastrowid
    work = db().execute("select is_work from trips where id=?", (tid,)).fetchone()[0]
    db().execute("insert or replace into trip_expenses(expense_id,trip_id,company) values(?,?,?)",
                 (eid, tid, 1 if (x["company"] and work) else 0))


def trip_sums(rows):
    mine = sum(net_total(r) for r in rows if not r["company"])
    comp = sum(net_total(r) for r in rows if r["company"])
    return mine, comp


@app.route("/trips")
def trips():
    uid = g.user["id"]
    sel = db().execute("select * from trips where id=?", (request.args.get("trip", type=int),)).fetchone()
    trip_cat = (trip_cat_ids() or [0])[0]
    if sel:
        rows = db().execute(BASE_Q + " where te.trip_id=? and (e.private=0 or e.owner_id=?) "
                            "order by e.date, e.id", (sel["id"], uid)).fetchall()
        mine, comp = trip_sums(rows)
        return render_template("trips.html", sel=sel, rows=rows, mine=mine, comp=comp, trip_cat=trip_cat,
                               first=rows[0]["date"] if rows else None, last=rows[-1]["date"] if rows else None)
    lst = db().execute(
        f"select t.*, coalesce(sum(case when te.company=0 then e.amount+e.voucher-{LENT} end),0) mine, "
        f"coalesce(sum(case when te.company=1 then e.amount+e.voucher-{LENT} end),0) comp, count(e.id) n, "
        "min(e.date) first, max(e.date) last from trips t "
        "left join trip_expenses te on te.trip_id=t.id "
        "left join expenses e on e.id=te.expense_id and (e.private=0 or e.owner_id=?) "
        "left join loans ln on ln.expense_id=e.id "
        "group by t.id order by max(e.date) is null, max(e.date) desc, lower(t.name)", (uid,)).fetchall()
    return render_template("trips.html", sel=None, trips=lst, trip_cat=trip_cat)


@app.post("/trips/add")
def trip_add():
    n = request.form["name"].strip()[:60]
    if not n:
        return redirect(url_for("trips"))
    if db().execute("select 1 from trips where lower(name)=lower(?)", (n,)).fetchone():
        flash("Esiste già un viaggio con questo nome")
        return redirect(url_for("trips"))
    cur = db().execute("insert into trips(name,is_work) values(?,?)", (n, 1 if request.form.get("work") else 0))
    db().commit()
    flash("Viaggio creato")
    return redirect(url_for("trips", trip=cur.lastrowid))


@app.post("/trips/<int:id>/update")
def trip_update(id):
    n = request.form["name"].strip()[:60]
    work = 1 if request.form.get("work") else 0
    if not n:
        return redirect(url_for("trips", trip=id))
    try:
        db().execute("update trips set name=?, is_work=? where id=?", (n, work, id))
        msg = "Viaggio aggiornato"
        if not work:   # un viaggio non di lavoro non ha spese dell'azienda: diventano tutte mie
            k = db().execute("update trip_expenses set company=0 where trip_id=? and company=1", (id,)).rowcount
            if k:
                msg += f": {k} spese pagate dall'azienda sono diventate tue"
        db().commit()
        flash(msg)
    except sqlite3.IntegrityError:
        flash("Esiste già un viaggio con questo nome")
    return redirect(url_for("trips", trip=id))


@app.post("/trips/<int:id>/delete")
def trip_delete(id):
    if db().execute("select count(*) from trip_expenses where trip_id=?", (id,)).fetchone()[0]:
        db().execute("update trips set archived=1 where id=?", (id,))
        flash("Viaggio archiviato (ha spese nello storico)")
    else:
        db().execute("delete from trips where id=?", (id,))
        flash("Viaggio eliminato")
    db().commit()
    return redirect(url_for("trips"))


@app.post("/trips/<int:id>/restore")
def trip_restore(id):
    db().execute("update trips set archived=0 where id=?", (id,))
    db().commit()
    flash("Viaggio ripristinato")
    return redirect(url_for("trips", trip=id))


@app.post("/trips/expense/<int:id>/payer")
def trip_payer(id):
    """Sposta una spesa di un viaggio di lavoro tra "mia" e "pagata dall'azienda"."""
    get_expense(id)   # Solo se l'utente la vede
    r = db().execute("select te.trip_id, te.company from trip_expenses te join trips t on t.id=te.trip_id "
                     "where te.expense_id=? and t.is_work=1", (id,)).fetchone()
    if not r:
        abort(404)
    new = 0 if r["company"] else 1
    db().execute("update trip_expenses set company=? where expense_id=?", (new, id))
    db().execute("update expenses set updated_by=?, updated_at=? where id=?", (g.user["id"], now(), id))
    db().commit()
    flash("Ora la paga l'azienda: non conta nei tuoi totali" if new else "Ora è una spesa tua: conta nei tuoi totali")
    return redirect(url_for("trips", trip=r["trip_id"]))


# --- Riepilogo ---
def summary(year, scope, uid):
    if scope == "shared":
        w, p = "e.private=0", ()
    elif scope == "mine":
        w, p = "e.private=1 and e.owner_id=?", (uid,)
    else:
        w, p = "(e.private=0 or e.owner_id=?)", (uid,)
    # La parte prestata non è una spesa dell'utente: si toglie prima dall'importo, poi dal buono
    rows = db().execute(f"""select cast(substr(e.date,6,2) as int) m, c.name cat,
        sum(e.amount - {LENT_A}) a, sum(e.voucher - ({LENT} - {LENT_A})) v
        from expenses e join categories c on c.id=e.category_id
        left join trip_expenses te on te.expense_id=e.id
        left join loans ln on ln.expense_id=e.id
        where substr(e.date,1,4)=? and coalesce(te.company,0)=0 and {w} group by m, c.name""", (str(year),) + p).fetchall()
    cats = sorted({r["cat"] for r in rows}, key=str.lower)
    cell = {(r["m"], r["cat"]): r["a"] + r["v"] for r in rows}
    table = []
    for m in range(1, 13):
        vals = [cell.get((m, c), 0) for c in cats]
        table.append((MONTHS[m - 1], vals, sum(vals)))
    coltot = [sum(r[1][i] for r in table) for i in range(len(cats))]
    mt = [(MONTHS[m - 1], sum(r["a"] for r in rows if r["m"] == m),
           sum(r["v"] for r in rows if r["m"] == m)) for m in range(1, 13)]
    ct = [(c, sum(r["a"] for r in rows if r["cat"] == c),
           sum(r["v"] for r in rows if r["cat"] == c)) for c in cats]
    return cats, table, coltot, mt, ct


def pie(cats, coltot):
    tot = sum(v for v in coltot if v > 0)
    acc, parts, items = 0.0, [], []
    for i, (c, v) in enumerate(sorted(zip(cats, coltot), key=lambda x: -x[1])):
        if v <= 0:
            continue
        pct = v * 100 / tot
        parts.append(f"{PALETTE[i % 12]} {acc:.2f}% {acc + pct:.2f}%")
        items.append((c, v, pct, PALETTE[i % 12]))
        acc += pct
    return ", ".join(parts) or "#ddd 0% 100%", items


@app.route("/summary")
def summary_page():
    year = request.args.get("year", type=int) or date.today().year
    scope = request.args.get("scope")
    if scope not in ("shared", "mine", "all"):
        scope = "shared"
    cats, table, coltot, mt, ct = summary(year, scope, g.user["id"])
    grad, items = pie(cats, coltot)
    lb = loan_balances(g.user["id"])
    w, p = scope_where(scope, g.user["id"])
    company_tot = db().execute("select coalesce(sum(e.amount+e.voucher),0) from expenses e "
                               "join trip_expenses te on te.expense_id=e.id where te.company=1 and "
                               "substr(e.date,1,4)=? and " + w, (str(year),) + p).fetchone()[0]
    n_exp = db().execute("select count(*) from expenses e left join trip_expenses te on te.expense_id=e.id "
                         "where substr(e.date,1,4)=? and coalesce(te.company,0)=0 and " + w,
                         (str(year),) + p).fetchone()[0]
    mtot = [a + v for _, a, v in mt]
    mmax = max(mtot)
    full = [t for t in mtot if t]
    mbars = [(name[:3], t, max(round(t * 100 / mmax), 3) if mmax > 0 and t > 0 else 0)
             for (name, _, _), t in zip(mt, mtot)]
    mstat = dict(top=MONTHS[mtot.index(mmax)] if mmax > 0 else None, top_v=mmax,
                 avg=sum(full) / len(full) if full else 0)
    return render_template("summary.html", year=year, scope=scope, years=years(), cats=cats,
                           mbars=mbars, mstat=mstat, company_tot=company_tot,
                           open_loans=[p for p in lb if p["bal"]], loan_tot=loan_totals(lb),
                           table=table, coltot=coltot, grand=sum(coltot), mt=mt, ct=ct,
                           tot_a=sum(a for _, a, _ in mt), tot_v=sum(v for _, _, v in mt), n_exp=n_exp,
                           grad=grad, items=items)


# --- Esportazione Excel ---
def fuel_sheet(wb, fr):
    s = wb.create_sheet("Rifornimenti")
    s.append(["Data", "Auto", "N. rifornimento", "Km totali", "Litri", "Prezzo/L", "Importo totale",
              "Km col pieno precedente", "Km dichiarati", "km/L", "L/100km"])
    for r in reversed(fr):
        s.append([date.fromisoformat(r["date"]), r["car"], r["fill_no"], r["odometer"], r["liters_l"],
                  r["price"], r["total"] / 100, r["trip"], r["claimed_km"],
                  round(r["kml"], 2) if r["kml"] else None, round(r["l100"], 2) if r["l100"] else None])
    for row in s.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"


def xlsx_response(wb, name):
    for sh in wb:
        for c in sh[1]:
            c.font = Font(bold=True)
        for col in sh.columns:
            sh.column_dimensions[col[0].column_letter].width = 16
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/backup")
def backup_page():
    cars = db().execute("select * from cars order by lower(name)").fetchall()
    people = db().execute("select * from people order by lower(name)").fetchall()
    trips = db().execute("select * from trips order by lower(name)").fetchall()
    return render_template("backup.html", years=years(), cars=cars, cur_car=current_car()["id"], people=people,
                           trips=trips)


# --- Importazione da Excel ---
@app.route("/import")
def import_page():
    trips = db().execute("select * from trips order by lower(name)").fetchall()
    return render_template("import.html", trips=trips)


IMPORT_TITLES = {"spese": "Importa spese", "viaggio": "Importa spese di un viaggio", "rifornimenti": "Importa rifornimenti"}


def _trip_opt():
    fm = request.form
    tid = fm.get("trip_id", "")
    if tid == "new":
        return dict(id=None, new=(fm.get("new_trip") or "").strip(), work=bool(fm.get("new_trip_work")))
    try:
        return dict(id=int(tid), new=None, work=False)
    except ValueError:
        return None


def _import_plan(kind, data, meta=None, opt=None):
    """Legge il file e prepara il piano senza scrivere nulla. Ritorna (piano, opzioni del viaggio).
    meta: correzioni ed esclusioni della procedura guidata.
    """
    meta = meta or {}
    if kind == "viaggio":
        opt = opt or _trip_opt()
        if not opt:
            raise importer.ImportFileError("Scegli il viaggio")
    try:
        fixes = {int(k): v for k, v in (meta.get("fix") or {}).items()}
        return importer.make_plan(db(), g.user["id"], kind, io.BytesIO(data), opt,
                                  fixes=fixes, exclude=meta.get("exclude") or [],
                                  force=meta.get("force") or []), opt
    except importer.ImportFileError:
        raise
    except Exception:
        app.logger.exception("import: lettura fallita")
        raise importer.ImportFileError("Non riesco a leggere il file: controlla che sia un Excel (.xlsx) con le intestazioni giuste")


def _clean_import_tmp():
    os.makedirs(IMP_DIR, exist_ok=True)
    limit = datetime.now().timestamp() - 86400
    for p in glob(os.path.join(IMP_DIR, "*.xlsx")) + glob(os.path.join(IMP_DIR, "*.json")):
        if os.path.getmtime(p) < limit:
            os.remove(p)


def _meta_path(token):
    return os.path.join(IMP_DIR, token + ".json")


def _imp_load(token):
    """Importazione in attesa di conferma per l'utente in sessione: (meta, bytes) o None."""
    path = os.path.join(IMP_DIR, token + ".xlsx")
    if not re.fullmatch(r"[A-Za-z0-9_-]{16,32}", token or "") or session.get("imp") != token \
            or not os.path.exists(path):
        return None
    try:
        with open(_meta_path(token), encoding="utf-8") as fh:
            meta = json.load(fh)
        with open(path, "rb") as fh:
            return meta, fh.read()
    except (OSError, ValueError):
        return None


def _imp_save(token, meta):
    with open(_meta_path(token), "w", encoding="utf-8") as fh:
        json.dump(meta, fh)


def _preview_page(token, meta, plan, opt):
    nfix = sum(1 for v in (meta.get("fix") or {}).values() if v)
    return render_template("import_preview.html", plan=plan, kind=meta["kind"], token=token, opt=opt,
                           title=IMPORT_TITLES[meta["kind"]], fname=meta.get("fname", ""),
                           counts=plan.counts, nfix=nfix)


@app.post("/import/preview")
def import_preview():
    kind = request.form.get("kind")
    if kind not in IMPORT_TITLES:
        abort(404)
    f = request.files.get("file")
    data = f.read(MAX_IMPORT + 1) if f and f.filename else b""
    err = ("Scegli il file Excel (.xlsx)" if not data else "File troppo grande (massimo 5 MB)" if len(data) > MAX_IMPORT
           else "Il file non è un Excel (.xlsx)" if not data.startswith(b"PK") else None)
    if not err:
        try:
            plan, opt = _import_plan(kind, data)
        except importer.ImportFileError as e:
            err = str(e)
    if err:
        flash(err)
        return redirect(url_for("import_page"))
    _clean_import_tmp()
    token = secrets.token_urlsafe(16)
    with open(os.path.join(IMP_DIR, token + ".xlsx"), "wb") as fh:
        fh.write(data)
    meta = dict(kind=kind, opt=opt, fname=f.filename, fix={}, exclude=[])
    _imp_save(token, meta)
    session["imp"] = token
    return _preview_page(token, meta, plan, opt)


@app.route("/import/review")
def import_review():
    """Mostra di nuovo l'anteprima di un file già caricato, dopo le correzioni."""
    token = request.args.get("token", "")
    ld = _imp_load(token)
    if not ld:
        flash("Importazione scaduta: carica di nuovo il file")
        return redirect(url_for("import_page"))
    meta, data = ld
    try:
        plan, opt = _import_plan(meta["kind"], data, meta, meta.get("opt"))
    except importer.ImportFileError as e:
        flash(str(e))
        return redirect(url_for("import_page"))
    return _preview_page(token, meta, plan, opt)


def _fix_ctx(cur, kind, meta):
    """Dati per il modulo di correzione. Ritorna (valori dei campi, campi in errore,
    scelte dei menu, viaggio di lavoro, dati per lo script).
    """
    yes = {"si", "s", "yes", "y", "true", "vero", "x", "1"}
    comp = {"azienda", "company", "ditta", "societa", "lavoro"}
    raw = cur["raw"]
    err = {"paidby" if k == "company" else k for k in cur["fix"]}
    v = {}
    for key, (_, typ) in importer.FIELDS.items():
        r = raw.get(key, "")
        if typ == "date":
            try:
                v[key] = importer.p_date(r).isoformat() if r else ""
            except importer.RowError:
                v[key] = ""
        elif typ == "num":
            try:
                d = importer._dec(r)
                v[key] = "" if d is None else format(d.normalize(), "f").replace(".", ",")
            except Exception:
                v[key] = r
        elif typ == "yn":
            v[key] = "Sì" if importer.norm(r) in yes else "No"
        elif typ == "paid":
            v[key] = "Azienda" if importer.norm(raw.get("company", "")) in yes or importer.norm(r) in comp else "Io"
        else:
            v[key] = r
    if v["loan"] or v["person"] or {"loan", "person"} & err:
        v["loanon"] = True

    def pick(key, rows, allow_new=True):
        """Menu a tendina con i nomi esistenti, più quello del file se è nuovo."""
        cur_name = raw.get(key, "").strip()
        match = next((n for n, _ in rows if importer.norm(n) == importer.norm(cur_name)), None)
        return dict(key=key, rows=rows, cur=match or "", extra="" if match else cur_name, allow_new=allow_new)

    c = db()
    cats = [(r[0], None) for r in c.execute("select name from categories where archived=0 and is_fuel=0 order by lower(name)")]
    cars = [(r[0], None) for r in c.execute("select name from cars order by archived, lower(name)")]
    people = [(r[0], None) for r in c.execute("select name from people order by archived, lower(name)")]
    trips = [(r[0], r[1]) for r in c.execute("select name, is_work from trips order by archived, lower(name)")]
    work = None  # Nell'import di un viaggio il viaggio è già stato scelto
    if kind == "viaggio":
        o = meta.get("opt") or {}
        t = c.execute("select is_work from trips where id=?", (o.get("id"),)).fetchone() if o.get("id") else None
        work = bool(t[0]) if t else bool(o.get("work"))
    P = dict(category=pick("category", cats, False), car=pick("car", cars), trip=pick("trip", trips),
             person=pick("person", people))
    js = dict(trip_cats=[r[0].lower() for r in c.execute("select name from categories where is_trip=1")],
              car_cats=[r[0].lower() for r in c.execute("select name from categories where is_car=1")])
    return v, err, P, work, js


@app.post("/import/force")
def import_force():
    """Esclude la riga dall'abbinamento con le spese già presenti ("È diversa")."""
    token = request.form.get("token", "")
    ld = _imp_load(token)
    row = request.form.get("row", type=int)
    if not ld or not row:
        flash("Importazione scaduta: carica di nuovo il file")
        return redirect(url_for("import_page"))
    meta = ld[0]
    if row not in meta.setdefault("force", []):
        meta["force"].append(row)
    _imp_save(token, meta)
    return redirect(url_for("import_review", token=token))


@app.route("/import/fix")
def import_fix():
    """Procedura guidata: una riga con errore alla volta, a partire da quella dopo "after"."""
    token = request.args.get("token", "")
    ld = _imp_load(token)
    if not ld:
        flash("Importazione scaduta: carica di nuovo il file")
        return redirect(url_for("import_page"))
    meta, data = ld
    try:
        plan, _ = _import_plan(meta["kind"], data, meta, meta.get("opt"))
    except importer.ImportFileError as e:
        flash(str(e))
        return redirect(url_for("import_page"))
    after = request.args.get("after", type=int) or 0
    errs = [i for i in plan.items if i["status"] == "error"]
    cur = next((i for i in errs if i["row"] > after), None)
    if not cur:
        flash("Errori risolti ✓" if not errs else
              "Restano %d %s con errori: non verranno importate" % (len(errs), "riga" if len(errs) == 1 else "righe"))
        return redirect(url_for("import_review", token=token))
    v, err, P, work, js = _fix_ctx(cur, meta["kind"], meta)
    return render_template("import_fix.html", token=token, it=cur, kind=meta["kind"], v=v, err=err, P=P,
                           work=work, js=js,
                           pos=sum(1 for i in errs if i["row"] <= cur["row"]), total=len(errs),
                           fname=meta.get("fname", ""))


@app.post("/import/fix")
def import_fix_save():
    token = request.form.get("token", "")
    ld = _imp_load(token)
    row = request.form.get("row", type=int)
    if not ld or not row:
        flash("Importazione scaduta: carica di nuovo il file")
        return redirect(url_for("import_page"))
    meta = ld[0]
    action = request.form.get("action")
    after = row  # Salta: si passa alla riga successiva
    if action == "exclude":
        if row not in meta.setdefault("exclude", []):
            meta["exclude"].append(row)
    elif action == "save":
        fx = meta.setdefault("fix", {}).setdefault(str(row), {})
        fm = request.form
        for key in importer.FIELDS:
            vals = fm.getlist("f_" + key)          # I checkbox inviano prima No (campo nascosto) e poi Sì, se spuntati
            if vals:
                val = vals[-1].strip()
                if val == "__new__":               # "+ Nuovo…": vale il nome scritto a parte
                    val = (fm.get("f_" + key + "_new") or "").strip()
                fx[key] = val[:200]
        if "f_paidby" in fm:                       # "Pagata da" vale anche per la colonna "Pagata dall'azienda"
            fx["company"] = "Sì" if fm["f_paidby"] == "Azienda" else "No"
        if fm.getlist("f_loanon") and fm.getlist("f_loanon")[-1] != "1":   # Prestito disattivato
            fx["loan"] = fx["person"] = ""
        after = row - 1                            # La stessa riga può avere un altro errore
    _imp_save(token, meta)
    return redirect(url_for("import_fix", token=token, after=after))


def _snapshot():
    """Copia del database prima di importare; si tengono le ultime 5."""
    snap = os.path.join(DATA, f"prima_import_{datetime.now():%Y%m%d_%H%M%S}.db")
    d = sqlite3.connect(snap)
    db().backup(d)
    d.close()
    for old in sorted(glob(os.path.join(DATA, "prima_import_*.db")))[:-5]:
        os.remove(old)


def _import_message(kind, st):
    parts = []
    if st["nuove"]:
        parts.append("%d %s" % (st["nuove"], ("rifornimenti importati" if kind == "rifornimenti" else "spese importate")
                                if st["nuove"] != 1 else ("rifornimento importato" if kind == "rifornimenti" else "spesa importata")))
    if st["completate"]:
        parts.append("%d spese Benzina completate con i dettagli" % st["completate"])
    msg = "Importazione fatta ✓ " + ", ".join(parts)
    made = ["%s: %s" % (k, ", ".join(v)) for k, v in st["creati"].items()]
    if made:
        msg += ". Creati - " + "; ".join(made)
    if st["senza_nomi"]:
        msg += ". %d righe non importate perché usano nomi che non hai voluto creare" % st["senza_nomi"]
    return msg


@app.post("/import/confirm")
def import_confirm():
    token = request.form.get("token", "")
    kind = request.form.get("kind")
    ld = _imp_load(token)
    if kind not in IMPORT_TITLES or not ld:
        flash("Importazione scaduta: carica di nuovo il file")
        return redirect(url_for("import_page"))
    meta, data = ld
    try:
        plan, _ = _import_plan(kind, data, meta)  # Si ricalcola: nel frattempo il database può essere cambiato
    except importer.ImportFileError as e:
        flash(str(e))
        return redirect(url_for("import_page"))
    try:
        _snapshot()
        st = importer.apply_plan(db(), g.user["id"], plan, create_missing=bool(request.form.get("create")))
        db().commit()
    except Exception:
        db().rollback()
        app.logger.exception("import: scrittura fallita")
        flash("Importazione non riuscita: non è stato modificato nulla")
        return redirect(url_for("import_page"))
    for p in (os.path.join(IMP_DIR, token + ".xlsx"), _meta_path(token)):
        if os.path.exists(p):
            os.remove(p)
    session.pop("imp", None)
    if not (st["nuove"] or st["completate"]):
        flash("Nessuna riga importata" + (": mancano i nomi da creare" if st["senza_nomi"] else ""))
        return redirect(url_for("import_page"))
    flash(_import_message(kind, st))
    if kind == "rifornimenti":
        return redirect(url_for("fuel", car=st["last_car"]) if st["last_car"] else url_for("fuel"))
    return redirect(url_for("history", year=(st["last_date"] or str(date.today()))[:4]))


@app.route("/backup/download")
def backup():
    year = request.args.get("year", type=int) or date.today().year
    scope = "mine" if request.args.get("scope") == "mine" else "shared"
    w, p = ("e.private=0", ()) if scope == "shared" else ("e.private=1 and e.owner_id=?", (g.user["id"],))
    rows = db().execute(BASE_Q + f" where substr(e.date,1,4)=? and {w} order by e.date, e.id",
                        (str(year),) + p).fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "Spese"
    ws.append(["Data", "Descrizione", "Categoria", "Importo", "Buono", "Totale", "Privata",
               "Inserita da", "Modificata da", "Modificata il", "Auto", "Persona", "Viaggio", "Pagata dall'azienda", "Prestito (+dato / -ricevuto)"])
    car_of = {r[0]: r[1] for r in db().execute(
        "select f.expense_id, k.name from fuel_logs f join cars k on k.id=f.car_id "
        "union all select x.expense_id, k.name from car_expenses x join cars k on k.id=x.car_id")}
    for r in rows:
        ws.append([date.fromisoformat(r["date"]), r["description"], r["cat"], r["amount"] / 100,
                   r["voucher"] / 100, (r["amount"] + r["voucher"]) / 100,
                   "Sì" if r["private"] else "No", r["owner"], r["updater"] or "", r["updated_at"] or "",
                   car_of.get(r["id"], ""), r["person"] or "", r["trip"] or "", "Sì" if r["company"] else "",
                   r["loan_part"] / 100 if r["loan_part"] is not None else None])
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"
        for c in row[3:6]:
            c.number_format = '#,##0.00'
        row[14].number_format = '#,##0.00'
    cats, table, coltot, mt, _ = summary(year, scope, g.user["id"])
    s = wb.create_sheet("Riepilogo")
    s.append(["Mese"] + cats + ["Totale"])
    for name, vals, tot in table:
        s.append([name] + [v / 100 for v in vals] + [tot / 100])
    s.append(["Totale"] + [v / 100 for v in coltot] + [sum(coltot) / 100])
    s.append([])
    s.append(["Mese", "Importo", "Buoni", "Totale"])
    for name, a, v in mt:
        s.append([name, a / 100, v / 100, (a + v) / 100])
    for c in s[1]:
        c.font = Font(bold=True)
    fr = fuel_data(g.user["id"], None, year, scope)
    if fr:
        fuel_sheet(wb, fr)
    pref = "private_" if scope == "mine" else ""
    return xlsx_response(wb, f"raspicash_{pref}{year}_{date.today():%Y%m%d}.xlsx")


@app.route("/backup/loans")
def backup_loans():
    """Prestiti per anno (0 = tutti) e persona (0 = tutte): movimenti, riporto e saldo finale."""
    year = request.args.get("year", type=int) or 0
    pid = request.args.get("person", type=int) or 0
    uid = g.user["id"]
    q = BASE_Q + " where ln.person_id is not null and (e.private=0 or e.owner_id=?)"
    args = [uid]
    if pid:
        q += " and ln.person_id=?"
        args.append(pid)
    allrows = db().execute(q + " order by e.date, e.id", args).fetchall()
    if year:
        allrows = [r for r in allrows if int(r["date"][:4]) <= year]
    by = {}
    for r in allrows:
        by.setdefault(r["person"], []).append(r)
    data = []   # (persona, riporto, movimenti dell'anno)
    for name in sorted(by, key=str.lower):
        old = [r for r in by[name] if year and int(r["date"][:4]) < year]
        cur = [r for r in by[name] if not year or int(r["date"][:4]) == year]
        carry = sum(r["loan_part"] for r in old)
        if cur or carry:
            data.append((name, carry, cur))
    if not data:
        flash("Nessun prestito per questa scelta")
        return redirect(url_for("backup_page"))

    def status(name, b):
        return f"{name} ti deve" if b > 0 else (f"Tu devi a {name}" if b < 0 else "In pari")

    wb = Workbook()
    rs = wb.active
    rs.title = "Riepilogo"
    rs.append(["Persona", "Riporto anni precedenti", "Prestati", "Anticipati", "Saldo finale", "Situazione"])
    ms = wb.create_sheet("Movimenti")
    ms.append(["Data", "Persona", "Descrizione", "Prestati", "Anticipati", "Saldo progressivo", "Privata"])
    finals = []
    for name, carry, cur in data:
        given = sum(r["loan_part"] for r in cur if r["loan_part"] > 0)
        recv = -sum(r["loan_part"] for r in cur if r["loan_part"] < 0)
        final = carry + given - recv
        finals.append(final)
        rs.append([name, carry / 100, given / 100, recv / 100, final / 100, status(name, final)])
        run = carry
        if carry:
            ms.append([None, name, "Riporto anni precedenti", None, None, run / 100, ""])
        for r in cur:
            run += r["loan_part"]
            ms.append([date.fromisoformat(r["date"]), name, r["description"],
                       r["loan_part"] / 100 if r["loan_part"] > 0 else None,
                       -r["loan_part"] / 100 if r["loan_part"] < 0 else None, run / 100,
                       "Sì" if r["private"] else ""])
    credit, debt = sum(f for f in finals if f > 0), -sum(f for f in finals if f < 0)
    rs.append([])
    rs.append(["Totale che ti devono", credit / 100])
    rs.append(["Totale che devi", debt / 100])
    rs.append(["Saldo netto (positivo = ti devono)", (credit - debt) / 100])
    rs.append(["Periodo", str(year) if year else "Tutti gli anni"])
    for row in rs.iter_rows(min_row=2):
        for c in row[1:5]:
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0.00"
    for row in ms.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"
        for c in row[3:6]:
            c.number_format = "#,##0.00"
    for r in range(rs.max_row - 3, rs.max_row):
        rs.cell(r, 1).font = Font(bold=True)
    who = safe_name(data[0][0]) if len(data) == 1 and pid else "tutte"
    return xlsx_response(wb, f"raspicash_prestiti_{who}_{year or 'tutti'}_{date.today():%Y%m%d}.xlsx")


@app.route("/backup/trip")
def backup_trip():
    """Spese di un viaggio con chi ha pagato e i totali."""
    t = db().execute("select * from trips where id=?", (request.args.get("trip", type=int),)).fetchone()
    if not t:
        flash("Scegli un viaggio")
        return redirect(url_for("backup_page"))
    rows = db().execute(BASE_Q + " where te.trip_id=? and (e.private=0 or e.owner_id=?) order by e.date, e.id",
                        (t["id"], g.user["id"])).fetchall()
    if not rows:
        flash(f"Nessuna spesa registrata per {t['name']}")
        return redirect(url_for("backup_page"))
    mine, comp = trip_sums(rows)
    wb = Workbook()
    rs = wb.active
    rs.title = "Riepilogo"
    ws = wb.create_sheet("Spese")
    ws.append(["Data", "Descrizione", "Importo", "Buono", "Totale", "Pagata da", "Inserita da", "Privata"])
    for r in rows:
        ws.append([date.fromisoformat(r["date"]), r["description"], r["amount"] / 100, r["voucher"] / 100,
                   (r["amount"] + r["voucher"]) / 100, "Azienda" if r["company"] else "Io", r["owner"],
                   "Sì" if r["private"] else ""])
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"
        for c in row[2:5]:
            c.number_format = "#,##0.00"
    f, l = date.fromisoformat(rows[0]["date"]), date.fromisoformat(rows[-1]["date"])
    rs.append(["Viaggio", t["name"]])
    rs.append(["Tipo", "Di lavoro" if t["is_work"] else "Personale"])
    rs.append(["Periodo", f"{f:%d/%m/%Y} - {l:%d/%m/%Y}"])
    rs.append(["Numero spese", len(rows)])
    rs.append([])
    rs.append(["Spese mie", mine / 100])
    if t["is_work"]:
        rs.append(["Pagate dall'azienda", comp / 100])
    rs.append(["Totale viaggio", (mine + comp) / 100])
    for row in rs.iter_rows(min_row=6):
        row[1].number_format = "#,##0.00"
    for row in rs.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
    return xlsx_response(wb, f"raspicash_viaggio_{safe_name(t['name'])}_{date.today():%Y%m%d}.xlsx")


def car_from_arg():
    return db().execute("select * from cars where id=?", (request.args.get("car", type=int),)).fetchone()


def safe_name(s):
    return "".join(ch if ch.isalnum() else "_" for ch in s).strip("_") or "auto"


@app.route("/backup/fuel")
def backup_fuel():
    """Rifornimenti di un'auto, tutti gli anni."""
    car = car_from_arg()
    if not car:
        flash("Scegli un'auto")
        return redirect(url_for("backup_page"))
    fr = fuel_data(g.user["id"], car["id"], None, "all")
    if not fr:
        flash(f"Nessun rifornimento registrato per {car['name']}")
        return redirect(url_for("backup_page"))
    wb = Workbook()
    wb.remove(wb.active)
    fuel_sheet(wb, fr)
    return xlsx_response(wb, f"raspicash_rifornimenti_{safe_name(car['name'])}_{date.today():%Y%m%d}.xlsx")


@app.route("/backup/car")
def backup_car():
    """Storico completo di un'auto: benzina e altre spese, tutti gli anni."""
    car = car_from_arg()
    if not car:
        flash("Scegli un'auto")
        return redirect(url_for("backup_page"))
    uid = g.user["id"]
    rows = db().execute(BASE_Q + " where e.id in (select expense_id from fuel_logs where car_id=? union "
                        "select expense_id from car_expenses where car_id=?) and (e.private=0 or e.owner_id=?) "
                        "order by e.date, e.id", (car["id"], car["id"], uid)).fetchall()
    if not rows:
        flash(f"Nessuna spesa registrata per {car['name']}")
        return redirect(url_for("backup_page"))
    fuel_ids = {r[0] for r in db().execute("select expense_id from fuel_logs where car_id=?", (car["id"],))}
    per_year = {}
    wb = Workbook()
    rs = wb.active
    rs.title = "Riepilogo"
    ws = wb.create_sheet("Spese")
    ws.append(["Data", "Descrizione", "Categoria", "Importo", "Buono", "Totale", "Inserita da"])
    for r in rows:
        tot = r["amount"] + r["voucher"]
        ws.append([date.fromisoformat(r["date"]), r["description"], r["cat"], r["amount"] / 100,
                   r["voucher"] / 100, tot / 100, r["owner"]])
        per_year.setdefault(r["date"][:4], [0, 0])[0 if r["id"] in fuel_ids else 1] += tot
    ws.append(["Totale", "", "", sum(r["amount"] for r in rows) / 100, sum(r["voucher"] for r in rows) / 100,
               sum(r["amount"] + r["voucher"] for r in rows) / 100])
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"
        for c in row[3:6]:
            c.number_format = "#,##0.00"
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    rs.append(["Anno", "Benzina", "Altre spese", "Totale"])
    for y in sorted(per_year):
        f, o = per_year[y]
        rs.append([int(y), f / 100, o / 100, (f + o) / 100])
    rs.append(["Totale", sum(v[0] for v in per_year.values()) / 100, sum(v[1] for v in per_year.values()) / 100,
               sum(v[0] + v[1] for v in per_year.values()) / 100])
    for row in rs.iter_rows(min_row=2):
        for c in row[1:]:
            c.number_format = "#,##0.00"
    for c in rs[rs.max_row]:
        c.font = Font(bold=True)
    fr = fuel_data(uid, car["id"], None, "all")
    if fr:
        fuel_sheet(wb, fr)
    return xlsx_response(wb, f"raspicash_storico_{safe_name(car['name'])}_{date.today():%Y%m%d}.xlsx")


# --- Amministrazione ---
def restore_db(f):
    """Sostituisce il database con quello del file caricato.
    Ritorna None se tutto ok, altrimenti il messaggio d'errore.
    """
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "up.db")
        f.save(path)
        try:
            src = sqlite3.connect(path)
            tables = {r[0] for r in src.execute("select name from sqlite_master where type='table'")}
            if not {"users", "categories", "expenses"} <= tables or \
                    not src.execute("select count(*) from users where is_admin=1").fetchone()[0]:
                src.close()
                return "Il file non è un backup valido di RaspiCash"
            if src.execute("pragma user_version").fetchone()[0] > SCHEMA_VERSION:
                src.close()
                return "Il backup è di una versione più recente di RaspiCash"
            dst = sqlite3.connect(DB)
            src.backup(dst)
            dst.close()
            src.close()
        except sqlite3.DatabaseError:
            return "File non valido o danneggiato"
    init_db()
    return None


@app.route("/setup/restore", methods=["POST"])
def restore_setup():
    """Ripristino su un'installazione nuova (solo se non ci sono ancora utenti)."""
    if db().execute("select count(*) from users").fetchone()[0]:
        return redirect(url_for("login"))
    f = request.files.get("backup")
    err = restore_db(f) if f else "Scegli il file di backup"
    flash(err or "Backup ripristinato: accedi con i tuoi utenti di prima")
    return redirect(url_for("setup" if err else "login"))


@app.route("/admin")
@admin_required
def admin():
    return render_template("admin.html", years=years())


@app.route("/admin/backup")
@admin_required
def admin_backup():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "b.db")
        dst = sqlite3.connect(path)
        db().backup(dst)  # copia coerente anche con l'app in uso
        dst.close()
        with open(path, "rb") as f:
            data = f.read()
    return send_file(io.BytesIO(data), as_attachment=True, mimetype="application/octet-stream",
                     download_name=f"raspicash_backup_{datetime.now():%Y%m%d_%H%M}.db")


@app.post("/admin/restore")
@admin_required
def admin_restore():
    f = request.files.get("backup")
    if not f or request.form.get("confirm", "").strip().upper() != "RIPRISTINA":
        flash("Scegli il file e scrivi RIPRISTINA per confermare")
        return redirect(url_for("admin"))
    snap = os.path.join(DATA, f"prima_del_ripristino_{datetime.now():%Y%m%d_%H%M}.db")
    d = sqlite3.connect(snap)
    db().backup(d)  # copia di sicurezza dello stato attuale
    d.close()
    err = restore_db(f)
    if err:
        flash(err)
        return redirect(url_for("admin"))
    session.clear()
    flash("Backup ripristinato: accedi di nuovo")
    return redirect(url_for("login"))


@app.post("/admin/delete-year")
@admin_required
def admin_delete_year():
    year = request.form.get("year", type=int)
    if not year or request.form.get("confirm", "").strip() != str(year):
        flash("Per confermare scrivi l'anno esatto")
    else:
        for t in ("fuel_logs", "car_expenses", "loans", "trip_expenses"):
            db().execute(f"delete from {t} where expense_id in "
                         "(select id from expenses where substr(date,1,4)=?)", (str(year),))
        n = db().execute("delete from expenses where substr(date,1,4)=?", (str(year),)).rowcount
        db().commit()
        db().execute("vacuum")
        flash(f"Anno {year} eliminato: {n} spese rimosse")
    return redirect(url_for("admin"))


# --- Utenti ---
@app.route("/users")
@admin_required
def users():
    us = db().execute("select u.*, (select count(*) from expenses where owner_id=u.id) n from users u order by lower(name)").fetchall()
    return render_template("users.html", us=us)


@app.post("/users/add")
@admin_required
def user_add():
    n, p = request.form["name"].strip(), request.form["password"]
    if n and len(p) >= 4:
        try:
            db().execute("insert into users(name,pw_hash,is_admin) values(?,?,?)",
                         (n, generate_password_hash(p), 1 if request.form.get("admin") else 0))
            db().commit()
            flash("Utente aggiunto")
        except sqlite3.IntegrityError:
            flash("Nome già esistente")
    else:
        flash("Servono un nome e una password di almeno 4 caratteri")
    return redirect(url_for("users"))


@app.post("/users/<int:id>/password")
@admin_required
def user_pw(id):
    if len(request.form["password"]) >= 4:
        db().execute("update users set pw_hash=? where id=?", (generate_password_hash(request.form["password"]), id))
        db().commit()
        flash("Password cambiata")
    else:
        flash("Password troppo corta")
    return redirect(url_for("users"))


@app.post("/users/<int:id>/delete")
@admin_required
def user_del(id):
    n = db().execute("select count(*) from expenses where owner_id=?", (id,)).fetchone()[0]
    if id == g.user["id"] or n:
        flash("Non puoi eliminare te stesso o un utente che ha inserito spese")
    else:
        db().execute("delete from users where id=?", (id,))
        db().execute("delete from user_car where user_id=?", (id,))
        db().commit()
        flash("Utente eliminato")
    return redirect(url_for("users"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
