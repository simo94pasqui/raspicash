"""Importazione da Excel.

Formati: "spese" (foglio Spese dell'esportazione), "viaggio" (elenco semplice delle spese di un
viaggio) e "rifornimenti" (foglio Rifornimenti). Due fasi:
  1. make_plan() legge il file e prepara un piano (nuove, duplicate, da completare, errori) senza scrivere;
  2. apply_plan() scrive il piano sulla connessione data; il commit spetta a chi chiama.
Non dipende da Flask: lavora su una connessione sqlite3 con row_factory sqlite3.Row.
"""
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from itertools import islice

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

MAX_ROWS = 5000


class ImportFileError(Exception):
    """File non leggibile o formato sconosciuto: blocca tutta l'importazione."""


class RowError(Exception):
    """Errore su una riga: la riga viene saltata, le altre vengono importate.
        fields: campi del foglio da correggere (usati dalla procedura guidata).
    """

    def __init__(self, msg, *fields):
        super().__init__(msg)
        self.fields = [f for f in fields if f]


# --- Lettura dei valori ---
def norm(s):
    """Minuscolo, senza accenti né simboli ("Prezzo/L" -> "prezzol")."""
    s = unicodedata.normalize("NFKD", str(s if s is not None else ""))
    return re.sub(r"[^a-z0-9]", "", "".join(c for c in s if not unicodedata.combining(c)).lower())


def clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def p_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 < v < 80000:
        return from_excel(v).date()
    s = clean(v)
    if not s:
        raise RowError("data mancante", "date")
    s = s.split(" ")[0].split("T")[0]
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise RowError(f"data non valida «{s}»", "date")


def _dec(v):
    if v is None:
        return None
    if isinstance(v, bool):
        raise InvalidOperation
    if isinstance(v, int):
        return Decimal(v)
    if isinstance(v, float):
        return Decimal(repr(v))
    s = str(v).strip().replace("€", "").replace("\xa0", "").replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    return Decimal(s)


def p_num(v, scale, what, field=None):
    """Numero opzionale come intero scalato (None se vuoto). Accetta virgola o punto."""
    try:
        d = _dec(v)
        if d is None:
            return None
        if not d.is_finite():
            raise InvalidOperation
        return int((d * scale).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError):
        raise RowError(f"{what} non valido «{clean(v)}»", field)


def p_yes(v, what, field=None):
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    n = norm(v)
    if n in ("", "no", "n", "false", "falso", "0"):
        return False
    if n in ("si", "s", "yes", "y", "true", "vero", "x", "1"):
        return True
    raise RowError(f"{what}: valore «{clean(v)}» non valido (usa Sì o No)", field)


def p_paidby(v):
    n = norm(v)
    if n in ("", "io", "me", "mio", "mia", "pagoio"):
        return False
    if n in ("azienda", "company", "ditta", "societa", "lavoro"):
        return True
    raise RowError(f"«Pagata da»: valore «{clean(v)}» non valido (usa Io o Azienda)", "paidby")


# --- Calcolo del rifornimento (usato anche dal modulo di inserimento) ---
def calc_fuel(total, liters, price):
    """Completa importo (centesimi), litri e prezzo/L (millesimi): bastano due valori su tre,
        oppure il solo importo. Ritorna (totale, litri, prezzo), o None se mancano importo e litri.
    """
    liters, price = liters or None, price or None
    if liters and total > 0 and not price:
        price = int(total * 10000 / liters + 0.5)
    elif liters and price and total == 0:
        total = int(liters * price / 10000 + 0.5)
    elif price and total > 0 and not liters:
        liters = int(total * 10000 / price + 0.5)
    if total <= 0 and not liters:
        return None
    return total, liters, price


# --- Lettura del foglio ---
COLS = {
    "date": ["data"], "description": ["descrizione", "desc"], "category": ["categoria"],
    "amount": ["importo"], "voucher": ["buono", "importobuono"], "private": ["privata", "privato"],
    "car": ["auto"], "person": ["persona"], "trip": ["viaggio"],
    "company": ["pagatadallazienda"], "paidby": ["pagatada"], "loan": ["prestito*"],
    # rifornimenti
    "fill_no": ["nrifornimento", "numerorifornimento", "nrif"], "odometer": ["kmtotali", "kmtotaliauto"],
    "liters": ["litri"], "price": ["prezzol", "prezzoalitro"], "fuel_amount": ["importototale", "importo"],
    "trip_km": ["kmcolpienoprecedente"], "claimed": ["kmdichiarati", "kmdichiaratidallauto"],
}
REQUIRED = {"spese": ["date", "description", "category", "amount"],
            "viaggio": ["date", "description", "amount"],
            "rifornimenti": ["date", "car"]}
LABELS = {"date": "Data", "description": "Descrizione", "category": "Categoria", "amount": "Importo",
          "car": "Auto"}


def _field_of(header, fields):
    n = norm(header)
    if not n:
        return None
    for f in fields:
        for a in COLS[f]:
            if (a.endswith("*") and n.startswith(a[:-1])) or n == a:
                return f
    return None


def read_table(fileobj, kind):
    """Ritorna dict(cols={campo: indice}, rows=[(riga Excel, valori)], sheet=nome del foglio)."""
    try:
        wb = load_workbook(fileobj, read_only=True, data_only=True)
    except Exception:
        raise ImportFileError("Il file non è un Excel (.xlsx) valido")
    wanted = "rifornimenti" if kind == "rifornimenti" else "spese"
    ws = next((w for w in wb.worksheets if norm(w.title) == wanted), None) or wb.worksheets[0]
    allrows = list(islice(ws.iter_rows(values_only=True), MAX_ROWS + 30))
    wb.close()
    fields = (["date", "car", "fill_no", "odometer", "liters", "price", "fuel_amount", "trip_km", "claimed"]
              if kind == "rifornimenti" else
              ["date", "description", "category", "amount", "voucher", "private", "car", "person", "trip",
               "company", "paidby", "loan"])
    best = None
    for i, r in enumerate(allrows[:10]):
        cols = {}
        for j, h in enumerate(r):
            f = _field_of(h, fields)
            if f and f not in cols:
                cols[f] = j
        if all(f in cols for f in REQUIRED[kind]):
            best = (i, cols)
            break
        if best is None or len(cols) > len(best[1]):
            best = (None, cols)
    if best is None or best[0] is None:
        found = best[1] if best else {}
        miss = [LABELS.get(f, f) for f in REQUIRED[kind] if f not in found]
        raise ImportFileError("Nel foglio «%s» mancano le colonne: %s" % (ws.title, ", ".join(miss) or "intestazioni"))
    hdr, cols = best
    if kind == "rifornimenti" and "liters" not in cols and "fuel_amount" not in cols:
        raise ImportFileError("Nel foglio «%s» serve almeno la colonna Litri oppure Importo totale" % ws.title)
    rows = [(i + 1, r) for i, r in enumerate(allrows) if i > hdr]
    rows = [(n, r) for n, r in rows if any(v is not None and str(v).strip() != "" for v in r)]
    if len(rows) > MAX_ROWS:
        raise ImportFileError("Troppe righe: al massimo %d per volta" % MAX_ROWS)
    if not rows:
        raise ImportFileError("Il foglio «%s» non contiene righe da importare" % ws.title)
    return dict(cols=cols, rows=rows, sheet=ws.title)


# --- Anagrafiche ---
class Lookup:
    """Ricerca per nome ignorando maiuscole, accenti e simboli; preferisce le voci non archiviate."""

    def __init__(self, conn, table):
        self.table, self.by = table, {}
        for r in conn.execute(f"select * from {table} order by archived, id"):
            self.by.setdefault(norm(r["name"]), dict(r))

    def find(self, name):
        return self.by.get(norm(name))

    def special(self, flag):
        return next((r for r in self.by.values() if r.get(flag)), None)


KINDS = {"categorie": "categories", "auto": "cars", "persone": "people", "viaggi": "trips"}


class Plan:
    """Esito della lettura del file: una voce per riga più i nomi da creare o riattivare."""

    def __init__(self, kind):
        self.kind, self.items = kind, []
        self.missing = {k: {} for k in KINDS}         # Nomi da creare: tipo -> {nome normalizzato: (nome, righe)}
        self.react = {k: {} for k in KINDS}           # Voci archiviate da riattivare
        self.ignored = 0                              # Righe di totale ignorate
        self.new_trip = None                          # Viaggio nuovo scelto nel modulo (import viaggio)
        self.excluded = 0                             # Righe escluse dall'utente
        self.used = set()                             # Spese già abbinate a una riga del file

    def add(self, row, status, msg="", warn=None, needs=None, label=None, d=None, react=None, fix=None, raw=None, part=None):
        it = dict(row=row, status=status, msg=msg, warn=warn or [], needs=needs or [], label=label or {},
                  d=d or {}, react=react or [], fix=fix or [], raw=raw or {}, part=part or {})
        self.items.append(it)
        return it

    def finish(self):
        """Raccoglie i nomi da creare o riattivare, contando solo le righe che verranno scritte."""
        for it in self.items:
            if it["status"] not in ("new", "complete"):
                continue
            for kind, name in it["needs"]:
                old = self.missing[kind].get(norm(name))
                self.missing[kind][norm(name)] = (old[0] if old else name, (old[1] if old else 0) + 1)
            for kind, name in it["react"]:
                self.react[kind][norm(name)] = name
        return self

    @property
    def counts(self):
        return Counter(i["status"] for i in self.items)


def _use(lk, kind, name, needs, react):
    """Cerca per nome: se manca lo aggiunge a needs, se è archiviato a react."""
    r = lk.find(name)
    if r is None:
        needs.append((kind, name))
    elif r.get("archived"):
        react.append((kind, r["name"]))
    return r


def _eur(c):
    return ("%.2f" % (c / 100)).replace(".", ",")


def _label(d, desc, total):
    return dict(date=d.isoformat() if d else "", desc=desc, total=total)


# --- Piano: spese e viaggio ---
def plan_expenses(conn, uid, table, mode="spese", trip_opt=None):
    plan = Plan(mode)
    cols = table["cols"]
    cats, cars, people, trips = (Lookup(conn, t) for t in ("categories", "cars", "people", "trips"))
    trip_cat = cats.special("is_trip")
    trip_by_id = {r["id"]: r for r in trips.by.values()}
    new_trip_work = defaultdict(bool)
    fixed_trip, fixed_react = None, []
    if mode == "viaggio":
        if trip_opt.get("id"):
            r = trip_by_id.get(trip_opt["id"])
            if not r:
                raise ImportFileError("Viaggio non trovato")
        else:
            name = clean(trip_opt.get("new"))[:60]
            if not name:
                raise ImportFileError("Indica il nome del nuovo viaggio")
            r = trips.find(name)
            if r is None:
                fixed_trip = dict(id=None, name=name, work=1 if trip_opt.get("work") else 0)
                plan.new_trip = name
        if r is not None:
            fixed_trip = dict(id=r["id"], name=r["name"])
            if r.get("archived"):
                fixed_react.append(("viaggi", r["name"]))

    def get(vals, f):
        i = cols.get(f)
        return vals[i] if i is not None and i < len(vals) else None

    parsed = []
    for rown, vals in table["rows"]:
        d0 = get(vals, "date")
        if isinstance(d0, str) and norm(d0).startswith(("totale", "totali", "riepilogo")):
            plan.ignored += 1
            continue
        warn, needs, react = [], [], []
        d = desc = None
        total = 0
        try:
            d = p_date(d0)
            desc = clean(get(vals, "description"))[:120]
            if not desc:
                raise RowError("descrizione mancante", "description")
            amount = p_num(get(vals, "amount"), 100, "importo", "amount")
            voucher = p_num(get(vals, "voucher"), 100, "buono", "voucher")
            if amount is None and voucher is None:
                raise RowError("importo mancante", "amount", "voucher")
            amount, voucher = amount or 0, voucher or 0
            total = amount + voucher
            private = p_yes(get(vals, "private"), "Privata", "private")
            company = p_yes(get(vals, "company"), "Pagata dall'azienda", "company")
            if get(vals, "paidby") is not None:
                company = company or p_paidby(get(vals, "paidby"))
            # categoria
            if mode == "viaggio":
                cat, cname = trip_cat, trip_cat["name"]
            else:
                cname = clean(get(vals, "category"))[:40]
                if not cname:
                    raise RowError("categoria mancante", "category")
                cat = _use(cats, "categorie", cname, needs, react)
            is_fuel, is_car, is_trip = (bool(cat and cat.get(k)) for k in ("is_fuel", "is_car", "is_trip"))
            if is_fuel:
                plan.add(rown, "skip", "Benzina: usa «Importa rifornimenti»", label=_label(d, desc, total))
                continue
            if is_car:
                private = False
            # auto
            car = None
            if is_car:
                carname = clean(get(vals, "car"))[:60]
                if not carname:
                    raise RowError("categoria Auto: indica l'auto", "car")
                r = _use(cars, "auto", carname, needs, react)
                car = dict(id=r["id"] if r else None, name=r["name"] if r else carname)
            elif clean(get(vals, "car")):
                warn.append("colonna Auto ignorata: la categoria non è Auto")
            # viaggio
            trip = None
            if is_trip:
                if mode == "viaggio":
                    trip = dict(fixed_trip)
                    react += fixed_react
                    work = bool(trip_by_id[trip["id"]]["is_work"]) if trip["id"] else bool(trip.get("work"))
                else:
                    tname = clean(get(vals, "trip"))[:60]
                    if not tname:
                        raise RowError("categoria Viaggi: indica il viaggio", "trip")
                    r = _use(trips, "viaggi", tname, needs, react)
                    trip = dict(id=r["id"], name=r["name"]) if r else dict(id=None, name=tname)
                    work = bool(r["is_work"]) if r else None      # viaggio nuovo: lo decide il file
                if work is not None and company and not work:
                    warn.append("«Azienda» ignorato: il viaggio non è di lavoro")
                    company = False
                if mode != "viaggio" and trip["id"] is None:
                    new_trip_work[norm(trip["name"])] |= company
            else:
                company = False
                if mode != "viaggio" and clean(get(vals, "trip")):
                    warn.append("colonna Viaggio ignorata: la categoria non è Viaggi")
            # prestito
            loan = None
            pname = clean(get(vals, "person"))[:40]
            lamt = p_num(get(vals, "loan"), 100, "prestito", "loan")
            if lamt or pname:
                if is_car or company:
                    warn.append("prestito ignorato (non previsto per spese Auto o pagate dall'azienda)")
                elif lamt:
                    if not pname:
                        raise RowError("prestito senza persona", "person")
                    if abs(lamt) > total:
                        raise RowError("il prestito supera il totale della spesa", "loan", "amount", "voucher")
                    r = _use(people, "persone", pname, needs, react)
                    loan = dict(id=r["id"] if r else None, name=r["name"] if r else pname, amount=lamt)
                else:
                    warn.append("Persona ignorata: manca l'importo del prestito")
        except RowError as e:
            plan.add(rown, "error", str(e), label=_label(d, desc, total), fix=e.fields,
                     raw={f: clean(get(vals, f)) for f in FIELDS}, part=_part(lambda f: get(vals, f), "spese"))
            continue
        dd = dict(date=d.isoformat(), desc=desc, amount=amount, voucher=voucher, private=int(bool(private)),
                  cat=dict(id=cat["id"] if cat else None, name=cat["name"] if cat else cname),
                  car=car, trip=trip, company=int(bool(company)), loan=loan)
        parsed.append(plan.add(rown, "new", warn=warn, needs=needs, react=react,
                               label=_label(d, desc, total), d=dd))
    # viaggi nuovi da un file Spese: di lavoro se almeno una riga è pagata dall'azienda
    for it in parsed:
        t = it["d"]["trip"]
        if t and t["id"] is None and "work" not in t:
            t["work"] = 1 if new_trip_work[norm(t["name"])] else 0
    _mark_duplicates(conn, uid, plan, parsed, trips)
    _soft_match(conn, uid, plan, table.get("force") or (), mode)
    return plan.finish()


def _soft(fn, *a):
    """Valore di fn(*a), o None se non valido."""
    try:
        return fn(*a)
    except (RowError, ValueError, ArithmeticError):
        return None


def _part(g, mode):
    """Valori validi di una riga in errore, da confrontare con le spese già salvate."""
    d = _soft(p_date, g("date"))
    d = d.isoformat() if d else None
    if mode == "rifornimenti":
        return dict(date=d, amount=_soft(p_num, g("fuel_amount"), 100, "importo"),
                    liters=_soft(p_num, g("liters"), 1000, "litri"), odometer=_soft(p_num, g("odometer"), 1, "km"))
    return dict(date=d, desc=norm(clean(g("description"))), amount=_soft(p_num, g("amount"), 100, "importo"),
                voucher=_soft(p_num, g("voucher"), 100, "buono"))


def _soft_match(conn, uid, plan, force, mode):
    """Cerca tra le spese salvate quelle che corrispondono alle righe in errore o alle righe nuove
        non riconosciute come duplicate (ad es. categoria cambiata dopo un primo import).
        Si confrontano solo i campi validi e non in errore: devono coincidere tutti, e almeno due.
        Le righe abbinate diventano "dup" con soft=True, così l'utente può dire che sono diverse.
        force: righe da non abbinare.
    """
    fuel = mode == "rifornimenti"
    rows = [i for i in plan.items if i["row"] not in force and
            ((i["status"] == "error" and i["part"]) or (i["status"] == "new" and not fuel))]
    if not rows:
        return
    if fuel:
        ex = [dict(r) for r in conn.execute(
            "select e.id, e.date, e.amount, f.liters, f.odometer from expenses e "
            "join fuel_logs f on f.expense_id=e.id where e.private=0 or e.owner_id=?", (uid,))]
    else:
        ex = [dict(r) for r in conn.execute(
            "select e.id, e.date, e.description, e.amount, e.voucher, c.name cat from expenses e "
            "join categories c on c.id=e.category_id where e.private=0 or e.owner_id=?", (uid,))]
        for e in ex:
            e["nd"] = norm(e["description"])
    for it in rows:
        is_err = it["status"] == "error"
        bad, p, c = (set(it["fix"]) if is_err else set()), it["part"], {}
        if fuel:
            for k, f in (("date", "date"), ("amount", "fuel_amount"), ("liters", "liters"), ("odometer", "odometer")):
                if p.get(k) and f not in bad:
                    c[k] = p[k]
        elif is_err:
            if p["date"] and "date" not in bad:
                c["date"] = p["date"]
            if p["desc"] and "description" not in bad:
                c["desc"] = p["desc"]
            if p["amount"] is not None and "amount" not in bad:
                c["money"] = (p["amount"], None if "voucher" in bad else (p["voucher"] or 0))
        else:
            d = it["d"]
            c = dict(date=d["date"], desc=norm(d["desc"]), money=(d["amount"], d["voucher"]))
        if len(c) < 2:
            continue

        def hit(e):
            if "date" in c and e["date"] != c["date"]:
                return False
            if fuel:
                return ("amount" not in c or e["amount"] == c["amount"]) and \
                    ("odometer" not in c or e["odometer"] == c["odometer"]) and \
                    ("liters" not in c or (e["liters"] is not None and abs(e["liters"] - c["liters"]) <= 20))
            if "desc" in c and e["nd"] != c["desc"]:
                return False
            if "money" in c:
                a, v = c["money"]
                return e["amount"] == a and (v is None or e["voucher"] == v)
            return True

        e = next((e for e in ex if e["id"] not in plan.used and hit(e)), None)
        if not e:
            continue
        plan.used.add(e["id"])
        day = "%s/%s/%s" % (e["date"][8:], e["date"][5:7], e["date"][:4])
        what = ("rifornimento del %s · %s €" % (day, _eur(e["amount"]))) if fuel else \
            "«%s» del %s · %s €, categoria %s" % (e["description"], day, _eur(e["amount"] + e["voucher"]), e["cat"])
        it["status"], it["soft"] = "dup", True
        it["msg"] = "già presente: " + what
        if bad:
            it["msg"] += " (errore ignorato: %s)" % ", ".join(sorted({FIELDS[k][0] for k in bad if k in FIELDS}))


def _mark_duplicates(conn, uid, plan, items, trips):
    """Una riga è duplicata se esiste già una spesa visibile con stessa data, descrizione,
        categoria, importi e viaggio. Si contano le occorrenze: due righe uguali nel file restano due.
    """
    if not items:
        return
    dmin, dmax = min(i["d"]["date"] for i in items), max(i["d"]["date"] for i in items)
    have = defaultdict(list)
    for r in conn.execute(
            "select e.id, e.date, e.description, e.category_id, e.amount, e.voucher, coalesce(te.trip_id,0) t "
            "from expenses e left join trip_expenses te on te.expense_id=e.id "
            "where e.date between ? and ? and (e.private=0 or e.owner_id=?)", (dmin, dmax, uid)):
        have[(r["date"], norm(r["description"]), r["category_id"], r["amount"], r["voucher"], r["t"])].append(r["id"])
    for it in items:
        d = it["d"]
        if d["cat"]["id"] is None or it["needs"] and any(k == "viaggi" for k, _ in it["needs"]):
            continue
        key = (d["date"], norm(d["desc"]), d["cat"]["id"], d["amount"], d["voucher"],
               (d["trip"] or {}).get("id") or 0)
        if have[key]:
            plan.used.add(have[key].pop())
            it["status"], it["msg"] = "dup", "già presente"


# --- Piano: rifornimenti ---
def plan_fuel(conn, uid, table):
    plan = Plan("rifornimenti")
    cols = table["cols"]
    cars = Lookup(conn, "cars")

    def get(vals, f):
        i = cols.get(f)
        return vals[i] if i is not None and i < len(vals) else None

    parsed = []
    for rown, vals in table["rows"]:
        d0 = get(vals, "date")
        if isinstance(d0, str) and norm(d0).startswith(("totale", "totali", "riepilogo")):
            plan.ignored += 1
            continue
        d, warn, needs, react, total = None, [], [], [], 0
        try:
            d = p_date(d0)
            carname = clean(get(vals, "car"))[:60]
            if not carname:
                raise RowError("auto mancante", "car")
            no = p_num(get(vals, "fill_no"), 1, "n. rifornimento", "fill_no")
            odo = p_num(get(vals, "odometer"), 1, "km totali", "odometer")
            liters = p_num(get(vals, "liters"), 1000, "litri", "liters")
            price = p_num(get(vals, "price"), 1000, "prezzo/L", "price")
            amount = p_num(get(vals, "fuel_amount"), 100, "importo", "fuel_amount")
            trip_km = p_num(get(vals, "trip_km"), 1, "km col pieno precedente", "trip_km")
            claimed = p_num(get(vals, "claimed"), 1, "km dichiarati", "claimed")
            neg = [k for k, x in (("fill_no", no), ("odometer", odo), ("liters", liters), ("price", price),
                                  ("fuel_amount", amount), ("trip_km", trip_km), ("claimed", claimed))
                   if x is not None and x < 0]
            if neg:
                raise RowError("valori negativi non ammessi", *neg)
            c = calc_fuel(amount or 0, liters, price)
            if not c:      # Senza importo né litri non è un errore: si salva a zero, da completare dopo
                total, liters, price = 0, None, price or None
                warn.append("né importo né litri indicati: salvato a zero, completalo dopo dalla modifica")
            else:
                total, liters, price = c
                if not total:
                    warn.append("importo non indicato: resta a zero")
            r = _use(cars, "auto", carname, needs, react)
        except RowError as e:
            plan.add(rown, "error", str(e), label=_label(d, "Rifornimento", total), fix=e.fields,
                     raw={f: clean(get(vals, f)) for f in FIELDS}, part=_part(lambda f: get(vals, f), "rifornimenti"))
            continue
        what = carname + (" · %s L" % ("%.2f" % (liters / 1000)).replace(".", ",") if liters else "")
        dd = dict(date=d.isoformat(), car=dict(id=r["id"] if r else None, name=r["name"] if r else carname),
                  no=no or None, odo=odo or None, liters=liters, price=price, trip=trip_km or None,
                  claimed=claimed or None, amount=total, explicit_total=amount is not None)
        parsed.append(plan.add(rown, "new", warn=warn, needs=needs, react=react,
                               label=_label(d, what, total), d=dd))
    if not parsed:
        _soft_match(conn, uid, plan, table.get("force") or (), "rifornimenti")
        return plan.finish()
    dmin, dmax = min(i["d"]["date"] for i in parsed), max(i["d"]["date"] for i in parsed)
    byday = defaultdict(list)
    for r in conn.execute(
            "select e.id, e.date, e.amount+e.voucher tot, f.expense_id fid, f.car_id, f.fill_no, f.odometer, "
            "f.liters, f.price_l, f.trip_km, f.claimed_km from expenses e "
            "left join fuel_logs f on f.expense_id=e.id "
            "where e.category_id in (select id from categories where is_fuel=1) and e.date between ? and ?",
            (dmin, dmax)):
        byday[r["date"]].append(dict(r))
    used = set()
    for it in parsed:
        d = it["d"]
        cid = d["car"]["id"]
        if cid is None:
            continue
        cand = [x for x in byday.get(d["date"], []) if x["id"] not in used and (x["car_id"] in (None, cid))]
        if d["amount"]:
            cand = [x for x in cand if abs(x["tot"] - d["amount"]) <= 1]
        if len(cand) > 1 and d["no"]:
            c2 = [x for x in cand if x["fill_no"] == d["no"]]
            cand = c2 if len(c2) == 1 else cand
        if len(cand) > 1 and d["odo"]:
            c2 = [x for x in cand if x["odometer"] == d["odo"]]
            cand = c2 if len(c2) == 1 else cand
        if not cand:
            continue
        if len(cand) > 1:
            it["status"], it["msg"] = "ambiguous", "più rifornimenti possibili nello stesso giorno: controllali a mano"
            continue
        x = cand[0]
        used.add(x["id"])
        add = {}
        if x["fid"] is None:
            add = dict(new_log=True)
        else:
            if x["odometer"] is None and d["odo"]:
                add["odometer"] = d["odo"]
            if x["liters"] is None and d["liters"]:
                add["liters"] = d["liters"]
                if x["price_l"] is None and d["price"]:
                    add["price_l"] = d["price"]
            if x["claimed_km"] is None and d["claimed"]:
                add["claimed_km"] = d["claimed"]
            if x["trip_km"] is None and x["odometer"] is None and "odometer" not in add and d["trip"]:
                add["trip_km"] = d["trip"]
            if x["fill_no"] is None and d["no"]:
                add["fill_no"] = d["no"]
            if x["liters"] and d["liters"] and abs(x["liters"] - d["liters"]) > 20:
                it["warn"].append("litri diversi da quelli già salvati: non modificati")
            if x["odometer"] and d["odo"] and x["odometer"] != d["odo"]:
                it["warn"].append("km totali diversi da quelli già salvati: non modificati")
        if add:
            it["status"] = "complete"
            it["msg"] = "completa la spesa Benzina esistente" + (": " + ", ".join(
                {"odometer": "km", "liters": "litri", "price_l": "prezzo", "claimed_km": "km dichiarati",
                 "trip_km": "km col pieno", "fill_no": "n.", "new_log": "dettagli"}[k] for k in add) if add else "")
            d["target"], d["add"] = x["id"], add
        else:
            it["status"], it["msg"] = "dup", "già presente"
    plan.used |= used
    _soft_match(conn, uid, plan, table.get("force") or (), "rifornimenti")
    return plan.finish()


# --- Scrittura ---
class Ensurer:
    """Trova o crea categorie, auto, persone e viaggi mentre si scrive il piano."""

    def __init__(self, conn):
        self.conn, self.lk, self.made = conn, {}, defaultdict(list)

    def _lookup(self, kind):
        if kind not in self.lk:
            self.lk[kind] = Lookup(self.conn, KINDS[kind])
        return self.lk[kind]

    def get(self, kind, name, work=0):
        lk = self._lookup(kind)
        r = lk.find(name)
        if r:
            if r.get("archived"):
                self.conn.execute(f"update {KINDS[kind]} set archived=0 where id=?", (r["id"],))
                r["archived"] = 0
            return r["id"]
        name = name[:60]
        if kind == "viaggi":
            cur = self.conn.execute("insert into trips(name,is_work) values(?,?)", (name, 1 if work else 0))
        else:
            cur = self.conn.execute(f"insert into {KINDS[kind]}(name) values(?)", (name,))
        lk.by[norm(name)] = dict(id=cur.lastrowid, name=name, archived=0)
        self.made[kind].append(name)
        return cur.lastrowid

    def ref(self, kind, ref, work=0):
        """Id di una voce del piano (già nota o da creare); la riattiva se era archiviata."""
        if ref["id"]:
            self.conn.execute(f"update {KINDS[kind]} set archived=0 where id=? and archived=1", (ref["id"],))
            return ref["id"]
        return self.get(kind, ref["name"], work)


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _next_no(conn, car):
    r = conn.execute("select max(fill_no) m, count(*) n from fuel_logs where car_id=?", (car,)).fetchone()
    return (r["m"] or r["n"]) + 1


def apply_plan(conn, uid, plan, create_missing=True):
    """Scrive il piano (senza commit) e ritorna le statistiche."""
    st = Counter()
    ens = Ensurer(conn)
    last_date = last_car = None
    for it in plan.items:
        if it["status"] not in ("new", "complete"):
            continue
        if it["needs"] and not create_missing:
            st["senza_nomi"] += 1
            continue
        d = it["d"]
        if plan.kind == "rifornimenti":
            car = ens.ref("auto", d["car"])
            last_car = car
            if it["status"] == "complete":
                add = d["add"]
                if add.get("new_log"):
                    conn.execute("insert into fuel_logs(expense_id,car_id,fill_no,odometer,liters,price_l,trip_km,claimed_km)"
                                 " values(?,?,?,?,?,?,?,?)",
                                 (d["target"], car, d["no"] or _next_no(conn, car), d["odo"], d["liters"],
                                  d["price"], d["trip"], d["claimed"]))
                else:
                    sets = ",".join(f"{k}=?" for k in add)
                    conn.execute(f"update fuel_logs set {sets} where expense_id=?", (*add.values(), d["target"]))
                conn.execute("update expenses set updated_by=?, updated_at=? where id=?", (uid, _now(), d["target"]))
                st["completate"] += 1
            else:
                fcat = conn.execute("select id from categories where is_fuel=1 order by archived, id").fetchone()[0]
                eid = conn.execute(
                    "insert into expenses(date,description,category_id,amount,voucher,private,owner_id,created_at)"
                    " values(?,?,?,?,?,?,?,?)", (d["date"], "Rifornimento", fcat, d["amount"], 0, 0, uid, _now())).lastrowid
                conn.execute("insert into fuel_logs(expense_id,car_id,fill_no,odometer,liters,price_l,trip_km,claimed_km)"
                             " values(?,?,?,?,?,?,?,?)",
                             (eid, car, d["no"] or _next_no(conn, car), d["odo"], d["liters"], d["price"],
                              d["trip"], d["claimed"]))
                st["nuove"] += 1
            last_date = max(last_date or d["date"], d["date"])
            continue
        # spese / viaggio
        cat = ens.ref("categorie", d["cat"])
        eid = conn.execute(
            "insert into expenses(date,description,category_id,amount,voucher,private,owner_id,created_at)"
            " values(?,?,?,?,?,?,?,?)",
            (d["date"], d["desc"], cat, d["amount"], d["voucher"], d["private"], uid, _now())).lastrowid
        if d["car"]:
            conn.execute("insert into car_expenses(expense_id,car_id) values(?,?)",
                         (eid, ens.ref("auto", d["car"])))
        if d["trip"]:
            t = d["trip"]
            tid = ens.ref("viaggi", t, work=t.get("work", 0))
            conn.execute("insert into trip_expenses(expense_id,trip_id,company) values(?,?,?)",
                         (eid, tid, d["company"]))
        if d["loan"]:
            ln = d["loan"]
            conn.execute("insert into loans(expense_id,person_id,amount) values(?,?,?)",
                         (eid, ens.ref("persone", ln), ln["amount"]))
        st["nuove"] += 1
        last_date = max(last_date or d["date"], d["date"])
    st["creati"] = {k: v for k, v in ens.made.items() if v}
    st["last_date"], st["last_car"] = last_date, last_car
    return st


# --- Correzione guidata ---
# Campo -> (etichetta, tipo); i tipi presenti in KINDS offrono i nomi già esistenti
FIELDS = {
    "date": ("Data", "date"), "description": ("Descrizione", "text"), "category": ("Categoria", "categorie"),
    "amount": ("Importo (€)", "num"), "voucher": ("Buono (€)", "num"), "private": ("Privata", "yn"),
    "company": ("Pagata dall'azienda", "yn"), "paidby": ("Pagata da", "paid"),
    "car": ("Auto", "auto"), "trip": ("Viaggio", "viaggi"), "person": ("Persona", "persone"),
    "loan": ("Prestito (€)", "num"), "fill_no": ("N. rifornimento", "num"), "odometer": ("Km totali", "num"),
    "liters": ("Litri", "num"), "price": ("Prezzo al litro (€)", "num"), "fuel_amount": ("Importo totale (€)", "num"),
    "trip_km": ("Km col pieno precedente", "num"), "claimed": ("Km dichiarati dall'auto", "num"),
}


def apply_fixes(table, fixes, exclude):
    """Applica le correzioni ({riga Excel: {campo: valore}}) ed esclude le righe indicate.
        Se manca la colonna di un campo corretto la aggiunge in coda. Ritorna quante righe sono state escluse.
    """
    cols, rows = table["cols"], table["rows"]
    width = max([len(v) for _, v in rows] + [max(cols.values()) + 1 if cols else 0])
    for f in sorted({f for fx in fixes.values() for f in fx if f not in cols and f in FIELDS}):
        cols[f] = width
        width += 1
    out, excluded = [], 0
    for n, vals in rows:
        if n in exclude:
            excluded += 1
            continue
        fx = fixes.get(n)
        if fx:
            vals = list(vals) + [None] * (width - len(vals))
            for f, v in fx.items():
                if f in cols:
                    vals[cols[f]] = v
        out.append((n, vals))
    table["rows"] = out
    return excluded


# --- Punto d'ingresso ---
def make_plan(conn, uid, kind, fileobj, trip_opt=None, fixes=None, exclude=None, force=None):
    table = read_table(fileobj, kind)
    excluded = apply_fixes(table, fixes or {}, set(exclude or ()))
    table["force"] = set(force or ())
    if kind == "rifornimenti":
        plan = plan_fuel(conn, uid, table)
    else:
        plan = plan_expenses(conn, uid, table, mode=kind, trip_opt=trip_opt)
    plan.excluded = excluded
    return plan
