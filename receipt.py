"""Lettura degli scontrini: OCR con Tesseract e ricerca di importo, negozio e data nel testo."""
import re
import shutil
import subprocess
from datetime import date, timedelta


class OcrError(Exception):
    pass


# --- OCR ---
_lang = None
PSM = "6"          # blocco di testo uniforme: regge meglio le foto con sfondo e leggere rotazioni


def _language(exe):
    """"ita" se installato, altrimenti "eng"."""
    global _lang
    if _lang is None:
        try:
            have = subprocess.run([exe, "--list-langs"], capture_output=True, text=True, timeout=15).stdout.split()
        except (OSError, subprocess.SubprocessError):
            have = []
        _lang = "ita" if "ita" in have else "eng"
    return _lang


def ocr_image(path, timeout=60):
    exe = shutil.which("tesseract")
    if not exe:
        raise OcrError("Tesseract non è installato sul Raspberry. Esegui: sudo apt install tesseract-ocr tesseract-ocr-ita")
    try:
        p = subprocess.run([exe, path, "stdout", "-l", _language(exe), "--psm", PSM],
                           capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise OcrError("La lettura richiede troppo tempo: riprova con una foto più nitida")
    except OSError:
        raise OcrError("Impossibile avviare Tesseract")
    if p.returncode != 0:
        raise OcrError("Non riesco a leggere questa immagine")
    return p.stdout


# --- Importi ---
AMOUNT = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+|\d{1,5})\s?[.,]\s?(\d{2})(?!\d)(?!\s?%)")
DATE = re.compile(r"(?<!\d)(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4}|\d{2})(?!\d)")
MAX_CENTS = 10_000_000          # oltre 100.000 € è quasi sicuramente un errore di lettura


def amounts(line):
    out = []
    for m in AMOUNT.finditer(line):
        c = int(m.group(1).replace(".", "")) * 100 + int(m.group(2))
        if 0 < c < MAX_CENTS:
            out.append(c)
    return out


def _kw(line):
    """Riga in maiuscolo con sole lettere (0->O, 1->I, 5->S), per riconoscere le parole chiave anche se lette male."""
    up = line.upper().translate(str.maketrans("015", "OIS"))
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z ]", " ", up)).strip()


EXCL = re.compile(r"SUB ?TOTALE|PARZIALE|\bIVA\b|IMPONIBILE|IMPOSTA|RESTO|ARTICOL|PEZZI|\bPZ\b|SCONT|ARROTOND|PUNTI|"
                  r"CONTANTI|\bCASH\b")
P1 = re.compile(r"TOTALE COMPLESSIVO|TOTALE DA PAGARE|DA PAGARE|TOTALE EUR|TOTALE IMPORTO|TOTALE DOVUTO")
P2 = re.compile(r"\bTOTALE\b|\bTOTAL\b|\bTOT\b")
P3 = re.compile(r"IMPORTO PAGATO|PAGAMENTO ELETTRONICO|PAGAMENTO|PAGATO|CARTA|BANCOMAT|\bPOS\b|ELETTRONICO")
VOUCHER = re.compile(r"BUON[OI] PASTO|TICKET|EDENRED|DAY ?UP|PELLEGRINI|BUONO PASTO")


def _amount_on(lines, i):
    """Importo della riga i (l'ultimo a destra) o, se manca, quello della riga seguente se contiene solo cifre."""
    a = amounts(lines[i])
    if a:
        return a[-1]
    if i + 1 < len(lines):
        nxt = lines[i + 1]
        if len(re.findall(r"[A-Za-z]", nxt)) <= 3 and len(amounts(nxt)) == 1:
            return amounts(nxt)[0]
    return None


def find_total(lines):
    """Ritorna (totale in centesimi, nota). Cerca per parola chiave, dalla più alla meno affidabile."""
    kws = [_kw(l) for l in lines]
    for pat in (P1, P2, P3):
        found = [_amount_on(lines, i) for i, k in enumerate(kws) if pat.search(k) and not EXCL.search(k)]
        found = [x for x in found if x]
        if found:
            return max(found), ""
    # nessuna parola chiave: l'importo più alto tra le righe che non sono date o resto/IVA
    cands = [a for i, l in enumerate(lines) if not EXCL.search(kws[i]) and not DATE.search(l) for a in amounts(l)]
    if cands:
        return max(cands), "Totale non riconosciuto: ho preso l'importo più alto, controllalo"
    return None, "Non trovo l'importo: scrivilo a mano"


def find_voucher(lines, total):
    """Parte pagata con buoni pasto, se indicata e inferiore al totale."""
    for i, l in enumerate(lines):
        if VOUCHER.search(_kw(l)):
            a = _amount_on(lines, i)
            if a and total and 0 < a < total:
                return a
    return 0


# --- Data ---
def find_date(text, today=None):
    today = today or date.today()
    for m in DATE.finditer(text):
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        y += 2000 if y < 100 else 0
        try:
            dt = date(y, mo, d)
        except ValueError:
            continue
        if today - timedelta(days=400) <= dt <= today + timedelta(days=1):
            return dt.isoformat()
    return None


# --- Negozio ---
CHAINS = [
    ("Conad", r"CONAD"), ("Coop", r"\bCOOP\b|COOP ALLEANZA|UNICOOP|IPERCOOP"), ("Esselunga", r"ESSELUNGA"),
    ("Lidl", r"\bLIDL\b"), ("Eurospin", r"EUROSPIN"), ("Carrefour", r"CARREFOUR"), ("Penny", r"\bPENNY\b"),
    ("Despar", r"DESPAR|EUROSPAR|INTERSPAR|\bSPAR\b"), ("Pam", r"\bPAM\b|PANORAMA"), ("Iper", r"\bIPER\b"),
    ("Tigros", r"TIGROS"), ("Famila", r"FAMILA"), ("Aldi", r"\bALDI\b"), ("Todis", r"TODIS"),
    ("Bennet", r"BENNET"), ("Simply", r"SIMPLY"), ("MD", r"\bDISCOUNT MD\b|\bMD S ?P ?A\b"),
    ("IKEA", r"\bIKEA\b"), ("Decathlon", r"DECATHLON"), ("Leroy Merlin", r"LEROY ?MERLIN"),
    ("Mediaworld", r"MEDIA ?WORLD"), ("Unieuro", r"UNIEURO"), ("OVS", r"\bOVS\b"), ("Zara", r"\bZARA\b"),
    ("Acqua e Sapone", r"ACQUA E SAPONE"), ("Sephora", r"SEPHORA"), ("Douglas", r"\bDOUGLAS\b"),
    ("McDonald's", r"MC ?DONALD"), ("Burger King", r"BURGER ?KING"), ("Autogrill", r"AUTOGRILL"),
    ("Starbucks", r"STARBUCKS"), ("Q8", r"\bQ8\b"), ("Eni", r"\bENI\b|ENI STATION"), ("Tamoil", r"TAMOIL"),
    ("Esso", r"\bESSO\b"), ("Shell", r"\bSHELL\b"), ("Farmacia", r"FARMACIA"), ("Tabaccheria", r"TABACCHERIA|TABACCHI"),
]
NOT_NAME = re.compile(r"\bVIA\b|\bV LE\b|VIALE|PIAZZA|\bP ?IVA\b|\bC ?F\b|\bTEL\b|\bFAX\b|WWW|@|SCONTRINO|DOCUMENTO|"
                      r"COMMERCIALE|FISCALE|\bCASSA\b|OPERATORE|\bCAP\b|VENDITA|PRESTAZIONE|ARRIVEDERCI|GRAZIE|"
                      r"\bN\b ?\d|^\d")
LEGAL = re.compile(r"\s+(S\.? ?R\.? ?L\.?S?|S\.? ?P\.? ?A|S\.? ?N\.? ?C|S\.? ?A\.? ?S|SOC\b.*|SOCIETA\b.*|COOP\b.*|A ?R ?L)\.?\s*$", re.I)


def find_store(lines):
    head = [l for l in lines if l.strip()][:12]
    up = "\n".join(h.upper() for h in head)
    best = None
    for name, pat in CHAINS:
        for variant in (up, up.replace("0", "O")):
            m = re.search(pat, variant)
            if m and (best is None or m.start() < best[0]):
                best = (m.start(), name)
    if best:
        return best[1]
    for l in head[:8]:                     # nessun nome noto: la prima riga che sembra un'insegna
        k = _kw(l)
        letters = len(re.findall(r"[A-Za-z]", l))
        if letters < 3 or letters < 0.6 * len(l.replace(" ", "")) or NOT_NAME.search(k) or len(amounts(l)):
            continue
        name = LEGAL.sub("", re.sub(r"[^\w&'’ .-]", " ", l).strip()).strip(" .-")
        name = re.sub(r"\s+", " ", name)
        if len(name) >= 3:
            return name.title()[:40]
    return None


# --- Risultato ---
def parse(text, today=None):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    total, note = find_total(lines)
    voucher = find_voucher(lines, total)
    return dict(total=total, voucher=voucher, store=find_store(lines), date=find_date(text, today), note=note)


def cents(c):
    return ("%.2f" % (c / 100)).replace(".", ",")
