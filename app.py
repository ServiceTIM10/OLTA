import re
import tempfile
from pathlib import Path
from io import BytesIO
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup

try:
    import extract_msg
except Exception:
    extract_msg = None

FINAL_COLUMNS = [
    "OLTA MITTENTE",
    "DATA NEWSLETTER",
    "TIPOLOGIA",
    "COMPAGNIA PROMOSSA",
    "DESTINAZIONE/MERCATO/TRATTA",
    "PROMOZIONE",
    "OGGETTO EMAIL",
    "FILE ORIGINE",
    "TESTO ESTRATTO",
    "CONFIDENCE",
    "NOTE",
]

KNOWN_OLTAS = {
    "AFERRY": ["aferry", "anna aferry"],
    "ALLOFERRY": ["allo ferry", "alloferry"],
    "FERRYHOPPER": ["ferryhopper"],
    "NETFERRY": ["netferry"],
    "TRAGHETTIPER": ["traghettiper"],
    "TRAGHETTI.COM": ["traghetti.com"],
    "DIRECT FERRIES": ["direct ferries"],
    "LA CENTRALE DES FERRIES": ["la centrale des ferries"],
}

KNOWN_COMPANIES = [
    "GNV", "Grimaldi Lines", "Moby", "Tirrenia", "Corsica Ferries", "Sardinia Ferries",
    "Corsica Linea", "La Méridionale", "La Meridionale", "Tallink Silja Line",
    "Trasmed GLE", "Trasmed", "Brittany Ferries", "DFDS", "Irish Ferries", "P&O Ferries",
    "P&O", "Color Line", "Superfast Ferries", "Liberty Lines", "Siremar", "Stena Line",
    "Armas", "Balearia", "SNAV", "Jadrolinija", "Blue Star Ferries", "Seajets",
    "Minoan Lines", "ANEK Lines", "Viking Line", "Finnlines",
]

KNOWN_DESTINATIONS = [
    "Sicilia", "Sardegna", "Grecia", "Italia-Grecia", "Italia ↔ Grecia", "Italia Grecia",
    "Inghilterra", "Irlanda", "Baleari", "Isole Baleari", "Ibiza", "Maiorca", "Minorca",
    "Tunisia", "Algeria", "Marocco", "Francia-Marocco", "Francia - Marocco",
    "Francia Algérie", "France Algérie", "Danimarca", "Norvegia", "Danimarca ↔ Norvegia",
    "Mar Baltico", "Corsica", "Elba", "Isola d'Elba", "Isole Pelagie", "Lampedusa", "Linosa",
    "Porto Empedocle", "Nord Europa", "Mare d'Irlanda", "Manica", "Fiordi",
    "Almeria Nador", "Motril Nador", "Almeria-Melilla", "Motril-Melilla", "Almeria Oran",
    "Almeria Ghazouet", "Algéciras Tanger", "Algésiras Tanger", "Algéciras Ceuta",
    "Sète Tanger", "Sète Nador", "Barcelone Tanger", "Barcelone Nador", "Gênes Tanger",
    "Marseille", "Sete", "Sète", "Alger", "Bejaia", "Nador", "Tanger", "Ceuta", "Oran", "Ghazouet",
    "Albania", "Croazia", "Isole Cicladi", "Svezia", "Finlandia", "Estonia",
]

COMMON_LINES = {
    "Prenota ora", "Trova il miglior prezzo", "Cerca la tua tratta su TraghettiPer.",
    "Confronta le offerte di tutte le compagnie di navigazione.", "Buona navigazione,",
    "Il Team di TraghettiPer", "Perché viaggiare con noi", "Professionale", "Sicuro", "Veloce",
    "Dicono di noi", "Metodi di pagamento", "Annulla iscrizione", "Cancella iscrizione alla newsletter",
}

PROMO_RE = re.compile(
    r"(?i)(?:fino\s+al|fino\s+a|jusqu['’]à|up\s+to|sconto\s+del|sconto|remise|"
    r"risparmia(?:\s+fino\s+al|\s+il)?|offerta|promo|coupon|codice|voucher|bon\s+d[’']achat|"
    r"a\s+partire\s+da|da|dès|à\s+partir\s+de|partire\s+da|-\s?\d{1,2}\s?%|"
    r"\d{1,2}\s?%|\d+(?:[,.]\d+)?\s?€|veicolo\s+gratis|gratis)"
)
COUPON_CODE_RE = re.compile(r"(?i)(?:codice|coupon|voucher|bon\s+d[’']achat)")
ROUTE_HINT_RE = re.compile(r"(?i)(?:\b[A-ZÀ-Ý][a-zà-ÿ]+(?:\s*[↔\-–]\s*[A-ZÀ-Ý][a-zà-ÿ]+)+\b)")


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = str(text)
    text = re.sub(r"<[^>]*(?:https?|mailto):[^>]+>", " ", text)
    text = re.sub(r"https?://\S+", " ", text)
    text = re.sub(r"mailto:\S+", " ", text)
    text = re.sub(r"[\u00ad\u034f\u061c\u180e\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]", " ", text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def raw_clean_lines(text: str) -> list[str]:
    text = clean_text(text)
    raw = [line.strip(" \t\r\n-*•·") for line in text.splitlines()]
    lines = []
    for line in raw:
        line = re.sub(r"\s+", " ", line).strip()
        if not line or len(line) < 2:
            continue
        if line.lower().startswith("http"):
            continue
        if re.search(
            r"(?i)annulla iscrizione|unsubscribe|copyright|tutti i diritti|ricevi questa email|"
            r"vous recevez ce mail|conformément|dicono di noi|metodi di pagamento|facebook|instagram|you received",
            line,
        ):
            continue
        lines.append(line)
    return lines


def filtered_lines(text: str) -> list[str]:
    return [line for line in raw_clean_lines(text) if line not in COMMON_LINES]


def parse_uploaded_file(uploaded_file) -> dict:
    filename = uploaded_file.name
    suffix = Path(filename).suffix.lower()
    data = uploaded_file.getvalue()

    if suffix == ".msg":
        if extract_msg is None:
            raise RuntimeError("Per leggere i file .msg devi installare la libreria extract-msg.")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        msg = extract_msg.Message(tmp_path)
        html_body = msg.htmlBody
        html_text = ""
        if isinstance(html_body, bytes):
            html_text = html_body.decode("utf-8", errors="ignore")
        elif html_body:
            html_text = str(html_body)
        body = msg.body or ""
        if body and len(body.strip()) > 30:
            text = body
        elif html_text:
            text = BeautifulSoup(html_text, "html.parser").get_text("\n", strip=True)
        else:
            text = ""
        return {"sender": msg.sender or "", "date": msg.date, "subject": msg.subject or "", "text": clean_text(text)}

    if suffix == ".eml":
        msg = BytesParser(policy=policy.default).parsebytes(data)
        sender = msg.get("from", "")
        subject = msg.get("subject", "")
        try:
            date = parsedate_to_datetime(msg.get("date")) if msg.get("date") else None
        except Exception:
            date = None
        plain_parts, html_parts = [], []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    continue
                try:
                    content = part.get_content()
                except Exception:
                    continue
                if part.get_content_type() == "text/plain":
                    plain_parts.append(content)
                elif part.get_content_type() == "text/html":
                    html_parts.append(content)
        else:
            content = msg.get_content()
            if msg.get_content_type() == "text/plain":
                plain_parts.append(content)
            elif msg.get_content_type() == "text/html":
                html_parts.append(content)
        text = "\n".join(plain_parts) if plain_parts else "\n".join(
            BeautifulSoup(h, "html.parser").get_text("\n", strip=True) for h in html_parts
        )
        return {"sender": sender, "date": date, "subject": subject, "text": clean_text(text)}

    if suffix in [".html", ".htm"]:
        raw = data.decode("utf-8", errors="ignore")
        text = BeautifulSoup(raw, "html.parser").get_text("\n", strip=True)
        return {"sender": "", "date": None, "subject": "", "text": clean_text(text)}

    if suffix == ".txt":
        raw = data.decode("utf-8", errors="ignore")
        return {"sender": "", "date": None, "subject": "", "text": clean_text(raw)}

    raw = data.decode("utf-8", errors="ignore")
    return {"sender": "", "date": None, "subject": "", "text": clean_text(raw)}


def get_olta(sender: str, filename: str) -> str:
    source = f"{sender} {filename}".lower()
    for olta, aliases in KNOWN_OLTAS.items():
        if any(alias.lower() in source for alias in aliases):
            return olta
    stem = Path(filename).stem
    match = re.match(r"([A-Za-z ._-]+)[_\-\s]\d", stem)
    if match:
        return match.group(1).replace("_", " ").strip().upper()
    return ""


def format_date(dt, filename: str) -> str:
    if dt:
        try:
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
    match = re.search(r"(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{4})", filename)
    if match:
        return f"{int(match.group(1)):02d}/{int(match.group(2)):02d}/{match.group(3)}"
    match = re.search(r"(\d{4})[.\-_](\d{1,2})[.\-_](\d{1,2})", filename)
    if match:
        return f"{int(match.group(3)):02d}/{int(match.group(2)):02d}/{match.group(1)}"
    return ""


def cleanup_field(value) -> str:
    if not value:
        return ""
    value = str(value)
    value = re.sub(r"\s+", " ", value).strip(" -;:,")
    return value.replace(" ,", ",")


def find_companies(context: str) -> str:
    found = []
    for company in KNOWN_COMPANIES:
        pattern = re.compile(r"(?i)(?<![A-Za-z])" + re.escape(company).replace("\\ ", r"\s+") + r"(?![A-Za-z])")
        if pattern.search(context):
            normalized = company.replace("La Meridionale", "La Méridionale")
            found.append(normalized)
    output = []
    for company in found:
        if company == "P&O" and "P&O Ferries" in found:
            continue
        if company not in output:
            output.append(company)
    return ", ".join(output)


def find_destinations(context: str) -> str:
    found = []
    for dest in KNOWN_DESTINATIONS:
        pattern = re.compile(r"(?i)(?<![A-Za-zÀ-ÿ])" + re.escape(dest).replace("\\ ", r"\s+") + r"(?![A-Za-zÀ-ÿ])")
        if pattern.search(context):
            found.append(dest)
    for route in ROUTE_HINT_RE.findall(context):
        if route not in found:
            found.append(route)
    return ", ".join(dict.fromkeys(found))


def extract_promo_phrase(context: str) -> str:
    context = " ".join(line.strip() for line in str(context).splitlines() if line.strip())

    match = re.search(r"(?i)(?:valore\s+del\s+coupon\s*:?\s*)?(\d+(?:[,.]\d+)?\s?€\s*(?:di\s+)?sconto(?:\s*\([^)]*\))?)", context)
    if match:
        return match.group(1).strip()

    if re.search(r"(?i)meilleur prix garanti|miglior prezzo garantito|best price guarantee", context):
        return "Meilleur Prix Garanti / rimborso differenza"

    match = re.search(r"(?i)(biglietto\s+da\s+\d+(?:[,.]\d+)?\s?€\s*\+\s*veicolo\s+gratis(?:\s+al\s+ritorno)?)", context)
    if match:
        return match.group(1).strip()

    match = re.search(r"(?i)(veicolo\s+gratis(?:\s+al\s+ritorno)?|vehicle\s+free|véhicule\s+gratuit)", context)
    if match:
        return match.group(1).strip()

    patterns = [
        r"(?i)(fino\s+al\s+\d{1,2}\s?%\s*(?:di\s+sconto)?)",
        r"(?i)(fino\s+a\s+\d{1,2}\s?%\s*(?:di\s+sconto)?)",
        r"(?i)(risparmia\s+fino\s+al\s+\d{1,2}\s?%)",
        r"(?i)(risparmia\s+il\s+\d{1,2}\s?%)",
        r"(?i)(sconto\s+del\s+\d{1,2}\s?%)",
        r"(?i)(\d{1,2}\s?%\s+di\s+sconto)",
        r"(?i)(-\s?\d{1,2}\s?%)",
        r"(?i)(jusqu['’]à\s+\d{1,2}\s?%)",
        r"(?i)(remise\s+.*?\d{1,2}\s?%)",
    ]
    for pattern in patterns:
        match = re.search(pattern, context)
        if match:
            return re.sub(r"\s+", " ", match.group(1)).strip()

    match = re.search(r"(?i)((?:a\s+partire\s+da|partire\s+da|da|dès|à\s+partir\s+de)\s*\d+(?:[,.]\d+)?\s?€)", context)
    if match:
        return match.group(1).strip()

    return ""


def infer_tipologia(subject: str, text: str, promo: str) -> str:
    sample = f"{subject}\n{text[:4000]}".lower()
    if re.search(r"\b(coupon|codice|voucher|bon d’achat|code promo|coupon sconto)\b", sample):
        return "Codice sconto"
    if re.search(r"\b(ultimi giorni|ultimo giorno|last chance|solo oggi|demain)\b", sample):
        return "Reminder"
    if promo or re.search(r"\b(sconto|offerta|promo|risparmia|remise|offre|a partire da|fino al|jusqu|discount)\b|-\s?\d{1,2}\s?%|\d+\s?€", sample):
        return "Promozione"
    return "Comunicazione generica"


def make_row(parsed: dict, filename: str, tipologia: str, company: str, destination: str, promo: str, confidence: float, note: str) -> dict:
    text = parsed.get("text", "") or ""
    return {
        "OLTA MITTENTE": get_olta(parsed.get("sender", ""), filename),
        "DATA NEWSLETTER": format_date(parsed.get("date"), filename),
        "TIPOLOGIA": cleanup_field(tipologia),
        "COMPAGNIA PROMOSSA": cleanup_field(company),
        "DESTINAZIONE/MERCATO/TRATTA": cleanup_field(destination),
        "PROMOZIONE": cleanup_field(promo),
        "OGGETTO EMAIL": parsed.get("subject", "") or "",
        "FILE ORIGINE": filename,
        "TESTO ESTRATTO": (text[:1000] + "…") if len(text) > 1000 else text,
        "CONFIDENCE": round(float(confidence), 2),
        "NOTE": note,
    }


def extract_fallback(parsed: dict, filename: str) -> dict:
    full = f"{parsed.get('subject', '')}\n{parsed.get('text', '')[:4000]}"
    promo = extract_promo_phrase(full) or "Comunicazione / promozione senza meccanica economica esplicita"
    company = find_companies(full)
    destination = find_destinations(full)
    tipologia = infer_tipologia(parsed.get("subject", ""), parsed.get("text", ""), promo)
    return make_row(parsed, filename, tipologia, company, destination, promo, 0.45, "Fallback: estrazione generica")


def extract_aferry(parsed: dict, filename: str) -> list[dict]:
    lines = raw_clean_lines(parsed.get("text", ""))
    positions = [i for i, line in enumerate(lines) if re.search(r"(?i)trova il miglior prezzo", line)]
    start = positions[-1] + 1 if positions else 0
    end = len(lines)
    for i, line in enumerate(lines[start:], start=start):
        if re.search(r"(?i)prenota più veloce|termini e condizioni", line):
            end = i
            break
    section = lines[start:end]
    blocks, current = [], []
    for line in section:
        if re.search(r"(?i)^prenota ora\b", line):
            if current:
                blocks.append(current)
                current = []
        else:
            current.append(line)
    if current:
        blocks.append(current)

    rows, seen = [], set()
    for block in blocks:
        block = [
            line for line in block
            if not re.search(r"(?i)^destinazioni|^traversate|^verso |^le tue prossime|^in famiglia|^viaggia a modo|^verso sud|^è arrivato|^più luce|^scopri le offerte|^baleari, tunisia|^corsica, marocco|^vacanze in famiglia", line)
        ]
        if not block:
            continue
        block_text = "\n".join(block)
        promo = extract_promo_phrase(block_text)
        if not promo:
            continue
        headline = ""
        for line in block:
            if re.search(r"(?i)^prenota ora|^trova|^scopri|^tutte le offerte|^approfitta|^risparmia|^fino al|^sconto|^corsica linea|^tallink|^gnv:|^brittany ferries|^dfds", line):
                continue
            if (":" in line and (PROMO_RE.search(line) or find_destinations(line))) or find_destinations(line):
                headline = line
                break
        if not headline:
            headline = block[0]
        match = re.match(r"^(.+?)\s*:\s*(.+)$", headline)
        destination = match.group(1).strip() if match else headline
        if re.search(r"(?i)tutte le offerte|risparmia|sconto|fino|approfitta", destination):
            destination = find_destinations(block_text)
        company = find_companies(block_text)
        tipologia = infer_tipologia(parsed.get("subject", ""), block_text, promo)
        key = (company, destination, promo)
        if key in seen:
            continue
        seen.add(key)
        rows.append(make_row(parsed, filename, tipologia, company, destination, promo, 0.90, "Regola specifica AFerry: blocco offerta"))
    return rows


def extract_alloferry(parsed: dict, filename: str) -> list[dict]:
    lines = filtered_lines(parsed.get("text", ""))
    full = "\n".join(lines[:120])
    rows = []

    if re.search(r"(?i)meilleur prix garanti", full):
        company = find_companies(full)
        match = re.search(
            r"(?i)Offre valable pour les traversées\s*:\s*(.+?)(?:\n\s*- Offre applicable|\n\s*- La garantie|\n\s*- L'offre)",
            full,
            re.S,
        )
        destination = re.sub(r"\s+", " ", match.group(1)).strip(" -") if match else find_destinations(full)
        rows.append(make_row(parsed, filename, "Promozione", company, destination, "Meilleur Prix Garanti / rimborso differenza", 0.80, "Regola specifica Alloferry: garanzia prezzo"))
        return rows

    company = ""
    match_company = re.search(r"(?i)compagnie\s*:\s*([^\n]+)", full)
    if match_company:
        company = match_company.group(1).strip()
    else:
        company = find_companies(full)

    seen = set()
    for line in lines[:40]:
        if not re.search(r"\d{1,2}\s?%", line):
            continue
        if re.search(r"(?i)prix hors taxe|non rétroactive|non cumulable|valable|places limitées|remise déduite|réduction", line):
            continue
        promo = extract_promo_phrase(line)
        if not promo:
            match_percent = re.search(r"(\d{1,2}\s?%)", line)
            promo = match_percent.group(1) if match_percent else ""
        destination = ""
        match = re.match(r"(?i)^(.+?)\s*[-: ]+\s*\d{1,2}\s?%\s*: ?", line)
        if match and not re.match(r"^\s*\d", line):
            destination = match.group(1).strip(" -:")
        if not destination:
            match = re.search(r"(?i)-?\s*\d{1,2}\s?%\s+(?:sur|sui|su|on)?\s*(.+)$", line)
            if match:
                destination = match.group(1).strip(" -:")
        key = (destination, promo, company)
        if promo and destination and key not in seen:
            rows.append(make_row(parsed, filename, "Promozione", company, destination, promo, 0.85, "Regola specifica Alloferry: headline con percentuale"))
            seen.add(key)
    return rows


def extract_traghettiper(parsed: dict, filename: str) -> list[dict]:
    lines = filtered_lines(parsed.get("text", ""))
    full = "\n".join(lines[:100])
    promo = extract_promo_phrase(full)
    match = re.search(r"(?i)valore\s+del\s+coupon\s*:?\s*([^\n]+)", full)
    if match:
        promo = match.group(1).strip()
    code = ""
    match_code = re.search(r"(?i)(?:il tuo codice traghettiper|codice traghettiper)\s*:?\s*\n?\s*([A-Z0-9-]{4,})", full)
    if match_code:
        code = match_code.group(1).strip()
    destination = find_destinations(f"{parsed.get('subject', '')}\n{full}")
    if not destination and re.search(r"(?i)tutte le prenotazioni", full):
        destination = "Tutte le prenotazioni"
    note = "Regola specifica TraghettiPer: coupon"
    if code:
        note += f" | Codice rilevato: {code}"
    tipologia = "Codice sconto" if COUPON_CODE_RE.search(full) else infer_tipologia(parsed.get("subject", ""), full, promo)
    return [make_row(parsed, filename, tipologia, "", destination, promo or "Coupon/codice sconto", 0.90 if promo else 0.65, note)]


def extract_netferry(parsed: dict, filename: str) -> list[dict]:
    full = f"{parsed.get('subject', '')}\n{parsed.get('text', '')[:3000]}"
    promo = extract_promo_phrase(full)
    company = find_companies(full)
    destination = find_destinations(full)
    if not promo:
        promo = "Comunicazione / promozione senza meccanica economica esplicita"
        tipologia = "Comunicazione generica"
        confidence = 0.50
    else:
        tipologia = infer_tipologia(parsed.get("subject", ""), parsed.get("text", ""), promo)
        confidence = 0.75
    return [make_row(parsed, filename, tipologia, company, destination, promo, confidence, "Regola specifica Netferry: sintesi newsletter")]


def extract_ferryhopper(parsed: dict, filename: str) -> list[dict]:
    lines = filtered_lines(parsed.get("text", ""))
    rows, seen = [], set()
    for i, line in enumerate(lines[:90]):
        if not PROMO_RE.search(line):
            continue
        if re.search(r"(?i)^promo traghetti|fino al -20% sui traghetti", line):
            continue
        if re.search(r"(?i)scade|assistenza|app di ferryhopper", line):
            continue
        prevs = [lines[j] for j in range(max(0, i - 3), i)]
        nexts = [lines[j] for j in range(i + 1, min(len(lines), i + 4))]
        context = "\n".join(prevs + [line] + nexts)
        promo_context = line + "\n" + ("\n".join(nexts[:1]) if re.search(r"(?i)veicolo|gratis", line) else "")
        promo = extract_promo_phrase(promo_context)
        if not promo:
            continue
        destination = ""
        for previous in reversed(prevs):
            if len(previous) > 3 and not PROMO_RE.search(previous) and not re.search(
                r"(?i)super promo|vamos|scopri|prenota|sugli itinerari|valido|traghetti scontati|mix perfetto|calette|segreto|itinerario",
                previous,
            ):
                destination = previous
                break
        if not destination:
            destination = find_destinations(context)
        company = find_companies(context)
        local_company = find_companies("\n".join([line] + nexts))
        if local_company:
            company = local_company
        key = (destination, promo, company)
        if key in seen:
            continue
        seen.add(key)
        rows.append(make_row(parsed, filename, "Promozione", company, destination, promo, 0.85, "Regola specifica Ferryhopper: riga promo + contesto"))
    return rows


def extract_rows(parsed: dict, filename: str) -> list[dict]:
    olta = get_olta(parsed.get("sender", ""), filename)
    try:
        if olta == "AFERRY":
            rows = extract_aferry(parsed, filename)
        elif olta == "ALLOFERRY":
            rows = extract_alloferry(parsed, filename)
        elif olta == "TRAGHETTIPER":
            rows = extract_traghettiper(parsed, filename)
        elif olta == "NETFERRY":
            rows = extract_netferry(parsed, filename)
        elif olta == "FERRYHOPPER":
            rows = extract_ferryhopper(parsed, filename)
        else:
            rows = []
        if not rows:
            rows = [extract_fallback(parsed, filename)]
    except Exception as exc:
        row = extract_fallback(parsed, filename)
        row["NOTE"] += f" | Errore regola specifica: {exc}"
        row["CONFIDENCE"] = 0.30
        rows = [row]

    cleaned = []
    seen = set()
    for row in rows:
        for column in FINAL_COLUMNS:
            row.setdefault(column, "")
        key = tuple(row[col] for col in ["OLTA MITTENTE", "DATA NEWSLETTER", "COMPAGNIA PROMOSSA", "DESTINAZIONE/MERCATO/TRATTA", "PROMOZIONE"])
        if key not in seen:
            cleaned.append(row)
            seen.add(key)
    return cleaned


def create_summary_text(df: pd.DataFrame) -> str:
    if df.empty:
        return "Nessuna newsletter analizzata."
    n_files = df["FILE ORIGINE"].nunique()
    n_rows = len(df)
    olta_counts = df["OLTA MITTENTE"].value_counts().to_dict()
    type_counts = df["TIPOLOGIA"].value_counts().to_dict()

    olta_text = ", ".join(f"{olta}: {count}" for olta, count in olta_counts.items())
    type_text = ", ".join(f"{typ}: {count}" for typ, count in type_counts.items())

    promo_rows = df[df["TIPOLOGIA"].isin(["Promozione", "Codice sconto", "Reminder"])]
    destinations = []
    for value in promo_rows["DESTINAZIONE/MERCATO/TRATTA"].dropna().astype(str):
        for item in re.split(r",|;", value):
            item = item.strip()
            if item and len(item) <= 60:
                destinations.append(item)
    top_dest = pd.Series(destinations).value_counts().head(5).to_dict() if destinations else {}
    dest_text = ", ".join(f"{dest} ({count})" for dest, count in top_dest.items()) if top_dest else "non disponibile"

    return (
        f"Nel periodo analizzato sono state caricate {n_files} newsletter e sono state estratte {n_rows} righe informative. "
        f"La distribuzione per OLTA è la seguente: {olta_text}. "
        f"Per tipologia, le comunicazioni risultano classificate come: {type_text}. "
        f"I mercati/tratte più ricorrenti risultano: {dest_text}. "
        "Il testo è generato automaticamente con regole Python e va considerato come bozza da revisionare."
    )


def build_excel(df: pd.DataFrame, summary_text: str) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Database", index=False)
        summary_rows = []
        if not df.empty:
            summary_rows.append({"Indicatore": "Newsletter caricate", "Valore": df["FILE ORIGINE"].nunique()})
            summary_rows.append({"Indicatore": "Righe estratte", "Valore": len(df)})
            for olta, count in df["OLTA MITTENTE"].value_counts().items():
                summary_rows.append({"Indicatore": f"Comunicazioni - {olta}", "Valore": count})
            for typ, count in df["TIPOLOGIA"].value_counts().items():
                summary_rows.append({"Indicatore": f"Tipologia - {typ}", "Valore": count})
        summary_df = pd.DataFrame(summary_rows or [{"Indicatore": "Note", "Valore": "Nessun dato"}])
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame([{"Descrizione automatica": summary_text}]).to_excel(writer, sheet_name="Text summary", index=False)
    output.seek(0)
    return output


st.set_page_config(page_title="OLTA Newsletter Extractor", layout="wide")
st.title("OLTA Newsletter Extractor")
st.caption("Versione 0.1 — estrazione rule-based, senza modelli AI")

st.markdown(
    "Carica le newsletter in formato `.msg`, `.eml`, `.html` o `.txt`. "
    "L'app estrarrà una tabella modificabile e genererà un file Excel con Database, Summary e descrizione testuale."
)

uploaded_files = st.file_uploader(
    "Carica newsletter",
    type=["msg", "eml", "html", "htm", "txt"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_rows = []
    errors = []
    for uploaded_file in uploaded_files:
        try:
            parsed = parse_uploaded_file(uploaded_file)
            all_rows.extend(extract_rows(parsed, uploaded_file.name))
        except Exception as exc:
            errors.append(f"{uploaded_file.name}: {exc}")

    if errors:
        st.warning("Alcuni file non sono stati letti correttamente:\n" + "\n".join(errors))

    df = pd.DataFrame(all_rows, columns=FINAL_COLUMNS)
    if df.empty:
        st.info("Nessuna informazione estratta dai file caricati.")
    else:
        st.subheader("Tabella estratta - modificabile")
        st.write("Controlla soprattutto le righe con confidence bassa o media.")
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "TIPOLOGIA": st.column_config.SelectboxColumn(
                    "TIPOLOGIA",
                    options=["Promozione", "Codice sconto", "Reminder", "Comunicazione generica"],
                ),
                "CONFIDENCE": st.column_config.NumberColumn("CONFIDENCE", min_value=0.0, max_value=1.0, step=0.05),
            },
        )

        summary_text = create_summary_text(edited_df)
        st.subheader("Descrizione testuale automatica")
        summary_text = st.text_area("Bozza riepilogativa", value=summary_text, height=160)

        excel_data = build_excel(edited_df, summary_text)
        st.download_button(
            "Scarica Excel",
            data=excel_data,
            file_name="OLTA_newsletter_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Scarica descrizione TXT",
            data=summary_text.encode("utf-8"),
            file_name="OLTA_newsletter_summary.txt",
            mime="text/plain",
        )
else:
    st.info("Carica uno o più file per iniziare.")
