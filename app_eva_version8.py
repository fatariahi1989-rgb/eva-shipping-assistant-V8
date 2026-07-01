import re
import os
import io
import json
import math
from datetime import datetime

import pandas as pd
import streamlit as st

# ============================================================
# EVA Version 8
# Erweiterungen gegenüber EVA 7:
# 1) Freitext-Verarbeitung mit OpenAI als optionale Eingabehilfe
# 2) Anzeige von Realgewicht, Volumengewicht und Abrechnungsgewicht
# 3) Gefahrgut-Sonderbehandlung über neue Excel-Spalte "Gefahrgut erlaubt"
# 4) Session-Protokoll mit CSV-Download
# 5) Optionale Online-Datenquelle für semi-automatische Datenaktualisierung
#
# Wichtig: Die bestehende Bewertungslogik bleibt regelbasiert.
# OpenAI interpretiert nur Freitext und füllt Felder vor.
# Die Entscheidung entsteht weiterhin aus Excel-Daten + Regeln + Scoring.
# ============================================================

st.set_page_config(page_title="EVA – Entscheidungsassistent", page_icon="📦", layout="wide")
st.title("📦 EVA – Entscheidungsassistent für die Versanddienstleisterauswahl")
st.caption("Hybrides System: OpenAI interpretiert Freitext, EVA entscheidet regelbasiert mit Excel-Daten, Ausschlussregeln und Scoring.")

DEFAULT_WEIGHTS = {
    "price": 0.21, "insurance_efficiency": 0.14,
    "runtime": 0.20, "otd": 0.08,
    "tracking": 0.07, "damage": 0.08, "receiver_flex": 0.05,
    "liability": 0.08, "goods_fit": 0.04,
    "international": 0.05,
}

GOODS_OPTIONS = ["Bücher", "Elektronik", "Laptop", "Smartphone", "Kleidung", "Schmuck", "Uhr", "Gefahrgut", "Sonstiges"]


# ============================================================
# Bestehende Hilfsfunktionen aus EVA 7
# ============================================================
def norm(text):
    return re.sub(r"[^a-z0-9]+", "", str(text).lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss"))


def find_col(df, *keywords):
    normalized_cols = {norm(c): c for c in df.columns}
    keys = [norm(k) for k in keywords]
    for ncol, original in normalized_cols.items():
        if all(k in ncol for k in keys):
            return original
    return None


def parse_number(value, default=0.0):
    if pd.isna(value):
        return default
    s = str(value).replace("€", "").replace("kg", "").replace("cm", "").replace("%", "")
    s = s.replace(".", "").replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else default


def parse_price(value):
    return parse_number(value, default=math.inf)


def parse_runtime_days(value):
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+(?:[,.]\d+)?", str(value))]
    if not nums:
        return 7.0
    return sum(nums) / len(nums)


def parse_weight_limit(value):
    return parse_number(value, default=0.0)


def parse_liability(value):
    s = str(value).lower()
    if "nein" in s or s.strip() in {"", "nan"}:
        return 0.0
    return parse_number(value, default=0.0)


def parse_dimensions(value):
    s = str(value).lower().replace("×", "x").replace("*", "x")
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+(?:[,.]\d+)?", s)]
    if not nums:
        return []
    return nums[:3] if "x" in s and len(nums) >= 3 else [max(nums)]


def package_fits(package_dims, limit_text):
    limits = parse_dimensions(limit_text)
    if not limits:
        return False
    dims = sorted(package_dims, reverse=True)
    if len(limits) == 1:
        return max(dims) <= limits[0]
    limits = sorted(limits, reverse=True)
    return all(d <= l for d, l in zip(dims, limits))


def yes(value):
    return str(value).strip().lower() in ["ja", "yes", "true", "1", "y", "erlaubt", "allowed"]


def allowed_value(value):
    s = str(value).strip().lower()
    return not (s in ["nein", "no", "false", "0", "nicht erlaubt", "verboten"] or "ablehnen" in s)


def score_runtime(days):
    return max(0, min(100, (7 - days) / 6 * 100))


def score_otd(value):
    otd = parse_number(value, default=0)
    if otd >= 97:
        return 100
    if otd >= 90:
        return 75
    if otd >= 80:
        return 50
    return 0


def score_damage(value):
    damage = parse_number(value, default=1.5)
    if damage < 0.1:
        return 100
    if damage <= 0.5:
        return 75
    if damage <= 1.0:
        return 50
    return 0


def score_receiver(row, notify_col, flex_col):
    text = f"{row.get(notify_col, '')} {row.get(flex_col, '')}".lower()
    has_notify = any(x in text for x in ["ja", "benachrichtigung", "notification", "voraus"])
    has_flex = any(x in text for x in ["paketshop", "packstation", "umleitung", "abstell", "zeitfenster", "flex"])
    if has_notify and has_flex:
        return 100
    if has_notify or has_flex:
        return 50
    return 0


def insurance_cost(carrier, value, liability):
    c = str(carrier).upper()
    if value <= liability:
        return 0.0
    if c == "DHL":
        if value <= 2500:
            return 6.99
        if value <= 25000:
            return 19.99
        return math.inf
    if c == "DPD":
        if value <= 10000:
            return max(5.0, value * 0.01)
        return math.inf
    if c == "GLS":
        if value <= 5000:
            return max(5.0, value * 0.01)
        return math.inf
    return math.inf


# ============================================================
# NEU 5: Optionale Online-Datenquelle
# Für die Abgabe defensiv formulieren als semi-automatisch:
# Preise können in einer gehosteten Excel-Datei aktualisiert werden,
# ohne dass der Nutzer jedes Mal eine lokale Datei neu hochladen muss.
# ============================================================
def get_excel_source(uploaded_file, online_url):
    if uploaded_file is not None:
        return uploaded_file
    if online_url and online_url.strip():
        return online_url.strip()
    return None


# ============================================================
# Bestehende Gewichtungslogik aus EVA 7
# ============================================================
def load_weights(xls):
    try:
        df = pd.read_excel(xls, sheet_name="Gewichtung")
        sub = find_col(df, "subkriterium")
        val = find_col(df, "normalisiert")
        if not sub or not val:
            return DEFAULT_WEIGHTS
        mapping = DEFAULT_WEIGHTS.copy()
        for _, r in df.iterrows():
            name = norm(r.get(sub, ""))
            w = r.get(val)
            if pd.isna(w):
                continue
            if "grundpreis" in name:
                mapping["price"] = float(w)
            elif "versicherung" in name:
                mapping["insurance_efficiency"] = float(w)
            elif "laufzeit" in name:
                mapping["runtime"] = float(w)
            elif "otd" in name or "lieferzuverlaessigkeit" in name:
                mapping["otd"] = float(w)
            elif "tracking" in name:
                mapping["tracking"] = float(w)
            elif "handling" in name and "warenart" not in name:
                mapping["damage"] = float(w)
            elif "empfaenger" in name:
                mapping["receiver_flex"] = float(w)
            elif "haftung" in name:
                mapping["liability"] = float(w)
            elif "warenart" in name:
                mapping["goods_fit"] = float(w)
            elif "auslandsversand" in name or "international" in name:
                mapping["international"] = float(w)
        total = sum(mapping.values())
        return {k: v / total for k, v in mapping.items()} if total else DEFAULT_WEIGHTS
    except Exception:
        return DEFAULT_WEIGHTS


def goods_allowed(goods_rules, goods_type, carrier):
    if goods_rules is None or goods_rules.empty:
        return True, "Keine Warenart-Regel gefunden"
    goods_col = find_col(goods_rules, "warenart")
    allow_col = find_col(goods_rules, carrier, "erlaubt")
    if not goods_col or not allow_col:
        return True, "Spalte für Warenart/Carrier fehlt"
    match = goods_rules[goods_rules[goods_col].astype(str).str.lower().str.contains(str(goods_type).lower(), na=False)]
    if match.empty:
        return True, "Keine spezifische Warenart-Regel"
    val = match.iloc[0].get(allow_col, "Ja")
    return allowed_value(val), f"Warenart-Regel: {carrier} erlaubt = {val}"


# ============================================================
# NEU 2: Die Berechnung existierte bereits.
# Sie wird jetzt zusätzlich transparent im UI angezeigt.
# ============================================================
def calculate_billable_weight(length, width, height, real_weight):
    volumetric = (length * width * height) / 5000
    return max(real_weight, volumetric), volumetric


# ============================================================
# NEU 3: Gefahrgut-Sonderbehandlung
# Erwartete neue Spalte im Sheet "Grundpreis": "Gefahrgut erlaubt"
# Ja/Nein pro Tarifzeile bzw. Carrier-Service.
# ============================================================
def dangerous_goods_allowed(row, danger_col):
    if not danger_col:
        return False, "Spalte 'Gefahrgut erlaubt' fehlt"
    return allowed_value(row.get(danger_col, "Nein")), row.get(danger_col, "Nein")


# ============================================================
# Bestehende Kernfunktion calculate_results bleibt im Grundaufbau erhalten.
# Ergänzt wurden nur:
# - danger_col in cols
# - Gefahrgut-Ausschluss vor normaler Warenart-Prüfung
# - Meta-Felder real/volumetric/billable/dangerous_goods_manual_check
# ============================================================
def calculate_results(xls, length, width, height, real_weight, goods_value, goods_type, dest_country):
    grund = pd.read_excel(xls, sheet_name="Grundpreis")
    try:
        goods_rules = pd.read_excel(xls, sheet_name="Sonderregeln nach Warenart")
    except Exception:
        goods_rules = pd.DataFrame()
    weights = load_weights(xls)

    cols = {
        "carrier": find_col(grund, "carrier"), "service": find_col(grund, "versandart"),
        "weight": find_col(grund, "gewicht", "max"), "dims": find_col(grund, "abmessungen"),
        "price": find_col(grund, "grundpreis"), "tracking": find_col(grund, "tracking"),
        "liability": find_col(grund, "haftung"), "runtime": find_col(grund, "laufzeit"),
        "international": find_col(grund, "auslandsversand"), "otd": find_col(grund, "otd"),
        "damage": find_col(grund, "schadensquote"),
        "notify": find_col(grund, "empfaengerbenachrichtigung") or find_col(grund, "empfängerbenachrichtigung"),
        "flex": find_col(grund, "zustellflexibilitaet") or find_col(grund, "zustellflexibilität"),
        "handling": find_col(grund, "handling"),
        # NEU 3: optionale Pflichtlogik nur für goods_type == "Gefahrgut"
        "dangerous": find_col(grund, "gefahrgut", "erlaubt"),
    }
    required = ["carrier", "service", "weight", "dims", "price", "tracking", "liability", "runtime", "international"]
    missing = [k for k in required if not cols[k]]
    if missing:
        st.error(f"Fehlende Pflichtspalten im Sheet Grundpreis: {missing}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {}

    package_dims = [length, width, height]
    billable_weight, volumetric_weight = calculate_billable_weight(length, width, height, real_weight)
    international = str(dest_country).strip().lower() not in ["de", "deutschland", "germany"]
    is_dangerous_goods = str(goods_type).strip().lower() == "gefahrgut"

    rows, rejected = [], []
    all_valid_tariffs = []
    best_candidates = []

    for _, r in grund.iterrows():
        carrier = str(r[cols["carrier"]]).strip().upper()
        if not carrier or carrier == "NAN":
            continue
        service = r[cols["service"]]

        if billable_weight > parse_weight_limit(r[cols["weight"]]):
            rejected.append([carrier, service, "Gewicht überschreitet Tariflimit"])
            continue
        if not package_fits(package_dims, r[cols["dims"]]):
            rejected.append([carrier, service, "Maße passen nicht zum Tarif"])
            continue
        if goods_value > 200 and not yes(r[cols["tracking"]]):
            rejected.append([carrier, service, "Tracking fehlt bei Warenwert > 200 €"])
            continue
        if international and not yes(r[cols["international"]]):
            rejected.append([carrier, service, "Kein Auslandsversand für internationale Sendung"])
            continue

        # NEU 3: Gefahrgut wird nicht wie normale Warenart behandelt.
        # Carrier/Tarife ohne Freigabe werden abgelehnt.
        if is_dangerous_goods:
            dg_ok, dg_value = dangerous_goods_allowed(r, cols["dangerous"])
            if not dg_ok:
                rejected.append([carrier, service, "Gefahrgut nicht zugelassen"])
                continue
            goods_reason = f"Gefahrgut erlaubt = {dg_value}; manuelle ADR-Prüfung erforderlich"
        else:
            ok_goods, goods_reason = goods_allowed(goods_rules, goods_type, carrier)
            if not ok_goods:
                rejected.append([carrier, service, goods_reason])
                continue

        price = parse_price(r[cols["price"]])
        liability = parse_liability(r[cols["liability"]])
        ins = insurance_cost(carrier, goods_value, liability)
        total_cost = price + ins if math.isfinite(ins) else math.inf
        runtime = str(r[cols["runtime"]])
        insurance_text = "Keine Zusatzversicherung nötig" if ins == 0 else (f"Zusatzversicherung: {ins:.2f} €" if math.isfinite(ins) else "Manuelle Versicherungsprüfung nötig")
        all_valid_tariffs.append({
            "Carrier": carrier,
            "Versandart": service,
            "Gesamtpreis mit Versicherung": f"{total_cost:.2f} €" if math.isfinite(total_cost) else "Manuell",
            "Versicherung": insurance_text,
            "Laufzeit": runtime,
        })
        best_candidates.append((carrier, total_cost, r, price, liability, ins, goods_reason))

    best_by_carrier = {}
    for item in best_candidates:
        carrier = item[0]
        if carrier not in best_by_carrier or item[1] < best_by_carrier[carrier][1]:
            best_by_carrier[carrier] = item
    candidates = list(best_by_carrier.values())

    meta = {
        "real_weight": real_weight,
        "volumetric_weight": volumetric_weight,
        "billable_weight": billable_weight,
        "dangerous_goods": is_dangerous_goods,
        "dangerous_goods_manual_check": is_dangerous_goods and bool(candidates),
    }

    if not candidates:
        return pd.DataFrame(), pd.DataFrame(rejected, columns=["Carrier", "Versandart", "Ablehnungsgrund"]), pd.DataFrame(), meta

    prices = [c[1] for c in candidates if math.isfinite(c[1])]
    min_price, max_price = min(prices), max(prices)

    for carrier, total_cost, r, base_price, liability, ins, goods_reason in candidates:
        price_score = 100 if max_price == min_price else (max_price - total_cost) / (max_price - min_price) * 100
        insurance_score = 100 if goods_value <= liability else max(0, min(100, liability / max(goods_value, 1) * 100))
        runtime_score = score_runtime(parse_runtime_days(r[cols["runtime"]]))
        otd_score = score_otd(r.get(cols["otd"], 0)) if cols["otd"] else 50
        tracking_score = 100 if yes(r[cols["tracking"]]) else 0
        damage_score = score_damage(r.get(cols["damage"], 1.5)) if cols["damage"] else 50
        receiver_score = score_receiver(r, cols["notify"], cols["flex"]) if cols["notify"] or cols["flex"] else 50
        liability_score = insurance_score
        handling_text = str(r.get(cols["handling"], "")).lower() if cols["handling"] else ""
        goods_fit_score = 100 if "spezial" in handling_text else 50
        intl_score = 100 if yes(r[cols["international"]]) else 50
        service_score = (tracking_score + damage_score + receiver_score) / 3
        safety_score = (liability_score + goods_fit_score) / 2
        score = (
            weights["price"] * price_score + weights["insurance_efficiency"] * insurance_score +
            weights["runtime"] * runtime_score + weights["otd"] * otd_score +
            weights["tracking"] * tracking_score + weights["damage"] * damage_score + weights["receiver_flex"] * receiver_score +
            weights["liability"] * liability_score + weights["goods_fit"] * goods_fit_score + weights["international"] * intl_score
        )
        rows.append({
            "Carrier": carrier,
            "Versandart": r[cols["service"]],
            "Grundpreis €": round(base_price, 2),
            "Versicherung €": 0 if ins == 0 else (round(ins, 2) if math.isfinite(ins) else "Manuell"),
            "Gesamtkosten €": round(total_cost, 2) if math.isfinite(total_cost) else "Manuell",
            "Score": round(score, 2),
            "Preis-Score": round(price_score, 1),
            "Laufzeit-Score": round(runtime_score, 1),
            "Service-Score": round(service_score, 1),
            "Sicherheits-Score": round(safety_score, 1),
            "Begründung": f"{goods_reason}; Haftung {liability:.0f} €; Laufzeit {r[cols['runtime']]}"
        })

    ranking = pd.DataFrame(rows).sort_values("Score", ascending=False)
    price_ranking = pd.DataFrame(all_valid_tariffs).sort_values("Gesamtpreis mit Versicherung") if all_valid_tariffs else pd.DataFrame()
    return ranking, pd.DataFrame(rejected, columns=["Carrier", "Versandart", "Ablehnungsgrund"]), price_ranking, meta


# ============================================================
# NEU 1: OpenAI-Freitext-Verarbeitung
# Fallback: Wenn kein API-Key vorhanden ist oder ein Fehler passiert,
# bleibt EVA vollständig manuell nutzbar.
# ============================================================
def get_openai_api_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY")


def normalize_extracted_payload(data):
    """Macht OpenAI-Ausgabe robust und Streamlit-freundlich."""
    if not isinstance(data, dict):
        return {}

    cleaned = {}
    numeric_fields = ["length", "width", "height", "weight", "goods_value"]
    for field in numeric_fields:
        value = data.get(field, None)
        if value in [None, "", "null"]:
            cleaned[field] = None
        else:
            try:
                cleaned[field] = float(str(value).replace(",", "."))
            except Exception:
                cleaned[field] = None

    goods_type = str(data.get("goods_type") or "").strip()
    if goods_type:
        # Mapping für typische Freitext-Wörter auf die bestehenden EVA-Optionen
        gt_low = goods_type.lower()
        if "laptop" in gt_low:
            goods_type = "Laptop"
        elif "smartphone" in gt_low or "handy" in gt_low:
            goods_type = "Smartphone"
        elif "elektr" in gt_low:
            goods_type = "Elektronik"
        elif "schmuck" in gt_low:
            goods_type = "Schmuck"
        elif "uhr" in gt_low:
            goods_type = "Uhr"
        elif "kleidung" in gt_low or "textil" in gt_low:
            goods_type = "Kleidung"
        elif "buch" in gt_low:
            goods_type = "Bücher"
        elif "gefahr" in gt_low or "adr" in gt_low:
            goods_type = "Gefahrgut"
        elif goods_type not in GOODS_OPTIONS:
            goods_type = "Sonstiges"
    cleaned["goods_type"] = goods_type if goods_type in GOODS_OPTIONS else None

    cleaned["country"] = str(data.get("country") or data.get("destination_country") or "").strip() or None
    cleaned["destination"] = str(data.get("destination") or "").strip() or None
    return cleaned


def extract_fields_with_openai(free_text):
    api_key = get_openai_api_key()
    if not api_key:
        return {}, "Kein OpenAI API-Key gefunden. Die manuelle Eingabe bleibt aktiv."

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        schema_instruction = {
            "length": "number or null, package length in cm",
            "width": "number or null, package width in cm",
            "height": "number or null, package height in cm",
            "weight": "number or null, real weight in kg",
            "goods_value": "number or null, value in EUR",
            "goods_type": "one of Bücher, Elektronik, Laptop, Smartphone, Kleidung, Schmuck, Uhr, Gefahrgut, Sonstiges, or null",
            "country": "destination country in German/English, default Deutschland if only a German city is mentioned",
            "destination": "city or full destination text if available"
        }

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Du bist ein Extraktionsmodul für eine Versand-App. "
                        "Extrahiere nur strukturierte Sendungsdaten. "
                        "Erfinde keine Maße und kein Gewicht. Wenn etwas fehlt, gib null zurück. "
                        f"Antworte ausschließlich als JSON mit diesem Schema: {json.dumps(schema_instruction, ensure_ascii=False)}"
                    ),
                },
                {"role": "user", "content": free_text},
            ],
            temperature=0,
        )
        raw = response.choices[0].message.content
        data = json.loads(raw)
        return normalize_extracted_payload(data), None
    except Exception as exc:
        return {}, f"OpenAI-Extraktion fehlgeschlagen. Manuelle Eingabe bleibt aktiv. Technischer Hinweis: {exc}"


# ============================================================
# Zusatzleistungen und Chat aus EVA 7
# ============================================================
def recommended_services(goods_type, goods_value, winner):
    services = []
    g = str(goods_type).lower()
    if goods_value > 500:
        services.append({"Zusatzleistung": "Zusatzversicherung prüfen", "Grund": "Der Warenwert liegt über 500 €.", "geschätzter Aufpreis": "abhängig vom Carrier"})
    if any(x in g for x in ["elektronik", "laptop", "smartphone", "uhr", "schmuck"]):
        services.append({"Zusatzleistung": "Sichere Verpackung / Handling-Hinweis", "Grund": "Die Warenart ist empfindlich oder hochwertig.", "geschätzter Aufpreis": "0,00 €"})
    if str(goods_type).lower() == "gefahrgut":
        services.append({"Zusatzleistung": "Manuelle ADR-Prüfung", "Grund": "Gefahrgut darf nicht rein automatisch freigegeben werden.", "geschätzter Aufpreis": "manuell"})
    if winner is not None and float(winner.get("Score", 0)) < 70:
        services.append({"Zusatzleistung": "Manuelle Prüfung", "Grund": "Der beste Score liegt unter 70 Punkten.", "geschätzter Aufpreis": "-"})
    return pd.DataFrame(services)


def rule_based_chat_answer(question, results, rejected):
    q = str(question).lower()
    if results is None or results.empty:
        if any(word in q for word in ["versicherung", "haftung", "warenwert", "zusatzversicherung"]):
            return "Die Versicherung hängt vom Warenwert und von der Standardhaftung des Carriers ab. Wenn der Warenwert höher ist als die Haftungsgrenze, empfiehlt EVA eine Zusatzversicherung oder eine manuelle Prüfung."
        return "Bitte lade zuerst eine Excel-Datei hoch und klicke auf **EVA berechnen lassen**. Danach kann ich die Empfehlung erklären."

    winner = results.iloc[0]
    if any(word in q for word in ["gefahrgut", "adr", "dangerous"]):
        return "Gefahrgut wird in EVA gesondert behandelt. Carrier ohne Freigabe in der Spalte **Gefahrgut erlaubt** werden ausgeschlossen. Auch bei erlaubten Carriern zeigt EVA nur eine Prüfungsempfehlung, weil eine manuelle ADR-Prüfung erforderlich bleibt."
    if any(word in q for word in ["versicherung", "haftung", "warenwert", "zusatzversicherung"]):
        ins = winner.get("Versicherung €", "")
        safety = winner.get("Sicherheits-Score", "")
        return (
            f"Bei der Versicherung prüft EVA, ob der Warenwert durch die Standardhaftung des Carriers gedeckt ist. "
            f"Für die empfohlene Option **{winner['Carrier']} – {winner['Versandart']}** wurde eine Versicherungsposition von **{ins} €** berücksichtigt. "
            f"Der Sicherheits-Score liegt bei **{safety} Punkten**. Wenn der Warenwert über der Haftungsgrenze liegt, wird eine Zusatzversicherung empfohlen."
        )
    if "warum" in q or "empfohlen" in q or "gewonnen" in q:
        return f"Empfohlen wird **{winner['Carrier']} – {winner['Versandart']}**, weil diese Option mit **{winner['Score']} Punkten** den höchsten Gesamtscore erreicht. Bewertet wurden Preis, Lieferzeit, Servicequalität, Sicherheit und internationale Versandfähigkeit."
    if "ausgeschlossen" in q or "eliminiert" in q or "abgelehnt" in q:
        if rejected is not None and not rejected.empty:
            return "Einige Tarife wurden ausgeschlossen, weil sie Muss-Kriterien nicht erfüllt haben, zum Beispiel Gewicht, Maße, Tracking, Auslandsversand, Warenart-Regeln oder Gefahrgut-Freigabe. Die genauen Gründe stehen in der Tabelle **Ausgeschlossene Tarife**."
        return "Es wurden keine Tarife ausgeschlossen."
    if "gewicht" in q or "volumen" in q or "abrechnung" in q:
        return "EVA berechnet das Volumengewicht mit Länge × Breite × Höhe / 5000. Abgerechnet wird das höhere Gewicht aus Realgewicht und Volumengewicht."
    if "spalten" in q or "excel" in q:
        return "Benötigt werden im Sheet Grundpreis mindestens: Carrier, Versandart, Gewicht max., Abmessungen, Grundpreis, Tracking, Haftung, Laufzeit und Auslandsversand. Für Gefahrgut zusätzlich: Gefahrgut erlaubt. Für das erweiterte Scoringsystem können OTD-Quote, Schadensquote, Empfängerbenachrichtigung, Zustellflexibilität und Handling-Programm ergänzt werden."
    return "EVA bewertet die verfügbaren Carrier mit einem regelbasierten Scoringsystem. OpenAI hilft nur bei der Interpretation von Freitext; die Entscheidung entsteht aus Excel-Daten, Ausschlussregeln und gewichteten Bewertungskriterien."


# ============================================================
# NEU 4: Session-Protokoll
# ============================================================
def init_session_state():
    defaults = {
        "eva_results": pd.DataFrame(),
        "eva_rejected": pd.DataFrame(),
        "eva_price_ranking": pd.DataFrame(),
        "eva_meta": {},
        "eva_inputs": {},
        "eva_log": [],
        "length_default": 25.0,
        "width_default": 20.0,
        "height_default": 3.0,
        "weight_default": 1.0,
        "value_default": 20.0,
        "goods_default": "Bücher",
        "country_default": "Deutschland",
        "receiver_default": "10115 Berlin, Deutschland",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def append_log(inputs, results):
    if results is not None and not results.empty:
        winner = results.iloc[0]
        recommended_carrier = f"{winner.get('Carrier', '')} – {winner.get('Versandart', '')}"
        score = winner.get("Score", "")
    else:
        recommended_carrier = "Keine gültige Option"
        score = ""

    row = {
        "Zeitstempel": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        **inputs,
        "Empfohlener Carrier": recommended_carrier,
        "Score": score,
    }
    st.session_state.eva_log.append(row)


def log_to_csv_bytes():
    df = pd.DataFrame(st.session_state.eva_log)
    return df.to_csv(index=False).encode("utf-8-sig")


init_session_state()

# ============================================================
# Datenquelle: Upload oder optionale Online-Datei
# ============================================================
st.subheader("📁 EVA-Datenbank")
uploaded = st.file_uploader("Excel-Datenbank hochladen", type=["xlsx"])
online_url = st.text_input(
    "Optionale Online-Excel-Quelle für semi-automatische Datenaktualisierung",
    value="",
    placeholder="z. B. öffentlicher .xlsx-Link oder Google-Sheets-Exportlink",
    help="Für die Abgabe: Das ist keine Live-Scraping-Lösung, sondern eine semi-automatische Aktualisierung über eine gepflegte Online-Datenquelle."
)
xls_source = get_excel_source(uploaded, online_url)

# ============================================================
# NEU 1: Freitext-Eingabe oberhalb des normalen Formulars
# ============================================================
st.subheader("🤖 Freitext-Verarbeitung mit OpenAI")
free_text = st.text_area(
    "Sendung als Satz beschreiben",
    placeholder="Beispiel: Ich möchte einen Laptop im Wert von 1500€ nach Berlin schicken. Das Paket ist 35x25x8 cm groß und wiegt 2 kg.",
    height=90,
)
extract_clicked = st.button("Freitext analysieren und Formular vorausfüllen")

if extract_clicked:
    if not free_text.strip():
        st.warning("Bitte zuerst einen Freitext eingeben.")
    else:
        extracted, error = extract_fields_with_openai(free_text)
        if error:
            st.warning(error)
        if extracted:
            if extracted.get("length"):
                st.session_state.length_default = extracted["length"]
            if extracted.get("width"):
                st.session_state.width_default = extracted["width"]
            if extracted.get("height"):
                st.session_state.height_default = extracted["height"]
            if extracted.get("weight"):
                st.session_state.weight_default = extracted["weight"]
            if extracted.get("goods_value") is not None:
                st.session_state.value_default = extracted["goods_value"]
            if extracted.get("goods_type"):
                st.session_state.goods_default = extracted["goods_type"]
            if extracted.get("country"):
                st.session_state.country_default = extracted["country"]
            if extracted.get("destination"):
                st.session_state.receiver_default = extracted["destination"]
            st.success("Erkannte Felder wurden in das Formular übernommen. Fehlende Felder kannst du manuell ergänzen.")
            with st.expander("Erkannte strukturierte Daten anzeigen"):
                st.json(extracted)
        else:
            st.info("Es konnten keine sicheren Felder extrahiert werden. Bitte manuell ausfüllen.")

# ============================================================
# Bestehendes Eingabeformular, jetzt mit Default-Werten aus Freitext
# ============================================================
st.subheader("📝 Sendungsdaten eingeben")
with st.form("eva_input_form"):
    c1, c2, c3 = st.columns(3)
    with c1:
        length = st.number_input("Länge in cm *", min_value=1.0, value=float(st.session_state.length_default))
        weight = st.number_input("Realgewicht / Rohgewicht in kg *", min_value=0.1, value=float(st.session_state.weight_default))
        sender = st.text_input("PLZ/Stadt Absender *", value="95028 Hof, Deutschland")
    with c2:
        width = st.number_input("Breite in cm *", min_value=1.0, value=float(st.session_state.width_default))
        value = st.number_input("Warenwert in € *", min_value=0.0, value=float(st.session_state.value_default))
        receiver = st.text_input("PLZ/Stadt Empfänger *", value=str(st.session_state.receiver_default))
    with c3:
        height = st.number_input("Höhe in cm *", min_value=1.0, value=float(st.session_state.height_default))
        goods_index = GOODS_OPTIONS.index(st.session_state.goods_default) if st.session_state.goods_default in GOODS_OPTIONS else 0
        goods = st.selectbox("Warenart *", GOODS_OPTIONS, index=goods_index)
        country = st.text_input("Zielland", value=str(st.session_state.country_default))

    calculate_clicked = st.form_submit_button("EVA berechnen lassen")

if calculate_clicked:
    if xls_source is None:
        st.warning("Bitte lade zuerst deine Excel-Datei hoch oder gib eine Online-Excel-Quelle an.")
    else:
        try:
            results, rejected, price_ranking, meta = calculate_results(xls_source, length, width, height, weight, value, goods, country)
            inputs = {
                "Länge": length, "Breite": width, "Höhe": height,
                "Realgewicht": weight, "Warenwert": value,
                "Warenart": goods, "Absender": sender, "Empfänger": receiver,
                "Zielland": country,
            }
            st.session_state.eva_results = results
            st.session_state.eva_rejected = rejected
            st.session_state.eva_price_ranking = price_ranking
            st.session_state.eva_meta = meta
            st.session_state.eva_inputs = inputs
            append_log(inputs, results)
        except Exception as exc:
            st.error(f"Berechnung fehlgeschlagen: {exc}")

results = st.session_state.eva_results
rejected = st.session_state.eva_rejected
price_ranking = st.session_state.eva_price_ranking
meta = st.session_state.eva_meta

# ============================================================
# Ergebnisdarstellung mit NEU 2 und NEU 3
# ============================================================
if xls_source is None:
    st.info("Bitte lade die EVA-Excel-Datenbank hoch oder nutze eine Online-Excel-Quelle und klicke danach auf **EVA berechnen lassen**.")
elif calculate_clicked or (results is not None and not results.empty):
    if results is None or results.empty:
        st.error("Kein Carrier erfüllt alle Muss-Kriterien.")
        if rejected is not None and not rejected.empty:
            with st.expander("Ausgeschlossene Tarife anzeigen"):
                st.dataframe(rejected, use_container_width=True)
    else:
        winner = results.iloc[0]
        cheapest = price_ranking.iloc[0] if price_ranking is not None and not price_ranking.empty else winner

        st.header("📊 Ergebnis der Kostenberechnung")

        # NEU 2: Gewichtstransparenz
        g1, g2, g3 = st.columns(3)
        with g1:
            st.metric("Realgewicht", f"{meta.get('real_weight', weight):.2f} kg")
        with g2:
            st.metric("Volumengewicht", f"{meta.get('volumetric_weight', 0):.2f} kg")
        with g3:
            st.metric("Abrechnungsgewicht", f"{meta.get('billable_weight', weight):.2f} kg")
        st.caption("Formel: Volumengewicht = Länge × Breite × Höhe / 5000. Abgerechnet wird das höhere Gewicht aus Realgewicht und Volumengewicht.")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Günstigste passende Option", f"{cheapest['Carrier']} – {cheapest['Versandart']}", cheapest.get("Gesamtpreis mit Versicherung", ""))
        with col2:
            st.metric("Bester Score", f"{winner['Carrier']} – {winner['Versandart']}", f"{winner['Score']} Punkte")

        if meta.get("dangerous_goods_manual_check"):
            # NEU 3: Keine normale Auto-Empfehlung bei Gefahrgut
            st.warning(
                "Gefahrgutversand erfordert manuelle ADR-Prüfung. "
                f"EVA zeigt als beste prüfbare Option: {winner['Carrier']} – {winner['Versandart']} | Score: {winner['Score']} | Kosten: {winner['Gesamtkosten €']} €."
            )
        else:
            st.success(f"Empfehlung: {winner['Carrier']} – {winner['Versandart']} | Score: {winner['Score']} | Kosten: {winner['Gesamtkosten €']} €")

        st.subheader("Score-Ranking")
        st.dataframe(results, use_container_width=True)

        st.subheader("Preisranking")
        st.dataframe(price_ranking, use_container_width=True)

        st.header("Finale Empfehlung")
        if meta.get("dangerous_goods_manual_check"):
            st.write(
                f"Für **Gefahrgut** gibt EVA keine vollständig automatische Freigabe. "
                f"Die beste prüfbare Option ist **{winner['Carrier']} – {winner['Versandart']}** mit **{winner['Score']} Punkten**. "
                "Vor Versand ist eine manuelle ADR-Prüfung erforderlich."
            )
        else:
            st.write(
                f"Ich empfehle **{winner['Carrier']} – {winner['Versandart']}**. "
                f"Diese Option erzielt mit **{winner['Score']} Punkten** den höchsten Gesamtscore. "
                "Die Bewertung basiert auf Preis, Lieferzeit, Servicequalität, Sicherheit und internationaler Versandfähigkeit."
            )

        st.header("Empfohlene Zusatzleistungen")
        services = recommended_services(st.session_state.eva_inputs.get("Warenart", goods), st.session_state.eva_inputs.get("Warenwert", value), winner)
        if services.empty:
            st.info("Für diese Sendung wurden keine zusätzlichen Services empfohlen.")
        else:
            st.dataframe(services, use_container_width=True)

        with st.expander("Ausgeschlossene Tarife anzeigen"):
            st.dataframe(rejected, use_container_width=True)

# ============================================================
# Chat mit EVA
# ============================================================
st.divider()
st.header("💬 Chat mit EVA")

if "eva_messages" not in st.session_state:
    st.session_state.eva_messages = [
        {"role": "assistant", "content": "Hallo, ich bin EVA – Entscheidungsassistent für die Versanddienstleisterauswahl. Lade deine Excel-Datei hoch, gib die Sendungsdaten ein oder nutze Freitext. Ich vergleiche die Carrier regelbasiert für dich."}
    ]

for msg in st.session_state.eva_messages:
    with st.chat_message(msg["role"], avatar="🤖" if msg["role"] == "assistant" else "👤"):
        st.write(msg["content"])

user_question = st.chat_input("Frage EVA etwas, z. B. 'Warum ist Versicherung nötig?', 'Was ist Volumengewicht?' oder 'Wie wird Gefahrgut behandelt?' ")

if user_question:
    st.session_state.eva_messages.append({"role": "user", "content": user_question})
    with st.chat_message("user", avatar="👤"):
        st.write(user_question)

    answer = rule_based_chat_answer(user_question, results, rejected)
    st.session_state.eva_messages.append({"role": "assistant", "content": answer})
    with st.chat_message("assistant", avatar="🤖"):
        st.write(answer)

# ============================================================
# NEU 4: Protokoll am Seitenende
# ============================================================
st.divider()
with st.expander("📋 Protokoll dieser Sitzung"):
    if not st.session_state.eva_log:
        st.info("Noch keine Berechnung in dieser Sitzung.")
    else:
        log_df = pd.DataFrame(st.session_state.eva_log)
        st.dataframe(log_df, use_container_width=True)
        st.download_button(
            label="Protokoll als CSV herunterladen",
            data=log_to_csv_bytes(),
            file_name="eva_session_log.csv",
            mime="text/csv",
        )
