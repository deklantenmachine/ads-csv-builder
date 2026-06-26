"""
Google Ads Editor CSV builder — genereert per klant twee bestanden:
  {klant}_lokaal_{datum}.csv
  {klant}_stad_{datum}.csv
"""

import re
import io
from datetime import date

import pandas as pd
import requests

# ── constanten sjabloon ──────────────────────────────────────────────────────

TEMPLATE_CITY  = "Groningen"
TEMPLATE_PHONE = "050-7820442"
TEMPLATE_NET   = "050"

SEP      = ";"
ENCODING = "utf-8-sig"

# kolommen waar de plaatsnaam vervangen wordt (beide campagnetypes)
AD_TEXT_COLS = (
    [f"Headline {i}" for i in range(1, 16)]
    + [f"Description {i}" for i in range(1, 5)]
    + ["Path 1", "Path 2", "Link Text", "Description Line 1", "Description Line 2"]
)

# ── helpers ──────────────────────────────────────────────────────────────────

def _today_str() -> str:
    d = date.today()
    return f"{d.day}-{d.month}-{d.year}"


# ── varianten laden ───────────────────────────────────────────────────────────

def load_variants(xlsx_path: str) -> dict:
    """
    Laad het Excel-bestand met korte/lange varianten.
    Geeft dict: {template_tekst_lowercase: (lange_variant_of_None, drempel_int)}

    De korte variant heeft 'Plaats' als placeholder (met of zonder quotes).
    We vervangen die door TEMPLATE_CITY zodat de sleutel overeenkomt met de
    werkelijke waarden uit de sjabloon-CSV.
    """
    df = pd.read_excel(xlsx_path, sheet_name=0, header=None)
    variants: dict[str, tuple[str | None, int]] = {}

    for _, row in df.iterrows():
        korte = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        lange = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        drempel = row.iloc[2] if pd.notna(row.iloc[2]) else None

        # sla header-rijen en lege rijen over
        if not korte or korte in ("nan",) or drempel is None:
            continue
        try:
            drempel = int(drempel)
        except (ValueError, TypeError):
            continue

        # normaliseer: verwijder quotes rond 'Plaats' en vervang door TEMPLATE_CITY
        korte_norm = korte.replace("'Plaats'", TEMPLATE_CITY).replace("Plaats'", TEMPLATE_CITY).replace("'Plaats", TEMPLATE_CITY).replace("Plaats", TEMPLATE_CITY)
        lange_val  = lange if lange and lange != "nan" else None

        variants[korte_norm.lower()] = (lange_val, drempel)

    return variants


def _apply_variant(val: str, city: str, variants: dict) -> str:
    """
    Geeft de juiste tekstvariant terug voor deze cel-waarde en plaatsnaam.
    - Als cel matcht met korte variant ÉN len(city) >= drempel → lange variant
    - Anders → normale plaatsnaam-vervanging
    - Lange variant is None → lege string (bijv. Path weggooien)
    """
    if not val or pd.isna(val):
        return val

    key = str(val).lower()
    if key in variants:
        lange, drempel = variants[key]
        if len(city) >= drempel:
            return lange if lange is not None else ""
    return _case_replace(str(val), TEMPLATE_CITY, city)


def _match_case(new: str, matched: str) -> str:
    """
    Pas het case-patroon van matched toe op new.
    - ALL CAPS  → new.upper()
    - all lower → new.lower()
    - Mixed/titel → new ongewijzigd (heeft al eigen correcte spelling)
    """
    if matched.isupper():
        return new.upper()
    if matched.islower():
        return new.lower()
    return new


def _case_replace(original: str, old: str, new: str) -> str:
    """Whole-word, case-insensitive vervanging; behoudt hoofdletterpatroon."""
    # Bouw pattern dat ook koppeltekens als woorddeel behandelt
    pattern = r"(?<!['\w])" + re.escape(old) + r"(?!['\w])"
    def replacer(m: re.Match) -> str:
        return _match_case(new, m.group())
    return re.sub(pattern, replacer, original, flags=re.IGNORECASE)


def _replace_in_col(series: pd.Series, old: str, new: str) -> pd.Series:
    def _safe(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        return _case_replace(str(val), old, new)
    return series.apply(_safe)


def _replace_exact(series: pd.Series, old: str, new: str) -> pd.Series:
    def _safe(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        return str(val).replace(old, new)
    return series.apply(_safe)


def _fix_labels(series: pd.Series, old_net: str, new_net: str) -> pd.Series:
    """Vervang netnummer alleen in cellen die 'Netnummer' bevatten."""
    def _safe(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        s = str(val)
        if "Netnummer" in s:
            return s.replace(f"Netnummer {old_net}", f"Netnummer {new_net}")
        return s
    return series.apply(_safe)


def _parse_locations(raw: str) -> list[tuple[str, str]]:
    """
    Parst 'Negatieve locaties' of 'Locatie' cel.
    Formaat: "ID, Naam[,Provincie,Land] | ID, Naam[,...] | ..."
    Geeft lijst van (id_str, naam_str) tuples.
    """
    result = []
    if not raw or pd.isna(raw) or str(raw).strip() == "":
        return result
    for part in str(raw).split("|"):
        part = part.strip()
        if not part:
            continue
        # eerste getal vóór de eerste komma is het ID
        idx = part.find(",")
        if idx == -1:
            continue
        loc_id   = part[:idx].strip()
        loc_name = part[idx + 1:].strip()
        result.append((loc_id, loc_name))
    return result


# ── Google Sheets ophalen ─────────────────────────────────────────────────────

def load_sheet(url: str, sheet_name: str = "DakPro NL Plaatsen") -> pd.DataFrame:
    """
    Laad Google Sheets tab als CSV.
    Accepteert /edit URL of directe export URL.
    """
    if "export?format=csv" not in url:
        base = url.split("/edit")[0]
        url  = f"{base}/export?format=csv&sheet={requests.utils.quote(sheet_name)}"

    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    df = pd.read_csv(io.StringIO(resp.text))
    df.columns = [c.strip() for c in df.columns]
    return df


def validate_sheet(df: pd.DataFrame) -> list[str]:
    """Controleer verplichte kolommen. Geeft lijst van foutmeldingen."""
    required = ["Plaats", "Dubbele plaats", "Netnummer", "Telefoonnummer",
                "URL", "Locatie", "Negatieve locaties", "Klant", "Ad Schedule"]
    missing = [c for c in required if c not in df.columns]
    return [f"Kolom ontbreekt in sheet: '{c}'" for c in missing]


# ── kern: locatierijen vervangen ──────────────────────────────────────────────

def _build_location_rows(template_row: pd.Series, targeting: list[tuple], negatives: list[tuple]) -> list[pd.Series]:
    """
    Bouw nieuwe locatierijen op basis van targeting + negatives.
    Kopieert niet-locatie-kolommen uit template_row (campagnenaam etc.).
    """
    rows = []
    loc_cols = {"Location", "ID", "Criterion Type", "Location groups", "Reach"}

    def _base(loc_id: str, loc_name: str, criterion: str) -> pd.Series:
        r = template_row.copy()
        # wis locatie-specifieke kolommen
        for c in loc_cols:
            if c in r.index:
                r[c] = None
        r["ID"]             = loc_id
        r["Location"]       = loc_name
        r["Criterion Type"] = criterion
        r["Location groups"] = None
        r["Reach"]          = None
        return r

    for loc_id, loc_name in targeting:
        rows.append(_base(loc_id, loc_name, ""))

    for loc_id, loc_name in negatives:
        rows.append(_base(loc_id, loc_name, "Campaign Negative"))

    return rows


# ── kern: één stad verwerken ─────────────────────────────────────────────────

def process_city(
    df_lokaal: pd.DataFrame,
    df_stad:   pd.DataFrame,
    row:       pd.Series,
    variants:  dict | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Geeft (lokaal_df, stad_df) voor één stad.
    Gooit ValueError als verplichte velden leeg zijn.
    """
    city        = str(row["Plaats"]).strip()
    ad_schedule = str(row.get("Ad Schedule", "")).strip()
    # Netnummer: haal uit telefoonnummer als leidende nul ontbreekt (sheets slaat 0413 op als 413.0)
    net_raw = str(row.get("Netnummer", "")).strip()
    if net_raw in ("", "nan"):
        net = ""
    else:
        # converteer float-notatie (413.0) naar int-string (413)
        try:
            net = str(int(float(net_raw)))
        except ValueError:
            net = net_raw
        # herstel leidende nul vanuit telefoonnummer indien aanwezig
        phone_raw = str(row.get("Telefoonnummer", "")).strip()
        if "-" in phone_raw:
            tel_prefix = phone_raw.split("-")[0]
            if tel_prefix.lstrip("0") == net:
                net = tel_prefix  # bijv. "0413"
    phone = str(row.get("Telefoonnummer", "")).strip()
    url   = str(row.get("URL", "")).strip()
    klant = str(row.get("Klant", "onbekend")).strip()

    # lege waarden: gebruik sjabloonwaarden als fallback
    if not net:   net   = TEMPLATE_NET
    if not phone: phone = TEMPLATE_PHONE

    # locaties parsen
    targeting_raw = str(row.get("Locatie", "")).strip()
    negatives_raw = str(row.get("Negatieve locaties", "")).strip()
    targeting_locs = _parse_locations(targeting_raw)
    negative_locs  = _parse_locations(negatives_raw)

    today = _today_str()

    lok = df_lokaal.copy()
    sta = df_stad.copy()

    for df in (lok, sta):
        # Start Date
        mask = df["Start Date"].notna() & (df["Start Date"].astype(str).str.strip() != "")
        df.loc[mask, "Start Date"] = today

        # End Date leegmaken
        mask = df["End Date"].notna() & (df["End Date"].astype(str).str.strip() != "")
        df.loc[mask, "End Date"] = ""

        # Ad Schedule
        mask = df["Ad Schedule"].notna() & (df["Ad Schedule"].astype(str).str.strip() != "")
        df.loc[mask, "Ad Schedule"] = ad_schedule

        # Labels: alleen Netnummer-cellen aanpassen
        df["Labels"] = _fix_labels(df["Labels"], TEMPLATE_NET, net)

        # Phone Number
        if phone and phone != TEMPLATE_PHONE:
            df["Phone Number"] = _replace_exact(df["Phone Number"], TEMPLATE_PHONE, phone)

        # Final URL
        if url:
            mask = df["Final URL"].notna() & (df["Final URL"].astype(str).str.strip() != "")
            df.loc[mask, "Final URL"] = url

        # Campaign Status → Paused
        mask = df["Campaign Status"].notna() & (df["Campaign Status"].astype(str).str.strip() != "")
        df.loc[mask, "Campaign Status"] = "Paused"

        # Plaatsnaam in advertentietekst (met varianten indien geladen)
        for col in AD_TEXT_COLS:
            if col not in df.columns:
                continue
            if variants:
                df[col] = df[col].apply(
                    lambda v: _apply_variant(v, city, variants)
                    if (pd.notna(v) and str(v).strip())
                    else v
                )
            else:
                df[col] = _replace_in_col(df[col], TEMPLATE_CITY, city)

    # ── + Stad specifiek ─────────────────────────────────────────────────────
    sta["Ad Group"] = _replace_in_col(sta["Ad Group"], TEMPLATE_CITY, city)
    sta["Keyword"]  = _replace_in_col(sta["Keyword"],  TEMPLATE_CITY, city)

    # ── Lokaal specifiek ─────────────────────────────────────────────────────
    lok["Campaign"] = _replace_in_col(lok["Campaign"], TEMPLATE_CITY, city)

    # Locatierijen vervangen (alleen als er locatiedata in de sheet staat)
    if targeting_locs or negative_locs:
        loc_filled = lok["Location"].notna() & (lok["Location"].astype(str).str.strip() != "")
        # bewaar één willekeurige locatierij als template voor campagnenaam etc.
        template_loc_row = lok[loc_filled].iloc[0] if loc_filled.any() else lok.iloc[0]

        # verwijder alle bestaande locatierijen
        lok = lok[~loc_filled].copy()

        # bouw nieuwe locatierijen
        new_rows = _build_location_rows(template_loc_row, targeting_locs, negative_locs)
        if new_rows:
            lok = pd.concat([lok, pd.DataFrame(new_rows)], ignore_index=True)
    else:
        # sheet heeft nog geen locatiedata: pas bestaande targeting-rij aan zoals eerder
        loc_filled = lok["Location"].notna() & (lok["Location"].astype(str).str.strip() != "")
        crit_col   = "Criterion Type"
        targeting_mask = loc_filled & (
            lok[crit_col].isna() | (lok[crit_col].astype(str).str.strip() == "")
        )
        lok["Location groups"] = lok["Location groups"].astype(object)
        if targeting_mask.any() and targeting_raw:
            lok.loc[targeting_mask, "Location"] = targeting_raw

    return lok, sta


# ── hoofd-functie ────────────────────────────────────────────────────────────

def build_all(
    lokaal_path:   str,
    stad_path:     str,
    sheet_url:     str,
    sheet_name:    str = "DakPro NL Plaatsen",
    variants_path: str | None = None,
    klant_filter:  list[str] | None = None,
    progress_cb=None,
) -> tuple[dict, list[str]]:
    """
    Verwerkt alle steden en groepeert output per klant.
    Geeft: ({klant: {"lokaal": df, "stad": df}}, [errors])
    """
    df_lokaal = pd.read_csv(lokaal_path, sep=SEP, encoding=ENCODING, low_memory=False)
    df_stad   = pd.read_csv(stad_path,   sep=SEP, encoding=ENCODING, low_memory=False)
    sheet     = load_sheet(sheet_url, sheet_name)

    variants = load_variants(variants_path) if variants_path else None

    # valideer kolommen
    val_errors = validate_sheet(sheet)
    if val_errors:
        raise ValueError("\n".join(val_errors))

    results: dict[str, dict[str, list]] = {}
    errors:  list[str] = []
    total = len(sheet)

    for i, (_, row) in enumerate(sheet.iterrows()):
        city  = str(row.get("Plaats", "")).strip()
        klant = str(row.get("Klant", "onbekend")).strip()
        if not klant or klant in ("nan", "onbekend"):
            klant = "DakPro"

        # Klantfilter
        if klant_filter and klant not in klant_filter:
            continue

        # Dubbele plaatsen overslaan (beide campagnetypes)
        if str(row.get("Dubbele plaats", "")).strip().lower() == "x":
            if progress_cb:
                progress_cb(city, i, total, skipped=True)
            continue

        if progress_cb:
            progress_cb(city, i, total, skipped=False)

        try:
            lok_df, sta_df = process_city(df_lokaal, df_stad, row, variants)
        except Exception as exc:
            errors.append(f"{city}: {exc}")
            continue

        if klant not in results:
            results[klant] = {"lokaal": [], "stad": []}
        results[klant]["lokaal"].append(lok_df)
        results[klant]["stad"].append(sta_df)

    merged: dict[str, dict[str, pd.DataFrame]] = {}
    for klant, parts in results.items():
        merged[klant] = {
            "lokaal": pd.concat(parts["lokaal"], ignore_index=True),
            "stad":   pd.concat(parts["stad"],   ignore_index=True),
        }

    return merged, errors


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_csv(buf, sep=SEP, encoding=ENCODING, index=False)
    return buf.getvalue()


def output_filename(klant: str, kind: str) -> str:
    return f"{klant}_{kind}_{_today_str()}.csv"
