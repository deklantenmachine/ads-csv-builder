"""
Google Ads Editor CSV builder — genereert per klant twee bestanden:
  {klant}_lokaal_{datum}.csv
  {klant}_stad_{datum}.csv
"""

import re
import io
from datetime import date
from urllib.parse import urlparse

import pandas as pd
import requests

from advice_parser import CpcAdviceImport, PauseAdviceImport, CampaignType, normalize_name
from rules_engine import CampaignBuildRulesEngine, extract_template_bid_strategy

# ── constanten sjabloon ──────────────────────────────────────────────────────

TEMPLATE_CITY           = "Groningen"
TEMPLATE_PHONE          = "050-7820442"
TEMPLATE_NET            = "050"
TEMPLATE_COMPANY        = "DCN"                  # bedrijfsnaam placeholder in eigen merk CSV/Excel
TEMPLATE_DOMAIN         = "dcn-dakdekkers.nl"    # domein in eigen merk sjabloon-CSV
TEMPLATE_USP            = "Familiebedrijf Sinds 1966"  # USP placeholder in eigen merk CSV
TEMPLATE_PRICE          = "€59,50"                 # prijsplaatshouder in eigen merk CSV
TEMPLATE_REVIEW_SCORE   = "9,6"                  # review score in portaal Excel-varianten
TEMPLATE_JAREN_GARANTIE = "10"                   # jaren garantie in portaal Excel-varianten

# Genormaliseerde placeholders voor variant-lookup (lowercase)
_KEY_REVIEW = "[reviewscore]"
_KEY_JAREN  = "[jarengarantie]"

# Plaatsen die Google als algemeen woord herkent → nooit +Stad campagne genereren
EXCLUDED_STAD_CITIES: frozenset[str] = frozenset({
    # Nederland
    "huizen", "zetten", "best", "dieren", "echt",
    "hoornaar", "heel", "enter", "zeeland", "noorden",
    "handel", "een",
    # België
    "menen", "balen", "geel", "halen", "as", "landen", "peer",
    "boom", "mol", "lint", "meer", "wortel", "paal", "beek",
    "lier", "schoten", "duffel", "weelde", "beringen", "ranst",
    "wichelen", "deerlijk", "heers", "wellen", "schaffen",
})

SEP      = ";"
ENCODING = "utf-8-sig"

# Klantnaam (plaatsensheet) → Google Ads accountnaam in het centrale CPC-blad (Schoorsteenveger).
_SCHOORSTEENVEGER_ACCOUNT_MAP: dict[str, dict[str, str]] = {
    "algemeen onderhoud nederland":              {"eigen": "Van Beynen Schoorsteenvegers",      "portaal": "De Schoorsteenbrigade (AON)"},
    "schoorsteenveegbedrijf van nieuwenhoven":   {"eigen": "VNH Schoorsteenvegers",             "portaal": "De Schoorsteenbrigade (VNH)"},
    "uw dakbeheerder":                           {"eigen": "Uw Schoorsteenveger",               "portaal": "DSB Uw Schoorsteenveger"},
    "schoorsteenvegersbedrijf de spil":          {"eigen": "De Spil Schoorsteenvegersbedrijf",  "portaal": "De Spil Schoorsteenbrigade"},
    "master cleaner bvba":                       {"eigen": "Schoorsteenveger Master Cleaner",   "portaal": "DSB Master Cleaner"},
    "schoorsteenreiniger filip":                 {"eigen": "Schoorsteenveger Filip",            "portaal": "De Schoorsteenbrigade BE Filip"},
    "multi-reiniging service":                   {"eigen": "MRS Schoorsteenveger",              "portaal": "De Schoorsteenbrigade (MRS)"},
    "multi-reiniging service oudwoude":          {"eigen": "MRS Schoorsteenveger",              "portaal": "De Schoorsteenbrigade (MRS)"},
    "schoorsteenveger bams":                     {"eigen": "Schoorsteenveger Bams",             "portaal": "De Schoorsteenbrigade (Bams)"},
    "schoorsteenveger johan":                    {"eigen": "Schoorsteenveger Johan",            "portaal": "De Schoorsteenbrigade (West-Vlaanderen)"},
}


def _resolve_sv_cpc_account(klant: str, is_portaal: bool) -> str:
    """Vertaal klantnaam (plaatsensheet) → accountnaam in het centrale Schoorsteenveger CPC-blad."""
    key   = "portaal" if is_portaal else "eigen"
    entry = _SCHOORSTEENVEGER_ACCOUNT_MAP.get(klant.strip().lower())
    return entry[key] if entry else klant


# Expliciete koppeling van klantnaam (spreadsheet) → accountnaam in CPC/pauzebestand.
# Wordt gebruikt vóór fuzzy matching zodat naamverschillen niet tot foute matches leiden.
_KLANT_TO_CPC_ACCOUNT: dict[str, str] = {
    "all in dakonderhoud":                      "All In Dakdekkers",
    "b en b dakonderhoud":                      "B&B Dakdekker",
    "van gestel onderhoud":                     "Boersma Dakdekkers",
    "woningverbetering buitenhuis":             "Dakdekker Buitenhuis",
    "dakcentrale nederland":                    "DCN Dakdekkers",
    "dakcentraal":                              "Dakcentraal",
    "debro loodgieterswerken":                  "Debro dakdekkers",
    "dakbeheer holland":                        "Holland Dakdekkers",
    "jouw dak expert b.v.":                     "Jouw Dak Expert",
    "dakreparatie nederland":                   "Jansen Dakdekker",
    "nvr allround":                             "NVR Dakdekkers",
    "rovaal dakbeheer":                         "SH Dakdekkers",
    "dakbedekking van eerd":                    "Dakbedekking van Eerd",
    "van den bogaert isolatie en daktechniek":  "Van den Bogaert Daktechniek",
    "highservice woningonderhoud":              "Dakdekkersbedrijf Voets",
    "dakspecialist vriezen bv":                 "Vriezen Dakdekkers",
}


def _resolve_cpc_account(klant: str) -> str:
    """Vertaal klantnaam uit spreadsheet naar accountnaam in het adviesbestand."""
    return _KLANT_TO_CPC_ACCOUNT.get(klant.strip().lower(), klant)

# kolommen waar de plaatsnaam (en bedrijfsnaam) vervangen wordt
AD_TEXT_COLS = (
    [f"Headline {i}" for i in range(1, 16)]
    + [f"Description {i}" for i in range(1, 5)]
    + ["Path 1", "Path 2", "Link Text", "Description Line 1", "Description Line 2",
       "Callout text"]
)

# ── helpers ──────────────────────────────────────────────────────────────────

def _today_str() -> str:
    d = date.today()
    return f"{d.day}-{d.month}-{d.year}"


def _match_case(new: str, matched: str) -> str:
    if matched.isupper():
        return new.upper()
    if matched.islower():
        return new.lower()
    return new


def _case_replace(original: str, old: str, new: str) -> str:
    """Whole-word, case-insensitive vervanging; behoudt hoofdletterpatroon."""
    pattern = r"(?<!['\w])" + re.escape(old) + r"(?!['\w])"
    return re.sub(pattern, lambda m: _match_case(new, m.group()), original, flags=re.IGNORECASE)


def _replace_in_col(series: pd.Series, old: str, new: str) -> pd.Series:
    def _safe(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        return _case_replace(str(val), old, new)
    return series.apply(_safe)


def _uses_explicit_placeholders(df: pd.DataFrame) -> bool:
    """Detecteer of het sjabloon expliciete placeholders gebruikt zoals [Plaats], [Klantnaam]."""
    for col in ["Final URL", "Campaign", "Headline 1", "Ad Group", "Keyword"]:
        if col in df.columns:
            vals = df[col].dropna().astype(str)
            if vals.str.contains(r"\[Plaats\]", regex=True).any():
                return True
    return False


def _apply_explicit_text(val, city: str, merk_info: dict | None, price: str | None, col: str = "") -> str:
    """Vervang expliciete placeholders [Plaats], [Klantnaam], [Prijs] in één cel."""
    if pd.isna(val) or not str(val).strip():
        return val
    s = str(val)
    s = s.replace("[Plaats]", city)
    if merk_info and "[Klantnaam]" in s:
        korte = merk_info.get("korte_naam", "")
        lange = merk_info.get("lange_naam", "") or korte
        limit = _COL_CHAR_LIMITS.get(col)
        if limit:
            bedrijfsnaam = korte if len(s.replace("[Klantnaam]", lange)) > limit else lange
        else:
            bedrijfsnaam = lange
        s = s.replace("[Klantnaam]", bedrijfsnaam)
    if price:
        s = re.sub(r"\[Prijs\]", price, s, flags=re.IGNORECASE)
    return s


def _replace_exact(series: pd.Series, old: str, new: str) -> pd.Series:
    def _safe(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        return str(val).replace(old, new)
    return series.apply(_safe)


def _ensure_netnummer_label(df: pd.DataFrame, net: str) -> pd.DataFrame:
    """Zet Labels op uitsluitend 'Netnummer {net}' voor alle rijen — verwijdert andere labels."""
    label = f"Netnummer {net}"

    def _add(val):
        return label  # altijd alleen netnummer-label bewaren

    has_campaign = df["Campaign"].notna() & (df["Campaign"].astype(str).str.strip() != "")
    has_location = df["Location"].notna() & (df["Location"].astype(str).str.strip() != "")
    mask = has_campaign & ~has_location

    df = df.copy()
    if "Labels" in df.columns and mask.any():
        df["Labels"] = df["Labels"].astype(object)
        df.loc[mask, "Labels"] = label
    return df


def _parse_locations(raw: str) -> list[tuple[str, str]]:
    """Parst 'ID, Naam | ID, Naam | ...' naar lijst van (id, naam) tuples."""
    result = []
    if not raw or pd.isna(raw) or str(raw).strip() == "":
        return result
    for part in str(raw).split("|"):
        part = part.strip()
        if not part:
            continue
        idx = part.find(",")
        if idx == -1:
            continue
        result.append((part[:idx].strip(), part[idx + 1:].strip()))
    return result


# ── varianten laden ───────────────────────────────────────────────────────────

def _parse_threshold(raw) -> dict:
    """
    Parst de drempelwaarde-cel uit het Excel-variantenbestand.
    Geeft een rule-dict:
      {"rule": "city",   "threshold": N, "name_max_len": M}  # stadslengte-drempel
      {"rule": "always"}                                      # altijd vervangen
      {"rule": "usp"}                                         # USP-logica
    """
    s = str(raw).strip() if pd.notna(raw) else ""
    if not s or s == "nan":
        return {}

    s_lower = s.lower()

    if "usp" in s_lower or "geen andere" in s_lower:
        return {"rule": "usp"}

    if "altijd" in s_lower:
        return {"rule": "always"}

    m = re.search(r"\d+", s)
    if not m:
        return {}

    result: dict = {"rule": "city", "threshold": int(m.group())}

    # bijv. "mocht 'lange bedrijfsnaam' langer zijn dan 20 tekens"
    m2 = re.search(r"langer.*?(\d+)\s*tekens", s_lower)
    if m2:
        result["name_max_len"] = int(m2.group(1))

    return result


def load_variants(xlsx_path: str) -> dict:
    """
    Laad het Excel-bestand met korte/lange varianten.

    Geeft dict:
      {lookup_key_lower: {"lange_raw": str|None, "rule_info": dict}}

    Lookup key = korte variant met:
      - 'Plaats' / Plaats → TEMPLATE_CITY
      - DCN → [Bedrijfsnaam]   (zodat CSV-waarden direct matchen)
    """
    df = pd.read_excel(xlsx_path, sheet_name=0, header=None)
    variants: dict = {}

    for _, row in df.iterrows():
        korte   = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        lange   = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        drempel = row.iloc[2] if len(row) > 2 else None

        if not korte or korte == "nan":
            continue

        rule_info = _parse_threshold(drempel)
        if not rule_info:
            continue

        # normaliseer de lookup-sleutel (oud formaat: 'Plaats'/DCN, nieuw: [Plaats]/[Klantnaam])
        key = korte
        key = key.replace("[Plaats]", TEMPLATE_CITY)           # nieuw: [Plaats] → Groningen
        key = re.sub(r"'?Plaats'?", TEMPLATE_CITY, key)        # oud: 'Plaats' → Groningen
        key = key.replace("[Klantnaam]", "[Bedrijfsnaam]")     # nieuw: [Klantnaam] → [Bedrijfsnaam]
        key = key.replace(TEMPLATE_COMPANY, "[Bedrijfsnaam]")  # oud: DCN → [Bedrijfsnaam]
        key = key.lower()
        # normaliseer portaal-specifieke waarden naar generieke placeholders
        key = key.replace(TEMPLATE_REVIEW_SCORE, _KEY_REVIEW)
        key = re.sub(r"\b" + re.escape(TEMPLATE_JAREN_GARANTIE) + r"\b jaar", _KEY_JAREN + " jaar", key)

        lange_val = lange if lange and lange != "nan" else None

        variants[key] = {"lange_raw": lange_val, "rule_info": rule_info}

    return variants


def _resolve_name(
    lange_raw: str,
    korte_naam: str,
    lange_naam: str,
    name_max_len: int | None,
    col_char_limit: int | None = None,
) -> str:
    """Vervang naam-placeholders in de lange-variant tekst (enkele quotes én vierkante haken).

    Kiest korte naam als:
    1. name_max_len is gezet en len(lange_naam) > name_max_len, OF
    2. col_char_limit is gezet en het resultaat met lange naam de limiet overschrijdt.
    """
    if name_max_len and len(lange_naam) > name_max_len:
        effective_lange = korte_naam
    elif col_char_limit:
        candidate = (lange_raw
            .replace("'Lange bedrijfsnaam'", lange_naam)
            .replace("[Lange bedrijfsnaam]", lange_naam)
            .replace("[Klantnaam]", lange_naam))
        effective_lange = korte_naam if len(candidate) > col_char_limit else lange_naam
    else:
        effective_lange = lange_naam
    result = lange_raw
    result = result.replace("'Lange bedrijfsnaam'", effective_lange)
    result = result.replace("[Lange bedrijfsnaam]", effective_lange)
    result = result.replace("[Klantnaam]", effective_lange)
    result = result.replace("'Korte bedrijfsnaam'", korte_naam)
    result = result.replace("[Korte bedrijfsnaam]", korte_naam)
    return result


def _apply_variant(val: str, city: str, variants: dict, merk_info: dict | None = None, col: str = "") -> str:
    """
    Bepaal de juiste tekstvariant voor één cel.

    Portaal mode (merk_info is None): alleen stadslengte-drempel.
    Eigen merk mode: ook bedrijfsnaam, USP, altijd-vervang regels.
    """
    if not val or pd.isna(val):
        return val

    s = str(val)
    # normaliseer lookup-sleutel: zelfde transformaties als load_variants
    key = s.lower()
    key = key.replace("[plaats]", TEMPLATE_CITY.lower())            # nieuw: [Plaats] → groningen
    key = key.replace("[klantnaam]", "[bedrijfsnaam]")              # nieuw: [Klantnaam] → [Bedrijfsnaam]
    key = key.replace(TEMPLATE_COMPANY.lower(), "[bedrijfsnaam]")   # oud: dcn → [bedrijfsnaam]
    key = key.replace("[reviewscore]", _KEY_REVIEW).replace("[jarengarantie]", _KEY_JAREN)

    if key in variants:
        entry     = variants[key]
        lange_raw = entry["lange_raw"]
        rule      = entry["rule_info"]
        r         = rule.get("rule")

        if r == "usp" and merk_info:
            usp = merk_info.get("usp", "").strip()
            if usp:
                return usp
            return lange_raw if lange_raw else s  # fallback uit Excel

        col_limit = _COL_CHAR_LIMITS.get(col)

        if r == "always" and merk_info:
            if not lange_raw:
                return s
            korte = merk_info.get("korte_naam", "")
            lange = merk_info.get("lange_naam", "")
            result = _resolve_name(lange_raw, korte, lange, rule.get("name_max_len"), col_limit)
            result = _apply_portaal_values(result, merk_info)
            return _case_replace(result, TEMPLATE_CITY, city)

        if r == "city":
            threshold    = rule.get("threshold", 999)
            name_max_len = rule.get("name_max_len")
            # Check ook of de volledige tekst na invullen de kolomlimiet overschrijdt
            _korte_est = merk_info.get("korte_naam", "") if merk_info else ""
            _estimated = (s
                .replace("[Klantnaam]", _korte_est)
                .replace("[Plaats]", city)
                .replace(TEMPLATE_CITY, city))
            _result_too_long = col_limit is not None and len(_estimated) > col_limit
            if len(city) >= threshold or _result_too_long:
                if not lange_raw:
                    return ""
                # Geketende drempel: volg de keten zolang:
                #   - de stad het volgende drempelniveau haalt, OF
                #   - het tussenresultaat de kolomlimiet al overschrijdt.
                # Voorbeeld: "Dak Vervangen In Groningen?" (t=13) → "Nieuw Dak In Groningen?" (t=18)
                #   Winterswijk Ratum (17 tekens): 17 < 18 maar "Nieuw Dak In Winterswijk Ratum?"
                #   is 31 tekens (> 30 kolomlimiet) → toch naar "Tijd Voor Een Nieuw Dak?".
                for _ in range(10):
                    sub_key = lange_raw.lower()
                    sub_key = sub_key.replace(TEMPLATE_COMPANY.lower(), "[bedrijfsnaam]")
                    sub_key = sub_key.replace(TEMPLATE_REVIEW_SCORE, _KEY_REVIEW)
                    sub_key = re.sub(r"\b" + re.escape(TEMPLATE_JAREN_GARANTIE) + r"\b jaar", _KEY_JAREN + " jaar", sub_key)
                    if sub_key not in variants:
                        break
                    sub_entry = variants[sub_key]
                    sub_rule  = sub_entry["rule_info"]
                    if sub_rule.get("rule") != "city":
                        break
                    candidate      = _case_replace(lange_raw, TEMPLATE_CITY, city)
                    threshold_ok   = len(city) >= sub_rule.get("threshold", 999)
                    result_too_long = col_limit is not None and len(candidate) > col_limit
                    if threshold_ok or result_too_long:
                        lange_raw    = sub_entry["lange_raw"] or lange_raw
                        name_max_len = sub_rule.get("name_max_len", name_max_len)
                    else:
                        break
                if merk_info:
                    korte = merk_info.get("korte_naam", "")
                    lange = merk_info.get("lange_naam", "")
                    result = _resolve_name(lange_raw, korte, lange, name_max_len, col_limit)
                    result = _apply_portaal_values(result, merk_info)
                else:
                    result = lange_raw
                return _case_replace(result, TEMPLATE_CITY, city)

    # geen variant match: standaard plaatsnaam-vervanging
    result = _case_replace(s, TEMPLATE_CITY, city)

    # eigen merk: vervang ook [Bedrijfsnaam] placeholder die niet via varianten liep
    if merk_info:
        result = _apply_eigen_placeholders(result, merk_info, city)

    return result


# Google Ads tekenlimiet per kolomtype
_COL_CHAR_LIMITS: dict[str, int] = {
    **{f"Headline {i}": 30 for i in range(1, 16)},
    **{f"Description {i}": 90 for i in range(1, 5)},
    "Description Line 1": 90,
    "Description Line 2": 90,
    "Path 1": 15,
    "Path 2": 15,
    "Callout text": 25,
    "Link Text": 25,
}


def _apply_eigen_placeholders(
    text: str, merk_info: dict, city: str, col: str = "", tmpl_company: str = TEMPLATE_COMPANY
) -> str:
    """Vervang [Bedrijfsnaam], [ReviewScore], [JarenGarantie] en sjabloon-bedrijfsnaam in één string.

    Kiest automatisch de korte bedrijfsnaam als de lange naam de tekenlimiet
    voor de betreffende kolom zou overschrijden.
    """
    korte = merk_info.get("korte_naam", "")
    lange = merk_info.get("lange_naam", "") or korte

    limit = _COL_CHAR_LIMITS.get(col)

    # Vervang sjabloon-bedrijfsnaam (bijv. "Van Beynen" of "DCN") door echte naam
    if tmpl_company and tmpl_company in text:
        # Gebruik lange naam tenzij die te lang maakt
        if limit:
            met_lange = text.replace(tmpl_company, lange)
            bedrijfsnaam = korte if len(met_lange) > limit else lange
        else:
            bedrijfsnaam = lange
        text = _case_replace(text, tmpl_company, bedrijfsnaam)

    # Vervang ook generieke [Bedrijfsnaam] placeholder (voor Dakdekker-stijl varianten)
    if limit and "[Bedrijfsnaam]" in text:
        met_lange = text.replace("[Bedrijfsnaam]", lange)
        bedrijfsnaam = korte if len(met_lange) > limit else lange
    else:
        bedrijfsnaam = lange
    text = text.replace("[Bedrijfsnaam]", bedrijfsnaam)

    text = text.replace("[ReviewScore]",   merk_info.get("review_score",    ""))
    text = text.replace("[JarenGarantie]", merk_info.get("jaren_garantie", ""))
    return text



def _apply_portaal_values(text: str, merk_info: dict) -> str:
    """Vervang portaal-specifieke vaste waarden in een lange variant door eigen merk waarden."""
    rs = merk_info.get("review_score", "").strip()
    jg = merk_info.get("jaren_garantie", "").strip()
    if rs and rs != TEMPLATE_REVIEW_SCORE:
        text = text.replace(TEMPLATE_REVIEW_SCORE, rs)
    if jg and jg != TEMPLATE_JAREN_GARANTIE:
        # vervang "10 jaar" / "10 Jaar" maar niet "100%" etc.
        text = re.sub(r"\b" + re.escape(TEMPLATE_JAREN_GARANTIE) + r"\b(?= [Jj]aar)", jg, text)
    return text


# ── Google Sheets ophalen ─────────────────────────────────────────────────────

def load_sheet(url: str, sheet_name: str = "DakPro NL Plaatsen") -> pd.DataFrame:
    if "export?format=csv" not in url:
        base = url.split("/edit")[0]
        url  = f"{base}/export?format=csv&sheet={requests.utils.quote(sheet_name)}"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    df = pd.read_csv(io.StringIO(resp.text))
    df.columns = [c.strip() for c in df.columns]
    return df


def validate_sheet(df: pd.DataFrame) -> list[str]:
    required = ["Plaats", "Dubbele plaats", "Netnummer", "Telefoonnummer",
                "URL", "Locatie", "Negatieve locaties", "Klant"]
    missing = [c for c in required if c not in df.columns]
    return [f"Kolom ontbreekt in sheet: '{c}'" for c in missing]


# ── locatierijen ─────────────────────────────────────────────────────────────

def _build_location_rows(template_row: pd.Series, targeting: list, negatives: list) -> list:
    rows = []
    loc_cols = {"Location", "ID", "Criterion Type", "Location groups", "Reach"}

    def _base(loc_id, loc_name, criterion):
        r = template_row.copy()
        for c in loc_cols:
            if c in r.index:
                r[c] = None
        r["ID"]              = loc_id
        r["Location"]        = loc_name
        r["Criterion Type"]  = criterion
        r["Location groups"] = None
        r["Reach"]           = None
        return r

    for loc_id, loc_name in targeting:
        rows.append(_base(loc_id, loc_name, ""))
    for loc_id, loc_name in negatives:
        rows.append(_base(loc_id, loc_name, "Campaign Negative"))
    return rows


# ── kern: één stad verwerken ──────────────────────────────────────────────────

def process_city(
    df_lokaal:  pd.DataFrame,
    df_stad:    pd.DataFrame,
    row:        pd.Series,
    variants:   dict | None = None,
    merk_info:  dict | None = None,   # None = portaal mode
    ad_schedule_override: str = "",   # tool-veld overschrijft sheet indien ingevuld
    tmpl_city:          str | None = None,
    tmpl_company:       str | None = None,
    tmpl_price:         str | None = None,
    portaal_klantnaam:  str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:

    city  = str(row["Plaats"]).strip()
    _tc   = tmpl_city    or TEMPLATE_CITY
    _tco  = tmpl_company or TEMPLATE_COMPANY
    _tp   = tmpl_price   or None

    # Ad Schedule: tool-veld heeft voorrang, anders uit sheet
    ad_schedule = ad_schedule_override.strip() or str(row.get("Ad Schedule", "")).strip()

    # Netnummer (herstel leidende nul via telefoonnummer)
    net_raw   = str(row.get("Netnummer", "")).strip()
    phone_raw = str(row.get("Telefoonnummer", "")).strip()
    if net_raw in ("", "nan"):
        net = ""
    else:
        # Strip non-numeric prefix (e.g. "Netnummer 56" → "56")
        import re as _re
        _digits_only = _re.sub(r'^[^\d]*', '', net_raw).strip()
        if _digits_only:
            net_raw = _digits_only
        try:
            net = str(int(float(net_raw)))
        except ValueError:
            net = net_raw
        # Herstel leidende nul via telefoonnummer (NL: 056-..., BE: 056 89 50 11 of 056/...)
        for _sep in ("-", "/", " "):
            if _sep in phone_raw:
                tel_prefix = phone_raw.split(_sep)[0]
                if tel_prefix.lstrip("0") == net:
                    net = tel_prefix
                break

    phone = phone_raw
    url   = str(row.get("URL", "")).strip()

    if not net:   net   = TEMPLATE_NET
    if not phone: phone = TEMPLATE_PHONE

    targeting_raw_val = str(row.get("Locatie", "")).strip()
    skip_lokaal    = targeting_raw_val.lower() == "niet in data"
    targeting_locs = _parse_locations(targeting_raw_val) if not skip_lokaal else []
    negative_locs  = _parse_locations(str(row.get("Negatieve locaties", "")))

    today    = _today_str()
    lok      = df_lokaal.copy()
    sta      = df_stad.copy()
    _explicit = _uses_explicit_placeholders(lok) or _uses_explicit_placeholders(sta)

    # Portaal + expliciet formaat: bouw minimale merk_info zodat [Klantnaam] vervangen wordt
    # met de klantnaam uit het UI-veld "Klantnaam campagne"
    if _explicit and merk_info is None and portaal_klantnaam:
        merk_info = {
            "korte_naam": portaal_klantnaam,
            "lange_naam": portaal_klantnaam,
            "review_score": "", "jaren_garantie": "", "usp": "", "usp_fallback": "",
        }

    for df in (lok, sta):
        # Start/End Date
        mask = df["Start Date"].notna() & (df["Start Date"].astype(str).str.strip() != "")
        df.loc[mask, "Start Date"] = today
        mask = df["End Date"].notna() & (df["End Date"].astype(str).str.strip() != "")
        df.loc[mask, "End Date"] = ""

        # Ad Schedule
        if ad_schedule:
            mask = df["Ad Schedule"].notna() & (df["Ad Schedule"].astype(str).str.strip() != "")
            df.loc[mask, "Ad Schedule"] = ad_schedule

        # Phone Number
        if _explicit:
            # Nieuw formaat: [Nummer] als placeholder
            for _pn_col in ("Phone Number", "Phone Number#Original"):
                if _pn_col in df.columns:
                    df[_pn_col] = df[_pn_col].apply(
                        lambda v, _ph=phone: str(v).replace("[Nummer]", _ph) if pd.notna(v) else v
                    )
        else:
            template_phone = TEMPLATE_PHONE
            if merk_info:
                phones = df["Phone Number"].dropna().astype(str)
                phones = phones[phones.str.strip() != ""]
                if len(phones):
                    template_phone = phones.iloc[0]
            if phone and phone != template_phone:
                df["Phone Number"] = _replace_exact(df["Phone Number"], template_phone, phone)
                if "Phone Number#Original" in df.columns:
                    df["Phone Number#Original"] = _replace_exact(df["Phone Number#Original"], template_phone, phone)

        # Final URL
        url_mask = df["Final URL"].notna() & (df["Final URL"].astype(str).str.strip() != "")
        if _explicit:
            # Nieuw formaat: [url]?kw=[Plaats]&tel=[Nummer]
            # Verwijder query-string en fragment zodat [url]?kw=... niet dubbel wordt
            _p = urlparse(url)
            _base_url = f"{_p.scheme}://{_p.netloc}{_p.path}".rstrip("/") if url else ""
            if url_mask.any():
                df.loc[url_mask, "Final URL"] = df.loc[url_mask, "Final URL"].astype(str).apply(
                    lambda v, _b=_base_url, _c=city, _ph=phone:
                        v.replace("[url]", _b).replace("[Plaats]", _c).replace("[Nummer]", _ph)
                )
            if "Final URL#Original" in df.columns:
                mask_o = df["Final URL#Original"].notna() & (df["Final URL#Original"].astype(str).str.strip() != "")
                if mask_o.any():
                    df.loc[mask_o, "Final URL#Original"] = df.loc[mask_o, "Final URL#Original"].astype(str).apply(
                        lambda v, _b=_base_url, _c=city, _ph=phone:
                            v.replace("[url]", _b).replace("[Plaats]", _c).replace("[Nummer]", _ph)
                    )
            # Price extension URLs (Final URL 1 t/m 8 en hun #Original varianten)
            for _n in range(1, 9):
                for _col in (f"Final URL {_n}", f"Final URL {_n}#Original"):
                    if _col in df.columns:
                        _m2 = df[_col].notna() & (df[_col].astype(str).str.strip() != "")
                        if _m2.any():
                            df.loc[_m2, _col] = df.loc[_m2, _col].astype(str).apply(
                                lambda v, _b=_base_url, _c=city, _ph=phone:
                                    v.replace("[url]", _b).replace("[Plaats]", _c).replace("[Nummer]", _ph)
                            )
        else:
            if url:
                if merk_info:
                    client_domain = urlparse(url).netloc or url.replace("https://", "").replace("http://", "").rstrip("/")
                    _first_urls = df.loc[url_mask, "Final URL"].astype(str)
                    _tmpl_domain = TEMPLATE_DOMAIN
                    if len(_first_urls):
                        _det = urlparse(_first_urls.iloc[0]).netloc
                        if _det:
                            _tmpl_domain = _det
                    df.loc[url_mask, "Final URL"] = df.loc[url_mask, "Final URL"].astype(str).str.replace(
                        _tmpl_domain, client_domain, regex=False
                    )
                else:
                    df.loc[url_mask, "Final URL"] = url
            if merk_info and url_mask.any():
                df.loc[url_mask, "Final URL"] = df.loc[url_mask, "Final URL"].astype(str).apply(
                    lambda v, _t=_tc, _c=city: _case_replace(v, _t, _c)
                )
                df.loc[url_mask, "Final URL"] = df.loc[url_mask, "Final URL"].astype(str).apply(
                    lambda v, _ph=phone: re.sub(r"(?<=tel=)[^&]+", _ph, v)
                )
            if "Final URL#Original" in df.columns:
                mask_o = df["Final URL#Original"].notna() & (df["Final URL#Original"].astype(str).str.strip() != "")
                if url and merk_info and mask_o.any():
                    _fo_urls = df.loc[mask_o, "Final URL#Original"].astype(str)
                    _fo_domain = TEMPLATE_DOMAIN
                    if len(_fo_urls):
                        _det = urlparse(_fo_urls.iloc[0]).netloc
                        if _det:
                            _fo_domain = _det
                    df.loc[mask_o, "Final URL#Original"] = df.loc[mask_o, "Final URL#Original"].astype(str).str.replace(
                        _fo_domain, client_domain, regex=False
                    )
                elif url and not merk_info and mask_o.any():
                    df.loc[mask_o, "Final URL#Original"] = url
                if merk_info and mask_o.any():
                    df.loc[mask_o, "Final URL#Original"] = df.loc[mask_o, "Final URL#Original"].astype(str).apply(
                        lambda v, _t=_tc, _c=city: _case_replace(v, _t, _c)
                    )
                    df.loc[mask_o, "Final URL#Original"] = df.loc[mask_o, "Final URL#Original"].astype(str).apply(
                        lambda v, _ph=phone: re.sub(r"(?<=tel=)[^&]+", _ph, v)
                    )

        # Campaign Status → Paused
        mask = df["Campaign Status"].notna() & (df["Campaign Status"].astype(str).str.strip() != "")
        df.loc[mask, "Campaign Status"] = "Paused"

        # Advertentieteksten
        for col in AD_TEXT_COLS:
            if col not in df.columns:
                continue
            if _explicit and variants:
                # Eerst variant opzoeken (stad-lengte logica), dan explicit placeholders invullen
                df[col] = df[col].apply(
                    lambda v, _col=col, _c=city, _v=variants, _m=merk_info, _p=_tp:
                        _apply_explicit_text(
                            _apply_variant(v, _c, _v, _m, col=_col),
                            _c, _m, _p, col=_col
                        ) if (pd.notna(v) and str(v).strip()) else v
                )
            elif _explicit:
                df[col] = df[col].apply(
                    lambda v, _col=col, _c=city, _m=merk_info, _p=_tp:
                        _apply_explicit_text(v, _c, _m, _p, col=_col)
                        if (pd.notna(v) and str(v).strip()) else v
                )
            elif variants:
                df[col] = df[col].apply(
                    lambda v, _col=col, _c=city, _v=variants, _m=merk_info:
                        _apply_variant(v, _c, _v, _m, col=_col)
                        if (pd.notna(v) and str(v).strip()) else v
                )
            elif merk_info:
                df[col] = df[col].apply(
                    lambda v, _col=col, _t=_tc, _c=city, _m=merk_info, _tco=_tco:
                        _apply_eigen_placeholders(
                            _case_replace(str(v), _t, _c), _m, _c, col=_col, tmpl_company=_tco
                        ) if (pd.notna(v) and str(v).strip()) else v
                )
            else:
                df[col] = _replace_in_col(df[col], _tc, city)

        # Tekenlimiet-vangnet (alleen oud formaat — nieuw formaat verwerkt dit al per cel)
        if not _explicit and merk_info:
            _korte = merk_info.get("korte_naam", "")
            _lange = merk_info.get("lange_naam", "") or _korte
            if _lange and _lange != _korte:
                for col in AD_TEXT_COLS:
                    limit = _COL_CHAR_LIMITS.get(col)
                    if not limit or col not in df.columns:
                        continue
                    df[col] = df[col].apply(
                        lambda v, _k=_korte, _l=_lange, _lim=limit:
                            (lambda s: s.replace(_l, _k) if len(s.replace(_l, _k)) <= _lim else s)(str(v))
                            if (pd.notna(v) and str(v).strip() and len(str(v)) > _lim and _l in str(v))
                            else v
                    )

        # USP-placeholder (alleen oud formaat — nieuw formaat heeft geen TEMPLATE_USP)
        if not _explicit and merk_info:
            usp = merk_info.get("usp", "").strip()
            usp_fallback = merk_info.get("usp_fallback", "")
            usp_val = usp or usp_fallback
            if usp_val:
                _usp_pattern = re.compile(re.escape(TEMPLATE_USP), re.IGNORECASE)
                _usp_lower   = TEMPLATE_USP.lower()
                for col in AD_TEXT_COLS:
                    if col in df.columns:
                        df[col] = df[col].apply(
                            lambda v, pat=_usp_pattern, rep=usp_val, low=_usp_lower:
                                pat.sub(rep, str(v))
                                if pd.notna(v) and low in str(v).lower()
                                else v
                        )

        # Prijsplaatshouder oud formaat (bijv. "€59,50")
        if not _explicit and _tp:
            for col in AD_TEXT_COLS:
                if col not in df.columns:
                    continue
                df[col] = df[col].apply(
                    lambda v, _pt=TEMPLATE_PRICE, _pn=_tp:
                        str(v).replace(_pt, _pn)
                        if pd.notna(v) and _pt in str(v)
                        else v
                )

        # Kopieer verwerkte advertentieteksten naar #Original kolommen
        for col in AD_TEXT_COLS:
            orig = col + "#Original"
            if col in df.columns and orig in df.columns:
                df[orig] = df[col]

    # Labels: zorg dat Netnummer {net} aanwezig is op alle niveaus
    lok = _ensure_netnummer_label(lok, net)
    sta = _ensure_netnummer_label(sta, net)

    # + Stad: Ad Group en Keyword
    if _explicit:
        sta["Ad Group"] = sta["Ad Group"].apply(
            lambda v, _c=city, _m=merk_info: _apply_explicit_text(v, _c, _m, None) if pd.notna(v) else v
        )
        sta["Keyword"] = sta["Keyword"].apply(
            lambda v, _c=city: str(v).replace("[Plaats]", _c) if pd.notna(v) else v
        )
    else:
        sta["Ad Group"] = _replace_in_col(sta["Ad Group"], _tc, city)
        sta["Keyword"]  = _replace_in_col(sta["Keyword"],  _tc, city)
        if merk_info:
            sta["Ad Group"] = sta["Ad Group"].apply(
                lambda v: _apply_eigen_placeholders(str(v), merk_info, city, tmpl_company=_tco)
                if pd.notna(v) else v
            )
            sta["Keyword"] = sta["Keyword"].apply(
                lambda v: _apply_eigen_placeholders(str(v), merk_info, city, tmpl_company=_tco)
                if pd.notna(v) else v
            )

    # Campaign namen (stad + lokaal)
    if _explicit:
        for _df in (sta, lok):
            _df["Campaign"] = _df["Campaign"].apply(
                lambda v, _c=city, _m=merk_info: _apply_explicit_text(v, _c, _m, None) if pd.notna(v) else v
            )
    else:
        sta["Campaign"] = _replace_in_col(sta["Campaign"], _tc, city)
        lok["Campaign"] = _replace_in_col(lok["Campaign"], _tc, city)
        if merk_info:
            _repl_camp = lambda v: _apply_eigen_placeholders(str(v), merk_info, city, tmpl_company=_tco) if pd.notna(v) else v  # noqa: E731
            sta["Campaign"] = sta["Campaign"].apply(_repl_camp)
            lok["Campaign"] = lok["Campaign"].apply(_repl_camp)

    # Lokale campagne overslaan als locatiedata ontbreekt
    if skip_lokaal:
        return df_lokaal.head(0).copy(), sta

    # Locatierijen
    targeting_raw = targeting_raw_val
    if targeting_locs or negative_locs:
        loc_filled       = lok["Location"].notna() & (lok["Location"].astype(str).str.strip() != "")
        template_loc_row = lok[loc_filled].iloc[0] if loc_filled.any() else lok.iloc[0]
        lok              = lok[~loc_filled].copy()
        new_rows         = _build_location_rows(template_loc_row, targeting_locs, negative_locs)
        if new_rows:
            lok = pd.concat([lok, pd.DataFrame(new_rows)], ignore_index=True)
    else:
        loc_filled     = lok["Location"].notna() & (lok["Location"].astype(str).str.strip() != "")
        targeting_mask = loc_filled & (
            lok["Criterion Type"].isna() | (lok["Criterion Type"].astype(str).str.strip() == "")
        )
        lok["Location groups"] = lok["Location groups"].astype(object)
        if targeting_mask.any() and targeting_raw:
            lok.loc[targeting_mask, "Location"] = targeting_raw

    return lok, sta


# ── CPC-toepassing ────────────────────────────────────────────────────────────

def _cents_to_str(cents: int) -> str:
    """Formatteer integer centen als decimaal getal met punt (Google Ads formaat)."""
    return f"{cents / 100:.2f}"


def _apply_cpc_to_df(
    df:                 pd.DataFrame,
    campaign_cpc_cents: int | None,
    ad_group_cpc_cents: int | None,
    keyword_cpc_cents:  int | None,
) -> pd.DataFrame:
    """
    Schrijf CPC-bedragen naar de juiste rijen.
    Rij-identificatie: Campaign, Ad Group, Keyword, Location kolommen.
    """
    if not any(v is not None for v in [campaign_cpc_cents, ad_group_cpc_cents, keyword_cpc_cents]):
        return df

    df = df.copy()

    def _filled(col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        return df[col].notna() & (df[col].astype(str).str.strip() != "")

    has_campaign = _filled("Campaign")
    has_ad_group = _filled("Ad Group")
    has_keyword  = _filled("Keyword")
    has_location = _filled("Location")

    # Campaign-rijen: Campaign gevuld, Ad Group leeg, geen Location
    campaign_mask = has_campaign & ~has_ad_group & ~has_location
    # Ad Group-rijen: Campaign + Ad Group gevuld, geen Keyword, geen Location
    ag_mask = has_campaign & has_ad_group & ~has_keyword & ~has_location
    # Keyword-rijen: Keyword gevuld
    kw_mask = has_keyword

    cpc_col = "Default max. CPC"
    kw_col  = "Max CPC"

    if campaign_cpc_cents is not None and cpc_col in df.columns:
        df.loc[campaign_mask, cpc_col] = _cents_to_str(campaign_cpc_cents)

    if ad_group_cpc_cents is not None and cpc_col in df.columns:
        df.loc[ag_mask, cpc_col] = _cents_to_str(ad_group_cpc_cents)

    if keyword_cpc_cents is not None and kw_col in df.columns:
        df.loc[kw_mask, kw_col] = _cents_to_str(keyword_cpc_cents)

    return df


def _apply_keyword_cpc_sheet(
    df:           pd.DataFrame,
    cpc_sheet,    # KeywordCpcSheet
    land:         str,
    merktype_cpc: str,
    account_cpc:  str,
    city:         str,
) -> pd.DataFrame:
    """Pas per-zoekwoord CPC toe vanuit het centrale keyword-adviesblad.
    Ad group Default max. CPC wordt op €0,50 gezet zodat zoekwoord-CPC altijd leidend is.
    """
    if "Keyword" not in df.columns:
        return df

    df       = df.copy()
    fallback = cpc_sheet.get_fallback_cpc(land, merktype_cpc, account_cpc)

    def _filled(col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        return df[col].notna() & (df[col].astype(str).str.strip() != "")

    has_campaign = _filled("Campaign")
    has_ad_group = _filled("Ad Group")
    has_keyword  = _filled("Keyword")
    has_location = _filled("Location")

    # Ad group-rijen: Max CPC op €0,50 zodat zoekwoord-CPC nooit overschreden wordt
    ag_mask = has_campaign & has_ad_group & ~has_keyword & ~has_location
    for _ag_col in ("Default max. CPC", "Max CPC"):
        if _ag_col in df.columns:
            df.loc[ag_mask, _ag_col] = 0.50

    # Zoekwoord-rijen: per-zoekwoord CPC opzoeken
    if "Max CPC" in df.columns:
        for idx in df.index[has_keyword]:
            kw  = str(df.at[idx, "Keyword"]).strip()
            cpc = cpc_sheet.get_cpc(land, merktype_cpc, account_cpc, city, kw)
            if cpc is None:
                cpc = fallback
            if cpc is not None:
                df.at[idx, "Max CPC"] = round(cpc / 100, 2)

    return df


def _force_campaign_status(df: pd.DataFrame, status: str) -> pd.DataFrame:
    """Forceer Campaign Status op rijen met een ingevuld Campaign-veld."""
    if "Campaign Status" not in df.columns:
        return df
    df = df.copy()
    mask = df["Campaign"].notna() & (df["Campaign"].astype(str).str.strip() != "")
    df.loc[mask, "Campaign Status"] = status
    return df


# ── hoofd-functie ─────────────────────────────────────────────────────────────

def build_all(
    lokaal_path:          str,
    stad_path:            str,
    sheet_url:            str,
    sheet_name:           str = "DakPro NL Plaatsen",
    variants_path:        str | None = None,
    klant_filter:         list[str] | None = None,
    merk_info:            dict | None = None,
    portaal_klantnaam:    str | None = None,
    ad_schedule_override: str = "",
    cpc_import:           CpcAdviceImport | None = None,
    pause_import:         PauseAdviceImport | None = None,
    dry_run:              bool = False,
    progress_cb=None,
    fallback_stad_cents:   int | None = None,
    fallback_lokaal_cents: int | None = None,
    template_city:        str | None = None,
    template_company:     str | None = None,
    template_price:       str | None = None,
    keyword_cpc_sheet=None,    # KeywordCpcSheet | None — centraal CPC-blad (Schoorsteenveger)
    land:                 str = "NL",  # "NL" of "BE" — alleen gebruikt met keyword_cpc_sheet
    merktype_cpc:         str = "",    # merktype-string zoals in CPC-blad
    is_portaal:           bool = False, # voor account-vertaling in keyword_cpc_sheet
) -> tuple[dict, list[str]]:

    # Overschrijf template-placeholders als de branche andere waarden gebruikt
    _tmpl_city    = template_city    or TEMPLATE_CITY
    _tmpl_company = template_company or TEMPLATE_COMPANY
    _tmpl_price   = template_price   or None

    df_lokaal = pd.read_csv(lokaal_path, sep=SEP, encoding=ENCODING, low_memory=False)
    df_stad   = pd.read_csv(stad_path,   sep=SEP, encoding=ENCODING, low_memory=False)
    sheet     = load_sheet(sheet_url, sheet_name)
    variants  = load_variants(variants_path) if variants_path else None

    val_errors = validate_sheet(sheet)
    if val_errors:
        raise ValueError("\n".join(val_errors))

    # Biedstrategie uit sjablonen lezen (lokaal en stad kunnen verschillen)
    bid_strategy_lok = extract_template_bid_strategy(df_lokaal)
    bid_strategy_sta = extract_template_bid_strategy(df_stad)
    bid_strategy = bid_strategy_lok or bid_strategy_sta

    engine = CampaignBuildRulesEngine(
        cpc_import             = cpc_import,
        pause_import           = pause_import,
        template_bid_strategy  = bid_strategy,
        fallback_stad_cents    = fallback_stad_cents,
        fallback_lokaal_cents  = fallback_lokaal_cents,
    )

    results:  dict      = {}
    errors:   list[str] = []
    dry_rows: list[dict] = []   # voor dry-run overzicht
    total = len(sheet)

    # Bepaal de kolom rechts van 'Dubbele plaats' voor grootte-aanduiding (Kleinste/Grootste)
    _sheet_cols = list(sheet.columns)
    _dup_idx = next((i for i, c in enumerate(_sheet_cols) if c.strip().lower() == "dubbele plaats"), -1)
    _grootte_col = _sheet_cols[_dup_idx + 1] if _dup_idx >= 0 and _dup_idx + 1 < len(_sheet_cols) else None

    # Verzamel bekende plaatsen en klanten NA filtering — warnings alleen voor actieve build
    def _klant_val(r):
        v = str(r.get("Klant", "")).strip()
        return v if v and v != "nan" else "DakPro"

    active_rows   = sheet[sheet.apply(lambda r: not klant_filter or _klant_val(r) in klant_filter, axis=1)]
    known_cities  = {str(r.get("Plaats", "")).strip() for _, r in active_rows.iterrows()}
    # Vertaal klantnamen naar CPC-accountnamen voor de warnings
    known_klanten = {_resolve_cpc_account(_klant_val(r)) for _, r in active_rows.iterrows()}

    # Unmatched CPC/pauze warnings — eenmalig vóór de loop
    for w in engine.unmatched_cpc_warnings(known_klanten):
        errors.append(f"Waarschuwing: {w}")
    for w in engine.unmatched_pause_warnings(known_cities, known_accounts=known_klanten):
        errors.append(f"Waarschuwing: {w}")

    for i, (_, row) in enumerate(sheet.iterrows()):
        city  = str(row.get("Plaats", "")).strip()
        klant = str(row.get("Klant", "")).strip()
        if not klant or klant == "nan":
            klant = "DakPro"

        if klant_filter and klant not in klant_filter:
            continue

        if str(row.get("Dubbele plaats", "")).strip().lower() == "x":
            # Alleen de kleinste van het duplo-paar overslaan; de grootste mag door
            grootte = str(row.get(_grootte_col, "")).strip().lower() if _grootte_col else ""
            if grootte != "grootste":
                if progress_cb:
                    progress_cb(city, i, total, skipped=True)
                continue

        # ── Rules engine: beslissingen per campagnetype ───────────────────────
        stad_excluded = city.lower() in EXCLUDED_STAD_CITIES
        cpc_account = _resolve_cpc_account(klant)  # klantnaam → adviesbestand-accountnaam

        stad_decision  = engine.get_city_decision(cpc_account, city)
        local_decision = engine.get_local_decision(cpc_account, city)

        # Blokkerende fouten (biedstrategie-incompatibiliteit)
        for err in stad_decision.blocking_errors + local_decision.blocking_errors:
            errors.append(f"Blokkerend: {err}")

        # Dry-run: registreer beslissingen en ga door zonder te bouwen
        if dry_run:
            stad_skip_reason = (
                "uitgesloten (algemeen woord)" if stad_excluded
                else ("niet bouwen" if not stad_decision.should_build else "")
            )
            dry_rows.append({
                "Klant":        klant,
                "Plaats":       city,
                "Stad":         "gepauzeerd" if (stad_decision.should_build and stad_decision.final_status == "Paused" and not stad_excluded)
                                else ("overgeslagen" if (not stad_decision.should_build or stad_excluded) else "normaal"),
                "Stad reden":   stad_skip_reason or next((getattr(r, "status_reason", "") for r in stad_decision.matched_rules if getattr(r, "status_reason", "")), ""),
                "Lokaal":       "gepauzeerd" if (local_decision.should_build and local_decision.final_status == "Paused")
                                else ("overgeslagen" if not local_decision.should_build else "normaal"),
                "Lokaal reden": next((getattr(r, "status_reason", "") for r in local_decision.matched_rules if getattr(r, "status_reason", "")), ""),
                "CPC stad":     f"€{stad_decision.campaign_cpc_cents/100:.2f}" if stad_decision.campaign_cpc_cents else "—",
                "CPC lokaal":   f"€{local_decision.campaign_cpc_cents/100:.2f}" if local_decision.campaign_cpc_cents else "—",
                "Waarschuwingen": "; ".join(stad_decision.warnings + local_decision.warnings),
            })
            if progress_cb:
                progress_cb(city, i, total, skipped=False, stad_excluded=stad_excluded)
            continue

        if progress_cb:
            progress_cb(city, i, total, skipped=False, stad_excluded=stad_excluded)

        # ── Campagnes bouwen ──────────────────────────────────────────────────
        try:
            lok_df, sta_df = process_city(
                df_lokaal, df_stad, row, variants, merk_info, ad_schedule_override,
                tmpl_city=_tmpl_city, tmpl_company=_tmpl_company, tmpl_price=_tmpl_price,
                portaal_klantnaam=portaal_klantnaam or None,
            )
        except Exception as exc:
            import traceback as _tb
            errors.append(f"[v2] {city}: {exc}\n{''.join(_tb.format_tb(exc.__traceback__))}")
            continue

        if klant not in results:
            results[klant] = {"lokaal": [], "stad": []}

        # Lokale campagne
        if local_decision.should_build and len(lok_df) > 0:
            if local_decision.final_status == "Paused":
                lok_df = _force_campaign_status(lok_df, "Paused")
            if keyword_cpc_sheet is not None:
                _sv_account = _resolve_sv_cpc_account(klant, is_portaal)
                lok_df = _apply_keyword_cpc_sheet(lok_df, keyword_cpc_sheet, land, merktype_cpc, _sv_account, city)
            else:
                lok_df = _apply_cpc_to_df(
                    lok_df,
                    local_decision.campaign_cpc_cents,
                    local_decision.ad_group_cpc_cents,
                    local_decision.keyword_cpc_cents,
                )
            results[klant]["lokaal"].append(lok_df)
        elif not local_decision.should_build:
            errors.append(f"Info: {city} (lokaal) niet gebouwd — {local_decision.audit_trail[-1].applied_value if local_decision.audit_trail else 'pauzeringsregel'}.")

        # +Stad campagne
        if not stad_excluded and stad_decision.should_build:
            if stad_decision.final_status == "Paused":
                sta_df = _force_campaign_status(sta_df, "Paused")
            if keyword_cpc_sheet is not None:
                _sv_account = _resolve_sv_cpc_account(klant, is_portaal)
                sta_df = _apply_keyword_cpc_sheet(sta_df, keyword_cpc_sheet, land, merktype_cpc, _sv_account, city)
            else:
                sta_df = _apply_cpc_to_df(
                    sta_df,
                    stad_decision.campaign_cpc_cents,
                    stad_decision.ad_group_cpc_cents,
                    stad_decision.keyword_cpc_cents,
                )
            results[klant]["stad"].append(sta_df)
        elif not stad_excluded and not stad_decision.should_build:
            errors.append(f"Info: {city} (+Stad) niet gebouwd — {stad_decision.audit_trail[-1].applied_value if stad_decision.audit_trail else 'pauzeringsregel'}.")

    # Dry-run: geef overzicht terug i.p.v. campagnedata
    if dry_run:
        return {"__dry_run__": dry_rows}, errors

    def _dedup_campaign_rows(df: pd.DataFrame) -> pd.DataFrame:
        """Verwijder dubbele campagne-definitierijen (Campaign#Original ingevuld).
        Bij meerdere steden in één campagne bevat elke stad een identieke campagne-rij.
        Google Ads Editor interpreteert elk als aparte campagne — bewaar er alleen de eerste per naam."""
        if "Campaign#Original" not in df.columns:
            return df
        is_camp = df["Campaign#Original"].notna() & (df["Campaign#Original"].astype(str).str.strip() != "")
        camp_rows = df[is_camp]
        other_rows = df[~is_camp]
        camp_deduped = camp_rows.drop_duplicates(subset=["Campaign#Original"], keep="first")
        return pd.concat([camp_deduped, other_rows], ignore_index=True)

    merged = {}
    for klant, parts in results.items():
        lokaal_frames = parts["lokaal"]
        stad_frames   = parts["stad"]
        merged[klant] = {
            "lokaal": _dedup_campaign_rows(pd.concat(lokaal_frames, ignore_index=True)) if lokaal_frames
                      else pd.read_csv(lokaal_path, sep=SEP, encoding=ENCODING, low_memory=False).head(0),
            "stad":   _dedup_campaign_rows(pd.concat(stad_frames, ignore_index=True)) if stad_frames
                      else pd.read_csv(stad_path, sep=SEP, encoding=ENCODING, low_memory=False).head(0),
        }

    return merged, errors


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_csv(buf, sep=SEP, encoding=ENCODING, index=False)
    return buf.getvalue()


def output_filename(klant: str, kind: str) -> str:
    return f"{klant}_{kind}_{_today_str()}.csv"
