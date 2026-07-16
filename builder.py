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

# ── constanten sjabloon ──────────────────────────────────────────────────────

TEMPLATE_CITY           = "Groningen"
TEMPLATE_PHONE          = "050-7820442"
TEMPLATE_NET            = "050"
TEMPLATE_COMPANY        = "DCN"                  # bedrijfsnaam placeholder in eigen merk CSV/Excel
TEMPLATE_DOMAIN         = "dcn-dakdekkers.nl"    # domein in eigen merk sjabloon-CSV
TEMPLATE_USP            = "Familiebedrijf Sinds 1966"  # USP placeholder in eigen merk CSV
TEMPLATE_REVIEW_SCORE   = "9,6"                  # review score in portaal Excel-varianten
TEMPLATE_JAREN_GARANTIE = "10"                   # jaren garantie in portaal Excel-varianten

# Genormaliseerde placeholders voor variant-lookup (lowercase)
_KEY_REVIEW = "[reviewscore]"
_KEY_JAREN  = "[jarengarantie]"

# Plaatsen die Google als algemeen woord herkent → nooit +Stad campagne genereren
EXCLUDED_STAD_CITIES: frozenset[str] = frozenset({
    "huizen", "zetten", "best", "dieren", "echt",
    "hoornaar", "heel", "enter", "zeeland", "noorden",
})

SEP      = ";"
ENCODING = "utf-8-sig"

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


def _replace_exact(series: pd.Series, old: str, new: str) -> pd.Series:
    def _safe(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        return str(val).replace(old, new)
    return series.apply(_safe)


def _ensure_netnummer_label(df: pd.DataFrame, net: str) -> pd.DataFrame:
    """Voeg 'Netnummer {net}' toe aan Labels voor alle campagne/ad group/keyword/ad/sitelink rijen."""
    label = f"Netnummer {net}"

    def _add(val):
        if pd.isna(val) or str(val).strip() == "":
            return label
        s = str(val).strip()
        if "Netnummer" in s:
            return re.sub(r"Netnummer \S+", label, s)
        return f"{s};{label}"

    has_campaign = df["Campaign"].notna() & (df["Campaign"].astype(str).str.strip() != "")
    has_location = df["Location"].notna() & (df["Location"].astype(str).str.strip() != "")
    mask = has_campaign & ~has_location

    df = df.copy()
    df.loc[mask, "Labels"] = df.loc[mask, "Labels"].apply(_add)
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

        # normaliseer de lookup-sleutel
        key = korte
        key = re.sub(r"'?Plaats'?", TEMPLATE_CITY, key)       # plaatsnaam placeholder
        key = key.replace(TEMPLATE_COMPANY, "[Bedrijfsnaam]")  # bedrijfsnaam placeholder
        key = key.lower()
        # normaliseer portaal-specifieke waarden naar generieke placeholders
        key = key.replace(TEMPLATE_REVIEW_SCORE, _KEY_REVIEW)
        key = re.sub(r"\b" + re.escape(TEMPLATE_JAREN_GARANTIE) + r"\b jaar", _KEY_JAREN + " jaar", key)

        lange_val = lange if lange and lange != "nan" else None

        variants[key] = {"lange_raw": lange_val, "rule_info": rule_info}

    return variants


def _resolve_name(lange_raw: str, korte_naam: str, lange_naam: str, name_max_len: int | None) -> str:
    """Vervang naam-placeholders in de lange-variant tekst (enkele quotes én vierkante haken)."""
    effective_lange = korte_naam if (name_max_len and len(lange_naam) > name_max_len) else lange_naam
    result = lange_raw
    result = result.replace("'Lange bedrijfsnaam'", effective_lange)
    result = result.replace("[Lange bedrijfsnaam]", effective_lange)
    result = result.replace("'Korte bedrijfsnaam'", korte_naam)
    result = result.replace("[Korte bedrijfsnaam]", korte_naam)
    return result


def _apply_variant(val: str, city: str, variants: dict, merk_info: dict | None = None) -> str:
    """
    Bepaal de juiste tekstvariant voor één cel.

    Portaal mode (merk_info is None): alleen stadslengte-drempel.
    Eigen merk mode: ook bedrijfsnaam, USP, altijd-vervang regels.
    """
    if not val or pd.isna(val):
        return val

    s = str(val)
    # normaliseer lookup-sleutel: eigen merk placeholders → zelfde vorm als Excel-sleutels
    key = s.lower().replace("[reviewscore]", _KEY_REVIEW).replace("[jarengarantie]", _KEY_JAREN)

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

        if r == "always" and merk_info:
            if not lange_raw:
                return s
            korte = merk_info.get("korte_naam", "")
            lange = merk_info.get("lange_naam", "")
            result = _resolve_name(lange_raw, korte, lange, rule.get("name_max_len"))
            result = _apply_portaal_values(result, merk_info)
            return _case_replace(result, TEMPLATE_CITY, city)

        if r == "city":
            threshold    = rule.get("threshold", 999)
            name_max_len = rule.get("name_max_len")
            if len(city) >= threshold:
                if not lange_raw:
                    return ""
                if merk_info:
                    korte = merk_info.get("korte_naam", "")
                    lange = merk_info.get("lange_naam", "")
                    result = _resolve_name(lange_raw, korte, lange, name_max_len)
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


def _apply_eigen_placeholders(text: str, merk_info: dict, city: str) -> str:
    """Vervang [Bedrijfsnaam], [ReviewScore], [JarenGarantie] in één string."""
    korte = merk_info.get("korte_naam", "")
    lange = merk_info.get("lange_naam", "")
    bedrijfsnaam = korte if (lange and len(lange) > 20) else (lange or korte)

    text = text.replace("[Bedrijfsnaam]", bedrijfsnaam)
    text = text.replace("[ReviewScore]",  merk_info.get("review_score",    ""))
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
) -> tuple[pd.DataFrame, pd.DataFrame]:

    city  = str(row["Plaats"]).strip()

    # Ad Schedule: tool-veld heeft voorrang, anders uit sheet
    ad_schedule = ad_schedule_override.strip() or str(row.get("Ad Schedule", "")).strip()

    # Netnummer (herstel leidende nul via telefoonnummer)
    net_raw   = str(row.get("Netnummer", "")).strip()
    phone_raw = str(row.get("Telefoonnummer", "")).strip()
    if net_raw in ("", "nan"):
        net = ""
    else:
        try:
            net = str(int(float(net_raw)))
        except ValueError:
            net = net_raw
        if "-" in phone_raw:
            tel_prefix = phone_raw.split("-")[0]
            if tel_prefix.lstrip("0") == net:
                net = tel_prefix

    phone = phone_raw
    url   = str(row.get("URL", "")).strip()

    if not net:   net   = TEMPLATE_NET
    if not phone: phone = TEMPLATE_PHONE

    targeting_raw_val = str(row.get("Locatie", "")).strip()
    skip_lokaal    = targeting_raw_val.lower() == "niet in data"
    targeting_locs = _parse_locations(targeting_raw_val) if not skip_lokaal else []
    negative_locs  = _parse_locations(str(row.get("Negatieve locaties", "")))

    today = _today_str()
    lok   = df_lokaal.copy()
    sta   = df_stad.copy()

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
        template_phone = TEMPLATE_PHONE
        if merk_info:
            # detecteer template-telefoonnummer uit eerste gevulde cel
            phones = df["Phone Number"].dropna().astype(str)
            phones = phones[phones.str.strip() != ""]
            if len(phones):
                template_phone = phones.iloc[0]
        if phone and phone != template_phone:
            df["Phone Number"] = _replace_exact(df["Phone Number"], template_phone, phone)

        # Final URL
        if url:
            if merk_info:
                # eigen merk: vervang alleen het domein, behoud pad en query
                client_domain = urlparse(url).netloc or url.replace("https://", "").replace("http://", "").rstrip("/")
                mask = df["Final URL"].notna() & (df["Final URL"].astype(str).str.strip() != "")
                df.loc[mask, "Final URL"] = df.loc[mask, "Final URL"].astype(str).str.replace(
                    TEMPLATE_DOMAIN, client_domain, regex=False
                )
                # vervang ook kw=Groningen in URL
                df.loc[mask, "Final URL"] = df.loc[mask, "Final URL"].astype(str).apply(
                    lambda v: _case_replace(v, TEMPLATE_CITY, city)
                )
                # vervang tel= parameter in URL (ongeacht welk nummer in sjabloon staat)
                _ph = phone  # capture voor lambda
                df.loc[mask, "Final URL"] = df.loc[mask, "Final URL"].astype(str).apply(
                    lambda v: re.sub(r"(?<=tel=)[^&]+", _ph, v)
                )
            else:
                mask = df["Final URL"].notna() & (df["Final URL"].astype(str).str.strip() != "")
                df.loc[mask, "Final URL"] = url

        # Campaign Status → Paused
        mask = df["Campaign Status"].notna() & (df["Campaign Status"].astype(str).str.strip() != "")
        df.loc[mask, "Campaign Status"] = "Paused"

        # Advertentieteksten
        for col in AD_TEXT_COLS:
            if col not in df.columns:
                continue
            if variants:
                df[col] = df[col].apply(
                    lambda v: _apply_variant(v, city, variants, merk_info)
                    if (pd.notna(v) and str(v).strip()) else v
                )
            elif merk_info:
                df[col] = df[col].apply(
                    lambda v: _apply_eigen_placeholders(
                        _case_replace(str(v), TEMPLATE_CITY, city), merk_info, city
                    ) if (pd.notna(v) and str(v).strip()) else v
                )
            else:
                df[col] = _replace_in_col(df[col], TEMPLATE_CITY, city)

        # vervang USP-placeholder (case-insensitief, ook in lange varianten)
        if merk_info:
            usp = merk_info.get("usp", "").strip()
            usp_fallback = merk_info.get("usp_fallback", "")
            usp_val = usp or usp_fallback
            if usp_val:
                _usp_pattern = re.compile(re.escape(TEMPLATE_USP), re.IGNORECASE)
                for col in AD_TEXT_COLS:
                    if col in df.columns:
                        df[col] = df[col].apply(
                            lambda v, pat=_usp_pattern, rep=usp_val:
                                pat.sub(rep, str(v))
                                if pd.notna(v) and TEMPLATE_USP.lower() in str(v).lower()
                                else v
                        )

    # Labels: zorg dat Netnummer {net} aanwezig is op alle niveaus
    lok = _ensure_netnummer_label(lok, net)
    sta = _ensure_netnummer_label(sta, net)

    # + Stad specifiek
    sta["Ad Group"] = _replace_in_col(sta["Ad Group"], TEMPLATE_CITY, city)
    sta["Keyword"]  = _replace_in_col(sta["Keyword"],  TEMPLATE_CITY, city)
    if merk_info:
        sta["Ad Group"] = sta["Ad Group"].apply(
            lambda v: _apply_eigen_placeholders(str(v), merk_info, city)
            if pd.notna(v) else v
        )
        sta["Keyword"] = sta["Keyword"].apply(
            lambda v: _apply_eigen_placeholders(str(v), merk_info, city)
            if pd.notna(v) else v
        )

    # Lokaal specifiek
    lok["Campaign"] = _replace_in_col(lok["Campaign"], TEMPLATE_CITY, city)

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


# ── hoofd-functie ─────────────────────────────────────────────────────────────

def build_all(
    lokaal_path:          str,
    stad_path:            str,
    sheet_url:            str,
    sheet_name:           str = "DakPro NL Plaatsen",
    variants_path:        str | None = None,
    klant_filter:         list[str] | None = None,
    merk_info:            dict | None = None,
    ad_schedule_override: str = "",
    progress_cb=None,
) -> tuple[dict, list[str]]:

    df_lokaal = pd.read_csv(lokaal_path, sep=SEP, encoding=ENCODING, low_memory=False)
    df_stad   = pd.read_csv(stad_path,   sep=SEP, encoding=ENCODING, low_memory=False)
    sheet     = load_sheet(sheet_url, sheet_name)
    variants  = load_variants(variants_path) if variants_path else None

    val_errors = validate_sheet(sheet)
    if val_errors:
        raise ValueError("\n".join(val_errors))

    results: dict = {}
    errors:  list[str] = []
    total = len(sheet)

    for i, (_, row) in enumerate(sheet.iterrows()):
        city  = str(row.get("Plaats", "")).strip()
        klant = str(row.get("Klant", "")).strip()
        if not klant or klant == "nan":
            klant = "DakPro"

        if klant_filter and klant not in klant_filter:
            continue

        if str(row.get("Dubbele plaats", "")).strip().lower() == "x":
            if progress_cb:
                progress_cb(city, i, total, skipped=True)
            continue

        if progress_cb:
            progress_cb(city, i, total, skipped=False,
                        stad_excluded=city.lower() in EXCLUDED_STAD_CITIES)

        try:
            lok_df, sta_df = process_city(
                df_lokaal, df_stad, row, variants, merk_info, ad_schedule_override
            )
        except Exception as exc:
            errors.append(f"{city}: {exc}")
            continue

        if klant not in results:
            results[klant] = {"lokaal": [], "stad": []}
        if len(lok_df) > 0:
            results[klant]["lokaal"].append(lok_df)
        if city.lower() not in EXCLUDED_STAD_CITIES:
            results[klant]["stad"].append(sta_df)

    merged = {}
    for klant, parts in results.items():
        lokaal_frames = parts["lokaal"]
        stad_frames   = parts["stad"]
        merged[klant] = {
            "lokaal": pd.concat(lokaal_frames, ignore_index=True) if lokaal_frames
                      else pd.read_csv(lokaal_path, sep=SEP, encoding=ENCODING, low_memory=False).head(0),
            "stad":   pd.concat(stad_frames, ignore_index=True) if stad_frames
                      else pd.read_csv(stad_path, sep=SEP, encoding=ENCODING, low_memory=False).head(0),
        }

    return merged, errors


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_csv(buf, sep=SEP, encoding=ENCODING, index=False)
    return buf.getvalue()


def output_filename(klant: str, kind: str) -> str:
    return f"{klant}_{kind}_{_today_str()}.csv"
