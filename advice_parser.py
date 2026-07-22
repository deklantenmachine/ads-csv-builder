"""
Parses and validates CPC advice and pause/termination advice files.
File type detection is based solely on sheet names and column presence — never filename.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum

import pandas as pd


# ── Enums ─────────────────────────────────────────────────────────────────────

class AdviceFileType(str, Enum):
    CPC_ADVICE   = "CPC_ADVICE"
    PAUSE_ADVICE = "PAUSE_ADVICE"
    UNKNOWN      = "UNKNOWN"
    AMBIGUOUS    = "AMBIGUOUS"


class CampaignType(str, Enum):
    REGULAR_CITY = "REGULAR_CITY"
    LOCAL        = "LOCAL"


class BuildAction(str, Enum):
    PAUSE        = "PAUSE"
    DO_NOT_BUILD = "DO_NOT_BUILD"


# ── Campagnetype mapping ───────────────────────────────────────────────────────

_CAMPAIGN_TYPE_MAP: dict[str, CampaignType] = {
    "regulier/stad": CampaignType.REGULAR_CITY,
    "regulier":      CampaignType.REGULAR_CITY,
    "stad":          CampaignType.REGULAR_CITY,
    "+stad":         CampaignType.REGULAR_CITY,
    "+ stad":        CampaignType.REGULAR_CITY,
    "lokaal":        CampaignType.LOCAL,
    "local":         CampaignType.LOCAL,
}

_PAUSE_ACTION_MAP: dict[str, BuildAction] = {
    "plaats gepauzeerd houden":                   BuildAction.PAUSE,
    "gepauzeerd bouwen":                          BuildAction.PAUSE,
    "lokale campagne gepauzeerd bouwen":          BuildAction.PAUSE,
    "niet bouwen":                                BuildAction.DO_NOT_BUILD,
    "niet bouwen – historisch beëindigd":         BuildAction.DO_NOT_BUILD,
    "niet bouwen - historisch beëindigd":         BuildAction.DO_NOT_BUILD,
    "niet bouwen / beëindigd houden":             BuildAction.DO_NOT_BUILD,
    "niet bouwen/beëindigd houden":               BuildAction.DO_NOT_BUILD,
    "beëindigd houden":                           BuildAction.DO_NOT_BUILD,
    "niet opnieuw bouwen":                        BuildAction.DO_NOT_BUILD,
}

# Sheet/column definitions
_CPC_REQUIRED_SHEET  = "Bouwplan CPC"
_CPC_REQUIRED_COLS   = {"Account", "Klant-ID", "Campagnetype", "Standaard CPC instellen"}
_CPC_OPTIONAL_SHEETS = {"Campagne-uitzonderingen", "Adv.groep-uitzonderingen", "Zoekwoord-uitzonderingen"}

_PAUSE_POSSIBLE_SHEETS = {"+ stad - plaatsen", "Lokale campagnes"}
_PAUSE_REQUIRED_COLS   = {"Account", "Klant-ID", "Campagne / plaats", "Definitieve bouwactie"}

# Bid strategies that support manual CPC
_MANUAL_CPC_STRATEGIES = {"manual_cpc", "manual cpc", "manuele cpc", "handmatige cpc", ""}


# ── Normalisatie-helpers ───────────────────────────────────────────────────────

def normalize_customer_id(raw: str) -> str:
    """Strip non-digits: '123-456-7890' → '1234567890'."""
    return re.sub(r"\D", "", str(raw).strip())


def normalize_name(raw: str) -> str:
    """Lowercase, strip, collapse whitespace."""
    return re.sub(r"\s+", " ", str(raw).strip().lower())


def parse_cpc_cents(raw) -> tuple[int | None, str | None]:
    """Parse CPC value to integer cents. Returns (cents, error_msg)."""
    if pd.isna(raw) or str(raw).strip() in ("", "nan"):
        return None, "Leeg bedrag"
    s = str(raw).strip().replace("€", "").replace(",", ".").strip()
    try:
        val = float(s)
        if val <= 0:
            return None, f"CPC moet positief zijn: {raw!r}"
        return round(val * 100), None
    except ValueError:
        return None, f"Ongeldig CPC-bedrag: {raw!r}"


def parse_campaign_type(raw: str) -> tuple[CampaignType | None, str | None]:
    key = normalize_name(raw)
    if key in _CAMPAIGN_TYPE_MAP:
        return _CAMPAIGN_TYPE_MAP[key], None
    # strip spaces and slashes for loose matching
    key_compact = key.replace(" ", "").replace("/", "/")
    for k, v in _CAMPAIGN_TYPE_MAP.items():
        if k.replace(" ", "") == key_compact:
            return v, None
    return None, f"Onbekend campagnetype: {raw!r} (ondersteund: REGULIER/STAD, LOKAAL)"


def parse_build_action(raw: str) -> tuple[BuildAction | None, str | None]:
    key = normalize_name(raw)
    if key in _PAUSE_ACTION_MAP:
        return _PAUSE_ACTION_MAP[key], None
    # partial match
    for k, v in _PAUSE_ACTION_MAP.items():
        if k in key:
            return v, None
    return None, f"Onbekende bouwactie: {raw!r}"


def normalize_match_type(raw: str) -> str:
    mapping = {
        "exact":  "EXACT",
        "phrase": "PHRASE",
        "broad":  "BROAD",
        "brede":  "BROAD",
        "breed":  "BROAD",
        "zin":    "PHRASE",
    }
    return mapping.get(raw.strip().lower(), raw.strip().upper())


def file_checksum(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


# ── Datamodellen ───────────────────────────────────────────────────────────────

@dataclass
class AdviceFileMetadata:
    original_file_name:   str
    detected_file_type:   AdviceFileType
    imported_at:          datetime
    checksum:             str
    workbook_sheet_names: list[str]
    row_counts:           dict[str, int]


@dataclass
class DefaultCpcRule:
    account_name:        str
    normalized_account:  str
    customer_id_raw:     str
    campaign_type:       CampaignType
    cpc_in_cents:        int
    source_sheet:        str
    source_row:          int


@dataclass
class CampaignCpcException:
    account_name:             str
    normalized_account:       str
    customer_id_raw:          str
    campaign_type:            CampaignType
    campaign_name:            str
    normalized_campaign_name: str
    place_name:               str
    normalized_place_name:    str
    cpc_in_cents:             int
    source_sheet:             str
    source_row:               int


@dataclass
class AdGroupCpcException:
    account_name:             str
    normalized_account:       str
    customer_id_raw:          str
    campaign_type:            CampaignType
    campaign_name:            str
    ad_group_name:            str
    normalized_ad_group_name: str
    cpc_in_cents:             int
    source_sheet:             str
    source_row:               int


@dataclass
class KeywordCpcException:
    account_name:             str
    normalized_account:       str
    customer_id_raw:          str
    campaign_type:            CampaignType
    campaign_name:            str
    ad_group_name:            str
    normalized_ad_group_name: str
    keyword_text:             str
    normalized_keyword_text:  str
    match_type:               str   # EXACT | PHRASE | BROAD
    cpc_in_cents:             int
    source_sheet:             str
    source_row:               int


@dataclass
class CpcAdviceImport:
    metadata:            AdviceFileMetadata
    default_rules:       list[DefaultCpcRule]       = field(default_factory=list)
    campaign_exceptions: list[CampaignCpcException] = field(default_factory=list)
    ad_group_exceptions: list[AdGroupCpcException]  = field(default_factory=list)
    keyword_exceptions:  list[KeywordCpcException]  = field(default_factory=list)
    validation_errors:   list[str]                  = field(default_factory=list)
    validation_warnings: list[str]                  = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return len(self.validation_errors) == 0

    @property
    def summary(self) -> dict:
        accounts = {r.normalized_account for r in self.default_rules}
        return {
            "accounts":            len(accounts),
            "standaard_regels":    len(self.default_rules),
            "campagne_uitzond":    len(self.campaign_exceptions),
            "adgroup_uitzond":     len(self.ad_group_exceptions),
            "zoekwoord_uitzond":   len(self.keyword_exceptions),
            "fouten":              len(self.validation_errors),
            "waarschuwingen":      len(self.validation_warnings),
        }


@dataclass
class CityPlaceRule:
    account_name:         str
    normalized_account:   str
    customer_id_raw:      str
    place_name:           str
    normalized_place_name: str
    action:               BuildAction
    status_reason:        str
    source_sheet:         str
    source_row:           int


@dataclass
class LocalCampaignRule:
    account_name:             str
    normalized_account:       str
    customer_id_raw:          str
    campaign_name:            str
    normalized_campaign_name: str
    place_name:               str
    normalized_place_name:    str
    action:                   BuildAction
    status_reason:            str
    source_sheet:             str
    source_row:               int


@dataclass
class PauseAdviceImport:
    metadata:             AdviceFileMetadata
    city_place_rules:     list[CityPlaceRule]    = field(default_factory=list)
    local_campaign_rules: list[LocalCampaignRule] = field(default_factory=list)
    validation_errors:    list[str]               = field(default_factory=list)
    validation_warnings:  list[str]               = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return len(self.validation_errors) == 0

    @property
    def summary(self) -> dict:
        accounts = {r.normalized_account for r in self.city_place_rules + self.local_campaign_rules}
        stad_pause    = sum(1 for r in self.city_place_rules if r.action == BuildAction.PAUSE)
        stad_skip     = sum(1 for r in self.city_place_rules if r.action == BuildAction.DO_NOT_BUILD)
        lok_pause     = sum(1 for r in self.local_campaign_rules if r.action == BuildAction.PAUSE)
        lok_skip      = sum(1 for r in self.local_campaign_rules if r.action == BuildAction.DO_NOT_BUILD)
        return {
            "accounts":          len(accounts),
            "stad_gepauzeerd":   stad_pause,
            "stad_niet_bouwen":  stad_skip,
            "lokaal_gepauzeerd": lok_pause,
            "lokaal_niet_bouwen": lok_skip,
            "fouten":            len(self.validation_errors),
            "waarschuwingen":    len(self.validation_warnings),
        }


# ── Bestandstype-detectie ─────────────────────────────────────────────────────

def _load_workbook(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    import io as _io
    return pd.read_excel(_io.BytesIO(file_bytes), sheet_name=None, header=0)


def _has_cpc_structure(wb_data: dict[str, pd.DataFrame]) -> bool:
    if _CPC_REQUIRED_SHEET not in wb_data:
        return False
    df = wb_data[_CPC_REQUIRED_SHEET]
    if df is None or df.empty:
        return False
    cols = {str(c).strip() for c in df.columns}
    return _CPC_REQUIRED_COLS.issubset(cols)


def _has_pause_structure(wb_data: dict[str, pd.DataFrame]) -> bool:
    for sheet_name in _PAUSE_POSSIBLE_SHEETS:
        df = wb_data.get(sheet_name)
        if df is not None and not df.empty:
            cols = {str(c).strip() for c in df.columns}
            if _PAUSE_REQUIRED_COLS.issubset(cols):
                return True
    return False


def detect_advice_file_type(
    file_bytes: bytes,
    file_name:  str = "onbekend",
) -> tuple[AdviceFileType, dict[str, pd.DataFrame], AdviceFileMetadata]:
    """
    Detect file type based on sheet names and column structure only.
    Returns (file_type, sheet_data_dict, metadata).
    """
    wb_data = _load_workbook(file_bytes)

    has_cpc   = _has_cpc_structure(wb_data)
    has_pause = _has_pause_structure(wb_data)

    if has_cpc and has_pause:
        file_type = AdviceFileType.AMBIGUOUS
    elif has_cpc:
        file_type = AdviceFileType.CPC_ADVICE
    elif has_pause:
        file_type = AdviceFileType.PAUSE_ADVICE
    else:
        file_type = AdviceFileType.UNKNOWN

    metadata = AdviceFileMetadata(
        original_file_name   = file_name,
        detected_file_type   = file_type,
        imported_at          = datetime.now(),
        checksum             = file_checksum(file_bytes),
        workbook_sheet_names = list(wb_data.keys()),
        row_counts           = {name: len(df) for name, df in wb_data.items()},
    )
    return file_type, wb_data, metadata


# ── CPC-bestand parseren ──────────────────────────────────────────────────────

def _get_col(row: pd.Series, *names: str):
    """Return first matching column value from a row."""
    for name in names:
        if name in row.index:
            val = row[name]
            if pd.notna(val) and str(val).strip() not in ("", "nan"):
                return val
    return None


def parse_cpc_advice(file_bytes: bytes, file_name: str = "onbekend") -> CpcAdviceImport:
    file_type, wb_data, metadata = detect_advice_file_type(file_bytes, file_name)

    errors:   list[str] = []
    warnings: list[str] = []

    if file_type == AdviceFileType.UNKNOWN:
        errors.append(
            f"Bestand '{file_name}' is niet herkend als CPC-adviesbestand. "
            f"Verwacht tabblad '{_CPC_REQUIRED_SHEET}' met kolommen: "
            f"{', '.join(sorted(_CPC_REQUIRED_COLS))}. "
            f"Aanwezige tabbladen: {', '.join(sorted(wb_data.keys()))}."
        )
        return CpcAdviceImport(metadata=metadata, validation_errors=errors)

    if file_type == AdviceFileType.PAUSE_ADVICE:
        errors.append(
            f"Bestand '{file_name}' is herkend als pauzeringsbestand, niet als CPC-adviesbestand."
        )
        return CpcAdviceImport(metadata=metadata, validation_errors=errors)

    if file_type == AdviceFileType.AMBIGUOUS:
        warnings.append(
            f"Bestand '{file_name}' voldoet aan zowel de CPC- als de pauzeringsstructuur. "
            f"Verwerkt als CPC-adviesbestand."
        )

    # ── Bouwplan CPC ──────────────────────────────────────────────────────────
    df_main = wb_data[_CPC_REQUIRED_SHEET].copy()
    df_main.columns = [str(c).strip() for c in df_main.columns]
    df_main = df_main.dropna(how="all")

    default_rules:       list[DefaultCpcRule]       = []
    campaign_exceptions: list[CampaignCpcException] = []
    ad_group_exceptions: list[AdGroupCpcException]  = []
    keyword_exceptions:  list[KeywordCpcException]  = []

    seen_defaults: dict[tuple, int] = {}  # (norm_account, campaign_type) → src_row

    for row_idx, row in df_main.iterrows():
        src_row     = int(row_idx) + 2
        account_raw = str(row.get("Account", "")).strip()
        cid_raw     = str(row.get("Klant-ID", "")).strip()
        ctype_raw   = str(row.get("Campagnetype", "")).strip()
        cpc_raw     = row.get("Standaard CPC instellen")

        if not account_raw or account_raw == "nan":
            continue  # lege rij overslaan
        norm_account = normalize_name(account_raw)

        ctype, ctype_err = parse_campaign_type(ctype_raw)
        if ctype_err:
            errors.append(f"Bouwplan CPC rij {src_row} ({account_raw}): {ctype_err}")
            continue

        cpc_cents, cpc_err = parse_cpc_cents(cpc_raw)
        if cpc_err:
            errors.append(f"Bouwplan CPC rij {src_row} ({account_raw}): {cpc_err}")
            continue

        dup_key = (norm_account, ctype)
        if dup_key in seen_defaults:
            errors.append(
                f"Bouwplan CPC rij {src_row}: Dubbele standaardregel voor "
                f"'{account_raw}' + {ctype.value} (eerste rij: {seen_defaults[dup_key]})."
            )
            continue
        seen_defaults[dup_key] = src_row

        default_rules.append(DefaultCpcRule(
            account_name       = account_raw,
            normalized_account = norm_account,
            customer_id_raw    = cid_raw,
            campaign_type      = ctype,
            cpc_in_cents       = cpc_cents,
            source_sheet       = _CPC_REQUIRED_SHEET,
            source_row         = src_row,
        ))

    # Stad/lokaal CPC-ratio validatie (minimaal €0,20 verschil)
    by_account: dict[str, dict[CampaignType, DefaultCpcRule]] = {}
    for rule in default_rules:
        by_account.setdefault(rule.normalized_account, {})[rule.campaign_type] = rule

    for norm_acc, type_map in by_account.items():
        stad  = type_map.get(CampaignType.REGULAR_CITY)
        lokaal = type_map.get(CampaignType.LOCAL)
        if stad and lokaal:
            if stad.cpc_in_cents < lokaal.cpc_in_cents + 20:
                errors.append(
                    f"Account '{stad.account_name}' (Klant-ID: {stad.customer_id_raw}): "
                    f"REGULIER/STAD CPC (€{stad.cpc_in_cents/100:.2f}) moet minimaal "
                    f"€0,20 hoger zijn dan LOKAAL CPC (€{lokaal.cpc_in_cents/100:.2f}). "
                    f"Bedragen worden niet automatisch aangepast."
                )
        elif stad and not lokaal:
            warnings.append(
                f"Account '{stad.account_name}': alleen REGULIER/STAD CPC-regel aanwezig, "
                f"stad/lokaal-validatie niet uitgevoerd."
            )
        elif lokaal and not stad:
            warnings.append(
                f"Account '{lokaal.account_name}': alleen LOKAAL CPC-regel aanwezig, "
                f"stad/lokaal-validatie niet uitgevoerd."
            )

    # ── Campagne-uitzonderingen (optioneel) ───────────────────────────────────
    df_camp = wb_data.get("Campagne-uitzonderingen")
    if df_camp is not None and not df_camp.dropna(how="all").empty:
        df_camp = df_camp.copy()
        df_camp.columns = [str(c).strip() for c in df_camp.columns]
        df_camp = df_camp.dropna(how="all")
        seen_camp: dict[tuple, int] = {}

        for row_idx, row in df_camp.iterrows():
            src_row     = int(row_idx) + 2
            account_raw = str(row.get("Account", "")).strip()
            cid_raw     = str(row.get("Klant-ID", "")).strip()
            ctype_raw   = str(row.get("Campagnetype", "")).strip()
            camp_raw    = str(row.get("Campagne / plaats", row.get("Campagne", ""))).strip()
            cpc_raw     = _get_col(row, "CPC", "Uitzondering CPC", "Standaard CPC instellen")

            if not account_raw or account_raw == "nan":
                continue
            norm_account = normalize_name(account_raw)
            ctype, ctype_err = parse_campaign_type(ctype_raw)
            if ctype_err:
                errors.append(f"Campagne-uitzonderingen rij {src_row}: {ctype_err}")
                continue
            cpc_cents, cpc_err = parse_cpc_cents(cpc_raw)
            if cpc_err:
                errors.append(f"Campagne-uitzonderingen rij {src_row} ({account_raw}): {cpc_err}")
                continue

            norm_camp = normalize_name(camp_raw)
            dup_key   = (norm_account, ctype, norm_camp)
            if dup_key in seen_camp and seen_camp[dup_key] != cpc_cents:
                errors.append(
                    f"Campagne-uitzonderingen rij {src_row}: Conflicterende CPC voor "
                    f"'{account_raw}' / '{camp_raw}' (twee verschillende bedragen)."
                )
                continue
            seen_camp[dup_key] = cpc_cents

            campaign_exceptions.append(CampaignCpcException(
                account_name             = account_raw,
                normalized_account       = norm_account,
                customer_id_raw          = cid_raw,
                campaign_type            = ctype,
                campaign_name            = camp_raw,
                normalized_campaign_name = norm_camp,
                place_name               = camp_raw,
                normalized_place_name    = norm_camp,
                cpc_in_cents             = cpc_cents,
                source_sheet             = "Campagne-uitzonderingen",
                source_row               = src_row,
            ))

    # ── Adv.groep-uitzonderingen (optioneel) ──────────────────────────────────
    df_ag = wb_data.get("Adv.groep-uitzonderingen")
    if df_ag is not None and not df_ag.dropna(how="all").empty:
        df_ag = df_ag.copy()
        df_ag.columns = [str(c).strip() for c in df_ag.columns]
        df_ag = df_ag.dropna(how="all")

        for row_idx, row in df_ag.iterrows():
            src_row     = int(row_idx) + 2
            account_raw = str(row.get("Account", "")).strip()
            cid_raw     = str(row.get("Klant-ID", "")).strip()
            ctype_raw   = str(row.get("Campagnetype", "")).strip()
            camp_raw    = str(row.get("Campagne", "")).strip()
            ag_raw      = str(row.get("Advertentiegroep", "")).strip()
            cpc_raw     = _get_col(row, "CPC", "Uitzondering CPC")

            if not account_raw or account_raw == "nan":
                continue
            norm_account = normalize_name(account_raw)
            ctype, ctype_err = parse_campaign_type(ctype_raw)
            if ctype_err:
                errors.append(f"Adv.groep-uitzonderingen rij {src_row}: {ctype_err}")
                continue
            cpc_cents, cpc_err = parse_cpc_cents(cpc_raw)
            if cpc_err:
                errors.append(f"Adv.groep-uitzonderingen rij {src_row} ({account_raw}): {cpc_err}")
                continue

            ad_group_exceptions.append(AdGroupCpcException(
                account_name             = account_raw,
                normalized_account       = norm_account,
                customer_id_raw          = cid_raw,
                campaign_type            = ctype,
                campaign_name            = camp_raw,
                ad_group_name            = ag_raw,
                normalized_ad_group_name = normalize_name(ag_raw),
                cpc_in_cents             = cpc_cents,
                source_sheet             = "Adv.groep-uitzonderingen",
                source_row               = src_row,
            ))

    # ── Zoekwoord-uitzonderingen (optioneel) ──────────────────────────────────
    df_kw = wb_data.get("Zoekwoord-uitzonderingen")
    if df_kw is not None and not df_kw.dropna(how="all").empty:
        df_kw = df_kw.copy()
        df_kw.columns = [str(c).strip() for c in df_kw.columns]
        df_kw = df_kw.dropna(how="all")

        for row_idx, row in df_kw.iterrows():
            src_row     = int(row_idx) + 2
            account_raw = str(row.get("Account", "")).strip()
            cid_raw     = str(row.get("Klant-ID", "")).strip()
            ctype_raw   = str(row.get("Campagnetype", "")).strip()
            camp_raw    = str(row.get("Campagne", "")).strip()
            ag_raw      = str(row.get("Advertentiegroep", "")).strip()
            kw_raw      = str(row.get("Zoekwoord", "")).strip()
            mt_raw      = str(row.get("Matchtype", "")).strip()
            cpc_raw     = _get_col(row, "CPC", "Uitzondering CPC")

            if not account_raw or account_raw == "nan":
                continue
            norm_account = normalize_name(account_raw)
            ctype, ctype_err = parse_campaign_type(ctype_raw)
            if ctype_err:
                errors.append(f"Zoekwoord-uitzonderingen rij {src_row}: {ctype_err}")
                continue
            cpc_cents, cpc_err = parse_cpc_cents(cpc_raw)
            if cpc_err:
                errors.append(f"Zoekwoord-uitzonderingen rij {src_row} ({account_raw}): {cpc_err}")
                continue

            keyword_exceptions.append(KeywordCpcException(
                account_name             = account_raw,
                normalized_account       = norm_account,
                customer_id_raw          = cid_raw,
                campaign_type            = ctype,
                campaign_name            = camp_raw,
                ad_group_name            = ag_raw,
                normalized_ad_group_name = normalize_name(ag_raw),
                keyword_text             = kw_raw,
                normalized_keyword_text  = normalize_name(kw_raw),
                match_type               = normalize_match_type(mt_raw),
                cpc_in_cents             = cpc_cents,
                source_sheet             = "Zoekwoord-uitzonderingen",
                source_row               = src_row,
            ))

    return CpcAdviceImport(
        metadata             = metadata,
        default_rules        = default_rules,
        campaign_exceptions  = campaign_exceptions,
        ad_group_exceptions  = ad_group_exceptions,
        keyword_exceptions   = keyword_exceptions,
        validation_errors    = errors,
        validation_warnings  = warnings,
    )


# ── Pauzeringsbestand parseren ────────────────────────────────────────────────

def parse_pause_advice(file_bytes: bytes, file_name: str = "onbekend") -> PauseAdviceImport:
    file_type, wb_data, metadata = detect_advice_file_type(file_bytes, file_name)

    errors:   list[str] = []
    warnings: list[str] = []

    if file_type == AdviceFileType.UNKNOWN:
        errors.append(
            f"Bestand '{file_name}' is niet herkend als pauzeringsbestand. "
            f"Verwacht minimaal één tabblad uit: {', '.join(sorted(_PAUSE_POSSIBLE_SHEETS))} "
            f"met kolommen: {', '.join(sorted(_PAUSE_REQUIRED_COLS))}. "
            f"Aanwezige tabbladen: {', '.join(sorted(wb_data.keys()))}."
        )
        return PauseAdviceImport(metadata=metadata, validation_errors=errors)

    if file_type == AdviceFileType.CPC_ADVICE:
        errors.append(
            f"Bestand '{file_name}' is herkend als CPC-adviesbestand, niet als pauzeringsbestand."
        )
        return PauseAdviceImport(metadata=metadata, validation_errors=errors)

    if file_type == AdviceFileType.AMBIGUOUS:
        warnings.append(
            f"Bestand '{file_name}' voldoet aan zowel de CPC- als de pauzeringsstructuur. "
            f"Verwerkt als pauzeringsbestand."
        )

    city_place_rules:     list[CityPlaceRule]    = []
    local_campaign_rules: list[LocalCampaignRule] = []

    # ── + stad - plaatsen ─────────────────────────────────────────────────────
    df_stad = wb_data.get("+ stad - plaatsen")
    if df_stad is not None and not df_stad.dropna(how="all").empty:
        df_stad = df_stad.copy()
        df_stad.columns = [str(c).strip() for c in df_stad.columns]
        df_stad = df_stad.dropna(how="all")
        seen_stad: dict[tuple, tuple[BuildAction, int]] = {}

        for row_idx, row in df_stad.iterrows():
            src_row     = int(row_idx) + 2
            account_raw = str(row.get("Account", "")).strip()
            cid_raw     = str(row.get("Klant-ID", "")).strip()
            place_raw   = str(row.get("Campagne / plaats", "")).strip()
            action_raw  = str(row.get("Definitieve bouwactie", "")).strip()
            reason_raw  = str(_get_col(row, "Statusreden", "Bouwinstructie") or "").strip()

            if not account_raw or account_raw == "nan":
                continue
            if not place_raw or place_raw == "nan":
                errors.append(f"+ stad - plaatsen rij {src_row}: 'Campagne / plaats' ontbreekt.")
                continue

            action, act_err = parse_build_action(action_raw)
            if act_err:
                errors.append(
                    f"+ stad - plaatsen rij {src_row} ({account_raw} / {place_raw}): {act_err}"
                )
                continue

            norm_acc   = normalize_name(account_raw)
            norm_place = normalize_name(place_raw)
            dup_key    = (norm_acc, norm_place)

            if dup_key in seen_stad:
                prev_action, prev_row = seen_stad[dup_key]
                if prev_action != action:
                    errors.append(
                        f"+ stad - plaatsen: Conflicterende bouwactie voor "
                        f"'{account_raw}' / '{place_raw}' "
                        f"(rij {src_row}: {action.value}, rij {prev_row}: {prev_action.value}). "
                        f"Los dit handmatig op."
                    )
                continue
            seen_stad[dup_key] = (action, src_row)

            city_place_rules.append(CityPlaceRule(
                account_name          = account_raw,
                normalized_account    = norm_acc,
                customer_id_raw       = cid_raw,
                place_name            = place_raw,
                normalized_place_name = norm_place,
                action                = action,
                status_reason         = reason_raw,
                source_sheet          = "+ stad - plaatsen",
                source_row            = src_row,
            ))

    # ── Lokale campagnes ──────────────────────────────────────────────────────
    df_lok = wb_data.get("Lokale campagnes")
    if df_lok is not None and not df_lok.dropna(how="all").empty:
        df_lok = df_lok.copy()
        df_lok.columns = [str(c).strip() for c in df_lok.columns]
        df_lok = df_lok.dropna(how="all")
        seen_lok: dict[tuple, tuple[BuildAction, int]] = {}

        for row_idx, row in df_lok.iterrows():
            src_row     = int(row_idx) + 2
            account_raw = str(row.get("Account", "")).strip()
            cid_raw     = str(row.get("Klant-ID", "")).strip()
            camp_raw    = str(row.get("Campagne / plaats", "")).strip()
            action_raw  = str(row.get("Definitieve bouwactie", "")).strip()
            reason_raw  = str(_get_col(row, "Statusreden", "Bouwinstructie") or "").strip()

            if not account_raw or account_raw == "nan":
                continue

            action, act_err = parse_build_action(action_raw)
            if act_err:
                errors.append(f"Lokale campagnes rij {src_row} ({account_raw}): {act_err}")
                continue

            norm_acc  = normalize_name(account_raw)
            norm_camp = normalize_name(camp_raw)
            dup_key   = (norm_acc, norm_camp)

            if dup_key in seen_lok:
                prev_action, prev_row = seen_lok[dup_key]
                if prev_action != action:
                    errors.append(
                        f"Lokale campagnes: Conflicterende bouwactie voor "
                        f"'{account_raw}' / '{camp_raw}' "
                        f"(rij {src_row}: {action.value}, rij {prev_row}: {prev_action.value}). "
                        f"Los dit handmatig op."
                    )
                continue
            seen_lok[dup_key] = (action, src_row)

            local_campaign_rules.append(LocalCampaignRule(
                account_name             = account_raw,
                normalized_account       = norm_acc,
                customer_id_raw          = cid_raw,
                campaign_name            = camp_raw,
                normalized_campaign_name = norm_camp,
                place_name               = camp_raw,
                normalized_place_name    = norm_camp,
                action                   = action,
                status_reason            = reason_raw,
                source_sheet             = "Lokale campagnes",
                source_row               = src_row,
            ))

    return PauseAdviceImport(
        metadata             = metadata,
        city_place_rules     = city_place_rules,
        local_campaign_rules = local_campaign_rules,
        validation_errors    = errors,
        validation_warnings  = warnings,
    )
