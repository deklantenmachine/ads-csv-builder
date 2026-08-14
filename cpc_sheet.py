"""CPC-lookup via het centrale zoekwoorden-adviesblad (Schoorsteenveger)."""

import pandas as pd

_LAND_NORMALIZE: dict[str, str] = {
    "nl":         "nederland",
    "nederland":  "nederland",
    "be":         "belgie",
    "belgie":     "belgie",
    "belgië":     "belgie",
}


def normalize_land(land: str) -> str:
    return _LAND_NORMALIZE.get(land.strip().lower(), land.strip().lower())


class KeywordCpcSheet:
    """
    Laadt het centrale keyword-adviesblad en biedt per-zoekwoord CPC-lookup.

    Kolommen in het blad:
      A  Land                          → normalize_land
      B  Merktype                      → lowercase
      C  Account (portaal)             → lowercase
      D  Plaats                        → lowercase
      H  Zoekwoord                     → lowercase
      I  CPC-advies (definitief)       → cent-waarde (float * 100, komma→punt)
      M  Is fallback voor nieuwe plaats → TRUE / FALSE
    """

    def __init__(self) -> None:
        self._lookup:   dict[tuple, int] = {}   # (land, merktype, account, plaats, zoekwoord) → cents
        self._fallback: dict[tuple, int] = {}   # (land, merktype, account) → cents

    def load(self, sheet_url: str, sheet_name: str) -> None:
        from builder import load_sheet  # lazy import to avoid circular dependency
        df = load_sheet(sheet_url, sheet_name)

        lookup:   dict[tuple, int] = {}
        fallback: dict[tuple, int] = {}

        for _, row in df.iterrows():
            land      = normalize_land(str(row.get("Land", "") or ""))
            merktype  = str(row.get("Merktype", "")                    or "").strip().lower()
            account   = str(row.get("Account (portaal)", "")           or "").strip().lower()
            plaats    = str(row.get("Plaats", "")                      or "").strip().lower()
            zoekwoord = str(row.get("Zoekwoord", "")                   or "").strip().lower()
            cpc_raw   = str(row.get("CPC-advies (definitief)", "")     or "").strip().replace(",", ".")
            is_fb     = str(row.get("Is fallback voor nieuwe plaats", "") or "").strip().upper() == "TRUE"

            try:
                cpc_cents = round(float(cpc_raw) * 100)
            except (ValueError, TypeError):
                continue

            key = (land, merktype, account, plaats, zoekwoord)
            lookup[key] = cpc_cents

            if is_fb:
                fallback[(land, merktype, account)] = cpc_cents

        self._lookup   = lookup
        self._fallback = fallback

    @property
    def loaded(self) -> bool:
        return bool(self._lookup)

    def get_cpc(
        self,
        land: str,
        merktype: str,
        account: str,
        plaats: str,
        zoekwoord: str,
    ) -> int | None:
        key = (
            normalize_land(land),
            merktype.strip().lower(),
            account.strip().lower(),
            plaats.strip().lower(),
            zoekwoord.strip().lower(),
        )
        return self._lookup.get(key)

    def get_fallback_cpc(self, land: str, merktype: str, account: str) -> int | None:
        key = (
            normalize_land(land),
            merktype.strip().lower(),
            account.strip().lower(),
        )
        return self._fallback.get(key)
