"""
CampaignBuildRulesEngine — determines per-city build decisions.
Reads parsed advice imports; never modifies templates or bid strategies.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field

import pandas as pd

from advice_parser import (
    BuildAction, CampaignType,
    CpcAdviceImport, PauseAdviceImport,
    DefaultCpcRule, CampaignCpcException, AdGroupCpcException, KeywordCpcException,
    CityPlaceRule, LocalCampaignRule,
    normalize_name, _MANUAL_CPC_STRATEGIES,
)


@dataclass
class AuditEntry:
    source_file:    str
    source_sheet:   str
    source_row:     int
    rule_type:      str
    applied_value:  str
    decision_level: str   # BUILD | CAMPAIGN | AD_GROUP | KEYWORD


@dataclass
class BuildDecision:
    should_build:       bool
    final_status:       str             # "Paused" | "Enabled"
    bidding_strategy:   str             # unchanged from template
    campaign_cpc_cents: int | None      # None = no CPC advice loaded
    ad_group_cpc_cents: int | None
    keyword_cpc_cents:  int | None      # None = inherit from ad group
    matched_rules:      list            = field(default_factory=list)
    warnings:           list[str]       = field(default_factory=list)
    blocking_errors:    list[str]       = field(default_factory=list)
    audit_trail:        list[AuditEntry] = field(default_factory=list)


_FUZZY_THRESHOLD = 0.60
_STOP_WORDS = {"de", "den", "der", "het", "een", "en", "of", "van", "voor", "in", "op"}


def _normalize_for_fuzzy(name: str) -> str:
    """Strip punctuation and stop-words for token-based similarity."""
    tokens = re.sub(r"[^a-z0-9\s]", " ", name).split()
    return " ".join(t for t in tokens if t not in _STOP_WORDS and len(t) > 1)


def _account_similarity(a: str, b: str) -> float:
    """Combined SequenceMatcher + token-Jaccard similarity on normalized strings."""
    seq_ratio = difflib.SequenceMatcher(None, a, b).ratio()
    a_tokens = set(_normalize_for_fuzzy(a).split())
    b_tokens = set(_normalize_for_fuzzy(b).split())
    union = a_tokens | b_tokens
    jaccard = len(a_tokens & b_tokens) / len(union) if union else 0.0
    return max(seq_ratio, jaccard)


def fuzzy_best_account_match(norm_query: str, norm_pool: set[str]) -> tuple[str | None, float]:
    """Return (best_match, score) from norm_pool, or (None, 0) if below threshold."""
    best, best_score = None, 0.0
    for candidate in norm_pool:
        score = _account_similarity(norm_query, candidate)
        if score > best_score:
            best, best_score = candidate, score
    if best_score >= _FUZZY_THRESHOLD:
        return best, best_score
    return None, 0.0


def extract_template_bid_strategy(df: pd.DataFrame) -> str:
    """Read bid strategy from template DataFrame column 'Bid Strategy Type'."""
    col = "Bid Strategy Type"
    if col not in df.columns:
        return ""
    values = df[col].dropna().astype(str)
    values = values[values.str.strip() != ""]
    return values.iloc[0].strip() if not values.empty else ""


class CampaignBuildRulesEngine:
    """
    Initialized once per build_all run.
    Call get_city_decision() for +Stad decisions and get_local_decision() for lokaal.
    """

    def __init__(
        self,
        cpc_import:            CpcAdviceImport | None,
        pause_import:          PauseAdviceImport | None,
        template_bid_strategy: str = "",
    ):
        self._cpc      = cpc_import
        self._pause    = pause_import
        self._strategy = template_bid_strategy.strip()
        self._supports_manual_cpc = (
            self._strategy.lower() in _MANUAL_CPC_STRATEGIES
        )
        # Build normalized-account pools for fuzzy matching
        self._cpc_norm_accounts: set[str] = set()
        self._cpc_account_map: dict[str, str] = {}   # norm_query → norm_in_file
        if cpc_import:
            for r in cpc_import.default_rules:
                self._cpc_norm_accounts.add(r.normalized_account)
        self._pause_norm_accounts: set[str] = set()
        self._pause_account_map: dict[str, str] = {}
        if pause_import:
            for r in pause_import.city_place_rules:
                self._pause_norm_accounts.add(r.normalized_account)
            for r in pause_import.local_campaign_rules:
                self._pause_norm_accounts.add(r.normalized_account)
        # Cache: norm_query → (resolved_norm_account, warning_msg | None)
        self._account_cache: dict[str, tuple[str | None, str | None]] = {}

    # ── Public API ────────────────────────────────────────────────────────────

    def get_city_decision(
        self,
        account: str,
        city:    str,
    ) -> BuildDecision:
        """Build decision for a +Stad city entry."""
        return self._decide(
            account       = account,
            campaign_type = CampaignType.REGULAR_CITY,
            city          = city,
            campaign_name = "",
        )

    def get_local_decision(
        self,
        account:       str,
        city:          str,
        campaign_name: str = "",
    ) -> BuildDecision:
        """Build decision for a lokale campagne entry."""
        return self._decide(
            account       = account,
            campaign_type = CampaignType.LOCAL,
            city          = city,
            campaign_name = campaign_name,
        )

    def validate_cpc_compatibility(self, account: str, campaign_type: CampaignType) -> str | None:
        """
        Returns a blocking error string if the template bid strategy does not support CPC,
        or None when compatible.
        """
        if not self._cpc:
            return None
        if not self._supports_manual_cpc:
            return (
                f"Account '{account}' / {campaign_type.value}: "
                f"het sjabloonbestand gebruikt biedstrategie '{self._strategy}' "
                f"die geen handmatige CPC-bedragen ondersteunt. "
                f"Er is wel een CPC-advies geladen, maar de biedstrategie wordt "
                f"niet automatisch aangepast."
            )
        return None

    def unmatched_cpc_warnings(self, known_accounts: set[str]) -> list[str]:
        """Return warnings for CPC rules whose account has no match in the build."""
        if not self._cpc:
            return []
        warnings = []
        norm_known = {normalize_name(a) for a in known_accounts}
        for rule in self._cpc.default_rules:
            na = rule.normalized_account
            if na not in norm_known:
                _, score = fuzzy_best_account_match(na, norm_known)
                if score < _FUZZY_THRESHOLD:
                    warnings.append(
                        f"CPC-advies: account '{rule.account_name}' uit het adviesbestand "
                        f"is niet gevonden in de plaatsenlijst."
                    )
        return warnings

    def unmatched_pause_warnings(
        self, known_cities: set[str], known_accounts: set[str] | None = None
    ) -> list[str]:
        """Return warnings for pause rules whose city has no match in the build.

        When known_accounts is provided, only rules for accounts in the active build
        are checked — rules for other accounts are silently ignored.
        """
        if not self._pause:
            return []
        warnings = []
        norm_cities   = {normalize_name(c) for c in known_cities}
        norm_accounts = {normalize_name(a) for a in known_accounts} if known_accounts else None
        for rule in self._pause.city_place_rules:
            if norm_accounts is not None:
                resolved, _ = fuzzy_best_account_match(rule.normalized_account, norm_accounts)
                if not resolved:
                    continue  # account niet in actieve build — sla over
            if rule.normalized_place_name not in norm_cities:
                warnings.append(
                    f"Pauzeringsregel voor '{rule.place_name}' (account: {rule.account_name}) "
                    f"kon niet worden toegepast: plaats niet gevonden in de actuele plaatsenlijst."
                )
        return warnings

    # ── Internal ──────────────────────────────────────────────────────────────

    def _decide(
        self,
        account:       str,
        campaign_type: CampaignType,
        city:          str,
        campaign_name: str,
    ) -> BuildDecision:
        norm_account  = normalize_name(account)
        norm_city     = normalize_name(city)
        norm_campaign = normalize_name(campaign_name)

        decision = BuildDecision(
            should_build       = True,
            final_status       = "Paused",   # templates always start paused
            bidding_strategy   = self._strategy,
            campaign_cpc_cents = None,
            ad_group_cpc_cents = None,
            keyword_cpc_cents  = None,
        )

        # Resolve account names (supports fuzzy matching)
        pause_account, pause_warn = self._resolve_account(
            norm_account, self._pause_norm_accounts, account,
        ) if self._pause else (norm_account, None)
        cpc_account, cpc_warn = self._resolve_account(
            norm_account, self._cpc_norm_accounts, account,
        ) if self._cpc else (norm_account, None)
        if pause_warn:
            decision.warnings.append(pause_warn)
        if cpc_warn and cpc_warn != pause_warn:
            decision.warnings.append(cpc_warn)

        # Use resolved account names in all subsequent lookups
        eff_pause_account = pause_account or norm_account
        eff_cpc_account   = cpc_account   or norm_account

        # ── Stap 1: pauze- en beëindigingsregels ─────────────────────────────
        if self._pause:
            if campaign_type == CampaignType.REGULAR_CITY:
                rule = self._find_city_rule(eff_pause_account, norm_city)
                if rule:
                    if rule.action == BuildAction.DO_NOT_BUILD:
                        decision.should_build = False
                        decision.audit_trail.append(self._entry(
                            self._pause.metadata.original_file_name,
                            rule, "DO_NOT_BUILD", "niet bouwen", "BUILD",
                        ))
                        return decision
                    decision.final_status = "Paused"
                    decision.matched_rules.append(rule)
                    decision.audit_trail.append(self._entry(
                        self._pause.metadata.original_file_name,
                        rule, "PAUSE", "gepauzeerd bouwen", "BUILD",
                    ))

            elif campaign_type == CampaignType.LOCAL:
                rule = self._find_local_rule(eff_pause_account, norm_city, norm_campaign)
                if rule:
                    if rule.action == BuildAction.DO_NOT_BUILD:
                        decision.should_build = False
                        decision.audit_trail.append(self._entry(
                            self._pause.metadata.original_file_name,
                            rule, "DO_NOT_BUILD", "niet bouwen", "BUILD",
                        ))
                        return decision
                    decision.final_status = "Paused"
                    decision.matched_rules.append(rule)
                    decision.audit_trail.append(self._entry(
                        self._pause.metadata.original_file_name,
                        rule, "PAUSE", "gepauzeerd bouwen", "BUILD",
                    ))

        # ── Stap 2: biedstrategie-compatibiliteit ────────────────────────────
        if self._cpc:
            compat_err = self.validate_cpc_compatibility(account, campaign_type)
            if compat_err:
                decision.blocking_errors.append(compat_err)
                return decision

            # ── Stap 3: standaard-CPC ────────────────────────────────────────
            default_rule = self._find_default(eff_cpc_account, campaign_type)
            if default_rule:
                decision.campaign_cpc_cents = default_rule.cpc_in_cents
                decision.ad_group_cpc_cents = default_rule.cpc_in_cents
                decision.matched_rules.append(default_rule)
                decision.audit_trail.append(self._entry(
                    self._cpc.metadata.original_file_name,
                    default_rule, "DEFAULT_CPC",
                    f"€{default_rule.cpc_in_cents/100:.2f}", "CAMPAIGN",
                ))
            else:
                if cpc_account:
                    decision.warnings.append(
                        f"Geen CPC-advies gevonden voor account '{account}' / "
                        f"{campaign_type.value}."
                    )
                # else: account not in CPC file at all, already warned via unmatched_cpc_warnings

            # ── Stap 4: campagne-uitzondering ─────────────────────────────────
            camp_exc = self._find_campaign_exc(eff_cpc_account, campaign_type, norm_city, norm_campaign)
            if camp_exc:
                decision.campaign_cpc_cents = camp_exc.cpc_in_cents
                decision.ad_group_cpc_cents = camp_exc.cpc_in_cents
                decision.matched_rules.append(camp_exc)
                decision.audit_trail.append(self._entry(
                    self._cpc.metadata.original_file_name,
                    camp_exc, "CAMPAIGN_EXCEPTION",
                    f"€{camp_exc.cpc_in_cents/100:.2f}", "CAMPAIGN",
                ))

        return decision

    # ── Account resolution (exact → fuzzy) ───────────────────────────────────

    def _resolve_account(
        self, norm_query: str, pool: set[str], original_name: str,
    ) -> tuple[str | None, str | None]:
        """
        Returns (resolved_norm_account, warning_msg).
        Exact match → no warning. Fuzzy match above threshold → warning. No match → (None, None).
        """
        cache_key = (norm_query, id(pool))
        if cache_key in self._account_cache:
            return self._account_cache[cache_key]

        if norm_query in pool:
            result = norm_query, None
        else:
            best, score = fuzzy_best_account_match(norm_query, pool)
            if best:
                # find a display name for the matched account
                result = best, (
                    f"Account '{original_name}' fuzzy-gematcht op adviesbestand-account "
                    f"(overeenkomst {score:.0%}). Controleer of dit klopt."
                )
            else:
                result = None, None
        self._account_cache[cache_key] = result
        return result

    # ── Lookup helpers ────────────────────────────────────────────────────────

    def _find_city_rule(self, norm_account: str, norm_city: str) -> CityPlaceRule | None:
        if not self._pause:
            return None
        for r in self._pause.city_place_rules:
            if r.normalized_account == norm_account and r.normalized_place_name == norm_city:
                return r
        return None

    def _find_local_rule(
        self, norm_account: str, norm_city: str, norm_campaign: str,
    ) -> LocalCampaignRule | None:
        if not self._pause:
            return None
        for r in self._pause.local_campaign_rules:
            if r.normalized_account != norm_account:
                continue
            # match by city name inside campaign name, or by exact campaign name
            if norm_city and norm_city in r.normalized_campaign_name:
                return r
            if norm_campaign and norm_campaign == r.normalized_campaign_name:
                return r
        return None

    def _find_default(
        self, norm_account: str, campaign_type: CampaignType,
    ) -> DefaultCpcRule | None:
        if not self._cpc:
            return None
        for r in self._cpc.default_rules:
            if r.normalized_account == norm_account and r.campaign_type == campaign_type:
                return r
        return None

    def _find_campaign_exc(
        self,
        norm_account:  str,
        campaign_type: CampaignType,
        norm_city:     str,
        norm_campaign: str,
    ) -> CampaignCpcException | None:
        if not self._cpc:
            return None
        matches = [
            r for r in self._cpc.campaign_exceptions
            if r.normalized_account == norm_account
            and r.campaign_type == campaign_type
            and (r.normalized_place_name == norm_city
                 or r.normalized_campaign_name == norm_campaign)
        ]
        if len(matches) == 1:
            return matches[0]
        return None  # 0 = geen match, >1 = ambigu → niet toepassen

    def get_ad_group_cpc(
        self,
        account:        str,
        campaign_type:  CampaignType,
        campaign_name:  str,
        ad_group_name:  str,
        fallback_cents: int | None,
    ) -> tuple[int | None, AuditEntry | None]:
        """Returns (cpc_cents, audit_entry) for an ad group; fallback when no exception."""
        if not self._cpc:
            return fallback_cents, None
        raw_norm = normalize_name(account)
        resolved, _ = self._resolve_account(raw_norm, self._cpc_norm_accounts, account)
        norm_account = resolved or raw_norm
        norm_ag      = normalize_name(ad_group_name)
        for r in self._cpc.ad_group_exceptions:
            if (r.normalized_account == norm_account
                    and r.campaign_type == campaign_type
                    and r.normalized_ad_group_name == norm_ag):
                entry = AuditEntry(
                    source_file    = self._cpc.metadata.original_file_name,
                    source_sheet   = r.source_sheet,
                    source_row     = r.source_row,
                    rule_type      = "AD_GROUP_EXCEPTION",
                    applied_value  = f"€{r.cpc_in_cents/100:.2f}",
                    decision_level = "AD_GROUP",
                )
                return r.cpc_in_cents, entry
        return fallback_cents, None

    def get_keyword_cpc(
        self,
        account:       str,
        campaign_type: CampaignType,
        ad_group_name: str,
        keyword_text:  str,
        match_type:    str,
    ) -> tuple[int | None, AuditEntry | None]:
        """Returns (cpc_cents, audit_entry) for a keyword; None when no exception."""
        if not self._cpc:
            return None, None
        raw_norm = normalize_name(account)
        resolved, _ = self._resolve_account(raw_norm, self._cpc_norm_accounts, account)
        norm_account = resolved or raw_norm
        norm_ag      = normalize_name(ad_group_name)
        norm_kw      = normalize_name(keyword_text)
        norm_mt      = match_type.upper()
        for r in self._cpc.keyword_exceptions:
            if (r.normalized_account == norm_account
                    and r.campaign_type == campaign_type
                    and r.normalized_ad_group_name == norm_ag
                    and r.normalized_keyword_text == norm_kw
                    and r.match_type == norm_mt):
                entry = AuditEntry(
                    source_file    = self._cpc.metadata.original_file_name,
                    source_sheet   = r.source_sheet,
                    source_row     = r.source_row,
                    rule_type      = "KEYWORD_EXCEPTION",
                    applied_value  = f"€{r.cpc_in_cents/100:.2f}",
                    decision_level = "KEYWORD",
                )
                return r.cpc_in_cents, entry
        return None, None

    @staticmethod
    def _entry(file_name: str, rule, rule_type: str, value: str, level: str) -> AuditEntry:
        return AuditEntry(
            source_file    = file_name,
            source_sheet   = rule.source_sheet,
            source_row     = rule.source_row,
            rule_type      = rule_type,
            applied_value  = value,
            decision_level = level,
        )
