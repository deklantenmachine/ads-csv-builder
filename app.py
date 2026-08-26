"""Streamlit interface voor de Google Ads CSV builder."""

import io
import json
import os
import tempfile

import streamlit as st
import pandas as pd

from builder import build_all, df_to_csv_bytes, output_filename, load_sheet
from advice_parser import (
    parse_cpc_advice, parse_pause_advice,
    detect_advice_file_type, AdviceFileType,
)
from cpc_sheet import KeywordCpcSheet

# ── Branch configuratie ────────────────────────────────────────────────────────
_BRANCHES_PATH = os.path.join(os.path.dirname(__file__), "branches.json")

def _load_branches() -> dict:
    try:
        with open(_BRANCHES_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _save_branches(data: dict):
    with open(_BRANCHES_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def _load_file_bytes(path: str) -> bytes | None:
    if not path:
        return None
    if not os.path.isabs(path):
        path = os.path.join(os.path.dirname(__file__), path)
    if os.path.isfile(path):
        with open(path, "rb") as f:
            return f.read()
    return None

def _branch_files_dir(branch_name: str) -> str:
    d = os.path.join(os.path.dirname(__file__), "branch_files", branch_name)
    os.makedirs(d, exist_ok=True)
    return d

def _save_branch_advice_file(branch_name: str, kind: str, file_bytes: bytes, orig_name: str) -> str:
    """Sla adviesbestand op als branch-standaard. Geeft het opgeslagen pad terug."""
    ext = os.path.splitext(orig_name)[1] or ".xlsx"
    dest = os.path.join(_branch_files_dir(branch_name), f"{kind}{ext}")
    with open(dest, "wb") as f:
        f.write(file_bytes)
    return dest

st.set_page_config(page_title="Ads CSV Builder", page_icon="📦", layout="centered")

# ── Authenticatie ─────────────────────────────────────────────────────────────
def _check_auth() -> bool:
    try:
        correct = st.secrets["password"]
    except Exception:
        return True  # geen secret ingesteld → lokaal gebruik, altijd toestaan

    if st.session_state.get("authenticated"):
        return True

    st.title("📦 Ads CSV Builder")
    pwd = st.text_input("Wachtwoord", type="password", key="pwd_input")
    if st.button("Inloggen"):
        if pwd == correct:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Onjuist wachtwoord.")
    st.stop()

_check_auth()

st.title("📦 Google Ads CSV Builder")

# ── Branch selector ────────────────────────────────────────────────────────────
_branches = _load_branches()
_branch_names   = list(_branches.keys())
_MERK_TYPES     = ["Eigen merk", "Overkoepelend merk (portaal)"]

st.header("0. Branche & merktype")
_col_branch, _col_merk, _col_land, _col_edit = st.columns([2, 2, 1, 1])
with _col_branch:
    selected_branch = st.selectbox("Branche", _branch_names, key="selected_branch")
with _col_merk:
    selected_merk = st.selectbox("Merktype", _MERK_TYPES, key="selected_merk")

# Land selector: toon alleen als de branche land-aware is (merktype heeft "NL"/"BE" sub-keys)
_merk_raw = _branches.get(selected_branch, {}).get(selected_merk, {})
_is_land_aware = any(k in _merk_raw for k in ["NL", "BE"])
_landen = [k for k in ["NL", "BE"] if k in _merk_raw] if _is_land_aware else []

with _col_land:
    if _is_land_aware:
        selected_land = st.selectbox("Land", _landen, key="selected_land")
    else:
        selected_land = "NL"
        st.write("")

with _col_edit:
    st.write("")
    show_branch_config = st.toggle("Bewerk standaarden", key="show_branch_config")

# Haal de juiste config op (land-aware of flat)
_branch_cfg = _merk_raw.get(selected_land, {}) if _is_land_aware else _merk_raw

# Branch-level UI configuratie (branch-niveau samenvoegen met land-niveau, land wint)
_ui_branch = _branches.get(selected_branch, {}).get("_ui", {})
_ui_land   = _branch_cfg.get("_ui", {})   # land-specifiek (NL/BE) overschrijft branch-niveau
_ui_cfg    = {**_ui_branch, **_ui_land}
_toon_uitgebreide_merk_info = _ui_cfg.get("toon_uitgebreide_merk_info", True)
_toon_adviesbestanden       = _ui_cfg.get("toon_adviesbestanden",        True)
_toon_fallback_cpc          = _ui_cfg.get("toon_fallback_cpc",           True)
_toon_pauze                 = _ui_cfg.get("toon_pauze",                  True)
# Placeholders voor bedrijfsnaam en review score (branche-specifiek instelbaar)
_placeholder_korte   = _ui_cfg.get("placeholder_korte_naam",   "DCN")
_placeholder_lange   = _ui_cfg.get("placeholder_lange_naam",   "DCN Dakdekkers")
_placeholder_review  = _ui_cfg.get("placeholder_review_score", "9,2")
_label_review        = _ui_cfg.get("label_review_score",       "Review score")

# ── Centraal CPC-blad (Schoorsteenveger) ─────────────────────────────────────
_kw_cpc_sheet_url  = _branches.get(selected_branch, {}).get("keyword_cpc_sheet_url", "")
_kw_cpc_sheet_name = _branches.get(selected_branch, {}).get("keyword_cpc_sheet_name", "")
_kw_cpc_cache_key  = f"{selected_branch}|{_kw_cpc_sheet_url}|{_kw_cpc_sheet_name}"

if _kw_cpc_sheet_url and _kw_cpc_sheet_name:
    if st.session_state.get("kw_cpc_sheet_key") != _kw_cpc_cache_key:
        try:
            _kw_sheet = KeywordCpcSheet()
            _kw_sheet.load(_kw_cpc_sheet_url, _kw_cpc_sheet_name)
            st.session_state["kw_cpc_sheet"]     = _kw_sheet
            st.session_state["kw_cpc_sheet_key"] = _kw_cpc_cache_key
        except Exception as _e:
            st.warning(f"CPC-blad kon niet geladen worden: {_e}")
            st.session_state["kw_cpc_sheet"]     = None
            st.session_state["kw_cpc_sheet_key"] = ""

_active_kw_cpc_sheet = st.session_state.get("kw_cpc_sheet") if _kw_cpc_sheet_url else None

if _active_kw_cpc_sheet and _active_kw_cpc_sheet.loaded:
    _n_lookup   = len(_active_kw_cpc_sheet._lookup)
    _n_fallback = len(_active_kw_cpc_sheet._fallback)
    st.success(f"✓ CPC-blad geladen: **{_n_lookup:,}** zoekwoord-CPC's, **{_n_fallback}** fallback-accounts")

    with st.expander("🔍 Controleer CPC-lookup", expanded=False):
        st.caption("Kies een account en plaats om te zien welke CPC's worden opgehaald.")
        _merktype_cpc_preview = {
            "Eigen merk": "eigen merk",
            "Overkoepelend merk (portaal)": "schoorsteenbrigade",
        }.get(selected_merk, selected_merk.lower())
        # Unieke accounts (voor dit merktype)
        _preview_accounts = sorted({
            k[2] for k in _active_kw_cpc_sheet._lookup
            if k[1] == _merktype_cpc_preview
        })
        _preview_landen = sorted({
            k[0] for k in _active_kw_cpc_sheet._lookup
            if k[1] == _merktype_cpc_preview
        })
        _pc1, _pc2, _pc3 = st.columns(3)
        with _pc1:
            _sel_land_p = st.selectbox("Land", _preview_landen or ["nederland"], key="cpc_prev_land")
        with _pc2:
            _accs_for_land = sorted({
                k[2] for k in _active_kw_cpc_sheet._lookup
                if k[1] == _merktype_cpc_preview and k[0] == _sel_land_p
            })
            _sel_acc = st.selectbox("Account", _accs_for_land or ["—"], key="cpc_prev_acc")
        with _pc3:
            _plaatsen_for_acc = sorted({
                k[3] for k in _active_kw_cpc_sheet._lookup
                if k[1] == _merktype_cpc_preview and k[0] == _sel_land_p and k[2] == _sel_acc
            })
            _sel_plaats = st.selectbox("Plaats", _plaatsen_for_acc or ["—"], key="cpc_prev_plaats")

        # Toon zoekwoorden + CPC voor deze combinatie
        _preview_rows = [
            {"Zoekwoord": k[4], "CPC": f"€{v/100:.2f}"}
            for k, v in _active_kw_cpc_sheet._lookup.items()
            if k[0] == _sel_land_p and k[1] == _merktype_cpc_preview and k[2] == _sel_acc and k[3] == _sel_plaats
        ]
        if _preview_rows:
            st.dataframe(pd.DataFrame(_preview_rows).sort_values("Zoekwoord"), use_container_width=True, hide_index=True)
        else:
            st.info("Geen zoekwoorden gevonden voor deze combinatie.")

        # Fallback CPC
        _fb_key = (_sel_land_p, _merktype_cpc_preview, _sel_acc)
        _fb_val = _active_kw_cpc_sheet._fallback.get(_fb_key)
        if _fb_val is not None:
            st.info(f"Fallback CPC voor dit account: **€{_fb_val/100:.2f}**")
        else:
            st.warning("Geen fallback CPC gevonden voor dit account.")

# ── Merktype-string voor CPC-blad lookup ─────────────────────────────────────
_MERKTYPE_CPC_MAP = {
    "Eigen merk":                   "eigen merk",
    "Overkoepelend merk (portaal)": "schoorsteenbrigade",
}
_merktype_cpc = _MERKTYPE_CPC_MAP.get(selected_merk, selected_merk.lower())

# CPC-positie toggle: toon alleen als er laagste-paden zijn geconfigureerd
_heeft_laagste = bool(
    _branch_cfg.get("lokaal_laagste_path") or _branch_cfg.get("stad_laagste_path")
)
if _heeft_laagste:
    _cpc_positie = st.radio(
        "CPC positie", ["Hoogste CPC", "Laagste CPC"],
        horizontal=True, key="cpc_positie",
    )
else:
    _cpc_positie = "Hoogste CPC"

if show_branch_config:
    with st.expander(f"Standaardwaarden voor {selected_branch} — {selected_merk}", expanded=True):
        bc_sheet_url  = st.text_input("Standaard sheet URL",   value=_branch_cfg.get("sheet_url", ""),         key="bc_sheet_url")
        bc_sheet_name = st.text_input("Standaard tabblad",     value=_branch_cfg.get("sheet_name", "Testtab"), key="bc_sheet_name")

        st.markdown("**Sjabloon- en variantenbestanden uploaden als standaard:**")
        _bfd = os.path.join(os.path.dirname(__file__), "branch_files", selected_branch, selected_merk)
        os.makedirs(_bfd, exist_ok=True)

        def _save_branch_file(upload, key_name: str, ext: str) -> str | None:
            if upload is None:
                return None
            dest = os.path.join(_bfd, f"{key_name}{ext}")
            with open(dest, "wb") as _f:
                _f.write(upload.getvalue())
            return dest

        st.markdown("*Hoogste CPC sjablonen:*")
        bc_col1, bc_col2 = st.columns(2)
        with bc_col1:
            bc_up_lokaal   = st.file_uploader("Lokaal hoogste CPC (.csv)",   type="csv",  key="bc_up_lokaal")
            bc_up_variants = st.file_uploader("Variantenbestand (.xlsx)",     type="xlsx", key="bc_up_variants")
        with bc_col2:
            bc_up_stad     = st.file_uploader("Stad hoogste CPC (.csv)",      type="csv",  key="bc_up_stad")
            bc_up_cpc      = st.file_uploader("CPC-adviesbestand (.xlsx)",    type="xlsx", key="bc_up_cpc")

        st.markdown("*Laagste CPC sjablonen (optioneel):*")
        bc_col3, bc_col4 = st.columns(2)
        with bc_col3:
            bc_up_lokaal_laagste = st.file_uploader("Lokaal laagste CPC (.csv)", type="csv", key="bc_up_lokaal_laagste")
        with bc_col4:
            bc_up_stad_laagste   = st.file_uploader("Stad laagste CPC (.csv)",   type="csv", key="bc_up_stad_laagste")

        bc_up_pause = st.file_uploader("Pauzeringsbestand (.xlsx)", type="xlsx", key="bc_up_pause")

        # Huidige paden tonen
        for _lbl, _key in [
            ("Lokaal (hoogste)", "lokaal_template_path"), ("Stad (hoogste)", "stad_template_path"),
            ("Lokaal (laagste)", "lokaal_laagste_path"),  ("Stad (laagste)", "stad_laagste_path"),
            ("Varianten", "variants_path"), ("CPC", "cpc_file_path"), ("Pauze", "pause_file_path"),
        ]:
            _p = _branch_cfg.get(_key, "")
            if _p:
                _ok = "✓" if os.path.isfile(_p) else "✗ niet gevonden"
                st.caption(f"{_lbl}: {os.path.basename(_p)} {_ok}")

        if st.button("💾 Sla bestanden op als standaard", key="save_branch_cfg"):
            _new_cfg = {
                "sheet_url":             bc_sheet_url.strip(),
                "sheet_name":            bc_sheet_name.strip(),
                "lokaal_template_path":  _save_branch_file(bc_up_lokaal,         "lokaal",         ".csv")  or _branch_cfg.get("lokaal_template_path", ""),
                "stad_template_path":    _save_branch_file(bc_up_stad,           "stad",           ".csv")  or _branch_cfg.get("stad_template_path", ""),
                "lokaal_laagste_path":   _save_branch_file(bc_up_lokaal_laagste, "lokaal_laagste", ".csv")  or _branch_cfg.get("lokaal_laagste_path", ""),
                "stad_laagste_path":     _save_branch_file(bc_up_stad_laagste,   "stad_laagste",   ".csv")  or _branch_cfg.get("stad_laagste_path", ""),
                "variants_path":         _save_branch_file(bc_up_variants,       "variants",       ".xlsx") or _branch_cfg.get("variants_path", ""),
                "cpc_file_path":         _save_branch_file(bc_up_cpc,            "cpc",            ".xlsx") or _branch_cfg.get("cpc_file_path", ""),
                "pause_file_path":       _save_branch_file(bc_up_pause,          "pause",          ".xlsx") or _branch_cfg.get("pause_file_path", ""),
            }
            if selected_branch not in _branches:
                _branches[selected_branch] = {}
            _branches[selected_branch][selected_merk] = _new_cfg
            _save_branches(_branches)
            _branch_cfg = _new_cfg
            st.success("Standaarden opgeslagen.")
            st.rerun()

st.divider()

# ── session state ─────────────────────────────────────────────────────────────
_defaults = {
    "results":          None,
    "errors":           [],
    "log":              [],
    "klanten":          [],
    "sheet_loaded":     False,
    "cpc_import":       None,
    "pause_import":     None,
    "cpc_confirmed":    False,
    "pause_confirmed":  False,
    "build_kwargs":     None,   # non-file kwargs for converting dry-run → real build
    "build_files":      None,   # raw bytes {"lokaal", "stad", "variants"} for re-run
    "kw_cpc_sheet":     None,   # KeywordCpcSheet instance (Schoorsteenveger)
    "kw_cpc_sheet_key": "",     # "{branch}|{url}|{tab}" — invalideer cache bij wijziging
}
for key, default in _defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ── resultaten-scherm ─────────────────────────────────────────────────────────
if st.session_state.results is not None:
    merged = st.session_state.results
    errors = st.session_state.errors
    log    = st.session_state.log

    if st.button("↩ Opnieuw beginnen"):
        for k in _defaults:
            st.session_state[k] = _defaults[k]
        st.rerun()

    # Dry-run resultaat
    if "__dry_run__" in merged:
        st.info("Dry-run voltooid — er zijn geen bestanden gegenereerd.")
        dry_rows = merged["__dry_run__"]
        if dry_rows:
            st.header("Bouwplan (dry-run)")
            st.dataframe(pd.DataFrame(dry_rows), use_container_width=True, hide_index=True)
        if errors:
            st.warning("**Meldingen:**\n" + "\n".join(errors))

        st.divider()
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("🚀 Nu echt bouwen", type="primary"):
                kwargs = st.session_state.build_kwargs
                files  = st.session_state.build_files
                if kwargs and files:
                    _pb  = st.progress(0)
                    _txt = st.empty()
                    def _real_progress_cb(city, i, total, skipped=False, stad_excluded=False):
                        _pb.progress(int((i / total) * 100))
                        _txt.text(f"Verwerken: {city} ({i+1}/{total})")
                    with tempfile.TemporaryDirectory() as tmp2:
                        lp = os.path.join(tmp2, "lokaal.csv")
                        sp = os.path.join(tmp2, "stad.csv")
                        with open(lp, "wb") as f: f.write(files["lokaal"])
                        with open(sp, "wb") as f: f.write(files["stad"])
                        vp = None
                        if files.get("variants"):
                            vp = os.path.join(tmp2, "variants.xlsx")
                            with open(vp, "wb") as f: f.write(files["variants"])
                        try:
                            real_merged, real_errors = build_all(
                                lp, sp, **kwargs,
                                variants_path=vp,
                                dry_run=False,
                                progress_cb=_real_progress_cb,
                            )
                        except Exception as exc:
                            st.error(f"Fout bij verwerken: {exc}")
                            st.stop()
                    _pb.progress(100)
                    _txt.text("Klaar!")
                    st.session_state.results = real_merged
                    st.session_state.errors  = real_errors
                    st.session_state.build_kwargs = None
                    st.session_state.build_files  = None
                    st.rerun()
                else:
                    st.error("Bouwparameters niet meer beschikbaar — begin opnieuw.")
        with col2:
            st.caption("Akkoord met het bouwplan? Klik om de bestanden te genereren.")
        st.stop()

    # Normale uitvoer
    st.success(
        f"{sum(len(v['lokaal']) for v in merged.values())} rijen gegenereerd "
        f"voor {len(merged)} klant(en)."
    )

    if errors:
        st.warning("**Meldingen:**\n" + "\n".join(errors))

    st.header("Downloads")
    for klant, dfs in sorted(merged.items()):
        st.subheader(f"Klant: {klant}")
        col_l, col_s = st.columns(2)
        with col_l:
            fname = output_filename(klant, "lokaal")
            st.download_button(f"⬇ {fname}", df_to_csv_bytes(dfs["lokaal"]),
                               file_name=fname, mime="text/csv", key=f"dl_lokaal_{klant}")
        with col_s:
            fname = output_filename(klant, "stad")
            st.download_button(f"⬇ {fname}", df_to_csv_bytes(dfs["stad"]),
                               file_name=fname, mime="text/csv", key=f"dl_stad_{klant}")

    st.header("Overzicht")
    rows = [{"Status": line} for line in log] + [{"Status": f"❌ {e}"} for e in errors]
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.stop()


# ── invoerformulier ───────────────────────────────────────────────────────────

st.header("1. Klantgegevens")
eigen_merk = selected_merk == "Eigen merk"

merk_info         = None
klant_prijs       = None
portaal_klantnaam = ""
if eigen_merk:
    st.subheader("Klantgegevens")
    if _toon_uitgebreide_merk_info:
        col1, col2 = st.columns(2)
        with col1:
            korte_naam     = st.text_input("Korte bedrijfsnaam", placeholder=_placeholder_korte)
            review_score   = st.text_input(_label_review, placeholder=_placeholder_review)
        with col2:
            lange_naam     = st.text_input("Lange bedrijfsnaam", placeholder=_placeholder_lange)
            jaren_garantie = st.text_input("Jaren garantie", placeholder="10")
        usp = st.text_input(
            "Overige USP (leeg = fallback uit variantenbestand)",
            placeholder="bijv. Al 30 jaar vakmanschap"
        )
        merk_info = {
            "korte_naam":     korte_naam.strip(),
            "lange_naam":     lange_naam.strip(),
            "review_score":   review_score.strip(),
            "jaren_garantie": jaren_garantie.strip(),
            "usp":            usp.strip(),
            "usp_fallback":   "100% Tevreden? Dan Wij Ook",
        }
    else:
        if selected_land == "BE":
            klantnaam = st.text_input("Klantnaam", placeholder="Schoorsteenvegers De Spil")
        else:
            col1, col2 = st.columns(2)
            with col1:
                klantnaam = st.text_input("Klantnaam", placeholder="Schoorsteenvegers De Spil")
            with col2:
                klant_prijs = st.text_input("Prijs (bijv. €59,50)", placeholder="€59,50") or None
        merk_info = {
            "korte_naam":     klantnaam.strip(),
            "lange_naam":     klantnaam.strip(),
            "review_score":   "",
            "jaren_garantie": "",
            "usp":            "",
            "usp_fallback":   "",
        }
else:
    if selected_land == "BE":
        portaal_klantnaam = st.text_input("Klantnaam campagne", placeholder="Multi-Reiniging Service")
    else:
        col1, col2 = st.columns(2)
        with col1:
            portaal_klantnaam = st.text_input("Klantnaam campagne", placeholder="Multi-Reiniging Service")
        with col2:
            klant_prijs = st.text_input("Prijs (bijv. €59,50)", placeholder="€59,50") or None

with st.expander("🔍 Debug: laadstatus bestanden", expanded=False):
    import tempfile as _tf, io as _io
    from builder import load_variants as _lv, _uses_explicit_placeholders as _uep, _apply_variant as _av, _apply_explicit_text as _aet
    _dbg_lok = _load_file_bytes(_branch_cfg.get("lokaal_template_path", ""))
    _dbg_sta = _load_file_bytes(_branch_cfg.get("stad_template_path", ""))
    _dbg_var = _load_file_bytes(_branch_cfg.get("variants_path", ""))
    st.write(f"**Lokaal sjabloon:** {'✓ geladen' if _dbg_lok else '✗ niet gevonden'} ({len(_dbg_lok) if _dbg_lok else 0} bytes)")
    st.write(f"**Stad sjabloon:** {'✓ geladen' if _dbg_sta else '✗ niet gevonden'} ({len(_dbg_sta) if _dbg_sta else 0} bytes)")
    st.write(f"**Varianten:** {'✓ geladen' if _dbg_var else '✗ niet gevonden'} ({len(_dbg_var) if _dbg_var else 0} bytes)")
    if _dbg_sta:
        import pandas as _pd
        _df_sta = _pd.read_csv(_io.BytesIO(_dbg_sta), sep=";", encoding="utf-8-sig", low_memory=False)
        st.write(f"**Explicit placeholders in stad-sjabloon:** {_uep(_df_sta)}")
    if _dbg_var:
        with _tf.NamedTemporaryFile(suffix=".xlsx", delete=False) as _f:
            _f.write(_dbg_var); _vpath = _f.name
        _vars = _lv(_vpath)
        st.write(f"**Aantal varianten geladen:** {len(_vars)}")
        _test_keys = [k for k in _vars if "bedrijfsnaam" in k and "in" in k]
        st.write(f"**Relevante keys:** {_test_keys}")
        if eigen_merk and merk_info and merk_info.get("korte_naam"):
            for _city in ["Alblasserdam", "Capelle aan den IJssel"]:
                _res = _av("[Klantnaam] In [Plaats]", _city, _vars, merk_info, col="Headline 2")
                _res2 = _aet(_res, _city, merk_info, None, col="Headline 2")
                st.write(f"[Klantnaam] In [Plaats] + {_city} → **{_res2}**")

st.header("2. Sjabloon-bestanden")
col1, col2 = st.columns(2)

# Auto-laden vanuit branch-configuratie als pad bestaat
# Bij "Laagste CPC" gebruik de laagste-paden als die bestaan, anders val terug op hoogste
if _cpc_positie == "Laagste CPC":
    _default_lokaal_bytes = (
        _load_file_bytes(_branch_cfg.get("lokaal_laagste_path", ""))
        or _load_file_bytes(_branch_cfg.get("lokaal_template_path", ""))
    )
    _default_stad_bytes = (
        _load_file_bytes(_branch_cfg.get("stad_laagste_path", ""))
        or _load_file_bytes(_branch_cfg.get("stad_template_path", ""))
    )
else:
    _default_lokaal_bytes = _load_file_bytes(_branch_cfg.get("lokaal_template_path", ""))
    _default_stad_bytes   = _load_file_bytes(_branch_cfg.get("stad_template_path", ""))

with col1:
    lokaal_file = st.file_uploader(
        "Lokaal sjabloon (.csv)", type="csv", key="lokaal",
        help="dak_lokaal.csv of eigen merk lokaal sjabloon"
    )
    if not lokaal_file and _default_lokaal_bytes:
        _lp = (_branch_cfg.get("lokaal_laagste_path") if _cpc_positie == "Laagste CPC" else None) or _branch_cfg.get("lokaal_template_path", "")
        st.success(f"✓ Geladen uit branche-standaard: **{os.path.basename(_lp)}**")
with col2:
    stad_file = st.file_uploader(
        "+ Stad sjabloon (.csv)", type="csv", key="stad",
        help="dak + stad.csv of eigen merk stad sjabloon"
    )
    if not stad_file and _default_stad_bytes:
        _sp = (_branch_cfg.get("stad_laagste_path") if _cpc_positie == "Laagste CPC" else None) or _branch_cfg.get("stad_template_path", "")
        st.success(f"✓ Geladen uit branche-standaard: **{os.path.basename(_sp)}**")

# Gebruik geüpload bestand of standaard uit branch-config
_lokaal_bytes = lokaal_file.getvalue() if lokaal_file else _default_lokaal_bytes
_stad_bytes   = stad_file.getvalue()   if stad_file   else _default_stad_bytes

variants_file = st.file_uploader(
    "Varianten plaatsnamen (optioneel .xlsx)", type="xlsx", key="variants"
)
_default_variants_bytes = _load_file_bytes(_branch_cfg.get("variants_path", ""))
if not variants_file and _default_variants_bytes:
    st.success(f"✓ Geladen uit branche-standaard: **{os.path.basename(_branch_cfg.get('variants_path', ''))}**")
_variants_bytes = variants_file.getvalue() if variants_file else _default_variants_bytes

# ── Adviesbestanden ───────────────────────────────────────────────────────────

if _toon_adviesbestanden:
    st.header("3. Adviesbestanden (optioneel)")
    st.caption(
        "Upload het CPC-adviesbestand en/of het pauzeringsbestand. "
        "De bestanden worden herkend op inhoud, niet op bestandsnaam."
    )
    if _toon_pauze:
        col_cpc, col_pause = st.columns(2)
    else:
        col_cpc = st.container()
    with col_cpc:
        cpc_file = st.file_uploader(
            "CPC-adviesbestand (.xlsx)", type="xlsx", key="cpc_file"
        )
    if _toon_pauze:
        with col_pause:
            pause_file = st.file_uploader(
                "Pauzeringen-/beëindigingenbestand (.xlsx)", type="xlsx", key="pause_file"
            )
    else:
        pause_file = None
else:
    cpc_file   = None
    pause_file = None


def _show_file_detection_error(file_name: str, file_type: AdviceFileType, wb_data: dict):
    if file_type == AdviceFileType.UNKNOWN:
        st.error(
            f"**{file_name}** is niet herkend. "
            f"Aanwezige tabbladen: {', '.join(sorted(wb_data.keys()))}."
        )
    elif file_type == AdviceFileType.AMBIGUOUS:
        st.warning(
            f"**{file_name}** voldoet aan zowel de CPC- als de pauzeringsstructuur. "
            f"Selecteer hieronder het juiste type."
        )


# CPC-bestand verwerken
_bc_cpc_path   = _branch_cfg.get("cpc_file_path", "")
_bc_cpc_bytes  = _load_file_bytes(_bc_cpc_path) if not cpc_file else None

if cpc_file or _bc_cpc_bytes:
    file_bytes = cpc_file.getvalue() if cpc_file else _bc_cpc_bytes
    file_name  = cpc_file.name       if cpc_file else os.path.basename(_bc_cpc_path)
    _from_default = not cpc_file and bool(_bc_cpc_bytes)

    file_type, wb_data, _ = detect_advice_file_type(file_bytes, file_name)

    if file_type == AdviceFileType.UNKNOWN:
        _show_file_detection_error(file_name, file_type, wb_data)
        st.session_state.cpc_import    = None
        st.session_state.cpc_confirmed = False

    elif file_type == AdviceFileType.AMBIGUOUS:
        _show_file_detection_error(file_name, file_type, wb_data)
        use_as = st.radio(
            f"Gebruik '{file_name}' als:",
            ["CPC-adviesbestand", "Pauzeringsbestand"],
            key="ambiguous_cpc_choice",
            horizontal=True,
        )
        if use_as == "CPC-adviesbestand":
            cpc_import = parse_cpc_advice(file_bytes, file_name)
            st.session_state.cpc_import = cpc_import
        else:
            pause_import = parse_pause_advice(file_bytes, file_name)
            st.session_state.pause_import = pause_import
        st.session_state.cpc_confirmed = True

    elif file_type == AdviceFileType.PAUSE_ADVICE:
        if not _from_default:
            st.warning(
                f"**{file_name}** is herkend als pauzeringsbestand, "
                f"niet als CPC-adviesbestand. Verplaats dit naar het pauzeringsveld."
            )
        st.session_state.cpc_import    = None
        st.session_state.cpc_confirmed = False

    elif file_type == AdviceFileType.CPC_ADGROUP_LOCAL:
        from advice_parser import parse_local_adgroup_cpc as _parse_local_cpc
        _local_lookup = _parse_local_cpc(wb_data)
        st.success(f"CPC-adviesbestand geladen: **{file_name}** ({len(_local_lookup):,} adv.groep-regels)")
        st.session_state.local_adgroup_cpc = _local_lookup
        st.session_state.cpc_import        = None
        st.session_state.cpc_confirmed     = True

    else:  # CPC_ADVICE
        cpc_import = parse_cpc_advice(file_bytes, file_name)
        summary = cpc_import.summary

        if _from_default:
            st.success(f"CPC-adviesbestand geladen uit branche-standaard: **{file_name}**")
            st.session_state.cpc_import    = cpc_import
            st.session_state.cpc_confirmed = True
        else:
            st.subheader(f"CPC-preview: {file_name}")
            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Accounts", summary["accounts"])
            col_b.metric("Standaard regels", summary["standaard_regels"])
            col_c.metric("Campagne-uitzond.", summary["campagne_uitzond"])
            col_d, col_e, col_f = st.columns(3)
            col_d.metric("Adv.groep-uitzond.", summary["adgroup_uitzond"])
            col_e.metric("Zoekw.-uitzond.", summary["zoekwoord_uitzond"])
            col_f.metric("Fouten", summary["fouten"],
                         delta_color="off" if summary["fouten"] == 0 else "inverse")

            if cpc_import.validation_errors:
                st.error("**Validatiefouten CPC-bestand:**\n" +
                         "\n".join(f"- {e}" for e in cpc_import.validation_errors))
                st.session_state.cpc_import    = None
                st.session_state.cpc_confirmed = False
            else:
                if cpc_import.validation_warnings:
                    st.warning("**Waarschuwingen:**\n" +
                               "\n".join(f"- {w}" for w in cpc_import.validation_warnings))
                if not st.session_state.cpc_confirmed:
                    if st.button("Bevestig CPC-adviesbestand", key="confirm_cpc"):
                        st.session_state.cpc_import    = cpc_import
                        st.session_state.cpc_confirmed = True
                        st.rerun()
                else:
                    st.success(f"CPC-adviesbestand bevestigd: {file_name}")
                    st.session_state.cpc_import = cpc_import
                    # Sla op als branche-standaard
                    if st.button(f"💾 Sla op als standaard voor '{selected_branch} — {selected_merk}'", key="save_cpc_default"):
                        saved = _save_branch_advice_file(selected_branch, "cpc", file_bytes, file_name)
                        _branches.setdefault(selected_branch, {}).setdefault(selected_merk, {})["cpc_file_path"] = saved
                        _save_branches(_branches)
                        st.success(f"Opgeslagen als standaard: {saved}")
else:
    st.session_state.cpc_import       = None
    st.session_state.local_adgroup_cpc = None
    st.session_state.cpc_confirmed    = False


# Pauzeringsbestand verwerken
_bc_pause_path  = _branch_cfg.get("pause_file_path", "")
_bc_pause_bytes = _load_file_bytes(_bc_pause_path) if not pause_file else None

if pause_file or _bc_pause_bytes:
    file_bytes = pause_file.getvalue() if pause_file else _bc_pause_bytes
    file_name  = pause_file.name       if pause_file else os.path.basename(_bc_pause_path)
    _from_default = not pause_file and bool(_bc_pause_bytes)

    file_type, wb_data, _ = detect_advice_file_type(file_bytes, file_name)

    if file_type == AdviceFileType.UNKNOWN:
        _show_file_detection_error(file_name, file_type, wb_data)
        st.session_state.pause_import    = None
        st.session_state.pause_confirmed = False

    elif file_type == AdviceFileType.AMBIGUOUS:
        _show_file_detection_error(file_name, file_type, wb_data)
        use_as = st.radio(
            f"Gebruik '{file_name}' als:",
            ["Pauzeringsbestand", "CPC-adviesbestand"],
            key="ambiguous_pause_choice",
            horizontal=True,
        )
        if use_as == "Pauzeringsbestand":
            pause_import = parse_pause_advice(file_bytes, file_name)
            st.session_state.pause_import = pause_import
        else:
            cpc_import = parse_cpc_advice(file_bytes, file_name)
            st.session_state.cpc_import = cpc_import
        st.session_state.pause_confirmed = True

    elif file_type == AdviceFileType.CPC_ADVICE:
        if not _from_default:
            st.warning(
                f"**{file_name}** is herkend als CPC-adviesbestand, "
                f"niet als pauzeringsbestand. Verplaats dit naar het CPC-veld."
            )
        st.session_state.pause_import    = None
        st.session_state.pause_confirmed = False

    else:  # PAUSE_ADVICE
        pause_import = parse_pause_advice(file_bytes, file_name)
        summary = pause_import.summary

        if _from_default:
            st.success(f"Pauzeringsbestand geladen uit branche-standaard: **{file_name}**")
            st.session_state.pause_import    = pause_import
            st.session_state.pause_confirmed = True
        else:
            st.subheader(f"Pauze-preview: {file_name}")
            col_a, col_b = st.columns(2)
            col_a.metric("Accounts", summary["accounts"])
            col_b.metric("Fouten", summary["fouten"],
                         delta_color="off" if summary["fouten"] == 0 else "inverse")
            col_c, col_d, col_e, col_f = st.columns(4)
            col_c.metric("+Stad gepauzeerd", summary["stad_gepauzeerd"])
            col_d.metric("+Stad niet bouwen", summary["stad_niet_bouwen"])
            col_e.metric("Lokaal gepauzeerd", summary["lokaal_gepauzeerd"])
            col_f.metric("Lokaal niet bouwen", summary["lokaal_niet_bouwen"])

            if pause_import.validation_errors:
                st.error("**Validatiefouten pauzeringsbestand:**\n" +
                         "\n".join(f"- {e}" for e in pause_import.validation_errors))
                st.session_state.pause_import    = None
                st.session_state.pause_confirmed = False
            else:
                if pause_import.validation_warnings:
                    st.warning("**Waarschuwingen:**\n" +
                               "\n".join(f"- {w}" for w in pause_import.validation_warnings))
                if not st.session_state.pause_confirmed:
                    if st.button("Bevestig pauzeringsbestand", key="confirm_pause"):
                        st.session_state.pause_import    = pause_import
                        st.session_state.pause_confirmed = True
                        st.rerun()
                else:
                    st.success(f"Pauzeringsbestand bevestigd: {file_name}")
                    st.session_state.pause_import = pause_import
                    # Sla op als branche-standaard
                    if st.button(f"💾 Sla op als standaard voor '{selected_branch} — {selected_merk}'", key="save_pause_default"):
                        saved = _save_branch_advice_file(selected_branch, "pause", file_bytes, file_name)
                        _branches.setdefault(selected_branch, {}).setdefault(selected_merk, {})["pause_file_path"] = saved
                        _save_branches(_branches)
                        st.success(f"Opgeslagen als standaard: {saved}")
else:
    st.session_state.pause_import    = None
    st.session_state.pause_confirmed = False


# ── Google Sheets ─────────────────────────────────────────────────────────────

st.header("4. Google Sheets")
sheet_url  = st.text_input(
    "URL van het Google Sheets (deelbaar via 'Iedereen met de link')",
    value=_branch_cfg.get("sheet_url", ""),
    placeholder="https://docs.google.com/spreadsheets/d/…/edit",
)
sheet_name = st.text_input("Tabblad naam", value=_branch_cfg.get("sheet_name", "Testtab"))

if sheet_url.strip():
    if st.button("🔄 Laad klanten uit sheet"):
        try:
            df_sheet = load_sheet(sheet_url, sheet_name)
            klanten_in_sheet = sorted(
                {str(k).strip() for k in df_sheet.get("Klant", pd.Series()).dropna()
                 if str(k).strip() not in ("", "nan")}
            )
            st.session_state.klanten      = klanten_in_sheet
            st.session_state.sheet_loaded = True
        except Exception as e:
            st.error(f"Kon sheet niet laden: {e}")

if st.session_state.sheet_loaded:
    st.header("5. Klant selecteren")
    geselecteerde_klanten = st.multiselect(
        "Voor welke klant(en) wil je campagnes genereren?",
        options=st.session_state.klanten,
        default=st.session_state.klanten,
    )
else:
    geselecteerde_klanten = None
    if sheet_url.strip():
        st.info("Klik op 'Laad klanten uit sheet' om klanten te selecteren.")

st.header("6. Ad Schedule")
ad_schedule_override = st.text_area(
    "Ad Schedule (leeg = waarde uit sheet wordt gebruikt)",
    placeholder="(Monday[08:00-17:00]);(Tuesday[08:00-17:00]);...",
    height=68,
)

# ── Fallback CPC ─────────────────────────────────────────────────────────────

def _parse_cpc_euro(s: str) -> int | None:
    try:
        return round(float(s.replace(",", ".")) * 100)
    except (ValueError, AttributeError):
        return None

if _toon_fallback_cpc:
    st.header("6b. Fallback CPC (optioneel)")
    st.caption("Wordt gebruikt wanneer een stad/klant geen CPC-regel heeft in het adviesbestand.")
    _fb_col1, _fb_col2 = st.columns(2)
    with _fb_col1:
        _fb_stad_str   = st.text_input("Fallback +Stad CPC (€)", value="6.52", key="fb_stad")
    with _fb_col2:
        _fb_lokaal_str = st.text_input("Fallback Lokaal CPC (€)", value="5.02", key="fb_lokaal")
    _fallback_stad_cents   = _parse_cpc_euro(_fb_stad_str)
    _fallback_lokaal_cents = _parse_cpc_euro(_fb_lokaal_str)
else:
    _fallback_stad_cents   = None
    _fallback_lokaal_cents = None

# ── Dry-run toggle ────────────────────────────────────────────────────────────

st.header("7. Genereer")
dry_run = st.checkbox(
    "Dry-run (toon bouwplan zonder bestanden te genereren)",
    value=False,
)

if st.button("🚀 Genereer campagnes", type="primary"):
    if not _lokaal_bytes or not _stad_bytes:
        st.error("Upload beide sjabloon-CSV bestanden (of stel standaardpaden in bij de branche).")
    elif not sheet_url.strip():
        st.error("Vul de Google Sheets URL in.")
    elif not geselecteerde_klanten:
        st.error("Selecteer minimaal één klant.")
    elif eigen_merk and not merk_info.get("korte_naam"):
        st.error("Vul minimaal de korte bedrijfsnaam in.")
    else:
        # Blokkeer bij niet-bevestigde of ongeldige adviesbestanden
        if cpc_file and not st.session_state.cpc_confirmed:
            st.error("Bevestig het CPC-adviesbestand of verwijder het eerst.")
            st.stop()
        if pause_file and not st.session_state.pause_confirmed:
            st.error("Bevestig het pauzeringsbestand of verwijder het eerst.")
            st.stop()

        with tempfile.TemporaryDirectory() as tmp:
            lok_path = os.path.join(tmp, "lokaal.csv")
            sta_path = os.path.join(tmp, "stad.csv")
            with open(lok_path, "wb") as f:
                f.write(_lokaal_bytes)
            with open(sta_path, "wb") as f:
                f.write(_stad_bytes)

            var_path = None
            if _variants_bytes:
                var_path = os.path.join(tmp, "variants.xlsx")
                with open(var_path, "wb") as f:
                    f.write(_variants_bytes)

            progress_bar = st.progress(0)
            status_text  = st.empty()
            log_lines: list[str] = []

            def progress_cb(city, i, total, skipped=False, stad_excluded=False):
                progress_bar.progress(int((i / total) * 100))
                status_text.text(f"Verwerken: {city} ({i+1}/{total})")
                if skipped:
                    log_lines.append(f"⏭ {city} (overgeslagen — dubbele plaatsnaam)")
                elif stad_excluded:
                    log_lines.append(f"✅ {city} (alleen lokaal — geen +Stad campagne)")
                else:
                    log_lines.append(f"✅ {city}")

            _shared_kwargs = dict(
                sheet_url              = sheet_url,
                sheet_name             = sheet_name,
                klant_filter           = geselecteerde_klanten,
                merk_info              = merk_info if eigen_merk else None,
                portaal_klantnaam      = portaal_klantnaam.strip() if not eigen_merk else None,
                ad_schedule_override   = ad_schedule_override,
                cpc_import             = st.session_state.cpc_import,
                pause_import           = st.session_state.pause_import,
                fallback_stad_cents    = _fallback_stad_cents,
                fallback_lokaal_cents  = _fallback_lokaal_cents,
                template_city          = _branch_cfg.get("template_city") or None,
                template_company       = _branch_cfg.get("template_company") or None,
                template_price         = klant_prijs,
                keyword_cpc_sheet      = _active_kw_cpc_sheet,
                local_adgroup_cpc      = st.session_state.get("local_adgroup_cpc"),
                land                   = selected_land,
                merktype_cpc           = _merktype_cpc,
                is_portaal             = not eigen_merk,
            )
            if dry_run:
                st.session_state.build_kwargs = _shared_kwargs
                st.session_state.build_files  = {
                    "lokaal":   _lokaal_bytes,
                    "stad":     _stad_bytes,
                    "variants": _variants_bytes,
                }
            try:
                merged, errors = build_all(
                    lok_path, sta_path,
                    **_shared_kwargs,
                    variants_path = var_path,
                    dry_run       = dry_run,
                    progress_cb   = progress_cb,
                )
            except Exception as exc:
                st.error(f"Fout bij verwerken: {exc}")
                st.stop()

            progress_bar.progress(100)
            status_text.text("Klaar!")

        st.session_state.results = merged
        st.session_state.errors  = errors
        st.session_state.log     = log_lines
        st.rerun()
