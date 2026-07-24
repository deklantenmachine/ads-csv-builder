"""Streamlit interface voor de Google Ads CSV builder."""

import io
import os
import tempfile

import streamlit as st
import pandas as pd

from builder import build_all, df_to_csv_bytes, output_filename, load_sheet
from advice_parser import (
    parse_cpc_advice, parse_pause_advice,
    detect_advice_file_type, AdviceFileType,
)

st.set_page_config(page_title="Ads CSV Builder", page_icon="📦", layout="centered")
st.title("📦 Google Ads CSV Builder")

# ── session state ─────────────────────────────────────────────────────────────
_defaults = {
    "results":       None,
    "errors":        [],
    "log":           [],
    "klanten":       [],
    "sheet_loaded":  False,
    "cpc_import":    None,
    "pause_import":  None,
    "cpc_confirmed": False,
    "pause_confirmed": False,
    "build_kwargs":  None,   # non-file kwargs for converting dry-run → real build
    "build_files":   None,   # raw bytes {"lokaal", "stad", "variants"} for re-run
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

st.header("1. Merktype")
merk_type = st.radio(
    "Kies het type account",
    ["Overkoepelend merk (portaal)", "Eigen merk"],
    horizontal=True,
)
eigen_merk = merk_type == "Eigen merk"

merk_info = None
if eigen_merk:
    st.subheader("Klantgegevens")
    col1, col2 = st.columns(2)
    with col1:
        korte_naam     = st.text_input("Korte bedrijfsnaam", placeholder="DCN")
        review_score   = st.text_input("Review score", placeholder="9.2")
    with col2:
        lange_naam     = st.text_input("Lange bedrijfsnaam", placeholder="DCN Dakdekkers")
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

st.header("2. Sjabloon-bestanden")
col1, col2 = st.columns(2)
with col1:
    lokaal_file = st.file_uploader(
        "Lokaal sjabloon (.csv)", type="csv", key="lokaal",
        help="dak_lokaal.csv of eigen merk lokaal sjabloon"
    )
with col2:
    stad_file = st.file_uploader(
        "+ Stad sjabloon (.csv)", type="csv", key="stad",
        help="dak + stad.csv of eigen merk stad sjabloon"
    )

variants_file = st.file_uploader(
    "Varianten plaatsnamen (optioneel .xlsx)", type="xlsx", key="variants"
)

# ── Adviesbestanden ───────────────────────────────────────────────────────────

st.header("3. Adviesbestanden (optioneel)")
st.caption(
    "Upload het CPC-adviesbestand en/of het pauzeringsbestand. "
    "De bestanden worden herkend op inhoud, niet op bestandsnaam."
)

col_cpc, col_pause = st.columns(2)

with col_cpc:
    cpc_file = st.file_uploader(
        "CPC-adviesbestand (.xlsx)", type="xlsx", key="cpc_file"
    )

with col_pause:
    pause_file = st.file_uploader(
        "Pauzeringen-/beëindigingenbestand (.xlsx)", type="xlsx", key="pause_file"
    )


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
if cpc_file:
    file_bytes = cpc_file.getvalue()
    file_name  = cpc_file.name
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

    else:
        # CPC_ADVICE of PAUSE_ADVICE
        if file_type == AdviceFileType.PAUSE_ADVICE:
            st.warning(
                f"**{file_name}** is herkend als pauzeringsbestand, "
                f"niet als CPC-adviesbestand. Verplaats dit naar het pauzeringsveld."
            )
            st.session_state.cpc_import    = None
            st.session_state.cpc_confirmed = False
        else:
            cpc_import = parse_cpc_advice(file_bytes, file_name)

            # Preview
            summary = cpc_import.summary
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
else:
    st.session_state.cpc_import    = None
    st.session_state.cpc_confirmed = False


# Pauzeringsbestand verwerken
if pause_file:
    file_bytes = pause_file.getvalue()
    file_name  = pause_file.name
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

    else:
        if file_type == AdviceFileType.CPC_ADVICE:
            st.warning(
                f"**{file_name}** is herkend als CPC-adviesbestand, "
                f"niet als pauzeringsbestand. Verplaats dit naar het CPC-veld."
            )
            st.session_state.pause_import    = None
            st.session_state.pause_confirmed = False
        else:
            pause_import = parse_pause_advice(file_bytes, file_name)

            summary = pause_import.summary
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
else:
    st.session_state.pause_import    = None
    st.session_state.pause_confirmed = False


# ── Google Sheets ─────────────────────────────────────────────────────────────

st.header("4. Google Sheets")
sheet_url  = st.text_input(
    "URL van het Google Sheets (deelbaar via 'Iedereen met de link')",
    placeholder="https://docs.google.com/spreadsheets/d/…/edit",
)
sheet_name = st.text_input("Tabblad naam", value="DakPro NL Plaatsen")

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

# ── Dry-run toggle ────────────────────────────────────────────────────────────

st.header("7. Genereer")
dry_run = st.checkbox(
    "Dry-run (toon bouwplan zonder bestanden te genereren)",
    value=False,
)

if st.button("🚀 Genereer campagnes", type="primary"):
    if not lokaal_file or not stad_file:
        st.error("Upload beide sjabloon-CSV bestanden.")
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
                f.write(lokaal_file.getvalue())
            with open(sta_path, "wb") as f:
                f.write(stad_file.getvalue())

            var_path = None
            if variants_file:
                var_path = os.path.join(tmp, "variants.xlsx")
                with open(var_path, "wb") as f:
                    f.write(variants_file.getvalue())

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
                sheet_url            = sheet_url,
                sheet_name           = sheet_name,
                klant_filter         = geselecteerde_klanten,
                merk_info            = merk_info if eigen_merk else None,
                ad_schedule_override = ad_schedule_override,
                cpc_import           = st.session_state.cpc_import,
                pause_import         = st.session_state.pause_import,
            )
            if dry_run:
                st.session_state.build_kwargs = _shared_kwargs
                st.session_state.build_files  = {
                    "lokaal":   lokaal_file.getvalue(),
                    "stad":     stad_file.getvalue(),
                    "variants": variants_file.getvalue() if variants_file else None,
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
