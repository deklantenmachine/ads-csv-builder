"""Streamlit interface voor de Google Ads CSV builder."""

import os
import tempfile

import streamlit as st
import pandas as pd

from builder import build_all, df_to_csv_bytes, output_filename, load_sheet

st.set_page_config(page_title="Ads CSV Builder", page_icon="📦", layout="centered")
st.title("📦 Google Ads CSV Builder")

# ── session state ─────────────────────────────────────────────────────────────
for key, default in [("results", None), ("errors", []), ("log", []),
                     ("klanten", []), ("sheet_loaded", False)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── resultaten-scherm ─────────────────────────────────────────────────────────
if st.session_state.results is not None:
    if st.button("↩ Opnieuw beginnen"):
        for key in ("results", "errors", "log", "klanten", "sheet_loaded"):
            st.session_state[key] = [] if key in ("errors", "log", "klanten") else (None if key == "results" else False)
        st.rerun()

    merged = st.session_state.results
    errors = st.session_state.errors
    log    = st.session_state.log

    st.success(f"{sum(len(v['lokaal']) for v in merged.values())} rijen gegenereerd voor {len(merged)} klant(en).")

    if errors:
        st.warning("**Fouten:**\n" + "\n".join(errors))

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

st.header("1. Sjabloon-bestanden")
col1, col2 = st.columns(2)
with col1:
    lokaal_file = st.file_uploader("dak_lokaal.csv", type="csv", key="lokaal")
with col2:
    stad_file = st.file_uploader("dak + stad.csv", type="csv", key="stad")

variants_file = st.file_uploader("Varianten plaatsnamen (optioneel .xlsx)", type="xlsx", key="variants")

st.header("2. Google Sheets")
sheet_url  = st.text_input(
    "URL van het Google Sheets (deelbaar via 'Iedereen met de link')",
    placeholder="https://docs.google.com/spreadsheets/d/…/edit",
)
sheet_name = st.text_input("Tabblad naam", value="DakPro NL Plaatsen")

# ── klanten laden ─────────────────────────────────────────────────────────────
if sheet_url.strip():
    if st.button("🔄 Laad klanten uit sheet"):
        try:
            df_sheet = load_sheet(sheet_url, sheet_name)
            klanten_in_sheet = sorted(
                {str(k).strip() for k in df_sheet["Klant"].dropna()
                 if str(k).strip() not in ("", "nan")}
            )
            st.session_state.klanten      = klanten_in_sheet
            st.session_state.sheet_loaded = True
        except Exception as e:
            st.error(f"Kon sheet niet laden: {e}")

if st.session_state.sheet_loaded:
    st.header("3. Klant selecteren")
    geselecteerde_klanten = st.multiselect(
        "Voor welke klant(en) wil je campagnes genereren?",
        options=st.session_state.klanten,
        default=st.session_state.klanten,
    )
else:
    geselecteerde_klanten = None
    if sheet_url.strip():
        st.info("Klik op 'Laad klanten uit sheet' om klanten te selecteren.")

st.header("4. Genereer")

if st.button("🚀 Genereer campagnes", type="primary"):
    if not lokaal_file or not stad_file:
        st.error("Upload beide sjabloon-CSV bestanden.")
    elif not sheet_url.strip():
        st.error("Vul de Google Sheets URL in.")
    elif not geselecteerde_klanten:
        st.error("Selecteer minimaal één klant.")
    else:
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

            def progress_cb(city, i, total, skipped=False):
                progress_bar.progress(int((i / total) * 100))
                status_text.text(f"Verwerken: {city} ({i+1}/{total})")
                if skipped:
                    log_lines.append(f"⏭ {city} (overgeslagen — dubbele plaatsnaam)")
                else:
                    log_lines.append(f"✅ {city}")

            try:
                merged, errors = build_all(
                    lok_path, sta_path, sheet_url,
                    sheet_name=sheet_name,
                    variants_path=var_path,
                    klant_filter=geselecteerde_klanten,
                    progress_cb=progress_cb,
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
