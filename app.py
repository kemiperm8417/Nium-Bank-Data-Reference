"""
Bank Reference Data Explorer — Streamlit front end.

Pick countries, pull bank and branch reference data from Nium's public
Reference Data API, and download one Excel workbook with a sheet per country.

Run:
    streamlit run app.py
"""

import streamlit as st

# ── must be first Streamlit call ──────────────────────────────────────────────
st.set_page_config(
    page_title="Bank Reference Data",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

import nium_theme                                            # noqa: E402
nium_theme.apply()

from refdata import (                                        # noqa: E402
    BASE_URL,
    COMMON_CORRIDORS,
    SEPA_COUNTRIES,
    RefDataError,
    build_workbook,
    choose_mode,
    fetch_country,
    list_countries,
    suggested_filename,
)

import refdata as _refdata                                   # noqa: E402

_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# In the browser build (GitHub Pages / stlite) allow `?api=<base url>` so the
# page can be pointed at a dev CORS proxy for testing. Ignored natively.
if _refdata.IN_BROWSER:
    _override = st.query_params.get("api")
    if _override:
        _refdata.BASE_URL = _override.rstrip("/")


def _init_state():
    defaults = {
        "refdata_results":   None,   # list[CountryResult] from the last fetch
        "refdata_xlsx":      None,   # bytes — must survive the download rerun
        "refdata_filename":  None,
        "refdata_selection": [],     # ISO codes bound to the multiselect
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()


@st.cache_data(ttl=86_400, show_spinner="Loading country list…")
def _cached_countries():
    """ISO country list — effectively static, so cached for a day."""
    return list_countries()


@st.cache_data(ttl=6 * 3600, max_entries=60, show_spinner=False)
def _cached_fetch(code: str):
    """One country's normalised rows, cached so re-runs of an overlapping
    selection are instant. Progress is not reported on a cache hit."""
    return fetch_country(code)


def _estimate(codes):
    """Rough, honest cost hint so a multi-minute run is never a surprise."""
    notes, slow = [], False
    for c in codes:
        try:
            mode = choose_mode(c)
        except RefDataError:
            continue
        if mode == "branch_chunked":
            notes.append("**%s** is fetched bank-by-bank (hundreds of requests)" % c)
            slow = True
    if slow:
        return "⚠️  " + "; ".join(notes) + " — expect several minutes."
    if len(codes) > 20:
        return "%d countries selected — expect 1–2 minutes." % len(codes)
    return "%d country(ies) selected — this should take well under a minute." % len(codes)


def main():
    st.title("Bank Reference Data")
    st.markdown(
        '<p class="lede">Bank and branch reference data from Nium\'s Reference Data '
        'API — ACH/ABA for the US, IFSC for India, Sort Code for the UK, BSB for '
        'Australia, bank + branch codes for Japan and Canada, BIC/SWIFT across SEPA. '
        'One Excel sheet per country.</p>',
        unsafe_allow_html=True,
    )

    try:
        countries = _cached_countries()
    except RefDataError as exc:
        st.error("Could not reach the Reference Data API at %s — %s" % (BASE_URL, exc))
        if _refdata.IN_BROWSER:
            st.warning(
                "This page runs in your browser, so the request went from `%s` "
                "straight to the API — and the browser blocked it.\n\n"
                "**Self-check:** open [this API link](%s/entity/Country?limit=1) in a new tab. "
                "If you see JSON, you are on the VPN and the only problem is **CORS**: "
                "the Reference Data team must add this page's origin to the API's "
                "`allowedOrigins`. If the link does not load, connect to the Nium VPN first."
                % (st.context.url.split("?")[0] if hasattr(st, "context") else "this page",
                   BASE_URL)
            )
            return
        st.warning(
            "This machine cannot reach `refdata.prod.nium.com`. If you are running "
            "on GitHub Codespaces or another cloud host, the API is most likely "
            "restricted to Nium's network — run the app from a machine on the Nium "
            "network/VPN instead. Falling back to built-in country codes so you can "
            "still see the UI, but fetches will fail until the API is reachable."
        )
        countries = [{"code": c, "name": c, "banned": False}
                     for c in sorted(set(SEPA_COUNTRIES) | set(COMMON_CORRIDORS))]

    names = {c["code"]: c["name"] for c in countries}
    options = [c["code"] for c in countries]

    # ── presets ──────────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns([1, 1, 1])
    if c1.button("SEPA (36)", use_container_width=True):
        st.session_state["refdata_selection"] = [c for c in SEPA_COUNTRIES if c in options]
        st.rerun()
    if c2.button("Common corridors", use_container_width=True):
        st.session_state["refdata_selection"] = [c for c in COMMON_CORRIDORS if c in options]
        st.rerun()
    if c3.button("Clear", use_container_width=True):
        st.session_state["refdata_selection"] = []
        st.rerun()

    selected = st.multiselect(
        "Countries",
        options=options,
        format_func=lambda c: "%s — %s" % (c, names.get(c, c)),
        key="refdata_selection",
        help="Pick any number of countries. Each becomes its own sheet in the workbook.",
    )
    st.caption(
        "SEPA membership has no API endpoint — the 36-member list is a maintained "
        "constant in `refdata.py`. Note the UK is a SEPA member but is exported with "
        "real sort codes, not BICs."
    )

    if selected:
        st.info(_estimate(selected))

    force = st.checkbox("Force refresh (bypass the 6-hour cache)", value=False)

    # ── fetch ────────────────────────────────────────────────────────────────
    if st.button("Fetch reference data", type="primary", disabled=not selected):
        if force:
            _cached_fetch.clear()

        results = []
        with st.status("Fetching %d countries…" % len(selected), expanded=True) as status:
            bar = st.progress(0.0, text="Starting…")
            for i, code in enumerate(selected):
                bar.progress(i / len(selected), text="%s — %s" % (code, names.get(code, code)))
                res = _cached_fetch(code)
                results.append(res)
                if res.error:
                    st.markdown('<span class="pill negative">Failed</span>&nbsp; %s — %s'
                                % (code, res.error), unsafe_allow_html=True)
                else:
                    st.markdown(
                        '<span class="pill positive">OK</span>&nbsp; %s — '
                        '<span class="num">%s</span> rows via %s in '
                        '<span class="num">%.1fs</span>%s'
                        % (code, format(len(res.rows), ","), res.mode, res.seconds,
                           " (%s)" % res.notes if res.notes else ""),
                        unsafe_allow_html=True)
            bar.progress(1.0, text="Building workbook…")
            status.update(label="Building workbook…")

            try:
                xlsx = build_workbook(results)
            except Exception as exc:                        # noqa: BLE001
                status.update(label="Workbook build failed", state="error")
                st.error("Could not build the workbook: %s" % exc)
                return

            total = sum(len(r.rows) for r in results)
            status.update(label="Done — %s rows" % format(total, ","),
                          state="complete", expanded=False)

        st.session_state["refdata_results"] = results
        st.session_state["refdata_xlsx"] = xlsx
        st.session_state["refdata_filename"] = suggested_filename(selected)

    # ── results (rendered from session state on every rerun) ─────────────────
    results = st.session_state["refdata_results"]
    if not results:
        return

    import pandas as pd

    st.subheader("Summary")
    st.dataframe(
        pd.DataFrame([{
            "Country": r.code,
            "Name": r.name,
            "Rows": len(r.rows),
            "Routing Code": next((x["Routing Code Type"] for x in r.rows[:50]
                                  if x.get("Routing Code Type")), ""),
            "Endpoint": r.mode,
            "Seconds": round(r.seconds, 1),
            "Status": "ERROR" if r.error else ("EMPTY" if not r.rows else "OK"),
            "Notes": r.error or r.notes,
        } for r in results]),
        use_container_width=True,
        hide_index=True,
    )

    failed = [r for r in results if r.error]
    if failed:
        st.warning(
            "%d country(ies) failed and are recorded as errors in the Summary sheet — "
            "the workbook still contains everything that succeeded." % len(failed)
        )

    ok = [r for r in results if r.rows]
    if ok:
        st.subheader("Preview")
        pick = st.selectbox(
            "Country", [r.code for r in ok],
            format_func=lambda c: "%s — %s" % (c, names.get(c, c)),
        )
        rows = next(r.rows for r in ok if r.code == pick)
        # Never build a DataFrame from a full country — India would lock the browser.
        st.caption("First 200 of %s rows." % format(len(rows), ","))
        st.dataframe(pd.DataFrame(rows[:200]), use_container_width=True, hide_index=True)

    if st.session_state["refdata_xlsx"]:
        if _refdata.IN_BROWSER:
            # stlite (GitHub Pages) has no media server for st.download_button,
            # so hand the bytes to the browser directly as a data: URL.
            import base64
            b64 = base64.b64encode(st.session_state["refdata_xlsx"]).decode("ascii")
            size_mb = len(st.session_state["refdata_xlsx"]) / 1e6
            st.markdown(
                '<a class="nium-btn" download="%s" href="data:%s;base64,%s">'
                'Download Excel workbook <span class="num">(%.1f MB)</span></a>'
                % (st.session_state["refdata_filename"], _MIME_XLSX, b64, size_mb),
                unsafe_allow_html=True,
            )
        else:
            st.download_button(
                "Download Excel workbook",
                data=st.session_state["refdata_xlsx"],
                file_name=st.session_state["refdata_filename"],
                mime=_MIME_XLSX,
                type="primary",
            )
        st.caption(
            "The workbook is held in this session, so downloading does not re-fetch. "
            "Refreshing the page during a fetch loses the run."
        )


if __name__ == "__main__":
    main()
