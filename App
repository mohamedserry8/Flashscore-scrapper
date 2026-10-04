"""
Flashscore ↔ Internal DB Comparator
Streamlit app that compares internal match fixtures (CSV export from DBeaver)
against Flashscore feed data (CSV export from the Tampermonkey script).
"""

import io
import re
import unicodedata
from datetime import datetime, timedelta, timezone

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from rapidfuzz import fuzz

# ─── Config ───────────────────────────────────────────────────────────────────
SHEET_ID        = "14tUgMxJI_glunJiyg8F61D5oduorWJRBRzHfT8zGRtU"
MAPPING_TAB     = "fs_mapping"          # competition_id → flashscore_tournament_id
TEAM_MAP_TAB    = "team mapping"        # manual name overrides (shared with SofaScore tool)
LOG_TAB_MATCHES = "fs_match_log"        # per-session match-level log
LOG_TAB_PLAYERS = "fs_session_log"      # per-session summary (optional)

CAIRO_TZ_OFFSET = 3                     # UTC+3
FUZZY_THRESHOLD = 80
KICK_OFF_TOLERANCE_MIN = 30             # minutes; within this → time match

# ─── Abbreviation expansion (shared with SofaScore tool) ─────────────────────
ABBR_MAP = {
    r"\bHFX\b": "Halifax",
    r"\bFC\b": "",
    r"\bSC\b": "",
    r"\bAFC\b": "",
    r"\bCFC\b": "",
    r"\bWFC\b": "",
    r"\bFSV\b": "",
    r"\bSV\b": "",
    r"\bSK\b": "",
    r"\bFK\b": "",
    r"\bBK\b": "",
    r"\bIK\b": "",
    r"\bGIF\b": "",
    r"\bIF\b": "",
    r"\bAC\b": "",
    r"\bAS\b": "",
    r"\bSSC\b": "",
    r"\bUS\b": "",
    r"\bUD\b": "",
    r"\bRCD\b": "",
    r"\bRC\b": "",
    r"\bCD\b": "",
    r"\bSD\b": "",
    r"\bAD\b": "",
    r"\bRFC\b": "",
    r"\bPFC\b": "",
    r"\bMFK\b": "",
    r"\bNK\b": "",
    r"\bHNK\b": "",
    r"\bGNK\b": "",
    r"\bDFK\b": "",
    r"\bOFK\b": "",
    r"\bJK\b": "",
    r"\bTJ\b": "",
    r"\bSFC\b": "",
    r"\bBFC\b": "",
    r"\bFBC\b": "",
    r"\bMSK\b": "",
    r"\bFSC\b": "",
    r"\bKSC\b": "",
    r"\bBSC\b": "",
    r"\bCSC\b": "",
    r"\bVfL\b": "",
    r"\bVfB\b": "",
    r"\bTSG\b": "",
    r"\bRB\b": "",
    r"\bSC\b": "",
    r"\b\d{4}\b": "",   # remove years like 1903
}

def normalize_name(name: str, team_map: dict | None = None) -> str:
    """Normalize a team name for fuzzy matching."""
    if not isinstance(name, str):
        return ""
    # Apply manual team map first
    if team_map and name in team_map:
        name = team_map[name]
    # Strip accents
    name = unicodedata.normalize("NFD", name)
    name = "".join(c for c in name if unicodedata.category(c) != "Mn")
    # Apply abbreviation expansions
    for pattern, replacement in ABBR_MAP.items():
        name = re.sub(pattern, replacement, name, flags=re.IGNORECASE)
    # Collapse whitespace, lowercase
    name = re.sub(r"\s+", " ", name).strip().lower()
    return name


def fuzzy_score(a: str, b: str) -> float:
    """Return best of token_set_ratio and partial_ratio."""
    return max(
        fuzz.token_set_ratio(a, b),
        fuzz.partial_ratio(a, b),
    )


# ─── Google Sheets helpers ────────────────────────────────────────────────────
@st.cache_resource
def get_gc():
    creds_dict = st.secrets["gcp_service_account"]
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


@st.cache_data(ttl=300)
def load_mapping_sheet():
    """
    Load fs_mapping tab.
    Expected columns: competition_id, flashscore_tournament_id (comma-separated if multiple)
    Returns: dict {competition_id (int) → list[str]}
    """
    gc = get_gc()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(MAPPING_TAB)
        df = pd.DataFrame(ws.get_all_records())
    except Exception:
        return {}
    if df.empty:
        return {}
    mapping = {}
    for _, row in df.iterrows():
        cid = row.get("competition_id")
        fids = str(row.get("flashscore_tournament_id", "")).strip()
        if cid and fids:
            mapping[int(cid)] = [x.strip() for x in fids.split(",") if x.strip()]
    return mapping


@st.cache_data(ttl=300)
def load_team_map():
    """Load manual team name overrides from 'team mapping' tab."""
    gc = get_gc()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(TEAM_MAP_TAB)
        df = pd.DataFrame(ws.get_all_records())
    except Exception:
        return {}
    if df.empty:
        return {}
    # Expected: internal_name, external_name
    team_map = {}
    for _, row in df.iterrows():
        src = str(row.get("internal_name", "")).strip()
        tgt = str(row.get("external_name", "")).strip()
        if src and tgt:
            team_map[src] = tgt
    return team_map


def append_to_log(rows: list[dict]):
    """Append match-level comparison rows to fs_match_log sheet."""
    if not rows:
        return
    gc = get_gc()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(LOG_TAB_MATCHES)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOG_TAB_MATCHES, rows=1, cols=len(rows[0]))
        ws.append_row(list(rows[0].keys()))
    values = [[str(v) for v in r.values()] for r in rows]
    ws.append_rows(values)


# ─── Core matching logic ──────────────────────────────────────────────────────
def match_fixtures(
    db_df: pd.DataFrame,
    fs_df: pd.DataFrame,
    comp_to_fs_ids: dict,
    team_map: dict,
    threshold: int = FUZZY_THRESHOLD,
    kick_off_tol_min: int = KICK_OFF_TOLERANCE_MIN,
) -> pd.DataFrame:
    """
    For each internal DB match, try to find its Flashscore counterpart.

    Matching strategy (in order):
      1. Filter Flashscore rows to tournament_id(s) mapped to this competition_id
      2. Filter Flashscore rows to same match_date
      3. Fuzzy match home+away (try both directions, take best)
      4. Accept if score >= threshold

    Returns a DataFrame with one row per DB match + match result columns.
    """
    results = []

    # Pre-normalise FS team names
    fs_df = fs_df.copy()
    fs_df["_home_norm"] = fs_df["home_team"].apply(lambda x: normalize_name(x, team_map))
    fs_df["_away_norm"] = fs_df["away_team"].apply(lambda x: normalize_name(x, team_map))

    for _, db_row in db_df.iterrows():
        comp_id   = int(db_row.get("competition_id", -1))
        db_home   = str(db_row.get("home_team", ""))
        db_away   = str(db_row.get("away_team", ""))
        db_date   = str(db_row.get("match_date", ""))[:10]
        db_kickoff = str(db_row.get("kick_off_time", ""))[:5]  # HH:MM

        # Get allowed flashscore tournament ids for this competition
        allowed_fs_ids = comp_to_fs_ids.get(comp_id, [])
        if not allowed_fs_ids:
            results.append(_no_match_row(db_row, reason="No FS mapping"))
            continue

        # Filter FS pool: tournament_id AND date
        # ⚠️ Per-row scoping: only use tournament IDs linked to THIS competition
        pool = fs_df[
            fs_df["tournament_id"].isin(allowed_fs_ids) &
            (fs_df["match_date"] == db_date)
        ]

        if pool.empty:
            results.append(_no_match_row(db_row, reason="No FS matches on this date/tournament"))
            continue

        db_home_norm = normalize_name(db_home, team_map)
        db_away_norm = normalize_name(db_away, team_map)

        best_score  = -1
        best_row    = None
        best_reversed = False

        for _, fs_row in pool.iterrows():
            fs_h = fs_row["_home_norm"]
            fs_a = fs_row["_away_norm"]

            # Normal direction
            score_fwd = (fuzzy_score(db_home_norm, fs_h) + fuzzy_score(db_away_norm, fs_a)) / 2
            # Reversed direction
            score_rev = (fuzzy_score(db_home_norm, fs_a) + fuzzy_score(db_away_norm, fs_h)) / 2

            score = score_fwd
            rev   = False
            if score_rev > score_fwd:
                score = score_rev
                rev   = True

            if score > best_score:
                best_score = score
                best_row   = fs_row
                best_reversed = rev

        if best_score < threshold:
            results.append(_no_match_row(db_row, reason=f"Best fuzzy={best_score:.0f}% < threshold"))
            continue

        # ─── Kick-off time check ──────────────────────────────────────────────
        fs_time   = str(best_row.get("kick_off_time", ""))[:5]
        time_diff = _time_diff_minutes(db_kickoff, fs_time)
        time_ok   = time_diff is not None and abs(time_diff) <= kick_off_tol_min
        time_note = (
            f"Match (+{time_diff} min)" if time_diff and time_diff > 0
            else f"Match ({time_diff} min)" if time_diff is not None
            else "N/A"
        )

        results.append({
            # DB fields
            "comp_id":         comp_id,
            "db_match_id":     db_row.get("match_id", ""),
            "db_home":         db_home,
            "db_away":         db_away,
            "db_date":         db_date,
            "db_kickoff":      db_kickoff,
            # FS fields
            "fs_match_id":     best_row.get("flashscore_id", ""),
            "fs_home":         best_row.get("home_team", ""),
            "fs_away":         best_row.get("away_team", ""),
            "fs_tournament":   best_row.get("tournament", ""),
            "fs_tournament_id":best_row.get("tournament_id", ""),
            "fs_kickoff":      fs_time,
            "fs_status":       best_row.get("status", ""),
            # Match quality
            "fuzzy_score":     round(best_score, 1),
            "teams_reversed":  best_reversed,
            "kickoff_diff_min":time_diff,
            "kickoff_ok":      time_ok,
            "kickoff_note":    time_note,
            "match_result":    "✅ Found" if time_ok else "⚠️ Time mismatch",
        })

    return pd.DataFrame(results)


def find_missing_in_db(db_df: pd.DataFrame, fs_df: pd.DataFrame, matched_fs_ids: set) -> pd.DataFrame:
    """Return FS matches that have no DB counterpart (potential missing fixtures)."""
    unmatched = fs_df[~fs_df["flashscore_id"].isin(matched_fs_ids)].copy()
    return unmatched


def _no_match_row(db_row, reason=""):
    return {
        "comp_id":          db_row.get("competition_id", ""),
        "db_match_id":      db_row.get("match_id", ""),
        "db_home":          db_row.get("home_team", ""),
        "db_away":          db_row.get("away_team", ""),
        "db_date":          str(db_row.get("match_date", ""))[:10],
        "db_kickoff":       str(db_row.get("kick_off_time", ""))[:5],
        "fs_match_id":      "",
        "fs_home":          "",
        "fs_away":          "",
        "fs_tournament":    "",
        "fs_tournament_id": "",
        "fs_kickoff":       "",
        "fs_status":        "",
        "fuzzy_score":      0,
        "teams_reversed":   False,
        "kickoff_diff_min": None,
        "kickoff_ok":       False,
        "kickoff_note":     reason,
        "match_result":     "❌ Not found",
    }


def _time_diff_minutes(t1: str, t2: str):
    """Return signed difference t2 - t1 in minutes, or None if parse fails."""
    try:
        h1, m1 = int(t1[:2]), int(t1[3:5])
        h2, m2 = int(t2[:2]), int(t2[3:5])
        return (h2 * 60 + m2) - (h1 * 60 + m1)
    except Exception:
        return None


# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Flashscore Comparator",
    page_icon="⚡",
    layout="wide",
)
st.title("⚡ Flashscore ↔ Internal DB Comparator")
st.caption("Compare your internal fixtures against Flashscore feed data.")

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")
    threshold = st.slider("Fuzzy match threshold (%)", 50, 100, FUZZY_THRESHOLD)
    kick_tol  = st.slider("Kick-off tolerance (min)", 0, 120, KICK_OFF_TOLERANCE_MIN, step=5)
    st.divider()
    st.subheader("Google Sheet")
    sheet_link = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
    st.markdown(f"[Open Sheet ↗]({sheet_link})")
    if st.button("🔄 Reload mappings"):
        load_mapping_sheet.clear()
        load_team_map.clear()
        st.success("Mappings reloaded!")

# ─── File uploads ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Internal DB Export (CSV)")
    st.caption("Export from DBeaver — needs columns: match_id, competition_id, home_team, away_team, match_date, kick_off_time")
    db_file = st.file_uploader("Upload DB CSV", type="csv", key="db_upload")

with col2:
    st.subheader("2️⃣ Flashscore Export (CSV)")
    st.caption("Export from Tampermonkey script — needs columns: flashscore_id, home_team, away_team, tournament_id, match_date, kick_off_time")
    fs_file = st.file_uploader("Upload Flashscore CSV", type="csv", key="fs_upload")

# ─── Run comparison ───────────────────────────────────────────────────────────
if db_file and fs_file:
    with st.spinner("Loading data…"):
        db_df = pd.read_csv(db_file)
        fs_df = pd.read_csv(fs_file)
        comp_to_fs_ids = load_mapping_sheet()
        team_map       = load_team_map()

    # ── Validate columns ──────────────────────────────────────────────────────
    DB_REQUIRED = {"match_id", "competition_id", "home_team", "away_team", "match_date", "kick_off_time"}
    FS_REQUIRED = {"flashscore_id", "home_team", "away_team", "tournament_id", "match_date", "kick_off_time"}
    missing_db = DB_REQUIRED - set(db_df.columns)
    missing_fs = FS_REQUIRED - set(fs_df.columns)
    if missing_db or missing_fs:
        if missing_db:
            st.error(f"DB CSV missing columns: {missing_db}")
        if missing_fs:
            st.error(f"Flashscore CSV missing columns: {missing_fs}")
        st.stop()

    # ── Filter DB to competitions that have FS mappings ────────────────────────
    mapped_comp_ids = set(comp_to_fs_ids.keys())
    db_mapped   = db_df[db_df["competition_id"].isin(mapped_comp_ids)]
    db_unmapped = db_df[~db_df["competition_id"].isin(mapped_comp_ids)]

    st.info(
        f"DB: **{len(db_df)}** matches total — "
        f"**{len(db_mapped)}** in mapped competitions, "
        f"**{len(db_unmapped)}** unmapped (skipped)."
    )

    if db_mapped.empty:
        st.warning("No DB matches belong to competitions with a Flashscore mapping.")
        st.stop()

    # ── Run matching ──────────────────────────────────────────────────────────
    with st.spinner("Matching fixtures…"):
        result_df = match_fixtures(db_mapped, fs_df, comp_to_fs_ids, team_map, threshold, kick_tol)

    # ── Summary metrics ───────────────────────────────────────────────────────
    total    = len(result_df)
    found    = (result_df["match_result"] == "✅ Found").sum()
    time_mis = (result_df["match_result"] == "⚠️ Time mismatch").sum()
    not_fnd  = (result_df["match_result"] == "❌ Not found").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total DB matches", total)
    c2.metric("✅ Found + time OK", found)
    c3.metric("⚠️ Time mismatch", time_mis)
    c4.metric("❌ Not found in FS", not_fnd)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_all, tab_issues, tab_missing, tab_unmapped, tab_diag = st.tabs([
        "All Results",
        "Issues (time mismatch / not found)",
        "Missing in DB (FS-only)",
        "Unmapped competitions",
        "Diagnostics",
    ])

    with tab_all:
        st.dataframe(result_df, use_container_width=True)
        csv_out = result_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download full results CSV", csv_out, "flashscore_comparison.csv", "text/csv")

    with tab_issues:
        issues = result_df[result_df["match_result"] != "✅ Found"]
        st.write(f"**{len(issues)} issues**")
        st.dataframe(issues, use_container_width=True)

    with tab_missing:
        matched_fs_ids = set(result_df[result_df["fs_match_id"] != ""]["fs_match_id"])
        missing_df = find_missing_in_db(db_mapped, fs_df, matched_fs_ids)
        st.write(f"**{len(missing_df)} Flashscore matches with no DB counterpart**")
        st.caption("These may be fixtures in your DB under unmapped competitions, or genuinely missing.")
        st.dataframe(missing_df, use_container_width=True)

    with tab_unmapped:
        st.write(f"**{len(db_unmapped)} DB matches in competitions without a Flashscore mapping**")
        if not db_unmapped.empty:
            unc = (
                db_unmapped.groupby("competition_id")
                .agg(count=("match_id", "count"), sample_home=("home_team", "first"))
                .reset_index()
            )
            st.dataframe(unc, use_container_width=True)
            st.markdown("### Add mappings to Google Sheet")
            st.info(
                f"Go to the **`{MAPPING_TAB}`** tab in your Sheet and add rows for these competition IDs. "
                "Format: `competition_id` | `flashscore_tournament_id` (comma-separated for multiple)."
            )

    with tab_diag:
        st.subheader("Mapping diagnostics")
        st.write(f"**{len(comp_to_fs_ids)}** competition → Flashscore ID mappings loaded")
        map_df = pd.DataFrame([
            {"competition_id": k, "flashscore_tournament_ids": ", ".join(v)}
            for k, v in comp_to_fs_ids.items()
        ])
        st.dataframe(map_df, use_container_width=True)

        st.subheader("Team map")
        st.write(f"**{len(team_map)}** manual team name overrides loaded")
        if team_map:
            st.dataframe(
                pd.DataFrame(list(team_map.items()), columns=["internal_name", "mapped_name"]),
                use_container_width=True,
            )

        st.subheader("FS feed sample")
        st.dataframe(fs_df.head(20), use_container_width=True)

    # ── Log to Google Sheet ───────────────────────────────────────────────────
    if st.button("📝 Log results to Google Sheet"):
        with st.spinner("Writing to Sheet…"):
            session_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
            log_rows = []
            for _, row in result_df.iterrows():
                log_rows.append({
                    "session_timestamp":  session_ts,
                    "comp_id":            row["comp_id"],
                    "db_match_id":        row["db_match_id"],
                    "db_home":            row["db_home"],
                    "db_away":            row["db_away"],
                    "db_date":            row["db_date"],
                    "db_kickoff":         row["db_kickoff"],
                    "fs_match_id":        row["fs_match_id"],
                    "fs_home":            row["fs_home"],
                    "fs_away":            row["fs_away"],
                    "fs_tournament":      row["fs_tournament"],
                    "fs_tournament_id":   row["fs_tournament_id"],
                    "fs_kickoff":         row["fs_kickoff"],
                    "fuzzy_score":        row["fuzzy_score"],
                    "kickoff_diff_min":   row["kickoff_diff_min"],
                    "match_result":       row["match_result"],
                    "kickoff_note":       row["kickoff_note"],
                })
            try:
                append_to_log(log_rows)
                st.success(f"✅ Logged {len(log_rows)} rows to `{LOG_TAB_MATCHES}`")
            except Exception as e:
                st.error(f"Failed to write to Sheet: {e}")

elif db_file or fs_file:
    st.info("Please upload both files to run the comparison.")
else:
    # ── Landing instructions ───────────────────────────────────────────────────
    st.markdown("""
    ## How to use

    ### Step 1 — Export from your DB (DBeaver)
    Run a query like:
    ```sql
    SELECT
        m.id            AS match_id,
        cs.competition_id,
        ht.name         AS home_team,
        at.name         AS away_team,
        m.start_date    AS match_date,
        TO_CHAR(m.start_date, 'HH24:MI') AS kick_off_time
    FROM matches m
    JOIN competition_season cs ON m.competition_season_id = cs.id
    JOIN teams ht ON m.home_team_id = ht.id
    JOIN teams at ON m.away_team_id = at.id
    WHERE m.start_date BETWEEN :from_date AND :to_date
    ORDER BY m.start_date;
    ```
    Export as **CSV**.

    ### Step 2 — Export from Flashscore
    Open **flashscore.com**, run the Tampermonkey script, set your day range,
    and click **Fetch & Export CSV**.

    ### Step 3 — Upload both and run
    Upload both CSVs above, then review the results.

    ---
    ### Google Sheet setup
    Your Sheet needs a tab called **`fs_mapping`** with these columns:

    | competition_id | flashscore_tournament_id |
    |---|---|
    | 123 | AbCd1234 |
    | 456 | XyZw5678,QrSt9012 |

    Multiple Flashscore tournament IDs can be comma-separated.
    """)
