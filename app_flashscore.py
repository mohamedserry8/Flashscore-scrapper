"""
Match Data Comparator — Combined (SofaScore + Flashscore)
Streamlit app that compares internal DB fixtures against SofaScore and Flashscore.

DB files (competitions + matches) are shared across both sources.
Navigate between sources using the tabs at the top.
"""

import io
import re
import time
import random
import unicodedata
from datetime import datetime, timedelta, timezone

try:
    import zoneinfo
    _CAIRO_TZ = zoneinfo.ZoneInfo("Africa/Cairo")
except Exception:
    _CAIRO_TZ = None

import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials
from rapidfuzz import fuzz, process

# ─── Config ───────────────────────────────────────────────────────────────────
SHEET_ID           = "14tUgMxJI_glunJiyg8F61D5oduorWJRBRzHfT8zGRtU"
FS_MAPPING_TAB     = "fs_mapping"
TEAM_MAP_TAB       = "team mapping"
FS_LOG_TAB         = "fs_match_log"
SOFASCORE_BASE     = "https://api.sofascore.com/api/v1"

FUZZY_DEFAULT      = 78
KICK_TOL_DEFAULT   = 30    # minutes (Flashscore)

# ─── Default name mappings ────────────────────────────────────────────────────
DEFAULT_MAPPINGS = {
    "Côte d'Ivoire": "Ivory Coast",
    "Congo DR": "DR Congo",
    "Korea Republic": "South Korea",
    "Korea DPR": "North Korea",
    "São Tomé e Príncipe": "Sao Tome and Principe",
    "Bosnia & Herzegovina": "Bosnia and Herzegovina",
    "Chinese Taipei": "Taiwan",
    "Kyrgyz Republic": "Kyrgyzstan",
    "Brazil - Seria A": "Brasileirão Série A",
    "Turkey Superliga": "Süper Lig",
    "Champions League": "UEFA Champions League",
    "PRIMERA DIVISIÓN": "Primera División",
    "Stoiximan Super League": "Super League",
    "Denmark Superliga": "Superliga",
    "Switzerland Super League": "Super League",
    "Chance Liga": "Czech First League",
}

# ─── Abbreviation expansions (merged from both apps) ─────────────────────────
ABBR_MAP = {
    r"\bHFX\b": "Halifax",
    r"\bFC\b": "",   r"\bCF\b": "",   r"\bSC\b": "",   r"\bAC\b": "",
    r"\bAFC\b": "",  r"\bWFC\b": "",  r"\bCFC\b": "",  r"\bFSV\b": "",
    r"\bSV\b": "",   r"\bSK\b": "",   r"\bFK\b": "",   r"\bBK\b": "",
    r"\bIK\b": "",   r"\bGIF\b": "",  r"\bIF\b": "",   r"\bAS\b": "",
    r"\bSSC\b": "",  r"\bUS\b": "",   r"\bUD\b": "",   r"\bRCD\b": "",
    r"\bRC\b": "",   r"\bCD\b": "",   r"\bSD\b": "",   r"\bAD\b": "",
    r"\bRFC\b": "",  r"\bPFC\b": "",  r"\bMFK\b": "",  r"\bNK\b": "",
    r"\bHNK\b": "",  r"\bGNK\b": "",  r"\bDFK\b": "",  r"\bOFK\b": "",
    r"\bJK\b": "",   r"\bTJ\b": "",   r"\bSFC\b": "",  r"\bBFC\b": "",
    r"\bFBC\b": "",  r"\bMSK\b": "",  r"\bFSC\b": "",  r"\bKSC\b": "",
    r"\bBSC\b": "",  r"\bCSC\b": "",  r"\bVfL\b": "",  r"\bVfB\b": "",
    r"\bTSG\b": "",  r"\bRB\b": "",   r"\bEC\b": "",   r"\bCA\b": "",
    r"\bATL\.?\b": "atletico",   r"\bATLETICO\b": "atletico",
    r"\bDEP\.?\b": "deportivo",  r"\bMAN\.?\b": "manchester",
    r"\bNOTT\.?\b": "nottingham", r"\bWOLVES\b": "wolverhampton",
    r"\bSPURS\b": "tottenham",   r"\bST\.?\b": "saint",
    r"\bUTD\.?\b": "united",     r"\bSP\.?\b": "sporting",
    r"\bPSG\b": "paris saint germain",
    r"\bNYC\b": "new york city", r"\bNY\b": "new york",
    r"\bLA\b": "los angeles",    r"\bKC\b": "kansas city",
    r"\b\d{4}\b": "",
}

# ─── Shared helpers ───────────────────────────────────────────────────────────

def normalize_name(name: str, team_map: dict = None) -> str:
    """Unified team name normalization for fuzzy matching."""
    if not isinstance(name, str):
        return ""
    if team_map and name in team_map:
        name = team_map[name]
    # Strip accents
    name = unicodedata.normalize("NFD", name)
    name = "".join(c for c in name if unicodedata.category(c) != "Mn")
    # Abbreviation expansions
    for pat, rep in ABBR_MAP.items():
        name = re.sub(pat, rep, name, flags=re.IGNORECASE)
    # Remove punctuation noise
    name = re.sub(r"[.\-–—&\']", " ", name)
    # Collapse + lowercase
    name = re.sub(r"\s+", " ", name).strip().lower()
    return name or unicodedata.normalize("NFC", str(name)).lower().strip()


def fuzzy_sim(a: str, b: str) -> float:
    return max(fuzz.token_set_ratio(a, b), fuzz.partial_ratio(a, b))


def time_diff_min(t1: str, t2: str):
    try:
        h1, m1 = int(t1[:2]), int(t1[3:5])
        h2, m2 = int(t2[:2]), int(t2[3:5])
        return (h2 * 60 + m2) - (h1 * 60 + m1)
    except Exception:
        return None


def safe_date(val) -> str:
    try:
        return pd.to_datetime(str(val), dayfirst=False).strftime("%Y-%m-%d")
    except Exception:
        try:
            return pd.to_datetime(str(val), dayfirst=True).strftime("%Y-%m-%d")
        except Exception:
            return str(val)


def read_csv_safe(file):
    """Try multiple encodings to handle DBeaver/Excel exports."""
    raw = file.read()
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except Exception:
            continue
    raise ValueError("تعذّر قراءة الملف — جرّب تصدّره بـ UTF-8")


def _today():
    if _CAIRO_TZ:
        return datetime.now(tz=_CAIRO_TZ).date()
    return (datetime.utcnow() + timedelta(hours=3)).date()


# ─── Google Sheets helpers ────────────────────────────────────────────────────

@st.cache_resource
def get_gc():
    try:
        creds_dict = st.secrets["gcp_service_account"]
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    except Exception:
        return None


@st.cache_data(ttl=60)
def discover_gids() -> dict[str, str]:
    """Return {tab_name_lower: gid} from the sheet's HTML."""
    try:
        r = requests.get(
            f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit", timeout=10
        )
        if r.status_code != 200:
            return {}
        ids   = re.findall(r'"sheetId":(\d+)', r.text)
        names = re.findall(r'"title":"([^"]+)"', r.text)
        return {n.lower(): g for n, g in zip(names, ids)}
    except Exception:
        return {}


def _fetch_csv_gid(gid: str) -> pd.DataFrame | None:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            df = pd.read_csv(io.StringIO(r.text))
            df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
            return df
    except Exception:
        pass
    return None


@st.cache_data(ttl=300)
def load_sofascore_mapping() -> tuple[pd.DataFrame, str]:
    """Load Master tab (gid=0) for SofaScore competition → tournament_id mapping."""
    gids = discover_gids()
    candidates = [gids.get("master", "0"), "0"] + list(gids.values())
    seen = set()
    for gid in candidates:
        if gid in seen:
            continue
        seen.add(gid)
        df = _fetch_csv_gid(gid)
        if df is not None and "competition_id" in df.columns:
            if any("sofascore" in c for c in df.columns):
                return df, ""
    # Fallback: gid=0
    df = _fetch_csv_gid("0")
    if df is not None:
        return df, ""
    return pd.DataFrame(), "تعذّر جلب Master tab"


@st.cache_data(ttl=300)
def load_fs_mapping() -> dict[int, list[str]]:
    """Load fs_mapping tab for Flashscore competition → flashscore_tournament_id."""
    # Try via gspread first (authenticated)
    gc = get_gc()
    if gc:
        try:
            sh = gc.open_by_key(SHEET_ID)
            ws = sh.worksheet(FS_MAPPING_TAB)
            df = pd.DataFrame(ws.get_all_records())
            df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        except Exception:
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()

    # Fallback: discover gid via HTML
    if df.empty:
        gids = discover_gids()
        gid = gids.get(FS_MAPPING_TAB.lower(), gids.get("fs_mapping", ""))
        if gid:
            df = _fetch_csv_gid(gid) or pd.DataFrame()

    if df.empty:
        return {}

    mapping = {}
    for _, row in df.iterrows():
        cid  = row.get("competition_id")
        fids = str(row.get("flashscore_tournament_id", "")).strip()
        if cid and fids:
            try:
                mapping[int(cid)] = [x.strip() for x in fids.split(",") if x.strip()]
            except Exception:
                pass
    return mapping


@st.cache_data(ttl=60)
def load_team_map() -> dict[str, str]:
    """
    Load manual team name overrides from the 'team mapping' tab.
    Supports column names from both apps:
      Flashscore: internal_name / external_name
      SofaScore:  db_team_name / sofascore_team_name
    """
    # Try gspread
    gc = get_gc()
    df = pd.DataFrame()
    if gc:
        try:
            sh = gc.open_by_key(SHEET_ID)
            ws = sh.worksheet(TEAM_MAP_TAB)
            df = pd.DataFrame(ws.get_all_records())
            df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        except Exception:
            pass

    if df.empty:
        # Scan all tabs via requests
        gids = discover_gids()
        candidates = list(dict.fromkeys(
            [gids.get(TEAM_MAP_TAB.lower(), ""), gids.get("team_mapping", "")]
            + list(gids.values())
        ))
        for gid in candidates:
            if not gid:
                continue
            candidate_df = _fetch_csv_gid(gid)
            if candidate_df is None:
                continue
            cols = candidate_df.columns
            if any(c in cols for c in ("internal_name", "db_team_name", "db_name")):
                df = candidate_df
                break

    if df.empty:
        return {}

    cols = df.columns
    # Detect column names
    src_col = next((c for c in cols if c in ("internal_name", "db_team_name", "db_name")), None)
    tgt_col = next((c for c in cols if c in ("external_name", "sofascore_team_name", "sf_name")), None)

    if not src_col or not tgt_col:
        return {}

    result = {}
    for _, row in df.iterrows():
        src = str(row.get(src_col, "")).strip()
        tgt = str(row.get(tgt_col, "")).strip()
        if src and tgt and src.lower() not in ("nan", "none") and tgt.lower() not in ("nan", "none"):
            result[src] = tgt
    return result


def build_sf_mapping(mapping_df: pd.DataFrame) -> dict[str, set]:
    """Build {competition_id_str → {sofascore_tournament_ids}} from Master tab."""
    if mapping_df.empty:
        return {}
    cols = mapping_df.columns.tolist()
    cid_col = next((c for c in cols if c == "competition_id"), None) or \
              next((c for c in cols if "competition_id" in c), None)
    sf_col  = next((c for c in cols if c == "sofascore_tournament_id"), None) or \
              next((c for c in cols if "sofascore" in c and "id" in c and "name" not in c), None)
    if not cid_col or not sf_col:
        return {}

    result = {}
    for _, row in mapping_df.iterrows():
        cid_raw = row.get(cid_col, "")
        if pd.isna(cid_raw) or str(cid_raw).strip() in ("", "nan", "None"):
            continue
        cid = (
            str(int(float(str(cid_raw).strip())))
            if str(cid_raw).strip().replace(".", "").isdigit()
            else str(cid_raw).strip()
        )
        sid_raw = row.get(sf_col, "")
        if pd.isna(sid_raw):
            continue
        raw = str(sid_raw).strip()
        if not raw or raw.lower() in ("nan", "none", "not in sofascore", "n/a", "-"):
            continue
        parts = re.split(r"[,\n;\s]+", raw)
        ids = {p.strip().replace(".0", "") for p in parts if p.strip().replace(".0", "").isdigit()}
        if ids:
            result.setdefault(cid, set()).update(ids)
    return result


def append_fs_log(rows: list[dict]):
    """Append match-level rows to fs_match_log sheet (Flashscore feature)."""
    if not rows:
        return
    gc = get_gc()
    if not gc:
        raise RuntimeError("gspread غير متاح — تأكد من إعداد st.secrets[gcp_service_account]")
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(FS_LOG_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=FS_LOG_TAB, rows=1, cols=len(rows[0]))
        ws.append_row(list(rows[0].keys()))
    ws.append_rows([[str(v) for v in r.values()] for r in rows])


# ─── SofaScore data fetching ──────────────────────────────────────────────────

UA_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
]

def _sf_headers():
    return {
        "User-Agent": random.choice(UA_POOL),
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.sofascore.com/",
        "Origin": "https://www.sofascore.com",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
    }


@st.cache_data(ttl=300, show_spinner=False)
def fetch_sf_day(date_str: str) -> tuple[list, str]:
    url = f"{SOFASCORE_BASE}/sport/football/scheduled-events/{date_str}"
    for attempt in range(3):
        try:
            if attempt > 0:
                time.sleep(1.5 * attempt)
            r = requests.get(url, headers=_sf_headers(), timeout=20)
            if r.status_code == 403:
                return [], "403_blocked"
            if r.status_code == 429:
                time.sleep(3)
                continue
            r.raise_for_status()
            rows = []
            for e in r.json().get("events", []):
                try:
                    ts = e.get("startTimestamp", 0)
                    ko = datetime.utcfromtimestamp(ts).strftime("%H:%M") if ts else ""
                    rows.append({
                        "sofascore_id":  e.get("id", ""),
                        "home_team":     e.get("homeTeam", {}).get("name", ""),
                        "away_team":     e.get("awayTeam", {}).get("name", ""),
                        "tournament":    e.get("tournament", {}).get("name", ""),
                        "tournament_id": str(e.get("tournament", {}).get("uniqueTournament", {}).get("id", "")),
                        "category":      e.get("tournament", {}).get("category", {}).get("name", ""),
                        "match_date":    date_str,
                        "kick_off_time": ko,
                        "status":        e.get("status", {}).get("description", ""),
                    })
                except Exception:
                    continue
            return rows, ""
        except requests.exceptions.ConnectionError:
            return [], "connection_error"
        except requests.exceptions.Timeout:
            if attempt == 2:
                return [], "timeout"
        except Exception as ex:
            return [], str(ex)
    return [], "max_retries"


def fetch_sf_range(date_from: datetime, date_to: datetime):
    all_rows = []
    total = (date_to - date_from).days + 1
    bar = st.progress(0, text="جاري جلب بيانات SofaScore...")
    for i in range(total):
        d = (date_from + timedelta(days=i)).strftime("%Y-%m-%d")
        bar.progress((i + 1) / total, text=f"جلب {d}...")
        rows, err = fetch_sf_day(d)
        if err == "403_blocked":
            bar.empty()
            return pd.DataFrame(), "403_blocked"
        if err:
            bar.empty()
            return pd.DataFrame(), err
        all_rows.extend(rows)
        time.sleep(0.4 + random.uniform(0, 0.3))
    bar.empty()
    return (pd.DataFrame(all_rows) if all_rows else pd.DataFrame()), ""


def parse_sf_csv(uploaded_file) -> tuple[pd.DataFrame, str]:
    try:
        df = pd.read_csv(uploaded_file)
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        col_map = {
            "home": "home_team", "home_team_name": "home_team", "hometeam": "home_team",
            "away": "away_team", "away_team_name": "away_team", "awayteam": "away_team",
            "date": "match_date", "game_date": "match_date", "start_date": "match_date",
            "time": "kick_off_time", "kickoff": "kick_off_time", "ko": "kick_off_time",
            "start_time": "kick_off_time", "kick_off": "kick_off_time",
            "league": "tournament", "competition": "tournament", "league_name": "tournament",
        }
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
        missing = [c for c in ("home_team", "away_team", "match_date") if c not in df.columns]
        if missing:
            return pd.DataFrame(), f"الأعمدة دي مش موجودة: {', '.join(missing)}"
        df["match_date"] = pd.to_datetime(df["match_date"], errors="coerce").dt.strftime("%Y-%m-%d")
        if "kick_off_time" not in df.columns:
            df["kick_off_time"] = ""
        if "tournament" not in df.columns:
            df["tournament"] = ""
        df["kick_off_time"] = df["kick_off_time"].astype(str).str[:5]
        df.attrs["already_local"] = "tz_applied" in df.columns
        return df, ""
    except Exception as ex:
        return pd.DataFrame(), str(ex)


def apply_tz_offset(df: pd.DataFrame, offset_hours: int) -> pd.DataFrame:
    if offset_hours == 0:
        return df
    df = df.copy()
    def shift_row(row):
        t = str(row.get("kick_off_time", ""))[:5]
        d = str(row.get("match_date", ""))
        if not t or len(t) < 5 or ":" not in t:
            return row
        try:
            h, m = int(t[:2]), int(t[3:5])
            total = h * 60 + m + offset_hours * 60
            day_shift = total // (24 * 60)
            total = total % (24 * 60)
            if total < 0:
                total += 24 * 60
                day_shift -= 1
            row["kick_off_time"] = f"{total // 60:02d}:{total % 60:02d}"
            if day_shift:
                try:
                    row["match_date"] = (
                        datetime.strptime(d, "%Y-%m-%d") + timedelta(days=day_shift)
                    ).strftime("%Y-%m-%d")
                except Exception:
                    pass
        except Exception:
            pass
        return row
    return df.apply(shift_row, axis=1)


# ─── SofaScore comparison engine ──────────────────────────────────────────────

def compare_sofascore(
    db_df: pd.DataFrame,
    sf_df: pd.DataFrame,
    competition_filter: list,
    fuzzy_threshold: int,
    mappings: dict,
    exclude_cancelled: bool,
    tz_offset: int = 0,
    tracked_comp_ids: set = None,
    sf_mapping: dict = None,
) -> pd.DataFrame:
    db_df = db_df.copy()
    sf_df = sf_df.copy()
    db_df["match_date"] = db_df["match_date"].apply(safe_date)
    sf_df["match_date"] = sf_df["match_date"].apply(safe_date)

    if tz_offset != 0:
        sf_df = apply_tz_offset(sf_df, tz_offset)

    # Deduplicate SofaScore rows
    if len(sf_df):
        key_cols = ["match_date", "home_team", "away_team"]
        if "tournament_id" in sf_df.columns:
            key_cols.append("tournament_id")
        sf_df["_dupkey"] = (
            sf_df[key_cols].astype(str)
            .apply(lambda r: "|".join(v.strip().lower() for v in r), axis=1)
        )
        rank_map = {"not started": 0, "postponed": 0, "canceled": 0, "cancelled": 0}
        sf_df["_rank"] = (
            sf_df.get("status", pd.Series([""] * len(sf_df), index=sf_df.index))
            .astype(str).str.strip().str.lower().map(rank_map).fillna(1)
        )
        n_before = len(sf_df)
        sf_df = (sf_df.sort_values("_rank", ascending=False)
                      .drop_duplicates(subset="_dupkey", keep="first")
                      .drop(columns=["_dupkey", "_rank"]))
        try:
            st.session_state.sf_dupes_removed = n_before - len(sf_df)
        except Exception:
            pass

    if competition_filter:
        db_df = db_df[db_df["competition"].isin(competition_filter)].copy()
    if exclude_cancelled and "match_play_status" in db_df.columns:
        db_df = db_df[db_df["match_play_status"].str.lower() != "cancelled"].copy()

    sf_tid_col = next((c for c in sf_df.columns if c == "tournament_id"), None)
    have_tid   = sf_tid_col is not None and bool(sf_mapping)
    if have_tid:
        sf_df["_tid"] = (
            sf_df[sf_tid_col].astype(str).str.replace(".0", "", regex=False).str.strip()
        )

    comp_id_to_name = {}
    if "competition_id" in db_df.columns and "competition" in db_df.columns:
        comp_id_to_name = dict(zip(db_df["competition_id"].astype(str), db_df["competition"]))

    results       = []
    sf_matched    = set()
    diag_unmatched = []

    for _, db in db_df.iterrows():
        db_date  = str(db.get("match_date", ""))
        db_home  = str(db.get("home_team", ""))
        db_away  = str(db.get("away_team", ""))
        db_comp  = str(db.get("competition", ""))
        db_kick  = str(db.get("kick_off_time", ""))[:5]
        db_home_n = normalize_name(db_home, mappings)
        db_away_n = normalize_name(db_away, mappings)
        db_comp_n = db_comp.lower().strip()

        try:
            d = datetime.strptime(db_date, "%Y-%m-%d")
            cand_dates = [(d + timedelta(days=x)).strftime("%Y-%m-%d") for x in (-1, 0, 1)]
        except Exception:
            cand_dates = [db_date]

        candidates = sf_df[sf_df["match_date"].isin(cand_dates)]

        db_cid = str(db.get("competition_id", "")).strip()
        if db_cid.replace(".", "").isdigit():
            db_cid = str(int(float(db_cid)))

        row_allowed = sf_mapping.get(db_cid, set()) if sf_mapping else set()
        use_comp_boost = True

        if have_tid and row_allowed:
            candidates = candidates[candidates["_tid"].isin(row_allowed)]
            use_comp_boost = False
        elif have_tid and not row_allowed:
            use_comp_boost = True

        if len(candidates) == 0:
            same_date = sf_df[sf_df["match_date"].isin(cand_dates)]
            diag_unmatched.append({
                "competition":     db_comp,
                "competition_id":  db_cid,
                "home_team":       db_home,
                "away_team":       db_away,
                "match_date":      db_date,
                "mapped_sf_ids":   ", ".join(sorted(row_allowed)) if row_allowed else "— مفيش mapping",
                "sf_on_that_date": len(same_date),
                "sf_tids_present": ", ".join(sorted(set(same_date["_tid"]))[:12]) if have_tid and len(same_date) else "",
                "reason": ("البطولة مش مربوطة في الـ Sheet" if not row_allowed
                           else "مفيش ماتش SofaScore بالـ tournament_id المربوط"),
            })

        best_score, best_idx, best_swapped = 0, None, False

        for idx, sf in candidates.iterrows():
            sf_h = normalize_name(str(sf["home_team"]), {})
            sf_a = normalize_name(str(sf["away_team"]), {})
            normal  = (fuzzy_sim(db_home_n, sf_h) + fuzzy_sim(db_away_n, sf_a)) / 2
            swapped = (fuzzy_sim(db_home_n, sf_a) + fuzzy_sim(db_away_n, sf_h)) / 2
            teams   = max(normal, swapped)
            is_sw   = swapped > normal
            if use_comp_boost:
                comp_s = fuzz.token_set_ratio(db_comp_n, str(sf.get("tournament", "")).lower())
                score  = teams * 0.70 + comp_s * 0.30
            else:
                score = teams
            if score > best_score:
                best_score, best_idx, best_swapped = score, idx, is_sw

        if best_idx is not None and best_score >= fuzzy_threshold:
            sf = sf_df.loc[best_idx]
            sf_matched.add(best_idx)
            sf_kick = str(sf.get("kick_off_time", ""))[:5]
            tdiff = None
            try:
                if db_kick and sf_kick and len(db_kick) == 5 and len(sf_kick) == 5:
                    tdiff = int(sf_kick[:2]) * 60 + int(sf_kick[3:]) - int(db_kick[:2]) * 60 - int(db_kick[3:])
            except Exception:
                pass
            status = "⏱ فرق كيك أوف" if (tdiff is not None and abs(tdiff) > 2) else "✓ متطابق"
            if best_swapped:
                status = "🔄 الهوم/أواي معكوس" if status == "✓ متطابق" else status + " + معكوس"
            results.append({
                "status": status, "match_date": db_date, "competition": db_comp,
                "home_team": db_home, "away_team": db_away,
                "db_kickoff": db_kick, "sf_kickoff": sf_kick, "time_diff_min": tdiff,
                "match_score": round(best_score),
                "db_id": db.get("id", ""), "home_team_id": db.get("home_team_id", ""),
                "away_team_id": db.get("away_team_id", ""), "competition_id": db.get("competition_id", ""),
                "sf_home": sf["home_team"], "sf_away": sf["away_team"],
                "sf_tournament": sf.get("tournament", ""), "sf_category": sf.get("category", ""),
            })
        else:
            results.append({
                "status": "🟡 مش في SofaScore", "match_date": db_date, "competition": db_comp,
                "home_team": db_home, "away_team": db_away,
                "db_kickoff": db_kick, "sf_kickoff": "", "time_diff_min": None,
                "match_score": round(best_score),
                "db_id": db.get("id", ""), "home_team_id": db.get("home_team_id", ""),
                "away_team_id": db.get("away_team_id", ""), "competition_id": db.get("competition_id", ""),
                "sf_home": "", "sf_away": "", "sf_tournament": "", "sf_category": "",
            })

    # SofaScore-only matches
    sf_id_to_db_comp = {}
    if sf_mapping:
        for cid, sids in sf_mapping.items():
            name = comp_id_to_name.get(cid, "")
            if name:
                for sid in sids:
                    sf_id_to_db_comp[sid] = (name, cid)

    selected_names = list(comp_id_to_name.values())
    db_name_to_id  = {n.lower().strip(): cid for cid, n in comp_id_to_name.items()}

    for idx, sf in sf_df.iterrows():
        if idx in sf_matched:
            continue
        sf_tourn = str(sf.get("tournament", ""))
        sf_cat   = str(sf.get("category", ""))
        sf_tid   = str(sf.get("tournament_id", "")).replace(".0", "").strip() if have_tid else ""
        if sf_tid and sf_tid in sf_id_to_db_comp:
            db_name, comp_id = sf_id_to_db_comp[sf_tid]
            results.append({
                "status": "🔴 ناقص في DB", "match_date": str(sf["match_date"]),
                "competition": db_name, "home_team": str(sf["home_team"]), "away_team": str(sf["away_team"]),
                "db_kickoff": "", "sf_kickoff": str(sf.get("kick_off_time", ""))[:5],
                "time_diff_min": None, "match_score": 100,
                "db_id": "", "home_team_id": "", "away_team_id": "", "competition_id": comp_id,
                "sf_home": str(sf["home_team"]), "sf_away": str(sf["away_team"]),
                "sf_tournament": sf_tourn, "sf_category": sf_cat,
            })
            continue
        if not selected_names:
            continue
        best = process.extractOne(
            sf_tourn.lower(), [c.lower() for c in selected_names], scorer=fuzz.token_sort_ratio
        )
        if not best or best[1] < 85:
            continue
        db_name = selected_names[[c.lower() for c in selected_names].index(best[0])]
        results.append({
            "status": "🔴 ناقص في DB", "match_date": str(sf["match_date"]),
            "competition": db_name, "home_team": str(sf["home_team"]), "away_team": str(sf["away_team"]),
            "db_kickoff": "", "sf_kickoff": str(sf.get("kick_off_time", ""))[:5],
            "time_diff_min": None, "match_score": best[1],
            "db_id": "", "home_team_id": "", "away_team_id": "",
            "competition_id": db_name_to_id.get(db_name.lower().strip(), ""),
            "sf_home": str(sf["home_team"]), "sf_away": str(sf["away_team"]),
            "sf_tournament": sf_tourn, "sf_category": sf_cat,
        })

    try:
        st.session_state.diag_unmatched = pd.DataFrame(diag_unmatched)
    except Exception:
        pass

    return pd.DataFrame(results) if results else pd.DataFrame()


# ─── Flashscore comparison engine ────────────────────────────────────────────

def _fs_no_match(db_row, reason=""):
    return {
        "comp_id":          db_row.get("competition_id", ""),
        "db_match_id":      db_row.get("match_id", ""),
        "db_home":          db_row.get("home_team", ""),
        "db_away":          db_row.get("away_team", ""),
        "db_date":          str(db_row.get("match_date", ""))[:10],
        "db_kickoff":       str(db_row.get("kick_off_time", ""))[:5],
        "fs_match_id":      "", "fs_home": "", "fs_away": "",
        "fs_tournament":    "", "fs_tournament_id": "",
        "fs_kickoff":       "", "fs_status":        "",
        "fuzzy_score":      0,  "teams_reversed":   False,
        "kickoff_diff_min": None, "kickoff_ok": False,
        "kickoff_note":     reason, "match_result": "❌ Not found",
    }


def compare_flashscore(
    db_df: pd.DataFrame,
    fs_df: pd.DataFrame,
    comp_to_fs_ids: dict,
    team_map: dict,
    threshold: int = 80,
    kick_tol: int = KICK_TOL_DEFAULT,
) -> pd.DataFrame:
    results = []

    def to_iso(val) -> str:
        s = str(val).strip()[:10]
        try:
            return pd.to_datetime(s, dayfirst=False).strftime("%Y-%m-%d")
        except Exception:
            return s

    fs_df = fs_df.copy()
    fs_df["_home_n"]    = fs_df["home_team"].apply(lambda x: normalize_name(x, team_map))
    fs_df["_away_n"]    = fs_df["away_team"].apply(lambda x: normalize_name(x, team_map))
    fs_df["_date_iso"]  = fs_df["match_date"].apply(to_iso)

    for _, db_row in db_df.iterrows():
        comp_id   = int(db_row.get("competition_id", -1))
        db_home   = str(db_row.get("home_team", ""))
        db_away   = str(db_row.get("away_team", ""))
        db_date   = to_iso(db_row.get("match_date", ""))
        db_kick   = str(db_row.get("kick_off_time", ""))[:5]

        allowed = comp_to_fs_ids.get(comp_id, [])
        if not allowed:
            results.append(_fs_no_match(db_row, reason="No FS mapping"))
            continue

        pool = fs_df[
            fs_df["tournament_id"].isin(allowed) & (fs_df["_date_iso"] == db_date)
        ]
        if pool.empty:
            results.append(_fs_no_match(db_row, reason="No FS matches on this date/tournament"))
            continue

        db_h = normalize_name(db_home, team_map)
        db_a = normalize_name(db_away, team_map)
        best_score, best_row, best_rev = -1, None, False

        for _, fs_row in pool.iterrows():
            fh, fa = fs_row["_home_n"], fs_row["_away_n"]
            fwd = (fuzzy_sim(db_h, fh) + fuzzy_sim(db_a, fa)) / 2
            rev = (fuzzy_sim(db_h, fa) + fuzzy_sim(db_a, fh)) / 2
            score, is_rev = (fwd, False) if fwd >= rev else (rev, True)
            if score > best_score:
                best_score, best_row, best_rev = score, fs_row, is_rev

        if best_score < threshold:
            results.append(_fs_no_match(db_row, reason=f"Best fuzzy={best_score:.0f}% < threshold"))
            continue

        fs_kick  = str(best_row.get("kick_off_time", ""))[:5]
        tdiff    = time_diff_min(db_kick, fs_kick)
        time_ok  = tdiff is not None and abs(tdiff) <= kick_tol
        results.append({
            "comp_id":         comp_id,
            "db_match_id":     db_row.get("match_id", ""),
            "db_home":         db_home,  "db_away":  db_away,
            "db_date":         db_date,  "db_kickoff": db_kick,
            "fs_match_id":     best_row.get("flashscore_id", ""),
            "fs_home":         best_row.get("home_team", ""),
            "fs_away":         best_row.get("away_team", ""),
            "fs_tournament":   best_row.get("tournament", ""),
            "fs_tournament_id":best_row.get("tournament_id", ""),
            "fs_kickoff":      fs_kick,
            "fs_status":       best_row.get("status", ""),
            "fuzzy_score":     round(best_score, 1),
            "teams_reversed":  best_rev,
            "kickoff_diff_min":tdiff,
            "kickoff_ok":      time_ok,
            "kickoff_note":    f"Match ({tdiff:+d} min)" if tdiff is not None else "N/A",
            "match_result":    "✅ Found" if time_ok else "⚠️ Time mismatch",
        })

    return pd.DataFrame(results)


def find_fs_missing(db_df, fs_df, matched_ids: set) -> pd.DataFrame:
    return fs_df[~fs_df["flashscore_id"].isin(matched_ids)].copy()


# ════════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Match Comparator",
    page_icon="⚽",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #0f1117; }
[data-testid="stSidebar"] { background: #161b27; border-right: 1px solid #2a3147; }
div[data-testid="stDataFrameResizable"] { border: 1px solid #2a3147; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

st.title("⚽⚡ Match Data Comparator")
st.caption("قارن بياناتك مع SofaScore و Flashscore في نفس الوقت")

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ الإعدادات")

    today = _today()

    st.subheader("📅 نطاق التاريخ")
    preset = st.radio("اختصارات", ["اليوم", "اليوم + بكرة", "أسبوع قادم", "مخصص"],
                      index=2, horizontal=True)
    if preset == "اليوم":
        date_from = date_to = today
    elif preset == "اليوم + بكرة":
        date_from, date_to = today, today + timedelta(days=1)
    elif preset == "أسبوع قادم":
        date_from, date_to = today, today + timedelta(days=7)
    else:
        c1, c2 = st.columns(2)
        date_from = c1.date_input("من", value=today)
        date_to   = c2.date_input("إلى", value=today + timedelta(days=7))

    st.divider()
    st.subheader("🎯 حساسية المطابقة")
    fuzzy_threshold = st.slider("Fuzzy Match %", 50, 100, FUZZY_DEFAULT)

    st.divider()
    with st.expander("⚽ إعدادات SofaScore"):
        tz_offset = st.number_input(
            "SofaScore UTC → توقيتك (ساعات)",
            min_value=-12, max_value=14, value=3, step=1,
            help="القاهرة = +3",
        )
        exclude_cancelled = st.checkbox("استبعاد Cancelled", value=True)

    with st.expander("⚡ إعدادات Flashscore"):
        kick_tol = st.slider("Kick-off tolerance (min)", 0, 120, KICK_TOL_DEFAULT, step=5)

    st.divider()
    st.subheader("📊 Google Sheet")
    st.markdown(f"[Open Sheet ↗](https://docs.google.com/spreadsheets/d/{SHEET_ID})")
    if st.button("🔄 Reload mappings"):
        load_sofascore_mapping.clear()
        load_fs_mapping.clear()
        load_team_map.clear()
        discover_gids.clear()
        st.success("Mappings reloaded!")

# ─── Shared DB uploads ────────────────────────────────────────────────────────
st.subheader("🗄️ بيانات الـ DB (مشتركة بين المصدرين)")
st.caption("ارفع الملفات دي مرة واحدة وهيتستخدموا في SofaScore و Flashscore")

uc1, uc2, uc3 = st.columns(3)

with uc1:
    st.markdown("**البطولات** *(اختياري)*")
    comp_file = st.file_uploader(
        "competitions CSV",
        type=["csv"], key="shared_comp",
        help="competition_id, competition, competition_country, season"
    )

with uc2:
    st.markdown("**الماتشات (DB)**")
    match_file = st.file_uploader(
        "matches CSV من DBeaver",
        type=["csv"], key="shared_matches",
        help="id / match_id, competition_id, home_team, away_team, match_date, kick_off_time"
    )

with uc3:
    st.markdown("**Flashscore export** *(للـ Flashscore tab)*")
    fs_file = st.file_uploader(
        "Flashscore CSV (Tampermonkey)",
        type=["csv"], key="shared_fs",
        help="flashscore_id, home_team, away_team, tournament_id, match_date, kick_off_time"
    )

# Load shared files
comp_df  = None
db_df    = None
fs_ext_df = None

if comp_file:
    try:
        comp_df = pd.read_csv(comp_file)
        comp_df.columns = [c.strip().lower() for c in comp_df.columns]
        uc1.success(f"✅ {len(comp_df)} بطولة")
    except Exception as e:
        uc1.error(f"خطأ: {e}")

if match_file:
    try:
        db_df = read_csv_safe(match_file)
        db_df.columns = [c.strip().lower() for c in db_df.columns]
        # Accept 'id' as alias for 'match_id'
        if "id" in db_df.columns and "match_id" not in db_df.columns:
            db_df = db_df.rename(columns={"id": "match_id"})
        db_df["match_date"] = pd.to_datetime(db_df["match_date"], errors="coerce").dt.strftime("%Y-%m-%d")
        uc2.success(f"✅ {len(db_df):,} ماتش")
    except Exception as e:
        uc2.error(f"خطأ: {e}")

if fs_file:
    try:
        fs_ext_df = read_csv_safe(fs_file)
        uc3.success(f"✅ {len(fs_ext_df):,} ماتش FS")
    except Exception as e:
        uc3.error(f"خطأ: {e}")

st.divider()

# ─── Main tabs ────────────────────────────────────────────────────────────────
tab_sf, tab_fs, tab_mapping, tab_help = st.tabs([
    "⚽ SofaScore", "⚡ Flashscore", "🔗 Name Mapping", "❓ مساعدة"
])


# ════════════════════════════════════════════════════════════════════════════════
# TAB: SofaScore
# ════════════════════════════════════════════════════════════════════════════════
with tab_sf:

    # Load mappings
    with st.spinner("جلب الـ SofaScore mapping..."):
        sf_map_df, sf_map_err = load_sofascore_mapping()

    if sf_map_err:
        st.error(f"❌ {sf_map_err}")
    elif not sf_map_df.empty:
        sf_mapping = build_sf_mapping(sf_map_df)
        mapped_count = sum(len(v) for v in sf_mapping.values())
        st.success(f"✅ {len(sf_mapping)} بطولة مربوطة ({mapped_count} tournament IDs)")

        team_map_sheet = load_team_map()
        if team_map_sheet:
            st.info(f"🔗 {len(team_map_sheet)} فريق مربوط يدوياً")
        st.session_state["sf_team_map"] = team_map_sheet

        with st.expander("🎯 IDs للـ Targeted Exporter"):
            all_sf_ids = sorted(
                {i for ids in sf_mapping.values() for i in ids},
                key=lambda x: int(x) if x.isdigit() else 0,
            )
            st.caption(f"{len(all_sf_ids)} tournament ID — الزقهم في الـ Tampermonkey script")
            st.text_area("انسخ من هنا", value=", ".join(all_sf_ids), height=110, key="sf_ids_out")
    else:
        sf_mapping = {}
        st.warning("⚠️ الـ SofaScore mapping فاضي")

    st.divider()

    # SofaScore data source
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🌐 ماتشات SofaScore")
        sf_manual_file = st.file_uploader(
            "CSV من Tampermonkey script",
            type=["csv"], key="sf_manual_tab",
            help="sofascore_id, home_team, away_team, tournament_id, match_date, kick_off_time"
        )

    sf_manual_df = None
    if sf_manual_file:
        sf_manual_df, err = parse_sf_csv(sf_manual_file)
        with col_a:
            if err:
                st.error(f"❌ {err}")
            else:
                st.success(f"✅ {len(sf_manual_df):,} ماتش")
                if "tournament_id" not in sf_manual_df.columns:
                    st.error("⚠️ الملف مفيهوش عمود tournament_id — حدّث الـ script")
                with st.expander("معاينة"):
                    st.dataframe(sf_manual_df.head(8), use_container_width=True)

    if db_df is not None and comp_df is not None:
        st.subheader("🏆 اختار البطولات")
        db_comps_all   = sorted(db_df["competition"].dropna().unique().tolist()) if "competition" in db_df.columns else []
        country_opts   = sorted(comp_df["competition_country"].dropna().unique().tolist()) if "competition_country" in comp_df.columns else []

        c1, c2, c3 = st.columns([3, 1, 1])
        with c1:
            country_filter = st.multiselect("فلتر بالدولة (اختياري)", options=country_opts, placeholder="كل الدول", key="sf_country")
        sel_all_sf = c2.button("✅ الكل", key="sf_sel_all")
        clr_all_sf = c3.button("❌ مسح", key="sf_clr_all")

        if country_filter and comp_df is not None and "competition_country" in comp_df.columns:
            in_country = comp_df[comp_df["competition_country"].isin(country_filter)]["competition"].dropna().unique().tolist()
            filtered = [c for c in db_comps_all if
                        process.extractOne(c.lower(), [x.lower() for x in in_country], scorer=fuzz.token_sort_ratio) and
                        process.extractOne(c.lower(), [x.lower() for x in in_country], scorer=fuzz.token_sort_ratio)[1] >= 70]
        else:
            filtered = db_comps_all

        if "sf_selected_comps" not in st.session_state or sel_all_sf:
            st.session_state.sf_selected_comps = filtered
        if clr_all_sf:
            st.session_state.sf_selected_comps = []

        selected_comps = st.multiselect(
            f"البطولات ({len(filtered)} متاحة)",
            options=filtered,
            default=[c for c in st.session_state.sf_selected_comps if c in filtered],
            key="sf_comp_sel",
        )
        st.session_state.sf_selected_comps = selected_comps

        st.divider()
        sf_ready = sf_manual_df is not None and not sf_manual_df.empty
        days_n   = (date_to - date_from).days + 1

        run_col, info_col = st.columns([1, 3])
        with run_col:
            run_sf = st.button("▶ تشغيل SofaScore", type="primary",
                               disabled=(not selected_comps or not sf_ready),
                               use_container_width=True)
        with info_col:
            if not sf_ready:
                st.warning("⬆️ ارفع ملف SofaScore")
            else:
                st.info(f"📅 {date_from} ← {date_to} ({days_n} يوم) | 🏆 {len(selected_comps)} بطولة | 🌐 {len(sf_manual_df):,} ماتش")

        if run_sf:
            sf_df_run = sf_manual_df.copy()
            already_local = "tz_applied" in sf_df_run.columns
            eff_tz        = 0 if already_local else tz_offset
            if already_local and tz_offset:
                st.info(f"ℹ️ الملف متحوّل محلياً — تم تجاهل offset ({tz_offset:+d})")

            db_filtered = db_df[
                (db_df["match_date"] >= date_from.strftime("%Y-%m-%d")) &
                (db_df["match_date"] <= date_to.strftime("%Y-%m-%d"))
            ].copy()

            with st.expander("🔍 Debug — إيه اللي بيتقارن", expanded=True):
                d1, d2, d3 = st.columns(3)
                d1.metric("DB بعد فلتر التاريخ", len(db_filtered))
                d2.metric("SofaScore ماتشات", len(sf_df_run))
                d3.metric("البطولات المختارة", len(selected_comps))
                if len(db_filtered) == 0:
                    st.error(f"⚠️ مفيش ماتشات في DB للفترة {date_from} ← {date_to}")

            combined_maps = {
                **DEFAULT_MAPPINGS,
                **st.session_state.get("mappings", {}),
                **st.session_state.get("sf_team_map", {}),
            }
            tracked = (
                set(comp_df["competition_id"].dropna().astype(str).unique())
                if comp_df is not None and "competition_id" in comp_df.columns
                else None
            )

            with st.spinner("جاري المقارنة..."):
                result_sf = compare_sofascore(
                    db_filtered, sf_df_run, selected_comps,
                    fuzzy_threshold, combined_maps, exclude_cancelled,
                    tz_offset=eff_tz, tracked_comp_ids=tracked, sf_mapping=sf_mapping,
                )

            if result_sf.empty:
                st.warning("⚠️ النتيجة فاضية — تأكد من الـ debug")
            else:
                st.session_state.result_sf  = result_sf
                st.session_state.sf_df_cache = sf_df_run

        # Show results
        if "result_sf" in st.session_state and not st.session_state.result_sf.empty:
            result_sf = st.session_state.result_sf
            counts   = result_sf["status"].value_counts()
            n_miss   = counts.get("🔴 ناقص في DB",       0)
            n_extra  = counts.get("🟡 مش في SofaScore",  0)
            n_time   = counts.get("⏱ فرق كيك أوف",      0)
            n_ok     = counts.get("✓ متطابق",            0)
            db_side  = n_extra + n_time + n_ok
            pct_ok   = round(n_ok / db_side * 100) if db_side else 0

            st.markdown(f"""
<div style="background:#1e2535;border:1px solid #2a3147;border-radius:10px;padding:14px 20px;margin-bottom:12px">
<b style="color:#e2e8f0">📊 ملخص SofaScore</b><br>
<span style="color:#8892a4;font-size:13px">
من <b style="color:#e2e8f0">{db_side}</b> ماتش DB:
&nbsp;✓ <b style="color:#22c55e">{n_ok} متطابق ({pct_ok}%)</b>
&nbsp;|&nbsp; ⏱ <b style="color:#a78bfa">{n_time} فرق وقت</b>
&nbsp;|&nbsp; 🟡 <b style="color:#f59e0b">{n_extra} مش في SF</b>
&nbsp;&nbsp;&nbsp; + 🔴 <b style="color:#ef4444">{n_miss} ناقص في DB</b>
</span></div>""", unsafe_allow_html=True)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("🔴 ناقص في DB",      n_miss)
            c2.metric("🟡 مش في SofaScore", n_extra)
            c3.metric("⏱ فرق كيك أوف",     n_time)
            c4.metric("✓ متطابق",           n_ok)
            c5.metric("دقة DB",             f"{pct_ok}%")
            st.progress(pct_ok / 100)

            st.divider()
            f1, f2, f3, f4, f5 = st.columns([2, 2, 2, 1, 1])
            search   = f1.text_input("🔍 بحث", placeholder="اسم الفريق...", key="sf_search")
            status_f = f2.multiselect("الحالة",   result_sf["status"].unique().tolist(),
                                       default=result_sf["status"].unique().tolist(), key="sf_status_f")
            comp_f   = f3.multiselect("البطولة",  sorted(result_sf["competition"].unique().tolist()), key="sf_comp_f")
            date_f   = f4.selectbox("التاريخ", ["الكل"] + sorted(result_sf["match_date"].unique().tolist()), key="sf_date_f")
            min_sc   = f5.number_input("أدنى تطابق%", 0, 100, 0, step=5, key="sf_min_sc")

            disp = result_sf.copy()
            if status_f:  disp = disp[disp["status"].isin(status_f)]
            if comp_f:    disp = disp[disp["competition"].isin(comp_f)]
            if date_f != "الكل": disp = disp[disp["match_date"] == date_f]
            if search:    disp = disp[
                disp["home_team"].str.contains(search, case=False, na=False) |
                disp["away_team"].str.contains(search, case=False, na=False) |
                disp["competition"].str.contains(search, case=False, na=False)
            ]
            if min_sc > 0: disp = disp[disp["match_score"] >= min_sc]

            st.caption(f"عرض **{len(disp):,}** من {len(result_sf):,}")
            st.dataframe(
                disp[["status","match_date","competition","home_team","away_team",
                       "db_kickoff","sf_kickoff","time_diff_min","match_score",
                       "sf_tournament","sf_category","db_id","competition_id"]]
                .rename(columns={
                    "status":"الحالة","match_date":"التاريخ","competition":"البطولة",
                    "home_team":"الهوم","away_team":"الأواي",
                    "db_kickoff":"كيك DB","sf_kickoff":"كيك SF","time_diff_min":"فرق(د)",
                    "match_score":"تطابق%","sf_tournament":"بطولة SF",
                    "sf_category":"دولة SF","db_id":"DB ID","competition_id":"Comp ID",
                }),
                use_container_width=True, height=500,
                column_config={
                    "تطابق%": st.column_config.ProgressColumn("تطابق%", min_value=0, max_value=100, format="%d%%"),
                    "فرق(د)": st.column_config.NumberColumn("فرق(د)", format="%+d د"),
                },
            )

            st.divider()
            ec1, ec2, ec3 = st.columns(3)
            ec1.download_button("⬇️ Export الفلتر", disp.to_csv(index=False).encode("utf-8-sig"),
                                f"sf_filtered_{date_from}_{date_to}.csv", "text/csv")
            ec2.download_button("🚨 Export المشاكل",
                                result_sf[result_sf["status"] != "✓ متطابق"].to_csv(index=False).encode("utf-8-sig"),
                                f"sf_problems_{date_from}_{date_to}.csv", "text/csv")
            ec3.download_button("✅ Export المتطابقات",
                                result_sf[result_sf["status"] == "✓ متطابق"].to_csv(index=False).encode("utf-8-sig"),
                                f"sf_matched_{date_from}_{date_to}.csv", "text/csv")

            # Diagnostics
            diag = st.session_state.get("diag_unmatched")
            if diag is not None and len(diag) > 0:
                with st.expander(f"🔎 تشخيص: {len(diag)} ماتش مالقاش مرشح"):
                    for r, c in diag["reason"].value_counts().items():
                        st.write(f"- **{c}** ماتش: {r}")
                    st.dataframe(diag, use_container_width=True, height=300)

    elif db_df is None or comp_df is None:
        st.info("⬆️ ارفع ملف الماتشات (+ البطولات اختياري) من الأعلى")


# ════════════════════════════════════════════════════════════════════════════════
# TAB: Flashscore
# ════════════════════════════════════════════════════════════════════════════════
with tab_fs:

    with st.spinner("جلب الـ Flashscore mapping..."):
        fs_comp_map = load_fs_mapping()
        fs_team_map = load_team_map()

    if fs_comp_map:
        st.success(f"✅ {len(fs_comp_map)} competition مربوطة بـ Flashscore")
    else:
        st.warning("⚠️ مفيش Flashscore mapping — تأكد من tab `fs_mapping` في الـ Sheet")

    if fs_ext_df is None or fs_ext_df.empty:
        st.info("⬆️ ارفع ملف Flashscore CSV من قسم الـ DB files فوق")
    else:
        # Validate columns
        FS_REQUIRED = {"flashscore_id", "home_team", "away_team", "tournament_id", "match_date", "kick_off_time"}
        missing_fs = FS_REQUIRED - set(fs_ext_df.columns)
        if missing_fs:
            st.error(f"Flashscore CSV ناقص أعمدة: {missing_fs}")
        elif db_df is None:
            st.info("⬆️ ارفع ملف الماتشات من الأعلى")
        else:
            DB_REQUIRED = {"match_id", "competition_id", "home_team", "away_team", "match_date", "kick_off_time"}
            missing_db = DB_REQUIRED - set(db_df.columns)
            if missing_db:
                st.error(f"DB CSV ناقص أعمدة: {missing_db}")
            else:
                mapped_ids    = set(fs_comp_map.keys())
                db_mapped     = db_df[db_df["competition_id"].apply(
                    lambda x: (int(float(str(x))) if str(x).replace(".", "").isdigit() else x)
                ).isin(mapped_ids)]
                db_unmapped   = db_df[~db_df.index.isin(db_mapped.index)]

                # Filter by date range
                db_mapped_dt = db_mapped[
                    (db_mapped["match_date"] >= date_from.strftime("%Y-%m-%d")) &
                    (db_mapped["match_date"] <= date_to.strftime("%Y-%m-%d"))
                ].copy()

                st.info(
                    f"DB: **{len(db_df)}** ماتش إجمالي — "
                    f"**{len(db_mapped_dt)}** في النطاق الزمني والـ mapping — "
                    f"**{len(db_unmapped)}** بدون mapping (هيتخطوا)"
                )

                if db_mapped_dt.empty:
                    st.warning("مفيش ماتشات في نطاق التاريخ المختار لبطولات مربوطة.")
                else:
                    run_col_fs, _ = st.columns([1, 3])
                    run_fs = run_col_fs.button("▶ تشغيل Flashscore", type="primary", use_container_width=True)

                    if run_fs:
                        with st.spinner("Matching fixtures…"):
                            combined_tm = {**fs_team_map, **st.session_state.get("mappings", {})}
                            result_fs = compare_flashscore(
                                db_mapped_dt, fs_ext_df, fs_comp_map,
                                combined_tm, fuzzy_threshold, kick_tol,
                            )
                        st.session_state.result_fs = result_fs

                    if "result_fs" in st.session_state and not st.session_state.result_fs.empty:
                        result_fs = st.session_state.result_fs
                        total    = len(result_fs)
                        found    = (result_fs["match_result"] == "✅ Found").sum()
                        time_mis = (result_fs["match_result"] == "⚠️ Time mismatch").sum()
                        not_fnd  = (result_fs["match_result"] == "❌ Not found").sum()

                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Total DB", total)
                        c2.metric("✅ Found", found)
                        c3.metric("⚠️ Time mismatch", time_mis)
                        c4.metric("❌ Not found", not_fnd)

                        fs_tabs = st.tabs(["All", "Issues", "Missing in DB", "Unmapped", "Diagnostics"])

                        with fs_tabs[0]:
                            st.dataframe(result_fs, use_container_width=True)
                            st.download_button(
                                "⬇️ Download CSV",
                                result_fs.to_csv(index=False).encode("utf-8"),
                                "flashscore_comparison.csv", "text/csv",
                            )

                        with fs_tabs[1]:
                            issues = result_fs[result_fs["match_result"] != "✅ Found"]
                            st.write(f"**{len(issues)} issues**")
                            st.dataframe(issues, use_container_width=True)

                        with fs_tabs[2]:
                            matched_ids = set(result_fs[result_fs["fs_match_id"] != ""]["fs_match_id"])
                            miss_df     = find_fs_missing(db_mapped_dt, fs_ext_df, matched_ids)
                            st.write(f"**{len(miss_df)} FS matches بدون مقابل في DB**")
                            st.dataframe(miss_df, use_container_width=True)

                        with fs_tabs[3]:
                            st.write(f"**{len(db_unmapped)} ماتش في بطولات بدون Flashscore mapping**")
                            if not db_unmapped.empty:
                                unc = (
                                    db_unmapped.groupby("competition_id")
                                    .agg(count=("match_id", "count"), sample_home=("home_team", "first"))
                                    .reset_index()
                                )
                                st.dataframe(unc, use_container_width=True)
                                st.info(
                                    f"أضف الـ competition_ids دي في tab **`{FS_MAPPING_TAB}`** في الـ Sheet. "
                                    "Format: `competition_id` | `flashscore_tournament_id`"
                                )

                        with fs_tabs[4]:
                            st.subheader("Mapping diagnostics")
                            map_df = pd.DataFrame([
                                {"competition_id": k, "flashscore_ids": ", ".join(v)}
                                for k, v in fs_comp_map.items()
                            ])
                            st.dataframe(map_df, use_container_width=True)
                            st.subheader("Team map")
                            if fs_team_map:
                                st.dataframe(
                                    pd.DataFrame(list(fs_team_map.items()), columns=["internal", "mapped"]),
                                    use_container_width=True,
                                )

                        # Log to Sheet
                        if st.button("📝 Log results to Google Sheet"):
                            with st.spinner("Writing…"):
                                ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
                                log_rows = [
                                    {
                                        "session_timestamp": ts,
                                        **{k: row[k] for k in (
                                            "comp_id","db_match_id","db_home","db_away",
                                            "db_date","db_kickoff","fs_match_id","fs_home",
                                            "fs_away","fs_tournament","fs_tournament_id",
                                            "fs_kickoff","fuzzy_score","kickoff_diff_min",
                                            "match_result","kickoff_note",
                                        )}
                                    }
                                    for _, row in result_fs.iterrows()
                                ]
                                try:
                                    append_fs_log(log_rows)
                                    st.success(f"✅ Logged {len(log_rows)} rows to `{FS_LOG_TAB}`")
                                except Exception as e:
                                    st.error(f"Failed: {e}")


# ════════════════════════════════════════════════════════════════════════════════
# TAB: Name Mapping
# ════════════════════════════════════════════════════════════════════════════════
with tab_mapping:
    st.subheader("جدول تعيين الأسماء")
    st.caption("لو اسم فريق في بياناتك مختلف عن اسمه على SofaScore أو Flashscore")

    if "mappings" not in st.session_state:
        st.session_state.mappings = DEFAULT_MAPPINGS.copy()

    mdf = pd.DataFrame([
        {"DB Name": k, "External Name": v}
        for k, v in st.session_state.mappings.items()
    ])
    edited = st.data_editor(mdf, num_rows="dynamic", use_container_width=True)
    if st.button("💾 حفظ"):
        st.session_state.mappings = {
            r["DB Name"]: r["External Name"]
            for _, r in edited.iterrows()
            if r["DB Name"] and r["External Name"]
        }
        st.success(f"✅ تم حفظ {len(st.session_state.mappings)} mapping")

    st.divider()
    st.caption("الـ mappings دي بتتطبق على الاثنين — SofaScore وFlashscore")


# ════════════════════════════════════════════════════════════════════════════════
# TAB: Help
# ════════════════════════════════════════════════════════════════════════════════
with tab_help:
    st.subheader("كيفية الاستخدام")
    st.markdown("""
### ملفات الـ DB (مشتركة)

**competitions CSV** *(اختياري)*
```
competition_id, competition, competition_country, season
```

**matches CSV** (من DBeaver)
```sql
SELECT
    m.id            AS match_id,
    cs.competition_id,
    cs.competition,
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

---

### ⚽ SofaScore Tab
ارفع CSV من Tampermonkey script — الأعمدة المطلوبة:
```
sofascore_id, home_team, away_team, tournament_id, match_date, kick_off_time
```

---

### ⚡ Flashscore Tab
ارفع CSV من Tampermonkey script — الأعمدة المطلوبة:
```
flashscore_id, home_team, away_team, tournament_id, match_date, kick_off_time
```

---

### Google Sheet setup
الـ Sheet محتاج الـ tabs دي:

| Tab | الأعمدة |
|-----|---------|
| **Master** (gid=0) | `competition_id`, `sofascore_tournament_id` |
| **fs_mapping** | `competition_id`, `flashscore_tournament_id` |
| **team mapping** | `internal_name`, `external_name` |
""")
