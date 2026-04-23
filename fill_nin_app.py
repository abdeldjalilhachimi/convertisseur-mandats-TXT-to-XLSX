#!/usr/bin/env python3
"""
Streamlit app: fill missing NIN / NIS / Nom / Prenom in personnel files
using the official NIN_A.XLS as source of truth.

Matching key: rightmost 10 digits of NUMCPT (works for both 10-digit
bare accounts and 18-digit prefixed accounts).

Run:
    streamlit run fill_nin_app.py
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
OFFICIAL_PATH = Path(__file__).parent / "data" / "NIN_A.XLS"
OFFICIAL_HEADER_ROW = 4  # header lives on Excel row 5 (0-indexed = 4)

FIELDS = ("NIN", "NIS", "Nom", "Prenom")

COL_ALIASES: dict[str, list[str]] = {
    "NUMCPT": ["NUMCPT", "ح ب ج", "CCP", "Compte", "N° Compte", "N Compte", "NumCpt"],
    "NIN":    ["NIN", "رقم التعريف الوطني", "N° NIN", "N NIN"],
    "NIS":    ["NIS", "رقم ض إ", "N° NIS", "N NIS", "SS"],
    "Nom":    ["Nom", "اللقب", "Last Name", "LastName", "NOM"],
    "Prenom": ["Prenom", "Prénom", "الاسم", "First Name", "FirstName", "PRENOM"],
}


# ─── normalization helpers ────────────────────────────────────────────────────

_DIGITS_RE = re.compile(r"\D+")


def norm_key(value: Any) -> str:
    """Return rightmost 10 digits of value, zero-padded. Empty string if no digits.

    Handles pandas float-cast of integer IDs ('11049487.0' → '0011049487'):
    the trailing '.0' is stripped before digit extraction so its zero is
    not mistaken for part of the account number.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    # Strip '.0' / '.00' suffix left by float→str of integer ids.
    if "." in s:
        left, _, right = s.partition(".")
        if right and set(right) == {"0"} and left.lstrip("-").isdigit():
            s = left
    digits = _DIGITS_RE.sub("", s)
    if not digits:
        return ""
    if len(digits) > 10:
        digits = digits[-10:]
    return digits.zfill(10)


def is_blank(v: Any) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return s == "" or s.lower() == "nan"


def clean_str(v: Any) -> str:
    """Turn a cell value into a clean string (strip trailing .0 from float-cast ints)."""
    if is_blank(v):
        return ""
    s = str(v).strip()
    # pandas float-cast of an int id: "1.19831355001e+17" or "119831355001300004.0"
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    return s


# ─── loaders ──────────────────────────────────────────────────────────────────


@st.cache_data(show_spinner=False)
def load_official(path_str: str, mtime: float) -> tuple[dict[str, dict], int, int]:
    """Load NIN_A.XLS into {key: {NIN,NIS,Nom,Prenom}} dict. mtime in cache key."""
    ext = Path(path_str).suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    try:
        df = pd.read_excel(
            path_str,
            sheet_name=0,
            header=OFFICIAL_HEADER_ROW,
            dtype=str,
            engine=engine,
        )
    except ImportError as e:
        raise ImportError(
            f"Dépendance manquante pour lire {ext} : {e}. "
            f"Installez la dans l'environnement de Streamlit : "
            f"`pip install xlrd==2.0.1` puis redémarrez `streamlit run`."
        ) from e
    # Drop fully-empty rows
    df = df.dropna(how="all")
    if "NUMCPT" not in df.columns:
        raise ValueError(
            f"Colonne NUMCPT introuvable dans {path_str}. Colonnes : {list(df.columns)}"
        )

    lookup: dict[str, dict] = {}
    duplicates = 0
    for _, row in df.iterrows():
        key = norm_key(row.get("NUMCPT"))
        if not key:
            continue
        entry = {f: clean_str(row.get(f)) for f in FIELDS}
        if key in lookup:
            duplicates += 1
            continue
        lookup[key] = entry
    return lookup, len(df), duplicates


def load_dynamic(file_bytes: bytes, filename: str) -> tuple[pd.DataFrame, str]:
    """Load a dynamic personnel file. Tries read_excel then read_html."""
    # 1) read_excel (handles real .xls via xlrd, .xlsx via openpyxl)
    ext = Path(filename).suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    try:
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=0,
            dtype=str,
            engine=engine,
        )
        return df, f"read_excel (engine={engine})"
    except Exception as excel_err:
        excel_msg = str(excel_err)

    # 2) read_html (for HTML saved as .xls)
    try:
        tables = pd.read_html(io.BytesIO(file_bytes), flavor=None)
    except Exception as html_err:
        raise ValueError(
            f"Impossible de lire {filename} (ni Excel ni HTML).\n"
            f"read_excel : {excel_msg}\nread_html : {html_err}"
        )

    if not tables:
        raise ValueError(f"Aucun tableau trouvé dans {filename}.")

    # pick the table with the most columns (most likely the data sheet)
    df = max(tables, key=lambda t: t.shape[1])
    # normalize column names: dropna headers
    df.columns = [str(c).strip() for c in df.columns]
    # cast everything to string
    df = df.astype(str)
    # pandas read_html turns NaN into "nan" strings — keep as-is, is_blank handles it
    return df, f"read_html ({len(tables)} table(s), widest retenue)"


# ─── column resolution ────────────────────────────────────────────────────────


def resolve_columns(df: pd.DataFrame) -> dict[str, str]:
    """Map canonical → actual column name. Missing canonicals simply absent."""
    actual_cols = list(df.columns)
    actual_lower = {str(c).strip().lower(): c for c in actual_cols}
    resolved: dict[str, str] = {}
    for canon, aliases in COL_ALIASES.items():
        for alias in aliases:
            key = alias.strip().lower()
            if key in actual_lower:
                resolved[canon] = actual_lower[key]
                break
    return resolved


# ─── fill ─────────────────────────────────────────────────────────────────────


def fill_missing(
    df: pd.DataFrame,
    col_map: dict[str, str],
    official: dict[str, dict],
) -> tuple[pd.DataFrame, dict[str, int], int, int, list[tuple[int, str]]]:
    stats = {f: 0 for f in FIELDS}
    matched = 0
    empty_key = 0
    unmatched: list[tuple[int, str]] = []

    numcpt_col = col_map["NUMCPT"]
    out = df.copy()

    for idx, raw_key in out[numcpt_col].items():
        key = norm_key(raw_key)
        if not key:
            empty_key += 1
            continue
        ref = official.get(key)
        if ref is None:
            unmatched.append((idx, str(raw_key)))
            continue
        matched += 1
        for field in FIELDS:
            if field not in col_map:
                continue
            target_col = col_map[field]
            cur = out.at[idx, target_col]
            if is_blank(cur) and not is_blank(ref.get(field)):
                out.at[idx, target_col] = ref[field]
                stats[field] += 1

    return out, stats, matched, empty_key, unmatched


# ─── output ───────────────────────────────────────────────────────────────────


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Personnel")
    return buf.getvalue()


# ─── UI ───────────────────────────────────────────────────────────────────────


st.set_page_config(
    page_title="Complétion NIN / NIS",
    page_icon="🧩",
    layout="centered",
)
st.title("Complétion NIN / NIS depuis fichier officiel")

if not OFFICIAL_PATH.exists():
    st.error(
        f"Fichier officiel introuvable : `{OFFICIAL_PATH}`.\n\n"
        "Placez `NIN_A.XLS` dans le dossier `data/` à côté du script."
    )
    st.stop()

try:
    official, official_rows, official_dups = load_official(
        str(OFFICIAL_PATH), OFFICIAL_PATH.stat().st_mtime
    )
except Exception as e:
    st.error(f"Erreur lecture fichier officiel : {e}")
    st.stop()

st.caption(
    f"Référence : **{OFFICIAL_PATH.name}** — "
    f"{len(official)} comptes indexés sur {official_rows} lignes"
    + (f" (⚠️ {official_dups} doublons ignorés)" if official_dups else "")
)

uploaded = st.file_uploader(
    "Fichiers personnel (.xls / .xlsx)",
    type=["xls", "xlsx", "XLS", "XLSX"],
    accept_multiple_files=True,
)

if not uploaded:
    st.stop()

if not st.button("Compléter", type="primary", use_container_width=True):
    st.stop()

results: list[tuple[str, bytes]] = []

for up in uploaded:
    st.divider()
    st.subheader(up.name)
    try:
        df, method = load_dynamic(up.read(), up.name)
    except Exception as e:
        st.error(f"Lecture échouée : {e}")
        continue

    col_map = resolve_columns(df)
    if "NUMCPT" not in col_map:
        st.error(
            "Colonne NUMCPT introuvable. Colonnes détectées : "
            + ", ".join(f"`{c}`" for c in df.columns)
        )
        continue

    missing_targets = [f for f in FIELDS if f not in col_map]
    filled_df, stats, matched, empty_key, unmatched = fill_missing(df, col_map, official)

    total = len(df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Lignes", total)
    c2.metric("Appariées", matched)
    c3.metric("Non trouvées", len(unmatched))
    c4.metric("Clé vide", empty_key)

    st.markdown("**Cellules remplies :**")
    fill_cols = st.columns(len(FIELDS))
    for col, field in zip(fill_cols, FIELDS):
        if field in col_map:
            col.metric(field, stats[field])
        else:
            col.metric(field, "—", help="colonne absente du fichier")

    with st.expander("Colonnes reconnues"):
        st.json({canon: col_map.get(canon, "— absente —") for canon in COL_ALIASES})
    if missing_targets:
        st.warning(
            "Colonnes non trouvées (pas remplies) : "
            + ", ".join(f"`{f}`" for f in missing_targets)
        )
    if unmatched:
        with st.expander(f"Voir {len(unmatched)} NUMCPT non trouvés"):
            st.dataframe(
                pd.DataFrame(unmatched, columns=["ligne (index)", "NUMCPT brut"]),
                use_container_width=True,
            )

    st.caption(f"Méthode de lecture : {method}")

    xlsx_bytes = df_to_xlsx_bytes(filled_df)
    out_name = Path(up.name).stem + "_filled.xlsx"
    results.append((out_name, xlsx_bytes))
    st.download_button(
        label=f"⬇️  Télécharger {out_name}",
        data=xlsx_bytes,
        file_name=out_name,
        mime=XLSX_MIME,
        use_container_width=True,
        key=f"dl-{up.name}",
    )

if len(results) > 1:
    st.divider()
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in results:
            zf.writestr(name, data)
    zip_buf.seek(0)
    st.download_button(
        label=f"⬇️  Télécharger tous les fichiers ({len(results)}) en .zip",
        data=zip_buf,
        file_name="personnel_filled.zip",
        mime="application/zip",
        use_container_width=True,
    )
