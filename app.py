# app.py
import io
import re
from typing import Dict, List, Tuple, Optional

import pandas as pd
import requests
import streamlit as st
import xml.etree.ElementTree as ET

# ---- Optional: interactive table with per-column filters (recommended) ----
AGGRID_AVAILABLE = True
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
except Exception:
    AGGRID_AVAILABLE = False


FEED_URL_DEFAULT = "https://platinumlist.net/xml-feed/partnership-program"


# ------------------------- Helpers: XML + encoding -------------------------
def clean_tag(tag: str) -> str:
    """Strip namespace if present: {ns}tag -> tag"""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


@st.cache_data(ttl=300)
def fetch_xml_bytes(url: str) -> Tuple[bytes, Dict[str, str], Optional[str], Optional[str]]:
    r = requests.get(
        url,
        timeout=45,
        headers={"User-Agent": "Mozilla/5.0 (Streamlit XML Dashboard)"},
    )
    r.raise_for_status()
    return r.content, dict(r.headers), r.encoding, getattr(r, "apparent_encoding", None)


def parse_xml_bytes(
    data: bytes,
    req_encoding: Optional[str] = None,
    apparent_encoding: Optional[str] = None,
) -> Tuple[ET.Element, str]:
    """
    Robust XML parsing:
    1) Try raw bytes (works when XML declares correct encoding)
    2) Detect encoding from XML declaration / HTTP / apparent
    3) Try common Cyrillic encodings (cp1251/windows-1251)
    Returns (root, used_encoding)
    """
    # Quick HTML guard (sometimes endpoints return HTML)
    head = data[:400].lstrip()
    if head.lower().startswith(b"<!doctype html") or head.lower().startswith(b"<html"):
        raise ValueError("Received HTML instead of XML (possible protection/redirect).")

    # 1) Best case: parse bytes directly
    try:
        root = ET.fromstring(data)
        return root, "bytes-as-is"
    except Exception:
        pass

    # 2) Try detect encoding from XML declaration
    enc_from_decl = None
    m = re.search(br'encoding=["\']([^"\']+)["\']', data[:500])
    if m:
        try:
            enc_from_decl = m.group(1).decode("ascii", errors="ignore").strip()
        except Exception:
            enc_from_decl = None

    # 3) Build candidates
    candidates = []
    for enc in [enc_from_decl, req_encoding, apparent_encoding]:
        if enc and enc not in candidates:
            candidates.append(enc)

    # Common Cyrillic fallbacks + safe fallbacks
    for enc in ["windows-1251", "cp1251", "utf-8", "iso-8859-1"]:
        if enc not in candidates:
            candidates.append(enc)

    last_err = None
    for enc in candidates:
        try:
            text = data.decode(enc)  # strict decode; will throw if invalid
            text = text.lstrip("\ufeff")  # drop BOM if any

            # Force UTF-8 for consistent downstream parsing
            if text.lstrip().startswith("<?xml"):
                text = re.sub(
                    r'encoding=["\']([^"\']+)["\']',
                    'encoding="utf-8"',
                    text,
                    count=1,
                    flags=re.IGNORECASE,
                )
            else:
                text = '<?xml version="1.0" encoding="utf-8"?>\n' + text

            root = ET.fromstring(text.encode("utf-8"))
            return root, enc
        except Exception as e:
            last_err = e

    raise last_err if last_err else ValueError("Failed to parse XML.")


# ------------------------- Flatten XML to rows -------------------------
def flatten_element(el: ET.Element, prefix: str = "") -> Dict[str, str]:
    """
    Flatten an element into a single row dict:
    - Leaf tags become columns
    - Attributes become columns with @
    - Nested structure becomes path-like columns: parent/child/grandchild
    - Repeated leaf tags are joined with " | "
    """
    row: Dict[str, str] = {}

    el_tag = clean_tag(el.tag)

    # Attributes of current node
    for k, v in (el.attrib or {}).items():
        row[f"{prefix}{el_tag}@{k}"] = str(v)

    children = list(el)
    if not children:
        key = f"{prefix}{el_tag}".rstrip("/")
        val = (el.text or "").strip()
        if key:
            row[key] = val
        return row

    # Children
    for ch in children:
        ch_tag = clean_tag(ch.tag)
        ch_children = list(ch)

        # Child attributes
        for ak, av in (ch.attrib or {}).items():
            row[f"{prefix}{el_tag}/{ch_tag}@{ak}"] = str(av)

        if ch_children:
            row.update(flatten_element(ch, prefix=f"{prefix}{el_tag}/"))
        else:
            key = f"{prefix}{el_tag}/{ch_tag}"
            val = (ch.text or "").strip()

            # If key repeats, join values
            if key in row and row[key]:
                row[key] = f"{row[key]} | {val}"
            else:
                row[key] = val

    return row


def find_repeating_tag_candidates(root: ET.Element, max_candidates: int = 12) -> List[Tuple[str, int]]:
    """
    Heuristic: find child tags that repeat under the same parent (common "row" pattern).
    Returns list of (tag_name, count) sorted desc.
    """
    counts: Dict[str, int] = {}
    for parent in root.iter():
        children = list(parent)
        if len(children) < 2:
            continue
        freq: Dict[str, int] = {}
        for ch in children:
            t = clean_tag(ch.tag)
            freq[t] = freq.get(t, 0) + 1
        for t, n in freq.items():
            if n >= 2:
                counts[t] = counts.get(t, 0) + n

    sorted_tags = sorted(counts.items(), key=lambda x: x[1], reverse=True)
    return sorted_tags[:max_candidates]


def collect_row_nodes(root: ET.Element, chosen_tag: Optional[str]) -> List[ET.Element]:
    """
    If chosen_tag provided -> return all nodes with that tag.
    Else -> choose best repeating tag; if none, fallback to direct children of root.
    """
    if chosen_tag:
        return [el for el in root.iter() if clean_tag(el.tag) == chosen_tag]

    candidates = find_repeating_tag_candidates(root)
    if candidates:
        best_tag = candidates[0][0]
        return [el for el in root.iter() if clean_tag(el.tag) == best_tag]

    return list(root)  # fallback


# ------------------------- Excel export -------------------------
def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()


# ------------------------- UI -------------------------
st.set_page_config(page_title="Platinumlist XML Dashboard", layout="wide")
st.title("Platinumlist XML → Table with filters and Excel export")

with st.sidebar:
    st.subheader("Source")
    url = st.text_input("XML URL", FEED_URL_DEFAULT)

    st.divider()
    st.subheader("Rows (what should be treated as a table row)")
    st.caption("Usually this is a repeating tag like offer/event/item. You can set it manually if needed.")
    manual_tag = st.text_input("Row tag (optional)", value="").strip() or None

    refresh = st.button("Reload (clear cache)", use_container_width=True)

if refresh:
    st.cache_data.clear()

# Fetch + parse
try:
    data, headers, req_enc, apparent_enc = fetch_xml_bytes(url)
    root, used_enc = parse_xml_bytes(data, req_enc, apparent_enc)
except Exception as e:
    st.error("XML load/parse error.")
    st.write(str(e))

    # Show a short response snippet for debugging (decoded safely)
    try:
        snippet = (data[:800]).decode("utf-8", errors="replace") if "data" in locals() else ""
        if snippet:
            st.caption("First ~800 bytes of the response (utf-8 with replacement):")
            st.code(snippet)
    except Exception:
        pass

    st.stop()

# Candidates hint
candidates = find_repeating_tag_candidates(root)
if candidates:
    with st.sidebar:
        st.caption("Detected repeating tag candidates:")
        st.write(", ".join([f"{t} ({c})" for t, c in candidates]))

# Collect rows
row_nodes = collect_row_nodes(root, manual_tag)

if not row_nodes:
    st.warning("No row nodes found (no repeating tag detected and root has no direct children).")
    st.stop()

# Flatten to DataFrame
rows = [flatten_element(n, prefix="") for n in row_nodes]
df = pd.DataFrame(rows).fillna("")

# Drop fully empty columns (helps readability)
non_empty_cols = [c for c in df.columns if (df[c].astype(str).str.len() > 0).any()]
df = df[non_empty_cols] if non_empty_cols else df

st.caption(f"Rows: {len(df)} • Columns: {len(df.columns)} • Encoding attempt: {used_enc}")

# Show table + filters
df_filtered = df

if AGGRID_AVAILABLE:
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, resizable=True, wrapText=True, autoHeight=True)
    gb.configure_grid_options(domLayout="normal")
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=50)
    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=False,
        height=650,
        allow_unsafe_jscode=False,
    )
    df_filtered = grid_response.get("data", df)
else:
    st.warning(
        "streamlit-aggrid is not installed, so the table is shown without Excel-like column filters. "
        "Add streamlit-aggrid to requirements.txt to enable per-column filtering."
    )
    q = st.text_input("Search across all fields (fallback)", value="")
    if q.strip():
        ql = q.strip().lower()
        mask = df.apply(lambda r: r.astype(str).str.lower().str.contains(ql, na=False).any(), axis=1)
        df_filtered = df[mask].copy()

    st.dataframe(df_filtered, use_container_width=True, height=650)

# Downloads
c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "Download FULL (Excel)",
        data=df_to_xlsx_bytes(df),
        file_name="platinumlist_xml_full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c2:
    st.download_button(
        "Download FILTERED (Excel)",
        data=df_to_xlsx_bytes(pd.DataFrame(df_filtered)),
        file_name="platinumlist_xml_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
