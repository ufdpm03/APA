import io
import re
import streamlit as st
from docx import Document

st.set_page_config(page_title="APA Reference Checker", layout="wide")
st.title("APA Reference Checker (DOCX)")

# -----------------------------
# Citation detection (best-effort)
# -----------------------------
# Parenthetical patterns: (Smith, 2021) / (Smith & Jones, 2021) / (Smith et al., 2021)
PAREN_CIT_RE = re.compile(
    r"\((?P<authors>[^()]+?),\s*(?P<year>\d{4}|n\.d\.)\)"
)

# Narrative patterns: Smith (2021) / Smith and Jones (2021)
NARR_CIT_RE = re.compile(
    r"\b(?P<author>[A-Z][A-Za-z’'\-]+)\s*\((?P<year>\d{4}|n\.d\.)\)"
)

def normalize_year(y: str) -> str:
    y = (y or "").strip().lower()
    if y == "n.d.":
        return "n.d."
    return y if re.fullmatch(r"\d{4}", y) else "n.d."

def normalize_author_token(s: str) -> str:
    s = (s or "").strip()
    # Keep only first author token, remove punctuation
    s = re.sub(r"[^A-Za-z’'\- ]", "", s)
    s = s.strip()
    if not s:
        return "unknown"
    return s.split()[0].lower()

def extract_intext_citations(all_text: str):
    """
    Returns set of keys like (author_token, year)
    author_token is simplified first author last name / org token.
    """
    keys = set()

    # Parenthetical citations
    for m in PAREN_CIT_RE.finditer(all_text):
        authors = m.group("authors")
        year = normalize_year(m.group("year"))
        # handle "Smith & Jones" or "Smith et al."
        first = authors.split("&")[0]
        first = first.split("and")[0]
        first = first.split("et al.")[0]
        keys.add((normalize_author_token(first), year))

    # Narrative citations
    for m in NARR_CIT_RE.finditer(all_text):
        author = m.group("author")
        year = normalize_year(m.group("year"))
        keys.add((normalize_author_token(author), year))

    return keys

# -----------------------------
# Reference section parsing (best-effort)
# -----------------------------

def find_references_start(paragraphs):
    """
    Find the paragraph index where 'References' header begins.
    """
    for i, p in enumerate(paragraphs):
        if p.text.strip().lower() == "references":
            return i
    return None

def split_reference_entries(ref_paragraphs):
    """
    Very common in Word: each reference is its own paragraph.
    Sometimes references wrap across paragraphs; we join if paragraph
    doesn't look like it starts a new reference.

    Heuristic: New entry likely begins with:
      - "Lastname,"  OR
      - Organization name (Capitalized word) and (YEAR) soon after
    """
    entries = []
    current = ""

    def looks_like_new_entry(t: str) -> bool:
        t = t.strip()
        if not t:
            return False
        # "Lastname," at start
        if re.match(r"^[A-Z][A-Za-z’'\-]+,\s", t):
            return True
        # Organization-like: starts with caps word and has (YYYY) in first ~60 chars
        if re.match(r"^[A-Z]", t) and re.search(r"\(\s*(\d{4}|n\.d\.)\s*\)", t[:80]):
            return True
        return False

    for p in ref_paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if looks_like_new_entry(t):
            if current:
                entries.append(current.strip())
            current = t
        else:
            # continuation line
            current = (current + " " + t).strip()

    if current:
        entries.append(current.strip())

    return entries

def extract_reference_keys(reference_entries):
    """
    For each reference entry, extract:
      - first author token (or org token)
      - year
    Best-effort.
    """
    keys = set()
    detailed = []

    for entry in reference_entries:
        # Year in parentheses
        ym = re.search(r"\(\s*(\d{4}|n\.d\.)\s*\)", entry)
        year = normalize_year(ym.group(1) if ym else "n.d.")

        # First author token: "Lastname," at start
        am = re.match(r"^([A-Z][A-Za-z’'\-]+),\s", entry)
        if am:
            author_tok = normalize_author_token(am.group(1))
        else:
            # fall back: first word before "("
            pre = entry.split("(")[0]
            author_tok = normalize_author_token(pre)

        keys.add((author_tok, year))
        detailed.append({"entry": entry, "key": (author_tok, year)})

    return keys, detailed

# -----------------------------
# UI
# -----------------------------

uploaded = st.file_uploader("Upload a Word document (.docx)", type=["docx"])

colA, colB = st.columns([1, 1])

run_check = colA.button("Check citations vs References", type="primary")
show_debug = colB.checkbox("Show debug details", value=False)

if run_check:
    if not uploaded:
        st.error("Please upload a .docx file first.")
        st.stop()

    doc = Document(io.BytesIO(uploaded.getvalue()))
    paragraphs = list(doc.paragraphs)

    # Full text for citation scan
    full_text = "\n".join(p.text for p in paragraphs)

    # In-text citations
    cited_keys = extract_intext_citations(full_text)

    # References section
    ref_start = find_references_start(paragraphs)
    if ref_start is None:
        st.warning("No 'References' heading found. I can’t reliably check the reference list.")
        st.info("Tip: Add a heading that is exactly 'References' at the end of the document.")
        st.stop()

    ref_paragraphs = paragraphs[ref_start + 1 :]
    ref_entries = split_reference_entries(ref_paragraphs)
    ref_keys, ref_details = extract_reference_keys(ref_entries)

    # Compare
    cited_not_in_refs = sorted(list(cited_keys - ref_keys))
    refs_not_cited = sorted(list(ref_keys - cited_keys))

    # Basic metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("In-text citations found", len(cited_keys))
    m2.metric("Reference entries found", len(ref_entries))
    m3.metric("Potential issues", len(cited_not_in_refs) + len(refs_not_cited))

    st.divider()

    st.subheader("Results (best-effort)")

    if not cited_not_in_refs and not refs_not_cited:
        st.success("Looks consistent: every in-text citation appears to have a matching reference, and vice versa.")
    else:
        if cited_not_in_refs:
            st.error("Cited in text but NOT found in References (possible missing reference):")
            st.write(cited_not_in_refs)
            st.caption("Keys shown as (first-author-token, year). Example: ('smith', '2021')")

        if refs_not_cited:
            st.warning("In References but NOT cited in text (possible unused reference):")
            st.write(refs_not_cited)

    if show_debug:
        st.divider()
        st.subheader("Debug details")
        st.write("**Cited keys:**", sorted(list(cited_keys))[:300])
        st.write("**Reference entries (parsed):**")
        for r in ref_details[:200]:
            st.markdown(f"- `{r['key']}`  →  {r['entry']}")
