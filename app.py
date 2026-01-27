import io
import json
import re
from datetime import date
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ============================
# APA helpers (v1)
# ============================

REF_TYPES = [
    "Journal article",
    "Book",
    "Book chapter",
    "Website",
    "Report / Government document",
    "Conference paper",
    "Thesis / Dissertation",
]

CITATION_PAREN_RE = re.compile(r"\(([A-Za-z][A-Za-z\-’' ]+?),\s*(\d{4}|n\.d\.)\)")
CITATION_NARR_RE = re.compile(r"\b([A-Z][A-Za-z\-’']+)\s*\((\d{4}|n\.d\.)\)")

def _strip(s): return (s or "").strip()

def sentence_case(title: str) -> str:
    t = _strip(title)
    if not t:
        return t
    # simple sentence-case: first char upper, rest lower
    return t[:1].upper() + t[1:].lower()

def title_case_keep(title: str) -> str:
    # For journals/books in APA: generally title case; we won't enforce aggressively.
    return _strip(title)

def normalize_year(y):
    if y is None:
        return "n.d."
    if isinstance(y, int):
        return str(y)
    y = _strip(str(y))
    if not y:
        return "n.d."
    if y.lower() == "n.d.":
        return "n.d."
    m = re.match(r"^\d{4}$", y)
    return y if m else "n.d."

def format_author_list_apa(authors):
    """
    authors can be dict: {"is_org": True, "org_name": "..."}
    or list of {"last": "...", "initials": "J. A.", "suffix": ""}
    """
    if isinstance(authors, dict) and authors.get("is_org"):
        return _strip(authors.get("org_name"))

    if not authors:
        return ""

    formatted = []
    for a in authors:
        last = _strip(a.get("last"))
        initials = _strip(a.get("initials"))
        suffix = _strip(a.get("suffix"))
        name = f"{last}, {initials}".strip().strip(",")
        if suffix:
            name = f"{name}, {suffix}"
        formatted.append(name)

    # APA: up to 20 authors; join with commas, ampersand before last
    if len(formatted) == 1:
        return formatted[0]
    if len(formatted) == 2:
        return f"{formatted[0]}, & {formatted[1]}"
    return ", ".join(formatted[:-1]) + f", & {formatted[-1]}"

def first_author_key(authors):
    if isinstance(authors, dict) and authors.get("is_org"):
        return _strip(authors.get("org_name")).lower() or "unknown"
    if not authors:
        return "unknown"
    return _strip(authors[0].get("last")).lower() or "unknown"

def intext_parenthetical(authors, year):
    y = normalize_year(year)
    if isinstance(authors, dict) and authors.get("is_org"):
        org = _strip(authors.get("org_name")) or "Unknown"
        return f"({org}, {y})"
    if not authors:
        return f"(Unknown, {y})"
    if len(authors) == 1:
        return f"({_strip(authors[0].get('last')) or 'Unknown'}, {y})"
    if len(authors) == 2:
        return f"({_strip(authors[0].get('last')) or 'Unknown'} & {_strip(authors[1].get('last')) or 'Unknown'}, {y})"
    return f"({_strip(authors[0].get('last')) or 'Unknown'} et al., {y})"

def intext_narrative(authors, year):
    y = normalize_year(year)
    if isinstance(authors, dict) and authors.get("is_org"):
        org = _strip(authors.get("org_name")) or "Unknown"
        return f"{org} ({y})"
    if not authors:
        return f"Unknown ({y})"
    if len(authors) == 1:
        return f"{_strip(authors[0].get('last')) or 'Unknown'} ({y})"
    if len(authors) == 2:
        return f"{_strip(authors[0].get('last')) or 'Unknown'} and {_strip(authors[1].get('last')) or 'Unknown'} ({y})"
    return f"{_strip(authors[0].get('last')) or 'Unknown'} et al. ({y})"

def apa_reference_string(ref):
    """
    Produces a decent APA 7 reference string for common types.
    This is not a perfect APA engine (APA has many edge cases),
    but it's strong enough for real use and easy to refine.
    """
    rtype = ref.get("type")
    authors = ref.get("authors")
    year = normalize_year(ref.get("year"))
    title = _strip(ref.get("title"))
    if ref.get("auto_sentence_case", True):
        title = sentence_case(title)

    doi = _strip(ref.get("doi"))
    url = _strip(ref.get("url"))

    author_str = format_author_list_apa(authors)
    base = f"{author_str} ({year}). {title}."

    if rtype == "Journal article":
        journal = title_case_keep(ref.get("journal"))
        volume = _strip(ref.get("volume"))
        issue = _strip(ref.get("issue"))
        pages = _strip(ref.get("pages"))
        issue_part = f"({issue})" if issue else ""
        jpart = f" {journal}, {volume}{issue_part}, {pages}."
        tail = f" {doi}" if doi else (f" {url}" if url else "")
        return (base + jpart + tail).replace("..", ".").strip()

    if rtype == "Book":
        publisher = _strip(ref.get("publisher"))
        edition = _strip(ref.get("edition"))
        ed_part = f" ({edition} ed.)." if edition else ""
        tail = f" {doi}" if doi else (f" {url}" if url else "")
        return f"{author_str} ({year}). {title}.{ed_part} {publisher}.{tail}".replace("..", ".").strip()

    if rtype == "Book chapter":
        chapter_title = title
        book_title = title_case_keep(ref.get("book_title"))
        editors = ref.get("editors") or []
        editors_str = format_author_list_apa(editors)
        pages = _strip(ref.get("pages"))
        publisher = _strip(ref.get("publisher"))
        tail = f" {doi}" if doi else (f" {url}" if url else "")
        # Simplified: In E. E. Editor (Ed.), Book title (pp. x–y). Publisher.
        ed_label = "(Ed.)" if len(editors) == 1 else "(Eds.)" if len(editors) > 1 else ""
        in_part = f"In {editors_str} {ed_label}, {book_title}"
        if pages:
            in_part += f" (pp. {pages})"
        in_part += f". {publisher}."
        return f"{author_str} ({year}). {chapter_title}. {in_part}{tail}".replace("..", ".").strip()

    if rtype == "Website":
        site = _strip(ref.get("site_name"))
        use_ret = bool(ref.get("use_retrieval_date"))
        retrieval_date = _strip(ref.get("retrieval_date"))
        ret_part = f" Retrieved {retrieval_date}," if (use_ret and retrieval_date) else ""
        tail = url or doi
        tail = f" {tail}" if tail else ""
        # If author same as site, APA sometimes omits site; we won't overdo it in v1.
        return f"{author_str} ({year}). {title}. {site}.{ret_part}{tail}".replace("..", ".").strip()

    if rtype == "Report / Government document":
        publisher = _strip(ref.get("publisher"))
        report_no = _strip(ref.get("report_number"))
        rep_part = f" (Report No. {report_no})." if report_no else ""
        tail = f" {url}" if url else (f" {doi}" if doi else "")
        return f"{author_str} ({year}). {title}.{rep_part} {publisher}.{tail}".replace("..", ".").strip()

    if rtype == "Conference paper":
        conf = _strip(ref.get("conference_name"))
        loc = _strip(ref.get("conference_location"))
        loc_part = f", {loc}" if loc else ""
        tail = f" {url}" if url else (f" {doi}" if doi else "")
        return f"{author_str} ({year}). {title}. {conf}{loc_part}.{tail}".replace("..", ".").strip()

    if rtype == "Thesis / Dissertation":
        uni = _strip(ref.get("university"))
        ttype = _strip(ref.get("thesis_type")) or "Dissertation"
        tail = f" {url}" if url else ""
        return f"{author_str} ({year}). {title} [{ttype}, {uni}].{tail}".replace("..", ".").strip()

    tail = f" {doi}" if doi else (f" {url}" if url else "")
    return (base + tail).replace("..", ".").strip()

# ============================
# DOCX formatting helpers
# ============================

def set_document_margins(doc: Document, inches=1.0):
    for section in doc.sections:
        section.top_margin = Inches(inches)
        section.bottom_margin = Inches(inches)
        section.left_margin = Inches(inches)
        section.right_margin = Inches(inches)

def set_default_font(doc: Document, font_name: str, font_size_pt: int):
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)
    # Ensure East Asia font set too (Word quirk)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

def set_double_spacing_and_first_line_indent(doc: Document, indent_inches=0.5):
    for p in doc.paragraphs:
        # Skip empty paragraphs
        p.paragraph_format.line_spacing = 2.0
        # First-line indent for body paragraphs (avoid for headings)
        if p.style and p.style.name and "Heading" not in p.style.name:
            p.paragraph_format.first_line_indent = Inches(indent_inches)

def clear_headers_footers(doc: Document):
    for section in doc.sections:
        header = section.header
        footer = section.footer
        for p in header.paragraphs:
            p.text = ""
        for p in footer.paragraphs:
            p.text = ""

def add_page_number_top_right(section):
    """
    Adds a PAGE field in header aligned right.
    """
    header = section.header
    if not header.paragraphs:
        p = header.add_paragraph()
    else:
        p = header.paragraphs[0]

    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # clear existing runs
    for r in p.runs:
        r.text = ""

    run = p.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = " PAGE "

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

def insert_student_title_page(doc: Document, tp: dict):
    """
    Simple APA 7 student title page:
    centered, double-spaced:
    Title
    Author
    Institution
    Course
    Instructor
    Due date
    """
    # Insert at start: create a new doc-like block by adding paragraphs at top.
    # python-docx doesn't support "insert paragraph at very top" cleanly,
    # so we prepend by creating a new Document in output step (we'll do that).
    pass

def find_references_heading_idx(doc: Document):
    for i, p in enumerate(doc.paragraphs):
        if _strip(p.text).lower() == "references":
            return i
    return None

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def rebuild_references_section(doc: Document, formatted_refs):
    """
    Ensures there is a "References" heading and then inserts references,
    hanging indent, double-spaced, alphabetized already.
    """
    idx = find_references_heading_idx(doc)

    # If a References heading exists, delete everything after it to the end
    if idx is not None:
        # delete paragraphs after heading
        for p in list(doc.paragraphs[idx+1:]):
            delete_paragraph(p)
        heading_p = doc.paragraphs[idx]
        heading_p.style = doc.styles["Normal"]
        heading_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading_p.text = "References"
        # insert refs after by just appending new paragraphs (since we deleted to end)
        for ref in formatted_refs:
            rp = doc.add_paragraph(ref)
            rp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            rp.paragraph_format.line_spacing = 2.0
            rp.paragraph_format.left_indent = Inches(0.5)
            rp.paragraph_format.first_line_indent = Inches(-0.5)
    else:
        # Add at end
        doc.add_page_break()
        hp = doc.add_paragraph("References")
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.line_spacing = 2.0

        for ref in formatted_refs:
            rp = doc.add_paragraph(ref)
            rp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            rp.paragraph_format.line_spacing = 2.0
            rp.paragraph_format.left_indent = Inches(0.5)
            rp.paragraph_format.first_line_indent = Inches(-0.5)

def scan_docx_for_citations(doc: Document):
    """
    Returns:
      - set of (author_token, year) from parenthetical citations
      - set of (author_token, year) from narrative citations
    Very approximate; good enough for a helpful report.
    """
    text = "\n".join(p.text for p in doc.paragraphs)
    paren = set((m.group(1).strip(), m.group(2)) for m in CITATION_PAREN_RE.finditer(text))
    narr = set((m.group(1).strip(), m.group(2)) for m in CITATION_NARR_RE.finditer(text))
    return paren, narr

def reference_key_guess(ref):
    """
    Makes a simple key that matches the scan logic:
    - first author's last name (or org)
    - year
    """
    authors = ref.get("authors")
    year = normalize_year(ref.get("year"))
    if isinstance(authors, dict) and authors.get("is_org"):
        name = _strip(authors.get("org_name")) or "Unknown"
        # take first word
        token = name.split()[0]
    elif isinstance(authors, list) and authors:
        token = _strip(authors[0].get("last")) or "Unknown"
    else:
        token = "Unknown"
    return (token, year)

# ============================
# Streamlit state init
# ============================

def init_state():
    if "refs" not in st.session_state:
        st.session_state.refs = []
    if "edit_idx" not in st.session_state:
        st.session_state.edit_idx = None
    if "uploaded_bytes" not in st.session_state:
        st.session_state.uploaded_bytes = None
    if "scan_results" not in st.session_state:
        st.session_state.scan_results = None
    if "paper_settings" not in st.session_state:
        st.session_state.paper_settings = {
            "paper_type": "Student",
            "font": "Times New Roman 12",
            "line_spacing": "Double",
            "running_head": False,
            "title_page": {
                "title": "",
                "author": "",
                "institution": "",
                "course": "",
                "instructor": "",
                "due_date": "",
            },
        }

init_state()

# ============================
# UI
# ============================

st.set_page_config(page_title="APA Formatter (OMSA)", layout="wide")
st.title("APA Formatter App (OMSA)")

# Sidebar settings
st.sidebar.header("Paper Settings")
paper_type = st.sidebar.radio("Paper type", ["Student", "Professional"], index=0)
st.session_state.paper_settings["paper_type"] = paper_type

font_choice = st.sidebar.selectbox("Font", ["Times New Roman 12", "Calibri 11", "Arial 11"], index=0)
st.session_state.paper_settings["font"] = font_choice

if paper_type == "Professional":
    running_head = st.sidebar.toggle("Running head (basic)", value=st.session_state.paper_settings.get("running_head", False))
else:
    running_head = False
st.session_state.paper_settings["running_head"] = running_head

with st.sidebar.expander("Title page fields", expanded=False):
    tp = st.session_state.paper_settings["title_page"]
    tp["title"] = st.text_input("Title", value=tp.get("title", ""))
    tp["author"] = st.text_input("Author", value=tp.get("author", ""))
    tp["institution"] = st.text_input("Institution", value=tp.get("institution", ""))
    tp["course"] = st.text_input("Course", value=tp.get("course", ""))
    tp["instructor"] = st.text_input("Instructor", value=tp.get("instructor", ""))
    tp["due_date"] = st.text_input("Due date", value=tp.get("due_date", ""))

st.sidebar.divider()
cA, cB = st.sidebar.columns(2)
if cA.button("New project"):
    st.session_state.refs = []
    st.session_state.edit_idx = None
    st.session_state.uploaded_bytes = None
    st.session_state.scan_results = None
    st.toast("New project created.", icon="🧹")

st.sidebar.caption(f"References: **{len(st.session_state.refs)}**")
st.sidebar.divider()

# Import/Export refs
st.sidebar.subheader("Import / Export references")
st.sidebar.download_button(
    "Export references (.json)",
    data=json.dumps(st.session_state.refs, indent=2),
    file_name="references.json",
    mime="application/json",
)

imp = st.sidebar.file_uploader("Import references (.json)", type=["json"])
if imp is not None:
    try:
        data = json.loads(imp.getvalue().decode("utf-8"))
        if isinstance(data, list):
            st.session_state.refs = data
            st.toast(f"Imported {len(data)} references.", icon="✅")
        else:
            st.sidebar.error("JSON must be a list.")
    except Exception as e:
        st.sidebar.error(f"Import failed: {e}")

tab1, tab2, tab3 = st.tabs(["1) Upload & Scan", "2) References Manager", "3) Format & Export"])

# ============================
# Tab 1: Upload & Scan
# ============================

with tab1:
    st.subheader("Upload your paper (.docx)")

    up = st.file_uploader("Upload .docx", type=["docx"])
    if up is not None:
        st.session_state.uploaded_bytes = up.getvalue()
        st.success("Uploaded.")
        st.session_state.scan_results = None

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Scan document"):
            if not st.session_state.uploaded_bytes:
                st.warning("Upload a .docx first.")
            else:
                doc = Document(io.BytesIO(st.session_state.uploaded_bytes))
                refs_idx = find_references_heading_idx(doc)
                paren, narr = scan_docx_for_citations(doc)
                st.session_state.scan_results = {
                    "references_heading_found": refs_idx is not None,
                    "paren_citations": sorted(list(paren)),
                    "narr_citations": sorted(list(narr)),
                }
                st.toast("Scan complete.", icon="🔎")

    with col2:
        if st.button("Clear upload"):
            st.session_state.uploaded_bytes = None
            st.session_state.scan_results = None
            st.toast("Cleared.", icon="🗑️")

    if st.session_state.scan_results:
        sr = st.session_state.scan_results
        st.markdown("### Scan results")
        a, b, c = st.columns(3)
        a.metric("References heading found", "Yes" if sr["references_heading_found"] else "No")
        b.metric("Parenthetical citations", len(sr["paren_citations"]))
        c.metric("Narrative citations", len(sr["narr_citations"]))

        with st.expander("Show citations found"):
            st.write("**Parenthetical:**")
            st.write(sr["paren_citations"][:200])
            st.write("**Narrative:**")
            st.write(sr["narr_citations"][:200])

# ============================
# Tab 2: References Manager
# ============================

def ref_short_label(ref):
    k = reference_key_guess(ref)
    title = _strip(ref.get("title"))[:60] or "(no title)"
    return f"{k[0]}, {k[1]} — {title}"

def reference_editor(default=None):
    if default is None:
        default = {
            "type": "Journal article",
            "authors": [{"last": "", "initials": "", "suffix": ""}],
            "org_author": False,
            "org_name": "",
            "year": "",
            "title": "",
            "auto_sentence_case": True,
            "journal": "",
            "volume": "",
            "issue": "",
            "pages": "",
            "publisher": "",
            "edition": "",
            "book_title": "",
            "chapter_title": "",
            "editors": [{"last": "", "initials": "", "suffix": ""}],
            "site_name": "",
            "use_retrieval_date": False,
            "retrieval_date": str(date.today()),
            "report_number": "",
            "conference_name": "",
            "conference_location": "",
            "university": "",
            "thesis_type": "Dissertation",
            "doi": "",
            "url": "",
        }

    with st.form("ref_form", clear_on_submit=False):
        st.markdown("### Add / Edit Reference")
        rtype = st.selectbox("Reference type", REF_TYPES, index=REF_TYPES.index(default.get("type", "Journal article")))
        org_author = st.checkbox("Organization as author", value=bool(default.get("org_author", False)))

        # Authors
        authors_payload = None
        if org_author:
            org_name = st.text_input("Organization name", value=default.get("org_name", ""))
            authors_payload = {"is_org": True, "org_name": org_name}
        else:
            n_authors = st.number_input("Number of authors", min_value=1, max_value=20,
                                        value=max(1, len(default.get("authors") or [])), step=1)
            authors = default.get("authors") or [{"last": "", "initials": "", "suffix": ""}]
            if len(authors) < n_authors:
                authors += [{"last": "", "initials": "", "suffix": ""} for _ in range(n_authors - len(authors))]
            if len(authors) > n_authors:
                authors = authors[:n_authors]

            for i in range(n_authors):
                st.markdown(f"**Author {i+1}**")
                c1, c2, c3 = st.columns([2, 2, 1])
                authors[i]["last"] = c1.text_input(f"Last name {i+1}", value=authors[i].get("last", ""), key=f"last_{i}")
                authors[i]["initials"] = c2.text_input(f"Initials {i+1} (e.g., J. A.)", value=authors[i].get("initials", ""), key=f"init_{i}")
                authors[i]["suffix"] = c3.selectbox(f"Suffix {i+1}", ["", "Jr.", "Sr.", "II", "III", "IV"],
                                                    index=["", "Jr.", "Sr.", "II", "III", "IV"].index(authors[i].get("suffix", "") or ""),
                                                    key=f"suf_{i}")
            authors_payload = authors

        year = st.text_input("Year (or n.d.)", value=str(default.get("year", "")))
        title = st.text_input("Title", value=default.get("title", ""))
        auto_sc = st.checkbox("Auto sentence-case title", value=bool(default.get("auto_sentence_case", True)))

        st.markdown("### Type-specific fields")
        doi = st.text_input("DOI (optional)", value=default.get("doi", ""))
        url = st.text_input("URL (optional)", value=default.get("url", ""))

        fields = {}
        if rtype == "Journal article":
            j1, j2 = st.columns(2)
            fields["journal"] = j1.text_input("Journal title", value=default.get("journal", ""))
            fields["volume"] = j2.text_input("Volume", value=default.get("volume", ""))
            j3, j4 = st.columns(2)
            fields["issue"] = j3.text_input("Issue (optional)", value=default.get("issue", ""))
            fields["pages"] = j4.text_input("Pages (e.g., 123–145)", value=default.get("pages", ""))

        elif rtype == "Book":
            b1, b2 = st.columns(2)
            fields["publisher"] = b1.text_input("Publisher", value=default.get("publisher", ""))
            fields["edition"] = b2.text_input("Edition (optional, e.g., 2nd)", value=default.get("edition", ""))

        elif rtype == "Book chapter":
            fields["chapter_title"] = st.text_input("Chapter title", value=default.get("chapter_title", ""))
            fields["book_title"] = st.text_input("Book title", value=default.get("book_title", ""))
            c1, c2 = st.columns(2)
            fields["publisher"] = c1.text_input("Publisher", value=default.get("publisher", ""))
            fields["pages"] = c2.text_input("Chapter pages (e.g., 12–34)", value=default.get("pages", ""))

            st.markdown("**Editors (optional)**")
            n_edit = st.number_input("Number of editors", min_value=0, max_value=10, value=0, step=1)
            editors = [{"last": "", "initials": "", "suffix": ""} for _ in range(n_edit)]
            for i in range(n_edit):
                e1, e2, e3 = st.columns([2,2,1])
                editors[i]["last"] = e1.text_input(f"Editor last {i+1}", key=f"ed_last_{i}")
                editors[i]["initials"] = e2.text_input(f"Editor initials {i+1}", key=f"ed_init_{i}")
                editors[i]["suffix"] = e3.selectbox(f"Editor suffix {i+1}", ["", "Jr.", "Sr.", "II", "III", "IV"], key=f"ed_suf_{i}")
            fields["editors"] = editors

        elif rtype == "Website":
            fields["site_name"] = st.text_input("Site name", value=default.get("site_name", ""))
            use_ret = st.checkbox("Include retrieval date (for changing content)", value=bool(default.get("use_retrieval_date", False)))
            fields["use_retrieval_date"] = use_ret
            if use_ret:
                fields["retrieval_date"] = st.text_input("Retrieval date (YYYY-MM-DD)", value=default.get("retrieval_date", str(date.today())))

        elif rtype == "Report / Government document":
            r1, r2 = st.columns(2)
            fields["publisher"] = r1.text_input("Agency / Publisher (if different from author)", value=default.get("publisher", ""))
            fields["report_number"] = r2.text_input("Report number (optional)", value=default.get("report_number", ""))

        elif rtype == "Conference paper":
            c1, c2 = st.columns(2)
            fields["conference_name"] = c1.text_input("Conference name", value=default.get("conference_name", ""))
            fields["conference_location"] = c2.text_input("Location (optional)", value=default.get("conference_location", ""))

        elif rtype == "Thesis / Dissertation":
            t1, t2 = st.columns(2)
            fields["university"] = t1.text_input("University", value=default.get("university", ""))
            fields["thesis_type"] = t2.selectbox("Type", ["Thesis", "Dissertation"],
                                                 index=["Thesis", "Dissertation"].index(default.get("thesis_type", "Dissertation")))

        # Build object
        yr = year.strip()
        yr_val = int(yr) if re.fullmatch(r"\d{4}", yr) else (None if yr.lower() == "n.d." or yr == "" else None)
        ref_obj = {
            "type": rtype,
            "authors": authors_payload,
            "org_author": org_author,
            "org_name": default.get("org_name", "") if not org_author else (authors_payload.get("org_name", "") if isinstance(authors_payload, dict) else ""),
            "year": yr_val if yr_val is not None else ("n.d." if yr.lower() == "n.d." else None),
            "title": title.strip(),
            "auto_sentence_case": auto_sc,
            "doi": doi.strip(),
            "url": url.strip(),
        }
        ref_obj.update(fields)

        # Previews
        st.markdown("### Live previews")
        st.text_area("APA reference preview", value=apa_reference_string(ref_obj), height=110)
        st.code(intext_parenthetical(ref_obj["authors"], ref_obj["year"]))
        st.code(intext_narrative(ref_obj["authors"], ref_obj["year"]))

        save = st.form_submit_button("Save reference")
        save_add = st.form_submit_button("Save & add another")
        if save or save_add:
            return ref_obj, save_add

    return None, False

with tab2:
    left, right = st.columns([1.2, 1.8], gap="large")

    with left:
        st.subheader("Your References")
        q = st.text_input("Search")
        ftype = st.selectbox("Filter", ["All"] + REF_TYPES, index=0)

        refs = st.session_state.refs

        def matches(ref):
            if ftype != "All" and ref.get("type") != ftype:
                return False
            if not q:
                return True
            return q.lower() in json.dumps(ref).lower()

        filtered = [r for r in refs if matches(r)]
        st.caption(f"Showing {len(filtered)} of {len(refs)}")

        for i, ref in enumerate(filtered):
            real_idx = refs.index(ref)
            with st.container(border=True):
                st.markdown(f"**{ref_short_label(ref)}**")
                st.caption(ref.get("type", ""))
                b1, b2, b3 = st.columns(3)
                if b1.button("Edit", key=f"edit_{real_idx}"):
                    st.session_state.edit_idx = real_idx
                if b2.button("Duplicate", key=f"dup_{real_idx}"):
                    st.session_state.refs.insert(real_idx + 1, json.loads(json.dumps(ref)))
                    st.toast("Duplicated.", icon="📄")
                    st.rerun()
                if b3.button("Delete", key=f"del_{real_idx}"):
                    st.session_state.refs.pop(real_idx)
                    st.session_state.edit_idx = None
                    st.toast("Deleted.", icon="🗑️")
                    st.rerun()

        st.divider()
        if st.button("Add reference"):
            st.session_state.edit_idx = "NEW"

    with right:
        mode = st.session_state.edit_idx
        if mode is None:
            st.subheader("Add or edit a reference")
            st.write("Click **Add reference** or **Edit** on the left.")
        else:
            default = None if mode == "NEW" else st.session_state.refs[mode]
            st.subheader("Add new reference" if mode == "NEW" else f"Editing reference #{mode+1}")

            ref_obj, add_another = reference_editor(default)

            if ref_obj:
                if mode == "NEW":
                    st.session_state.refs.append(ref_obj)
                    st.toast("Saved new reference.", icon="✅")
                else:
                    st.session_state.refs[mode] = ref_obj
                    st.toast("Saved changes.", icon="✅")

                st.session_state.edit_idx = "NEW" if add_another else None
                st.rerun()

# ============================
# Tab 3: Format & Export
# ============================

with tab3:
    st.subheader("Format & Export your APA paper")

    opt_margins = st.checkbox("Apply APA margins + font + double-spacing", value=True)
    opt_pagenum = st.checkbox("Add page numbers (top-right)", value=True)
    opt_refs = st.checkbox("Create/replace References section using saved references", value=True)
    opt_alpha = st.checkbox("Alphabetize references", value=True)
    opt_hanging = st.checkbox("Hanging indent in references", value=True)
    run_checks = st.checkbox("Generate citation/reference report", value=True)

    st.divider()

    if st.button("Build formatted DOCX"):
        if not st.session_state.uploaded_bytes:
            st.error("Upload a .docx first (Tab 1).")
        else:
            # Load
            doc = Document(io.BytesIO(st.session_state.uploaded_bytes))

            # Apply formatting
            if opt_margins:
                set_document_margins(doc, 1.0)
                if "Times New Roman" in st.session_state.paper_settings["font"]:
                    set_default_font(doc, "Times New Roman", 12)
                elif "Calibri" in st.session_state.paper_settings["font"]:
                    set_default_font(doc, "Calibri", 11)
                else:
                    set_default_font(doc, "Arial", 11)

                set_double_spacing_and_first_line_indent(doc, indent_inches=0.5)

            # Page numbers
            if opt_pagenum:
                for section in doc.sections:
                    add_page_number_top_right(section)

            # References rebuild
            report = {"missing_in_references": [], "never_cited": []}
            if opt_refs:
                # Build formatted refs
                refs = st.session_state.refs[:]
                if opt_alpha:
                    refs.sort(key=lambda r: (first_author_key(r.get("authors")), normalize_year(r.get("year"))))
                formatted = [apa_reference_string(r) for r in refs]

                rebuild_references_section(doc, formatted)

                if run_checks:
                    paren, narr = scan_docx_for_citations(doc)
                    cited = set((a, y) for (a, y) in paren) | set((a, y) for (a, y) in narr)
                    ref_keys = set(reference_key_guess(r) for r in refs)

                    # missing citations
                    missing = []
                    for (a, y) in cited:
                        # compare only first token of author string
                        a_token = a.split()[0]
                        if (a_token, y) not in ref_keys:
                            missing.append((a, y))
                    # never cited refs
                    never = []
                    for rk in ref_keys:
                        if rk not in set((c[0].split()[0], c[1]) for c in cited):
                            never.append(rk)

                    report["missing_in_references"] = sorted(list(set(missing)))
                    report["never_cited"] = sorted(list(set(never)))

            # Output to bytes
            out = io.BytesIO()
            doc.save(out)
            out.seek(0)

            st.success("Formatted DOCX generated.")
            st.download_button(
                "Download formatted paper (.docx)",
                data=out.getvalue(),
                file_name="paper_APA_formatted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            if run_checks and opt_refs:
                st.markdown("### Citation / Reference Report (best-effort)")
                st.write("**Citations found in paper but missing from your reference library:**")
                st.write(report["missing_in_references"] if report["missing_in_references"] else "None found ✅")
                st.write("**References in your library that were never cited in text:**")
                st.write(report["never_cited"] if report["never_cited"] else "None found ✅")

    st.divider()
    st.markdown("### Current reference list preview")
    if st.session_state.refs:
        refs = st.session_state.refs[:]
        refs.sort(key=lambda r: (first_author_key(r.get("authors")), normalize_year(r.get("year"))))
        st.text_area("References (APA v1)", value="\n\n".join(apa_reference_string(r) for r in refs), height=260)
    else:
        st.caption("No references yet.")
