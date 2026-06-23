import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd
from docx2python import docx2python

# =========================================================
# DOCX2PYTHON-ONLY PARSER
# =========================================================
# This parser avoids python-docx / docx completely.
# It uses docx2python, like the original parser.
#
# Output columns:
# Name | Reference | Description | Parent | Authority document |
# Supplemental guidance | Active
# =========================================================
OUTPUT_COLUMNS = [
    "Name",
    "Reference",
    "Description",
    "Parent",
    "Authority document",
    "Supplemental guidance",
    "Active",
]

# =========================================================
# BASIC HELPERS
# =========================================================
def extract_text_from_docx(path: str) -> str:
    """
    Extract text from a .docx using docx2python.
    """
    doc = docx2python(path)
    return doc.text

def clean_text(text: str) -> str:
    """
    Normalize whitespace.
    """
    if text is None:
        return ""
    text = str(text)
    text = text.replace("\\r\\n", "\\n").replace("\\r", "\\n").replace("\\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def remove_inline_artifacts(text: str) -> str:
    """
    Remove artifacts produced by docx2python extraction:
    - media placeholders
    - footnote placeholders
    - simple HTML link wrappers
    """
    txt = clean_text(text)
    # Media / image placeholders
    txt = re.sub(r"----media/.*?----", "", txt, flags=re.I)
    txt = re.sub(r"----Image alt text---->?", "", txt, flags=re.I)
    txt = re.sub(r"Description automatically generated", "", txt, flags=re.I)
    # Footnote / endnote placeholders
    txt = re.sub(r"----footnote\d+----", "", txt, flags=re.I)
    txt = re.sub(r"footnote\d+\)", "", txt, flags=re.I)
    txt = re.sub(r"endnote\d+\)", "", txt, flags=re.I)
    # Remove simple HTML tags while keeping visible text
    # Example:
    # <a href="mailto:test@test.com">test@test.com</a>
    # becomes:
    # test@test.com
    txt = re.sub(r"<a\b[^>]*>", "", txt, flags=re.I)
    txt = re.sub(r"</a>", "", txt, flags=re.I)
    # Remove empty markdown-ish remnants
    txt = txt.replace("[]", "")
    return clean_text(txt)

def is_noise_line(text: str) -> bool:
    """
    Identify extraction artifacts that should not become rows.
    """
    txt = clean_text(text)
    if not txt:
        return True
    if re.search(r"----media/.*?----", txt, flags=re.I):
        return True
    if txt.startswith("----Image alt text----"):
        return True
    if txt.lower().startswith("description automatically generated"):
        return True
    if txt.startswith("![]") or txt.startswith("[][image_"):
        return True
    if re.fullmatch(r"_?\[?Internal Use Only\]?_?", txt, flags=re.I):
        return True
    return False

def split_lines(raw_text: str) -> List[str]:
    """
    Split docx2python output into clean, non-empty lines.
    """
    lines = [remove_inline_artifacts(x) for x in raw_text.splitlines()]
    return [x for x in lines if x and not is_noise_line(x)]

def pad_reference(ref: str) -> str:
    """
    Pad numeric reference parts:
    3.2.1 -> 03.02.01
    3.2.1.a -> 03.02.01.a
    """
    if not ref:
        return ""
    parts = []
    for token in str(ref).split("."):
        if token.isdigit():
            parts.append(token.zfill(2))
        else:
            parts.append(token)
    return ".".join(parts)

def make_row(
    name: str,
    reference: str,
    description: str,
    parent: str,
    authority_document: str,
    supplemental_guidance: str = "",
    active: str = "FALSE",
) -> Dict[str, str]:
    """
    Create one output row.
    """
    return {
        "Name": remove_inline_artifacts(name),
        "Reference": pad_reference(reference),
        "Description": remove_inline_artifacts(description),
        "Parent": pad_reference(parent) if parent else "",
        "Authority document": authority_document,
        "Supplemental guidance": remove_inline_artifacts(supplemental_guidance),
        "Active": active,
    }

def reference_exists(rows: List[Dict[str, str]], ref: str) -> bool:
    """
    Check if a reference already exists.
    Useful because docx2python can sometimes repeat bullets/letters.
    """
    padded = pad_reference(ref)
    return any(r.get("Reference") == padded for r in rows)

def append_to_row(rows: List[Dict[str, str]], reference: str, extra_text: str) -> None:
    """
    Append continuation text to an existing row.
    """
    padded_ref = pad_reference(reference)
    for i in range(len(rows) - 1, -1, -1):
        if rows[i]["Reference"] == padded_ref:
            rows[i]["Description"] = remove_inline_artifacts(
                rows[i]["Description"] + "\n" + extra_text
            )
            return

def root_parent(ref: str) -> str:
    """
    Return parent reference:
    03.02.01 -> 03.02
    """
    parts = ref.split(".")
    return ".".join(parts[:-1]) if len(parts) > 1 else ""

# =========================================================
# REGEX PATTERNS
# =========================================================
GUIDE_SECTION_RE = re.compile(r"^\d+(?:\.\d+)+\.?\s+.+$")
TOC_ENTRY_RE = re.compile(r"^\d+(?:\.\d+)*\s+.+?(?:\s+\d+)?$")
TOC_ENTRY_WITH_TITLE_RE = re.compile(r"^(\d+(?:\.\d+)*)\s+(.+?)(?:\s+\d+)?$")
HEADER_NOISE_RE = re.compile(
    r"^FINAL REPORT ON GUIDELINES ON ICT AND SECURITY RISK MANAGEMENT!?$",
    re.I,
)
SECTION_RE = re.compile(r"^(\d+(?:\.\d+)+)\.?\s+(.+)$")
PARAGRAPH_RE = re.compile(r"^(\d+)\.\s+(.+)$")
EXTRACTED_NUMBERED_PARAGRAPH_RE = re.compile(r"^(\d+)\)\s+(.+)$")
LETTER_RE = re.compile(r"^\(?([a-z])\)?[\).]\s+(.+)$", re.I)
ROMAN_RE = re.compile(r"^\(?([ivxlcdm]+)\)?[\).]\s+(.+)$", re.I)
BULLET_RE = re.compile(r"^(?:•|--|-|–)\s*(.+)$")

# =========================================================
# CORPORATE / UC HEADING HELPERS
# =========================================================
CORP_FIXED_HEADINGS = {
    "policy requirement and purpose",
    "applicability and scope",
    "minimum operational requirements",
    "methodology",
    "framework",
    "attachments",
    "references",
    "risk identification",
    "risk assessment and measurement",
    "risk response",
    "risk monitoring and reporting",
}
UC_MAIN_HEADINGS = {
    "policy requirement and purpose",
    "applicability and scope",
    "minimum operational requirements",
    "attachments",
    "references",
}
UC_LEVEL2_HEADINGS = {
    "methodology",
    "framework",
}
UC_LEVEL3_HEADINGS = {
    "risk identification",
    "risk assessment and measurement",
    "risk response",
    "risk monitoring and reporting",
}

def strip_trailing_page_number(text: str) -> str:
    """
    Remove trailing page number from TOC-like headings:
    'Risk Monitoring and Reporting 17' -> 'Risk Monitoring and Reporting'
    """
    return re.sub(r"\s+\d+$", "", clean_text(text)).strip()

def normalize_heading_text(text: str) -> str:
    """
    Normalize heading text for matching.
    """
    txt = strip_trailing_page_number(text)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip().lower()

def is_toc_start(line: str) -> bool:
    return normalize_heading_text(line) in {"table of contents", "contents"}

def is_probably_table_noise(line: str) -> bool:
    """
    Avoid treating table labels as document headings.
    """
    low = normalize_heading_text(line)
    noisy = {
        "perimeter",
        "exception",
        "approval date",
        "contacts",
        "description",
        "definitions/acronyms",
        "change type",
        "summary of main changes",
    }
    return low in noisy

def looks_like_corporate_heading(line: str) -> bool:
    """
    Detect UC-style headings from plain text.
    This is necessarily heuristic because docx2python does not expose Word
    heading styles like python-docx does.
    """
    txt = strip_trailing_page_number(line)
    low = normalize_heading_text(txt)
    if not txt:
        return False
    if low in CORP_FIXED_HEADINGS:
        return True
    if is_probably_table_noise(txt):
        return False
    if BULLET_RE.match(txt):
        return False
    if len(txt) > 90:
        return False
    if txt.endswith(".") or txt.endswith(":"):
        return False
    if re.match(r"^\d", txt):
        return False
    # Uppercase headings such as METHODOLOGY / FRAMEWORK.
    if txt.isupper() and len(txt.split()) <= 8:
        return True
    # Title-like subheadings.
    words = txt.split()
    if 2 <= len(words) <= 10:
        return True
    return False

def classify_corporate_heading(line: str) -> int:
    """
    Assign a level to UC-style headings.
    """
    low = normalize_heading_text(line)
    if low in UC_MAIN_HEADINGS:
        return 1
    if low in UC_LEVEL2_HEADINGS:
        return 2
    if low in UC_LEVEL3_HEADINGS:
        return 3
    return 4

# =========================================================
# DOCUMENT TYPE DETECTION
# =========================================================
def detect_document_type(docx_path: str) -> str:
    """
    Route document to:
    - corporate_framework
    - numbered_guideline
    """
    raw = extract_text_from_docx(docx_path)
    lines = split_lines(raw)
    has_toc = any(is_toc_start(x) for x in lines)
    has_uc_markers = any(
        normalize_heading_text(x) in CORP_FIXED_HEADINGS
        for x in lines
    )
    # UC policies usually contain a TOC and known business headings.
    if has_toc and has_uc_markers:
        return "corporate_framework"
    # Numbered guidelines contain real numbered section headings like:
    # 3.2.1. Governance
    real_guideline_sections = 0
    for line in lines:
        m = SECTION_RE.match(line)
        if not m:
            continue
        title = m.group(2).strip()
        # Avoid UC TOC entries ending with page numbers.
        if not re.search(r"\s+\d+$", title):
            real_guideline_sections += 1
    if real_guideline_sections >= 3:
        return "numbered_guideline"
    return "corporate_framework"

# =========================================================
# NUMBERED GUIDELINE PARSER
# =========================================================
def is_page_noise(text: str) -> bool:
    """
    Filter page headers, page numbers, dates, images.
    """
    txt = clean_text(text)
    if not txt:
        return True
    if is_noise_line(txt):
        return True
    if HEADER_NOISE_RE.match(txt):
        return True
    if re.fullmatch(r"\d+", txt):
        return True
    if re.fullmatch(r"\d{1,2}\s+[A-Za-z]+\s+\d{4}", txt):
        return True
    return False

def is_title_candidate(text: str) -> bool:
    """
    Detect short preface headings.
    """
    txt = clean_text(text)
    if not txt:
        return False
    if len(txt) > 85:
        return False
    if txt.endswith(":") or txt.endswith(";") or txt.startswith("("):
        return False
    if re.match(r"^\d", txt):
        return False
    blockers = ["this ", "the ", "according ", "all ", "financial ", "for "]
    if any(txt.lower().startswith(x) for x in blockers):
        return False
    return len(txt.split()) <= 10

def parse_preface(
    preface_lines: List[str],
    authority_document: str
) -> List[Dict[str, str]]:
    """
    Store preface/front matter under 00.*.
    """
    rows: List[Dict[str, str]] = []
    if not preface_lines:
        return rows
    rows.append(
        make_row(
            "Preface",
            "0",
            "Preface / preliminary material",
            "",
            authority_document,
        )
    )
    current_ref: Optional[str] = None
    current_title = "Preface"
    root_counter = 0
    child_counter = 0
    supp_counter: Dict[str, int] = {}
    for text in preface_lines:
        text = remove_inline_artifacts(text)
        if not text:
            continue
        # Footnotes / bibliography-like references.
        if (
            re.match(r"^\d+\s+Definition from ", text)
            or re.match(r"^\d+\s+Regulation \(EU\)", text)
        ):
            parent_ref = current_ref or "0"
            supp_counter[parent_ref] = supp_counter.get(parent_ref, 98) + 1
            rows.append(
                make_row(
                    "Supplemental guidance",
                    f"{parent_ref}.{supp_counter[parent_ref]}",
                    "Supplemental guidance",
                    parent_ref,
                    authority_document,
                    supplemental_guidance=text,
                )
            )
            continue
        if is_title_candidate(text):
            root_counter += 1
            current_ref = f"0.{root_counter}"
            current_title = text
            child_counter = 0
            rows.append(
                make_row(
                    text,
                    current_ref,
                    text,
                    "0",
                    authority_document,
                )
            )
            continue
        parent_ref = current_ref or "0"
        child_counter += 1
        rows.append(
            make_row(
                f"{current_title} - paragraph {child_counter}",
                f"{parent_ref}.{child_counter}",
                text,
                parent_ref,
                authority_document,
            )
        )
    return rows

def add_generated_child(
    rows: List[Dict[str, str]],
    authority_document: str,
    parent_ref: str,
    parent_title: str,
    text: str,
    counters: Dict[str, int],
) -> str:
    """
    Add a numeric child row under a parent reference.
    Used when bullet/letter extraction is ambiguous.
    """
    counters[parent_ref] = counters.get(parent_ref, 0) + 1
    child_ref = f"{parent_ref}.{counters[parent_ref]}"
    rows.append(
        make_row(
            f"{parent_title} - point {counters[parent_ref]}",
            child_ref,
            text,
            parent_ref,
            authority_document,
        )
    )
    return child_ref

def parse_numbered_guideline(docx_path: str) -> List[Dict[str, str]]:
    """
    Parse EBA/Guideline-style documents.
    """
    raw = extract_text_from_docx(docx_path)
    authority_document = Path(docx_path).name
    lines = [
        x for x in split_lines(raw)
        if not is_page_noise(x)
    ]
    # Split front matter from real numbered body.
    main_start = None
    for idx, line in enumerate(lines):
        m = SECTION_RE.match(line)
        if m and m.group(1).count(".") >= 1:
            main_start = idx
            break
    if main_start is None:
        main_start = len(lines)
    preface = lines[:main_start]
    body = lines[main_start:]
    rows: List[Dict[str, str]] = []
    rows.extend(parse_preface(preface, authority_document))
    current_section_ref: Optional[str] = None
    current_section_title = ""
    current_para_ref: Optional[str] = None
    current_para_title = ""
    current_letter_ref: Optional[str] = None
    current_letter_title = ""
    active_list_parent: Optional[str] = None
    active_list_parent_title = ""
    generated_children_count: Dict[str, int] = {}

    def close_list_mode() -> None:
        nonlocal active_list_parent, active_list_parent_title
        active_list_parent = None
        active_list_parent_title = ""

    for text in body:
        text = remove_inline_artifacts(text)
        if not text:
            continue
        # Section heading: 3.2.1 Governance
        m_sec = SECTION_RE.match(text)
        if m_sec:
            close_list_mode()
            sec_ref, sec_title = m_sec.groups()
            current_section_ref = sec_ref
            current_section_title = strip_trailing_page_number(sec_title)
            current_para_ref = None
            current_para_title = ""
            current_letter_ref = None
            current_letter_title = ""
            rows.append(
                make_row(
                    current_section_title,
                    current_section_ref,
                    current_section_title,
                    root_parent(current_section_ref),
                    authority_document,
                )
            )
            continue
        # Numbered paragraphs may appear as "10. text" or "10) text".
        m_para = (
            PARAGRAPH_RE.match(text)
            or EXTRACTED_NUMBERED_PARAGRAPH_RE.match(text)
        )
        if current_section_ref and m_para and not GUIDE_SECTION_RE.match(text):
            close_list_mode()
            num, body_text = m_para.groups()
            current_para_ref = f"{current_section_ref}.{num}"
            current_para_title = f"{current_section_title} - paragraph {num}"
            current_letter_ref = None
            current_letter_title = ""
            rows.append(
                make_row(
                    current_para_title,
                    current_para_ref,
                    body_text,
                    current_section_ref,
                    authority_document,
                )
            )
            if body_text.endswith(":"):
                active_list_parent = current_para_ref
                active_list_parent_title = current_para_title
            continue
        # Lettered point under current paragraph.
        m_letter = LETTER_RE.match(text)
        if current_para_ref and m_letter:
            close_list_mode()
            letter, letter_text = m_letter.groups()
            candidate_ref = f"{current_para_ref}.{letter.lower()}"
            if reference_exists(rows, candidate_ref):
                current_letter_ref = add_generated_child(
                    rows,
                    authority_document,
                    current_para_ref,
                    current_para_title,
                    letter_text,
                    generated_children_count,
                )
                current_letter_title = f"{current_section_title} - point"
            else:
                current_letter_ref = candidate_ref
                current_letter_title = f"{current_section_title} - point {letter.lower()}"
                rows.append(
                    make_row(
                        current_letter_title,
                        current_letter_ref,
                        letter_text,
                        current_para_ref,
                        authority_document,
                    )
                )
            if letter_text.endswith(":"):
                active_list_parent = current_letter_ref
                active_list_parent_title = current_letter_title
            continue
        # Roman point under current letter/paragraph.
        m_roman = ROMAN_RE.match(text)
        if (current_letter_ref or current_para_ref) and m_roman:
            close_list_mode()
            roman, roman_text = m_roman.groups()
            parent_ref = current_letter_ref or current_para_ref
            parent_title = current_letter_title or current_para_title or current_section_title
            candidate_ref = f"{parent_ref}.{roman.lower()}"
            if reference_exists(rows, candidate_ref):
                child_ref = add_generated_child(
                    rows,
                    authority_document,
                    parent_ref,
                    parent_title,
                    roman_text,
                    generated_children_count,
                )
            else:
                child_ref = candidate_ref
                rows.append(
                    make_row(
                        f"{parent_title} - point {roman.lower()}",
                        child_ref,
                        roman_text,
                        parent_ref,
                        authority_document,
                    )
                )
            current_letter_ref = child_ref
            current_letter_title = f"{parent_title} - point {roman.lower()}"
            if roman_text.endswith(":"):
                active_list_parent = child_ref
                active_list_parent_title = current_letter_title
            continue
        # Bullets / list items after a parent ending with ":"
        if active_list_parent:
            bullet_match = BULLET_RE.match(text)
            item_text = bullet_match.group(1) if bullet_match else text
            # Long narrative continuation after a list parent.
            if (
                not bullet_match
                and text.startswith("The ")
                and text.endswith(".")
                and len(text) > 80
            ):
                append_to_row(rows, active_list_parent, text)
                close_list_mode()
                continue
            add_generated_child(
                rows,
                authority_document,
                active_list_parent,
                active_list_parent_title or current_section_title,
                item_text,
                generated_children_count,
            )
            continue
        # Continuation text goes to the most specific current node.
        target_ref = current_letter_ref or current_para_ref or current_section_ref
        target_title = current_letter_title or current_para_title or current_section_title
        if target_ref is None:
            root_count = sum(1 for r in rows if r["Parent"] == "00") + 1
            rows.append(
                make_row(
                    f"Preface - paragraph {root_count}",
                    f"0.{root_count}",
                    text,
                    "0",
                    authority_document,
                )
            )
            continue
        append_to_row(rows, target_ref, text)
        if text.endswith(":"):
            active_list_parent = target_ref
            active_list_parent_title = target_title
    return rows

# =========================================================
# UC CORPORATE / FRAMEWORK PARSER
# =========================================================
def parse_toc_entry(line: str) -> Optional[Tuple[str, str]]:
    """
    Parse TOC line such as:
    3.2.4 Risk Monitoring and Reporting 17
    Returns:
    ("3.2.4", "Risk Monitoring and Reporting")
    """
    m = TOC_ENTRY_WITH_TITLE_RE.match(line)
    if not m:
        return None
    ref = m.group(1)
    title = strip_trailing_page_number(m.group(2))
    if not title:
        return None
    return ref, title

def split_corporate_sections(
    lines: List[str]
) -> Tuple[List[str], List[str], Dict[str, str]]:
    """
    Split a UC document into:
    - metadata lines
    - body lines
    - TOC map: normalized heading -> reference
    """
    metadata_lines: List[str] = []
    body_lines: List[str] = []
    toc_map: Dict[str, str] = {}
    in_toc = False
    body_started = False
    for raw_line in lines:
        line = remove_inline_artifacts(raw_line)
        if not line:
            continue
        norm = normalize_heading_text(line)
        if is_toc_start(line):
            in_toc = True
            continue
        if in_toc:
            parsed = parse_toc_entry(line)
            if parsed:
                ref, title = parsed
                toc_map[normalize_heading_text(title)] = ref
                continue
            # Start body at first real main heading.
            if norm == "policy requirement and purpose":
                in_toc = False
                body_started = True
                body_lines.append(strip_trailing_page_number(line))
                continue
            # Ignore remaining TOC noise.
            continue
        if not body_started:
            if norm == "policy requirement and purpose":
                body_started = True
                body_lines.append(strip_trailing_page_number(line))
            else:
                metadata_lines.append(line)
            continue
        body_lines.append(line)
    return metadata_lines, body_lines, toc_map

def get_corporate_reference(
    heading: str,
    level: int,
    toc_map: Dict[str, str],
    counters: Dict[str, int],
    current_refs: Dict[int, str],
) -> str:
    """
    Build or reuse corporate-style reference.
    Prefer TOC reference when available.
    """
    norm = normalize_heading_text(heading)
    # Prefer TOC reference.
    if norm in toc_map:
        ref = toc_map[norm]
        current_refs[level] = ref
        # Clear lower-level refs.
        for k in list(current_refs.keys()):
            if k > level:
                del current_refs[k]
        return ref
    # Otherwise generate.
    if level == 1:
        counters["h1"] = counters.get("h1", 0) + 1
        counters["h2"] = 0
        counters["h3"] = 0
        counters["h4"] = 0
        ref = str(counters["h1"])
    elif level == 2:
        if counters.get("h1", 0) == 0:
            counters["h1"] = 1
        counters["h2"] = counters.get("h2", 0) + 1
        counters["h3"] = 0
        counters["h4"] = 0
        ref = f"{counters['h1']}.{counters['h2']}"
    elif level == 3:
        if counters.get("h1", 0) == 0:
            counters["h1"] = 1
        if counters.get("h2", 0) == 0:
            counters["h2"] = 1
        counters["h3"] = counters.get("h3", 0) + 1
        counters["h4"] = 0
        ref = f"{counters['h1']}.{counters['h2']}.{counters['h3']}"
    else:
        parent = (
            current_refs.get(3)
            or current_refs.get(2)
            or current_refs.get(1)
            or "1"
        )
        counter_key = f"h4:{parent}"
        counters[counter_key] = counters.get(counter_key, 0) + 1
        ref = f"{parent}.{counters[counter_key]}"
    current_refs[level] = ref
    for k in list(current_refs.keys()):
        if k > level:
            del current_refs[k]
    return ref

def parse_corporate_framework(docx_path: str) -> List[Dict[str, str]]:
    """
    Parse UC-style policy/framework documents using text only.
    """
    raw = extract_text_from_docx(docx_path)
    lines = split_lines(raw)
    authority_document = Path(docx_path).name
    metadata_lines, body_lines, toc_map = split_corporate_sections(lines)
    rows: List[Dict[str, str]] = []
    # Metadata bucket.
    if metadata_lines:
        rows.append(
            make_row(
                "Metadata",
                "0",
                "Document metadata / front matter",
                "",
                authority_document,
            )
        )
        md_count = 0
        seen_md = set()
        for line in metadata_lines:
            clean = remove_inline_artifacts(line)
            if not clean:
                continue
            dedupe_key = normalize_heading_text(clean)
            # Avoid repeated table-cell noise.
            if dedupe_key in seen_md and len(clean) < 120:
                continue
            seen_md.add(dedupe_key)
            md_count += 1
            rows.append(
                make_row(
                    f"Metadata - item {md_count}",
                    f"0.{md_count}",
                    clean,
                    "0",
                    authority_document,
                )
            )
    counters: Dict[str, int] = {
        "h1": 0,
        "h2": 0,
        "h3": 0,
        "h4": 0,
    }
    current_refs: Dict[int, str] = {}
    current_section_ref: Optional[str] = None
    current_section_title = ""
    child_counter_by_ref: Dict[str, int] = {}
    seen_heading_refs = set()
    for raw_line in body_lines:
        line = remove_inline_artifacts(raw_line)
        if not line:
            continue
        # Skip residual TOC lines if they leaked through.
        if is_toc_start(line):
            continue
        parsed_toc = parse_toc_entry(line)
        if parsed_toc and normalize_heading_text(parsed_toc[1]) in toc_map:
            continue
        heading_line = strip_trailing_page_number(line)
        if looks_like_corporate_heading(heading_line):
            level = classify_corporate_heading(heading_line)
            ref = get_corporate_reference(
                heading_line,
                level,
                toc_map,
                counters,
                current_refs,
            )
            parent = root_parent(ref)
            heading_key = (
                pad_reference(ref),
                normalize_heading_text(heading_line),
            )
            if heading_key not in seen_heading_refs:
                rows.append(
                    make_row(
                        heading_line,
                        ref,
                        heading_line,
                        parent,
                        authority_document,
                    )
                )
                seen_heading_refs.add(heading_key)
            current_section_ref = ref
            current_section_title = heading_line
            child_counter_by_ref.setdefault(current_section_ref, 0)
            continue
        if current_section_ref is None:
            current_section_ref = "1"
            current_section_title = "Main section"
            rows.append(
                make_row(
                    current_section_title,
                    current_section_ref,
                    current_section_title,
                    "",
                    authority_document,
                )
            )
            child_counter_by_ref.setdefault(current_section_ref, 0)
        bullet_match = BULLET_RE.match(line)
        desc = bullet_match.group(1) if bullet_match else line
        label = "list item" if bullet_match else "paragraph"
        child_counter_by_ref[current_section_ref] = (
            child_counter_by_ref.get(current_section_ref, 0) + 1
        )
        child_no = child_counter_by_ref[current_section_ref]
        rows.append(
            make_row(
                f"{current_section_title} - {label} {child_no}",
                f"{current_section_ref}.{child_no}",
                desc,
                current_section_ref,
                authority_document,
            )
        )
    return rows

# =========================================================
# PUBLIC API / EXPORT
# =========================================================
def parse_docx_to_rows(docx_path: str) -> List[Dict[str, str]]:
    """
    Parse one .docx and return output rows.
    """
    document_type = detect_document_type(docx_path)
    if document_type == "numbered_guideline":
        rows = parse_numbered_guideline(docx_path)
    else:
        rows = parse_corporate_framework(docx_path)

    # --- POST-PROCESSING ---
    # Creiamo un dizionario di mappatura Reference -> Name
    ref_to_name = {row["Reference"]: row["Name"] for row in rows}
    
    # Sostituiamo il valore di Parent (reference) con il corrispondente Name
    for row in rows:
        parent_ref = row["Parent"]
        if parent_ref in ref_to_name:
            row["Parent"] = ref_to_name[parent_ref]
    # -----------------------

    return rows

def export_rows_to_excel(rows: List[Dict[str, str]], output_path: str) -> None:
    """
    Export rows to Excel.
    """
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    df.to_excel(output_path, index=False, engine="openpyxl")

def parse_docx_to_excel(docx_path: str, output_path: Optional[str] = None) -> str:
    """
    Parse one .docx and export to Excel.
    """
    docx_path = str(docx_path)
    if output_path is None:
        output_path = str(Path(docx_path).with_suffix(".xlsx"))
    rows = parse_docx_to_rows(docx_path)
    export_rows_to_excel(rows, output_path)
    return output_path

def parse_many(docx_paths: List[str], output_dir: str = ".") -> List[str]:
    """
    Parse multiple .docx files.
    """
    output_dir_path = Path(output_dir)
    output_dir_path.mkdir(parents=True, exist_ok=True)
    outputs: List[str] = []
    for path in docx_paths:
        output_file = output_dir_path / (Path(path).stem + ".xlsx")
        parse_docx_to_excel(path, str(output_file))
        outputs.append(str(output_file))
    return outputs

def main() -> None:
    """
    Default execution.
    If the sample files are in the same folder, parse them.
    """
    samples = [
        "Final_draft_Guidelines_on_ICT_and_security_risk.docx",
    ]
    existing = [x for x in samples if Path(x).exists()]
    if existing:
        outputs = parse_many(existing)
        print("Generated files:")
        for output in outputs:
            print(output)
    else:
        print("No sample files found. Use parse_docx_to_excel(path) or parse_many(paths).")

if __name__ == "__main__":
    main()
