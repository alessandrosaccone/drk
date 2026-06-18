import re # imports regular expression library --> used to find text
from docx2python import docx2python # import function that can read and extract their text
 
# =========================================================
# IO
# =========================================================
 
def extract_text_from_docx(path: str) -> str: # input: sting, output: string
    doc = docx2python(path) # reads .docx file
    return doc.text # returns the document text
# AI Act as one huge string
 
 
def normalize(text: str) -> str:
    # \r\n → \n prima di tutto (alcuni articoli usano \r\n tra numero e titolo)
    # converts different ways to write a line break (\r\n,\n,\r) into \n
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "")
 
    # Rimuovi intestazioni di pagina ELI / OJ L / numeri pagina / EN
    # re.sub(pattern, replacement, text)
    text = re.sub(r"\t?ELI:\s*(?:<[^>]+>)?https?://[^\n<]+(?:</[^>]+>)?\t?[^\n]*", "", text)
    text = re.sub(r"\tOJ L,[^\n]+", "", text)
    text = re.sub(r"\t\d{1,3}/\d{1,3}(?:\t[^\n]*|\n)", "", text)
    text = re.sub(r"\tEN(?:\t[^\n]*)?\n", "\n", text)
    text = re.sub(r"OJ L,.*?\n", "", text)
 
    # OCR artefacts
    text = text.replace("A!", "AI")
    text = text.replace("`", "")
 
    # Residui " EN" isolati dopo normalizzazione tab
    text = re.sub(r"\n EN\n", "\n", text)
    # Stops before ANNEX content begins
    annex_start = re.search(r"\nANNEX\s+[IVXLCDM]+\n", text, re.I) # searches for the first annex
    if annex_start:
        text = text[:annex_start.start()] # annex_start.start() --> position in the text where the annex begins
 
    # Rimuovi intestazioni CHAPTER / SECTION / ANNEX
    text = re.sub(r"\n\s*CHAPTER\s+[IVXLCDM\d]+\s*\n+\s*[^\n]+\n","\n",text, flags=re.I)
    # text = re.sub(r"\nCHAPTER\s+[IVXLCDM\d]+\n\n[^\n]+\n", "\n", text)
    # text = re.sub(r"\nCHAPTER\s+[IVXLCDM\d]+\n", "\n", text)
    text = re.sub(r"\nSECTION\s+\d+\n\n[^\n]+\n", "\n", text)
    text = re.sub(r"\nSECTION\s+\d+\n", "\n", text)
    text = re.sub(r"\nANNEX\s+[IVXLCDM]+\n\n[^\n]+\n", "\n", text)
    text = re.sub(r"\nANNEX\s+[IVXLCDM]+\n", "\n", text)
    # Rimouvi footnotes
    # Remove only real legal footnotes, not Article 3 definitions like (45), (46), etc.
    text = re.sub(
        r"\n\s*\(\d{1,3}\)\s+"
        r"(?=(?:OJ\s+[A-Z]|Regulation|Directive|Decision|Council|European Parliament|Position of the European Parliament))"
        r".*?"
        r"(?=\n\s*(?:Article\s+\d+|CHAPTER\s+|SECTION\s+|ANNEX\s+|\(\d{1,3}\)\s+(?:OJ\s+[A-Z]|Regulation|Directive|Decision|Council|European Parliament|Position of the European Parliament))|\Z)",
        "\n",
        text,
        flags=re.S | re.I
    )
    #text = re.sub(r"\n\s*\(\d{1,3}\)\s+.*?(?=\n\s*(?:[A-Z][a-z]|Article\s+\d+|\d+[.\)]\s+|\([a-z]\)|[a-z][\).]\s+|CHAPTER\s+|SECTION\s+|ANNEX\s+)|\Z)","\n",text, flags=re.S)
    # text = re.sub("\n\s*\(\d{1,3}\)\s+.*?(?=\n\s*(?:Article\s+\d+|\d+[.\)]\s+|\([a-z]\)|[a-z][\).]\s+|CHAPTER\s+|SECTION\s+|ANNEX\s+)|\Z)", "\n", text, flags=re.S | re.I)
    # text = re.sub(r"\n\s*\(\d{1,3}\)\s+.*?(?=\n\s*(?:Article\s+\d+|\d+[.\)]\s+|\([a-z]\)|[a-z][\).]\s+|CHAPTER\s+|SECTION\s+|ANNEX\s+)|\Z),"\n",text, flags=re.S|re.I")
 
    # Paragrafi consecutivi sulla stessa riga separati da " N.\t" (es. Art 64)
    text = re.sub(r"(?<!\n)( \d+)\.\t", r"\n\1. ", text)
    # (?<!\n) = not immediately preceded by a new line
    # ( \d+)\.\t = space, number, period, tab
    # r"\n\1. " = (replacement) newline before the number
 
    # Normalizza tab → spazio (mantieni \n)
    text = re.sub(r"[ \t]+", " ", text) # "[ \t]+" = multiple spaces or tabs
    text = re.sub(r"\n{3,}", "\n\n", text) # \n{3,} = 3 or more newlines

 
    return text.strip()
 
 
# =========================================================
# MACRO STRUCTURE
# =========================================================
 
ARTICLE_RE = re.compile(
    r"(?:^|\n)Article\s+(\d+)\s*\n"      # Article header
    r"(?!\s)"                            # no indent (evita inline/ref)
    r"(.*?)"
    r"(?=(?:\nArticle\s+\d+\s*\n)|\Z)",  # next true article
    re.S | re.I
)
 
 
#ARTICLE_RE = re.compile(
#    r"\nArticle\s+(\d+)\n(.*?)(?=\nArticle\s+\d+|\Z)",
#    re.S | re.I
#)
# \nArticle\s+(\d+)\n(.*?) = find articles (caputing number [(\d+)] + article content [(.*?)], stops ASAP [?])
# (?=\nArticle\s+\d+|\Z) = stop when article starts or doc ends ([\Z] = end of string)
# re.S = "." match any character (including a new line)
# re.I = ignore case (case-insensitive matching)
 
# =========================================================
# SUBSTRUCTURE PATTERNS
# =========================================================
 
# Paragrafi numerati: (N), N., N)
NUMBER_RE = re.compile(
    r"(?:^|\n) ?(?:\((\d+)\)|(\d+)[.\)])\s+(.*?)"
    r"(?=\n ?(?:\(\d+\)|\d+[.\)])\s|\Z)",
    re.S
)
 
# Lettere: (a), a), a.
LETTER_RE = re.compile(
    r"(?:^|\n) ?(?:\(([a-z])\)|([a-z])[)\.])\s+(.*?)"
    r"(?=\n ?(?:\([a-z]\)|[a-z][)\.]) |\n ?(?:\(\d+\)|\d+[.\)])\s|\Z)",
    re.S | re.I
)
 
# Romani: (i), (ii), i), ii), ecc.
ROMAN_RE = re.compile(
    r"(?:^|\n) ?(?:\(([ivxlcdm]+)\)|([ivxlcdm]+)[)\.])\s+(.*?)"
    r"(?=\n ?(?:\([ivxlcdm]+\)|[ivxlcdm]+[)\.]) |\n ?(?:\([a-z]\)|[a-z][)\.]) |\Z)",
    re.S | re.I
)
 
VALID_ROMANS = re.compile(
    r"^(i{1,3}|iv|vi{0,3}|ix|xi{0,2}|xii|xiii|xiv|xv)$",
    re.I
)
 
 
# =========================================================
# UTILITIES
# =========================================================
 
# returns True if "s" is a valid roman numeral
def is_roman(s: str) -> bool:
    return bool(VALID_ROMANS.match(s))
 
# seperates main paragraph text from the subpoints
def extract_items(pattern, text, roman=False):
    items = []
    for m in pattern.finditer(text):
        groups = m.groups()
        ident = next((g for g in groups[:-1] if g), None)
        if ident is None:
            continue
        if roman and not is_roman(ident):
            continue
        items.append({"id": ident, "text": groups[-1].strip()})
    cleaned = pattern.sub("", text).strip()
    return items, cleaned
 
#  function solves the ambiguity between letter i and roman i.
# a sequence like a, i, ii, iii, b should become letter a with roman subpoints,
# while a sequence like a, b, ..., h, i, j should treat i as the ninth letter.  
 
 
def process_letters(letters_raw, pid):
    """
    Classificazione sequenziale lettere vs sub-romani.
 
    LETTER_RE può catturare 'i)' come "lettera i" quando in realtà è il primo
    romano di una sub-lista della lettera precedente. Distinzione tramite la
    sequenza alfabetica attesa:
      - se l'id matcha la lettera alfabetica attesa (a, b, c, ...) → lettera vera
      - altrimenti, se è un romano valido → sub-romano della lettera precedente
        (ed eventuali romani nidificati nel suo testo sono i successivi sub-romani)
    Es. struttura "a, i, ii, iii, b" → lettere = [a, b] con a.supplementary = [i, ii, iii]
    Es. struttura "a, b, ..., h, i, j, k, l" → tutte lettere alfabetiche (i = nona lettera)
    """
    letter_objs = []
    expected = "a"
 
    for l in letters_raw: # Loop through each detected letter-like item
        iid = l["id"]
 
        if iid == expected:
            # Lettera alfabetica vera: estrai sub-romani dal suo testo
            romans_in_text, l_text_clean = extract_items(ROMAN_RE, l["text"], roman=True)
            letter_objs.append({
                "id": f"{pid}.{iid}",
                "text": l_text_clean,
                "supplementary": [
                    {"id": f"{pid}.{iid}.{r['id']}", "text": r["text"]}
                    for r in romans_in_text
                ]
            })
            expected = chr(ord(expected) + 1) # moves to the next expected letter
            # ord = number code for the letter
            # chr(...) = turns it back into a character
        elif is_roman(iid) and letter_objs:
            # Sub-romano della lettera precedente.
            # Il testo dell'item può contenere altri romani nidificati (ii, iii...).
            prev = letter_objs[-1] # get the last real letter
            prev_letter_id = prev["id"].split(".")[-1] #
            inner_romans, l_text_clean = extract_items(ROMAN_RE, l["text"], roman=True)
            # L'item stesso è il primo sub-romano
            prev["supplementary"].append({
                "id": f"{pid}.{prev_letter_id}.{iid}",
                "text": l_text_clean
            })
            # Add this roman numeral under the previous letter.
 
            # I romani nidificati nel suo testo sono i successivi
            for r in inner_romans:
                prev["supplementary"].append({
                    "id": f"{pid}.{prev_letter_id}.{r['id']}",
                    "text": r["text"]
                })
        # else: caso non gestito (raro) → skip
 
    return letter_objs
 
 
# =========================================================
# ARTICLE PARSER
# =========================================================
 
def parse_article(body: str, article_num: str):
    paragraphs, remainder = extract_items(NUMBER_RE, body)
 
    structured = []
 
    for p in paragraphs:
        pid = f"{article_num}.{p['id']}"
        letters_raw, p_text = extract_items(LETTER_RE, p["text"])
        letter_objs = process_letters(letters_raw, pid)
 
        structured.append({
            "id": pid,
            "text": p_text if letter_objs else p["text"],
            # If the paragraph has letters, use the cleaned paragraph text without the letter items.
            # If it has no letters, keep the full paragraph text.
 
            "letters": letter_objs
            })
    if not structured and remainder: # if NO numbered paragraphs, but some text is left
        letters_raw, p_text = extract_items(LETTER_RE, remainder)
        letter_objs = process_letters(letters_raw, f"{article_num}.1")
 
    # CASE: letters exist → create XX.1
        if letter_objs:
            structured.append({
                "id": f"{article_num}.1", # creating non-existent XX.1 for articles with sub lettered paragphas
                "text": p_text,
                "letters": letter_objs
            })
 
    return structured

 
    # if not structured and remainder:
    #     # Articoli senza paragrafi numerati ma con lista di lettere (es. Art 16, 66)
    #     pid = f"{article_num}.1"
    #     letters_raw, p_text = extract_items(LETTER_RE, remainder)
    #     letter_objs = process_letters(letters_raw, pid)
 
    #     structured.append({
    #         "id": pid,
    #         "text": p_text if letter_objs else remainder,
    #         "letters": letter_objs
            #})
 
    return structured
 
 
# =========================================================
# DOCUMENT PARSER
# =========================================================
 
def parse_document(text: str):
    articles = []
 
    for m in ARTICLE_RE.finditer(text):
        num = m.group(1)
        content = m.group(2).strip()
 
        lines = content.split("\n")
        first_line = lines[0].strip() if lines else ""
 
        # Alcuni articoli hanno titolo e paragrafo 1 sulla stessa riga
        # (es. Art 65: "Establishment...Board 1. A European AI Board...")
        inline_par = re.search(r"\s(\d+[.\)])\s", first_line)
        if (inline_par
                and not re.match(r" ?(?:\(?\d+\)?[.\)])", first_line)
                and len(first_line) < 300):
        # Only treat this as an inline paragraph if:
            # A number like 1. appears inside the first line.
            # The first line does not already start with a paragraph number.
            # The line is not too long.
            # Split title an paragraph text when they are not seperated
            split_pos = inline_par.start()
            title = first_line[:split_pos].strip()
            rest_of_first = first_line[inline_par.start() + 1:].strip()
            body = (rest_of_first + "\n" + "\n".join(lines[1:])).strip()
        # elif (first_line
        #       and not re.match(r" ?(?:\(?\d+\)?[.\)])", first_line)
        #       and len(first_line) < 120):
        #     # If the first line exists, does not start with a paragraph number, and is short, treat it as the title.
        #     # PROBLEM: Art. 59 & Art. 111 have two-line titles and they re cut-off
        #     title = first_line
        #     body = "\n".join(lines[1:]).strip()
    #     elif first_line and not re.match(r" ?(?:\(?\d+\)?[.\)])", first_line):
    #         title_lines = [first_line]
    #         body_start_idx = 1
 
        elif first_line and not re.match(r" ?(?:\(?\d+\)?[.\)])", first_line):
            title_lines = []
            body_start_idx = 0
            for i, line in enumerate(lines):
                stripped = line.strip()
                # BODY START = first numbered paragraph or letter
                if re.match(r"^\s*(\(?\d+\)?[.\)])\s+", stripped):
                    body_start_idx = i
                    break
                if re.match(r"^\s*(\([a-z]\)|[a-z][\).])\s+", stripped, re.I):
                    body_start_idx = i
                    break
                # otherwise → still title
                title_lines.append(stripped)
            # if no body found → entire content is title
            if not title_lines:
                title = ""
                body = content
            else:
                title = re.sub(r"\s+", " ", " ".join(l.strip() for l in title_lines)).strip()
                body = "\n".join(lines[body_start_idx:]).strip()
 
        else: # no title detected
            title = ""
            body = content
        structure = parse_article(body, num)
 
        if not structure:
            full_text = (title + "\n" + body).strip()
            articles.append({
                "id": f"Article {num}",
                "title": full_text,
                "structure": []
            })
        else:
            articles.append({
                "id": f"Article {num}",
                "title": title,
                "structure": structure
            })
        # articles.append({
        #     "id": f"Article {num}",
        #     "title": title,
        #     "structure": parse_article(body, num)
        # })
 
    return articles
 
 
# =========================================================
# FLATTEN (OUTPUT MODEL)
# =========================================================
 
def flatten(parsed):
    articles = {}
    paragraphs = {}
    letters = {}
    supplementary = {}
 
    for art in parsed:
        art_num = art["id"].split()[1]
        articles[art_num] = art["title"]
 
        for p in art["structure"]:
            paragraphs[p["id"]] = p["text"]
 
            for l in p["letters"]:
                letters[l["id"]] = l["text"]
 
                for r in l["supplementary"]:
                    supplementary[r["id"]] = r["text"]
 
    return articles, paragraphs, letters, supplementary
 
 
# =========================================================
# MAIN
# =========================================================
 
def main(docx_path: str, output_path="parsed_output.txt"):
    raw = extract_text_from_docx(docx_path) #read Word file
    text = normalize(raw) #clean the text
 
    parsed = parse_document(text) #parse clean text into their sub-sections
    articles, paragraphs, letters, supplementary = flatten(parsed) # convert into dictionaries
 
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("=== ARTICLES ===\n")
        for k, v in articles.items():
            f.write(f"{k} → {v}\n")
 
        f.write("\n=== PARAGRAPHS ===\n")
        for k, v in paragraphs.items():
            f.write(f"{k} → {v}\n\n")
 
        f.write("\n=== LETTERS ===\n")
        for k, v in letters.items():
            f.write(f"{k} → {v}\n\n")
 
        f.write("\n=== SUPPLEMENTARY ===\n")
        for k, v in supplementary.items():
            f.write(f"{k} → {v}\n\n")
 
    print(f"Output written to {output_path}")
    print(f"  Articoli:      {len(articles)}")
    print(f"  Paragrafi:     {len(paragraphs)}")
    print(f"  Lettere:       {len(letters)}")
    print(f"  Supplementary: {len(supplementary)}")
 
 
import pandas as pd
 
 
# =========================================================
# BUILD ROWS FOR EXCEL
# =========================================================
 
import pandas as pd
 
 
# =========================================================
# BUILD ROWS (STRICT TEMPLATE MATCH)
# =========================================================
 
def pad_token(token: str) -> str:
    # "4" -> "04", "4.1" -> "04.01", lascia lettere e romani intatti ("a", "i")
    if not token:
        return token
    return ".".join(p.zfill(2) if p.isdigit() else p for p in token.split("."))
 
def build_rows(parsed):
    rows = []
 
    for art in parsed:
        art_num = art["id"].split()[1]   # "Article 4" -> "4"
        art_title = art["title"]
 
        # =========================
        # ARTICLE
        # =========================
        rows.append({
            "Name": art_num,
            "Reference": art_num,
            "Description": art_title,
            "Parent": "",
            "Authority document": "AI ACT",
            "Supplemental guidance": art_title,
            "Active": False
        })
 
        # =========================
        # PARAGRAPHS (4.1)
        # =========================
        for p in art["structure"]:
            pid = p["id"]
            short_ref = pid.split(".")[-1]
 
            rows.append({
                "Name": pid,
                "Reference": short_ref,
                "Description": art_title,
                "Parent": art_num,
                "Authority document": "AI ACT",
                "Supplemental guidance": p["text"],
                "Active": False
            })
 
            # =========================
            # LETTERS (4.1.a)
            # =========================
            for l in p["letters"]:
                lid = l["id"]
                short_ref = lid.split(".")[-1]
 
                rows.append({
                    "Name": lid,
                    "Reference": short_ref,
                    "Description": art_title,
                    "Parent": pid,
                    "Authority document": "AI ACT",
                    "Supplemental guidance": l["text"],
                    "Active": False
                })
 
                # =========================
                # ROMAN (4.1.a.i)
                # =========================
                for r in l["supplementary"]:
                    rid = r["id"]
                    short_ref = rid.split(".")[-1]
 
                    rows.append({
                        "Name": rid,
                        "Reference": short_ref,
                        "Description": art_title,
                        "Parent": lid,
                        "Authority document": "AI ACT",
                        "Supplemental guidance": r["text"],
                        "Active": False
                    })
 
        for r in rows:
            r["Name"] = pad_token(r["Name"])
            r["Parent"] = pad_token(r["Parent"])
            r["Reference"] = pad_token(r["Reference"])
 
    return rows
 
 
# =========================================================
# EXPORT TO EXCEL (MATCH TEMPLATE EXACTLY)
# =========================================================
 
def export_to_excel(parsed, output_path="citation_output.xlsx"):
    rows = build_rows(parsed)
 
    # ⚠️ ORDINE COLONNE IDENTICO AL TEMPLATE
    columns = [
        "Name",
        "Reference",
        "Description",
        "Parent",
        "Authority document",
        "Supplemental guidance",
        "Active"
    ]
 
    df = pd.DataFrame(rows)
    df = df[columns]   # forza ordine colonne
 
    # Scrittura Excel
    df.to_excel(output_path, index=False, engine="openpyxl")
 
    print(f"✅ File Excel creato: {output_path}")
    print(f"Totale righe: {len(df)}")
 
 
# =========================================================
# MAIN
# =========================================================
 
def main(docx_path: str):
    raw = extract_text_from_docx(docx_path)
    text = normalize(raw)
 
    parsed = parse_document(text)
 
    export_to_excel(parsed, "DORA.xlsx")
 
 
if __name__ == "__main__":
    main("AI_Act_OJ_L_202401689_EN_TXT.docx")
    #main("CELEX_32022R2554_EN_TXT.docx")
