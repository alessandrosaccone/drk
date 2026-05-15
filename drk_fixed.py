import re
from docx2python import docx2python


# =========================================================
# IO
# =========================================================

def extract_text_from_docx(path: str) -> str:
    doc = docx2python(path)
    return doc.text


def normalize(text: str) -> str:
    # \r\n → \n prima di tutto (alcuni articoli usano \r\n tra numero e titolo)
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "")

    # Rimuovi intestazioni di pagina ELI / OJ L / numeri pagina / EN
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

    # Rimuovi intestazioni CHAPTER / SECTION / ANNEX
    text = re.sub(r"\nCHAPTER\s+[IVXLCDM\d]+\n\n[^\n]+\n", "\n", text)
    text = re.sub(r"\nCHAPTER\s+[IVXLCDM\d]+\n", "\n", text)
    text = re.sub(r"\nSECTION\s+\d+\n\n[^\n]+\n", "\n", text)
    text = re.sub(r"\nSECTION\s+\d+\n", "\n", text)
    text = re.sub(r"\nANNEX\s+[IVXLCDM]+\n\n[^\n]+\n", "\n", text)
    text = re.sub(r"\nANNEX\s+[IVXLCDM]+\n", "\n", text)

    # Paragrafi consecutivi sulla stessa riga separati da " N.\t" (es. Art 64)
    text = re.sub(r"(?<!\n)( \d+)\.\t", r"\n\1. ", text)

    # Normalizza tab → spazio (mantieni \n)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


# =========================================================
# MACRO STRUCTURE
# =========================================================

ARTICLE_RE = re.compile(
    r"\nArticle\s+(\d+)\n(.*?)(?=\nArticle\s+\d+|\Z)",
    re.S | re.I
)


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

def is_roman(s: str) -> bool:
    return bool(VALID_ROMANS.match(s))


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

    for l in letters_raw:
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
            expected = chr(ord(expected) + 1)
        elif is_roman(iid) and letter_objs:
            # Sub-romano della lettera precedente.
            # Il testo dell'item può contenere altri romani nidificati (ii, iii...).
            prev = letter_objs[-1]
            prev_letter_id = prev["id"].split(".")[-1]
            inner_romans, l_text_clean = extract_items(ROMAN_RE, l["text"], roman=True)
            # L'item stesso è il primo sub-romano
            prev["supplementary"].append({
                "id": f"{pid}.{prev_letter_id}.{iid}",
                "text": l_text_clean
            })
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
            "letters": letter_objs
        })

    if not structured and remainder:
        # Articoli senza paragrafi numerati ma con lista di lettere (es. Art 16, 66)
        pid = f"{article_num}.1"
        letters_raw, p_text = extract_items(LETTER_RE, remainder)
        letter_objs = process_letters(letters_raw, pid)

        structured.append({
            "id": pid,
            "text": p_text if letter_objs else remainder,
            "letters": letter_objs
        })

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
            split_pos = inline_par.start()
            title = first_line[:split_pos].strip()
            rest_of_first = first_line[inline_par.start() + 1:].strip()
            body = (rest_of_first + "\n" + "\n".join(lines[1:])).strip()
        elif (first_line
              and not re.match(r" ?(?:\(?\d+\)?[.\)])", first_line)
              and len(first_line) < 120):
            title = first_line
            body = "\n".join(lines[1:]).strip()
        else:
            title = ""
            body = content

        articles.append({
            "id": f"Article {num}",
            "title": title,
            "structure": parse_article(body, num)
        })

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
    raw = extract_text_from_docx(docx_path)
    text = normalize(raw)

    parsed = parse_document(text)
    articles, paragraphs, letters, supplementary = flatten(parsed)

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


if __name__ == "__main__":
    main("AI_Act_OJ_L_202401689_EN_TXT.docx")
