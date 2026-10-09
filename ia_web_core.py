from pathlib import Path
import sys, re, math, json, unicodedata, urllib.request, zipfile
from collections import Counter, defaultdict
from lxml import etree

ROOT = Path(__file__).resolve().parent
NORMATIVA = ROOT / "Normativa"
TOP_K = 5
MAX_THEMATIC_HITS = 40

if hasattr(sys.stdout, "reconfigure"):
    try: sys.stdout.reconfigure(encoding="utf-8")
    except Exception: pass

STOPWORDS = set("a al algo algun alguna algunas alguno algunos ante bajo con contra cual cuando de del desde donde durante e el ella ellas ellos en entre era es esa esas ese eso esos esta estas este esto estos fue ha hay la las le les lo los mas me mi mis muy no o para pero por que se sea segun ser si sin sobre su sus tambien te tiene un una unas uno unos y ya dice establece dispone refiere respecto".split())

ARTICLE_LINE_RE = re.compile(r"^\s*(?:art(?:í|i)?culo|art\s*\.?)\s*(?:n\s*(?:ro\s*\.?|[º°o.]?)\s*)?(?P<num>\d+(?:\s*(?:bis|ter|quater|quinquies))?)\s*[º°o.]?\s*[-–—.:)]*", re.I)
NORM_ANY_RE = re.compile(r"(?P<tipo>decreto[\s-]*ley|ley|decreto|resoluci[oó]n|disposici[oó]n|acordada)\s*(?:n(?:ro\.?|[º°o.]?)\s*)?(?P<num>\d[\d.\-/]*)", re.I)
QUERY_ART_RE = re.compile(r"(?:art(?:í|i)?culo|art\s*\.?)\s*(?:n\s*(?:ro\s*\.?|[º°o.]?)\s*)?(?P<num>\d+(?:\s*(?:bis|ter|quater|quinquies))?)", re.I)
QUERY_NORM_RE = re.compile(r"(?P<tipo>decreto[\s-]*ley|ley|decreto|resoluci[oó]n|disposici[oó]n|acordada)\s*(?:n(?:ro\.?|[º°o.]?)\s*)?(?P<num>\d[\d.\-/]*)", re.I)

def normnum(s):
    return re.sub(r"[.\s]", "", s.lower())

def normalize(text):
    text = unicodedata.normalize("NFD", text.lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9ñü\s/.-]", " ", text)).strip()

def toks(text):
    return [t for t in normalize(text).split() if len(t) > 2 and t not in STOPWORDS]


def _xml_bool(node, qname):
    if node is None:
        return False
    val = node.get(qname)
    if val is None:
        return True
    return str(val).lower() not in {"0", "false", "off", "none"}

def extract_docx_blocks(path):
    """
    Extrae bloques respetando el orden real del DOCX.

    - Los párrafos se conservan como párrafos.
    - Las tablas se conservan como UNA unidad con filas y columnas,
      en lugar de convertir cada celda en líneas independientes.
    - display_text mantiene tabulaciones para que la interfaz pueda
      reproducir visualmente la disposición del Compendio.
    """
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml")

    root = etree.fromstring(xml)
    ns = {"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    blocks = []
    body = root.find("w:body", ns)
    if body is None:
        return blocks

    def paragraph_data(p):
        runs = p.xpath(".//w:r", namespaces=ns)
        pieces = []
        bold_chars = 0
        italic_chars = 0
        total_chars = 0

        for r in runs:
            texts = r.xpath(".//w:t/text()", namespaces=ns)
            txt = "".join(t for t in texts if t)
            if not txt:
                continue

            pieces.append(txt)
            chars = len(txt.strip())
            if chars <= 0:
                continue

            rpr = r.find("w:rPr", ns)
            is_bold = False
            is_italic = False
            if rpr is not None:
                is_bold = _xml_bool(rpr.find("w:b", ns), W + "val")
                is_italic = _xml_bool(rpr.find("w:i", ns), W + "val")

            total_chars += chars
            if is_bold:
                bold_chars += chars
            if is_italic:
                italic_chars += chars

        text = " ".join(x.strip() for x in pieces if x and x.strip()).strip()
        if not text:
            return None

        total = max(1, total_chars)
        return {
            "text": text,
            "display_text": text,
            "bold_ratio": bold_chars / total,
            "italic_ratio": italic_chars / total,
            "is_table": False,
            "table_rows": None,
        }

    def table_data(tbl):
        rows = []
        search_rows = []
        bold_chars = 0
        italic_chars = 0
        total_chars = 0

        for tr in tbl.xpath("./w:tr", namespaces=ns):
            cells = []
            search_cells = []

            for tc in tr.xpath("./w:tc", namespaces=ns):
                cell_parts = []
                cell_search = []

                for p in tc.xpath(".//w:p", namespaces=ns):
                    pdata = paragraph_data(p)
                    if not pdata:
                        continue

                    cell_parts.append(pdata["text"])
                    cell_search.append(pdata["text"])

                    chars = max(1, len(pdata["text"].strip()))
                    total_chars += chars
                    bold_chars += int(chars * pdata["bold_ratio"])
                    italic_chars += int(chars * pdata["italic_ratio"])

                # En una misma celda varios párrafos se muestran juntos,
                # separados por salto de línea si fuera necesario.
                cell_text = " / ".join(cell_parts).strip()
                search_text = " ".join(cell_search).strip()
                cells.append(cell_text)
                search_cells.append(search_text)

            # Conservar incluso celdas vacías: son las que mantienen la alineación.
            rows.append(cells)
            search_rows.append(search_cells)

        if not rows:
            return None

        # Quitar solo columnas completamente vacías al extremo derecho.
        max_cols = max(len(r) for r in rows)
        while max_cols > 0:
            if all((len(r) < max_cols or not str(r[max_cols-1]).strip()) for r in rows):
                max_cols -= 1
            else:
                break
        rows = [r[:max_cols] for r in rows]
        search_rows = [r[:max_cols] for r in search_rows]

        # Una tabulación por celda conserva el esquema de columnas en la GUI.
        display_lines = ["\t".join(r) for r in rows]
        display_text = "\n".join(display_lines).strip()
        search_text = " ".join(
            cell for row in search_rows for cell in row if cell
        ).strip()

        total = max(1, total_chars)
        return {
            "text": search_text or display_text,
            "display_text": display_text,
            "bold_ratio": bold_chars / total,
            "italic_ratio": italic_chars / total,
            "is_table": True,
            "table_rows": rows,
        }

    for child in body:
        tag = etree.QName(child).localname

        if tag == "p":
            pdata = paragraph_data(child)
            if pdata:
                blocks.append(pdata)

        elif tag == "tbl":
            tdata = table_data(child)
            if tdata:
                blocks.append(tdata)

    return blocks

def build_units_formatted(blocks, source):
    """
    Construye unidades jurídicas preservando la relación Ley 931 / Decreto 1324/93.

    Además de mirar negrita/cursiva, reconoce expresamente encabezados como:
        Decreto N° 1.324/93 – ARTÍCULO 92°: ...
    y los indexa como un artículo propio del Decreto 1324/93.
    """
    units = []
    current_norms = []
    current_article = None
    current_segments = []
    ley931_context = []
    current_is_regulation = False

    def dedupe_norms(items):
        out, seen = [], set()
        for item in items:
            key = item.get("key")
            if key not in seen:
                seen.add(key)
                out.append(item)
        return out

    def classify_segment(seg):
        br = seg.get("bold_ratio", 0.0)
        ir = seg.get("italic_ratio", 0.0)
        if ir >= 0.35 and ir > br:
            return "reglamentacion"
        if br >= 0.35 and br >= ir:
            return "ley"
        return "neutral"

    def flush():
        nonlocal current_segments, current_article
        if not current_article or not current_segments:
            current_segments = []
            return

        full_text = "\n".join(s.get("display_text", s["text"]) for s in current_segments).strip()
        if not full_text:
            current_segments = []
            return

        keys = [x["key"] for x in current_norms]
        labels = [x["label"] for x in current_norms] or ["Norma no identificada"]

        law_parts = []
        reg_parts = []
        neutral_parts = []

        for seg in current_segments:
            role = classify_segment(seg)
            if role == "ley":
                law_parts.append(seg.get("display_text", seg["text"]))
            elif role == "reglamentacion":
                reg_parts.append(seg.get("display_text", seg["text"]))
            else:
                neutral_parts.append(seg.get("display_text", seg["text"]))

        if "931" in keys:
            if not law_parts:
                first = current_segments[0].get("display_text", current_segments[0]["text"])
                law_parts.append(first)
                neutral_parts = [x for x in neutral_parts if x != first]

            law_text = "\n".join(law_parts + neutral_parts).strip()
            reg_text = "\n".join(reg_parts).strip()
        elif "1324/93" in keys:
            # Para una unidad propia del Decreto, todo su contenido es reglamentación.
            law_text = ""
            reg_text = full_text
        else:
            law_text = full_text
            reg_text = ""

        units.append({
            "source": source,
            "norm_labels": labels,
            "norm_keys": keys,
            "article": current_article,
            "text": full_text,
            "law_text": law_text,
            "reg_text": reg_text,
            "has_regulation": bool(reg_text),
        })
        current_segments = []

    for block in blocks:
        line = block["text"]
        ms = norm_matches(line)
        ms_keys = {x["key"] for x in ms}

        # CASO CLAVE:
        # El Compendio escribe la reglamentación como un párrafo que comienza con
        # "Decreto N° 1.324/93 – ARTÍCULO XX°: ...".
        # Antes se lo tomaba solo como encabezado y se perdía su contenido.
        inline_article = QUERY_ART_RE.search(line) if ms else None
        if "1324/93" in ms_keys and inline_article:
            flush()
            current_norms = [x for x in ms if x["key"] == "1324/93"]
            current_article = inline_article.group("num").strip()
            current_segments = [block]
            current_is_regulation = True
            continue

        # Encabezados generales de normas.
        if ms and is_heading(line):
            flush()
            current_article = None
            keys = ms_keys

            if "931" in keys:
                ley931_context = list(ms)
                current_norms = list(ms)
                current_is_regulation = False

            elif "1324/93" in keys and ley931_context:
                current_norms = dedupe_norms(ley931_context + ms)
                current_is_regulation = False

            else:
                current_norms = list(ms)
                current_is_regulation = False
                if keys & {"971", "572/85", "1613", "26206"}:
                    ley931_context = []
            continue

        # Artículo de la Ley 931 (o de la norma actual) al comienzo del párrafo.
        am = ARTICLE_LINE_RE.match(line)
        if am:
            flush()

            # Si venimos de una unidad reglamentaria del Decreto 1324/93,
            # el siguiente ARTÍCULO sin prefijo vuelve a ser de la Ley 931.
            if current_is_regulation and ley931_context:
                current_norms = list(ley931_context)
                current_is_regulation = False

            current_article = am.group("num").strip()
            current_segments = [block]
            continue

        if current_article:
            current_segments.append(block)

    flush()
    return units

def extract_docx_lines(path):
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    ns = {"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    lines = []
    body = root.find("w:body", ns)
    for child in body:
        tag = etree.QName(child).localname
        if tag == "p":
            texts = child.xpath(".//w:t/text()", namespaces=ns)
            txt = " ".join(t.strip() for t in texts if t and t.strip()).strip()
            if txt: lines.append(txt)
        elif tag == "tbl":
            for p in child.xpath(".//w:p", namespaces=ns):
                texts = p.xpath(".//w:t/text()", namespaces=ns)
                txt = " ".join(t.strip() for t in texts if t and t.strip()).strip()
                if txt: lines.append(txt)
    return lines

def norm_matches(line):
    out = []
    for m in NORM_ANY_RE.finditer(line):
        tipo = re.sub(r"\s+", " ", m.group("tipo").replace("-", " ")).title()
        nr = m.group("num")
        out.append({"label":f"{tipo} {nr}", "key":normnum(nr)})
    return out

def is_heading(line):
    return bool(re.match(r"^\s*(decreto[\s-]*ley|ley|decreto|resoluci[oó]n|disposici[oó]n|acordada)\b", line, re.I)) and len(line) <= 240

def build_units(lines, source):
    units, current_norms, current_article, current_text = [], [], None, []
    ley931_context = []

    def dedupe_norms(items):
        out, seen = [], set()
        for item in items:
            key = item.get("key")
            if key not in seen:
                seen.add(key)
                out.append(item)
        return out

    def flush():
        nonlocal current_text, current_article
        txt = "\n".join(current_text).strip()
        if txt and current_article:
            units.append({
                "source": source,
                "norm_labels":[x["label"] for x in current_norms] or ["Norma no identificada"],
                "norm_keys":[x["key"] for x in current_norms],
                "article":current_article,
                "text":txt,
                "law_text":txt,
                "reg_text":"",
                "has_regulation":False
            })
        current_text = []

    for line in lines:
        ms = norm_matches(line)

        if ms and is_heading(line):
            flush()
            current_article = None
            keys = {x["key"] for x in ms}

            # El Compendio presenta la Ley 931 junto con su reglamentación
            # por Decreto 1324/93. Si aparecen en encabezados consecutivos,
            # no se debe perder la identidad de la Ley 931.
            if "931" in keys:
                ley931_context = list(ms)
                current_norms = list(ms)

            elif "1324/93" in keys and ley931_context:
                current_norms = dedupe_norms(ley931_context + ms)

            else:
                current_norms = list(ms)
                # Al entrar a otra norma principal, termina el contexto Ley 931.
                if keys & {"971", "572/85", "1613", "26206"}:
                    ley931_context = []

            continue

        am = ARTICLE_LINE_RE.match(line)
        if am:
            flush()
            current_article = am.group("num").strip()
            current_text = [line]
            continue

        if current_article:
            current_text.append(line)

    flush()
    return units

def parse_query(q):
    am = QUERY_ART_RE.search(q)
    nm = QUERY_NORM_RE.search(q)
    article = am.group("num").strip().lower() if am else None
    norm_key = normnum(nm.group("num")) if nm else None
    return article, norm_key

def build_index(units):
    docs, df = [], defaultdict(int)
    for u in units:
        c = Counter(toks(" ".join(u["norm_labels"])+" "+u["article"]+" "+u["text"]))
        docs.append(c)
        for t in c: df[t] += 1
    n=max(1,len(units))
    idf={t:math.log((n+1)/(f+1))+1 for t,f in df.items()}
    return docs,idf

def retrieve(q, units, docs, idf):
    article, norm_key = parse_query(q)
    if article and norm_key:
        exact=[u for u in units if u["article"].lower()==article and norm_key in u["norm_keys"]]
        if exact:
            return exact[:TOP_K], "exacta"

        # Compatibilidad específica del Compendio:
        # Ley 931 está presentada junto al Decreto 1324/93.
        if norm_key == "931":
            linked=[
                u for u in units
                if u["article"].lower()==article and "1324/93" in u["norm_keys"]
            ]
            if linked:
                fixed=[]
                for u in linked[:TOP_K]:
                    c=dict(u)
                    labels=list(c["norm_labels"])
                    keys=list(c["norm_keys"])
                    if "931" not in keys:
                        keys.insert(0, "931")
                        labels.insert(0, "Ley 931")
                    c["norm_keys"]=keys
                    c["norm_labels"]=labels
                    fixed.append(c)
                return fixed, "exacta vinculada"
    ids=list(range(len(units)))
    if norm_key:
        ids=[i for i,u in enumerate(units) if norm_key in u["norm_keys"]]
    if article:
        aids=[i for i in ids if units[i]["article"].lower()==article]
        if aids: return [units[i] for i in aids[:TOP_K]], "artículo"
    qt=toks(q)
    scored=[]
    for i in ids:
        c=docs[i]
        score=sum(idf.get(t,1)*(1+min(c[t],5)*0.15) for t in qt if t in c)
        if score>0: scored.append((score,i))
    scored.sort(reverse=True)
    return ([units[i] for _,i in scored[:TOP_K]], "temática") if scored else ([], "sin coincidencias")


def build_contexts(lines, source, window=10, step=5):
    """Ventanas de texto que incluyen comentarios, explicaciones e índices,
    no solo artículos. Sirven para temas como 'plazo de gracia' que pueden
    aparecer explicados fuera del texto literal del artículo."""
    contexts = []
    if not lines:
        return contexts
    for start in range(0, len(lines), step):
        end = min(len(lines), start + window)
        frag = "\n".join(lines[start:end]).strip()
        if frag:
            contexts.append({
                "source": source,
                "start": start + 1,
                "end": end,
                "text": frag
            })
        if end >= len(lines):
            break
    return contexts

def build_context_index(contexts):
    docs, df = [], defaultdict(int)
    for c in contexts:
        cnt = Counter(toks(c["text"]))
        docs.append(cnt)
        for t in cnt:
            df[t] += 1
    n = max(1, len(contexts))
    idf = {t: math.log((n + 1) / (f + 1)) + 1 for t, f in df.items()}
    return docs, idf


def meaningful_phrases(q):
    nq = normalize(q)

    for p in [
        r"^\s*que\s+es\s+el\s+",
        r"^\s*que\s+es\s+la\s+",
        r"^\s*que\s+es\s+",
        r"^\s*como\s+funciona\s+",
        r"^\s*como\s+se\s+aplica\s+",
        r"^\s*que\s+significa\s+",
    ]:
        nq = re.sub(p, "", nq)

    nq = re.sub(r"\s+y\s+como\s+funciona\s*$", "", nq)
    nq = re.sub(r"\s+y\s+como\s+se\s+aplica\s*$", "", nq)
    nq = re.sub(r"\s+y\s+para\s+que\s+sirve\s*$", "", nq)
    nq = nq.strip(" ?¿!¡.,;:")

    phrases = []
    if nq:
        phrases.append(nq)

    words = nq.split()
    for n in (4, 3, 2):
        if len(words) >= n:
            for i in range(len(words)-n+1):
                ph = " ".join(words[i:i+n]).strip()
                if len(ph) >= 6 and ph not in phrases:
                    phrases.append(ph)
    return phrases


def local_snippet_around_phrase(text, phrase, radius=900):
    nt = normalize(text)
    nph = normalize(phrase)
    pos = nt.find(nph)
    if pos < 0:
        return None

    # Mapeo aproximado: normalize conserva casi la misma longitud salvo acentos/puntuación.
    # Para mayor seguridad buscamos también en el texto original sin acentos.
    raw_norm = unicodedata.normalize("NFD", text.lower())
    raw_norm = "".join(ch for ch in raw_norm if unicodedata.category(ch) != "Mn")
    raw_pos = raw_norm.find(nph)
    if raw_pos < 0:
        raw_pos = pos

    start = max(0, raw_pos - radius)
    end = min(len(text), raw_pos + len(phrase) + radius)
    return text[start:end].strip()

def refs_ranked_by_distance(text, phrase):
    """Extrae pares artículo+norma dentro del fragmento local y los ordena por
    cercanía a la frase consultada."""
    ntext = normalize(text)
    nphrase = normalize(phrase)
    ppos = ntext.find(nphrase)
    if ppos < 0:
        ppos = len(ntext) // 2

    art_pat = re.compile(
        r"(?:art(?:í|i)?culo|art\s*\.?)\s*(?:n\s*(?:ro\s*\.?|[º°o.]?)\s*)?"
        r"(\d+(?:\s*(?:bis|ter|quater|quinquies))?)\s*[º°o.]?",
        re.I
    )
    norm_pat = re.compile(
        r"(?:decreto[\s-]*ley|ley|decreto|resoluci[oó]n|disposici[oó]n)"
        r"\s*(?:n(?:ro\.?|[º°o.]?)\s*)?(\d[\d.\-/]*)",
        re.I
    )

    arts = [(m.start(), m.group(1).strip().lower()) for m in art_pat.finditer(text)]
    norms = [(m.start(), normnum(m.group(1))) for m in norm_pat.finditer(text)]

    candidates = []
    for apos, article in arts:
        for npos, nkey in norms:
            # Solo empareja referencias razonablemente cercanas.
            if abs(apos - npos) <= 220:
                dist = min(abs(apos - ppos), abs(npos - ppos))
                candidates.append((dist, article, nkey))

    candidates.sort(key=lambda x: x[0])
    out, seen = [], set()
    for _, article, nkey in candidates:
        pair = (article, nkey)
        if pair not in seen:
            seen.add(pair)
            out.append(pair)
    return out

def exact_phrase_context(q, contexts):
    phrases = meaningful_phrases(q)
    if not phrases:
        return None, None

    best = None
    best_phrase = None
    best_score = -1

    for c in contexts:
        nt = normalize(c["text"])
        for ph in phrases:
            nph = normalize(ph)
            if nph and nph in nt:
                score = len(nph)
                if re.search(r"(?:art(?:í|i)?culo|art\s*\.?)\s*\d+", c["text"], re.I):
                    score += 15
                if re.search(r"(?:ley|decreto|resoluci[oó]n)\s*(?:n\s*(?:ro\s*\.?|[º°o.]?)\s*)?\d+", c["text"], re.I):
                    score += 10
                if score > best_score:
                    best = c
                    best_phrase = ph
                    best_score = score

    return best, best_phrase

def top_contexts(q, contexts, docs, idf, top_k=8):
    qt = toks(q)
    nq = normalize(q)
    scored = []
    for i, cnt in enumerate(docs):
        score = sum(idf.get(t, 1.0) * (1 + min(cnt[t], 5) * 0.18) for t in qt if t in cnt)
        ntext = normalize(contexts[i]["text"])
        if nq and len(nq) > 6 and nq in ntext:
            score += 15
        if score > 0:
            scored.append((score, i))
    scored.sort(reverse=True)
    return [(s, contexts[i]) for s, i in scored[:top_k]]

def extract_cross_refs(text):
    """Extrae referencias del tipo 'Art. 71 Ley 971' o 'Ley 971 ... Art. 71'
    dentro de un tramo corto de texto."""
    refs = []
    art_pat = r"(?:art(?:í|i)?culo|art\s*\.?)\s*(?:n\s*(?:ro\s*\.?|[º°o.]?)\s*)?(\d+(?:\s*(?:bis|ter|quater|quinquies))?)\s*[º°o.]?"
    norm_pat = r"(?:decreto[\s-]*ley|ley|decreto|resoluci[oó]n|disposici[oó]n)\s*(?:n(?:ro\.?|[º°o.]?)\s*)?(\d[\d.\-/]*)"

    # Artículo antes que norma.
    for m in re.finditer(art_pat + r"(?P<mid>.{0,140}?)" + norm_pat, text, re.I | re.S):
        refs.append((m.group(1).strip().lower(), normnum(m.group(3))))

    # Norma antes que artículo.
    for m in re.finditer(norm_pat + r"(?P<mid>.{0,140}?)" + art_pat, text, re.I | re.S):
        refs.append((m.group(3).strip().lower(), normnum(m.group(1))))

    # Quitar duplicados conservando orden.
    seen, out = set(), []
    for r in refs:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out

def resolve_refs(refs, units):
    found = []
    seen = set()
    for article, norm_key in refs:
        for u in units:
            key = (u["source"], tuple(u["norm_keys"]), u["article"])
            if u["article"].lower() == article and norm_key in u["norm_keys"] and key not in seen:
                seen.add(key)
                found.append(u)
    return found

def _thematic_unit_key(u):
    return (
        u.get("source", ""),
        tuple(u.get("norm_keys", [])),
        str(u.get("article", "")),
    )


def _dedupe_units(items):
    out = []
    seen = set()
    for u in items:
        key = _thematic_unit_key(u)
        if key in seen:
            continue
        seen.add(key)
        out.append(u)
    return out


def direct_phrase_units(q, units):
    """Devuelve TODOS los artículos cuyo texto contiene literalmente
    el tema consultado. Es la base de la búsqueda temática exhaustiva."""
    phrases = meaningful_phrases(q)
    if not phrases:
        return [], None

    phrase = phrases[0]
    nph = normalize(phrase)
    if len(nph) < 3:
        return [], phrase

    found = []
    for u in units:
        searchable = " ".join([
            " ".join(u.get("norm_labels", [])),
            str(u.get("article", "")),
            u.get("text", "") or "",
            u.get("law_text", "") or "",
            u.get("reg_text", "") or "",
        ])
        if nph in normalize(searchable):
            found.append(u)

    return _dedupe_units(found), phrase


def retrieve_thematic_bridge(q, contexts, cdocs, cidf, units):
    """
    Búsqueda temática exhaustiva.

    Antes se elegía una sola aparición de la frase y eso podía ocultar otros
    artículos importantes. Ahora:
      1) se buscan todos los artículos cuyo texto contiene el tema;
      2) se recorren todas las ventanas del Compendio donde aparece la frase;
      3) se resuelven todas las referencias artículo+norma encontradas;
      4) se combinan y eliminan duplicados.
    """
    direct, phrase = direct_phrase_units(q, units)

    exact_ctx, best_phrase = exact_phrase_context(q, contexts)
    phrase = best_phrase or phrase

    resolved_all = []
    matching_contexts = []

    if phrase:
        nph = normalize(phrase)
        for c in contexts:
            if nph and nph in normalize(c.get("text", "")):
                matching_contexts.append(c)
                local = local_snippet_around_phrase(c["text"], phrase, radius=1100) or c["text"]

                refs = refs_ranked_by_distance(local, phrase)
                # No nos limitamos a la primera estrategia: agregamos también
                # las referencias cruzadas detectables en el mismo fragmento.
                for ref in extract_cross_refs(local):
                    if ref not in refs:
                        refs.append(ref)

                resolved_all.extend(resolve_refs(refs, units))

    combined = _dedupe_units(direct + resolved_all)

    if combined:
        # Conservamos un fragmento explicativo legible para la cabecera.
        ctx = dict(exact_ctx or matching_contexts[0]) if (exact_ctx or matching_contexts) else None
        if ctx is not None and phrase:
            ctx["matched_phrase"] = phrase
            ctx["text"] = local_snippet_around_phrase(ctx["text"], phrase, radius=1100) or ctx["text"]
            ctx["total_referencias"] = len(combined)
            ctx["total_contextos"] = len(matching_contexts)

        return combined[:MAX_THEMATIC_HITS], ctx, "temática exhaustiva"

    # Si la frase aparece en el Compendio, pero no pudo vincularse a un artículo,
    # informamos esa coincidencia sin saltar a asociaciones lejanas.
    if exact_ctx and phrase:
        ctx = dict(exact_ctx)
        ctx["matched_phrase"] = phrase
        ctx["text"] = local_snippet_around_phrase(exact_ctx["text"], phrase, radius=1100) or exact_ctx["text"]
        return [], ctx, "frase exacta sin referencia resuelta"

    # Segunda vía: consulta sin coincidencia literal. Acumulamos referencias de
    # varios contextos relacionados en vez de quedarnos con el primer resultado.
    hits = top_contexts(q, contexts, cdocs, cidf, top_k=16)
    fallback = []
    first_ctx = None
    for score, c in hits:
        refs = extract_cross_refs(c["text"])
        resolved = resolve_refs(refs, units)
        if resolved:
            if first_ctx is None:
                first_ctx = c
            fallback.extend(resolved)

    fallback = _dedupe_units(fallback)
    if fallback:
        return fallback[:MAX_THEMATIC_HITS], first_ctx, "temática referenciada"

    return [], None, None


def sentence_around_phrase(text, phrase):
    """Devuelve la oración o tramo más cercano que contiene la frase."""
    low = normalize(text)
    ph = normalize(phrase)
    pos = low.find(ph)
    if pos < 0:
        return text[:1200].strip()

    # Como la normalización altera mínimamente las posiciones, usamos una búsqueda
    # directa insensible a mayúsculas/acentos cuando sea posible.
    raw = unicodedata.normalize("NFD", text.lower())
    raw = "".join(ch for ch in raw if unicodedata.category(ch) != "Mn")
    rpos = raw.find(ph)
    if rpos < 0:
        rpos = pos

    # Busca límites de oración razonables.
    left_candidates = [raw.rfind(".", 0, rpos), raw.rfind(";", 0, rpos), raw.rfind("\n", 0, rpos)]
    left = max(left_candidates)
    left = 0 if left < 0 else left + 1

    right_candidates = [x for x in (raw.find(".", rpos + len(ph)),
                                    raw.find(";", rpos + len(ph)),
                                    raw.find("\n", rpos + len(ph))) if x >= 0]
    right = min(right_candidates) + 1 if right_candidates else min(len(text), rpos + 900)

    frag = text[left:right].strip()
    return frag if frag else text[max(0, rpos-400):min(len(text), rpos+900)].strip()

def exact_thematic_safe_answer(question, found, context_hit):
    """Para una frase jurídica encontrada literalmente y vinculada a un artículo,
    responde solo con texto del Compendio y de la norma, sin usar el modelo."""
    phrase = context_hit.get("matched_phrase") or ""
    explanation = sentence_around_phrase(context_hit["text"], phrase)
    u = found[0]

    return (
        "MODO SEGURO — CONSULTA TEMÁTICA\n\n"
        f"Tema localizado:\n{phrase}\n\n"
        "Explicación encontrada en el Compendio:\n"
        f"{explanation}\n\n"
        f"Normativa relacionada:\n{' + '.join(u['norm_labels'])} — Artículo {u['article']}\n\n"
        "Texto normativo recuperado:\n"
        f"{u['text'].strip()}\n\n"
        "Observación:\n"
        "La vinculación entre el tema y la norma fue tomada del propio Compendio. "
        "En este modo no se usa Ollama para reinterpretar el contenido."
    )

def thematic_safe_answer(question, found, context_hit):
    """Usa la explicación del Compendio + el artículo exacto. El modelo puede
    resumir, pero se controla que no agregue numeración jurídica ajena."""
    context_text = context_hit["text"] if context_hit else ""
    blocks = []
    for i, u in enumerate(found, 1):
        blocks.append(
            f"[NORMA {i}]\n"
            f"Norma: {' + '.join(u['norm_labels'])}\n"
            f"Artículo: {u['article']}\n"
            f"Texto: {u['text']}"
        )

    system = """Sos una IA local especializada en legislación escolar de Formosa.
Tu tarea es explicar una consulta temática usando EXCLUSIVAMENTE:
1) la explicación del Compendio proporcionada;
2) el/los artículos exactos proporcionados.
No uses conocimiento externo. No inventes normas, números, plazos, excepciones ni jurisprudencia.
No menciones ningún número jurídico que no aparezca en las fuentes.
Si la explicación y el artículo no bastan, decí: "No encontré respaldo suficiente en la normativa cargada."
Respondé en español y de forma breve."""

    user = f"""PREGUNTA:
{question}

EXPLICACIÓN DEL COMPENDIO:
{context_text}

ARTÍCULO/S RELACIONADO/S:
{chr(10).join(blocks)}

FORMATO:
Tema:
[tema consultado]

Normativa relacionada:
[norma y artículo]

Explicación:
[explicación fiel a las fuentes]

Texto normativo relevante:
[síntesis fiel, sin agregar nada externo]
"""

    payload = json.dumps({
        "model": MODELO,
        "messages": [{"role":"system","content":system},{"role":"user","content":user}],
        "stream": False,
        "think": False,
        "options": {"temperature":0.0,"num_ctx":4096}
    }).encode("utf-8")
    req = urllib.request.Request(OLLAMA_URL, data=payload, headers={"Content-Type":"application/json"})
    try:
        with urllib.request.urlopen(req, timeout=600) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return ((data.get("message") or {}).get("content") or "").strip()
    except Exception as e:
        return f"ERROR al consultar Ollama: {e}"

def exact_safe_answer(question, found):
    """Para una consulta exacta norma+artículo, no usa el modelo.
    Devuelve únicamente la fuente recuperada para evitar alucinaciones."""
    u = found[0]
    body = u["text"].strip()
    # Quita solo el encabezado "ART. xx.-" al comienzo para una lectura más limpia,
    # conservando el texto íntegro restante.
    clean = re.sub(
        r"^\s*(?:art(?:í|i)?culo|art\s*\.?)\s*(?:n\s*(?:ro\s*\.?|[º°o.]?)\s*)?"
        r"\d+(?:\s*(?:bis|ter|quater|quinquies))?\s*[º°o.]?\s*[-–—.:)]*\s*",
        "",
        body,
        flags=re.I
    ).strip()

    return (
        "MODO SEGURO — CONSULTA EXACTA\n\n"
        f"Normativa encontrada:\n{' + '.join(u['norm_labels'])} — Artículo {u['article']}\n\n"
        "Respuesta basada exclusivamente en el texto recuperado:\n"
        f"{clean}\n\n"
        "Observación:\n"
        "Para consultas exactas de norma + artículo, esta versión no utiliza el modelo "
        "para reinterpretar el contenido. Así se evita que agregue leyes, números o conceptos "
        "que no estén en la fuente."
    )

def answer_has_suspicious_numbers(answer, found, question):
    """Bloquea respuestas temáticas que introduzcan números jurídicos no presentes
    en la pregunta o en las fuentes (p. ej. una ley inventada)."""
    allowed_text = question + "\n" + "\n".join(
        " ".join(u["norm_labels"]) + " " + u["article"] + " " + u["text"]
        for u in found
    )
    allowed = set(re.findall(r"\b\d{3,6}(?:/\d{2,4})?\b", allowed_text))
    produced = set(re.findall(r"\b\d{3,6}(?:/\d{2,4})?\b", answer))
    return bool(produced - allowed), produced - allowed

def main():
    print("="*78)
    print(" IA LEGISLACIÓN ESCOLAR DE FORMOSA - LOCAL (VERSIÓN 3.6 - REFERENCIAS FLEXIBLES)")
    print("="*78)
    NORMATIVA.mkdir(parents=True,exist_ok=True)
    files=sorted(f for f in NORMATIVA.glob("*.docx") if not f.name.startswith("~$"))
    if not files:
        print(f"\nNo encontré archivos .docx en {NORMATIVA}")
        input("\nEnter para salir...")
        return
    units=[]
    contexts=[]
    for f in files:
        try:
            lines = extract_docx_lines(f)
            units.extend(build_units(lines, f.name))
            contexts.extend(build_contexts(lines, f.name))
        except zipfile.BadZipFile:
            print(f"ADVERTENCIA: se omitió {f.name} porque no es un archivo .docx válido.")
        except Exception as e:
            print(f"ADVERTENCIA: no pude leer {f.name}: {e}")
    if not units:
        print("\nNo pude extraer unidades normativas de los archivos .docx.")
        print("Cierre Word si tiene abierto el Compendio y verifique que el archivo sea .docx real.")
        input("\nEnter para salir...")
        return
    docs,idf=build_index(units)
    cdocs,cidf=build_context_index(contexts)
    print(f"\nDocumentos cargados: {len(files)}")
    for f in files: print(f"  - {f.name}")
    print(f"Unidades normativas indexadas: {len(units)}")
    print(f"Modelo Ollama: {MODELO}")
    counts=Counter()
    for u in units:
        for label in u["norm_labels"]: counts[label]+=1
    print("\nPrincipales normas detectadas:")
    for label,n in counts.most_common(12):
        print(f"  - {label}: {n} artículos/unidades")
    print("\nEscriba una consulta. Para salir escriba: salir")
    while True:
        q=input("\nCONSULTA > ").strip()
        if not q: continue
        if normalize(q) in {"salir","exit","quit"}: break
        article_q, norm_q = parse_query(q)
        if article_q or norm_q:
            found,mode=retrieve(q,units,docs,idf)
            context_hit = None
        else:
            found, context_hit, mode = retrieve_thematic_bridge(q, contexts, cdocs, cidf, units)
            if not found:
                found,mode=retrieve(q,units,docs,idf)
                context_hit = None
        if not found:
            if mode == "frase exacta sin referencia resuelta" and context_hit:
                print("\nCoincidencia: frase temática exacta")
                print("\nREFERENCIA TEMÁTICA ENCONTRADA EN EL COMPENDIO:")
                print("-"*78)
                print(f'Frase localizada: "{context_hit.get("matched_phrase", "")}"')
                print(sentence_around_phrase(context_hit["text"], context_hit.get("matched_phrase", "")))
                print("-"*78)
                print("\nNo pude vincular automáticamente esa frase con un artículo exacto.")
                print("Para evitar una respuesta incorrecta, no se asociará a otra norma.")
                if context_hit.get("detected_refs"):
                    print("Referencias detectadas:", context_hit["detected_refs"])
                continue
            print("\nNo encontré la norma/artículo solicitado en el material indexado.")
            continue
        print(f"\nCoincidencia: {mode}")
        first=found[0]
        print("\nTEXTO FUENTE RECUPERADO:")
        print("-"*78)
        print(f"Norma: {' + '.join(first['norm_labels'])}")
        print(f"Artículo: {first['article']}")
        print(first["text"][:3500])
        print("-"*78)
        # En consultas exactas de norma + artículo se evita completamente la
        # reinterpretación del modelo pequeño: se devuelve la norma recuperada.
        if mode == "exacta" and article_q and norm_q:
            print("\nRESPUESTA SEGURA:")
            print("-"*78)
            print(exact_safe_answer(q, found))
            print("-"*78)
        elif mode == "frase temática exacta" and context_hit:
            print("\nREFERENCIA TEMÁTICA ENCONTRADA EN EL COMPENDIO:")
            print("-"*78)
            print(f'Frase localizada: "{context_hit.get("matched_phrase", "")}"')
            print(sentence_around_phrase(context_hit["text"], context_hit.get("matched_phrase", "")))
            print("-"*78)
            print("\nRESPUESTA TEMÁTICA SEGURA:")
            print("-"*78)
            print(exact_thematic_safe_answer(q, found, context_hit))
            print("-"*78)
        elif mode == "temática referenciada" and context_hit:
            print("\nREFERENCIA TEMÁTICA ENCONTRADA EN EL COMPENDIO:")
            print("-"*78)
            print(context_hit["text"][:2800])
            print("-"*78)
            print("\nAnalizando la referencia y el artículo exacto...")
            ans = thematic_safe_answer(q, found, context_hit)
            suspicious, extras = answer_has_suspicious_numbers(ans, found, q + "\n" + context_hit["text"])
            print("\nRESPUESTA TEMÁTICA:")
            print("-"*78)
            if not ans:
                print("Ollama no devolvió texto.")
            elif suspicious:
                print("La explicación del modelo fue descartada porque introdujo numeración jurídica "
                      "que no aparece en las fuentes.")
                print("Números detectados:", ", ".join(sorted(extras)))
                print("\nUse como respaldo la REFERENCIA TEMÁTICA y el TEXTO FUENTE recuperados.")
            else:
                print(ans)
            print("-"*78)
        else:
            print("\nAnalizando con Ollama (consulta temática general)...")
            ans=ollama_answer(q,found)
            suspicious, extras = answer_has_suspicious_numbers(ans, found, q)
            print("\nRESPUESTA DE LA IA:")
            print("-"*78)
            if not ans:
                print("Ollama no devolvió texto.")
            elif suspicious:
                print("La respuesta del modelo fue descartada por seguridad porque introdujo "
                      "números jurídicos no presentes en las fuentes recuperadas.")
                print("Números detectados:", ", ".join(sorted(extras)))
                print("\nUse como respaldo el TEXTO FUENTE RECUPERADO mostrado arriba.")
            else:
                print(ans)
            print("-"*78)

        print("\nFuentes utilizadas:")
        for u in found:
            print(f"  • {' + '.join(u['norm_labels'])} — Artículo {u['article']} — {u['source']}")

if __name__=="__main__":
    main()
