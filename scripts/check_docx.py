"""Collect the facts needed to check a Word document against СТО 02068410-004-2018.

Usage:
    python check_docx.py report.docx            # .docx only
    python check_docx.py report.doc             # .doc/.docm/.rtf/.odt: converted first (Word or LibreOffice)
    python check_docx.py report.docx --pdf      # also render a PDF and check what Word actually prints

The script never edits the input; conversions go to a temp folder next to nothing.
Output is a plain-text report: measured values plus "!" lines for likely violations.
The "!" lines are candidates for Claude to confirm, not final verdicts.

Needs: python-docx. Optional: pdftotext (poppler) for the --pdf checks.
"""
import os, re, shutil, subprocess, sys, tempfile
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
try:
    import docx
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("python-docx is missing: pip install python-docx")

W = lambda t: qn("w:" + t)
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


# ---------------------------------------------------------------- conversion
def convert(src, fmt, outdir):
    """Convert src to fmt ('docx' or 'pdf') with LibreOffice, or Word on Windows. Returns the new path."""
    base = os.path.splitext(os.path.basename(src))[0]
    dst = os.path.join(outdir, f"{base}.{fmt}")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", fmt, "--outdir", outdir, src],
                       check=True, capture_output=True, timeout=300)
    elif os.name == "nt":
        code = 16 if fmt == "docx" else 17  # wdFormatDocumentDefault / wdFormatPDF
        ps = (f"$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;"
              f"try{{$d=$w.Documents.Open('{os.path.abspath(src)}',$false,$true);"
              f"$d.SaveAs2('{os.path.abspath(dst)}',{code});$d.Close($false)}}finally{{$w.Quit()}}")
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True, timeout=300)
    else:
        sys.exit(f"cannot convert {src}: install LibreOffice (soffice)")
    if not os.path.exists(dst):
        sys.exit(f"conversion to {fmt} failed for {src}")
    return dst


# ---------------------------------------------------------------- style resolution
class Doc:
    def __init__(self, path):
        self.d = docx.Document(path)
        dd = self.d.styles.element.find(W("docDefaults"))
        self.dd_r = dd.find(f"{W('rPrDefault')}/{W('rPr')}") if dd is not None else None
        self.dd_p = dd.find(f"{W('pPrDefault')}/{W('pPr')}") if dd is not None else None

    @staticmethod
    def chain(style):
        while style is not None:
            yield style.element
            style = style.base_style

    def run_props(self, p, r):
        srcs = [r._r.rPr]
        rs = r._r.rPr.find(W("rStyle")) if r._r.rPr is not None else None
        if rs is not None:
            st = next((s for s in self.d.styles if s.style_id == rs.get(W("val"))), None)
            if st is not None:
                srcs += [e.find(W("rPr")) for e in self.chain(st)]
        srcs += [e.find(W("rPr")) for e in self.chain(p.style)] + [self.dd_r]
        font = sz = b = None
        for s in srcs:
            if s is None:
                continue
            if font is None and (f := s.find(W("rFonts"))) is not None:
                font = next((f.get(W(a)) for a in ("hAnsi", "ascii", "hAnsiTheme", "asciiTheme") if f.get(W(a))), None)
            if sz is None and (e := s.find(W("sz"))) is not None:
                sz = int(e.get(W("val"))) / 2
            if b is None and (e := s.find(W("b"))) is not None:
                b = e.get(W("val")) not in ("0", "false")
        return font, sz or 10, bool(b)

    def para_props(self, p):
        srcs = [p._p.pPr] + [e.find(W("pPr")) for e in self.chain(p.style)] + [self.dd_p]
        def first(tag, need=None):
            for s in srcs:
                if s is None:
                    continue
                e = s.find(W(tag))
                if e is not None and (need is None or any(e.get(W(a)) for a in need)):
                    return e
            return None
        jc = first("jc"); jc = jc.get(W("val")) if jc is not None else "left"
        ind = first("ind", ("firstLine", "hanging"))
        if ind is not None:
            ind = round(int(ind.get(W("firstLine"))) / 567, 2) if ind.get(W("firstLine")) else -round(int(ind.get(W("hanging"))) / 567, 2)
        sp = first("spacing", ("line",))
        if sp is not None:
            v, rule = int(sp.get(W("line"))), sp.get(W("lineRule")) or "auto"
            sp = round(v / 240, 2) if rule == "auto" else f"{v / 20}pt-{rule}"
        ol = first("outlineLvl"); ol = int(ol.get(W("val"))) if ol is not None else None
        return dict(jc=jc, ind=ind, sp=sp, ol=ol, num=first("numPr"))

    def main_run(self, p):
        c = Counter()
        for r in p.runs:
            if r.text.strip():
                c[self.run_props(p, r)] += len(r.text)
        return c.most_common(1)[0][0] if c else (None, None, False)


# ---------------------------------------------------------------- helpers
STRUCTURAL = ("РЕФЕРАТ", "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "ОПРЕДЕЛЕНИЯ", "ОБОЗНАЧЕНИЯ", "СОКРАЩЕНИЯ", "СПИСОК", "ПРИЛОЖЕНИЕ")

def is_toc(p):
    """A table-of-contents entry (not the СОДЕРЖАНИЕ title itself)."""
    name = p.style.name.lower()
    return name.startswith(("toc", "оглавление", "содержание")) and "heading" not in name and "заголов" not in name

def toc_title(t):
    parts = [x.strip() for x in t.split("\t") if x.strip() and not re.fullmatch(r"[.…\s]+", x)]
    if parts and re.fullmatch(r"\d+", parts[-1]):
        parts = parts[:-1]
    s = " ".join(parts)
    return re.sub(r"\s*[.…]{3,}\s*\d*\s*$", "", s).strip()

def is_heading(D, p):
    name = p.style.name.lower()
    return D.para_props(p)["ol"] is not None and D.para_props(p)["ol"] < 9 or name.startswith(("heading", "заголовок"))

def num_key(n):
    return tuple(int(x) for x in n.split("."))

def seq_problems(nums):
    """nums: list of 'N' or 'N.M' strings in document order. Returns duplicates and gaps."""
    out, seen = [], Counter(nums)
    out += [f"duplicate {n} (x{c})" for n, c in seen.items() if c > 1]
    groups = {}
    for n in dict.fromkeys(nums):
        prefix, last = (n.rsplit(".", 1) if "." in n else ("", n))
        groups.setdefault(prefix, []).append(int(last))
    for prefix, lst in groups.items():
        for a, b in zip(lst, lst[1:]):
            if b != a + 1:
                out.append(f"jump {prefix + '.' if prefix else ''}{a} -> {prefix + '.' if prefix else ''}{b}")
        if lst and lst[0] != 1:
            out.append(f"starts at {prefix + '.' if prefix else ''}{lst[0]}")
    return out

def heading_seq_problems(nums):
    """Check 1, 1.1, 1.1.1 ... nesting: each number must be the next sibling or first child."""
    out, prev = [], ()
    for n in nums:
        k = num_key(n)
        ok = (len(k) <= len(prev) + 1 and k[:-1] == prev[:len(k) - 1] and
              (k[-1] == (prev[len(k) - 1] + 1 if len(k) <= len(prev) else 1)))
        if not ok and prev:
            out.append(f"{'.'.join(map(str, prev))} -> {n}")
        prev = k
    return out


# ---------------------------------------------------------------- checks
def check(path, pdf_text=None):
    D = Doc(path)
    d = D.d
    paras = [Paragraph(e, d) for e in d.element.body.iter(W("p"))]
    in_table = lambda p: any(a.tag == W("tbl") for a in p._p.iterancestors())
    body = [p for p in paras if not in_table(p)]
    txt = lambda p: p.text.strip()

    # landmarks (skip TOC entries)
    def find(pred):
        return next((i for i, p in enumerate(body) if not is_toc(p) and pred(txt(p))), None)
    i_ref = find(lambda t: t.upper() == "РЕФЕРАТ")
    i_toc = find(lambda t: t.upper() in ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"))
    i_intro = find(lambda t: t.upper() == "ВВЕДЕНИЕ")
    i_concl = next((i for i, p in enumerate(body) if i_intro and i > i_intro and not is_toc(p) and txt(p).upper() == "ЗАКЛЮЧЕНИЕ"), None)
    i_src = next((i for i, p in enumerate(body) if i_intro and i > i_intro and not is_toc(p) and len(txt(p)) < 60
                  and "СПИСОК" in txt(p).upper() and re.search("ИСТОЧНИК|ЛИТЕРАТУР", txt(p).upper())), None)
    apps = [i for i, p in enumerate(body) if i_intro and i > i_intro and not is_toc(p) and re.fullmatch(r"ПРИЛОЖЕНИЕ\s+[А-ЯA-Z0-9]", txt(p).upper())]
    end_main = apps[0] if apps else len(body)
    main = body[i_intro or 0:end_main]
    main_text = "\n".join(p.text for p in main if p._p.find(f".//{M}oMath") is None)
    main_text = re.sub(r"https?://\S+", "URL", main_text)

    print("== PAGE SETUP (п. 3)")
    for i, s in enumerate(d.sections):
        fx = s.footer._element.xml
        print(f"section {i}: {s.page_width.mm:.0f}x{s.page_height.mm:.0f} mm, margins L{s.left_margin.mm:.0f} R{s.right_margin.mm:.0f} "
              f"T{s.top_margin.mm:.0f} B{s.bottom_margin.mm:.0f}, PAGE field in footer={'PAGE' in fx}, title page separate={s.different_first_page_header_footer}")
        if (round(s.left_margin.mm), round(s.right_margin.mm), round(s.top_margin.mm), round(s.bottom_margin.mm)) != (30, 15, 20, 20):
            print("! margins must be L30 R15 T20 B20 mm (п. 3.2)")

    print("\n== BODY TEXT (main part, paragraphs > 80 chars, share of characters)")
    c = Counter()
    for p in main:
        if len(txt(p)) < 80 or is_heading(D, p):
            continue
        pp = D.para_props(p)
        c[D.main_run(p) + (pp["jc"], pp["ind"], pp["sp"])] += len(txt(p))
    tot = sum(c.values()) or 1
    for (font, sz, b, jc, ind, sp), v in c.most_common(6):
        print(f"{v * 100 / tot:5.1f}%  {font} {sz}pt bold={b} align={jc} indent={ind}cm spacing={sp}")
    for (font, sz, b, jc, ind, sp), v in c.items():
        if v / tot > 0.02:
            bad = [m for m, cond in (("font not Times New Roman", font != "Times New Roman"), ("size outside 12-14", not 12 <= sz <= 14),
                                     ("bold", b), ("not justified", jc != "both"), ("indent != 1.25 cm", ind != 1.25), ("spacing != 1.5", sp != 1.5)) if cond]
            if bad:
                print(f"! {v * 100 / tot:.0f}% of text: {', '.join(bad)} (п. 3.2)")
    bold = sum(len(r.text) for p in main if not is_heading(D, p) for r in p.runs if r.text.strip() and D.run_props(p, r)[2])
    if bold:
        print(f"! {bold} bold characters in the main part (п. 3.2: bold is not used)")

    print("\n== STRUCTURE (п. 4.1)")
    order = [("РЕФЕРАТ", i_ref), ("СОДЕРЖАНИЕ", i_toc), ("ВВЕДЕНИЕ", i_intro), ("ЗАКЛЮЧЕНИЕ", i_concl), ("СПИСОК ИСТОЧНИКОВ", i_src)] + \
            [(txt(body[i]), i) for i in apps]
    print("found:", ", ".join(f"{n}@{i}" for n, i in order if i is not None) or "nothing")
    missing = [n for n, i in order[:4] if i is None]
    if missing:
        print("! not found:", ", ".join(missing))
    found = [i for _, i in order if i is not None]
    if found != sorted(found):
        print("! structural elements out of order")
    for n, i in order:
        if i is None:
            continue
        p = body[i]
        pp, (font, sz, b) = D.para_props(p), D.main_run(p)
        bad = [m for m, cond in (("not centred", pp["jc"] != "center"), ("ends with period", txt(p).endswith(".")), ("bold", b)) if cond]
        if bad:
            print(f"! heading {txt(p)[:40]!r}: {', '.join(bad)} (п. 4.1.2)")

    print("\n== HEADINGS (п. 4.1.4-4.1.7)")
    heads = [p for p in main if is_heading(D, p) and txt(p) and not txt(p).upper().startswith(STRUCTURAL)]
    manual = [re.match(r"^(\d+(?:\.\d+)*)\.?\s", txt(p)) for p in heads]
    auto = sum(1 for p, m in zip(heads, manual) if not m and D.para_props(p)["num"] is not None)
    print(f"{len(heads)} headings; {len(heads) - auto} with typed numbers, {auto} numbered automatically by Word")
    for p in heads:
        t, pp = txt(p), D.para_props(p)
        bad = [m for m, cond in (("trailing period", t.endswith(".")), ("centred", pp["jc"] == "center"), ("no paragraph indent", not pp["ind"]),
                                 ("period after number", bool(re.match(r"^\d+(\.\d+)*\.\s", t)))) if cond]
        if bad:
            print(f"! {t[:60]!r}: {', '.join(bad)}")
    nums = [m.group(1) for m in manual if m]
    if pdf_text and auto:  # take the numbers Word actually prints, matching headings in document order
        lines = [re.sub(r"\s+", " ", l.strip()) for l in pdf_text.splitlines()]
        nums, pos = [], 0
        for p in heads:
            title = re.sub(r"\s+", " ", re.sub(r"^\d+(\.\d+)*\.?\s*", "", txt(p)))[:30]
            for j in range(pos, len(lines)):
                m = re.match(r"^(\d+(?:\.\d+)*)\.?\s(.*)", lines[j])
                if m and title and m.group(2).startswith(title) and not re.search(r"\.{4,}|…", lines[j]):
                    nums.append(m.group(1)); pos = j + 1
                    break
    elif auto:
        nums = []
        print("  (some numbers are automatic: rerun with --pdf to check numbering)")
    for prob in heading_seq_problems(nums):
        print(f"! heading numbering breaks: {prob} (п. 4.1.7)")

    print("\n== EMPTY NUMBERED PARAGRAPHS (print as a stray '1.' or bullet)")
    stray = [i for i, p in enumerate(body) if not txt(p) and D.para_props(p)["num"] is not None and not p._p.xpath(".//w:drawing|.//w:pict")]
    for i in stray:
        near = next((txt(q) for q in body[i + 1:i + 4] if txt(q)), "")
        print(f"! empty auto-numbered paragraph before {near[:50]!r}")
    if not stray:
        print("none")

    print("\n== CAPTIONS (п. 6.1, 6.2)")
    caps = {"Рисунок": [], "Таблица": []}
    for p in paras:
        m = re.match(r"^(Рисунок|Таблица)\s+([А-ЯA-Z]?\.?\d+(?:\.\d+)?)\s*(.?)", txt(p))
        if not m or is_toc(p):
            continue
        kind, n, sep = m.groups()
        caps[kind].append(n)
        pp = D.para_props(p)
        bad = []
        if sep and sep != "–":
            bad.append(f"separator {sep!r}, use en dash ' – '")
        if txt(p).endswith("."):
            bad.append("trailing period")
        if kind == "Рисунок" and pp["jc"] != "center":
            bad.append("not centred")
        if kind == "Рисунок" and pp["ind"]:
            bad.append(f"first-line indent {pp['ind']} cm shifts the centred caption")
        if kind == "Таблица" and (pp["jc"] == "center" or pp["ind"]):
            bad.append("table title must be left-aligned with no indent")
        print(("! " if bad else "  ") + f"{txt(p)[:70]!r}" + (f": {', '.join(bad)}" if bad else ""))
    for kind, lst in caps.items():
        plain = [n for n in lst if not re.match(r"^[А-ЯA-Z]", n)]  # appendix ones are numbered per appendix
        for prob in seq_problems(plain):
            print(f"! {kind} numbering: {prob} (п. 6.1.5 / 6.2.6)")
        low = main_text.lower()
        stem = "рисун" if kind == "Рисунок" else "таблиц"
        caption_lines = {txt(p) for p in paras if re.match(r"^(Рисунок|Таблица)\s", txt(p))}
        body_low = "\n".join(l for l in low.split("\n") if l.strip() not in {c.lower() for c in caption_lines})
        unref = [n for n in dict.fromkeys(plain) if not re.search(
            rf"{stem}\w*\s+(?:[\d.]+\s*(?:,|и|–|-)\s*)*{re.escape(n)}(?!\d)(?!\.\d)", body_low)]
        if unref:
            print(f"! {kind} not referenced in the text: {', '.join(unref)} (п. 6.1.3 / 6.2.3)")

    print("\n== FORMULAS (п. 6.3)")
    fpar = [p for p in paras if p._p.find(f".//{M}oMath") is not None or "Equation." in p._p.xml]
    print(f"{len(fpar)} formula paragraphs")
    gde = [txt(p) for p in paras if re.match(r"^где\b", txt(p))]
    if any(t.startswith("где:") for t in gde):
        print("! 'где:' with a colon; write 'где' (п. 6.3.2)")
    fnums = [m.group(1) for p in fpar if (m := re.search(r"\((\d+(?:\.\d+)?)\)\s*$", p.text))]
    for prob in seq_problems(fnums):
        print(f"! formula numbering: {prob}")
    if re.search(r"формул\w*\s+\d", main_text):
        print("! formula referenced without parentheses, write 'по формуле (1)' (п. 6.3.4)")

    print("\n== CITATIONS (п. 6.4.1)")
    refs = []
    for m in re.finditer(r"\[(\d+(?:\s*[,–-]\s*\d+)*)\]", main_text):
        for part in re.split(r"\s*,\s*", m.group(1)):
            a, b = (re.split(r"\s*[–-]\s*", part) + [None])[:2]
            refs += list(range(int(a), int(b) + 1)) if b else [int(a)]
    first = list(dict.fromkeys(refs))
    print("first-mention order:", first[:40])
    if first != sorted(first):
        print("! sources are not cited in order of first mention")

    print("\n== SOURCES (п. 5.9)")
    src = []
    if i_src is not None:
        for p in body[i_src + 1:end_main]:
            if txt(p):
                src.append(p)
    print(f"{len(src)} entries")
    if first and src and max(first) != len(src):
        print(f"! highest citation [{max(first)}] but {len(src)} entries")
    never = sorted(set(range(1, len(src) + 1)) - set(refs))
    if src and never:
        print(f"! entries never cited: {never}")
    numfmt = Counter()
    numbering = d.part.numbering_part.element if any("numbering" in r.reltype for r in d.part.rels.values()) else None
    for p in src:
        t = txt(p)
        if re.match(r"^\d+\.\s", t):
            numfmt["typed 'N.'"] += 1
        elif (num := D.para_props(p)["num"]) is not None and numbering is not None:
            nid = num.find(W("numId"))
            lvl = num.find(W("ilvl"))
            nm = numbering.find(f"{W('num')}[@{W('numId')}='{nid.get(W('val'))}']") if nid is not None else None
            an = numbering.find(f"{W('abstractNum')}[@{W('abstractNumId')}='{nm.find(W('abstractNumId')).get(W('val'))}']") if nm is not None else None
            L = an.find(f"{W('lvl')}[@{W('ilvl')}='{lvl.get(W('val')) if lvl is not None else '0'}']") if an is not None else None
            lt = L.find(W("lvlText")).get(W("val")) if L is not None and L.find(W("lvlText")) is not None else "?"
            numfmt[f"auto {lt!r}"] += 1
    print("numbering:", dict(numfmt))
    if any("." in k for k in numfmt):
        print("! source numbers must have no period: '1 Фамилия...' (п. 5.9.1)")
    issues = Counter()
    for p in src:
        t = txt(p)
        if "http" in t and "URL:" not in t:
            issues["link without 'URL:'"] += 1
        if re.search(r"[Дд]ата обращения(?!:)", t):
            issues["'дата обращения' without colon"] += 1
        if "Дата обращения" in t:
            issues["capitalised 'Дата обращения'"] += 1
        if re.search(r"\s-\s", t):
            issues["hyphen ' - ' instead of en dash ' – '"] += 1
        if "/ URL" in t:
            issues["'/ URL:' instead of '. – URL:'"] += 1
    for k, v in issues.items():
        print(f"! {v} entries: {k} (ГОСТ 7.1)")
    for p in src[:3]:
        print("  e.g.", txt(p)[:130])

    print("\n== APPENDICES (п. 5.10)")
    for i in apps:
        letter = txt(body[i]).split()[-1]
        nxt = next((txt(q) for q in body[i + 1:i + 3] if txt(q)), "")
        refd = re.search(rf"приложени\w*\s+{letter.lower()}\b", main_text.lower())
        bad = [m for m, cond in (("not referenced in the text", not refd), ("forbidden letter", letter in "ЁЗЙОЧЪЫЬ")) if cond]
        print(("! " if bad else "  ") + f"ПРИЛОЖЕНИЕ {letter} — {nxt[:50]!r}" + (f": {', '.join(bad)} (п. 5.10.4/5.10.6)" if bad else ""))

    print("\n== CONTENTS vs HEADINGS (п. 5.4)")
    toc = [toc_title(txt(p)) for p in paras if is_toc(p) and txt(p)]
    if toc:
        norm = lambda s: re.sub(r"^\d+(\.\d+)*\.?\s*", "", re.sub(r"^приложение\s+\S+\.?\s*", "", s.lower())).strip(" .")
        tlist = [norm(t) for t in toc]
        # compare by prefix: long entries wrap or get cut in hand-typed contents
        miss = [txt(p) for p in heads if not any(t[:25] and norm(txt(p)).startswith(t[:25]) for t in tlist)]
        if miss:
            print("! headings missing from contents:", [m[:50] for m in miss])
        caps_toc = [t for t in toc if t.isupper() and not re.fullmatch(r"ПРИЛОЖЕНИЕ\s+\S+", t)]
        if caps_toc:
            print("! contents entries in CAPS, write sentence case (п. 5.4.1):", caps_toc)
        bare = [t for t in toc if re.fullmatch(r"(?i)приложение\s+\S+", t)]
        if bare:
            print("! appendix listed without its title (п. 5.10.9):", bare)
        if not (miss or caps_toc or bare):
            print(f"{len(toc)} entries, consistent")
    else:
        print("no TOC-styled paragraphs (typed by hand? check page numbers manually)")

    print("\n== РЕФЕРАТ (п. 5.3)")
    if i_ref is not None:
        block = [txt(p) for p in body[i_ref + 1:(i_toc or i_ref + 15)] if txt(p)]
        head = next((t for t in block if re.match(r"^(Пояснительная записка|Отч[её]т|Курсов|Выпускная)", t)), "")
        print("  ", head[:150])
        def count(word):
            m = re.search(rf"(\d+)\s+{word}", head)
            return int(m.group(1)) if m else None
        real = {"рисун": len(set(n for n in caps["Рисунок"] if not re.match(r"^[А-ЯA-Z]", n))),
                "таблиц": len(set(n for n in caps["Таблица"] if not re.match(r"^[А-ЯA-Z]", n))),
                "источник": len(src), "приложени": len(apps)}
        for w, n in real.items():
            stated = count(w)
            if stated is not None and stated != n:
                print(f"! реферат says {stated} {w}…, document has {n}")
            elif stated is None and n:
                print(f"! реферат does not state the number of {w}… ({n} in document)")
        if pdf_text is not None:
            pages = pdf_text.count("\f") + (0 if pdf_text.endswith("\f") else 1)
            stated = re.search(r"(\d+)\s*(с\.?|стр|страниц)", head)
            if stated and int(stated.group(1)) != pages:
                print(f"! реферат says {stated.group(1)} pages, PDF has {pages}")
        if re.search(r"\d+\s+(страниц|стр\b)|\d+\s+с(?!\.)", head):
            print("! write pages as '85 с.' (Appendix В)")
        kw = next((t for t in block if t.isupper() and "," in t), None)
        if kw:
            k = len([x for x in kw.split(",") if x.strip()])
            print(f"   keywords: {k}")
            if not 5 <= k <= 15:
                print("! 5-15 keywords required (п. 5.3.2)")
            if kw.endswith("."):
                print("! no period after the keyword list (п. 5.3.2)")
        else:
            print("! keyword line in CAPS not found (п. 5.3.2)")
        text_len = sum(len(t) for t in block if t is not kw and t is not head and not t.startswith("Графическая"))
        print(f"   abstract text: {text_len} characters")
        if text_len > 850:
            print("! longer than the recommended 850 characters (п. 5.3.4)")
        boldrun = [r.text for p in body[i_ref + 1:(i_toc or i_ref + 15)] for r in p.runs if r.text.strip() and D.run_props(p, r)[2]]
        if boldrun:
            print("! bold in реферат:", "".join(boldrun)[:60])
    else:
        print("! РЕФЕРАТ not found")

    print("\n== TEXT RULES (п. 7, main part only; candidates, confirm by eye)")
    def ctx(pat, n=5):
        return [main_text[max(0, m.start() - 30):m.end() + 15].replace("\n", " ") for m in re.finditer(pat, main_text)][:n]
    rules = [
        ("decimal point, use a comma (п. 7.12)", r"(?<![\w.])\d+\.\d+(?=\s?(%|[а-яА-Я]{1,4}\b|\$))(?<!\d\.\d\.\d)"),
        ("range with a hyphen, write 'от 10 до 20 Гц' (п. 7.9)", r"(?<![\w.])\d+(?:[.,]\d+)?\s?[-–]\s?\d+(?:[.,]\d+)?\s?(?:[кМмГ]?Гц|дБ\w*|мм|см|м|кг|В|мВ|с|%)(?!\w)"),
        ("sign without a number: ~ ≈ ≥ ≤ (п. 7.3)", r"[~≈≥≤]\s?\d"),
        ("'%' or '№' without a number (п. 7.3)", r"(?<![\d\s])\s?%|№(?!\s?\d)"),
        ("bare 'ГОСТ' without its number (п. 7.3)", r"ГОСТ(?![\s ]*Р?[\s ]*\d)"),
        ("number glued to its unit, add a (non-breaking) space", r"(?<![\w.])\d+(?:Гц|кГц|дБ|мм|см|кг|мВ|мс)\b"),
        ("negative number with '-', write 'минус' in text (п. 7.3)", r"(?<![\w\d)])[-−]\d+(?:[.,]\d+)?\s?(?:[а-яА-Я]|%)"),
        ("'глава' — the standard has разделы/подразделы/пункты", r"\bглав[аеуы]\s+\d"),
    ]
    for label, pat in rules:
        hits = ctx(pat)
        if hits:
            print(f"! {label}: {len(re.findall(pat, main_text))} hit(s)")
            for h in hits[:3]:
                print("     …" + h + "…")

    if pdf_text is not None:
        print("\n== RENDERED PDF")
        pages = pdf_text.split("\f")
        if pages and not pages[-1].strip():
            pages = pages[:-1]
        foot = [next((l.strip() for l in reversed(pg.splitlines()) if l.strip()), "") for pg in pages]
        nums = [f if re.fullmatch(r"\d+", f) else "·" for f in foot]
        print(f"{len(pages)} pages; footer numbers: {' '.join(nums[:12])}{' …' if len(nums) > 12 else ''}")
        if nums and nums[0] != "·":
            print("! number printed on the title page (п. 4.2.2)")
        expected = [str(i) for i in range(2, len(nums) + 1)]
        if nums[1:] != expected:
            bad = [f"p{i + 2}={n}" for i, (n, e) in enumerate(zip(nums[1:], expected)) if n != e][:6]
            print("! page numbers not continuous:", bad, "(check: landscape pages or images can hide the number)")
        empty = [i + 1 for i, pg in enumerate(pages) if len(pg.strip()) < 5]
        if empty:
            print("  pages without text (scans/images, or accidentally blank — look at them):", empty)


def selftest():
    assert seq_problems(["1", "2", "3"]) == []
    assert seq_problems(["1", "2", "2", "4"]) == ["duplicate 2 (x2)", "jump 2 -> 4"]
    assert seq_problems(["1.1", "1.2", "2.1"]) == []
    assert heading_seq_problems(["1", "1.1", "1.1.1", "1.2", "2", "2.1"]) == []
    assert heading_seq_problems(["2.2", "2.2.1", "2.2.3"]) == ["2.2.1 -> 2.2.3"]
    assert toc_title("1\tСтруктура базы данных\t7") == "1 Структура базы данных"
    assert toc_title("Введение\t……………\t7") == "Введение"
    print("selftest ok")


def main():
    if "--selftest" in sys.argv:
        return selftest()
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        sys.exit(__doc__)
    src = args[0]
    tmp = tempfile.mkdtemp(prefix="sto_check_")
    path = src if src.lower().endswith(".docx") else convert(src, "docx", tmp)
    pdf_text = None
    if "--pdf" in sys.argv:
        if not shutil.which("pdftotext"):
            print("(pdftotext not found: install poppler for the rendered-PDF checks)\n")
        else:
            pdf = convert(path, "pdf", tmp)
            out = os.path.join(tmp, "out.txt")
            subprocess.run(["pdftotext", "-layout", "-enc", "UTF-8", pdf, out], check=True)
            pdf_text = open(out, encoding="utf-8").read()
            print(f"(PDF rendered to {pdf} — open it to look at pages)\n")
    check(path, pdf_text)


if __name__ == "__main__":
    main()
