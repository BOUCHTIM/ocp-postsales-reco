"""docx -> HTML (for Chrome headless PDF). Supports paragraphs, headings, runs (bold/italic/size/color), bullet/numbered lists,
tables (widths, shading), page breaks, header/footer text. Layout approximates Word (A4, 2 cm margins, Arial)."""
import sys, html
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def run_html(r):
    t = html.escape(r.text)
    if not t:
        if r._r.find(qn("w:br")) is not None and r._r.find(qn("w:br")).get(qn("w:type")) == "page":
            return '<div class="pb"></div>'
        return ""
    st = []
    if r.bold: st.append("font-weight:bold")
    if r.italic: st.append("font-style:italic")
    if r.font.size: st.append(f"font-size:{r.font.size.pt}pt")
    try:
        if r.font.color is not None and r.font.color.rgb is not None: st.append(f"color:#{r.font.color.rgb}")
    except Exception:
        pass
    return f'<span style="{";".join(st)}">{t}</span>'


def para_html(p, in_table=False):
    inner = "".join(run_html(r) for r in p.runs)
    if '<div class="pb"></div>' in inner and not p.text.strip():
        return '<div class="pb"></div>'
    pPr = p._p.pPr
    style = p.style.name if p.style is not None else ""
    tag = "p"
    cls = []
    if style.startswith("Heading 1"): tag, cls = "h1", ["h1"]
    elif style.startswith("Heading 2"): tag, cls = "h2", ["h2"]
    num = pPr.find(qn("w:numPr")) if pPr is not None else None
    if num is not None:
        numId = num.find(qn("w:numId")).get(qn("w:val"))
        cls.append("li-num" if NUMFMT.get(numId) == "decimal" else "li-bul")
        if NUMFMT.get(numId) == "decimal": cls.append(f"n{numId}")
    sp = ""
    if pPr is not None and pPr.find(qn("w:spacing")) is not None:
        s = pPr.find(qn("w:spacing"))
        a = s.get(qn("w:after")); b = s.get(qn("w:before"))
        if a: sp += f"margin-bottom:{int(a)/20}pt;"
        if b: sp += f"margin-top:{int(b)/20}pt;"
    al = ""
    if p.alignment is not None:
        al = {1: "center", 2: "right"}.get(int(p.alignment), "")
        if al: al = f"text-align:{al};"
    if not inner.strip():
        inner = "&nbsp;"
    return f'<{tag} class="{" ".join(cls)}" style="{sp}{al}">{inner}</{tag}>'


def table_html(t):
    cols = t._tbl.tblGrid.findall(qn("w:gridCol"))
    widths = [int(c.get(qn("w:w"))) for c in cols]
    total = sum(widths) or 1
    out = ['<table>']
    out.append("<colgroup>" + "".join(f'<col style="width:{w / total * 100:.2f}%">' for w in widths) + "</colgroup>")
    for i, row in enumerate(t.rows):
        hdr = row._tr.trPr is not None and row._tr.trPr.find(qn("w:tblHeader")) is not None
        out.append("<thead>" if hdr else "")
        cells = []
        for c in row.cells:
            fill = ""
            tcPr = c._tc.tcPr
            if tcPr is not None and tcPr.find(qn("w:shd")) is not None:
                f = tcPr.find(qn("w:shd")).get(qn("w:fill"))
                if f and f != "auto": fill = f"background:#{f};"
            inner = "".join(para_html(p, True) for p in c.paragraphs)
            cells.append(f'<td style="{fill}">{inner}</td>')
        out.append("<tr>" + "".join(cells) + "</tr>")
        out.append("</thead>" if hdr else "")
    out.append("</table>")
    return "".join(out)


NUMFMT = {}


def load_numbering(d):
    np_ = d.part.numbering_part.element
    abs_fmt = {}
    for a in np_.findall(qn("w:abstractNum")):
        lvl = a.find(qn("w:lvl"))
        abs_fmt[a.get(qn("w:abstractNumId"))] = lvl.find(qn("w:numFmt")).get(qn("w:val"))
    for n in np_.findall(qn("w:num")):
        NUMFMT[n.get(qn("w:numId"))] = abs_fmt.get(n.find(qn("w:abstractNumId")).get(qn("w:val")))


def convert(path, out):
    d = Document(path)
    load_numbering(d)
    numids = sorted({k for k, v in NUMFMT.items() if v == "decimal"})
    body = []
    for el in d.element.body.iterchildren():
        if el.tag == W + "p":
            from docx.text.paragraph import Paragraph
            body.append(para_html(Paragraph(el, d)))
        elif el.tag == W + "tbl":
            from docx.table import Table
            body.append(table_html(Table(el, d)))
    hdr = d.sections[0].header.paragraphs[0].text if d.sections[0].header.paragraphs else ""
    css = """
    @page { size: A4; margin: 20mm 20mm 18mm 20mm; }
    body { font-family: Arial, Helvetica, sans-serif; font-size: 10.5pt; color: #0B1F3A; line-height: 1.25; }
    p { margin: 0 0 6pt 0; }
    h1.h1 { font-size: 15pt; color: #0F4C81; margin: 18pt 0 8pt; page-break-after: avoid; }
    h2.h2 { font-size: 12pt; color: #0F4C81; margin: 12pt 0 6pt; page-break-after: avoid; }
    .li-bul, .li-num { margin-left: 27pt; text-indent: -13.5pt; margin-bottom: 4pt; }
    .li-bul::before { content: "•"; display: inline-block; width: 13.5pt; }
    """ + "".join(f"body {{ counter-reset: c{n}; }} .li-num.n{n} {{ counter-increment: c{n}; }} .li-num.n{n}::before {{ content: counter(c{n}) '.'; display: inline-block; width: 18pt; }}\n" for n in numids) + """
    @page { @top-right { content: "HDRTEXT"; font: 8pt Arial; color: #4A5568; } @bottom-center { content: "Page " counter(page); font: 8pt Arial; color: #4A5568; } }
    .pb { page-break-after: always; }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; margin: 0 0 4pt 0; font-size: 9pt; page-break-inside: auto; }
    thead { display: table-header-group; }
    tr { page-break-inside: avoid; }
    td { border: 1px solid #C9D3DE; padding: 3pt 4.5pt; vertical-align: middle; word-wrap: break-word; }
    td p { margin: 0; }
    .hdr { position: running(hdr); }
    """
    css = css.replace("HDRTEXT", hdr.replace('"', ""))
    doc = f"""<!doctype html><html><head><meta charset="utf-8"><title>Rapport</title><style>{css}</style></head>
<body>{"".join(body)}</body></html>"""
    open(out, "w").write(doc)
    print("html written", out, "header:", hdr)


if __name__ == "__main__":
    convert(sys.argv[1], sys.argv[2])
