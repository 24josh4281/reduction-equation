#!/usr/bin/env python3
"""
PPT Master - Editable Equations inside an Existing Word Document

Back end of `office_equations.py inject-docx`. Takes a methodology (방법론),
business plan (사업계획서) or monitoring report the user already drafted in
Word and turns its equation markers into native Office Math (OMML), keeping
every other paragraph, style, table and image untouched.

Markers (typed in Word as ordinary text):
    {{eq:<id>}}            alone in a paragraph -> numbered display equation;
                           inside a sentence -> inline equation (spec needed)
    {{sym:<param id>}}     inline parameter symbol (spec needed)
    {{where:<scenario>}}   "여기서," symbol list of that scenario (spec needed)
    {{monitoring}}         monitoring-parameter table (spec needed)
    {{fixed}}              fixed-parameter table (spec needed)
    $$ LaTeX $$            alone in a paragraph -> numbered display equation
    $ LaTeX $              inline equation (only when it parses and holds
                           _ ^ \\ or =, so "$100" stays text)

With --convert-plain, formulas already typed as plain text are converted too:
a whole-paragraph formula such as "ERy = BEy − PEy (식 3)" with Word
subscript formatting becomes a display equation labelled (식 3), and a
subscripted symbol inside a sentence ("여기서 BEy는") becomes inline math.

Usage:
    python3 scripts/office_equations.py inject-docx draft.docx [spec.json] -o draft_eq.docx [--convert-plain] [--no-number]

Dependencies:
    None (only uses standard library)
"""

from __future__ import annotations

import bisect
import re
import zipfile
from dataclasses import dataclass, field
from html import unescape
from pathlib import Path
from typing import Any, Optional
from xml.sax.saxutils import escape

from omml_math import FORMULA_LIKE, M_NS, EquationError, docx_rpr, omml_math, parse_latex

PARTS = re.compile(r"word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$")
# Innermost paragraphs only: a paragraph holding a text box is left as is,
# while the text box's own paragraphs are processed.
PARA = re.compile(r"<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>/]).)*?</w:p>", re.S)
RUN = re.compile(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", re.S)
RPR = re.compile(r"<w:rPr>.*?</w:rPr>|<w:rPr/>", re.S)
TEXT = re.compile(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>|<w:t\s*/>")
PURE_REST = re.compile(r"(?:<w:t(?:\s[^>]*)?>[^<]*</w:t>|<w:t\s*/>|<w:lastRenderedPageBreak/>)*")
BOUNDARY = re.compile(r"</?w:(?:hyperlink|ins|del|moveFrom|moveTo|sdt|sdtContent|fldSimple|smartTag)\b")
SENTINEL = "\x00"

BLOCK = re.compile(r"\{\{(eq|where|monitoring|fixed)(?::([A-Za-z0-9_.-]+))?\}\}|\$\$([^$\x00]+)\$\$")
INLINE = re.compile(r"\{\{(eq|sym):([A-Za-z0-9_.-]+)\}\}|(?<![\\$])\$(?![\s$])([^$\n\x00]+?)(?<!\s)\$(?!\$)")
LABEL_TAIL = re.compile(r"\s*[(（]\s*((?:식|수식|Eq\.?)?\s*[0-9]+(?:[-.][0-9]+)*)\s*[)）]\s*$")
IDENT_TAIL = re.compile(r"[A-Za-zΑ-Ωα-ω]{1,8}$")
LATEX_HINT = re.compile(r"[_^\\=]")
UNICODE_TO_LATEX = {"×": r" \times ", "·": r" \cdot ", "÷": r" \div ", "–": "-", "∑": r"\sum ",
                    "Σ": r"\sum ", "{": r"\{", "}": r"\}", "%": r"\%", " ": " "}


@dataclass
class Report:
    display: int = 0
    inline: int = 0
    tables: int = 0
    converted: int = 0
    errors: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)


@dataclass
class _Run:
    start: int
    end: int
    rpr: str
    text: str
    pure: bool
    valign: Optional[str]


@dataclass
class _Ctx:
    spec: dict[str, Any]
    font: str
    size: float
    text_width: int
    number: bool
    convert_plain: bool
    report: Report
    counter: int = 0
    prefix: str = ""


# ---------------------------------------------------------------------------
# Spec lookups
# ---------------------------------------------------------------------------

def _equations(spec: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {eq["id"]: eq for sc in spec.get("scenarios", []) for eq in sc.get("equations", [])
            if eq.get("id")}


def _symbols(spec: dict[str, Any]) -> dict[str, str]:
    out = {}
    for key in ("monitoring_parameters", "fixed_parameters"):
        for param in spec.get(key, []):
            if param.get("id"):
                out[param["id"]] = param["symbol"]
    return out


def _scenario(spec: dict[str, Any], sid: str) -> Optional[dict[str, Any]]:
    return next((sc for sc in spec.get("scenarios", []) if sc.get("id") == sid), None)


# ---------------------------------------------------------------------------
# Document geometry and defaults
# ---------------------------------------------------------------------------

def _text_width(document_xml: str) -> int:
    pg = re.search(r'<w:pgSz[^>]*\bw:w="(\d+)"', document_xml)
    left = re.search(r'<w:pgMar[^>]*\bw:left="(\d+)"', document_xml)
    right = re.search(r'<w:pgMar[^>]*\bw:right="(\d+)"', document_xml)
    if not (pg and left and right):
        return 9026
    return max(3000, int(pg.group(1)) - int(left.group(1)) - int(right.group(1)))


def _doc_defaults(styles_xml: str) -> tuple[Optional[str], Optional[float]]:
    block = re.search(r"<w:rPrDefault>.*?</w:rPrDefault>", styles_xml, re.S)
    if not block:
        return None, None
    font = re.search(r'w:eastAsia="([^"]+)"', block.group(0))
    size = re.search(r'<w:sz w:val="(\d+)"', block.group(0))
    return (font.group(1) if font else None, int(size.group(1)) / 2 if size else None)


# ---------------------------------------------------------------------------
# Run parsing
# ---------------------------------------------------------------------------

def _runs(para: str) -> list[_Run]:
    runs = []
    for m in RUN.finditer(para):
        inner = m.group(1)
        rpr_m = RPR.search(inner)
        rpr = rpr_m.group(0) if rpr_m else ""
        rest = inner.replace(rpr, "", 1) if rpr else inner
        pure = PURE_REST.fullmatch(rest.strip()) is not None and "<w:t" in rest
        text = unescape("".join(t or "" for t in TEXT.findall(rest))) if pure else ""
        va = re.search(r'<w:vertAlign w:val="(subscript|superscript)"', rpr)
        runs.append(_Run(m.start(), m.end(), rpr, text, pure, va.group(1) if va else None))
    return runs


def _joined(para: str, runs: list[_Run]) -> tuple[str, list[Optional[int]]]:
    """Concatenate pure-run text; SENTINEL marks anything a marker must not cross."""
    chars: list[str] = []
    owner: list[Optional[int]] = []
    prev_end = None
    for idx, run in enumerate(runs):
        between = para[prev_end:run.start] if prev_end is not None else ""
        if not run.pure or BOUNDARY.search(between):
            chars.append(SENTINEL)
            owner.append(None)
        for ch in run.text:
            chars.append(ch)
            owner.append(idx)
        prev_end = run.end
    return "".join(chars), owner


def _run_size(rpr: str, default: float) -> float:
    m = re.search(r'<w:sz w:val="(\d+)"', rpr)
    return int(m.group(1)) / 2 if m else default


def _run_color(rpr: str) -> Optional[str]:
    m = re.search(r'<w:color w:val="([0-9A-Fa-f]{6})"', rpr)
    return m.group(1) if m else None


def _text_run(rpr: str, text: str) -> str:
    if not text:
        return ""
    clean = re.sub(r'<w:vertAlign w:val="[^"]*"/>', "", rpr) if rpr else ""
    return f'<w:r>{clean}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


# ---------------------------------------------------------------------------
# Plain-text formula conversion
# ---------------------------------------------------------------------------

def _latex_from_runs(runs: list[_Run]) -> str:
    out = []
    for run in runs:
        if not run.pure:
            continue
        text = "".join(UNICODE_TO_LATEX.get(ch, ch) for ch in run.text)
        if run.valign == "subscript" and text.strip():
            out.append("_{" + text.strip() + "}")
        elif run.valign == "superscript" and text.strip():
            out.append("^{" + text.strip() + "}")
        else:
            out.append(text)
    return "".join(out)


def _plain_display(runs: list[_Run], text: str) -> Optional[tuple[str, Optional[str]]]:
    """Return (latex, label) when the whole paragraph is a plain-text formula."""
    body = text.replace(SENTINEL, "").strip()
    if not body or len(body) > 200 or "=" not in body and "＝" not in body:
        return None
    has_script = any(r.valign for r in runs if r.pure and r.text.strip())
    if not has_script and not FORMULA_LIKE.search(body):
        return None
    if re.search(r"[가-힣]{2,}\s+[가-힣]{2,}\s+[가-힣]{2,}", body):
        return None  # a sentence that mentions a formula, not a formula line
    label_m = LABEL_TAIL.search(body)
    latex = _latex_from_runs(runs).replace("＝", "=")
    label = None
    if label_m:
        label = label_m.group(1).strip()
        latex = LABEL_TAIL.sub("", latex)
    try:
        parse_latex(latex)
    except EquationError:
        return None
    return latex.strip(), label


def _plain_inline_spans(text: str, owner: list[Optional[int]], runs: list[_Run]) -> list[tuple[int, int, str]]:
    """Symbol + Word-subscript runs inside a sentence ("BEy" with y subscripted)."""
    def valign(pos: int) -> Optional[str]:
        idx = owner[pos]
        return "boundary" if idx is None else runs[idx].valign

    spans = []
    i = 0
    while i < len(text):
        if valign(i) != "subscript" or not text[i].strip():
            i += 1
            continue
        j = i
        while j < len(text) and valign(j) == "subscript":
            j += 1
        k = j
        while k < len(text) and valign(k) == "superscript":
            k += 1
        head = IDENT_TAIL.search(text[:i])
        sub = text[i:j].strip()
        # CO2, CH4, N2O are chemical formulas and stay text
        chemical = head is not None and sub.isdigit() and head.group(0).isupper() and len(head.group(0)) <= 2
        if head and not chemical and all(valign(p) is None for p in range(head.start(), i)):
            sup = text[j:k].strip()
            latex = f"{head.group(0)}_{{{sub}}}" + (f"^{{{sup}}}" if sup else "")
            try:
                parse_latex(latex)
                spans.append((head.start(), k, latex))
            except EquationError:
                pass
        i = k
    return spans


# ---------------------------------------------------------------------------
# Paragraph rewriting
# ---------------------------------------------------------------------------

def _ppr(para: str) -> str:
    m = re.search(r"<w:pPr>.*?</w:pPr>|<w:pPr/>", para, re.S)
    return m.group(0) if m else ""


def _display_para(para: str, latex: str, label: Optional[str], ctx: _Ctx, width: int,
                  rpr_src: str = "") -> str:
    size = _run_size(rpr_src, ctx.size) + 1.0
    math = omml_math(latex, docx_rpr(size, color=_run_color(rpr_src), ea_font=ctx.font))
    mark = re.search(r"<w:rPr>.*?</w:rPr>", _ppr(para), re.S)
    mark_rpr = mark.group(0) if mark else ""  # paragraph-mark properties stay last in pPr
    spacing = '<w:spacing w:before="80" w:after="80"/>'
    if ctx.number:
        if label:
            num = re.search(r"(\d+)\s*$", label)
            if num:  # continue the document's own numbering ("식 1" -> "식 2")
                ctx.counter, ctx.prefix = int(num.group(1)), label[:num.start()]
        else:
            ctx.counter += 1
            label = f"{ctx.prefix}{ctx.counter}"
        tabs = (f'<w:tabs><w:tab w:val="center" w:pos="{width // 2}"/>'
                f'<w:tab w:val="right" w:pos="{width}"/></w:tabs>')
        label_rpr = re.sub(r'<w:vertAlign w:val="[^"]*"/>', "", rpr_src)
        return (f'<w:p><w:pPr>{tabs}{spacing}<w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>'
                f'{mark_rpr}</w:pPr><w:r><w:tab/></w:r>{math}<w:r><w:tab/></w:r>'
                f'{_text_run(label_rpr, f"({label})")}</w:p>')
    return (f'<w:p><w:pPr>{spacing}<w:jc w:val="center"/>{mark_rpr}</w:pPr>'
            f'<m:oMathPara xmlns:m="{M_NS}"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
            f'{math}</m:oMathPara></w:p>')


def _block_tables(kind: str, ident: Optional[str], ctx: _Ctx, width: int) -> Optional[str]:
    import office_equations as oe  # lazy: office_equations imports this module's caller
    spec = ctx.spec
    if kind == "where":
        scenario = _scenario(spec, ident or "")
        if not scenario or not scenario.get("where"):
            ctx.report.errors.append(f"{{{{where:{ident}}}}}: no scenario with a where-list")
            return None
        return oe.where_table(scenario["where"], ctx.font, ctx.size, total=width)
    key = "monitoring_parameters" if kind == "monitoring" else "fixed_parameters"
    if not spec.get(key):
        ctx.report.errors.append(f"{{{{{kind}}}}}: spec has no {key}")
        return None
    return oe.param_table(spec[key], kind == "fixed", ctx.font, ctx.size - 1, total=width)


def _rewrite_para(para: str, ctx: _Ctx, width: int) -> str:
    if "<m:oMath" in para:
        return para
    runs = _runs(para)
    if not runs:
        return para
    text, owner = _joined(para, runs)
    stripped = text.replace(SENTINEL, "").strip()
    first_rpr = next((r.rpr for r in runs if r.pure and r.text.strip()), "")

    block = BLOCK.fullmatch(stripped)
    if block and SENTINEL not in text.strip(SENTINEL + " "):
        kind, ident, latex = block.group(1), block.group(2), block.group(3)
        if latex is not None or kind == "eq":
            label = None
            if kind == "eq":
                eq = _equations(ctx.spec).get(ident or "")
                if not eq:
                    ctx.report.errors.append(f"unknown equation {{{{eq:{ident}}}}}")
                    return para
                latex, label = eq["latex"], eq.get("label")
            latex = latex or ""
            try:
                out = _display_para(para, latex.strip(), label, ctx, width, first_rpr)
            except EquationError as exc:
                ctx.report.errors.append(f"$$ {latex.strip()[:60]} $$: {exc}")
                return para
            ctx.report.display += 1
            return out
        table = _block_tables(kind, ident, ctx, width)
        if table is None:
            return para
        ctx.report.tables += 1
        return table

    if ctx.convert_plain:
        plain = _plain_display(runs, text)
        if plain:
            ctx.report.converted += 1
            ctx.report.display += 1
            return _display_para(para, plain[0], plain[1], ctx, width, first_rpr)

    spans = _inline_spans(text, ctx)
    if ctx.convert_plain:
        taken = [(s, e) for s, e, _ in spans]
        for s, e, latex in _plain_inline_spans(text, owner, runs):
            if not any(s < te and ts < e for ts, te in taken):
                spans.append((s, e, latex))
                ctx.report.converted += 1
    if not spans:
        return para
    return _splice(para, runs, text, owner, sorted(spans), ctx)


def _inline_spans(text: str, ctx: _Ctx) -> list[tuple[int, int, str]]:
    eqs, syms = _equations(ctx.spec), _symbols(ctx.spec)
    spans = []
    for m in INLINE.finditer(text):
        kind, ident, latex = m.group(1), m.group(2), m.group(3)
        if kind == "eq":
            if ident not in eqs:
                ctx.report.errors.append(f"unknown equation {{{{eq:{ident}}}}}")
                continue
            latex = eqs[ident]["latex"]
        elif kind == "sym":
            if ident not in syms:
                ctx.report.errors.append(f"unknown parameter {{{{sym:{ident}}}}}")
                continue
            latex = syms[ident]
        elif not LATEX_HINT.search(latex):
            continue
        try:
            parse_latex(latex)
        except EquationError as exc:
            if kind:
                ctx.report.errors.append(f"{m.group(0)}: {exc}")
            else:
                ctx.report.skipped.append(f"left as text (not LaTeX): {m.group(0)[:60]}")
            continue
        spans.append((m.start(), m.end(), latex))
    return spans


def _splice(para: str, runs: list[_Run], text: str, owner: list[Optional[int]],
            spans: list[tuple[int, int, str]], ctx: _Ctx) -> str:
    """Rebuild only the runs a span touches; math lands where the span starts."""
    pieces: dict[int, list[str]] = {}
    touched: set[int] = set()
    for s, e, _ in spans:
        touched.update(i for i in owner[s:e] if i is not None)
    starts = {s: (e, latex) for s, e, latex in spans}
    i = 0
    buf_run: Optional[int] = None
    buf = ""

    def flush() -> None:
        nonlocal buf
        if buf_run is not None and buf:
            pieces.setdefault(buf_run, []).append(_text_run(runs[buf_run].rpr, buf))
        buf = ""

    while i < len(text):
        idx = owner[i]
        if i in starts and idx is not None:  # spans never start on a SENTINEL
            end, latex = starts[i]
            flush()
            rpr = runs[idx].rpr
            rfn = docx_rpr(_run_size(rpr, ctx.size), color=_run_color(rpr), ea_font=ctx.font)
            pieces.setdefault(idx, []).append(omml_math(latex, rfn))
            ctx.report.inline += 1
            i = end
            continue
        if idx is not None and idx in touched:
            if idx != buf_run:
                flush()
                buf_run = idx
            buf += text[i]
        i += 1
    flush()
    out, pos = [], 0
    for idx, run in enumerate(runs):
        if idx not in touched:
            continue
        out.append(para[pos:run.start])
        out.append("".join(pieces.get(idx, [])))
        pos = run.end
    out.append(para[pos:])
    return "".join(out)


# ---------------------------------------------------------------------------
# Part / package level
# ---------------------------------------------------------------------------

def _cell_widths(xml: str) -> tuple[list[int], list[int], list[tuple[int, int]]]:
    opens = [m.start() for m in re.finditer(r"<w:tc[\s>]", xml)]
    closes = [m.start() for m in re.finditer(r"</w:tc>", xml)]
    widths = [(m.start(), int(m.group(1))) for m in re.finditer(r'<w:tcW w:w="(\d+)" w:type="dxa"', xml)]
    return opens, closes, widths


def rewrite_part(xml: str, ctx: _Ctx) -> str:
    opens, closes, widths = _cell_widths(xml)
    width_pos = [p for p, _ in widths]
    out, pos = [], 0
    for m in PARA.finditer(xml):
        depth = bisect.bisect_left(opens, m.start()) - bisect.bisect_left(closes, m.start())
        width = ctx.text_width
        if depth > 0:
            k = bisect.bisect_left(width_pos, m.start()) - 1
            width = max(1500, widths[k][1] - 200) if k >= 0 else ctx.text_width // 2
        new = _rewrite_para(m.group(0), ctx, width)
        out.append(xml[pos:m.start()])
        out.append(new)
        if new.startswith("<w:tbl>"):
            # 6pt gap below the table; inside a cell it also keeps the cell ending in a paragraph
            out.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="120" '
                       'w:lineRule="exact"/></w:pPr></w:p>')
        pos = m.end()
    out.append(xml[pos:])
    result = "".join(out)
    if "<m:oMath" in result and 'xmlns:m="' not in result[:2000]:
        result = re.sub(r"(<w:(?:document|ftr|hdr|footnotes|endnotes)\b)",
                        r'\1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
                        result, count=1)
    return result


def inject_docx(src: Path, spec: dict[str, Any], out: Path, number: bool = True,
                convert_plain: bool = False, font: Optional[str] = None) -> Report:
    from office_equations import require_new_output
    require_new_output(out, src)
    report = Report()
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        styles = zin.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in names else ""
        ea_font, default_size = _doc_defaults(styles)
        ctx = _Ctx(spec=spec, font=font or ea_font or spec.get("font") or "Malgun Gothic",
                   size=float(spec.get("font_size_pt") or default_size or 10.5),
                   text_width=_text_width(doc_xml), number=number,
                   convert_plain=convert_plain, report=report)
        labels = [eq.get("label", "") for eq in _equations(spec).values()]
        if any(label.startswith("식") for label in labels) or re.search(r"[(（]식\s*\d", doc_xml):
            ctx.prefix = "식 "  # auto numbers follow the document's "(식 n)" style
        patched: dict[str, bytes] = {}
        for name in names:
            data = zin.read(name)
            if PARTS.match(name):
                xml = data.decode("utf-8")
                new = rewrite_part(xml, ctx)
                if new != xml:
                    patched[name] = new.encode("utf-8")
        if report.errors:
            return report
        if not patched:
            report.errors.append("No equation markers or plain formulas found; output was not created.")
            return report
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "x", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zout.writestr(info, patched.get(info.filename, zin.read(info.filename)))
    return report
