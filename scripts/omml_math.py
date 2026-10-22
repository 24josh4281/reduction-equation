#!/usr/bin/env python3
"""
PPT Master - Editable Equation (OMML) Builder

Converts a restricted LaTeX subset into Office Math Markup (OMML) so that
equations land in PowerPoint and Word as native, editable equation objects
(Insert > Equation), never as pictures. The same parse tree also yields a
Unicode linear fallback and SVG preview tspans for the SVG authoring stage.

Supported LaTeX subset (enough for GHG methodology / reduction-project math):
    groups {..}, sub/superscripts _ ^, \\frac \\dfrac \\tfrac, \\sqrt[n]{..},
    \\sum \\prod \\int with limits, \\left( .. \\right) delimiters,
    \\text{..} \\mathrm{..} \\textrm{..} \\operatorname{..} (upright text),
    \\bar \\hat \\tilde \\overline accents, Greek letters, common operators
    (\\times \\cdot \\pm \\le \\ge \\neq \\approx \\rightarrow \\infty ...),
    function names (\\ln \\log \\exp \\max \\min), spacing (\\, \\; \\quad),
    Hangul and other non-Latin text (kept upright, one run per word).
Anything outside the subset raises EquationError instead of guessing.

Usage:
    python3 scripts/omml_math.py check "<latex>"
    python3 scripts/omml_math.py omml "<latex>" [--target pptx|docx] [--size-pt 12]
    python3 scripts/omml_math.py linear "<latex>"
    python3 scripts/omml_math.py preview "<latex>" [--font-size 20]

Examples:
    python3 scripts/omml_math.py check "ER_y = BE_y - PE_y - LE_y"
    python3 scripts/omml_math.py omml "BE_y = \\sum_{i} EG_{i,y} \\times EF_{grid,y}" --target docx

Dependencies:
    None (only uses standard library)
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from typing import Callable, Optional, Union
from xml.sax.saxutils import escape


M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MATH_FONT = "Cambria Math"


# Formula-like plain text that should have been an equation: a subscripted
# identifier (BE_y, EF_{grid,y}) on either side of '=', a Σ, or raw LaTeX.
FORMULA_LIKE = re.compile(
    r"(?:[A-Za-z]{1,6}_(?:\{[^}]+\}|[A-Za-z0-9,]+)\s*[=＝]"
    r"|[=＝]\s*[A-Za-z]{1,6}_(?:\{[^}]+\}|[A-Za-z0-9]+)|∑|\\sum|\\frac)"
)

# A short "X = A − B" line written in words (감축량 = 베이스라인 배출량 − 사업 배출량).
WORD_FORMULA = re.compile(r"^[^=＝]{1,30}[=＝][^=＝]{1,90}[−\-×+÷/][^=＝]{0,90}$")


class EquationError(ValueError):
    """Raised when LaTeX falls outside the supported subset."""


# ---------------------------------------------------------------------------
# Symbol tables
# ---------------------------------------------------------------------------

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι",
    "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π",
    "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "φ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

OPERATORS = {
    "times": "×", "cdot": "⋅", "pm": "±", "mp": "∓", "div": "÷",
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "neq": "≠", "ne": "≠",
    "approx": "≈", "equiv": "≡", "sim": "∼", "propto": "∝",
    "rightarrow": "→", "to": "→", "leftarrow": "←", "Rightarrow": "⇒",
    "infty": "∞", "partial": "∂", "nabla": "∇", "in": "∈", "forall": "∀",
    "ldots": "…", "cdots": "⋯", "degree": "°", "circ": "∘",
    "%": "%", "_": "_", "{": "{", "}": "}", "#": "#", "&": "&", "$": "$",
}

SPACES = {",": " ", ";": " ", ":": " ", "quad": " ",
          "qquad": "  ", " ": " ", "!": ""}

NARY = {"sum": "∑", "prod": "∏", "int": "∫", "oint": "∮", "iint": "∬"}

FUNCTIONS = {"ln", "log", "exp", "max", "min", "sin", "cos", "tan", "lim",
             "det", "arg", "mod"}

ACCENTS = {"hat": "̂", "tilde": "̃", "dot": "̇",
           "vec": "⃗", "bar": "̅"}

TEXT_COMMANDS = {"text", "mathrm", "textrm", "operatorname", "mbox", "textnormal"}

FRACTIONS = {"frac", "dfrac", "tfrac"}

DELIMS = {"(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}",
          "|": "|", ".": "", "\\langle": "⟨", "\\rangle": "⟩"}

# A top-level operator ends the operand of a preceding n-ary (Σ ... + ...).
NARY_STOP = set("+-=<>,;") | {"−", "±", "≤", "≥", "≠", "≈", "→", "≡"}


# ---------------------------------------------------------------------------
# AST
# ---------------------------------------------------------------------------

@dataclass
class Run:
    text: str
    kind: str = "var"  # var | num | op | text | func


@dataclass
class Frac:
    num: list
    den: list


@dataclass
class Script:
    base: list
    sub: Optional[list] = None
    sup: Optional[list] = None


@dataclass
class Nary:
    chr: str
    sub: Optional[list]
    sup: Optional[list]
    body: list = field(default_factory=list)


@dataclass
class Rad:
    deg: Optional[list]
    body: list


@dataclass
class Delim:
    open: str
    close: str
    body: list


@dataclass
class Accent:
    chr: str
    body: list
    bar: bool = False


Node = Union[Run, Frac, Script, Nary, Rad, Delim, Accent]


# ---------------------------------------------------------------------------
# Tokenizer / parser
# ---------------------------------------------------------------------------

def _is_hangul(ch: str) -> bool:
    code = ord(ch)
    return 0xAC00 <= code <= 0xD7A3 or 0x3130 <= code <= 0x318F or 0x1100 <= code <= 0x11FF


def _tokenize(src: str) -> list[str]:
    tokens: list[str] = []
    i = 0
    while i < len(src):
        ch = src[i]
        if ch == "\\":
            j = i + 1
            if j < len(src) and src[j].isalpha():
                while j < len(src) and src[j].isalpha():
                    j += 1
                tokens.append(src[i:j])
                i = j
            elif j < len(src):
                tokens.append(src[i:j + 1])
                i = j + 1
            else:
                raise EquationError("trailing backslash")
            continue
        if ch.isspace():
            tokens.append(" ")
            while i < len(src) and src[i].isspace():
                i += 1
            continue
        if ch.isdigit() or (ch == "." and i + 1 < len(src) and src[i + 1].isdigit()):
            j = i
            while j < len(src) and (src[j].isdigit() or src[j] == "."):
                j += 1
            tokens.append(src[i:j])
            i = j
            continue
        if ch.isascii() and ch.isalpha():
            j = i
            while j < len(src) and src[j].isascii() and src[j].isalpha():
                j += 1
            tokens.append(src[i:j])
            i = j
            continue
        if _is_hangul(ch):
            j = i
            while j < len(src) and _is_hangul(src[j]):
                j += 1
            tokens.append(src[i:j])
            i = j
            continue
        tokens.append(ch)
        i += 1
    return tokens


class _Parser:
    def __init__(self, src: str):
        self.src = src
        self.tokens = _tokenize(src)
        self.pos = 0

    # -- helpers ---------------------------------------------------------
    def _peek(self, skip_space: bool = True) -> Optional[str]:
        pos = self.pos
        while skip_space and pos < len(self.tokens) and self.tokens[pos] == " ":
            pos += 1
        return self.tokens[pos] if pos < len(self.tokens) else None

    def _next(self, skip_space: bool = True) -> Optional[str]:
        while skip_space and self.pos < len(self.tokens) and self.tokens[self.pos] == " ":
            self.pos += 1
        if self.pos >= len(self.tokens):
            return None
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def _expect(self, tok: str) -> None:
        got = self._next()
        if got != tok:
            raise EquationError(f"expected {tok!r}, got {got!r} in {self.src!r}")

    def _raw_group(self) -> str:
        """Return the literal text of a {...} group (for \\text)."""
        self._expect("{")
        depth, parts = 1, []
        while True:
            if self.pos >= len(self.tokens):
                raise EquationError(f"unclosed text group in {self.src!r}")
            tok = self.tokens[self.pos]
            self.pos += 1
            if tok == "{":
                depth += 1
            elif tok == "}":
                depth -= 1
                if depth == 0:
                    return "".join(parts)
            parts.append(tok[1:] if tok in ("\\{", "\\}", "\\%", "\\_", "\\&") else tok)

    def _argument(self) -> list:
        tok = self._peek()
        if tok == "{":
            self._next()
            body = self.parse_sequence(stop={"}"})
            self._expect("}")
            return body
        if tok is None:
            raise EquationError(f"missing argument in {self.src!r}")
        atom = self._atom()
        if atom is None:
            raise EquationError(f"missing argument in {self.src!r}")
        return [atom]

    def _optional_bracket(self) -> Optional[list]:
        if self._peek() != "[":
            return None
        self._next()
        body = self.parse_sequence(stop={"]"})
        self._expect("]")
        return body

    def _scripts(self) -> tuple[Optional[list], Optional[list]]:
        sub = sup = None
        while self._peek() in ("_", "^"):
            tok = self._next()
            arg = self._argument()
            if tok == "_":
                if sub is not None:
                    raise EquationError(f"double subscript in {self.src!r}")
                sub = arg
            else:
                if sup is not None:
                    raise EquationError(f"double superscript in {self.src!r}")
                sup = arg
        return sub, sup

    # -- grammar ---------------------------------------------------------
    def parse_sequence(self, stop: set[str]) -> list:
        items: list = []
        while True:
            tok = self._peek()
            if tok is None or tok in stop:
                return items
            if tok == "\\right":
                return items
            if tok in ("_", "^"):
                if not items:
                    raise EquationError(f"script without base in {self.src!r}")
                base = items.pop()
                sub, sup = self._scripts()
                if isinstance(base, Script):
                    raise EquationError(f"double script in {self.src!r}")
                if isinstance(base, Nary):
                    raise EquationError(f"scripts after n-ary body in {self.src!r}")
                items.append(Script([base], sub, sup))
                continue
            atom = self._atom()
            if atom is None:
                continue
            if isinstance(atom, Nary) and not atom.body:
                atom.body = self._nary_body(stop)
            items.append(atom)

    def _nary_body(self, stop: set[str]) -> list:
        if self._peek() == "{":
            return self._argument()
        body: list = []
        while True:
            tok = self._peek()
            if tok is None or tok in stop or tok == "\\right":
                break
            if tok in NARY_STOP or OPERATORS.get(tok[1:] if tok.startswith("\\") else "") in NARY_STOP:
                break
            if tok in ("_", "^"):
                if not body:
                    raise EquationError(f"script without base in {self.src!r}")
                base = body.pop()
                sub, sup = self._scripts()
                body.append(Script([base], sub, sup))
                continue
            atom = self._atom()
            if atom is not None:
                if isinstance(atom, Nary) and not atom.body:
                    atom.body = self._nary_body(stop)
                body.append(atom)
        if not body:
            raise EquationError(f"n-ary operator without operand in {self.src!r}")
        return body

    def _atom(self) -> Optional[Node]:
        tok = self._next()
        if tok is None:
            return None
        if tok == "{":
            body = self.parse_sequence(stop={"}"})
            self._expect("}")
            if len(body) == 1:
                return body[0]
            return Delim("", "", body)
        if tok == "}":
            raise EquationError(f"unbalanced '}}' in {self.src!r}")
        if tok.startswith("\\"):
            return self._command(tok)
        if tok[0].isdigit() or tok[0] == ".":
            return Run(tok, "num")
        if tok.isascii() and tok.isalpha():
            return Run(tok, "var")
        if _is_hangul(tok[0]):
            return Run(tok, "text")
        if tok == "-":
            return Run("−", "op")
        if tok == "*":
            return Run("×", "op")
        if tok in "+=<>/(),.;:|![]'":
            return Run(tok, "op")
        if tok == "~":
            return Run(" ", "op")
        if tok == "&":
            raise EquationError(f"alignment '&' is not supported; split the equation in {self.src!r}")
        return Run(tok, "op" if not tok.isalnum() else "var")

    def _command(self, tok: str) -> Optional[Node]:
        name = tok[1:]
        if name in SPACES:
            return Run(SPACES[name], "op") if SPACES[name] else None
        if name == "\\":
            raise EquationError(f"line break '\\\\' is not supported; use one equation per line in {self.src!r}")
        if name in FRACTIONS:
            return Frac(self._argument(), self._argument())
        if name == "sqrt":
            deg = self._optional_bracket()
            return Rad(deg, self._argument())
        if name in NARY:
            sub, sup = self._scripts()
            return Nary(NARY[name], sub, sup)
        if name in TEXT_COMMANDS:
            text = self._raw_group()
            return Run(text, "func" if name == "operatorname" else "text")
        if name in ("mathit",):
            text = self._raw_group()
            return Run(text, "var")
        if name in ("mathbf", "boldsymbol"):
            return self._argument()[0] if self._peek() != "{" else Delim("", "", self._argument())
        if name in FUNCTIONS:
            return Run(name, "func")
        if name in ("overline",):
            return Accent("", self._argument(), bar=True)
        if name in ACCENTS:
            return Accent(ACCENTS[name], self._argument())
        if name == "left":
            open_tok = self._next()
            if open_tok not in DELIMS:
                raise EquationError(f"unsupported \\left delimiter {open_tok!r}")
            body = self.parse_sequence(stop=set())
            if self._next() != "\\right":
                raise EquationError(f"\\left without \\right in {self.src!r}")
            close_tok = self._next()
            if close_tok not in DELIMS:
                raise EquationError(f"unsupported \\right delimiter {close_tok!r}")
            return Delim(DELIMS[open_tok], DELIMS[close_tok], body)
        if name == "right":
            raise EquationError(f"\\right without \\left in {self.src!r}")
        if name in GREEK:
            return Run(GREEK[name], "var")
        if name in OPERATORS:
            return Run(OPERATORS[name], "op")
        raise EquationError(f"unsupported command {tok!r} in {self.src!r}")


def parse_latex(src: str) -> list:
    """Parse the LaTeX subset into a node list. Raises EquationError."""
    text = src.strip()
    if text.startswith("$") and text.endswith("$"):
        text = text.strip("$").strip()
    if not text:
        raise EquationError("empty equation")
    parser = _Parser(text)
    nodes = parser.parse_sequence(stop=set())
    if parser._peek() is not None:
        raise EquationError(f"unexpected {parser._peek()!r} in {src!r}")
    return _merge_runs(nodes)


def _merge_runs(nodes: list, in_script: bool = False) -> list:
    """Merge adjacent plain runs of the same kind (fewer, editable runs).

    Inside sub/superscripts a multi-letter Latin word (grid, CM, PJ) is a
    label, not a product of variables, so it is set upright.
    """
    out: list = []
    for node in nodes:
        for attr in ("num", "den", "base", "sub", "sup", "body", "deg"):
            value = getattr(node, attr, None)
            if isinstance(value, list):
                setattr(node, attr, _merge_runs(value, in_script or attr in ("sub", "sup")))
        if in_script and isinstance(node, Run) and node.kind == "var"                 and len(node.text) > 1 and node.text.isascii():
            node = Run(node.text, "label")
        if (isinstance(node, Run) and out and isinstance(out[-1], Run)
                and out[-1].kind == node.kind and node.kind in ("text", "num")):
            out[-1] = Run(out[-1].text + node.text, node.kind)
        else:
            out.append(node)
    return out


# ---------------------------------------------------------------------------
# OMML emission
# ---------------------------------------------------------------------------

RprFn = Callable[[str], str]


def pptx_rpr(size_pt: float = 12.0, color: Optional[str] = None,
             ea_font: Optional[str] = None, bold: bool = False) -> RprFn:
    """Return a run-property factory for DrawingML (PowerPoint) math runs."""
    size = int(round(size_pt * 100))
    fill = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else ""
    ea = f'<a:ea typeface="{escape(ea_font)}"/>' if ea_font else ""
    b_attr = ' b="1"' if bold else ""

    def factory(kind: str) -> str:
        italic = ' i="1"' if kind == "var" else ""
        lang = "ko-KR" if kind == "text" else "en-US"
        return (f'<a:rPr lang="{lang}" altLang="en-US" sz="{size}"{italic}{b_attr} dirty="0">'
                f'{fill}<a:latin typeface="{MATH_FONT}"/>{ea}<a:cs typeface="{MATH_FONT}"/></a:rPr>')
    return factory


def docx_rpr(size_pt: float = 11.0, color: Optional[str] = None,
             ea_font: Optional[str] = None) -> RprFn:
    """Return a run-property factory for WordprocessingML math runs."""
    half = int(round(size_pt * 2))
    ea = f' w:eastAsia="{escape(ea_font)}"' if ea_font else ""
    clr = f'<w:color w:val="{color}"/>' if color else ""

    def factory(kind: str) -> str:
        return (f'<w:rPr><w:rFonts w:ascii="{MATH_FONT}" w:hAnsi="{MATH_FONT}"{ea} '
                f'w:cs="{MATH_FONT}"/>{clr}<w:sz w:val="{half}"/><w:szCs w:val="{half}"/></w:rPr>')
    return factory


def _ctrl(rpr: RprFn) -> str:
    return f"<m:ctrlPr>{rpr('var')}</m:ctrlPr>"


def _emit_run(run: Run, rpr: RprFn) -> str:
    if run.kind in ("text", "func", "label"):
        mpr = '<m:rPr><m:sty m:val="p"/></m:rPr>'
    else:
        mpr = ""
    space = ' xml:space="preserve"' if run.text != run.text.strip() or " " in run.text else ""
    return f"<m:r>{mpr}{rpr(run.kind)}<m:t{space}>{escape(run.text)}</m:t></m:r>"


def _emit_seq(nodes: list, rpr: RprFn) -> str:
    return "".join(_emit_node(n, rpr) for n in nodes)


def _emit_node(node: Node, rpr: RprFn) -> str:
    if isinstance(node, Run):
        return _emit_run(node, rpr)
    if isinstance(node, Frac):
        return (f"<m:f><m:fPr>{_ctrl(rpr)}</m:fPr><m:num>{_emit_seq(node.num, rpr)}</m:num>"
                f"<m:den>{_emit_seq(node.den, rpr)}</m:den></m:f>")
    if isinstance(node, Script):
        base = f"<m:e>{_emit_seq(node.base, rpr)}</m:e>"
        if node.sub is not None and node.sup is not None:
            return (f"<m:sSubSup><m:sSubSupPr>{_ctrl(rpr)}</m:sSubSupPr>{base}"
                    f"<m:sub>{_emit_seq(node.sub, rpr)}</m:sub>"
                    f"<m:sup>{_emit_seq(node.sup, rpr)}</m:sup></m:sSubSup>")
        if node.sub is not None:
            return (f"<m:sSub><m:sSubPr>{_ctrl(rpr)}</m:sSubPr>{base}"
                    f"<m:sub>{_emit_seq(node.sub, rpr)}</m:sub></m:sSub>")
        return (f"<m:sSup><m:sSupPr>{_ctrl(rpr)}</m:sSupPr>{base}"
                f"<m:sup>{_emit_seq(node.sup or [], rpr)}</m:sup></m:sSup>")
    if isinstance(node, Nary):
        hide = ""
        if node.sub is None:
            hide += '<m:subHide m:val="1"/>'
        if node.sup is None:
            hide += '<m:supHide m:val="1"/>'
        lim = '<m:limLoc m:val="undOvr"/>' if node.chr != "∫" else '<m:limLoc m:val="subSup"/>'
        return (f'<m:nary><m:naryPr><m:chr m:val="{node.chr}"/>{lim}{hide}{_ctrl(rpr)}</m:naryPr>'
                f"<m:sub>{_emit_seq(node.sub or [], rpr)}</m:sub>"
                f"<m:sup>{_emit_seq(node.sup or [], rpr)}</m:sup>"
                f"<m:e>{_emit_seq(node.body, rpr)}</m:e></m:nary>")
    if isinstance(node, Rad):
        if node.deg:
            return (f"<m:rad><m:radPr>{_ctrl(rpr)}</m:radPr><m:deg>{_emit_seq(node.deg, rpr)}</m:deg>"
                    f"<m:e>{_emit_seq(node.body, rpr)}</m:e></m:rad>")
        return (f'<m:rad><m:radPr><m:degHide m:val="1"/>{_ctrl(rpr)}</m:radPr><m:deg/>'
                f"<m:e>{_emit_seq(node.body, rpr)}</m:e></m:rad>")
    if isinstance(node, Delim):
        if not node.open and not node.close:
            return _emit_seq(node.body, rpr)
        return (f'<m:d><m:dPr><m:begChr m:val="{escape(node.open)}"/>'
                f'<m:endChr m:val="{escape(node.close)}"/>{_ctrl(rpr)}</m:dPr>'
                f"<m:e>{_emit_seq(node.body, rpr)}</m:e></m:d>")
    if isinstance(node, Accent):
        if node.bar:
            return (f'<m:bar><m:barPr><m:pos m:val="top"/>{_ctrl(rpr)}</m:barPr>'
                    f"<m:e>{_emit_seq(node.body, rpr)}</m:e></m:bar>")
        return (f'<m:acc><m:accPr><m:chr m:val="{node.chr}"/>{_ctrl(rpr)}</m:accPr>'
                f"<m:e>{_emit_seq(node.body, rpr)}</m:e></m:acc>")
    raise EquationError(f"unknown node {node!r}")


def omml_math(latex: str, rpr: RprFn) -> str:
    """Return one <m:oMath> element (namespace declared inline)."""
    body = _emit_seq(parse_latex(latex), rpr)
    return f'<m:oMath xmlns:m="{M_NS}">{body}</m:oMath>'


def omml_para(latex: str, rpr: RprFn, jc: str = "centerGroup") -> str:
    """Return a display <m:oMathPara> wrapping one equation."""
    body = _emit_seq(parse_latex(latex), rpr)
    return (f'<m:oMathPara xmlns:m="{M_NS}"><m:oMathParaPr><m:jc m:val="{jc}"/></m:oMathParaPr>'
            f"<m:oMath>{body}</m:oMath></m:oMathPara>")


def pptx_math_paragraph(latex: str, rpr: RprFn, align: str = "l",
                        end_rpr: str = "") -> str:
    """Return a DrawingML <a:p> holding a native a14:m equation."""
    jc = {"l": "left", "ctr": "center", "r": "right"}.get(align, "left")
    algn = f' algn="{align}"' if align != "l" else ""
    return (f'<a:p><a:pPr{algn}/><a14:m xmlns:a14="{A14_NS}">{omml_para(latex, rpr, jc=jc)}</a14:m>'
            f"{end_rpr}</a:p>")


# ---------------------------------------------------------------------------
# Linear / preview forms
# ---------------------------------------------------------------------------

def _linear_seq(nodes: list) -> str:
    return "".join(_linear(n) for n in nodes)


def _wrap_if_complex(nodes: list) -> str:
    text = _linear_seq(nodes)
    simple = len(nodes) == 1 and isinstance(nodes[0], Run)
    return text if simple else f"({text})"


def _linear(node: Node) -> str:
    if isinstance(node, Run):
        if node.kind == "op" and node.text in "=+−×≤≥≠≈→±":
            return f" {node.text} "
        return node.text
    if isinstance(node, Frac):
        return f"{_wrap_if_complex(node.num)}/{_wrap_if_complex(node.den)}"
    if isinstance(node, Script):
        out = _linear_seq(node.base)
        if node.sub is not None:
            out += f"_{_wrap_if_complex(node.sub)}"
        if node.sup is not None:
            out += f"^{_wrap_if_complex(node.sup)}"
        return out
    if isinstance(node, Nary):
        out = node.chr
        if node.sub is not None:
            out += f"_{_wrap_if_complex(node.sub)}"
        if node.sup is not None:
            out += f"^{_wrap_if_complex(node.sup)}"
        return f"{out} {_linear_seq(node.body)}"
    if isinstance(node, Rad):
        return f"√({_linear_seq(node.body)})"
    if isinstance(node, Delim):
        return f"{node.open}{_linear_seq(node.body)}{node.close}"
    if isinstance(node, Accent):
        return _linear_seq(node.body) + (node.chr or "̅")
    raise EquationError(f"unknown node {node!r}")


def linear_text(latex: str) -> str:
    """Unicode linear form, e.g. 'ER_y = BE_y − PE_y'. Used as fallback text."""
    return " ".join(_linear_seq(parse_latex(latex)).split())


def _preview_seq(nodes: list, size: float, out: list) -> None:
    for node in nodes:
        _preview(node, size, out)


def _preview(node: Node, size: float, out: list) -> None:
    small = round(size * 0.68, 2)
    if isinstance(node, Run):
        style = ' font-style="italic"' if node.kind == "var" else ""
        out.append(f"<tspan{style}>{escape(_linear(node))}</tspan>")
        return
    if isinstance(node, Script):
        _preview_seq(node.base, size, out)
        if node.sub is not None:
            out.append(f'<tspan baseline-shift="sub" font-size="{small}">')
            _preview_seq(node.sub, small, out)
            out.append("</tspan>")
        if node.sup is not None:
            out.append(f'<tspan baseline-shift="super" font-size="{small}">')
            _preview_seq(node.sup, small, out)
            out.append("</tspan>")
        return
    if isinstance(node, Nary):
        out.append(f"<tspan>{node.chr}</tspan>")
        if node.sub is not None:
            out.append(f'<tspan baseline-shift="sub" font-size="{small}">')
            _preview_seq(node.sub, small, out)
            out.append("</tspan>")
        if node.sup is not None:
            out.append(f'<tspan baseline-shift="super" font-size="{small}">')
            _preview_seq(node.sup, small, out)
            out.append("</tspan>")
        out.append("<tspan> </tspan>")
        _preview_seq(node.body, size, out)
        return
    if isinstance(node, Frac):
        out.append("<tspan>(</tspan>")
        _preview_seq(node.num, size, out)
        out.append("<tspan>)/(</tspan>")
        _preview_seq(node.den, size, out)
        out.append("<tspan>)</tspan>")
        return
    if isinstance(node, Rad):
        out.append("<tspan>√(</tspan>")
        _preview_seq(node.body, size, out)
        out.append("<tspan>)</tspan>")
        return
    if isinstance(node, Delim):
        out.append(f"<tspan>{escape(node.open)}</tspan>")
        _preview_seq(node.body, size, out)
        out.append(f"<tspan>{escape(node.close)}</tspan>")
        return
    if isinstance(node, Accent):
        _preview_seq(node.body, size, out)
        return
    raise EquationError(f"unknown node {node!r}")


def svg_preview_tspans(latex: str, font_size: float) -> str:
    """Non-positional tspans approximating the equation for SVG previews."""
    out: list[str] = []
    _preview_seq(parse_latex(latex), font_size, out)
    return "".join(out)


def _mathml_seq(nodes: list) -> str:
    inner = "".join(_mathml(n) for n in nodes)
    return inner if len(nodes) == 1 else f"<mrow>{inner}</mrow>"


def _mathml(node: Node) -> str:
    if isinstance(node, Run):
        text = escape(node.text)
        if node.kind == "num":
            return f"<mn>{text}</mn>"
        if node.kind == "op":
            return f"<mo>{text}</mo>" if node.text.strip() else "<mspace width='0.25em'/>"
        if node.kind == "var":
            # A multi-letter <mi> renders upright by default; symbols stay italic.
            return f"<mi>{text}</mi>" if len(node.text) == 1 else f"<mi mathvariant='italic'>{text}</mi>"
        return f"<mtext>{text}</mtext>" if node.kind == "text" else f"<mi mathvariant='normal'>{text}</mi>"
    if isinstance(node, Frac):
        return f"<mfrac>{_mathml_seq(node.num)}{_mathml_seq(node.den)}</mfrac>"
    if isinstance(node, Script):
        base = _mathml_seq(node.base)
        if node.sub is not None and node.sup is not None:
            return f"<msubsup>{base}{_mathml_seq(node.sub)}{_mathml_seq(node.sup)}</msubsup>"
        if node.sub is not None:
            return f"<msub>{base}{_mathml_seq(node.sub)}</msub>"
        return f"<msup>{base}{_mathml_seq(node.sup or [])}</msup>"
    if isinstance(node, Nary):
        op = f"<mo largeop='true' movablelimits='false'>{node.chr}</mo>"
        if node.sub is not None and node.sup is not None:
            op = f"<munderover>{op}{_mathml_seq(node.sub)}{_mathml_seq(node.sup)}</munderover>"
        elif node.sub is not None:
            op = f"<munder>{op}{_mathml_seq(node.sub)}</munder>"
        elif node.sup is not None:
            op = f"<mover>{op}{_mathml_seq(node.sup)}</mover>"
        return f"<mrow>{op}{''.join(_mathml(n) for n in node.body)}</mrow>"
    if isinstance(node, Rad):
        if node.deg:
            return f"<mroot>{_mathml_seq(node.body)}{_mathml_seq(node.deg)}</mroot>"
        return f"<msqrt>{''.join(_mathml(n) for n in node.body)}</msqrt>"
    if isinstance(node, Delim):
        inner = "".join(_mathml(n) for n in node.body)
        if not node.open and not node.close:
            return f"<mrow>{inner}</mrow>"
        return (f"<mrow><mo fence='true'>{escape(node.open)}</mo>{inner}"
                f"<mo fence='true'>{escape(node.close)}</mo></mrow>")
    if isinstance(node, Accent):
        return f"<mover accent='true'>{_mathml_seq(node.body)}<mo>{node.chr or '‾'}</mo></mover>"
    raise EquationError(f"unknown node {node!r}")


def mathml(latex: str, display: bool = True) -> str:
    """Presentation MathML for browser previews (web equation editor)."""
    body = "".join(_mathml(n) for n in parse_latex(latex))
    mode = "block" if display else "inline"
    return f'<math xmlns="http://www.w3.org/1998/Math/MathML" display="{mode}">{body}</math>'


def equation_height_factor(latex: str) -> float:
    """Approximate rendered height in multiples of the font size."""
    def depth(nodes: list) -> float:
        best = 1.0
        for n in nodes:
            if isinstance(n, Frac):
                best = max(best, depth(n.num) + depth(n.den) + 0.2)
            elif isinstance(n, Nary):
                lim = 0.6 * ((n.sub is not None) + (n.sup is not None))
                best = max(best, 1.3 + lim, depth(n.body))
            elif isinstance(n, Script):
                best = max(best, depth(n.base) + 0.35)
            elif isinstance(n, (Delim, Rad, Accent)):
                best = max(best, depth(n.body) + 0.1)
        return best
    return depth(parse_latex(latex))


def estimate_width_em(latex: str) -> float:
    """Approximate rendered width in em (font-size multiples)."""
    def width(nodes: list) -> float:
        total = 0.0
        for n in nodes:
            if isinstance(n, Run):
                for ch in n.text:
                    total += 1.0 if _is_hangul(ch) else (0.9 if ch in "=+−×≤≥≠≈→±" else 0.55)
            elif isinstance(n, Frac):
                total += max(width(n.num), width(n.den)) + 0.3
            elif isinstance(n, Script):
                total += width(n.base) + 0.7 * max(width(n.sub or []), width(n.sup or []))
            elif isinstance(n, Nary):
                total += 1.2 + 0.4 * max(width(n.sub or []), width(n.sup or [])) + width(n.body)
            elif isinstance(n, Rad):
                total += width(n.body) + 0.8
            elif isinstance(n, Delim):
                total += width(n.body) + 0.7
            elif isinstance(n, Accent):
                total += width(n.body)
        return total
    return width(parse_latex(latex))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a LaTeX subset into editable Office Math (OMML).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", required=True)
    check = sub.add_parser("check", help="parse only; exit 1 on unsupported LaTeX")
    check.add_argument("latex")
    omml = sub.add_parser("omml", help="print the OMML for PowerPoint or Word")
    omml.add_argument("latex")
    omml.add_argument("--target", choices=("pptx", "docx"), default="pptx")
    omml.add_argument("--size-pt", type=float, default=12.0)
    linear = sub.add_parser("linear", help="print the Unicode linear fallback")
    linear.add_argument("latex")
    preview = sub.add_parser("preview", help="print SVG preview tspans")
    preview.add_argument("latex")
    preview.add_argument("--font-size", type=float, default=20.0)
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    from console_encoding import configure_utf8_stdio
    configure_utf8_stdio()
    args = build_parser().parse_args(argv)
    try:
        if args.command == "check":
            parse_latex(args.latex)
            print(f"OK  {linear_text(args.latex)}")
        elif args.command == "omml":
            factory = pptx_rpr(args.size_pt) if args.target == "pptx" else docx_rpr(args.size_pt)
            print(omml_para(args.latex, factory))
        elif args.command == "linear":
            print(linear_text(args.latex))
        else:
            print(svg_preview_tspans(args.latex, args.font_size))
    except EquationError as exc:
        print(f"ERROR  {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
