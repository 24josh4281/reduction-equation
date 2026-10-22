#!/usr/bin/env python3
"""
PPT Master - Rough-Text Equation Shorthand

Lets a consultant type reduction-project formulas the way they would in a
memo and turns them into the LaTeX subset of `omml_math.py`, which then
becomes native, editable Office Math (OMML) in Word and PowerPoint.

Shorthand (all of these are accepted and may be mixed):
    BE_MFG,y                 subscript runs to the next space/operator
    Q_{PJ,j,y}               braces for a subscript holding spaces or symbols
    BEy  ERy  n0             short forms: BE_y, ER_y, n_0
    a * b   a × b            multiplication sign (×);  a · b  a ∙ b -> ·
    A/B   (a+b)/2            fraction; one outer pair of parentheses is dropped
    x^2   (…)^x_a  e^(a+b)   superscript
    sum_j(…)  Σ_j …  sum_{i=1}^{n} …   n-ary sum (prod for ∏)
    sqrt(x)  min(a, b)  max(…)  ln(…)  exp(…)  bar(x)  hat(x)
    sigma  σ  Delta  Δ       Greek letters by name or character
    <=  >=  !=  ≈            relations
    (식 3) or [식 3] at the end -> equation label
    Korean words stay upright text.
A line that already looks like LaTeX (contains a backslash command) is used
as LaTeX unchanged.

Usage:
    python3 scripts/math_shorthand.py latex "<shorthand>"
    python3 scripts/math_shorthand.py check "<shorthand>"
    python3 scripts/math_shorthand.py text "<latex>"        (LaTeX -> shorthand)

Examples:
    python3 scripts/math_shorthand.py latex "BE_MFG,y = sum_j((Q_PJ,j,y * QW_BL,j) * DHC_PJ,j/DHC_BL,j * EF_BL,MFG,j)"
    python3 scripts/math_shorthand.py latex "n = n0/(1 + (n0 - 1)/N)"

Dependencies:
    None (only uses standard library)
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from omml_math import (  # noqa: E402
    GREEK,
    Accent,
    Delim,
    EquationError,
    Frac,
    Nary,
    Rad,
    Run,
    Script,
    _is_hangul,
    parse_latex,
)

GREEK_CHARS = {v: k for k, v in GREEK.items() if not k.startswith("var")}
FUNCS = {"min", "max", "ln", "log", "exp", "sin", "cos", "tan", "lim"}
UNARY_FUNCS = {"sqrt", "bar", "hat", "tilde", "overline"}
NARY_WORDS = {"sum": "sum", "Σ": "sum", "∑": "sum", "prod": "prod", "∏": "prod", "Π": "prod"}
RELATIONS = {"=": "=", "<=": "\\le", ">=": "\\ge", "!=": "\\neq", "≤": "\\le", "≥": "\\ge",
             "≠": "\\neq", "≈": "\\approx", "<": "<", ">": ">", "→": "\\rightarrow"}
MUL_OPS = {"*": "\\times", "×": "\\times", "·": "\\cdot", "∙": "\\cdot", "⋅": "\\cdot"}
ADD_OPS = {"+": "+", "-": "-", "−": "-", "±": "\\pm"}
LABEL = re.compile(r"\s*[(\[（]\s*((?:식|수식|Eq\.?)\s*[0-9]+(?:[-.][0-9]+)*)\s*[)\]）]\s*$")
LATEX_HINT = re.compile(r"\\(?:frac|sum|times|cdot|left|text|mathrm|sqrt|prod|sigma|[a-zA-Z]{2,})")
SUB_CHARS = re.compile(r"[A-Za-z0-9,.'·]")
SHORT_SUB = re.compile(r"^(?:([A-Z]{2,})([a-z])|([a-z])([0-9]))$")


# ---------------------------------------------------------------------------
# Tokens
# ---------------------------------------------------------------------------

@dataclass
class Tok:
    kind: str           # ident num op open close text nary func comma
    text: str
    sub: Optional[str] = None   # LaTeX of an identifier subscript


def _normalize(src: str) -> str:
    # Mathematical italic letters pasted from Word/PowerPoint (𝐵𝐸, ℎ) -> ASCII.
    out = []
    for ch in src:
        if 0x1D400 <= ord(ch) <= 0x1D7FF or ch == "ℎ":
            out.append(unicodedata.normalize("NFKC", ch))
        else:
            out.append(ch)
    return "".join(out).replace(" ", " ")


def _braced(src: str, i: int) -> tuple[str, int]:
    """Return the content of the {...} group starting at src[i] and the index after it."""
    depth, j = 0, i
    while j < len(src):
        if src[j] == "{":
            depth += 1
        elif src[j] == "}":
            depth -= 1
            if depth == 0:
                return src[i + 1:j], j + 1
        j += 1
    raise EquationError(f"unclosed '{{' in {src!r}")


def _sub_latex(raw: str) -> str:
    raw = raw.strip()
    if not raw:
        raise EquationError("empty subscript")
    if "_" in raw or "^" in raw or "(" in raw or " " in raw:
        return to_latex_expr(raw)
    return "".join(f"\\{GREEK_CHARS[c]} " if c in GREEK_CHARS else c for c in raw).strip()


def _read_sub(src: str, j: int) -> tuple[Optional[str], int]:
    """Read the subscript after src[j] == "_"; return (LaTeX or None, index after it)."""
    n = len(src)
    if j + 1 < n and src[j + 1] == "{":
        raw, k = _braced(src, j + 1)
        return _sub_latex(raw), k
    k = j + 1
    while k < n and (SUB_CHARS.match(src[k]) or src[k] in GREEK_CHARS):
        k += 1
    raw = src[j + 1:k]
    # "min(A_x, B)": a trailing comma/period belongs to the sentence.
    while raw and raw[-1] in ",.":
        raw, k = raw[:-1], k - 1
    return (_sub_latex(raw), k) if raw else (None, j + 1)


def _tokenize(src: str) -> list[Tok]:
    toks: list[Tok] = []
    i, n = 0, len(src)
    while i < n:
        ch = src[i]
        if ch.isspace():
            i += 1
            continue
        two = src[i:i + 2]
        if two in ("<=", ">=", "!="):
            toks.append(Tok("rel", two))
            i += 2
            continue
        if ch in RELATIONS:
            toks.append(Tok("rel", ch))
            i += 1
            continue
        if ch in MUL_OPS:
            toks.append(Tok("mul", ch))
            i += 1
            continue
        if ch in ADD_OPS:
            toks.append(Tok("add", ch))
            i += 1
            continue
        if ch in "/÷":
            # " / " with spaces divides the whole product on its left; "a/b" only the atoms.
            spaced = (i > 0 and src[i - 1].isspace()) or (i + 1 < n and src[i + 1].isspace())
            toks.append(Tok("div", "/", "wide" if spaced else None))
            i += 1
            continue
        if ch == "^":
            toks.append(Tok("pow", "^"))
            i += 1
            continue
        if ch in "([{":
            toks.append(Tok("open", ch))
            i += 1
            continue
        if ch in ")]}":
            toks.append(Tok("close", ch))
            i += 1
            continue
        if ch == ",":
            toks.append(Tok("comma", ","))
            i += 1
            continue
        if ch.isdigit() or (ch == "." and i + 1 < n and src[i + 1].isdigit()):
            j = i
            while j < n and (src[j].isdigit() or (src[j] == "." and j + 1 < n and src[j + 1].isdigit())):
                j += 1
            toks.append(Tok("num", src[i:j]))
            i = j
            continue
        if ch in NARY_WORDS and not ch.isascii():
            toks.append(Tok("nary", NARY_WORDS[ch]))
            i += 1
            continue
        if _is_hangul(ch):
            j = i
            while j < n and (_is_hangul(src[j]) or (src[j] == " " and j + 1 < n and _is_hangul(src[j + 1]))):
                j += 1
            toks.append(Tok("text", src[i:j]))
            i = j
            continue
        if ch.isalpha() or ch in GREEK_CHARS:
            j = i
            if ch.isascii():
                while j < n and src[j].isascii() and src[j].isalnum():
                    j += 1
            else:
                j += 1
            word = src[i:j]
            sub = None
            if j < n and src[j] == "_" and word not in ("sum", "prod"):
                sub, j = _read_sub(src, j)
                if sub is None:
                    raise EquationError(f"empty subscript after {word!r}")
            if word in ("sum", "prod"):
                toks.append(Tok("nary", word))
            elif word in FUNCS or word in UNARY_FUNCS:
                toks.append(Tok("func", word))
            else:
                if sub is None:
                    short = SHORT_SUB.match(word)
                    if short:
                        word, sub = (short.group(1), short.group(2)) if short.group(1) else (short.group(3), short.group(4))
                toks.append(Tok("ident", word, sub))
            i = j
            continue
        if ch == "_":
            # After sum/prod or a closing bracket: "sum_j", "bar(DP)_BL,a".
            sub, i = _read_sub(src, i)
            toks.append(Tok("under", "_", sub))
            continue
        if ch == "%":
            toks.append(Tok("ident", "%"))
            i += 1
            continue
        raise EquationError(f"unsupported character {ch!r} in {src!r}")
    return toks


# ---------------------------------------------------------------------------
# Expression tree
# ---------------------------------------------------------------------------

@dataclass
class Node:
    kind: str                     # atom group frac pow nary func seq
    latex: str = ""
    items: list = field(default_factory=list)
    open: str = ""
    close: str = ""


def _strip_group(node: Node) -> Node:
    if node.kind == "group" and node.open in "({":
        return node.items[0]
    return node


class _Parser:
    def __init__(self, toks: list[Tok], src: str):
        self.toks, self.pos, self.src = toks, 0, src

    def peek(self, offset: int = 0) -> Optional[Tok]:
        idx = self.pos + offset
        return self.toks[idx] if idx < len(self.toks) else None

    def _is(self, *kinds: str) -> bool:
        tok = self.peek()
        return tok is not None and tok.kind in kinds

    def next(self) -> Tok:
        tok = self.peek()
        if tok is None:
            raise EquationError(f"unexpected end of {self.src!r}")
        self.pos += 1
        return tok

    def expect_close(self, opener: str) -> str:
        tok = self.next()
        want = {"(": ")", "[": "]", "{": "}"}[opener]
        if tok.kind != "close" or tok.text != want:
            raise EquationError(f"expected {want!r} in {self.src!r}")
        return tok.text

    def parse(self) -> Node:
        node = self.relation()
        extra = self.peek()
        if extra is not None:
            raise EquationError(f"unexpected {extra.text!r} in {self.src!r}")
        return node

    def relation(self) -> Node:
        parts = [self.additive()]
        while self._is("rel", "comma"):
            tok = self.next()
            parts.append(Node("atom", RELATIONS.get(tok.text, ",")))
            parts.append(self.additive())
        return parts[0] if len(parts) == 1 else Node("seq", items=parts)

    def additive(self) -> Node:
        parts: list[Node] = []
        if self._is("add"):
            parts.append(Node("atom", ADD_OPS[self.next().text]))
        parts.append(self.multiplicative())
        while self._is("add"):
            parts.append(Node("atom", ADD_OPS[self.next().text]))
            parts.append(self.multiplicative())
        return parts[0] if len(parts) == 1 else Node("seq", items=parts)

    def _starts_operand(self) -> bool:
        tok = self.peek()
        return tok is not None and (tok.kind in ("ident", "num", "nary", "func", "text")
                                    or (tok.kind == "open"))

    def multiplicative(self) -> Node:
        parts = [self.division()]
        while True:
            tok = self.peek()
            if tok and tok.kind == "div" and tok.sub == "wide":
                self.next()
                num = parts[0] if len(parts) == 1 else Node("seq", items=parts)
                parts = [Node("frac", items=[_strip_group(num), _strip_group(self.division())])]
            elif tok and tok.kind == "mul":
                self.next()
                parts.append(Node("atom", MUL_OPS[tok.text]))
                parts.append(self.division())
            elif self._starts_operand():
                parts.append(self.division())      # implicit product: "2 x"
            else:
                break
        return parts[0] if len(parts) == 1 else Node("seq", items=parts)

    def division(self) -> Node:
        node = self.power()
        while self._is("div") and self.peek().sub != "wide":
            self.next()
            node = Node("frac", items=[_strip_group(node), _strip_group(self.power())])
        return node

    def power(self) -> Node:
        base = self.primary()
        if self._is("under"):
            sub = self.next().sub
            if sub is None:
                raise EquationError(f"empty subscript in {self.src!r}")
            base = Node("atom", f"{{{_emit(base)}}}_{{{sub}}}" if base.kind != "func"
                        else f"{_emit(base)}_{{{sub}}}")
        if self._is("pow"):
            self.next()
            if self._is("add"):     # x^-1
                sign = ADD_OPS[self.next().text]
                exp = Node("seq", items=[Node("atom", sign), self.primary()])
            else:
                exp = _strip_group(self.primary())
            return Node("pow", items=[base, exp])
        return base

    def primary(self) -> Node:
        tok = self.next()
        if tok.kind == "num":
            return Node("atom", tok.text)
        if tok.kind == "text":
            return Node("atom", f"\\text{{{tok.text}}}")
        if tok.kind == "ident":
            base = _ident_latex(tok.text)
            return Node("atom", f"{base}_{{{tok.sub}}}" if tok.sub is not None else base)
        if tok.kind == "open":
            inner = self.relation()
            close = self.expect_close(tok.text)
            return Node("group", items=[inner], open=tok.text, close=close)
        if tok.kind == "func":
            arg = self.primary()
            return Node("func", latex=tok.text, items=[arg])
        if tok.kind == "nary":
            sub = sup = None
            if self._is("under"):
                tok_sub = self.next().sub
                sub = Node("atom", tok_sub) if tok_sub is not None else self._limit()
            if self._is("pow"):
                self.next()
                sup = self._limit()
            body = self.multiplicative()
            return Node("nary", latex=tok.text, items=[sub, sup, body])
        raise EquationError(f"unexpected {tok.text!r} in {self.src!r}")

    def _limit(self) -> Node:
        tok = self.peek()
        if tok and tok.kind == "open":
            return _strip_group(self.primary())
        tok = self.next()
        if tok.kind not in ("ident", "num"):
            raise EquationError(f"bad sum limit {tok.text!r} in {self.src!r}")
        base = _ident_latex(tok.text) if tok.kind == "ident" else tok.text
        return Node("atom", f"{base}_{{{tok.sub}}}" if tok.sub else base)


def _ident_latex(word: str) -> str:
    if word in GREEK:
        return f"\\{word}"
    if word in GREEK_CHARS:
        return f"\\{GREEK_CHARS[word]}"
    if word == "%":
        return "\\%"
    return word


DELIM_LATEX = {"(": ("\\left(", "\\right)"), "[": ("\\left[", "\\right]"),
               "{": ("\\left\\{", "\\right\\}")}


def _emit(node: Node) -> str:
    if node.kind == "atom":
        return node.latex
    if node.kind == "seq":
        out = []
        for item in node.items:
            text = _emit(item)
            if item.kind == "atom" and (text.startswith("\\") and text in
                                        ("\\times", "\\cdot", "\\le", "\\ge", "\\neq", "\\approx", "\\pm",
                                         "\\rightarrow")):
                out.append(f" {text} ")
            elif item.kind == "atom" and text in ("=", "+", "-", "<", ">"):
                out.append(f" {text} ")
            elif text == ",":
                out.append(", ")
            else:
                if out and not out[-1].endswith(" "):
                    out.append(" ")
                out.append(text)
        return "".join(out).strip()
    if node.kind == "group":
        left, right = DELIM_LATEX[node.open]
        return f"{left}{_emit(node.items[0])}{right}"
    if node.kind == "frac":
        return f"\\frac{{{_emit(node.items[0])}}}{{{_emit(node.items[1])}}}"
    if node.kind == "pow":
        return f"{_emit(node.items[0])}^{{{_emit(node.items[1])}}}"
    if node.kind == "func":
        arg = node.items[0]
        if node.latex == "sqrt":
            return f"\\sqrt{{{_emit(_strip_group(arg))}}}"
        if node.latex in ("bar", "hat", "tilde", "overline"):
            cmd = "overline" if node.latex == "bar" else node.latex
            return f"\\{cmd}{{{_emit(_strip_group(arg))}}}"
        return f"\\{node.latex}{_emit(arg)}"
    if node.kind == "nary":
        sub, sup, body = node.items
        out = f"\\{node.latex}"
        if sub is not None:
            out += f"_{{{_emit(sub)}}}"
        if sup is not None:
            out += f"^{{{_emit(sup)}}}"
        return f"{out} {_emit(body)}"
    raise EquationError(f"unknown node {node.kind}")


def split_label(text: str) -> tuple[str, Optional[str]]:
    """Strip a trailing "(식 3)" label; return (formula, label)."""
    match = LABEL.search(text)
    if not match:
        return text.strip(), None
    return text[:match.start()].strip(), re.sub(r"\s+", " ", match.group(1)).strip()


def to_latex_expr(text: str) -> str:
    toks = _tokenize(_normalize(text))
    if not toks:
        raise EquationError("empty equation")
    return _emit(_Parser(toks, text).parse())


def to_latex(text: str) -> str:
    """Convert shorthand (or pass through LaTeX) and verify it parses for OMML."""
    body = _normalize(text).strip()
    if body.startswith("$") and body.endswith("$"):
        body = body.strip("$").strip()
    latex = body if LATEX_HINT.search(body) else to_latex_expr(body)
    parse_latex(latex)
    return latex


# ---------------------------------------------------------------------------
# LaTeX tree -> shorthand (round trip for imported Word/PPT equations)
# ---------------------------------------------------------------------------

def _plain(nodes: list) -> Optional[str]:
    """Subscript text when it can stay unbraced (PJ,j,y)."""
    out = []
    for n in nodes:
        if not isinstance(n, Run) or n.kind == "text":
            return None
        if n.kind == "op" and n.text not in ",.'":
            return None
        out.append(n.text)
    text = "".join(out)
    return text if text and all(SUB_CHARS.match(c) or c in GREEK_CHARS for c in text) \
        and text[-1] not in ",." else None


def _sh_seq(nodes: list) -> str:
    out = []
    for n in nodes:
        piece = _sh(n)
        if isinstance(n, Run) and n.kind == "op" and n.text in "=+−×⋅≤≥≠≈±<>→":
            out.append(f" {piece} ")
        elif isinstance(n, Run) and n.kind == "op" and n.text == ",":
            out.append(", ")
        else:
            if out and not out[-1].endswith((" ", "(", "[", "{")) and not piece.startswith((")", "]", "}")):
                prev_word = out[-1][-1:].isalnum()
                if prev_word and (piece[:1].isalnum() or piece.startswith(("sum", "prod", "Σ"))):
                    out.append(" ")
            out.append(piece)
    return re.sub(r"\s+", " ", "".join(out)).strip()


def _atomic(nodes: list) -> bool:
    if len(nodes) != 1:
        return False
    n = nodes[0]
    return isinstance(n, (Script, Rad)) or (isinstance(n, Run) and n.kind != "op") \
        or (isinstance(n, Delim) and n.open == "(")


def _operand(nodes: list) -> str:
    text = _sh_seq(nodes)
    return text if _atomic(nodes) else f"({text})"


def _sh(node) -> str:
    if isinstance(node, Run):
        if node.kind == "op":
            return {"−": "-", "⋅": "·"}.get(node.text, node.text)
        return node.text
    if isinstance(node, Frac):
        return f"{_operand(node.num)}/{_operand(node.den)}"
    if isinstance(node, Script):
        out = _operand(node.base) if not (len(node.base) == 1 and isinstance(node.base[0], Run)) \
            else node.base[0].text
        if node.sub is not None:
            plain = _plain(node.sub)
            out += f"_{plain}" if plain else f"_{{{_sh_seq(node.sub)}}}"
        if node.sup is not None:
            out += f"^{_operand(node.sup)}"
        return out
    if isinstance(node, Nary):
        word = "sum" if node.chr == "∑" else "prod" if node.chr == "∏" else None
        if word is None:
            raise EquationError(f"no shorthand for {node.chr}")
        out = word
        if node.sub is not None:
            plain = _plain(node.sub)
            out += f"_{plain}" if plain else f"_({_sh_seq(node.sub)})"
        if node.sup is not None:
            out += f"^{_operand(node.sup)}"
        return f"{out} {_sh_seq(node.body)}"
    if isinstance(node, Rad):
        return f"sqrt({_sh_seq(node.body)})"
    if isinstance(node, Delim):
        return f"{node.open}{_sh_seq(node.body)}{node.close}"
    if isinstance(node, Accent):
        return f"{'bar' if node.bar else 'hat'}({_sh_seq(node.body)})"
    raise EquationError(f"unknown node {node!r}")


def to_shorthand(latex: str) -> str:
    """Readable shorthand for a LaTeX formula; falls back to LaTeX when it cannot round-trip."""
    try:
        text = _sh_seq(parse_latex(latex))
        if same_structure(to_latex(text), latex):
            return text
    except EquationError:
        pass
    return latex


def same_structure(latex_a: str, latex_b: str) -> bool:
    """True when both render the same equation (auto-sized vs literal brackets count as equal)."""
    from omml_math import omml_math

    def norm(latex: str) -> str:
        xml = omml_math(re.sub(r"\\(?:left|right)(?![a-zA-Z])", "", latex), lambda kind: "")
        return re.sub(r"<m:ctrlPr>.*?</m:ctrlPr>", "", xml)
    try:
        return norm(latex_a) == norm(latex_b)
    except EquationError:
        return False


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Rough-text equation shorthand -> LaTeX subset for OMML.")
    sub = parser.add_subparsers(dest="command", required=True)
    for name, help_text in (("latex", "print LaTeX"), ("check", "print LaTeX and linear form; exit 1 on error"),
                            ("text", "LaTeX -> shorthand")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("source")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    from console_encoding import configure_utf8_stdio
    from omml_math import linear_text
    configure_utf8_stdio()
    args = build_parser().parse_args(argv)
    try:
        if args.command == "text":
            print(to_shorthand(args.source))
            return 0
        formula, label = split_label(args.source)
        latex = to_latex(formula)
        print(latex)
        if args.command == "check":
            print(f"OK  {linear_text(latex)}" + (f"  ({label})" if label else ""))
    except EquationError as exc:
        print(f"ERROR  {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
