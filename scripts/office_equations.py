#!/usr/bin/env python3
"""
PPT Master - Reduction-Project Equations for Word and PowerPoint

Builds the editable-math deliverables of a GHG reduction project (감축사업):
scenario formulas (baseline / project / leakage / reduction) and the
monitoring-parameter tables. Every formula and every parameter symbol is
written as native Office Math (OMML), so it opens in Word and PowerPoint as
an equation object that can be edited with Insert > Equation, never as an
image or plain text.

Subcommands:
    example                  print the JSON spec schema with sample values
    docx <spec.json> -o X    write a Word document (scenarios + where-lists +
                             monitoring / fixed parameter tables)
    inject-pptx <in.pptx> <spec.json> -o X
                             replace text runs holding {{eq:<id>}} or
                             {{sym:<parameter id>}} tokens with OMML (for
                             direct-PPTX routes such as ppt-template-fill)
    inject-docx <in.docx> [spec.json] -o X [--convert-plain] [--no-number]
                             turn markers in a drafted Word methodology or
                             business plan into OMML: {{eq:id}}, {{sym:id}},
                             {{where:scenario}}, {{monitoring}}, {{fixed}},
                             $$LaTeX$$, $LaTeX$; --convert-plain also converts
                             formulas typed as text with Word subscripts
    verify <file.docx|file.pptx> [--expect N]
                             count OMML equations and flag formula-like plain
                             text (e.g. "ER_y = BE_y - PE_y") that escaped
                             conversion; exit 1 on failure

Usage:
    python3 scripts/office_equations.py example > equations.json
    python3 scripts/office_equations.py docx equations.json -o exports/monitoring.docx
    python3 scripts/office_equations.py inject-pptx deck.pptx equations.json -o deck_eq.pptx
    python3 scripts/office_equations.py inject-docx 사업계획서.docx equations.json -o 사업계획서_수식.docx
    python3 scripts/office_equations.py inject-docx 방법론_초안.docx -o 방법론_수식.docx --convert-plain
    python3 scripts/office_equations.py verify exports/monitoring.docx

Examples:
    python3 scripts/office_equations.py docx projects/demo/notes/equations.json -o projects/demo/exports/모니터링계획.docx

Dependencies:
    None (only uses standard library)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any, Optional
from xml.sax.saxutils import escape

_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from console_encoding import configure_utf8_stdio  # noqa: E402
from omml_math import (  # noqa: E402
    FORMULA_LIKE,
    A14_NS,
    M_NS,
    EquationError,
    docx_rpr,
    omml_math,
    omml_para,
    parse_latex,
    pptx_rpr,
)

configure_utf8_stdio()

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DEFAULT_FONT = "Pretendard"
HEADER_FILL = "1F3A5F"
HEADER_TEXT = "FFFFFF"
BAND_FILL = "F3F5F8"
BORDER = "BFC6D1"

EXAMPLE_SPEC: dict[str, Any] = {
    "title": "감축사업 모니터링 계획 — 산정식 및 모니터링 인자",
    "font": DEFAULT_FONT,
    "font_size_pt": 10.5,
    "scenarios": [
        {
            "id": "baseline",
            "name": "베이스라인 시나리오",
            "description": "사업이 없을 때 계통전력으로 같은 양을 공급한 경우의 배출량",
            "equations": [
                {"id": "be", "label": "식 1", "latex": "BE_y = EG_{PJ,y} \\times EF_{grid,CM,y}"},
            ],
            "where": [
                {"symbol": "BE_y", "desc": "y년 베이스라인 배출량", "unit": "tCO2eq/yr"},
                {"symbol": "EG_{PJ,y}", "desc": "y년 사업 순발전량", "unit": "MWh/yr"},
                {"symbol": "EF_{grid,CM,y}", "desc": "y년 계통 복합배출계수", "unit": "tCO2/MWh"},
            ],
        },
        {
            "id": "project",
            "name": "사업 시나리오",
            "description": "사업 설비의 보조연료 연소 배출량",
            "equations": [
                {"id": "pe", "label": "식 2", "latex": "PE_y = \\sum_{i} FC_{i,y} \\times NCV_{i} \\times EF_{CO_2,i}"},
            ],
            "where": [
                {"symbol": "PE_y", "desc": "y년 사업 배출량", "unit": "tCO2eq/yr"},
                {"symbol": "FC_{i,y}", "desc": "연료 i의 y년 사용량", "unit": "kL/yr"},
            ],
        },
        {
            "id": "reduction",
            "name": "감축량",
            "description": "베이스라인 배출량에서 사업·누출 배출량을 뺀 값",
            "equations": [
                {"id": "er", "label": "식 3", "latex": "ER_y = BE_y - PE_y - LE_y"},
            ],
            "where": [
                {"symbol": "ER_y", "desc": "y년 온실가스 감축량", "unit": "tCO2eq/yr"},
                {"symbol": "LE_y", "desc": "y년 누출 배출량(해당 시)", "unit": "tCO2eq/yr"},
            ],
        },
    ],
    "monitoring_parameters": [
        {
            "id": "eg",
            "symbol": "EG_{PJ,y}",
            "name": "사업 순발전량",
            "unit": "MWh/yr",
            "source": "전력량계(거래용) 검침 기록",
            "method": "연속 계측, 월 단위 집계",
            "frequency": "연속 / 월 집계",
            "qaqc": "한전 거래 명세서와 월별 교차 확인, 계량기 검교정 주기 준수",
        },
        {
            "id": "fc",
            "symbol": "FC_{i,y}",
            "name": "보조연료 사용량",
            "unit": "kL/yr",
            "source": "연료 구매 영수증, 유량계",
            "method": "구매량 기준, 재고 변동 보정",
            "frequency": "월",
            "qaqc": "구매 기록과 유량계 값 차이 5% 초과 시 원인 기록",
        },
    ],
    "fixed_parameters": [
        {
            "id": "efgrid",
            "symbol": "EF_{grid,CM,y}",
            "name": "계통 복합배출계수",
            "value": "(공표값 입력)",
            "unit": "tCO2/MWh",
            "source": "국가 온실가스 종합정보센터 공표값",
        },
    ],
}


def example_spec(kind: Optional[str] = None) -> dict[str, Any]:
    """EXAMPLE_SPEC plus, for a document kind, the blocks that kind reads."""
    spec = json.loads(json.dumps(EXAMPLE_SPEC))
    if kind not in DOC_KINDS:
        return spec
    layout = DOC_KINDS[kind]
    spec["title"] = layout["title"]
    texts = {heading: "" for heading, content, *_ in layout["sections"] if content == "text"}
    texts["모니터링 절차 및 QA/QC"] = ""
    spec["document"] = {"kind": kind, "meta": {key: "" for key in layout["meta"]}, "texts": texts}
    if kind == "business_plan":
        spec["expected_reductions"] = [{"year": "", "be": "", "pe": "", "le": "", "er": ""}]
        spec["expected_total"] = {"be": "", "pe": "", "le": "", "er": ""}
    return spec


# ---------------------------------------------------------------------------
# Spec validation
# ---------------------------------------------------------------------------

def validate_spec(spec: dict[str, Any]) -> list[str]:
    """Return error strings; empty when every formula and symbol parses."""
    errors: list[str] = []
    seen: set[str] = set()

    def check(latex: Any, where: str) -> None:
        if not isinstance(latex, str) or not latex.strip():
            errors.append(f"{where}: empty LaTeX")
            return
        try:
            parse_latex(latex)
        except EquationError as exc:
            errors.append(f"{where}: {exc}")

    for s_idx, scenario in enumerate(spec.get("scenarios", [])):
        for e_idx, eq in enumerate(scenario.get("equations", [])):
            eq_id = eq.get("id") or f"{scenario.get('id', s_idx)}-{e_idx}"
            if eq_id in seen:
                errors.append(f"duplicate equation id {eq_id!r}")
            seen.add(eq_id)
            check(eq.get("latex"), f"scenario {scenario.get('id', s_idx)} equation {eq_id}")
        for w_idx, item in enumerate(scenario.get("where", [])):
            check(item.get("symbol"), f"scenario {scenario.get('id', s_idx)} where[{w_idx}]")
    for key in ("monitoring_parameters", "fixed_parameters"):
        for p_idx, param in enumerate(spec.get(key, [])):
            check(param.get("symbol"), f"{key}[{p_idx}] {param.get('id', '')}")
    if not spec.get("scenarios") and not spec.get("monitoring_parameters"):
        errors.append("spec has neither scenarios nor monitoring_parameters")
    return errors


# ---------------------------------------------------------------------------
# Word (docx)
# ---------------------------------------------------------------------------

def _w_run(text: str, font: str, size_pt: float, bold: bool = False,
           color: Optional[str] = None) -> str:
    half = int(round(size_pt * 2))
    b = "<w:b/><w:bCs/>" if bold else ""
    c = f'<w:color w:val="{color}"/>' if color else ""
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
            f'{b}{c}<w:sz w:val="{half}"/><w:szCs w:val="{half}"/></w:rPr>'
            f'<w:t{space}>{escape(text)}</w:t></w:r>')


def _w_para(content: str, align: str = "left", before: int = 0, after: int = 80,
            style: Optional[str] = None, keep_next: bool = False) -> str:
    st = f'<w:pStyle w:val="{style}"/>' if style else ""
    kn = "<w:keepNext/>" if keep_next else ""
    return (f'<w:p><w:pPr>{st}{kn}<w:spacing w:before="{before}" w:after="{after}"/>'
            f'<w:jc w:val="{align}"/></w:pPr>{content}</w:p>')


def _tc(content: str, width: int, fill: Optional[str] = None, align: str = "left",
        valign: str = "center", borders: bool = True) -> str:
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    if borders:
        b = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="{BORDER}"/>'
                    for side in ("top", "left", "bottom", "right"))
    else:
        b = "".join(f'<w:{side} w:val="nil"/>' for side in ("top", "left", "bottom", "right"))
    paragraphs = content if content.startswith("<w:p>") or content.startswith("<w:p ") else \
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="{align}"/></w:pPr>{content}</w:p>'
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:tcBorders>{b}</w:tcBorders>{shd}'
            f'<w:vAlign w:val="{valign}"/></w:tcPr>{paragraphs}</w:tc>')


def _table(rows: list[str], widths: list[int], borders: bool = True) -> str:
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl_b = "" if borders else (
        "<w:tblBorders>" + "".join(f'<w:{s} w:val="nil"/>' for s in
                                   ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>")
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:jc w:val="center"/>{tbl_b}'
            f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/>'
            f'<w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid>{"".join(rows)}</w:tbl>')


def _tr(cells: list[str], header: bool = False) -> str:
    hdr = "<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>" if header else "<w:trPr><w:cantSplit/></w:trPr>"
    body = "".join(cells)
    if header:  # keep the header row on the same page as the first data row
        body = body.replace("<w:pPr><w:spacing", "<w:pPr><w:keepNext/><w:spacing")
    return f"<w:tr>{hdr}{body}</w:tr>"


def _param_table(params: list[dict[str, Any]], columns: list[tuple[str, str, int]],
                 font: str, size_pt: float) -> str:
    rpr = docx_rpr(size_pt, ea_font=font)
    header = _tr([_tc(_w_run(label, font, size_pt, bold=True, color=HEADER_TEXT), w,
                      fill=HEADER_FILL, align="center") for _, label, w in columns], header=True)
    rows = [header]
    for idx, param in enumerate(params):
        fill = BAND_FILL if idx % 2 else None
        cells = []
        for key, _, w in columns:
            if key == "symbol":
                cells.append(_tc(omml_math(param.get("symbol", ""), rpr), w, fill=fill, align="center"))
            else:
                cells.append(_tc(_w_run(str(param.get(key, "")), font, size_pt), w, fill=fill,
                                 align="center" if key in ("unit", "frequency", "value") else "left"))
        rows.append(_tr(cells))
    return _table(rows, [w for _, _, w in columns])


def _scaled(widths: list[int], total: Optional[int]) -> list[int]:
    if not total:
        return widths
    factor = total / sum(widths)
    return [max(200, int(w * factor)) for w in widths]


def where_table(where: list[dict[str, Any]], font: str, size: float,
                total: Optional[int] = None) -> str:
    """Borderless "symbol : description [unit]" list under an equation."""
    rpr_sym = docx_rpr(size, ea_font=font)
    widths = _scaled([1900, 300, 5000, 1800], total)
    rows = []
    for item in where:
        rows.append(_tr([
            _tc(omml_math(item["symbol"], rpr_sym), widths[0], align="left", borders=False),
            _tc(_w_run(":", font, size), widths[1], align="center", borders=False),
            _tc(_w_run(item.get("desc", ""), font, size), widths[2], borders=False),
            _tc(_w_run(f"[{item['unit']}]" if item.get("unit") else "", font, size),
                widths[3], align="left", borders=False),
        ]))
    return _table(rows, widths, borders=False)


MONITORING_COLUMNS = [("symbol", "인자", 1400), ("name", "설명", 1600), ("unit", "단위", 900),
                      ("source", "자료 출처", 1500), ("method", "측정 방법", 1500),
                      ("frequency", "주기", 900), ("qaqc", "QA/QC", 1700)]
FIXED_COLUMNS = [("symbol", "인자", 1600), ("name", "설명", 2400), ("value", "값", 1400),
                 ("unit", "단위", 1200), ("source", "출처", 2900)]


def param_table(params: list[dict[str, Any]], fixed: bool, font: str, size: float,
                total: Optional[int] = None) -> str:
    columns = FIXED_COLUMNS if fixed else MONITORING_COLUMNS
    widths = _scaled([w for _, _, w in columns], total)
    return _param_table(params, [(k, label, w) for (k, label, _), w in zip(columns, widths)],
                        font, size)


# Section skeletons for whole documents. "text" sections take the user's text
# from spec["document"]["texts"][heading] or show a grey placeholder to fill in.
# They follow the usual order of Korean offset methodologies and project plans;
# match them to the scheme's official form before submission.
DOC_KINDS: dict[str, dict[str, Any]] = {
    "methodology": {
        "title": "감축 방법론",
        "meta": ["방법론명", "버전", "적용 분야", "사업 유형", "작성 기관", "작성일"],
        "sections": [
            ("방법론 개요", "meta"),
            ("적용 조건", "text", "이 방법론을 적용할 수 있는 사업 조건을 항목별로 작성"),
            ("사업 경계", "text", "사업 경계에 포함되는 배출원과 온실가스(CO2·CH4·N2O 등)를 정리"),
            ("베이스라인 시나리오 설정", "text", "사업이 없을 때의 대안 시나리오와 선정 근거를 작성"),
            ("추가성 입증", "text", "법적·경제적 추가성 등 입증 방법과 근거를 작성"),
            ("배출량 및 감축량 산정", "scenarios"),
            ("모니터링 계획", "monitoring"),
            ("참고 문헌", "text", "인용한 지침·계수 출처를 작성"),
        ],
    },
    "business_plan": {
        "title": "외부사업 사업계획서",
        "meta": ["사업명", "사업자", "사업 위치", "사업 유형", "적용 방법론", "사업 기간",
                 "인증 유효기간", "예상 연평균 감축량"],
        "sections": [
            ("사업 개요", "meta"),
            ("사업 내용 및 기술 설명", "text", "설비 구성, 적용 기술, 가동 계획을 작성"),
            ("사업 경계", "text", "사업 경계 도식과 포함 배출원·온실가스를 작성"),
            ("베이스라인 시나리오 및 추가성", "text", "베이스라인 시나리오와 추가성 입증 내용을 작성"),
            ("감축량 산정", "scenarios"),
            ("모니터링 계획", "monitoring"),
            ("환경영향 및 이해관계자 의견", "text", "환경영향과 이해관계자 의견 수렴 결과를 작성"),
        ],
    },
}
PLACEHOLDER = "7F8794"


def _heading(text: str, level: int, font: str, size: float) -> str:
    if level == 1:
        return _w_para(_w_run(text, font, size + 3, bold=True), before=280, after=120,
                       style="Heading1", keep_next=True)
    return _w_para(_w_run(text, font, size + 1.5, bold=True, color=HEADER_FILL),
                   before=160, after=60, style="Heading2", keep_next=True)


def _texts(value: Any, font: str, size: float, hint: str) -> list[str]:
    if not value:
        return [_w_para(_w_run(f"(작성: {hint})", font, size, color=PLACEHOLDER))]
    items = value if isinstance(value, list) else [value]
    return [_w_para(_w_run(str(item), font, size)) for item in items]


def _scenario_blocks(spec: dict[str, Any], prefix: str, font: str, size: float) -> list[str]:
    rpr_eq = docx_rpr(float(spec.get("equation_size_pt") or size + 1.5), ea_font=font)
    body: list[str] = []
    for s_idx, scenario in enumerate(spec.get("scenarios", []), start=1):
        body.append(_heading(f"{prefix}{s_idx} {scenario.get('name', '')}", 2, font, size))
        if scenario.get("description"):
            body.append(_w_para(_w_run(scenario["description"], font, size), after=80, keep_next=True))
        for eq in scenario.get("equations", []):
            eq_par = f"<w:p><w:pPr><w:spacing w:before=\"60\" w:after=\"60\"/><w:jc w:val=\"center\"/></w:pPr>" \
                     f"{omml_para(eq['latex'], rpr_eq)}</w:p>"
            label = _w_run(f"({eq.get('label', '')})" if eq.get("label") else "", font, size)
            body.append(_table([_tr([_tc(eq_par, 8000, borders=False),
                                     _tc(label, 1000, align="right", borders=False)])],
                               [8000, 1000], borders=False))
        where = scenario.get("where", [])
        if where:
            body.append(_w_para(_w_run("여기서,", font, size), before=60, after=40, keep_next=True))
            body.append(where_table(where, font, size))
    return body


def _value_run(value: Any, font: str, size: float) -> str:
    if value in (None, ""):
        return _w_run("(입력)", font, size, color=PLACEHOLDER)
    return _w_run(str(value), font, size)


def _meta_table(keys: list[str], meta: dict[str, Any], font: str, size: float) -> str:
    rows = [_tr([_tc(_w_run(key, font, size, bold=True), 2600, fill=BAND_FILL),
                 _tc(_value_run(meta.get(key), font, size), 6400)]) for key in keys]
    return _table(rows, [2600, 6400])


def _expected_table(spec: dict[str, Any], font: str, size: float) -> str:
    """Yearly expected reductions; values come only from the spec, never computed here."""
    rpr = docx_rpr(size, color=HEADER_TEXT, ea_font=font)
    cols = [("year", "", 1400), ("be", "BE_y", 1900), ("pe", "PE_y", 1900),
            ("le", "LE_y", 1900), ("er", "ER_y", 1900)]
    header = [_tc(_w_run("연도", font, size, bold=True, color=HEADER_TEXT), cols[0][2],
                  fill=HEADER_FILL, align="center")]
    header += [_tc(omml_math(sym, rpr), w, fill=HEADER_FILL, align="center") for _, sym, w in cols[1:]]
    rows = [_tr(header, header=True)]
    for idx, item in enumerate(spec.get("expected_reductions") or [{} for _ in range(3)]):
        fill = BAND_FILL if idx % 2 else None
        rows.append(_tr([_tc(_value_run(item.get(key), font, size), w, fill=fill, align="center")
                         for key, _, w in cols]))
    total = spec.get("expected_total") or {}
    rows.append(_tr([_tc(_w_run("합계", font, size, bold=True), cols[0][2], fill=BAND_FILL, align="center")]
                    + [_tc(_value_run(total.get(key), font, size), w, fill=BAND_FILL, align="center")
                       for key, _, w in cols[1:]]))
    unit = _w_para(_w_run("단위: tCO2eq/yr", font, size - 1, color=PLACEHOLDER), align="right", after=40)
    return unit + _table(rows, [w for _, _, w in cols])


def _monitoring_blocks(spec: dict[str, Any], prefix: str, font: str, size: float,
                       texts: dict[str, Any]) -> list[str]:
    body: list[str] = []
    sub = 0
    if spec.get("monitoring_parameters"):
        sub += 1
        body.append(_heading(f"{prefix}{sub} 모니터링 인자", 2, font, size))
        body.append(param_table(spec["monitoring_parameters"], False, font, size - 1))
    if spec.get("fixed_parameters"):
        sub += 1
        body.append(_heading(f"{prefix}{sub} 사전 고정 인자", 2, font, size))
        body.append(param_table(spec["fixed_parameters"], True, font, size - 1))
    sub += 1
    body.append(_heading(f"{prefix}{sub} 모니터링 절차 및 QA/QC", 2, font, size))
    body += _texts(texts.get("모니터링 절차 및 QA/QC"), font, size,
                   "측정·기록·검교정·자료 보관 절차와 담당 조직을 작성")
    return body


def _kind_body(spec: dict[str, Any], kind: str, font: str, size: float) -> list[str]:
    layout = DOC_KINDS[kind]
    doc = spec.get("document") or {}
    texts = doc.get("texts") or {}
    body: list[str] = []
    for n, (heading, content, *hint) in enumerate(layout["sections"], start=1):
        body.append(_heading(f"{n}. {heading}", 1, font, size))
        if content == "meta":
            body.append(_meta_table(layout["meta"], doc.get("meta") or {}, font, size))
        elif content == "text":
            body += _texts(texts.get(heading), font, size, hint[0] if hint else "내용")
        elif content == "scenarios":
            if texts.get(heading):
                body += _texts(texts[heading], font, size, "")
            body += _scenario_blocks(spec, f"{n}.", font, size)
            if kind == "business_plan":
                body.append(_heading(f"{n}.{len(spec.get('scenarios', [])) + 1} 연도별 예상 감축량",
                                     2, font, size))
                body.append(_expected_table(spec, font, size - 1))
        elif content == "monitoring":
            body += _monitoring_blocks(spec, f"{n}.", font, size, texts)
    return body


def build_document_xml(spec: dict[str, Any]) -> str:
    font = spec.get("font") or DEFAULT_FONT
    size = float(spec.get("font_size_pt") or 10.5)
    kind = (spec.get("document") or {}).get("kind") or "monitoring"
    body: list[str] = []
    title = spec.get("title") or (DOC_KINDS[kind]["title"] if kind in DOC_KINDS else "")
    if title:
        body.append(_w_para(_w_run(title, font, size + 5, bold=True, color=HEADER_FILL),
                            after=200, style="Title"))
    if kind in DOC_KINDS:
        body += _kind_body(spec, kind, font, size)
    else:
        section = 0
        if spec.get("scenarios"):
            section += 1
            body.append(_heading(f"{section}. 시나리오별 산정식", 1, font, size))
            body += _scenario_blocks(spec, f"{section}.", font, size)
        if spec.get("monitoring_parameters"):
            section += 1
            body.append(_heading(f"{section}. 모니터링 인자", 1, font, size))
            body.append(param_table(spec["monitoring_parameters"], False, font, size - 1))
        if spec.get("fixed_parameters"):
            section += 1
            body.append(_heading(f"{section}. 사전 고정 인자", 1, font, size))
            body.append(param_table(spec["fixed_parameters"], True, font, size - 1))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1200" w:bottom="1440" w:left="1200" w:header="708" '
            'w:footer="708" w:gutter="0"/></w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<w:body>{"".join(body)}{sect}</w:body></w:document>')


def _styles_xml(font: str, size_pt: float) -> str:
    half = int(round(size_pt * 2))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:styles xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/>'
            f'</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="276" '
            f'w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            f'<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>'
            f'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr></w:style>'
            f'<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>')


def _settings_xml() -> str:
    # mathPr makes Word treat display math with Cambria Math by default.
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings xmlns:w="{W_NS}" xmlns:m="{M_NS}"><m:mathPr><m:mathFont m:val="Cambria Math"/>'
            f'<m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/>'
            f'<m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr></w:settings>')


def require_new_output(out: Path, source: Optional[Path] = None) -> None:
    """Protect source aliases and existing deliverables; ZIP exclusive creation closes the race."""
    if source is not None and out.resolve() == source.resolve():
        raise ValueError("Input and output must differ. Choose a new versioned output path.")
    if out.exists() or out.is_symlink():
        raise FileExistsError(f"Output already exists; choose a new versioned path: {out}")


def write_docx(spec: dict[str, Any], out: Path) -> None:
    require_new_output(out)
    font = spec.get("font") or DEFAULT_FONT
    size = float(spec.get("font_size_pt") or 10.5)
    out.parent.mkdir(parents=True, exist_ok=True)
    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    doc_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            '</Types>'),
        "_rels/.rels": (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{rels_ns}">'
            f'<Relationship Id="rId1" Type="{doc_type}/officeDocument" Target="word/document.xml"/></Relationships>'),
        "word/_rels/document.xml.rels": (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{rels_ns}">'
            f'<Relationship Id="rId1" Type="{doc_type}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{doc_type}/settings" Target="settings.xml"/></Relationships>'),
        "word/document.xml": build_document_xml(spec),
        "word/styles.xml": _styles_xml(font, size),
        "word/settings.xml": _settings_xml(),
    }
    with zipfile.ZipFile(out, "x", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)


# ---------------------------------------------------------------------------
# PowerPoint token injection (direct-PPTX routes)
# ---------------------------------------------------------------------------

TOKEN = re.compile(r"\{\{(eq|sym):([A-Za-z0-9_.-]+)\}\}")


def _token_table(spec: dict[str, Any]) -> dict[tuple[str, str], str]:
    table: dict[tuple[str, str], str] = {}
    for scenario in spec.get("scenarios", []):
        for eq in scenario.get("equations", []):
            if eq.get("id"):
                table[("eq", eq["id"])] = eq["latex"]
    for key in ("monitoring_parameters", "fixed_parameters"):
        for param in spec.get(key, []):
            if param.get("id"):
                table[("sym", param["id"])] = param["symbol"]
    return table


PARA_RE = re.compile(r"<a:p>(?:(?!</a:p>).)*?</a:p>|<a:p [^>]*>(?:(?!</a:p>).)*?</a:p>", re.S)


def _inject_slide(xml: str, table: dict[tuple[str, str], str], font: str,
                  missing: list[str]) -> tuple[str, int]:
    count = 0

    def replace_para(match: re.Match) -> str:
        nonlocal count
        para = match.group(0)
        text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", para))
        token = TOKEN.fullmatch(text.strip())
        if not token:
            if TOKEN.search(text):
                return _inline_tokens(para)
            return para
        key = (token.group(1), token.group(2))
        if key not in table:
            missing.append(f"unknown token {text.strip()!r}")
            return para
        rpr = _rpr_from(para)
        algn_m = re.search(r'<a:pPr[^>]*\balgn="(\w+)"', para)
        ppr_m = re.search(r"<a:pPr[^>]*/>|<a:pPr[^>]*>.*?</a:pPr>", para, re.S)
        ppr = ppr_m.group(0) if ppr_m else "<a:pPr/>"
        jc = {"ctr": "center", "r": "right"}.get(algn_m.group(1) if algn_m else "l", "left")
        math = omml_para(table[key], rpr, jc=jc) if key[0] == "eq" else omml_math(table[key], rpr)
        count += 1
        return f'<a:p>{ppr}<a14:m xmlns:a14="{A14_NS}">{math}</a14:m></a:p>'

    def _rpr_from(fragment: str):
        size_m = re.search(r'<a:rPr[^>]*\bsz="(\d+)"', fragment)
        color_m = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', fragment)
        size_pt = int(size_m.group(1)) / 100 if size_m else 14.0
        return pptx_rpr(size_pt, color=color_m.group(1) if color_m else None, ea_font=font)

    def _inline_tokens(para: str) -> str:
        """Split runs around tokens inside a sentence; math goes inline."""
        nonlocal count

        def split_run(run_match: re.Match) -> str:
            nonlocal count
            run, props, text = run_match.group(0), run_match.group(1), run_match.group(2)
            if not TOKEN.search(text):
                return run
            pieces, pos = [], 0
            for tok in TOKEN.finditer(text):
                key = (tok.group(1), tok.group(2))
                if key not in table:
                    missing.append(f"unknown token {tok.group(0)!r}")
                    continue
                if tok.start() > pos:
                    pieces.append(f"<a:r>{props}<a:t>{text[pos:tok.start()]}</a:t></a:r>")
                pieces.append(f'<a14:m xmlns:a14="{A14_NS}">{omml_math(table[key], _rpr_from(props))}</a14:m>')
                count += 1
                pos = tok.end()
            if pos < len(text):
                pieces.append(f"<a:r>{props}<a:t>{text[pos:]}</a:t></a:r>")
            return "".join(pieces)

        new_para = re.sub(r"<a:r>\s*(<a:rPr[^>]*/>|<a:rPr[^>]*>.*?</a:rPr>)?\s*<a:t>([^<]*)</a:t>\s*</a:r>",
                          split_run, para, flags=re.S)
        leftover = "".join(re.findall(r"<a:t>([^<]*)</a:t>", new_para))
        if TOKEN.search(leftover):
            missing.append(f"token split across runs (retype it as one run): {leftover.strip()!r}")
        return new_para

    new_xml = PARA_RE.sub(replace_para, xml)
    if count:
        new_xml = _wrap_alternate_content(new_xml)
    return new_xml, count


def _wrap_alternate_content(xml: str) -> str:
    """Wrap each <p:sp> or table <p:graphicFrame> holding a14:m in mc:AlternateContent (PowerPoint's own form)."""
    out, pos = [], 0
    for match in re.finditer(r"<p:sp>.*?</p:sp>|<p:graphicFrame>.*?</p:graphicFrame>", xml, re.S):
        shape = match.group(0)
        out.append(xml[pos:match.start()])
        if "<a14:m" in shape and "mc:AlternateContent" not in xml[max(0, match.start() - 200):match.start()]:
            fallback = re.sub(r"<a14:m .*?</a14:m>", "", shape, flags=re.S)
            out.append(
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                f'xmlns:a14="{A14_NS}"><mc:Choice Requires="a14">{shape}</mc:Choice>'
                f"<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>")
        else:
            out.append(shape)
        pos = match.end()
    out.append(xml[pos:])
    return "".join(out)


def inject_pptx(src: Path, spec: dict[str, Any], out: Path) -> tuple[int, list[str]]:
    require_new_output(out, src)
    table = _token_table(spec)
    font = spec.get("font") or DEFAULT_FONT
    missing: list[str] = []
    total = 0
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as zin:
        patched: dict[str, bytes] = {}
        for info in zin.infolist():
            data = zin.read(info.filename)
            if re.match(r"ppt/slides/slide\d+\.xml$", info.filename):
                xml, count = _inject_slide(data.decode("utf-8"), table, font, missing)
                total += count
                data = xml.encode("utf-8")
            patched[info.filename] = data
        if missing:
            return total, missing
        if total == 0:
            return 0, ["No equation/symbol tokens found; output was not created."]
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "x", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zout.writestr(info, patched[info.filename])
    return total, missing


# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------

def verify_office(path: Path) -> dict[str, Any]:
    """Count OMML objects and list formula-like plain-text runs."""
    result: dict[str, Any] = {"file": str(path), "equations": 0, "plain_formulas": []}
    with zipfile.ZipFile(path) as zf:
        if path.suffix.lower() == ".docx":
            parts = [n for n in zf.namelist() if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)]
            text_tag = "w:t"
        else:
            parts = sorted((n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                           key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[-1]).group(1)))
            text_tag = "a:t"
        for name in parts:
            xml = zf.read(name).decode("utf-8")
            if path.suffix.lower() == ".pptx":
                # Fallback branches repeat the equation as plain text by design.
                xml = re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", xml, flags=re.S)
                result["equations"] += len(re.findall(r"<a14:m\b", xml))
            else:
                result["equations"] += len(re.findall(r"<m:oMath[ >]", xml))
            paragraphs = re.split(r"</(?:w|a):p>", xml)
            for para in paragraphs:
                para_no_math = re.sub(r"<m:oMath\b.*?</m:oMath>", "", para, flags=re.S)
                text = "".join(re.findall(rf"<{text_tag}(?: [^>]*)?>([^<]*)</{text_tag}>", para_no_math))
                if text and FORMULA_LIKE.search(text):
                    result["plain_formulas"].append({"part": name, "text": text.strip()[:120]})
    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _load_spec(path: Path) -> dict[str, Any]:
    spec = json.loads(path.read_text(encoding="utf-8"))
    errors = validate_spec(spec)
    if errors:
        raise SystemExit("ERROR  equation spec:\n  " + "\n  ".join(errors))
    return spec


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Editable OMML equations and monitoring tables for Word/PowerPoint.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", required=True)
    ex = sub.add_parser("example", help="print a sample spec")
    ex.add_argument("--kind", choices=["monitoring", *DOC_KINDS],
                    help="add the document block for a methodology or business plan")
    docx = sub.add_parser("docx", help="write a Word document")
    docx.add_argument("spec", type=Path)
    docx.add_argument("-o", "--output", type=Path, required=True)
    docx.add_argument("--kind", choices=["monitoring", *DOC_KINDS],
                      help="document skeleton (default: spec document.kind, else monitoring)")
    inj = sub.add_parser("inject-pptx", help="replace {{eq:id}} / {{sym:id}} tokens in a PPTX")
    inj.add_argument("pptx", type=Path)
    inj.add_argument("spec", type=Path)
    inj.add_argument("-o", "--output", type=Path, required=True)
    injd = sub.add_parser("inject-docx", help="turn equation markers in a Word draft into OMML")
    injd.add_argument("docx", type=Path)
    injd.add_argument("spec", type=Path, nargs="?", default=None,
                      help="equation spec (needed for {{eq:}}, {{sym:}}, {{where:}}, tables)")
    injd.add_argument("-o", "--output", type=Path, required=True)
    injd.add_argument("--convert-plain", action="store_true",
                      help="also convert formulas typed as plain text (Word subscripts)")
    injd.add_argument("--no-number", action="store_true", help="display equations without (n) labels")
    injd.add_argument("--font", default=None, help="East Asian font for Hangul inside equations")
    ver = sub.add_parser("verify", help="count OMML and flag plain-text formulas")
    ver.add_argument("file", type=Path)
    ver.add_argument("--expect", type=int, default=None, help="minimum equation count")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "example":
        print(json.dumps(example_spec(args.kind), ensure_ascii=False, indent=2))
        return 0
    if args.command == "docx":
        spec = _load_spec(args.spec)
        if args.kind:
            spec.setdefault("document", {})["kind"] = args.kind
        write_docx(spec, args.output)
        report = verify_office(args.output)
        print(f"OK  {args.output}  equations={report['equations']}")
        return 0
    if args.command == "inject-pptx":
        spec = _load_spec(args.spec)
        count, problems = inject_pptx(args.pptx, spec, args.output)
        for problem in problems:
            print(f"ERROR  {problem}", file=sys.stderr)
        print(f"{'OK' if not problems else 'FAIL'}  {args.output}  injected={count}")
        return 1 if problems else 0
    if args.command == "inject-docx":
        from docx_equations import inject_docx
        spec = _load_spec(args.spec) if args.spec else {}
        rep = inject_docx(args.docx, spec, args.output, number=not args.no_number,
                          convert_plain=args.convert_plain, font=args.font)
        for problem in rep.errors:
            print(f"ERROR  {problem}", file=sys.stderr)
        for note in rep.skipped:
            print(f"NOTE   {note}")
        if rep.errors:
            print(f"FAIL  {args.output} not written")
            return 1
        check = verify_office(args.output)
        for item in check["plain_formulas"]:
            print(f"WARN   plain-text formula left in {item['part']}: {item['text']}")
        print(f"OK  {args.output}  display={rep.display} inline={rep.inline} tables={rep.tables} "
              f"converted={rep.converted} equations={check['equations']}")
        return 0
    report = verify_office(args.file)
    failed = bool(report["plain_formulas"]) or (
        args.expect is not None and report["equations"] < args.expect)
    for item in report["plain_formulas"]:
        print(f"ERROR  plain-text formula in {item['part']}: {item['text']}")
    if args.expect is not None and report["equations"] < args.expect:
        print(f"ERROR  expected at least {args.expect} equations, found {report['equations']}")
    print(f"{'FAIL' if failed else 'OK'}  {report['file']}  equations={report['equations']}")
    return 1 if failed else 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(f"ERROR  {exc}", file=sys.stderr)
        raise SystemExit(1)
