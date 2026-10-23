---
name: reduction-equations
description: Turn roughly typed GHG reduction-project formulas (감축사업 산정식) into native, editable Office Math in Word and PowerPoint, and auto-build the symbol tables, monitoring summary and per-parameter 데이터/인자 tables in the KCS (대한상의 탄소감축인증센터) or 환경부 외부사업 methodology form. Use when the user asks to make, convert, import or fix 수식·모니터링 인자·모니터링 계획·기호 정의표·데이터/인자 표 for a 방법론, 사업계획서, 모니터링 보고서 or 인증위원회 발표자료 — "수식 만들어줘", "수식을 편집 가능하게", "모니터링 인자 표 자동화", "방법론 수식 워드로", "KCS 서식으로", "환경부 외부사업 서식", "수식 편집기 열어줘" — or supplies an existing .docx/.pptx whose formulas and tables should be reused.
---

# 감축사업 수식·모니터링 인자 자동화

목적은 **제출 서식에 맞춘 수식과 표를 빠르게 만드는 것**이다. 수식은 대충 텍스트로 받고, 결과는 Word·PowerPoint에서 **수식 편집기로 바로 고칠 수 있는 Office Math(OMML)** 로 만든다. 그림·일반 텍스트·윤곽선 수식은 결과물로 인정하지 않는다.

감축사업 자료(방법론, 사업계획서, 모니터링 보고서, 인증위원회 발표)를 만드는 요청이면 사용자가 말하지 않아도 첫 답변에서 이 기능을 추가 작업으로 제안한다.

도구 위치: 이 스킬 폴더의 `scripts/`(아래 명령의 스크립트는 모두 이 폴더 기준). 실행 환경별 준비:

- **내 PC (Claude Code·Codex)**: 처음 한 번 `setup.ps1`(Windows) 또는 `sh setup.sh`로 `.venv`를 만들고 그 Python으로 실행한다.
- **claude.ai 채팅·Cowork**: 설치 단계는 건너뛴다. `python-pptx`가 없을 때만 `pip install python-pptx`. 결과 파일은 `/mnt/user-data/outputs/`에 저장해 링크로 준다. 웹 편집기(1장 7단계)는 쓸 수 없으니 수식은 LaTeX로, 인자는 채팅 표로 보여 주고 고칠 점을 받는다.
- **ChatGPT**: `CHATGPT.md`를 먼저 읽고 그 바꿔 읽기 표를 따른다.

| 도구 | 역할 |
|---|---|
| `math_shorthand.py` | 대충 텍스트 → LaTeX 부분집합 (OMML 변환 가능 여부까지 검사) |
| `monitoring_plan.py` | 계획 파일(JSON 또는 빠른 텍스트) → 기호 자동 매칭, 검수, Word·PPT 출력, 기존 문서 가져오기 |
| `equation_editor/server.py` | 같은 기능의 로컬 웹 편집기 |
| `office_equations.py` | 이미 쓴 Word·PPT 초안의 `{{eq:id}}`·`{{sym:id}}` 자리표시를 수식으로 바꾸기, `verify` 검사 |

## 1. 작업 순서

1. **원본이 있으면 먼저 가져온다.** 사용자가 방법론·사업계획서·발표자료를 주면 그 파일의 수식·기호표·모니터링 표를 읽어 계획 파일을 만든다. 원본은 고치지 않는다.
   ```
   python monitoring_plan.py import <원본.docx|pptx> -o projects/<사업>/notes/plan.json
   ```
   방법론·사업계획서(Word)와 인증위원회 발표(PPT) 모두 읽는다. 사업계획서의 `적용된 값`도 가져오므로 기호표에 값 열을 켤 수 있다(`options.show_values`). 가져온 결과의 수식 개수·인자 개수를 원본과 대조하고, 출력되는 WARN을 사용자에게 알린다.
2. **원본이 없으면 빠른 텍스트로 받는다.** 사용자가 적은 수식을 아래 2장 규칙으로 옮겨 `plan.txt`에 쓰고 `from-text`로 계획 파일을 만든다. 원문에 없는 계수·값·출처를 지어내지 않는다. 비어 있으면 비워 두고 검수 경고로 남긴다.
3. **서식을 정한다.** `options.form`: `"kcs"`(기본, 대한상의 탄소감축인증센터) 또는 `"moe"`(환경부 외부사업 방법론). 사용자가 서식을 말하지 않았으면 원본 파일의 서식을 따르고, 원본도 없으면 KCS로 만든 뒤 바꿀 수 있다고 알린다.
4. **검수한다.** `python monitoring_plan.py check plan.json` — ERROR가 하나라도 있으면 출력하지 않는다. WARN(정의·단위·출처 빈칸, 같은 기호 중복 정의, 수식에 없는 인자)은 목록으로 보고한다.
5. **출력한다.** 파일명은 새 버전으로 정한다(기존 파일 덮어쓰기 불가).
   ```
   python monitoring_plan.py docx plan.json -o projects/<사업>/exports/<사업>_수식_모니터링계획_v1.docx
   python monitoring_plan.py pptx plan.json -o projects/<사업>/exports/<사업>_수식_인자_v1.pptx
   python office_equations.py verify <출력 파일>
   ```
6. **눈으로 확인한다.** PPT와 Word를 PowerPoint·Word(또는 LibreOffice)에서 PDF로 내보내 페이지를 본다. 수식 잘림·표 넘침·기호 깨짐이 없어야 완료다. 렌더하지 못했으면 "렌더 미확인"이라고 보고한다.
7. **사용자가 직접 고치고 싶어 하면** 웹 편집기를 띄운다: `powershell -ExecutionPolicy Bypass -File scripts/equation_editor/start_equation_editor.ps1 [--plan plan.json]` (macOS·Linux: `.venv/bin/python scripts/equation_editor/server.py`). 편집기에서 저장한 JSON을 다시 4~5단계에 쓴다. 팀이 함께 쓸 상시 주소가 필요하면 Tailscale이 깔린 PC에서 `equation_editor/serve_tailscale.ps1 -Install`을 실행한다(`https://<PC>.<tailnet>.ts.net:10001/`, 로그인 때 자동 시작, 해제는 `-Uninstall`).

기존 Word·PPT 초안 안에 수식만 넣을 때는 `office_equations.py inject-docx` / `inject-pptx`를 쓰고, 토큰 표는 `monitoring_plan.py spec plan.json -o spec.json`으로 만든다(`{{sym:Q_PJ.j.y}}`처럼 쉼표는 마침표).

## 2. 빠른 텍스트 규칙

```
# 베이스라인 배출량                 ← 섹션 (## 는 하위 섹션)
> 사업이 없을 때의 배출량            ← 섹션 설명
BEy = BE_MAT,y + BE_ENE,y           ← 수식 (끝에 (식 3)을 쓰면 그 번호 사용)
BE_MFG,y = sum_j((Q_PJ,j,y * QW_BL,j) * DHC_PJ,j/DHC_BL,j * EF_BL,MFG,j)
Q_PJ,j,y | y년도 사업필터 제품군 j의 순판매량 | 개/년 | 모니터링 | 값
```

| 입력 | 결과 |
|---|---|
| `BE_MFG,y`, `Q_{PJ, j}` | 아래첨자 (빈칸·연산자까지, 중괄호면 그 안 전체) |
| `BEy`, `ERy`, `n0` | `BE_y`, `ER_y`, `n_0` 축약 |
| `a * b`, `a × b` / `a · b` | × / · |
| `A/B` | A, B만 분수 (`DHC_PJ,j/DHC_BL,j`) |
| `a × b / c` (빈칸 있는 /) | 왼쪽 곱 전체가 분자 |
| `(a+b)/2` | 바깥 괄호 하나는 지우고 분수 |
| `x^2`, `(…)^x_a`, `e^(a+b)` | 위첨자 |
| `sum_j(…)`, `sum_i sum_j {…}`, `sum_{i=1}^{n}` | Σ (괄호·중괄호는 그대로 보임) |
| `bar(DP)_BL,a` | 윗줄 기호 D̅P̅ |
| `sqrt(x)`, `min(a, b)`, `max`, `ln`, `exp` | 함수 |
| `sigma`, `σ`, `<=`, `>=`, `!=` | 그리스 문자, 부등호 |
| 한글 | 직립 글자 그대로 |
| 표·설명 칸의 `$…$` | 문장 속 인라인 수식 |

인자 행: `기호 | 정의 | 단위 | 구분(산출/모니터링/고정) | 값 | 데이터 출처 | 측정 방법 | 모니터링 주기 | QA/QC | 데이터 목적 | 기타 의견`. 앞에서부터 필요한 만큼만 쓴다.

## 3. 자동 매칭 규칙

- 수식에 나온 기호를 순서대로 모아 섹션마다 **기호·정의·단위** 표(환경부 서식은 "여기서," 목록)를 만든다. Σ의 첨자(j, a, h)는 기호에서 뺀다.
- 인자 사전에 없는 기호는 자동으로 추가하고 노란색(편집기)·WARN(검수)으로 표시한다.
- 구분 추정: 등호 왼쪽 기호와 BE·PE·LE·ER 계열은 **산출**, 아래첨자에 `y`가 있으면 **모니터링**, 나머지는 **고정**. 사용자가 정한 구분이 항상 우선한다.
- 데이터 목적이 비어 있으면 그 기호가 쓰인 상위 섹션 이름으로 채운다(예: 베이스라인 배출량 산정).
- 산출 인자는 데이터/인자 세부표에서 빠지고, 모니터링·고정 인자만 요약표와 세부표에 들어간다.

## 4. 서식

| 항목 | KCS (`kcs`) | 환경부 외부사업 (`moe`) |
|---|---|---|
| 수식 아래 | 기호·정의·단위 표 | "여기서," 목록 (기호 : 정의 [단위]) |
| 요약표 | 구분·기호·항목·고정/모니터링·비고 | 번호·데이터/변수·설명·단위·구분 |
| 세부표 머리 | 데이터/인자 | 데이터/변수 |
| 세부표 구분 | 고정 데이터 및 인자 / 모니터링 데이터 및 인자 | 사전에 결정되는 데이터 및 변수 / 모니터링 대상 데이터 및 변수 |

KCS 서식은 뉴라이즌 KCS 제출본(2026-09)을 기준으로 맞췄다. 환경부 서식의 표 머리말은 외부사업 방법론의 일반 구성을 따른 것이므로, 제출 전 최신 공식 서식과 대조하라고 사용자에게 알린다. 다른 표현이 필요하면 `options.labels`로 해당 항목만 바꾼다(예: `{"param_word": "데이터/인자"}`).

## 5. 완료 보고

- 수식 수, 인자 수(산출·모니터링·고정), 남은 WARN, 서식, 출력 파일 경로.
- `verify` 결과의 수식 객체 수, 렌더 확인 여부.
- 계산값·계수의 정확성은 이 도구가 검증하지 않는다는 점. 수식 구조와 기호·단위 대응만 확인했다고 구분해서 말한다.
