# 감축사업 수식·모니터링 인자 자동화 (reduction-equations)

대충 적은 감축사업 산정식을 **Word·PowerPoint에서 바로 고칠 수 있는 수식(Office Math)** 으로 바꾸고,
기호·정의·단위 표, 모니터링 요약표, 인자별 `데이터/인자` 표를 제출 서식에 맞춰 자동으로 만듭니다.

- 서식: 대한상의 탄소감축인증센터(KCS), 환경부 외부사업 방법론
- 출력: Word(.docx), PowerPoint(.pptx) — 수식·기호는 모두 편집 가능한 수식 개체
- 가져오기: 기존 방법론·사업계획서(.docx), 인증위원회 발표(.pptx)의 수식과 표
- 웹 편집기: 입력하는 대로 수식 미리보기, 인자 표 편집, Word·PPT 내보내기

> 이 저장소는 Slide Master(`.claude/skills/ppt-master/scripts/build_reduction_equations_package.py`)에서
> 생성됩니다. 수정은 Slide Master에서 하고 다시 생성하세요.

## 설치

Windows: `powershell -ExecutionPolicy Bypass -File setup.ps1`
macOS·Linux: `sh setup.sh`

Python 3.10 이상이 필요합니다. `.venv`를 만들고 `python-pptx`를 설치한 뒤 예시를 검사합니다.

## 스킬로 쓰기

- **Claude Code / Codex**: 이 폴더를 `.claude/skills/reduction-equations/`(Codex는 `.codex/skills/`)에 두면 자동 인식됩니다.
- **claude.ai**: 설정 → 기능 → 스킬에서 `reduction-equations.zip`을 올립니다.

"감축사업 수식 만들어줘", "모니터링 인자 표 KCS 서식으로", "환경부 외부사업 서식으로 워드 만들어줘",
"이 방법론 수식 가져와서 PPT로" 같은 요청에 쓰입니다.

## 명령

```
python scripts/monitoring_plan.py from-text examples/plan_example.txt -o plan.json
python scripts/monitoring_plan.py import 방법론.docx -o plan.json
python scripts/monitoring_plan.py check plan.json
python scripts/monitoring_plan.py docx plan.json -o exports/수식_모니터링계획_v1.docx
python scripts/monitoring_plan.py pptx plan.json -o exports/수식_인자_v1.pptx
python scripts/office_equations.py verify exports/수식_모니터링계획_v1.docx
```

서식은 계획 파일의 `options.form`으로 고릅니다: `"kcs"`(기본) 또는 `"moe"`(환경부 외부사업).

## 빠른 입력 규칙

```
# 베이스라인 배출량
> 섹션 설명
BEy = BE_MAT,y + BE_ENE,y
BE_MFG,y = sum_j((Q_PJ,j,y * QW_BL,j) * DHC_PJ,j/DHC_BL,j * EF_BL,MFG,j)
Q_PJ,j,y | y년도 사업필터 순판매량 | 개/년 | 모니터링 | 값
```

| 입력 | 결과 |
|---|---|
| `BE_MFG,y`, `Q_{PJ, j}` | 아래첨자 |
| `BEy`, `n0` | `BE_y`, `n_0` |
| `a * b` | × |
| `A/B` / `a × b / c` | 분수 (빈칸 있는 /는 왼쪽 곱 전체가 분자) |
| `x^2`, `(…)^x_a` | 위첨자 |
| `sum_j(…)` | Σ |
| `bar(DP)_BL,a` | 윗줄 기호 |
| `(식 3)` | 식 번호 |

## 확인 범위와 한계

- 수식 구조, 기호·단위 대응, 수식 개체 생성은 확인합니다. 계산값·계수의 정확성은 확인하지 않습니다.
- 환경부 서식의 표 머리말은 외부사업 방법론의 일반 구성을 따릅니다. 제출 전 최신 공식 서식과 대조하세요.
  다르면 `options.labels`로 해당 문구만 바꿀 수 있습니다.
