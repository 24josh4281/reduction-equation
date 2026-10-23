# ChatGPT에서 실행하기

`SKILL.md`를 ChatGPT의 Python 도구에서 실행할 때 바꿔 읽을 점이다. 여기 없는 내용은 `SKILL.md`를 따른다.

## 준비 (채팅마다 한 번)

```python
import glob, os, subprocess, sys, zipfile
src = glob.glob('/mnt/data/reduction-equations-chatgpt*.zip')[0]
zipfile.ZipFile(src).extractall('/mnt/data/re')
TOOLS = '/mnt/data/re/reduction-equations/scripts'
ENV = dict(os.environ, PYTHONUTF8='1')
import pptx  # 없으면 "코드 실행 환경에 python-pptx가 없다"고 알리고 멈춘다

def run(script, *args):
    p = subprocess.run([sys.executable, f'{TOOLS}/{script}', *args], capture_output=True, text=True, env=ENV)
    print(p.stdout, p.stderr)
    return p.returncode
```

## 바꿔 읽기

| SKILL.md | ChatGPT |
|---|---|
| `python monitoring_plan.py import 원본 -o plan.json` | `run('monitoring_plan.py', 'import', '/mnt/data/<첨부 파일>', '-o', '/mnt/data/work/plan.json')` |
| `projects/<사업>/notes/` | `/mnt/data/work/` |
| `projects/<사업>/exports/` | `/mnt/data/exports/` |
| 사용자가 준 원본 | 채팅에 첨부된 `/mnt/data/<파일명>` |
| 빠른 텍스트 `plan.txt` | `/mnt/data/work/plan.txt`에 써서 `from-text` |
| `setup.ps1`, `.venv` | 쓰지 않는다 |
| 1장 6단계 렌더 확인 | `soffice`가 있으면 PDF로 바꿔 페이지 이미지를 보고, 없으면 "렌더 미확인"이라고 보고한다 |
| 1장 7단계 웹 편집기 | 쓸 수 없다. 수식은 LaTeX(`$$…$$`)로, 인자는 채팅 표로 보여 주고 고칠 점을 받아 계획 파일을 고친 뒤 새 버전으로 다시 낸다 |

## 결과 전달

- 파일은 `/mnt/data/exports/`에 새 버전 이름(`_v1`, `_v2` …)으로 저장하고 다운로드 링크(`sandbox:/mnt/data/exports/<파일명>`)를 준다.
- 수식 수, 인자 수(산출·모니터링·고정), 남은 WARN, 서식, `verify` 수식 개수를 짧게 보고한다.
