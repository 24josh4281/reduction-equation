# AGENTS.md

이 저장소는 감축사업 수식·모니터링 인자 자동화 스킬이다. 작업 절차는 `SKILL.md`가 정한다.
Codex(GPT)는 수식·모니터링 인자·기호 정의표·데이터/인자 표 요청을 받으면 `SKILL.md`를 먼저 읽고 그대로 따른다.

- Python: `.venv`가 없으면 `setup.ps1`(Windows) 또는 `sh setup.sh`를 먼저 실행하고, 이후 `.venv`의 Python으로 실행한다.
- 수식은 Word·PPT 네이티브 Office Math로만 만든다. 그림·일반 텍스트 수식은 결과물이 아니다.
- 원문에 없는 계수·값·출처를 지어내지 않는다.
- 이 저장소는 Slide Master에서 생성된다. 여기서 고친 코드는 다음 생성 때 덮어써진다.
