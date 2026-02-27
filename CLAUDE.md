# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

수소 에너지 및 연료전지 산업에 대한 종합 연구 프로젝트. 도서 분석(3권), 산업 동향 분석, Python 기반 PPT 자동 생성을 포함한다. 모든 문서는 한국어로 작성되어 있다.

## 프로젝트 구조

- `수소에너지/` — Part 1: 도서 3권 분석(기술·투자·지정학), 연구 동향, 한국 전략, PPT 생성
- `연료전지/` — Part 2: 글로벌 연료전지 시장, 미·중 경쟁, 한국 R&D·인프라 현황, PPT 생성

각 디렉토리에 마크다운 분석 문서들과 `create_ppt.py` / `create_fuelcell_ppt.py` PPT 생성 스크립트가 있다.

## PPT 생성 명령

```bash
# 수소에너지 발표자료 생성 (24장)
cd 수소에너지 && python create_ppt.py

# 연료전지 발표자료 생성 (27장)
cd 연료전지 && python create_fuelcell_ppt.py
```

의존성: `python-pptx` (pip install python-pptx)

## PPT 스크립트 아키텍처

두 스크립트(`create_ppt.py`, `create_fuelcell_ppt.py`)는 동일한 패턴을 따른다:

- **색상 상수**: NAVY(27,58,92), GREEN(46,204,113), WHITE 등 모듈 상단에 정의
- **스타일링 함수**: `set_font()`, `add_background()`, `add_navy_header_bar()`, `add_green_accent_line()`, `add_slide_number()`
- **콘텐츠 함수**: `add_body_textbox()`, `add_bullet()` / `add_bullet_point()`, `set_cell()`, `style_header_row()`, `style_data_rows()`
- **슬라이드 함수**: `slide_01_cover()` ~ `slide_NN_*()` — 각 슬라이드가 독립된 함수
- **템플릿 함수**: `setup_slide()` / `setup_content_slide()` — 슬라이드 기본 구조(헤더바, 악센트라인, 번호)를 설정
- **진입점**: `main()` → 각 슬라이드 함수를 순차 호출 → `.pptx` 파일 저장

슬라이드 추가 시: 새 `slide_NN_topic()` 함수 작성 → `main()`에 호출 추가 → `TOTAL_SLIDES` 상수 업데이트(연료전지 스크립트).

## 마크다운 문서 관계

수소에너지 디렉토리:
- `01~03`: 개별 도서 분석 (기술→투자→지정학 순)
- `04`: 세 권 비교분석
- `05`: 글로벌 연구 동향 (2020~2025)
- `06`: 미래 전망 및 한국 전략
- `수소에너지_도서분석.md`: 위 내용을 하나로 통합한 전체본

연료전지 디렉토리:
- 글로벌 동향 → 미·중 비교 → 한국 현황 순서로 범위가 좁아지는 구조

## Git 커밋/푸시 규칙

**`.md` 파일은 이 프로젝트의 핵심 산출물이다.** git commit 또는 push 시 반드시 다음을 지킨다:

- 변경된 `.md` 파일이 있으면 **절대 누락하지 않고 함께 커밋**한다. `git status`로 unstaged 상태의 `.md` 파일이 없는지 반드시 확인한다.
- `.py` 스크립트만 커밋하고 관련 `.md` 파일을 빠뜨리는 실수를 하지 않는다.
- 커밋 전 체크리스트: `git status`에서 modified/untracked `.md` 파일이 모두 staged 되었는지 확인 → 빠진 것이 있으면 추가 → 커밋

## GitHub 한국어 파일명 렌더링 문제 대응

GitHub 웹 UI에서 한국어 폴더/파일명이 `"\354\210\230\354\206\214..."` 같은 8진수 이스케이프 코드로 표시되는 경우가 있다. Git 데이터 자체는 정상(UTF-8)이며, clone하면 한국어가 올바르게 나온다. GitHub 트리 렌더링 캐시 문제이다.

**해결 방법**: 한국어 이름을 가진 폴더 내 파일에 작은 변경을 커밋 & push하여 새로운 트리 SHA를 생성하면 GitHub가 트리를 다시 파싱하면서 정상 렌더링된다.

```bash
# 1. 한국어 폴더 내 파일에 사소한 변경 추가
echo "" >> "수소에너지/아무파일.md"
echo "" >> "연료전지/아무파일.md"
# 2. 커밋 & push
git add -A && git commit -m "chore: GitHub 트리 인덱스 갱신" && git push
# 3. GitHub 웹에서 파일명 정상 표시 확인
```

**사전 예방**: `git config --global core.quotepath false` 설정이 되어 있는지 확인한다 (현재 설정됨).

## 작업 시 유의사항

- 폰트는 `맑은 고딕`(Malgun Gothic) 통일
- PPT 단위: `python-pptx`의 `Inches()`, `Pt()` 사용
- 두 PPT 스크립트 간 유틸리티 함수가 유사하지만 별도 파일로 관리됨 — 수정 시 양쪽 일관성 확인
- 슬라이드 번호는 1-indexed, 목차(Table of Contents) 슬라이드에서 전체 장수를 참조함
