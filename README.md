# slide-converter

슬라이드 장표(.pptx)를 입력하면 **본사 템플릿 양식**에 맞게 내용을 재구성해 새 .pptx 로 만들어 주는 CLI 도구입니다.

```
원본.pptx ──▶ [1] 내용 추출 ──▶ [2] Claude 재구성 ──▶ [3] 템플릿 렌더링 ──▶ 결과.pptx
              제목/글머리/표/      본사 스타일 가이드·      본사 템플릿 레이아웃의
              이미지/노트          레이아웃 카탈로그 기준    placeholder 서식 상속
```

- **[1] 추출** (`extract.py`): 제목, 글머리(들여쓰기 수준 포함), 표, 이미지, 발표자 노트를 읽습니다. 차트처럼 옮길 수 없는 요소는 기록해 두었다가 노트에 `[원본 확인 필요]`로 남깁니다.
- **[2] 재구성** (`planner.py`): Claude 가 템플릿 레이아웃 목록과 본사 스타일 가이드를 보고 슬라이드 계획(JSON)을 만듭니다. 문체 통일(개조식 등), 분량 초과 시 분할, 표 전환, 레이아웃 선택을 맡습니다. 결과는 Pydantic 스키마(structured outputs)로 검증되고, 잘못된 레이아웃 이름이나 이미지 id 는 자동 보정됩니다.
- **[3] 렌더링** (`render.py`): 템플릿 파일의 예시 슬라이드를 지우고, 계획대로 레이아웃을 골라 제목/부제/본문/표/이미지를 배치합니다. 글꼴·색·글머리 모양은 템플릿 placeholder 서식을 그대로 따릅니다.

## 설치

```bash
pip install -e ".[dev]"
```

Claude 재구성을 쓰려면 Anthropic API 키가 필요합니다.

```bash
export ANTHROPIC_API_KEY=sk-ant-...
```

## 빠른 시작 (예시 파일)

```bash
python scripts/make_samples.py
```

```bash
python -m slide_converter input/sample_input.pptx -c config/company.example.toml
```

결과: `output/sample_input_converted.pptx`

## 본사 템플릿 연결하기

1. 본사 템플릿(.pptx 또는 .potx)을 `templates/` 에 넣습니다.
2. 레이아웃 이름을 확인합니다.
   ```bash
   python -m slide_converter --inspect-template templates/본사템플릿.potx
   ```
3. `config/company.example.toml` 을 `config/company.toml` 로 복사하고 다음을 채웁니다.
   - `[template] path` : 템플릿 경로
   - `[style] guide` : 본사 문서 작성 규칙 (문체, 분량, 표기법 등) — **변환 품질을 가장 크게 좌우합니다**
   - `[layouts.descriptions]` : 각 레이아웃의 용도 설명 (Claude 가 레이아웃을 고를 때 참고)
   - `[layouts.defaults]` : AI 미사용 변환·보정 시 쓸 기본 레이아웃

## 사용법

```bash
# Claude 로 재구성 (기본)
python -m slide_converter input/보고서.pptx -c config/company.toml -o output/보고서_본사양식.pptx

# 계획을 JSON 으로 저장 → 사람이 검토/수정 → 그 계획으로 다시 렌더링 (API 재호출 없음)
python -m slide_converter input/보고서.pptx -c config/company.toml --dump-plan output/plan.json
python -m slide_converter input/보고서.pptx -c config/company.toml --from-plan output/plan.json

# AI 없이 서식만 1:1 교체
python -m slide_converter input/보고서.pptx -c config/company.toml --no-ai
```

| 옵션 | 설명 |
|---|---|
| `-c, --config` | 회사 설정 TOML |
| `-t, --template` | 템플릿 경로 (설정 파일보다 우선) |
| `-o, --output` | 출력 경로 (기본 `output/<원본>_converted.pptx`) |
| `--model` | Claude 모델 ID (기본 `claude-opus-5-5`) |
| `--no-ai` | Claude 없이 단순 변환 |
| `--dump-plan` / `--from-plan` | 변환 계획 저장 / 재사용 |
| `--inspect-template` | 템플릿 레이아웃 목록 출력 |

Claude 호출은 adaptive thinking 과 server-side fallback(`fallbacks: "default"`)을 사용합니다. 안전 분류기가 요청을 거절하면 서버가 대체 모델로 자동 재시도하고, 그래도 거절되면 오류로 종료합니다.

## 현재 한계

- 차트, SmartArt, 도형으로 그린 다이어그램은 옮기지 않고 발표자 노트에 표시만 합니다.
- 텍스트 넘침은 PowerPoint 의 '도형에 맞춤' 자동 맞춤 설정에 맡깁니다 (LibreOffice 미리보기에서는 다르게 보일 수 있음).
- 원본이 매우 길면(수백 장) 한 번의 호출 출력 한도를 넘을 수 있으니 나눠서 변환하세요.

## 테스트

```bash
python -m pytest -q
```

Claude 호출은 가짜 클라이언트로 대체해 API 키 없이 실행됩니다.

## 구조

```
slide_converter/
  cli.py        명령행 진입점
  extract.py    원본 .pptx 내용 추출
  template.py   템플릿 레이아웃 카탈로그
  planner.py    Claude 재구성 / 단순 변환 / 계획 보정
  render.py     템플릿 위에 렌더링
  models.py     데이터 모델 (Pydantic 스키마 포함)
  config.py     TOML 설정 로딩
  pptx_io.py    .potx 열기, 예시 슬라이드 제거
config/company.example.toml
scripts/make_samples.py   예시 템플릿·원본 생성
tests/
```
