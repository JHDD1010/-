# 본사 템플릿 분석 (Mercedes-Benz 코퍼레이트 마스터)

분석 대상 (2026-11-25 업로드)

| 파일 | 장수 | 마스터 | 비고 |
|---|---|---|---|
| A. `Korea_input_2026_10_Academia_Landscape_Overview…pptx` | 1 | 1개, 레이아웃 71개 | **라이트(흰 배경) 마스터 전체 세트** |
| B. `261029_StarSeeker_InternationalKickoff_…first_half.pptx` | 18 | 6개 (Office Theme 1 + MB 5) | 라이트·다크 혼용, 복사·붙여넣기로 마스터 중복 |
| C. `second_half.pptx` | 15 | 2개 (MB 다크 2) | **다크 'Pioneering Next' 마스터** |

세 파일 모두 같은 MB 코퍼레이트 템플릿(테마 `Mercedes Benz`, 색 구성 `MB 2025-05-21`)에서 나왔고, 마스터 텍스트 스타일·그리드·레이아웃 이름이 같습니다. 다른 점은 **라이트/다크 색 구성**과, 다크 마스터에서 일부 레이아웃이 빠진 것뿐입니다.

---

## 1. 공통 기본값

| 항목 | 값 |
|---|---|
| 슬라이드 크기 | 33.87 × 19.05 cm (16:9, 13.333 × 7.5 in) |
| 제목 폰트 (major) | **MB Corpo A Title Cond Office** |
| 본문 폰트 (minor) | **MB Corpo S Text Office Light** |
| 보조 폰트 (슬라이드에서 직접 지정) | MB Corpo S Text Office (Regular: 라벨·카드 제목), CorpoS (구버전) |
| 한글(EA/Hang) 폰트 | **테마에 지정 없음** → 한글은 PC 기본 글꼴(보통 맑은 고딕)로 대체 |
| 폰트 임베드 | 없음. MB Corpo 폰트가 없는 PC에서는 대체 폰트로 보임 |
| 레이아웃 안내 문구 | 독일어 (Titel, Headline, Subheadline, Kapitel, Fußnote, „Bild durch Klicken…") |
| 서드파티 | 모든 마스터에 `think-cell data - do not delete` OLE 개체 (TCLayout). 슬라이드에도 think-cell 개체가 있음 (A 38개, B 66개, C 16개) |

### 마스터 텍스트 스타일 (라이트·다크 동일)

| 수준 | 크기 | 폰트 | 글머리 | 왼쪽 여백 / 내어쓰기 |
|---|---|---|---|---|
| 제목 | 36 pt | major (Title Cond) | – | – |
| 본문 1~4단계 | 18 pt (모든 단계 동일) | minor (Light) | • | 0.51 / 1.00 / 1.50 / 1.99 cm, 내어쓰기 약 −0.5 cm |
| 단락 앞 간격 | 0 pt | | | |
| 날짜 / 바닥글 | 8 pt, **대문자, 자간 1.5 pt** | minor | | 날짜 왼쪽 아래, 바닥글 오른쪽 정렬 |
| 슬라이드 번호 | `— # —` 가운데 아래 | | | |

날짜 placeholder는 날짜 대신 **태그라인**으로 씁니다 (`PIONEERING NEXT`, `HEART OF THE FUTURE`).
바닥글은 `<프로젝트명> | <월 연도 또는 날짜>` 형식입니다. 예: `STARSEEKER INTERNATIONAL KICKOFF | NOVEMBER 2026`, `RD INNO STEERING PROCESS | 09.10.2026`, `PARTNERSHIP@RD STATUS / OCTOBER 16, 2026`.

### 그리드 (cm, 마스터·레이아웃 placeholder 기준)

```
  0 ─ 1.53 ──────────────────────────────────── 32.34 ─ 33.87   (좌우 여백 1.53)
 0.5  [Mercedes-Benz 워드마크 / "Mercedes-Benz Pioneering Next"]   (가운데 위)
 1.74 ┌ 제목 (h 2.9) ───────────────────────────────┐
 3.44 │ 부제(Subheadline, h 1.2~1.55)               │
 5.42 ├ 본문 영역 (h 12.1 → 17.52) ──────────────────┤
16.93 │ 각주 Fußnote (h 0.6, 8 pt, 아래 정렬)        │
17.83 └ 태그라인(왼쪽) · 쪽번호(가운데) · 바닥글(오른쪽) ┘  (h 0.5, 8 pt)
```

- 2단: 폭 15.1 cm × 2, 간격 0.6 (x = 1.53 / 17.23)
- 3단: 폭 8.8 cm × 3, 간격 1.1 (x = 2.63 / 12.53 / 22.43), 좌우 1.1 더 들여 씀
- 4단: 폭 6.6 cm × 4 (x = 2.63 / 9.97 / 17.30 / 24.63)
- 전면 사진 레이아웃은 0,0 부터 33.87 × 19.05 전체

---

## 2. 색 구성 — 라이트 vs 다크

### 라이트 (A, B의 M4~M6) — 흰 배경, 상단 `Mercedes-Benz` 워드마크, 태그라인 `HEART OF THE FUTURE`

| 슬롯 | 값 | 용도 |
|---|---|---|
| dk1 / lt1 | `000000` / `FFFFFF` | 본문 검정 / 배경 흰색 |
| dk2 / lt2 | `8D8D8D` / `606060` | 회색 |
| accent1 / 2 | `0078D6` / `1398FF` | 파랑 |
| accent3 / 4 | `075959` / `038281` | 청록 |
| accent5 / 6 | `522099` / `7B30E5` | 보라 |

### 다크 (C, B의 M2~M3) — 검정 배경, 상단 `Mercedes-Benz Pioneering Next`, 태그라인 `PIONEERING NEXT`

**dk1과 lt1이 뒤집혀 있습니다.** 그래서 `tx1`(글자)이 밝은색, `bg1`(배경)이 어두운색이 됩니다. 변환기에서 색을 RGB로 직접 넣지 말고 테마 색(`tx1`, `accent1` 등)으로 지정해야 라이트/다크 어느 쪽에서도 맞게 나옵니다.

| 슬롯 | 값 | 용도 |
|---|---|---|
| dk1 (=tx1) | `E7EDF0` | 본문 글자 (밝은 회백색) |
| lt1 (=bg1) | `0A0E12` | 배경 (거의 검정) |
| dk2 / lt2 | `B8C2C8` | 보조 글자 (부제, 설명) |
| accent1 / 4 | `00A6B8` | 청록 (테두리) |
| accent2 | `00D1E0` | **강조 시안** (강조 바, 현재 아젠다 항목, 번호) |
| accent3 | `121A20` | 카드 배경 |
| accent5 | `9AA5AD` | 흐린 회색 |
| accent6 | `16212A` | 짙은 패널 |

슬라이드에서 실제로 많이 쓴 색: `FFFFFF`, `E7EDF0`, `B8C2C8`, `00D1E0`, `01C9D7`(시안 변형), `9FB3BC`.

---

## 3. 레이아웃 카탈로그

라이트 마스터(A)에 71개가 있습니다. 같은 구성의 `1_` 접두 사본이 18개 있으니 실제로는 약 53종입니다. 다크 마스터(C)에는 그중 17~18개만 있습니다.

| 분류 | 레이아웃 | 핵심 서식 |
|---|---|---|
| **표지** | Title, Subheadline (White/Black Background) · Title, Subheadline, Picture (White/Black) · Title, Picture (White/Black) · Title, Subheadline, Full-Page Picture (Bright/Dark) 01·02 | 제목 **64 pt, 모두 대문자, 자간 6.4 pt, 가운데, 아래 정렬** · 부제 20 pt 가운데 · 사진 placeholder 30.81 × 8.8 |
| **아젠다** | Agenda Centre-Aligned · Agenda Left-Aligned · Agenda, Picture · Agenda, 2 Pictures | 본문 36 pt major 폰트, 가운데. 실제 슬라이드에서는 **표**(시간 / 항목, 행 구분선)로 만들고 현재 항목만 시안 `00D1E0`으로 강조 |
| **간지(Chapter)** | Chapter (White/Black Background) · Chapter, Full-Page Picture (Bright/Dark) | 48 pt, 모두 대문자, 자간 7 pt, 가운데 |
| **제목만** | Headline Only 01 (제목 가운데 정렬, 위) · 02 (제목이 전체 영역) · 03 (왼쪽 절반) | 제목 36 pt, 각주 8 pt |
| **제목 + 부제** | Headline, Subheadline (부제 24 pt) · **Headline, Subheadline (Smaller Text)** (제목 36 pt / 다크는 24 pt, 부제 12 pt, 가운데) | B·C에서 가장 많이 쓴 본문 레이아웃 |
| **본문** | Headline, Subheadline, Content · Headline, Subheadline, 2-Column-Content · Headline, Subheadline, 3-Column-Content (단마다 소제목) · Headline, Content 01/02/03 (제목 왼쪽 + 본문 오른쪽 등) · Headline, Content, 4 Icons | 본문은 OBJECT placeholder (텍스트·표·차트 모두 가능) |
| **표·차트** | Headline, Content (TABLE placeholder 30.81 × 11.5) · Headline, Diagram (제목 왼쪽 + CHART) | |
| **사진** | Headline, Picture · Headline, Content, Picture · Headline, 2/3/5 Pictures · Headline, 3-Column Content, 3 Pictures · Headline, 4-Column Content, 4 Pictures 01·02 · 2 Columns: Headline & 1/3 Pictures Each · Headline/Subheadline, Full-Page Picture (Bright/Dark) · 2 Headlines, 2 Pictures (Dark) · 4 Headlines, 4 Pictures (Dark) | |
| **팀** | Team Slide, 4 Pictures · Team Slide, 9 Pictures | 원형 사진 3.09 cm + 이름 칸 |
| **기타** | Blank Slide · DEFAULT · Title with Content (다크 M1) | |

### 다크 마스터의 레이아웃 이름 주의

다크 테마는 dk1/lt1 이 뒤집혀 있어서 **이름과 실제 배경이 반대**로 나옵니다.
`Title, Subheadline (White Background)` 가 어두운 배경, `(Black Background)` 가 밝은 회색 배경입니다.
다크 덱 표지에는 `(White Background)` 레이아웃을 씁니다 (`config/mb_dark.toml` 에 반영).

### placeholder 역할 주의 (변환기에 직접 영향)

- **부제(Subheadline)는 SUBTITLE이 아니라 BODY placeholder**입니다 (제목 바로 아래, 높이 1.2~1.55 cm). 표지에서만 SUBTITLE 유형을 씁니다.
- **각주(Fußnote)도 BODY placeholder**입니다 (y 16.93, 높이 0.6 cm, 8 pt).
- 전면 사진 레이아웃에는 **로고·워드마크 자리용 작은 BODY placeholder**(3.04 × 0.35, 1.4 × 1.4)가 있습니다.
- 본문은 BODY가 아니라 **OBJECT** placeholder입니다.

---

## 4. 실제 슬라이드에서 쓰는 패턴 (B·C)

| 요소 | 서식 |
|---|---|
| 사용한 레이아웃 | Title, Subheadline, Picture (Black) → 표지 / **Agenda Centre-Aligned** → 아젠다 표 / **Headline, Subheadline (Smaller Text)** · Headline Only 01·02 → 본문 / Chapter (Black) → 간지 |
| 제목 (레이아웃 방식) | 제목 placeholder, 가운데 정렬, 32 pt (다크) / 36 pt (라이트) |
| 제목 (새 방식, B 5·6·7, C 13~15) | placeholder 대신 `Title` 텍스트 상자, **왼쪽 정렬, 28~32 pt**, 위에 **`Accent Marker`**(시안 `00D1E0` 막대 2.92 × 0.14 cm, x 1.52, y 1.52). 그 아래 부제는 `B8C2C8` 12 pt |
| 카드 | 둥근 사각형 (모서리 4%) · 채움 `121A20` · 테두리 0.75 pt `00A6B8` · 번호 `01`, `02`는 40 pt major 폰트 시안 · 카드 라벨은 대문자 작은 글씨 시안 |
| 배경 이미지 | 전면 이미지 위에 `0A0E12` 80% 불투명 덮개(`Background Scrim`) |
| 보안 등급 라벨 | 오른쪽 위 (x 29.7~29.9, y 0.16~0.22) · `INTERNAL`: 둥근 사각형, accent1 글자, 12 pt, 자간 0.9 · `CONFIDENTIAL`: 채움 `741218`(진한 빨강), 흰 글자 대문자 12 pt, 자간 0.8 |
| 글자 크기 분포 | 본문 **12 pt**가 가장 많음 → 18 pt → 36 pt(제목) → 14·16 pt → 8~11 pt(표·주석). 마스터 기본 18 pt보다 실제 본문이 훨씬 작음 |
| 표 | 아젠다(시간 / 항목, 가로 구분선만), 역할 매트릭스, 요약표(머리행 시안 글자, 체크박스 □) |
| 차트·다이어그램 | think-cell, 그룹 도형, 연결선을 많이 씀 (B에 LINE 72개, GROUP 18개) |

---

## 5. 문제점 / 정리 필요 사항

1. **마스터 중복**: B에는 마스터가 6개(`Office Theme`, `1_`·`2_`·`4_`·`10_`·`15_Mercedes Benz`) 있습니다. 슬라이드를 다른 파일에서 복사해 오면서 생긴 것으로, 레이아웃 이름이 마스터마다 겹칩니다 (`Headline Only 01`이 여러 개). 변환기용으로는 **깨끗한 기준 템플릿 1개**가 필요합니다.
2. **제목 방식이 두 가지**: 레이아웃 제목 placeholder(가운데)와 텍스트 상자 + 시안 강조 바(왼쪽)가 섞여 있습니다. 어느 쪽이 표준인지 정해야 합니다.
3. **한글 폰트 미지정**: 테마의 EA/Hang 폰트가 비어 있어서 한글이 PC마다 다르게 보입니다.
4. **A 파일 본문이 비표준**: 템플릿은 MB인데 슬라이드 본문은 Calibri / Aptos Narrow와 보라 `522099`로 되어 있습니다 (한국 입력본으로 보이며, 바로 이 변환기가 고쳐야 할 대상의 예).
5. **think-cell 개체**: 마스터의 `think-cell data - do not delete`는 지우면 안 됩니다. 슬라이드의 think-cell 차트는 python-pptx로 옮길 수 없습니다.

---

## 6. 변환기에 반영할 사항

| # | 변경 | 이유 |
|---|---|---|
| 1 | 레이아웃 카탈로그에서 placeholder **역할 판별** (제목 바로 아래 작은 BODY = 부제, 하단 0.6 cm BODY = 각주, 아주 작은 BODY = 로고 자리 → 제외) | 지금 코드는 부제·각주·로고 자리를 모두 '본문 영역'으로 보고 글머리를 넣을 수 있음 |
| 2 | 새 슬라이드에 **바닥글·태그라인 채우기** (설정의 `footer`, `tagline`) | python-pptx로 추가한 슬라이드에는 날짜/바닥글 placeholder가 복사되지 않음 |
| 3 | 기준 템플릿 1개 지정 + 사용할 레이아웃 화이트리스트(약 10종)와 용도 설명을 설정 파일로 | 71개 전부 Claude에 주면 선택이 흔들림. `1_` 사본 제외 |
| 4 | 본문 크기 12~14 pt 기준, 글머리 수 제한을 스타일 가이드에 | 실제 사용 패턴 |
| 5 | 한글 EA 폰트 지정 옵션 (예: 맑은 고딕 또는 사내 지정 한글 폰트) | 테마에 Hang 미지정 |
| 6 | 색은 테마 슬롯(`tx1`, `accent2` …)으로 지정 | 라이트/다크 공용 |
| 7 | 보안 등급 라벨(`INTERNAL` / `CONFIDENTIAL`), 강조 바 같은 장식을 옵션으로 | 실사용 패턴 |
| 8 | think-cell·그룹 도형 다이어그램은 '[원본 확인 필요]'로 표시 (현재 동작 유지) | 변환 불가 |

→ 1~3, 5~7 은 반영 완료 (2026-11-25). 기준 템플릿: `scripts/build_base_template.py`, 설정: `config/mb_dark.toml`.

분석 스크립트: `python scripts/analyze_template.py <템플릿.pptx>` (JSON 출력)
