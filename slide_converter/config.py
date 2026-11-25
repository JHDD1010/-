"""회사 템플릿 설정(TOML) 로딩."""

from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from pathlib import Path

DEFAULT_STYLE_GUIDE = """\
- 문체: 개조식, 명사형 종결(예: '~함', '~임', '~ 추진'). 존댓말 문장체 금지.
- 슬라이드 제목은 20자 이내, 핵심 메시지를 담을 것.
- 글머리는 슬라이드당 최대 6개, 한 줄 40자 이내. 넘치면 슬라이드를 나눌 것.
- 숫자·단위·고유명사는 원문 그대로 유지하고, 원문에 없는 사실을 만들지 말 것.
- 표로 정리하는 것이 더 명확한 비교·수치 데이터는 표로 바꿀 것.
"""


@dataclass
class CompanyConfig:
    template_path: Path | None = None
    model: str = "claude-opus-5-5"
    effort: str = "high"
    style_guide: str = DEFAULT_STYLE_GUIDE

    # 레이아웃
    layout_descriptions: dict[str, str] = field(default_factory=dict)
    allowed_layouts: list[str] = field(default_factory=list)  # 비우면 전부
    # 단순 변환(--no-ai) 및 보정용 기본 레이아웃: title / section / content / two_content
    preferred_layouts: dict[str, str] = field(default_factory=dict)
    # content 칸이 없는 레이아웃에서 본문을 놓을 영역 [left, top, width, height] (cm)
    content_area_cm: list[float] | None = None

    # 표지
    add_cover: bool = True
    cover_title: str = ""
    cover_subtitle: str = ""

    # 서식
    title_style: str = "layout"  # "layout" = 레이아웃 그대로 / "left_accent" = 왼쪽 정렬 + 강조 바
    title_size: float | None = None  # pt, None 이면 레이아웃 값
    subheadline_size: float | None = None
    body_sizes: list[float] = field(default_factory=list)  # 글머리 수준별 pt, 비우면 레이아웃 값
    table_font_size: int = 12
    table_style: str = "default"  # "default" = PowerPoint 기본 표 스타일 / "lines" = 가로 구분선만
    max_bullets: int = 6

    # 반복 요소
    footer_text: str = ""  # 오른쪽 아래 바닥글
    tagline: str = ""  # 왼쪽 아래 (날짜 placeholder 자리)
    slide_numbers: bool = True
    classification: str = ""  # 예: "INTERNAL", "CONFIDENTIAL" — 오른쪽 위 라벨


def load_config(path: str | Path | None) -> CompanyConfig:
    cfg = CompanyConfig()
    if not path:
        return cfg
    path = Path(path)
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    base = path.parent

    tpl = data.get("template", {})
    if "path" in tpl:
        p = Path(tpl["path"])
        cfg.template_path = p if p.is_absolute() else (base / p).resolve()

    ai = data.get("ai", {})
    cfg.model = ai.get("model", cfg.model)
    cfg.effort = ai.get("effort", cfg.effort)

    style = data.get("style", {})
    cfg.style_guide = style.get("guide", cfg.style_guide)
    cfg.table_font_size = style.get("table_font_size", cfg.table_font_size)
    cfg.table_style = style.get("table_style", cfg.table_style)
    cfg.max_bullets = style.get("max_bullets", cfg.max_bullets)
    cfg.title_style = style.get("title_style", cfg.title_style)
    cfg.title_size = style.get("title_size", cfg.title_size)
    cfg.subheadline_size = style.get("subheadline_size", cfg.subheadline_size)
    cfg.body_sizes = list(style.get("body_sizes", cfg.body_sizes))

    cover = data.get("cover", {})
    cfg.add_cover = cover.get("enabled", cfg.add_cover)
    cfg.cover_title = cover.get("title", "")
    cfg.cover_subtitle = cover.get("subtitle", "")

    deck = data.get("deck", {})
    cfg.footer_text = deck.get("footer", cfg.footer_text)
    cfg.tagline = deck.get("tagline", cfg.tagline)
    cfg.slide_numbers = deck.get("slide_numbers", cfg.slide_numbers)
    cfg.classification = deck.get("classification", cfg.classification)

    layouts = data.get("layouts", {})
    cfg.layout_descriptions = dict(layouts.get("descriptions", {}))
    cfg.preferred_layouts = dict(layouts.get("defaults", {}))
    cfg.allowed_layouts = list(layouts.get("allowed", []))
    cfg.content_area_cm = layouts.get("content_area_cm")
    return cfg
