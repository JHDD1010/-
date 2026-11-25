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
    layout_descriptions: dict[str, str] = field(default_factory=dict)
    # 단순 변환(--no-ai) 및 보정용 기본 레이아웃: title / section / content / two_content
    preferred_layouts: dict[str, str] = field(default_factory=dict)
    add_cover: bool = True
    cover_title: str = ""
    cover_subtitle: str = ""
    table_font_size: int = 12
    max_bullets: int = 6


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
    cfg.max_bullets = style.get("max_bullets", cfg.max_bullets)

    cover = data.get("cover", {})
    cfg.add_cover = cover.get("enabled", cfg.add_cover)
    cfg.cover_title = cover.get("title", "")
    cfg.cover_subtitle = cover.get("subtitle", "")

    cfg.layout_descriptions = dict(data.get("layouts", {}).get("descriptions", {}))
    cfg.preferred_layouts = dict(data.get("layouts", {}).get("defaults", {}))
    return cfg
