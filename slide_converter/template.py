"""본사 템플릿의 레이아웃 카탈로그를 만든다."""

from __future__ import annotations

from dataclasses import dataclass, field

from pptx.enum.shapes import PP_PLACEHOLDER

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE}
SUBTITLE_TYPES = {PP_PLACEHOLDER.SUBTITLE}
CONTENT_TYPES = {
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.PICTURE,
    PP_PLACEHOLDER.TABLE,
    PP_PLACEHOLDER.CHART,
    PP_PLACEHOLDER.VERTICAL_BODY,
    PP_PLACEHOLDER.VERTICAL_OBJECT,
}


@dataclass
class LayoutInfo:
    name: str
    has_title: bool
    has_subtitle: bool
    content_areas: list[str] = field(default_factory=list)  # 예: ["BODY", "PICTURE"]
    description: str = ""

    def to_prompt_dict(self) -> dict:
        d = {
            "name": self.name,
            "title": self.has_title,
            "subtitle": self.has_subtitle,
            "content_areas": self.content_areas,
        }
        if self.description:
            d["description"] = self.description
        return d


def build_catalog(prs, descriptions: dict[str, str] | None = None) -> list[LayoutInfo]:
    descriptions = descriptions or {}
    catalog = []
    for layout in prs.slide_layouts:
        info = LayoutInfo(name=layout.name, has_title=False, has_subtitle=False)
        for ph in sorted(layout.placeholders, key=lambda p: (p.left or 0, p.top or 0)):
            t = ph.placeholder_format.type
            if t in TITLE_TYPES:
                info.has_title = True
            elif t in SUBTITLE_TYPES:
                info.has_subtitle = True
            elif t in CONTENT_TYPES:
                info.content_areas.append(t.name if hasattr(t, "name") else str(t))
        info.description = descriptions.get(layout.name, "")
        catalog.append(info)
    return catalog


def find_layout(prs, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return None


def pick_layout_name(catalog: list[LayoutInfo], *, kind: str, preferred: dict[str, str] | None = None) -> str:
    """kind: 'title' | 'section' | 'content' | 'two_content' 에 맞는 레이아웃을 고른다."""
    preferred = preferred or {}
    if kind in preferred and any(l.name == preferred[kind] for l in catalog):
        return preferred[kind]

    def score(l: LayoutInfo) -> int:
        n = len(l.content_areas)
        if kind == "title":
            return (3 if l.has_subtitle else 0) + (2 if l.has_title else 0) - n
        if kind == "section":
            return (2 if l.has_title else 0) - 2 * n + (1 if "section" in l.name.lower() else 0)
        if kind == "two_content":
            return (2 if l.has_title else 0) + (3 if n == 2 else -abs(n - 2))
        # content
        return (2 if l.has_title else 0) + (3 if n == 1 else -abs(n - 1)) - (1 if l.has_subtitle else 0)

    return max(catalog, key=score).name
