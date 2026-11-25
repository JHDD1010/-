"""본사 템플릿의 레이아웃 카탈로그와 placeholder 역할 판별.

MB 템플릿처럼 부제·각주·로고 자리를 모두 BODY placeholder 로 만든 경우가 있어,
유형만 보지 않고 위치·크기로 역할을 정한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Cm

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE}
SUBTITLE_TYPES = {PP_PLACEHOLDER.SUBTITLE}
FOOTER_TYPES = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.SLIDE_NUMBER}
BODY_TYPES = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.VERTICAL_BODY}
CONTENT_TYPES = {
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.PICTURE,
    PP_PLACEHOLDER.TABLE,
    PP_PLACEHOLDER.CHART,
    PP_PLACEHOLDER.VERTICAL_BODY,
    PP_PLACEHOLDER.VERTICAL_OBJECT,
}

# 역할 이름
TITLE, SUBTITLE, SUBHEADLINE, FOOTNOTE, DECOR, FOOTER, CONTENT, PICTURE = (
    "title", "subtitle", "subheadline", "footnote", "decor", "footer", "content", "picture",
)


def placeholder_role(ph, slide_height: int, title_ph=None) -> str:
    """placeholder 하나의 역할을 돌려준다 (레이아웃·슬라이드 양쪽에서 사용)."""
    t = ph.placeholder_format.type
    if t in TITLE_TYPES:
        return TITLE
    if t in SUBTITLE_TYPES:
        return SUBTITLE
    if t in FOOTER_TYPES:
        return FOOTER
    if t == PP_PLACEHOLDER.PICTURE:
        return PICTURE
    if t in BODY_TYPES:
        top, h, w = ph.top or 0, ph.height or 0, ph.width or 0
        # 아주 작은 BODY: 로고/워드마크 자리
        if h < Cm(0.5) or (w < Cm(2) and h < Cm(2)):
            return DECOR
        # 슬라이드 아래쪽의 낮은 BODY: 각주
        if h <= Cm(0.9) and top > slide_height * 0.8:
            return FOOTNOTE
        # 제목 바로 아래의 낮은 BODY: 부제(Subheadline)
        if title_ph is not None and h <= Cm(2.0) and top < slide_height * 0.35:
            return SUBHEADLINE
    if t in CONTENT_TYPES:
        return CONTENT
    return DECOR


def roles_of(placeholders, slide_height: int) -> list[tuple]:
    """[(placeholder, role)]. 왼쪽→위 순서로 정렬된다."""
    phs = sorted(placeholders, key=lambda p: ((p.left or 0), (p.top or 0)))
    title_ph = next((p for p in phs if p.placeholder_format.type in TITLE_TYPES), None)
    return [(p, placeholder_role(p, slide_height, title_ph)) for p in phs]


@dataclass
class LayoutInfo:
    name: str
    has_title: bool
    has_subtitle: bool  # SUBTITLE 또는 Subheadline 칸
    has_footnote: bool = False
    content_areas: list[str] = field(default_factory=list)  # 예: ["OBJECT", "PICTURE"]
    description: str = ""

    def to_prompt_dict(self) -> dict:
        d = {
            "name": self.name,
            "title": self.has_title,
            "subtitle": self.has_subtitle,
            "footnote": self.has_footnote,
            "content_areas": self.content_areas or "없음 (본문은 기본 본문 영역에 자유 배치)",
        }
        if self.description:
            d["description"] = self.description
        return d


def build_catalog(
    prs,
    descriptions: dict[str, str] | None = None,
    allowed: list[str] | None = None,
) -> list[LayoutInfo]:
    descriptions = descriptions or {}
    catalog = []
    seen = set()
    for layout in prs.slide_layouts:
        if allowed and layout.name not in allowed:
            continue
        if layout.name in seen:  # 같은 이름 레이아웃은 첫 번째만
            continue
        seen.add(layout.name)
        info = LayoutInfo(name=layout.name, has_title=False, has_subtitle=False)
        for ph, role in roles_of(layout.placeholders, prs.slide_height):
            if role == TITLE:
                info.has_title = True
            elif role in (SUBTITLE, SUBHEADLINE):
                info.has_subtitle = True
            elif role == FOOTNOTE:
                info.has_footnote = True
            elif role in (CONTENT, PICTURE):
                t = ph.placeholder_format.type
                info.content_areas.append(t.name if hasattr(t, "name") else str(t))
        info.description = descriptions.get(layout.name, "")
        catalog.append(info)
    if allowed:
        missing = [n for n in allowed if n not in seen]
        if missing:
            raise ValueError(f"템플릿에 없는 레이아웃 이름: {missing}")
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
    if kind == "two_content" and "content" in preferred and any(l.name == preferred["content"] for l in catalog):
        return preferred["content"]

    def score(l: LayoutInfo) -> int:
        n = len(l.content_areas)
        if kind == "title":
            return (3 if l.has_subtitle else 0) + (2 if l.has_title else 0) - n
        if kind == "section":
            return (2 if l.has_title else 0) - 2 * n + (1 if "section" in l.name.lower() else 0)
        if kind == "two_content":
            return (2 if l.has_title else 0) + (3 if n == 2 else -abs(n - 2))
        return (2 if l.has_title else 0) + (3 if n == 1 else -abs(n - 1)) - (1 if l.has_subtitle else 0)

    return max(catalog, key=score).name
