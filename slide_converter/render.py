"""DeckPlan 을 본사 템플릿 위에 그려 .pptx 로 저장한다.

글꼴·색상·글머리 모양은 템플릿 레이아웃의 placeholder 서식을 그대로 상속하므로
여기서는 '어디에 무엇을 넣을지'만 결정한다.
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt

from .config import CompanyConfig
from .models import Bullet, DeckPlan, PlannedSlide, SourceImage, SourceSlide, Table
from .pptx_io import open_presentation, remove_all_slides
from .template import CONTENT_TYPES, SUBTITLE_TYPES, TITLE_TYPES, find_layout

_TEXT_OK = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.VERTICAL_BODY, PP_PLACEHOLDER.VERTICAL_OBJECT}
_SUITS = {
    "bullets": _TEXT_OK,
    "table": {PP_PLACEHOLDER.TABLE, PP_PLACEHOLDER.OBJECT},
    "image": {PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.OBJECT},
}


@dataclass
class Box:
    left: int
    top: int
    width: int
    height: int

    def split(self, n: int, horizontal: bool, gap: int) -> list["Box"]:
        if n <= 1:
            return [self]
        if horizontal:
            w = (self.width - gap * (n - 1)) // n
            return [Box(self.left + i * (w + gap), self.top, w, self.height) for i in range(n)]
        h = (self.height - gap * (n - 1)) // n
        return [Box(self.left, self.top + i * (h + gap), self.width, h) for i in range(n)]


def _box_of(shape) -> Box:
    return Box(shape.left, shape.top, shape.width, shape.height)


def _remove(shape) -> None:
    el = shape._element
    el.getparent().remove(el)


# ---------------------------------------------------------------------------
# 요소별 그리기
# ---------------------------------------------------------------------------


def _fill_text_frame(tf, bullets: list[Bullet], lead: str | None = None) -> None:
    tf.clear()
    tf.word_wrap = True
    lines: list[tuple[str, int, bool]] = []
    if lead:
        lines.append((lead, 0, True))
    lines += [(b.text, b.level, False) for b in bullets]
    for i, (text, level, bold) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = text
        p.level = level
        if bold:
            run.font.bold = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _draw_bullets(slide, ph, box: Box, bullets: list[Bullet], lead: str | None) -> None:
    if ph is not None and ph.placeholder_format.type in _TEXT_OK:
        if _box_of(ph) != box:
            ph.left, ph.top, ph.width, ph.height = box.left, box.top, box.width, box.height
        _fill_text_frame(ph.text_frame, bullets, lead)
        return
    if ph is not None:
        _remove(ph)
    tb = slide.shapes.add_textbox(box.left, box.top, box.width, box.height)
    _fill_text_frame(tb.text_frame, bullets, lead)


def _draw_table(slide, ph, box: Box, table: Table, font_size: int) -> None:
    if ph is not None:
        _remove(ph)
    data = [table.header] + table.rows if table.header else table.rows
    if not data:
        return
    n_cols = max(len(r) for r in data)
    row_h = Pt(font_size * 2.2)
    height = min(box.height, int(row_h * len(data)))
    shape = slide.shapes.add_table(len(data), n_cols, box.left, box.top, box.width, height)
    tbl = shape.table
    tbl.first_row = bool(table.header)
    for r, row in enumerate(data):
        for c in range(n_cols):
            cell = tbl.cell(r, c)
            cell.text = row[c] if c < len(row) else ""
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(font_size)


def _draw_image(slide, ph, box: Box, image: SourceImage) -> None:
    if ph is not None:
        _remove(ph)
    ratio = image.width / image.height if image.height else 1.0
    w, h = box.width, int(box.width / ratio)
    if h > box.height:
        h, w = box.height, int(box.height * ratio)
    left = box.left + (box.width - w) // 2
    top = box.top + (box.height - h) // 2
    pic = slide.shapes.add_picture(io.BytesIO(image.blob), left, top, w, h)
    if image.description:
        pic._element._nvXxPr.cNvPr.set("descr", image.description)


# ---------------------------------------------------------------------------
# 슬라이드 단위
# ---------------------------------------------------------------------------


def _default_area(prs) -> Box:
    W, H = prs.slide_width, prs.slide_height
    return Box(int(W * 0.06), int(H * 0.2), int(W * 0.88), int(H * 0.72))


def _render_slide(prs, ps: PlannedSlide, images: dict[str, SourceImage], cfg: CompanyConfig) -> None:
    layout = find_layout(prs, ps.layout)
    slide = prs.slides.add_slide(layout)
    phs = list(slide.placeholders)

    title_ph = next((p for p in phs if p.placeholder_format.type in TITLE_TYPES), None)
    sub_ph = next((p for p in phs if p.placeholder_format.type in SUBTITLE_TYPES), None)
    regions = sorted(
        (p for p in phs if p.placeholder_format.type in CONTENT_TYPES),
        key=lambda p: (p.left, p.top),
    )

    # 제목 / 부제
    if title_ph is not None:
        if ps.title:
            title_ph.text_frame.text = ps.title
        else:
            _remove(title_ph)
    lead = None
    if sub_ph is not None:
        if ps.subtitle:
            sub_ph.text_frame.text = ps.subtitle
        else:
            _remove(sub_ph)
    elif ps.subtitle:
        lead = ps.subtitle  # 부제 영역이 없으면 본문 첫 줄(굵게)로

    # 배치할 항목
    items: list[tuple[str, object]] = []
    if ps.bullets or lead:
        items.append(("bullets", ps.bullets))
    if ps.table:
        items.append(("table", ps.table))
    items += [("image", images[i]) for i in ps.image_ids if i in images]

    if not items:
        for r in regions:
            _remove(r)
        _set_notes(slide, ps.notes)
        return

    # 항목 ↔ 영역 매칭
    free = list(regions)
    assigned: list[tuple[str, object, object, Box]] = []  # kind, payload, placeholder, box

    def take_region(kind: str):
        for r in free:
            if r.placeholder_format.type in _SUITS[kind]:
                free.remove(r)
                return r
        return free.pop(0) if free else None

    direct = items if len(items) <= len(regions) else items[: max(len(regions) - 1, 0)]
    overflow = items[len(direct):]

    for kind, payload in direct:
        ph = take_region(kind)
        assigned.append((kind, payload, ph, _box_of(ph)))

    if overflow:
        host = free.pop(0) if free else None
        area = _box_of(host) if host is not None else _default_area(prs)
        horizontal = not any(k == "table" for k, _ in overflow) and area.width > area.height
        boxes = area.split(len(overflow), horizontal, gap=int(prs.slide_width * 0.015))
        for idx, ((kind, payload), box) in enumerate(zip(overflow, boxes)):
            # host placeholder 는 글머리가 있으면 글머리에, 없으면 첫 항목에서 제거된다
            ph = host if (host is not None and (kind == "bullets" or idx == 0)) else None
            if ph is host:
                host = None
            assigned.append((kind, payload, ph, box))
        if host is not None:
            _remove(host)

    for kind, payload, ph, box in assigned:
        if kind == "bullets":
            _draw_bullets(slide, ph, box, payload, lead)
        elif kind == "table":
            _draw_table(slide, ph, box, payload, cfg.table_font_size)
        else:
            _draw_image(slide, ph, box, payload)

    for r in free:
        _remove(r)
    _set_notes(slide, ps.notes)


def _set_notes(slide, notes: str) -> None:
    if notes:
        slide.notes_slide.notes_text_frame.text = notes


def render_deck(
    plan: DeckPlan,
    source_slides: list[SourceSlide],
    template_path: str | Path,
    output_path: str | Path,
    cfg: CompanyConfig,
) -> Path:
    prs = open_presentation(template_path)
    remove_all_slides(prs)
    images = {im.id: im for s in source_slides for im in s.images}
    for ps in plan.slides:
        _render_slide(prs, ps, images, cfg)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path
