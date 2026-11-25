"""DeckPlan 을 본사 템플릿 위에 그려 .pptx 로 저장한다.

글꼴·색·글머리 모양은 템플릿 placeholder 서식을 상속하고, 색을 지정할 때는
RGB 대신 테마 색(tx1, tx2, accent2 …)을 써서 라이트/다크 템플릿 모두에서 맞게 보이게 한다.
"""

from __future__ import annotations

import copy
import io
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from .config import CompanyConfig
from .models import Bullet, DeckPlan, PlannedSlide, SourceImage, SourceSlide, Table
from .pptx_io import open_presentation, remove_all_slides
from .template import (
    CONTENT,
    DECOR,
    FOOTER,
    FOOTNOTE,
    PICTURE,
    SUBHEADLINE,
    SUBTITLE,
    TITLE,
    find_layout,
    roles_of,
)

_TEXT_OK = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.VERTICAL_BODY, PP_PLACEHOLDER.VERTICAL_OBJECT}
_SUITS = {
    "bullets": _TEXT_OK,
    "table": {PP_PLACEHOLDER.TABLE, PP_PLACEHOLDER.OBJECT},
    "image": {PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.OBJECT},
}
_NO_STYLE_TABLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "No Style, No Grid"

# MB 템플릿 실사용 값 (cm)
_ACCENT_BAR = (1.52, 1.52, 2.92, 0.14)
_LEFT_TITLE = (1.52, 2.03, 30.82, 3.0)
_LABEL = (29.74, 0.2, 3.77, 0.62)


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


def _cm_box(vals) -> Box:
    return Box(*(int(Cm(v)) for v in vals))


# ---------------------------------------------------------------------------
# 텍스트
# ---------------------------------------------------------------------------


def _bulletize(p, level: int) -> None:
    """placeholder 가 아닌 텍스트 상자에 마스터 본문과 같은 글머리(•, 0.5 cm 단위 들여쓰기)를 단다."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Cm(0.5 * (level + 1)))))
    pPr.set("indent", str(-int(Cm(0.5))))
    for child in list(pPr):
        if etree.QName(child).localname in ("buNone", "buChar", "buAutoNum"):
            pPr.remove(child)
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", "•")


def _fill_text_frame(tf, bullets: list[Bullet], lead: str | None, sizes: list[float], bulleted: bool = False) -> None:
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
        elif bulleted:
            _bulletize(p, level)
        if sizes:
            run.font.size = Pt(sizes[min(level, len(sizes) - 1)])
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _draw_bullets(slide, ph, box: Box, bullets, lead, cfg: CompanyConfig) -> None:
    if ph is not None and ph.placeholder_format.type in _TEXT_OK:
        if _box_of(ph) != box:
            ph.left, ph.top, ph.width, ph.height = box.left, box.top, box.width, box.height
        # BODY 칸(아젠다·간지 등)은 레이아웃 고유 크기를 유지하고, 일반 본문(OBJECT)에만 본문 크기 적용
        sizes = [] if ph.placeholder_format.type == PP_PLACEHOLDER.BODY else cfg.body_sizes
        _fill_text_frame(ph.text_frame, bullets, lead, sizes)
        return
    if ph is not None:
        _remove(ph)
    tb = slide.shapes.add_textbox(box.left, box.top, box.width, box.height)
    _fill_text_frame(tb.text_frame, bullets, lead, cfg.body_sizes, bulleted=True)


# ---------------------------------------------------------------------------
# 표
# ---------------------------------------------------------------------------


def _ln(tag: str, scheme: str | None, width: int = 9525):
    ln = etree.SubElement(etree.Element("dummy"), qn(f"a:{tag}"))
    if scheme is None:
        etree.SubElement(ln, qn("a:noFill"))
    else:
        ln.set("w", str(width))
        fill = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:schemeClr")).set("val", scheme)
    return ln


def _style_table_lines(shape, header: bool) -> None:
    """MB 다크 덱 스타일: 배경 없음, 가로 구분선(tx2)만, 머리행 글자 accent2."""
    tbl = shape.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "1" if header else "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = _NO_STYLE_TABLE

    n_rows = len(tbl.rows)
    for r, row in enumerate(tbl.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if etree.QName(child).localname in ("lnL", "lnR", "lnT", "lnB", "noFill", "solidFill", "gradFill"):
                    tcPr.remove(child)
            bottom = "tx2" if (r < n_rows - 1 or n_rows == 1) else None
            if header and r == 0:
                bottom = "accent2"
            for el in (_ln("lnL", None), _ln("lnR", None), _ln("lnT", None), _ln("lnB", bottom)):
                tcPr.append(el)
            tcPr.append(etree.Element(qn("a:noFill")))
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.color.theme_color = (
                        MSO_THEME_COLOR.ACCENT_2 if header and r == 0 else MSO_THEME_COLOR.TEXT_1
                    )


def _draw_table(slide, ph, box: Box, table: Table, cfg: CompanyConfig) -> None:
    if ph is not None:
        _remove(ph)
    data = [table.header] + table.rows if table.header else table.rows
    if not data:
        return
    n_cols = max(len(r) for r in data)
    row_h = Pt(cfg.table_font_size * 2.2)
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
                    run.font.size = Pt(cfg.table_font_size)
    if cfg.table_style == "lines":
        _style_table_lines(shape, bool(table.header))


# ---------------------------------------------------------------------------
# 이미지
# ---------------------------------------------------------------------------


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
# 제목 / 반복 요소
# ---------------------------------------------------------------------------


def _layout_font_pt(ph, default: float) -> float:
    """placeholder 가 상속받는 1단계 글자 크기 (슬라이드 → 레이아웃 순으로 찾음)."""
    try:
        layout_ph = ph.part.slide_layout.placeholders.get(idx=ph.placeholder_format.idx)
    except Exception:  # noqa: BLE001
        layout_ph = None
    for src in (ph, layout_ph):
        if src is None:
            continue
        r = src._element.find(".//" + qn("a:lstStyle") + "/" + qn("a:lvl1pPr") + "/" + qn("a:defRPr"))
        if r is not None and r.get("sz"):
            return int(r.get("sz")) / 100
    return default


def _fit_size(text: str, width_emu: int, start_pt: float, max_lines: int = 2, min_pt: float = 24) -> float:
    """대문자·넓은 자간 제목이 max_lines 줄 안에 들어가는 가장 큰 크기 (단어 단위 줄바꿈 추정)."""
    width_pt = width_emu / 12700
    size = start_pt
    while size > min_pt:
        per_char = size * 0.74  # 대문자 평균 폭 + 자간 10%
        cap = max(1, int(width_pt / per_char))
        lines, cur = 1, 0
        for word in text.split():
            need = len(word) + (1 if cur else 0)
            if cur + need > cap:
                lines, cur = lines + 1, len(word)
            else:
                cur += need
        if lines <= max_lines:
            return size
        size -= 4
    return min_pt


def _set_cover_title(title_ph, text: str) -> None:
    title_ph.text_frame.text = text
    start = _layout_font_pt(title_ph, 64)
    size = _fit_size(text, title_ph.width, start)
    if size < start:
        for r in title_ph.text_frame.paragraphs[0].runs:
            r.font.size = Pt(size)


def _left_accent_title(slide, title_ph, title: str, subtitle: str | None, cfg: CompanyConfig) -> None:
    """MB 다크 덱 방식: 왼쪽 정렬 제목 + 같은 상자 안 두 번째 단락에 부제 + 위쪽 시안 막대."""
    l, t, w, h = (int(Cm(v)) for v in _LEFT_TITLE)
    title_ph.left, title_ph.top, title_ph.width, title_ph.height = l, t, w, h
    tf = title_ph.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title
    if cfg.title_size:
        r.font.size = Pt(cfg.title_size)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        p2.space_before = Pt(4)
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.size = Pt(cfg.subheadline_size or 14)
        r2.font.color.theme_color = MSO_THEME_COLOR.TEXT_2

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *(int(Cm(v)) for v in _ACCENT_BAR))
    bar.name = "Accent Marker"
    bar.fill.solid()
    bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
    bar.line.fill.background()
    bar.shadow.inherit = False


def _add_footers(slide, layout, cfg: CompanyConfig) -> None:
    """python-pptx 는 새 슬라이드에 날짜/바닥글/번호 placeholder 를 만들지 않으므로 레이아웃에서 복제한다."""
    wanted = {
        PP_PLACEHOLDER.FOOTER: cfg.footer_text,
        PP_PLACEHOLDER.DATE: cfg.tagline,
        PP_PLACEHOLDER.SLIDE_NUMBER: None if cfg.slide_numbers else "",
    }
    for ph in layout.placeholders:
        t = ph.placeholder_format.type
        if t not in wanted:
            continue
        text = wanted[t]
        if text == "":
            continue
        el = copy.deepcopy(ph._element)
        slide.shapes._spTree.append(el)
        if text is not None:  # 바닥글/태그라인: 텍스트 교체 (서식은 레이아웃 상속)
            new = next(s for s in slide.placeholders if s._element is el)
            new.text_frame.text = text


def _add_label(slide, text: str) -> None:
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *(int(Cm(v)) for v in _LABEL))
    shp.name = "Classification Label"
    shp.shadow.inherit = False
    strong = text.strip().upper() in {"CONFIDENTIAL", "SECRET", "STRICTLY CONFIDENTIAL"}
    if strong:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(0x74, 0x12, 0x18)
        shp.line.fill.background()
    else:
        shp.fill.background()
        shp.line.width = Pt(0.75)
        shp.line.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    tf = shp.text_frame
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text.upper()
    r.font.size = Pt(12)
    if strong:
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    else:
        r.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    r.font._rPr.set("spc", "80")


_NOTES_BODY_XML = (
    '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<p:nvSpPr><p:cNvPr id="{id}" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
    '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
)


def _set_notes(slide, text: str) -> None:
    """발표자 노트. 노트 마스터에 본문 칸이 없는 템플릿이면 본문 placeholder 를 직접 만든다."""
    notes = slide.notes_slide
    tf = notes.notes_text_frame
    if tf is None:
        ids = [int(x) for x in notes._element.xpath("//p:cNvPr/@id")] or [1]
        notes.shapes._spTree.append(etree.fromstring(_NOTES_BODY_XML.format(id=max(ids) + 1)))
        tf = notes.notes_text_frame
    tf.text = text


# ---------------------------------------------------------------------------
# 슬라이드 단위
# ---------------------------------------------------------------------------


def _default_area(prs, cfg: CompanyConfig) -> Box:
    if cfg.content_area_cm:
        return _cm_box(cfg.content_area_cm)
    W, H = prs.slide_width, prs.slide_height
    return Box(int(W * 0.06), int(H * 0.2), int(W * 0.88), int(H * 0.72))


def _render_slide(prs, ps: PlannedSlide, images: dict[str, SourceImage], cfg: CompanyConfig) -> None:
    layout = find_layout(prs, ps.layout)
    slide = prs.slides.add_slide(layout)
    roles = roles_of(slide.placeholders, prs.slide_height)

    def first(*want):
        return next((p for p, r in roles if r in want), None)

    title_ph = first(TITLE)
    sub_ph = first(SUBTITLE) or first(SUBHEADLINE)
    foot_ph = first(FOOTNOTE)
    regions = [p for p, r in roles if r in (CONTENT, PICTURE)]
    for p, r in roles:
        if r == DECOR or (r in (SUBTITLE, SUBHEADLINE) and p is not sub_ph):
            _remove(p)

    # --- 제목 / 부제 -------------------------------------------------------
    lead = None
    use_left = (
        cfg.title_style == "left_accent"
        and title_ph is not None
        and title_ph.placeholder_format.type == PP_PLACEHOLDER.TITLE
        and title_ph.height <= Cm(4)  # 'Headline Only 02' 처럼 화면 전체 제목은 제외
    )
    if title_ph is not None and ps.title and use_left:
        _left_accent_title(slide, title_ph, ps.title, ps.subtitle, cfg)
        if sub_ph is not None:
            _remove(sub_ph)
    else:
        if title_ph is not None:
            if ps.title and title_ph.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE:
                _set_cover_title(title_ph, ps.title)
            elif ps.title:
                title_ph.text_frame.text = ps.title
                if cfg.title_size and title_ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                    for r in title_ph.text_frame.paragraphs[0].runs:
                        r.font.size = Pt(cfg.title_size)
            else:
                _remove(title_ph)
        if sub_ph is not None:
            if ps.subtitle:
                sub_ph.text_frame.text = ps.subtitle
            else:
                _remove(sub_ph)
        elif ps.subtitle:
            lead = ps.subtitle  # 부제 칸이 없으면 본문 첫 줄(굵게)로

    # 제목 칸이 없는 레이아웃(간지 등): 본문이 없으면 제목을 첫 영역에 넣는다
    if title_ph is None and ps.title and regions and not (ps.bullets or ps.table or ps.image_ids):
        region = regions.pop(0)
        if region.placeholder_format.type in _TEXT_OK:
            region.text_frame.text = ps.title
        else:
            _remove(region)

    # --- 각주 ---------------------------------------------------------------
    if foot_ph is not None:
        if ps.footnote:
            foot_ph.text_frame.text = ps.footnote
        else:
            _remove(foot_ph)

    # --- 본문 항목 배치 -------------------------------------------------------
    items: list[tuple[str, object]] = []
    if ps.bullets or lead:
        items.append(("bullets", ps.bullets))
    if ps.table:
        items.append(("table", ps.table))
    items += [("image", images[i]) for i in ps.image_ids if i in images]

    free = list(regions)
    assigned: list[tuple[str, object, object, Box]] = []

    def take_region(kind: str):
        for r in free:
            if r.placeholder_format.type in _SUITS[kind]:
                free.remove(r)
                return r
        return free.pop(0) if free else None

    if items:
        direct = items if len(items) <= len(regions) else items[: max(len(regions) - 1, 0)]
        overflow = items[len(direct):]
        for kind, payload in direct:
            ph = take_region(kind)
            assigned.append((kind, payload, ph, _box_of(ph)))
        if overflow:
            host = free.pop(0) if free else None
            area = _box_of(host) if host is not None else _default_area(prs, cfg)
            horizontal = not any(k == "table" for k, _ in overflow) and area.width > area.height
            boxes = area.split(len(overflow), horizontal, gap=int(Cm(0.6)))
            for idx, ((kind, payload), box) in enumerate(zip(overflow, boxes)):
                ph = host if (host is not None and (kind == "bullets" or idx == 0)) else None
                if ph is host:
                    host = None
                assigned.append((kind, payload, ph, box))
            if host is not None:
                _remove(host)

    for kind, payload, ph, box in assigned:
        if kind == "bullets":
            _draw_bullets(slide, ph, box, payload, lead, cfg)
        elif kind == "table":
            _draw_table(slide, ph, box, payload, cfg)
        else:
            _draw_image(slide, ph, box, payload)
    for r in free:
        _remove(r)

    # --- 반복 요소 ----------------------------------------------------------
    _add_footers(slide, layout, cfg)
    if cfg.classification:
        _add_label(slide, cfg.classification)
    if ps.notes:
        _set_notes(slide, ps.notes)


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
