"""원본 .pptx 에서 슬라이드별 내용(제목, 본문, 표, 이미지, 노트)을 추출한다."""

from __future__ import annotations

from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from .models import SourceImage, SourceParagraph, SourceSlide, SourceTable
from .pptx_io import open_presentation

_TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
_SKIP_PH_TYPES = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER}


def _iter_shapes(shapes):
    """그룹 도형을 펼치고, 위→아래 / 왼쪽→오른쪽 순으로 정렬해 돌려준다."""
    flat = []
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            flat.extend(_iter_shapes(shp.shapes))
        else:
            flat.append(shp)
    return sorted(flat, key=lambda s: ((s.top or 0), (s.left or 0)))


def _ph_type(shape):
    if not shape.is_placeholder:
        return None
    try:
        return shape.placeholder_format.type
    except (AttributeError, ValueError):
        return None


def _clean(text: str) -> str:
    return " ".join(text.replace("\x0b", " ").split())


def extract_slides(path: str | Path) -> list[SourceSlide]:
    prs = open_presentation(path)
    slides: list[SourceSlide] = []

    for n, slide in enumerate(prs.slides, start=1):
        src = SourceSlide(number=n)
        img_count = 0

        for shape in _iter_shapes(slide.shapes):
            ph_type = _ph_type(shape)
            if ph_type in _SKIP_PH_TYPES:
                continue

            if ph_type in _TITLE_TYPES and shape.has_text_frame:
                title = _clean(shape.text_frame.text)
                if title and not src.title:
                    src.title = title
                    continue

            if getattr(shape, "has_table", False) and shape.has_table:
                rows = [[_clean(c.text) for c in r.cells] for r in shape.table.rows]
                src.tables.append(SourceTable(rows=rows))
                continue

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or (
                ph_type == PP_PLACEHOLDER.PICTURE and hasattr(shape, "image")
            ):
                try:
                    image = shape.image
                except (AttributeError, ValueError):
                    continue
                img_count += 1
                desc = ""
                try:
                    desc = shape._element._nvXxPr.cNvPr.get("descr") or ""
                except AttributeError:
                    pass
                src.images.append(
                    SourceImage(
                        id=f"s{n}_img{img_count}",
                        blob=image.blob,
                        ext=image.ext,
                        width=shape.width or 1,
                        height=shape.height or 1,
                        description=desc or shape.name,
                    )
                )
                continue

            if getattr(shape, "has_chart", False) and shape.has_chart:
                chart = shape.chart
                name = chart.chart_title.text_frame.text if chart.has_title else shape.name
                src.other.append(f"차트: {_clean(name)}")
                continue

            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    text = _clean(p.text)
                    if text:
                        src.paragraphs.append(SourceParagraph(text=text, level=p.level))

        if slide.has_notes_slide:
            src.notes = slide.notes_slide.notes_text_frame.text.strip()

        slides.append(src)

    return slides
