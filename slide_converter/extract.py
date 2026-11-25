"""원본 .pptx 에서 슬라이드별 내용(제목, 본문, 표, 이미지, 노트)을 추출한다.

회사 덱에서 흔한 잡음은 걸러낸다:
- 바닥글·태그라인·쪽번호 (placeholder 이거나, 일반 텍스트 상자로 복사된 경우 모두)
- 위/아래 띠 영역의 워드마크·보안 등급 라벨(INTERNAL, CONFIDENTIAL …)
- 화면 전체를 덮는 배경 이미지, 아주 작은 아이콘
제목 placeholder 가 비어 있으면, 위쪽에서 가장 큰 글자의 텍스트 상자를 제목으로 본다.
"""

from __future__ import annotations

import re
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Cm

from .models import SourceImage, SourceParagraph, SourceSlide, SourceTable
from .pptx_io import open_presentation

_TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
_SKIP_PH_TYPES = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER}
_FOOTER_NAME = re.compile(r"(Fußzeilen|Datums|Foliennummern|Footer|Date|Slide Number)\w*", re.I)
_LABELS = {"INTERNAL", "CONFIDENTIAL", "SECRET", "STRICTLY CONFIDENTIAL", "PUBLIC", "DRAFT"}
_ICON_MAX = Cm(2.2)


def _iter_shapes(shapes, box=None):
    """그룹을 펼쳐 (도형, 슬라이드 기준 위치) 를 돌려준다.

    그룹 안 도형의 좌표는 그룹 좌표계이므로, 위치 판단에는 최상위 그룹의 위치를 쓴다.
    """
    out = []
    for shp in shapes:
        own = (shp.left or 0, shp.top or 0, shp.width or 0, shp.height or 0)
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            out.extend(_iter_shapes(shp.shapes, box or own))
        else:
            out.append((shp, box or own))
    return out


def _ph_type(shape):
    if not shape.is_placeholder:
        return None
    try:
        return shape.placeholder_format.type
    except (AttributeError, ValueError):
        return None


def _clean(text: str) -> str:
    return " ".join(text.replace("\x0b", " ").split())


def _max_font_pt(shape) -> float:
    sizes = [
        r.font.size.pt
        for p in shape.text_frame.paragraphs
        for r in p.runs
        if r.font.size is not None
    ]
    return max(sizes, default=0.0)


def _in_band(box, slide_h) -> bool:
    """위 1.4 cm / 아래 1.6 cm 띠 안의 낮은 도형 (워드마크, 라벨, 바닥글)."""
    _, top, _, h = box
    return h < Cm(1.2) and (top + h < Cm(1.4) or top > slide_h - Cm(1.6))


def extract_slides(path: str | Path) -> list[SourceSlide]:
    prs = open_presentation(path)
    W, H = prs.slide_width, prs.slide_height
    slides: list[SourceSlide] = []

    for n, slide in enumerate(prs.slides, start=1):
        src = SourceSlide(number=n, layout_name=slide.slide_layout.name)
        img_count = 0
        skipped_icons = 0
        shapes = sorted(_iter_shapes(slide.shapes), key=lambda sb: (sb[1][1], sb[1][0]))

        # 1) 제목: placeholder → 없으면 위쪽의 가장 큰 글자 텍스트 상자
        title_shape = None
        for shape, _ in shapes:
            if _ph_type(shape) in _TITLE_TYPES and shape.has_text_frame and _clean(shape.text_frame.text):
                title_shape = shape
                src.is_title_slide = _ph_type(shape) == PP_PLACEHOLDER.CENTER_TITLE
                break
        if title_shape is None:
            cands = [
                (shape, _max_font_pt(shape), box)
                for shape, box in shapes
                if shape.has_text_frame
                and len(_clean(shape.text_frame.paragraphs[0].text)) >= 4  # '01' 같은 번호 제외
                and box[1] < H * 0.25
                and not _in_band(box, H)
            ]
            named = [c for c in cands if re.match(r"^(Title|Titel|Headline)\b", c[0].name or "", re.I)]
            # 크기를 지정하지 않은 텍스트는 24pt 로 간주. 제목 띠(위 20%)에 있는 후보를 먼저 본다.
            scored = [(c[0], c[1] or 24.0, c[2]) for c in cands]
            zone = [c for c in scored if c[2][1] < H * 0.2 and c[1] >= 18]
            pool = named or zone or [c for c in scored if c[1] >= 20]
            if pool:
                title_shape = max(pool, key=lambda c: (c[1], -c[2][1]))[0]

        if title_shape is not None:
            paras = [_clean(p.text) for p in title_shape.text_frame.paragraphs]
            paras = [p for p in paras if p]
            if title_shape.is_placeholder:
                src.title = " ".join(paras)
            else:
                # 텍스트 상자 제목: 첫 단락 = 제목, 나머지 = 부제
                src.title = paras[0]
                src.subtitle = " ".join(paras[1:])

        # 2) 나머지 도형
        block = 0
        for shape, box in shapes:
            if shape is title_shape:
                continue
            ph_type = _ph_type(shape)
            if ph_type in _SKIP_PH_TYPES or _FOOTER_NAME.match(shape.name or ""):
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
                _, _, w, h = box if box[2] else (0, 0, shape.width or 0, shape.height or 0)
                if _in_band(box, H):
                    continue  # 워드마크/로고
                if w >= W * 0.8 and h >= H * 0.8:
                    src.other.append(f"배경 이미지: {shape.name}")
                    continue
                if (shape.width or 0) < _ICON_MAX and (shape.height or 0) < _ICON_MAX:
                    skipped_icons += 1
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

            if shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT:
                if re.search(r"think-?cell data", shape.name or "", re.I):
                    continue  # think-cell 내부 데이터 개체 (내용 없음)
                if shape.width and shape.height:  # 마스터용 0 크기 think-cell 데이터 개체는 무시
                    prog = getattr(shape.ole_format, "prog_id", "") or "OLE"
                    kind = "think-cell 차트" if "TCLayout" in prog else f"OLE 개체({prog})"
                    src.other.append(f"{kind}: {shape.name}")
                continue

            if getattr(shape, "has_chart", False) and shape.has_chart:
                chart = shape.chart
                name = chart.chart_title.text_frame.text if chart.has_title else shape.name
                src.other.append(f"차트: {_clean(name)}")
                continue

            if shape.has_text_frame:
                text_all = _clean(shape.text_frame.text)
                if not text_all:
                    continue
                if text_all.upper() in _LABELS:
                    continue
                if _in_band(box, H) and len(text_all) < 80:
                    continue  # 바닥글/태그라인이 일반 텍스트 상자로 들어간 경우
                block += 1
                for p in shape.text_frame.paragraphs:
                    text = _clean(p.text)
                    if text:
                        src.paragraphs.append(SourceParagraph(text=text, level=p.level, block=block))

        if skipped_icons:
            src.other.append(f"작은 아이콘 {skipped_icons}개 생략")

        if slide.has_notes_slide:
            ntf = slide.notes_slide.notes_text_frame
            if ntf is not None:
                src.notes = ntf.text.strip()

        slides.append(src)

    return slides
