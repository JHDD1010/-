"""MB 템플릿 대응 기능 테스트 (사내 템플릿 없이 합성 파일로 검증)."""

import io
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import make_samples  # noqa: E402
from slide_converter.config import CompanyConfig  # noqa: E402
from slide_converter.extract import extract_slides  # noqa: E402
from slide_converter.models import Bullet, DeckPlan, PlannedSlide, Table  # noqa: E402
from slide_converter.render import _fit_size, render_deck  # noqa: E402
from slide_converter.template import (  # noqa: E402
    CONTENT,
    DECOR,
    FOOTNOTE,
    SUBHEADLINE,
    TITLE,
    placeholder_role,
)

H = int(Cm(19.05))


def _ph(t, top, h, w=Cm(30.81)):
    return SimpleNamespace(placeholder_format=SimpleNamespace(type=t), top=int(Cm(top)), height=int(Cm(h)), width=int(w), left=int(Cm(1.53)))


def test_placeholder_roles_mb_style():
    title = _ph(PP_PLACEHOLDER.TITLE, 1.74, 2.9)
    assert placeholder_role(title, H, title) == TITLE
    assert placeholder_role(_ph(PP_PLACEHOLDER.BODY, 3.51, 1.55), H, title) == SUBHEADLINE
    assert placeholder_role(_ph(PP_PLACEHOLDER.BODY, 16.93, 0.6), H, title) == FOOTNOTE
    assert placeholder_role(_ph(PP_PLACEHOLDER.BODY, 0.71, 0.35, Cm(3.04)), H, title) == DECOR
    assert placeholder_role(_ph(PP_PLACEHOLDER.BODY, 2.13, 14.8), H, None) == CONTENT  # 아젠다/간지 본문
    assert placeholder_role(_ph(PP_PLACEHOLDER.OBJECT, 5.42, 12.1), H, title) == CONTENT


def test_fit_size_shrinks_long_cover_titles():
    w = int(Cm(26.41))
    assert _fit_size("NEXT GEN", w, 64) == 64
    assert _fit_size("NEXT GEN RD SCOUTING INTERNATIONAL KICKOFF 29.10.2026", w, 64) < 64


@pytest.fixture()
def template(tmp_path):
    p = tmp_path / "template.pptx"
    make_samples.make_template(p)
    return p


def test_left_accent_title_footer_and_label(template, tmp_path):
    cfg = CompanyConfig(
        title_style="left_accent",
        title_size=28,
        subheadline_size=14,
        footer_text="KICKOFF | NOVEMBER 2026",
        tagline="Pioneering Next",
        slide_numbers=False,
        classification="INTERNAL",
        table_style="lines",
        content_area_cm=[1.53, 5.6, 30.81, 11.1],
    )
    plan = DeckPlan(
        slides=[
            PlannedSlide(
                layout="Title Only",
                title="One network, one heartbeat",
                subtitle="Five regions, one voice",
                bullets=[Bullet(text="Germany", level=0), Bullet(text="Korea", level=1)],
                table=Table(header=["Hub", "Focus"], rows=[["Korea", "Academia"]]),
                footnote="Source: workshop",
                notes="speaker note",
            )
        ]
    )
    out = render_deck(plan, [], template, tmp_path / "o.pptx", cfg)
    s = Presentation(str(out)).slides[0]
    names = [sh.name for sh in s.shapes]
    assert "Accent Marker" in names and "Classification Label" in names

    title = s.shapes.title
    assert title.text_frame.paragraphs[0].alignment == PP_ALIGN.LEFT
    assert title.text_frame.paragraphs[0].runs[0].font.size == Pt(28)
    assert title.text_frame.paragraphs[1].text == "Five regions, one voice"

    footers = {sh.placeholder_format.type: sh.text_frame.text for sh in s.placeholders if sh.placeholder_format.type in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.SLIDE_NUMBER)}
    assert footers == {PP_PLACEHOLDER.FOOTER: "KICKOFF | NOVEMBER 2026", PP_PLACEHOLDER.DATE: "Pioneering Next"}

    tbl = next(sh for sh in s.shapes if sh.has_table)
    assert tbl.top >= Cm(5.6)  # 기본 본문 영역 안
    assert "2D5ABB26" in tbl.table._tbl.tblPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId").text
    assert s.notes_slide.notes_text_frame.text == "speaker note"


def test_extract_textbox_title_and_noise_filter(tmp_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.87), Cm(19.05)
    s = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

    def tb(text, l, t, w, h, size=None, name=None):
        box = s.shapes.add_textbox(Cm(l), Cm(t), Cm(w), Cm(h))
        box.text_frame.text = text
        if size:
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(size)
        if name:
            box.name = name
        return box

    t = tb("Teams channels: new set-up", 1.5, 2.0, 30, 2.0, 28)
    p2 = t.text_frame.add_paragraph()
    p2.text = "Workshop summary"
    tb("01", 7.6, 5.7, 3, 1.7, 40)
    tb("All Internationals can now post", 1.5, 8, 9, 2, 14)
    tb("INTERNAL", 29.9, 0.2, 3.7, 0.6, 12)
    tb("PROJECT | NOVEMBER 2026", 19.1, 17.83, 13.2, 0.5, 8)
    png = make_samples._png(20, 20)
    s.shapes.add_picture(io.BytesIO(png), 0, 0, prs.slide_width, prs.slide_height)  # 배경
    s.shapes.add_picture(io.BytesIO(png), Cm(3), Cm(10), Cm(1), Cm(1))  # 아이콘
    s.shapes.add_picture(io.BytesIO(png), Cm(15), Cm(6), Cm(8), Cm(8))  # 본문 이미지
    path = tmp_path / "in.pptx"
    prs.save(str(path))

    src = extract_slides(path)[0]
    assert src.title == "Teams channels: new set-up"
    assert src.subtitle == "Workshop summary"
    texts = [p.text for p in src.paragraphs]
    assert texts == ["01", "All Internationals can now post"]
    assert len(src.images) == 1
    assert any("배경 이미지" in o for o in src.other)
    assert any("아이콘 1개" in o for o in src.other)
