"""테스트용 예시 템플릿과 원본 장표를 만든다.

    python scripts/make_samples.py

- templates/sample_company_template.pptx : 마스터에 색 띠·로고 텍스트를 넣은 예시 '본사 템플릿'
- input/sample_input.pptx                : 제목/글머리/표/이미지/차트가 섞인 예시 원본

실제 사용 시에는 templates/ 에 본사 템플릿을 넣고 config/company.toml 의 경로를 바꾸면 된다.
"""

from __future__ import annotations

import io
import struct
import sys
import zlib
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from slide_converter.pptx_io import remove_all_slides  # noqa: E402
NAVY = RGBColor(0x0B, 0x2A, 0x4A)


def _png(w: int, h: int, rgb=(70, 130, 180)) -> bytes:
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def make_template(path: Path) -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    master = prs.slide_master
    # MasterShapes 에는 add_shape 가 없으므로 임시 슬라이드에 그린 뒤 마스터로 옮긴다
    tmp = prs.slides.add_slide(prs.slide_layouts[6])

    band = tmp.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.25))
    band.fill.solid()
    band.fill.fore_color.rgb = NAVY
    band.line.fill.background()
    logo = tmp.shapes.add_textbox(Inches(11.3), Inches(6.95), Inches(1.9), Inches(0.4))
    logo.text_frame.text = "SAMPLE CORP"
    logo.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
    logo.text_frame.paragraphs[0].runs[0].font.bold = True
    logo.text_frame.paragraphs[0].runs[0].font.color.rgb = NAVY
    for shp in (band, logo):
        master.shapes._spTree.append(shp._element)
    remove_all_slides(prs)

    # 기본 레이아웃 placeholder 들을 16:9 폭에 맞게 늘림
    sx = prs.slide_width / Inches(10)
    for layout in list(prs.slide_layouts) + [master]:
        for ph in layout.placeholders:
            # 상속 좌표를 먼저 읽고 네 값을 함께 써야 위치가 깨지지 않는다
            l, t, w, h = ph.left, ph.top, ph.width, ph.height
            ph.left, ph.top, ph.width, ph.height = int(l * sx), t, int(w * sx), h

    # 제목 서식: 남색, 맑은 고딕
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            if ph.placeholder_format.type in (1, 3):  # TITLE, CENTER_TITLE
                for p in ph.text_frame.paragraphs:
                    for r in p.runs or [p.add_run()]:
                        r.font.color.rgb = NAVY
                        r.font.name = "맑은 고딕"

    # 예시 슬라이드 하나 (변환 시 삭제되는지 확인용)
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "템플릿 예시 슬라이드"
    prs.save(str(path))


def make_input(path: Path) -> None:
    prs = Presentation()  # 4:3, 기본 서식 = '다른 양식'
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "2026년 4분기 영업 실적 보고"
    s.placeholders[1].text = "영업기획팀 | 2026.11"

    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "요약"
    tf = s.placeholders[1].text_frame
    tf.text = "4분기 매출은 전년 동기 대비 12% 증가하여 152억원을 기록하였습니다."
    for text, lvl in [
        ("신규 고객 확보가 성장의 주요 요인이었습니다.", 1),
        ("특히 수도권 지역 B2B 계약이 크게 늘었습니다.", 1),
        ("영업이익률은 원자재 가격 상승으로 인해 1.2%p 하락하였습니다.", 0),
        ("내년 1분기에는 가격 정책 재검토가 필요할 것으로 보입니다.", 0),
    ]:
        p = tf.add_paragraph()
        p.text, p.level = text, lvl
    s.notes_slide.notes_text_frame.text = "수치는 11월 20일 마감 기준"

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "지역별 실적"
    rows = [["지역", "매출(억원)", "전년比"], ["수도권", "88", "+18%"], ["영남", "37", "+6%"], ["호남", "27", "+3%"]]
    tbl = s.shapes.add_table(len(rows), 3, Inches(1), Inches(2), Inches(8), Inches(2)).table
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            tbl.cell(r, c).text = v

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "신규 매장 전경 및 추진 현황"
    pic = s.shapes.add_picture(io.BytesIO(_png(160, 90)), Inches(0.5), Inches(2), Inches(4.5))
    pic._element._nvXxPr.cNvPr.set("descr", "판교 신규 매장 외관 사진")
    tb = s.shapes.add_textbox(Inches(5.3), Inches(2), Inches(4.2), Inches(3))
    tb.text_frame.text = "판교 매장 10월 오픈 완료"
    tb.text_frame.add_paragraph().text = "12월 송도 매장 오픈 예정"

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "월별 매출 추이"
    cd = CategoryChartData()
    cd.categories = ["10월", "11월", "12월"]
    cd.add_series("매출", (48, 50, 54))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(8), Inches(4.5), cd).chart.has_title = False

    prs.save(str(path))


if __name__ == "__main__":
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT
    (out_dir / "templates").mkdir(parents=True, exist_ok=True)
    (out_dir / "input").mkdir(parents=True, exist_ok=True)
    make_template(out_dir / "templates" / "sample_company_template.pptx")
    make_input(out_dir / "input" / "sample_input.pptx")
    print("생성 완료:", out_dir / "templates" / "sample_company_template.pptx", out_dir / "input" / "sample_input.pptx")
