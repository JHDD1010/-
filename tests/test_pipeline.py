import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import make_samples  # noqa: E402
from slide_converter.cli import main  # noqa: E402
from slide_converter.config import load_config  # noqa: E402
from slide_converter.extract import extract_slides  # noqa: E402
from slide_converter.models import Bullet, DeckPlan, PlannedSlide, Table  # noqa: E402
from slide_converter.planner import plan_with_claude, sanitize_plan  # noqa: E402
from slide_converter.pptx_io import open_presentation  # noqa: E402
from slide_converter.render import render_deck  # noqa: E402
from slide_converter.template import build_catalog  # noqa: E402


@pytest.fixture(scope="module")
def samples(tmp_path_factory):
    d = tmp_path_factory.mktemp("samples")
    make_samples.make_template(d / "template.pptx")
    make_samples.make_input(d / "input.pptx")
    return d


def test_extract(samples):
    slides = extract_slides(samples / "input.pptx")
    assert [s.title for s in slides][:3] == ["2026년 4분기 영업 실적 보고", "요약", "지역별 실적"]
    assert slides[1].paragraphs[1].level == 1
    assert slides[1].notes == "수치는 11월 20일 마감 기준"
    assert slides[2].tables[0].rows[1] == ["수도권", "88", "+18%"]
    assert slides[3].images[0].description == "판교 신규 매장 외관 사진"
    assert slides[4].other == ["차트: Chart 2"] or slides[4].other[0].startswith("차트")


def test_no_ai_end_to_end(samples, tmp_path):
    out = tmp_path / "out.pptx"
    rc = main([str(samples / "input.pptx"), "-t", str(samples / "template.pptx"), "-o", str(out), "--no-ai"])
    assert rc == 0
    prs = Presentation(str(out))
    assert len(prs.slides) == 5  # 템플릿 예시 슬라이드는 제거됨
    assert prs.slide_width == make_samples.Inches(13.333)  # 템플릿 크기 사용
    assert prs.slides[0].slide_layout.name == "Title Slide"
    kinds = [sh.shape_type for sh in prs.slides[3].shapes]
    assert 13 in kinds  # 그림 이관


class _FakeStream:
    def __init__(self, message):
        self._m = message

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def get_final_message(self):
        return self._m


class _FakeClient:
    def __init__(self, plan, stop_reason="end_turn"):
        self.kwargs = None
        msg = SimpleNamespace(stop_reason=stop_reason, parsed_output=plan)
        outer = self

        class _Messages:
            def stream(self, **kw):
                outer.kwargs = kw
                return _FakeStream(msg)

        self.beta = SimpleNamespace(messages=_Messages())


def test_claude_plan_and_render(samples, tmp_path):
    slides = extract_slides(samples / "input.pptx")
    cfg = load_config(None)
    catalog = build_catalog(open_presentation(samples / "template.pptx"))
    plan = DeckPlan(
        slides=[
            PlannedSlide(layout="Title Slide", title="4Q 영업 실적", subtitle="영업기획팀 | 2026.11"),
            PlannedSlide(
                layout="Two Content",
                title="지역별 실적",
                subtitle="수도권 중심 성장",
                bullets=[Bullet(text="매출 152억원(전년比 +12%)", level=0), Bullet(text="x", level=9)],
                table=Table(header=["지역", "매출"], rows=[["수도권", "88"]]),
                image_ids=["s4_img1", "nope"],
            ),
            PlannedSlide(layout="없는레이아웃", title="기타", bullets=[Bullet(text="a", level=0)]),
        ]
    )
    client = _FakeClient(plan)
    got = plan_with_claude(slides, catalog, cfg, client=client)

    kw = client.kwargs
    assert kw["model"] == "claude-opus-5-5"
    assert kw["output_format"] is DeckPlan
    assert kw["fallbacks"] == "default"
    assert "Two Content" in kw["messages"][0]["content"]

    got, warnings = sanitize_plan(got, slides, catalog, cfg)
    assert got.slides[1].image_ids == ["s4_img1"]
    assert got.slides[1].bullets[1].level == 4
    assert got.slides[2].layout in {l.name for l in catalog}
    assert any("없는레이아웃" in w for w in warnings)

    out = render_deck(got, slides, samples / "template.pptx", tmp_path / "ai.pptx", cfg)
    prs = Presentation(str(out))
    s2 = prs.slides[1]
    types = {sh.shape_type for sh in s2.shapes}
    assert 19 in types and 13 in types  # 표 + 그림
    texts = " ".join(sh.text_frame.text for sh in s2.shapes if sh.has_text_frame)
    assert "수도권 중심 성장" in texts  # 부제 영역이 없으면 본문 첫 줄로


def test_refusal_raises(samples):
    from slide_converter.planner import PlanningError

    slides = extract_slides(samples / "input.pptx")
    catalog = build_catalog(open_presentation(samples / "template.pptx"))
    with pytest.raises(PlanningError):
        plan_with_claude(slides, catalog, load_config(None), client=_FakeClient(None, "refusal"))
