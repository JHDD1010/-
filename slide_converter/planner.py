"""원본 슬라이드 → 템플릿 슬라이드 계획(DeckPlan).

- plan_with_claude : Claude 가 본사 스타일 가이드에 맞게 내용을 재구성
- plan_simple      : AI 없이 1:1 로 옮기는 단순 변환 (오프라인/테스트용)
"""

from __future__ import annotations

import json

from .config import CompanyConfig
from .models import Bullet, DeckPlan, PlannedSlide, SourceSlide, Table
from .template import LayoutInfo, pick_layout_name

SYSTEM_PROMPT = """\
당신은 회사 보고서 슬라이드를 본사 표준 템플릿에 맞게 재구성하는 편집자입니다.

입력으로 (1) 본사 템플릿의 레이아웃 카탈로그, (2) 본사 문서 스타일 가이드, (3) 원본 슬라이드에서 추출한 내용(JSON)을 받습니다.
원본 내용을 본사 템플릿 슬라이드 계획으로 변환하세요.

규칙:
- layout 값은 반드시 카탈로그의 name 중 하나를 그대로 사용합니다.
- 각 레이아웃의 content_areas 개수와 종류를 고려해 배치합니다. 표·이미지가 있는 슬라이드는 이를 담을 영역이 있는 레이아웃을 우선 고릅니다.
- 원본의 사실·수치·고유명사를 보존하고, 원본에 없는 내용을 지어내지 않습니다. 원본이 모호하면 원문 표현을 유지합니다.
- 스타일 가이드의 문체·분량 규칙에 맞게 문장을 다듬고, 분량이 넘치면 슬라이드를 나누고, 지나치게 빈약한 연속 슬라이드는 합칠 수 있습니다.
- 원본 이미지는 image_ids 로만 참조하며, 존재하는 id 만 사용합니다. 의미 있는 이미지는 빠뜨리지 않습니다.
- 차트 등 옮길 수 없는 요소(other_elements)는 발표자 노트에 '[원본 확인 필요] ...' 로 남깁니다.
- 원본 발표자 노트는 notes 에 보존합니다. 출처·주석은 footnote 에 넣습니다.
- 출력 언어는 스타일 가이드를 따르고, 가이드에 언급이 없으면 원본 언어를 그대로 유지합니다 (번역하지 않음).
- 카탈로그에서 content_areas 가 없는 레이아웃이라도 글머리·표·이미지는 기본 본문 영역에 배치되므로 사용할 수 있습니다.
- source_slides 에 근거가 된 원본 슬라이드 번호를 적습니다.
"""


def _build_user_prompt(
    slides: list[SourceSlide], catalog: list[LayoutInfo], cfg: CompanyConfig
) -> str:
    cover_hint = ""
    if cfg.add_cover:
        cover_hint = (
            "\n첫 슬라이드는 표지(title 성격의 레이아웃)로 만드세요."
            + (f" 표지 제목: {cfg.cover_title}" if cfg.cover_title else "")
            + (f" 표지 부제: {cfg.cover_subtitle}" if cfg.cover_subtitle else "")
        )
    return (
        "## 템플릿 레이아웃 카탈로그\n"
        + json.dumps([l.to_prompt_dict() for l in catalog], ensure_ascii=False, indent=1)
        + "\n\n## 본사 스타일 가이드\n"
        + cfg.style_guide.strip()
        + f"\n- 슬라이드당 글머리 최대 {cfg.max_bullets}개."
        + cover_hint
        + "\n\n## 원본 슬라이드\n"
        + json.dumps([s.to_prompt_dict() for s in slides], ensure_ascii=False, indent=1)
    )


class PlanningError(RuntimeError):
    pass


def plan_with_claude(
    slides: list[SourceSlide],
    catalog: list[LayoutInfo],
    cfg: CompanyConfig,
    client=None,
) -> DeckPlan:
    import anthropic

    client = client or anthropic.Anthropic()

    with client.beta.messages.stream(
        model=cfg.model,
        max_tokens=64000,
        thinking={"type": "adaptive"},
        output_config={"effort": cfg.effort},
        # 안전 분류기가 요청을 거절하면 서버가 대체 모델로 자동 재시도
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": _build_user_prompt(slides, catalog, cfg)}],
        output_format=DeckPlan,
    ) as stream:
        message = stream.get_final_message()

    if message.stop_reason == "refusal":
        raise PlanningError("Claude 가 요청을 처리하지 않았습니다 (refusal).")
    if message.stop_reason == "max_tokens":
        raise PlanningError("출력이 max_tokens 에서 잘렸습니다. 원본을 나눠서 변환하세요.")
    plan = message.parsed_output
    if plan is None:
        raise PlanningError("Claude 응답을 변환 계획으로 해석하지 못했습니다.")
    return plan


def plan_simple(
    slides: list[SourceSlide], catalog: list[LayoutInfo], cfg: CompanyConfig
) -> DeckPlan:
    """AI 없이 원본 슬라이드를 1:1 로 템플릿 레이아웃에 옮긴다."""
    pref = cfg.preferred_layouts
    out: list[PlannedSlide] = []

    if cfg.add_cover and cfg.cover_title:
        out.append(
            PlannedSlide(
                layout=pick_layout_name(catalog, kind="title", preferred=pref),
                title=cfg.cover_title,
                subtitle=cfg.cover_subtitle or None,
            )
        )

    for i, s in enumerate(slides):
        first_like = i == 0 and not s.tables and len(s.paragraphs) <= 3
        is_cover = cfg.add_cover and not cfg.cover_title and (s.is_title_slide or first_like) and i < 2
        n_items = (1 if s.paragraphs else 0) + len(s.tables[:1]) + len(s.images)

        table = None
        if s.tables and s.tables[0].rows:
            rows = s.tables[0].rows
            # 첫 행에 숫자가 있으면 데이터 행으로 본다 (예: 아젠다 '15min | Intro')
            has_header = len(rows) > 1 and not any(ch.isdigit() for cell in rows[0] for ch in cell)
            table = Table(header=rows[0], rows=rows[1:]) if has_header else Table(header=[], rows=rows)

        notes = s.notes
        if s.other:
            notes = (notes + "\n" if notes else "") + "\n".join(f"[원본 확인 필요] {o}" for o in s.other)

        if is_cover:
            kind = "cover_picture" if s.images and "cover_picture" in pref else "title"
            subtitle = s.subtitle or " / ".join(p.text for p in s.paragraphs[:2])
            out.append(
                PlannedSlide(
                    layout=pref.get(kind) or pick_layout_name(catalog, kind="title", preferred=pref),
                    title=s.title or (s.paragraphs[0].text if s.paragraphs else ""),
                    subtitle=subtitle or None,
                    image_ids=[im.id for im in s.images[:1]] if kind == "cover_picture" else [],
                    notes=notes,
                    source_slides=[s.number],
                )
            )
            continue

        if n_items == 0:
            kind = "section"
        elif n_items >= 2:
            kind = "two_content"
        else:
            kind = "content"
        out.append(
            PlannedSlide(
                layout=pick_layout_name(catalog, kind=kind, preferred=pref),
                title=s.title,
                subtitle=s.subtitle or None,
                bullets=[Bullet(text=p.text, level=p.level) for p in s.paragraphs],
                table=table,
                image_ids=[im.id for im in s.images],
                notes=notes,
                source_slides=[s.number],
            )
        )
    return DeckPlan(slides=out)


def sanitize_plan(
    plan: DeckPlan, slides: list[SourceSlide], catalog: list[LayoutInfo], cfg: CompanyConfig
) -> tuple[DeckPlan, list[str]]:
    """존재하지 않는 레이아웃/이미지 id 등을 보정하고 경고 목록을 돌려준다."""
    warnings: list[str] = []
    layout_names = {l.name for l in catalog}
    image_ids = {im.id for s in slides for im in s.images}
    used_images: set[str] = set()

    for i, ps in enumerate(plan.slides, start=1):
        if ps.layout not in layout_names:
            kind = "two_content" if (ps.table or ps.image_ids) and ps.bullets else "content"
            new = pick_layout_name(catalog, kind=kind, preferred=cfg.preferred_layouts)
            warnings.append(f"슬라이드 {i}: 알 수 없는 레이아웃 '{ps.layout}' → '{new}'")
            ps.layout = new
        bad = [x for x in ps.image_ids if x not in image_ids]
        if bad:
            warnings.append(f"슬라이드 {i}: 없는 이미지 id 제거 {bad}")
            ps.image_ids = [x for x in ps.image_ids if x in image_ids]
        used_images.update(ps.image_ids)
        for b in ps.bullets:
            b.level = max(0, min(4, b.level))

    missing = sorted(image_ids - used_images)
    if missing:
        warnings.append(f"계획에 포함되지 않은 원본 이미지: {missing}")
    return plan, warnings
