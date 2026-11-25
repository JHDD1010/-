"""명령행 진입점.

사용 예:
    python -m slide_converter input/보고서.pptx -c config/company.toml -o output/보고서_본사양식.pptx
    python -m slide_converter input/보고서.pptx -t templates/본사.potx --no-ai
    python -m slide_converter --inspect-template templates/본사.pptx
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .config import load_config
from .extract import extract_slides
from .models import DeckPlan
from .planner import PlanningError, plan_simple, plan_with_claude, sanitize_plan
from .pptx_io import open_presentation
from .render import render_deck
from .template import build_catalog


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="slide_converter",
        description="슬라이드 장표(.pptx)를 본사 템플릿 양식으로 변환합니다.",
    )
    p.add_argument("input", nargs="?", help="변환할 원본 .pptx")
    p.add_argument("-c", "--config", help="회사 설정 TOML (템플릿 경로, 스타일 가이드 등)")
    p.add_argument("-t", "--template", help="본사 템플릿 .pptx/.potx (설정 파일보다 우선)")
    p.add_argument("-o", "--output", help="출력 .pptx 경로 (기본: output/<원본>_converted.pptx)")
    p.add_argument("--no-ai", action="store_true", help="Claude 없이 1:1 서식 변환만 수행")
    p.add_argument("--model", help="Claude 모델 ID (기본: 설정 파일 또는 claude-opus-5-5)")
    p.add_argument("--dump-plan", metavar="JSON", help="변환 계획을 JSON 으로 저장 (검토/수정용)")
    p.add_argument("--from-plan", metavar="JSON", help="저장된 변환 계획으로 렌더링만 수행")
    p.add_argument("--inspect-template", metavar="TEMPLATE", help="템플릿 레이아웃 목록만 출력")
    return p


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)

    if args.inspect_template:
        prs = open_presentation(args.inspect_template)
        for info in build_catalog(prs):
            print(json.dumps(info.to_prompt_dict(), ensure_ascii=False))
        return 0

    if not args.input:
        _parser().error("원본 .pptx 경로가 필요합니다.")

    cfg = load_config(args.config)
    if args.model:
        cfg.model = args.model
    template = Path(args.template) if args.template else cfg.template_path
    if not template or not template.exists():
        print(f"템플릿을 찾을 수 없습니다: {template}", file=sys.stderr)
        return 2

    src_path = Path(args.input)
    output = Path(args.output) if args.output else Path("output") / f"{src_path.stem}_converted.pptx"

    print(f"[1/3] 원본 추출: {src_path}")
    slides = extract_slides(src_path)
    catalog = build_catalog(open_presentation(template), cfg.layout_descriptions)
    print(f"      원본 {len(slides)}장, 템플릿 레이아웃 {len(catalog)}개")

    if args.from_plan:
        print(f"[2/3] 저장된 계획 사용: {args.from_plan}")
        plan = DeckPlan.model_validate_json(Path(args.from_plan).read_text(encoding="utf-8"))
    elif args.no_ai:
        print("[2/3] 단순 변환 (AI 미사용)")
        plan = plan_simple(slides, catalog, cfg)
    else:
        print(f"[2/3] Claude 로 내용 재구성 ({cfg.model})")
        try:
            plan = plan_with_claude(slides, catalog, cfg)
        except PlanningError as e:
            print(f"오류: {e}", file=sys.stderr)
            return 1

    plan, warnings = sanitize_plan(plan, slides, catalog, cfg)
    for w in warnings:
        print(f"      경고: {w}")

    if args.dump_plan:
        Path(args.dump_plan).write_text(plan.model_dump_json(indent=2), encoding="utf-8")
        print(f"      계획 저장: {args.dump_plan}")

    print(f"[3/3] 템플릿 적용: {template}")
    out = render_deck(plan, slides, template, output, cfg)
    print(f"완료: {out} ({len(plan.slides)}장)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
