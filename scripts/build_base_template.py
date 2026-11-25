"""회사 덱 하나에서 '깨끗한 기준 템플릿'을 만든다.

- 슬라이드를 모두 지우고, 지정한 마스터 하나만 남긴다 (나머지 마스터·레이아웃은 버림)
- 테마의 동아시아(EA)/한글 폰트를 지정한다 (기본: 맑은 고딕)

    python scripts/build_base_template.py <원본덱.pptx> templates/mb_dark.pptx --master 2
    python scripts/build_base_template.py <원본덱.pptx> --list      # 마스터 목록만 보기
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from slide_converter.pptx_io import open_presentation, remove_all_slides  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A}


def master_theme(master):
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            return rel.target_part
    return None


def list_masters(prs) -> None:
    used = {}
    for s in prs.slides:
        m = s.slide_layout.slide_master
        used[id(m)] = used.get(id(m), 0) + 1
    for i, m in enumerate(prs.slide_masters, 1):
        theme = etree.fromstring(master_theme(m).blob)
        clr = theme.find(".//a:clrScheme/a:lt1/*", NS)
        bg = clr.get("val") or clr.get("lastClr")
        names = [l.name for l in m.slide_layouts]
        print(f"[{i}] theme={theme.get('name')!r} bg(lt1)={bg} slides={used.get(id(m), 0)} layouts={len(names)}")
        print("     " + ", ".join(names))


def keep_only_master(prs, keep_index: int) -> None:
    lst = prs.slide_masters._sldMasterIdLst
    entries = list(lst)
    if not 1 <= keep_index <= len(entries):
        raise SystemExit(f"--master 는 1~{len(entries)} 사이여야 합니다.")
    for i, entry in enumerate(entries, 1):
        if i != keep_index:
            prs.part.drop_rel(entry.rId)
            lst.remove(entry)


def set_ea_font(theme_part, typeface: str) -> None:
    """major/minor 폰트의 <a:ea> 와 Hang 스크립트 폰트를 지정한다."""
    root = etree.fromstring(theme_part.blob)
    for kind in ("majorFont", "minorFont"):
        f = root.find(f".//a:fontScheme/a:{kind}", NS)
        ea = f.find("a:ea", NS)
        ea.set("typeface", typeface)
        hang = next((x for x in f.findall("a:font", NS) if x.get("script") == "Hang"), None)
        if hang is None:
            hang = etree.SubElement(f, f"{{{A}}}font")
            hang.set("script", "Hang")
        hang.set("typeface", typeface)
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--master", type=int, help="남길 마스터 번호 (--list 로 확인)")
    ap.add_argument("--ea-font", default="맑은 고딕", help="한글/동아시아 폰트 (기본: 맑은 고딕)")
    ap.add_argument("--list", action="store_true")
    args = ap.parse_args()

    prs = open_presentation(args.source)
    if args.list or not args.output or not args.master:
        list_masters(prs)
        return

    remove_all_slides(prs)
    keep_only_master(prs, args.master)
    master = prs.slide_masters[0]
    set_ea_font(master_theme(master), args.ea_font)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"저장: {out} (레이아웃 {len(master.slide_layouts)}개)")
    for l in master.slide_layouts:
        print("  -", l.name)


if __name__ == "__main__":
    main()
