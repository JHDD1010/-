"""다른 .pptx 의 슬라이드를 '있는 그대로' 가져온다.

슬라이드만 복사하면 원본의 배경·서식(마스터/레이아웃/테마)이 사라지므로,
슬라이드가 참조하는 마스터·레이아웃·테마·이미지·OLE 등 연결된 부품을 함께 복제해
대상 프레젠테이션에 새 마스터로 추가한다. 발표자 노트는 가져오지 않는다.
"""

from __future__ import annotations

import re
from pathlib import Path

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, PackURI, _Relationship
from pptx.oxml.ns import qn

from .pptx_io import open_presentation

# 따라가지 않는 관계 (복제하면 엉뚱한 부품까지 딸려 오거나 대상에 이미 있는 것)
_SKIP_RELTYPES = {RT.NOTES_SLIDE, RT.NOTES_MASTER, RT.HANDOUT_MASTER, RT.SLIDE}


def _template_of(partname: str) -> str:
    """'/ppt/slides/slide12.xml' → '/ppt/slides/slide%d.xml'"""
    m = re.match(r"^(.*?)(\d+)(\.[^./]+)$", partname)
    if m:
        return f"{m.group(1)}%d{m.group(3)}"
    stem, dot, ext = partname.rpartition(".")
    return f"{stem}%d.{ext}" if dot else partname + "%d"


class _Cloner:
    def __init__(self, target_package):
        self.pkg = target_package
        self.used = {str(p.partname) for p in target_package.iter_parts()}
        self.done: dict[int, object] = {}

    def _partname(self, src_partname: str) -> PackURI:
        tmpl = _template_of(str(src_partname))
        n = 1
        while (tmpl % n) in self.used:
            n += 1
        name = tmpl % n
        self.used.add(name)
        return PackURI(name)

    def clone(self, part):
        if id(part) in self.done:
            return self.done[id(part)]
        new = PartFactory(self._partname(part.partname), part.content_type, self.pkg, part.blob)
        self.done[id(part)] = new
        for rId, rel in part.rels.items():
            if rel.is_external:
                new.rels._rels[rId] = _Relationship(new.rels._base_uri, rId, rel.reltype, RTM.EXTERNAL, rel.target_ref)
                continue
            if rel.reltype in _SKIP_RELTYPES:
                continue
            target = self.clone(rel.target_part)
            new.rels._rels[rId] = _Relationship(new.rels._base_uri, rId, rel.reltype, RTM.INTERNAL, target)
        return new


def _max_id(prs) -> int:
    ids = [int(x.get("id")) for x in prs.part._element.iter(qn("p:sldMasterId"))]
    for master in prs.slide_masters:
        ids += [int(x.get("id")) for x in master._element.iter(qn("p:sldLayoutId"))]
    return max(ids, default=2147483647)


def import_slides(target_prs, source_path: str | Path, indices: list[int] | None = None, position: int | None = None) -> int:
    """source 의 슬라이드(0부터 시작하는 indices, 없으면 전부)를 target 의 position 에 그대로 넣는다.

    돌려주는 값은 가져온 슬라이드 수.
    """
    src = open_presentation(source_path)
    if (src.slide_width, src.slide_height) != (target_prs.slide_width, target_prs.slide_height):
        raise ValueError("슬라이드 크기가 달라 그대로 가져올 수 없습니다.")
    picks = [s for i, s in enumerate(src.slides) if indices is None or i in indices]
    cloner = _Cloner(target_prs.part.package)
    pres_part = target_prs.part
    next_id = _max_id(target_prs)

    # 1) 쓰이는 마스터를 복제하고 프레젠테이션에 등록 (마스터·레이아웃 id 는 새로 매김)
    masters_needed = []
    for s in picks:
        m = s.slide_layout.slide_master.part
        if m not in masters_needed:
            masters_needed.append(m)
    lst = pres_part._element.find(qn("p:sldMasterIdLst"))
    for m in masters_needed:
        new_master = cloner.clone(m)
        for lid in new_master._element.iter(qn("p:sldLayoutId")):
            next_id += 1
            lid.set("id", str(next_id))
        next_id += 1
        rId = pres_part.relate_to(new_master, RT.SLIDE_MASTER)
        el = etree.SubElement(lst, qn("p:sldMasterId"))
        el.set("id", str(next_id))
        el.set(qn("r:id"), rId)

    # 2) 슬라이드 복제 후 순서 목록에 삽입
    sld_lst = target_prs.slides._sldIdLst
    slide_ids = [int(x.get("id")) for x in sld_lst] or [255]
    pos = len(sld_lst) if position is None else position
    for k, s in enumerate(picks):
        new_slide = cloner.clone(s.part)
        rId = pres_part.relate_to(new_slide, RT.SLIDE)
        el = etree.Element(qn("p:sldId"))
        el.set("id", str(max(slide_ids) + 1))
        slide_ids.append(max(slide_ids) + 1)
        el.set(qn("r:id"), rId)
        sld_lst.insert(pos + k, el)
    return len(picks)
