""".pptx / .potx 열기와 템플릿 슬라이드 정리 유틸리티."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from pptx import Presentation

_TEMPLATE_CT = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
_PRES_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"


def open_presentation(path: str | Path):
    """.pptx 는 그대로, .potx 는 content-type 을 바꿔 python-pptx 로 연다."""
    path = Path(path)
    if path.suffix.lower() != ".potx":
        return Presentation(str(path))

    src = zipfile.ZipFile(path)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_TEMPLATE_CT.encode(), _PRES_CT.encode())
            dst.writestr(item, data)
    buf.seek(0)
    return Presentation(buf)


def remove_all_slides(prs) -> None:
    """템플릿 파일에 들어 있는 예시 슬라이드를 모두 삭제한다 (레이아웃/마스터는 유지)."""
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        prs.part.drop_rel(sld_id.rId)
        sld_id_lst.remove(sld_id)
