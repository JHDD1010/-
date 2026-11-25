"""파이프라인 전반에서 쓰는 데이터 모델.

- Source*  : 원본 .pptx 에서 추출한 내용
- Planned* : 템플릿에 배치할 슬라이드 계획 (Claude 또는 단순 변환기가 생성)
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# 원본 추출 결과
# ---------------------------------------------------------------------------


@dataclass
class SourceParagraph:
    text: str
    level: int = 0


@dataclass
class SourceTable:
    rows: list[list[str]]


@dataclass
class SourceImage:
    id: str
    blob: bytes
    ext: str
    width: int  # EMU
    height: int  # EMU
    description: str = ""


@dataclass
class SourceSlide:
    number: int
    title: str = ""
    paragraphs: list[SourceParagraph] = field(default_factory=list)
    tables: list[SourceTable] = field(default_factory=list)
    images: list[SourceImage] = field(default_factory=list)
    notes: str = ""
    other: list[str] = field(default_factory=list)  # 차트 등 옮기지 못한 요소 메모

    def to_prompt_dict(self) -> dict:
        """Claude 에 넘길 텍스트 표현 (이미지 바이너리 제외)."""
        return {
            "number": self.number,
            "title": self.title,
            "paragraphs": [{"text": p.text, "level": p.level} for p in self.paragraphs],
            "tables": [t.rows for t in self.tables],
            "images": [
                {
                    "id": im.id,
                    "description": im.description,
                    "aspect_ratio": round(im.width / im.height, 2) if im.height else None,
                }
                for im in self.images
            ],
            "notes": self.notes,
            "other_elements": self.other,
        }


# ---------------------------------------------------------------------------
# 변환 계획 (Claude structured output 스키마)
# ---------------------------------------------------------------------------


class Bullet(BaseModel):
    text: str = Field(description="글머리 한 줄. 템플릿 문체 규칙을 따른다.")
    level: int = Field(description="들여쓰기 수준. 0이 최상위, 최대 4.")


class Table(BaseModel):
    header: list[str]
    rows: list[list[str]]


class PlannedSlide(BaseModel):
    layout: str = Field(description="템플릿 레이아웃 카탈로그에 있는 레이아웃 이름 그대로")
    title: str
    subtitle: str | None = Field(default=None, description="부제/핵심 메시지(헤드라인). 레이아웃에 부제 영역이 있거나 스타일 가이드가 요구할 때 사용")
    bullets: list[Bullet] = Field(default_factory=list)
    table: Table | None = None
    image_ids: list[str] = Field(default_factory=list, description="원본에서 가져올 이미지 id 목록")
    notes: str = Field(default="", description="발표자 노트")
    source_slides: list[int] = Field(default_factory=list, description="이 슬라이드의 근거가 된 원본 슬라이드 번호")


class DeckPlan(BaseModel):
    slides: list[PlannedSlide]
