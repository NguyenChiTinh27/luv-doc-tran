"""
models.py
─────────
Data classes and geometrical models for Phase 1 Extractor and document processing.
Represents immutable ground-truth geometric representation extracted from PDF:
- BBox coordinate arithmetic and twips/EMU conversions
- TextSpan: detailed typography flags, origin, color, rotation
- TextBlock: grouped text line/block with bounding box
- GridCellData: table cell representation with span, border, background shading
- TableGridData: table grid topology, cell coordinates, border detection
- DrawingData: vector stroke/fill segments and curve classification
- ImageData: raster images with preserved DPI and dimensions
- ExtractedPage: page-level aggregated primitives
- ExtractedDocument: top-level multi-page container
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict, Any


@dataclass
class BBox:
    """Bounding box in PDF Points (1 pt = 1/72 inch)."""
    x0: float
    y0: float
    x1: float
    y1: float

    def __post_init__(self):
        # Normalize in case coordinates are inverted
        if self.x0 > self.x1:
            self.x0, self.x1 = self.x1, self.x0
        if self.y0 > self.y1:
            self.y0, self.y1 = self.y1, self.y0

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(0.0, self.y1 - self.y0)

    @property
    def area(self) -> float:
        return self.width * self.height

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2.0

    def intersects(self, other: "BBox", tolerance: float = 0.0) -> bool:
        """Check if this bbox intersects with another."""
        return not (
            self.x1 + tolerance < other.x0 or
            self.x0 - tolerance > other.x1 or
            self.y1 + tolerance < other.y0 or
            self.y0 - tolerance > other.y1
        )

    def contains(self, other: "BBox", tolerance: float = 1.0) -> bool:
        """Check if this bbox completely encloses another."""
        return (
            self.x0 - tolerance <= other.x0 and
            self.y0 - tolerance <= other.y0 and
            self.x1 + tolerance >= other.x1 and
            self.y1 + tolerance >= other.y1
        )

    def union(self, other: "BBox") -> "BBox":
        """Return the union bounding box."""
        return BBox(
            x0=min(self.x0, other.x0),
            y0=min(self.y0, other.y0),
            x1=max(self.x1, other.x1),
            y1=max(self.y1, other.y1)
        )

    def intersection(self, other: "BBox") -> Optional["BBox"]:
        """Return intersection bbox if any."""
        ix0 = max(self.x0, other.x0)
        iy0 = max(self.y0, other.y0)
        ix1 = min(self.x1, other.x1)
        iy1 = min(self.y1, other.y1)
        if ix1 > ix0 and iy1 > iy0:
            return BBox(ix0, iy0, ix1, iy1)
        return None

    def iou(self, other: "BBox") -> float:
        """Intersection over union metric."""
        inter = self.intersection(other)
        if not inter:
            return 0.0
        inter_area = inter.area
        union_area = self.area + other.area - inter_area
        return inter_area / union_area if union_area > 0 else 0.0

    def to_twips(self) -> Tuple[int, int, int, int]:
        """Convert points (x0, y0, x1, y1) to twips (1 pt = 20 twips)."""
        return (
            round(self.x0 * 20),
            round(self.y0 * 20),
            round(self.x1 * 20),
            round(self.y1 * 20),
        )

    def to_emus(self) -> Tuple[int, int, int, int]:
        """Convert points (x0, y0, x1, y1) to EMUs (1 pt = 12,700 EMUs)."""
        return (
            round(self.x0 * 12700),
            round(self.y0 * 12700),
            round(self.x1 * 12700),
            round(self.y1 * 12700),
        )

    def to_tuple(self) -> Tuple[float, float, float, float]:
        return (self.x0, self.y0, self.x1, self.y1)

    @classmethod
    def from_tuple(cls, t: Tuple[float, float, float, float]) -> "BBox":
        return cls(t[0], t[1], t[2], t[3])


@dataclass
class TextSpan:
    """Atomic text span extracted from PDF with rich font attributes."""
    text: str
    bbox: Tuple[float, float, float, float]
    font_name: str = "Arial"
    font_size: float = 10.0
    flags: int = 0
    color_hex: str = "000000"
    origin: Tuple[float, float] = (0.0, 0.0)
    rotation: float = 0.0
    is_bold: bool = False
    is_italic: bool = False
    is_monospace: bool = False

    def __post_init__(self):
        # Auto-detect typography flags if flags or font_name indicate bold/italic/mono
        lower_font = self.font_name.lower()
        if bool(self.flags & 16) or "bold" in lower_font or "black" in lower_font or "heavy" in lower_font:
            self.is_bold = True
        if bool(self.flags & 2) or "italic" in lower_font or "oblique" in lower_font:
            self.is_italic = True
        if bool(self.flags & 8) or "courier" in lower_font or "mono" in lower_font or "consolas" in lower_font:
            self.is_monospace = True

    def to_bbox(self) -> BBox:
        return BBox.from_tuple(self.bbox)


@dataclass
class TextBlock:
    """A line or cohesive text block consisting of one or more text spans."""
    spans: List[TextSpan] = field(default_factory=list)
    bbox: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    text: str = ""
    font_name: str = "Arial"
    font_size: float = 10.0
    is_bold: bool = False
    is_italic: bool = False
    is_bullet: bool = False
    bullet_char: str = ""
    line_index: int = 0
    block_type: int = 0  # 0: text, 1: image/other

    def update_from_spans(self) -> None:
        """Recalculate text and bounding box from contained spans."""
        if not self.spans:
            return
        self.text = "".join(s.text for s in self.spans)
        x0 = min(s.bbox[0] for s in self.spans)
        y0 = min(s.bbox[1] for s in self.spans)
        x1 = max(s.bbox[2] for s in self.spans)
        y1 = max(s.bbox[3] for s in self.spans)
        self.bbox = (x0, y0, x1, y1)

        # Dominant font attributes from first non-whitespace span
        for s in self.spans:
            if s.text.strip():
                self.font_name = s.font_name
                self.font_size = s.font_size
                self.is_bold = s.is_bold
                self.is_italic = s.is_italic
                break

        # Check for list bullet
        stripped = self.text.strip()
        if stripped and stripped[0] in {'•', '●', '▪', '▫', '-', '+', '*', '–', '—', '➢', '❖', '✓', '✔'}:
            self.is_bullet = True
            self.bullet_char = stripped[0]

    def to_bbox(self) -> BBox:
        return BBox.from_tuple(self.bbox)


@dataclass
class GridCellData:
    """Table cell representation with bounds, text, and styling."""
    row_idx: int
    col_idx: int
    row_span: int = 1
    col_span: int = 1
    bbox: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    text: str = ""
    blocks: List[TextBlock] = field(default_factory=list)
    bg_color: Optional[str] = None
    borders: Dict[str, Any] = field(default_factory=dict)  # "top", "bottom", "left", "right"

    def to_bbox(self) -> BBox:
        return BBox.from_tuple(self.bbox)


@dataclass
class TableGridData:
    """Table grid topology, cell coordinates, and border metadata."""
    rows: int
    cols: int
    col_widths: List[float] = field(default_factory=list)
    cells: List[GridCellData] = field(default_factory=list)
    bbox: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    is_bordered: bool = True

    def get_cell(self, row: int, col: int) -> Optional[GridCellData]:
        for cell in self.cells:
            if cell.row_idx == row and cell.col_idx == col:
                return cell
        return None

    def to_bbox(self) -> BBox:
        return BBox.from_tuple(self.bbox)


@dataclass
class DrawingData:
    """Vector drawing stroke or fill segment."""
    bbox: Tuple[float, float, float, float]
    stroke_color: Optional[str] = None
    fill_color: Optional[str] = None
    line_width: float = 1.0
    is_horizontal: bool = False
    is_vertical: bool = False
    is_rect: bool = False
    is_curve: bool = False
    items: List[Any] = field(default_factory=list)

    def to_bbox(self) -> BBox:
        return BBox.from_tuple(self.bbox)


@dataclass
class ImageData:
    """Raster image extracted from PDF or rasterized vector cluster."""
    bbox: Tuple[float, float, float, float]
    image_bytes: bytes = b""
    ext: str = "png"
    width: int = 0
    height: int = 0
    dpi: int = 150

    def to_bbox(self) -> BBox:
        return BBox.from_tuple(self.bbox)


@dataclass
class ExtractedPage:
    """Extracted data representation of a single PDF page."""
    page_num: int
    width: float
    height: float
    margins: Tuple[float, float, float, float] = (36.0, 36.0, 36.0, 36.0)  # top, bottom, left, right in pt
    blocks: List[TextBlock] = field(default_factory=list)
    tables: List[TableGridData] = field(default_factory=list)
    drawings: List[DrawingData] = field(default_factory=list)
    images: List[ImageData] = field(default_factory=list)
    rotation: float = 0.0

    @property
    def text(self) -> str:
        """Concatenated plain text of all body blocks on the page."""
        return "\n".join(b.text for b in self.blocks if b.text)


@dataclass
class ExtractedDocument:
    """Top-level multi-page container of extracted PDF content."""
    pages: List[ExtractedPage] = field(default_factory=list)
    pdf_path: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)

    @property
    def page_count(self) -> int:
        return len(self.pages)

    def get_page(self, page_num: int) -> Optional[ExtractedPage]:
        """Get 1-based page number."""
        for p in self.pages:
            if p.page_num == page_num:
                return p
        return None
