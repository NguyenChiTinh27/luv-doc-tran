"""
schema.py
─────────
Data structures and models for the Hybrid Geometry-to-OpenXML PDF conversion engine.
Represents geometrical coordinates, typography metrics, and document layout tree.
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


@dataclass
class SpanData:
    """Text span extracted from PDF."""
    text: str
    bbox: BBox
    font_name: str
    font_size: float
    flags: int = 0
    color_hex: str = "000000"
    origin: Tuple[float, float] = (0.0, 0.0)

    @property
    def is_bold(self) -> bool:
        return bool(self.flags & 16) or "bold" in self.font_name.lower()

    @property
    def is_italic(self) -> bool:
        return bool(self.flags & 2) or "italic" in self.font_name.lower() or "oblique" in self.font_name.lower()

    @property
    def is_monospace(self) -> bool:
        return bool(self.flags & 8) or "courier" in self.font_name.lower() or "mono" in self.font_name.lower()


@dataclass
class LineData:
    """A horizontal text line consisting of one or more spans."""
    spans: List[SpanData] = field(default_factory=list)
    bbox: BBox = field(default_factory=lambda: BBox(0, 0, 0, 0))
    text: str = ""
    is_bullet: bool = False
    bullet_char: str = ""

    def update_text_and_bbox(self) -> None:
        if not self.spans:
            return
        self.text = "".join(s.text for s in self.spans)
        x0 = min(s.bbox.x0 for s in self.spans)
        y0 = min(s.bbox.y0 for s in self.spans)
        x1 = max(s.bbox.x1 for s in self.spans)
        y1 = max(s.bbox.y1 for s in self.spans)
        self.bbox = BBox(x0, y0, x1, y1)


@dataclass
class ParagraphData:
    """A semantic paragraph formed by grouping lines."""
    lines: List[LineData] = field(default_factory=list)
    bbox: BBox = field(default_factory=lambda: BBox(0, 0, 0, 0))
    text: str = ""
    align: str = "left"  # left, center, right, justify
    style_name: str = "Normal"
    line_spacing: float = 1.15
    space_before: float = 0.0
    space_after: float = 2.0
    indent_left: float = 0.0


@dataclass
class BorderStyle:
    """Border styling for Word table cells."""
    val: str = "single"  # single, dashed, double, nil
    sz: int = 4          # 1/8 pt (4 = 0.5 pt)
    space: int = 0
    color: str = "auto"  # HEX color or auto


@dataclass
class GridCellData:
    """Table cell representation with spans and borders."""
    row_idx: int
    col_idx: int
    row_span: int = 1
    col_span: int = 1
    bbox: BBox = field(default_factory=lambda: BBox(0, 0, 0, 0))
    text: str = ""
    lines: List[LineData] = field(default_factory=list)
    paragraphs: List[ParagraphData] = field(default_factory=list)
    bg_color: Optional[str] = None
    borders: Dict[str, BorderStyle] = field(default_factory=dict)  # top, bottom, left, right
    padding: Tuple[float, float, float, float] = (2.0, 2.0, 4.0, 4.0)  # top, bottom, left, right (pt)


@dataclass
class TableData:
    """Table layout structure."""
    rows: int
    cols: int
    col_widths: List[float] = field(default_factory=list)  # in points
    cells: List[GridCellData] = field(default_factory=list)
    bbox: BBox = field(default_factory=lambda: BBox(0, 0, 0, 0))
    is_bordered: bool = True

    def get_cell(self, row: int, col: int) -> Optional[GridCellData]:
        for cell in self.cells:
            if cell.row_idx == row and cell.col_idx == col:
                return cell
        return None


@dataclass
class VectorPathData:
    """Vector drawing stroke or fill segment."""
    bbox: BBox
    stroke_color: Optional[str] = None
    fill_color: Optional[str] = None
    line_width: float = 1.0
    is_horizontal: bool = False
    is_vertical: bool = False
    is_rect: bool = False
    is_curve: bool = False


@dataclass
class VectorCluster:
    """Dense cluster of vector curves representing logos, stamps, or diagrams."""
    bbox: BBox
    path_count: int
    is_dense: bool = False
    rendered_png: Optional[bytes] = None


@dataclass
class ImageData:
    """Raster image extracted from PDF."""
    bbox: BBox
    image_bytes: bytes
    ext: str = "png"
    width: int = 0
    height: int = 0


@dataclass
class PageData:
    """Extracted data representation of a single PDF page."""
    page_num: int
    width: float
    height: float
    margins: Tuple[float, float, float, float] = (36.0, 36.0, 36.0, 36.0)  # top, bottom, left, right in pt
    lines: List[LineData] = field(default_factory=list)
    paragraphs: List[ParagraphData] = field(default_factory=list)
    tables: List[TableData] = field(default_factory=list)
    vector_paths: List[VectorPathData] = field(default_factory=list)
    vector_clusters: List[VectorCluster] = field(default_factory=list)
    images: List[ImageData] = field(default_factory=list)


@dataclass
class ConversionConfig:
    """Runtime engine parameters."""
    engine_preference: str = "auto"       # "auto", "engine_b" (Word COM), "engine_a" (PyMuPDF)
    enable_surgical_pass: bool = True     # Strict OpenXML surgical normalization
    enable_line_wrap_guard: bool = True   # Anti-wordwrap character tracking & scaling
    enable_vector_recovery: bool = True   # High-DPI transparent rasterization for dense vectors
    enable_visual_qa: bool = True         # Headless verification loop
    enable_self_heal: bool = True         # Auto-compacting if page count expands
    target_dpi: int = 150                 # DPI for visual verification
    vector_raster_dpi: int = 300          # DPI for recovered vector graphics
    max_heal_passes: int = 3              # Maximum self-healing iteration passes


@dataclass
class ConversionReport:
    """Diagnostic conversion report."""
    input_file: str
    output_file: str
    success: bool = False
    engine_used: str = "unknown"
    pdf_pages: int = 0
    docx_pages: int = 0
    page_count_match: bool = False
    ssim_score: float = 0.0
    passes_executed: int = 1
    tables_processed: int = 0
    vectors_recovered: int = 0
    execution_time_seconds: float = 0.0
    warnings: List[str] = field(default_factory=list)
