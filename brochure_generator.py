#!/usr/bin/env python3
"""
Professional Brochure Generator with Booklet Imposition
Creates 1, 2, 4, or 8 page A5 brochures with proper print layout for saddle-stitch binding.

Features:
- 5 complex templates with unique designs
- Proper booklet imposition (pages arranged for double-sided A4 printing)
- Spread preview (see adjacent pages as they'll appear in final product)
- Greek and English text support
- Many customization options
- Print-ready PDF export
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, font as tkfont
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageFilter, ImageEnhance
import os
import io
import json
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any
from enum import Enum
import math

# Try to import reportlab for PDF generation
try:
    from reportlab.lib.pagesizes import A4, A5
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import HexColor, Color
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False
    print("Warning: reportlab not installed. PDF export will be limited.")

# Constants
A5_WIDTH_MM = 148
A5_HEIGHT_MM = 210
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297

# Screen preview dimensions (pixels)
A5_PREVIEW_WIDTH = 296
A5_PREVIEW_HEIGHT = 420
A4_PREVIEW_WIDTH = 420
A4_PREVIEW_HEIGHT = 594

# High resolution for export (300 DPI)
A5_EXPORT_WIDTH = 1748
A5_EXPORT_HEIGHT = 2480
A4_EXPORT_WIDTH = 2480
A4_EXPORT_HEIGHT = 3508


class PageCount(Enum):
    ONE = 1
    TWO = 2
    FOUR = 4
    EIGHT = 8


@dataclass
class TextElement:
    """Represents a text element on a page"""
    text: str
    x: float  # Percentage of page width
    y: float  # Percentage of page height
    font_family: str = "Arial"
    font_size: int = 24
    font_color: str = "#000000"
    bold: bool = False
    italic: bool = False
    alignment: str = "center"  # left, center, right
    max_width: float = 90  # Percentage of page width
    rotation: float = 0  # Rotation in degrees
    layer_index: int = 0  # For z-ordering (higher = on top)


@dataclass
class ImageElement:
    """Represents an image element on a page"""
    image_path: Optional[str] = None
    image_data: Optional[bytes] = None
    x: float = 50
    y: float = 50
    width: float = 50  # Percentage of page width
    height: float = 50  # Percentage of page height
    opacity: float = 1.0
    rotation: float = 0  # Rotation in degrees
    fit_mode: str = "cover"  # cover, contain, stretch
    layer_index: int = 0  # For z-ordering (higher = on top)


@dataclass
class ShapeElement:
    """Represents a shape element on a page"""
    shape_type: str  # rectangle, circle, line, triangle
    x: float = 50
    y: float = 50
    width: float = 20
    height: float = 20
    fill_color: str = "#3498db"
    stroke_color: str = "#2980b9"
    stroke_width: int = 2
    opacity: float = 1.0
    rotation: float = 0
    layer_index: int = 0  # For z-ordering (higher = on top)


@dataclass
class PageData:
    """Represents a single page with all its elements"""
    background_color: str = "#ffffff"
    background_gradient: Optional[Tuple[str, str, str]] = None  # (color1, color2, direction)
    background_image: Optional[ImageElement] = None
    text_elements: List[TextElement] = field(default_factory=list)
    image_elements: List[ImageElement] = field(default_factory=list)
    shape_elements: List[ShapeElement] = field(default_factory=list)


@dataclass
class BrochureData:
    """Represents the entire brochure"""
    page_count: int = 4
    pages: List[PageData] = field(default_factory=list)
    template_name: str = "blank"

    def __post_init__(self):
        # Initialize pages if empty
        while len(self.pages) < self.page_count:
            self.pages.append(PageData())


class Template:
    """Base class for brochure templates"""

    def __init__(self, name: str, description: str):
        self.name = name
        self.description = description

    def apply(self, brochure: BrochureData) -> BrochureData:
        """Apply template to brochure - override in subclasses"""
        raise NotImplementedError


class CorporateTemplate(Template):
    """Professional corporate/business template"""

    def __init__(self):
        super().__init__(
            "Corporate Professional",
            "Clean, professional design with geometric accents - perfect for business"
        )

    def apply(self, brochure: BrochureData) -> BrochureData:
        primary_color = "#1a365d"  # Dark blue
        accent_color = "#ed8936"   # Orange
        light_bg = "#f7fafc"

        for i, page in enumerate(brochure.pages):
            if i == 0:  # Cover page
                page.background_gradient = (primary_color, "#2d4a77", "vertical")
                # Header accent bar
                page.shape_elements = [
                    ShapeElement("rectangle", 50, 8, 100, 8, accent_color, accent_color, 0, 1.0),
                    ShapeElement("rectangle", 50, 92, 100, 8, accent_color, accent_color, 0, 0.8),
                    # Decorative circles
                    ShapeElement("circle", 85, 20, 15, 15, "#ffffff", "#ffffff", 0, 0.1),
                    ShapeElement("circle", 90, 25, 20, 20, "#ffffff", "#ffffff", 0, 0.05),
                ]
                page.text_elements = [
                    TextElement("COMPANY NAME", 50, 35, "Arial", 36, "#ffffff", True, False, "center"),
                    TextElement("Professional Services", 50, 45, "Arial", 18, accent_color, False, True, "center"),
                    TextElement("Your tagline goes here", 50, 55, "Arial", 14, "#ffffff", False, False, "center"),
                    TextElement("www.example.com", 50, 85, "Arial", 12, "#ffffff", False, False, "center"),
                ]
            elif i == len(brochure.pages) - 1:  # Back cover
                page.background_gradient = (primary_color, "#2d4a77", "vertical")
                page.shape_elements = [
                    ShapeElement("rectangle", 50, 8, 100, 8, accent_color, accent_color, 0, 1.0),
                ]
                page.text_elements = [
                    TextElement("Contact Us", 50, 30, "Arial", 28, "#ffffff", True, False, "center"),
                    TextElement("Location: 123 Business Street", 50, 45, "Arial", 14, "#ffffff", False, False, "center"),
                    TextElement("Phone: +30 210 1234567", 50, 52, "Arial", 14, "#ffffff", False, False, "center"),
                    TextElement("Email: info@example.com", 50, 59, "Arial", 14, "#ffffff", False, False, "center"),
                    TextElement("© 2024 Company Name", 50, 85, "Arial", 10, "#ffffff", False, False, "center"),
                ]
            else:  # Content pages
                page.background_color = light_bg
                page.shape_elements = [
                    ShapeElement("rectangle", 50, 5, 100, 6, primary_color, primary_color, 0, 1.0),
                    ShapeElement("rectangle", 15, 5, 20, 6, accent_color, accent_color, 0, 1.0),
                    # Content divider
                    ShapeElement("rectangle", 50, 50, 60, 0.5, accent_color, accent_color, 0, 0.5),
                ]
                page.text_elements = [
                    TextElement(f"Section {i}", 50, 15, "Arial", 24, primary_color, True, False, "center"),
                    TextElement("Add your content here. This template provides\na clean, professional look for your business.", 50, 35, "Arial", 12, "#4a5568", False, False, "center"),
                    TextElement("Key Features:", 20, 55, "Arial", 16, primary_color, True, False, "left"),
                    TextElement("• Professional design\n• Easy to customize\n• Print-ready format", 20, 70, "Arial", 12, "#4a5568", False, False, "left"),
                ]

        brochure.template_name = self.name
        return brochure


class CreativeTemplate(Template):
    """Bold, creative design with dynamic shapes"""

    def __init__(self):
        super().__init__(
            "Creative Bold",
            "Dynamic design with bold colors and geometric patterns - for creative businesses"
        )

    def apply(self, brochure: BrochureData) -> BrochureData:
        primary_color = "#9b2c2c"   # Deep red
        secondary_color = "#f6e05e" # Yellow
        accent_color = "#2d3748"    # Dark gray

        for i, page in enumerate(brochure.pages):
            if i == 0:  # Cover
                page.background_gradient = (primary_color, "#c53030", "diagonal")
                page.shape_elements = [
                    # Large decorative triangle
                    ShapeElement("triangle", 80, 70, 60, 60, secondary_color, secondary_color, 0, 0.3),
                    ShapeElement("triangle", 20, 30, 40, 40, "#ffffff", "#ffffff", 0, 0.1),
                    # Accent circles
                    ShapeElement("circle", 10, 10, 30, 30, secondary_color, secondary_color, 0, 0.8),
                    ShapeElement("circle", 90, 90, 25, 25, "#ffffff", "#ffffff", 0, 0.2),
                    # Stripe
                    ShapeElement("rectangle", 50, 65, 100, 3, secondary_color, secondary_color, 0, 1.0),
                ]
                page.text_elements = [
                    TextElement("CREATIVE", 50, 30, "Arial", 48, "#ffffff", True, False, "center"),
                    TextElement("STUDIO", 50, 42, "Arial", 48, secondary_color, True, False, "center"),
                    TextElement("Design • Innovation • Excellence", 50, 75, "Arial", 14, "#ffffff", False, True, "center"),
                ]
            elif i == len(brochure.pages) - 1:  # Back
                page.background_gradient = (accent_color, "#1a202c", "vertical")
                page.shape_elements = [
                    ShapeElement("circle", 50, 30, 40, 40, primary_color, primary_color, 0, 0.3),
                    ShapeElement("rectangle", 50, 70, 60, 2, secondary_color, secondary_color, 0, 1.0),
                ]
                page.text_elements = [
                    TextElement("Let's Create Together", 50, 30, "Arial", 24, "#ffffff", True, False, "center"),
                    TextElement("Contact us for your next project", 50, 45, "Arial", 14, "#a0aec0", False, False, "center"),
                    TextElement("hello@creativestudio.com", 50, 75, "Arial", 16, secondary_color, True, False, "center"),
                    TextElement("+30 210 9876543", 50, 82, "Arial", 14, "#ffffff", False, False, "center"),
                ]
            else:  # Content pages
                page.background_color = "#ffffff"
                page.shape_elements = [
                    # Side accent
                    ShapeElement("rectangle", 3, 50, 3, 100, primary_color, primary_color, 0, 1.0),
                    # Top decorative bar
                    ShapeElement("rectangle", 50, 3, 94, 3, secondary_color, secondary_color, 0, 1.0),
                    # Corner accent
                    ShapeElement("triangle", 95, 95, 15, 15, primary_color, primary_color, 0, 0.7),
                ]
                page.text_elements = [
                    TextElement(f"Our Work #{i}", 50, 12, "Arial", 28, primary_color, True, False, "center"),
                    TextElement("We bring your ideas to life with\ninnovative design solutions.", 50, 30, "Arial", 13, accent_color, False, False, "center"),
                    TextElement("Services:", 15, 45, "Arial", 18, primary_color, True, False, "left"),
                    TextElement("- Brand Identity\n- Web Design\n- Print Materials\n- Marketing", 15, 60, "Arial", 12, accent_color, False, False, "left"),
                ]

        brochure.template_name = self.name
        return brochure


class ElegantTemplate(Template):
    """Sophisticated elegant design with subtle gradients"""

    def __init__(self):
        super().__init__(
            "Elegant Luxury",
            "Sophisticated design with gold accents - for premium brands and services"
        )

    def apply(self, brochure: BrochureData) -> BrochureData:
        dark_color = "#1a1a2e"
        gold_color = "#d4a574"
        cream_color = "#faf5f0"

        for i, page in enumerate(brochure.pages):
            if i == 0:  # Cover
                page.background_gradient = (dark_color, "#16213e", "vertical")
                page.shape_elements = [
                    # Elegant border
                    ShapeElement("rectangle", 50, 50, 90, 90, "transparent", gold_color, 2, 0.6),
                    # Corner ornaments
                    ShapeElement("rectangle", 10, 10, 15, 1, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 10, 10, 1, 15, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 90, 10, 15, 1, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 90, 10, 1, 15, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 10, 90, 15, 1, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 10, 90, 1, 15, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 90, 90, 15, 1, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 90, 90, 1, 15, gold_color, gold_color, 0, 1.0),
                    # Center divider
                    ShapeElement("rectangle", 50, 55, 40, 0.5, gold_color, gold_color, 0, 1.0),
                ]
                page.text_elements = [
                    TextElement("*", 50, 25, "Arial", 24, gold_color, False, False, "center"),
                    TextElement("MAISON", 50, 38, "Arial", 42, "#ffffff", True, False, "center"),
                    TextElement("ÉLÉGANCE", 50, 50, "Arial", 28, gold_color, False, True, "center"),
                    TextElement("Luxury Redefined", 50, 65, "Arial", 14, "#a0a0a0", False, False, "center"),
                    TextElement("Est. 2024", 50, 85, "Arial", 12, gold_color, False, False, "center"),
                ]
            elif i == len(brochure.pages) - 1:  # Back
                page.background_gradient = (dark_color, "#16213e", "vertical")
                page.shape_elements = [
                    ShapeElement("rectangle", 50, 50, 90, 90, "transparent", gold_color, 1, 0.4),
                ]
                page.text_elements = [
                    TextElement("Visit Us", 50, 30, "Arial", 22, gold_color, True, False, "center"),
                    TextElement("123 Luxury Avenue\nAthens, Greece", 50, 45, "Arial", 13, "#ffffff", False, False, "center"),
                    TextElement("By Appointment Only", 50, 60, "Arial", 11, "#a0a0a0", False, True, "center"),
                    TextElement("+30 210 555 0000", 50, 75, "Arial", 14, gold_color, False, False, "center"),
                    TextElement("www.maisonelegance.com", 50, 82, "Arial", 11, "#ffffff", False, False, "center"),
                ]
            else:
                page.background_color = cream_color
                page.shape_elements = [
                    ShapeElement("rectangle", 50, 8, 30, 0.5, gold_color, gold_color, 0, 1.0),
                    ShapeElement("rectangle", 50, 92, 30, 0.5, gold_color, gold_color, 0, 1.0),
                    # Subtle corner
                    ShapeElement("rectangle", 5, 5, 8, 0.5, gold_color, gold_color, 0, 0.5),
                    ShapeElement("rectangle", 5, 5, 0.5, 8, gold_color, gold_color, 0, 0.5),
                ]
                page.text_elements = [
                    TextElement(f"Collection {i}", 50, 15, "Arial", 26, dark_color, True, False, "center"),
                    TextElement("Experience the extraordinary", 50, 25, "Arial", 12, gold_color, False, True, "center"),
                    TextElement("Discover our curated selection of\nexquisite pieces, crafted with precision\nand an unwavering commitment to excellence.", 50, 45, "Arial", 11, "#4a4a4a", False, False, "center"),
                    TextElement("*  *  *", 50, 65, "Arial", 14, gold_color, False, False, "center"),
                    TextElement("Quality • Craftsmanship • Heritage", 50, 80, "Arial", 10, dark_color, False, False, "center"),
                ]

        brochure.template_name = self.name
        return brochure


class ModernMinimalTemplate(Template):
    """Clean, modern minimalist design"""

    def __init__(self):
        super().__init__(
            "Modern Minimal",
            "Clean lines and whitespace - for tech companies and modern brands"
        )

    def apply(self, brochure: BrochureData) -> BrochureData:
        black = "#000000"
        white = "#ffffff"
        gray = "#f5f5f5"
        accent = "#00d9ff"  # Cyan

        for i, page in enumerate(brochure.pages):
            if i == 0:  # Cover
                page.background_color = white
                page.shape_elements = [
                    # Minimalist accent
                    ShapeElement("rectangle", 50, 80, 100, 40, black, black, 0, 1.0),
                    ShapeElement("circle", 75, 30, 30, 30, accent, accent, 0, 0.15),
                    ShapeElement("rectangle", 15, 50, 0.5, 30, black, black, 0, 1.0),
                ]
                page.text_elements = [
                    TextElement("NEXT", 50, 30, "Arial", 56, black, True, False, "center"),
                    TextElement("GEN", 50, 45, "Arial", 56, accent, True, False, "center"),
                    TextElement("Technology for Tomorrow", 50, 88, "Arial", 12, white, False, False, "center"),
                ]
            elif i == len(brochure.pages) - 1:  # Back
                page.background_color = black
                page.shape_elements = [
                    ShapeElement("circle", 80, 20, 25, 25, accent, accent, 0, 0.3),
                ]
                page.text_elements = [
                    TextElement("Get Started", 50, 35, "Arial", 28, white, True, False, "center"),
                    TextElement("hello@nextgen.tech", 50, 50, "Arial", 16, accent, False, False, "center"),
                    TextElement("nextgen.tech", 50, 60, "Arial", 14, white, False, False, "center"),
                    TextElement("@nextgentech", 50, 85, "Arial", 12, "#666666", False, False, "center"),
                ]
            else:
                page.background_color = gray
                page.shape_elements = [
                    # Accent line
                    ShapeElement("rectangle", 10, 15, 15, 0.8, accent, accent, 0, 1.0),
                    # Bottom bar
                    ShapeElement("rectangle", 50, 95, 100, 5, black, black, 0, 1.0),
                ]
                page.text_elements = [
                    TextElement(f"0{i}.", 10, 20, "Arial", 36, black, True, False, "left"),
                    TextElement(f"Feature", 30, 20, "Arial", 28, black, False, False, "left"),
                    TextElement("Innovative solutions designed\nfor the modern world.", 50, 40, "Arial", 13, "#333333", False, False, "center"),
                    TextElement("→ Fast\n→ Secure\n→ Scalable", 20, 60, "Arial", 12, black, False, False, "left"),
                    TextElement(f"Page {i}", 90, 95, "Arial", 10, white, False, False, "center"),
                ]

        brochure.template_name = self.name
        return brochure


class NatureOrganicTemplate(Template):
    """Organic, nature-inspired design"""

    def __init__(self):
        super().__init__(
            "Nature Organic",
            "Earthy tones and organic shapes - for eco-friendly and wellness brands"
        )

    def apply(self, brochure: BrochureData) -> BrochureData:
        forest_green = "#2d5016"
        sage = "#9caf88"
        cream = "#f5f2eb"
        earth = "#8b7355"

        for i, page in enumerate(brochure.pages):
            if i == 0:  # Cover
                page.background_gradient = (forest_green, "#3d6b1e", "vertical")
                page.shape_elements = [
                    # Organic circles
                    ShapeElement("circle", 80, 75, 50, 50, sage, sage, 0, 0.3),
                    ShapeElement("circle", 20, 85, 35, 35, cream, cream, 0, 0.15),
                    ShapeElement("circle", 90, 20, 20, 20, "#ffffff", "#ffffff", 0, 0.1),
                    # Leaf-like accent
                    ShapeElement("circle", 50, 25, 8, 8, cream, cream, 0, 0.8),
                ]
                page.text_elements = [
                    TextElement("*", 50, 20, "Arial", 32, cream, False, False, "center"),
                    TextElement("TERRA", 50, 38, "Arial", 44, cream, True, False, "center"),
                    TextElement("BOTANICA", 50, 52, "Arial", 24, sage, False, False, "center"),
                    TextElement("Natural - Organic - Sustainable", 50, 70, "Arial", 11, "#c8d6b9", False, False, "center"),
                ]
            elif i == len(brochure.pages) - 1:  # Back
                page.background_gradient = (earth, "#6b5344", "vertical")
                page.shape_elements = [
                    ShapeElement("circle", 50, 50, 60, 60, cream, cream, 0, 0.1),
                ]
                page.text_elements = [
                    TextElement("Find Us", 50, 30, "Arial", 24, cream, True, False, "center"),
                    TextElement("123 Garden Path\nNature Valley", 50, 45, "Arial", 13, "#ffffff", False, False, "center"),
                    TextElement("Open Daily: 9am - 6pm", 50, 60, "Arial", 11, sage, False, True, "center"),
                    TextElement("www.terrabotanica.eco", 50, 80, "Arial", 13, cream, False, False, "center"),
                ]
            else:
                page.background_color = cream
                page.shape_elements = [
                    # Organic shape top
                    ShapeElement("circle", 90, 5, 20, 20, sage, sage, 0, 0.4),
                    ShapeElement("circle", 10, 95, 15, 15, sage, sage, 0, 0.3),
                    # Divider
                    ShapeElement("rectangle", 50, 50, 50, 0.5, earth, earth, 0, 0.5),
                ]
                page.text_elements = [
                    TextElement("+", 50, 10, "Arial", 24, forest_green, False, False, "center"),
                    TextElement(f"Our Promise #{i}", 50, 20, "Arial", 24, forest_green, True, False, "center"),
                    TextElement("We believe in harmony with nature.\nEvery product is crafted with care\nfor you and the environment.", 50, 38, "Arial", 11, "#5a5a5a", False, False, "center"),
                    TextElement("+ 100% Organic\n+ Sustainably Sourced\n+ Eco-Friendly Packaging", 25, 65, "Arial", 11, forest_green, False, False, "left"),
                ]

        brochure.template_name = self.name
        return brochure


# Template registry
TEMPLATES = {
    "blank": Template("Blank", "Empty template - start from scratch"),
    "corporate": CorporateTemplate(),
    "creative": CreativeTemplate(),
    "elegant": ElegantTemplate(),
    "modern": ModernMinimalTemplate(),
    "nature": NatureOrganicTemplate(),
}


class BrochureRenderer:
    """Renders brochure pages to images"""

    def __init__(self, preview_mode=True):
        self.preview_mode = preview_mode
        self.width = A5_PREVIEW_WIDTH if preview_mode else A5_EXPORT_WIDTH
        self.height = A5_PREVIEW_HEIGHT if preview_mode else A5_EXPORT_HEIGHT

        # Try to load fonts
        self.fonts = {}
        self._load_fonts()

    def _load_fonts(self):
        """Load system fonts"""
        # Common font paths
        font_paths = [
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            "~/.fonts",
            "C:/Windows/Fonts",
            "/Library/Fonts",
            "~/Library/Fonts",
        ]

        # Default to PIL's default font
        self.default_font = ImageFont.load_default()

        # Try to find a good Unicode font for Greek support
        unicode_fonts = [
            "DejaVuSans.ttf",
            "NotoSans-Regular.ttf",
            "Arial.ttf",
            "arial.ttf",
            "FreeSans.ttf",
            "LiberationSans-Regular.ttf",
        ]

        for font_dir in font_paths:
            font_dir = os.path.expanduser(font_dir)
            if os.path.exists(font_dir):
                for root, dirs, files in os.walk(font_dir):
                    for font_file in files:
                        if font_file.endswith(('.ttf', '.TTF', '.otf', '.OTF')):
                            font_path = os.path.join(root, font_file)
                            try:
                                # Test if font supports Greek
                                test_font = ImageFont.truetype(font_path, 12)
                                self.fonts[font_file.lower()] = font_path

                                # Check for preferred Unicode fonts
                                if font_file in unicode_fonts and 'default_unicode' not in self.fonts:
                                    self.fonts['default_unicode'] = font_path
                            except Exception:
                                pass

    def get_font(self, family: str, size: int, bold: bool = False, italic: bool = False):
        """Get a font object"""
        size = int(size * (self.width / A5_PREVIEW_WIDTH))

        # Try to find the requested font
        font_key = family.lower() + ".ttf"

        if font_key in self.fonts:
            try:
                return ImageFont.truetype(self.fonts[font_key], size)
            except Exception:
                pass

        # Fall back to default unicode font
        if 'default_unicode' in self.fonts:
            try:
                return ImageFont.truetype(self.fonts['default_unicode'], size)
            except Exception:
                pass

        # Last resort - default font
        return self.default_font

    def hex_to_rgb(self, hex_color: str) -> Tuple[int, int, int]:
        """Convert hex color to RGB tuple"""
        hex_color = hex_color.lstrip('#')
        if len(hex_color) == 3:
            hex_color = ''.join([c*2 for c in hex_color])
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def hex_to_rgba(self, hex_color: str, alpha: float = 1.0) -> Tuple[int, int, int, int]:
        """Convert hex color to RGBA tuple"""
        rgb = self.hex_to_rgb(hex_color)
        return (*rgb, int(alpha * 255))

    def create_gradient(self, color1: str, color2: str, direction: str) -> Image.Image:
        """Create a gradient background"""
        img = Image.new('RGB', (self.width, self.height))
        draw = ImageDraw.Draw(img)

        r1, g1, b1 = self.hex_to_rgb(color1)
        r2, g2, b2 = self.hex_to_rgb(color2)

        if direction == "vertical":
            for y in range(self.height):
                ratio = y / self.height
                r = int(r1 + (r2 - r1) * ratio)
                g = int(g1 + (g2 - g1) * ratio)
                b = int(b1 + (b2 - b1) * ratio)
                draw.line([(0, y), (self.width, y)], fill=(r, g, b))
        elif direction == "horizontal":
            for x in range(self.width):
                ratio = x / self.width
                r = int(r1 + (r2 - r1) * ratio)
                g = int(g1 + (g2 - g1) * ratio)
                b = int(b1 + (b2 - b1) * ratio)
                draw.line([(x, 0), (x, self.height)], fill=(r, g, b))
        elif direction == "diagonal":
            for y in range(self.height):
                for x in range(self.width):
                    ratio = (x + y) / (self.width + self.height)
                    r = int(r1 + (r2 - r1) * ratio)
                    g = int(g1 + (g2 - g1) * ratio)
                    b = int(b1 + (b2 - b1) * ratio)
                    draw.point((x, y), fill=(r, g, b))

        return img

    def render_page(self, page: PageData) -> Image.Image:
        """Render a single page to an image with layer support"""
        # Create base image with background
        if page.background_gradient:
            img = self.create_gradient(*page.background_gradient)
        else:
            img = Image.new('RGB', (self.width, self.height), self.hex_to_rgb(page.background_color))

        # Convert to RGBA for compositing
        img = img.convert('RGBA')

        # Collect all elements with their layer indices for z-ordering
        all_elements = []

        # Add shapes
        for shape in page.shape_elements:
            layer_idx = getattr(shape, 'layer_index', 0)
            all_elements.append(('shape', shape, layer_idx))

        # Add images
        for image in page.image_elements:
            layer_idx = getattr(image, 'layer_index', 0)
            all_elements.append(('image', image, layer_idx))

        # Add text elements
        for text in page.text_elements:
            layer_idx = getattr(text, 'layer_index', 0)
            all_elements.append(('text', text, layer_idx))

        # Sort by layer index (lower layers rendered first, higher on top)
        all_elements.sort(key=lambda x: x[2])

        # Render each element in order
        for elem_type, elem, _ in all_elements:
            if elem_type == 'shape':
                img = self._render_shape_to_image(img, elem)
            elif elem_type == 'image':
                img = self._render_image_element(img, elem)
            elif elem_type == 'text':
                img = self._render_text_to_image(img, elem)

        return img.convert('RGB')

    def _render_shape_to_image(self, img: Image.Image, shape: ShapeElement) -> Image.Image:
        """Render a shape element with rotation support"""
        x = int(shape.x / 100 * self.width)
        y = int(shape.y / 100 * self.height)
        w = int(shape.width / 100 * self.width)
        h = int(shape.height / 100 * self.height)

        # Create a temporary image for the shape (larger to accommodate rotation)
        padding = max(w, h)
        temp_size = (w + padding * 2, h + padding * 2)
        temp_img = Image.new('RGBA', temp_size, (0, 0, 0, 0))
        temp_draw = ImageDraw.Draw(temp_img)

        if shape.fill_color == "transparent":
            fill = None
        else:
            fill = self.hex_to_rgba(shape.fill_color, shape.opacity)

        outline = self.hex_to_rgba(shape.stroke_color, shape.opacity) if shape.stroke_width > 0 else None

        # Draw shape centered in temp image
        cx, cy = temp_size[0] // 2, temp_size[1] // 2

        if shape.shape_type == "rectangle":
            x1, y1 = cx - w // 2, cy - h // 2
            x2, y2 = cx + w // 2, cy + h // 2
            temp_draw.rectangle([x1, y1, x2, y2], fill=fill, outline=outline, width=shape.stroke_width)
        elif shape.shape_type == "circle":
            x1, y1 = cx - w // 2, cy - h // 2
            x2, y2 = cx + w // 2, cy + h // 2
            temp_draw.ellipse([x1, y1, x2, y2], fill=fill, outline=outline, width=shape.stroke_width)
        elif shape.shape_type == "triangle":
            points = [
                (cx, cy - h // 2),
                (cx - w // 2, cy + h // 2),
                (cx + w // 2, cy + h // 2)
            ]
            temp_draw.polygon(points, fill=fill, outline=outline)
        elif shape.shape_type == "line":
            temp_draw.line([(cx - w // 2, cy), (cx + w // 2, cy)], fill=fill, width=shape.stroke_width)

        # Apply rotation if needed
        rotation = getattr(shape, 'rotation', 0)
        if rotation != 0:
            temp_img = temp_img.rotate(-rotation, resample=Image.Resampling.BICUBIC, expand=False)

        # Paste onto main image
        paste_x = x - temp_size[0] // 2
        paste_y = y - temp_size[1] // 2
        img.paste(temp_img, (paste_x, paste_y), temp_img)

        return img

    def _render_image_element(self, img: Image.Image, image_elem: ImageElement) -> Image.Image:
        """Render an image element with rotation, resize, and opacity support"""
        try:
            # Load image from path or data
            if image_elem.image_path and os.path.exists(image_elem.image_path):
                source_img = Image.open(image_elem.image_path)
            elif image_elem.image_data:
                source_img = Image.open(io.BytesIO(image_elem.image_data))
            else:
                return img

            # Convert to RGBA
            source_img = source_img.convert('RGBA')

            # Calculate target size
            target_w = int(image_elem.width / 100 * self.width)
            target_h = int(image_elem.height / 100 * self.height)

            # Resize based on fit mode
            if image_elem.fit_mode == "stretch":
                source_img = source_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            elif image_elem.fit_mode == "contain":
                source_img.thumbnail((target_w, target_h), Image.Resampling.LANCZOS)
            else:  # cover
                # Scale to cover the target area
                img_ratio = source_img.width / source_img.height
                target_ratio = target_w / target_h
                if img_ratio > target_ratio:
                    new_h = target_h
                    new_w = int(new_h * img_ratio)
                else:
                    new_w = target_w
                    new_h = int(new_w / img_ratio)
                source_img = source_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                # Crop to target size
                left = (source_img.width - target_w) // 2
                top = (source_img.height - target_h) // 2
                source_img = source_img.crop((left, top, left + target_w, top + target_h))

            # Apply opacity
            if image_elem.opacity < 1.0:
                alpha = source_img.split()[3]
                alpha = alpha.point(lambda p: int(p * image_elem.opacity))
                source_img.putalpha(alpha)

            # Apply rotation
            rotation = getattr(image_elem, 'rotation', 0)
            if rotation != 0:
                source_img = source_img.rotate(-rotation, resample=Image.Resampling.BICUBIC, expand=True)

            # Calculate position
            x = int(image_elem.x / 100 * self.width) - source_img.width // 2
            y = int(image_elem.y / 100 * self.height) - source_img.height // 2

            # Paste onto main image
            img.paste(source_img, (x, y), source_img)

        except Exception as e:
            print(f"Error rendering image: {e}")

        return img

    def _render_text_to_image(self, img: Image.Image, text_elem: TextElement) -> Image.Image:
        """Render a text element with rotation support"""
        font = self.get_font(text_elem.font_family, text_elem.font_size, text_elem.bold, text_elem.italic)

        # Calculate position
        x = int(text_elem.x / 100 * self.width)
        y = int(text_elem.y / 100 * self.height)

        # Handle multiline text
        lines = text_elem.text.split('\n')

        # Create temporary draw to measure text
        temp_draw = ImageDraw.Draw(img)

        # Calculate total dimensions
        line_heights = []
        line_widths = []
        for line in lines:
            bbox = temp_draw.textbbox((0, 0), line, font=font)
            line_heights.append(bbox[3] - bbox[1])
            line_widths.append(bbox[2] - bbox[0])

        total_height = sum(line_heights) + (len(lines) - 1) * 5
        max_width = max(line_widths) if line_widths else 0

        # Create temporary image for text (to support rotation)
        padding = 20
        temp_size = (max_width + padding * 2, total_height + padding * 2)
        text_img = Image.new('RGBA', temp_size, (0, 0, 0, 0))
        text_draw = ImageDraw.Draw(text_img)

        fill_color = self.hex_to_rgb(text_elem.font_color)
        current_y = padding

        for i, line in enumerate(lines):
            bbox = text_draw.textbbox((0, 0), line, font=font)
            text_width = bbox[2] - bbox[0]

            if text_elem.alignment == "center":
                text_x = (temp_size[0] - text_width) // 2
            elif text_elem.alignment == "right":
                text_x = temp_size[0] - padding - text_width
            else:  # left
                text_x = padding

            text_draw.text((text_x, current_y), line, font=font, fill=fill_color)
            current_y += line_heights[i] + 5

        # Apply rotation if needed
        rotation = getattr(text_elem, 'rotation', 0)
        if rotation != 0:
            text_img = text_img.rotate(-rotation, resample=Image.Resampling.BICUBIC, expand=True)

        # Calculate paste position
        paste_x = x - text_img.width // 2
        paste_y = y - text_img.height // 2

        # Paste onto main image
        img.paste(text_img, (paste_x, paste_y), text_img)

        return img

    def _render_shape(self, draw: ImageDraw.Draw, shape: ShapeElement):
        """Render a shape element"""
        x = int(shape.x / 100 * self.width)
        y = int(shape.y / 100 * self.height)
        w = int(shape.width / 100 * self.width)
        h = int(shape.height / 100 * self.height)

        if shape.fill_color == "transparent":
            fill = None
        else:
            fill = self.hex_to_rgba(shape.fill_color, shape.opacity)

        outline = self.hex_to_rgba(shape.stroke_color, shape.opacity) if shape.stroke_width > 0 else None

        if shape.shape_type == "rectangle":
            x1, y1 = x - w // 2, y - h // 2
            x2, y2 = x + w // 2, y + h // 2
            draw.rectangle([x1, y1, x2, y2], fill=fill, outline=outline, width=shape.stroke_width)

        elif shape.shape_type == "circle":
            x1, y1 = x - w // 2, y - h // 2
            x2, y2 = x + w // 2, y + h // 2
            draw.ellipse([x1, y1, x2, y2], fill=fill, outline=outline, width=shape.stroke_width)

        elif shape.shape_type == "triangle":
            points = [
                (x, y - h // 2),
                (x - w // 2, y + h // 2),
                (x + w // 2, y + h // 2)
            ]
            draw.polygon(points, fill=fill, outline=outline)

        elif shape.shape_type == "line":
            draw.line([(x - w // 2, y), (x + w // 2, y)], fill=fill, width=shape.stroke_width)

    def _render_text(self, img: Image.Image, text_elem: TextElement):
        """Render a text element"""
        draw = ImageDraw.Draw(img)
        font = self.get_font(text_elem.font_family, text_elem.font_size, text_elem.bold, text_elem.italic)

        # Calculate position
        x = int(text_elem.x / 100 * self.width)
        y = int(text_elem.y / 100 * self.height)
        max_width = int(text_elem.max_width / 100 * self.width)

        # Handle multiline text
        lines = text_elem.text.split('\n')

        # Calculate total height
        line_heights = []
        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            line_heights.append(bbox[3] - bbox[1])

        total_height = sum(line_heights) + (len(lines) - 1) * 5
        current_y = y - total_height // 2

        fill_color = self.hex_to_rgb(text_elem.font_color)

        for i, line in enumerate(lines):
            bbox = draw.textbbox((0, 0), line, font=font)
            text_width = bbox[2] - bbox[0]

            if text_elem.alignment == "center":
                text_x = x - text_width // 2
            elif text_elem.alignment == "right":
                text_x = x - text_width
            else:  # left
                text_x = x

            draw.text((text_x, current_y), line, font=font, fill=fill_color)
            current_y += line_heights[i] + 5


class BookletImposition:
    """
    Handles booklet imposition for saddle-stitch printing.

    For an 8-page A5 brochure printed on A4:
    - Sheet 1, Front: Pages 8, 1 (left, right)
    - Sheet 1, Back: Pages 2, 7 (left, right)
    - Sheet 2, Front: Pages 6, 3 (left, right)
    - Sheet 2, Back: Pages 4, 5 (left, right)

    The sheets are nested (sheet 2 inside sheet 1) and folded.
    """

    @staticmethod
    def get_imposition_order(page_count: int) -> List[Tuple[int, int]]:
        """
        Get the page pairs for each print side.
        Returns list of (left_page, right_page) tuples.
        For page numbers, 0 = first page (cover), page_count-1 = last page (back)
        """
        if page_count == 1:
            return [(0, -1)]  # -1 means blank
        elif page_count == 2:
            return [(1, 0)]  # Back, Front on same side
        elif page_count == 4:
            # Sheet 1 Front: 4, 1 -> indices 3, 0
            # Sheet 1 Back: 2, 3 -> indices 1, 2
            return [
                (3, 0),  # Sheet 1 Front
                (1, 2),  # Sheet 1 Back
            ]
        elif page_count == 8:
            # Sheet 1 Front: 8, 1 -> indices 7, 0
            # Sheet 1 Back: 2, 7 -> indices 1, 6
            # Sheet 2 Front: 6, 3 -> indices 5, 2
            # Sheet 2 Back: 4, 5 -> indices 3, 4
            return [
                (7, 0),  # Sheet 1 Front
                (1, 6),  # Sheet 1 Back
                (5, 2),  # Sheet 2 Front
                (3, 4),  # Sheet 2 Back
            ]
        else:
            # For other counts, just pair sequentially
            pairs = []
            for i in range(0, page_count, 2):
                if i + 1 < page_count:
                    pairs.append((i, i + 1))
                else:
                    pairs.append((i, -1))
            return pairs

    @staticmethod
    def get_spread_pairs(page_count: int) -> List[Tuple[int, int, str]]:
        """
        Get page pairs as they appear in the final folded product.
        Returns list of (left_page, right_page, description) tuples.
        """
        if page_count == 1:
            return [(0, -1, "Single Page")]
        elif page_count == 2:
            return [(0, 1, "Front & Back")]
        elif page_count == 4:
            return [
                (3, 0, "Back Cover & Front Cover"),
                (1, 2, "Inside Spread"),
            ]
        elif page_count == 8:
            return [
                (7, 0, "Back Cover & Front Cover"),
                (1, 2, "First Inside Spread"),
                (3, 4, "Center Spread"),
                (5, 6, "Last Inside Spread"),
            ]
        else:
            pairs = []
            for i in range(0, page_count, 2):
                if i + 1 < page_count:
                    pairs.append((i, i + 1, f"Pages {i+1}-{i+2}"))
                else:
                    pairs.append((i, -1, f"Page {i+1}"))
            return pairs


class RealTimeColorPicker:
    """Custom color picker with real-time preview callback"""

    def __init__(self, parent, initial_color="#ffffff", callback=None, title="Choose Color"):
        self.callback = callback
        self.result = None
        self.current_color = initial_color

        # Create popup window
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("400x350")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        # Parse initial color
        self.h, self.s, self.v = self._hex_to_hsv(initial_color)

        self._create_ui()
        self._update_preview()

    def _hex_to_hsv(self, hex_color):
        """Convert hex color to HSV"""
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) / 255.0 for i in (0, 2, 4))

        max_c = max(r, g, b)
        min_c = min(r, g, b)
        diff = max_c - min_c

        if diff == 0:
            h = 0
        elif max_c == r:
            h = (60 * ((g - b) / diff) + 360) % 360
        elif max_c == g:
            h = (60 * ((b - r) / diff) + 120) % 360
        else:
            h = (60 * ((r - g) / diff) + 240) % 360

        s = 0 if max_c == 0 else diff / max_c
        v = max_c

        return h, s, v

    def _hsv_to_hex(self, h, s, v):
        """Convert HSV to hex color"""
        c = v * s
        x = c * (1 - abs((h / 60) % 2 - 1))
        m = v - c

        if h < 60:
            r, g, b = c, x, 0
        elif h < 120:
            r, g, b = x, c, 0
        elif h < 180:
            r, g, b = 0, c, x
        elif h < 240:
            r, g, b = 0, x, c
        elif h < 300:
            r, g, b = x, 0, c
        else:
            r, g, b = c, 0, x

        r, g, b = int((r + m) * 255), int((g + m) * 255), int((b + m) * 255)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _create_ui(self):
        """Create the color picker UI"""
        main_frame = tk.Frame(self.window, bg='#2d2d2d')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Hue slider
        tk.Label(main_frame, text="Hue:", bg='#2d2d2d', fg='white').pack(anchor=tk.W)
        self.hue_var = tk.DoubleVar(value=self.h)
        self.hue_scale = tk.Scale(main_frame, from_=0, to=359, orient=tk.HORIZONTAL,
                                   variable=self.hue_var, command=self._on_color_change,
                                   bg='#2d2d2d', fg='white', highlightthickness=0,
                                   troughcolor='#444444', length=350)
        self.hue_scale.pack(fill=tk.X)

        # Saturation slider
        tk.Label(main_frame, text="Saturation:", bg='#2d2d2d', fg='white').pack(anchor=tk.W)
        self.sat_var = tk.DoubleVar(value=self.s * 100)
        self.sat_scale = tk.Scale(main_frame, from_=0, to=100, orient=tk.HORIZONTAL,
                                   variable=self.sat_var, command=self._on_color_change,
                                   bg='#2d2d2d', fg='white', highlightthickness=0,
                                   troughcolor='#444444', length=350)
        self.sat_scale.pack(fill=tk.X)

        # Value/Brightness slider
        tk.Label(main_frame, text="Brightness:", bg='#2d2d2d', fg='white').pack(anchor=tk.W)
        self.val_var = tk.DoubleVar(value=self.v * 100)
        self.val_scale = tk.Scale(main_frame, from_=0, to=100, orient=tk.HORIZONTAL,
                                   variable=self.val_var, command=self._on_color_change,
                                   bg='#2d2d2d', fg='white', highlightthickness=0,
                                   troughcolor='#444444', length=350)
        self.val_scale.pack(fill=tk.X)

        # Color preview
        preview_frame = tk.Frame(main_frame, bg='#2d2d2d')
        preview_frame.pack(fill=tk.X, pady=10)

        tk.Label(preview_frame, text="Preview:", bg='#2d2d2d', fg='white').pack(side=tk.LEFT)
        self.preview_label = tk.Label(preview_frame, width=10, height=2, bg=self.current_color)
        self.preview_label.pack(side=tk.LEFT, padx=10)

        self.hex_label = tk.Label(preview_frame, text=self.current_color, bg='#2d2d2d', fg='white')
        self.hex_label.pack(side=tk.LEFT)

        # Hex entry
        hex_entry_frame = tk.Frame(main_frame, bg='#2d2d2d')
        hex_entry_frame.pack(fill=tk.X, pady=5)

        tk.Label(hex_entry_frame, text="Hex:", bg='#2d2d2d', fg='white').pack(side=tk.LEFT)
        self.hex_entry = tk.Entry(hex_entry_frame, width=10, bg='#444444', fg='white')
        self.hex_entry.insert(0, self.current_color)
        self.hex_entry.pack(side=tk.LEFT, padx=5)
        self.hex_entry.bind('<Return>', self._on_hex_entry)

        # Buttons
        button_frame = tk.Frame(main_frame, bg='#2d2d2d')
        button_frame.pack(fill=tk.X, pady=10)

        tk.Button(button_frame, text="OK", command=self._on_ok,
                 bg='#4CAF50', fg='white', width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel,
                 bg='#f44336', fg='white', width=10).pack(side=tk.LEFT, padx=5)

    def _on_color_change(self, *args):
        """Handle color change from sliders"""
        self.h = self.hue_var.get()
        self.s = self.sat_var.get() / 100
        self.v = self.val_var.get() / 100
        self._update_preview()

        # Call real-time callback
        if self.callback:
            self.callback(self.current_color)

    def _update_preview(self):
        """Update color preview"""
        self.current_color = self._hsv_to_hex(self.h, self.s, self.v)
        self.preview_label.configure(bg=self.current_color)
        self.hex_label.configure(text=self.current_color)
        self.hex_entry.delete(0, tk.END)
        self.hex_entry.insert(0, self.current_color)

    def _on_hex_entry(self, event):
        """Handle hex entry"""
        hex_val = self.hex_entry.get().strip()
        if not hex_val.startswith('#'):
            hex_val = '#' + hex_val
        if len(hex_val) == 7:
            try:
                self.h, self.s, self.v = self._hex_to_hsv(hex_val)
                self.hue_var.set(self.h)
                self.sat_var.set(self.s * 100)
                self.val_var.set(self.v * 100)
                self._update_preview()
                if self.callback:
                    self.callback(self.current_color)
            except ValueError:
                pass

    def _on_ok(self):
        """Handle OK button"""
        self.result = self.current_color
        self.window.destroy()

    def _on_cancel(self):
        """Handle Cancel button"""
        self.result = None
        self.window.destroy()

    def get_color(self):
        """Wait for dialog and return selected color"""
        self.window.wait_window()
        return self.result


class BrochureGeneratorApp:
    """Main application class"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Professional Brochure Generator - Booklet Imposition")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 700)

        # Application state
        self.brochure = BrochureData(page_count=4)
        self.current_page_index = 0
        self.selected_element = None
        self.preview_mode = "single"  # single, spread, print
        self.renderer = BrochureRenderer(preview_mode=True)

        # Initialize pages
        self._initialize_pages()

        # Setup UI
        self._setup_styles()
        self._create_ui()
        self._update_preview()

    def _initialize_pages(self):
        """Initialize brochure pages"""
        self.brochure.pages = [PageData() for _ in range(self.brochure.page_count)]

    def _setup_styles(self):
        """Setup ttk styles"""
        style = ttk.Style()
        style.theme_use('clam')

        # Colors
        self.colors = {
            'bg_dark': '#1a1a2e',
            'bg_medium': '#16213e',
            'bg_light': '#0f3460',
            'accent': '#e94560',
            'accent_hover': '#ff6b6b',
            'text': '#eaeaea',
            'text_dim': '#a0a0a0',
            'success': '#4ade80',
            'warning': '#fbbf24',
        }

        self.root.configure(bg=self.colors['bg_dark'])

        # Configure styles
        style.configure('Dark.TFrame', background=self.colors['bg_dark'])
        style.configure('Medium.TFrame', background=self.colors['bg_medium'])
        style.configure('Light.TFrame', background=self.colors['bg_light'])

        style.configure('Dark.TLabel',
                       background=self.colors['bg_dark'],
                       foreground=self.colors['text'],
                       font=('Segoe UI', 10))

        style.configure('Title.TLabel',
                       background=self.colors['bg_dark'],
                       foreground=self.colors['text'],
                       font=('Segoe UI', 14, 'bold'))

        style.configure('Accent.TButton',
                       background=self.colors['accent'],
                       foreground='white',
                       font=('Segoe UI', 10, 'bold'),
                       padding=(10, 5))

        style.configure('Dark.TButton',
                       background=self.colors['bg_light'],
                       foreground=self.colors['text'],
                       font=('Segoe UI', 10),
                       padding=(10, 5))

        style.configure('Dark.TCombobox',
                       background=self.colors['bg_light'],
                       foreground=self.colors['text'],
                       fieldbackground=self.colors['bg_light'])

        style.configure('Dark.TEntry',
                       fieldbackground=self.colors['bg_light'],
                       foreground=self.colors['text'])

        style.configure('Dark.TSpinbox',
                       fieldbackground=self.colors['bg_light'],
                       foreground=self.colors['text'])

        style.configure("Dark.TNotebook", background=self.colors['bg_dark'])
        style.configure("Dark.TNotebook.Tab",
                       background=self.colors['bg_medium'],
                       foreground=self.colors['text'],
                       padding=(10, 5))
        style.map("Dark.TNotebook.Tab",
                 background=[('selected', self.colors['accent'])],
                 foreground=[('selected', 'white')])

    def _create_ui(self):
        """Create the main UI"""
        # Main container
        main_frame = ttk.Frame(self.root, style='Dark.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left sidebar - Controls
        self._create_left_sidebar(main_frame)

        # Center - Preview area
        self._create_preview_area(main_frame)

        # Right sidebar - Element properties
        self._create_right_sidebar(main_frame)

    def _create_left_sidebar(self, parent):
        """Create the left sidebar with controls"""
        sidebar = ttk.Frame(parent, style='Dark.TFrame', width=280)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        sidebar.pack_propagate(False)

        # Title
        title_label = ttk.Label(sidebar, text="📖 Brochure Generator",
                               style='Title.TLabel')
        title_label.pack(pady=(10, 15))

        # Notebook for tabs
        notebook = ttk.Notebook(sidebar, style='Dark.TNotebook')
        notebook.pack(fill=tk.BOTH, expand=True, padx=5)

        # Tab 1: Setup
        setup_frame = ttk.Frame(notebook, style='Dark.TFrame')
        notebook.add(setup_frame, text='  Setup  ')
        self._create_setup_tab(setup_frame)

        # Tab 2: Templates
        templates_frame = ttk.Frame(notebook, style='Dark.TFrame')
        notebook.add(templates_frame, text='  Templates  ')
        self._create_templates_tab(templates_frame)

        # Tab 3: Elements
        elements_frame = ttk.Frame(notebook, style='Dark.TFrame')
        notebook.add(elements_frame, text='  Elements  ')
        self._create_elements_tab(elements_frame)

        # Export section at bottom
        export_frame = ttk.Frame(sidebar, style='Dark.TFrame')
        export_frame.pack(fill=tk.X, pady=10, padx=5)

        ttk.Label(export_frame, text="Export", style='Title.TLabel').pack(anchor=tk.W)

        export_btn = tk.Button(export_frame, text="📄 Export Print-Ready PDF",
                              bg=self.colors['success'], fg='white',
                              font=('Segoe UI', 10, 'bold'),
                              command=self._export_pdf, cursor='hand2')
        export_btn.pack(fill=tk.X, pady=5)

        export_img_btn = tk.Button(export_frame, text="🖼️ Export Pages as Images",
                                   bg=self.colors['bg_light'], fg='white',
                                   font=('Segoe UI', 10),
                                   command=self._export_images, cursor='hand2')
        export_img_btn.pack(fill=tk.X, pady=2)

    def _create_setup_tab(self, parent):
        """Create the setup tab content"""
        # Page count selection
        ttk.Label(parent, text="Number of Pages:", style='Dark.TLabel').pack(anchor=tk.W, pady=(10, 5), padx=10)

        page_frame = ttk.Frame(parent, style='Dark.TFrame')
        page_frame.pack(fill=tk.X, padx=10)

        self.page_count_var = tk.IntVar(value=4)
        for count in [1, 2, 4, 8]:
            rb = tk.Radiobutton(page_frame, text=f"{count} page{'s' if count > 1 else ''}",
                               variable=self.page_count_var, value=count,
                               bg=self.colors['bg_dark'], fg=self.colors['text'],
                               selectcolor=self.colors['bg_light'],
                               activebackground=self.colors['bg_dark'],
                               activeforeground=self.colors['accent'],
                               command=self._on_page_count_change)
            rb.pack(anchor=tk.W)

        # Page size info
        ttk.Label(parent, text="\nPage Size: A5 (148×210mm)", style='Dark.TLabel').pack(anchor=tk.W, padx=10)
        ttk.Label(parent, text="Print Size: A4 (210×297mm)", style='Dark.TLabel').pack(anchor=tk.W, padx=10)

        # Info about imposition
        info_frame = ttk.Frame(parent, style='Medium.TFrame')
        info_frame.pack(fill=tk.X, padx=10, pady=15)

        info_text = """ℹ️ Booklet Imposition Info:

Pages are arranged for
double-sided A4 printing.

After printing, nest the sheets
and fold in half to create your
A5 booklet with pages in order.

8 pages = 2 A4 sheets
4 pages = 1 A4 sheet
2 pages = 1 A4 side
1 page = Half A4"""

        info_label = tk.Label(info_frame, text=info_text,
                             bg=self.colors['bg_medium'], fg=self.colors['text_dim'],
                             font=('Segoe UI', 9), justify=tk.LEFT, padx=10, pady=10)
        info_label.pack(fill=tk.X)

    def _create_templates_tab(self, parent):
        """Create the templates tab content"""
        ttk.Label(parent, text="Select Template:", style='Dark.TLabel').pack(anchor=tk.W, pady=(10, 5), padx=10)

        # Scrollable template list
        canvas = tk.Canvas(parent, bg=self.colors['bg_dark'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Dark.TFrame')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Template buttons
        for key, template in TEMPLATES.items():
            btn_frame = ttk.Frame(scrollable_frame, style='Medium.TFrame')
            btn_frame.pack(fill=tk.X, pady=3)

            btn = tk.Button(btn_frame, text=template.name,
                           bg=self.colors['bg_light'], fg='white',
                           font=('Segoe UI', 10, 'bold'),
                           anchor='w', padx=10, cursor='hand2',
                           command=lambda k=key: self._apply_template(k))
            btn.pack(fill=tk.X)

            desc = tk.Label(btn_frame, text=template.description,
                           bg=self.colors['bg_medium'], fg=self.colors['text_dim'],
                           font=('Segoe UI', 8), anchor='w', padx=10, wraplength=230)
            desc.pack(fill=tk.X)

    def _create_elements_tab(self, parent):
        """Create the elements tab content"""
        ttk.Label(parent, text="Add Elements:", style='Dark.TLabel').pack(anchor=tk.W, pady=(10, 5), padx=10)

        # Add text button
        add_text_btn = tk.Button(parent, text="[T] Add Text",
                                bg=self.colors['bg_light'], fg='white',
                                font=('Segoe UI', 10),
                                command=self._add_text_element, cursor='hand2')
        add_text_btn.pack(fill=tk.X, padx=10, pady=3)

        # Add image button
        add_img_btn = tk.Button(parent, text="[+] Import Image",
                               bg=self.colors['accent'], fg='white',
                               font=('Segoe UI', 10, 'bold'),
                               command=self._import_image, cursor='hand2')
        add_img_btn.pack(fill=tk.X, padx=10, pady=3)

        # Add shape buttons
        shapes_frame = ttk.Frame(parent, style='Dark.TFrame')
        shapes_frame.pack(fill=tk.X, padx=10, pady=5)

        for shape, label in [("rectangle", "[=] Rectangle"), ("circle", "[O] Circle"), ("triangle", "[^] Triangle")]:
            btn = tk.Button(shapes_frame, text=label,
                           bg=self.colors['bg_light'], fg='white',
                           font=('Segoe UI', 9),
                           command=lambda s=shape: self._add_shape_element(s), cursor='hand2')
            btn.pack(fill=tk.X, pady=2)

        # Page background
        ttk.Label(parent, text="\nPage Background:", style='Dark.TLabel').pack(anchor=tk.W, padx=10)

        bg_color_btn = tk.Button(parent, text="🎨 Background Color",
                                bg=self.colors['bg_light'], fg='white',
                                font=('Segoe UI', 10),
                                command=self._change_bg_color, cursor='hand2')
        bg_color_btn.pack(fill=tk.X, padx=10, pady=3)

        # Gradient options
        ttk.Label(parent, text="Gradient:", style='Dark.TLabel').pack(anchor=tk.W, padx=10, pady=(10, 5))

        grad_frame = ttk.Frame(parent, style='Dark.TFrame')
        grad_frame.pack(fill=tk.X, padx=10)

        self.grad_color1_btn = tk.Button(grad_frame, text="Color 1", width=10,
                                        bg='#3498db', fg='white',
                                        command=lambda: self._change_grad_color(1), cursor='hand2')
        self.grad_color1_btn.pack(side=tk.LEFT, padx=2)

        self.grad_color2_btn = tk.Button(grad_frame, text="Color 2", width=10,
                                        bg='#9b59b6', fg='white',
                                        command=lambda: self._change_grad_color(2), cursor='hand2')
        self.grad_color2_btn.pack(side=tk.LEFT, padx=2)

        self.grad_dir_var = tk.StringVar(value="vertical")
        grad_dir_combo = ttk.Combobox(parent, textvariable=self.grad_dir_var,
                                     values=["vertical", "horizontal", "diagonal"],
                                     state="readonly", width=15)
        grad_dir_combo.pack(padx=10, pady=5)
        grad_dir_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_gradient())

        apply_grad_btn = tk.Button(parent, text="Apply Gradient",
                                  bg=self.colors['accent'], fg='white',
                                  command=self._apply_gradient, cursor='hand2')
        apply_grad_btn.pack(fill=tk.X, padx=10, pady=3)

        clear_grad_btn = tk.Button(parent, text="Clear Gradient",
                                  bg=self.colors['bg_light'], fg='white',
                                  command=self._clear_gradient, cursor='hand2')
        clear_grad_btn.pack(fill=tk.X, padx=10, pady=3)

    def _create_preview_area(self, parent):
        """Create the central preview area"""
        preview_container = ttk.Frame(parent, style='Dark.TFrame')
        preview_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Preview mode selector
        mode_frame = ttk.Frame(preview_container, style='Dark.TFrame')
        mode_frame.pack(fill=tk.X, pady=5)

        ttk.Label(mode_frame, text="Preview Mode:", style='Dark.TLabel').pack(side=tk.LEFT, padx=10)

        self.preview_mode_var = tk.StringVar(value="single")
        modes = [("Single Page", "single"), ("Spread View", "spread"), ("Print Layout", "print")]

        for text, value in modes:
            rb = tk.Radiobutton(mode_frame, text=text, variable=self.preview_mode_var,
                               value=value, bg=self.colors['bg_dark'], fg=self.colors['text'],
                               selectcolor=self.colors['bg_light'],
                               activebackground=self.colors['bg_dark'],
                               command=self._on_preview_mode_change)
            rb.pack(side=tk.LEFT, padx=10)

        # Page navigation
        nav_frame = ttk.Frame(preview_container, style='Dark.TFrame')
        nav_frame.pack(fill=tk.X, pady=5)

        self.prev_btn = tk.Button(nav_frame, text="◀ Previous",
                                 bg=self.colors['bg_light'], fg='white',
                                 command=self._prev_page, cursor='hand2')
        self.prev_btn.pack(side=tk.LEFT, padx=10)

        self.page_label = ttk.Label(nav_frame, text="Page 1 of 4", style='Dark.TLabel')
        self.page_label.pack(side=tk.LEFT, expand=True)

        self.next_btn = tk.Button(nav_frame, text="Next ▶",
                                 bg=self.colors['bg_light'], fg='white',
                                 command=self._next_page, cursor='hand2')
        self.next_btn.pack(side=tk.RIGHT, padx=10)

        # Canvas for preview
        self.preview_canvas = tk.Canvas(preview_container,
                                        bg=self.colors['bg_medium'],
                                        highlightthickness=1,
                                        highlightbackground=self.colors['bg_light'])
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Bind canvas events
        self.preview_canvas.bind('<Configure>', self._on_canvas_resize)
        self.preview_canvas.bind('<Button-1>', self._on_canvas_click)
        self.preview_canvas.bind('<B1-Motion>', self._on_canvas_drag)

        # Current page info label
        self.info_label = ttk.Label(preview_container, text="", style='Dark.TLabel')
        self.info_label.pack(pady=5)

    def _create_right_sidebar(self, parent):
        """Create the right sidebar for element properties"""
        sidebar = ttk.Frame(parent, style='Dark.TFrame', width=280)
        sidebar.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        sidebar.pack_propagate(False)

        ttk.Label(sidebar, text="Element Properties", style='Title.TLabel').pack(pady=(10, 15))

        # Properties frame (populated when element selected)
        self.properties_frame = ttk.Frame(sidebar, style='Dark.TFrame')
        self.properties_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        # Default message
        self.no_selection_label = ttk.Label(self.properties_frame,
                                            text="Select an element\nto edit its properties",
                                            style='Dark.TLabel')
        self.no_selection_label.pack(pady=50)

        # Element list with page indicator
        self.elements_header_frame = ttk.Frame(sidebar, style='Dark.TFrame')
        self.elements_header_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        self.elements_header_label = tk.Label(self.elements_header_frame,
                                               text="Page Elements for ",
                                               bg=self.colors['bg_dark'],
                                               fg=self.colors['text'],
                                               font=('Segoe UI', 10))
        self.elements_header_label.pack(side=tk.LEFT)

        self.page_number_label = tk.Label(self.elements_header_frame,
                                           text=f"Page {self.current_page_index + 1}",
                                           bg=self.colors['bg_dark'],
                                           fg=self.colors['accent'],
                                           font=('Segoe UI', 10, 'bold'))
        self.page_number_label.pack(side=tk.LEFT)

        self.elements_listbox = tk.Listbox(sidebar, bg=self.colors['bg_light'],
                                           fg=self.colors['text'],
                                           selectbackground=self.colors['accent'],
                                           height=8)
        self.elements_listbox.pack(fill=tk.X, padx=10, pady=5)
        self.elements_listbox.bind('<<ListboxSelect>>', self._on_element_select)

        # Delete button
        delete_btn = tk.Button(sidebar, text="🗑️ Delete Selected",
                              bg=self.colors['accent'], fg='white',
                              command=self._delete_selected_element, cursor='hand2')
        delete_btn.pack(fill=tk.X, padx=10, pady=5)

    def _on_page_count_change(self):
        """Handle page count change"""
        new_count = self.page_count_var.get()
        self.brochure.page_count = new_count

        # Adjust pages list
        while len(self.brochure.pages) < new_count:
            self.brochure.pages.append(PageData())
        while len(self.brochure.pages) > new_count:
            self.brochure.pages.pop()

        # Reset to first page
        self.current_page_index = 0
        self._update_preview()
        self._update_page_label()
        self._update_elements_list()

    def _on_preview_mode_change(self):
        """Handle preview mode change"""
        self.preview_mode = self.preview_mode_var.get()
        self._update_preview()
        self._update_page_label()

    def _apply_template(self, template_key: str):
        """Apply a template to the brochure"""
        if template_key == "blank":
            # Reset to blank
            self._initialize_pages()
        else:
            template = TEMPLATES[template_key]
            self.brochure = template.apply(self.brochure)

        self._update_preview()
        self._update_elements_list()

    def _add_text_element(self):
        """Add a new text element to the current page"""
        page = self.brochure.pages[self.current_page_index]

        new_text = TextElement(
            text="New Text",
            x=50, y=50,
            font_family="Arial",
            font_size=24,
            font_color="#000000"
        )
        page.text_elements.append(new_text)

        self._update_preview()
        self._update_elements_list()

    def _add_shape_element(self, shape_type: str):
        """Add a new shape element to the current page"""
        page = self.brochure.pages[self.current_page_index]

        new_shape = ShapeElement(
            shape_type=shape_type,
            x=50, y=50,
            width=20, height=20,
            fill_color="#3498db",
            stroke_color="#2980b9",
            stroke_width=2
        )
        page.shape_elements.append(new_shape)

        self._update_preview()
        self._update_elements_list()

    def _change_bg_color(self):
        """Change the background color of the current page with real-time preview"""
        page = self.brochure.pages[self.current_page_index]
        original_color = page.background_color
        original_gradient = page.background_gradient

        def on_color_change(color):
            page.background_color = color
            page.background_gradient = None
            self._update_preview()

        picker = RealTimeColorPicker(self.root, original_color, on_color_change, "Background Color")
        result = picker.get_color()

        if result is None:
            # User cancelled - revert to original
            page.background_color = original_color
            page.background_gradient = original_gradient
            self._update_preview()

    def _change_grad_color(self, which: int):
        """Change gradient color with real-time preview"""
        current = self.grad_color1_btn.cget('bg') if which == 1 else self.grad_color2_btn.cget('bg')
        page = self.brochure.pages[self.current_page_index]

        def on_color_change(color):
            if which == 1:
                self.grad_color1_btn.configure(bg=color)
            else:
                self.grad_color2_btn.configure(bg=color)
            # Apply gradient preview
            color1 = self.grad_color1_btn.cget('bg')
            color2 = self.grad_color2_btn.cget('bg')
            direction = self.grad_dir_var.get()
            page.background_gradient = (color1, color2, direction)
            self._update_preview()

        picker = RealTimeColorPicker(self.root, current, on_color_change, f"Gradient Color {which}")
        picker.get_color()

    def _apply_gradient(self):
        """Apply gradient to current page"""
        page = self.brochure.pages[self.current_page_index]
        color1 = self.grad_color1_btn.cget('bg')
        color2 = self.grad_color2_btn.cget('bg')
        direction = self.grad_dir_var.get()

        page.background_gradient = (color1, color2, direction)
        self._update_preview()

    def _clear_gradient(self):
        """Clear gradient from current page"""
        page = self.brochure.pages[self.current_page_index]
        page.background_gradient = None
        self._update_preview()

    def _prev_page(self):
        """Navigate to previous page"""
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self._update_preview()
            self._update_page_label()
            self._update_elements_list()

    def _next_page(self):
        """Navigate to next page"""
        if self.current_page_index < self.brochure.page_count - 1:
            self.current_page_index += 1
            self._update_preview()
            self._update_page_label()
            self._update_elements_list()

    def _update_page_label(self):
        """Update the page navigation label"""
        if self.preview_mode == "single":
            self.page_label.configure(
                text=f"Page {self.current_page_index + 1} of {self.brochure.page_count}")
        elif self.preview_mode == "spread":
            spreads = BookletImposition.get_spread_pairs(self.brochure.page_count)
            spread_idx = self.current_page_index // 2
            if spread_idx < len(spreads):
                self.page_label.configure(text=spreads[spread_idx][2])
            else:
                self.page_label.configure(text="Spread View")
        else:  # print
            self.page_label.configure(text="Print Layout (A4 sheets)")

    def _update_elements_list(self):
        """Update the elements listbox"""
        self.elements_listbox.delete(0, tk.END)

        # Update page number label
        self.page_number_label.configure(text=f"Page {self.current_page_index + 1}")

        page = self.brochure.pages[self.current_page_index]

        # Collect all elements with layer info for display
        elements_display = []

        for i, shape in enumerate(page.shape_elements):
            layer = getattr(shape, 'layer_index', 0)
            elements_display.append(('shape', i, f"[L{layer}] {shape.shape_type.title()} {i+1}", layer))

        for i, img in enumerate(page.image_elements):
            layer = getattr(img, 'layer_index', 0)
            name = os.path.basename(img.image_path) if img.image_path else f"Image {i+1}"
            name = name[:15] + "..." if len(name) > 15 else name
            elements_display.append(('image', i, f"[L{layer}] IMG: {name}", layer))

        for i, text in enumerate(page.text_elements):
            layer = getattr(text, 'layer_index', 0)
            preview = text.text[:15] + "..." if len(text.text) > 15 else text.text
            preview = preview.replace('\n', ' ')
            elements_display.append(('text', i, f"[L{layer}] TXT: {preview}", layer))

        # Sort by layer for display
        elements_display.sort(key=lambda x: x[3])

        for elem_type, idx, display_text, _ in elements_display:
            self.elements_listbox.insert(tk.END, display_text)

    def _on_element_select(self, event):
        """Handle element selection from listbox"""
        selection = self.elements_listbox.curselection()
        if not selection:
            return

        idx = selection[0]
        page = self.brochure.pages[self.current_page_index]

        # Rebuild elements_display to find which element was selected
        elements_display = []

        for i, shape in enumerate(page.shape_elements):
            layer = getattr(shape, 'layer_index', 0)
            elements_display.append(('shape', i, layer))

        for i, img in enumerate(page.image_elements):
            layer = getattr(img, 'layer_index', 0)
            elements_display.append(('image', i, layer))

        for i, text in enumerate(page.text_elements):
            layer = getattr(text, 'layer_index', 0)
            elements_display.append(('text', i, layer))

        # Sort by layer (same order as displayed)
        elements_display.sort(key=lambda x: x[2])

        if idx < len(elements_display):
            elem_type, elem_idx, _ = elements_display[idx]
            self.selected_element = (elem_type, elem_idx)

            if elem_type == 'shape':
                self._show_shape_properties(page.shape_elements[elem_idx])
            elif elem_type == 'image':
                self._show_image_properties(page.image_elements[elem_idx])
            elif elem_type == 'text':
                self._show_text_properties(page.text_elements[elem_idx])

        self._update_preview()

    def _show_text_properties(self, text_elem: TextElement):
        """Show text element properties in the right sidebar"""
        # Clear existing properties
        for widget in self.properties_frame.winfo_children():
            widget.destroy()

        # Text content
        ttk.Label(self.properties_frame, text="Text:", style='Dark.TLabel').pack(anchor=tk.W, pady=(0, 5))

        text_entry = tk.Text(self.properties_frame, height=4, width=25,
                            bg=self.colors['bg_light'], fg=self.colors['text'],
                            insertbackground=self.colors['text'])
        text_entry.insert('1.0', text_elem.text)
        text_entry.pack(fill=tk.X, pady=(0, 10))

        def update_text(*args):
            text_elem.text = text_entry.get('1.0', 'end-1c')
            self._update_preview()

        text_entry.bind('<KeyRelease>', update_text)

        # Font size
        ttk.Label(self.properties_frame, text="Font Size:", style='Dark.TLabel').pack(anchor=tk.W)

        size_var = tk.IntVar(value=text_elem.font_size)
        size_spin = tk.Spinbox(self.properties_frame, from_=8, to=72, textvariable=size_var,
                              width=10, bg=self.colors['bg_light'], fg=self.colors['text'])
        size_spin.pack(anchor=tk.W, pady=(0, 10))

        def update_size(*args):
            text_elem.font_size = size_var.get()
            self._update_preview()

        size_var.trace('w', update_size)

        # Color with real-time preview
        ttk.Label(self.properties_frame, text="Color:", style='Dark.TLabel').pack(anchor=tk.W)

        color_btn = tk.Button(self.properties_frame, text="Choose Color",
                             bg=text_elem.font_color, fg='white',
                             cursor='hand2')
        color_btn.pack(anchor=tk.W, pady=(0, 10))

        def change_color():
            def on_color_change(color):
                text_elem.font_color = color
                color_btn.configure(bg=color)
                self._update_preview()

            picker = RealTimeColorPicker(self.root, text_elem.font_color, on_color_change, "Text Color")
            picker.get_color()

        color_btn.configure(command=change_color)

        # Position
        ttk.Label(self.properties_frame, text="Position (%):", style='Dark.TLabel').pack(anchor=tk.W)

        pos_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        pos_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(pos_frame, text="X:", style='Dark.TLabel').pack(side=tk.LEFT)
        x_var = tk.IntVar(value=int(text_elem.x))
        x_spin = tk.Spinbox(pos_frame, from_=0, to=100, textvariable=x_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        x_spin.pack(side=tk.LEFT, padx=5)

        ttk.Label(pos_frame, text="Y:", style='Dark.TLabel').pack(side=tk.LEFT)
        y_var = tk.IntVar(value=int(text_elem.y))
        y_spin = tk.Spinbox(pos_frame, from_=0, to=100, textvariable=y_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        y_spin.pack(side=tk.LEFT, padx=5)

        def update_pos(*args):
            text_elem.x = x_var.get()
            text_elem.y = y_var.get()
            self._update_preview()

        x_var.trace('w', update_pos)
        y_var.trace('w', update_pos)

        # Alignment
        ttk.Label(self.properties_frame, text="Alignment:", style='Dark.TLabel').pack(anchor=tk.W)

        align_var = tk.StringVar(value=text_elem.alignment)
        align_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        align_frame.pack(fill=tk.X, pady=(0, 10))

        for align in ["left", "center", "right"]:
            rb = tk.Radiobutton(align_frame, text=align.title(), variable=align_var,
                               value=align, bg=self.colors['bg_dark'], fg=self.colors['text'],
                               selectcolor=self.colors['bg_light'])
            rb.pack(side=tk.LEFT)

        def update_align(*args):
            text_elem.alignment = align_var.get()
            self._update_preview()

        align_var.trace('w', update_align)

        # Rotation
        ttk.Label(self.properties_frame, text="Rotation (degrees):", style='Dark.TLabel').pack(anchor=tk.W)

        rotation_var = tk.DoubleVar(value=getattr(text_elem, 'rotation', 0))
        rotation_scale = tk.Scale(self.properties_frame, from_=0, to=360, resolution=1,
                                 orient=tk.HORIZONTAL, variable=rotation_var,
                                 bg=self.colors['bg_dark'], fg=self.colors['text'],
                                 highlightthickness=0, troughcolor=self.colors['bg_light'])
        rotation_scale.pack(fill=tk.X, pady=(0, 10))

        def update_rotation(*args):
            text_elem.rotation = rotation_var.get()
            self._update_preview()

        rotation_var.trace('w', update_rotation)

        # Layer
        ttk.Label(self.properties_frame, text="Layer (z-order):", style='Dark.TLabel').pack(anchor=tk.W)

        layer_var = tk.IntVar(value=getattr(text_elem, 'layer_index', 0))
        layer_spin = tk.Spinbox(self.properties_frame, from_=-100, to=100, textvariable=layer_var, width=10,
                               bg=self.colors['bg_light'], fg=self.colors['text'])
        layer_spin.pack(anchor=tk.W, pady=(0, 10))

        def update_layer(*args):
            text_elem.layer_index = layer_var.get()
            self._update_preview()
            self._update_elements_list()

        layer_var.trace('w', update_layer)

    def _show_shape_properties(self, shape: ShapeElement):
        """Show shape element properties in the right sidebar"""
        # Clear existing properties
        for widget in self.properties_frame.winfo_children():
            widget.destroy()

        ttk.Label(self.properties_frame, text=f"{shape.shape_type.title()} Properties",
                 style='Title.TLabel').pack(pady=(0, 10))

        # Fill color with real-time preview
        ttk.Label(self.properties_frame, text="Fill Color:", style='Dark.TLabel').pack(anchor=tk.W)

        fill_btn = tk.Button(self.properties_frame, text="Choose Fill",
                            bg=shape.fill_color if shape.fill_color != "transparent" else "#888888",
                            fg='white', cursor='hand2')
        fill_btn.pack(anchor=tk.W, pady=(0, 10))

        def change_fill():
            def on_color_change(color):
                shape.fill_color = color
                fill_btn.configure(bg=color)
                self._update_preview()

            initial = shape.fill_color if shape.fill_color != "transparent" else "#ffffff"
            picker = RealTimeColorPicker(self.root, initial, on_color_change, "Fill Color")
            result = picker.get_color()
            if result is None:
                # User cancelled, revert (already applied during preview)
                pass

        fill_btn.configure(command=change_fill)

        # Position
        ttk.Label(self.properties_frame, text="Position (%):", style='Dark.TLabel').pack(anchor=tk.W)

        pos_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        pos_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(pos_frame, text="X:", style='Dark.TLabel').pack(side=tk.LEFT)
        x_var = tk.IntVar(value=int(shape.x))
        x_spin = tk.Spinbox(pos_frame, from_=0, to=100, textvariable=x_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        x_spin.pack(side=tk.LEFT, padx=5)

        ttk.Label(pos_frame, text="Y:", style='Dark.TLabel').pack(side=tk.LEFT)
        y_var = tk.IntVar(value=int(shape.y))
        y_spin = tk.Spinbox(pos_frame, from_=0, to=100, textvariable=y_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        y_spin.pack(side=tk.LEFT, padx=5)

        def update_pos(*args):
            shape.x = x_var.get()
            shape.y = y_var.get()
            self._update_preview()

        x_var.trace('w', update_pos)
        y_var.trace('w', update_pos)

        # Size
        ttk.Label(self.properties_frame, text="Size (%):", style='Dark.TLabel').pack(anchor=tk.W)

        size_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        size_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(size_frame, text="W:", style='Dark.TLabel').pack(side=tk.LEFT)
        w_var = tk.IntVar(value=int(shape.width))
        w_spin = tk.Spinbox(size_frame, from_=1, to=100, textvariable=w_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        w_spin.pack(side=tk.LEFT, padx=5)

        ttk.Label(size_frame, text="H:", style='Dark.TLabel').pack(side=tk.LEFT)
        h_var = tk.IntVar(value=int(shape.height))
        h_spin = tk.Spinbox(size_frame, from_=1, to=100, textvariable=h_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        h_spin.pack(side=tk.LEFT, padx=5)

        def update_size(*args):
            shape.width = w_var.get()
            shape.height = h_var.get()
            self._update_preview()

        w_var.trace('w', update_size)
        h_var.trace('w', update_size)

        # Rotation
        ttk.Label(self.properties_frame, text="Rotation (degrees):", style='Dark.TLabel').pack(anchor=tk.W)

        rotation_var = tk.DoubleVar(value=getattr(shape, 'rotation', 0))
        rotation_scale = tk.Scale(self.properties_frame, from_=0, to=360, resolution=1,
                                 orient=tk.HORIZONTAL, variable=rotation_var,
                                 bg=self.colors['bg_dark'], fg=self.colors['text'],
                                 highlightthickness=0, troughcolor=self.colors['bg_light'])
        rotation_scale.pack(fill=tk.X, pady=(0, 10))

        def update_rotation(*args):
            shape.rotation = rotation_var.get()
            self._update_preview()

        rotation_var.trace('w', update_rotation)

        # Opacity
        ttk.Label(self.properties_frame, text="Opacity:", style='Dark.TLabel').pack(anchor=tk.W)

        opacity_var = tk.DoubleVar(value=shape.opacity)
        opacity_scale = tk.Scale(self.properties_frame, from_=0, to=1, resolution=0.1,
                                orient=tk.HORIZONTAL, variable=opacity_var,
                                bg=self.colors['bg_dark'], fg=self.colors['text'],
                                highlightthickness=0, troughcolor=self.colors['bg_light'])
        opacity_scale.pack(fill=tk.X, pady=(0, 10))

        def update_opacity(*args):
            shape.opacity = opacity_var.get()
            self._update_preview()

        opacity_var.trace('w', update_opacity)

        # Layer
        ttk.Label(self.properties_frame, text="Layer (z-order):", style='Dark.TLabel').pack(anchor=tk.W)

        layer_var = tk.IntVar(value=getattr(shape, 'layer_index', 0))
        layer_spin = tk.Spinbox(self.properties_frame, from_=-100, to=100, textvariable=layer_var, width=10,
                               bg=self.colors['bg_light'], fg=self.colors['text'])
        layer_spin.pack(anchor=tk.W, pady=(0, 10))

        def update_layer(*args):
            shape.layer_index = layer_var.get()
            self._update_preview()
            self._update_elements_list()

        layer_var.trace('w', update_layer)

    def _import_image(self):
        """Import an image file and add it to the current page"""
        filetypes = [
            ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"),
            ("PNG files", "*.png"),
            ("JPEG files", "*.jpg *.jpeg"),
            ("All files", "*.*")
        ]

        filepath = filedialog.askopenfilename(
            title="Select Image",
            filetypes=filetypes
        )

        if filepath:
            # Create image element
            page = self.brochure.pages[self.current_page_index]

            # Get max layer index to place new image on top
            max_layer = 0
            for elem in page.shape_elements + page.text_elements + page.image_elements:
                layer = getattr(elem, 'layer_index', 0)
                max_layer = max(max_layer, layer)

            new_image = ImageElement(
                image_path=filepath,
                x=50,
                y=50,
                width=30,
                height=30,
                opacity=1.0,
                rotation=0,
                fit_mode="contain",
                layer_index=max_layer + 1
            )
            page.image_elements.append(new_image)

            self._update_preview()
            self._update_elements_list()

    def _show_image_properties(self, image_elem: ImageElement):
        """Show image element properties in the right sidebar"""
        # Clear existing properties
        for widget in self.properties_frame.winfo_children():
            widget.destroy()

        name = os.path.basename(image_elem.image_path) if image_elem.image_path else "Image"
        ttk.Label(self.properties_frame, text=f"Image: {name[:20]}",
                 style='Title.TLabel').pack(pady=(0, 10))

        # Position
        ttk.Label(self.properties_frame, text="Position (%):", style='Dark.TLabel').pack(anchor=tk.W)

        pos_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        pos_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(pos_frame, text="X:", style='Dark.TLabel').pack(side=tk.LEFT)
        x_var = tk.IntVar(value=int(image_elem.x))
        x_spin = tk.Spinbox(pos_frame, from_=0, to=100, textvariable=x_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        x_spin.pack(side=tk.LEFT, padx=5)

        ttk.Label(pos_frame, text="Y:", style='Dark.TLabel').pack(side=tk.LEFT)
        y_var = tk.IntVar(value=int(image_elem.y))
        y_spin = tk.Spinbox(pos_frame, from_=0, to=100, textvariable=y_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        y_spin.pack(side=tk.LEFT, padx=5)

        def update_pos(*args):
            image_elem.x = x_var.get()
            image_elem.y = y_var.get()
            self._update_preview()

        x_var.trace('w', update_pos)
        y_var.trace('w', update_pos)

        # Size
        ttk.Label(self.properties_frame, text="Size (%):", style='Dark.TLabel').pack(anchor=tk.W)

        size_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        size_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(size_frame, text="W:", style='Dark.TLabel').pack(side=tk.LEFT)
        w_var = tk.IntVar(value=int(image_elem.width))
        w_spin = tk.Spinbox(size_frame, from_=1, to=100, textvariable=w_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        w_spin.pack(side=tk.LEFT, padx=5)

        ttk.Label(size_frame, text="H:", style='Dark.TLabel').pack(side=tk.LEFT)
        h_var = tk.IntVar(value=int(image_elem.height))
        h_spin = tk.Spinbox(size_frame, from_=1, to=100, textvariable=h_var, width=5,
                           bg=self.colors['bg_light'], fg=self.colors['text'])
        h_spin.pack(side=tk.LEFT, padx=5)

        def update_size(*args):
            image_elem.width = w_var.get()
            image_elem.height = h_var.get()
            self._update_preview()

        w_var.trace('w', update_size)
        h_var.trace('w', update_size)

        # Rotation
        ttk.Label(self.properties_frame, text="Rotation (degrees):", style='Dark.TLabel').pack(anchor=tk.W)

        rotation_var = tk.DoubleVar(value=getattr(image_elem, 'rotation', 0))
        rotation_scale = tk.Scale(self.properties_frame, from_=0, to=360, resolution=1,
                                 orient=tk.HORIZONTAL, variable=rotation_var,
                                 bg=self.colors['bg_dark'], fg=self.colors['text'],
                                 highlightthickness=0, troughcolor=self.colors['bg_light'])
        rotation_scale.pack(fill=tk.X, pady=(0, 10))

        def update_rotation(*args):
            image_elem.rotation = rotation_var.get()
            self._update_preview()

        rotation_var.trace('w', update_rotation)

        # Opacity
        ttk.Label(self.properties_frame, text="Opacity:", style='Dark.TLabel').pack(anchor=tk.W)

        opacity_var = tk.DoubleVar(value=image_elem.opacity)
        opacity_scale = tk.Scale(self.properties_frame, from_=0, to=1, resolution=0.1,
                                orient=tk.HORIZONTAL, variable=opacity_var,
                                bg=self.colors['bg_dark'], fg=self.colors['text'],
                                highlightthickness=0, troughcolor=self.colors['bg_light'])
        opacity_scale.pack(fill=tk.X, pady=(0, 10))

        def update_opacity(*args):
            image_elem.opacity = opacity_var.get()
            self._update_preview()

        opacity_var.trace('w', update_opacity)

        # Fit mode
        ttk.Label(self.properties_frame, text="Fit Mode:", style='Dark.TLabel').pack(anchor=tk.W)

        fit_var = tk.StringVar(value=image_elem.fit_mode)
        fit_frame = ttk.Frame(self.properties_frame, style='Dark.TFrame')
        fit_frame.pack(fill=tk.X, pady=(0, 10))

        for mode in ["contain", "cover", "stretch"]:
            rb = tk.Radiobutton(fit_frame, text=mode.title(), variable=fit_var,
                               value=mode, bg=self.colors['bg_dark'], fg=self.colors['text'],
                               selectcolor=self.colors['bg_light'])
            rb.pack(side=tk.LEFT)

        def update_fit(*args):
            image_elem.fit_mode = fit_var.get()
            self._update_preview()

        fit_var.trace('w', update_fit)

        # Layer
        ttk.Label(self.properties_frame, text="Layer (z-order):", style='Dark.TLabel').pack(anchor=tk.W)

        layer_var = tk.IntVar(value=getattr(image_elem, 'layer_index', 0))
        layer_spin = tk.Spinbox(self.properties_frame, from_=-100, to=100, textvariable=layer_var, width=10,
                               bg=self.colors['bg_light'], fg=self.colors['text'])
        layer_spin.pack(anchor=tk.W, pady=(0, 10))

        def update_layer(*args):
            image_elem.layer_index = layer_var.get()
            self._update_preview()
            self._update_elements_list()

        layer_var.trace('w', update_layer)

        # Replace image button
        replace_btn = tk.Button(self.properties_frame, text="Replace Image",
                               bg=self.colors['bg_light'], fg='white',
                               cursor='hand2')
        replace_btn.pack(fill=tk.X, pady=5)

        def replace_image():
            filetypes = [
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"),
                ("All files", "*.*")
            ]
            filepath = filedialog.askopenfilename(title="Select Image", filetypes=filetypes)
            if filepath:
                image_elem.image_path = filepath
                image_elem.image_data = None
                self._update_preview()
                self._update_elements_list()

        replace_btn.configure(command=replace_image)

    def _delete_selected_element(self):
        """Delete the selected element"""
        if not self.selected_element:
            return

        page = self.brochure.pages[self.current_page_index]
        elem_type, idx = self.selected_element

        if elem_type == 'shape' and idx < len(page.shape_elements):
            page.shape_elements.pop(idx)
        elif elem_type == 'image' and idx < len(page.image_elements):
            page.image_elements.pop(idx)
        elif elem_type == 'text' and idx < len(page.text_elements):
            page.text_elements.pop(idx)

        self.selected_element = None
        self._update_preview()
        self._update_elements_list()

        # Reset properties panel
        for widget in self.properties_frame.winfo_children():
            widget.destroy()
        self.no_selection_label = ttk.Label(self.properties_frame,
                                            text="Select an element\nto edit its properties",
                                            style='Dark.TLabel')
        self.no_selection_label.pack(pady=50)

    def _on_canvas_resize(self, event):
        """Handle canvas resize"""
        self._update_preview()

    def _on_canvas_click(self, event):
        """Handle canvas click for element selection"""
        pass  # TODO: Implement click-to-select

    def _on_canvas_drag(self, event):
        """Handle canvas drag for element movement"""
        pass  # TODO: Implement drag-to-move

    def _update_preview(self):
        """Update the preview canvas"""
        self.preview_canvas.delete('all')

        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()

        if canvas_width <= 1 or canvas_height <= 1:
            return

        if self.preview_mode == "single":
            self._render_single_page_preview(canvas_width, canvas_height)
        elif self.preview_mode == "spread":
            self._render_spread_preview(canvas_width, canvas_height)
        else:  # print
            self._render_print_preview(canvas_width, canvas_height)

    def _render_single_page_preview(self, canvas_width, canvas_height):
        """Render single page preview"""
        if self.current_page_index >= len(self.brochure.pages):
            return

        page = self.brochure.pages[self.current_page_index]

        # Calculate preview size to fit canvas with padding
        padding = 40
        max_width = canvas_width - padding * 2
        max_height = canvas_height - padding * 2

        # Maintain A5 aspect ratio
        aspect_ratio = A5_PREVIEW_WIDTH / A5_PREVIEW_HEIGHT

        if max_width / max_height > aspect_ratio:
            preview_height = max_height
            preview_width = int(preview_height * aspect_ratio)
        else:
            preview_width = max_width
            preview_height = int(preview_width / aspect_ratio)

        # Render page
        self.renderer.width = preview_width
        self.renderer.height = preview_height

        img = self.renderer.render_page(page)

        # Center on canvas
        x = (canvas_width - preview_width) // 2
        y = (canvas_height - preview_height) // 2

        # Convert to PhotoImage and display
        self.preview_image = ImageTk.PhotoImage(img)
        self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.preview_image)

        # Draw border
        self.preview_canvas.create_rectangle(x, y, x + preview_width, y + preview_height,
                                            outline=self.colors['text_dim'], width=1)

        # Page number indicator
        self.info_label.configure(
            text=f"Page {self.current_page_index + 1} - A5 ({A5_WIDTH_MM}×{A5_HEIGHT_MM}mm)")

    def _render_spread_preview(self, canvas_width, canvas_height):
        """Render spread (two facing pages) preview"""
        spreads = BookletImposition.get_spread_pairs(self.brochure.page_count)
        spread_idx = min(self.current_page_index // 2, len(spreads) - 1)

        if spread_idx >= len(spreads):
            return

        left_idx, right_idx, description = spreads[spread_idx]

        # Calculate preview size
        padding = 40
        gap = 10  # Gap between pages
        max_width = canvas_width - padding * 2
        max_height = canvas_height - padding * 2

        # Two A5 pages side by side
        aspect_ratio = (A5_PREVIEW_WIDTH * 2 + gap) / A5_PREVIEW_HEIGHT

        if max_width / max_height > aspect_ratio:
            preview_height = max_height
            total_width = int(preview_height * aspect_ratio)
        else:
            total_width = max_width
            preview_height = int(total_width / aspect_ratio)

        page_width = (total_width - gap) // 2

        self.renderer.width = page_width
        self.renderer.height = preview_height

        # Starting position
        x_start = (canvas_width - total_width) // 2
        y = (canvas_height - preview_height) // 2

        # Render left page
        if left_idx >= 0 and left_idx < len(self.brochure.pages):
            left_img = self.renderer.render_page(self.brochure.pages[left_idx])
            self.spread_left_image = ImageTk.PhotoImage(left_img)
            self.preview_canvas.create_image(x_start, y, anchor=tk.NW, image=self.spread_left_image)
            self.preview_canvas.create_rectangle(x_start, y, x_start + page_width, y + preview_height,
                                                outline=self.colors['text_dim'], width=1)
            # Page label
            self.preview_canvas.create_text(x_start + page_width // 2, y + preview_height + 15,
                                           text=f"Page {left_idx + 1}", fill=self.colors['text_dim'])
        else:
            # Blank page
            self.preview_canvas.create_rectangle(x_start, y, x_start + page_width, y + preview_height,
                                                fill=self.colors['bg_light'],
                                                outline=self.colors['text_dim'], width=1)

        # Render right page
        right_x = x_start + page_width + gap
        if right_idx >= 0 and right_idx < len(self.brochure.pages):
            right_img = self.renderer.render_page(self.brochure.pages[right_idx])
            self.spread_right_image = ImageTk.PhotoImage(right_img)
            self.preview_canvas.create_image(right_x, y, anchor=tk.NW, image=self.spread_right_image)
            self.preview_canvas.create_rectangle(right_x, y, right_x + page_width, y + preview_height,
                                                outline=self.colors['text_dim'], width=1)
            # Page label
            self.preview_canvas.create_text(right_x + page_width // 2, y + preview_height + 15,
                                           text=f"Page {right_idx + 1}", fill=self.colors['text_dim'])
        else:
            # Blank page
            self.preview_canvas.create_rectangle(right_x, y, right_x + page_width, y + preview_height,
                                                fill=self.colors['bg_light'],
                                                outline=self.colors['text_dim'], width=1)

        # Center fold line
        fold_x = x_start + page_width + gap // 2
        self.preview_canvas.create_line(fold_x, y - 10, fold_x, y + preview_height + 10,
                                        fill=self.colors['accent'], dash=(4, 4))

        self.info_label.configure(text=f"Spread: {description}")

    def _render_print_preview(self, canvas_width, canvas_height):
        """Render print layout preview (A4 sheets with imposed pages)"""
        imposition = BookletImposition.get_imposition_order(self.brochure.page_count)

        # Calculate preview size for A4 (landscape for two A5 side by side)
        padding = 40
        max_width = canvas_width - padding * 2
        max_height = canvas_height - padding * 2

        # A4 landscape aspect ratio
        aspect_ratio = A4_HEIGHT_MM / A4_WIDTH_MM  # 297/210 in landscape

        if max_width / max_height > aspect_ratio:
            sheet_height = max_height
            sheet_width = int(sheet_height * aspect_ratio)
        else:
            sheet_width = max_width
            sheet_height = int(sheet_width / aspect_ratio)

        # Each A5 page takes half the A4 width
        page_width = sheet_width // 2
        page_height = sheet_height

        self.renderer.width = page_width
        self.renderer.height = page_height

        # Center position
        x_start = (canvas_width - sheet_width) // 2
        y = (canvas_height - sheet_height) // 2

        # Draw A4 sheet background
        self.preview_canvas.create_rectangle(x_start, y, x_start + sheet_width, y + sheet_height,
                                            fill='white', outline=self.colors['text_dim'], width=2)

        # Get current sheet (based on current page index)
        sheet_idx = min(self.current_page_index // 2, (len(imposition) - 1) // 2) * 2
        if sheet_idx < len(imposition):
            left_idx, right_idx = imposition[sheet_idx]

            # Render left half
            if left_idx >= 0 and left_idx < len(self.brochure.pages):
                left_img = self.renderer.render_page(self.brochure.pages[left_idx])
                self.print_left_image = ImageTk.PhotoImage(left_img)
                self.preview_canvas.create_image(x_start, y, anchor=tk.NW, image=self.print_left_image)
                self.preview_canvas.create_text(x_start + page_width // 2, y + page_height + 15,
                                               text=f"Page {left_idx + 1}", fill=self.colors['text'])

            # Render right half
            if right_idx >= 0 and right_idx < len(self.brochure.pages):
                right_img = self.renderer.render_page(self.brochure.pages[right_idx])
                self.print_right_image = ImageTk.PhotoImage(right_img)
                self.preview_canvas.create_image(x_start + page_width, y, anchor=tk.NW, image=self.print_right_image)
                self.preview_canvas.create_text(x_start + page_width + page_width // 2, y + page_height + 15,
                                               text=f"Page {right_idx + 1}", fill=self.colors['text'])

        # Center fold line
        fold_x = x_start + page_width
        self.preview_canvas.create_line(fold_x, y, fold_x, y + page_height,
                                        fill=self.colors['accent'], dash=(4, 4), width=2)

        # Sheet info
        total_sheets = (len(imposition) + 1) // 2
        current_sheet = sheet_idx // 2 + 1
        side = "Front" if sheet_idx % 2 == 0 else "Back"

        self.info_label.configure(
            text=f"A4 Sheet {current_sheet}/{total_sheets} ({side}) - Print double-sided, flip on short edge")

    def _export_pdf(self):
        """Export brochure as print-ready PDF with imposition"""
        if not HAS_REPORTLAB:
            messagebox.showerror("Error",
                "PDF export requires the 'reportlab' library.\n"
                "Install it with: pip install reportlab")
            return

        # Ask for save location
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Print-Ready PDF"
        )

        if not filename:
            return

        try:
            self._generate_pdf(filename)
            messagebox.showinfo("Success", f"PDF saved to:\n{filename}\n\n"
                              "Print double-sided, flip on short edge.\n"
                              "Then nest sheets and fold to create booklet.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF:\n{str(e)}")

    def _generate_pdf(self, filename: str):
        """Generate the PDF with proper imposition"""
        # Use high resolution renderer
        export_renderer = BrochureRenderer(preview_mode=False)

        # Create PDF with A4 landscape pages
        c = canvas.Canvas(filename, pagesize=(A4[1], A4[0]))  # Landscape A4

        imposition = BookletImposition.get_imposition_order(self.brochure.page_count)

        for i, (left_idx, right_idx) in enumerate(imposition):
            # Render pages at export resolution
            page_width = A5_EXPORT_WIDTH
            page_height = A5_EXPORT_HEIGHT

            export_renderer.width = page_width
            export_renderer.height = page_height

            # Left page
            if left_idx >= 0 and left_idx < len(self.brochure.pages):
                left_img = export_renderer.render_page(self.brochure.pages[left_idx])
                left_buffer = io.BytesIO()
                left_img.save(left_buffer, format='PNG')
                left_buffer.seek(0)
                left_reader = ImageReader(left_buffer)
                c.drawImage(left_reader, 0, 0, width=A4[1]/2, height=A4[0])

            # Right page
            if right_idx >= 0 and right_idx < len(self.brochure.pages):
                right_img = export_renderer.render_page(self.brochure.pages[right_idx])
                right_buffer = io.BytesIO()
                right_img.save(right_buffer, format='PNG')
                right_buffer.seek(0)
                right_reader = ImageReader(right_buffer)
                c.drawImage(right_reader, A4[1]/2, 0, width=A4[1]/2, height=A4[0])

            # Add crop marks (optional)
            c.setStrokeColor(HexColor('#CCCCCC'))
            c.setLineWidth(0.5)
            c.line(A4[1]/2, 0, A4[1]/2, 10*mm)  # Bottom center
            c.line(A4[1]/2, A4[0]-10*mm, A4[1]/2, A4[0])  # Top center

            if i < len(imposition) - 1:
                c.showPage()

        c.save()

    def _export_images(self):
        """Export individual pages as images"""
        # Ask for save directory
        directory = filedialog.askdirectory(title="Select folder to save images")

        if not directory:
            return

        try:
            export_renderer = BrochureRenderer(preview_mode=False)

            for i, page in enumerate(self.brochure.pages):
                img = export_renderer.render_page(page)
                filepath = os.path.join(directory, f"page_{i+1:02d}.png")
                img.save(filepath, 'PNG', dpi=(300, 300))

            messagebox.showinfo("Success", f"Exported {len(self.brochure.pages)} pages to:\n{directory}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export images:\n{str(e)}")

    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    # Check for required dependencies
    try:
        from PIL import Image, ImageDraw, ImageFont, ImageTk
    except ImportError:
        print("Error: Pillow library is required.")
        print("Install it with: pip install Pillow")
        return

    if not HAS_REPORTLAB:
        print("Warning: reportlab not installed. PDF export will be limited.")
        print("Install it with: pip install reportlab")

    app = BrochureGeneratorApp()
    app.run()


if __name__ == "__main__":
    main()
