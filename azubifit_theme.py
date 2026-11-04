"""
========================================================================================
AZUBIFIT CORPORATE DESIGN THEME ENGINE (azubifit_theme.py)
========================================================================================
Updated Specification:
  1. Typography (Title & Text):
     - Main titles, headers, and slide body text MUST remain Deep Navy #012641.
     - Pure White #FFFFFF is strictly preserved for text on dark backgrounds for contrast.

  2. Graphic & Object Colors:
     - ONLY these 4 official Azubifit colors are permitted for graphics, shapes, lines,
       diagrams, and charts:
         • Grün (Green)   : #26966D (RGB: 38, 150, 109)
         • Gelb (Yellow)  : #F2C230 (RGB: 242, 194, 48)
         • Orange (Orange): #D8892E (RGB: 216, 137, 46)
         • Rot (Red)      : #E8383E (RGB: 232, 56, 62)
     - Pure White #FFFFFF is preserved for neutral canvas backgrounds & white line dividers.

  3. Isolated Execution:
     - Completely separated from IfG Vorlage Kurzfassung.
========================================================================================
"""

import math
import colorsys
from collections import Counter

# ==============================================================================
# 1. AZUBIFIT COLOR CONSTANTS
# ==============================================================================

# 🖋️ Typography (Titles & Body Text)
AZUBI_TEXT_NAVY = "012641"  # Official Deep Navy for Title and Text
AZUBI_WHITE = "FFFFFF"      # White text on dark shapes / pure white background

# 🎨 The ONLY 4 Colors for Graphics, Shapes, Diagrams & Charts
AZUBI_GREEN = "26966D"   # Grün (#26966D) - Primary Graphic & Accent Color
AZUBI_YELLOW = "F2C230"  # Gelb (#F2C230) - Yellow Accent
AZUBI_ORANGE = "D8892E"  # Orange (#D8892E) - Orange Accent
AZUBI_RED = "E8383E"     # Rot (#E8383E) - Red Accent / Signal

# Graphic Swatches strictly limited to the 4 colors
AZUBI_GRAPHIC_SWATCHES = [
    AZUBI_GREEN,   # #26966D
    AZUBI_YELLOW,  # #F2C230
    AZUBI_ORANGE,  # #D8892E
    AZUBI_RED,     # #E8383E
]

# Official 4-color Chart & Diagram Palette
AZUBIFIT_CHART_DIAGRAM_PALETTE = [
    AZUBI_GREEN,   # Series 1: Green (#26966D)
    AZUBI_YELLOW,  # Series 2: Yellow (#F2C230)
    AZUBI_ORANGE,  # Series 3: Orange (#D8892E)
    AZUBI_RED,     # Series 4: Red (#E8383E)
]


# ==============================================================================
# 2. COLOR CALCULATION HELPERS
# ==============================================================================

def hex_to_rgb(hex_str):
    h = str(hex_str).strip().lstrip("#")
    if len(h) != 6:
        return (0, 0, 0)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return (0, 0, 0)


def perceived_brightness(rgb):
    return (rgb[0] * 0.299 + rgb[1] * 0.587 + rgb[2] * 0.114) / 255.0


def rgb_saturation(rgb):
    mx = max(rgb)
    mn = min(rgb)
    if mx == 0:
        return 0.0
    return (mx - mn) / float(mx)


def rgb_distance(rgb1, rgb2):
    return math.sqrt(
        (rgb1[0] - rgb2[0]) ** 2
        + (rgb1[1] - rgb2[1]) ** 2
        + (rgb1[2] - rgb2[2]) ** 2
    )


# ==============================================================================
# 3. AZUBIFIT COLOR MAPPER ENGINE
# ==============================================================================

class AzubifitColorMapper:
    """
    Dedicated Color Mapper for Azubifit:
      - Titles & Text: Deep Navy #012641 (White on dark shapes & images)
      - Graphics / Shapes / Accents: Strictly restricted to #26966D, #F2C230, #D8892E, #E8383E
    """
    is_azubifit = True

    def __init__(self, user_prs=None):
        self.is_azubifit = True
        self.brand_navy = AZUBI_TEXT_NAVY       # Titles & Headings (#012641)
        self.brand_charcoal = AZUBI_TEXT_NAVY   # Slide Text (#012641)
        self.brand_muted = AZUBI_TEXT_NAVY      # Subtitles & Kickers (#012641)
        self.brand_white = AZUBI_WHITE          # White (#FFFFFF)
        self.brand_light_blue = AZUBI_GREEN     # Primary accent for shapes (#26966D)

        # Graphic swatches strictly limited to the 4 requested colors
        self.graphic_swatches = AZUBI_GRAPHIC_SWATCHES
        self.chart_palette = AZUBIFIT_CHART_DIAGRAM_PALETTE

        self.clusters = []
        if user_prs is not None:
            self._scan_presentation(user_prs)

    def _scan_presentation(self, user_prs):
        """Scans the source presentation to map accent shapes consistently into the 4 colors."""
        counts = Counter()
        seen_layouts = set()

        for slide in user_prs.slides:
            roots_to_scan = [slide._element]
            try:
                lay_el = slide.slide_layout._element
                if id(lay_el) not in seen_layouts:
                    seen_layouts.add(id(lay_el))
                    roots_to_scan.append(lay_el)
            except Exception:
                pass

            for root in roots_to_scan:
                for parent in root.iter():
                    ptag = parent.tag.split("}")[-1]
                    if ptag in ("solidFill", "gs", "fillRef", "lnRef"):
                        anc_tags = {a.tag.split("}")[-1] for a in parent.iterancestors()}
                        if "effectLst" in anc_tags:
                            continue
                        for ch in parent.iter():
                            if ch.tag.endswith("srgbClr"):
                                val = ch.get("val")
                                if val:
                                    rgb = hex_to_rgb(val)
                                    sat = rgb_saturation(rgb)
                                    if sat >= 0.12:
                                        q = (rgb[0] // 16 * 16, rgb[1] // 16 * 16, rgb[2] // 16 * 16)
                                        counts[q] += 1

        for q_rgb, _ in counts.most_common():
            merged = False
            for existing_rgb, _ in self.clusters:
                if rgb_distance(q_rgb, existing_rgb) < 45:
                    merged = True
                    break
            if not merged:
                swatch = self.graphic_swatches[len(self.clusters) % len(self.graphic_swatches)]
                self.clusters.append((q_rgb, swatch))

    def map_fill_or_line(self, rgb, role="fill"):
        """
        Maps graphic shapes, fills, and lines strictly to the 4 colors:
          - Green: #26966D
          - Yellow: #F2C230
          - Orange: #D8892E
          - Red: #E8383E
        (Preserves pure white #FFFFFF for neutral backdrops).
        """
        sat = rgb_saturation(rgb)
        lum = perceived_brightness(rgb)

        # 1. Pure white or very pale neutral backdrops
        if sat < 0.08 and lum >= 0.95:
            return self.brand_white

        # 2. Semantic Hue Mapping into the 4 official colors
        h, _, _ = colorsys.rgb_to_hls(rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)
        hue_deg = h * 360.0

        # Yellow / Amber (40° <= hue < 68°)
        if 40.0 <= hue_deg < 68.0:
            return AZUBI_YELLOW

        # Orange (15° <= hue < 40°)
        if 15.0 <= hue_deg < 40.0:
            return AZUBI_ORANGE

        # Red / Crimson / Magenta / Pink (hue < 15° or hue >= 320°)
        if hue_deg < 15.0 or hue_deg >= 320.0:
            return AZUBI_RED

        # Blues, Cyans, Teals, and Greens (68° <= hue < 320°)
        # All standard corporate blues/cyans/greens map to Azubifit Green #26966D
        if 68.0 <= hue_deg < 320.0:
            return AZUBI_GREEN

        # 3. Cluster fallback strictly within the 4 swatches
        for c_rgb, hex_val in self.clusters:
            if rgb_distance(rgb, c_rgb) < 50:
                return hex_val

        return AZUBI_GREEN
