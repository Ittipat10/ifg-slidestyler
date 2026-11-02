import streamlit as st
import collections
import collections.abc
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
import os
import io
import re
import posixpath
import tempfile
import warnings
import base64
import colorsys
import hashlib
from PIL import Image
from collections import Counter
from copy import deepcopy
from lxml import etree
import pptx
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

# Critical patch for python-pptx: templates containing SVGs lack sha1 on generic Part,
# which causes get_or_add_image_part() / add_picture() to crash with AttributeError
if not hasattr(Part, 'sha1'):
    Part.sha1 = property(lambda self: hashlib.sha1(self.blob).hexdigest() if hasattr(self, 'blob') else '')

warnings.filterwarnings('ignore', category=UserWarning, module='zipfile')

# Set Streamlit Page Config
st.set_page_config(
    page_title="IfG SlideStyler | PowerPoint Vorlagen-Engine",
    page_icon="🏢",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Load IfG Logo as base64 for reliable inline rendering
logo_b64 = ""
if os.path.exists("ifg-logo.png"):
    with open("ifg-logo.png", "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode("utf-8")

# Custom CSS matching https://qr.ifg-zukunftswerkstatt.de/
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Global Background & Typography */
    html, body, [data-testid="stAppViewContainer"], .main {
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
    }

    /* Hide Streamlit Header, GitHub ribbon & Watermarks */
    #MainMenu {visibility: hidden !important; display: none !important;}
    header {visibility: hidden !important; display: none !important;}
    footer {visibility: hidden !important; display: none !important;}
    [data-testid="stHeader"] {display: none !important;}
    [data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}
    [data-testid="stStatusWidget"] {display: none !important;}
    .stAppDeployButton {display: none !important;}

    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 860px !important;
    }

    /* Typography */
    h1, h2, h3, h4, h5, h6 {
        color: #0f172a !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 700 !important;
        letter-spacing: -0.025em !important;
    }

    p, span, label, li {
        color: #334155;
    }

    /* Top Brand Navigation Header */
    .ifg-top-nav {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        padding: 14px 22px;
        margin-bottom: 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 12px;
    }

    .ifg-nav-left {
        display: flex;
        align-items: center;
        gap: 16px;
    }

    .ifg-logo-img {
        height: 38px;
        width: auto;
        object-fit: contain;
    }

    .ifg-nav-divider {
        height: 32px;
        width: 1px;
        background-color: #e2e8f0;
    }

    .ifg-nav-title {
        font-size: 1.18rem;
        font-weight: 700;
        color: #0f172a;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .ifg-badge-tool {
        display: inline-flex;
        align-items: center;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 0.72rem;
        font-weight: 600;
        background-color: #eff6ff;
        color: #012641;
        border: 1px solid #bfdbfe;
    }

    .ifg-nav-subtitle {
        font-size: 0.8rem;
        color: #64748b;
        margin: 2px 0 0 0;
    }

    .ifg-nav-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 0.75rem;
        font-weight: 500;
        color: #047857;
        background-color: #ecfdf5;
        border: 1px solid #a7f3d0;
        padding: 5px 12px;
        border-radius: 9999px;
    }

    /* Cards */
    .ifg-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        padding: 22px 24px;
        margin-bottom: 22px;
    }

    .ifg-card-header {
        display: flex;
        align-items: center;
        gap: 10px;
        padding-bottom: 12px;
        border-bottom: 1px solid #f1f5f9;
        margin-bottom: 14px;
    }

    .ifg-card-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #012641;
        margin: 0;
    }

    .ifg-feature-list {
        list-style: none;
        padding-left: 0;
        margin: 0;
    }

    .ifg-feature-item {
        display: flex;
        align-items: flex-start;
        gap: 10px;
        padding: 6px 0;
        font-size: 0.92rem;
        color: #334155;
        line-height: 1.5;
    }

    .ifg-bullet-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: #3EB9C0;
        margin-top: 7px;
        flex-shrink: 0;
    }

    /* Streamlit Form Components */
    [data-testid="stSelectbox"] > div > div {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        color: #0f172a !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.04) !important;
        font-size: 0.92rem !important;
    }
    [data-testid="stSelectbox"] > div > div:hover {
        border-color: #012641 !important;
    }

    [data-testid="stFileUploader"] section {
        background-color: #ffffff !important;
        border: 2px dashed #cbd5e1 !important;
        border-radius: 14px !important;
        padding: 20px !important;
        transition: all 0.2s ease !important;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #012641 !important;
        background-color: #f8fafc !important;
    }

    label, [data-testid="stWidgetLabel"] p {
        color: #0f172a !important;
        font-weight: 600 !important;
        font-size: 0.92rem !important;
    }

    /* Buttons */
    .stButton > button {
        background: #012641 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.85rem 2rem !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        box-shadow: 0 2px 4px 0 rgba(1, 38, 65, 0.18) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        background: #083254 !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 10px 0 rgba(1, 38, 65, 0.25) !important;
    }

    .stDownloadButton > button {
        background: #059669 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.85rem 2rem !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        box-shadow: 0 2px 4px 0 rgba(5, 150, 105, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton > button:hover {
        background: #047857 !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 10px 0 rgba(5, 150, 105, 0.28) !important;
    }

    /* Success Card */
    .ifg-success-box {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 14px;
        padding: 20px;
        margin-top: 20px;
        margin-bottom: 20px;
        text-align: center;
    }
    .ifg-success-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #15803d;
        margin-bottom: 6px;
    }
    .ifg-success-desc {
        font-size: 0.92rem;
        color: #166534;
        margin: 0;
    }

    /* Expander Styling */
    [data-testid="stExpander"] {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.03) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ----------------- OPENXML NAMESPACES & CORPORATE CONSTANTS -----------------

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"

CORPORATE_FONT = "Abadi"
TITLE_FONT_SZ = "2800"         # Official Vorlage 'Titel' size: 28 pt
KAPITEL_FONT_SZ = "1200"       # Official Vorlage 'Kapitel' size: 12 pt
VORLAGE_KAPITEL_GRAY = "808080"  # Matches Vorlage slideLayout1 bg1 lumMod=50000

# Official IfG / Arsipa Corporate Color Palette (from Vorlage theme1.xml)
IFG_NAVY = "012641"       # Primary Brand Dark Navy (Headings, Primary Boxes, Headers)
IFG_CHARCOAL = "3A4A63"   # Body Text & Neutral Lines
IFG_MUTED = "8C96A5"      # Kickers, Subtitles, Muted Captions
IFG_WHITE = "FFFFFF"

DARK_BRAND_SWATCHES = ["012641", "006174", "3E3D5E", "007A8C"]
MID_BRAND_SWATCHES = ["00A8B7", "3EB9C0", "57577C", "2B7A9B"]
LIGHT_BRAND_SWATCHES = ["D9F0F2", "C5E4EE", "DCE3F0", "E8F5F6"]

SYMBOL_FONTS = {
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "marlett", "ms outlook", "bookshelf symbol 7", "segoe mdl2 assets",
    "segoe fluent icons", "fontawesome", "font awesome 5 free",
    "font awesome 6 free", "material icons"
}

PRESET_COLORS = {
    "black": (0, 0, 0),
    "white": (255, 255, 255),
    "red": (255, 0, 0),
    "green": (0, 128, 0),
    "blue": (0, 0, 255),
    "yellow": (255, 255, 0),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "lightgray": (211, 211, 211),
    "darkgray": (105, 105, 105),
    "navy": (0, 0, 128),
    "teal": (0, 128, 128),
    "orange": (255, 165, 0),
}


def qn_a(tag):
    return f"{{{A_NS}}}{tag}"


def qn_p(tag):
    return f"{{{P_NS}}}{tag}"


def safe_shape_type(shape):
    """Safely retrieves shape_type without raising NotImplementedError on freeforms or exotic shapes."""
    try:
        return shape.shape_type
    except Exception:
        return None


def scan_arsipa_template(template_prs):
    """Finds the cover layout ('Kapitel mit großem Bild') and corporate canvas layout ('Leere Folie') in the Vorlage."""
    canvas_idx = 0
    cover_idx = 0
    found_cover = False
    for idx, layout in enumerate(template_prs.slide_layouts):
        name_lower = layout.name.lower() if layout.name else ""
        if "leere folie" in name_lower or "blank" in name_lower:
            canvas_idx = idx
        if not found_cover and ("kapitel mit großem bild" in name_lower or "titel" in name_lower or "cover" in name_lower):
            cover_idx = idx
            found_cover = True
    return {
        "canvas_leere_folie": canvas_idx,
        "cover_layout": cover_idx,
    }


# ----------------- 1. COLOR & THEME RESOLUTION ENGINE -----------------

def hex_to_rgb(hex_str):
    h = hex_str.strip().lstrip("#")
    if len(h) != 6:
        return (0, 0, 0)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb_to_hex(rgb):
    return "{:02X}{:02X}{:02X}".format(
        max(0, min(255, int(round(rgb[0])))),
        max(0, min(255, int(round(rgb[1])))),
        max(0, min(255, int(round(rgb[2]))))
    )


def perceived_brightness(rgb):
    """Perceived brightness in [0.0, 1.0] matching human contrast perception."""
    return (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255.0


def rgb_saturation(rgb):
    _, s, _ = colorsys.rgb_to_hsv(rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)
    return s


def rgb_distance(c1, c2):
    return ((c1[0] - c2[0]) ** 2 + (c1[1] - c2[1]) ** 2 + (c1[2] - c2[2]) ** 2) ** 0.5


def shift_hex_lightness(hex_str, delta_l):
    r, g, b = hex_to_rgb(hex_str)
    h, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
    l_new = max(0.08, min(0.95, l + delta_l))
    rn, gn, bn = colorsys.hls_to_rgb(h, l_new, s)
    return rgb_to_hex((rn * 255.0, gn * 255.0, bn * 255.0))


def apply_oxml_color_mods(base_rgb, clr_el):
    """Applies OpenXML DrawingML color transforms (lumMod, lumOff, tint, shade, satMod, alpha)."""
    r, g, b = [c / 255.0 for c in base_rgb]
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    alpha_val = None

    for ch in clr_el:
        local = ch.tag.split("}")[-1]
        raw_val = ch.get("val")
        if raw_val is None:
            continue
        try:
            v = int(raw_val) / 100000.0
        except ValueError:
            continue

        if local == "alpha":
            alpha_val = ch.get("val")
        elif local == "lumMod":
            l = max(0.0, min(1.0, l * v))
        elif local == "lumOff":
            l = max(0.0, min(1.0, l + v))
        elif local == "satMod":
            s = max(0.0, min(1.0, s * v))
        elif local == "satOff":
            s = max(0.0, min(1.0, s + v))
        elif local == "tint":
            r, g, b = colorsys.hls_to_rgb(h, l, s)
            r = 1.0 - (1.0 - r) * v
            g = 1.0 - (1.0 - g) * v
            b = 1.0 - (1.0 - b) * v
            h, l, s = colorsys.rgb_to_hls(r, g, b)
        elif local == "shade":
            r, g, b = colorsys.hls_to_rgb(h, l, s)
            r = r * v
            g = g * v
            b = b * v
            h, l, s = colorsys.rgb_to_hls(r, g, b)

    rf, gf, bf = colorsys.hls_to_rgb(h, l, s)
    return (int(round(rf * 255)), int(round(gf * 255)), int(round(bf * 255))), alpha_val


def extract_slide_theme_map(slide):
    """Resolves the source slide's active theme color scheme + clrMap aliases into RGB tuples."""
    theme_map = {
        "dk1": (0, 0, 0),
        "lt1": (255, 255, 255),
        "dk2": (68, 84, 106),
        "lt2": (231, 230, 230),
        "accent1": (68, 114, 196),
        "accent2": (237, 125, 49),
        "accent3": (165, 165, 165),
        "accent4": (255, 192, 0),
        "accent5": (91, 155, 213),
        "accent6": (112, 173, 71),
        "hlink": (5, 99, 193),
        "folHlink": (149, 79, 114),
    }

    try:
        master = slide.slide_layout.slide_master
        for rel in master.part.rels.values():
            if not rel.is_external and rel.reltype.endswith("/theme"):
                t_root = etree.fromstring(rel.target_part.blob)
                clr_scheme = t_root.find(f".//{{{A_NS}}}clrScheme")
                if clr_scheme is not None:
                    for slot in clr_scheme:
                        slot_name = slot.tag.split("}")[-1]
                        for c in slot:
                            ctag = c.tag.split("}")[-1]
                            if ctag == "srgbClr" and c.get("val"):
                                theme_map[slot_name] = hex_to_rgb(c.get("val"))
                            elif ctag == "sysClr":
                                last = c.get("lastClr") or "000000"
                                theme_map[slot_name] = hex_to_rgb(last)
                break

        # Apply master clrMap (e.g. bg1->lt1, tx1->dk1, bg2->lt2, tx2->dk2)
        clr_map_el = master._element.find(f"{{{P_NS}}}clrMap")
        aliases = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
        if clr_map_el is not None:
            for k, v in clr_map_el.attrib.items():
                aliases[k] = v

        # Check layout & slide clrMapOvr if overridden
        for obj in (slide.slide_layout._element, slide._element):
            ovr = obj.find(f"{{{P_NS}}}clrMapOvr/{{{A_NS}}}overrideClrMapping")
            if ovr is not None:
                for k, v in ovr.attrib.items():
                    aliases[k] = v

        for alias_k, target_k in aliases.items():
            if target_k in theme_map:
                theme_map[alias_k] = theme_map[target_k]
    except Exception:
        pass

    if "tx1" not in theme_map:
        theme_map["tx1"] = theme_map["dk1"]
    if "bg1" not in theme_map:
        theme_map["bg1"] = theme_map["lt1"]
    if "tx2" not in theme_map:
        theme_map["tx2"] = theme_map["dk2"]
    if "bg2" not in theme_map:
        theme_map["bg2"] = theme_map["lt2"]

    return theme_map


def resolve_color_from_parent(parent_el, theme_map, ph_rgb=None):
    """Finds and resolves any DrawingML color child under parent_el into (rgb, alpha_val)."""
    if parent_el is None:
        return None
    for ch in parent_el:
        ctag = ch.tag.split("}")[-1]
        base = None
        if ctag == "srgbClr" and ch.get("val"):
            base = hex_to_rgb(ch.get("val"))
        elif ctag == "schemeClr" and ch.get("val"):
            sval = ch.get("val")
            if sval == "phClr" and ph_rgb is not None:
                base = ph_rgb
            else:
                base = theme_map.get(sval, (0, 0, 0))
        elif ctag == "sysClr":
            base = hex_to_rgb(ch.get("lastClr") or "000000")
        elif ctag == "prstClr" and ch.get("val"):
            base = PRESET_COLORS.get(ch.get("val").lower(), (0, 0, 0))
        elif ctag == "scrgbClr":
            try:
                r = int(round(int(ch.get("r", "0")) / 100000.0 * 255))
                g = int(round(int(ch.get("g", "0")) / 100000.0 * 255))
                b = int(round(int(ch.get("b", "0")) / 100000.0 * 255))
                base = (r, g, b)
            except ValueError:
                base = (0, 0, 0)
        elif ctag == "hslClr":
            try:
                h_val = (int(ch.get("hue", "0")) / 60000.0 % 360.0) / 360.0
                s_val = max(0.0, min(1.0, int(ch.get("sat", "0")) / 100000.0))
                l_val = max(0.0, min(1.0, int(ch.get("lum", "0")) / 100000.0))
                rf, gf, bf = colorsys.hls_to_rgb(h_val, l_val, s_val)
                base = (int(round(rf * 255)), int(round(gf * 255)), int(round(bf * 255)))
            except ValueError:
                base = (0, 0, 0)
        if base is not None:
            return apply_oxml_color_mods(base, ch)
    return None


def write_srgb_into_parent(parent_el, hex_val, alpha_val=None):
    """Replaces any color node under parent_el with explicit <a:srgbClr val='HEX'>."""
    if parent_el is None:
        return
    color_tags = {"srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"}
    for ch in list(parent_el):
        if ch.tag.split("}")[-1] in color_tags:
            parent_el.remove(ch)
    srgb = etree.SubElement(parent_el, qn_a("srgbClr"))
    srgb.set("val", hex_val.upper())
    if alpha_val is not None:
        a_el = etree.SubElement(srgb, qn_a("alpha"))
        a_el.set("val", str(alpha_val))


class DeckColorMapper:
    """
    Maps all colors from an arbitrary source presentation into the official
    IfG / Arsipa Corporate Design palette while strictly preserving luminance
    tiers (dark / medium / light) and distinguishing multiple accent colors.
    """
    def __init__(self, user_prs):
        self.dark_clusters = []   # [(center_rgb, mapped_hex)]
        self.mid_clusters = []
        self.light_clusters = []
        self._scan_presentation(user_prs)

    def _scan_presentation(self, user_prs):
        dark_counts = Counter()
        mid_counts = Counter()
        light_counts = Counter()

        seen_layouts = set()
        for slide in user_prs.slides:
            tmap = extract_slide_theme_map(slide)
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
                        # Skip drop shadows
                        anc_tags = {a.tag.split("}")[-1] for a in parent.iterancestors()}
                        if "effectLst" in anc_tags:
                            continue
                        res = resolve_color_from_parent(parent, tmap)
                        if res is None:
                            continue
                        rgb, _ = res
                        sat = rgb_saturation(rgb)
                        lum = perceived_brightness(rgb)
                        if sat < 0.12:
                            continue
                        # Quantize slightly to group identical brand colors
                        q = (rgb[0] // 8 * 8, rgb[1] // 8 * 8, rgb[2] // 8 * 8)
                        if lum < 0.38:
                            dark_counts[q] += 1
                        elif lum < 0.76:
                            mid_counts[q] += 1
                        else:
                            light_counts[q] += 1

        self.dark_clusters = self._assign_clusters(dark_counts, DARK_BRAND_SWATCHES)
        self.mid_clusters = self._assign_clusters(mid_counts, MID_BRAND_SWATCHES)
        self.light_clusters = self._assign_clusters(light_counts, LIGHT_BRAND_SWATCHES)

    @staticmethod
    def _assign_clusters(counter, swatches):
        clusters = []
        for q_rgb, _ in counter.most_common():
            merged = False
            for existing_rgb, _ in clusters:
                if rgb_distance(q_rgb, existing_rgb) < 34:
                    merged = True
                    break
            if not merged:
                swatch = swatches[len(clusters) % len(swatches)]
                clusters.append((q_rgb, swatch))
        return clusters

    def _match_cluster(self, rgb, clusters, swatches):
        for c_rgb, hex_val in clusters:
            if rgb_distance(rgb, c_rgb) < 40:
                return hex_val
        idx = len(clusters) % len(swatches)
        swatch = swatches[idx]
        clusters.append((rgb, swatch))
        return swatch

    def map_fill_or_line(self, rgb, role="fill"):
        sat = rgb_saturation(rgb)
        lum = perceived_brightness(rgb)

        # Neutral / Achromatic colors
        if sat < 0.12:
            if lum >= 0.94:
                return IFG_WHITE
            elif lum >= 0.86:
                return "F1F5F9"
            elif lum >= 0.70:
                return "CBD5E1"
            elif lum >= 0.32:
                return IFG_MUTED
            else:
                return IFG_NAVY if role == "fill" else IFG_CHARCOAL

        # Chromatic colors mapped by luminance tier so contrast never inverts
        if lum < 0.38:
            return self._match_cluster(rgb, self.dark_clusters, DARK_BRAND_SWATCHES)
        elif lum < 0.76:
            return self._match_cluster(rgb, self.mid_clusters, MID_BRAND_SWATCHES)
        else:
            return self._match_cluster(rgb, self.light_clusters, LIGHT_BRAND_SWATCHES)


# ----------------- 2. FONT & TEXT RESTYLING HELPERS -----------------

def is_symbol_text(text):
    if not text:
        return False
    stripped = text.strip()
    if not stripped:
        return False
    # Private Use Area glyphs (\uF000-\uF8FF) used by Wingdings/Symbol in OpenXML
    return any(0xF000 <= ord(ch) <= 0xF8FF for ch in stripped)


def is_symbol_font_name(typeface):
    if not typeface:
        return False
    return typeface.strip().lower() in SYMBOL_FONTS


def apply_corporate_font_to_rpr(rpr_el, run_text=""):
    """Sets Abadi on <a:rPr> / <a:defRPr> / <a:endParaRPr> while protecting symbol glyphs/fonts."""
    if rpr_el is None or is_symbol_text(run_text):
        return

    latin = rpr_el.find(qn_a("latin"))
    if latin is not None and is_symbol_font_name(latin.get("typeface")):
        return

    if latin is None:
        # Insert <a:latin> in proper DrawingML child order (after Fill/Effect/Highlight, before ea/cs/sym)
        latin = etree.Element(qn_a("latin"))
        insert_after_tags = {
            "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
            "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill"
        }
        idx = 0
        for i, ch in enumerate(rpr_el):
            if ch.tag.split("}")[-1] in insert_after_tags:
                idx = i + 1
        rpr_el.insert(idx, latin)

    latin.attrib.clear()
    latin.set("typeface", CORPORATE_FONT)

    for tag_name in ("cs", "ea"):
        f_el = rpr_el.find(qn_a(tag_name))
        if f_el is not None and not is_symbol_font_name(f_el.get("typeface")):
            f_el.attrib.clear()
            f_el.set("typeface", CORPORATE_FONT)


def restyle_txbody(txbody, theme_map, mapper, bg_lum, is_heading_shape=False, style_font_rgb=None, font_scale=1.0, header_size_mode=None):
    """
    Restyles all runs in a <p:txBody> / <a:txBody> to Abadi + Corporate Colors
    while preserving 100% of original paragraphs, line breaks, spacing, bullets,
    alignments, bold/italic, and ensuring readable contrast against bg_lum.
    When header_size_mode is 'title' or 'kapitel', enforces official 28pt (2800) / 12pt (1200) sizes.
    """
    if txbody is None:
        return

    # Restyle lstStyle defRPr fonts if present
    lst_style = txbody.find(qn_a("lstStyle"))
    lst_lvl_colors = {}
    lst_lvl_sizes = {}
    if lst_style is not None:
        for ch in lst_style:
            local = ch.tag.split("}")[-1]
            if local.startswith("lvl") and local.endswith("pPr"):
                lvl_num = local[3]
                def_rpr = ch.find(qn_a("defRPr"))
                if def_rpr is not None:
                    res = resolve_color_from_parent(def_rpr.find(qn_a("solidFill")), theme_map)
                    if res is not None:
                        lst_lvl_colors[lvl_num] = res[0]
                    if def_rpr.get("sz"):
                        try:
                            lst_lvl_sizes[lvl_num] = int(def_rpr.get("sz"))
                            if header_size_mode == "title":
                                def_rpr.set("sz", TITLE_FONT_SZ)
                            elif header_size_mode == "kapitel":
                                def_rpr.set("sz", KAPITEL_FONT_SZ)
                            elif abs(font_scale - 1.0) > 0.02:
                                def_rpr.set("sz", str(int(round(lst_lvl_sizes[lvl_num] * font_scale))))
                        except ValueError:
                            pass
                    elif header_size_mode in ("title", "kapitel"):
                        def_rpr.set("sz", TITLE_FONT_SZ if header_size_mode == "title" else KAPITEL_FONT_SZ)
                    def_sf = def_rpr.find(qn_a("solidFill"))
                    if def_sf is not None:
                        write_srgb_into_parent(def_sf, IFG_WHITE if bg_lum < 0.48 else (IFG_NAVY if is_heading_shape else IFG_CHARCOAL))
                    apply_corporate_font_to_rpr(def_rpr)

    # Determine largest font size in this text body to distinguish main title vs kicker line
    max_sz = max(lst_lvl_sizes.values(), default=0)
    for r in txbody.findall(f".//{{{A_NS}}}r"):
        rpr = r.find(qn_a("rPr"))
        if rpr is not None and rpr.get("sz"):
            try:
                max_sz = max(max_sz, int(rpr.get("sz")))
            except ValueError:
                pass

    default_fallback_rgb = style_font_rgb if style_font_rgb is not None else theme_map.get("tx1", (0, 0, 0))

    for p in txbody.findall(qn_a("p")):
        ppr = p.find(qn_a("pPr"))
        lvl_key = str(int(ppr.get("lvl", "0")) + 1) if ppr is not None else "1"
        p_def_rgb = lst_lvl_colors.get(lvl_key, default_fallback_rgb)
        p_def_sz = lst_lvl_sizes.get(lvl_key, 0)

        if ppr is not None:
            def_rpr = ppr.find(qn_a("defRPr"))
            if def_rpr is not None:
                res = resolve_color_from_parent(def_rpr.find(qn_a("solidFill")), theme_map)
                if res is not None:
                    p_def_rgb = res[0]
                if def_rpr.get("sz"):
                    try:
                        p_def_sz = int(def_rpr.get("sz"))
                    except ValueError:
                        pass
                if header_size_mode in ("title", "kapitel"):
                    def_rpr.set("sz", TITLE_FONT_SZ if header_size_mode == "title" else KAPITEL_FONT_SZ)
                def_sf = def_rpr.find(qn_a("solidFill"))
                if def_sf is not None:
                    write_srgb_into_parent(def_sf, IFG_WHITE if bg_lum < 0.48 else (IFG_NAVY if is_heading_shape else IFG_CHARCOAL))
                apply_corporate_font_to_rpr(def_rpr)

            # Restyle explicit bullet color (<a:buClr>) if present; never touch <a:buFont>!
            bu_clr = ppr.find(qn_a("buClr"))
            if bu_clr is not None:
                new_bu_hex = IFG_WHITE if bg_lum < 0.48 else IFG_NAVY
                write_srgb_into_parent(bu_clr, new_bu_hex)

        for child in p:
            ctag = child.tag.split("}")[-1]
            if ctag == "r":
                t_el = child.find(qn_a("t"))
                r_txt = t_el.text if t_el is not None and t_el.text else ""
                rpr = child.find(qn_a("rPr"))
                if rpr is None:
                    rpr = etree.Element(qn_a("rPr"))
                    child.insert(0, rpr)

                run_sz = p_def_sz
                if rpr.get("sz"):
                    try:
                        run_sz = int(rpr.get("sz"))
                    except ValueError:
                        pass

                is_kicker_run_in_title = (
                    header_size_mode == "title"
                    and max_sz > 0
                    and run_sz > 0
                    and run_sz <= max_sz - 400
                    and run_sz <= 1800
                )

                # Enforce official Title (28pt) / Kapitel (12pt) sizes or scale proportionally
                if header_size_mode == "kapitel" or is_kicker_run_in_title:
                    rpr.set("sz", KAPITEL_FONT_SZ)
                    rpr.set("b", "0")
                elif header_size_mode == "title":
                    rpr.set("sz", TITLE_FONT_SZ)
                    rpr.set("b", "0")
                elif abs(font_scale - 1.0) > 0.02 and rpr.get("sz"):
                    try:
                        rpr.set("sz", str(int(round(int(rpr.get("sz")) * font_scale))))
                    except ValueError:
                        pass

                # Resolve original text color
                orig_rgb = p_def_rgb
                orig_sf = rpr.find(qn_a("solidFill"))
                if orig_sf is not None:
                    res = resolve_color_from_parent(orig_sf, theme_map)
                    if res is not None:
                        orig_rgb = res[0]

                # Choose corporate text color with strict contrast guarantee
                if bg_lum < 0.48:
                    new_txt_hex = IFG_WHITE
                elif header_size_mode == "kapitel" or is_kicker_run_in_title:
                    new_txt_hex = VORLAGE_KAPITEL_GRAY
                elif header_size_mode == "title":
                    new_txt_hex = IFG_NAVY
                else:
                    orig_sat = rgb_saturation(orig_rgb)
                    orig_lum = perceived_brightness(orig_rgb)
                    if orig_sat >= 0.20 and orig_lum < 0.78:
                        # Highlighted / colored text run -> map to dark/readable corporate brand color
                        mapped = mapper.map_fill_or_line(orig_rgb, role="fill")
                        if perceived_brightness(hex_to_rgb(mapped)) > 0.55:
                            mapped = "006174"
                        new_txt_hex = mapped
                    elif is_heading_shape:
                        if max_sz > 0 and run_sz > 0 and run_sz < max_sz - 300 and run_sz < 1800:
                            new_txt_hex = IFG_MUTED
                        else:
                            new_txt_hex = IFG_NAVY
                    elif run_sz >= 2000:
                        new_txt_hex = IFG_NAVY
                    elif 0.40 <= orig_lum <= 0.78 and orig_sat < 0.20 and (run_sz == 0 or run_sz < 1800):
                        # Small muted gray kicker / caption
                        new_txt_hex = IFG_MUTED
                    else:
                        new_txt_hex = IFG_CHARCOAL

                # Ensure <a:solidFill> exists in proper DrawingML schema position inside <a:rPr>
                for ftag in ("noFill", "solidFill", "gradFill", "pattFill"):
                    old_f = rpr.find(qn_a(ftag))
                    if old_f is not None:
                        rpr.remove(old_f)

                sf = etree.Element(qn_a("solidFill"))
                write_srgb_into_parent(sf, new_txt_hex)
                insert_pos = 1 if (len(rpr) > 0 and rpr[0].tag.split("}")[-1] == "ln") else 0
                rpr.insert(insert_pos, sf)

                apply_corporate_font_to_rpr(rpr, r_txt)

            elif ctag == "fld":
                rpr = child.find(qn_a("rPr"))
                if rpr is not None:
                    fld_sf = rpr.find(qn_a("solidFill"))
                    if fld_sf is not None:
                        write_srgb_into_parent(fld_sf, IFG_WHITE if bg_lum < 0.48 else IFG_CHARCOAL)
                    apply_corporate_font_to_rpr(rpr)
            elif ctag == "endParaRPr":
                if header_size_mode == "kapitel":
                    child.set("sz", KAPITEL_FONT_SZ)
                elif header_size_mode == "title":
                    child.set("sz", TITLE_FONT_SZ)
                end_sf = child.find(qn_a("solidFill"))
                if end_sf is not None:
                    write_srgb_into_parent(end_sf, IFG_WHITE if bg_lum < 0.48 else (IFG_NAVY if is_heading_shape else IFG_CHARCOAL))
                apply_corporate_font_to_rpr(child)


# ----------------- 3. PLACEHOLDER & GEOMETRY MATERIALIZATION -----------------

def find_inherited_placeholder(old_slide, ph_idx, ph_type):
    """
    Finds matching layout and master placeholder elements for style/fill/geometry inheritance.
    Strictly prefers exact placeholder idx matches on each container before falling back to type.
    """
    matches = []
    for container in (old_slide.slide_layout, old_slide.slide_layout.slide_master):
        try:
            exact_match = None
            type_match = None
            for sh in container.placeholders:
                c_idx = sh.placeholder_format.idx
                c_type = sh.placeholder_format.type
                if ph_idx is not None and c_idx == ph_idx:
                    exact_match = sh._element
                    break
                if type_match is None and ph_type is not None and c_type == ph_type:
                    type_match = sh._element
            chosen = exact_match if exact_match is not None else type_match
            if chosen is not None:
                matches.append(chosen)
        except Exception:
            pass
    return matches


def extract_inherited_level_props(ph_elems, is_title_ph=False):
    """Extracts inherited {lvl_int: {'sz': str, 'b': str, 'algn': str, 'buNone': bool}} from layout/master."""
    props = {i: {} for i in range(9)}
    # Walk master first, then layout so layout overrides master
    for el in reversed(ph_elems):
        lst = el.find(f".//{{{P_NS}}}txBody/{{{A_NS}}}lstStyle")
        if lst is None:
            continue
        for lvl_idx in range(9):
            lp = lst.find(qn_a(f"lvl{lvl_idx + 1}pPr"))
            if lp is None:
                continue
            if lp.get("algn"):
                props[lvl_idx]["algn"] = lp.get("algn")
            if lp.find(qn_a("buNone")) is not None:
                props[lvl_idx]["buNone"] = True
            elif any(lp.find(qn_a(bt)) is not None for bt in ("buChar", "buAutoNum", "buBlip")):
                props[lvl_idx]["buNone"] = False
            dr = lp.find(qn_a("defRPr"))
            if dr is not None:
                if dr.get("sz"):
                    props[lvl_idx]["sz"] = dr.get("sz")
                if dr.get("b"):
                    props[lvl_idx]["b"] = dr.get("b")

    if is_title_ph:
        if "sz" not in props[0]:
            props[0]["sz"] = "2800"
        props[0].setdefault("buNone", True)
    return props


def shape_has_visual_fill(el):
    """Returns True if <p:sp> has an explicit visual fill in <p:spPr> or <p:style>/<a:fillRef>."""
    if el is None:
        return False
    sp_pr = el.find(qn_p("spPr"))
    if sp_pr is not None:
        if sp_pr.find(qn_a("noFill")) is not None:
            return False
        if any(sp_pr.find(qn_a(ft)) is not None for ft in ("solidFill", "gradFill", "blipFill", "pattFill")):
            return True
    fill_ref = el.find(f"{{{P_NS}}}style/{{{A_NS}}}fillRef")
    if fill_ref is not None and fill_ref.get("idx", "0") != "0":
        return True
    return False


def get_pic_embed_sha1(sh, slide):
    """Returns SHA1 of an embedded picture on slide if small/checkable, else None."""
    try:
        if sh._element.tag.split("}")[-1] != "pic":
            return None
        blip = sh._element.find(f".//{{{A_NS}}}blip")
        if blip is None:
            return None
        rid = blip.get(f"{{{R_NS}}}embed")
        if not rid or rid not in slide.part.rels:
            return None
        rel = slide.part.rels[rid]
        if rel.is_external:
            return None
        return hashlib.sha1(rel.target_part.blob).hexdigest()
    except Exception:
        return None


def detect_repeated_corner_logos(user_prs, scale, dx, dy):
    """
    Identifies small top-right corner logo images repeated across >= 3 content slides
    so old template corner logos don't double-stack over the official Vorlage 'a.' logo.
    """
    sha_slide_sets = collections.defaultdict(set)
    for s_idx, slide in enumerate(user_prs.slides):
        if s_idx == 0:
            continue
        for sh in slide.shapes:
            if sh._element.tag.split("}")[-1] != "pic":
                continue
            x1 = (sh.left or 0) * scale + dx
            y1 = (sh.top or 0) * scale + dy
            w = (sh.width or 0) * scale
            h = (sh.height or 0) * scale
            if (
                x1 >= int(9.9 * 914400)
                and (y1 + h) <= int(1.60 * 914400)
                and 0 < w <= int(3.15 * 914400)
                and 0 < h <= int(1.40 * 914400)
            ):
                sha = get_pic_embed_sha1(sh, slide)
                if sha:
                    sha_slide_sets[sha].add(s_idx)
    return {sha for sha, slides in sha_slide_sets.items() if len(slides) >= 3}


def should_skip_source_shape(sh, slide=None, repeated_corner_logos=None, scale=1.0, dx=0.0, dy=0.0):
    """Skips hidden 0x0 think-cell metadata, repeated old corner logos, and empty/system placeholders."""
    name_low = (sh.name or "").lower()
    if "think-cell" in name_low:
        return True

    cNvPr = sh._element.find(f".//{{{P_NS}}}cNvPr")
    if cNvPr is not None and cNvPr.get("hidden") == "1":
        w = sh.width or 0
        h = sh.height or 0
        if w <= 5000 and h <= 5000:
            return True

    if repeated_corner_logos and slide is not None and sh._element.tag.split("}")[-1] == "pic":
        x1 = (sh.left or 0) * scale + dx
        y1 = (sh.top or 0) * scale + dy
        w = (sh.width or 0) * scale
        h = (sh.height or 0) * scale
        if (
            x1 >= int(9.9 * 914400)
            and (y1 + h) <= int(1.60 * 914400)
            and 0 < w <= int(3.15 * 914400)
            and 0 < h <= int(1.40 * 914400)
        ):
            sha = get_pic_embed_sha1(sh, slide)
            if sha in repeated_corner_logos:
                return True

    if sh.is_placeholder:
        try:
            pht = sh.placeholder_format.type
            if pht in (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.HEADER, PP_PLACEHOLDER.FOOTER):
                return True
        except Exception:
            pass

        tag = sh._element.tag.split("}")[-1]
        if tag == "sp":
            has_txt = getattr(sh, "has_text_frame", False) and sh.text_frame and sh.text_frame.text.strip()
            has_fill = shape_has_visual_fill(sh._element)
            if not has_txt and not has_fill:
                return True

    if any(k in name_low for k in ("slide number", "foliennummer", "page number", "datum")):
        if not (getattr(sh, "has_text_frame", False) and len(sh.text_frame.text.strip()) > 12):
            return True

    return False


def insert_sppr_child(sp_pr, new_el, slot_rank):
    """
    Inserts a DrawingML child into <p:spPr> strictly respecting CT_ShapeProperties sequence:
    0: xfrm | 1: custGeom, prstGeom | 2: fills | 3: ln | 4: effects | 5: scene3d | 6: sp3d | 7: extLst
    """
    rank_map = {
        "xfrm": 0,
        "custGeom": 1, "prstGeom": 1,
        "noFill": 2, "solidFill": 2, "gradFill": 2, "blipFill": 2, "pattFill": 2, "grpFill": 2,
        "ln": 3,
        "effectLst": 4, "effectDag": 4,
        "scene3d": 5, "sp3d": 6, "extLst": 7,
    }
    pos = len(sp_pr)
    for i, ch in enumerate(sp_pr):
        if rank_map.get(ch.tag.split("}")[-1], 99) > slot_rank:
            pos = i
            break
    sp_pr.insert(pos, new_el)


def insert_ppr_bunone(ppr):
    """Inserts <a:buNone/> into <a:pPr> strictly before tabLst/defRPr/extLst."""
    after_bullet_tags = {"tabLst", "defRPr", "extLst"}
    pos = len(ppr)
    for i, ch in enumerate(ppr):
        if ch.tag.split("}")[-1] in after_bullet_tags:
            pos = i
            break
    ppr.insert(pos, etree.Element(qn_a("buNone")))


def ensure_explicit_xfrm(elem, x, y, cx, cy):
    """Ensures a shape/picture/graphicFrame XML element has explicit coordinates and valid geometry."""
    tag = elem.tag.split("}")[-1]
    if tag == "graphicFrame":
        xfrm = elem.find(qn_p("xfrm"))
        if xfrm is None:
            xfrm = etree.Element(qn_p("xfrm"))
            elem.insert(1, xfrm)
    elif tag == "grpSp":
        grp_pr = elem.find(qn_p("grpSpPr"))
        if grp_pr is None:
            grp_pr = etree.Element(qn_p("grpSpPr"))
            elem.insert(1, grp_pr)
        xfrm = grp_pr.find(qn_a("xfrm"))
        if xfrm is None:
            xfrm = etree.Element(qn_a("xfrm"))
            grp_pr.insert(0, xfrm)
    else:
        sp_pr = elem.find(qn_p("spPr"))
        if sp_pr is None:
            sp_pr = etree.Element(qn_p("spPr"))
            elem.insert(1, sp_pr)
        xfrm = sp_pr.find(qn_a("xfrm"))
        if xfrm is None:
            xfrm = etree.Element(qn_a("xfrm"))
            sp_pr.insert(0, xfrm)
        # Critical: both <p:pic> and materialized <p:sp> require <a:prstGeom> or <a:custGeom> to render fills/borders!
        if tag in ("pic", "sp") and sp_pr.find(qn_a("prstGeom")) is None and sp_pr.find(qn_a("custGeom")) is None:
            prst = etree.Element(qn_a("prstGeom"))
            prst.set("prst", "rect")
            etree.SubElement(prst, qn_a("avLst"))
            insert_sppr_child(sp_pr, prst, slot_rank=1)

    off = xfrm.find(qn_a("off"))
    if off is None:
        off = etree.Element(qn_a("off"))
        xfrm.insert(0, off)
    off.set("x", str(int(round(x))))
    off.set("y", str(int(round(y))))

    ext = xfrm.find(qn_a("ext"))
    if ext is None:
        ext = etree.Element(qn_a("ext"))
        xfrm.insert(1, ext)
    ext.set("cx", str(max(0, int(round(cx)))))
    ext.set("cy", str(max(0, int(round(cy)))))


def materialize_placeholder_and_transform(sh, elem, old_slide, scale, dx, dy, is_top_left_header=False, max_header_right=None):
    """
    Materializes inherited placeholder coordinates, geometry, fills, borders, bodyPr margins/anchor,
    and typography if sh.is_placeholder, strips <p:ph> so it never collides with Vorlage placeholders,
    and applies aspect-ratio-safe slide coordinate mapping.
    """
    orig_x = sh.left if sh.left is not None else 0
    orig_y = sh.top if sh.top is not None else 0
    orig_cx = sh.width if sh.width is not None else 0
    orig_cy = sh.height if sh.height is not None else 0

    if sh.is_placeholder:
        ph_idx = None
        ph_type = None
        is_title_ph = False
        try:
            ph_idx = sh.placeholder_format.idx
            ph_type = sh.placeholder_format.type
            is_title_ph = ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
        except Exception:
            pass

        inherited_els = find_inherited_placeholder(old_slide, ph_idx, ph_type)

        # Fallback coordinates from layout/master placeholder if python-pptx had 0x0
        if orig_cx == 0 and orig_cy == 0:
            for inh_el in inherited_els:
                inh_bbox = get_local_bbox(inh_el)
                if inh_bbox is not None and (inh_bbox[2] - inh_bbox[0]) > 0:
                    orig_x, orig_y = inh_bbox[0], inh_bbox[1]
                    orig_cx, orig_cy = inh_bbox[2] - inh_bbox[0], inh_bbox[3] - inh_bbox[1]
                    break

        # Inherit geometry, fill, border, and <p:style> from layout/master placeholder onto <p:sp>
        if elem.tag.split("}")[-1] == "sp" and inherited_els:
            sp_pr = elem.find(qn_p("spPr"))
            if sp_pr is None:
                sp_pr = etree.Element(qn_p("spPr"))
                elem.insert(1, sp_pr)

            # 1) Inherit geometry (prstGeom / custGeom)
            if sp_pr.find(qn_a("prstGeom")) is None and sp_pr.find(qn_a("custGeom")) is None:
                for inh_el in inherited_els:
                    inh_sppr = inh_el.find(qn_p("spPr"))
                    if inh_sppr is not None:
                        geom = inh_sppr.find(qn_a("custGeom"))
                        if geom is None:
                            geom = inh_sppr.find(qn_a("prstGeom"))
                        if geom is not None:
                            insert_sppr_child(sp_pr, deepcopy(geom), slot_rank=1)
                            break

            # 2) Inherit fill (solidFill / gradFill / pattFill / noFill)
            fill_tags = ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")
            if not any(sp_pr.find(qn_a(ft)) is not None for ft in fill_tags):
                for inh_el in inherited_els:
                    inh_sppr = inh_el.find(qn_p("spPr"))
                    if inh_sppr is not None:
                        found_fill = None
                        for ft in fill_tags:
                            cand = inh_sppr.find(qn_a(ft))
                            if cand is not None:
                                found_fill = cand
                                break
                        if found_fill is not None:
                            if found_fill.tag.split("}")[-1] != "blipFill":
                                insert_sppr_child(sp_pr, deepcopy(found_fill), slot_rank=2)
                            break

            # 3) Inherit border (<a:ln>)
            if sp_pr.find(qn_a("ln")) is None:
                for inh_el in inherited_els:
                    inh_ln = inh_el.find(f"{{{P_NS}}}spPr/{{{A_NS}}}ln")
                    if inh_ln is not None:
                        insert_sppr_child(sp_pr, deepcopy(inh_ln), slot_rank=3)
                        break

            # 4) Inherit <p:style> (<a:fillRef>, <a:lnRef>, <a:fontRef>)
            if elem.find(qn_p("style")) is None:
                for inh_el in inherited_els:
                    inh_style = inh_el.find(qn_p("style"))
                    if inh_style is not None:
                        sp_pr_idx = list(elem).index(sp_pr)
                        elem.insert(sp_pr_idx + 1, deepcopy(inh_style))
                        break

        lvl_props = extract_inherited_level_props(inherited_els, is_title_ph=is_title_ph)

        txbody = elem.find(qn_p("txBody"))
        if txbody is not None:
            # 5) Inherit <a:bodyPr> vertical alignment (anchor="ctr"/"b"/"t"), padding insets, wrap, vert
            body_pr = txbody.find(qn_a("bodyPr"))
            if body_pr is None:
                body_pr = etree.Element(qn_a("bodyPr"))
                txbody.insert(0, body_pr)

            inh_bpr_attrs = {}
            inh_autofit = None
            for inh_el in reversed(inherited_els):
                ibpr = inh_el.find(f".//{{{P_NS}}}txBody/{{{A_NS}}}bodyPr")
                if ibpr is not None:
                    for k, v in ibpr.attrib.items():
                        inh_bpr_attrs[k] = v
                    for af_tag in ("noAutofit", "normAutofit", "spAutoFit"):
                        af_el = ibpr.find(qn_a(af_tag))
                        if af_el is not None:
                            inh_autofit = af_el

            for attr_k in ("anchor", "anchorCtr", "vert", "wrap", "lIns", "tIns", "rIns", "bIns", "numCol", "spcCol", "rtlCol", "spcFirstLastPara"):
                if attr_k in inh_bpr_attrs and body_pr.get(attr_k) is None:
                    body_pr.set(attr_k, inh_bpr_attrs[attr_k])

            if abs(scale - 1.0) > 0.02:
                for ins_k in ("lIns", "tIns", "rIns", "bIns"):
                    if body_pr.get(ins_k) is not None:
                        try:
                            body_pr.set(ins_k, str(int(round(int(body_pr.get(ins_k)) * scale))))
                        except ValueError:
                            pass

            if inh_autofit is not None and not any(body_pr.find(qn_a(t)) is not None for t in ("noAutofit", "normAutofit", "spAutoFit")):
                ins_idx = 1 if (len(body_pr) > 0 and body_pr[0].tag.split("}")[-1] == "prstTxWarp") else 0
                body_pr.insert(ins_idx, deepcopy(inh_autofit))

            for p in txbody.findall(qn_a("p")):
                ppr = p.find(qn_a("pPr"))
                lvl = int(ppr.get("lvl", "0")) if ppr is not None else 0
                inh = lvl_props.get(min(lvl, 8), {})

                if inh.get("algn") or inh.get("buNone"):
                    if ppr is None:
                        ppr = etree.Element(qn_a("pPr"))
                        p.insert(0, ppr)
                    if inh.get("algn") and not ppr.get("algn"):
                        ppr.set("algn", inh["algn"])
                    if inh.get("buNone") and not any(
                        ppr.find(qn_a(bt)) is not None for bt in ("buNone", "buChar", "buAutoNum", "buBlip")
                    ):
                        insert_ppr_bunone(ppr)

                for r in p.findall(qn_a("r")):
                    rpr = r.find(qn_a("rPr"))
                    if rpr is None:
                        rpr = etree.Element(qn_a("rPr"))
                        r.insert(0, rpr)
                    if inh.get("sz") and not rpr.get("sz"):
                        rpr.set("sz", inh["sz"])
                    if inh.get("b") and not rpr.get("b"):
                        rpr.set("b", inh["b"])

        # Remove <p:ph> from <p:nvPr> so it stands as an exact-coordinate native object
        for nv_pr in elem.findall(f".//{{{P_NS}}}nvPr"):
            for ph in list(nv_pr.findall(qn_p("ph"))):
                nv_pr.remove(ph)

    # Compute target coordinates
    new_y = orig_y * scale + dy
    new_cy = orig_cy * scale
    new_cx = orig_cx * scale

    if is_top_left_header:
        # Align top-left header neatly next to Vorlage's left brand bar (0.47") and clear of top-right logo
        new_x = max(int(0.47 * 914400), int(round(orig_x * scale)))
        if max_header_right is not None and (new_x + new_cx) > max_header_right:
            new_cx = max(int(2.0 * 914400), max_header_right - new_x)
    else:
        new_x = orig_x * scale + dx

    ensure_explicit_xfrm(elem, new_x, new_y, new_cx, new_cy)

    # Scale table column widths and row heights only if uniform scale != 1.0
    if abs(scale - 1.0) > 0.02 and elem.tag.split("}")[-1] == "graphicFrame":
        for gc in elem.findall(f".//{{{A_NS}}}gridCol"):
            if gc.get("w"):
                try:
                    gc.set("w", str(int(round(int(gc.get("w")) * scale))))
                except ValueError:
                    pass
        for tr in elem.findall(f".//{{{A_NS}}}tr"):
            if tr.get("h"):
                try:
                    tr.set("h", str(int(round(int(tr.get("h")) * scale))))
                except ValueError:
                    pass

    return (new_x, new_y, new_x + new_cx, new_y + new_cy)


def extract_layout_background_panels(old_slide, old_w, old_h):
    """
    Extracts large non-placeholder colored background panels (e.g., full-height sidebars
    or header banners) defined on old_slide.slide_layout so they are preserved on 'Leere Folie'.
    """
    if old_slide._element.get("showMasterSp") == "0":
        return []
    panels = []
    try:
        lay = old_slide.slide_layout
        tmap = extract_slide_theme_map(old_slide)
        slide_boxes = []
        for s_el in old_slide._element.findall(f".//{{{P_NS}}}spTree/{{{P_NS}}}sp"):
            sb = get_local_bbox(s_el)
            if sb is not None and shape_has_visual_fill(s_el):
                slide_boxes.append(sb)

        for lel in lay._element.findall(f".//{{{P_NS}}}spTree/{{{P_NS}}}sp"):
            if lel.find(f".//{{{P_NS}}}ph") is not None:
                continue
            txt = "".join(t.text or "" for t in lel.iter(qn_a("t"))).strip()
            if txt:
                continue
            if not shape_has_visual_fill(lel):
                continue
            lb = get_local_bbox(lel)
            if lb is None:
                continue
            x1, y1, x2, y2 = lb
            w, h = x2 - x1, y2 - y1
            # Must be an on-canvas structural rectangular panel (excludes tiny dashes & huge off-canvas decorations)
            if (
                x1 < -300000 or y1 < -300000
                or x2 > old_w + 300000 or y2 > old_h + 300000
                or w < 0.18 * old_w or h < 0.18 * old_h
                or (w * h) < 0.10 * (old_w * old_h)
            ):
                continue

            # Resolve fill brightness to skip plain white background rectangles
            sp_pr = lel.find(qn_p("spPr"))
            res = None
            if sp_pr is not None:
                sf = sp_pr.find(qn_a("solidFill"))
                gf = sp_pr.find(qn_a("gradFill"))
                if sf is not None:
                    res = resolve_color_from_parent(sf, tmap)
                elif gf is not None:
                    gs = gf.find(f".//{{{A_NS}}}gs")
                    res = resolve_color_from_parent(gs, tmap)
            if res is None:
                fref = lel.find(f"{{{P_NS}}}style/{{{A_NS}}}fillRef")
                res = resolve_color_from_parent(fref, tmap)
            if res is not None and perceived_brightness(res[0]) >= 0.95 and rgb_saturation(res[0]) < 0.08:
                continue

            # Skip if old_slide already has a filled shape covering essentially the same panel box
            already_covered = False
            for sx1, sy1, sx2, sy2 in slide_boxes:
                ix = max(0, min(x2, sx2) - max(x1, sx1))
                iy = max(0, min(y2, sy2) - max(y1, sy1))
                if (ix * iy) > 0.75 * (w * h):
                    already_covered = True
                    break
            if not already_covered:
                panels.append((lel, (x1, y1, w, h), lay))
    except Exception:
        pass
    return panels


# ----------------- 4. DEEP SHAPE, GROUP, TABLE, SMARTART & CHART RESTYLER -----------------

def get_bg_lum_under_box(bbox, z_bg_stack, default_lum=1.0):
    """Checks previously drawn opaque shapes in Z-order to find the effective background luminance."""
    if not bbox:
        return default_lum
    cx = (bbox[0] + bbox[2]) / 2.0
    cy = (bbox[1] + bbox[3]) / 2.0
    for bx1, by1, bx2, by2, blum in reversed(z_bg_stack):
        if bx1 <= cx <= bx2 and by1 <= cy <= by2:
            return blum
    return default_lum


def get_local_bbox(el):
    """Reads (x1, y1, x2, y2) from an element's local xfrm."""
    xfrm = el.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
    if xfrm is None:
        xfrm = el.find(f"{{{P_NS}}}grpSpPr/{{{A_NS}}}xfrm")
    if xfrm is None:
        xfrm = el.find(qn_p("xfrm"))
    if xfrm is None:
        return None
    off = xfrm.find(qn_a("off"))
    ext = xfrm.find(qn_a("ext"))
    if off is None or ext is None:
        return None
    try:
        x = int(off.get("x", "0"))
        y = int(off.get("y", "0"))
        cx = int(ext.get("cx", "0"))
        cy = int(ext.get("cy", "0"))
        return (x, y, x + cx, y + cy)
    except ValueError:
        return None


def restyle_sp_or_cxn(el, theme_map, mapper, z_bg_stack, default_bg_lum=1.0, is_heading_shape=False, font_scale=1.0, header_size_mode=None):
    """Restyles a <p:sp> or <p:cxnSp> (fills, gradients, borders, style refs, and text runs)."""
    tag = el.tag.split("}")[-1]
    sp_pr = el.find(qn_p("spPr"))
    style_el = el.find(qn_p("style"))
    bbox = get_local_bbox(el)
    underlying_lum = get_bg_lum_under_box(bbox, z_bg_stack, default_bg_lum)

    effective_bg_lum = underlying_lum
    has_opaque_fill = False

    # 1. Restyle Fill (<p:spPr> or <p:style>/<a:fillRef>)
    if sp_pr is not None:
        no_fill = sp_pr.find(qn_a("noFill"))
        solid_fill = sp_pr.find(qn_a("solidFill"))
        grad_fill = sp_pr.find(qn_a("gradFill"))
        blip_fill = sp_pr.find(qn_a("blipFill"))

        if no_fill is not None and (solid_fill is not None or grad_fill is not None or blip_fill is not None):
            sp_pr.remove(no_fill)
            no_fill = None

        if no_fill is not None:
            effective_bg_lum = underlying_lum
        elif blip_fill is not None:
            has_opaque_fill = True
            effective_bg_lum = 0.5
        elif solid_fill is not None:
            res = resolve_color_from_parent(solid_fill, theme_map)
            if res is not None:
                rgb, alpha_val = res
                new_hex = mapper.map_fill_or_line(rgb, role="fill")
                write_srgb_into_parent(solid_fill, new_hex, alpha_val)
                fill_lum = perceived_brightness(hex_to_rgb(new_hex))
                alpha_f = (int(alpha_val) / 100000.0) if alpha_val is not None else 1.0
                effective_bg_lum = fill_lum * alpha_f + underlying_lum * (1.0 - alpha_f)
                if alpha_f > 0.25:
                    has_opaque_fill = True
        elif grad_fill is not None:
            gs_list = grad_fill.findall(f".//{{{A_NS}}}gs")
            stop_hexes = []
            for gs in gs_list:
                res = resolve_color_from_parent(gs, theme_map)
                if res is not None:
                    rgb, alpha_val = res
                    new_hex = mapper.map_fill_or_line(rgb, role="fill")
                    if stop_hexes and new_hex == stop_hexes[-1]:
                        new_hex = shift_hex_lightness(new_hex, 0.10 if perceived_brightness(rgb) > 0.3 else 0.14)
                    stop_hexes.append(new_hex)
                    write_srgb_into_parent(gs, new_hex, alpha_val)
            if stop_hexes:
                avg_lum = sum(perceived_brightness(hex_to_rgb(h)) for h in stop_hexes) / len(stop_hexes)
                effective_bg_lum = avg_lum
                has_opaque_fill = True
        elif style_el is not None:
            fill_ref = style_el.find(qn_a("fillRef"))
            if fill_ref is not None and fill_ref.get("idx", "0") != "0":
                res = resolve_color_from_parent(fill_ref, theme_map)
                if res is not None:
                    rgb, alpha_val = res
                    new_hex = mapper.map_fill_or_line(rgb, role="fill")
                    write_srgb_into_parent(fill_ref, new_hex, alpha_val)
                    # Also materialize explicit solidFill so Vorlage fmtScheme doesn't alter brightness
                    sf = etree.Element(qn_a("solidFill"))
                    write_srgb_into_parent(sf, new_hex, alpha_val)
                    insert_sppr_child(sp_pr, sf, slot_rank=2)
                    effective_bg_lum = perceived_brightness(hex_to_rgb(new_hex))
                    has_opaque_fill = True

    # 2. Restyle Border / Connector Line (<a:ln> and <p:style>/<a:lnRef>)
    ln_el = sp_pr.find(qn_a("ln")) if sp_pr is not None else None
    if ln_el is not None:
        ln_sf = ln_el.find(qn_a("solidFill"))
        if ln_sf is not None:
            res = resolve_color_from_parent(ln_sf, theme_map)
            if res is not None:
                rgb, alpha_val = res
                new_ln_hex = mapper.map_fill_or_line(rgb, role="line")
                write_srgb_into_parent(ln_sf, new_ln_hex, alpha_val)

    if style_el is not None:
        ln_ref = style_el.find(qn_a("lnRef"))
        if ln_ref is not None and ln_ref.get("idx", "0") != "0":
            res = resolve_color_from_parent(ln_ref, theme_map)
            if res is not None:
                rgb, alpha_val = res
                new_ln_hex = mapper.map_fill_or_line(rgb, role="line")
                write_srgb_into_parent(ln_ref, new_ln_hex, alpha_val)

    # 3. Resolve <p:style>/<a:fontRef> default text color & restyle <p:txBody>
    style_font_rgb = None
    if style_el is not None:
        font_ref = style_el.find(qn_a("fontRef"))
        if font_ref is not None:
            res = resolve_color_from_parent(font_ref, theme_map)
            if res is not None:
                style_font_rgb = res[0]
            write_srgb_into_parent(font_ref, IFG_WHITE if effective_bg_lum < 0.48 else IFG_CHARCOAL)

    if tag == "sp":
        txbody = el.find(qn_p("txBody"))
        restyle_txbody(
            txbody,
            theme_map=theme_map,
            mapper=mapper,
            bg_lum=effective_bg_lum,
            is_heading_shape=is_heading_shape,
            style_font_rgb=style_font_rgb,
            font_scale=font_scale,
            header_size_mode=header_size_mode
        )

    if has_opaque_fill and bbox is not None:
        z_bg_stack.append((bbox[0], bbox[1], bbox[2], bbox[3], effective_bg_lum))


def restyle_table_element(tbl_el, theme_map, mapper, font_scale=1.0):
    """Restyles an existing <a:tbl> in-place (preserving merged cells, column widths, row heights, borders)."""
    tbl_pr = tbl_el.find(qn_a("tblPr"))
    has_first_row_style = tbl_pr is not None and tbl_pr.get("firstRow") == "1" and tbl_pr.find(qn_a("tableStyleId")) is not None

    for r_idx, tr in enumerate(tbl_el.findall(qn_a("tr"))):
        for tc in tr.findall(qn_a("tc")):
            tc_pr = tc.find(qn_a("tcPr"))
            cell_bg_lum = 0.12 if (r_idx == 0 and has_first_row_style) else 1.0

            if tc_pr is not None:
                sf = tc_pr.find(qn_a("solidFill"))
                if sf is not None:
                    res = resolve_color_from_parent(sf, theme_map)
                    if res is not None:
                        rgb, alpha_val = res
                        new_hex = mapper.map_fill_or_line(rgb, role="fill")
                        write_srgb_into_parent(sf, new_hex, alpha_val)
                        cell_bg_lum = perceived_brightness(hex_to_rgb(new_hex))
                elif tc_pr.find(qn_a("noFill")) is not None:
                    cell_bg_lum = 1.0

                # Restyle explicit cell border colors cleanly
                for b_tag in ("lnL", "lnR", "lnT", "lnB"):
                    b_el = tc_pr.find(qn_a(b_tag))
                    if b_el is not None:
                        b_sf = b_el.find(qn_a("solidFill"))
                        if b_sf is not None:
                            res = resolve_color_from_parent(b_sf, theme_map)
                            if res is not None:
                                rgb, alpha_val = res
                                write_srgb_into_parent(b_sf, mapper.map_fill_or_line(rgb, role="line"), alpha_val)

            txbody = tc.find(qn_a("txBody"))
            restyle_txbody(
                txbody,
                theme_map=theme_map,
                mapper=mapper,
                bg_lum=cell_bg_lum,
                is_heading_shape=(r_idx == 0),
                style_font_rgb=(255, 255, 255) if cell_bg_lum < 0.48 else (0, 0, 0),
                font_scale=font_scale
            )


def restyle_any_element(el, theme_map, mapper, z_bg_stack, default_bg_lum=1.0, is_heading_shape=False, font_scale=1.0, header_size_mode=None):
    """Dispatches deep color & font restyling across shapes, connectors, pictures, groups, and tables."""
    tag = el.tag.split("}")[-1]
    if tag in ("sp", "cxnSp"):
        restyle_sp_or_cxn(el, theme_map, mapper, z_bg_stack, default_bg_lum, is_heading_shape, font_scale, header_size_mode=header_size_mode)
    elif tag == "pic":
        sp_pr = el.find(qn_p("spPr"))
        if sp_pr is not None:
            ln_sf = sp_pr.find(f"{{{A_NS}}}ln/{{{A_NS}}}solidFill")
            if ln_sf is not None:
                res = resolve_color_from_parent(ln_sf, theme_map)
                if res is not None:
                    write_srgb_into_parent(ln_sf, mapper.map_fill_or_line(res[0], role="line"), res[1])
    elif tag == "grpSp":
        grp_stack = []
        for child in el:
            ctag = child.tag.split("}")[-1]
            if ctag in ("sp", "cxnSp", "pic", "grpSp", "graphicFrame"):
                restyle_any_element(child, theme_map, mapper, grp_stack, default_bg_lum, False, font_scale)
    elif tag == "graphicFrame":
        tbl = el.find(f".//{{{A_NS}}}tbl")
        if tbl is not None:
            restyle_table_element(tbl, theme_map, mapper, font_scale=font_scale)


# ----------------- 5. SMARTART, CHART & RELATIONSHIP SAFE CLONER -----------------

def restyle_smartart_or_chart_xml(xml_bytes, partname_str, theme_map, mapper):
    """Updates fonts to Abadi and harmonizes colors inside cloned SmartArt / Chart XML parts."""
    try:
        root = etree.fromstring(xml_bytes)
    except Exception:
        return xml_bytes

    pname = partname_str.lower()
    old_acc1_lum = perceived_brightness(theme_map.get("accent1", (68, 114, 196)))

    # 1. Update all text fonts to Abadi (skipping symbol fonts)
    for rpr in root.findall(f".//{{{A_NS}}}rPr") + root.findall(f".//{{{A_NS}}}defRPr") + root.findall(f".//{{{A_NS}}}endParaRPr"):
        parent = rpr.getparent()
        t_el = parent.find(qn_a("t")) if parent is not None else None
        r_txt = t_el.text if t_el is not None and t_el.text else ""
        apply_corporate_font_to_rpr(rpr, r_txt)

    # 2. If SmartArt colors.xml or drawing.xml relied on a light/medium accent1 in source theme,
    # map accent1 -> accent2 (#3EB9C0 Corporate Cyan/Teal in Vorlage) so shapes stay vibrant & readable!
    if "/diagrams/" in pname and old_acc1_lum >= 0.45:
        for sc in root.findall(f".//{{{A_NS}}}schemeClr"):
            if sc.get("val") == "accent1":
                sc.set("val", "accent2")

    # 3. Map explicit <a:srgbClr> colors (excluding shadow effectLst) into corporate palette
    for parent in root.iter():
        ptag = parent.tag.split("}")[-1]
        if ptag in ("solidFill", "gs"):
            anc_tags = {a.tag.split("}")[-1] for a in parent.iterancestors()}
            if "effectLst" in anc_tags:
                continue
            srgb = parent.find(qn_a("srgbClr"))
            if srgb is not None and srgb.get("val"):
                res = resolve_color_from_parent(parent, theme_map)
                if res is not None:
                    rgb, alpha_val = res
                    # In rPr (text run), keep dark/white readability
                    if "rPr" in anc_tags or "defRPr" in anc_tags:
                        lum = perceived_brightness(rgb)
                        new_hex = IFG_WHITE if lum > 0.85 else IFG_NAVY
                    else:
                        new_hex = mapper.map_fill_or_line(rgb, role="fill")
                    write_srgb_into_parent(parent, new_hex, alpha_val)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def allocate_unique_packuri(orig_uri_str, used_partnames):
    """Generates a guaranteed unique PackURI in destination package."""
    if orig_uri_str not in used_partnames:
        used_partnames.add(orig_uri_str)
        return PackURI(orig_uri_str)

    base, ext = posixpath.splitext(orig_uri_str)
    m = re.match(r"^(.*?)(\d+)$", base)
    prefix = m.group(1) if m else base + "_"
    counter = int(m.group(2)) + 100 if m else 101
    while True:
        candidate = f"{prefix}{counter}{ext}"
        if candidate not in used_partnames:
            used_partnames.add(candidate)
            return PackURI(candidate)
        counter += 1


def make_part(uri, content_type, blob, package):
    """Constructs an OPC Part instance compatible with all python-pptx versions (0.6.x and 1.0.x)."""
    try:
        return Part(uri, content_type, package=package, blob=blob)
    except TypeError:
        return Part(uri, content_type, blob=blob, package=package)


def clone_part_into_package(old_part, dst_pkg, used_partnames, part_memo, theme_map, mapper, extra_xml_transform=None):
    """Recursively clones an OPC Part (image, SVG, SmartArt, chart, workbook, media) without name collisions."""
    key = id(old_part)
    if key in part_memo and extra_xml_transform is None:
        return part_memo[key]

    ctype = getattr(old_part, "content_type", "") or ""
    blob = getattr(old_part, "blob", b"")

    # Raster images -> deduplicated native ImagePart via SHA1
    if ctype.startswith("image/") and "svg" not in ctype and "emf" not in ctype and "wmf" not in ctype and "wdp" not in ctype:
        try:
            img_part = dst_pkg.get_or_add_image_part(io.BytesIO(blob))
            used_partnames.add(str(img_part.partname))
            part_memo[key] = img_part
            return img_part
        except Exception:
            pass

    orig_uri = str(old_part.partname)
    new_uri = allocate_unique_packuri(orig_uri, used_partnames)

    new_blob = blob
    if ctype.endswith("+xml") and ("/diagrams/" in orig_uri or "/charts/" in orig_uri):
        new_blob = restyle_smartart_or_chart_xml(new_blob, orig_uri, theme_map, mapper)
    if extra_xml_transform is not None:
        new_blob = extra_xml_transform(new_blob)

    new_part = make_part(new_uri, ctype, new_blob, dst_pkg)
    if extra_xml_transform is None:
        part_memo[key] = new_part

    # Recursively clone sub-relationships (e.g. embedded Excel in chart, images in SmartArt)
    sub_rid_map = {}
    for sub_rid, sub_rel in old_part.rels.items():
        if sub_rel.is_external:
            nr = new_part.relate_to(sub_rel.target_ref, sub_rel.reltype, is_external=True)
        else:
            cloned_sub = clone_part_into_package(
                sub_rel.target_part, dst_pkg, used_partnames, part_memo, theme_map, mapper
            )
            nr = new_part.relate_to(cloned_sub, sub_rel.reltype)
        if nr != sub_rid:
            sub_rid_map[sub_rid] = nr

    if sub_rid_map and ctype.endswith("+xml"):
        try:
            sub_root = etree.fromstring(new_part._blob)
            for node in sub_root.iter():
                for attr_k, attr_v in list(node.attrib.items()):
                    if attr_v in sub_rid_map:
                        node.set(attr_k, sub_rid_map[attr_v])
            new_part._blob = etree.tostring(sub_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        except Exception:
            pass

    return new_part


def copy_and_restyle_shape(sh, elem, old_slide, new_slide, theme_map, mapper, used_partnames, part_memo):
    """
    Transfers all rId relationships of elem from old_slide to new_slide without
    any partname or rId collisions, including SmartArt diagramDrawing parts,
    and assigns unique cNvPr IDs on new_slide.
    """
    dst_pkg = new_slide.part.package

    # Collect all rId values referenced in attributes of elem
    referenced_rids = set()
    for node in elem.iter():
        for attr_v in node.attrib.values():
            if re.match(r"^rId\d+$", attr_v):
                referenced_rids.add(attr_v)

    # SmartArt special handling: data.xml contains <dsp:dataModelExt relId="rIdX"/>
    # referencing a diagramDrawing rel on the SLIDE!
    rid_map = {}
    for rid in sorted(referenced_rids):
        if rid not in old_slide.part.rels:
            continue
        old_rel = old_slide.part.rels[rid]
        if old_rel.is_external:
            new_rid = new_slide.part.relate_to(old_rel.target_ref, old_rel.reltype, is_external=True)
            rid_map[rid] = new_rid
        elif old_rel.reltype.endswith("/diagramData"):
            # Inspect data.xml for dsp:dataModelExt relId
            data_blob = old_rel.target_part.blob
            ext_m = re.search(rb'relId="(rId\d+)"', data_blob)
            draw_rid_new = None
            draw_rid_old = None
            if ext_m:
                draw_rid_old = ext_m.group(1).decode("ascii")
                if draw_rid_old in old_slide.part.rels:
                    d_rel = old_slide.part.rels[draw_rid_old]
                    cloned_draw = clone_part_into_package(
                        d_rel.target_part, dst_pkg, used_partnames, {}, theme_map, mapper
                    )
                    draw_rid_new = new_slide.part.relate_to(cloned_draw, d_rel.reltype)

            def _patch_data_relid(b, old_r=draw_rid_old, new_r=draw_rid_new):
                if old_r and new_r and old_r != new_r:
                    return b.replace(f'relId="{old_r}"'.encode("ascii"), f'relId="{new_r}"'.encode("ascii"))
                return b

            cloned_data = clone_part_into_package(
                old_rel.target_part, dst_pkg, used_partnames, {}, theme_map, mapper,
                extra_xml_transform=_patch_data_relid
            )
            new_rid = new_slide.part.relate_to(cloned_data, old_rel.reltype)
            rid_map[rid] = new_rid
        else:
            # For diagram layout/quickStyle/colors, clone per-diagram instance cleanly
            memo_to_use = {} if "/diagrams/" in str(getattr(old_rel.target_part, "partname", "")) else part_memo
            cloned_target = clone_part_into_package(
                old_rel.target_part, dst_pkg, used_partnames, memo_to_use, theme_map, mapper
            )
            new_rid = new_slide.part.relate_to(cloned_target, old_rel.reltype)
            rid_map[rid] = new_rid

    # Rewrite rId attributes in-place on elem
    if rid_map:
        for node in elem.iter():
            for attr_k, attr_v in list(node.attrib.items()):
                if attr_v in rid_map:
                    node.set(attr_k, rid_map[attr_v])

    # Assign unique cNvPr IDs across the slide
    existing_ids = set()
    for c in new_slide.shapes._spTree.iter(qn_p("cNvPr")):
        try:
            existing_ids.add(int(c.get("id")))
        except (TypeError, ValueError):
            pass
    next_id = max(existing_ids | {1}) + 1
    for c in elem.iter(qn_p("cNvPr")):
        c.set("id", str(next_id))
        next_id += 1

    new_slide.shapes._spTree.append(elem)


def sanitize_pptx_package(buf, slide_count):
    """
    Cleans template leftovers that PowerPoint Desktop (Windows 11) and SharePoint Online refuse:
    - Deduplicates all ZIP entry filenames (prevents corrupt package warnings)
    - Removes ppt/changesInfos/* co-authoring history pointing at deleted template slides
    - Updates docProps/app.xml slide counts
    - Strips orphaned content-type overrides and broken relationships
    """
    import zipfile
    buf.seek(0)
    src = zipfile.ZipFile(buf)

    seen_names = set()
    deduped_names = []
    for n in src.namelist():
        if n not in seen_names:
            seen_names.add(n)
            deduped_names.append(n)

    drop = {
        n for n in deduped_names
        if n.startswith("ppt/changesInfos/") or n.startswith("ppt/revisionInfo") or n.startswith("[trash]/")
    }
    keep = [n for n in deduped_names if n not in drop]
    keep_set = set(keep)

    notes_count = sum(1 for n in keep if n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml"))

    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<TotalTime>0</TotalTime><Application>Microsoft Office PowerPoint</Application>'
        '<PresentationFormat>Breitbild</PresentationFormat>'
        f'<Slides>{slide_count}</Slides><Notes>{notes_count}</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips>'
        '<ScaleCrop>false</ScaleCrop><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
        '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>'
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        written_entries = set()
        order = ["[Content_Types].xml"] + [n for n in keep if n != "[Content_Types].xml"]
        for n in order:
            if n in written_entries:
                continue
            written_entries.add(n)
            data = src.read(n)
            if n == "[Content_Types].xml":
                xml = data.decode("utf-8")
                seen_overrides = set()

                def _ok(m):
                    pname = m.group(1)
                    if pname not in keep_set or pname in seen_overrides:
                        return ""
                    seen_overrides.add(pname)
                    return m.group(0)

                xml = re.sub(r'<Override PartName="/([^"]+)"[^>]*/>', _ok, xml)
                data = xml.encode("utf-8")
            elif n == "docProps/app.xml":
                data = app_xml.encode("utf-8")
            elif n.endswith(".rels"):
                xml = data.decode("utf-8")
                base_dir = posixpath.dirname(posixpath.dirname(n))

                def _rel_ok(m):
                    tag = m.group(0)
                    if 'TargetMode="External"' in tag:
                        return tag
                    t = re.search(r'Target="([^"]+)"', tag).group(1)
                    full = posixpath.normpath(posixpath.join(base_dir, t)).lstrip("/") if not t.startswith("/") else t.lstrip("/")
                    return tag if full in keep_set else ""

                xml = re.sub(r"<Relationship [^>]*/>", _rel_ok, xml)
                data = xml.encode("utf-8")
            dst.writestr(n, data)
    out.seek(0)
    return out


# ----------------- 6. 1:1 EXACT-LAYOUT CORPORATE TRANSFORMATION ENGINE -----------------

def detect_slide_heading_shape(shapes, exclude_shapes=None, old_w=None, old_h=None):
    """Finds the primary top heading/title shape on a slide (if any) without altering layout."""
    best = None
    best_score = -999
    excl = list(exclude_shapes or ())
    ref_h = float(old_h) if old_h else (7.5 * 914400.0)
    ref_w = float(old_w) if old_w else (13.333 * 914400.0)

    for sh in shapes:
        if any(sh is x for x in excl) or should_skip_source_shape(sh):
            continue
        if not (getattr(sh, "has_text_frame", False) and sh.text_frame and sh.text_frame.text.strip()):
            continue
        txt = sh.text_frame.text.strip()
        if len(txt) > 180:
            continue

        is_title_ph = False
        ph_idx = None
        if sh.is_placeholder:
            try:
                ph_idx = sh.placeholder_format.idx
                if sh.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) or ph_idx == 10:
                    is_title_ph = True
            except Exception:
                pass

        # Skip Vorlage top-strip Kapitel placeholder (idx=12 at y < 0.65") when searching for main Title
        if ph_idx == 12 and (sh.top or 0) < int(0.65 * 914400):
            continue

        # Skip tall multi-paragraph body/quote boxes that merely start near the top
        if not is_title_ph and (sh.height or 0) > 0.28 * ref_h:
            continue

        rel_top = (sh.top or 0) / ref_h
        rel_left = (sh.left or 0) / ref_w
        score = 0
        if rel_top < 0.16:
            score += 45
        elif rel_top < 0.30:
            score += 15
        else:
            score -= 25

        if rel_left < 0.20:
            score += 10
        if len(txt) < 90:
            score += 20

        max_pt = 0
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size:
                    max_pt = max(max_pt, r.font.size.pt)
        if max_pt >= 20:
            score += 25

        if is_title_ph:
            score += 50

        if score > best_score and score >= 25:
            best_score = score
            best = sh
    return best


def collect_repeated_header_kickers(user_prs, old_w, old_h):
    """
    Pre-scans content slides to identify repeating first-line kicker/chapter strings
    inside 2-line top header boxes (e.g., 'Arbeitssystemanalyse [ASystA]' across slides).
    """
    counts = Counter()
    for s_idx, slide in enumerate(user_prs.slides):
        if s_idx == 0:
            continue
        for sh in slide.shapes:
            if should_skip_source_shape(sh):
                continue
            if not (getattr(sh, "has_text_frame", False) and sh.text_frame):
                continue
            if (sh.top or 0) / old_h > 0.18 or (sh.left or 0) / old_w > 0.35:
                continue
            paras = [" ".join(p.text.split()) for p in sh.text_frame.paragraphs if p.text and p.text.strip()]
            if len(paras) >= 2 and 2 <= len(paras[0]) <= 85:
                counts[paras[0].lower()] += 1
    return {k for k, cnt in counts.items() if cnt >= 2}


def extract_slide_header_components(old_slide, source_shapes, old_w, old_h, repeated_first_lines):
    """
    Identifies top-bar 'Kapitel' (12pt) and 'Titel' (28pt) components on a content slide:
    Returns dict with:
      - use_vorlage_header: bool (True when slide uses standard top header zone)
      - kapitel_sh: shape consumed as Kapitel (or None)
      - top_title_sh: shape consumed as Titel (or None)
      - heading_sh: general heading shape on slide (may be middle section divider or banner)
      - kapitel_lines: list[str]
      - title_lines: list[str]
      - subtitle_lines: list[str]
    """
    kapitel_sh = None
    # 1. Detect dedicated Kapitel shape (idx=12 at top, or small kicker box above title)
    potential_other_top_titles = []
    for sh in source_shapes:
        if should_skip_source_shape(sh):
            continue
        if not (getattr(sh, "has_text_frame", False) and sh.text_frame and sh.text_frame.text.strip()):
            continue
        if (sh.top or 0) / old_h <= 0.18 and (sh.left or 0) / old_w <= 0.38:
            potential_other_top_titles.append(sh)

    for sh in potential_other_top_titles:
        if sh._element.tag.split("}")[-1] != "sp":
            continue
        ph_idx = None
        if sh.is_placeholder:
            try:
                ph_idx = sh.placeholder_format.idx
            except Exception:
                pass
        rel_top = (sh.top or 0) / old_h
        rel_left = (sh.left or 0) / old_w
        raw_txt = " ".join(sh.text_frame.text.split())

        # Explicit Vorlage Kapitel placeholder (idx=12 in top strip y <= 0.64")
        if ph_idx == 12 and rel_top <= 0.085:
            kapitel_sh = sh
            break

        # Non-title small kicker box at very top-left above a lower title or at y <= 0.31"
        if ph_idx != 10 and rel_top <= 0.065 and rel_left <= 0.20 and not shape_has_visual_fill(sh._element):
            non_empty_p = [p for p in sh.text_frame.paragraphs if p.text and p.text.strip()]
            max_sz = get_shape_max_font_sz(sh, old_slide)
            has_lower_title = any(
                other is not sh
                and (other.top or 0) > (sh.top or 0) + int(0.10 * 914400)
                and len(other.text_frame.text.strip()) <= 160
                for other in potential_other_top_titles
            )
            if len(raw_txt) <= 110 and len(non_empty_p) <= 2 and (max_sz == 0 or max_sz <= 1700):
                if has_lower_title or (rel_top <= 0.042 and 0 < max_sz <= 1600):
                    kapitel_sh = sh
                    break

    excl = [kapitel_sh] if kapitel_sh is not None else []
    heading_sh = detect_slide_heading_shape(source_shapes, exclude_shapes=excl, old_w=old_w, old_h=old_h)

    top_title_sh = None
    if heading_sh is not None and heading_sh._element.tag.split("}")[-1] == "sp":
        rel_top = (heading_sh.top or 0) / old_h
        rel_left = (heading_sh.left or 0) / old_w
        rel_h = (heading_sh.height or 0) / old_h
        is_title_ph = False
        if heading_sh.is_placeholder:
            try:
                pht = heading_sh.placeholder_format.type
                phi = heading_sh.placeholder_format.idx
                is_title_ph = pht in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) or phi == 10
            except Exception:
                pass

        has_bullets = any(
            heading_sh._element.find(f".//{{{A_NS}}}{bt}") is not None
            for bt in ("buChar", "buAutoNum", "buBlip")
        )
        non_empty_p = [p for p in heading_sh.text_frame.paragraphs if p.text and p.text.strip()]

        if (
            rel_top <= 0.18
            and rel_left <= 0.38
            and not shape_has_visual_fill(heading_sh._element)
            and (is_title_ph or (rel_h <= 0.26 and len(non_empty_p) <= 3 and not has_bullets))
        ):
            top_title_sh = heading_sh

    kapitel_lines = []
    title_lines = []
    subtitle_lines = []

    if kapitel_sh is not None:
        for p in kapitel_sh.text_frame.paragraphs:
            for ln in p.text.splitlines():
                cleaned = " ".join(ln.split())
                if cleaned and cleaned.lower() != "kapitel":
                    kapitel_lines.append(cleaned)

    if top_title_sh is not None:
        tmap = extract_slide_theme_map(old_slide)
        txbody = top_title_sh._element.find(qn_p("txBody"))
        lst = txbody.find(qn_a("lstStyle")) if txbody is not None else None
        lvl_sz = {}
        lvl_b = {}
        lvl_rgb = {}
        if top_title_sh.is_placeholder:
            try:
                inh_els = find_inherited_placeholder(
                    old_slide,
                    top_title_sh.placeholder_format.idx,
                    top_title_sh.placeholder_format.type,
                )
                inh_props = extract_inherited_level_props(inh_els, is_title_ph=True)
                for l_i, d in inh_props.items():
                    if d.get("sz"):
                        lvl_sz[l_i] = int(d["sz"])
                    if d.get("b"):
                        lvl_b[l_i] = d["b"]
            except Exception:
                pass
        if lst is not None:
            for l_i in range(9):
                lp = lst.find(qn_a(f"lvl{l_i + 1}pPr"))
                if lp is not None:
                    dr = lp.find(qn_a("defRPr"))
                    if dr is not None:
                        if dr.get("sz"):
                            try:
                                lvl_sz[l_i] = int(dr.get("sz"))
                            except ValueError:
                                pass
                        if dr.get("b"):
                            lvl_b[l_i] = dr.get("b")
                        res = resolve_color_from_parent(dr.find(qn_a("solidFill")), tmap)
                        if res is not None:
                            lvl_rgb[l_i] = res[0]

        para_meta = []
        if txbody is not None:
            for p_el in txbody.findall(qn_a("p")):
                ppr = p_el.find(qn_a("pPr"))
                lvl = int(ppr.get("lvl", "0")) if (ppr is not None and ppr.get("lvl", "0").isdigit()) else 0
                eff_sz = lvl_sz.get(lvl, 0)
                eff_b = lvl_b.get(lvl, "0")
                eff_rgb = lvl_rgb.get(lvl, tmap.get("tx1", (0, 0, 0)))

                runs = p_el.findall(qn_a("r"))
                run_szs = []
                for r_el in runs:
                    rpr = r_el.find(qn_a("rPr"))
                    if rpr is not None:
                        if rpr.get("sz"):
                            try:
                                run_szs.append(int(rpr.get("sz")))
                            except ValueError:
                                pass
                        if rpr.get("b") is not None:
                            eff_b = rpr.get("b")
                        res = resolve_color_from_parent(rpr.find(qn_a("solidFill")), tmap)
                        if res is not None:
                            eff_rgb = res[0]
                if run_szs:
                    eff_sz = max(run_szs)

                # Split explicit <a:br/> line breaks cleanly inside the paragraph
                chunks = [[]]
                for ch in p_el:
                    ctag = ch.tag.split("}")[-1]
                    if ctag == "r":
                        t_node = ch.find(qn_a("t"))
                        if t_node is not None and t_node.text:
                            chunks[-1].append(t_node.text)
                    elif ctag == "br":
                        chunks.append([])
                joined_chunks = [" ".join("".join(c).split()) for c in chunks if "".join(c).strip()]
                if joined_chunks:
                    para_meta.append((" |BR| ".join(joined_chunks), joined_chunks, eff_sz, eff_b, eff_rgb))

        if len(para_meta) == 1:
            title_lines.extend(para_meta[0][1])
        elif len(para_meta) == 2 and not kapitel_lines:
            (_, ch0, sz0, b0, rgb0), (_, ch1, sz1, b1, rgb1) = para_meta
            t0 = " ".join(ch0)
            t1 = " ".join(ch1)
            is_kicker_first = (
                len(ch0) == 1
                and len(t0) <= 85
                and (
                    (t0.lower() in repeated_first_lines)
                    or (sz0 > 0 and sz1 > 0 and sz0 < sz1)
                )
            )
            if is_kicker_first:
                if t0.lower() != "kapitel":
                    kapitel_lines.append(t0)
                title_lines.extend(ch1)
            elif len(t0) + len(t1) <= 68 and len(ch0) == 1 and len(ch1) == 1:
                sep = " " if t0.endswith((":", "–", "-", "?", "!")) else " – "
                title_lines.append(f"{t0}{sep}{t1}")
            else:
                title_lines.extend(ch0)
                title_lines.extend(ch1)
        elif len(para_meta) >= 2:
            if not kapitel_lines:
                first_t = " ".join(para_meta[0][1])
                if len(para_meta[0][1]) == 1 and (
                    first_t.lower() in repeated_first_lines
                    or (0 < para_meta[0][2] < para_meta[1][2])
                ):
                    if first_t.lower() != "kapitel":
                        kapitel_lines.append(first_t)
                    for pm in para_meta[1:]:
                        title_lines.extend(pm[1])
                else:
                    for pm in para_meta:
                        title_lines.extend(pm[1])
            else:
                for pm in para_meta:
                    title_lines.extend(pm[1])

    use_vorlage_header = bool(title_lines or kapitel_lines or subtitle_lines)
    return {
        "use_vorlage_header": use_vorlage_header,
        "kapitel_sh": kapitel_sh,
        "top_title_sh": top_title_sh,
        "heading_sh": heading_sh,
        "kapitel_lines": kapitel_lines,
        "title_lines": title_lines,
        "subtitle_lines": subtitle_lines,
    }


def is_url_only_text(txt):
    t = (txt or "").strip().lower()
    if not t or " " in t or "\n" in t:
        return False
    return t.startswith("www.") or t.startswith("http://") or t.startswith("https://") or t.endswith(".de") or t.endswith(".com")


def get_shape_max_font_sz(sh, old_slide):
    """Returns effective maximum font size in hundredths of a point (e.g. 3600 = 36pt)."""
    max_sz = 0
    el = sh._element
    for rpr in el.findall(f".//{{{A_NS}}}rPr") + el.findall(f".//{{{A_NS}}}defRPr"):
        if rpr.get("sz"):
            try:
                max_sz = max(max_sz, int(rpr.get("sz")))
            except ValueError:
                pass
    if max_sz == 0 and sh.is_placeholder:
        ph_idx = None
        ph_type = None
        is_title_ph = False
        try:
            ph_idx = sh.placeholder_format.idx
            ph_type = sh.placeholder_format.type
            is_title_ph = ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) or ph_idx == 10
        except Exception:
            pass
        inh_els = find_inherited_placeholder(old_slide, ph_idx, ph_type)
        lvl_props = extract_inherited_level_props(inh_els, is_title_ph=is_title_ph)
        if lvl_props.get(0, {}).get("sz"):
            try:
                max_sz = int(lvl_props[0]["sz"])
            except ValueError:
                pass
        elif is_title_ph:
            max_sz = 3200
    return max_sz


def is_cover_slide_candidate(first_slide):
    """
    Returns True when Slide 1 represents a presentation cover/title slide
    (no tables, charts, or SmartArt diagrams, and concise title/subtitle text).
    """
    root = first_slide._element
    if root.find(f".//{{{A_NS}}}tbl") is not None:
        return False
    for gd in root.findall(f".//{{{A_NS}}}graphicData"):
        uri = (gd.get("uri") or "").lower()
        if "/chart" in uri or "/diagram" in uri:
            return False

    non_url_chars = 0
    for sh in first_slide.shapes:
        if should_skip_source_shape(sh):
            continue
        if getattr(sh, "has_text_frame", False) and sh.text_frame:
            txt = sh.text_frame.text.strip()
            if txt and not is_url_only_text(txt):
                non_url_chars += len(txt)

    return 1 <= non_url_chars <= 420


def extract_cover_title_lines(first_slide):
    """
    Extracts (primary_lines, subtitle_lines) from Slide 1 of the source presentation.
    """
    candidates = []
    for sh in first_slide.shapes:
        if should_skip_source_shape(sh):
            continue
        if not (getattr(sh, "has_text_frame", False) and sh.text_frame):
            continue
        raw_txt = sh.text_frame.text.strip()
        if not raw_txt or is_url_only_text(raw_txt):
            continue

        ph_idx = None
        ph_type = None
        if sh.is_placeholder:
            try:
                ph_idx = sh.placeholder_format.idx
                ph_type = sh.placeholder_format.type
            except Exception:
                pass

        max_sz = get_shape_max_font_sz(sh, first_slide)
        score = max_sz / 50.0  # e.g. 3600 -> +72

        if ph_idx == 10 or ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            score += 120
        elif ph_type == PP_PLACEHOLDER.SUBTITLE:
            score -= 20
        elif ph_idx == 11:
            score -= 40

        if len(raw_txt) <= 140:
            score += 25
        top_in = (sh.top or 0) / 914400.0
        if top_in < 5.0:
            score += 15

        candidates.append((score, sh, ph_idx, ph_type, max_sz))

    if not candidates:
        return (["Präsentation"], [])

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, main_sh, _, _, main_max_sz = candidates[0]

    para_items = []
    for p in main_sh.text_frame.paragraphs:
        t = " ".join(p.text.split())
        if not t:
            continue
        p_sz = 0
        for r in p.runs:
            if r.font.size:
                p_sz = max(p_sz, int(round(r.font.size.pt * 100)))
        if p_sz == 0:
            p_sz = main_max_sz
        para_items.append((t, p_sz))

    top_p_sz = max((sz for _, sz in para_items), default=0)
    primary_lines = []
    subtitle_lines = []
    for t, sz in para_items:
        if primary_lines and top_p_sz > 0 and sz > 0 and (top_p_sz - sz) >= 400:
            subtitle_lines.append(t)
        else:
            primary_lines.append(t)

    # Check if there is a separate subtitle shape on Slide 1
    if not subtitle_lines and len(candidates) > 1:
        for _, sub_sh, sub_idx, sub_type, _ in candidates[1:]:
            if sub_idx == 11:
                continue
            sub_paras = [" ".join(p.text.split()) for p in sub_sh.text_frame.paragraphs if p.text.strip()]
            joined = " ".join(sub_paras)
            if not sub_paras or is_url_only_text(joined):
                continue
            if sub_type == PP_PLACEHOLDER.SUBTITLE or (len(joined) <= 100 and len(sub_paras) <= 2):
                subtitle_lines.extend(sub_paras)
                break

    if not primary_lines:
        primary_lines = [" ".join(main_sh.text_frame.text.split())]

    return (primary_lines, subtitle_lines)


def populate_vorlage_cover_slide(cover_slide, primary_lines, subtitle_lines):
    """
    Writes the presentation title (and optional subtitle) onto the official Vorlage
    Cover Slide ('Kapitel mit großem Bild') in pure White (#FFFFFF) Abadi font.
    """
    title_sp = None
    url_sp = None
    for el in cover_slide.shapes._spTree.findall(qn_p("sp")):
        ph = el.find(f".//{{{P_NS}}}ph")
        if ph is not None:
            idx_str = ph.get("idx")
            ptype = ph.get("type", "")
            if idx_str == "10" or ptype in ("title", "ctrTitle"):
                if title_sp is None or idx_str == "10":
                    title_sp = el
            elif idx_str == "11":
                url_sp = el

    # Fallback if custom template cover had no idx=10 placeholder
    if title_sp is None:
        tb = cover_slide.shapes.add_textbox(690605, 3242373, 4950000, 1780000)
        title_sp = tb._element
    else:
        ensure_explicit_xfrm(title_sp, 690605, 3242373, 4950000, 1780000)

    main_chars = sum(len(x) for x in primary_lines)
    if main_chars <= 32 and len(primary_lines) <= 2 and not subtitle_lines:
        main_sz = 3600
    elif main_chars <= 55:
        main_sz = 3000 if not subtitle_lines else 2800
    elif main_chars <= 90:
        main_sz = 2600
    else:
        main_sz = 2200
    sub_sz = max(1800, main_sz - 600)

    txbody = title_sp.find(qn_p("txBody"))
    if txbody is None:
        txbody = etree.SubElement(title_sp, qn_p("txBody"))

    body_pr = txbody.find(qn_a("bodyPr"))
    if body_pr is None:
        body_pr = etree.Element(qn_a("bodyPr"))
        txbody.insert(0, body_pr)
    body_pr.attrib.clear()
    body_pr.set("wrap", "square")
    body_pr.set("lIns", "0")
    body_pr.set("rIns", "0")
    body_pr.set("tIns", "0")
    body_pr.set("bIns", "0")
    body_pr.set("anchor", "t")
    for ch in list(body_pr):
        body_pr.remove(ch)
    etree.SubElement(body_pr, qn_a("noAutofit"))

    if txbody.find(qn_a("lstStyle")) is None:
        txbody.insert(1, etree.Element(qn_a("lstStyle")))

    for old_p in list(txbody.findall(qn_a("p"))):
        txbody.remove(old_p)

    def _append_white_para(text_str, sz_val, space_before_pts=0):
        p = etree.SubElement(txbody, qn_a("p"))
        ppr = etree.SubElement(p, qn_a("pPr"))
        ppr.set("marL", "0")
        ppr.set("lvl", "0")
        ppr.set("indent", "0")
        ppr.set("algn", "l")
        ln_spc = etree.SubElement(ppr, qn_a("lnSpc"))
        spc_pct = etree.SubElement(ln_spc, qn_a("spcPct"))
        spc_pct.set("val", "105000")
        if space_before_pts > 0:
            sb = etree.SubElement(ppr, qn_a("spcBef"))
            sp = etree.SubElement(sb, qn_a("spcPts"))
            sp.set("val", str(space_before_pts))
        etree.SubElement(ppr, qn_a("buNone"))

        r = etree.SubElement(p, qn_a("r"))
        rpr = etree.SubElement(r, qn_a("rPr"))
        rpr.set("lang", "de-DE")
        rpr.set("sz", str(sz_val))
        rpr.set("b", "0")
        rpr.set("dirty", "0")
        sf = etree.SubElement(rpr, qn_a("solidFill"))
        write_srgb_into_parent(sf, IFG_WHITE)
        for ftag in ("latin", "ea", "cs"):
            fe = etree.SubElement(rpr, qn_a(ftag))
            fe.set("typeface", CORPORATE_FONT)

        t_el = etree.SubElement(r, qn_a("t"))
        t_el.text = text_str

        ep = etree.SubElement(p, qn_a("endParaRPr"))
        ep.set("lang", "de-DE")
        ep.set("sz", str(sz_val))
        ep.set("dirty", "0")
        esf = etree.SubElement(ep, qn_a("solidFill"))
        write_srgb_into_parent(esf, IFG_WHITE)
        elat = etree.SubElement(ep, qn_a("latin"))
        elat.set("typeface", CORPORATE_FONT)

    for line in primary_lines:
        _append_white_para(line, main_sz, space_before_pts=0)

    for s_idx, s_line in enumerate(subtitle_lines):
        _append_white_para(s_line, sub_sz, space_before_pts=(800 if s_idx == 0 else 200))

    # Ensure bottom domain placeholder (www.gesundheitsmanagement.com) is also crisp White Abadi
    if url_sp is not None:
        url_tx = url_sp.find(qn_p("txBody"))
        if url_tx is not None:
            for r in url_tx.findall(f".//{{{A_NS}}}r"):
                rpr = r.find(qn_a("rPr"))
                if rpr is None:
                    rpr = etree.Element(qn_a("rPr"))
                    r.insert(0, rpr)
                for ftag in ("noFill", "solidFill", "gradFill", "pattFill"):
                    old_f = rpr.find(qn_a(ftag))
                    if old_f is not None:
                        rpr.remove(old_f)
                sf = etree.Element(qn_a("solidFill"))
                write_srgb_into_parent(sf, IFG_WHITE)
                rpr.insert(0, sf)
                apply_corporate_font_to_rpr(rpr)

    # Guarantee unique <p:cNvPr id="..."> across the entire cover slide XML (fixes native think-cell duplicate IDs)
    for idx_id, cnv in enumerate(cover_slide._element.iter(qn_p("cNvPr")), start=1):
        cnv.set("id", str(idx_id))


def create_white_logo_variant(png_blob):
    """Generates a white-on-transparent variant of the 'a.' corner logo while keeping the cyan bar intact."""
    try:
        img = Image.open(io.BytesIO(png_blob)).convert("RGBA")
        px = img.load()
        w, h = img.size
        for y in range(h):
            for x in range(w):
                r, g, b, a = px[x, y]
                if a > 15:
                    # Keep cyan bar pixels (high green/blue relative to red), turn dark navy 'a.' to crisp white
                    if perceived_brightness((r, g, b)) < 0.35 and not (g > r + 40 and b > r + 40):
                        px[x, y] = (255, 255, 255, a)
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return png_blob


def create_cyan_bg_logo_variant(png_blob):
    """Keeps Dark Navy 'a.' and turns the bottom cyan underline bar to crisp White so both pop on cyan sidebars."""
    try:
        img = Image.open(io.BytesIO(png_blob)).convert("RGBA")
        px = img.load()
        w, h = img.size
        for y in range(h):
            for x in range(w):
                r, g, b, a = px[x, y]
                if a > 15 and g > r + 35 and b > r + 35:
                    px[x, y] = (255, 255, 255, a)
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return png_blob


def extract_vorlage_brand_assets(output_prs):
    """
    Extracts the official top-right 'a.' corner logo (dark-navy, cyan-bg, and white variants),
    bottom-right slide number coordinates, and theme heading/body font references from the Vorlage.
    """
    assets = {
        "dark_logo_blob": None,
        "dark_logo_box": (11089638, 406399, 733425, 733425),
        "cyan_logo_blob": None,
        "light_logo_blob": None,
        "light_logo_box": (11089638, 406399, 733425, 733425),
        "snum_box": (11101546, 6299200, 743702, 365125),
        "heading_typeface": CORPORATE_FONT,
        "body_typeface": CORPORATE_FONT,
    }
    try:
        master = output_prs.slide_masters[0]
        # Check if Vorlage theme1.xml defines Abadi as majorFont (+mj-lt -> 'Abadi (Überschriften)') & minorFont (+mn-lt)
        for rel in master.part.rels.values():
            if not rel.is_external and rel.reltype.endswith("/theme"):
                t_root = etree.fromstring(rel.target_part.blob)
                maj_lat = t_root.find(f".//{{{A_NS}}}fontScheme/{{{A_NS}}}majorFont/{{{A_NS}}}latin")
                min_lat = t_root.find(f".//{{{A_NS}}}fontScheme/{{{A_NS}}}minorFont/{{{A_NS}}}latin")
                if maj_lat is not None and (maj_lat.get("typeface") or "").strip().lower() == CORPORATE_FONT.lower():
                    assets["heading_typeface"] = "+mj-lt"
                if min_lat is not None and (min_lat.get("typeface") or "").strip().lower() == CORPORATE_FONT.lower():
                    assets["body_typeface"] = "+mn-lt"
                break

        for el in master._element.findall(f".//{{{P_NS}}}spTree/*"):
            tag = el.tag.split("}")[-1]
            bbox = get_local_bbox(el)
            if bbox is None:
                continue
            x1, y1, x2, y2 = bbox
            if tag == "pic" and x1 >= int(9.5 * 914400) and y1 <= int(1.5 * 914400):
                blip = el.find(f".//{{{A_NS}}}blip")
                rid = blip.get(f"{{{R_NS}}}embed") if blip is not None else None
                if rid and rid in master.part.rels and not master.part.rels[rid].is_external:
                    assets["dark_logo_blob"] = master.part.rels[rid].target_part.blob
                    assets["dark_logo_box"] = (x1, y1, x2 - x1, y2 - y1)
            elif tag == "sp":
                ph = el.find(f".//{{{P_NS}}}ph")
                if ph is not None and ph.get("type") == "sldNum":
                    assets["snum_box"] = (x1, y1, x2 - x1, y2 - y1)

        # Look for official white corner logo variant in dark layouts (e.g. slideLayout17 -> image5.png)
        for lay in output_prs.slide_layouts:
            for el in lay._element.findall(f".//{{{P_NS}}}spTree/{{{P_NS}}}pic"):
                bbox = get_local_bbox(el)
                if bbox is None:
                    continue
                x1, y1, x2, y2 = bbox
                if x1 >= int(9.5 * 914400) and y1 <= int(1.5 * 914400):
                    blip = el.find(f".//{{{A_NS}}}blip")
                    rid = blip.get(f"{{{R_NS}}}embed") if blip is not None else None
                    if rid and rid in lay.part.rels and not lay.part.rels[rid].is_external:
                        cand_blob = lay.part.rels[rid].target_part.blob
                        if assets["dark_logo_blob"] is None or cand_blob != assets["dark_logo_blob"]:
                            assets["light_logo_blob"] = cand_blob
                            assets["light_logo_box"] = (x1, y1, x2 - x1, y2 - y1)
                            break
            if assets["light_logo_blob"] is not None:
                break

        if assets["dark_logo_blob"] is not None:
            assets["cyan_logo_blob"] = create_cyan_bg_logo_variant(assets["dark_logo_blob"])
        if assets["light_logo_blob"] is None and assets["dark_logo_blob"] is not None:
            assets["light_logo_blob"] = create_white_logo_variant(assets["dark_logo_blob"])
            assets["light_logo_box"] = assets["dark_logo_box"]
    except Exception:
        pass
    return assets


def populate_vorlage_content_header(
    new_slide,
    ph_titel_el,
    ph_kapitel_el,
    title_lines,
    kapitel_lines,
    subtitle_lines,
    z_bg_stack,
    slide_bg_lum,
    brand_assets,
):
    """
    Populates the official Vorlage 'Leere Folie' header placeholders on content slides:
      - Kapitel (idx=12 at x=0.47", y=0.23"): Abadi, size 12 pt (sz='1200'), muted gray (#808080) or white on dark bg
      - Titel   (idx=10 at x=0.47", y=0.50"): Abadi (Überschriften), size 28 pt (sz='2800'), IfG Navy (#012641) or white on dark bg
    Automatically narrows header width before any right-side sidebar/split image and prevents overlap with body content.
    """
    hdr_x = 431799
    kap_y = 208228
    kap_cy = 307777
    tit_y = 455313
    default_cx = 9441249  # 10.325" (ends at 10.80", safely clear of 12.13" 'a.' logo)
    max_x2 = hdr_x + default_cx

    # Check if a right-side colored sidebar or full-height split image starts before 10.80"
    sp_tree = new_slide.shapes._spTree
    top_backdrop_idx = -1
    for tree_i, el in enumerate(list(sp_tree)):
        if el is ph_titel_el or el is ph_kapitel_el:
            continue
        tag = el.tag.split("}")[-1]
        if tag not in ("sp", "pic", "grpSp", "graphicFrame"):
            continue
        bbox = get_local_bbox(el)
        if bbox is None:
            continue
        bx1, by1, bx2, by2 = bbox
        bw, bh = bx2 - bx1, by2 - by1
        is_right_split = (
            bx1 >= int(4.5 * 914400)
            and by1 <= int(1.10 * 914400)
            and bh >= int(2.4 * 914400)
            and ((tag == "sp" and shape_has_visual_fill(el)) or (tag == "pic" and by2 >= int(6.5 * 914400)))
        )
        if is_right_split:
            max_x2 = min(max_x2, bx1 - int(0.25 * 914400))

        # Track top-covering full-width header banners so Kapitel/Titel sit above them in Z-order
        if tag == "sp" and shape_has_visual_fill(el) and bx1 <= int(1.0 * 914400) and by1 <= int(0.65 * 914400) and bw >= int(5.0 * 914400):
            top_backdrop_idx = max(top_backdrop_idx, tree_i)

    header_cx = max(int(3.2 * 914400), max_x2 - hdr_x)
    header_bg_lum = get_bg_lum_under_box((hdr_x, tit_y, hdr_x + min(header_cx, int(3.5 * 914400)), tit_y + 523220), z_bg_stack, slide_bg_lum)

    titel_hex = IFG_WHITE if header_bg_lum < 0.48 else IFG_NAVY
    kapitel_hex = IFG_WHITE if header_bg_lum < 0.48 else VORLAGE_KAPITEL_GRAY
    heading_tf = brand_assets.get("heading_typeface", CORPORATE_FONT)
    body_tf = brand_assets.get("body_typeface", CORPORATE_FONT)

    # Estimate wrapped line count of Titel at 28pt (~0.148" per average character)
    chars_per_line = max(18, int((header_cx / 914400.0) / 0.148))
    est_lines = 0
    for tl in title_lines:
        est_lines += max(1, (len(tl) + chars_per_line - 1) // chars_per_line)
    if subtitle_lines:
        est_lines += len(subtitle_lines)
    titel_cy = 523220 if est_lines <= 1 else int(round((0.46 * est_lines + 0.08) * 914400))

    def _set_placeholder_xfrm(el, x, y, cx, cy, is_placeholder=True):
        sp_pr = el.find(qn_p("spPr"))
        if sp_pr is None:
            sp_pr = etree.Element(qn_p("spPr"))
            el.insert(1, sp_pr)
        for ch in list(sp_pr):
            if ch.tag.split("}")[-1] in ("xfrm", "solidFill", "gradFill", "blipFill", "pattFill", "noFill", "ln"):
                sp_pr.remove(ch)
        xfrm = etree.Element(qn_a("xfrm"))
        off = etree.SubElement(xfrm, qn_a("off"))
        off.set("x", str(int(round(x))))
        off.set("y", str(int(round(y))))
        ext = etree.SubElement(xfrm, qn_a("ext"))
        ext.set("cx", str(max(0, int(round(cx)))))
        ext.set("cy", str(max(0, int(round(cy)))))
        sp_pr.insert(0, xfrm)
        if not is_placeholder:
            if sp_pr.find(qn_a("prstGeom")) is None:
                prst = etree.Element(qn_a("prstGeom"))
                prst.set("prst", "rect")
                etree.SubElement(prst, qn_a("avLst"))
                insert_sppr_child(sp_pr, prst, slot_rank=1)
            insert_sppr_child(sp_pr, etree.Element(qn_a("noFill")), slot_rank=2)

    def _populate_header_sp(sp_el, x, y, cx, cy, main_paras, sub_paras, main_sz, main_hex, main_tf, sub_sz, sub_hex, sub_tf, is_ph):
        _set_placeholder_xfrm(sp_el, x, y, cx, cy, is_placeholder=is_ph)
        txbody = sp_el.find(qn_p("txBody"))
        if txbody is None:
            txbody = etree.SubElement(sp_el, qn_p("txBody"))

        body_pr = txbody.find(qn_a("bodyPr"))
        if body_pr is None:
            body_pr = etree.Element(qn_a("bodyPr"))
            txbody.insert(0, body_pr)
        body_pr.attrib.clear()
        body_pr.set("wrap", "square")
        body_pr.set("lIns", "0")
        body_pr.set("rIns", "0")
        body_pr.set("tIns", "0")
        body_pr.set("bIns", "0")
        body_pr.set("anchor", "t")
        for ch in list(body_pr):
            body_pr.remove(ch)
        etree.SubElement(body_pr, qn_a("noAutofit"))

        lst = txbody.find(qn_a("lstStyle"))
        if lst is None:
            lst = etree.Element(qn_a("lstStyle"))
            txbody.insert(1, lst)
        else:
            for ch in list(lst):
                lst.remove(ch)
        for lvl_i in range(1, 6):
            lp = etree.SubElement(lst, qn_a(f"lvl{lvl_i}pPr"))
            lp.set("marL", "0")
            lp.set("indent", "0")
            lp.set("algn", "l")
            etree.SubElement(lp, qn_a("buNone"))
            dr = etree.SubElement(lp, qn_a("defRPr"))
            dr.set("lang", "de-DE")
            dr.set("sz", main_sz)
            dr.set("b", "0")
            dr.set("dirty", "0")
            dsf = etree.SubElement(dr, qn_a("solidFill"))
            write_srgb_into_parent(dsf, main_hex)
            dlat = etree.SubElement(dr, qn_a("latin"))
            dlat.set("typeface", main_tf)
            dea = etree.SubElement(dr, qn_a("ea"))
            dea.set("typeface", "+mn-ea")
            dcs = etree.SubElement(dr, qn_a("cs"))
            dcs.set("typeface", CORPORATE_FONT)

        for old_p in list(txbody.findall(qn_a("p"))):
            txbody.remove(old_p)

        def _add_p(txt_val, sz_v, hex_v, tf_v, spc_bef_pts=0):
            p = etree.SubElement(txbody, qn_a("p"))
            ppr = etree.SubElement(p, qn_a("pPr"))
            ppr.set("marL", "0")
            ppr.set("indent", "0")
            ppr.set("lvl", "0")
            ppr.set("algn", "l")
            if spc_bef_pts > 0:
                sb = etree.SubElement(ppr, qn_a("spcBef"))
                sp_pts = etree.SubElement(sb, qn_a("spcPts"))
                sp_pts.set("val", str(spc_bef_pts))
            etree.SubElement(ppr, qn_a("buNone"))

            if txt_val is not None:
                r = etree.SubElement(p, qn_a("r"))
                rpr = etree.SubElement(r, qn_a("rPr"))
                rpr.set("lang", "de-DE")
                rpr.set("sz", sz_v)
                rpr.set("b", "0")
                rpr.set("dirty", "0")
                sf = etree.SubElement(rpr, qn_a("solidFill"))
                write_srgb_into_parent(sf, hex_v)
                lat = etree.SubElement(rpr, qn_a("latin"))
                lat.set("typeface", tf_v)
                ea = etree.SubElement(rpr, qn_a("ea"))
                ea.set("typeface", "+mn-ea")
                cs = etree.SubElement(rpr, qn_a("cs"))
                cs.set("typeface", CORPORATE_FONT)
                t_el = etree.SubElement(r, qn_a("t"))
                t_el.text = txt_val

            ep = etree.SubElement(p, qn_a("endParaRPr"))
            ep.set("lang", "de-DE")
            ep.set("sz", sz_v)
            ep.set("b", "0")
            ep.set("dirty", "0")
            esf = etree.SubElement(ep, qn_a("solidFill"))
            write_srgb_into_parent(esf, hex_v)
            elat = etree.SubElement(ep, qn_a("latin"))
            elat.set("typeface", tf_v)

        if main_paras:
            for m_txt in main_paras:
                _add_p(m_txt, main_sz, main_hex, main_tf)
            for s_idx, s_txt in enumerate(sub_paras):
                _add_p(s_txt, sub_sz, sub_hex, sub_tf, spc_bef_pts=(200 if s_idx == 0 else 0))
        else:
            # Empty native placeholder: shows 'Kapitel'/'Titel' prompt in Edit Mode at 12pt/28pt, hidden in Slide Show
            _add_p(None, main_sz, main_hex, main_tf)

    # Fallback textbox creation if custom template lacked idx=12 or idx=10 on Leere Folie
    kap_is_ph = ph_kapitel_el is not None
    if ph_kapitel_el is None and kapitel_lines:
        ph_kapitel_el = new_slide.shapes.add_textbox(hdr_x, kap_y, header_cx, kap_cy)._element

    tit_is_ph = ph_titel_el is not None
    if ph_titel_el is None and (title_lines or subtitle_lines):
        ph_titel_el = new_slide.shapes.add_textbox(hdr_x, tit_y, header_cx, titel_cy)._element

    kap_joined = [" – ".join(kapitel_lines)] if kapitel_lines else []
    if ph_kapitel_el is not None:
        _populate_header_sp(
            ph_kapitel_el, hdr_x, kap_y, header_cx, kap_cy,
            kap_joined, [],
            KAPITEL_FONT_SZ, kapitel_hex, body_tf,
            KAPITEL_FONT_SZ, kapitel_hex, body_tf,
            is_ph=kap_is_ph,
        )

    if ph_titel_el is not None:
        _populate_header_sp(
            ph_titel_el, hdr_x, tit_y, header_cx, titel_cy,
            title_lines, subtitle_lines,
            TITLE_FONT_SZ, titel_hex, heading_tf,
            KAPITEL_FONT_SZ, kapitel_hex, body_tf,
            is_ph=tit_is_ph,
        )

    # Ensure Kapitel & Titel sit cleanly in front of any background panels/rectangles so they are never obscured or hard to click
    for hdr_el in (ph_kapitel_el, ph_titel_el):
        if hdr_el is not None and hdr_el in sp_tree:
            sp_tree.remove(hdr_el)
            sp_tree.append(hdr_el)

    # If a 2+ line 28pt Titel extends down to ~1.36", gently clear any colliding foreground element right below it
    if est_lines >= 2 and title_lines:
        safe_title_bottom = tit_y + titel_cy + int(0.06 * 914400)
        for el in list(sp_tree):
            if el is ph_titel_el or el is ph_kapitel_el:
                continue
            tag = el.tag.split("}")[-1]
            if tag not in ("sp", "pic", "grpSp", "graphicFrame"):
                continue
            bbox = get_local_bbox(el)
            if bbox is None:
                continue
            bx1, by1, bx2, by2 = bbox
            bw, bh = bx2 - bx1, by2 - by1
            if bh >= int(5.4 * 914400) or by1 < int(0.55 * 914400):
                continue
            if bx1 < hdr_x + header_cx and bx2 > hdr_x and by1 < safe_title_bottom:
                dy_nudge = safe_title_bottom - by1
                if by2 + dy_nudge <= int(7.18 * 914400):
                    ensure_explicit_xfrm(el, bx1, by1 + dy_nudge, bw, bh)

    return [el for el in (ph_kapitel_el, ph_titel_el) if el is not None]


def sample_pic_region_lum(pic_el, slide, target_box):
    """Samples average luminance of an embedded raster picture at target_box (tx1, ty1, tx2, ty2)."""
    try:
        pb = get_local_bbox(pic_el)
        if pb is None:
            return None
        px1, py1, px2, py2 = pb
        pw, ph = px2 - px1, py2 - py1
        if pw <= 0 or ph <= 0:
            return None
        blip = pic_el.find(f".//{{{A_NS}}}blip")
        rid = blip.get(f"{{{R_NS}}}embed") if blip is not None else None
        if not rid or rid not in slide.part.rels or slide.part.rels[rid].is_external:
            return None
        img = Image.open(io.BytesIO(slide.part.rels[rid].target_part.blob)).convert("RGB")
        iw, ih = img.size
        u1 = max(0.0, min(1.0, (target_box[0] - px1) / pw))
        v1 = max(0.0, min(1.0, (target_box[1] - py1) / ph))
        u2 = max(0.0, min(1.0, (target_box[2] - px1) / pw))
        v2 = max(0.0, min(1.0, (target_box[3] - py1) / ph))
        if u2 <= u1 or v2 <= v1:
            return None
        crop = img.crop((int(u1 * iw), int(v1 * ih), max(int(u1 * iw) + 1, int(u2 * iw)), max(int(v1 * ih) + 1, int(v2 * ih))))
        crop = crop.resize((16, 16))
        pixels = list(crop.getdata())
        return sum(perceived_brightness(rgb) for rgb in pixels) / len(pixels)
    except Exception:
        return None


def append_front_slide_number(slide, snum_box, slide_num_int, text_hex):
    """Appends a crisp Abadi slide-number field on top of any right-side background bar/photo."""
    sx, sy, scx, scy = snum_box
    tb = slide.shapes.add_textbox(int(sx), int(sy), int(scx), int(scy))
    sp = tb._element
    sp_pr = sp.find(qn_p("spPr"))
    if sp_pr is not None and sp_pr.find(qn_a("noFill")) is None:
        insert_sppr_child(sp_pr, etree.Element(qn_a("noFill")), slot_rank=2)

    txbody = sp.find(qn_p("txBody"))
    body_pr = txbody.find(qn_a("bodyPr"))
    body_pr.attrib.clear()
    body_pr.set("vert", "horz")
    body_pr.set("lIns", "91440")
    body_pr.set("tIns", "45720")
    body_pr.set("rIns", "91440")
    body_pr.set("bIns", "45720")
    body_pr.set("rtlCol", "0")
    body_pr.set("anchor", "b")

    for old_p in list(txbody.findall(qn_a("p"))):
        txbody.remove(old_p)

    p = etree.SubElement(txbody, qn_a("p"))
    ppr = etree.SubElement(p, qn_a("pPr"))
    ppr.set("algn", "r")
    etree.SubElement(ppr, qn_a("buNone"))

    fld = etree.SubElement(p, qn_a("fld"))
    fld.set("id", "{B6F15528-F159-4107-2D62-56D95E031A43}")
    fld.set("type", "slidenum")
    rpr = etree.SubElement(fld, qn_a("rPr"))
    rpr.set("lang", "de-DE")
    rpr.set("sz", "1000")
    rpr.set("b", "0")
    rpr.set("dirty", "0")
    sf = etree.SubElement(rpr, qn_a("solidFill"))
    write_srgb_into_parent(sf, text_hex)
    lat = etree.SubElement(rpr, qn_a("latin"))
    lat.set("typeface", CORPORATE_FONT)
    t_el = etree.SubElement(fld, qn_a("t"))
    t_el.text = str(slide_num_int)

    ep = etree.SubElement(p, qn_a("endParaRPr"))
    ep.set("lang", "de-DE")
    ep.set("sz", "1000")
    ep.set("dirty", "0")


def resolve_logo_and_sidebar_layout(new_slide, slide_bg_lum, brand_assets, slide_num_int, heading_elem=None, protected_elems=None):
    """
    Ensures harmonious coexistence of right-side colored bars, portraits/images, sidebar text,
    the official top-right Vorlage 'a.' logo, and the bottom-right slide number:
    1. Foreground content colliding with the top-right 'a.' logo (e.g. a row of portraits)
       shifts cleanly below the logo as an aligned group without overlapping lower elements.
    2. Text inside a right-side colored sidebar automatically centers within the open window
       between top/bottom elements with balanced corporate margins.
    3. Whenever a colored sidebar, banner, or full-bleed backdrop covers the SlideMaster's
       corner logo or slide number, places the official 'a.' logo (dark or white for contrast)
       and slide number cleanly on the frontmost Z-layer.
    """
    protected_list = list(protected_elems or ())
    if heading_elem is not None:
        protected_list.append(heading_elem)

    lx1, ly1, lcx, lcy = brand_assets["dark_logo_box"]
    lx2, ly2 = lx1 + lcx, ly1 + lcy
    safe_lx1 = lx1 - int(0.10 * 914400)
    safe_lx2 = lx2 + int(0.25 * 914400)
    safe_ly2 = ly2 + int(0.15 * 914400)  # ~1.40" clear below 'a.' logo

    snx1, sny1, sncx, sncy = brand_assets["snum_box"]
    snx2, sny2 = snx1 + sncx, sny1 + sncy
    safe_sny1 = sny1 - int(0.08 * 914400)  # ~6.81" clear above slide number

    items = []
    for el in list(new_slide.shapes._spTree):
        tag = el.tag.split("}")[-1]
        if tag not in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp"):
            continue
        bbox = get_local_bbox(el)
        if bbox is None:
            continue
        x1, y1, x2, y2 = bbox
        w, h = x2 - x1, y2 - y1
        if w <= 15000 and h <= 15000:
            continue

        is_backdrop = False
        if tag == "sp" and shape_has_visual_fill(el):
            if w >= int(2.2 * 914400) and h >= int(1.4 * 914400) and x2 >= lx1 and y1 <= ly1 + int(0.30 * 914400):
                is_backdrop = True
        elif tag in ("pic", "grpSp", "graphicFrame"):
            if y1 <= int(0.20 * 914400) and y2 >= int(7.10 * 914400) and w >= int(3.6 * 914400) and x2 >= lx1:
                is_backdrop = True

        ph_node = el.find(f".//{{{P_NS}}}ph")
        is_hdr_ph = ph_node is not None and ph_node.get("idx") in ("10", "12")

        items.append({
            "el": el,
            "tag": tag,
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "w": w, "h": h,
            "is_backdrop": is_backdrop,
            "is_heading": any(el is p for p in protected_list) or is_hdr_ph,
        })

    def _overlaps(ax1, ay1, ax2, ay2, bx1, by1, bx2, by2):
        return ax2 > bx1 and ax1 < bx2 and ay2 > by1 and ay1 < by2

    # 1. Reflow foreground elements colliding with the top-right 'a.' logo
    fg_items = [it for it in items if not it["is_backdrop"] and not it["is_heading"]]
    colliding_fg = [
        it for it in fg_items
        if _overlaps(it["x1"], it["y1"], it["x2"], it["y2"], safe_lx1, int(0.10 * 914400), safe_lx2, safe_ly2)
    ]

    # Handle wide left-anchored tables/boxes that barely graze safe_lx1 horizontally
    right_colliders = []
    for it in colliding_fg:
        if it["x1"] < int(4.2 * 914400) and it["w"] > int(7.2 * 914400):
            if it["y1"] >= int(0.95 * 914400) and it["x2"] <= safe_lx1 + int(0.65 * 914400):
                new_w = max(int(3.0 * 914400), safe_lx1 - it["x1"])
                it["w"] = new_w
                it["x2"] = it["x1"] + new_w
                ensure_explicit_xfrm(it["el"], it["x1"], it["y1"], it["w"], it["h"])
            elif it["y2"] + (safe_ly2 - it["y1"]) <= safe_sny1:
                dy_push = safe_ly2 - it["y1"]
                it["y1"] += dy_push
                it["y2"] += dy_push
                ensure_explicit_xfrm(it["el"], it["x1"], it["y1"], it["w"], it["h"])
        else:
            right_colliders.append(it)

    if right_colliders:
        # Determine the right-column horizontal span (matching sidebar panel if present, else row span)
        right_panels = [
            it for it in items
            if it["is_backdrop"] and it["tag"] == "sp" and it["x1"] >= int(4.5 * 914400)
        ]
        if right_panels:
            rp = right_panels[0]
            col_x1 = rp["x1"] - int(0.15 * 914400)
            col_x2 = rp["x2"] + int(0.15 * 914400)
        else:
            row_mates = list(right_colliders)
            for cand in fg_items:
                if cand["x1"] >= int(4.2 * 914400):
                    for rc in right_colliders:
                        v_ov = min(cand["y2"], rc["y2"]) - max(cand["y1"], rc["y1"])
                        if v_ov >= 0.45 * min(max(1, cand["h"]), max(1, rc["h"])) or abs(cand["y1"] - rc["y1"]) <= int(0.22 * 914400):
                            if cand not in row_mates:
                                row_mates.append(cand)
            col_x1 = min(m["x1"] for m in row_mates) - int(0.10 * 914400)
            col_x2 = max(m["x2"] for m in row_mates) + int(0.10 * 914400)

        col_fg = [
            it for it in fg_items
            if col_x1 <= (it["x1"] + it["x2"]) / 2.0 <= col_x2
        ]
        col_fg.sort(key=lambda z: (z["y1"], z["x1"]))

        # Group horizontally-aligned items (e.g., 3 circular portraits) into row bands
        bands = []
        for it in col_fg:
            placed = False
            for b in bands:
                b_h = max(1, b["y2"] - b["y1"])
                v_ov = min(b["y2"], it["y2"]) - max(b["y1"], it["y1"])
                if v_ov >= 0.45 * min(b_h, max(1, it["h"])) or abs(it["y1"] - b["y1"]) <= int(0.20 * 914400):
                    b["members"].append(it)
                    b["y1"] = min(b["y1"], it["y1"])
                    b["y2"] = max(b["y2"], it["y2"])
                    b["x2"] = max(b["x2"], it["x2"])
                    placed = True
                    break
            if not placed:
                bands.append({
                    "members": [it],
                    "y1": it["y1"],
                    "y2": it["y2"],
                    "x2": it["x2"],
                })

        if bands and bands[0]["y1"] < safe_ly2:
            target_top = safe_ly2
            min_gap = int(0.15 * 914400)
            touches_snum_x = bool(right_panels) or any(b["x2"] > safe_lx1 for b in bands)
            max_bottom = safe_sny1 if touches_snum_x else int(7.30 * 914400)

            for b in bands:
                b["orig_h"] = max(1, b["y2"] - b["y1"])

            # Try pure downward shift/cascade without shrinking heights first
            cur_y = target_top
            trial_positions = []
            for idx_b, b in enumerate(bands):
                ny1 = max(b["y1"], cur_y) if idx_b > 0 else target_top
                ny2 = ny1 + b["orig_h"]
                trial_positions.append((ny1, ny2))
                cur_y = ny2 + min_gap

            h_scale = 1.0
            if trial_positions[-1][1] > max_bottom:
                # Pack tighter or gently scale heights so nothing touches the slide number or bottom edge
                sum_h = sum(b["orig_h"] for b in bands)
                avail_h = max(int(1.5 * 914400), max_bottom - target_top)
                req_gaps = min_gap * max(0, len(bands) - 1)
                if sum_h + req_gaps <= avail_h:
                    extra = (avail_h - sum_h - req_gaps) / max(1, len(bands))
                    cur_y = target_top
                    trial_positions = []
                    for b in bands:
                        ny1 = cur_y
                        ny2 = ny1 + b["orig_h"]
                        trial_positions.append((ny1, ny2))
                        cur_y = ny2 + min_gap + extra
                else:
                    h_scale = max(0.75, (avail_h - req_gaps) / max(1, sum_h))
                    cur_y = target_top
                    trial_positions = []
                    for b in bands:
                        bh = int(round(b["orig_h"] * h_scale))
                        trial_positions.append((cur_y, cur_y + bh))
                        cur_y += bh + min_gap

            for b, (ny1, ny2) in zip(bands, trial_positions):
                old_by1 = b["y1"]
                for m in b["members"]:
                    rel_y = (m["y1"] - old_by1) * h_scale
                    new_my1 = ny1 + rel_y
                    new_mh = m["h"] * h_scale
                    new_mw = m["w"]
                    new_mx1 = m["x1"]
                    if m["tag"] == "pic" and h_scale < 0.995:
                        new_mw = m["w"] * h_scale
                        new_mx1 = (m["x1"] + m["x2"]) / 2.0 - new_mw / 2.0
                    m["x1"] = int(round(new_mx1))
                    m["y1"] = int(round(new_my1))
                    m["w"] = int(round(new_mw))
                    m["h"] = int(round(new_mh))
                    m["x2"] = m["x1"] + m["w"]
                    m["y2"] = m["y1"] + m["h"]
                    ensure_explicit_xfrm(m["el"], m["x1"], m["y1"], m["w"], m["h"])

    # 2. If a right-side colored panel itself contains text (like Screenshot 1's sidebar!),
    # position its internal text frame into the largest open vertical window between top/bottom photos & logo
    for rp in [it for it in items if it["is_backdrop"] and it["tag"] == "sp" and it["x1"] >= int(4.5 * 914400)]:
        txbody = rp["el"].find(qn_p("txBody"))
        if txbody is None:
            continue
        raw_t = "".join(t.text or "" for t in txbody.iter(qn_a("t"))).strip()
        if not raw_t:
            continue

        zone_y1 = max(rp["y1"] + int(0.18 * 914400), safe_ly2 if rp["x2"] > safe_lx1 else rp["y1"] + int(0.18 * 914400))
        zone_y2 = min(rp["y2"] - int(0.18 * 914400), safe_sny1 if rp["x2"] > safe_lx1 else rp["y2"] - int(0.18 * 914400))

        # Subtract foreground shapes sitting on top of rp
        pad = int(0.12 * 914400)
        blocked = []
        for fg in fg_items:
            hov = min(rp["x2"], fg["x2"]) - max(rp["x1"], fg["x1"])
            if hov > 0.30 * max(1, fg["w"]):
                blocked.append((fg["y1"] - pad, fg["y2"] + pad))
        blocked.sort()

        free_windows = []
        cur = zone_y1
        for by1, by2 in blocked:
            if by1 > cur:
                free_windows.append((cur, min(zone_y2, by1)))
            cur = max(cur, by2)
        if cur < zone_y2:
            free_windows.append((cur, zone_y2))

        if free_windows:
            win_y1, win_y2 = max(free_windows, key=lambda w_iv: w_iv[1] - w_iv[0])
            if (win_y2 - win_y1) >= int(0.75 * 914400):
                body_pr = txbody.find(qn_a("bodyPr"))
                if body_pr is None:
                    body_pr = etree.Element(qn_a("bodyPr"))
                    txbody.insert(0, body_pr)
                body_pr.set("tIns", str(max(0, int(round(win_y1 - rp["y1"])))))
                body_pr.set("bIns", str(max(0, int(round(rp["y2"] - win_y2)))))
                min_lr = int(0.32 * 914400)
                for lr_k in ("lIns", "rIns"):
                    try:
                        cur_lr = int(body_pr.get(lr_k, "91440"))
                    except ValueError:
                        cur_lr = 91440
                    if cur_lr < min_lr:
                        body_pr.set(lr_k, str(min_lr))
                if blocked or body_pr.get("anchor") in (None, "ctr"):
                    body_pr.set("anchor", "ctr")

    # 3. Check if any shape covers the SlideMaster's top-right 'a.' logo or bottom-right Slide Number
    def _effective_color_at_box(box_tuple):
        bx1, by1, bx2, by2 = box_tuple
        cx = (bx1 + bx2) / 2.0
        cy = (by1 + by2) / 2.0
        lum = slide_bg_lum
        rgb_hit = (255, 255, 255)
        covered = False
        for it in items:
            if _overlaps(it["x1"], it["y1"], it["x2"], it["y2"], bx1, by1, bx2, by2):
                covered = True
                if it["x1"] <= cx <= it["x2"] and it["y1"] <= cy <= it["y2"]:
                    if it["tag"] == "sp" and shape_has_visual_fill(it["el"]):
                        sf = it["el"].find(f"{{{P_NS}}}spPr/{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr")
                        gf = it["el"].find(f"{{{P_NS}}}spPr/{{{A_NS}}}gradFill//{{{A_NS}}}srgbClr")
                        clr_node = sf if sf is not None else gf
                        if clr_node is not None and clr_node.get("val"):
                            rgb_hit = hex_to_rgb(clr_node.get("val"))
                            lum = perceived_brightness(rgb_hit)
                    elif it["tag"] == "pic":
                        sampled = sample_pic_region_lum(it["el"], new_slide, box_tuple)
                        if sampled is not None:
                            lum = sampled
                            rgb_hit = (int(lum * 255), int(lum * 255), int(lum * 255))
        return covered, lum, rgb_hit

    logo_covered, logo_lum, logo_rgb = _effective_color_at_box((lx1, ly1, lx2, ly2))
    if logo_covered and (brand_assets["dark_logo_blob"] is not None or brand_assets["light_logo_blob"] is not None):
        is_cyan_bg = (0.42 <= logo_lum <= 0.72) and (logo_rgb[1] > logo_rgb[0] + 35 and logo_rgb[2] > logo_rgb[0] + 35)
        if logo_lum < 0.42 and brand_assets["light_logo_blob"] is not None:
            blob = brand_assets["light_logo_blob"]
            bx, by, bcx, bcy = brand_assets["light_logo_box"]
        elif is_cyan_bg and brand_assets.get("cyan_logo_blob") is not None:
            blob = brand_assets["cyan_logo_blob"]
            bx, by, bcx, bcy = brand_assets["dark_logo_box"]
        else:
            blob = brand_assets["dark_logo_blob"]
            bx, by, bcx, bcy = brand_assets["dark_logo_box"]
        if blob is not None:
            new_slide.shapes.add_picture(io.BytesIO(blob), int(bx), int(by), int(bcx), int(bcy))

    snum_covered, snum_lum, _ = _effective_color_at_box((snx1, sny1, snx2, sny2))
    if snum_covered:
        sn_hex = IFG_WHITE if snum_lum < 0.48 else IFG_NAVY
        append_front_slide_number(new_slide, (snx1, sny1, sncx, sncy), slide_num_int, sn_hex)

    # Guarantee unique <p:cNvPr id="..."> across the entire slide XML
    for idx_id, cnv in enumerate(new_slide._element.iter(qn_p("cNvPr")), start=1):
        cnv.set("id", str(idx_id))


def build_arsipa_presentation(template_file, user_file):
    """
    Converts every slide (N -> N) from user_file into template_file (Vorlage):
    - Slide 1 (Cover): Places the presentation title onto the official Vorlage Cover Slide
      ('Kapitel mit großem Bild') in crisp White (#FFFFFF) Abadi text.
    - Slides 2..N: Preserves 100% of the exact original layout, shapes, groups, tables,
      images, SmartArt diagrams, and charts while restyling ONLY fonts (Abadi) and colors
      (IfG / Arsipa Corporate Design).
    """
    output_prs = Presentation(template_file)
    user_prs = Presentation(user_file)

    catalog = scan_arsipa_template(output_prs)
    canvas_layout = output_prs.slide_layouts[catalog["canvas_leere_folie"]]
    cover_layout = output_prs.slide_layouts[catalog["cover_layout"]]
    brand_assets = extract_vorlage_brand_assets(output_prs)

    # Pre-scan deck colors for consistent 1:1 corporate color mapping
    mapper = DeckColorMapper(user_prs)

    user_slides = list(user_prs.slides)
    use_cover_for_slide_1 = len(user_slides) > 0 and is_cover_slide_candidate(user_slides[0])

    # Clear template sample slides safely (preserving Vorlage Slide 1 when used as Cover Slide!)
    slide_ids = list(output_prs.slides._sldIdLst)
    keep_template_slide_1 = use_cover_for_slide_1 and len(slide_ids) > 0
    ids_to_drop = slide_ids[1:] if keep_template_slide_1 else slide_ids
    for sld_id in ids_to_drop:
        r_id = sld_id.rId
        output_prs.part.drop_rel(r_id)
        output_prs.slides._sldIdLst.remove(sld_id)

    # Strip stale template section lists referencing deleted slide IDs
    for ext in output_prs.part._element.xpath("./p:extLst/p:ext"):
        if ext.xpath('./*[local-name()="sectionLst"]'):
            ext.getparent().remove(ext)

    # Compute aspect-ratio-safe coordinate mapping from user_prs -> Vorlage
    old_w = float(user_prs.slide_width)
    old_h = float(user_prs.slide_height)
    new_w = float(output_prs.slide_width)
    new_h = float(output_prs.slide_height)

    sx = new_w / old_w
    sy = new_h / old_h
    scale = min(sx, sy)
    dx = (new_w - old_w * scale) / 2.0
    dy = (new_h - old_h * scale) / 2.0
    max_header_right = int(11.85 * 914400)  # Keep top headers clear of top-right Vorlage logo

    repeated_corner_logos = detect_repeated_corner_logos(user_prs, scale, dx, dy)
    repeated_header_kickers = collect_repeated_header_kickers(user_prs, old_w, old_h)

    pkg = output_prs.part.package
    pkg_parts = list(pkg.iter_parts()) if hasattr(pkg, "iter_parts") else getattr(pkg, "parts", [])
    used_partnames = {str(p.partname) for p in pkg_parts}
    part_memo = {}
    processed_slides = []

    for idx, old_slide in enumerate(user_slides):
        if idx == 0 and use_cover_for_slide_1:
            cover_slide = output_prs.slides[0] if keep_template_slide_1 else output_prs.slides.add_slide(cover_layout)
            primary_lines, subtitle_lines = extract_cover_title_lines(old_slide)
            populate_vorlage_cover_slide(cover_slide, primary_lines, subtitle_lines)
            full_cover_title = " — ".join(primary_lines + subtitle_lines)
            processed_slides.append({
                "index": idx,
                "title": full_cover_title or "Titelfolie",
                "layout_name": cover_slide.slide_layout.name,
                "reason": "Offizielle IfG-Titelfolie (Kapitel mit großem Bild) • Titel in Weiß (#FFFFFF) & Abadi"
            })
            continue

        new_slide = output_prs.slides.add_slide(canvas_layout)

        source_shapes = list(old_slide.shapes)
        hdr_info = extract_slide_header_components(old_slide, source_shapes, old_w, old_h, repeated_header_kickers)
        use_vorlage_header = hdr_info["use_vorlage_header"]
        heading_sh = hdr_info["heading_sh"]

        # Manage cloned layout placeholders on new_slide:
        # Keep official Vorlage Titel (idx=10, 28pt) & Kapitel (idx=12, 12pt) when slide uses top header zone;
        # always remove idx=3 (slide number) and unused placeholders so no ghost boxes overlap user shapes.
        ph_titel_el = None
        ph_kapitel_el = None
        for ph in list(new_slide.placeholders):
            ph_idx = None
            try:
                ph_idx = ph.placeholder_format.idx
            except Exception:
                pass
            if use_vorlage_header and ph_idx == 10 and ph_titel_el is None:
                ph_titel_el = ph._element
                continue
            if use_vorlage_header and ph_idx == 12 and ph_kapitel_el is None:
                ph_kapitel_el = ph._element
                continue
            try:
                new_slide.shapes._spTree.remove(ph._element)
            except Exception:
                pass

        theme_map = extract_slide_theme_map(old_slide)
        slide_bg_lum = 1.0

        # If old_slide has an explicit custom slide background (<p:bg>), restyle its color too
        old_bg = old_slide._element.find(f"{{{P_NS}}}cSld/{{{P_NS}}}bg")
        if old_bg is not None:
            new_bg = deepcopy(old_bg)
            bg_sf = new_bg.find(f".//{{{A_NS}}}solidFill")
            if bg_sf is not None:
                res = resolve_color_from_parent(bg_sf, theme_map)
                if res is not None:
                    mapped_bg = mapper.map_fill_or_line(res[0], role="fill")
                    write_srgb_into_parent(bg_sf, mapped_bg, res[1])
                    slide_bg_lum = perceived_brightness(hex_to_rgb(mapped_bg))
                    csld = new_slide._element.find(qn_p("cSld"))
                    existing_bg = csld.find(qn_p("bg"))
                    if existing_bg is not None:
                        csld.remove(existing_bg)
                    csld.insert(0, new_bg)

        slide_title_label = " ".join(hdr_info["title_lines"]).strip()
        if not slide_title_label and heading_sh is not None and getattr(heading_sh, "has_text_frame", False):
            slide_title_label = " ".join(heading_sh.text_frame.text.split())

        z_bg_stack = []
        copied_count = 0
        heading_elem_on_new = None

        # 0. Materialize any colored sidebar / header banner panels defined on old_slide.slide_layout
        for lay_panel_el, (lx, ly, lw, lh), lay_obj in extract_layout_background_panels(old_slide, old_w, old_h):
            p_elem = deepcopy(lay_panel_el)
            ensure_explicit_xfrm(p_elem, lx * scale + dx, ly * scale + dy, lw * scale, lh * scale)
            restyle_any_element(
                p_elem,
                theme_map=theme_map,
                mapper=mapper,
                z_bg_stack=z_bg_stack,
                default_bg_lum=slide_bg_lum,
                is_heading_shape=False,
                font_scale=scale
            )
            copy_and_restyle_shape(
                None, p_elem, lay_obj, new_slide,
                theme_map=theme_map,
                mapper=mapper,
                used_partnames=used_partnames,
                part_memo=part_memo
            )
            copied_count += 1

        for sh in source_shapes:
            if should_skip_source_shape(
                sh, slide=old_slide, repeated_corner_logos=repeated_corner_logos, scale=scale, dx=dx, dy=dy
            ):
                continue

            if sh is hdr_info["kapitel_sh"] and not hdr_info["kapitel_lines"]:
                continue

            if use_vorlage_header:
                # Consumed into official Vorlage Kapitel (12pt) & Titel (28pt) placeholders
                if sh is hdr_info["kapitel_sh"] or sh is hdr_info["top_title_sh"]:
                    continue
                # Skip duplicate tiny empty left-edge accent bar when Leere Folie already provides Rectangle 3
                if (
                    sh._element.tag.split("}")[-1] == "sp"
                    and not (getattr(sh, "has_text_frame", False) and sh.text_frame and sh.text_frame.text.strip())
                    and ((sh.left or 0) * scale + dx) <= int(0.08 * 914400)
                    and int(0.35 * 914400) <= ((sh.top or 0) * scale + dy) <= int(0.90 * 914400)
                    and (sh.width or 0) * scale <= int(0.75 * 914400)
                    and (sh.height or 0) * scale <= int(0.35 * 914400)
                ):
                    continue

            if not slide_title_label and getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
                slide_title_label = " ".join(sh.text_frame.text.split())[:60]

            elem = deepcopy(sh._element)

            is_top_left_hdr = (
                sh is heading_sh
                and (sh.top or 0) / old_h <= 0.18
                and (sh.left or 0) / old_w <= 0.25
            )
            hdr_sz_mode = "title" if (sh is heading_sh and (sh.top or 0) / old_h <= 0.18) else None

            bbox = materialize_placeholder_and_transform(
                sh, elem, old_slide,
                scale=scale, dx=dx, dy=dy,
                is_top_left_header=is_top_left_hdr,
                max_header_right=max_header_right
            )

            restyle_any_element(
                elem,
                theme_map=theme_map,
                mapper=mapper,
                z_bg_stack=z_bg_stack,
                default_bg_lum=slide_bg_lum,
                is_heading_shape=(sh is heading_sh),
                font_scale=scale,
                header_size_mode=hdr_sz_mode
            )

            copy_and_restyle_shape(
                sh, elem, old_slide, new_slide,
                theme_map=theme_map,
                mapper=mapper,
                used_partnames=used_partnames,
                part_memo=part_memo
            )
            if sh is heading_sh:
                heading_elem_on_new = elem
            copied_count += 1

        hdr_elems = []
        if use_vorlage_header:
            hdr_elems = populate_vorlage_content_header(
                new_slide,
                ph_titel_el=ph_titel_el,
                ph_kapitel_el=ph_kapitel_el,
                title_lines=hdr_info["title_lines"],
                kapitel_lines=hdr_info["kapitel_lines"],
                subtitle_lines=hdr_info["subtitle_lines"],
                z_bg_stack=z_bg_stack,
                slide_bg_lum=slide_bg_lum,
                brand_assets=brand_assets,
            )
            copied_count += len(hdr_elems)

        # Harmonize right-sidebar bars, foreground portraits/boxes, and front-layer Vorlage 'a.' logo + slide number
        resolve_logo_and_sidebar_layout(
            new_slide,
            slide_bg_lum=slide_bg_lum,
            brand_assets=brand_assets,
            slide_num_int=idx + 1,
            heading_elem=heading_elem_on_new,
            protected_elems=hdr_elems,
        )

        processed_slides.append({
            "index": idx,
            "title": slide_title_label or f"Folie {idx + 1}",
            "layout_name": canvas_layout.name,
            "reason": f"1:1 Original-Aufbau erhalten ({copied_count} Elemente) • Schriftart Abadi & IfG-Farben angewendet"
        })

    out_buf = io.BytesIO()
    output_prs.save(out_buf)
    out_buf = sanitize_pptx_package(out_buf, len(processed_slides))

    return out_buf, processed_slides


# ----------------- STREAMLIT UI: 1-CLICK CLEAN WORKFLOW -----------------

logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="IfG Logo" class="ifg-logo-img" />' if logo_b64 else '<strong style="color: #012641; font-size: 1.3rem;">IfG</strong>'

st.markdown(f"""
<div class="ifg-top-nav">
    <div class="ifg-nav-left">
        {logo_html}
        <div class="ifg-nav-divider"></div>
        <div>
            <div class="ifg-nav-title">
                IfG SlideStyler
                <span class="ifg-badge-tool">Internes Tool</span>
            </div>
            <p class="ifg-nav-subtitle">PowerPoint Vorlagen-Engine nach offiziellem Corporate Design</p>
        </div>
    </div>
    <div class="ifg-nav-badge">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
            <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
        </svg>
        <span>100% DSGVO-konform (Lokale Verarbeitung)</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ifg-card">
    <div class="ifg-card-header">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#012641" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
            <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
        </svg>
        <h3 class="ifg-card-title">1:1 Layout-Erhaltung & Corporate Design (Nur Farben & Schriftart)</h3>
    </div>
    <ul class="ifg-feature-list">
        <li class="ifg-feature-item">
            <div class="ifg-bullet-dot"></div>
            <div><strong>Exakte 1:1 Anordnung</strong>: Alle Grafiken, Formen, Gruppen, Tabellen, Bilder und SmartArt-Diagramme bleiben exakt in ihrer ursprünglichen Struktur und Position erhalten.</div>
        </li>
        <li class="ifg-feature-item">
            <div class="ifg-bullet-dot"></div>
            <div><strong>Offizielle Typografie (<code>Abadi</code>)</strong>: Sämtliche Texte (inkl. Tabellen, Gruppen und Diagramme) werden auf <code>Abadi</code> umgestellt, während Symbole & Icons erhalten bleiben.</div>
        </li>
        <li class="ifg-feature-item">
            <div class="ifg-bullet-dot"></div>
            <div><strong>Intelligente IfG / Arsipa Farbpalette</strong>: Alte Farben werden kontrastgetreu in Deep Navy (<code>#012641</code>), Corporate Teal (<code>#00A8B7</code> / <code>#3EB9C0</code>) und Slate Charcoal (<code>#3A4A63</code>) überführt.</div>
        </li>
        <li class="ifg-feature-item">
            <div class="ifg-bullet-dot"></div>
            <div><strong>100% Editierbar & Kompatibel</strong>: Alle Folien (egal ob 10 oder 50+ Seiten) bleiben vollständig nativ bearbeitbar und öffnen fehlerfrei unter Windows 11 & SharePoint Online.</div>
        </li>
    </ul>
</div>
""", unsafe_allow_html=True)

DEFAULT_TEMPLATES = {
    "⭐ IfG Vorlage Kurzfassung (Empfohlen – 6,7 MB, schnelle Ladezeit)": "Arsipa Vorlage/ifg_PowerPoint_Vorlage_Kurzfassung.pptx",
    "🏢 IfG Vorlage Vollversion (47 Layouts – 37 MB)": "Arsipa Vorlage/Vorlage_Power_Point_Arsipa.pptx",
    "📁 Eigene Vorlagendatei (.pptx) hochladen": None
}

template_choice = st.selectbox(
    "1. Wählen Sie die gewünschte Unternehmensvorlage aus (Template Selection):",
    list(DEFAULT_TEMPLATES.keys()),
    index=0
)

template_file_to_use = None

if DEFAULT_TEMPLATES[template_choice] is not None:
    template_path = DEFAULT_TEMPLATES[template_choice]
    if os.path.exists(template_path):
        template_file_to_use = template_path
        st.success(f"✅ Vorlage „{os.path.basename(template_path)}“ erfolgreich geladen (Einsatzbereit)")
    else:
        st.error(f"Vorlagendatei nicht gefunden unter: {template_path}")
else:
    template_uploaded = st.file_uploader("Laden Sie Ihre eigene Vorlagendatei (.pptx) hoch:", type=["pptx"])
    if template_uploaded:
        template_file_to_use = template_uploaded

user_file = st.file_uploader("2. Ausgangspräsentation hochladen, die konvertiert werden soll (Ihre .pptx-Datei):", type=["pptx"])

if template_file_to_use and user_file:
    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    current_key = f"{getattr(user_file, 'name', '')}_{getattr(user_file, 'size', 0)}_{template_choice}"
    if st.session_state.get('last_processed_file') != current_key:
        st.session_state.pop('transformed_bytes', None)
        st.session_state.pop('transformed_slides', None)
        st.session_state.pop('transformed_filename', None)

    if st.button("🚀 Präsentation jetzt konvertieren (1-Click Transform)"):
        with st.spinner("⚡ Folien werden 1:1 in die Vorlage übertragen (Farben & Schriftarten werden angepasst)..."):
            try:
                out_buf, processed_slides = build_arsipa_presentation(template_file_to_use, user_file)
                st.session_state['transformed_bytes'] = out_buf.getvalue()
                st.session_state['transformed_slides'] = processed_slides
                st.session_state['transformed_filename'] = f"IfG_{user_file.name}"
                st.session_state['last_processed_file'] = current_key
            except Exception as ex:
                import traceback
                tb_str = traceback.format_exc()
                st.error(f"Verarbeitungsfehler aufgetreten: {ex}")
                st.code(tb_str)
                print(tb_str)

    if 'transformed_bytes' in st.session_state:
        processed_slides = st.session_state['transformed_slides']
        out_bytes = st.session_state['transformed_bytes']
        out_filename = st.session_state['transformed_filename']

        st.markdown(f"""
        <div class="ifg-success-box">
            <div class="ifg-success-title">🎉 Konvertierung erfolgreich abgeschlossen!</div>
            <p class="ifg-success-desc">
                Alle <strong>{len(processed_slides)} Folien</strong> wurden im exakten 1:1 Original-Layout in die IfG / Arsipa Vorlage übertragen – mit angepasster <code>Abadi</code>-Typografie und offiziellen Unternehmensfarben.
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="📥 Konvertierte Präsentation herunterladen (.pptx)",
            data=out_bytes,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        with st.expander("🔍 Folien-Übersicht anzeigen"):
            for s in processed_slides:
                st.markdown(f"**Folie {s['index']+1}:** `{s['title'][:55]}` ➔ {s['reason']}")
else:
    st.info("Bitte wählen Sie eine Vorlage aus und laden Sie Ihre Ausgangspräsentation hoch, um die Konvertierung zu starten.")
