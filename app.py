"""
Informa Connect Session Slide Builder

Generate professional holding slides for events.
Usage: streamlit run app.py
"""

import base64
import io
import os
import tempfile
from pathlib import Path

import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
SLIDE_ASPECT = SLIDE_WIDTH_IN / SLIDE_HEIGHT_IN  # ~1.778 (16:9)

COLORS = {
    "title_text_dark": "283857",  # Carbon
    "title_text_light": "FFFFFF",  # White
    "speaker_text_dark": "283857",
    "speaker_text_light": "FFFFFF",
    "stage_text_dark": "283857",
    "stage_text_light": "FFFFFF",
    "panel_bg": "E3F4FF",
    "accent": "003CB2",
    "default_bg": "F5F5F5",
    "cta_bg": "002244",  # Indigo
}

CONTENT_PANEL = {
    "left": 1.5,
    "top": 2.5,
    "width": 10.333,
    "height": 3.5,
}

STAGE_TEXTBOX = {
    "left": 0.5,
    "bottom_offset": 0.5,
    "width": 4.0,
    "height": 0.6,
}

BRAND_LOGO_PLACEMENT = {
    "left": 0.56,
    "top": 0.34,
    "max_width": 2.68,
    "max_height": 0.64,
}

INFORMA_LOGO_PLACEMENT = {
    "left": 11.79,
    "top": 6.63,
    "max_width": 0.98,
    "max_height": 0.31,
}

TITLE_FONT = "Open Sans"
SPEAKER_FONT = "Open Sans"
STAGE_FONT = "Open Sans"
TITLE_MAX_PT = 36
TITLE_MIN_PT = 24
SPEAKER_MAX_PT = 21
SPEAKER_MIN_PT = 16
STAGE_PT = 18

ASSETS_DIR = Path(__file__).parent / "Informa Logo"
INFORMA_LOGO_LIGHT = ASSETS_DIR / "Informa_Logo_OneLine_Graduated_White_RGB.png"
INFORMA_LOGO_DARK = ASSETS_DIR / "Informa_Logo_OneLine_Graduated_Indigo_RGB.png"


# -----------------------------------------------------------------------------
# DATA LOADING
# -----------------------------------------------------------------------------

def parse_is_moderator(value):
    if not value:
        return False
    return str(value).strip().lower() in ("yes", "true", "1", "y")


def load_events(csv_path):
    df = pd.read_csv(csv_path)
    events = df.groupby("event_name").agg({"brand": "first"}).reset_index()
    return events.to_dict("records")


def load_sessions_for_event(csv_path, event_name):
    df = pd.read_csv(csv_path)
    df = df[df["event_name"] == event_name]
    sessions = {}
    for _, row in df.iterrows():
        session_id = str(row.get("session_id", "")).strip()
        if not session_id:
            continue
        if session_id not in sessions:
            sessions[session_id] = {
                "session_id": session_id,
                "session_title": str(row.get("session_title", "")).strip(),
                "stage": str(row.get("stage", "")).strip(),
                "speakers": [],
            }
        sessions[session_id]["speakers"].append({
            "name": str(row.get("speaker_name", "")).strip(),
            "job_title": str(row.get("job_title", "")).strip(),
            "company": str(row.get("company", "")).strip(),
            "is_moderator": parse_is_moderator(row.get("is_moderator", "")),
        })
    return list(sessions.values())


def get_event_brand(csv_path, event_name):
    df = pd.read_csv(csv_path)
    event_row = df[df["event_name"] == event_name]
    if not event_row.empty:
        return str(event_row.iloc[0].get("brand", "Unknown")).strip()
    return "Unknown"


# -----------------------------------------------------------------------------
# SLIDE GENERATION
# -----------------------------------------------------------------------------

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    return RGBColor.from_string(hex_str)


def calculate_font_size(text, max_size, min_size, available_chars):
    if len(text) <= available_chars:
        return max_size
    ratio = available_chars / max(len(text), 1)
    return max(min_size, min(max_size, int(max_size * ratio)))


def fit_content(session, panel_height_in=CONTENT_PANEL["height"]):
    title = session["session_title"]
    speakers = session["speakers"]
    speakers = sorted(speakers, key=lambda s: (not s["is_moderator"], s["name"]))
    panel_width = CONTENT_PANEL["width"]
    title_max_chars = int(panel_width * 15 * 2)
    title_size = calculate_font_size(title, TITLE_MAX_PT, TITLE_MIN_PT, title_max_chars)
    if not speakers:
        return title_size, SPEAKER_MAX_PT
    speaker_lines = []
    for spk in speakers:
        line_len = len(spk["name"])
        if spk["job_title"]:
            line_len += len(spk["job_title"]) + 2
        if spk["company"]:
            line_len += len(spk["company"]) + 2
        if spk["is_moderator"]:
            line_len += 11
        lines_needed = max(1, (line_len + 24) // 25)
        speaker_lines.append(lines_needed)
    total_speaker_lines = sum(speaker_lines)
    speaker_area = panel_height_in * 0.6
    line_height_in = 0.25
    max_lines = int(speaker_area / line_height_in)
    if total_speaker_lines <= max_lines:
        return title_size, SPEAKER_MAX_PT
    ratio = max_lines / max(total_speaker_lines, 1)
    speaker_size = max(SPEAKER_MIN_PT, int(SPEAKER_MAX_PT * ratio))
    return title_size, speaker_size


def add_slide_with_branding(prs, session, brand_logo_bytes=None, informa_logo_path=None,
                            background_image_bytes=None, background_color=None,
                            panel_color=None, panel_opacity=100,
                            panel_text_color=None, stage_text_color=None,
                            include_stage=True):
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    slide_width = Inches(SLIDE_WIDTH_IN)
    slide_height = Inches(SLIDE_HEIGHT_IN)

    # Background
    if background_image_bytes:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(background_image_bytes)
            tmp_path = tmp.name
        try:
            pic = slide.shapes.add_picture(tmp_path, 0, 0, width=slide_width, height=slide_height)
            sp_tree = slide.shapes._spTree
            sp_tree.remove(pic._element)
            sp_tree.insert(2, pic._element)
        except Exception as e:
            st.warning(f"Could not add background image: {e}")
        finally:
            os.unlink(tmp_path)
    elif background_color:
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_width, slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = hex_to_rgb(background_color.lstrip("#"))
        bg.line.fill.background()
        sp_tree = slide.shapes._spTree
        sp_tree.remove(bg._element)
        sp_tree.insert(2, bg._element)
    else:
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_width, slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = hex_to_rgb(COLORS["default_bg"])
        bg.line.fill.background()
        sp_tree = slide.shapes._spTree
        sp_tree.remove(bg._element)
        sp_tree.insert(2, bg._element)

    # Brand logo
    if brand_logo_bytes:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(brand_logo_bytes)
            tmp_path = tmp.name
        try:
            img = Image.open(io.BytesIO(brand_logo_bytes))
            img_w, img_h = img.size
            aspect = img_w / img_h
            max_w = BRAND_LOGO_PLACEMENT["max_width"]
            max_h = BRAND_LOGO_PLACEMENT["max_height"]
            if aspect >= max_w / max_h:
                new_w = Inches(max_w)
                new_h = Inches(max_w / aspect)
            else:
                new_w = Inches(max_h * aspect)
                new_h = Inches(max_h)
            slide.shapes.add_picture(
                tmp_path,
                Inches(BRAND_LOGO_PLACEMENT["left"]),
                Inches(BRAND_LOGO_PLACEMENT["top"]),
                width=new_w, height=new_h
            )
        except Exception as e:
            st.warning(f"Could not add brand logo: {e}")
        finally:
            os.unlink(tmp_path)

    # Informa logo
    if informa_logo_path and os.path.exists(informa_logo_path):
        try:
            img = Image.open(informa_logo_path)
            img_w, img_h = img.size
            aspect = img_w / img_h
            max_w = INFORMA_LOGO_PLACEMENT["max_width"]
            max_h = INFORMA_LOGO_PLACEMENT["max_height"]
            if aspect >= max_w / max_h:
                new_w = Inches(max_w)
                new_h = Inches(max_w / aspect)
            else:
                new_w = Inches(max_h * aspect)
                new_h = Inches(max_h)
            slide.shapes.add_picture(
                informa_logo_path,
                Inches(INFORMA_LOGO_PLACEMENT["left"]),
                Inches(INFORMA_LOGO_PLACEMENT["top"]),
                width=new_w, height=new_h
            )
        except Exception as e:
            st.warning(f"Could not add Informa logo: {e}")

    # Stage name - conditional with independent color
    if include_stage and session["stage"]:
        sh = Inches(SLIDE_HEIGHT_IN)
        left = Inches(STAGE_TEXTBOX["left"])
        top = sh - Inches(STAGE_TEXTBOX["bottom_offset"]) - Inches(STAGE_TEXTBOX["height"])
        textbox = slide.shapes.add_textbox(left, top, Inches(STAGE_TEXTBOX["width"]), Inches(STAGE_TEXTBOX["height"]))
        tf = textbox.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = session["stage"].upper()
        p.font.size = Pt(STAGE_PT)
        p.font.name = STAGE_FONT
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(stage_text_color or COLORS["stage_text_light"])
        p.alignment = PP_ALIGN.LEFT

    # Content panel with configurable color and opacity
    from pptx.enum.dml import MSO_THEME_COLOR
    panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(CONTENT_PANEL["left"]),
        Inches(CONTENT_PANEL["top"]),
        Inches(CONTENT_PANEL["width"]),
        Inches(CONTENT_PANEL["height"]),
    )
    panel.fill.solid()
    panel.fill.fore_color.rgb = hex_to_rgb((panel_color or COLORS["panel_bg"]).lstrip("#"))
    # Apply transparency (0-100%, where 100% = opaque, 0% = fully transparent)
    # In PowerPoint, transparency is stored as 0-100000 (where 0 = opaque)
    transparency = int((100 - panel_opacity) * 1000)  # Convert to PowerPoint units
    if transparency > 0:
        panel.fill.fore_color.brightness = 0.0  # Reset brightness
        # Set transparency via the color's alpha/tint
        panel.fill.transparency = (100 - panel_opacity) / 100.0
    panel.line.fill.background()

    speakers = sorted(session["speakers"], key=lambda s: (not s["is_moderator"], s["name"]))
    title_size, speaker_size = fit_content(session)

    # Session title with configurable text color (independent)
    title_left = Inches(CONTENT_PANEL["left"] + 0.3)
    title_top = Inches(CONTENT_PANEL["top"] + 0.3)
    title_width = Inches(CONTENT_PANEL["width"] - 0.6)
    title_height = Inches(CONTENT_PANEL["height"] * 0.4)
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = session["session_title"]
    p.font.size = Pt(title_size)
    p.font.name = TITLE_FONT
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(panel_text_color or COLORS["title_text_dark"])
    p.alignment = PP_ALIGN.LEFT

    # Speakers with configurable text color and proper bolding (uses same panel_text_color)
    if speakers:
        speaker_left = title_left
        speaker_top = Inches(CONTENT_PANEL["top"] + CONTENT_PANEL["height"] * 0.45)
        speaker_width = title_width
        speaker_height = Inches(CONTENT_PANEL["height"] * 0.5)
        speaker_box = slide.shapes.add_textbox(speaker_left, speaker_top, speaker_width, speaker_height)
        stf = speaker_box.text_frame
        stf.word_wrap = True
        for i, speaker in enumerate(speakers):
            if i == 0:
                sp = stf.paragraphs[0]
            else:
                sp = stf.add_paragraph()
            sp.space_before = Pt(6)
            sp.clear()
            # Moderator label (not bold)
            if speaker["is_moderator"]:
                r = sp.add_run()
                r.text = "Moderator: "
                r.font.size = Pt(speaker_size)
                r.font.name = SPEAKER_FONT
                r.font.color.rgb = hex_to_rgb(panel_text_color or COLORS["speaker_text_dark"])
            # Name (bold)
            r = sp.add_run()
            r.text = speaker["name"]
            r.font.size = Pt(speaker_size)
            r.font.name = SPEAKER_FONT
            r.font.bold = True
            r.font.color.rgb = hex_to_rgb(panel_text_color or COLORS["speaker_text_dark"])
            # Job title (not bold)
            if speaker["job_title"]:
                r = sp.add_run()
                r.text = f", {speaker['job_title']}"
                r.font.size = Pt(speaker_size)
                r.font.name = SPEAKER_FONT
                r.font.color.rgb = hex_to_rgb(panel_text_color or COLORS["speaker_text_dark"])
            # Company (bold)
            if speaker["company"]:
                r = sp.add_run()
                r.text = f", {speaker['company']}"
                r.font.size = Pt(speaker_size)
                r.font.name = SPEAKER_FONT
                r.font.bold = True
                r.font.color.rgb = hex_to_rgb(panel_text_color or COLORS["speaker_text_dark"])

    return slide


def generate_slides(csv_path, event_name, brand_logo_bytes=None,
                    informa_logo_path=None, background_image_bytes=None,
                    background_color=None, panel_color=None, panel_opacity=100,
                    panel_text_color=None, stage_text_color=None,
                    include_stage=True):
    sessions = load_sessions_for_event(csv_path, event_name)
    prs = Presentation()
    prs.slide_width = Emu(int(Inches(SLIDE_WIDTH_IN)))
    prs.slide_height = Emu(int(Inches(SLIDE_HEIGHT_IN)))
    for session in sessions:
        add_slide_with_branding(
            prs, session,
            brand_logo_bytes=brand_logo_bytes,
            informa_logo_path=informa_logo_path,
            background_image_bytes=background_image_bytes,
            background_color=background_color,
            panel_color=panel_color,
            panel_opacity=panel_opacity,
            panel_text_color=panel_text_color,
            stage_text_color=stage_text_color,
            include_stage=include_stage,
        )
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output, len(sessions)


# -----------------------------------------------------------------------------
# PIL PREVIEW RENDERING
# -----------------------------------------------------------------------------

def hex_to_rgb_tuple(hex_str):
    hex_str = hex_str.lstrip("#")
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))


def render_slide_preview_pil(target_width_px, session, brand_logo_bytes, informa_logo_path,
                               background_image_bytes=None, background_color=None,
                               panel_color=None, panel_opacity=100,
                               title_text_color=None, speaker_text_color=None, stage_text_color=None):
    """Render a slide preview using PIL Image at exact slide proportions with accurate text styling."""
    target_height_px = int(target_width_px / SLIDE_ASPECT)

    # Background
    if background_image_bytes:
        try:
            bg = Image.open(io.BytesIO(background_image_bytes))
            bg = bg.resize((target_width_px, target_height_px), Image.LANCZOS)
        except:
            bg_color = hex_to_rgb_tuple(background_color or COLORS["default_bg"])
            bg = Image.new("RGB", (target_width_px, target_height_px), bg_color)
    elif background_color:
        bg_color = hex_to_rgb_tuple(background_color.lstrip("#"))
        bg = Image.new("RGB", (target_width_px, target_height_px), bg_color)
    else:
        bg_color = hex_to_rgb_tuple(COLORS["default_bg"])
        bg = Image.new("RGB", (target_width_px, target_height_px), bg_color)

    draw = ImageDraw.Draw(bg, 'RGBA')

    # Content panel with opacity
    panel_left = int((CONTENT_PANEL["left"] / SLIDE_WIDTH_IN) * target_width_px)
    panel_top = int((CONTENT_PANEL["top"] / SLIDE_HEIGHT_IN) * target_height_px)
    panel_right = int(((CONTENT_PANEL["left"] + CONTENT_PANEL["width"]) / SLIDE_WIDTH_IN) * target_width_px)
    panel_bottom = int(((CONTENT_PANEL["top"] + CONTENT_PANEL["height"]) / SLIDE_HEIGHT_IN) * target_height_px)
    panel_rgb = hex_to_rgb_tuple(panel_color or COLORS["panel_bg"])
    alpha = int(255 * panel_opacity / 100)
    draw.rounded_rectangle(
        [panel_left, panel_top, panel_right, panel_bottom],
        radius=12,
        fill=panel_rgb + (alpha,)
    )

    # Load fonts with bold variants
    try:
        title_font = ImageFont.truetype("OpenSans-Bold.ttf", int(target_width_px * 0.04))
        speaker_font_regular = ImageFont.truetype("OpenSans-Regular.ttf", int(target_width_px * 0.023))
        speaker_font_bold = ImageFont.truetype("OpenSans-Bold.ttf", int(target_width_px * 0.023))
        stage_font = ImageFont.truetype("OpenSans-Bold.ttf", int(target_width_px * 0.024))
    except:
        try:
            title_font = ImageFont.truetype("arial.ttf", int(target_width_px * 0.04))
            speaker_font_regular = ImageFont.truetype("arial.ttf", int(target_width_px * 0.023))
            speaker_font_bold = ImageFont.truetype("arialbd.ttf", int(target_width_px * 0.023))
            stage_font = ImageFont.truetype("arialbd.ttf", int(target_width_px * 0.024))
        except:
            title_font = ImageFont.load_default()
            speaker_font_regular = ImageFont.load_default()
            speaker_font_bold = ImageFont.load_default()
            stage_font = ImageFont.load_default()

    # Session title (bold) - positioned inside the panel with proper margins
    title_x = panel_left + int(target_width_px * 0.025)
    title_y = panel_top + int(target_height_px * 0.08)
    title = session.get("session_title", "")
    title_color = hex_to_rgb_tuple(title_text_color or COLORS["title_text_dark"])

    # Simple text wrapping for title to avoid overflow
    max_title_width = panel_right - panel_left - int(target_width_px * 0.05)
    words = title.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = current_line + " " + word if current_line else word
        bbox = draw.textbbox((0, 0), test_line, font=title_font)
        if bbox[2] - bbox[0] <= max_title_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)

    # Draw wrapped title
    line_height = int(target_height_px * 0.06)
    for i, line in enumerate(lines[:2]):  # Max 2 lines
        draw.text((title_x, title_y + i * line_height), line, font=title_font, fill=title_color)

    # Speakers with proper bolding - start below title area
    speakers = sorted(session.get("speakers", []), key=lambda s: (not s["is_moderator"], s["name"]))
    speaker_y = panel_top + int(target_height_px * 0.28)
    text_color = hex_to_rgb_tuple(speaker_text_color or COLORS["speaker_text_dark"])
    line_spacing = int(target_height_px * 0.055)
    space_width = int(target_width_px * 0.008)

    for speaker in speakers:
        current_x = title_x

        # Moderator label (not bold)
        if speaker["is_moderator"]:
            mod_text = "Moderator: "
            draw.text((current_x, speaker_y), mod_text, font=speaker_font_regular, fill=text_color)
            bbox = draw.textbbox((current_x, speaker_y), mod_text, font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0]) + space_width

        # Name (bold)
        name = speaker["name"]
        draw.text((current_x, speaker_y), name, font=speaker_font_bold, fill=text_color)
        bbox = draw.textbbox((current_x, speaker_y), name, font=speaker_font_bold)
        current_x += (bbox[2] - bbox[0])

        # Job title (not bold)
        if speaker["job_title"]:
            comma = ", "
            draw.text((current_x, speaker_y), comma, font=speaker_font_regular, fill=text_color)
            bbox = draw.textbbox((current_x, speaker_y), comma, font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0])
            draw.text((current_x, speaker_y), speaker["job_title"], font=speaker_font_regular, fill=text_color)
            bbox = draw.textbbox((current_x, speaker_y), speaker["job_title"], font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0])

        # Company (bold)
        if speaker["company"]:
            comma = ", "
            draw.text((current_x, speaker_y), comma, font=speaker_font_regular, fill=text_color)
            bbox = draw.textbbox((current_x, speaker_y), comma, font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0])
            draw.text((current_x, speaker_y), speaker["company"], font=speaker_font_bold, fill=text_color)

        speaker_y += line_spacing

    # Brand logo (top-left)
    if brand_logo_bytes:
        try:
            logo = Image.open(io.BytesIO(brand_logo_bytes))
            max_logo_w = int((BRAND_LOGO_PLACEMENT["max_width"] / SLIDE_WIDTH_IN) * target_width_px)
            max_logo_h = int((BRAND_LOGO_PLACEMENT["max_height"] / SLIDE_HEIGHT_IN) * target_height_px)
            aspect = logo.width / logo.height
            if max_logo_w / aspect <= max_logo_h:
                logo_w = max_logo_w
                logo_h = int(logo_w / aspect)
            else:
                logo_h = max_logo_h
                logo_w = int(logo_h * aspect)
            logo = logo.resize((logo_w, logo_h), Image.LANCZOS)
            logo_x = int((BRAND_LOGO_PLACEMENT["left"] / SLIDE_WIDTH_IN) * target_width_px)
            logo_y = int((BRAND_LOGO_PLACEMENT["top"] / SLIDE_HEIGHT_IN) * target_height_px)
            if logo.mode == 'RGBA':
                bg.paste(logo, (logo_x, logo_y), logo)
            else:
                bg.paste(logo, (logo_x, logo_y))
        except:
            pass

    # Informa logo (bottom-right)
    if informa_logo_path and os.path.exists(informa_logo_path):
        try:
            informa = Image.open(informa_logo_path)
            max_logo_w = int((INFORMA_LOGO_PLACEMENT["max_width"] / SLIDE_WIDTH_IN) * target_width_px)
            max_logo_h = int((INFORMA_LOGO_PLACEMENT["max_height"] / SLIDE_HEIGHT_IN) * target_height_px)
            aspect = informa.width / informa.height
            if max_logo_w / aspect <= max_logo_h:
                logo_w = max_logo_w
                logo_h = int(logo_w / aspect)
            else:
                logo_h = max_logo_h
                logo_w = int(logo_h * aspect)
            informa = informa.resize((logo_w, logo_h), Image.LANCZOS)
            logo_x = int((INFORMA_LOGO_PLACEMENT["left"] / SLIDE_WIDTH_IN) * target_width_px)
            logo_y = int((INFORMA_LOGO_PLACEMENT["top"] / SLIDE_HEIGHT_IN) * target_height_px)
            if informa.mode == 'RGBA':
                bg.paste(informa, (logo_x, logo_y), informa)
            else:
                bg.paste(informa, (logo_x, logo_y))
        except:
            pass

    # Stage name (bottom-left)
    if session.get("stage"):
        stage_text = session["stage"].upper()
        stage_x = int((STAGE_TEXTBOX["left"] / SLIDE_WIDTH_IN) * target_width_px)
        stage_y = int(((SLIDE_HEIGHT_IN - STAGE_TEXTBOX["bottom_offset"] - 0.2) / SLIDE_HEIGHT_IN) * target_height_px)
        stage_color_rgb = hex_to_rgb_tuple(stage_text_color or COLORS["stage_text_dark"])
        draw.text((stage_x, stage_y), stage_text, font=stage_font, fill=stage_color_rgb)

    return bg


# -----------------------------------------------------------------------------
# STREAMLIT INTERFACE
# -----------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Informa Connect Session Slide Builder", layout="wide")

    # Large professional title
    st.markdown("<h1 style='font-size: 2.5rem; margin-bottom: 0.5rem;'>Informa Connect Session Slide Builder</h1>", unsafe_allow_html=True)

    # Unequal columns - controls thinner, preview larger
    left_col, right_col = st.columns([1, 2])

    # Initialize session state
    if "csv_path" not in st.session_state:
        st.session_state.csv_path = "sample_sessions.csv"

    # Left column - Compact Controls
    with left_col:
        # Data & Event Card
        with st.container(border=True):
            st.markdown("**Data & Event**")

            use_sample = st.checkbox("Use sample data", value=True)

            if not use_sample:
                uploaded_csv = st.file_uploader("CSV file", type=["csv"], label_visibility="collapsed")
                if uploaded_csv:
                    st.session_state.csv_path = "/tmp/uploaded_sessions.csv"
                    with open(st.session_state.csv_path, "wb") as f:
                        f.write(uploaded_csv.getvalue())

            csv_path = st.session_state.csv_path if use_sample else st.session_state.csv_path
            selected_event = None
            event_brand = None

            if csv_path:
                try:
                    events = load_events(csv_path)
                    if events:
                        event_names = [e["event_name"] for e in events]
                        selected_event = st.selectbox(
                            "Event",
                            event_names,
                        )
                        event_brand = get_event_brand(csv_path, selected_event)
                    else:
                        st.warning("No events found")
                except Exception as e:
                    st.error(f"Could not load: {e}")

        # Branding Card - compressed layout
        with st.container(border=True):
            st.markdown("**Branding**")

            # 1. Brand Logo (compact)
            brand_logo_file = st.file_uploader(
                "Brand logo",
                type=["png", "jpg", "jpeg"],
                help="Appears top-left on slides",
            )
            brand_logo_bytes = brand_logo_file.getvalue() if brand_logo_file else None
            if brand_logo_bytes:
                st.image(brand_logo_bytes, width=50)

            # 2. Informa Logo
            informa_choice = st.segmented_control(
                "Informa logo",
                ["Light", "Dark", "None"],
                default="Light",
            )
            informa_logo_path = None
            if informa_choice == "Light":
                informa_logo_path = str(INFORMA_LOGO_LIGHT) if INFORMA_LOGO_LIGHT.exists() else None
            elif informa_choice == "Dark":
                informa_logo_path = str(INFORMA_LOGO_DARK) if INFORMA_LOGO_DARK.exists() else None

            # 3. Background choice
            bg_choice = st.segmented_control(
                "Background",
                ["Default", "Color", "Image"],
                default="Default",
            )
            bg_image_bytes = None
            bg_color = None

            # 4. COLORS SECTION - All in one row
            st.markdown("**Colors**")
            color_cols = st.columns(4)
            with color_cols[0]:
                if bg_choice == "Color":
                    bg_color = st.color_picker("Background", value="#F5F5F5")
                else:
                    bg_color = st.color_picker("Background", value="#F5F5F5", disabled=True)
                    bg_color = None
            with color_cols[1]:
                panel_color = st.color_picker("Panel", value="#" + COLORS["panel_bg"])
            with color_cols[2]:
                panel_text_color = st.color_picker("Panel text", value="#" + COLORS["title_text_dark"])
            with color_cols[3]:
                stage_text_color = st.color_picker("Stage text", value="#" + COLORS["stage_text_light"])

            # Advanced settings in expander
            with st.expander("Advanced"):
                panel_opacity = st.slider(
                    "Panel opacity",
                    min_value=0,
                    max_value=100,
                    value=100,
                    step=5,
                    format="%d%%",
                    help="Only affects panel background, text stays fully opaque",
                )
                if bg_choice == "Image":
                    bg_file = st.file_uploader("Background image", type=["png", "jpg", "jpeg"])
                    if bg_file:
                        bg_image_bytes = bg_file.getvalue()

            # 5. Stage toggle (compact)
            include_stage = st.checkbox("Show stage name", value=True)

        # Generate with Indigo CTA color
        if selected_event:
            st.markdown("<style>.stButton>button {background-color: #002244; border-color: #002244;}</style>", unsafe_allow_html=True)
            generate_clicked = st.button(
                "Generate Slides",
                type="primary",
                use_container_width=True,
            )

            if generate_clicked:
                with st.spinner("Generating..."):
                    try:
                        output_bytes, num_slides = generate_slides(
                            csv_path, selected_event,
                            brand_logo_bytes=brand_logo_bytes,
                            informa_logo_path=informa_logo_path,
                            background_image_bytes=bg_image_bytes,
                            background_color=bg_color,
                            panel_color=panel_color,
                            panel_opacity=panel_opacity,
                            panel_text_color=panel_text_color if 'panel_text_color' in locals() else COLORS["title_text_dark"],
                            stage_text_color=stage_text_color if 'stage_text_color' in locals() else COLORS["stage_text_light"],
                            include_stage=include_stage,
                        )
                        st.success(f"Generated {num_slides} slides!")
                        st.download_button(
                            label="Download PPTX",
                            data=output_bytes,
                            file_name=f"{selected_event.replace(' ', '_')}_holding_slides.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")

    # Right column - Larger Preview with exact slide proportions
    with right_col:
        st.markdown("**Preview**")

        preview_session = {
            "session_title": "Sample Session: The Future of AI in Healthcare",
            "stage": "Main Stage" if include_stage else "",
            "speakers": [
                {"name": "Dr. Sarah Chen", "job_title": "CMO", "company": "HealthTech", "is_moderator": True},
                {"name": "James Rodriguez", "job_title": "VP Innovation", "company": "MedGlobal", "is_moderator": False},
            ]
        }

        if csv_path and selected_event:
            try:
                sessions = load_sessions_for_event(csv_path, selected_event)
                if sessions:
                    preview_session = sessions[0]
                    if not include_stage:
                        preview_session["stage"] = ""
            except:
                pass

        # Get colors - panel text and stage are now independent
        panel_text_hex = panel_text_color if 'panel_text_color' in locals() else COLORS["title_text_dark"]
        stage_text_hex = stage_text_color if 'stage_text_color' in locals() else COLORS["stage_text_light"]

        preview_img = render_slide_preview_pil(
            900,
            preview_session,
            brand_logo_bytes,
            informa_logo_path,
            bg_image_bytes,
            bg_color,
            panel_color if 'panel_color' in locals() else None,
            panel_opacity if 'panel_opacity' in locals() else 100,
            panel_text_hex,  # Title text color (same as speaker text in this simplified version)
            panel_text_hex,  # Speaker text color
            stage_text_hex,  # Stage text color - INDEPENDENT
        )
        st.image(preview_img, use_container_width=True)


if __name__ == "__main__":
    main()
