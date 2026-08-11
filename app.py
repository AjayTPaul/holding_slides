"""
Informa Connect Session Slide Builder

Generate professional holding slides for events.
Usage: streamlit run app.py
"""

import base64
import io
import os
import subprocess
import tempfile
from copy import deepcopy
from pathlib import Path

import streamlit as st
from streamlit.components.v1 import html as st_html
import streamlit.components.v1 as components
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


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
    "default_bg": "F5F5F5",
    "cta_bg": "002244",  # Indigo
}

CONTENT_PANEL = {
    "left": 1.5,
    "top": 2.5,
    "width": 10.333,
    "height": 3.5,
}

TITLE_FONT = "Open Sans"
SPEAKER_FONT = "Open Sans"
TITLE_MAX_PT = 36
TITLE_MIN_PT = 24
SPEAKER_MAX_PT = 21
SPEAKER_MIN_PT = 16


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


def load_sessions_for_event(csv_path, event_name, stream=None):
    df = pd.read_csv(csv_path)
    df = df[df["event_name"] == event_name]
    if stream:
        df = df[df["stage"] == stream]
    sessions = {}
    for _, row in df.iterrows():
        session_id = str(row.get("session_id", "")).strip()
        if not session_id:
            continue
        if session_id not in sessions:
            sessions[session_id] = {
                "session_id": session_id,
                "session_title": str(row.get("session_title", "")).strip(),
                "speakers": [],
            }
        sessions[session_id]["speakers"].append({
            "name": str(row.get("speaker_name", "")).strip(),
            "job_title": str(row.get("job_title", "")).strip(),
            "company": str(row.get("company", "")).strip(),
            "is_moderator": parse_is_moderator(row.get("is_moderator", "")),
        })
    return list(sessions.values())


def load_streams_for_event(csv_path, event_name):
    df = pd.read_csv(csv_path)
    df = df[df["event_name"] == event_name]
    streams = df["stage"].dropna().unique().tolist()
    return sorted([s.strip() for s in streams if s.strip()])


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


def substitute_tokens_in_paragraph(p, session, speaker=None, scale=1.0, context=""):
    """Substitute tokens in a paragraph, handling PowerPoint's run splitting.

    PowerPoint may split tokens across runs (e.g., "{session" in run0, "_title}" in run1).
    This function merges runs, substitutes, then rebuilds runs to preserve formatting.
    """
    from copy import deepcopy

    # Debug: print raw run text with context
    print(f"DEBUG paragraph runs {context}:")
    for i, run in enumerate(p.runs):
        print(f"  run[{i}] = {repr(run.text)}")

    # Get full paragraph text
    full_text = get_paragraph_text(p)
    print(f"  full_text = {repr(full_text)}")

    # Check if any tokens present
    tokens = ["{session_title}", "{speaker_name}", "{job_title}", "{company}"]
    if not any(t in full_text for t in tokens):
        print(f"  No tokens found, skipping")
        return

    # Prepare replacement values
    session_title = session.get("session_title", "")
    speaker_name = speaker.get("name", "") if speaker else "{NAME}"
    job_title = speaker.get("job_title", "") if speaker else "{JOB}"
    company = speaker.get("company", "") if speaker else "{COMPANY}"

    # Perform substitutions on the full text
    new_text = full_text
    new_text = new_text.replace("{session_title}", session_title)
    new_text = new_text.replace("{speaker_name}", speaker_name)
    new_text = new_text.replace("{job_title}", job_title)
    new_text = new_text.replace("{company}", company)

    print(f"  after substitution = {repr(new_text)}")

    # Rebuild runs: keep first run's formatting, put all text there, clear others
    if p.runs:
        first_run = p.runs[0]
        first_run.text = new_text

        # Clear remaining runs
        for run in p.runs[1:]:
            run.text = ""

        # Apply font scaling
        if scale < 1.0 and first_run.font.size:
            from pptx.util import Pt
            first_run.font.size = Pt(int(first_run.font.size.pt * scale))


def substitute_tokens_in_shape(shape, session):
    """Substitute tokens in a shape's text frame. Returns True if this is a speaker template."""
    if not shape.has_text_frame:
        return False

    tf = shape.text_frame
    has_speaker_template = False

    for p in tf.paragraphs:
        full_text = get_paragraph_text(p)
        if "{speaker_name}" in full_text:
            has_speaker_template = True

    return has_speaker_template


def get_paragraph_text(p):
    """Get paragraph text by concatenating all runs.

    Handles PowerPoint splitting tokens across multiple runs.
    Same logic used by both generation and preview scanning.
    """
    return "".join(run.text for run in p.runs)


def scan_template_for_tokens(template_bytes):
    """Scan template PPTX to find which tokens are actually present.

    Returns a dict with token names as keys and True/False for presence.
    Tokens: session_title, speaker_name, job_title, company
    """
    import re
    from pptx import Presentation
    import tempfile
    import os

    TOKEN_PATTERNS = {
        'session_title': r'\{session_title\}',
        'speaker_name': r'\{speaker_name\}',
        'job_title': r'\{job_title\}',
        'company': r'\{company\}',
    }

    found_tokens = {
        'session_title': False,
        'speaker_name': False,
        'job_title': False,
        'company': False,
    }

    # Write to temp file and load
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        data = template_bytes.getvalue() if hasattr(template_bytes, 'getvalue') else template_bytes
        tmp.write(data)
        tmp_path = tmp.name

    try:
        prs = Presentation(tmp_path)
        # Check all slides for tokens
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for p in shape.text_frame.paragraphs:
                    para_text = get_paragraph_text(p)
                    for token_name, pattern in TOKEN_PATTERNS.items():
                        if re.search(pattern, para_text):
                            found_tokens[token_name] = True
    finally:
        os.unlink(tmp_path)

    return found_tokens


def extract_slide_confirmation(pptx_bytes, template_bytes=None):
    """Extract text confirmation from generated PPTX showing substituted values with per-token status.

    Only checks and displays tokens that exist in the original template.

    Args:
        pptx_bytes: The generated output PPTX
        template_bytes: The original template PPTX (optional, for token scanning)

    Returns a list of dicts with 'slide_num', 'session_title', 'session_title_resolved',
    'session_title_in_template', 'speakers'.
    Each speaker has per-token resolution status and template presence flags.
    """
    import re
    from pptx import Presentation
    import tempfile
    import os

    TOKEN_PATTERNS = {
        'session_title': r'\{session_title\}',
        'speaker_name': r'\{speaker_name\}',
        'job_title': r'\{job_title\}',
        'company': r'\{company\}',
    }

    # First, scan template to find which tokens exist (if template_bytes provided)
    tokens_in_template = {
        'session_title': template_bytes is None,  # If no template provided, assume all exist
        'speaker_name': template_bytes is None,
        'job_title': template_bytes is None,
        'company': template_bytes is None,
    }
    if template_bytes is not None:
        tokens_in_template = scan_template_for_tokens(template_bytes)
        print(f"DEBUG scan result: {tokens_in_template}")

    def check_token_in_text(text, token_pattern):
        """Check if token pattern still exists in text (unresolved)."""
        return bool(re.search(token_pattern, text))

    def is_substantial_text(text):
        """Check if text is substantial (not placeholder like '.')."""
        return len(text) > 2 and any(c.isalnum() for c in text)

    slides = []

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        data = pptx_bytes.getvalue() if hasattr(pptx_bytes, 'getvalue') else pptx_bytes
        tmp.write(data)
        tmp_path = tmp.name

    try:
        prs = Presentation(tmp_path)
        for slide_idx, slide in enumerate(prs.slides, 1):
            session_title_value = None
            session_title_resolved = False
            session_title_in_template = tokens_in_template.get('session_title', False)
            speakers = []
            current_speaker = None

            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                tf = shape.text_frame

                for p in tf.paragraphs:
                    para_text = get_paragraph_text(p).strip()

                    if not is_substantial_text(para_text):
                        continue

                    # Check for unresolved tokens (placeholder still present)
                    has_session_title = check_token_in_text(para_text, TOKEN_PATTERNS['session_title'])
                    has_speaker_name = check_token_in_text(para_text, TOKEN_PATTERNS['speaker_name'])
                    has_job_title = check_token_in_text(para_text, TOKEN_PATTERNS['job_title'])
                    has_company = check_token_in_text(para_text, TOKEN_PATTERNS['company'])

                    # Classify paragraph based on what tokens are present
                    if has_speaker_name:
                        # This is a speaker line
                        parts = para_text.replace('|', ',').split(',')
                        name_part = parts[0].strip() if len(parts) > 0 else ""
                        job_part = parts[1].strip() if len(parts) > 1 else ""
                        company_part = parts[2].strip() if len(parts) > 2 else ""

                        speakers.append({
                            "name": name_part,
                            "name_resolved": not has_speaker_name,
                            "name_in_template": tokens_in_template.get('speaker_name', False),
                            "job_title": job_part,
                            "job_title_resolved": not has_job_title,
                            "job_title_in_template": tokens_in_template.get('job_title', False),
                            "company": company_part,
                            "company_resolved": not has_company,
                            "company_in_template": tokens_in_template.get('company', False),
                        })
                    elif has_session_title and session_title_value is None:
                        # This is the session title
                        session_title_value = para_text
                        session_title_resolved = not has_session_title

            slides.append({
                "slide_num": slide_idx,
                "session_title": session_title_value or "(no title found)",
                "session_title_resolved": session_title_resolved,
                "session_title_in_template": session_title_in_template,
                "speakers": speakers
            })
    finally:
        os.unlink(tmp_path)

    return slides



def generate_slides_from_template(template_bytes, sessions):
    """Generate slides using a PPTX template with token replacement.

    Tokens: {session_title}, {speaker_name}, {job_title}, {company}

    Classifies paragraphs by token content, not by shape identity:
    - Paragraphs with {session_title} get single substitution
    - Paragraphs with {speaker_name} are the repeatable template paragraph
    - Paragraphs with no tokens are left untouched
    """
    from copy import deepcopy
    output = io.BytesIO()

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_template:
        tmp_template.write(template_bytes)
        tmp_template_path = tmp_template.name

    try:
        # Load template presentation
        template_prs = Presentation(tmp_template_path)
        template_slide = template_prs.slides[0]

        for session in sessions:
            speakers = session["speakers"]
            speaker_count = len(speakers)

            # Process each shape in the template slide
            for shape in template_slide.shapes:
                if not shape.has_text_frame:
                    continue

                tf = shape.text_frame

                # Collect paragraphs by type based on tokens actually present
                speaker_template_paras = []  # Has {speaker_name}
                title_paras = []             # Has {session_title} but not {speaker_name}
                other_paras = []             # No tokens

                for p in tf.paragraphs:
                    full_text = get_paragraph_text(p)
                    has_speaker = "{speaker_name}" in full_text
                    has_title = "{session_title}" in full_text

                    if has_speaker:
                        speaker_template_paras.append(p)
                    elif has_title:
                        title_paras.append(p)
                    else:
                        other_paras.append(p)

                # Process speaker paragraphs (clone per speaker if needed)
                if speaker_template_paras and speaker_count > 0:
                    template_para = speaker_template_paras[0]

                    # Calculate scale based on total speakers
                    total_lines = speaker_count
                    available_height = shape.height
                    estimated_line_height = 360000  # ~0.3 inches in EMUs
                    available_lines = max(3, available_height / estimated_line_height)
                    scale = max(0.6, available_lines / total_lines) if total_lines > available_lines else 1.0
                    print(f"DEBUG: scale={scale:.2f} for {speaker_count} speakers in shape")

                    # First speaker: substitute in-place
                    substitute_tokens_in_paragraph(template_para, session, speakers[0], scale, context=f"SPEAKER-1/{speaker_count}")

                    # Additional speakers: clone the template paragraph
                    for idx, speaker in enumerate(speakers[1:], start=2):
                        new_p_el = deepcopy(template_para._element)
                        tf._element.append(new_p_el)
                        new_p = tf.paragraphs[-1]
                        substitute_tokens_in_paragraph(new_p, session, speaker, scale, context=f"SPEAKER-{idx}/{speaker_count}")

                # Process title paragraphs (single substitution, no cloning)
                for p in title_paras:
                    substitute_tokens_in_paragraph(p, session, None, 1.0, context="TITLE")

                # Other paragraphs: leave untouched

        # Save modified presentation
        template_prs.save(output)
        output.seek(0)

    finally:
        os.unlink(tmp_template_path)

    return output, len(sessions)



def generate_slides(csv_path, event_name, stream, template_image_bytes, text_color):
    sessions = load_sessions_for_event(csv_path, event_name, stream)
    prs = Presentation()
    prs.slide_width = Emu(int(Inches(SLIDE_WIDTH_IN)))
    prs.slide_height = Emu(int(Inches(SLIDE_HEIGHT_IN)))
    for session in sessions:
        add_slide_with_template(
            prs, session,
            template_image_bytes=template_image_bytes,
            text_color=text_color,
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


def render_slide_preview_pil(target_width_px, session, template_image_bytes, text_color):
    """Render a slide preview using PIL Image with pre-designed template.

    Just overlays dynamic text (session title, speaker list) onto the template image.
    """
    target_height_px = int(target_width_px / SLIDE_ASPECT)

    # Template as background
    if template_image_bytes:
        try:
            bg = Image.open(io.BytesIO(template_image_bytes))
            bg = bg.resize((target_width_px, target_height_px), Image.LANCZOS)
        except:
            bg_color = hex_to_rgb_tuple(COLORS["default_bg"])
            bg = Image.new("RGB", (target_width_px, target_height_px), bg_color)
    else:
        bg_color = hex_to_rgb_tuple(COLORS["default_bg"])
        bg = Image.new("RGB", (target_width_px, target_height_px), bg_color)

    if bg.mode != 'RGBA':
        bg = bg.convert('RGBA')

    draw = ImageDraw.Draw(bg, 'RGBA')

    # Load bundled fonts
    script_dir = Path(__file__).parent
    fonts_dir = script_dir / "fonts"
    try:
        title_font = ImageFont.truetype(str(fonts_dir / "OpenSans-Bold.ttf"), int(target_width_px * 0.04))
        speaker_font_regular = ImageFont.truetype(str(fonts_dir / "OpenSans-Regular.ttf"), int(target_width_px * 0.023))
        speaker_font_bold = ImageFont.truetype(str(fonts_dir / "OpenSans-Bold.ttf"), int(target_width_px * 0.023))
    except Exception as e:
        st.warning(f"Bundled fonts not found at {fonts_dir}, using default font: {e}")
        title_font = ImageFont.load_default()
        speaker_font_regular = ImageFont.load_default()
        speaker_font_bold = ImageFont.load_default()

    # Content panel positions
    panel_left = int((CONTENT_PANEL["left"] / SLIDE_WIDTH_IN) * target_width_px)
    panel_top = int((CONTENT_PANEL["top"] / SLIDE_HEIGHT_IN) * target_height_px)
    panel_right = int(((CONTENT_PANEL["left"] + CONTENT_PANEL["width"]) / SLIDE_WIDTH_IN) * target_width_px)

    # Session title (bold)
    title_x = panel_left + int(target_width_px * 0.025)
    title_y = panel_top + int(target_height_px * 0.08)
    title = session.get("session_title", "")
    text_color_tuple = hex_to_rgb_tuple(text_color or COLORS["title_text_dark"])

    # Simple text wrapping for title
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
        draw.text((title_x, title_y + i * line_height), line, font=title_font, fill=text_color_tuple)

    # Speakers with proper bolding
    speakers = sorted(session.get("speakers", []), key=lambda s: (not s["is_moderator"], s["name"]))
    speaker_y = panel_top + int(target_height_px * 0.28)
    line_spacing = int(target_height_px * 0.055)
    space_width = int(target_width_px * 0.008)

    for speaker in speakers:
        current_x = title_x

        # Moderator label (not bold)
        if speaker["is_moderator"]:
            mod_text = "Moderator: "
            draw.text((current_x, speaker_y), mod_text, font=speaker_font_regular, fill=text_color_tuple)
            bbox = draw.textbbox((current_x, speaker_y), mod_text, font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0]) + space_width

        # Name (bold)
        name = speaker["name"]
        draw.text((current_x, speaker_y), name, font=speaker_font_bold, fill=text_color_tuple)
        bbox = draw.textbbox((current_x, speaker_y), name, font=speaker_font_bold)
        current_x += (bbox[2] - bbox[0])

        # Job title (not bold)
        if speaker["job_title"]:
            comma = ", "
            draw.text((current_x, speaker_y), comma, font=speaker_font_regular, fill=text_color_tuple)
            bbox = draw.textbbox((current_x, speaker_y), comma, font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0])
            draw.text((current_x, speaker_y), speaker["job_title"], font=speaker_font_regular, fill=text_color_tuple)
            bbox = draw.textbbox((current_x, speaker_y), speaker["job_title"], font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0])

        # Company (bold)
        if speaker["company"]:
            comma = ", "
            draw.text((current_x, speaker_y), comma, font=speaker_font_regular, fill=text_color_tuple)
            bbox = draw.textbbox((current_x, speaker_y), comma, font=speaker_font_regular)
            current_x += (bbox[2] - bbox[0])
            draw.text((current_x, speaker_y), speaker["company"], font=speaker_font_bold, fill=text_color_tuple)

        speaker_y += line_spacing

    return bg


# -----------------------------------------------------------------------------
# STREAMLIT INTERFACE
# -----------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="Holding Slide Generator | Informa",
        page_icon="Informa_Orbit_RGB.png",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Informa logo and title header (stacked vertically)
    informa_logo_path = Path(__file__).parent / "Informa Logo" / "Informa_Logo_OneLine_Graduated_Indigo_RGB.png"
    if informa_logo_path.exists():
        st.image(str(informa_logo_path), width=120)

    # Title with brand color and subtitle
    st.markdown(
        f"""
        <h1 style='font-size: 2.5rem; margin-bottom: 0.25rem; margin-top: 0.5rem; color: #{COLORS["cta_bg"]};'>
            Holding Slide Generator
        </h1>
        <p style='color: #6B7280; font-size: 1rem; margin-bottom: 0.5rem;'>
            Generate professional holding slides for your event sessions
        </p>
        <p style='color: #9CA3AF; font-size: 0.85rem; margin-bottom: 1.5rem;'>
            Upload a PPTX template with placeholder tokens: {{session_title}}, {{speaker_name}}, {{job_title}}, {{company}}.<br>
            Font styles, weights, spacing, and colors from your template will carry over to generated slides.
        </p>
        """,
        unsafe_allow_html=True
    )

    # Initialize session state
    if "csv_path" not in st.session_state:
        st.session_state.csv_path = "sample_sessions.csv"

    # Initialize control values in session state
    if "selected_event" not in st.session_state:
        st.session_state.selected_event = None
    if "selected_stream" not in st.session_state:
        st.session_state.selected_stream = None

# Controls section
    csv_path = "sample_sessions.csv"
    selected_event = None
    selected_stream = None
    template_bytes = None
    text_color = "#" + COLORS["title_text_dark"]

    # Combine Event and Stream into one container with two columns
    with st.container(border=True):
            event_stream_cols = st.columns(2)

            with event_stream_cols[0]:
                st.markdown("**Event**")
                try:
                    events = load_events(csv_path)
                    if events:
                        event_names = [e["event_name"] for e in events]
                        selected_event = st.selectbox(
                            "",
                            event_names,
                            key="event_select",
                            label_visibility="collapsed"
                        )
                        st.session_state.selected_event = selected_event
                    else:
                        st.warning("No events found")
                except Exception as e:
                    st.error(f"Could not load events: {e}")

            with event_stream_cols[1]:
                st.markdown("**Stream**")
                if selected_event:
                    try:
                        streams = load_streams_for_event(csv_path, selected_event)
                        if streams:
                            selected_stream = st.selectbox(
                                "",
                                streams,
                                key="stream_select",
                                label_visibility="collapsed"
                            )
                            st.session_state.selected_stream = selected_stream
                        else:
                            st.warning("No streams found")
                    except Exception as e:
                        st.error(f"Could not load streams: {e}")
                else:
                    st.selectbox("", ["Select event first"], disabled=True, label_visibility="collapsed")

    # Template upload - PPTX with token replacement
    if selected_event and selected_stream:
            with st.container(border=True):
                st.markdown("**Template**")
                template_file = st.file_uploader(
                    "",
                    type=["pptx"],
                    help="Upload a PPTX template with {session_title}, {speaker_name}, {job_title}, {company} tokens",
                    label_visibility="collapsed"
                )
                template_bytes = template_file.getvalue() if template_file else None

    # Generate button with Indigo CTA color and Ultramarine hover
    if selected_event and selected_stream:
            st.markdown("""
            <style>
                .stButton>button {
                    background-color: #002244 !important;
                    border-color: #002244 !important;
                }
                .stButton>button:hover {
                    background-color: #003CB2 !important;
                    border-color: #003CB2 !important;
                }
                .stButton>button:focus {
                    background-color: #003CB2 !important;
                    border-color: #003CB2 !important;
                }
                .stButton>button:active {
                    background-color: #003CB2 !important;
                    border-color: #003CB2 !important;
                }
            </style>
            """, unsafe_allow_html=True)
            generate_clicked = st.button(
                "Generate Slides",
                type="primary",
                use_container_width=True,
            )

            if generate_clicked:
                with st.spinner("Generating slides..."):
                    try:
                        sessions = load_sessions_for_event(csv_path, selected_event, selected_stream)
                        output_bytes, num_slides = generate_slides_from_template(
                            template_bytes, sessions
                        )

                        # Auto-trigger download via hidden link
                        filename = f"{selected_event.replace(' ', '_')}_{selected_stream.replace(' ', '_')}_holding_slides.pptx"
                        b64 = base64.b64encode(output_bytes.getvalue()).decode()
                        st_html(f"""
                            <a id="auto-dl" href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}"></a>
                            <script>document.getElementById('auto-dl').click();</script>
                        """, height=0)

                        # Generation complete - download triggered above
                    except Exception as e:
                        st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
