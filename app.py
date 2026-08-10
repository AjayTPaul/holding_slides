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


def add_slide_with_template(prs, session, template_image_bytes, text_color):
    """Generate a slide using pre-designed template image.

    Dynamic text (session title, speaker list) is rendered directly onto
the template at fixed positions. No additional shapes or logos are added.
    """
    layout_idx = min(6, len(prs.slide_layouts) - 1)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    slide_width = Inches(SLIDE_WIDTH_IN)
    slide_height = Inches(SLIDE_HEIGHT_IN)

    # Template image as full slide background
    if template_image_bytes:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(template_image_bytes)
            tmp_path = tmp.name
        try:
            pic = slide.shapes.add_picture(tmp_path, 0, 0, width=slide_width, height=slide_height)
            sp_tree = slide.shapes._spTree
            sp_tree.remove(pic._element)
            sp_tree.insert(2, pic._element)
        except Exception as e:
            st.warning(f"Could not add template image: {e}")
        finally:
            os.unlink(tmp_path)

    speakers = sorted(session["speakers"], key=lambda s: (not s["is_moderator"], s["name"]))
    title_size, speaker_size = fit_content(session)

    # Session title - positioned on template
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
    p.font.color.rgb = hex_to_rgb(text_color)
    p.alignment = PP_ALIGN.LEFT

    # Speakers with proper bolding
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
                r.font.color.rgb = hex_to_rgb(text_color)
            # Name (bold)
            r = sp.add_run()
            r.text = speaker["name"]
            r.font.size = Pt(speaker_size)
            r.font.name = SPEAKER_FONT
            r.font.bold = True
            r.font.color.rgb = hex_to_rgb(text_color)
            # Job title (not bold)
            if speaker["job_title"]:
                r = sp.add_run()
                r.text = f", {speaker['job_title']}"
                r.font.size = Pt(speaker_size)
                r.font.name = SPEAKER_FONT
                r.font.color.rgb = hex_to_rgb(text_color)
            # Company (bold)
            if speaker["company"]:
                r = sp.add_run()
                r.text = f", {speaker['company']}"
                r.font.size = Pt(speaker_size)
                r.font.name = SPEAKER_FONT
                r.font.bold = True
                r.font.color.rgb = hex_to_rgb(text_color)

    return slide


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
        page_icon="🎨",
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
        <p style='color: #6B7280; font-size: 1rem; margin-bottom: 1.5rem;'>
            Generate professional holding slides for your event sessions
        </p>
        """,
        unsafe_allow_html=True
    )

    # Unequal columns - controls thinner, preview larger
    left_col, right_col = st.columns([1, 2])

    # Initialize session state
    if "csv_path" not in st.session_state:
        st.session_state.csv_path = "sample_sessions.csv"

    # Initialize control values in session state
    if "selected_event" not in st.session_state:
        st.session_state.selected_event = None
    if "selected_stream" not in st.session_state:
        st.session_state.selected_stream = None

    # Left column - Compact Controls (simplified to 4 controls)
    with left_col:
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

        # Combine Template and Text Color into one container with two columns
        if selected_event and selected_stream:
            with st.container(border=True):
                template_color_cols = st.columns([3, 2])

                with template_color_cols[0]:
                    st.markdown("**Template**")
                    template_file = st.file_uploader(
                        "",
                        type=["png", "jpg", "jpeg"],
                        help="Upload a template with all fixed design elements baked in",
                        label_visibility="collapsed"
                    )
                    template_bytes = template_file.getvalue() if template_file else None
                    if template_bytes:
                        st.image(template_bytes, use_container_width=True)

                with template_color_cols[1]:
                    st.markdown("**Text Color**")
                    text_color = st.color_picker(
                        "",
                        value="#" + COLORS["title_text_dark"],
                        help="Applies to session title and speaker list",
                        label_visibility="collapsed"
                    )

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
                        output_bytes, num_slides = generate_slides(
                            csv_path, selected_event, selected_stream,
                            template_image_bytes=template_bytes,
                            text_color=text_color.lstrip("#"),
                        )

                        # Auto-trigger download via hidden link
                        filename = f"{selected_event.replace(' ', '_')}_{selected_stream.replace(' ', '_')}_holding_slides.pptx"
                        b64 = base64.b64encode(output_bytes.getvalue()).decode()
                        st_html(f"""
                            <a id="auto-dl" href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}"></a>
                            <script>document.getElementById('auto-dl').click();</script>
                        """, height=0)
                    except Exception as e:
                        st.error(f"Error: {e}")

    # Right column - Preview with styled container
    with right_col:
        st.markdown("**Preview**")

        # Check if template is uploaded for empty state
        has_template = template_bytes is not None

        # Default preview session
        preview_session = {
            "session_title": "Sample Session: The Future of AI in Healthcare",
            "speakers": [
                {"name": "Dr. Sarah Chen", "job_title": "CMO", "company": "HealthTech", "is_moderator": True, "headshot_url": None},
                {"name": "James Rodriguez", "job_title": "VP Innovation", "company": "MedGlobal", "is_moderator": False, "headshot_url": None},
            ]
        }

        # Load actual session data if available
        if selected_event and selected_stream:
            try:
                sessions = load_sessions_for_event(csv_path, selected_event, selected_stream)
                if sessions:
                    preview_session = sessions[0]
            except:
                pass

        # Render preview or show empty state
        if has_template:
            preview_img = render_slide_preview_pil(
                900,
                preview_session,
                template_bytes,
                text_color.lstrip("#") if 'text_color' in locals() else COLORS["title_text_dark"],
            )
            st.image(preview_img, use_container_width=True)
        else:
            # Empty state with subtle styling
            st.markdown(
                """
                <div style='
                    background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
                    border: 2px dashed #cbd5e1;
                    border-radius: 12px;
                    padding: 4rem 2rem;
                    text-align: center;
                    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
                '>
                    <div style='font-size: 3rem; margin-bottom: 1rem;'>🖼️</div>
                    <div style='color: #64748b; font-size: 1.1rem; font-weight: 500;'>
                        Upload a template to see your preview
                    </div>
                    <div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>
                        Your slide design will appear here
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )


if __name__ == "__main__":
    main()
