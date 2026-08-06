# Informa Connect Session Slide Builder

A Streamlit web app that generates PowerPoint holding slides for conference sessions.
Built for Informa events, following brand guidelines.

## Live Demo

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)

## Features

- **Web-based interface** - No Python knowledge required
- **Live preview** - See changes in real-time before generating
- **Brand-compliant** - Follows Informa brand guidelines (colors, typography, spacing)
- **Customizable branding:**
  - Brand logo (top-left)
  - Informa logo (Light/Dark variants, bottom-right)
  - Background (default, solid color, or custom image)
  - Content panel color with opacity control
  - Text theme (Dark/Light)
  - Optional stage name display

## Installation (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Usage

1. **Select Data Source**
   - Use sample demo data, or
   - Upload your own CSV with session/speaker data

2. **Select Event**
   - Auto-detects event from CSV
   - Brand auto-detected from event data

3. **Customize Branding**
   - Upload brand logo
   - Choose Informa logo treatment
   - Set background
   - Configure content panel color and opacity
   - Choose text color theme
   - Toggle stage name visibility

4. **Preview & Generate**
   - Live preview updates as you edit
   - Click "Generate Slides" to create PPTX
   - Download slides

## CSV Data Schema

| Column | Description |
|--------|-------------|
| `event_name` | Event name (used for filtering) |
| `session_id` | Groups multiple speakers into one slide |
| `session_title` | Session title displayed on slide |
| `stage` | Stage/track name (e.g., "MAIN STAGE") |
| `speaker_name` | Speaker full name |
| `job_title` | Speaker job title |
| `company` | Speaker company |
| `is_moderator` | "yes"/"no" - affects ordering and "Moderator:" prefix |
| `brand` | Event brand (auto-detected) |

## Demo Data

`sample_sessions.csv` contains fake/demo data for testing:
- NetworkX London 2026 (Informa Tech brand)
- PharmaSummit 2026 (Informa Pharma brand)

All speaker names, companies, and events are fictional.

## Tech Stack

- **Streamlit** - Web UI
- **python-pptx** - PowerPoint generation
- **Pillow** - Image preview rendering
- **pandas** - CSV data handling

## Deployment

This app is deployed on [Streamlit Community Cloud](https://streamlit.io/cloud).

To deploy your own instance:
1. Fork this repository
2. Connect to Streamlit Cloud
3. Set `app.py` as the main file

## Project Structure

```
.
├── app.py                      # Main Streamlit application
├── requirements.txt            # Python dependencies
├── sample_sessions.csv         # Demo data (fictional)
├── BRAND_GUIDELINES.md        # Design reference
├── Informa Logo/              # Informa logo assets
│   ├── Informa_Logo_OneLine_Graduated_White_RGB.png
│   └── Informa_Logo_OneLine_Graduated_Indigo_RGB.png
└── README.md                  # This file
```

## License

Internal use - Informa AI Pilot Group
