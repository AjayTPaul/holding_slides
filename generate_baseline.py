"""Generate baseline PPTX files for comparison before/after cleanup."""

import sys
sys.path.insert(0, '.')

from app import generate_slides

# Test configuration: multi-speaker, transparency, custom colors
output1, count1 = generate_slides(
    "sample_sessions.csv",
    "NetworkX London 2026",  # Multi-speaker session
    brand_logo_bytes=None,
    informa_logo_path=None,
    background_image_bytes=None,
    background_color=None,
    panel_color="#E3F4FF",
    panel_opacity=60,  # 60% transparency
    panel_text_color="#283857",
    stage_text_color="#FFFFFF",  # Different from panel text
    include_stage=True
)

with open("baseline_multi_speaker_transparent.pptx", "wb") as f:
    f.write(output1.getvalue())

print(f"Baseline 1: Multi-speaker with transparency - {count1} slides")

# Test configuration 2: Light text theme, no stage
output2, count2 = generate_slides(
    "sample_sessions.csv",
    "PharmaSummit 2026",
    brand_logo_bytes=None,
    informa_logo_path=None,
    background_image_bytes=None,
    background_color="#1a1a2e",  # Dark background
    panel_color="#4a4a6a",
    panel_opacity=80,
    panel_text_color="#FFFFFF",  # Light text
    stage_text_color="#FFD700",  # Gold stage text
    include_stage=False  # No stage
)

with open("baseline_light_theme_nostage.pptx", "wb") as f:
    f.write(output2.getvalue())

print(f"Baseline 2: Light theme, no stage - {count2} slides")

# Test configuration 3: Default settings
output3, count3 = generate_slides(
    "sample_sessions.csv",
    "NetworkX London 2026",
    brand_logo_bytes=None,
    informa_logo_path=None,
    background_image_bytes=None,
    background_color=None,
    panel_color=None,  # Default
    panel_opacity=100,  # Opaque
    panel_text_color=None,  # Default dark
    stage_text_color=None,  # Default
    include_stage=True
)

with open("baseline_default_opaque.pptx", "wb") as f:
    f.write(output3.getvalue())

print(f"Baseline 3: Default opaque settings - {count3} slides")

print("\nAll baseline files generated:")
print("- baseline_multi_speaker_transparent.pptx")
print("- baseline_light_theme_nostage.pptx")
print("- baseline_default_opaque.pptx")
