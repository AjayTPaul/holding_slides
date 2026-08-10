"""Verify PPTX outputs match baselines."""

import hashlib
import sys

def file_checksum(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()

baselines = [
    "baseline_multi_speaker_transparent.pptx",
    "baseline_light_theme_nostage.pptx",
    "baseline_default_opaque.pptx"
]

print("Baseline checksums:")
for f in baselines:
    cs = file_checksum(f)
    print(f"  {f}: {cs[:16]}...")

print("\nSave these checksums. After cleanup, regenerate and compare.")
