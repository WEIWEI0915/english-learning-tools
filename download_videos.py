#!/usr/bin/env python3
"""Download English Singsing videos from Bilibili and merge DASH streams"""

import subprocess
import os
import glob
import shutil
from pathlib import Path
import time

VIDEO_DIR = Path(__file__).parent / "videos"
VIDEO_DIR.mkdir(exist_ok=True)

# Get ffmpeg path
import imageio_ffmpeg
FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()

# Videos to download: (url, output_name, description)
DOWNLOADS = [
    # 28-theme vocabulary collection (BV1f3411B7cx)
    # These are all Kids Vocabulary videos
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=1", "family_vocab", "Family members"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=19", "body_vocab", "Body parts"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=13", "colors_vocab", "Colors"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=22", "colors2_vocab", "Colors 2"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=12", "numbers_vocab", "Numbers"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=11", "shapes_vocab", "Shapes"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=14", "animals_vocab", "Sea animals"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=21", "food_vocab", "Food"),
    ("https://www.bilibili.com/video/BV1f3411B7cx/?p=17", "weather_vocab", "Weather"),

    # Dialogue videos (standalone)
    ("https://www.bilibili.com/video/BV1L4411h77Q/", "greetings_dialogue", "Greetings dialogue"),

    # Weather theme song + dialogue
    ("https://www.bilibili.com/video/BV1V64y1F76f/", "weather_theme", "Weather theme"),

    # Daily Routine - try standalone BV
    ("https://www.bilibili.com/video/BV1yB4WzZEwt/", "daily_routine", "Daily Routine"),
]

def download_and_merge(url, name, desc):
    """Download video with you-get and merge DASH streams"""
    print(f"\n{'='*60}")
    print(f"Downloading: {desc} -> {name}.mp4")
    print(f"URL: {url}")

    # Check if already merged
    merged = VIDEO_DIR / f"{name}.mp4"
    if merged.exists():
        print(f"  Already exists: {merged} ({merged.stat().st_size / 1024 / 1024:.1f} MB)")
        return True

    # Download with you-get
    os.chdir(VIDEO_DIR)
    result = subprocess.run(
        ["you-get", "--format=dash-flv480-AVC", "-o", str(VIDEO_DIR), url],
        capture_output=True, timeout=300, env={**os.environ, 'PYTHONIOENCODING': 'utf-8'}
    )

    if result.returncode != 0:
        # Try without specific format
        result = subprocess.run(
            ["you-get", "-o", str(VIDEO_DIR), url],
            capture_output=True, timeout=300, env={**os.environ, 'PYTHONIOENCODING': 'utf-8'}
        )
        if result.returncode != 0:
            print(f"  ERROR downloading: {result.stderr[:200] if isinstance(result.stderr, str) else result.stderr[:200].decode('utf-8', errors='replace')}")
            return False

    # Find downloaded files - look for [00].mp4 (video) and [01].mp4 (audio)
    time.sleep(1)  # Wait for file system

    # Find the most recent [00] and [01] files
    video_files = sorted(VIDEO_DIR.glob("*[00].mp4"), key=os.path.getmtime, reverse=True)
    audio_files = sorted(VIDEO_DIR.glob("*[01].mp4"), key=os.path.getmtime, reverse=True)

    if not video_files or not audio_files:
        print(f"  Could not find downloaded DASH parts")
        return False

    video_file = video_files[0]
    # Find matching audio - same prefix
    prefix = str(video_file).replace("[00].mp4", "")
    audio_file = Path(prefix + "[01].mp4")

    if not audio_file.exists():
        print(f"  Audio file not found: {audio_file}")
        return False

    print(f"  Video: {video_file.name}")
    print(f"  Audio: {audio_file.name}")

    # Merge with ffmpeg
    merge_cmd = [
        FFMPEG, "-i", str(video_file), "-i", str(audio_file),
        "-c", "copy", "-map", "0:v:0", "-map", "1:a:0",
        "-shortest", str(merged), "-y"
    ]

    merge_result = subprocess.run(merge_cmd, capture_output=True, timeout=120, env={**os.environ, 'PYTHONIOENCODING': 'utf-8'})
    if merge_result.returncode != 0:
        print(f"  ERROR merging: {merge_result.stderr[-200:]}")
        return False

    # Clean up DASH parts
    try:
        video_file.unlink()
        audio_file.unlink()
        # Also clean up XML
        for xml_file in VIDEO_DIR.glob(f"{Path(prefix).name}*.cmt.xml"):
            xml_file.unlink()
    except Exception as e:
        print(f"  Warning cleaning up: {e}")

    print(f"  SUCCESS: {merged.name} ({merged.stat().st_size / 1024 / 1024:.1f} MB)")
    return True

def main():
    print("English Singsing Video Downloader")
    print(f"FFmpeg: {FFMPEG}")
    print(f"Output: {VIDEO_DIR}")

    success = 0
    failed = 0

    for url, name, desc in DOWNLOADS:
        try:
            if download_and_merge(url, name, desc):
                success += 1
            else:
                failed += 1
        except Exception as e:
            print(f"  EXCEPTION: {e}")
            failed += 1

    print(f"\n{'='*60}")
    print(f"Done! Success: {success}, Failed: {failed}")

    # List all merged files
    print("\nMerged videos:")
    for f in sorted(VIDEO_DIR.glob("*_*.mp4")):
        print(f"  {f.name} ({f.stat().st_size / 1024 / 1024:.1f} MB)")

if __name__ == '__main__':
    main()
