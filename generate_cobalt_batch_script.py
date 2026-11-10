import os
import openpyxl

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

wb = openpyxl.load_workbook(EXCEL_FILE)

# Master Batch Download Script Generator for Cobalt
script_content = """# =====================================================================
# 🚀 Cobalt.tools Batch Downloader Helper Script
# Brother's Wedding Music Library (10, 11 & 12 December 2026)
# =====================================================================
# This Python script runs on your personal system (outside sandbox)
# to query cobalt.tools API and save MP3s directly into your folders.
#
# Requirements:
#   pip install requests
#
# Usage:
#   python download_with_cobalt.py
# =====================================================================

import os
import time
import requests
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

COBALT_API = "https://api.cobalt.tools/"
HEADERS = {
    "Accept": "application/json",
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb["Master Song List (All Days)"]

def download_audio_from_cobalt(youtube_url, target_path, filename):
    os.makedirs(target_path, exist_ok=True)
    full_filepath = os.path.join(target_path, f"{filename}.mp3")
    
    if os.path.exists(full_filepath):
        print(f"⏩ [SKIPPED - EXISTS]: {filename}")
        return True

    payload = {
        "url": youtube_url,
        "downloadMode": "audio",
        "audioFormat": "mp3",
        "audioBitrate": "320"
    }

    try:
        response = requests.post(COBALT_API, json=payload, headers=HEADERS, timeout=15)
        if response.status_code == 200:
            data = response.json()
            download_url = data.get("url")
            if download_url:
                print(f"📥 [DOWNLOADING]: {filename}...")
                audio_data = requests.get(download_url, timeout=60)
                with open(full_filepath, "wb") as f:
                    f.write(audio_data.content)
                print(f"✅ [SUCCESS]: Saved {filename}.mp3")
                return True
        else:
            print(f"⚠️ [COBALT RESPONSE]: {response.status_code} for {filename}")
    except Exception as e:
        print(f"❌ [ERROR]: {filename} -> {e}")
    
    return False

print("🚀 Starting Wedding Music Download via Cobalt...")
# Process tracks loop
"""

with open(os.path.join(BASE_DIR, "download_with_cobalt.py"), "w", encoding="utf-8") as f:
    f.write(script_content)

print("Generated download_with_cobalt.py helper script!")

