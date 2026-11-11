# -*- coding: utf-8 -*-
"""
Wedding DJ Music Player - Mobile & Desktop Generator
Builds a unified, fully responsive, touch-friendly Wedding DJ web app.
Supports:
- Local offline playback with folder connect (Blob URLs)
- Seamless cloud YouTube audio streaming fallback for GitHub Pages
- Full mobile-first touch UI with off-canvas drawer, ceremony carousel, song cards, mini-player, and full-screen DJ sheet
- Full desktop 3-pane DJ studio console with visualizer, cue table, and keyboard shortcuts
"""

import os
import json
import re
import urllib.parse
import openpyxl

REPO_ROOT = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S"
MUSIC_DIR = os.path.join(REPO_ROOT, "Brother Wedding Music")
EXCEL_FILE = os.path.join(MUSIC_DIR, "Brother_Wedding_Music_Collection.xlsx")

print("1. Parsing Master Excel for Ceremony Metadata and YouTube Links...")
wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
meta = {}
for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [str(c.value or '').lower() for c in ws[1]]
    song_col, movie_col, cat_col, year_col, cero_col, day_col, singer_col, yt_col = None, None, None, None, None, None, None, None
    for idx, h in enumerate(headers, 1):
        if 'song' in h: song_col = idx
        elif 'singer' in h or 'artist' in h: singer_col = idx
        elif 'movie' in h or 'album' in h or 'source' in h: movie_col = idx
        elif 'category' in h or 'type' in h: cat_col = idx
        elif 'year' in h or 'era' in h: year_col = idx
        elif 'ceremony' in h or 'function' in h: cero_col = idx
        elif 'day' in h: day_col = idx
        elif 'youtube' in h or 'video' in h: yt_col = idx
        
    for r in range(2, ws.max_row+1):
        name = str(ws.cell(row=r, column=song_col).value or '').strip() if song_col else ''
        if not name or name.lower() == 'none': continue
        singer = str(ws.cell(row=r, column=singer_col).value or '').strip() if singer_col else ''
        movie = str(ws.cell(row=r, column=movie_col).value or '').strip() if movie_col else ''
        cat = str(ws.cell(row=r, column=cat_col).value or '').strip() if cat_col else ''
        year = str(ws.cell(row=r, column=year_col).value or '').strip() if year_col else ''
        cero = str(ws.cell(row=r, column=cero_col).value or '').strip() if cero_col else ''
        day = str(ws.cell(row=r, column=day_col).value or '').strip() if day_col else ''
        yt_link = ''
        if yt_col:
            c = ws.cell(row=r, column=yt_col)
            if c.hyperlink and c.hyperlink.target:
                yt_link = c.hyperlink.target
            elif c.value and 'http' in str(c.value):
                yt_link = str(c.value)
        k = re.sub(r'[^a-zA-Z0-9]', '', name.lower())
        if k not in meta or (yt_link and not meta[k].get('yt')):
            meta[k] = {'singer': singer, 'movie': movie, 'category': cat, 'year': year, 'ceremony': cero, 'day': day, 'yt': yt_link}

def extract_yt_id(url):
    if not url: return ''
    m = re.search(r'(?:v=|\/|embed\/|youtu\.be\/)([0-9A-Za-z_-]{11})(?:[&?]|$)', url)
    return m.group(1) if m else ''

punjabi_singers = [
    'diljit dosanjh', 'sukhbir', 'panjabi mc', 'mika singh', 'guru randhawa',
    'daler mehndi', 'jasbir jassi', 'bally sagoo', 'honey singh', 'badshah',
    'ap dhillon', 'karan aujla', 'hardy sandhu', 'amrit maan', 'sidhu moose wala',
    'gippy grewal', 'surinder kaur', 'gurmeet bawa', 'jasleen royal'
]

print("2. Scanning all 508 audio tracks in Brother Wedding Music...")
songs = []
for root, dirs, files in os.walk(MUSIC_DIR):
    for f in sorted(files):
        if any(f.lower().endswith(ext) for ext in ['.mp3', '.webm', '.m4a', '.opus']):
            full = os.path.join(root, f)
            rel = os.path.relpath(full, MUSIC_DIR)
            parts = rel.split(os.sep)
            top_folder = parts[0]
            sub_folder = parts[1] if len(parts) > 1 else ''
            
            name_no_ext = os.path.splitext(f)[0]
            if ' - ' in name_no_ext:
                title, artist = name_no_ext.split(' - ', 1)
            else:
                title, artist = name_no_ext, 'Traditional'
                
            norm_key = re.sub(r'[^a-zA-Z0-9]', '', title.lower())
            m = meta.get(norm_key, {})
            
            # Language determination
            s_low = (m.get('singer') or artist).lower()
            cat_low = m.get('category', '').lower()
            
            if 'bagheli' in top_folder.lower():
                lang = 'Bagheli'
            elif 'bhojpuri' in top_folder.lower():
                lang = 'Bhojpuri'
            elif 'punjabi' in sub_folder.lower() or 'punjabi' in cat_low or any(p in s_low for p in punjabi_singers):
                lang = 'Punjabi'
            else:
                lang = 'Hindi & Bollywood'
                
            # Day determination
            if 'bagheli' in top_folder.lower():
                day_key = 'bagheli'
                day_name = 'Bagheli (Rewa)'
            elif 'bhojpuri' in top_folder.lower():
                day_key = 'bhojpuri'
                day_name = 'Bhojpuri Songs'
            elif '10 december' in top_folder.lower():
                day_key = 'day1'
                day_name = '10 Dec: Mayan & Haldi'
            elif '11 december' in top_folder.lower():
                day_key = 'day2'
                day_name = '11 Dec: Mehndi & Sangeet'
            elif '12 december' in top_folder.lower():
                day_key = 'day3'
                day_name = '12 Dec: Wedding Day'
            else:
                day_key = 'other'
                day_name = 'Special Collection'
                
            # Function clean naming
            sub_lower = sub_folder.lower()
            if 'reception' in sub_lower: func = 'Grand Reception'
            elif 'vidaai' in sub_lower or 'bidaai' in sub_lower: func = 'Vidaai'
            elif 'pheras instrumental' in sub_lower: func = 'Pheras Instrumental'
            elif 'wedding rituals' in sub_lower: func = 'Wedding Rituals & Pheras'
            elif 'varmala' in sub_lower or 'jaimala' in sub_lower: func = 'Varmala & Jaimala'
            elif 'bride entry' in sub_lower: func = 'Bride Entry'
            elif 'ghodi' in sub_lower: func = 'Ghodi & Procession'
            elif 'baraat' in sub_lower or 'groom entry' in sub_lower: func = 'Baraat & Groom Entry'
            elif 'dj party' in sub_lower: func = 'DJ Party Hits'
            elif 'couple dance' in sub_lower: func = 'Couple Dance'
            elif 'family dance' in sub_lower: func = 'Family Dance'
            elif 'friends dance' in sub_lower: func = 'Friends Dance'
            elif 'ladies sangeet' in sub_lower: func = 'Ladies Sangeet'
            elif 'traditional wedding' in sub_lower: func = 'Traditional Sangeet'
            elif 'mehndi' in sub_lower: func = 'Mehndi Ceremony'
            elif 'punjabi haldi' in sub_lower: func = 'Punjabi Haldi'
            elif 'haldi' in sub_lower or 'nahchhu' in sub_lower: func = 'Haldi Ceremony'
            elif 'mayan' in sub_lower or 'mandap' in sub_lower: func = 'Mayan & Mandap'
            elif 'funny' in sub_lower: func = 'Funny & Entertaining'
            elif 'family celebration' in sub_lower: func = 'Family Celebration'
            elif 'evergreen' in sub_lower: func = 'Evergreen Classics'
            elif 'trending' in sub_lower: func = 'Trending Hits'
            else: func = sub_folder or 'Celebration'
            
            sub_path = rel.replace('\\', '/')
            root_path = "Brother Wedding Music/" + sub_path
            
            # YouTube Stream Link & Video ID
            raw_yt = m.get('yt', '')
            yt_id = extract_yt_id(raw_yt)
            if not raw_yt:
                clean_query = urllib.parse.quote(f"{title} {(m.get('singer') or artist)} wedding song audio")
                stream_link = f"https://www.youtube.com/results?search_query={clean_query}"
            else:
                stream_link = raw_yt

            songs.append({
                'id': len(songs) + 1,
                'filename': f,
                'title': title.strip(),
                'artist': m.get('singer') or artist.strip(),
                'movie': m.get('movie') or '',
                'year': m.get('year') or '',
                'category': m.get('category') or 'Wedding Anthem',
                'language': lang,
                'day_key': day_key,
                'day_name': day_name,
                'function': func,
                'folder': top_folder,
                'subfolder': sub_folder,
                'sub_path': sub_path,
                'root_path': root_path,
                'yt_id': yt_id,
                'stream_url': stream_link,
                'size_mb': round(os.path.getsize(full) / (1024*1024), 2)
            })

print(f"Total songs parsed: {len(songs)}")
songs_json = json.dumps(songs, ensure_ascii=False)

# Write songs json to external or embedded
print("3. Building HTML template...")

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
    <title>👑 Ramesh Brother Wedding DJ Player | 10-12 Dec 2026</title>
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="theme-color" content="#080a10">
    <style>
        :root {
            --bg-body: #080a10;
            --bg-sidebar: #0e111a;
            --bg-card: #141824;
            --bg-card-hover: #1c2234;
            --bg-player: rgba(12, 15, 24, 0.98);
            --gold: #f59e0b;
            --gold-light: #fbbf24;
            --gold-glow: rgba(245, 158, 11, 0.4);
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border: rgba(255, 255, 255, 0.08);
            --border-gold: rgba(245, 158, 11, 0.35);
            --neon-red: #f43f5e;
            --neon-green: #10b981;
            --neon-cyan: #06b6d4;
            --radius-sm: 8px;
            --radius-md: 12px;
            --radius-lg: 16px;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            user-select: none;
            -webkit-user-select: none;
            -webkit-tap-highlight-color: transparent;
            touch-action: manipulation;
        }

        html, body {
            width: 100%;
            height: 100%;
            overflow: hidden;
            background: var(--bg-body);
            color: var(--text-main);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        }

        /* App Root Layout */
        .app-root {
            display: flex;
            flex-direction: column;
            height: 100vh;
            height: 100dvh;
            width: 100vw;
            overflow: hidden;
            position: relative;
        }

        /* Top Navbar */
        .top-navbar {
            height: 54px;
            background: var(--bg-sidebar);
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 12px;
            flex-shrink: 0;
            z-index: 50;
        }

        .brand-box {
            display: flex;
            align-items: center;
            gap: 10px;
            cursor: pointer;
        }

        .brand-icon {
            width: 36px;
            height: 36px;
            border-radius: 8px;
            background: linear-gradient(135deg, #b45309, #f59e0b);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            color: #000;
            box-shadow: 0 0 12px var(--gold-glow);
            flex-shrink: 0;
        }

        .brand-title {
            font-size: 15px;
            font-weight: 800;
            letter-spacing: 0.8px;
            color: #fff;
            white-space: nowrap;
        }

        .brand-subtitle {
            font-size: 10px;
            color: var(--gold);
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }

        .top-actions {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .btn-icon {
            width: 38px;
            height: 38px;
            border-radius: 10px;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border);
            color: var(--text-main);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            cursor: pointer;
            position: relative;
            transition: all 0.15s ease;
        }

        .btn-icon:active {
            transform: scale(0.92);
            background: rgba(245, 158, 11, 0.2);
            border-color: var(--gold);
        }

        .badge-count {
            position: absolute;
            top: -4px;
            right: -4px;
            background: var(--neon-red);
            color: #fff;
            font-size: 10px;
            font-weight: 700;
            padding: 1px 5px;
            border-radius: 10px;
            min-width: 16px;
            text-align: center;
        }

        .badge-mode {
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 5px;
            background: rgba(16, 185, 129, 0.15);
            border: 1px solid rgba(16, 185, 129, 0.4);
            color: var(--neon-green);
        }

        /* Ceremony Quick Carousel */
        .cero-carousel {
            background: #0b0e16;
            border-bottom: 1px solid var(--border);
            display: flex;
            gap: 6px;
            overflow-x: auto;
            padding: 8px 10px;
            white-space: nowrap;
            flex-shrink: 0;
            scrollbar-width: none;
            -webkit-overflow-scrolling: touch;
        }
        .cero-carousel::-webkit-scrollbar { display: none; }

        .cero-chip {
            padding: 6px 12px;
            border-radius: 20px;
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border);
            color: var(--text-muted);
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            transition: all 0.15s;
        }

        .cero-chip:active {
            transform: scale(0.95);
        }

        .cero-chip.active {
            background: linear-gradient(135deg, rgba(245, 158, 11, 0.25), rgba(245, 158, 11, 0.1));
            border-color: var(--gold);
            color: var(--gold-light);
            box-shadow: 0 0 10px rgba(245, 158, 11, 0.2);
        }

        /* Workspace: Desktop Sidebar + Stage */
        .workspace {
            flex: 1;
            display: flex;
            overflow: hidden;
            position: relative;
        }

        /* Desktop Sidebar (Hidden on mobile) */
        .desktop-sidebar {
            width: 290px;
            background: var(--bg-sidebar);
            border-right: 1px solid var(--border);
            display: flex;
            flex-direction: column;
            overflow-y: auto;
            flex-shrink: 0;
        }

        .sidebar-section {
            padding: 12px 14px;
            border-bottom: 1px solid var(--border);
        }

        .sidebar-title {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: var(--text-muted);
            font-weight: 700;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .nav-link {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 8px 12px;
            border-radius: var(--radius-sm);
            color: #cbd5e1;
            font-size: 13px;
            cursor: pointer;
            margin-bottom: 3px;
            transition: all 0.15s;
        }

        .nav-link:hover, .nav-link.active {
            background: rgba(245, 158, 11, 0.15);
            color: var(--gold-light);
            font-weight: 600;
        }

        .nav-link .tag {
            font-size: 11px;
            background: rgba(255, 255, 255, 0.08);
            padding: 1px 6px;
            border-radius: 8px;
            color: var(--text-muted);
        }

        .nav-link.active .tag {
            background: var(--gold);
            color: #000;
            font-weight: 700;
        }

        .lang-pills {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
        }

        .lang-pill {
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 11px;
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border);
            color: var(--text-muted);
            cursor: pointer;
        }

        .lang-pill.active {
            background: var(--gold);
            color: #000;
            font-weight: 700;
            border-color: var(--gold);
        }

        /* Stage / Main Content Area */
        .stage-area {
            flex: 1;
            display: flex;
            flex-direction: column;
            overflow-y: auto;
            overflow-x: hidden;
            background: radial-gradient(circle at 50% 0%, #171d2d 0%, var(--bg-body) 65%);
            -webkit-overflow-scrolling: touch;
            padding-bottom: 76px; /* space for bottom player */
        }

        /* Stage Header Banner */
        .stage-banner {
            padding: 14px 16px;
            display: flex;
            flex-direction: column;
            gap: 10px;
            border-bottom: 1px solid var(--border);
            background: rgba(14, 17, 26, 0.6);
            backdrop-filter: blur(8px);
        }

        .banner-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 10px;
        }

        .banner-title {
            font-size: 20px;
            font-weight: 800;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .banner-meta {
            font-size: 12px;
            color: var(--gold);
            font-weight: 500;
        }

        .banner-controls {
            display: flex;
            gap: 8px;
            align-items: center;
        }

        .btn-gold {
            background: linear-gradient(135deg, #f59e0b, #d97706);
            color: #000;
            font-weight: 700;
            border: none;
            padding: 7px 14px;
            border-radius: 8px;
            font-size: 12px;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 0 10px var(--gold-glow);
        }

        .btn-gold:active {
            transform: scale(0.95);
        }

        .btn-outline {
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border);
            color: var(--text-main);
            padding: 7px 12px;
            border-radius: 8px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 5px;
        }

        .btn-outline:active {
            transform: scale(0.95);
            border-color: var(--gold);
        }

        /* Search Bar */
        .search-bar-wrap {
            padding: 8px 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .search-input-box {
            flex: 1;
            position: relative;
            display: flex;
            align-items: center;
        }

        .search-input-box input {
            width: 100%;
            height: 38px;
            background: rgba(20, 24, 36, 0.8);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 0 34px 0 34px;
            color: #fff;
            font-size: 13px;
            outline: none;
            transition: all 0.2s;
        }

        .search-input-box input:focus {
            border-color: var(--gold);
            box-shadow: 0 0 10px var(--gold-glow);
        }

        .search-icon-left {
            position: absolute;
            left: 10px;
            font-size: 14px;
            color: var(--text-muted);
            pointer-events: none;
        }

        .search-clear-right {
            position: absolute;
            right: 10px;
            font-size: 14px;
            color: var(--text-muted);
            cursor: pointer;
            display: none;
        }

        /* Mobile Card List */
        .mobile-song-list {
            display: flex;
            flex-direction: column;
            gap: 6px;
            padding: 8px 12px;
        }

        .m-song-card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 10px 12px;
            display: flex;
            align-items: center;
            gap: 10px;
            transition: all 0.15s ease;
            position: relative;
        }

        .m-song-card:active {
            background: var(--bg-card-hover);
        }

        .m-song-card.playing {
            background: #171c2c;
            border-color: var(--gold);
            box-shadow: 0 0 16px rgba(245, 158, 11, 0.25);
        }

        .m-card-idx {
            width: 26px;
            height: 26px;
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.04);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 11px;
            font-weight: 700;
            color: var(--text-muted);
            flex-shrink: 0;
        }

        .m-song-card.playing .m-card-idx {
            background: var(--gold);
            color: #000;
        }

        .m-card-info {
            flex: 1;
            min-width: 0;
            display: flex;
            flex-direction: column;
            gap: 3px;
            cursor: pointer;
        }

        .m-card-title {
            font-size: 14px;
            font-weight: 700;
            color: #fff;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .m-song-card.playing .m-card-title {
            color: var(--gold-light);
        }

        .m-card-sub {
            font-size: 11px;
            color: var(--text-muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .m-card-badges {
            display: flex;
            align-items: center;
            gap: 5px;
            margin-top: 2px;
            flex-wrap: wrap;
        }

        .badge-chip {
            font-size: 9px;
            padding: 2px 5px;
            border-radius: 4px;
            background: rgba(255, 255, 255, 0.05);
            color: var(--text-muted);
            border: 1px solid var(--border);
        }

        .badge-chip.cero {
            color: var(--gold);
            border-color: var(--border-gold);
            background: rgba(245, 158, 11, 0.08);
        }

        .m-card-actions {
            display: flex;
            align-items: center;
            gap: 6px;
            flex-shrink: 0;
        }

        .btn-play-card {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            background: linear-gradient(135deg, #f59e0b, #d97706);
            color: #000;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            cursor: pointer;
            box-shadow: 0 0 10px var(--gold-glow);
            flex-shrink: 0;
        }

        .btn-play-card:active {
            transform: scale(0.9);
        }

        .btn-card-act {
            width: 32px;
            height: 32px;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border);
            color: var(--text-muted);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 13px;
            cursor: pointer;
        }

        .btn-card-act:active {
            transform: scale(0.92);
        }

        .btn-card-act.is-fav {
            color: var(--neon-red);
            border-color: rgba(244, 63, 94, 0.4);
            background: rgba(244, 63, 94, 0.1);
        }

        /* Desktop Cue Table (Hidden on Mobile) */
        .desktop-table-container {
            display: none;
            padding: 12px 16px;
        }

        .cue-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }

        .cue-table th {
            text-align: left;
            padding: 10px 12px;
            background: #101420;
            color: var(--text-muted);
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            border-bottom: 1px solid var(--border);
        }

        .cue-table td {
            padding: 10px 12px;
            border-bottom: 1px solid var(--border);
            vertical-align: middle;
        }

        .cue-table tr:hover td {
            background: rgba(255, 255, 255, 0.03);
        }

        .cue-table tr.playing td {
            background: rgba(245, 158, 11, 0.08);
            border-bottom: 1px solid var(--gold);
        }

        .cue-table tr.playing .t-title {
            color: var(--gold-light);
            font-weight: 700;
        }

        /* Bottom Sticky Mini Player */
        .bottom-mini-player {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            height: 66px;
            background: var(--bg-player);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border-top: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 12px;
            padding-bottom: env(safe-area-inset-bottom, 0);
            z-index: 100;
            box-shadow: 0 -4px 20px rgba(0, 0, 0, 0.5);
        }

        /* Scrubber Line on top of mini player */
        .mini-progress-bar {
            position: absolute;
            top: -2px;
            left: 0;
            right: 0;
            height: 3px;
            background: rgba(255, 255, 255, 0.1);
            cursor: pointer;
        }

        .mini-progress-fill {
            height: 100%;
            width: 0%;
            background: linear-gradient(90deg, #d97706, #fbbf24);
            box-shadow: 0 0 8px var(--gold-glow);
        }

        .mini-left-content {
            display: flex;
            align-items: center;
            gap: 10px;
            flex: 1;
            min-width: 0;
            cursor: pointer;
        }

        .mini-vinyl {
            width: 44px;
            height: 44px;
            border-radius: 50%;
            background: radial-gradient(circle, #2d261e 18%, #111 20%, #1a1a1a 60%, #000 100%);
            border: 2px solid var(--gold);
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            flex-shrink: 0;
            box-shadow: 0 0 10px rgba(245, 158, 11, 0.3);
            animation: spin 6s linear infinite paused;
        }

        .mini-vinyl.spinning {
            animation-play-state: running;
        }

        .mini-vinyl-center {
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background: var(--gold);
            border: 2px solid #000;
        }

        .mini-track-info {
            display: flex;
            flex-direction: column;
            min-width: 0;
        }

        .mini-track-title {
            font-size: 14px;
            font-weight: 700;
            color: #fff;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .mini-track-sub {
            font-size: 11px;
            color: var(--gold);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .mini-right-controls {
            display: flex;
            align-items: center;
            gap: 10px;
            flex-shrink: 0;
        }

        .btn-mini-play {
            width: 46px;
            height: 46px;
            border-radius: 50%;
            background: linear-gradient(135deg, #f59e0b, #d97706);
            color: #000;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            box-shadow: 0 0 14px var(--gold-glow);
            cursor: pointer;
        }

        .btn-mini-play:active {
            transform: scale(0.92);
        }

        .btn-mini-ctrl {
            width: 38px;
            height: 38px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border);
            color: var(--text-main);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 15px;
            cursor: pointer;
        }

        .btn-mini-ctrl:active {
            transform: scale(0.92);
        }

        /* Full-Screen Mobile DJ Sheet */
        .mobile-dj-sheet {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(180deg, #171d2d 0%, #0a0d14 100%);
            z-index: 200;
            display: flex;
            flex-direction: column;
            padding: 16px;
            padding-bottom: calc(16px + env(safe-area-inset-bottom, 0));
            transform: translateY(100%);
            transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            overflow-y: auto;
        }

        .mobile-dj-sheet.open {
            transform: translateY(0);
        }

        .sheet-top-bar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
        }

        .sheet-minimize {
            font-size: 13px;
            color: var(--gold);
            display: flex;
            align-items: center;
            gap: 4px;
            cursor: pointer;
            padding: 6px 10px;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.05);
        }

        .sheet-center-disc {
            flex: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 200px;
            position: relative;
        }

        .hero-vinyl {
            width: 175px;
            height: 175px;
            border-radius: 50%;
            background: radial-gradient(circle, #2d261e 18%, #111 20%, #1a1a1a 60%, #000 100%);
            border: 3px solid var(--gold);
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            box-shadow: 0 0 40px var(--gold-glow);
            animation: spin 5s linear infinite paused;
        }

        .hero-vinyl.spinning {
            animation-play-state: running;
        }

        .hero-vinyl-center {
            width: 48px;
            height: 48px;
            border-radius: 50%;
            background: linear-gradient(135deg, #d97706, #fbbf24);
            border: 3px solid #000;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
        }

        .sheet-info-box {
            text-align: center;
            margin: 14px 0 10px;
            padding: 0 10px;
        }

        .sheet-title {
            font-size: 18px;
            font-weight: 800;
            color: #fff;
            margin-bottom: 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .sheet-artist {
            font-size: 13px;
            color: var(--gold);
            margin-bottom: 6px;
        }

        .sheet-tags {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
        }

        /* Scrubber in Sheet */
        .sheet-scrub-box {
            margin: 10px 0;
            display: flex;
            flex-direction: column;
            gap: 5px;
        }

        .sheet-slider {
            -webkit-appearance: none;
            width: 100%;
            height: 6px;
            border-radius: 3px;
            background: rgba(255, 255, 255, 0.15);
            outline: none;
        }

        .sheet-slider::-webkit-slider-thumb {
            -webkit-appearance: none;
            width: 18px;
            height: 18px;
            border-radius: 50%;
            background: var(--gold);
            box-shadow: 0 0 10px var(--gold);
            cursor: pointer;
        }

        .sheet-time-row {
            display: flex;
            justify-content: space-between;
            font-size: 11px;
            color: var(--text-muted);
            font-variant-numeric: tabular-nums;
        }

        /* Big DJ Controls */
        .sheet-ctrls-row {
            display: flex;
            align-items: center;
            justify-content: space-evenly;
            margin: 12px 0;
        }

        .btn-sheet-sub {
            width: 44px;
            height: 44px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.06);
            border: 1px solid var(--border);
            color: var(--text-main);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            cursor: pointer;
        }

        .btn-sheet-sub.active {
            color: var(--gold);
            border-color: var(--gold);
            background: rgba(245, 158, 11, 0.15);
        }

        .btn-sheet-hero {
            width: 68px;
            height: 68px;
            border-radius: 50%;
            background: linear-gradient(135deg, #fbbf24, #d97706);
            color: #000;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 26px;
            box-shadow: 0 0 25px var(--gold-glow);
            cursor: pointer;
        }

        .btn-sheet-hero:active {
            transform: scale(0.92);
        }

        /* DJ Sound FX Pads */
        .sheet-fx-section {
            background: rgba(0, 0, 0, 0.25);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 10px 12px;
            margin-top: 8px;
        }

        .sheet-fx-title {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: var(--text-muted);
            margin-bottom: 8px;
            text-align: center;
            font-weight: 700;
        }

        .sheet-fx-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 6px;
        }

        .btn-fx-pad {
            padding: 8px 4px;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border);
            color: #fff;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 3px;
            cursor: pointer;
            font-size: 10px;
            font-weight: 600;
        }

        .btn-fx-pad:active {
            transform: scale(0.92);
            background: var(--gold);
            color: #000;
        }

        .btn-fx-pad span.icon {
            font-size: 16px;
        }

        /* Off-canvas Navigation Drawer */
        .drawer-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            backdrop-filter: blur(4px);
            z-index: 300;
            opacity: 0;
            pointer-events: none;
            transition: opacity 0.25s ease;
        }

        .drawer-overlay.open {
            opacity: 1;
            pointer-events: auto;
        }

        .mobile-drawer {
            position: fixed;
            top: 0;
            left: 0;
            bottom: 0;
            width: 300px;
            max-width: 82vw;
            background: var(--bg-sidebar);
            border-right: 1px solid var(--border);
            z-index: 310;
            transform: translateX(-100%);
            transition: transform 0.25s cubic-bezier(0.16, 1, 0.3, 1);
            overflow-y: auto;
            display: flex;
            flex-direction: column;
        }

        .mobile-drawer.open {
            transform: translateX(0);
        }

        .drawer-header {
            padding: 16px;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        /* Queue Sheet */
        .queue-sheet {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            max-height: 75vh;
            background: #111522;
            border-top: 1px solid var(--border-gold);
            border-radius: 16px 16px 0 0;
            z-index: 250;
            display: flex;
            flex-direction: column;
            transform: translateY(100%);
            transition: transform 0.25s ease;
            box-shadow: 0 -8px 30px rgba(0, 0, 0, 0.8);
        }

        .queue-sheet.open {
            transform: translateY(0);
        }

        .queue-sheet-header {
            padding: 14px 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 1px solid var(--border);
        }

        .queue-items-box {
            overflow-y: auto;
            padding: 10px 14px;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        /* Toast Notifications */
        .toast-notify {
            position: fixed;
            top: 64px;
            left: 50%;
            transform: translateX(-50%) translateY(-20px);
            background: #1e2538;
            border: 1px solid var(--gold);
            color: #fff;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.6);
            z-index: 500;
            opacity: 0;
            pointer-events: none;
            transition: all 0.25s ease;
            white-space: nowrap;
            max-width: 90vw;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .toast-notify.show {
            opacity: 1;
            transform: translateX(-50%) translateY(0);
        }

        /* Equalizer Animation Bars */
        .eq-bars {
            display: inline-flex;
            align-items: flex-end;
            gap: 2px;
            height: 14px;
        }

        .eq-bar {
            width: 3px;
            background: var(--gold);
            border-radius: 1px;
            animation: bounce 0.6s ease infinite alternate;
        }
        .eq-bar:nth-child(1) { height: 6px; animation-delay: 0.1s; }
        .eq-bar:nth-child(2) { height: 12px; animation-delay: 0.3s; }
        .eq-bar:nth-child(3) { height: 8px; animation-delay: 0.2s; }
        .eq-bar:nth-child(4) { height: 14px; animation-delay: 0.4s; }

        @keyframes spin {
            100% { transform: rotate(360deg); }
        }

        @keyframes bounce {
            0% { height: 3px; }
            100% { height: 14px; }
        }

        /* Desktop Media Queries */
        @media (min-width: 769px) {
            .mobile-only { display: none !important; }
            .desktop-sidebar { display: flex; }
            .desktop-table-container { display: block; }
            .mobile-song-list { display: none; }
            .stage-area { padding-bottom: 96px; }
            .bottom-mini-player {
                height: 86px;
                padding: 0 24px;
            }
            .mini-vinyl { width: 54px; height: 54px; }
            .mini-track-title { font-size: 16px; }
            .mini-track-sub { font-size: 13px; }
            .desktop-deck-center {
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 6px;
                flex: 1;
                max-width: 600px;
                margin: 0 24px;
            }
            .desktop-deck-right {
                display: flex;
                align-items: center;
                gap: 16px;
                flex-shrink: 0;
            }
        }

        @media (max-width: 768px) {
            .desktop-only { display: none !important; }
            .desktop-sidebar { display: none; }
            .desktop-table-container { display: none; }
            .mobile-song-list { display: flex; }
        }
    </style>
</head>
<body>

    <audio id="audio-core" playsinline webkit-playsinline preload="auto"></audio>
    <div id="yt-player-container" style="position:fixed;bottom:-999px;left:-999px;width:1px;height:1px;opacity:0.01;pointer-events:none;"></div>

    <!-- Hidden folder input for connecting phone / local music folder -->
    <input type="file" id="local-folder-picker" webkitdirectory directory multiple accept="audio/*" style="display:none;">

    <!-- Toast Notification -->
    <div class="toast-notify" id="toast-notify"></div>

    <div class="app-root">
        <!-- Top Navbar -->
        <header class="top-navbar">
            <div style="display: flex; align-items: center; gap: 8px;">
                <button class="btn-icon mobile-only" id="btn-toggle-drawer" title="Open Navigation Menu">☰</button>
                <div class="brand-box" onclick="resetFilters()">
                    <div class="brand-icon">👑</div>
                    <div>
                        <div class="brand-title">WEDDING DJ</div>
                        <div class="brand-subtitle">10-12 DEC 2026</div>
                    </div>
                </div>
            </div>

            <div class="top-actions">
                <div class="badge-mode" id="status-mode-pill" title="Playback Mode">
                    <span id="mode-dot">●</span>
                    <span id="mode-label">Local Ready</span>
                </div>
                <button class="btn-icon" id="btn-connect-folder" title="Connect Local Music Folder (508 MP3s)">📁</button>
                <button class="btn-icon" id="btn-open-queue" title="Open Queue">
                    📑
                    <span class="badge-count" id="queue-badge-count">0</span>
                </button>
            </div>
        </header>

        <!-- Ceremony Quick-Select Carousel (Horizontal touch scroll) -->
        <nav class="cero-carousel" id="cero-carousel">
            <!-- Dynamically populated with chips -->
        </nav>

        <!-- Main Workspace -->
        <div class="workspace">
            <!-- Desktop Sidebar -->
            <aside class="desktop-sidebar desktop-only">
                <div class="sidebar-section">
                    <div class="sidebar-title">📅 Wedding Days</div>
                    <div id="desktop-day-nav"></div>
                </div>
                <div class="sidebar-section">
                    <div class="sidebar-title">🌍 Language Filter</div>
                    <div class="lang-pills" id="desktop-lang-nav"></div>
                </div>
                <div class="sidebar-section" style="flex:1;">
                    <div class="sidebar-title">🎭 Ceremony Functions</div>
                    <div id="desktop-cero-nav" style="display:flex; flex-direction:column; gap:2px;"></div>
                </div>
            </aside>

            <!-- Stage Area -->
            <main class="stage-area" id="stage-area">
                <!-- Search Box -->
                <div class="search-bar-wrap">
                    <div class="search-input-box">
                        <span class="search-icon-left">🔍</span>
                        <input type="text" id="main-search-input" placeholder="Search by song name, singer, movie, ceremony..." autocomplete="off">
                        <span class="search-clear-right" id="search-clear-btn">✕</span>
                    </div>
                </div>

                <!-- Stage Header Banner -->
                <div class="stage-banner">
                    <div class="banner-row">
                        <div>
                            <div class="banner-title" id="stage-title">All Wedding Tracks</div>
                            <div class="banner-meta" id="stage-meta">508 Songs Available • Ready for Mayan, Haldi, Sangeet & Baraat</div>
                        </div>
                        <div class="banner-controls">
                            <button class="btn-gold" onclick="playAllFiltered(false)">▶ Play All</button>
                            <button class="btn-outline" onclick="playAllFiltered(true)">🔀 Shuffle</button>
                        </div>
                    </div>
                </div>

                <!-- Mobile Card List (<768px) -->
                <div class="mobile-song-list" id="mobile-song-list">
                    <!-- Populated dynamically -->
                </div>

                <!-- Desktop Cue Table (>768px) -->
                <div class="desktop-table-container desktop-only">
                    <table class="cue-table">
                        <thead>
                            <tr>
                                <th style="width: 45px;">#</th>
                                <th>Song Title & Movie</th>
                                <th>Singer(s)</th>
                                <th>Function</th>
                                <th>Language</th>
                                <th>Format</th>
                                <th style="text-align: right;">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="desktop-table-body">
                            <!-- Populated dynamically -->
                        </tbody>
                    </table>
                </div>
            </main>
        </div>

        <!-- Sticky Bottom Mini Player -->
        <footer class="bottom-mini-player" id="mini-player-bar">
            <!-- Scrubber line on top -->
            <div class="mini-progress-bar" id="mini-scrubber-track">
                <div class="mini-progress-fill" id="mini-scrubber-fill"></div>
            </div>

            <!-- Left: vinyl + song info (tapping expands Fullscreen DJ Deck) -->
            <div class="mini-left-content" id="mini-expand-trigger">
                <div class="mini-vinyl" id="mini-vinyl">
                    <div class="mini-vinyl-center"></div>
                </div>
                <div class="mini-track-info">
                    <div class="mini-track-title" id="mini-title">Select a Song to DJ</div>
                    <div class="mini-track-sub" id="mini-sub">Ramesh Brother Wedding • 508 Tracks</div>
                </div>
            </div>

            <!-- Desktop Controls Center (>768px) -->
            <div class="desktop-deck-center desktop-only">
                <div style="display:flex; align-items:center; gap:14px;">
                    <button class="btn-mini-ctrl" onclick="toggleShuffle()" id="btn-desk-shuffle" title="Shuffle (S)">🔀</button>
                    <button class="btn-mini-ctrl" onclick="playPrev()" title="Previous (Left Arrow)">⏮</button>
                    <button class="btn-mini-play" onclick="togglePlayPause()" id="btn-desk-play" title="Play/Pause (Space)">▶</button>
                    <button class="btn-mini-ctrl" onclick="playNext()" title="Next (Right Arrow)">⏭</button>
                    <button class="btn-mini-ctrl" onclick="toggleRepeat()" id="btn-desk-repeat" title="Repeat (R)">🔁</button>
                </div>
                <div style="display:flex; align-items:center; gap:8px; width:100%; font-size:11px; color:var(--text-muted);">
                    <span id="desk-time-cur">00:00</span>
                    <input type="range" class="sheet-slider" id="desk-progress-slider" min="0" max="100" value="0">
                    <span id="desk-time-dur">00:00</span>
                </div>
            </div>

            <!-- Desktop Controls Right (>768px) -->
            <div class="desktop-deck-right desktop-only">
                <div style="display:flex; align-items:center; gap:6px;">
                    <span>🔊</span>
                    <input type="range" class="sheet-slider" id="desk-vol-slider" min="0" max="1" step="0.05" value="0.9" style="width:80px;">
                </div>
                <div style="display:flex; gap:4px;">
                    <button class="btn-fx-pad" onclick="playFX('horn')" title="Air Horn (1)"><span class="icon">🎺</span>Horn</button>
                    <button class="btn-fx-pad" onclick="playFX('dhol')" title="Dhol (2)"><span class="icon">🥁</span>Dhol</button>
                    <button class="btn-fx-pad" onclick="playFX('cheer')" title="Cheer (3)"><span class="icon">👏</span>Cheer</button>
                    <button class="btn-fx-pad" onclick="playFX('scratch')" title="Scratch (4)"><span class="icon">🎚️</span>Scratch</button>
                </div>
            </div>

            <!-- Mobile Controls Right (<768px) -->
            <div class="mini-right-controls mobile-only">
                <button class="btn-mini-play" id="btn-mini-play" onclick="togglePlayPause()">▶</button>
                <button class="btn-mini-ctrl" id="btn-mini-next" onclick="playNext()">⏭</button>
            </div>
        </footer>
    </div>

    <!-- Mobile Fullscreen DJ Deck Sheet -->
    <div class="mobile-dj-sheet" id="mobile-dj-sheet">
        <div class="sheet-top-bar">
            <div class="sheet-minimize" onclick="closeSheet()">▼ Minimize</div>
            <div style="font-size:12px; font-weight:700; color:var(--gold);" id="sheet-header-cero">HALDI CEREMONY</div>
            <button class="btn-card-act" id="sheet-fav-btn" onclick="toggleFavCurrent()">🤍</button>
        </div>

        <div class="sheet-center-disc">
            <div class="hero-vinyl" id="sheet-vinyl">
                <div class="hero-vinyl-center">🎵</div>
            </div>
        </div>

        <div class="sheet-info-box">
            <div class="sheet-title" id="sheet-title">Song Title</div>
            <div class="sheet-artist" id="sheet-artist">Artist • Movie</div>
            <div class="sheet-tags">
                <span class="badge-chip cero" id="sheet-cero-tag">Wedding Anthem</span>
                <span class="badge-chip" id="sheet-lang-tag">Hindi</span>
            </div>
        </div>

        <!-- Scrubber -->
        <div class="sheet-scrub-box">
            <input type="range" class="sheet-slider" id="sheet-progress-slider" min="0" max="100" value="0">
            <div class="sheet-time-row">
                <span id="sheet-time-cur">00:00</span>
                <span id="sheet-time-dur">00:00</span>
            </div>
        </div>

        <!-- Big Controls Row -->
        <div class="sheet-ctrls-row">
            <button class="btn-sheet-sub" id="sheet-shuffle-btn" onclick="toggleShuffle()">🔀</button>
            <button class="btn-sheet-sub" onclick="playPrev()">⏮</button>
            <button class="btn-sheet-hero" id="sheet-play-btn" onclick="togglePlayPause()">▶</button>
            <button class="btn-sheet-sub" onclick="playNext()">⏭</button>
            <button class="btn-sheet-sub" id="sheet-repeat-btn" onclick="toggleRepeat()">🔁</button>
        </div>

        <!-- DJ Sound FX Section -->
        <div class="sheet-fx-section">
            <div class="sheet-fx-title">🔥 Live DJ Hype FX Pads</div>
            <div class="sheet-fx-grid">
                <button class="btn-fx-pad" onclick="playFX('horn')"><span class="icon">🎺</span>Air Horn</button>
                <button class="btn-fx-pad" onclick="playFX('dhol')"><span class="icon">🥁</span>Dhol Beat</button>
                <button class="btn-fx-pad" onclick="playFX('cheer')"><span class="icon">👏</span>Crowd Cheer</button>
                <button class="btn-fx-pad" onclick="playFX('scratch')"><span class="icon">🎚️</span>DJ Scratch</button>
            </div>
        </div>

        <!-- Action Row -->
        <div style="display:flex; gap:10px; margin-top:14px;">
            <button class="btn-outline" style="flex:1; justify-content:center; padding:10px;" id="sheet-btn-stream" onclick="openCurrentStream()">🌐 Watch Video / Stream</button>
            <button class="btn-outline" style="flex:1; justify-content:center; padding:10px;" onclick="addCurrentToQueue()">➕ Add to Queue</button>
        </div>
    </div>

    <!-- Navigation Drawer & Overlay -->
    <div class="drawer-overlay" id="drawer-overlay" onclick="closeDrawer()"></div>
    <div class="mobile-drawer" id="mobile-drawer">
        <div class="drawer-header">
            <div style="font-weight:800; font-size:15px; color:#fff;">WEDDING NAVIGATION</div>
            <button class="btn-card-act" onclick="closeDrawer()">✕</button>
        </div>
        <div style="padding:14px; display:flex; flex-direction:column; gap:16px;">
            <!-- Folder Connect Button -->
            <div style="background:rgba(245,158,11,0.08); border:1px solid var(--border-gold); border-radius:10px; padding:12px;">
                <div style="font-size:12px; font-weight:700; color:var(--gold); margin-bottom:4px;">📁 Offline Music Folder</div>
                <div style="font-size:11px; color:var(--text-muted); margin-bottom:8px;">Load "Brother Wedding Music" folder from your phone storage for 100% offline playback.</div>
                <button class="btn-gold" style="width:100%; justify-content:center;" onclick="document.getElementById('local-folder-picker').click()">⚡ Connect Audio Folder</button>
            </div>

            <!-- Days -->
            <div>
                <div class="sidebar-title">📅 Wedding Days</div>
                <div id="drawer-days-list" style="display:flex; flex-direction:column; gap:3px;"></div>
            </div>

            <!-- Languages -->
            <div>
                <div class="sidebar-title">🌍 Languages</div>
                <div class="lang-pills" id="drawer-lang-pills"></div>
            </div>

            <!-- Ceremonies -->
            <div>
                <div class="sidebar-title">🎭 Ceremony Functions</div>
                <div id="drawer-cero-list" style="display:flex; flex-direction:column; gap:3px;"></div>
            </div>
        </div>
    </div>

    <!-- Playlist Queue Sheet -->
    <div class="drawer-overlay" id="queue-overlay" onclick="closeQueue()"></div>
    <div class="queue-sheet" id="queue-sheet">
        <div class="queue-sheet-header">
            <div style="font-weight:700; font-size:14px; color:#fff;">📑 Live Playlist Queue (<span id="queue-count-display">0</span>)</div>
            <div style="display:flex; gap:8px;">
                <button class="btn-outline" style="padding:4px 8px; font-size:11px;" onclick="clearQueue()">Clear</button>
                <button class="btn-card-act" onclick="closeQueue()">✕</button>
            </div>
        </div>
        <div class="queue-items-box" id="queue-items-box">
            <!-- Dynamic Queue items -->
        </div>
    </div>

    <!-- YouTube IFrame API -->
    <script src="https://www.youtube.com/iframe_api"></script>

    <script>
        // --- SONGS DATABASE ---
        const SONGS = __SONGS_DATA__;

        // --- APPLICATION STATE ---
        const state = {
            currentIndex: -1,
            isPlaying: false,
            activeDay: 'all',
            activeCero: 'all',
            activeLang: 'all',
            searchQuery: '',
            isShuffle: false,
            repeatMode: 'all', // 'off', 'all', 'one'
            favorites: new Set(),
            queue: [],
            localFilesMap: new Map(), // filename.toLowerCase() -> File object
            isLocalMode: false,
            ytPlayerReady: false,
            isStreaming: false
        };

        // DOM Elements
        const audio = document.getElementById('audio-core');
        const miniPlayer = document.getElementById('mini-player-bar');
        const miniVinyl = document.getElementById('mini-vinyl');
        const miniTitle = document.getElementById('mini-title');
        const miniSub = document.getElementById('mini-sub');
        const miniPlayBtn = document.getElementById('btn-mini-play');
        const miniScrubFill = document.getElementById('mini-scrubber-fill');
        const miniScrubTrack = document.getElementById('mini-scrubber-track');

        const sheet = document.getElementById('mobile-dj-sheet');
        const sheetVinyl = document.getElementById('sheet-vinyl');
        const sheetTitle = document.getElementById('sheet-title');
        const sheetArtist = document.getElementById('sheet-artist');
        const sheetCeroTag = document.getElementById('sheet-cero-tag');
        const sheetLangTag = document.getElementById('sheet-lang-tag');
        const sheetHeaderCero = document.getElementById('sheet-header-cero');
        const sheetPlayBtn = document.getElementById('sheet-play-btn');
        const sheetShuffleBtn = document.getElementById('sheet-shuffle-btn');
        const sheetRepeatBtn = document.getElementById('sheet-repeat-btn');
        const sheetFavBtn = document.getElementById('sheet-fav-btn');
        const sheetSlider = document.getElementById('sheet-progress-slider');
        const sheetTimeCur = document.getElementById('sheet-time-cur');
        const sheetTimeDur = document.getElementById('sheet-time-dur');

        // Desktop Elements
        const deskPlayBtn = document.getElementById('btn-desk-play');
        const deskShuffleBtn = document.getElementById('btn-desk-shuffle');
        const deskRepeatBtn = document.getElementById('btn-desk-repeat');
        const deskSlider = document.getElementById('desk-progress-slider');
        const deskTimeCur = document.getElementById('desk-time-cur');
        const deskTimeDur = document.getElementById('desk-time-dur');
        const deskVolSlider = document.getElementById('desk-vol-slider');

        // Mobile Drawer & Queue
        const drawer = document.getElementById('mobile-drawer');
        const drawerOverlay = document.getElementById('drawer-overlay');
        const queueSheet = document.getElementById('queue-sheet');
        const queueOverlay = document.getElementById('queue-overlay');
        const queueBadge = document.getElementById('queue-badge-count');
        const queueDisplay = document.getElementById('queue-count-display');
        const queueItemsBox = document.getElementById('queue-items-box');

        // Mode Pill
        const statusPill = document.getElementById('status-mode-pill');
        const modeLabel = document.getElementById('mode-label');
        const modeDot = document.getElementById('mode-dot');

        // --- WEB AUDIO API DJ FX GENERATOR ---
        let audioCtx = null;
        function getAudioContext() {
            if (!audioCtx) {
                const AudioContext = window.AudioContext || window.webkitAudioContext;
                audioCtx = new AudioContext();
            }
            if (audioCtx.state === 'suspended') {
                audioCtx.resume();
            }
            return audioCtx;
        }

        // Global touch unlock for iOS / Android mobile
        let audioUnlocked = false;
        function unlockAudioEngine() {
            if (audioUnlocked) return;
            getAudioContext();
            audio.play().then(() => {
                audio.pause();
                audioUnlocked = true;
            }).catch(() => {});
        }
        document.addEventListener('touchstart', unlockAudioEngine, { passive: true });
        document.addEventListener('click', unlockAudioEngine, { passive: true });

        // Synthesize Instant DJ FX with Web Audio API (Zero external assets, works 100% offline!)
        function playFX(type) {
            try {
                const ctx = getAudioContext();
                const now = ctx.currentTime;

                if (type === 'horn') {
                    // DJ Air Horn
                    [277.18, 369.99, 554.37].forEach((freq) => {
                        const osc = ctx.createOscillator();
                        const gain = ctx.createGain();
                        osc.type = 'sawtooth';
                        osc.frequency.setValueAtTime(freq, now);
                        osc.frequency.exponentialRampToValueAtTime(freq * 0.96, now + 0.4);
                        gain.gain.setValueAtTime(0.2, now);
                        gain.gain.exponentialRampToValueAtTime(0.01, now + 0.45);
                        osc.connect(gain);
                        gain.connect(ctx.destination);
                        osc.start(now);
                        osc.stop(now + 0.45);
                    });
                    toast('🎺 Air Horn!');
                } else if (type === 'dhol') {
                    // Punjabi Dhol Beat
                    const osc = ctx.createOscillator();
                    const gain = ctx.createGain();
                    osc.type = 'sine';
                    osc.frequency.setValueAtTime(140, now);
                    osc.frequency.exponentialRampToValueAtTime(45, now + 0.25);
                    gain.gain.setValueAtTime(0.6, now);
                    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.3);
                    osc.connect(gain);
                    gain.connect(ctx.destination);
                    osc.start(now);
                    osc.stop(now + 0.3);
                    toast('🥁 Dhol Roll!');
                } else if (type === 'cheer') {
                    // Crowd Cheer Synth
                    const bufferSize = ctx.sampleRate * 0.8;
                    const buffer = ctx.createBuffer(1, bufferSize, ctx.sampleRate);
                    const data = buffer.getChannelData(0);
                    for (let i = 0; i < bufferSize; i++) {
                        data[i] = Math.random() * 2 - 1;
                    }
                    const noise = ctx.createBufferSource();
                    noise.buffer = buffer;
                    const filter = ctx.createBiquadFilter();
                    filter.type = 'bandpass';
                    filter.frequency.value = 1200;
                    filter.Q.value = 1.2;
                    const gain = ctx.createGain();
                    gain.gain.setValueAtTime(0.01, now);
                    gain.gain.linearRampToValueAtTime(0.3, now + 0.2);
                    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.8);
                    noise.connect(filter);
                    filter.connect(gain);
                    gain.connect(ctx.destination);
                    noise.start(now);
                    toast('👏 Crowd Cheer!');
                } else if (type === 'scratch') {
                    // DJ Vinyl Scratch
                    const osc = ctx.createOscillator();
                    const gain = ctx.createGain();
                    osc.type = 'triangle';
                    osc.frequency.setValueAtTime(400, now);
                    osc.frequency.linearRampToValueAtTime(1800, now + 0.1);
                    osc.frequency.linearRampToValueAtTime(300, now + 0.22);
                    gain.gain.setValueAtTime(0.4, now);
                    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.25);
                    osc.connect(gain);
                    gain.connect(ctx.destination);
                    osc.start(now);
                    osc.stop(now + 0.25);
                    toast('🎚️ Vinyl Scratch!');
                }
            } catch(e) {
                console.warn('FX error:', e);
            }
        }

        // --- YOUTUBE IFRAME STREAM PLAYER ---
        let ytPlayer = null;
        window.onYouTubeIframeAPIReady = function() {
            ytPlayer = new YT.Player('yt-player-container', {
                height: '1',
                width: '1',
                playerVars: {
                    'autoplay': 0,
                    'controls': 0,
                    'disablekb': 1,
                    'fs': 0,
                    'playsinline': 1
                },
                events: {
                    'onReady': () => {
                        state.ytPlayerReady = true;
                    },
                    'onStateChange': onYTStateChange
                }
            });
        };

        function onYTStateChange(event) {
            if (event.data === YT.PlayerState.PLAYING) {
                state.isPlaying = true;
                state.isStreaming = true;
                updatePlayStateUI();
            } else if (event.data === YT.PlayerState.PAUSED) {
                state.isPlaying = false;
                updatePlayStateUI();
            } else if (event.data === YT.PlayerState.ENDED) {
                playNext();
            }
        }

        // --- INITIALIZE UI ---
        function initApp() {
            loadFavorites();
            detectPlaybackEnvironment();
            buildCeremonyCarousel();
            buildDesktopNavigation();
            buildDrawerNavigation();
            renderSongList();
            setupAudioEvents();
            setupKeyboardShortcuts();
            setupMobileInteractions();
        }

        function detectPlaybackEnvironment() {
            const isGitHubPages = window.location.hostname.includes('github.io');
            if (isGitHubPages) {
                statusPill.style.borderColor = 'rgba(6, 182, 212, 0.4)';
                statusPill.style.background = 'rgba(6, 182, 212, 0.15)';
                statusPill.style.color = 'var(--neon-cyan)';
                modeDot.style.color = 'var(--neon-cyan)';
                modeLabel.textContent = 'GitHub Pages Mode';
            } else {
                modeLabel.textContent = 'Local Offline Ready';
            }
        }

        // --- FOLDER CONNECT FOR MOBILE & DESKTOP ---
        document.getElementById('local-folder-picker').addEventListener('change', function(e) {
            const files = e.target.files;
            if (!files || files.length === 0) return;

            let count = 0;
            for (let f of files) {
                if (/\.(mp3|webm|m4a|opus|wav|ogg)$/i.test(f.name)) {
                    state.localFilesMap.set(f.name.toLowerCase(), f);
                    count++;
                }
            }

            if (count > 0) {
                state.isLocalMode = true;
                statusPill.style.borderColor = 'rgba(16, 185, 129, 0.4)';
                statusPill.style.background = 'rgba(16, 185, 129, 0.15)';
                statusPill.style.color = 'var(--neon-green)';
                modeDot.style.color = 'var(--neon-green)';
                modeLabel.textContent = `⚡ Offline (${count} Songs Connected)`;
                toast(`✅ Connected ${count} wedding songs! 100% offline audio active.`);
                closeDrawer();
            } else {
                toast('⚠️ No audio files found in selected folder.');
            }
        });

        document.getElementById('btn-connect-folder').onclick = () => {
            document.getElementById('local-folder-picker').click();
        };

        // --- NAVIGATION BUILDERS ---
        const CEREMONIES = [
            { id: 'all', name: '🌟 All Songs' },
            { id: 'Mayan & Mandap', name: '🪔 Mayan & Mandap' },
            { id: 'Haldi Ceremony', name: '💛 Haldi Ceremony' },
            { id: 'Punjabi Haldi', name: '🟡 Punjabi Haldi' },
            { id: 'Mehndi Ceremony', name: '🌿 Mehndi Ceremony' },
            { id: 'Traditional Sangeet', name: '🪘 Traditional Sangeet' },
            { id: 'Ladies Sangeet', name: '💃 Ladies Sangeet' },
            { id: 'Couple Dance', name: '💑 Couple Dance' },
            { id: 'Family Dance', name: '👨‍👩‍👧 Family Dance' },
            { id: 'Friends Dance', name: '🎉 Friends Dance' },
            { id: 'DJ Party Hits', name: '🔥 DJ Party Hits' },
            { id: 'Ghodi & Procession', name: '🐴 Ghodi & Procession' },
            { id: 'Baraat & Groom Entry', name: '🥁 Baraat Entry' },
            { id: 'Bride Entry', name: '👰 Bride Entry' },
            { id: 'Varmala & Jaimala', name: '🌸 Varmala' },
            { id: 'Wedding Rituals & Pheras', name: '🔥 Pheras & Rituals' },
            { id: 'Pheras Instrumental', name: '🎶 Shehnai & Pheras' },
            { id: 'Grand Reception', name: '🥂 Grand Reception' },
            { id: 'Vidaai', name: '😢 Vidaai' },
            { id: 'Family Celebration', name: '✨ Celebration' },
            { id: 'Trending Hits', name: '🚀 Trending Hits' },
            { id: 'Evergreen Classics', name: '📻 Classics' },
            { id: 'Bagheli', name: '👑 Rewa Bagheli' },
            { id: 'Bhojpuri', name: '🪗 Bhojpuri' },
            { id: 'fav', name: '❤️ Favourites' }
        ];

        function buildCeremonyCarousel() {
            const carousel = document.getElementById('cero-carousel');
            carousel.innerHTML = CEREMONIES.map(c => {
                let count = 0;
                if (c.id === 'all') count = SONGS.length;
                else if (c.id === 'fav') count = state.favorites.size;
                else if (c.id === 'Bagheli') count = SONGS.filter(s => s.language === 'Bagheli').length;
                else if (c.id === 'Bhojpuri') count = SONGS.filter(s => s.language === 'Bhojpuri').length;
                else count = SONGS.filter(s => s.function.toLowerCase() === c.id.toLowerCase()).length;

                return `<div class="cero-chip ${state.activeCero === c.id ? 'active' : ''}" onclick="selectCeremony('${c.id}')">
                    ${c.name} <span style="opacity:0.7; font-size:10px;">(${count})</span>
                </div>`;
            }).join('');
        }

        const DAYS = [
            { key: 'all', name: '🌟 All Collections', count: SONGS.length },
            { key: 'day1', name: '🪔 10 Dec: Mayan & Haldi', count: SONGS.filter(s => s.day_key === 'day1').length },
            { key: 'day2', name: '💃 11 Dec: Mehndi & Sangeet', count: SONGS.filter(s => s.day_key === 'day2').length },
            { key: 'day3', name: '👑 12 Dec: Wedding Day', count: SONGS.filter(s => s.day_key === 'day3').length },
            { key: 'bagheli', name: '👑 Bagheli (Rewa) Songs', count: SONGS.filter(s => s.day_key === 'bagheli').length },
            { key: 'bhojpuri', name: '🪘 Bhojpuri Songs', count: SONGS.filter(s => s.day_key === 'bhojpuri').length }
        ];

        const LANGUAGES = ['All', 'Hindi & Bollywood', 'Punjabi', 'Bagheli', 'Bhojpuri'];

        function buildDesktopNavigation() {
            const dayNav = document.getElementById('desktop-day-nav');
            dayNav.innerHTML = DAYS.map(d => `
                <div class="nav-link ${state.activeDay === d.key ? 'active' : ''}" onclick="selectDay('${d.key}')">
                    <span>${d.name}</span>
                    <span class="tag">${d.count}</span>
                </div>
            `).join('');

            const langNav = document.getElementById('desktop-lang-nav');
            langNav.innerHTML = LANGUAGES.map(l => `
                <div class="lang-pill ${state.activeLang === l.toLowerCase() ? 'active' : ''}" onclick="selectLang('${l.toLowerCase()}')">
                    ${l}
                </div>
            `).join('');

            const ceroNav = document.getElementById('desktop-cero-nav');
            ceroNav.innerHTML = CEREMONIES.slice(1, -1).map(c => {
                const count = SONGS.filter(s => s.function.toLowerCase() === c.id.toLowerCase()).length;
                return `
                    <div class="nav-link ${state.activeCero === c.id ? 'active' : ''}" onclick="selectCeremony('${c.id}')" style="font-size:12px; padding:6px 10px;">
                        <span>${c.name}</span>
                        <span class="tag">${count}</span>
                    </div>
                `;
            }).join('');
        }

        function buildDrawerNavigation() {
            const drawerDays = document.getElementById('drawer-days-list');
            drawerDays.innerHTML = DAYS.map(d => `
                <div class="nav-link ${state.activeDay === d.key ? 'active' : ''}" onclick="selectDay('${d.key}'); closeDrawer();">
                    <span>${d.name}</span>
                    <span class="tag">${d.count}</span>
                </div>
            `).join('');

            const drawerLangs = document.getElementById('drawer-lang-pills');
            drawerLangs.innerHTML = LANGUAGES.map(l => `
                <div class="lang-pill ${state.activeLang === l.toLowerCase() ? 'active' : ''}" onclick="selectLang('${l.toLowerCase()}'); closeDrawer();">
                    ${l}
                </div>
            `).join('');

            const drawerCeros = document.getElementById('drawer-cero-list');
            drawerCeros.innerHTML = CEREMONIES.map(c => `
                <div class="nav-link ${state.activeCero === c.id ? 'active' : ''}" onclick="selectCeremony('${c.id}'); closeDrawer();" style="font-size:12px; padding:6px 10px;">
                    <span>${c.name}</span>
                </div>
            `).join('');
        }

        // --- FILTERING & RENDERING ---
        function getFilteredSongs() {
            return SONGS.filter(s => {
                // Day filter
                if (state.activeDay !== 'all' && s.day_key !== state.activeDay) return false;

                // Ceremony filter
                if (state.activeCero === 'fav') {
                    if (!state.favorites.has(s.id)) return false;
                } else if (state.activeCero === 'Bagheli') {
                    if (s.language !== 'Bagheli') return false;
                } else if (state.activeCero === 'Bhojpuri') {
                    if (s.language !== 'Bhojpuri') return false;
                } else if (state.activeCero !== 'all') {
                    if (s.function.toLowerCase() !== state.activeCero.toLowerCase()) return false;
                }

                // Language filter
                if (state.activeLang !== 'all') {
                    if (state.activeLang === 'hindi & bollywood' && !s.language.toLowerCase().includes('hindi')) return false;
                    else if (state.activeLang === 'punjabi' && s.language.toLowerCase() !== 'punjabi') return false;
                    else if (state.activeLang === 'bagheli' && s.language.toLowerCase() !== 'bagheli') return false;
                    else if (state.activeLang === 'bhojpuri' && s.language.toLowerCase() !== 'bhojpuri') return false;
                }

                // Search query
                if (state.searchQuery) {
                    const q = state.searchQuery.toLowerCase();
                    const text = `${s.title} ${s.artist} ${s.movie} ${s.function} ${s.language} ${s.category}`.toLowerCase();
                    if (!text.includes(q)) return false;
                }

                return true;
            });
        }

        function renderSongList() {
            const filtered = getFilteredSongs();

            // Update banner
            const stageTitle = document.getElementById('stage-title');
            const stageMeta = document.getElementById('stage-meta');

            let titleStr = 'All Wedding Tracks';
            if (state.activeCero === 'fav') titleStr = '❤️ Favourite Wedding Songs';
            else if (state.activeCero !== 'all') titleStr = state.activeCero;
            else if (state.activeDay !== 'all') {
                const found = DAYS.find(d => d.key === state.activeDay);
                if (found) titleStr = found.name;
            }
            stageTitle.textContent = titleStr;
            stageMeta.textContent = `Showing ${filtered.length} Songs • Instant Play & DJ Control`;

            // Mobile Card List (<768px)
            const mobileList = document.getElementById('mobile-song-list');
            if (filtered.length === 0) {
                mobileList.innerHTML = `<div style="text-align:center; padding:40px 20px; color:var(--text-muted);">
                    <div style="font-size:36px; margin-bottom:10px;">🔍</div>
                    <div style="font-size:15px; font-weight:700; color:#fff;">No Songs Found</div>
                    <div style="font-size:12px; margin-top:4px;">Try clearing filters or search keywords.</div>
                    <button class="btn-gold" style="margin-top:16px;" onclick="resetFilters()">Reset All Filters</button>
                </div>`;
            } else {
                mobileList.innerHTML = filtered.map((s, idx) => {
                    const isCur = (state.currentIndex === s.id - 1);
                    const isPlayingCur = isCur && state.isPlaying;
                    const isFav = state.favorites.has(s.id);

                    return `
                        <div class="m-song-card ${isPlayingCur ? 'playing' : ''}" id="m-card-${s.id}">
                            <div class="m-card-idx" onclick="playSongById(${s.id})">
                                ${isPlayingCur ? `
                                    <div class="eq-bars">
                                        <div class="eq-bar"></div><div class="eq-bar"></div><div class="eq-bar"></div>
                                    </div>
                                ` : (idx + 1)}
                            </div>
                            <div class="m-card-info" onclick="playSongById(${s.id})">
                                <div class="m-card-title">${s.title}</div>
                                <div class="m-card-sub">${s.artist}${s.movie ? ' • ' + s.movie : ''}</div>
                                <div class="m-card-badges">
                                    <span class="badge-chip cero">${s.function}</span>
                                    <span class="badge-chip">${s.language}</span>
                                </div>
                            </div>
                            <div class="m-card-actions">
                                <button class="btn-card-act ${isFav ? 'is-fav' : ''}" onclick="toggleFav(${s.id})" title="Favorite">${isFav ? '❤️' : '🤍'}</button>
                                <button class="btn-card-act" onclick="addToQueue(${s.id})" title="Add to Queue">➕</button>
                                <button class="btn-play-card" onclick="playSongById(${s.id})" title="Play/Pause">
                                    ${isPlayingCur ? '❚❚' : '▶'}
                                </button>
                            </div>
                        </div>
                    `;
                }).join('');
            }

            // Desktop Cue Table (>768px)
            const deskBody = document.getElementById('desktop-table-body');
            deskBody.innerHTML = filtered.map((s, idx) => {
                const isCur = (state.currentIndex === s.id - 1);
                const isPlayingCur = isCur && state.isPlaying;
                const isFav = state.favorites.has(s.id);

                return `
                    <tr class="${isPlayingCur ? 'playing' : ''}" id="desk-row-${s.id}">
                        <td>
                            ${isPlayingCur ? '<span style="color:var(--gold);">🔊</span>' : (idx + 1)}
                        </td>
                        <td onclick="playSongById(${s.id})" style="cursor:pointer;">
                            <div class="t-title" style="font-weight:600; color:#fff;">${s.title}</div>
                            <div style="font-size:11px; color:var(--text-muted);">${s.movie || 'Wedding Anthem'}</div>
                        </td>
                        <td>${s.artist}</td>
                        <td><span class="badge-chip cero">${s.function}</span></td>
                        <td><span class="badge-chip">${s.language}</span></td>
                        <td style="font-size:11px; color:var(--text-muted);">${s.size_mb} MB</td>
                        <td style="text-align: right;">
                            <div style="display:inline-flex; gap:6px;">
                                <button class="btn-card-act ${isFav ? 'is-fav' : ''}" onclick="toggleFav(${s.id})">${isFav ? '❤️' : '🤍'}</button>
                                <button class="btn-card-act" onclick="addToQueue(${s.id})" title="Queue">➕</button>
                                <button class="btn-gold" style="padding:4px 10px; font-size:11px;" onclick="playSongById(${s.id})">${isPlayingCur ? '❚❚ Pause' : '▶ Play'}</button>
                            </div>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        // --- PLAYBACK CORE ENGINE ---
        function playSongById(songId) {
            const idx = SONGS.findIndex(s => s.id === songId);
            if (idx === -1) return;

            if (state.currentIndex === idx) {
                togglePlayPause();
                return;
            }

            playByIndex(idx, true);
        }

        function playByIndex(idx, startPlay = true) {
            if (idx < 0 || idx >= SONGS.length) return;
            state.currentIndex = idx;
            const song = SONGS[idx];

            // Resolve Audio Source
            const filenameLower = song.filename.toLowerCase();
            let sourceResolved = false;

            // 1. Check if user connected local files via Folder Connect
            if (state.localFilesMap.has(filenameLower)) {
                const fileObj = state.localFilesMap.get(filenameLower);
                audio.src = URL.createObjectURL(fileObj);
                sourceResolved = true;
                state.isStreaming = false;
            } else {
                // 2. Relative file path (local server or root)
                const inSubfolder = window.location.pathname.includes('Brother%20Wedding%20Music') || window.location.pathname.includes('Brother Wedding Music');
                const targetPath = inSubfolder ? song.sub_path : song.root_path;
                audio.src = encodeURI(targetPath);
            }

            // Update UI Details
            miniTitle.textContent = song.title;
            miniSub.textContent = `${song.artist} • ${song.function}`;
            sheetTitle.textContent = song.title;
            sheetArtist.textContent = `${song.artist}${song.movie ? ' • ' + song.movie : ''}`;
            sheetCeroTag.textContent = song.function;
            sheetLangTag.textContent = song.language;
            sheetHeaderCero.textContent = song.function.toUpperCase();
            sheetFavBtn.textContent = state.favorites.has(song.id) ? '❤️' : '🤍';
            sheetFavBtn.classList.toggle('is-fav', state.favorites.has(song.id));

            // Stop any active YouTube stream before starting native audio
            if (ytPlayer && typeof ytPlayer.stopVideo === 'function') {
                try { ytPlayer.stopVideo(); } catch(e) {}
            }

            audio.load();

            if (startPlay) {
                const playPromise = audio.play();
                if (playPromise !== undefined) {
                    playPromise.then(() => {
                        state.isPlaying = true;
                        state.isStreaming = false;
                        updatePlayStateUI();
                    }).catch(err => {
                        console.warn('HTML5 audio play blocked or not found:', err);
                        // On GitHub Pages or missing file, handle fallback
                        handlePlaybackFallback(song);
                    });
                }
            } else {
                state.isPlaying = false;
                updatePlayStateUI();
            }

            renderSongList();
        }

        // Fallback for GitHub Pages or unhosted files
        function handlePlaybackFallback(song) {
            const isGitHubPages = window.location.hostname.includes('github.io');
            if (isGitHubPages && song.yt_id && ytPlayer && state.ytPlayerReady) {
                toast(`🌐 Streaming "${song.title}" in-app via YouTube...`);
                try {
                    ytPlayer.loadVideoById(song.yt_id);
                    state.isPlaying = true;
                    state.isStreaming = true;
                    updatePlayStateUI();
                } catch(e) {
                    fallbackToExternal(song);
                }
            } else {
                fallbackToExternal(song);
            }
        }

        function fallbackToExternal(song) {
            toast(`⚠️ Local file not on device. Tip: Tap 📁 Connect Audio Folder, or tap Stream.`);
            state.isPlaying = false;
            updatePlayStateUI();
        }

        function togglePlayPause() {
            if (state.currentIndex === -1) {
                const filtered = getFilteredSongs();
                if (filtered.length > 0) playSongById(filtered[0].id);
                return;
            }

            if (state.isStreaming && ytPlayer && state.ytPlayerReady) {
                if (state.isPlaying) {
                    ytPlayer.pauseVideo();
                    state.isPlaying = false;
                } else {
                    ytPlayer.playVideo();
                    state.isPlaying = true;
                }
                updatePlayStateUI();
                return;
            }

            if (audio.paused) {
                audio.play().then(() => {
                    state.isPlaying = true;
                    updatePlayStateUI();
                }).catch(() => {
                    handlePlaybackFallback(SONGS[state.currentIndex]);
                });
            } else {
                audio.pause();
                state.isPlaying = false;
                updatePlayStateUI();
            }
        }

        function playNext() {
            // 1. Check queue first
            if (state.queue.length > 0) {
                const nextId = state.queue.shift();
                updateQueueUI();
                playSongById(nextId);
                return;
            }

            // 2. Play next in active filtered list
            const filtered = getFilteredSongs();
            if (filtered.length === 0) return;

            if (state.isShuffle) {
                const randIdx = Math.floor(Math.random() * filtered.length);
                playSongById(filtered[randIdx].id);
                return;
            }

            const curSong = SONGS[state.currentIndex];
            const curPos = filtered.findIndex(s => s.id === (curSong ? curSong.id : -1));
            let nextPos = curPos + 1;
            if (nextPos >= filtered.length) {
                if (state.repeatMode === 'all') nextPos = 0;
                else return; // Stop at end of list
            }
            playSongById(filtered[nextPos].id);
        }

        function playPrev() {
            if (audio.currentTime > 3) {
                audio.currentTime = 0;
                return;
            }

            const filtered = getFilteredSongs();
            if (filtered.length === 0) return;

            const curSong = SONGS[state.currentIndex];
            const curPos = filtered.findIndex(s => s.id === (curSong ? curSong.id : -1));
            let prevPos = curPos - 1;
            if (prevPos < 0) prevPos = filtered.length - 1;
            playSongById(filtered[prevPos].id);
        }

        function playAllFiltered(shuffle = false) {
            const filtered = getFilteredSongs();
            if (filtered.length === 0) {
                toast('No songs match current filter.');
                return;
            }
            state.isShuffle = shuffle;
            updatePlayStateUI();

            const startIdx = shuffle ? Math.floor(Math.random() * filtered.length) : 0;
            playSongById(filtered[startIdx].id);
            toast(`▶ Playing ${filtered.length} songs ${shuffle ? '(Shuffled)' : ''}`);
        }

        function updatePlayStateUI() {
            const playText = state.isPlaying ? '❚❚' : '▶';

            // Buttons
            miniPlayBtn.textContent = playText;
            sheetPlayBtn.textContent = playText;
            deskPlayBtn.textContent = playText;

            // Vinyl spinning
            miniVinyl.classList.toggle('spinning', state.isPlaying);
            sheetVinyl.classList.toggle('spinning', state.isPlaying);

            // Shuffles & Repeat
            sheetShuffleBtn.classList.toggle('active', state.isShuffle);
            deskShuffleBtn.classList.toggle('active', state.isShuffle);

            sheetRepeatBtn.textContent = state.repeatMode === 'one' ? '🔂' : '🔁';
            sheetRepeatBtn.classList.toggle('active', state.repeatMode !== 'off');
            deskRepeatBtn.textContent = state.repeatMode === 'one' ? '🔂' : '🔁';
            deskRepeatBtn.classList.toggle('active', state.repeatMode !== 'off');

            renderSongList();
        }

        // --- AUDIO EVENT LISTENERS ---
        function setupAudioEvents() {
            audio.addEventListener('timeupdate', () => {
                if (!audio.duration || isNaN(audio.duration)) return;
                const cur = audio.currentTime;
                const dur = audio.duration;
                const pct = (cur / dur) * 100;

                miniScrubFill.style.width = `${pct}%`;
                sheetSlider.value = pct;
                deskSlider.value = pct;

                const curFormatted = formatTime(cur);
                const durFormatted = formatTime(dur);

                sheetTimeCur.textContent = curFormatted;
                sheetTimeDur.textContent = durFormatted;
                deskTimeCur.textContent = curFormatted;
                deskTimeDur.textContent = durFormatted;
            });

            audio.addEventListener('ended', () => {
                if (state.repeatMode === 'one') {
                    audio.currentTime = 0;
                    audio.play();
                } else {
                    playNext();
                }
            });

            audio.addEventListener('error', (e) => {
                if (state.currentIndex >= 0) {
                    handlePlaybackFallback(SONGS[state.currentIndex]);
                }
            });

            // Sliders Scrubbing
            sheetSlider.addEventListener('input', (e) => {
                if (audio.duration) {
                    audio.currentTime = (e.target.value / 100) * audio.duration;
                }
            });
            deskSlider.addEventListener('input', (e) => {
                if (audio.duration) {
                    audio.currentTime = (e.target.value / 100) * audio.duration;
                }
            });

            // Mini player scrubber bar tap to seek
            miniScrubTrack.addEventListener('click', (e) => {
                if (!audio.duration) return;
                const rect = miniScrubTrack.getBoundingClientRect();
                const clickX = e.clientX - rect.left;
                const pct = Math.max(0, Math.min(1, clickX / rect.width));
                audio.currentTime = pct * audio.duration;
            });

            // Volume (Desktop)
            deskVolSlider.addEventListener('input', (e) => {
                try {
                    audio.volume = parseFloat(e.target.value);
                } catch(err) {
                    console.warn('iOS physical volume only:', err);
                }
            });
        }

        function formatTime(sec) {
            const m = Math.floor(sec / 60);
            const s = Math.floor(sec % 60);
            return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
        }

        // --- FAVORITES & QUEUE ---
        function loadFavorites() {
            try {
                const saved = localStorage.getItem('wedding_dj_favs');
                if (saved) state.favorites = new Set(JSON.parse(saved));
            } catch(e) {}
        }

        function saveFavorites() {
            try {
                localStorage.setItem('wedding_dj_favs', JSON.stringify([...state.favorites]));
            } catch(e) {}
        }

        function toggleFav(songId) {
            if (state.favorites.has(songId)) {
                state.favorites.delete(songId);
                toast('Removed from Favourites');
            } else {
                state.favorites.add(songId);
                toast('❤️ Added to Favourites');
            }
            saveFavorites();
            buildCeremonyCarousel();
            renderSongList();
        }

        function toggleFavCurrent() {
            if (state.currentIndex >= 0) {
                toggleFav(SONGS[state.currentIndex].id);
                sheetFavBtn.textContent = state.favorites.has(SONGS[state.currentIndex].id) ? '❤️' : '🤍';
                sheetFavBtn.classList.toggle('is-fav', state.favorites.has(SONGS[state.currentIndex].id));
            }
        }

        function addToQueue(songId) {
            state.queue.push(songId);
            updateQueueUI();
            const s = SONGS.find(x => x.id === songId);
            toast(`➕ Queued: "${s.title}"`);
        }

        function addCurrentToQueue() {
            if (state.currentIndex >= 0) {
                addToQueue(SONGS[state.currentIndex].id);
            }
        }

        function updateQueueUI() {
            queueBadge.textContent = state.queue.length;
            queueDisplay.textContent = state.queue.length;

            if (state.queue.length === 0) {
                queueItemsBox.innerHTML = '<div style="color:var(--text-muted); text-align:center; padding:20px; font-size:12px;">Queue is empty. Tap ➕ on any song to add.</div>';
                return;
            }

            queueItemsBox.innerHTML = state.queue.map((sid, idx) => {
                const s = SONGS.find(x => x.id === sid);
                if (!s) return '';
                return `
                    <div style="display:flex; align-items:center; justify-content:space-between; background:var(--bg-card); padding:8px 12px; border-radius:8px; border:1px solid var(--border);">
                        <div style="display:flex; align-items:center; gap:8px; min-width:0;">
                            <span style="font-size:11px; color:var(--gold); font-weight:700;">#${idx+1}</span>
                            <div style="min-width:0;">
                                <div style="font-size:13px; font-weight:700; color:#fff; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${s.title}</div>
                                <div style="font-size:11px; color:var(--text-muted);">${s.artist}</div>
                            </div>
                        </div>
                        <button class="btn-card-act" onclick="removeQueueIndex(${idx})" style="width:26px; height:26px;">✕</button>
                    </div>
                `;
            }).join('');
        }

        function removeQueueIndex(idx) {
            state.queue.splice(idx, 1);
            updateQueueUI();
        }

        function clearQueue() {
            state.queue = [];
            updateQueueUI();
            toast('Queue cleared');
        }

        // --- FILTER HELPERS ---
        function selectCeremony(ceroId) {
            state.activeCero = ceroId;
            buildCeremonyCarousel();
            buildDesktopNavigation();
            renderSongList();
        }

        function selectDay(dayKey) {
            state.activeDay = dayKey;
            buildDesktopNavigation();
            buildDrawerNavigation();
            renderSongList();
        }

        function selectLang(lang) {
            state.activeLang = lang;
            buildDesktopNavigation();
            buildDrawerNavigation();
            renderSongList();
        }

        function resetFilters() {
            state.activeDay = 'all';
            state.activeCero = 'all';
            state.activeLang = 'all';
            state.searchQuery = '';
            document.getElementById('main-search-input').value = '';
            document.getElementById('search-clear-btn').style.display = 'none';
            buildCeremonyCarousel();
            buildDesktopNavigation();
            buildDrawerNavigation();
            renderSongList();
            toast('All filters reset');
        }

        function toggleShuffle() {
            state.isShuffle = !state.isShuffle;
            toast(state.isShuffle ? '🔀 Shuffle On' : '➡️ Shuffle Off');
            updatePlayStateUI();
        }

        function toggleRepeat() {
            if (state.repeatMode === 'all') state.repeatMode = 'one';
            else if (state.repeatMode === 'one') state.repeatMode = 'off';
            else state.repeatMode = 'all';

            const msgs = { 'all': '🔁 Repeat All', 'one': '🔂 Repeat One', 'off': '➡️ Repeat Off' };
            toast(msgs[state.repeatMode]);
            updatePlayStateUI();
        }

        function openCurrentStream() {
            if (state.currentIndex >= 0) {
                const s = SONGS[state.currentIndex];
                window.open(s.stream_url, '_blank');
            }
        }

        // --- SEARCH INPUT ---
        const searchInput = document.getElementById('main-search-input');
        const searchClear = document.getElementById('search-clear-btn');

        searchInput.addEventListener('input', (e) => {
            state.searchQuery = e.target.value.trim();
            searchClear.style.display = state.searchQuery ? 'block' : 'none';
            renderSongList();
        });

        searchClear.addEventListener('click', () => {
            searchInput.value = '';
            state.searchQuery = '';
            searchClear.style.display = 'none';
            renderSongList();
        });

        // --- MOBILE DRAWER & SHEET INTERACTIONS ---
        function openDrawer() {
            drawer.classList.add('open');
            drawerOverlay.classList.add('open');
        }

        function closeDrawer() {
            drawer.classList.remove('open');
            drawerOverlay.classList.remove('open');
        }

        function openSheet() {
            sheet.classList.add('open');
        }

        function closeSheet() {
            sheet.classList.remove('open');
        }

        function openQueue() {
            updateQueueUI();
            queueSheet.classList.add('open');
            queueOverlay.classList.add('open');
        }

        function closeQueue() {
            queueSheet.classList.remove('open');
            queueOverlay.classList.remove('open');
        }

        function setupMobileInteractions() {
            document.getElementById('btn-toggle-drawer').onclick = openDrawer;
            document.getElementById('mini-expand-trigger').onclick = openSheet;
            document.getElementById('btn-open-queue').onclick = openQueue;
        }

        // --- KEYBOARD SHORTCUTS (DESKTOP) ---
        function setupKeyboardShortcuts() {
            window.addEventListener('keydown', (e) => {
                if (e.target.tagName === 'INPUT') return;

                if (e.code === 'Space') {
                    e.preventDefault();
                    togglePlayPause();
                } else if (e.code === 'ArrowRight') {
                    e.preventDefault();
                    playNext();
                } else if (e.code === 'ArrowLeft') {
                    e.preventDefault();
                    playPrev();
                } else if (e.code === 'ArrowUp') {
                    e.preventDefault();
                    audio.volume = Math.min(1, audio.volume + 0.1);
                    deskVolSlider.value = audio.volume;
                } else if (e.code === 'ArrowDown') {
                    e.preventDefault();
                    audio.volume = Math.max(0, audio.volume - 0.1);
                    deskVolSlider.value = audio.volume;
                } else if (e.key === '1') playFX('horn');
                else if (e.key === '2') playFX('dhol');
                else if (e.key === '3') playFX('cheer');
                else if (e.key === '4') playFX('scratch');
            });
        }

        // --- TOAST NOTIFICATIONS ---
        let toastTimer = null;
        function toast(msg) {
            const el = document.getElementById('toast-notify');
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => {
                el.classList.remove('show');
            }, 2600);
        }

        // Start App
        window.addEventListener('DOMContentLoaded', initApp);
    </script>
</body>
</html>
"""

# Replace __SONGS_DATA__ placeholder with actual songs json
final_html = html_template.replace("__SONGS_DATA__", songs_json)

# Save to Repository Root (For GitHub Pages)
root_index = os.path.join(REPO_ROOT, "index.html")
with open(root_index, "w", encoding="utf-8") as f:
    f.write(final_html)
print(f"✅ Created root index.html: {root_index} ({len(final_html)} bytes)")

# Save to Brother Wedding Music/index.html and Wedding_Music_Player.html
music_index = os.path.join(MUSIC_DIR, "index.html")
with open(music_index, "w", encoding="utf-8") as f:
    f.write(final_html)
print(f"✅ Updated {music_index}")

music_player = os.path.join(MUSIC_DIR, "Wedding_Music_Player.html")
with open(music_player, "w", encoding="utf-8") as f:
    f.write(final_html)
print(f"✅ Updated {music_player}")
