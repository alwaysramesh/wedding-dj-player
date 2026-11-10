import os
import json
import re
import openpyxl

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

# Load Excel metadata
wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
meta = {}
for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [str(c.value or '').lower() for c in ws[1]]
    song_col, movie_col, cat_col, year_col, cero_col, day_col, singer_col = None, None, None, None, None, None, None
    for idx, h in enumerate(headers, 1):
        if 'song' in h: song_col = idx
        elif 'singer' in h or 'artist' in h: singer_col = idx
        elif 'movie' in h or 'album' in h or 'source' in h: movie_col = idx
        elif 'category' in h or 'type' in h: cat_col = idx
        elif 'year' in h or 'era' in h: year_col = idx
        elif 'ceremony' in h or 'function' in h: cero_col = idx
        elif 'day' in h: day_col = idx
        
    for r in range(2, ws.max_row+1):
        name = str(ws.cell(row=r, column=song_col).value or '').strip() if song_col else ''
        if not name or name.lower() == 'none': continue
        singer = str(ws.cell(row=r, column=singer_col).value or '').strip() if singer_col else ''
        movie = str(ws.cell(row=r, column=movie_col).value or '').strip() if movie_col else ''
        cat = str(ws.cell(row=r, column=cat_col).value or '').strip() if cat_col else ''
        year = str(ws.cell(row=r, column=year_col).value or '').strip() if year_col else ''
        cero = str(ws.cell(row=r, column=cero_col).value or '').strip() if cero_col else ''
        day = str(ws.cell(row=r, column=day_col).value or '').strip() if day_col else ''
        k = re.sub(r'[^a-zA-Z0-9]', '', name.lower())
        if k not in meta:
            meta[k] = {'singer': singer, 'movie': movie, 'category': cat, 'year': year, 'ceremony': cero, 'day': day}

punjabi_singers = [
    'diljit dosanjh', 'sukhbir', 'panjabi mc', 'mika singh', 'guru randhawa',
    'daler mehndi', 'jasbir jassi', 'bally sagoo', 'honey singh', 'badshah',
    'ap dhillon', 'karan aujla', 'hardy sandhu', 'amrit maan', 'sidhu moose wala',
    'gippy grewal', 'surinder kaur', 'gurmeet bawa', 'jasleen royal'
]

songs = []
for root, dirs, files in os.walk(BASE_DIR):
    for f in sorted(files):
        if any(f.lower().endswith(ext) for ext in ['.mp3', '.webm', '.m4a', '.opus']):
            full = os.path.join(root, f)
            rel = os.path.relpath(full, BASE_DIR)
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
                
            # Day determination (check bagheli/bhojpuri first!)
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
                
            # Function / Ceremony clean naming
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
            
            # Clean relative web path
            web_path = rel.replace('\\', '/')
            
            songs.append({
                'id': len(songs) + 1,
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
                'path': web_path,
                'size_mb': round(os.path.getsize(full) / (1024*1024), 2)
            })

songs_json = json.dumps(songs, ensure_ascii=False)

html_code = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Wedding DJ Studio Console • Ramesh Brother's Wedding</title>
    <style>
        :root {
            --bg-body: #080a10;
            --bg-sidebar: #0e111a;
            --bg-card: #141824;
            --bg-card-hover: #1c2234;
            --bg-player: rgba(12, 15, 24, 0.96);
            --gold: #f59e0b;
            --gold-light: #fbbf24;
            --gold-glow: rgba(245, 158, 11, 0.35);
            --text-main: #f8fafc;
            --text-muted: #8e9bb0;
            --border: rgba(255, 255, 255, 0.08);
            --border-gold: rgba(245, 158, 11, 0.28);
            --neon-red: #f43f5e;
            --neon-green: #10b981;
            --neon-cyan: #06b6d4;
            --radius-sm: 6px;
            --radius-md: 10px;
            --radius-lg: 14px;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            user-select: none;
            -webkit-user-select: none;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: var(--bg-body);
            color: var(--text-main);
            height: 100vh;
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }

        /* Top Bar */
        .top-navbar {
            height: 58px;
            background: var(--bg-sidebar);
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 20px;
            flex-shrink: 0;
            z-index: 50;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 12px;
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
            font-weight: 800;
            box-shadow: 0 0 14px var(--gold-glow);
        }

        .brand-title {
            font-size: 1rem;
            font-weight: 800;
            letter-spacing: 0.5px;
            color: #fff;
        }

        .brand-subtitle {
            font-size: 0.72rem;
            color: var(--gold-light);
            font-weight: 600;
        }

        .top-search-box {
            flex: 1;
            max-width: 480px;
            margin: 0 20px;
            position: relative;
        }

        .top-search-input {
            width: 100%;
            background: var(--bg-card);
            border: 1px solid var(--border);
            color: #fff;
            padding: 9px 16px 9px 36px;
            border-radius: var(--radius-md);
            font-size: 0.86rem;
            outline: none;
            transition: all 0.2s;
        }

        .top-search-input:focus {
            border-color: var(--gold);
            box-shadow: 0 0 12px var(--gold-glow);
            background: #192030;
        }

        .search-icon {
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            font-size: 14px;
        }

        .top-actions {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .btn-top {
            background: var(--bg-card);
            border: 1px solid var(--border);
            color: var(--text-main);
            padding: 7px 12px;
            border-radius: var(--radius-sm);
            font-size: 0.78rem;
            font-weight: 700;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.15s;
        }

        .btn-top:hover {
            border-color: var(--gold);
            color: var(--gold-light);
        }

        /* Workspace Grid (Sidebar + Main Stage) */
        .workspace {
            display: flex;
            flex: 1;
            height: calc(100vh - 58px - 94px);
            overflow: hidden;
        }

        /* Left Deck Sidebar */
        .sidebar {
            width: 290px;
            background: var(--bg-sidebar);
            border-right: 1px solid var(--border);
            display: flex;
            flex-direction: column;
            overflow-y: auto;
            flex-shrink: 0;
        }

        .sidebar-section {
            padding: 14px 16px;
            border-bottom: 1px solid var(--border);
        }

        .sidebar-label {
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            font-weight: 800;
            color: var(--text-muted);
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .nav-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 9px 12px;
            border-radius: var(--radius-sm);
            color: var(--text-muted);
            font-size: 0.84rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s;
            margin-bottom: 3px;
        }

        .nav-item:hover {
            background: var(--bg-card);
            color: #fff;
        }

        .nav-item.active {
            background: rgba(245, 158, 11, 0.14);
            border-left: 3px solid var(--gold);
            color: var(--gold-light);
            font-weight: 700;
        }

        .badge-count {
            font-size: 0.7rem;
            background: rgba(255, 255, 255, 0.08);
            padding: 2px 7px;
            border-radius: 10px;
            color: inherit;
        }

        /* Language Filter Chips in Sidebar */
        .lang-chips {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
        }

        .chip {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border);
            color: var(--text-muted);
            padding: 5px 10px;
            border-radius: var(--radius-sm);
            font-size: 0.76rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s;
        }

        .chip:hover {
            border-color: var(--gold);
            color: #fff;
        }

        .chip.active {
            background: var(--gold);
            color: #000;
            font-weight: 800;
            border-color: var(--gold);
        }

        /* Center Main Stage */
        .stage {
            flex: 1;
            display: flex;
            flex-direction: column;
            overflow-y: auto;
            background: var(--bg-body);
            position: relative;
        }

        .stage-header {
            padding: 20px 24px;
            background: linear-gradient(180deg, #101420 0%, transparent 100%);
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
        }

        .stage-header-title {
            font-size: 1.4rem;
            font-weight: 900;
            color: #fff;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .stage-header-meta {
            font-size: 0.8rem;
            color: var(--gold-light);
            margin-top: 4px;
            font-weight: 600;
        }

        .stage-actions {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .btn-play-all {
            background: linear-gradient(135deg, #b45309, #f59e0b);
            border: none;
            color: #000;
            padding: 10px 20px;
            border-radius: var(--radius-md);
            font-size: 0.88rem;
            font-weight: 800;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 0 16px var(--gold-glow);
            transition: all 0.2s;
        }

        .btn-play-all:hover {
            transform: scale(1.03);
            box-shadow: 0 0 24px var(--gold-glow);
        }

        .btn-secondary {
            background: var(--bg-card);
            border: 1px solid var(--border);
            color: var(--text-main);
            padding: 10px 16px;
            border-radius: var(--radius-md);
            font-size: 0.86rem;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.15s;
        }

        .btn-secondary:hover {
            border-color: var(--gold);
            color: var(--gold-light);
        }

        /* Song Cue Table */
        .cue-table {
            width: 100%;
            border-collapse: collapse;
        }

        .cue-table th {
            text-align: left;
            padding: 10px 20px;
            font-size: 0.72rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: var(--text-muted);
            border-bottom: 1px solid var(--border);
            background: rgba(8, 10, 16, 0.75);
            position: sticky;
            top: 0;
            z-index: 10;
        }

        .cue-row {
            border-bottom: 1px solid rgba(255, 255, 255, 0.03);
            transition: background 0.15s;
            cursor: pointer;
        }

        .cue-row:hover {
            background: var(--bg-card-hover);
        }

        .cue-row.playing {
            background: rgba(245, 158, 11, 0.09);
            border-left: 3px solid var(--gold);
        }

        .cue-row td {
            padding: 11px 20px;
            font-size: 0.88rem;
            vertical-align: middle;
        }

        .cue-index {
            width: 45px;
            color: var(--text-muted);
            font-size: 0.8rem;
            font-weight: 700;
            text-align: center;
        }

        .song-title-col {
            font-weight: 700;
            color: #fff;
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .cue-row.playing .song-title-col .main-name {
            color: var(--gold-light);
        }

        .song-artist-sub {
            font-size: 0.76rem;
            color: var(--text-muted);
            font-weight: normal;
        }

        .tag-pill {
            display: inline-block;
            padding: 3px 8px;
            border-radius: var(--radius-sm);
            font-size: 0.72rem;
            font-weight: 700;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(255, 255, 255, 0.08);
            color: var(--text-muted);
        }

        .tag-pill.lang-hindi { color: #f8fafc; border-color: rgba(255, 255, 255, 0.15); }
        .tag-pill.lang-punjabi { color: #38bdf8; border-color: rgba(56, 189, 248, 0.25); background: rgba(56, 189, 248, 0.08); }
        .tag-pill.lang-bagheli { color: #fbbf24; border-color: rgba(251, 191, 36, 0.25); background: rgba(251, 191, 36, 0.08); }
        .tag-pill.lang-bhojpuri { color: #f43f5e; border-color: rgba(244, 63, 94, 0.25); background: rgba(244, 63, 94, 0.08); }

        .row-btns {
            display: flex;
            align-items: center;
            gap: 6px;
            justify-content: flex-end;
        }

        .btn-act {
            width: 32px;
            height: 32px;
            border-radius: var(--radius-sm);
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border);
            color: var(--text-muted);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.15s;
        }

        .btn-act:hover {
            border-color: var(--gold);
            color: #fff;
            background: rgba(255, 255, 255, 0.08);
        }

        .btn-act.is-fav {
            color: var(--neon-red);
            border-color: var(--neon-red);
        }

        /* Slide-out Queue Panel */
        .queue-deck {
            position: fixed;
            top: 58px;
            right: -360px;
            width: 350px;
            height: calc(100vh - 58px - 94px);
            background: rgba(14, 17, 26, 0.98);
            border-left: 1px solid var(--border);
            box-shadow: -10px 0 30px rgba(0, 0, 0, 0.7);
            z-index: 80;
            display: flex;
            flex-direction: column;
            transition: right 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }

        .queue-deck.open {
            right: 0;
        }

        .queue-deck-header {
            padding: 14px 18px;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .queue-deck-title {
            font-size: 0.92rem;
            font-weight: 800;
            color: var(--gold-light);
        }

        .queue-items {
            flex: 1;
            overflow-y: auto;
            padding: 10px 14px;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        .queue-item {
            padding: 9px 12px;
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            cursor: pointer;
            transition: all 0.15s;
        }

        .queue-item:hover {
            border-color: var(--gold);
        }

        /* Bottom Master DJ Console Deck */
        .master-deck {
            height: 94px;
            background: var(--bg-player);
            backdrop-filter: blur(20px);
            border-top: 1px solid var(--border-gold);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 24px;
            gap: 20px;
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            z-index: 100;
            box-shadow: 0 -8px 30px rgba(0, 0, 0, 0.7);
        }

        /* Left: Current Track */
        .master-left {
            width: 320px;
            display: flex;
            align-items: center;
            gap: 14px;
            flex-shrink: 0;
        }

        .vinyl-art {
            width: 58px;
            height: 58px;
            border-radius: 50%;
            background: radial-gradient(circle, #222 25%, #050505 60%, #333 65%, #111 90%);
            border: 2px solid var(--gold);
            box-shadow: 0 0 14px var(--gold-glow);
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            animation: spin-vinyl 4s linear infinite;
            animation-play-state: paused;
        }

        .vinyl-art.spinning {
            animation-play-state: running;
        }

        .vinyl-needle {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            background: var(--gold-light);
            border: 2px solid #000;
        }

        @keyframes spin-vinyl {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }

        .now-playing-info {
            overflow: hidden;
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .now-title {
            font-size: 0.94rem;
            font-weight: 800;
            color: #fff;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .now-artist {
            font-size: 0.78rem;
            color: var(--gold-light);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .now-cero {
            font-size: 0.68rem;
            color: var(--text-muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        /* Center: Transport & Scrubber */
        .master-center {
            flex: 1;
            max-width: 680px;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
        }

        .transport-bar {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .btn-ctrl {
            background: none;
            border: none;
            color: var(--text-muted);
            cursor: pointer;
            font-size: 16px;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.15s;
        }

        .btn-ctrl:hover {
            color: #fff;
            background: rgba(255, 255, 255, 0.08);
        }

        .btn-ctrl.active {
            color: var(--gold-light);
            background: rgba(245, 158, 11, 0.15);
        }

        .btn-play-hero {
            width: 52px;
            height: 52px;
            border-radius: 50%;
            background: linear-gradient(135deg, #b45309, #f59e0b);
            border: none;
            color: #000;
            font-size: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 0 20px var(--gold-glow);
            transition: all 0.2s;
        }

        .btn-play-hero:hover {
            transform: scale(1.08);
            box-shadow: 0 0 28px var(--gold-glow);
        }

        .scrubber-bar {
            width: 100%;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .time-label {
            font-size: 0.72rem;
            color: var(--text-muted);
            font-family: monospace;
            width: 42px;
            text-align: center;
        }

        .track-timeline {
            flex: 1;
            height: 6px;
            background: rgba(255, 255, 255, 0.12);
            border-radius: 3px;
            cursor: pointer;
            position: relative;
            overflow: hidden;
            transition: height 0.15s;
        }

        .track-timeline:hover {
            height: 8px;
        }

        .track-progress {
            height: 100%;
            background: linear-gradient(90deg, #b45309, #f59e0b);
            width: 0%;
            border-radius: 3px;
        }

        /* Right: Tools & Visualizer */
        .master-right {
            width: 320px;
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 14px;
            flex-shrink: 0;
        }

        .canvas-spectrum {
            width: 80px;
            height: 32px;
            background: rgba(0, 0, 0, 0.3);
            border-radius: 4px;
            border: 1px solid var(--border-gold);
        }

        .volume-deck {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .volume-range {
            width: 80px;
            -webkit-appearance: none;
            height: 4px;
            border-radius: 2px;
            background: rgba(255, 255, 255, 0.18);
            outline: none;
            cursor: pointer;
        }

        .volume-range::-webkit-slider-thumb {
            -webkit-appearance: none;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: var(--gold-light);
            cursor: pointer;
            box-shadow: 0 0 8px var(--gold-glow);
        }

        /* Fullscreen Stage Overlay */
        .fullscreen-stage {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: #06080e;
            z-index: 2000;
            display: none;
            flex-direction: column;
            justify-content: space-between;
            padding: 30px 60px;
            background-image: radial-gradient(circle at 50% 30%, rgba(245, 158, 11, 0.14) 0%, transparent 65%);
        }

        .fullscreen-stage.open {
            display: flex;
        }

        .stage-title-huge {
            font-size: 2.6rem;
            font-weight: 900;
            background: linear-gradient(90deg, #fff, #fbbf24);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-align: center;
        }

        .stage-artist-huge {
            font-size: 1.4rem;
            color: var(--gold-light);
            text-align: center;
            margin-top: 6px;
        }

        /* Toast notifications */
        .toast-box {
            position: fixed;
            bottom: 110px;
            right: 24px;
            background: rgba(18, 22, 34, 0.96);
            border: 1px solid var(--gold);
            color: #fff;
            padding: 10px 18px;
            border-radius: var(--radius-sm);
            font-size: 0.84rem;
            font-weight: 600;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.6);
            z-index: 5000;
            display: none;
            align-items: center;
            gap: 8px;
        }
    </style>
</head>
<body>

    <!-- Hidden HTML5 Audio Engine -->
    <audio id="audio-core" preload="auto"></audio>

    <!-- Top Navbar -->
    <div class="top-navbar">
        <div class="brand">
            <div class="brand-icon">👑</div>
            <div>
                <div class="brand-title">WEDDING DJ STUDIO</div>
                <div class="brand-subtitle">Ramesh Brother's Wedding • 10-12 Dec 2026</div>
            </div>
        </div>

        <div class="top-search-box">
            <span class="search-icon">🔍</span>
            <input type="text" class="top-search-input" id="search-input" placeholder="Search 508 wedding songs by title, singer, movie... (Press '/' to focus)">
        </div>

        <div class="top-actions">
            <button class="btn-top" id="btn-shuffle-all" title="Shuffle Entire 508 Song Library">🔀 Shuffle All</button>
            <button class="btn-top" id="btn-shortcuts" title="Keyboard Shortcuts">⌨️ Shortcuts</button>
            <button class="btn-top" id="btn-stage-toggle" title="Toggle Fullscreen Stage (F)">⛶ DJ Stage</button>
        </div>
    </div>

    <!-- Workspace Deck -->
    <div class="workspace">

        <!-- Sidebar Deck -->
        <div class="sidebar">

            <!-- Day Selectors -->
            <div class="sidebar-section">
                <div class="sidebar-label">📅 Wedding Days</div>
                <div class="nav-item active" data-type="day" data-val="all">
                    <span>🌟 Complete Library</span>
                    <span class="badge-count">508</span>
                </div>
                <div class="nav-item" data-type="day" data-val="day1">
                    <span>🟡 Day 1: Mayan & Haldi</span>
                    <span class="badge-count">115</span>
                </div>
                <div class="nav-item" data-type="day" data-val="day2">
                    <span>🟢 Day 2: Mehndi & Sangeet</span>
                    <span class="badge-count">155</span>
                </div>
                <div class="nav-item" data-type="day" data-val="day3">
                    <span>🔴 Day 3: Wedding Day</span>
                    <span class="badge-count">167</span>
                </div>
                <div class="nav-item" data-type="day" data-val="bagheli">
                    <span>🌾 Bagheli (Rewa) Songs</span>
                    <span class="badge-count">32</span>
                </div>
                <div class="nav-item" data-type="day" data-val="bhojpuri">
                    <span>🪕 Bhojpuri Celebrations</span>
                    <span class="badge-count">39</span>
                </div>
                <div class="nav-item" data-type="day" data-val="favorites">
                    <span>⭐ Starred Favourites</span>
                    <span class="badge-count" id="fav-badge">0</span>
                </div>
            </div>

            <!-- Language Filter -->
            <div class="sidebar-section">
                <div class="sidebar-label">🌍 Language Filter</div>
                <div class="lang-chips">
                    <div class="chip active" data-type="lang" data-val="all">All (508)</div>
                    <div class="chip" data-type="lang" data-val="Hindi & Bollywood">Hindi</div>
                    <div class="chip" data-type="lang" data-val="Punjabi">Punjabi</div>
                    <div class="chip" data-type="lang" data-val="Bagheli">Bagheli</div>
                    <div class="chip" data-type="lang" data-val="Bhojpuri">Bhojpuri</div>
                </div>
            </div>

            <!-- Functions / Ceremonies List -->
            <div class="sidebar-section" style="border-bottom: none; flex: 1;">
                <div class="sidebar-label">
                    <span>🎛️ Functions & Rituals</span>
                    <span style="font-size: 0.68rem; color: var(--gold-light); cursor: pointer;" onclick="resetFunctionFilter()">Reset</span>
                </div>
                <div id="sidebar-cero-list">
                    <!-- Injected by JavaScript -->
                </div>
            </div>

        </div>

        <!-- Main Stage Deck -->
        <div class="stage">

            <!-- Active Ceremony Banner -->
            <div class="stage-header">
                <div>
                    <div class="stage-header-title" id="stage-banner-title">
                        <span>🌟</span>
                        <span id="banner-title-text">Complete Wedding Collection</span>
                    </div>
                    <div class="stage-header-meta" id="stage-banner-meta">
                        508 Tracks • 4.08 GB Offline Library • Family-Safe Selections
                    </div>
                </div>

                <div class="stage-actions">
                    <button class="btn-play-all" id="btn-play-deck">
                        <span>▶ Play All</span>
                        <span id="deck-play-count" style="background: rgba(0,0,0,0.25); padding: 2px 7px; border-radius: 10px; font-size: 0.76rem;">508</span>
                    </button>
                    <button class="btn-secondary" id="btn-shuffle-deck">🔀 Shuffle</button>
                    <button class="btn-secondary" id="btn-queue-deck">➕ Queue All</button>
                </div>
            </div>

            <!-- Song Cue Table -->
            <table class="cue-table">
                <thead>
                    <tr>
                        <th style="width: 45px; text-align: center;">#</th>
                        <th>Track Title & Artist</th>
                        <th>Ceremony / Function</th>
                        <th>Day</th>
                        <th>Language</th>
                        <th>Size</th>
                        <th style="text-align: right; width: 120px;">Actions</th>
                    </tr>
                </thead>
                <tbody id="cue-table-body">
                    <!-- Injected by JavaScript -->
                </tbody>
            </table>

        </div>

    </div>

    <!-- Slide-out Queue Deck -->
    <div class="queue-deck" id="queue-deck">
        <div class="queue-deck-header">
            <div class="queue-deck-title">📑 Live Playlist Queue (<span id="queue-count-head">0</span>)</div>
            <button class="btn-top" id="btn-clear-q" style="padding: 3px 8px; font-size: 0.7rem;">Clear</button>
            <button class="btn-top" id="btn-close-q" style="padding: 3px 8px; font-size: 0.7rem;">✖</button>
        </div>
        <div class="queue-items" id="queue-items-box">
            <!-- Queue items -->
        </div>
    </div>

    <!-- Master Bottom Player Deck -->
    <div class="master-deck">

        <!-- Left Track Info -->
        <div class="master-left">
            <div class="vinyl-art" id="vinyl-icon">
                <div class="vinyl-needle"></div>
            </div>
            <div class="now-playing-info">
                <div class="now-title" id="now-title">Ready to Play</div>
                <div class="now-artist" id="now-artist">Ramesh Brother's Wedding</div>
                <div class="now-cero" id="now-cero">508 Offline Songs Available</div>
            </div>
            <button class="btn-act" id="btn-now-fav" title="Mark Favourite" style="margin-left: auto;">🤍</button>
        </div>

        <!-- Center Transport & Scrubber -->
        <div class="master-center">
            <div class="transport-bar">
                <button class="btn-ctrl" id="btn-shuffle-toggle" title="Shuffle (S)">🔀</button>
                <button class="btn-ctrl" id="btn-prev-track" title="Previous Track (Left Arrow)">⏮</button>
                <button class="btn-play-hero" id="btn-hero-play" title="Play / Pause (Space)">▶</button>
                <button class="btn-ctrl" id="btn-next-track" title="Next Track (Right Arrow)">⏭</button>
                <button class="btn-ctrl" id="btn-repeat-toggle" title="Repeat Mode: Off (R)">🔁</button>
            </div>

            <div class="scrubber-bar">
                <span class="time-label" id="lbl-time-cur">00:00</span>
                <div class="track-timeline" id="scrub-timeline">
                    <div class="track-progress" id="scrub-fill"></div>
                </div>
                <span class="time-label" id="lbl-time-dur">00:00</span>
            </div>
        </div>

        <!-- Right Tools & Volume -->
        <div class="master-right">
            <canvas class="canvas-spectrum" id="canvas-eq" width="80" height="32"></canvas>

            <div class="volume-deck">
                <button class="btn-ctrl" id="btn-vol-mute" title="Mute (M)">🔊</button>
                <input type="range" class="volume-range" id="vol-slider" min="0" max="1" step="0.01" value="0.85">
            </div>

            <button class="btn-top" id="btn-toggle-q-deck">
                <span>📑 Queue</span>
                <span class="badge-count" id="q-dock-badge" style="background: var(--gold); color: #000; font-weight: 800;">0</span>
            </button>
        </div>

    </div>

    <!-- Fullscreen DJ Stage Mode -->
    <div class="fullscreen-stage" id="fullscreen-stage">
        <div style="display: flex; align-items: center; justify-content: space-between;">
            <div style="font-size: 1.1rem; font-weight: 800; color: var(--gold-light);">👑 RAMESH WEDDING DJ CONCERT STAGE</div>
            <button class="btn-top" id="btn-exit-stage" style="padding: 8px 16px;">✖ Exit Fullscreen (Esc)</button>
        </div>

        <div style="display: flex; flex-direction: column; align-items: center; gap: 20px;">
            <div class="vinyl-art" id="stage-vinyl" style="width: 180px; height: 180px; box-shadow: 0 0 50px var(--gold-glow);">
                <div class="vinyl-needle" style="width: 50px; height: 50px;"></div>
            </div>

            <div>
                <div class="stage-title-huge" id="stage-huge-title">Ready to Play</div>
                <div class="stage-artist-huge" id="stage-huge-artist">Ramesh Brother's Wedding</div>
            </div>

            <div style="display: flex; gap: 10px;" id="stage-huge-tags">
                <span class="tag-pill" style="font-size: 0.82rem; padding: 5px 12px; background: rgba(245,158,11,0.2); color: var(--gold-light);" id="stage-tag-day">Day 1</span>
                <span class="tag-pill" style="font-size: 0.82rem; padding: 5px 12px; background: rgba(244,63,94,0.2); color: var(--neon-red);" id="stage-tag-cero">Haldi</span>
            </div>

            <!-- Big Stage Play Controls -->
            <div style="display: flex; align-items: center; gap: 24px; margin-top: 10px;">
                <button class="btn-ctrl" id="stage-btn-prev" style="font-size: 26px; width: 48px; height: 48px;">⏮</button>
                <button class="btn-play-hero" id="stage-btn-play" style="width: 72px; height: 72px; font-size: 28px;">▶</button>
                <button class="btn-ctrl" id="stage-btn-next" style="font-size: 26px; width: 48px; height: 48px;">⏭</button>
            </div>
        </div>

        <div style="text-align: center; color: var(--text-muted); font-size: 0.82rem;">
            Space: Play/Pause • Right Arrow: Next • Left Arrow: Previous • Up/Down: Volume • Esc: Exit
        </div>
    </div>

    <!-- Toast Notice -->
    <div class="toast-box" id="toast-box">
        <span>🎵</span>
        <span id="toast-text">Message</span>
    </div>

    <!-- App Engine JavaScript -->
    <script>
        // Database of all 508 verified offline songs
        const SONGS = __SONGS_DATA__;

        // DJ Console State
        const state = {
            currentIndex: -1,
            isPlaying: false,
            isShuffle: false,
            repeat: 'off', // 'off', 'all', 'one'
            volume: 0.85,
            isMuted: false,
            filterDay: 'all',
            filterLang: 'all',
            filterCero: 'all',
            searchQuery: '',
            queue: [],
            favorites: new Set(JSON.parse(localStorage.getItem('wedding_favs') || '[]'))
        };

        // Audio Engine
        const audio = document.getElementById('audio-core');

        // Elements
        const btnHeroPlay = document.getElementById('btn-hero-play');
        const btnPrevTrack = document.getElementById('btn-prev-track');
        const btnNextTrack = document.getElementById('btn-next-track');
        const btnShuffleToggle = document.getElementById('btn-shuffle-toggle');
        const btnRepeatToggle = document.getElementById('btn-repeat-toggle');
        const scrubTimeline = document.getElementById('scrub-timeline');
        const scrubFill = document.getElementById('scrub-fill');
        const lblTimeCur = document.getElementById('lbl-time-cur');
        const lblTimeDur = document.getElementById('lbl-time-dur');
        const volSlider = document.getElementById('vol-slider');
        const btnVolMute = document.getElementById('btn-vol-mute');
        const nowTitle = document.getElementById('now-title');
        const nowArtist = document.getElementById('now-artist');
        const nowCero = document.getElementById('now-cero');
        const btnNowFav = document.getElementById('btn-now-fav');
        const vinylIcon = document.getElementById('vinyl-icon');
        const searchInput = document.getElementById('search-input');
        const cueTableBody = document.getElementById('cue-table-body');
        const sidebarCeroList = document.getElementById('sidebar-cero-list');
        const bannerTitleText = document.getElementById('banner-title-text');
        const stageBannerMeta = document.getElementById('stage-banner-meta');
        const deckPlayCount = document.getElementById('deck-play-count');
        const btnPlayDeck = document.getElementById('btn-play-deck');
        const btnShuffleDeck = document.getElementById('btn-shuffle-deck');
        const btnQueueDeck = document.getElementById('btn-queue-deck');
        const queueDeck = document.getElementById('queue-deck');
        const btnToggleQDeck = document.getElementById('btn-toggle-q-deck');
        const btnCloseQ = document.getElementById('btn-close-q');
        const btnClearQ = document.getElementById('btn-clear-q');
        const qDockBadge = document.getElementById('q-dock-badge');
        const queueCountHead = document.getElementById('queue-count-head');
        const queueItemsBox = document.getElementById('queue-items-box');
        const favBadge = document.getElementById('fav-badge');
        const fullscreenStage = document.getElementById('fullscreen-stage');
        const btnStageToggle = document.getElementById('btn-stage-toggle');
        const btnExitStage = document.getElementById('btn-exit-stage');
        const stageHugeTitle = document.getElementById('stage-huge-title');
        const stageHugeArtist = document.getElementById('stage-huge-artist');
        const stageTagDay = document.getElementById('stage-tag-day');
        const stageTagCero = document.getElementById('stage-tag-cero');
        const stageVinyl = document.getElementById('stage-vinyl');
        const stageBtnPlay = document.getElementById('stage-btn-play');
        const stageBtnPrev = document.getElementById('stage-btn-prev');
        const stageBtnNext = document.getElementById('stage-btn-next');

        // Audio Spectrum Analyser
        let audioCtx, analyser, dataArray;
        function initSpectrum() {
            if (audioCtx) return;
            try {
                audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                analyser = audioCtx.createAnalyser();
                const src = audioCtx.createMediaElementSource(audio);
                src.connect(analyser);
                analyser.connect(audioCtx.destination);
                analyser.fftSize = 64;
                dataArray = new Uint8Array(analyser.frequencyBinCount);
                renderSpectrum();
            } catch(e) {
                console.warn('Spectrum init:', e);
            }
        }

        function renderSpectrum() {
            requestAnimationFrame(renderSpectrum);
            if (!analyser || !dataArray) return;
            analyser.getByteFrequencyData(dataArray);
            const canvas = document.getElementById('canvas-eq');
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            const barW = (canvas.width / 16) - 1;
            for (let i = 0; i < 16; i++) {
                const h = (dataArray[i * 2] / 255) * canvas.height;
                ctx.fillStyle = '#f59e0b';
                ctx.fillRect(i * (barW + 1), canvas.height - h, barW, h);
            }
        }

        // Format MM:SS
        function fmtTime(sec) {
            if (isNaN(sec) || sec < 0) return '00:00';
            const m = Math.floor(sec / 60);
            const s = Math.floor(sec % 60);
            return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
        }

        // Toast
        let toastTimer;
        function toast(msg) {
            const b = document.getElementById('toast-box');
            document.getElementById('toast-text').textContent = msg;
            b.style.display = 'flex';
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => { b.style.display = 'none'; }, 2400);
        }

        // Filter Songs
        function getFiltered() {
            return SONGS.filter(s => {
                // Day filter
                if (state.filterDay !== 'all') {
                    if (state.filterDay === 'favorites') {
                        if (!state.favorites.has(s.id)) return false;
                    } else if (state.filterDay !== s.day_key) {
                        return false;
                    }
                }
                // Language filter
                if (state.filterLang !== 'all' && s.language !== state.filterLang) {
                    return false;
                }
                // Ceremony filter
                if (state.filterCero !== 'all' && s.function !== state.filterCero) {
                    return false;
                }
                // Search query
                if (state.searchQuery) {
                    const q = state.searchQuery.toLowerCase();
                    const txt = `${s.title} ${s.artist} ${s.movie} ${s.function} ${s.language}`.toLowerCase();
                    if (!txt.includes(q)) return false;
                }
                return true;
            });
        }

        // Render Sidebar Ceremony List
        function renderSidebarCeremonies() {
            // Count songs per ceremony within current day & lang filter
            const counts = {};
            SONGS.forEach(s => {
                if (state.filterDay !== 'all' && state.filterDay !== 'favorites' && s.day_key !== state.filterDay) return;
                if (state.filterDay === 'favorites' && !state.favorites.has(s.id)) return;
                if (state.filterLang !== 'all' && s.language !== state.filterLang) return;
                counts[s.function] = (counts[s.function] || 0) + 1;
            });

            const items = Object.entries(counts).sort((a,b) => b[1] - a[1]);
            sidebarCeroList.innerHTML = items.map(([funcName, count]) => {
                const isActive = state.filterCero === funcName;
                return `
                    <div class="nav-item ${isActive ? 'active' : ''}" onclick="selectCeremony('${funcName}')">
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding-right: 6px;">${funcName}</span>
                        <span class="badge-count">${count}</span>
                    </div>
                `;
            }).join('');
        }

        function selectCeremony(funcName) {
            state.filterCero = (state.filterCero === funcName) ? 'all' : funcName;
            renderSidebarCeremonies();
            renderMainDeck();
        }

        window.resetFunctionFilter = function() {
            state.filterCero = 'all';
            renderSidebarCeremonies();
            renderMainDeck();
        };

        // Render Main Deck
        function renderMainDeck() {
            const list = getFiltered();
            deckPlayCount.textContent = list.length;
            favBadge.textContent = state.favorites.size;

            // Update Header Banner
            if (state.searchQuery) {
                bannerTitleText.textContent = `Search: "${state.searchQuery}"`;
                stageBannerMeta.textContent = `Found ${list.length} matching songs`;
            } else if (state.filterCero !== 'all') {
                bannerTitleText.textContent = state.filterCero;
                stageBannerMeta.textContent = `${list.length} Songs • Dedicated Ceremony Collection`;
            } else if (state.filterDay !== 'all') {
                const labels = {
                    day1: 'Day 1: Mayan & Haldi',
                    day2: 'Day 2: Mehndi & Sangeet',
                    day3: 'Day 3: Wedding Day',
                    bagheli: 'Bagheli (Rewa) Traditional Collection',
                    bhojpuri: 'Bhojpuri Wedding Collection',
                    favorites: 'Starred Favourites'
                };
                bannerTitleText.textContent = labels[state.filterDay] || 'Wedding Selection';
                stageBannerMeta.textContent = `${list.length} Songs in this category`;
            } else {
                bannerTitleText.textContent = 'Complete Wedding Collection';
                stageBannerMeta.textContent = '508 Tracks • 4.08 GB Offline Library • Family-Safe Selections';
            }

            if (list.length === 0) {
                cueTableBody.innerHTML = `
                    <tr>
                        <td colspan="7" style="text-align: center; padding: 60px 20px; color: var(--text-muted);">
                            <div style="font-size: 2.2rem; margin-bottom: 8px;">🔍</div>
                            <div style="font-weight: 700; color: #fff; font-size: 1.1rem; margin-bottom: 4px;">No Songs Match Current Filters</div>
                            <div style="font-size: 0.84rem;">Try resetting your language or ceremony filter on the left.</div>
                        </td>
                    </tr>
                `;
                return;
            }

            cueTableBody.innerHTML = list.map((s, idx) => {
                const isPlaying = state.currentIndex !== -1 && SONGS[state.currentIndex].id === s.id;
                const isFav = state.favorites.has(s.id);
                const masterIdx = SONGS.findIndex(item => item.id === s.id);
                const langClass = s.language === 'Punjabi' ? 'lang-punjabi' : 
                                 s.language === 'Bagheli' ? 'lang-bagheli' : 
                                 s.language === 'Bhojpuri' ? 'lang-bhojpuri' : 'lang-hindi';

                return `
                    <tr class="cue-row ${isPlaying ? 'playing' : ''}" data-id="${s.id}" onclick="playByIndex(${masterIdx}, true)">
                        <td class="cue-index">
                            ${isPlaying && state.isPlaying ? '🔊' : (idx + 1)}
                        </td>
                        <td>
                            <div class="song-title-col">
                                <span class="main-name">${s.title}</span>
                                <span class="song-artist-sub">${s.artist}${s.movie ? ' • ' + s.movie : ''}</span>
                            </div>
                        </td>
                        <td>
                            <span class="tag-pill">${s.function}</span>
                        </td>
                        <td>
                            <span class="tag-pill" style="color: var(--gold-light);">${s.day_name}</span>
                        </td>
                        <td>
                            <span class="tag-pill ${langClass}">${s.language}</span>
                        </td>
                        <td>
                            <span style="font-size: 0.76rem; color: var(--text-muted);">${s.size_mb} MB</span>
                        </td>
                        <td>
                            <div class="row-btns" onclick="event.stopPropagation()">
                                <button class="btn-act ${isFav ? 'is-fav' : ''}" onclick="toggleFav(${s.id})" title="Favourite">${isFav ? '❤️' : '🤍'}</button>
                                <button class="btn-act" onclick="queueNext(${s.id})" title="Play Next">⏭</button>
                                <button class="btn-act" onclick="queuePush(${s.id})" title="Add to Queue">➕</button>
                            </div>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        // Playback Controller
        function playByIndex(idx, startPlay = true) {
            if (idx < 0 || idx >= SONGS.length) return;
            state.currentIndex = idx;
            const song = SONGS[idx];

            audio.src = encodeURI(song.path);
            audio.load();

            nowTitle.textContent = song.title;
            nowArtist.textContent = song.artist + (song.movie ? ` • ${song.movie}` : '');
            nowCero.textContent = `${song.day_name} • ${song.function} • ${song.language}`;
            btnNowFav.textContent = state.favorites.has(song.id) ? '❤️' : '🤍';
            btnNowFav.classList.toggle('is-fav', state.favorites.has(song.id));

            stageHugeTitle.textContent = song.title;
            stageHugeArtist.textContent = song.artist + (song.movie ? ` (${song.movie})` : '');
            stageTagDay.textContent = song.day_name;
            stageTagCero.textContent = song.function;

            if (startPlay) {
                audio.play().then(() => {
                    state.isPlaying = true;
                    updateDeckState();
                    initSpectrum();
                }).catch(err => {
                    console.warn(err);
                    state.isPlaying = false;
                    updateDeckState();
                });
            }

            renderMainDeck();
            renderQueue();
        }

        function togglePlay() {
            if (state.currentIndex === -1) {
                const list = getFiltered();
                if (list.length > 0) {
                    playByIndex(SONGS.findIndex(s => s.id === list[0].id), true);
                }
                return;
            }

            if (audio.paused) {
                audio.play().then(() => {
                    state.isPlaying = true;
                    updateDeckState();
                    initSpectrum();
                });
            } else {
                audio.pause();
                state.isPlaying = false;
                updateDeckState();
            }
        }

        function nextTrack() {
            if (state.queue.length > 0) {
                const nextId = state.queue.shift();
                const idx = SONGS.findIndex(s => s.id === nextId);
                if (idx !== -1) {
                    playByIndex(idx, true);
                    return;
                }
            }

            const list = getFiltered();
            if (list.length === 0) return;

            if (state.isShuffle) {
                const randPos = Math.floor(Math.random() * list.length);
                playByIndex(SONGS.findIndex(s => s.id === list[randPos].id), true);
                return;
            }

            const curPos = list.findIndex(s => s.id === SONGS[state.currentIndex]?.id);
            if (curPos !== -1 && curPos < list.length - 1) {
                playByIndex(SONGS.findIndex(s => s.id === list[curPos + 1].id), true);
            } else if (state.repeat === 'all') {
                playByIndex(SONGS.findIndex(s => s.id === list[0].id), true);
            } else {
                toast('Playlist complete');
            }
        }

        function prevTrack() {
            if (audio.currentTime > 3) {
                audio.currentTime = 0;
                return;
            }
            const list = getFiltered();
            if (list.length === 0) return;

            const curPos = list.findIndex(s => s.id === SONGS[state.currentIndex]?.id);
            if (curPos > 0) {
                playByIndex(SONGS.findIndex(s => s.id === list[curPos - 1].id), true);
            } else {
                audio.currentTime = 0;
            }
        }

        function updateDeckState() {
            if (state.isPlaying) {
                btnHeroPlay.innerHTML = '⏸';
                stageBtnPlay.innerHTML = '⏸';
                vinylIcon.classList.add('spinning');
                stageVinyl.classList.add('spinning');
            } else {
                btnHeroPlay.innerHTML = '▶';
                stageBtnPlay.innerHTML = '▶';
                vinylIcon.classList.remove('spinning');
                stageVinyl.classList.remove('spinning');
            }
        }

        function toggleFav(id) {
            if (state.favorites.has(id)) {
                state.favorites.delete(id);
                toast('Removed from favourites');
            } else {
                state.favorites.add(id);
                toast('Added to favourites ⭐');
            }
            localStorage.setItem('wedding_favs', JSON.stringify(Array.from(state.favorites)));
            favBadge.textContent = state.favorites.size;
            if (state.currentIndex !== -1 && SONGS[state.currentIndex].id === id) {
                btnNowFav.textContent = state.favorites.has(id) ? '❤️' : '🤍';
                btnNowFav.classList.toggle('is-fav', state.favorites.has(id));
            }
            renderMainDeck();
        }

        // Queue Actions
        function queuePush(id) {
            state.queue.push(id);
            renderQueue();
            const s = SONGS.find(item => item.id === id);
            toast(`Queued: ${s.title}`);
        }

        function queueNext(id) {
            state.queue.unshift(id);
            renderQueue();
            const s = SONGS.find(item => item.id === id);
            toast(`Up Next: ${s.title}`);
        }

        function renderQueue() {
            qDockBadge.textContent = state.queue.length;
            queueCountHead.textContent = state.queue.length;

            if (state.queue.length === 0) {
                queueItemsBox.innerHTML = '<div style="color: var(--text-muted); text-align: center; padding: 40px 10px; font-size: 0.84rem;">Queue is empty.<br>Click ➕ on any track to add it.</div>';
                return;
            }

            queueItemsBox.innerHTML = state.queue.map((id, qIdx) => {
                const s = SONGS.find(item => item.id === id);
                if (!s) return '';
                return `
                    <div class="queue-item">
                        <div style="flex: 1; overflow: hidden;" onclick="playByIndex(${SONGS.findIndex(item => item.id === id)}, true); state.queue.splice(${qIdx}, 1); renderQueue();">
                            <div style="font-weight: 700; font-size: 0.84rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #fff;">${s.title}</div>
                            <div style="font-size: 0.72rem; color: var(--gold-light);">${s.artist} • ${s.function}</div>
                        </div>
                        <button class="btn-act" onclick="state.queue.splice(${qIdx}, 1); renderQueue();" style="width: 24px; height: 24px; font-size: 11px;">✖</button>
                    </div>
                `;
            }).join('');
        }

        // Audio Engine Event Listeners
        audio.addEventListener('timeupdate', () => {
            const cur = audio.currentTime;
            const dur = audio.duration;
            lblTimeCur.textContent = fmtTime(cur);
            if (!isNaN(dur) && dur > 0) {
                scrubFill.style.width = (cur / dur * 100) + '%';
                lblTimeDur.textContent = fmtTime(dur);
            }
        });

        audio.addEventListener('ended', () => {
            if (state.repeat === 'one') {
                audio.currentTime = 0;
                audio.play();
            } else {
                nextTrack();
            }
        });

        audio.addEventListener('play', () => {
            state.isPlaying = true;
            updateDeckState();
        });

        audio.addEventListener('pause', () => {
            state.isPlaying = false;
            updateDeckState();
        });

        // Scrubber
        scrubTimeline.addEventListener('click', (e) => {
            const rect = scrubTimeline.getBoundingClientRect();
            const pct = (e.clientX - rect.left) / rect.width;
            if (!isNaN(audio.duration)) {
                audio.currentTime = pct * audio.duration;
            }
        });

        // Volume
        volSlider.addEventListener('input', (e) => {
            state.volume = parseFloat(e.target.value);
            audio.volume = state.volume;
            state.isMuted = state.volume === 0;
            btnVolMute.innerHTML = state.isMuted ? '🔇' : (state.volume > 0.5 ? '🔊' : '🔉');
        });

        btnVolMute.addEventListener('click', () => {
            if (state.isMuted) {
                state.isMuted = false;
                audio.volume = state.volume || 0.85;
                volSlider.value = audio.volume;
                btnVolMute.innerHTML = '🔊';
            } else {
                state.isMuted = true;
                audio.volume = 0;
                volSlider.value = 0;
                btnVolMute.innerHTML = '🔇';
            }
        });

        // Shuffle & Repeat
        btnShuffleToggle.addEventListener('click', () => {
            state.isShuffle = !state.isShuffle;
            btnShuffleToggle.classList.toggle('active', state.isShuffle);
            toast(state.isShuffle ? '🔀 Shuffle Activated' : '➡️ Sequential Playback');
        });

        btnRepeatToggle.addEventListener('click', () => {
            if (state.repeat === 'off') {
                state.repeat = 'all';
                btnRepeatToggle.classList.add('active');
                btnRepeatToggle.innerHTML = '🔁';
                toast('🔁 Repeat Ceremony All');
            } else if (state.repeat === 'all') {
                state.repeat = 'one';
                btnRepeatToggle.classList.add('active');
                btnRepeatToggle.innerHTML = '🔂';
                toast('🔂 Repeat Current Track');
            } else {
                state.repeat = 'off';
                btnRepeatToggle.classList.remove('active');
                btnRepeatToggle.innerHTML = '🔁';
                toast('➡️ Repeat Off');
            }
        });

        // Search Input
        searchInput.addEventListener('input', (e) => {
            state.searchQuery = e.target.value.trim();
            renderMainDeck();
        });

        // Sidebar Day Clicks
        document.querySelectorAll('.sidebar .nav-item[data-type="day"]').forEach(el => {
            el.addEventListener('click', () => {
                document.querySelectorAll('.sidebar .nav-item[data-type="day"]').forEach(item => item.classList.remove('active'));
                el.classList.add('active');
                state.filterDay = el.getAttribute('data-val');
                state.filterCero = 'all';
                renderSidebarCeremonies();
                renderMainDeck();
            });
        });

        // Sidebar Language Chips
        document.querySelectorAll('.lang-chips .chip').forEach(chip => {
            chip.addEventListener('click', () => {
                document.querySelectorAll('.lang-chips .chip').forEach(c => c.classList.remove('active'));
                chip.classList.add('active');
                state.filterLang = chip.getAttribute('data-val');
                renderSidebarCeremonies();
                renderMainDeck();
            });
        });

        // Deck Action Buttons
        btnPlayDeck.addEventListener('click', () => {
            const list = getFiltered();
            if (list.length > 0) {
                playByIndex(SONGS.findIndex(s => s.id === list[0].id), true);
                toast(`Playing ${list.length} tracks`);
            }
        });

        btnShuffleDeck.addEventListener('click', () => {
            state.isShuffle = true;
            btnShuffleToggle.classList.add('active');
            const list = getFiltered();
            if (list.length > 0) {
                const randPos = Math.floor(Math.random() * list.length);
                playByIndex(SONGS.findIndex(s => s.id === list[randPos].id), true);
                toast(`🔀 Shuffled & playing ${list.length} tracks`);
            }
        });

        btnQueueDeck.addEventListener('click', () => {
            const list = getFiltered();
            list.forEach(s => state.queue.push(s.id));
            renderQueue();
            toast(`Added ${list.length} songs to Queue`);
        });

        document.getElementById('btn-shuffle-all').addEventListener('click', () => {
            state.filterDay = 'all';
            state.filterLang = 'all';
            state.filterCero = 'all';
            state.isShuffle = true;
            btnShuffleToggle.classList.add('active');
            document.querySelectorAll('.sidebar .nav-item[data-type="day"]').forEach(item => item.classList.remove('active'));
            document.querySelector('.sidebar .nav-item[data-val="all"]').classList.add('active');
            document.querySelectorAll('.lang-chips .chip').forEach(c => c.classList.remove('active'));
            document.querySelector('.lang-chips .chip[data-val="all"]').classList.add('active');
            renderSidebarCeremonies();
            renderMainDeck();
            const randPos = Math.floor(Math.random() * SONGS.length);
            playByIndex(randPos, true);
            toast('🔀 Shuffled entire 508-song wedding library!');
        });

        // Master Deck Buttons
        btnHeroPlay.addEventListener('click', togglePlay);
        stageBtnPlay.addEventListener('click', togglePlay);
        btnNextTrack.addEventListener('click', nextTrack);
        stageBtnNext.addEventListener('click', nextTrack);
        btnPrevTrack.addEventListener('click', prevTrack);
        stageBtnPrev.addEventListener('click', prevTrack);
        btnNowFav.addEventListener('click', () => {
            if (state.currentIndex !== -1) toggleFav(SONGS[state.currentIndex].id);
        });

        // Queue Panel Toggle
        btnToggleQDeck.addEventListener('click', () => queueDeck.classList.toggle('open'));
        btnCloseQ.addEventListener('click', () => queueDeck.classList.remove('open'));
        btnClearQ.addEventListener('click', () => {
            state.queue = [];
            renderQueue();
            toast('Queue cleared');
        });

        // Stage Mode
        btnStageToggle.addEventListener('click', () => {
            fullscreenStage.classList.add('open');
            if (document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
        });

        btnExitStage.addEventListener('click', () => {
            fullscreenStage.classList.remove('open');
            if (document.exitFullscreen) {
                document.exitFullscreen().catch(() => {});
            }
        });

        // Keyboard Shortcuts
        window.addEventListener('keydown', (e) => {
            if (e.target === searchInput) {
                if (e.key === 'Escape') searchInput.blur();
                return;
            }

            switch(e.code) {
                case 'Space':
                    e.preventDefault();
                    togglePlay();
                    break;
                case 'ArrowRight':
                    e.preventDefault();
                    nextTrack();
                    break;
                case 'ArrowLeft':
                    e.preventDefault();
                    prevTrack();
                    break;
                case 'ArrowUp':
                    e.preventDefault();
                    state.volume = Math.min(1, state.volume + 0.05);
                    audio.volume = state.volume;
                    volSlider.value = state.volume;
                    toast(`🔊 Volume ${Math.round(state.volume * 100)}%`);
                    break;
                case 'ArrowDown':
                    e.preventDefault();
                    state.volume = Math.max(0, state.volume - 0.05);
                    audio.volume = state.volume;
                    volSlider.value = state.volume;
                    toast(`🔉 Volume ${Math.round(state.volume * 100)}%`);
                    break;
                case 'KeyM':
                    btnVolMute.click();
                    break;
                case 'KeyS':
                    btnShuffleToggle.click();
                    break;
                case 'KeyR':
                    btnRepeatToggle.click();
                    break;
                case 'KeyF':
                    if (fullscreenStage.classList.contains('open')) {
                        btnExitStage.click();
                    } else {
                        btnStageToggle.click();
                    }
                    break;
                case 'Slash':
                    e.preventDefault();
                    searchInput.focus();
                    break;
                case 'Escape':
                    if (fullscreenStage.classList.contains('open')) btnExitStage.click();
                    if (queueDeck.classList.contains('open')) queueDeck.classList.remove('open');
                    break;
            }
        });

        document.getElementById('btn-shortcuts').addEventListener('click', () => {
            toast('Shortcuts: Space (Play), → (Next), ← (Prev), ↑/↓ (Vol), M (Mute), S (Shuffle), R (Repeat), F (Stage)');
        });

        // Boot
        audio.volume = state.volume;
        renderSidebarCeremonies();
        renderMainDeck();
        renderQueue();
        console.log('👑 Wedding DJ Studio loaded with', SONGS.length, 'offline tracks.');
    </script>
</body>
</html>
"""

final_html = html_code.replace("__SONGS_DATA__", songs_json)

# Write to Wedding_Music_Player.html & index.html
player_path = os.path.join(BASE_DIR, "Wedding_Music_Player.html")
with open(player_path, "w", encoding="utf-8") as f:
    f.write(final_html)

index_path = os.path.join(BASE_DIR, "index.html")
with open(index_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print("✅ Premium Wedding DJ Studio Web App successfully created at:", player_path)

