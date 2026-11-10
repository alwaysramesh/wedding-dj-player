import os
import json
import re
import urllib.parse
import openpyxl

REPO_ROOT = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S"
MUSIC_DIR = os.path.join(REPO_ROOT, "Brother Wedding Music")
EXCEL_FILE = os.path.join(MUSIC_DIR, "Brother_Wedding_Music_Collection.xlsx")

# Load Excel metadata including YouTube streaming links
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
        k = re.sub(r'[^a-zA-Z0-9]', '', name.lower())
        if k not in meta:
            meta[k] = {'singer': singer, 'movie': movie, 'category': cat, 'year': year, 'ceremony': cero, 'day': day, 'yt': yt_link}

punjabi_singers = [
    'diljit dosanjh', 'sukhbir', 'panjabi mc', 'mika singh', 'guru randhawa',
    'daler mehndi', 'jasbir jassi', 'bally sagoo', 'honey singh', 'badshah',
    'ap dhillon', 'karan aujla', 'hardy sandhu', 'amrit maan', 'sidhu moose wala',
    'gippy grewal', 'surinder kaur', 'gurmeet bawa', 'jasleen royal'
]

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
            
            # YouTube Stream Link
            stream_link = m.get('yt') or f"https://www.youtube.com/results?search_query={urllib.parse.quote(title + ' ' + (m.get('singer') or artist) + ' audio song')}"

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
                'sub_path': sub_path,
                'root_path': root_path,
                'stream_url': stream_link,
                'size_mb': round(os.path.getsize(full) / (1024*1024), 2)
            })

songs_json = json.dumps(songs, ensure_ascii=False)

# Read HTML template from create_premium_dj_player.py
with open("create_premium_dj_player.py", "r", encoding="utf-8") as f:
    orig = f.read()

start_str = 'html_code = """'
end_str = '"""\n\nfinal_html'
s_idx = orig.find(start_str) + len(start_str)
e_idx = orig.find(end_str)
html_code = orig[s_idx:e_idx]

# Update path resolver inside playByIndex
old_src = "audio.src = encodeURI(song.path);"
new_src = """
            // Smart Path Resolver (Works at Repo Root, Subfolder & GitHub Pages)
            const isGitHubPages = window.location.hostname.includes('github.io');
            const inSubfolder = window.location.pathname.includes('Brother%20Wedding%20Music') || window.location.pathname.includes('Brother Wedding Music');
            const targetAudioPath = inSubfolder ? song.sub_path : song.root_path;

            audio.src = encodeURI(targetAudioPath);
            audio.load();
"""

# Update error handler to suggest or auto-fallback to stream if on GitHub Pages
old_err = "showToast('⚠️ Audio loading error. Trying fallback...');"
new_err = """
        audio.addEventListener('error', () => {
            const song = SONGS[state.currentIndex];
            if (window.location.hostname.includes('github.io')) {
                toast(`🌐 GitHub Pages Mode: Opening official stream for ${song.title}`);
                window.open(song.stream_url, '_blank');
            } else {
                toast(`⚠️ File not found locally. Click "▶ Stream" to play online.`);
            }
        });
"""

html_code = html_code.replace(old_src, new_src)

# Add "Stream Online" button in table row actions
old_row_actions = """                                <button class="btn-act ${isFav ? 'is-fav' : ''}" onclick="toggleFav(${s.id})" title="Favourite">${isFav ? '❤️' : '🤍'}</button>
                                <button class="btn-act" onclick="queueNext(${s.id})" title="Play Next">⏭</button>
                                <button class="btn-act" onclick="queuePush(${s.id})" title="Add to Queue">➕</button>"""

new_row_actions = """                                <button class="btn-act ${isFav ? 'is-fav' : ''}" onclick="toggleFav(${s.id})" title="Favourite">${isFav ? '❤️' : '🤍'}</button>
                                <button class="btn-act" onclick="queueNext(${s.id})" title="Play Next">⏭</button>
                                <button class="btn-act" onclick="queuePush(${s.id})" title="Add to Queue">➕</button>
                                <a href="${s.stream_url}" target="_blank" class="btn-act" title="Stream Official Song on YouTube / YouTube Music" style="text-decoration: none; font-size: 11px;">▶</a>"""

html_code = html_code.replace(old_row_actions, new_row_actions)

final_html = html_code.replace("__SONGS_DATA__", songs_json)

# 1. Write to Repo Root index.html (for GitHub Pages!)
root_index = os.path.join(REPO_ROOT, "index.html")
with open(root_index, "w", encoding="utf-8") as f:
    f.write(final_html)

# 2. Write to Brother Wedding Music/index.html & Wedding_Music_Player.html
with open(os.path.join(MUSIC_DIR, "index.html"), "w", encoding="utf-8") as f:
    f.write(final_html)

with open(os.path.join(MUSIC_DIR, "Wedding_Music_Player.html"), "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"✅ Created Root index.html for GitHub Pages: {root_index}")
print(f"✅ Updated Brother Wedding Music/index.html")
print(f"✅ Updated Brother Wedding Music/Wedding_Music_Player.html")
