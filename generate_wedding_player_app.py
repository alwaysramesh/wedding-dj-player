import os
import json
import re
import urllib.parse
import openpyxl

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

def load_metadata():
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
                meta[k] = {
                    'singer': singer,
                    'movie': movie,
                    'category': cat,
                    'year': year,
                    'ceremony': cero,
                    'day': day
                }
    return meta

punjabi_singers = [
    'diljit dosanjh', 'sukhbir', 'panjabi mc', 'mika singh', 'guru randhawa',
    'daler mehndi', 'jasbir jassi', 'bally sagoo', 'honey singh', 'badshah',
    'ap dhillon', 'karan aujla', 'hardy sandhu', 'amrit maan', 'sidhu moose wala',
    'gippy grewal', 'surinder kaur', 'gurmeet bawa', 'jasleen royal'
]

def scan_songs(metadata):
    songs = []
    for root, dirs, files in os.walk(BASE_DIR):
        for f in sorted(files):
            if any(f.lower().endswith(ext) for ext in ['.mp3', '.webm', '.m4a', '.opus']):
                full = os.path.join(root, f)
                rel = os.path.relpath(full, BASE_DIR).replace('\\', '/')
                parts = rel.split('/')
                top_folder = parts[0]
                sub_folder = parts[1] if len(parts) > 1 else ''
                
                name_no_ext = os.path.splitext(f)[0]
                if ' - ' in name_no_ext:
                    title, artist = name_no_ext.split(' - ', 1)
                else:
                    title, artist = name_no_ext, 'Traditional'
                    
                norm_key = re.sub(r'[^a-zA-Z0-9]', '', title.lower())
                meta = metadata.get(norm_key, {})
                
                # Language determination
                singer_lower = (meta.get('singer') or artist).lower()
                cat_lower = meta.get('category', '').lower()
                
                if 'bagheli' in top_folder.lower() or 'bagheli' in rel.lower():
                    lang = 'Bagheli (Rewa)'
                elif 'bhojpuri' in top_folder.lower() or 'bhojpuri' in rel.lower():
                    lang = 'Bhojpuri'
                elif 'punjabi' in sub_folder.lower() or 'punjabi' in cat_lower or any(p in singer_lower for p in punjabi_singers):
                    lang = 'Punjabi'
                else:
                    lang = 'Hindi & Bollywood'
                    
                # Day determination
                if '10 december' in top_folder.lower() or 'mayan' in top_folder.lower() or 'haldi' in top_folder.lower():
                    day_key = 'day1'
                    day_label = '10 Dec: Mayan & Haldi'
                elif '11 december' in top_folder.lower() or 'mehndi' in top_folder.lower() or 'sangeet' in top_folder.lower():
                    day_key = 'day2'
                    day_label = '11 Dec: Mehndi & Sangeet'
                elif '12 december' in top_folder.lower() or 'wedding' in top_folder.lower():
                    day_key = 'day3'
                    day_label = '12 Dec: Wedding Day'
                elif 'bagheli' in top_folder.lower():
                    day_key = 'bagheli'
                    day_label = 'Bagheli Traditions (Rewa)'
                elif 'bhojpuri' in top_folder.lower():
                    day_key = 'bhojpuri'
                    day_label = 'Bhojpuri Celebrations'
                else:
                    day_key = 'other'
                    day_label = 'Wedding Collection'
                    
                # Function / Ceremony determination
                f_low = (sub_folder + ' ' + meta.get('ceremony', '')).lower()
                if 'mayan' in f_low or 'mandap' in f_low or 'tilak' in f_low:
                    func = 'Mayan & Mandap'
                elif 'haldi' in f_low or 'nahchhu' in f_low:
                    func = 'Haldi Ceremony'
                elif 'mehndi' in f_low:
                    func = 'Mehndi'
                elif 'ladies sangeet' in f_low:
                    func = 'Ladies Sangeet'
                elif 'family dance' in f_low or 'family celebration' in f_low:
                    func = 'Family Dance'
                elif 'friends dance' in f_low:
                    func = 'Friends Dance'
                elif 'couple dance' in f_low:
                    func = 'Couple Dance'
                elif 'dj party' in f_low or 'dj mix' in f_low or 'dj dance' in f_low or 'blast' in f_low:
                    func = 'DJ Party'
                elif 'ghodi' in f_low or 'procession' in f_low:
                    func = 'Ghodi & Procession'
                elif 'baraat' in f_low or 'groom entry' in f_low:
                    func = 'Baraat & Groom Entry'
                elif 'bride entry' in f_low:
                    func = 'Bride Entry'
                elif 'varmala' in f_low or 'jaimala' in f_low:
                    func = 'Varmala & Jaimala'
                elif 'pheras' in f_low or 'instrumental' in f_low or 'shehnai' in f_low:
                    func = 'Pheras & Instrumental'
                elif 'ritual' in f_low:
                    func = 'Wedding Rituals'
                elif 'reception' in f_low:
                    func = 'Reception'
                elif 'vidaai' in f_low or 'bidaai' in f_low:
                    func = 'Vidaai'
                elif 'funny' in f_low:
                    func = 'Funny & Entertaining'
                else:
                    func = sub_folder or 'Celebration'

                songs.append({
                    'id': len(songs) + 1,
                    'title': title.strip(),
                    'artist': meta.get('singer') or artist.strip(),
                    'movie': meta.get('movie') or '',
                    'year': meta.get('year') or '',
                    'category': meta.get('category') or 'Wedding Hit',
                    'language': lang,
                    'day_key': day_key,
                    'day': day_label,
                    'function': func,
                    'folder': top_folder,
                    'subfolder': sub_folder,
                    'path': rel,
                    'size_mb': round(os.path.getsize(full) / (1024*1024), 2)
                })
    return songs

print("Metadata & Song scanner ready!")

