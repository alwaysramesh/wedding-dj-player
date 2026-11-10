import os
import re
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
import yt_dlp

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

print_lock = threading.Lock()

def safe_print(*args, **kwargs):
    with print_lock:
        print(*args, **kwargs)

def clean_filename(text):
    clean = re.sub(r'[\\/*?:"<>|]', "", text).strip()
    return clean[:120]

def get_target_subfolder(sheet_name, day, func, cat):
    f = (func or '').lower()
    c = (cat or '').lower()
    d = (day or '').lower()
    
    if sheet_name == 'Bagheli Wedding Songs (Rewa)':
        if 'mayan' in f or 'mandap' in f or 'tilak' in f:
            return os.path.join("Bagheli Wedding Songs", "01 - Mayan & Mandap Pujan")
        elif 'haldi' in f or 'nahchhu' in f:
            return os.path.join("Bagheli Wedding Songs", "02 - Haldi & Nahchhu")
        elif 'mehndi' in f or 'sangeet' in f:
            return os.path.join("Bagheli Wedding Songs", "03 - Mehndi & Ladies Sangeet")
        elif 'baraat' in f or 'ghodi' in f:
            return os.path.join("Bagheli Wedding Songs", "04 - Baraat & Ghodi Chadai")
        else:
            return os.path.join("Bagheli Wedding Songs", "05 - Pheras, Varmala & Bidaai")
            
    if sheet_name == 'Bhojpuri Wedding Songs':
        if 'classic' in c or 'evergreen' in c:
            return os.path.join("Bhojpuri Wedding Songs", "01 - Evergreen Classics")
        else:
            return os.path.join("Bhojpuri Wedding Songs", "02 - Latest Trending Hits")
            
    # Master Song List / Days
    if '10 dec' in d or 'mayan' in f or 'haldi' in f:
        if 'mayan' in f or 'tilak' in f:
            return os.path.join("10 December - Mayan & Haldi", "Traditional Mayan Songs")
        elif 'punjabi' in f or 'punjabi' in c:
            return os.path.join("10 December - Mayan & Haldi", "Punjabi Haldi Songs")
        elif 'funny' in f:
            return os.path.join("10 December - Mayan & Haldi", "Funny & Entertaining Songs")
        elif 'family celebration' in f:
            return os.path.join("10 December - Mayan & Haldi", "Family Celebration")
        else:
            return os.path.join("10 December - Mayan & Haldi", "Haldi Bollywood Songs")
            
    if '11 dec' in d or 'mehndi' in f or 'sangeet' in f:
        if 'couple' in f:
            return os.path.join("11 December - Mehndi & Sangeet", "Couple Dance")
        elif 'dj' in f:
            return os.path.join("11 December - Mehndi & Sangeet", "DJ Party")
        elif 'family' in f:
            return os.path.join("11 December - Mehndi & Sangeet", "Family Dance")
        elif 'friends' in f:
            return os.path.join("11 December - Mehndi & Sangeet", "Friends Dance")
        elif 'ladies' in f:
            return os.path.join("11 December - Mehndi & Sangeet", "Ladies Sangeet")
        elif 'traditional' in f:
            return os.path.join("11 December - Mehndi & Sangeet", "Traditional Wedding Songs")
        else:
            return os.path.join("11 December - Mehndi & Sangeet", "Mehndi Songs")
            
    if '12 dec' in d or 'wedding' in d or 'baraat' in f:
        if 'ghodi' in f or 'procession' in f:
            return os.path.join("12 December - Wedding", "Ghodi & Procession")
        elif 'baraat' in f or 'groom entry' in f:
            return os.path.join("12 December - Wedding", "Baraat & Groom Entry")
        elif 'bride entry' in f:
            return os.path.join("12 December - Wedding", "Bride Entry")
        elif 'varmala' in f or 'jaimala' in f:
            return os.path.join("12 December - Wedding", "Varmala & Jaimala")
        elif 'ritual' in f:
            return os.path.join("12 December - Wedding", "Wedding Rituals")
        elif 'pheras' in f or 'instrumental' in f:
            return os.path.join("12 December - Wedding", "Pheras Instrumental")
        elif 'reception' in f:
            return os.path.join("12 December - Wedding", "Reception")
        elif 'vidaai' in f:
            return os.path.join("12 December - Wedding", "Vidaai")
            
    return os.path.join("10 December - Mayan & Haldi", "Family Celebration")

def check_file_exists(target_dir, base_name):
    if not os.path.exists(target_dir):
        return False
    # Check if a file starting with first 15 chars or matching name exists and > 100KB
    prefix = clean_filename(base_name.split(" - ")[0]).lower()[:15]
    for f in os.listdir(target_dir):
        f_lower = f.lower()
        if any(f_lower.endswith(ext) for ext in [".mp3", ".webm", ".m4a", ".opus"]):
            if base_name.lower() in f_lower or f_lower.startswith(prefix):
                full_path = os.path.join(target_dir, f)
                try:
                    if os.path.getsize(full_path) > 102400:
                        return True
                except:
                    pass
    return False

def download_task(item, idx, total):
    rel_folder, song_filename, query = item
    target_dir = os.path.join(BASE_DIR, rel_folder)
    os.makedirs(target_dir, exist_ok=True)
    
    if check_file_exists(target_dir, song_filename):
        safe_print(f"[{idx}/{total}] ⏩ [ALREADY EXISTS]: {song_filename} in {rel_folder}")
        return "EXISTS", song_filename
        
    safe_print(f"[{idx}/{total}] ⬇️ [DOWNLOADING]: {song_filename} -> {rel_folder}")
    out_template = os.path.join(target_dir, f"{song_filename}.%(ext)s")
    
    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': out_template,
        'quiet': True,
        'no_warnings': True,
        'ignoreerrors': True,
        'retries': 2,
        'socket_timeout': 15,
        'default_search': 'ytsearch'
    }
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([query])
            
        if check_file_exists(target_dir, song_filename):
            safe_print(f"[{idx}/{total}] ✅ [SUCCESS]: {song_filename}")
            return "SUCCESS", song_filename
        else:
            safe_print(f"[{idx}/{total}] ⚠️ [UNAVAILABLE/SKIPPED]: {song_filename}")
            return "FAILED", song_filename
    except Exception as e:
        safe_print(f"[{idx}/{total}] ❌ [ERROR]: {song_filename} -> {e}")
        return "ERROR", song_filename

def main():
    safe_print("="*70)
    safe_print("🎵 DOWNLOADING ENTIRE WEDDING MUSIC COLLECTION (ALL FOLDERS)")
    safe_print("="*70)

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    sheets_to_process = [
        "Master Song List (All Days)",
        "Bagheli Wedding Songs (Rewa)",
        "Bhojpuri Wedding Songs"
    ]
    
    queue = []
    seen = set()
    
    for sheet_name in sheets_to_process:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [str(cell.value or '').strip().lower() for cell in ws[1]]
        
        song_col = None
        singer_col = None
        movie_col = None
        cero_col = None
        day_col = None
        cat_col = None
        yt_col = None
        
        for idx, h in enumerate(headers, start=1):
            if "song" in h: song_col = idx
            elif "singer" in h or "artist" in h: singer_col = idx
            elif "movie" in h or "album" in h or "source" in h: movie_col = idx
            elif "ceremony" in h or "function" in h: cero_col = idx
            elif "category" in h or "type" in h: cat_col = idx
            elif "day" in h: day_col = idx
            elif "youtube" in h or "video" in h or "link" in h: yt_col = idx

        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(row=r, column=song_col).value or '').strip() if song_col else ""
            singer = str(ws.cell(row=r, column=singer_col).value or '').strip() if singer_col else ""
            movie = str(ws.cell(row=r, column=movie_col).value or '').strip() if movie_col else ""
            cero = str(ws.cell(row=r, column=cero_col).value or '').strip() if cero_col else ""
            cat = str(ws.cell(row=r, column=cat_col).value or '').strip() if cat_col else ""
            day = str(ws.cell(row=r, column=day_col).value or '').strip() if day_col else ""
            
            # YouTube Link from hyperlink if present
            yt_link = ""
            if yt_col:
                cell = ws.cell(row=r, column=yt_col)
                if cell.hyperlink and cell.hyperlink.target:
                    target_url = cell.hyperlink.target
                    if "watch?v=" in target_url:
                        yt_link = target_url
            
            if not name or name.lower() == "none":
                continue
                
            first_singer = singer.split(",")[0].split("&")[0].split("/")[0].strip()
            raw_name = f"{name} - {first_singer}" if first_singer else name
            clean_name = clean_filename(raw_name)
            
            folder_rel = get_target_subfolder(sheet_name, day, cero, cat)
            key = (folder_rel, clean_name.lower())
            if key in seen:
                continue
            seen.add(key)
            
            # If valid direct watch link exists, prefer it, else search
            if yt_link:
                query = yt_link
            else:
                search_terms = f"{name} {movie} {first_singer} audio song".strip()
                query = f"ytsearch1:{search_terms}"
                
            queue.append((folder_rel, clean_name, query))

    total = len(queue)
    safe_print(f"Total Unique Songs Queued across all folders: {total}")
    
    # Check already existing
    already_done = 0
    to_download = []
    for item in queue:
        rel_folder, song_filename, query = item
        target_dir = os.path.join(BASE_DIR, rel_folder)
        if check_file_exists(target_dir, song_filename):
            already_done += 1
        else:
            to_download.append(item)
            
    safe_print(f"Already Downloaded: {already_done}")
    safe_print(f"Remaining to Download: {len(to_download)}")
    safe_print("="*70)
    
    success_count = already_done
    fail_count = 0
    
    # Process with 3 concurrent worker threads
    start_time = time.time()
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = {executor.submit(download_task, item, idx, len(to_download)): item for idx, item in enumerate(to_download, 1)}
        for future in as_completed(futures):
            res, s_name = future.result()
            if res in ["SUCCESS", "EXISTS"]:
                success_count += 1
            else:
                fail_count += 1

    elapsed = time.time() - start_time
    safe_print("\n" + "="*70)
    safe_print("🎉 ALL FOLDERS DOWNLOAD CYCLE COMPLETED!")
    safe_print(f"Total Collection: {total}")
    safe_print(f"Successfully Present/Downloaded: {success_count}")
    safe_print(f"Failed / Unavailable: {fail_count}")
    safe_print(f"Total Time Taken: {elapsed/60:.2f} minutes")
    safe_print("="*70)

if __name__ == "__main__":
    main()

