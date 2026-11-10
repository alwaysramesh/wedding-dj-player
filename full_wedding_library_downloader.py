import os
import re
import time
import openpyxl
import yt_dlp

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", text).strip()

def get_target_subfolder(day, func, lang):
    f_lower = func.lower()
    
    if "mayan" in f_lower or "tilak" in f_lower or "mandap" in f_lower:
        return os.path.join("10 December - Mayan & Haldi", "Traditional Mayan Songs")
    elif "haldi" in f_lower and "punjabi" in f_lower:
        return os.path.join("10 December - Mayan & Haldi", "Punjabi Haldi Songs")
    elif "haldi" in f_lower:
        return os.path.join("10 December - Mayan & Haldi", "Haldi Bollywood Songs")
    elif "family celebration" in f_lower:
        return os.path.join("10 December - Mayan & Haldi", "Family Celebration")
    elif "funny" in f_lower:
        return os.path.join("10 December - Mayan & Haldi", "Funny & Entertaining Songs")
    elif "mehndi" in f_lower:
        return os.path.join("11 December - Mehndi & Sangeet", "Mehndi Songs")
    elif "ladies sangeet" in f_lower:
        return os.path.join("11 December - Mehndi & Sangeet", "Ladies Sangeet")
    elif "family dance" in f_lower:
        return os.path.join("11 December - Mehndi & Sangeet", "Family Dance")
    elif "friends dance" in f_lower:
        return os.path.join("11 December - Mehndi & Sangeet", "Friends Dance")
    elif "couple dance" in f_lower:
        return os.path.join("11 December - Mehndi & Sangeet", "Couple Dance")
    elif "dj party" in f_lower:
        return os.path.join("11 December - Mehndi & Sangeet", "DJ Party")
    elif "baraat" in f_lower or "groom entry" in f_lower:
        return os.path.join("12 December - Wedding", "Baraat & Groom Entry")
    elif "ghodi" in f_lower or "procession" in f_lower:
        return os.path.join("12 December - Wedding", "Ghodi & Procession")
    elif "varmala" in f_lower or "jaimala" in f_lower:
        return os.path.join("12 December - Wedding", "Varmala & Jaimala")
    elif "bride entry" in f_lower:
        return os.path.join("12 December - Wedding", "Bride Entry")
    elif "ritual" in f_lower or "pheras" in f_lower:
        return os.path.join("12 December - Wedding", "Pheras Instrumental")
    elif "reception" in f_lower:
        return os.path.join("12 December - Wedding", "Reception")
    elif "vidaai" in f_lower:
        return os.path.join("12 December - Wedding", "Vidaai")
    
    return os.path.join("10 December - Mayan & Haldi", "Family Celebration")

def check_file_exists(target_dir, base_name):
    if not os.path.exists(target_dir):
        return False
    # Check if a file starting with the clean name exists and has a size > 100KB
    for f in os.listdir(target_dir):
        if f.lower().startswith(base_name.lower()) and any(f.lower().endswith(ext) for ext in [".mp3", ".webm", ".m4a", ".opus"]):
            full_path = os.path.join(target_dir, f)
            if os.path.getsize(full_path) > 102400:
                return True
    return False

def main():
    print("="*65)
    print("🎵 FULL WEDDING LIBRARY AUTO-DOWNLOADER (ALL SHEETS & FOLDERS)")
    print("="*65)

    wb = openpyxl.load_workbook(EXCEL_FILE)
    
    # Target sheets to process
    sheets_to_process = [
        "Master Song List (All Days)",
        "Bagheli Wedding Songs (Rewa)",
        "Bhojpuri Wedding Songs"
    ]
    
    tasks = []
    
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
        
        for idx, h in enumerate(headers, start=1):
            if "song" in h: song_col = idx
            elif "singer" in h or "artist" in h: singer_col = idx
            elif "movie" in h or "album" in h or "source" in h: movie_col = idx
            elif "ceremony" in h or "function" in h or "category" in h: cero_col = idx
            elif "day" in h: day_col = idx
            
        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(row=r, column=song_col).value or '').strip() if song_col else ""
            singer = str(ws.cell(row=r, column=singer_col).value or '').strip() if singer_col else ""
            movie = str(ws.cell(row=r, column=movie_col).value or '').strip() if movie_col else ""
            cero = str(ws.cell(row=r, column=cero_col).value or '').strip() if cero_col else ""
            day = str(ws.cell(row=r, column=day_col).value or '').strip() if day_col else ""
            
            if not name or name.lower() == "none":
                continue
                
            first_singer = singer.split(",")[0].split("&")[0].strip()
            raw_name = f"{name} - {first_singer}" if first_singer else name
            clean_name = clean_filename(raw_name)
            
            # Destination directory determination
            if sheet_name == "Bagheli Wedding Songs (Rewa)":
                if "mayan" in cero.lower() or "mandap" in cero.lower() or "tilak" in cero.lower():
                    sub = os.path.join("Bagheli Wedding Songs", "01 - Mayan & Mandap Pujan")
                elif "haldi" in cero.lower() or "nahchhu" in cero.lower():
                    sub = os.path.join("Bagheli Wedding Songs", "02 - Haldi & Nahchhu")
                elif "mehndi" in cero.lower() or "sangeet" in cero.lower():
                    sub = os.path.join("Bagheli Wedding Songs", "03 - Mehndi & Ladies Sangeet")
                elif "baraat" in cero.lower() or "ghodi" in cero.lower():
                    sub = os.path.join("Bagheli Wedding Songs", "04 - Baraat & Ghodi Chadai")
                else:
                    sub = os.path.join("Bagheli Wedding Songs", "05 - Pheras, Varmala & Bidaai")
            elif sheet_name == "Bhojpuri Wedding Songs":
                if "classic" in cero.lower() or "evergreen" in cero.lower() or "mayan" in cero.lower() or "haldi" in cero.lower():
                    sub = os.path.join("Bhojpuri Wedding Songs", "01 - Evergreen Classics")
                else:
                    sub = os.path.join("Bhojpuri Wedding Songs", "02 - Latest Trending Hits")
            else:
                sub = get_target_subfolder(day, cero, "")
                
            search_query = f"ytsearch1:{name} {movie} {first_singer} audio song".strip()
            tasks.append((sub, clean_name, search_query))

    print(f"Total songs queued across all sheets: {len(tasks)}")
    
    downloaded = 0
    already_present = 0
    errors = 0

    ydl_opts_base = {
        'format': 'bestaudio/best',
        'quiet': True,
        'no_warnings': True,
        'ignoreerrors': True,
        'retries': 2,
        'socket_timeout': 15,
        'default_search': 'ytsearch'
    }

    for idx, (folder_rel, song_filename, query) in enumerate(tasks, 1):
        target_dir = os.path.join(BASE_DIR, folder_rel)
        os.makedirs(target_dir, exist_ok=True)
        
        if check_file_exists(target_dir, song_filename):
            print(f"[{idx}/{len(tasks)}] ⏩ [EXISTS]: {song_filename}")
            already_present += 1
            continue
            
        print(f"[{idx}/{len(tasks)}] ⬇️ [DOWNLOADING]: {song_filename}...")
        out_template = os.path.join(target_dir, f"{song_filename}.%(ext)s")
        ydl_opts = dict(ydl_opts_base)
        ydl_opts['outtmpl'] = out_template
        
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([query])
            if check_file_exists(target_dir, song_filename):
                print(f"   ✅ [SUCCESS]: Saved {song_filename}")
                downloaded += 1
            else:
                print(f"   ⚠️ [SKIPPED/UNAVAILABLE]: {song_filename}")
                errors += 1
        except Exception as e:
            print(f"   ❌ [ERROR]: {song_filename} -> {e}")
            errors += 1
            
        time.sleep(0.3)

    print("\n" + "="*65)
    print("🎉 FULL DOWNLOAD EXECUTION FINISHED!")
    print(f"Already Existed:        {already_present}")
    print(f"Successfully Downloaded: {downloaded}")
    print(f"Unavailable/Skipped:     {errors}")
    print("="*65)

if __name__ == "__main__":
    main()

