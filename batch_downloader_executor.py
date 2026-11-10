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
    
    if "mayan" in f_lower or "tilak" in f_lower:
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
    # Check if any audio file with base_name exists
    if not os.path.exists(target_dir):
        return False
    for f in os.listdir(target_dir):
        if f.lower().startswith(base_name.lower()) and any(f.lower().endswith(ext) for ext in [".mp3", ".webm", ".m4a", ".opus"]):
            return True
    return False

def main():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Master Song List (All Days)"]
    
    total = ws.max_row - 1
    print(f"Total songs to process: {total}")
    
    downloaded_count = 0
    skipped_count = 0
    failed_count = 0
    
    for r in range(2, ws.max_row + 1):
        day = str(ws.cell(row=r, column=2).value or "").strip()
        func = str(ws.cell(row=r, column=3).value or "").strip()
        lang = str(ws.cell(row=r, column=4).value or "").strip()
        name = str(ws.cell(row=r, column=5).value or "").strip()
        singer = str(ws.cell(row=r, column=7).value or "").strip()
        
        c_yt = ws.cell(row=r, column=9)
        yt_url = c_yt.hyperlink.target if (c_yt and c_yt.hyperlink) else ""
        
        if not name or not yt_url:
            continue
            
        first_singer = singer.split(",")[0].split("&")[0].strip()
        raw_name = f"{name} - {first_singer}" if first_singer else name
        clean_name = clean_filename(raw_name)
        
        subfolder = get_target_subfolder(day, func, lang)
        target_dir = os.path.join(BASE_DIR, subfolder)
        os.makedirs(target_dir, exist_ok=True)
        
        if check_file_exists(target_dir, clean_name):
            print(f"[{r-1}/{total}] ⏩ SKIPPED (Exists): {clean_name}")
            skipped_count += 1
            continue
            
        print(f"[{r-1}/{total}] 📥 DOWNLOADING: {clean_name} from {yt_url}")
        
        out_template = os.path.join(target_dir, f"{clean_name}.%(ext)s")
        ydl_opts = {
            'format': 'bestaudio/best',
            'outtmpl': out_template,
            'quiet': True,
            'no_warnings': True,
            'ignoreerrors': True,
            'retries': 3,
            'socket_timeout': 15
        }
        
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([yt_url])
            if check_file_exists(target_dir, clean_name):
                print(f"   ✅ SUCCESS: {clean_name}")
                downloaded_count += 1
            else:
                print(f"   ⚠️ FAILED: {clean_name}")
                failed_count += 1
        except Exception as e:
            print(f"   ❌ ERROR: {clean_name} -> {e}")
            failed_count += 1
            
        time.sleep(0.5)

    print("\n" + "="*50)
    print(f"DOWNLOAD PROCESS COMPLETE")
    print(f"Successfully Downloaded: {downloaded_count}")
    print(f"Already Existed:        {skipped_count}")
    print(f"Failed / Unavailable:   {failed_count}")
    print("="*50)

if __name__ == "__main__":
    main()

