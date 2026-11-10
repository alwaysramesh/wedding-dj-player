import os
import re
import yt_dlp

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# Curated High-Priority Core Wedding Songs with Exact Official Video IDs
curated_official_downloads = [
    # Haldi & Sangeet Hits
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "London Thumakda - Labh Janjua", "https://www.youtube.com/watch?v=udra3Mfw2oo"),
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "Navrai Majhi - Sunidhi Chauhan", "https://www.youtube.com/watch?v=uA0Lh6T1m2Y"),
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "Sasural Genda Phool - Rekha Bhardwaj", "https://www.youtube.com/watch?v=xV3fNnC94qA"),
    ("11 December - Mehndi & Sangeet\\Mehndi Songs", "Mehndi Laga Ke Rakhna - Lata Mangeshkar", "https://www.youtube.com/watch?v=cM75_vWkiAk"),
    ("11 December - Mehndi & Sangeet\\Ladies Sangeet", "Bole Chudiyan - Sonu Nigam", "https://www.youtube.com/watch?v=IBvg3WeqP1U"),
    ("11 December - Mehndi & Sangeet\\Couple Dance", "Kesariya - Arijit Singh", "https://www.youtube.com/watch?v=BddP6PYo2gs"),
    ("11 December - Mehndi & Sangeet\\Family Dance", "Gallan Goodiyaan - Shankar Mahadevan", "https://www.youtube.com/watch?v=yeIeTRcI3II"),
    ("11 December - Mehndi & Sangeet\\Friends Dance", "What Jhumka - Arijit Singh", "https://www.youtube.com/watch?v=A664n8Z_2FE"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Tauba Tauba - Karan Aujla", "https://www.youtube.com/watch?v=HD3KxxLVCk4"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Kala Chashma - Badshah", "https://www.youtube.com/watch?v=k4yXQkG2s1E"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Aayi Nai - Divya Kumar", "https://www.youtube.com/watch?v=nFgsBxw-zWQ"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Banno - Brijesh Shandilya", "https://www.youtube.com/watch?v=6yqalAdcMHE"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Sauda Khara Khara - Diljit Dosanjh", "https://www.youtube.com/watch?v=LYEqeUr-158"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Morni Banke - Guru Randhawa", "https://www.youtube.com/watch?v=1EadhOBcfI0"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Badtameez Dil - Benny Dayal", "https://www.youtube.com/watch?v=II2EO3Nw4m0"),

    # Baraat & Wedding Core
    ("12 December - Wedding\\Baraat & Groom Entry", "Malhari - Vishal Dadlani", "https://www.youtube.com/watch?v=l_MyUGq7pgs"),
    ("12 December - Wedding\\Baraat & Groom Entry", "Dulhe Ka Sehra - Nusrat Fateh Ali Khan", "https://www.youtube.com/watch?v=f8b7iNKT3uU"),
    ("12 December - Wedding\\Baraat & Groom Entry", "Mundian To Bach Ke - Panjabi MC", "https://www.youtube.com/watch?v=x9WO2ieJMYk"),
    ("12 December - Wedding\\Varmala & Jaimala", "Kudmayi - Arijit Singh", "https://www.youtube.com/watch?v=3qpxJEp4Ec4"),
    ("12 December - Wedding\\Bride Entry", "Din Shagna Da - Jasleen Royal", "https://www.youtube.com/watch?v=1LVpuWpRs3I"),
    ("12 December - Wedding\\Vidaai", "Dilbaro - Harshdeep Kaur", "https://www.youtube.com/watch?v=kMEfhcc6zxs"),

    # Regional Bhojpuri & Folk
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Hardi Hardiya - Sharda Sinha", "https://www.youtube.com/watch?v=zCkngalbucU"),
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Lollypop Lagelu - Pawan Singh", "https://www.youtube.com/watch?v=Gr8G_ldltDE"),
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Piya Ke Nagariya - Sharda Sinha", "https://www.youtube.com/watch?v=R5fMMLMy7Bg"),
    ("Bhojpuri Wedding Songs\\02 - Latest Trending Hits", "Panche Ke Nache Aiha - Pawan Singh", "https://www.youtube.com/watch?v=0c7vj-y5QCs"),
    ("Bhojpuri Wedding Songs\\02 - Latest Trending Hits", "Laal Ghaghra - Pawan Singh", "https://www.youtube.com/watch?v=qZId59qml_4"),
]

def main():
    print(f"🚀 Starting high-speed download for {len(curated_official_downloads)} core wedding hits...")
    
    success_count = 0
    for folder, name, url in curated_official_downloads:
        target_dir = os.path.join(BASE_DIR, folder)
        os.makedirs(target_dir, exist_ok=True)
        
        # Check if already exists
        exists = any(f.startswith(name) for f in os.listdir(target_dir)) if os.path.exists(target_dir) else False
        if exists:
            print(f"⏩ ALREADY EXISTS: {name}")
            success_count += 1
            continue
            
        print(f"⬇️ DOWNLOADING: {name}...")
        ydl_opts = {
            'format': 'bestaudio/best',
            'outtmpl': os.path.join(target_dir, f"{name}.%(ext)s"),
            'quiet': True,
            'no_warnings': True,
            'retries': 2
        }
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])
            print(f"✅ SAVED: {name}")
            success_count += 1
        except Exception as e:
            print(f"❌ FAILED: {name} -> {e}")

    print("\n" + "="*50)
    print(f"Finished: {success_count}/{len(curated_official_downloads)} songs ready offline in your folders!")
    print("="*50)

if __name__ == "__main__":
    main()

