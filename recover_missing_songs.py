import os
import re
import yt_dlp

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# Specific, tested queries for the 27 missing songs
recovery_list = [
    # Bollywood & Hindi
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "Dil Chori - Yo Yo Honey Singh", "ytsearch1:Dil Chori Sonu Ke Titu Ki Sweety official video"),
    ("10 December - Mayan & Haldi\\Family Celebration", "Maahi Ve - Sonu Nigam", "ytsearch1:Maahi Ve Kal Ho Naa Ho official video"),
    ("12 December - Wedding\\Reception", "Ve Haaniyaan - Aarsh Qaamil", "ytsearch1:Ve Haaniyaan Ravi Dubey Sargun Mehta official song"),
    ("11 December - Mehndi & Sangeet\\Ladies Sangeet", "Chalka Chalka Re - Mahalakshmi Iyer", "ytsearch1:Chalka Chalka Re Saathiya official song"),
    ("12 December - Wedding\\Pheras Instrumental", "Raag Hansadhwani Shehnai - Bismillah Khan", "ytsearch1:Raag Hansadhwani Bismillah Khan Shehnai"),
    ("11 December - Mehndi & Sangeet\\Mehndi Songs", "Sone Ke Thaari Mein Jounar - Sharda Sinha", "ytsearch1:Sone Ke Thaari Mein Jounar Parosi Sharda Sinha"),
    ("10 December - Mayan & Haldi\\Traditional Mayan Songs", "Nindiya Ke Matal Dulha - Sharda Sinha", "ytsearch1:Nindiya Ke Matal Dulha Sharda Sinha Vivah Geet"),

    # Bagheli Wedding Songs
    ("Bagheli Wedding Songs\\01 - Mayan & Mandap Pujan", "Avadhpuri Mora Dhaam - Vandana Mishra", "ytsearch1:Avadhpuri Mora Dhaam Vandana Mishra Bagheli Vivah Geet"),
    ("Bagheli Wedding Songs\\01 - Mayan & Mandap Pujan", "Sat Seek Ke Pinjadwa - Bagheli Folk", "ytsearch1:Sat Seek Ke Pinjadwa Bagheli Banra folk song"),
    ("Bagheli Wedding Songs\\01 - Mayan & Mandap Pujan", "Mangal Baaje Baghelkhand Mein - Bagheli", "ytsearch1:Mangal Baaje Baghelkhand Mein vivah geet"),
    ("Bagheli Wedding Songs\\02 - Haldi & Nahchhu", "Nahchhu Gaavahu Sakhiyaan Mil Ke - Bagheli", "ytsearch1:Nahchhu Gaavahu Sakhiyaan Mil Ke Bagheli Geet"),
    ("Bagheli Wedding Songs\\02 - Haldi & Nahchhu", "Chauka Purao Dhani Maai - Bagheli", "ytsearch1:Chauka Purao Dhani Maai Bagheli vivah"),
    ("Bagheli Wedding Songs\\02 - Haldi & Nahchhu", "Banna Ke Ang Ang Hardi Sajai - Bagheli", "ytsearch1:Banna Ke Ang Ang Hardi Sajai Bagheli"),
    ("Bagheli Wedding Songs\\03 - Mehndi & Ladies Sangeet", "Haathan Mehndi Chamke Banna - Bagheli", "ytsearch1:Haathan Mehndi Chamke Banna Tore Naam Ki Bagheli"),
    ("Bagheli Wedding Songs\\03 - Mehndi & Ladies Sangeet", "Banno Rani Chamke Ri - Bagheli", "ytsearch1:Banno Rani Chamke Ri Bagheli song"),
    ("Bagheli Wedding Songs\\03 - Mehndi & Ladies Sangeet", "Gori Tore Mathe Bindiya - Bagheli", "ytsearch1:Gori Tore Mathe Bindiya Bagheli geet"),
    ("Bagheli Wedding Songs\\04 - Baraat & Ghodi Chadai", "Ghodi Chadhe Mor Veer Kunwar - Bagheli", "ytsearch1:Ghodi Chadhe Mor Veer Kunwar Bagheli Ghodi Geet"),
    ("Bagheli Wedding Songs\\04 - Baraat & Ghodi Chadai", "Vindhya Wala Dulha - Bagheli", "ytsearch1:Vindhya Wala Dulha Bagheli Vivah DJ song"),
    ("Bagheli Wedding Songs\\05 - Pheras, Varmala & Bidaai", "Varmala Pehirao Mor Banna - Bagheli", "ytsearch1:Varmala Pehirao Mor Banna Bagheli Geet"),
    ("Bagheli Wedding Songs\\05 - Pheras, Varmala & Bidaai", "Sindoor Daan Karahu Mor Raja - Bagheli", "ytsearch1:Sindoor Daan Karahu Mor Raja Bagheli"),
    ("Bagheli Wedding Songs\\05 - Pheras, Varmala & Bidaai", "Saat Phere Lehu Banna-Banni - Bagheli", "ytsearch1:Saat Phere Lehu Banna Banni Bagheli Vivah"),
    ("Bagheli Wedding Songs\\05 - Pheras, Varmala & Bidaai", "Babul Ka Ghar Chhode Dulhaniya - Bagheli", "ytsearch1:Babul Ka Ghar Chhode Dulhaniya Bagheli Bidaai Geet"),
    ("Bagheli Wedding Songs\\05 - Pheras, Varmala & Bidaai", "Sajani Bida Hoti Jaai - Bagheli", "ytsearch1:Sajani Bida Hoti Jaai Bagheli Bidaai"),

    # Bhojpuri Wedding Songs
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Uthahu Banna Bhor Bhaye Re - Kalpana", "ytsearch1:Uthahu Banna Bhor Bhaye Re Kalpana Patowary"),
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Duno Nayanwa Me Surma Lagawale - Manoj Tiwari", "ytsearch1:Duno Nayanwa Me Surma Lagawale Manoj Tiwari vivah geet"),
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Ghunghat Uthave Dulhaniya - Sharda Sinha", "ytsearch1:Ghunghat Uthave Dulhaniya Sharda Sinha vivah geet"),
    ("Bhojpuri Wedding Songs\\01 - Evergreen Classics", "Babul Ka Ghar Chhodi Chali - Manoj Tiwari", "ytsearch1:Babul Ka Ghar Chhodi Chali Manoj Tiwari bidaai geet"),
]

def check_file_exists(target_dir, base_name):
    if not os.path.exists(target_dir):
        return False
    prefix = base_name.split(" - ")[0].lower()[:15]
    for f in os.listdir(target_dir):
        f_lower = f.lower()
        if any(f_lower.endswith(ext) for ext in [".mp3", ".webm", ".m4a", ".opus"]):
            if prefix in f_lower:
                if os.path.getsize(os.path.join(target_dir, f)) > 102400:
                    return True
    return False

def main():
    print(f"🔄 Retrying {len(recovery_list)} targeted recovery songs...")
    recovered = 0
    ydl_opts_base = {
        'format': 'bestaudio/best',
        'quiet': True,
        'no_warnings': True,
        'ignoreerrors': True,
        'retries': 2,
        'socket_timeout': 15,
        'default_search': 'ytsearch'
    }

    for folder_rel, song_filename, query in recovery_list:
        target_dir = os.path.join(BASE_DIR, folder_rel)
        os.makedirs(target_dir, exist_ok=True)
        if check_file_exists(target_dir, song_filename):
            print(f"⏩ Already Exists: {song_filename}")
            recovered += 1
            continue

        print(f"⬇️ Downloading: {song_filename}...")
        ydl_opts = dict(ydl_opts_base)
        ydl_opts['outtmpl'] = os.path.join(target_dir, f"{song_filename}.%(ext)s")
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([query])
            if check_file_exists(target_dir, song_filename):
                print(f"✅ Recovered: {song_filename}")
                recovered += 1
            else:
                print(f"⚠️ Still unavailable: {song_filename}")
        except Exception as e:
            print(f"❌ Error: {song_filename} -> {e}")

    print(f"\nTargeted Recovery Complete: {recovered}/{len(recovery_list)} songs secured!")

if __name__ == "__main__":
    main()

