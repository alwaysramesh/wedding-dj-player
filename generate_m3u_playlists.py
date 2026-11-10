import os
import json

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
HTML_PLAYER = os.path.join(BASE_DIR, "Wedding_Music_Player.html")

# Read songs database from Wedding_Music_Player.html
with open(HTML_PLAYER, "r", encoding="utf-8") as f:
    text = f.read()

start = text.find("const SONGS = ") + len("const SONGS = ")
end = text.find(";\n\n        // DJ Console State")
songs = json.loads(text[start:end])

# Create Playlists Directory
PL_DIR = os.path.join(BASE_DIR, "DJ Playlists (M3U)")
os.makedirs(PL_DIR, exist_ok=True)

# 1. Master All Songs Playlist
with open(os.path.join(PL_DIR, "00_Master_Wedding_Library_All_508_Songs.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 2. Day 1: Mayan & Haldi Playlist
with open(os.path.join(PL_DIR, "01_Day_1_Mayan_and_Haldi_Complete.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["day_key"] == "day1":
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 3. Day 2: Mehndi & Sangeet Playlist
with open(os.path.join(PL_DIR, "02_Day_2_Mehndi_and_Sangeet_Complete.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["day_key"] == "day2":
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 4. Day 3: Wedding Day & Baraat Playlist
with open(os.path.join(PL_DIR, "03_Day_3_Wedding_Baraat_and_Rituals_Complete.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["day_key"] == "day3":
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 5. Bagheli (Rewa) Folk Playlist
with open(os.path.join(PL_DIR, "04_Bagheli_Rewa_Traditional_Folk.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["day_key"] == "bagheli" or s["language"] == "Bagheli":
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 6. Bhojpuri Wedding Playlist
with open(os.path.join(PL_DIR, "05_Bhojpuri_Wedding_Celebrations.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["day_key"] == "bhojpuri" or s["language"] == "Bhojpuri":
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 7. DJ Dance & Baraat High Energy Nonstop Playlist
with open(os.path.join(PL_DIR, "06_DJ_Party_and_Baraat_Blast.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["function"] in ["DJ Party Hits", "Baraat & Groom Entry", "Ghodi & Procession", "Friends Dance"]:
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 8. Sacred Wedding Rituals & Pheras Playlist
with open(os.path.join(PL_DIR, "07_Sacred_Pheras_and_Rituals_Devotional.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["function"] in ["Wedding Rituals & Pheras", "Pheras Instrumental", "Varmala & Jaimala", "Mayan & Mandap"]:
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

# 9. Emotional Vidaai Playlist
with open(os.path.join(PL_DIR, "08_Vidaai_Emotional_and_Blessings.m3u"), "w", encoding="utf-8") as f:
    f.write("#EXTM3U\n")
    for s in songs:
        if s["function"] == "Vidaai":
            f.write(f"#EXTINF:-1,{s['artist']} - {s['title']}\n../{s['path']}\n")

print(f"✅ Generated 9 master .m3u DJ playlists in: {PL_DIR}")

