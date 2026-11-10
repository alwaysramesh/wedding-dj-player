import os

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# List of high-priority ceremony direct video links for HTML DJ Player
direct_songs = [
    # Haldi & Mayan
    {"day": "Day 1 (10 Dec)", "ceremony": "Haldi Ceremony", "name": "London Thumakda", "movie": "Queen", "singer": "Labh Janjua, Sonu Kakkar, Neha Kakkar", "url": "https://www.youtube.com/watch?v=udra3Mfw2oo"},
    {"day": "Day 1 (10 Dec)", "ceremony": "Haldi Ceremony", "name": "Navrai Majhi", "movie": "English Vinglish", "singer": "Sunidhi Chauhan", "url": "https://www.youtube.com/watch?v=uA0Lh6T1m2Y"},
    {"day": "Day 1 (10 Dec)", "ceremony": "Haldi Ceremony", "name": "Sasural Genda Phool", "movie": "Delhi-6", "singer": "Rekha Bhardwaj", "url": "https://www.youtube.com/watch?v=xV3fNnC94qA"},
    {"day": "Day 1 (10 Dec)", "ceremony": "Haldi Ceremony", "name": "Laal Peeli Akhiyaan", "movie": "Teri Baaton Mein", "singer": "Shadab Faridi", "url": "https://www.youtube.com/watch?v=8L9pX_m2_uE"},
    {"day": "Day 1 (10 Dec)", "ceremony": "Mayan & Tilak", "name": "Chaap Tilak Sab Cheeni", "movie": "Coke Studio", "singer": "Rahul Vaidya", "url": "https://www.youtube.com/watch?v=qXgq2wE8-w0"},
    {"day": "Day 1 (10 Dec)", "ceremony": "Mayan & Tilak", "name": "Chhote Chhote Bhaiyon Ke", "movie": "Hum Saath-Saath Hain", "singer": "Kavita K, Udit Narayan", "url": "https://www.youtube.com/watch?v=yeIeTRcI3II"},

    # Mehndi & Sangeet
    {"day": "Day 2 (11 Dec)", "ceremony": "Mehndi Ceremony", "name": "Mehndi Laga Ke Rakhna", "movie": "DDLJ", "singer": "Lata Mangeshkar, Udit Narayan", "url": "https://www.youtube.com/watch?v=cM75_vWkiAk"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Ladies Sangeet", "name": "Bole Chudiyan", "movie": "Kabhi Khushi Kabhie Gham", "singer": "Sonu Nigam, Alka Yagnik, Kavita K.", "url": "https://www.youtube.com/watch?v=IBvg3WeqP1U"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Couple Dance", "name": "Sajni", "movie": "Laapataa Ladies", "singer": "Arijit Singh", "url": "https://www.youtube.com/watch?v=AFTIVN8rRbI"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Couple Dance", "name": "Kesariya", "movie": "Brahmastra", "singer": "Arijit Singh", "url": "https://www.youtube.com/watch?v=BddP6PYo2gs"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Couple Dance", "name": "Tum Hi Ho", "movie": "Aashiqui 2", "singer": "Arijit Singh", "url": "https://www.youtube.com/watch?v=Umqb9KENgmk"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Couple Dance", "name": "Heeriye", "movie": "Single", "singer": "Jasleen Royal, Arijit Singh", "url": "https://www.youtube.com/watch?v=RLzC55ai0eo"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Family Dance", "name": "Gallan Goodiyaan", "movie": "Dil Dhadakne Do", "singer": "Shankar Mahadevan, Yashita", "url": "https://www.youtube.com/watch?v=yeIeTRcI3II"},
    {"day": "Day 2 (11 Dec)", "ceremony": "Friends Dance", "name": "What Jhumka?", "movie": "Rocky Aur Rani", "singer": "Arijit Singh, Jonita Gandhi", "url": "https://www.youtube.com/watch?v=A664n8Z_2FE"},
    {"day": "Day 2 (11 Dec)", "ceremony": "DJ Party", "name": "Tauba Tauba", "movie": "Bad Newz", "singer": "Karan Aujla", "url": "https://www.youtube.com/watch?v=HD3KxxLVCk4"},
    {"day": "Day 2 (11 Dec)", "ceremony": "DJ Party", "name": "Kala Chashma", "movie": "Baar Baar Dekho", "singer": "Amar Arshi, Badshah", "url": "https://www.youtube.com/watch?v=k4yXQkG2s1E"},
    {"day": "Day 2 (11 Dec)", "ceremony": "DJ Party", "name": "Aayi Nai", "movie": "Stree 2", "singer": "Sachin-Jigar, Divya Kumar", "url": "https://www.youtube.com/watch?v=nFgsBxw-zWQ"},

    # Baraat & Wedding
    {"day": "Day 3 (12 Dec)", "ceremony": "Baraat & Groom Entry", "name": "Malhari", "movie": "Bajirao Mastani", "singer": "Vishal Dadlani", "url": "https://www.youtube.com/watch?v=l_MyUGq7pgs"},
    {"day": "Day 3 (12 Dec)", "ceremony": "Baraat & Groom Entry", "name": "Dulhe Ka Sehra Suhana Lagta Hai", "movie": "Dhadkan", "singer": "Nusrat Fateh Ali Khan", "url": "https://www.youtube.com/watch?v=f8b7iNKT3uU"},
    {"day": "Day 3 (12 Dec)", "ceremony": "Baraat & Groom Entry", "name": "Aaj Mere Yaar Ki Shaadi Hai", "movie": "Aadmi Sadak Ka", "singer": "Mohammed Rafi", "url": "https://www.youtube.com/watch?v=H7yC_iG8jvg"},
    {"day": "Day 3 (12 Dec)", "ceremony": "Varmala / Jaimala", "name": "Kudmayi", "movie": "Rocky Aur Rani", "singer": "Shreya Ghoshal, Arijit Singh", "url": "https://www.youtube.com/watch?v=3qpxJEp4Ec4"},
    {"day": "Day 3 (12 Dec)", "ceremony": "Bride Entry", "name": "Din Shagna Da", "movie": "Phillauri", "singer": "Jasleen Royal", "url": "https://www.youtube.com/watch?v=1LVpuWpRs3I"},
    {"day": "Day 3 (12 Dec)", "ceremony": "Vidaai", "name": "Dilbaro", "movie": "Raazi", "singer": "Harshdeep Kaur, Vibha Saraf", "url": "https://www.youtube.com/watch?v=kMEfhcc6zxs"},

    # Regional Folk (Bhojpuri & Bagheli)
    {"day": "Regional", "ceremony": "Bhojpuri Haldi", "name": "Hardi Hardiya", "movie": "Shubh Vivah", "singer": "Sharda Sinha", "url": "https://www.youtube.com/watch?v=zCkngalbucU"},
    {"day": "Regional", "ceremony": "Bhojpuri Baraat", "name": "Lollypop Lagelu", "movie": "Classic Hit", "singer": "Pawan Singh", "url": "https://www.youtube.com/watch?v=Gr8G_ldltDE"},
    {"day": "Regional", "ceremony": "Bhojpuri Vidaai", "name": "Piya Ke Nagariya", "movie": "Dulhin", "singer": "Sharda Sinha", "url": "https://www.youtube.com/watch?v=R5fMMLMy7Bg"},
    {"day": "Regional", "ceremony": "Bhojpuri DJ", "name": "Panche Ke Nache Aiha", "movie": "Baraat Single", "singer": "Pawan Singh, Shilpi Raj", "url": "https://www.youtube.com/watch?v=0c7vj-y5QCs"},
    {"day": "Regional", "ceremony": "Bhojpuri DJ", "name": "Laal Ghaghra", "movie": "Party Single", "singer": "Pawan Singh, Shilpi Raj", "url": "https://www.youtube.com/watch?v=qZId59qml_4"},
    {"day": "Regional", "ceremony": "Bagheli Mayan", "name": "Avadhpuri Mora Dhaam", "movie": "Bagheli Folk", "singer": "Vandana Mishra", "url": "https://www.youtube.com/watch?v=AUZIYQE5Vl1r8ZL9XI-rbKUGdV8N0iuugKBVDa9TTXS9n9HF_7JD5-B_6cfjFa7crfM9vdez3ZYelqBEtRW6U6kUsQcrmGs3d_fM-esNgsmOVzC5dO3qNw7kNqgwwzRKDqxzsIzf"},
    {"day": "Regional", "ceremony": "Bagheli Ghodi", "name": "Ghodi Chadhe Mor Veer Kunwar", "movie": "Bagheli Vivah", "singer": "Rewa Folk Troupe", "url": "https://www.youtube.com/watch?v=AUZIYQFryDfAah0gMpcL0nsZe1NysZb92NLjfai6q8b9Zj2UcFMa__PUDze9NrTg3YQQiS9QTAs72AhP1BOuZAqfwaNOUAai1iI6DUBtGBPR_t_sSiRbR1v79ZrizUMaOBDdbp8C"},
]

# Generate 1-Click HTML Jukebox Player
html_content = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Wedding Music 1-Click Direct Player</title>
<style>
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 24px; }
  h1 { color: #f59e0b; font-size: 26px; margin-bottom: 4px; }
  p.subtitle { color: #94a3b8; font-size: 14px; margin-top: 0; margin-bottom: 24px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 16px; }
  .card { background: #1e293b; border-radius: 12px; padding: 16px; border: 1px solid #334155; transition: transform 0.15s, border-color 0.15s; }
  .card:hover { transform: translateY(-2px); border-color: #f59e0b; }
  .tag { display: inline-block; padding: 4px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; }
  .tag-day1 { background: #b45309; color: #fef3c7; }
  .tag-day2 { background: #047857; color: #d1fae5; }
  .tag-day3 { background: #b91c1c; color: #fee2e2; }
  .tag-reg { background: #6b21a8; color: #f3e8ff; }
  .song-title { font-size: 18px; font-weight: 700; color: #ffffff; text-decoration: none; display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
  .song-title:hover { color: #38bdf8; }
  .details { color: #94a3b8; font-size: 13px; line-height: 1.4; margin-bottom: 12px; }
  .play-btn { display: block; text-align: center; background: #e11d48; color: #ffffff; font-weight: bold; text-decoration: none; padding: 10px 14px; border-radius: 8px; font-size: 14px; transition: background 0.2s; }
  .play-btn:hover { background: #be123c; }
  .filter-bar { margin-bottom: 20px; }
  .filter-btn { background: #334155; color: #f8fafc; border: none; padding: 8px 16px; border-radius: 6px; cursor: pointer; margin-right: 8px; font-size: 13px; font-weight: 600; }
  .filter-btn:hover, .filter-btn.active { background: #f59e0b; color: #0f172a; }
</style>
</head>
<body>
<h1>Brother's Wedding - 1-Click Direct Play Music Player</h1>
<p class="subtitle">Click any song name or red button below to immediately start playing the official song video.</p>

<div class="grid">
"""

for s in direct_songs:
    tag_class = "tag-day1"
    if "Day 2" in s["day"]:
        tag_class = "tag-day2"
    elif "Day 3" in s["day"]:
        tag_class = "tag-day3"
    elif "Regional" in s["day"]:
        tag_class = "tag-reg"
        
    html_content += f"""
  <div class="card">
    <span class="tag {tag_class}">{s['day']} &bull; {s['ceremony']}</span>
    <a href="{s['url']}" target="_blank" class="song-title">
      <span>{s['name']}</span>
      <span style="font-size: 14px; color: #38bdf8;">&#9654; Play</span>
    </a>
    <div class="details">
      <div><strong>Movie/Album:</strong> {s['movie']}</div>
      <div><strong>Singer(s):</strong> {s['singer']}</div>
    </div>
    <a href="{s['url']}" target="_blank" class="play-btn">&#9658; Click Here to Direct Play Song</a>
  </div>
"""

html_content += """
</div>
</body>
</html>
"""

html_path = os.path.join(BASE_DIR, "Click_To_Direct_Play_Songs.html")
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Generated 1-Click Direct Play Player at: {html_path}")

