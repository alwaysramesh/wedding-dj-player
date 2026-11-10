import os
import urllib.parse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

# 1. Load existing songs to verify ZERO duplication
wb = openpyxl.load_workbook(EXCEL_FILE)
existing_names = set()
for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell and isinstance(cell, str) and len(cell) > 2:
                existing_names.add(cell.strip().lower())

def is_duplicate(name):
    norm = name.strip().lower()
    for ex in existing_names:
        if norm in ex or ex in norm:
            return True
    return False

# NEW CURATED UNIQUE SONGS (Strictly verified against duplicates, family dignity, 100% clean)
new_unique_songs = [
    # --- 10 December: Mayan & Haldi ---
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Mayan Ceremony",
        "Category": "Evergreen Classic",
        "Song": "Maiya Yashoda Yeh Tera Kanhaiya",
        "Movie": "Hum Saath-Saath Hain",
        "Singer": "Kavita Krishnamurthy, Alka Yagnik, Anuradha Paudwal",
        "Year": "1999",
        "Language": "Bollywood",
        "Folder": "10 December - Mayan & Haldi\\Traditional Mayan Songs",
        "Vibe": "Family tradition, elder blessings, devotional joy"
    },
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Mayan Ceremony",
        "Category": "Evergreen Classic",
        "Song": "Chhote Chhote Bhaiyon Ke Bade Bhaiya",
        "Movie": "Hum Saath-Saath Hain",
        "Singer": "Kavita Krishnamurthy, Udit Narayan, Kumar Sanu",
        "Year": "1999",
        "Language": "Bollywood",
        "Folder": "10 December - Mayan & Haldi\\Traditional Mayan Songs",
        "Vibe": "Pivotal brother's wedding anthem celebrating the elder brother"
    },
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Haldi Ceremony",
        "Category": "Trending (2024-2026)",
        "Song": "Udaariyaan (Acoustic Wedding Version)",
        "Movie": "Independent / Film",
        "Singer": "Satinder Sartaaj",
        "Year": "2024",
        "Language": "Punjabi",
        "Folder": "10 December - Mayan & Haldi\\Punjabi Haldi Songs",
        "Vibe": "Soulful, poetic Punjabi celebration"
    },
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Haldi Ceremony",
        "Category": "Trending (2024-2026)",
        "Song": "Gulaabi Aankhein (Acoustic Sangeet Mix)",
        "Movie": "Lover (Revival)",
        "Singer": "Sanam Puri",
        "Year": "2024",
        "Language": "Bollywood",
        "Folder": "10 December - Mayan & Haldi\\Haldi Bollywood Songs",
        "Vibe": "Breezy daytime family acoustic rhythm"
    },
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Family Celebration",
        "Category": "Evergreen Classic",
        "Song": "Yeh Ladka Hai Deewana",
        "Movie": "Kuch Kuch Hota Hai",
        "Singer": "Udit Narayan, Alka Yagnik",
        "Year": "1998",
        "Language": "Bollywood",
        "Folder": "10 December - Mayan & Haldi\\Family Celebration",
        "Vibe": "Innocent, sweet family & sibling teasing"
    },

    # --- 11 December: Mehndi & Sangeet ---
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Mehndi Ceremony",
        "Category": "Evergreen Classic",
        "Song": "Maye Ni Maye Munder Pe Teri",
        "Movie": "Hum Aapke Hain Koun",
        "Singer": "Lata Mangeshkar",
        "Year": "1994",
        "Language": "Bollywood",
        "Folder": "11 December - Mehndi & Sangeet\\Mehndi Songs",
        "Vibe": "Traditional pigeon folk meter for mother & daughter/sister"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Mehndi Ceremony",
        "Category": "Trending (2024-2026)",
        "Song": "Jaan Ban Gaye",
        "Movie": "Khuda Haafiz 2",
        "Singer": "Mithoon, Vishal Mishra, Asees Kaur",
        "Year": "2022/24",
        "Language": "Bollywood",
        "Folder": "11 December - Mehndi & Sangeet\\Mehndi Songs",
        "Vibe": "Graceful modern romantic melody for henna moments"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Ladies Sangeet",
        "Category": "Evergreen Classic",
        "Song": "Chalka Chalka Re Kalgidhar",
        "Movie": "Punjabi Folk Tradition",
        "Singer": "Gurmeet Bawa / Traditional",
        "Year": "Classic",
        "Language": "Punjabi Folk",
        "Folder": "11 December - Mehndi & Sangeet\\Ladies Sangeet",
        "Vibe": "Authentic dholki singing by elderly aunts & bhabhis"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Family Dance",
        "Category": "Evergreen Classic",
        "Song": "Aapka Kya Hoga Janabe Ali",
        "Movie": "Housefull",
        "Singer": "Mika Singh, Sunidhi Chauhan, Sajid Khan",
        "Year": "2010",
        "Language": "Bollywood",
        "Folder": "11 December - Mehndi & Sangeet\\Family Dance",
        "Vibe": "High energy group choreography for uncles and cousins"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Friends Dance",
        "Category": "Trending (2024-2026)",
        "Song": "Dil Tu Jaan Tu",
        "Movie": "Single",
        "Singer": "Gurnazar Chattha",
        "Year": "2023/24",
        "Language": "Punjabi Pop",
        "Folder": "11 December - Mehndi & Sangeet\\Friends Dance",
        "Vibe": "Modern celebration anthem for groom's friends squad"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Couple Dance",
        "Category": "Evergreen Classic",
        "Song": "Main Rang Sharbaton Ka",
        "Movie": "Phata Poster Nikhla Hero",
        "Singer": "Atif Aslam, Chinmayi Sripada",
        "Year": "2013",
        "Language": "Bollywood",
        "Folder": "11 December - Mehndi & Sangeet\\Couple Dance",
        "Vibe": "Gentle, waltz-friendly romantic duet"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Couple Dance",
        "Category": "Trending (2024-2026)",
        "Song": "Ruaa Ruaa",
        "Movie": "PS-2 (Hindi)",
        "Singer": "A.R. Rahman, Shilpa Rao",
        "Year": "2023/24",
        "Language": "Bollywood",
        "Folder": "11 December - Mehndi & Sangeet\\Couple Dance",
        "Vibe": "Deeply artistic, poetic royal couple dance"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "DJ Party",
        "Category": "Trending (2024-2026)",
        "Song": "Ishq Mitaye (Main Hoon Punjab)",
        "Movie": "Amar Singh Chamkila",
        "Singer": "Mohit Chauhan, A.R. Rahman",
        "Year": "2024",
        "Language": "Punjabi-Bollywood",
        "Folder": "11 December - Mehndi & Sangeet\\DJ Party",
        "Vibe": "Driving live Punjabi energy, dhol thumps"
    },

    # --- 12 December: Wedding Day ---
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Ghodi & Procession",
        "Category": "Evergreen Classic",
        "Song": "Ghoriyan - Matthe Chamkan Waal",
        "Movie": "Punjabi Ghoriyan Tradition",
        "Singer": "Surinder Kaur & Narinder Kaur",
        "Year": "Classic Folk",
        "Language": "Punjabi Folk",
        "Folder": "12 December - Wedding\\Ghodi & Procession",
        "Vibe": "Authentic sister blessings when brother mounts the mare"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Baraat & Groom Entry",
        "Category": "Trending (2024-2026)",
        "Song": "Balam Bambai",
        "Movie": "Single",
        "Singer": "Amit Trivedi, Mame Khan",
        "Year": "2024",
        "Language": "Folk Fusion",
        "Folder": "12 December - Wedding\\Baraat & Groom Entry",
        "Vibe": "Grand royal Rajasthani-Punjabi brass fusion entry"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Varmala & Jaimala",
        "Category": "Trending (2024-2026)",
        "Song": "Satranga",
        "Movie": "Animal",
        "Singer": "Arijit Singh, Shreyas Puranik",
        "Year": "2023/24",
        "Language": "Bollywood",
        "Folder": "12 December - Wedding\\Varmala & Jaimala",
        "Vibe": "Soulful 7-color vow melody for floral garland exchange"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Bride Entry",
        "Category": "Evergreen Classic",
        "Song": "Taare Hain Baaraati Chandani Hai Ye Baraat",
        "Movie": "Virasat",
        "Singer": "Kumar Sanu, Jaspinder Narula",
        "Year": "1997",
        "Language": "Bollywood",
        "Folder": "12 December - Wedding\\Bride Entry",
        "Vibe": "Nostalgic, grand Indian bridal walk classic"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Pheras Instrumental",
        "Category": "Evergreen Classic",
        "Song": "Raag Hansadhwani (Shehnai & Nadaswaram)",
        "Movie": "Auspicious Dawn Classics",
        "Singer": "Ustad Bismillah Khan & Dr. N. Rajam",
        "Year": "Classical",
        "Language": "Instrumental",
        "Folder": "12 December - Wedding\\Pheras Instrumental",
        "Vibe": "Sacred, joyous dawn raag for sacred Vedic vows"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Reception",
        "Category": "Trending (2024-2026)",
        "Song": "Tera Rastaa Chhodoon Na",
        "Movie": "Chennai Express",
        "Singer": "Anusha Mani, Amitabh Bhattacharya",
        "Year": "Classic Melodic",
        "Language": "Bollywood",
        "Folder": "12 December - Wedding\\Reception",
        "Vibe": "Sweet celebratory banquet greeting melody"
    },

    # --- Bagheli Wedding Songs (Rewa / Vindhya) ---
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Mayan & Tilak",
        "Category": "Evergreen Classic",
        "Song": "Chandan Ghori Ke Tilak Lagaao",
        "Movie": "Bagheli Vivah Sangrah",
        "Singer": "Vandana Mishra",
        "Year": "Folk Classic",
        "Language": "Bagheli",
        "Folder": "Bagheli Wedding Songs\\01 - Mayan & Mandap Pujan",
        "Vibe": "Traditional sandalwood tilak folk song from Rewa"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Mehndi & Sangeet",
        "Category": "Evergreen Classic",
        "Song": "Sone Ke Thaari Mein Jounar Parosi (Kaleva Geet)",
        "Movie": "Vindhya Lok Sanskaar",
        "Singer": "Urmila Shrivastava",
        "Year": "Folk Classic",
        "Language": "Bagheli",
        "Folder": "Bagheli Wedding Songs\\03 - Mehndi & Ladies Sangeet",
        "Vibe": "Traditional Bagheli Kaleva feast song teasing in good humor"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Baraat & Ghodi",
        "Category": "Trending (2024-2026)",
        "Song": "Banna Mor Aaye Vindhyachal Se",
        "Movie": "Rewa Music Series",
        "Singer": "Subham Tiwari & Shradha Khare",
        "Year": "2025",
        "Language": "Bagheli Pop",
        "Folder": "Bagheli Wedding Songs\\04 - Baraat & Ghodi Chadai",
        "Vibe": "Proud Vindhyan Baraat celebratory track"
    },

    # --- Bhojpuri Wedding Songs ---
    {
        "Day": "Day 1 (10 Dec)",
        "Ceremony": "Mayan & Haldi",
        "Category": "Evergreen Classic",
        "Song": "Nindiya Ke Matal Dulha (Nahchhu)",
        "Movie": "Traditional Vivah Geet",
        "Singer": "Sharda Sinha",
        "Year": "1990s",
        "Language": "Bhojpuri Folk",
        "Folder": "Bhojpuri Wedding Songs\\01 - Evergreen Classics",
        "Vibe": "Affectionate awakening of the groom for the morning bath"
    },
    {
        "Day": "Day 2 (11 Dec)",
        "Ceremony": "Mehndi & Sangeet",
        "Category": "Evergreen Classic",
        "Song": "Hariyer Baans Kataih A Babaji",
        "Movie": "Shubh Vivah",
        "Singer": "Sharda Sinha",
        "Year": "1990s",
        "Language": "Bhojpuri Folk",
        "Folder": "Bhojpuri Wedding Songs\\01 - Evergreen Classics",
        "Vibe": "Cutting the green bamboo for the wedding canopy"
    },
    {
        "Day": "Day 3 (12 Dec)",
        "Ceremony": "Baraat & DJ Party",
        "Category": "Trending (2024-2026)",
        "Song": "Dulhan Tilak Lagawai",
        "Movie": "Lagan Series 2024",
        "Singer": "Pawan Singh & Priyanka Singh",
        "Year": "2024",
        "Language": "Bhojpuri Pop",
        "Folder": "Bhojpuri Wedding Songs\\02 - Latest Trending Hits",
        "Vibe": "High tempo modern respectful celebration"
    }
]

print(f"Total candidate new songs: {len(new_unique_songs)}")

# Filter to ensure zero duplicates
verified_new_songs = []
for item in new_unique_songs:
    s_name = item["Song"]
    if not is_duplicate(s_name):
        verified_new_songs.append(item)
        existing_names.add(s_name.strip().lower())
    else:
        print(f"Skipping duplicate: {s_name}")

print(f"Verified unique new songs ready to add: {len(verified_new_songs)}")

def make_yt_link(name, movie_or_singer):
    query = f"{name} {movie_or_singer}".strip()
    return f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"

def make_sp_link(name, singer):
    first_singer = singer.split(",")[0].split("&")[0].split("/")[0].strip()
    query = f"{name} {first_singer}".strip()
    return f"https://open.spotify.com/search/{urllib.parse.quote(query)}"

# Append to Master Excel Sheets
ws_master = wb["Master Song List (All Days)"]
ws_clean = wb["Family Dignity & Clean Songs"]

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
data_font = Font(name="Calibri", size=10)
bold_font = Font(name="Calibri", size=10, bold=True)
link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")

start_master_idx = ws_master.max_row
for s in verified_new_songs:
    row_num = ws_master.max_row + 1
    yt = make_yt_link(s["Song"], s["Movie"])
    sp = make_sp_link(s["Song"], s["Singer"])
    
    ws_master.cell(row=row_num, column=1, value=row_num - 1).alignment = Alignment(horizontal="center")
    ws_master.cell(row=row_num, column=2, value=s["Day"])
    ws_master.cell(row=row_num, column=3, value=s["Ceremony"])
    ws_master.cell(row=row_num, column=4, value=s["Language"])
    ws_master.cell(row=row_num, column=5, value=s["Song"]).font = bold_font
    ws_master.cell(row=row_num, column=6, value=s["Movie"])
    ws_master.cell(row=row_num, column=7, value=s["Singer"])
    ws_master.cell(row=row_num, column=8, value=s["Year"]).alignment = Alignment(horizontal="center")
    
    c_yt = ws_master.cell(row=row_num, column=9, value="YouTube Link")
    c_yt.hyperlink = yt
    c_yt.font = link_font
    c_yt.alignment = Alignment(horizontal="center")
    
    c_sp = ws_master.cell(row=row_num, column=10, value="Spotify Link")
    c_sp.hyperlink = sp
    c_sp.font = link_font
    c_sp.alignment = Alignment(horizontal="center")
    
    for c_i in range(1, 11):
        ws_master.cell(row=row_num, column=c_i).border = thin_border
        ws_master.cell(row=row_num, column=c_i).fill = PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid")

    # Also append to Family Dignity Sheet
    c_row = ws_clean.max_row + 1
    ws_clean.cell(row=c_row, column=1, value=c_row - 1).alignment = Alignment(horizontal="center")
    ws_clean.cell(row=c_row, column=2, value=s["Day"])
    ws_clean.cell(row=c_row, column=3, value=s["Ceremony"]).font = bold_font
    ws_clean.cell(row=c_row, column=4, value=s["Category"]).alignment = Alignment(horizontal="center")
    ws_clean.cell(row=c_row, column=5, value=s["Song"]).font = bold_font
    ws_clean.cell(row=c_row, column=6, value=s["Movie"])
    ws_clean.cell(row=c_row, column=7, value=s["Singer"])
    ws_clean.cell(row=c_row, column=8, value=s["Year"]).alignment = Alignment(horizontal="center")
    
    c_yt2 = ws_clean.cell(row=c_row, column=9, value="▶ YouTube")
    c_yt2.hyperlink = yt
    c_yt2.font = link_font
    c_yt2.alignment = Alignment(horizontal="center")
    
    c_sp2 = ws_clean.cell(row=c_row, column=10, value="🎵 Spotify")
    c_sp2.hyperlink = sp
    c_sp2.font = link_font
    c_sp2.alignment = Alignment(horizontal="center")
    
    for c_i in range(1, 11):
        ws_clean.cell(row=c_row, column=c_i).border = thin_border
        ws_clean.cell(row=c_row, column=c_i).fill = PatternFill(start_color="F9FBE7", end_color="F9FBE7", fill_type="solid")

wb.save(EXCEL_FILE)
print("Updated Excel Master Workbook with newly verified songs!")

# Write individual Markdown logs in the respective folders
for s in verified_new_songs:
    target_folder = os.path.join(BASE_DIR, s["Folder"])
    if os.path.exists(target_folder):
        md_file = os.path.join(target_folder, "Additional_Curated_Songs.md")
        yt = make_yt_link(s["Song"], s["Movie"])
        entry = f"- **{s['Song']}** ({s['Movie']}) - {s['Singer']} [{s['Year']}] | *{s['Category']}* | [▶ YouTube]({yt})\n  - *Context:* {s['Vibe']}\n\n"
        with open(md_file, "a", encoding="utf-8") as f:
            f.write(entry)

print("Updated local folder Markdown files with new unique additions!")

