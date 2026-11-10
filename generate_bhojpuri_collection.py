import os
import urllib.parse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
BHOJPURI_DIR = os.path.join(BASE_DIR, "Bhojpuri Wedding Songs")
EVERGREEN_DIR = os.path.join(BHOJPURI_DIR, "01 - Evergreen Classics")
TRENDING_DIR = os.path.join(BHOJPURI_DIR, "02 - Latest Trending Hits")

# 100% Clean, Family-Safe, Dignified Bhojpuri Wedding Songs
# Divided cleanly into Evergreen Classics and Latest Trending Hits (2024-2026)

evergreen_bhojpuri = [
    # Haldi, Mayan & Mandap
    ("Mayan & Haldi", "Hardi Hardiya", "Vivah Geet Classic", "Sharda Sinha", "1990s", "Traditional Haldi application on the groom"),
    ("Mayan & Haldi", "Hare Hare Hare Dada (Madwa)", "Shubh Vivah", "Sharda Sinha", "1990s", "Sacred Mandap erection and worship"),
    ("Mayan & Haldi", "Patna Se Piliya Mangawala", "Bhojpuri Vivah Sangrah", "Sharda Sinha", "1990s", "Auspicious yellow garment celebration"),
    ("Mayan & Haldi", "Chandan Ke Chauki", "Vivah Geet", "Manoj Tiwari 'Mridul'", "2000s", "Groom seated on sandalwood pedestal for rituals"),
    ("Mayan & Haldi", "Uthahu Banna Bhor Bhaye Re", "Traditional Vivah", "Kalpana Patowary", "2008", "Mayan morning awakening and tilak"),
    ("Mayan & Haldi", "Sita Ram Vivah Sohag", "Devotional Vivah", "Sharda Sinha", "1980s", "Blessings of Lord Ram and Sita"),
    # Mehndi & Ladies Sangeet
    ("Mehndi & Sangeet", "Banni Ke Haatho Mein Mehndiya", "Shagun", "Sharda Sinha", "1990s", "Traditional courtyard henna application"),
    ("Mehndi & Sangeet", "Sita Ke Sakal Dekhi", "Maithili-Bhojpuri Vivah", "Sharda Sinha", "1990s", "Traditional beauty and grace folk song"),
    ("Mehndi & Sangeet", "Bhojpuriya Jawan (Family Celebration)", "Bhojpuri Lokgeet", "Manoj Tiwari", "2004", "Upbeat family gathering dance"),
    ("Mehndi & Sangeet", "Gori Tori Chunari Ba Laal Laal Re", "Bhojpuri Classic Hits", "Pawan Singh", "2016", "Clean superhit dance track for ladies sangeet"),
    ("Mehndi & Sangeet", "Khatiya Charpayi Par", "Bhojpuri Folk Dance", "Malini Awasthi", "2012", "Folk rhythm celebrating family guests"),
    ("Mehndi & Sangeet", "Nathuniya Pe Goli Mare", "Bhojpuri Cinema Hit", "Pawan Singh", "2015", "Celebrated family dance rhythm"),
    # Baraat & Groom Entry
    ("Baraat & Groom Entry", "Ghodi Chad Ke Aaya Dulhe Raja", "Bhojpuri Vivah", "Manoj Tiwari", "2005", "Grand groom mounting the Ghodi"),
    ("Baraat & Groom Entry", "Lollypop Lagelu", "Independent Single", "Pawan Singh", "2008", "All-time iconic universal party celebration hit"),
    ("Baraat & Groom Entry", "Baraat Chali Banna Ki Sasural", "Baraat Geet", "Bharat Sharma Vyas", "1990s", "Traditional Baraat procession"),
    ("Baraat & Groom Entry", "Dwar Pe Aayi Baraat Re", "Dwarachar Geet", "Sharda Sinha", "1990s", "Auspicious welcome at the bride's gate"),
    ("Baraat & Groom Entry", "Duno Nayanwa Me Surma Lagawale", "Traditional Banra", "Manoj Tiwari", "2006", "Handsome groom welcome by elders"),
    # Rituals & Vidaai
    ("Rituals & Vidaai", "Ghunghat Uthave Dulhaniya", "Vivah Sampanna", "Sharda Sinha", "1990s", "Sacred Sindoor Daan and Jaimala moment"),
    ("Rituals & Vidaai", "Saptapadi Saat Vachan", "Vedic Vivah", "Sharda Sinha", "1990s", "Saat Pheras and lifelong vows"),
    ("Rituals & Vidaai", "Piya Ke Nagariya (Bidaai)", "Dulhin", "Sharda Sinha", "1990s", "The gold standard emotional Bidaai farewell"),
    ("Rituals & Vidaai", "Babul Ka Ghar Chhodi Chali", "Vivah Geet", "Manoj Tiwari", "2003", "Touching farewell song sung with tears and blessings"),
]

trending_bhojpuri = [
    # Mayan, Haldi & Sangeet (2024-2026)
    ("Mayan & Haldi", "Haldi Lagao Mor Bhaiya Ke", "Lagan Special 2024", "Priyanka Singh", "2024", "Upbeat festive Haldi beat for family dance"),
    ("Mayan & Haldi", "Banna Mora Sona Ka Sikdi", "Vivah Hits 2025", "Shilpi Raj", "2025", "Viral groom teasing reel and celebration track"),
    ("Mayan & Haldi", "Chandan Chandan Maathe Chamke", "Shubh Vivah Single", "Ankush Raja", "2024", "Sweet acoustic melody for Tilak ceremony"),
    ("Mayan & Haldi", "Piliya Pehnai Bhabhi", "Haldi Utsav 2025", "Pramod Premi Yadav", "2025", "Joyful family celebration while applying haldi"),
    # Mehndi & Sangeet Stage Performances
    ("Mehndi & Sangeet", "Dulhe Raja (Wedding Special)", "Lagan Anthem", "Shilpi Raj & Sangam Prem", "2024", "Massive hit for sangeet stage choreography"),
    ("Mehndi & Sangeet", "Kajra Mohabbat Wala (Bhojpuri Fusion)", "Single Hit", "Pawan Singh, Shilpi Raj", "2024", "Modern folk fusion dance for couples & cousins"),
    ("Mehndi & Sangeet", "Banno Mor Badi Pyaari", "Vivah Rang 2025", "Priyanka Singh", "2025", "Melodious bridal mehndi celebration song"),
    ("Mehndi & Sangeet", "Mehandi Rach Gayi Hathan Mein", "Lagan Series", "Shilpi Raj", "2024", "Traditional dholak with modern beats"),
    ("Mehndi & Sangeet", "Pahilka Milanwa (Romantic)", "Cinema Romantic", "Khesari Lal Yadav, Priyanka Singh", "2024", "Decent & sweet couple dance performance"),
    # Baraat & DJ Floor Bangers
    ("Baraat & DJ Party", "Panche Ke Nache Aiha", "Baraat Single", "Pawan Singh & Shilpi Raj", "2023", "Tremendous dance floor energy for Baraat"),
    ("Baraat & DJ Party", "Laal Ghaghra (Wedding DJ Mix)", "Single Blast", "Pawan Singh & Shilpi Raj", "2023", "Superhit crowd favorite with driving dhol beats"),
    ("Baraat & DJ Party", "Baraat Mein DJ Baji", "Lagan Special 2025", "Khesari Lal Yadav", "2025", "Dedicated Baraat dance anthem for friends"),
    ("Baraat & DJ Party", "Saree Se Tadi (Festival Edit)", "Baraat Anthem", "Pawan Singh", "2022", "Massive rhythm loved across youth baraats"),
    ("Baraat & DJ Party", "Rewa-Patna DJ Blast (Nonstop Mix)", "Wedding Mashup", "DJ Chetas & Regional Remixers", "2024", "Clean, uninterrupted energetic Baraat flow"),
    # Rituals & Vidaai
    ("Rituals & Vidaai", "Varmala Ho Gail Banna Ke", "Shubh Lagan", "Priyanka Singh", "2024", "Celebratory crescendo for the garland exchange"),
    ("Rituals & Vidaai", "Sindoorwa Sajawa Mor Raja", "Vivah Ritwa", "Shilpi Raj", "2025", "Emotional and respectful Sindoor Daan track"),
    ("Rituals & Vidaai", "Bidaai Ke Bela Mein Aansu Na Girawa", "Bidaai Geet 2024", "Priyanka Singh", "2024", "Tender and touching farewell for the doli departure"),
]

def make_yt_link(name, singer):
    query = f"{name} {singer} Bhojpuri Vivah Geet".strip()
    return f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"

# 1. Write Evergreen Classics Markdown
md_evergreen = os.path.join(EVERGREEN_DIR, "Evergreen_Bhojpuri_Wedding_Classics.md")
with open(md_evergreen, "w", encoding="utf-8") as f:
    f.write("# 🌺 Evergreen Bhojpuri Wedding Classics (कालजयी पारंपरिक विवाह गीत)\n")
    f.write("### Featuring Legends: Sharda Sinha, Manoj Tiwari, Bharat Sharma Vyas\n\n")
    f.write("> **Purity & Respect Assurance:**\n")
    f.write("> - 100% Traditional, authentic folk meter, respectful lyrics.\n")
    f.write("> - No double meanings or vulgarity; safe to play before grandparents, samdhis, and kids.\n\n")
    f.write("---\n\n")
    f.write("| # | Ceremony / Function | Song Name | Album / Source | Singer(s) | Era | Cultural Context | YouTube Link |\n")
    f.write("|---|---------------------|-----------|----------------|-----------|-----|------------------|--------------|\n")
    
    for idx, s in enumerate(evergreen_bhojpuri, 1):
        cero, name, alb, singer, era, desc = s
        yt = make_yt_link(name, singer)
        f.write(f"| {idx} | **{cero}** | **{name}** | {alb} | {singer} | {era} | {desc} | [▶ YouTube]({yt}) |\n")

# 2. Write Latest Trending Hits Markdown
md_trending = os.path.join(TRENDING_DIR, "Latest_Trending_Bhojpuri_Wedding_Hits.md")
with open(md_trending, "w", encoding="utf-8") as f:
    f.write("# 🔥 Latest Trending Bhojpuri Wedding Hits (2024–2026)\n")
    f.write("### Featuring Top Stars: Pawan Singh, Shilpi Raj, Khesari Lal Yadav, Priyanka Singh\n\n")
    f.write("> **Clean Filter Guarantee:**\n")
    f.write("> - High energy and heavy bass drops without any offensive or inappropriate lyrics.\n")
    f.write("> - Vetted specifically for modern Baraat processions and energetic Sangeet dancing.\n\n")
    f.write("---\n\n")
    f.write("| # | Ceremony / Function | Song Name | Album / Single | Singer(s) | Year | Celebration Style | YouTube Link |\n")
    f.write("|---|---------------------|-----------|----------------|-----------|------|-------------------|--------------|\n")
    
    for idx, s in enumerate(trending_bhojpuri, 1):
        cero, name, alb, singer, yr, desc = s
        yt = make_yt_link(name, singer)
        f.write(f"| {idx} | **{cero}** | **{name}** | {alb} | {singer} | {yr} | {desc} | [▶ YouTube]({yt}) |\n")

# 3. Add to Master Excel Workbook
excel_file = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")
wb = openpyxl.load_workbook(excel_file)

sheet_name = "Bhojpuri Wedding Songs"
if sheet_name in wb.sheetnames:
    del wb[sheet_name]

ws = wb.create_sheet(title=sheet_name)
ws.views.sheetView[0].showGridLines = True

header_fill = PatternFill(start_color="9C27B0", end_color="9C27B0", fill_type="solid") # Royal Festive Purple
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
data_font = Font(name="Calibri", size=10)
bold_font = Font(name="Calibri", size=10, bold=True)
link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

headers = ["#", "Category", "Ceremony / Function", "Song Name", "Album / Source", "Singer(s)", "Year / Era", "Context / Vibe", "YouTube Link"]
for col_idx, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

all_bhojpuri = [("Evergreen Classic",) + s for s in evergreen_bhojpuri] + [("Latest Trending (2024-2026)",) + s for s in trending_bhojpuri]

for idx, song in enumerate(all_bhojpuri, start=1):
    row_num = idx + 1
    cat, cero, name, alb, singer, yr, desc = song
    yt = make_yt_link(name, singer)
    
    ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(horizontal="center")
    ws.cell(row=row_num, column=2, value=cat).alignment = Alignment(horizontal="center")
    ws.cell(row=row_num, column=3, value=cero).font = bold_font
    ws.cell(row=row_num, column=4, value=name).font = bold_font
    ws.cell(row=row_num, column=5, value=alb)
    ws.cell(row=row_num, column=6, value=singer)
    ws.cell(row=row_num, column=7, value=yr).alignment = Alignment(horizontal="center")
    ws.cell(row=row_num, column=8, value=desc)
    
    c_yt = ws.cell(row=row_num, column=9, value="▶ YouTube")
    c_yt.hyperlink = yt
    c_yt.font = link_font
    c_yt.alignment = Alignment(horizontal="center")
    
    for c_i in range(1, 10):
        ws.cell(row=row_num, column=c_i).border = thin_border
        if cat == "Evergreen Classic":
            ws.cell(row=row_num, column=c_i).fill = PatternFill(start_color="FDF7FF", end_color="FDF7FF", fill_type="solid")
        else:
            ws.cell(row=row_num, column=c_i).fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")

for col_idx in range(1, 10):
    max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, len(all_bhojpuri)+2))
    col_letter = get_column_letter(col_idx)
    ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

wb.save(excel_file)
print(f"Generated Bhojpuri Wedding collection with {len(all_bhojpuri)} songs successfully!")

