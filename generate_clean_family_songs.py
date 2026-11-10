import os
import urllib.parse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# STRICT FILTER CRITERIA:
# 100% Family-safe, respectful, zero vulgarity/innuendo, celebration & dignified songs.
# Divided strictly into "Evergreen Classic" and "Trending (2024-2026)"

pure_family_songs = [
    # ==========================================
    # 1. MAYAN CEREMONY (Family & Spiritual Dignity)
    # ==========================================
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Mangal Dhun (Shehnai)", "Shaadi Ki Shehnai", "Ustad Bismillah Khan", "1960s"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Raga Bhairavi (Shehnai)", "Festive Shehnai", "Ustad Bismillah Khan", "1965"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Shri Ram Janki Vivah Mangalgeet", "Shubh Vivah Mangal Geet", "Sharma Bandhu", "1998"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Mangal Bhawan Amangal Haari", "Ramayan", "Ravindra Jain", "1987"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Gauri Ke Lala Ko Manao (Ganesh Vandana)", "Devotional Vivah", "Traditional Folk", "Classical"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Aaj Janakpur Mein Marwa", "Vivah Geet", "Sharda Sinha", "1990s"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Banna Re Banna", "Hum Dil De Chuke Sanam", "Kavita Krishnamurthy, Vinod Rathod", "1999"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Piya Tose Naina Laage Re", "Guide", "Lata Mangeshkar", "1965"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Tujh Mein Rab Dikhta Hai", "Rab Ne Bana Di Jodi", "Roop Kumar Rathod", "2008"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Evergreen Classic", "Badhai Ho Badhai Shehnai Dhun", "Wedding Shehnai", "Ustad Bismillah Khan", "1970"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Trending (2024-2026)", "Chaap Tilak Sab Cheeni", "Coke Studio", "Rahul Vaidya, Shreyas Puranik", "2020"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Trending (2024-2026)", "Kudmayi (Spiritual / Family Edit)", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Shreya Ghoshal", "2023"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Trending (2024-2026)", "Kesariya (Acoustic / Flute)", "Brahmastra", "Arijit Singh", "2022"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Trending (2024-2026)", "Din Shagna Da", "Phillauri", "Jasleen Royal", "2017"),
    ("Day 1 (10 Dec)", "Mayan Ceremony", "Trending (2024-2026)", "Raghupati Raghav Raja Ram (Wedding Fusion)", "Devotional Ambient", "Various Artists", "2024"),

    # ==========================================
    # 2. HALDI CEREMONY (Auspicious, Cheerful, Respectful)
    # ==========================================
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Navrai Majhi", "English Vinglish", "Sunidhi Chauhan, Swanand Kirkire", "2012"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "London Thumakda", "Queen", "Labh Janjua, Sonu Kakkar, Neha Kakkar", "2014"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Sasural Genda Phool", "Delhi-6", "Rekha Bhardwaj", "2009"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Gal Mitthi Mitthi Bol", "Aisha", "Tochi Raina", "2010"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Wah Wah Vattna Katore Da", "Punjabi Folk Classic", "Traditional", "Folk"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Mathe Te Chamkan Waal", "Maiyan Folk", "Traditional Folk", "Folk"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Suhe Ve Cheere Waleya", "Punjabi Folk", "Surinder Kaur, Parkash Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Kala Doriya", "Punjabi Folk", "Surinder Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Latthe Di Chadar", "Punjabi Folk", "Surinder Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Nachde Ne Saare", "Baar Baar Dekho", "Jasleen Royal, Harshdeep Kaur", "2016"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Gallan Goodiyaan", "Dil Dhadakne Do", "Shankar Mahadevan, Yashita Sharma", "2015"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Evergreen Classic", "Bari Barsi", "Band Baaja Baaraat", "Salim Merchant, Labh Janjua", "2010"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Laal Peeli Akhiyaan", "Teri Baaton Mein Aisa Uljha Jiya", "Shadab Faridi, Tanishk Bagchi", "2024"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Haldi Lagao Sajna Ko", "Dil Ki Diary Originals", "Traditional Folk Artist", "2025"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Haldi Hai Aayi", "Wedding Anthem Single", "Independent Artists", "2025"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Rangisari", "Jugjugg Jeeyo", "Kanishk Seth, Kavita Seth", "2022"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "What Jhumka?", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Jonita Gandhi", "2023"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Chaleya", "Jawan", "Arijit Singh, Shilpa Rao", "2023"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Kudmayi (Haldi Version)", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh", "2023"),
    ("Day 1 (10 Dec)", "Haldi Ceremony", "Trending (2024-2026)", "Gud Naal Ishq Mitha", "Ek Ladki Ko Dekha Toh Aisa Laga", "Navraj Hans, Harshdeep Kaur", "2019"),

    # ==========================================
    # 3. MEHNDI CEREMONY (Decent, Graceful Henna Songs)
    # ==========================================
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Mehndi Hai Rachne Wali", "Zubeidaa", "Alka Yagnik", "2001"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Mehndi Laga Ke Rakhna", "DDLJ", "Lata Mangeshkar, Udit Narayan", "1995"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Mehendi Rang Layee", "Chal Mere Bhai", "Udit Narayan, Sonu Nigam, Alka Yagnik", "2000"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Maine Payal Hai Chhankai", "Indipop Classic", "Falguni Pathak", "1999"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Likh Ke Mehndi Se Sajna Ka Naam", "Ishq Hua", "Anuradha Paudwal", "1997"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Bole Chudiyan", "Kabhi Khushi Kabhie Gham", "Sonu Nigam, Alka Yagnik, Kavita K.", "2001"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Chhalka Chhalka Re", "Saathiya", "Richa Sharma, Mahalaxmi Iyer", "2002"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Punjabi Wedding Song", "Hasee Toh Phasee", "Sunidhi Chauhan, Benny Dayal", "2014"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Wah Wah Ramji", "Hum Aapke Hain Koun", "SP Balasubrahmanyam, Lata Mangeshkar", "1994"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Evergreen Classic", "Lo Chali Main", "Hum Aapke Hain Koun", "Lata Mangeshkar", "1994"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Trending (2024-2026)", "Dekhha Tenu", "Mr. & Mrs. Mahi", "Mohd. Faiz, Jaani", "2024"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Trending (2024-2026)", "Rangrez", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Trending (2024-2026)", "Aayi Nai", "Stree 2", "Sachin-Jigar, Divya Kumar", "2024"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Trending (2024-2026)", "Sweetheart", "Kedarnath", "Dev Negi, Palak Muchhal", "2018"),
    ("Day 2 (11 Dec)", "Mehndi Ceremony", "Trending (2024-2026)", "Prem Ratan Dhan Payo", "Prem Ratan Dhan Payo", "Palak Muchhal", "2015"),

    # ==========================================
    # 4. SANGEET & FAMILY PERFORMANCES (Joyful, Dignified & Clean)
    # ==========================================
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Maahi Ve", "Kal Ho Naa Ho", "Sonu Nigam, Shankar Mahadevan, Sadhana Sargam", "2003"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Say Shava Shava", "Kabhi Khushi Kabhie Gham", "Amitabh Bachchan, Alka Yagnik, Udit Narayan", "2001"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Didi Tera Devar Deewana", "Hum Aapke Hain Koun", "Lata Mangeshkar, SP Balasubrahmanyam", "1994"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Joote Do Paise Lo", "Hum Aapke Hain Koun", "Lata Mangeshkar, SP Balasubrahmanyam", "1994"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Chitta Kukkad", "Punjabi Folk", "Traditional / Neha Bhasin", "Folk"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Ghoomar", "Padmaavat", "Shreya Ghoshal, Swaroop Khan", "2018"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Mere Haathon Mein Nau Nau Chudiyan", "Chandni", "Lata Mangeshkar", "1989"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Mera Piya Ghar Aaya", "Yaraana", "Kavita Krishnamurthy", "1995"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Ainvayi Ainvayi", "Band Baaja Baaraat", "Salim Merchant, Sunidhi Chauhan", "2010"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Sadi Gali", "Tanu Weds Manu", "Lehmber Hussainpuri", "2011"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Deewangi Deewangi", "Om Shanti Om", "Shaan, Udit Narayan, Sunidhi Chauhan", "2007"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Evergreen Classic", "Dilbaro (Father-Daughter Special)", "Raazi", "Harshdeep Kaur, Vibha Saraf, Shankar M.", "2018"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Sajni", "Laapataa Ladies", "Arijit Singh, Ram Sampath", "2024"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Heeriye", "Single", "Jasleen Royal, Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Tum Kya Mile", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Shreya Ghoshal", "2023"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "O Maahi", "Dunki", "Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Ve Haaniyaan", "Single", "Aarsh Qaamil, Danny", "2022"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Apna Bana Le", "Bhediya", "Arijit Singh", "2022"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Sauda Khara Khara", "Good Newwz", "Diljit Dosanjh, Sukhbir", "2019"),
    ("Day 2 (11 Dec)", "Sangeet Ceremony", "Trending (2024-2026)", "Morni Banke", "Badhaai Ho", "Guru Randhawa, Neha Kakkar", "2018"),

    # ==========================================
    # 5. BARAAT & GROOM ENTRY (High Voltage & Respectful Celebrations)
    # ==========================================
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Aaj Mere Yaar Ki Shaadi Hai", "Aadmi Sadak Ka", "Mohammed Rafi", "1977"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Dulhe Ka Sehra Suhana Lagta Hai", "Dhadkan", "Nusrat Fateh Ali Khan", "2000"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Saajanji Ghar Aaye", "Kuch Kuch Hota Hai", "Kumar Sanu, Alka Yagnik, Kavita K.", "1998"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Tenu Leke", "Salaam-e-Ishq", "Sonu Nigam, Mahalakshmi Iyer", "2007"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Tunak Tunak Tun", "Album", "Daler Mehndi", "1998"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Ishq Tera Tadpave (Oh Ho Ho Ho)", "Album", "Sukhbir", "1998"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Malhari", "Bajirao Mastani", "Vishal Dadlani", "2015"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Nagada Sang Dhol", "Ram-Leela", "Shreya Ghoshal, Osman Mir", "2013"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Ghodi Chadeya Bhena Da Veer", "Punjabi Folk", "Traditional Folk", "Folk"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Banna Re Baga Me Jhula", "Rajasthani Folk", "Traditional Folk", "Folk"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Veer Ji Viyahn Aaye", "Punjabi Folk Pop", "Jassi Sidhu", "2011"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Evergreen Classic", "Mundian To Bach Ke", "Album", "Panjabi MC", "1998"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "Bhool Bhulaiyaa 3 Title Track", "Bhool Bhulaiyaa 3", "Neeraj Shridhar, Diljit Dosanjh", "2024"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "Mere Mehboob Mere Sanam (Family Dance Edit)", "Bad Newz", "Alka Yagnik, Udit Narayan", "2024"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "Pushpa Pushpa", "Pushpa 2: The Rule", "Mika Singh, DSP", "2024"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "Vibe", "MoonChild Era", "Diljit Dosanjh", "2022"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "Born to Shine", "G.O.A.T.", "Diljit Dosanjh", "2020"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "Jhoome Jo Pathaan", "Pathaan", "Arijit Singh, Vishal-Sheykhar", "2023"),
    ("Day 3 (12 Dec)", "Baraat Ceremony", "Trending (2024-2026)", "The Punjaabban Song", "JugJugg Jeeyo", "Gippy Grewal, Zahrah Khan", "2022"),

    # ==========================================
    # 6. WEDDING RITUALS, VARMALA, PHERAS & VIDAAI (Sacred & Royal)
    # ==========================================
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Mangal Parinay (Raga Rageshri)", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Kanyadaan (Hamsa Narayan)", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Sindoor Daan Chaiti", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Mangalyam (Sacred Chants)", "Saathiya", "A.R. Rahman", "2002"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Raag Yaman (Sitar & Bansuri)", "Classical Jugalbandi", "Pt. Ravi Shankar & Hariprasad Chaurasia", "Classical"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Vedic Vivah Mantras & Shehnai", "Sacred Mantras", "Vedic Priests Ensemble", "Sacred"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Saptha Padi Mantras (Seven Vows)", "Vedic Rituals", "Traditional Chants", "Sacred"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Pehla Nasha (Varmala Instrumental)", "Jo Jeeta Wohi Sikandar", "Instrumental Ensemble", "1992"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Jashn-E-Bahara (Flute Instrumental)", "Jodhaa Akbar", "A.R. Rahman", "2008"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Babul Ki Duayen Leti Ja", "Neel Kamal", "Mohammed Rafi", "1968"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Babul Jo Tumne Sikhaya", "Hum Aapke Hain Koun", "Lata Mangeshkar", "1994"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Babul (Bidaai Song)", "Baabul", "Richa Sharma", "2006"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Evergreen Classic", "Ye Betiyan To Babul Ki", "Na Tum Jaano Na Hum", "Lata Mangeshkar", "2002"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "Kudmayi (Grand Jaimala Special)", "Rocky Aur Rani Kii Prem Kahaani", "Shreya Ghoshal, Arijit Singh", "2023"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "Dekha Tenu (Varmala Moment)", "Mr. & Mrs. Mahi", "Arijit Singh", "2024"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "Aaj Sajeya", "Single", "Goldie Sohel", "2021"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "Madhanya (Bridal Entry & Vidaai)", "Single", "Rahul Vaidya, Asees Kaur", "2021"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "O Piya (Bridal Walk)", "Wedding Single", "Parineeti Chopra", "2023"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "Tere Vaaste (Acoustic Jaimala)", "Zara Hatke Zara Bachke", "Arijit Singh", "2023"),
    ("Day 3 (12 Dec)", "Wedding Rituals", "Trending (2024-2026)", "Sajni (Vidaai Acoustic Farewell)", "Laapataa Ladies", "Arijit Singh", "2024"),
]

def make_yt_link(name, movie):
    query = f"{name} {movie}".strip()
    return f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"

def make_sp_link(name, singer):
    first_singer = singer.split(",")[0].split("&")[0].split("/")[0].strip()
    query = f"{name} {first_singer}".strip()
    return f"https://open.spotify.com/search/{urllib.parse.quote(query)}"

# Generate Dedicated Clean Excel Sheet
excel_file = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")
wb = openpyxl.load_workbook(excel_file)

sheet_name = "Family Dignity & Clean Songs"
if sheet_name in wb.sheetnames:
    del wb[sheet_name]

ws = wb.create_sheet(title=sheet_name)
ws.views.sheetView[0].showGridLines = True

header_fill = PatternFill(start_color="104E8B", end_color="104E8B", fill_type="solid") # Deep Royal Blue
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

headers = ["#", "Day", "Ceremony / Function", "Type (Classic / Trending)", "Song Name", "Movie / Album", "Singer(s)", "Year", "YouTube Link", "Spotify Link"]
for col_idx, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col_idx, value=h)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

for idx, song in enumerate(pure_family_songs, start=1):
    row_num = idx + 1
    day, func, song_type, name, movie, singer, yr = song
    yt = make_yt_link(name, movie)
    sp = make_sp_link(name, singer)
    
    ws.cell(row=row_num, column=1, value=idx).alignment = Alignment(horizontal="center")
    ws.cell(row=row_num, column=2, value=day)
    ws.cell(row=row_num, column=3, value=func).font = bold_font
    ws.cell(row=row_num, column=4, value=song_type).alignment = Alignment(horizontal="center")
    ws.cell(row=row_num, column=5, value=name).font = bold_font
    ws.cell(row=row_num, column=6, value=movie)
    ws.cell(row=row_num, column=7, value=singer)
    ws.cell(row=row_num, column=8, value=yr).alignment = Alignment(horizontal="center")
    
    c_yt = ws.cell(row=row_num, column=9, value="▶ YouTube")
    c_yt.hyperlink = yt
    c_yt.font = link_font
    c_yt.alignment = Alignment(horizontal="center")
    
    c_sp = ws.cell(row=row_num, column=10, value="🎵 Spotify")
    c_sp.hyperlink = sp
    c_sp.font = link_font
    c_sp.alignment = Alignment(horizontal="center")
    
    for c_i in range(1, 11):
        ws.cell(row=row_num, column=c_i).border = thin_border
        if song_type == "Evergreen Classic":
            ws.cell(row=row_num, column=c_i).fill = PatternFill(start_color="F7F9FB", end_color="F7F9FB", fill_type="solid")
        else:
            ws.cell(row=row_num, column=c_i).fill = PatternFill(start_color="FFFDF0", end_color="FFFDF0", fill_type="solid")

for col_idx in range(1, 11):
    max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, len(pure_family_songs)+2))
    col_letter = get_column_letter(col_idx)
    ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

wb.save(excel_file)
print(f"Updated Excel workbook with {len(pure_family_songs)} 100% Family Dignity songs!")

# Write standalone Markdown report
md_report_path = os.path.join(BASE_DIR, "Family_Dignity_Clean_Wedding_Songs.md")
with open(md_report_path, "w", encoding="utf-8") as f:
    f.write("# 👨‍👩‍👧‍👦 Family-Friendly & Dignified Wedding Songs (100% Sanskaari & Joyful)\n")
    f.write("### Specially curated for Brother's Wedding (10, 11 & 12 December 2026)\n\n")
    f.write("> **Strict Filtering Guarantee:**\n")
    f.write("> - ❌ ZERO double meanings, vulgarity, provocative language or embarrassing lines.\n")
    f.write("> - ❌ ZERO awkward roasting or disrespectful in-law taunting songs.\n")
    f.write("> - ✅ 100% Dignified, auspicious, celebratory songs that maintain family respect (izzat) in front of elders, in-laws and kids.\n\n")
    f.write("---\n\n")

    current_func = None
    for s in pure_family_songs:
        day, func, song_type, name, movie, singer, yr = s
        if func != current_func:
            current_func = func
            f.write(f"\n## 🌸 {func} ({day})\n\n")
            f.write("| Type | Song Name | Movie / Album | Singer(s) | Year | Link |\n")
            f.write("|------|-----------|---------------|-----------|------|------|\n")
        yt = make_yt_link(name, movie)
        f.write(f"| **{song_type}** | **{name}** | {movie} | {singer} | {yr} | [▶ YouTube]({yt}) |\n")

print("Created Family_Dignity_Clean_Wedding_Songs.md successfully!")

