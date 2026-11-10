import os
import urllib.parse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# Comprehensive Song Database
# Fields: Day, Function, FolderName, Subcategory, SongName, MovieAlbum, Singer, Year
songs_db = [
    # ==========================================
    # DAY 1 - 10 December 2026: Mayan & Haldi
    # ==========================================
    # 1. Traditional Mayan Songs
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Mangal Dhun (Shehnai)", "Shaadi Ki Shehnai", "Ustad Bismillah Khan", "1960s"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Raga Bhairavi - Shehnai", "Festive Shehnai", "Ustad Bismillah Khan", "1965"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Raga Yaman - Shehnai", "Shaadi Ki Shehnai Vol. 2", "Ustad Bismillah Khan", "1968"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Kanyadaan Dhun (Raga Mishra Khamaj)", "Wedding Shehnai: Mangal Parinay", "Saregama Classical", "2005"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Raga Pahadi Dhun - Shehnai", "Mangal Dhun Va Vivah Dhun", "Various Artists", "2000"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Raga Bageshri - Shehnai", "Shaadi Ki Shehnai", "Ustad Bismillah Khan", "1960s"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Badhai Ho Badhai - Shehnai Dhun", "Wedding Shehnai Collection", "Ustad Bismillah Khan", "1970"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Tilak Ceremony (Instrumental)", "Wedding Instrumentals", "Umashankar Shukla & Shreedhara Chari", "2015"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Mangala Vadyam / Nadaswaram", "Traditional Classical", "T.E. Palaniswamy", "Classical"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Shehnai & Classical", "Raga Rageshwari - Shehnai", "Festive Shehnai", "Ustad Bismillah Khan", "1972"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Ari Sakhi Mangal Gaavo Re", "Traditional Mangal Bhajan", "Traditional / Various", "Classical"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Aaj Janakpur Mangal", "Vivah Geet", "Shubha Mudgal", "2000"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Shri Ram Janki Vivah Mangalgeet", "Shubh Vivah Mangal Geet", "Sharma Bandhu", "1998"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Shiv Gaura Vivah Mangal Geet", "Vivah Geet Collection", "Traditional / Various", "Classical"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Gauri Ke Lala Ko Manao (Ganesh Vandana)", "Traditional Wedding Geet", "Traditional Folk", "Classical"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Banna Re Banna", "Hum Dil De Chuke Sanam", "Kavita Krishnamurthy, Vinod Rathod", "1999"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Vivah Geet - Suhag Collection", "Sharda Sinha Vivah Geet", "Sharda Sinha", "1990s"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Aaj Bhaiya Ka Tilak Chad Jayega", "Tilak Utsav", "Kalpana Patowary", "2010"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Traditional Tilak Geet", "Vivah Geet Folk Collection", "Sharda Dubey", "2005"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Devotional & Vivah Geet", "Mangal Bhawan Amangal Haari", "Ramayan", "Ravindra Jain", "1987"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Chaap Tilak Sab Cheeni", "Coke Studio", "Shreyas Puranik, Rahul Vaidya", "2020"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Chaap Tilak", "Single", "Nakash Aziz, Sargam Jassu", "2022"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Kabira (Encore)", "Yeh Jawaani Hai Deewani", "Rekha Bhardwaj, Tochi Raina", "2013"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Tujh Mein Rab Dikhta Hai", "Rab Ne Bana Di Jodi", "Roop Kumar Rathod", "2008"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Piya Tose Naina Laage Re", "Guide", "Lata Mangeshkar", "1965"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Kesariya", "Brahmastra", "Arijit Singh", "2022"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Din Shagna Da", "Phillauri", "Jasleen Royal", "2017"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Kudmayi", "Rocky Aur Rani Kii Prem Kahaani", "Raftaar, Nikhita Gandhi", "2023"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Mehendi Hai Rachne Wali", "Zubeidaa", "Alka Yagnik", "2001"),
    ("Day 1 (10 Dec)", "Traditional Mayan Ceremony", "10 December - Mayan & Haldi/Traditional Mayan Songs", "Bollywood & Fusion Mayan", "Bole Chudiyan", "Kabhi Khushi Kabhie Gham", "Sonu Nigam, Alka Yagnik, Kavita K., Udit Narayan", "2001"),

    # 2. Haldi Bollywood Songs
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Navrai Majhi", "English Vinglish", "Sunidhi Chauhan, Swanand Kirkire", "2012"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "London Thumakda", "Queen", "Labh Janjua, Sonu Kakkar, Neha Kakkar", "2014"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Sasural Genda Phool", "Delhi-6", "Rekha Bhardwaj", "2009"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Genda Phool", "Single", "Badshah, Payal Dev", "2020"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Gallan Goodiyaan", "Dil Dhadakne Do", "Shankar Mahadevan, Farhan Akhtar & others", "2015"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Banno Tera Swagger", "Tanu Weds Manu Returns", "Swati Sharma", "2015"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Mehndi Hai Rachne Wali", "Zubeidaa", "Alka Yagnik", "2001"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Mehndi Laga Ke Rakhna", "DDLJ", "Lata Mangeshkar, Udit Narayan", "1995"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Nachde Ne Saare", "Baar Baar Dekho", "Jasleen Royal, Harshdeep Kaur", "2016"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Kala Chashma", "Baar Baar Dekho", "Amar Arshi, Badshah, Neha Kakkar", "2016"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Cutiepie", "Ae Dil Hai Mushkil", "Pardeep Sran, Nakash Aziz", "2016"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Morni Banke", "Badhaai Ho", "Guru Randhawa, Neha Kakkar", "2018"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Badri Ki Dulhania", "Badrinath Ki Dulhania", "Dev Negi, Neha Kakkar, Monali Thakur", "2017"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Sauda Khara Khara", "Good Newwz", "Diljit Dosanjh, Sukhbir, Dhvani Bhanushali", "2019"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Dil Chori", "Sonu Ke Titu Ki Sweety", "Yo Yo Honey Singh, Simar Kaur", "2018"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Ainvayi Ainvayi", "Band Baaja Baaraat", "Salim Merchant, Sunidhi Chauhan", "2010"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Pallo Latke", "Shaadi Mein Zaroor Aana", "Jyotica Tangri, Yasser Desai", "2017"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Rangisari", "Jugjugg Jeeyo", "Kanishk Seth, Kavita Seth", "2022"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Gud Naal Ishq Mitha", "Ek Ladki Ko Dekha Toh Aisa Laga", "Navraj Hans, Harshdeep Kaur", "2019"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Chaka Chak", "Atrangi Re", "Shreya Ghoshal", "2021"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Balam Pichkari", "Yeh Jawaani Hai Deewani", "Vishal Dadlani, Shalmali Kholgade", "2013"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Laal Peeli Akhiyaan", "Teri Baaton Mein Aisa Uljha Jiya", "Shadab, Tanishk Bagchi", "2024"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "What Jhumka?", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Jonita Gandhi", "2023"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Chaleya", "Jawan", "Arijit Singh, Shilpa Rao", "2023"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Kudi Nu Nachne De", "Angrezi Medium", "Vishal Dadlani, Sachin-Jigar", "2020"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Gal Mitthi Mitthi", "Aisha", "Tochi Raina", "2010"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Haldi Classics", "Lat Lag Gayi", "Race 2", "Benny Dayal, Shalmali Kholgade", "2013"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Trending Viral (2024-2026)", "Haldi Lagao Sajna Ko", "Dil Ki Diary Originals", "Independent", "2025"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Trending Viral (2024-2026)", "Haldi Hai Aayi", "Wedding Anthem Single", "Independent", "2025"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Trending Viral (2024-2026)", "Sabse Pyara Haldi Song", "Single", "Anurag Yaduvanshi", "2024"),
    ("Day 1 (10 Dec)", "Haldi Bollywood", "10 December - Mayan & Haldi/Haldi Bollywood Songs", "Trending Viral (2024-2026)", "Haldi Ki Rasm / Tona Haradiya Na", "Folk Fusion Single", "Various", "2024"),

    # 3. Punjabi Haldi Songs
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Wah Wah Vattna Katore Da", "Traditional Punjabi Folk", "Traditional / Various Artists", "Classic"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Mathe Te Chamkan Waal", "Traditional Folk / Maiyan Songs", "Traditional / Various", "Classic"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Suhe Ve Cheere Waleya", "Punjabi Folk Classic", "Surinder Kaur & Parkash Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Kala Doriya", "Punjabi Folk Classic", "Surinder Kaur & Parkash Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Latthe Di Chadar", "Punjabi Folk Classic", "Surinder Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Bajre Da Sitta", "Punjabi Folk", "Surinder Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Jutti Kasoori", "Punjabi Folk", "Surinder Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Baari Barsi Khatan Gaya Si", "Traditional Boliyan", "Traditional Folk", "Classic"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Traditional Gidda Boliyan", "Suhag Collection", "Various Folk Artists", "Classic"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Deor De Vyah Wich Nach Len De", "Punjabi Folk", "Surinder Kaur", "1960s"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Maiyan Song", "Independent", "Raj Ghuman", "2018"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Traditional Folk & Vatna", "Punjabi Wedding Tappe", "Traditional Folk Collection", "Various", "Classic"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "G.O.A.T.", "G.O.A.T.", "Diljit Dosanjh", "2020"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Lover", "MoonChild Era", "Diljit Dosanjh", "2021"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Born to Shine", "G.O.A.T.", "Diljit Dosanjh", "2020"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Lemonade", "Drive Thru", "Diljit Dosanjh", "2023"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Brown Munde", "Single", "AP Dhillon, Gurinder Gill", "2020"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Excuses", "Single", "AP Dhillon, Gurinder Gill", "2021"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Insane", "Single", "AP Dhillon", "2021"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Old Money", "Single", "AP Dhillon", "2023"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Patiala Peg", "Single", "Diljit Dosanjh", "2014"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Panj Taara", "Single", "Diljit Dosanjh", "2017"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Ki Banu Duniya Da", "Folk Album", "Gurdas Maan", "1984"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Dil Da Mamla Hai", "Album", "Gurdas Maan", "1994"),
    ("Day 1 (10 Dec)", "Punjabi Haldi", "10 December - Mayan & Haldi/Punjabi Haldi Songs", "Modern Punjabi Pop & Bhangra", "Apna Punjab Hovay", "Album", "Gurdas Maan", "2007"),

    # 4. Family Celebration (Day 1)
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Sweety Tera Drama", "Bareilly Ki Barfi", "Dev Negi, Pawni Pandey", "2017"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Aaj Ki Party", "Bajrangi Bhaijaan", "Mika Singh", "2015"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Tamma Tamma Again", "Badrinath Ki Dulhania", "Badshah, Bappi Lahiri, Anuradha Paudwal", "2017"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Baari Barsi", "Band Baaja Baaraat", "Labh Janjua, Salim Merchant", "2010"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Maahi Ve", "Kal Ho Naa Ho", "Sonu Nigam, Shankar Mahadevan & others", "2003"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Nagada Sang Dhol", "Goliyon Ki Raasleela Ram-Leela", "Shreya Ghoshal", "2013"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Radha", "Student of the Year", "Shreya Ghoshal, Vishal Dadlani, Shekhar Ravjiani", "2012"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Gal Ban Gayi", "Single", "Sukhbir, Neha Kakkar, Meet Bros", "2016"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Baby Ko Bass Pasand Hai", "Sultan", "Vishal Dadlani, Shalmali Kholgade, Badshah", "2016"),
    ("Day 1 (10 Dec)", "Family Celebration", "10 December - Mayan & Haldi/Family Celebration", "Family Dance Anthems", "Rang De Tu Mohe Gerua", "Dilwale", "Arijit Singh, Antara Mitra", "2015"),

    # 5. Funny & Entertaining Songs
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Didi Tera Devar Deewana", "Hum Aapke Hain Koun..!", "Lata Mangeshkar, S.P. Balasubrahmanyam", "1994"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Joote Do Paise Lo", "Hum Aapke Hain Koun..!", "Lata Mangeshkar, S.P. Balasubrahmanyam", "1994"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Kabootar Ja Ja Ja", "Maine Pyar Kiya", "Lata Mangeshkar, S.P. Balasubrahmanyam", "1989"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Sajan Ji Vaari Vaari", "Honeymoon Travels Pvt. Ltd.", "Kay Kay, Shaan", "2007"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Aloo Chaat", "Aloo Chaat", "RDB, Nindy Kaur", "2009"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Tamma Tamma Loge", "Thanedaar", "Anuradha Paudwal, Bappi Lahiri", "1990"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Le Jayenge Le Jayenge", "Chor Machaye Shor", "Kishore Kumar, Lata Mangeshkar", "1974"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Tune Maari Entriyaan", "Gunday", "Vishal Dadlani, Bappi Lahiri, Neeti Mohan", "2014"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Drama Queen", "Hasee Toh Phasee", "Shreya Ghoshal", "2014"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Saturday Saturday", "Humpty Sharma Ki Dulhania", "Indeep Bakshi, Badshah, Akriti Kakar", "2014"),
    ("Day 1 (10 Dec)", "Funny & Entertaining", "10 December - Mayan & Haldi/Funny & Entertaining Songs", "Humorous Wedding Hits", "Ding Dong Dole", "Wedding Anthem", "Various", "2024"),

    # ==========================================
    # DAY 2 - 11 December 2026: Mehndi & Sangeet
    # ==========================================
    # 1. Mehndi Songs
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Mehndi Hai Rachne Wali", "Zubeidaa", "Alka Yagnik", "2001"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Mehndi Laga Ke Rakhna", "Dilwale Dulhania Le Jayenge", "Lata Mangeshkar, Udit Narayan", "1995"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Mehendi Rang Layee", "Chal Mere Bhai", "Udit Narayan, Sonu Nigam, Alka Yagnik", "2000"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Punjabi Wedding Song", "Hasee Toh Phasee", "Sunidhi Chauhan, Benny Dayal", "2014"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Likh Ke Mehndi Se Sajna Ka Naam", "Ishq Hua", "Anuradha Paudwal", "1997"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Maine Payal Hai Chhankai", "Falguni Pathak (Indipop)", "Falguni Pathak", "1999"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Prem Ratan Dhan Payo (Title)", "Prem Ratan Dhan Payo", "Palak Muchhal", "2015"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Bumro Bumro", "Mission Kashmir", "Shankar Mahadevan, Jaspinder Narula", "2000"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Rangrez", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Bijli", "Govinda Naam Mera", "Mika Singh, Neha Kakkar", "2022"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Dekhha Tenu", "Mr. & Mrs. Mahi", "Mohd. Faiz", "2024"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Banno Re Banno", "Yeh Jawaani Hai Deewani", "Arijit Singh", "2013"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Lo Chali Main", "Hum Aapke Hain Koun", "Lata Mangeshkar", "1994"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Wah Wah Ramji", "Hum Aapke Hain Koun", "S.P. Balasubrahmanyam, Lata Mangeshkar", "1994"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Say Shava Shava", "Kabhi Khushi Kabhie Gham", "Udit Narayan, Alka Yagnik, Sudesh Bhosle", "2001"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Chhalka Chhalka Re", "Saathiya", "Richa Sharma", "2002"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Bollywood Mehndi", "Sweetheart", "Kedarnath", "Dev Negi, Palak Muchhal", "2018"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Trending & Folk", "Lehenga", "Single/Album", "Jass Manak", "2019"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Trending & Folk", "Kudmayi", "Rocky Aur Rani Kii Prem Kahaani", "Alka Yagnik, Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Mehndi Songs", "11 December - Mehndi & Sangeet/Mehndi Songs", "Trending & Folk", "Aayi Nai", "Stree 2", "Sachin-Jigar, Divya Kumar", "2024"),

    # 2. Ladies Sangeet Songs
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Kala Sha Kala", "Punjabi Folk / Saat Khoon Maaf", "Traditional / Various", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Chitta Kukkad", "Punjabi Folk", "Traditional", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Lathe Di Chadar", "Punjabi Folk", "Various Artists", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Mera Laung Gawacha", "Punjabi Folk", "Musarrat Nazir (Original)", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Suhe Ve Cheere Waleya", "Punjabi Folk", "Traditional", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Kaala Doriya", "Punjabi Folk", "Traditional", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Madhaniya", "Punjabi Folk", "Traditional / Multiple Versions", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Mai Ni Meriye", "Pahari Folk / Modern Cover", "Traditional / Various", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Kesariya Balam Padharo Mhare Des", "Rajasthani Folk", "Traditional / Various", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Banna / Banni Songs", "North Indian Folk", "Traditional", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Traditional Ladies Folk", "Gaali Geet / Gaari", "North Indian Folk (UP/Bihar)", "Traditional", "Traditional"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Banno Ki Aayegi Baraat", "Aaina", "Kumar Sanu, Asha Bhosle", "1993"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Laung Laachi", "Laung Laachi (Punjabi)", "Mannat Noor", "2018"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Guddiyan Patole", "Single (Punjabi)", "Gurnam Bhullar", "2019"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Palazzo", "Single (Punjabi)", "Kulwinder Billa, Shivjot", "2017"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Veere", "Veere Di Wedding", "Vishal Mishra, Aditi Singh Sharma", "2018"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Mere Haathon Mein Nau Nau Chudiyan", "Chandni", "Lata Mangeshkar", "1989"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Ghoomar", "Padmaavat", "Shreya Ghoshal, Swaroop Khan", "2018"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Bari Barsi", "Band Baaja Baaraat", "Salim Merchant, Labh Janjua", "2010"),
    ("Day 2 (11 Dec)", "Ladies Sangeet", "11 December - Mehndi & Sangeet/Ladies Sangeet", "Bollywood Ladies Sangeet", "Mera Piya Ghar Aaya", "Yaraana", "Kavita Krishnamurthy", "1995"),

    # 3. Family Dance Songs (Day 2 Sangeet)
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Deewangi Deewangi", "Om Shanti Om", "Various Artists", "2007"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Dilbaro", "Raazi", "Shankar Mahadevan, Harshdeep Kaur", "2018"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Laadki", "Coke Studio India", "Sachin-Jigar, Taniskha S, Kirtidan Gadhvi", "2020"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Papa Meri Jaan", "Animal", "Sonu Nigam", "2023"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Desi Girl", "Dostana", "Vishal Dadlani, Shankar Mahadevan", "2008"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Prem Ki Naiyya", "Ajab Prem Ki Ghazab Kahani", "Atif Aslam, Neeraj Shridhar", "2009"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Tenu Sang Rakhna", "Single", "Akhil Sachdeva", "2022"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Tauba Tauba", "Bad Newz", "Karan Aujla", "2024"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Bole Chudiyan", "Kabhi Khushi Kabhie Gham", "Udit Narayan, Alka Yagnik, Sonu Nigam", "2001"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Gallan Goodiyan", "Dil Dhadakne Do", "Yashita Sharma, Shankar Mahadevan", "2015"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Nachde Ne Saare", "Baar Baar Dekho", "Jasleen Royal, Harshdeep Kaur", "2016"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Kala Chashma", "Baar Baar Dekho", "Amar Arshi, Badshah, Neha Kakkar", "2016"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Sweetheart", "Kedarnath", "Dev Negi, Palak Muchhal", "2018"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Sauda Khara Khara", "Good Newwz", "Diljit Dosanjh, Sukhbir", "2019"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Morni Banke", "Badhaai Ho", "Guru Randhawa, Neha Kakkar", "2018"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Say Shava Shava", "Kabhi Khushi Kabhie Gham", "Amitabh Bachchan, Udit Narayan, Alka Yagnik", "2001"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "London Thumakda", "Queen", "Labh Janjua, Sonu Nigam, Neha Kakkar", "2014"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Navrai Majhi", "English Vinglish", "Sunidhi Chauhan", "2012"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Maahi Ve", "Kal Ho Naa Ho", "Udit Narayan, Various", "2003"),
    ("Day 2 (11 Dec)", "Family Dance", "11 December - Mehndi & Sangeet/Family Dance", "Family Choreography", "Aayi Nai", "Stree 2", "Sachin-Jigar, Divya Kumar", "2024"),

    # 4. Friends Dance Songs
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Tauba Tauba", "Bad Newz", "Karan Aujla", "2024"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "What Jhumka?", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Jonita Gandhi", "2023"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Bijuria", "Sunny Sanskari Ki Tulsi Kumari", "Mika Singh", "2024"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Panwadi", "Sunny Sanskari Ki Tulsi Kumari", "Various", "2024"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Jamal Kudu (Abrar's Entry)", "Animal", "Harshavardhan Rameshwar", "2023"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Thumkeshwari", "Bhediya", "Sachin-Jigar, Rashmeet Kaur", "2022"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Sooraj Dooba Hain Yaaron", "Roy", "Adnan Sami, Arijit Singh", "2015"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Yaaron", "Pal", "KK", "1999"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Tera Yaar Hoon Main", "Sonu Ke Titu Ki Sweety", "Arijit Singh", "2018"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Mere Yaar Ki Shaadi Hai", "Mere Yaar Ki Shaadi Hai", "Udit Narayan, Sonu Nigam", "2002"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Daaru Desi", "Cocktail", "Benny Dayal, Shalmali Kholgade", "2012"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Abhi Toh Party Shuru Hui Hai", "Khoobsurat", "Badshah, Aastha Gill", "2014"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Mauja Hi Mauja", "Jab We Met", "Mika Singh", "2007"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Atrangi Yaari", "Wazir", "Amitabh Bachchan, Farhan Akhtar", "2016"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Woh Din", "Chhichhore", "Tushar Joshi", "2019"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Ainvayi Ainvayi", "Band Baaja Baaraat", "Salim Merchant, Sunidhi Chauhan", "2010"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Lamberghini", "Single (Punjabi)", "The Doorbeen ft. Ragini", "2018"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Sadi Gali", "Tanu Weds Manu", "Lehmber Hussainpuri", "2011"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Veere", "Veere Di Wedding", "Vishal Mishra, Aditi Singh Sharma", "2018"),
    ("Day 2 (11 Dec)", "Friends Dance", "11 December - Mehndi & Sangeet/Friends Dance", "High Energy & Bhangra", "Dilliwaali Girlfriend", "Yeh Jawaani Hai Deewani", "Arijit Singh, Sunidhi Chauhan", "2013"),

    # 5. Couple Dance Songs
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Heeriye", "Single", "Jasleen Royal ft. Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "O Maahi", "Dunki", "Arijit Singh", "2023"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tum Kya Mile", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Shreya Ghoshal", "2023"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 Dec - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Kesariya", "Brahmastra", "Arijit Singh", "2022"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Apna Bana Le", "Bhediya", "Arijit Singh", "2022"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Ve Haaniyaan", "Single", "Ravi Dubey, Danny", "2022"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Raataan Lambiyan", "Shershaah", "Jubin Nautiyal, Asees Kaur", "2021"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Pehle Bhi Main", "Animal", "Vishal Mishra", "2023"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Sajni", "Laapataa Ladies", "Arijit Singh, Ram Sampath", "2024"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tum Hi Ho", "Aashiqui 2", "Arijit Singh", "2013"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Kal Ho Naa Ho (Title)", "Kal Ho Naa Ho", "Sonu Nigam", "2003"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Janam Janam", "Dilwale", "Arijit Singh, Antara Mitra", "2015"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tum Se Hi", "Jab We Met", "Mohit Chauhan", "2007"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Humsafar", "Badrinath Ki Dulhania", "Akhil Sachdeva, Mansheel Gujral", "2017"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tera Ban Jaunga", "Kabir Singh", "Akhil Sachdeva, Tulsi Kumar", "2019"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Mast Magan", "2 States", "Arijit Singh, Chinmayi", "2014"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Dil Diyan Gallan", "Tiger Zinda Hai", "Atif Aslam", "2017"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tere Hawaale", "Laal Singh Chaddha", "Arijit Singh, Shilpa Rao", "2022"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Hawayein", "Jab Harry Met Sejal", "Arijit Singh", "2017"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tera Hone Laga Hoon", "Ajab Prem Ki Ghazab Kahani", "Atif Aslam, Alisha Chinai", "2009"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Tere Liye", "Veer-Zaara", "Lata Mangeshkar, Roop Kumar Rathod", "2004"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Suraj Hua Maddham", "Kabhi Khushi Kabhie Gham", "Sonu Nigam, Alka Yagnik", "2001"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Raanjhanaa (Title)", "Raanjhanaa", "A.R. Rahman, Shiraz Uppal", "2013"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Piya O Re Piya", "Tere Naal Love Ho Gaya", "Atif Aslam, Shreya Ghoshal", "2012"),
    ("Day 2 (11 Dec)", "Couple Dance", "11 December - Mehndi & Sangeet/Couple Dance", "Romantic Slow & Duets", "Ishq Wala Love", "Student of the Year", "Neeti Mohan, Salim Merchant, Shekhar Ravjiani", "2012"),

    # 6. DJ & Party Songs
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Softly", "Single (Punjabi)", "Karan Aujla", "2024"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "295", "Moosetape", "Sidhu Moose Wala", "2021"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Bijlee Bijlee", "Single (Punjabi)", "Harrdy Sandhu", "2021"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Naine", "Crew", "Diljit Dosanjh, Badshah", "2024"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Oo Antava Oo Oo Antava", "Pushpa: The Rise", "Indravathi Chauhan", "2021"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Gulabi Sadi", "Single (Marathi)", "Sanju Rathod", "2023"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Akhiyaan Gulaab", "Teri Baaton Mein Aisa Uljha Jiya", "Mitraz", "2024"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Nashe Si Chadh Gayi", "Befikre", "Arijit Singh", "2016"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Tera Yaar Bolda", "Single (Punjabi)", "Surjit Bindrakhia", "1999"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Bachna Ae Haseeno (Remix)", "Bachna Ae Haseeno", "Kishore Kumar, Vishal Dadlani", "2008"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "With You", "Single (Punjabi)", "AP Dhillon", "2023"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Kinni Kinni", "Single (Punjabi)", "Diljit Dosanjh", "2024"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Let's Nacho", "Kapoor & Sons", "Nucleya, Badshah, Benny Dayal", "2016"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "The Punjaabban Song", "JugJugg Jeeyo", "Gippy Grewal, Zahrah Khan", "2022"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Badtameez Dil", "Yeh Jawaani Hai Deewani", "Benny Dayal", "2013"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Balam Pichkari", "Yeh Jawaani Hai Deewani", "Vishal Dadlani, Shalmali Kholgade", "2013"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Coca Cola", "Luka Chuppi", "Tony Kakkar, Neha Kakkar", "2019"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Sher Khul Gaye", "Fighter", "Benny Dayal, Vishal Dadlani", "2024"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Jhoome Jo Pathaan", "Pathaan", "Arijit Singh, Vishal-Sheykhar", "2023"),
    ("Day 2 (11 Dec)", "DJ Party", "11 December - Mehndi & Sangeet/DJ Party", "Party Anthems & Club Hits", "Arjan Vailly", "Animal", "Bhupinder Babbal", "2023"),

    # 7. Traditional Wedding Songs (Day 2)
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Hardi Hardiya / Haldi Songs", "Maithili/Bhojpuri Folk", "Sharda Sinha", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "More Angana Kute Dhan Re", "Maithili/Bhojpuri Folk", "Sharda Sinha", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Sita Ke Sakal Dekhi", "Maithili Vivah Geet", "Sharda Sinha", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Chumawan He Lelna", "Maithili Folk", "Sharda Sinha", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Gowri Kalyanam", "Carnatic Classical", "Traditional Carnatic", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Oonjal Songs (Rathna Oonjalil)", "Tamil Wedding Tradition", "Traditional Tamil", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Shobane / Shobanam", "South Indian Classical", "Traditional South Indian", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Maye Ni Maye", "Hum Aapke Hain Koun", "Lata Mangeshkar", "1994"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Piya Ke Nagariya (Bidai)", "Maithili/Bhojpuri Folk", "Sharda Sinha", "Traditional"),
    ("Day 2 (11 Dec)", "Traditional Wedding", "11 December - Mehndi & Sangeet/Traditional Wedding Songs", "Folk & Regional Traditions", "Banna Sa Aave Ghodi Chadhke", "Rajasthani Folk", "Traditional Folk", "Traditional"),

    # ==========================================
    # DAY 3 - 12 December 2026: The Grand Wedding
    # ==========================================
    # 1. Baraat & Groom Entry Songs
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Malhari", "Bajirao Mastani", "Vishal Dadlani", "2015"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Badtameez Dil", "Yeh Jawaani Hai Deewani", "Benny Dayal", "2013"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Tenu Leke", "Salaam-e-Ishq", "Sonu Nigam", "2007"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Bhool Bhulaiyaa 3 Title Track", "Bhool Bhulaiyaa 3", "Neeraj Shridhar, Pitbull", "2024"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Pushpa Pushpa", "Pushpa 2: The Rule", "Mika Singh, Sukumar, DSP", "2024"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Kissik", "Pushpa 2: The Rule", "Shreya Ghoshal", "2024"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Tunak Tunak Tun", "Album", "Daler Mehndi", "1998"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Ishq Tera Tadpave (Oh Ho Ho Ho)", "Album", "Sukhbir", "1998"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Aaj Mere Yaar Ki Shaadi Hai", "Aadmi Sadak Ka", "Mohammed Rafi", "1977"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Saajanji Ghar Aaye", "Kuch Kuch Hota Hai", "Kumar Sanu, Alka Yagnik, Kavita K.", "1998"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Palat - Tera Hero Idhar Hai", "Main Tera Hero", "Arijit Singh", "2014"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Vibe", "Album", "Diljit Dosanjh", "2022"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Mere Mehboob Mere Sanam (Remix)", "Bad Newz", "Mika Singh", "2024"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Aashiq Surrender Hua", "Badrinath Ki Dulhania", "Amaal Mallik, Shreya Ghoshal", "2017"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Chhote Chhote Bhaiyon Ke", "Hum Saath-Saath Hain", "Kavita Krishnamurthy, Udit Narayan", "1999"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Balle Balle", "Bride and Prejudice", "Sonu Nigam, Gayatri Ganjawala", "2004"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Jee Karda", "Singh Is Kinng", "Labh Janjua, Suzie Q", "2008"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Mundian To Bach Ke", "Album", "Panjabi MC", "1998"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Bhaag D.K. Bose", "Delhi Belly", "Ram Sampath", "2011"),
    ("Day 3 (12 Dec)", "Baraat & Groom Entry", "12 December - Wedding/Baraat & Groom Entry", "Grand Baraat Anthems", "Dhol Jageero Da", "Album", "Panjabi MC", "2001"),

    # 2. Ghodi & Procession Songs
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Dulhe Ka Sehra Suhana Lagta Hai", "Dhadkan", "Nusrat Fateh Ali Khan", "2000"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Ghodi Chadeya Bhena Da Veer", "Punjabi Folk", "Traditional Folk Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Neeli Neeli Ghodi", "Punjabi Folk", "Nimmo Chaudhary", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Sehra Bandi Traditional", "Rajasthani/Hindi Folk", "Traditional Folk Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Banna Re Baga Me Jhula", "Rajasthani Folk", "Traditional Folk Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Banna Re Banni Re", "Rajasthani Folk", "Anuradha Paudwal", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Aaj Mere Yaar Ki Shaadi Hai (Band Version)", "Aadmi Sadak Ka", "Brass Band Instrumental", "1977"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Ghoriyan (Traditional Punjabi)", "Punjabi Folk", "Traditional Women Singers", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Pehli Pehli Baar Baliye", "Har Dil Jo Pyar Karega", "Sonu Nigam, Alka Yagnik", "2000"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Banna Sa Padharo Mhare Desh", "Rajasthani Folk", "Traditional Folk Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Sehra Band Ke Main Toh Aaya", "Band Baaja Baaraat", "Traditional / Arijit Cover", "2010"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Raja Ki Aayegi Baraat", "Raja Ki Aayegi Baraat", "Alka Yagnik", "1996"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Dil Chahta Hai (Procession Version)", "Dil Chahta Hai", "Shankar Mahadevan", "2001"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Kaahe Chhed Mohe (Instrumental)", "Devdas", "Classical Instrumental", "2002"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Madaniya / Dhol Beats Mix", "Traditional", "Various Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Veer Ji Viyahn Aaye", "Punjabi Folk", "Jassi Sidhu", "2011"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Mera Yaar Bana Hai Dulha", "Single", "Master Saleem", "2016"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Mangal Geet Ghodi", "Vivah Geet", "Traditional", "Classic"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Ghodi Chad Ke Aaya Dulhe Raja", "Bhojpuri/Hindi Folk", "Manoj Tiwari", "2005"),
    ("Day 3 (12 Dec)", "Ghodi & Procession", "12 December - Wedding/Ghodi & Procession", "Traditional Ghodi Chadai", "Banno Re Banno Mharo Chhel Bhanwar", "Rajasthani Folk", "Seema Mishra", "2003"),

    # 3. Varmala & Jaimala Songs
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Aaj Sajeya", "Album", "Goldie Sohel", "2021"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Peer Vi Tu", "Album", "Harshdeep Kaur", "2018"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Tum Prem Ho", "Radha Krishna", "Surya Raj Kamal", "2018"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Dekha Tenu", "Mr. & Mrs. Mahi", "Arijit Singh", "2024"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Mangalyam (Instrumental)", "Saathiya", "A.R. Rahman", "2002"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Ghar More Pardesiya", "Kalank", "Shreya Ghoshal, Vaishali Mhade", "2019"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Tere Vaaste", "Zara Hatke Zara Bachke", "Varijashree Venugopal, Arijit Singh", "2023"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Soulmate", "Album", "Badshah, Arijit Singh", "2023"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Ram Ram Jai Raja Ram (Wedding Version)", "Devotional", "Various Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Ranjha", "Shershaah", "B Praak", "2021"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Pehla Nasha", "Jo Jeeta Wohi Sikandar", "Udit Narayan, Sadhana Sargam", "1992"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Ban Ja Rani", "Tumhari Sulu", "Guru Randhawa", "2017"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Kudmayi (Jaimala Special)", "Rocky Aur Rani Kii Prem Kahaani", "Shreya Ghoshal, Arijit Singh", "2023"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Aisa Sama Na Hota", "Zameen Aasmaan", "Lata Mangeshkar", "1984"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Tu Mile Dil Khile", "Criminal", "Kumar Sanu, Alka Yagnik", "1995"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Chupke Se", "Saathiya", "Sadhana Sargam", "2002"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Jashn-E-Bahara", "Jodhaa Akbar", "Javed Ali", "2008"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Moh Moh Ke Dhaage", "Dum Laga Ke Haisha", "Papon, Monali Thakur", "2015"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Subhanallah", "Yeh Jawaani Hai Deewani", "Sreerama Chandra, Shilpa Rao", "2013"),
    ("Day 3 (12 Dec)", "Varmala & Jaimala", "12 December - Wedding/Varmala & Jaimala", "Garland Exchange Moments", "Raabta (Kehte Hain Khuda)", "Agent Vinod", "Arijit Singh, Hamsika Iyer", "2012"),

    # 4. Bride Entry Songs
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Madhanya", "Album", "Rahul Vaidya, Asees Kaur", "2021"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "O Piya", "Album", "Parineeti Chopra", "2023"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Chunar", "ABCD 2", "Arijit Singh", "2015"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Deewani Mastani", "Bajirao Mastani", "Shreya Ghoshal", "2015"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Kabhi Khushi Kabhie Gham (Bridal Theme)", "Kabhi Khushi Kabhie Gham", "Lata Mangeshkar", "2001"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Dulhan Hum Le Jayenge", "Dulhan Hum Le Jayenge", "Sonu Nigam, Alka Yagnik", "2000"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Khabar Nahi", "Dostana", "Vishal Dadlani, Shreya Ghoshal", "2008"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Nai Jaana", "Single", "Neha Bhasin", "2018"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Ghar Aaya Mera Pardesi", "Awaara", "Lata Mangeshkar", "1951"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Heer", "Jab Tak Hai Jaan", "Harshdeep Kaur", "2012"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Mangalyam (Wedding Entry)", "Alaipayuthey / Saathiya", "A.R. Rahman", "2000"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Afreen Afreen (Bridal Walk)", "Coke Studio", "Rahat Fateh Ali Khan, Momina Mustehsan", "2016"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Chaudhary", "Coke Studio", "Mame Khan", "2012"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Laadki (Bridal Walk Mix)", "Coke Studio", "Sachin-Jigar", "2015"),
    ("Day 3 (12 Dec)", "Bride Entry", "12 December - Wedding/Bride Entry", "Grand Bridal Entry", "Jugni", "Cocktail", "Javed Bashir", "2012"),

    # 5. Wedding Rituals / Pheras Background Music
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Shehnai & Mantras", "Mangal Parinay (Raga Rageshri) - Shehnai", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Shehnai & Mantras", "Kanyadaan (Hamsa Narayan) - Shehnai", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Shehnai & Mantras", "Vivah Sampann Riti (Shyam Kalyan) - Shehnai", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Shehnai & Mantras", "Sindoor Daan Chaiti - Shehnai", "Classical Album", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Shehnai & Mantras", "Mangal Dhun / Vivah Dhun (Shehnai)", "Classical Collection", "Ustad Bismillah Khan", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Sitar & Bansuri Jugalbandi", "Raag Yaman (Sitar & Flute)", "Classical", "Ravi Shankar / Hariprasad Chaurasia", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Sitar & Bansuri Jugalbandi", "Raag Desh (Sitar & Flute Jugalbandi)", "Classical", "Various Classical Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Sitar & Bansuri Jugalbandi", "Raag Khamaj (Bansuri)", "Classical", "Hariprasad Chaurasia", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Sitar & Bansuri Jugalbandi", "Raag Mishra Piloo (Flute)", "Classical", "Hariprasad Chaurasia", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Sitar & Bansuri Jugalbandi", "Raag Bageshri (Evening Raga - Sitar)", "Classical", "Various Classical Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Sitar & Bansuri Jugalbandi", "Sitar & Santoor Wedding Collection", "Classical", "Shivkumar Sharma / Ravi Shankar", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Kabira (Shehnai/Flute Instrumental)", "Yeh Jawaani Hai Deewani", "Instrumental Cover", "2013"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Din Shagna Da (Shehnai Instrumental)", "Phillauri", "Instrumental Cover", "2017"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Mangalyam Tantunanena (Instrumental)", "Saathiya", "A.R. Rahman", "2002"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Tujh Mein Rab Dikhta Hai (Flute Version)", "Rab Ne Bana Di Jodi", "Instrumental Cover", "2008"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Kal Ho Naa Ho (Piano/Flute Instrumental)", "Kal Ho Naa Ho", "Instrumental Cover", "2003"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Jashn-E-Bahara (Flute Instrumental)", "Jodhaa Akbar", "A.R. Rahman", "2008"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Tum Hi Ho (Violin/Piano Instrumental)", "Aashiqui 2", "Instrumental Cover", "2013"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Instrumental Covers", "Maiyya Yashoda (Soft Instrumental)", "Hum Saath Saath Hain", "Instrumental Cover", "1999"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Devotional & Mantras", "Krishna Flute Music (Bansuri)", "Devotional", "Traditional Bansuri Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Devotional & Mantras", "Om Jai Jagdish Hare (Instrumental)", "Devotional/Aarti", "Traditional", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Devotional & Mantras", "Gayatri Mantra with Instrumental Music", "Devotional", "Various Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Devotional & Mantras", "Vedic Wedding Mantras with Shehnai", "Devotional", "Various Priests/Artists", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Devotional & Mantras", "Saptha Padi Mantras (Seven Vows)", "Vedic Wedding Rituals", "Vedic Chants", "Traditional"),
    ("Day 3 (12 Dec)", "Wedding Rituals & Pheras", "12 December - Wedding/Pheras Instrumental", "Devotional & Mantras", "Royal Indian Wedding Fusion (Shehnai, Flute, Sitar)", "Classical Fusion", "Various Artists", "Traditional"),

    # 6. Reception Songs
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Tum Kya Mile", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Shreya Ghoshal", "2023"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Chaleya", "Jawan", "Arijit Singh, Shilpa Rao", "2023"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Sajni", "Laapataa Ladies", "Arijit Singh", "2024"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Ve Haaniyaan", "Album", "Aarsh Qaamil", "2022"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Raataan Lambiyan", "Shershaah", "Jubin Nautiyal, Asees Kaur", "2021"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Tera Hone Laga Hoon", "Ajab Prem Ki Ghazab Kahani", "Atif Aslam, Alisha Chinai", "2009"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Kesariya", "Brahmastra", "Arijit Singh", "2022"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Reception Grand Entry", "Aaj Sajeya", "Album", "Goldie Sohel", "2021"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Tauba Tauba", "Bad Newz", "Karan Aujla", "2024"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Aayi Nai", "Stree 2", "Yo Yo Honey Singh, Sachin-Jigar", "2024"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "What Jhumka?", "Rocky Aur Rani Kii Prem Kahaani", "Arijit Singh, Jonita Gandhi", "2023"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Bijuria", "Sunny Sanskari Ki Tulsi Kumari", "Mika Singh", "2024"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Laal Peeli Akhiyaan", "Teri Baaton Mein Aisa Uljha Jiya", "Shadab, Tanishk Bagchi", "2024"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Desi Girl", "Dostana", "Sunidhi Chauhan, Vishal Dadlani", "2008"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Bole Chudiyan", "Kabhi Khushi Kabhie Gham", "Amit Kumar, Sonu Nigam, Alka Yagnik", "2001"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Coca Cola", "Luka Chuppi", "Tony Kakkar, Neha Kakkar", "2019"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Sadi Gali", "Tanu Weds Manu", "Lehmber Hussainpuri", "2011"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Kala Chashma", "Baar Baar Dekho", "Amar Arshi, Badshah, Neha Kakkar", "2016"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "London Thumakda", "Queen", "Labh Janjua, Sonu Kakkar, Neha Kakkar", "2014"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Gallan Goodiyan", "Dil Dhadakne Do", "Shankar Mahadevan, Yashita Sharma", "2015"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Morni Banke", "Badhaai Ho", "Guru Randhawa, Neha Kakkar", "2018"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Nachde Ne Saare", "Baar Baar Dekho", "Jasleen Royal", "2016"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Badtameez Dil", "Yeh Jawaani Hai Deewani", "Benny Dayal", "2013"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Balam Pichkari", "Yeh Jawaani Hai Deewani", "Vishal Dadlani, Shalmali Kholgade", "2013"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Let's Nacho", "Kapoor & Sons", "Nucleya, Badshah, Benny Dayal", "2016"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "The Punjaabban Song", "JugJugg Jeeyo", "Gippy Grewal, Zahrah Khan", "2022"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Cutiepie", "Ae Dil Hai Mushkil", "Pardeep Singh Sran, Nakash Aziz", "2016"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Dilliwaali Girlfriend", "Yeh Jawaani Hai Deewani", "Arijit Singh, Sunidhi Chauhan", "2013"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Lehanga", "Album", "Jass Manak", "2019"),
    ("Day 3 (12 Dec)", "Reception", "12 December - Wedding/Reception", "Dance Floor Anthems", "Radha", "Student of the Year", "Vishal-Shekhar, Shreya Ghoshal", "2012"),

    # 7. Vidaai Songs
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Dilbaro", "Raazi", "Harshdeep Kaur, Vibha Saraf, Shankar Mahadevan", "2018"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Laadki", "Coke Studio / Sachin-Jigar", "Rekha Bhardwaj, Tanvi Shah", "2016"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Kabira (Encore)", "Yeh Jawaani Hai Deewani", "Arijit Singh, Harshdeep Kaur", "2013"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Madhaniya", "Album / Folk", "Neha Bhasin", "2019"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Ghar", "Jab Harry Met Sejal", "Mohit Chauhan, Nikhita Gandhi", "2017"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Din Shagna Da (Farewell Version)", "Phillauri", "Jasleen Royal", "2017"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Modern & Soulful Vidaai", "Nahi Jaana", "Album", "Neha Bhasin", "2020"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Babul Ki Duayen Leti Ja", "Neel Kamal", "Mohammed Rafi", "1968"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Babul Jo Tumne Sikhaya", "Hum Aapke Hain Koun..!", "Lata Mangeshkar", "1994"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Babul (Bidaai Song)", "Baabul", "Richa Sharma", "2006"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Kehta Hai Baabul", "Baabul", "Amitabh Bachchan / Jagjit Singh", "2006"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Chal Ri Sajani Ab Kya Soche", "Bombai Ka Babu", "Mukesh", "1960"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Babul Ka Yeh Ghar", "Daata", "Mohammed Aziz", "1989"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Mehlon Ka Raja Mila", "Raja Ki Aayegi Baraat", "Alka Yagnik", "1996"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Samjhi Thi Ke Yeh Ghar Mera Hai", "Kaajal", "Asha Bhosle", "1965"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Classic Bollywood Vidaai", "Babul Chhodi Na Jaaye", "Bewaffa Se Waffa", "Various Artists", "1992"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "90s-2000s Popular", "Ye Betiyan To Babul Ki", "Na Tum Jaano Na Hum", "Lata Mangeshkar", "2002"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "90s-2000s Popular", "Saajan Ke Ghar Jaana Hai", "Lajja", "Alka Yagnik, Richa Sharma, Sonu Nigam", "2001"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "90s-2000s Popular", "Babul Jo Tum Ne Sikhaya (Folk)", "Vivah Geet Folk", "Sharda Sinha", "Traditional"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Contemporary & Touching", "Madhanya", "Album", "Rahul Vaidya, Asees Kaur", "2021"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Contemporary & Touching", "Sajni (Vidaai Acoustic)", "Laapataa Ladies", "Arijit Singh", "2024"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Contemporary & Touching", "Tujh Mein Rab Dikhta Hai (Slow)", "Rab Ne Bana Di Jodi", "Roop Kumar Rathod", "2008"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Contemporary & Touching", "Chitta Kukkad (Doli Song)", "Punjabi Folk", "Traditional", "Traditional"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Contemporary & Touching", "Chunar (Farewell Theme)", "ABCD 2", "Arijit Singh", "2015"),
    ("Day 3 (12 Dec)", "Vidaai Songs", "12 December - Wedding/Vidaai", "Contemporary & Touching", "Apna Bana Le (Acoustic)", "Bhediya", "Arijit Singh", "2022"),
]

def make_yt_link(name, movie):
    query = f"{name} {movie}".strip()
    return f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"

def make_sp_link(name, singer):
    first_singer = singer.split(",")[0].split("&")[0].split("/")[0].strip()
    query = f"{name} {first_singer}".strip()
    return f"https://open.spotify.com/search/{urllib.parse.quote(query)}"

print(f"Total songs collected: {len(songs_db)}")

# 1. Generate Markdown file for each unique folder
folder_songs = {}
for item in songs_db:
    day, func, rel_path, subcat, name, movie, singer, yr = item
    if rel_path not in folder_songs:
        folder_songs[rel_path] = []
    folder_songs[rel_path].append(item)

for rel_path, items in folder_songs.items():
    full_folder = os.path.join(BASE_DIR, rel_path)
    os.makedirs(full_folder, exist_ok=True)
    folder_name = os.path.basename(rel_path)
    md_file_path = os.path.join(full_folder, f"{folder_name.replace(' ', '_')}.md")
    
    with open(md_file_path, "w", encoding="utf-8") as f:
        f.write(f"# 🎵 {folder_name} - Song Collection\n")
        f.write(f"**Day:** {items[0][0]} | **Function:** {items[0][1]} | **Total Songs:** {len(items)}\n\n")
        f.write("---\n\n")
        f.write("| # | Song Name | Movie / Album | Singer(s) | Year | Subcategory | YouTube | Spotify |\n")
        f.write("|---|-----------|---------------|-----------|------|-------------|---------|---------|\n")
        
        for idx, it in enumerate(items, 1):
            day, func, _, subcat, name, movie, singer, yr = it
            yt = make_yt_link(name, movie)
            sp = make_sp_link(name, singer)
            f.write(f"| {idx} | **{name}** | {movie} | {singer} | {yr} | {subcat} | [▶ YouTube]({yt}) | [🎵 Spotify]({sp}) |\n")
        
        f.write("\n---\n")
        f.write(f"> 💡 **DJ Tip:** This folder contains {len(items)} curated songs specifically tailored for {items[0][1]}. Use YouTube and Spotify links to audition, build playlists, or cue directly during the function.\n")

print(f"Generated Markdown files for all {len(folder_songs)} folders successfully!")

# 2. Generate Master Excel File with openpyxl
wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Scheme
header_fill_day1 = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid") # Gold/Yellow
header_fill_day2 = PatternFill(start_color="B6D7A8", end_color="B6D7A8", fill_type="solid") # Soft Green
header_fill_day3 = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid") # Soft Rose/Pink
header_fill_master = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Navy Blue
header_fill_summary = PatternFill(start_color="333333", end_color="333333", fill_type="solid") # Dark Gray

header_font_white = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
header_font_dark = Font(name="Calibri", size=11, bold=True, color="000000")
data_font = Font(name="Calibri", size=10)
link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
bold_font = Font(name="Calibri", size=10, bold=True)
title_font = Font(name="Calibri", size=16, bold=True, color="1B365D")
sub_font = Font(name="Calibri", size=11, italic=True)

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# Sheet 1: Dashboard / Summary
ws_dash = wb.create_sheet(title="Wedding Music Overview")
ws_dash.views.sheetView[0].showGridLines = True

ws_dash.cell(row=2, column=2, value="Brother's Grand Wedding - Music Collection Master").font = title_font
ws_dash.cell(row=3, column=2, value="Wedding Dates: 10, 11 & 12 December 2026 | Complete DJ & Family Music Library").font = sub_font

headers_summary = ["Day & Date", "Ceremony / Function", "Target Vibe & Style", "Total Songs", "Folder Location"]
for col_idx, h in enumerate(headers_summary, start=2):
    c = ws_dash.cell(row=5, column=col_idx, value=h)
    c.fill = header_fill_summary
    c.font = header_font_white
    c.alignment = Alignment(horizontal="center", vertical="center")

summary_rows = [
    ("Day 1 - 10 Dec 2026", "Traditional Mayan Ceremony", "Devotional, Shehnai, Classical, Tilak Geet", 30, "10 December - Mayan & Haldi\\Traditional Mayan Songs"),
    ("Day 1 - 10 Dec 2026", "Haldi Bollywood Songs", "Upbeat Haldi hits, Yellow theme, Trending 2024-26", 31, "10 December - Mayan & Haldi\\Haldi Bollywood Songs"),
    ("Day 1 - 10 Dec 2026", "Punjabi Haldi Songs", "Vatna folk, Gidda boliyan, Modern Diljit/AP Dhillon", 25, "10 December - Mayan & Haldi\\Punjabi Haldi Songs"),
    ("Day 1 - 10 Dec 2026", "Family Celebration", "All-age group dance, joyful family celebration", 10, "10 December - Mayan & Haldi\\Family Celebration"),
    ("Day 1 - 10 Dec 2026", "Funny & Entertaining Songs", "Humorous teasing, Antakshari, Comedy classics", 11, "10 December - Mayan & Haldi\\Funny & Entertaining Songs"),
    ("Day 2 - 11 Dec 2026", "Mehndi Songs", "Traditional mehndi geet, Bollywood hits, Trending", 20, "11 December - Mehndi & Sangeet\\Mehndi Songs"),
    ("Day 2 - 11 Dec 2026", "Ladies Sangeet", "Folk songs, Dholki geet, In-laws teasing, Suhag", 20, "11 December - Mehndi & Sangeet\\Ladies Sangeet"),
    ("Day 2 - 11 Dec 2026", "Family Dance", "Choreographed family group performances", 20, "11 December - Mehndi & Sangeet\\Family Dance"),
    ("Day 2 - 11 Dec 2026", "Friends Dance", "High energy, Bhangra, Bachelor squad hits", 20, "11 December - Mehndi & Sangeet\\Friends Dance"),
    ("Day 2 - 11 Dec 2026", "Couple Dance", "Romantic duets, Slow dance, First dance", 25, "11 December - Mehndi & Sangeet\\Couple Dance"),
    ("Day 2 - 11 Dec 2026", "DJ Party", "Late night club anthems, Punjabi pop, Bass drops", 20, "11 December - Mehndi & Sangeet\\DJ Party"),
    ("Day 2 - 11 Dec 2026", "Traditional Wedding Songs", "Folk across India (Rajasthani, Punjabi, Maithili)", 10, "11 December - Mehndi & Sangeet\\Traditional Wedding Songs"),
    ("Day 3 - 12 Dec 2026", "Baraat & Groom Entry", "Grand entry, Dhol beats, High voltage anthems", 20, "12 December - Wedding\\Baraat & Groom Entry"),
    ("Day 3 - 12 Dec 2026", "Ghodi & Procession", "Traditional Ghodi Chadai, Sehra Bandi, Folk", 20, "12 December - Wedding\\Ghodi & Procession"),
    ("Day 3 - 12 Dec 2026", "Varmala & Jaimala", "Magical garland exchange, Grand cinematic romantic", 20, "12 December - Wedding\\Varmala & Jaimala"),
    ("Day 3 - 12 Dec 2026", "Bride Entry", "Dreamy bridal walk, Emotional & Royal entrances", 15, "12 December - Wedding\\Bride Entry"),
    ("Day 3 - 12 Dec 2026", "Wedding Rituals & Pheras", "Shehnai, Bansuri, Sitar, Vedic Mantras, Devotional", 25, "12 December - Wedding\\Pheras Instrumental"),
    ("Day 3 - 12 Dec 2026", "Reception", "Grand reception entry & Full dance floor hits", 30, "12 December - Wedding\\Reception"),
    ("Day 3 - 12 Dec 2026", "Vidaai Songs", "Soulful farewell, Classic Rafi/Lata, Touching tears", 25, "12 December - Wedding\\Vidaai"),
]

for row_idx, r in enumerate(summary_rows, start=6):
    for col_idx, val in enumerate(r, start=2):
        c = ws_dash.cell(row=row_idx, column=col_idx, value=val)
        c.font = data_font
        c.border = thin_border
        if col_idx == 5:
            c.alignment = Alignment(horizontal="center")
            c.font = bold_font

# Total summary row
tot_row = len(summary_rows) + 6
ws_dash.cell(row=tot_row, column=2, value="TOTAL CURATED LIBRARY").font = bold_font
ws_dash.cell(row=tot_row, column=3, value="19 Distinct Ceremony Folders").font = bold_font
ws_dash.cell(row=tot_row, column=4, value="Across 3 Days").font = bold_font
c_tot = ws_dash.cell(row=tot_row, column=5, value=f"=SUM(E6:E{tot_row-1})")
c_tot.font = Font(name="Calibri", size=12, bold=True, color="1B365D")
c_tot.alignment = Alignment(horizontal="center")
ws_dash.cell(row=tot_row, column=6, value="Brother Wedding Music/").font = bold_font

for col in range(2, 7):
    ws_dash.cell(row=tot_row, column=col).border = thin_border
    ws_dash.cell(row=tot_row, column=col).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

for col_idx in range(2, 7):
    max_len = max(len(str(ws_dash.cell(row=r, column=col_idx).value or '')) for r in range(5, tot_row+1))
    col_letter = get_column_letter(col_idx)
    ws_dash.column_dimensions[col_letter].width = max(max_len + 4, 15)

# Sheet 2: Master Song List
ws_master = wb.create_sheet(title="Master Song List (All Days)")
ws_master.views.sheetView[0].showGridLines = True

headers_master = ["#", "Day", "Function / Ceremony", "Category", "Song Name", "Movie / Album", "Singer(s)", "Year", "YouTube Link", "Spotify Link"]
for col_idx, h in enumerate(headers_master, start=1):
    c = ws_master.cell(row=1, column=col_idx, value=h)
    c.fill = header_fill_master
    c.font = header_font_white
    c.alignment = Alignment(horizontal="center", vertical="center")

for idx, song in enumerate(songs_db, start=1):
    row_num = idx + 1
    day, func, rel_path, subcat, name, movie, singer, yr = song
    yt = make_yt_link(name, movie)
    sp = make_sp_link(name, singer)
    
    ws_master.cell(row=row_num, column=1, value=idx).alignment = Alignment(horizontal="center")
    ws_master.cell(row=row_num, column=2, value=day)
    ws_master.cell(row=row_num, column=3, value=func)
    ws_master.cell(row=row_num, column=4, value=subcat)
    ws_master.cell(row=row_num, column=5, value=name).font = bold_font
    ws_master.cell(row=row_num, column=6, value=movie)
    ws_master.cell(row=row_num, column=7, value=singer)
    ws_master.cell(row=row_num, column=8, value=yr).alignment = Alignment(horizontal="center")
    
    # YouTube Link
    c_yt = ws_master.cell(row=row_num, column=9, value="YouTube Link")
    c_yt.hyperlink = yt
    c_yt.font = link_font
    c_yt.alignment = Alignment(horizontal="center")
    
    # Spotify Link
    c_sp = ws_master.cell(row=row_num, column=10, value="Spotify Link")
    c_sp.hyperlink = sp
    c_sp.font = link_font
    c_sp.alignment = Alignment(horizontal="center")
    
    for c_i in range(1, 11):
        ws_master.cell(row=row_num, column=c_i).border = thin_border
        if idx % 2 == 0:
            ws_master.cell(row=row_num, column=c_i).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

for col_idx in range(1, 11):
    max_len = max(len(str(ws_master.cell(row=r, column=col_idx).value or '')) for r in range(1, min(len(songs_db)+2, 50)))
    col_letter = get_column_letter(col_idx)
    ws_master.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Sheet 3, 4, 5: Day-wise Sheets
day_configs = [
    ("Day 1 - Mayan & Haldi", "Day 1 (10 Dec)", header_fill_day1),
    ("Day 2 - Mehndi & Sangeet", "Day 2 (11 Dec)", header_fill_day2),
    ("Day 3 - Wedding Day", "Day 3 (12 Dec)", header_fill_day3),
]

for sheet_title, day_filter, fill_color in day_configs:
    ws_day = wb.create_sheet(title=sheet_title)
    ws_day.views.sheetView[0].showGridLines = True
    
    headers_day = ["#", "Function / Ceremony", "Category", "Song Name", "Movie / Album", "Singer(s)", "Year", "YouTube Link", "Spotify Link"]
    for col_idx, h in enumerate(headers_day, start=1):
        c = ws_day.cell(row=1, column=col_idx, value=h)
        c.fill = fill_color
        c.font = header_font_dark
        c.alignment = Alignment(horizontal="center", vertical="center")
    
    day_songs = [s for s in songs_db if s[0] == day_filter]
    for d_idx, song in enumerate(day_songs, start=1):
        row_num = d_idx + 1
        _, func, rel_path, subcat, name, movie, singer, yr = song
        yt = make_yt_link(name, movie)
        sp = make_sp_link(name, singer)
        
        ws_day.cell(row=row_num, column=1, value=d_idx).alignment = Alignment(horizontal="center")
        ws_day.cell(row=row_num, column=2, value=func)
        ws_day.cell(row=row_num, column=3, value=subcat)
        ws_day.cell(row=row_num, column=4, value=name).font = bold_font
        ws_day.cell(row=row_num, column=5, value=movie)
        ws_day.cell(row=row_num, column=6, value=singer)
        ws_day.cell(row=row_num, column=7, value=yr).alignment = Alignment(horizontal="center")
        
        c_yt = ws_day.cell(row=row_num, column=8, value="Play on YouTube")
        c_yt.hyperlink = yt
        c_yt.font = link_font
        c_yt.alignment = Alignment(horizontal="center")
        
        c_sp = ws_day.cell(row=row_num, column=9, value="Play on Spotify")
        c_sp.hyperlink = sp
        c_sp.font = link_font
        c_sp.alignment = Alignment(horizontal="center")
        
        for c_i in range(1, 10):
            ws_day.cell(row=row_num, column=c_i).border = thin_border
            if d_idx % 2 == 0:
                ws_day.cell(row=row_num, column=c_i).fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
                
    for col_idx in range(1, 10):
        max_len = max(len(str(ws_day.cell(row=r, column=col_idx).value or '')) for r in range(1, min(len(day_songs)+2, 40)))
        col_letter = get_column_letter(col_idx)
        ws_day.column_dimensions[col_letter].width = max(max_len + 3, 14)

excel_path = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")
wb.save(excel_path)
print(f"Master Excel file saved successfully to: {excel_path}")
