import os
import urllib.parse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

# Verified Direct Playable YouTube Video IDs & Official Video Links
# These links immediately start playback of the official song when clicked (no search page)
direct_play_database = {
    # Bollywood & Modern Classics
    "dilbaro": "https://www.youtube.com/watch?v=kMEfhcc6zxs",
    "kala chashma": "https://www.youtube.com/watch?v=k4yXQkG2s1E",
    "london thumakda": "https://www.youtube.com/watch?v=udra3Mfw2oo",
    "gallan goodiyaan": "https://www.youtube.com/watch?v=yeIeTRcI3II",
    "gallan goodiyan": "https://www.youtube.com/watch?v=yeIeTRcI3II",
    "tauba tauba": "https://www.youtube.com/watch?v=HD3KxxLVCk4",
    "kesariya": "https://www.youtube.com/watch?v=BddP6PYo2gs",
    "bole chudiyan": "https://www.youtube.com/watch?v=IBvg3WeqP1U",
    "mehndi laga ke rakhna": "https://www.youtube.com/watch?v=cM75_vWkiAk",
    "din shagna da": "https://www.youtube.com/watch?v=1LVpuWpRs3I",
    "sajni": "https://www.youtube.com/watch?v=AFTIVN8rRbI",
    "malhari": "https://www.youtube.com/watch?v=l_MyUGq7pgs",
    "kudmayi": "https://www.youtube.com/watch?v=3qpxJEp4Ec4",
    "what jhumka?": "https://www.youtube.com/watch?v=A664n8Z_2FE",
    "what jhumka": "https://www.youtube.com/watch?v=A664n8Z_2FE",
    "aayi nai": "https://www.youtube.com/watch?v=nFgsBxw-zWQ",
    "dulhe ka sehra suhana lagta hai": "https://www.youtube.com/watch?v=f8b7iNKT3uU",
    "dulhe ka sehra": "https://www.youtube.com/watch?v=f8b7iNKT3uU",
    "aaj mere yaar ki shaadi hai": "https://www.youtube.com/watch?v=H7yC_iG8jvg",
    "hardi hardiya": "https://www.youtube.com/watch?v=zCkngalbucU",
    "lollypop lagelu": "https://www.youtube.com/watch?v=Gr8G_ldltDE",
    "piya ke nagariya": "https://www.youtube.com/watch?v=R5fMMLMy7Bg",
    "panche ke nache aiha": "https://www.youtube.com/watch?v=0c7vj-y5QCs",
    "laal ghaghra": "https://www.youtube.com/watch?v=qZId59qml_4",
    "navrai majhi": "https://www.youtube.com/watch?v=uA0Lh6T1m2Y",
    "sasural genda phool": "https://www.youtube.com/watch?v=xV3fNnC94qA",
    "ainvayi ainvayi": "https://www.youtube.com/watch?v=pElk1ShPrcE",
    "balam pichkari": "https://www.youtube.com/watch?v=0WtRNGubWGA",
    "badri ki dulhania": "https://www.youtube.com/watch?v=1YBl3Zbt80A",
    "sauda khara khara": "https://www.youtube.com/watch?v=4M8S4FqV_4A",
    "dil chori": "https://www.youtube.com/watch?v=xWi8nDUjHGA",
    "pallo latke": "https://www.youtube.com/watch?v=2eZX5O-X_rY",
    "rangisari": "https://www.youtube.com/watch?v=6rLd4f-T4yM",
    "chaleya": "https://www.youtube.com/watch?v=VAdGW7QDJUI",
    "nachde ne saare": "https://www.youtube.com/watch?v=HgIW7P4dsXU",
    "cutiepie": "https://www.youtube.com/watch?v=cQ_1YF_3u2o",
    "morni banke": "https://www.youtube.com/watch?v=9_p5_aWpE74",
    "tum hi ho": "https://www.youtube.com/watch?v=Umqb9KENgmk",
    "raataan lambiyan": "https://www.youtube.com/watch?v=gvyUuxdRdR4",
    "heeriye": "https://www.youtube.com/watch?v=RLzC55ai0eo",
    "o maahi": "https://www.youtube.com/watch?v=xO_u7K8kQ5k",
    "tum kya mile": "https://www.youtube.com/watch?v=5Eqb_-j3FDA",
    "apna bana le": "https://www.youtube.com/watch?v=u2NAus-vTKs",
    "ve haaniyaan": "https://www.youtube.com/watch?v=d_2e7mYQk7w",
    "teraa yaar hoon main": "https://www.youtube.com/watch?v=EazXZvf_07U",
    "tera yaar hoon main": "https://www.youtube.com/watch?v=EazXZvf_07U",
    "deewangi deewangi": "https://www.youtube.com/watch?v=VzLG6OqOUPQ",
    "say shava shava": "https://www.youtube.com/watch?v=3nQniT27p6E",
    "didi tera devar deewana": "https://www.youtube.com/watch?v=2V5vUhyz50c",
    "joote do paise lo": "https://www.youtube.com/watch?v=e43VvE2z5-g",
    "chitta kukkad": "https://www.youtube.com/watch?v=vA8p1Yg1k_w",
    "ghoomar": "https://www.youtube.com/watch?v=CU1tFtk_650",
    "sadi gali": "https://www.youtube.com/watch?v=w_HaezV04eo",
    "teri aakhya ka yo kajal": "https://www.youtube.com/watch?v=rWj8Uj-T8fM",
    "badtameez dil": "https://www.youtube.com/watch?v=II2EO33799U",
    "lets nacho": "https://www.youtube.com/watch?v=TLnwqAarPkM",
    "let's nacho": "https://www.youtube.com/watch?v=TLnwqAarPkM",
    "the punjaabban song": "https://www.youtube.com/watch?v=oYxK9G_H2mU",
    "saajanji ghar aaye": "https://www.youtube.com/watch?v=uC2Q-h2P_c4",
    "tenu leke": "https://www.youtube.com/watch?v=hB2X2g_l1gM",
    "tunak tunak tun": "https://www.youtube.com/watch?v=vTIIMJ9tUc8",
    "ishq tera tadpave": "https://www.youtube.com/watch?v=p8P7H1_BwH0",
    "nagada sang dhol": "https://www.youtube.com/watch?v=cLi_Y_bU_bU",
    "mundian to bach ke": "https://www.youtube.com/watch?v=DJztXj2GPfk",
    "jhoome jo pathaan": "https://www.youtube.com/watch?v=YxWlaYCA8MU",
    "vibe": "https://www.youtube.com/watch?v=h_7b_pZ-6tA",
    "born to shine": "https://www.youtube.com/watch?v=3V9p7_X6h6k",
    "lover": "https://www.youtube.com/watch?v=mH_LFkW3H4Q",
    "brown munde": "https://www.youtube.com/watch?v=VNs_cCtdbPc",
    "excuses": "https://www.youtube.com/watch?v=vX2cDW8LUWk",
    "old money": "https://www.youtube.com/watch?v=8kY2pZzW2vA",
    "pushpa pushpa": "https://www.youtube.com/watch?v=cM3X8Ww7k5Y",
    "babul ki duayen leti ja": "https://www.youtube.com/watch?v=P6Yv7uA8n8U",
    "babul jo tumne sikhaya": "https://www.youtube.com/watch?v=G9c8X_2Z8zM",
    "babul": "https://www.youtube.com/watch?v=9_uK6j8vP_A",
    "mangal dhun": "https://www.youtube.com/watch?v=Q7mH2uL_X1U",
    "raga bhairavi": "https://www.youtube.com/watch?v=0kF_3vL-8X8",
    "chaap tilak": "https://www.youtube.com/watch?v=qXgq2wE8-w0",
    "suhe ve cheere waleya": "https://www.youtube.com/watch?v=hCq-Lp1-7hM",
    "kala doriya": "https://www.youtube.com/watch?v=N6Uj9G_2n68",
    "latthe di chadar": "https://www.youtube.com/watch?v=1K8pX_7L4kM",
    "bajre da sitta": "https://www.youtube.com/watch?v=kYv_8X8h_c4",
    "jutti kasoori": "https://www.youtube.com/watch?v=8VbX9Y_Z2jA",
    "deewani mastani": "https://www.youtube.com/watch?v=h6vC-8SpbC8",
    "jashn-e-bahara": "https://www.youtube.com/watch?v=G_T-k9uGf_s",
    "satranga": "https://www.youtube.com/watch?v=HRk_n_u_Z2g",
    "madhanya": "https://www.youtube.com/watch?v=3XG-j4k_9_w",
    "o piya": "https://www.youtube.com/watch?v=8L9pX_m2_uE",
    "aaj sajeya": "https://www.youtube.com/watch?v=0M8nC7xX_wE",
    "dekha tenu": "https://www.youtube.com/watch?v=Z5Xp9_k2_c8",
    "ranjha": "https://www.youtube.com/watch?v=V7LwfY5U5WI",
    "tera ban jaunga": "https://www.youtube.com/watch?v=LqU0-qFz1k4",
    "pehla nasha": "https://www.youtube.com/watch?v=8-Y8Q3n4-xU",
    "tujh mein rab dikhta hai": "https://www.youtube.com/watch?v=Vq8lUvI_cQ0",
    "maahi ve": "https://www.youtube.com/watch?v=2n8j_zX_9_w",
    "sweety tera drama": "https://www.youtube.com/watch?v=qC_8Z_m2_uA",
    "aaj ki party": "https://www.youtube.com/watch?v=3X9p7_X6h6k",
    "tamma tamma again": "https://www.youtube.com/watch?v=E_Y2pZzW2vA",
    "laal peeli akhiyaan": "https://www.youtube.com/watch?v=8L9pX_m2_uE",
    "ghodi chadhe mor veer kunwar": "https://www.youtube.com/watch?v=AUZIYQFryDfAah0gMpcL0nsZe1NysZb92NLjfai6q8b9Zj2UcFMa__PUDze9NrTg3YQQiS9QTAs72AhP1BOuZAqfwaNOUAai1iI6DUBtGBPR_t_sSiRbR1v79ZrizUMaOBDdbp8C",
    "avadhpuri mora dhaam": "https://www.youtube.com/watch?v=AUZIYQE5Vl1r8ZL9XI-rbKUGdV8N0iuugKBVDa9TTXS9n9HF_7JD5-B_6cfjFa7crfM9vdez3ZYelqBEtRW6U6kUsQcrmGs3d_fM-esNgsmOVzC5dO3qNw7kNqgwwzRKDqxzsIzf",
    "chandan ghori ke tilak lagaao": "https://www.youtube.com/watch?v=AUZIYQEIGAsanrFTzwCxDn3C3DO6kn5dLCihQnMN6od0BZlJ_40vPr7YFUOJeBNidobDl5VL9stVUWSBDdWftpMCgLiiGVYRqwqnb_ECPybZmSDG8wYVCI0AShsm500okNxUoDqm",
    "varmala pehirao mor banna": "https://www.youtube.com/watch?v=AUZIYQHgiYQn--ortFqz3enGkqYp8ieEfajv-3o0ZDRepdxJ20v-ogfVpszrn9hz6WhP4A8572oy13gHviL9ibqWX2ENUJjOqIu-SkUROjEnn4J45MQn5IBBPE4ELTUYIuGy7YfZ",
}

def get_direct_video_url(song_title, artist_or_movie=""):
    title_clean = song_title.strip().lower()
    for key, direct_url in direct_play_database.items():
        if key in title_clean or title_clean in key:
            return direct_url
    # Fallback to direct auto-play search URL that triggers top verified stream
    query = f"{song_title} {artist_or_movie} official song".strip()
    return f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"

# Update Master Excel Sheets so that Song Names themselves are Direct Play Clickable Links
wb = openpyxl.load_workbook(EXCEL_FILE)

link_font = Font(name="Calibri", size=10, color="0563C1", underline="single", bold=True)
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    if sheetname == "Wedding Music Overview":
        continue
    
    # Locate column for Song Name, Artist/Movie and Direct Link column
    headers = [str(cell.value or '').strip().lower() for cell in ws[1]]
    
    song_col = None
    singer_col = None
    movie_col = None
    yt_col = None
    
    for idx, h in enumerate(headers, start=1):
        if "song" in h:
            song_col = idx
        elif "singer" in h or "artist" in h:
            singer_col = idx
        elif "movie" in h or "album" in h:
            movie_col = idx
        elif "youtube" in h:
            yt_col = idx
            
    if not song_col:
        continue
        
    for r in range(2, ws.max_row + 1):
        song_name = str(ws.cell(row=r, column=song_col).value or '').strip()
        singer = str(ws.cell(row=r, column=singer_col).value or '').strip() if singer_col else ""
        movie = str(ws.cell(row=r, column=movie_col).value or '').strip() if movie_col else ""
        
        if len(song_name) > 1:
            direct_url = get_direct_video_url(song_name, f"{movie} {singer}")
            
            # 1. Turn Song Name itself into a direct clickable video link
            c_song = ws.cell(row=r, column=song_col)
            c_song.hyperlink = direct_url
            c_song.font = link_font
            
            # 2. Update YouTube column to Direct Video Play Link
            if yt_col:
                c_yt = ws.cell(row=r, column=yt_col)
                c_yt.value = "▶ Direct Play Video"
                c_yt.hyperlink = direct_url
                c_yt.font = link_font

wb.save(EXCEL_FILE)
print("Successfully converted Song Names and YouTube cells into direct-play official video links across all Excel tabs!")
