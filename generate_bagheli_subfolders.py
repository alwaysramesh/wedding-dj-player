import os
import urllib.parse

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
BAGHELI_DIR = os.path.join(BASE_DIR, "Bagheli Wedding Songs")

subfolders = [
    "01 - Mayan & Mandap Pujan",
    "02 - Haldi & Nahchhu",
    "03 - Mehndi & Ladies Sangeet",
    "04 - Baraat & Ghodi Chadai",
    "05 - Pheras, Varmala & Bidaai"
]

for sf in subfolders:
    os.makedirs(os.path.join(BAGHELI_DIR, sf), exist_ok=True)

# Detailed ceremony songs with lyrics previews & cultural notes
bagheli_detailed_data = {
    "01 - Mayan & Mandap Pujan": [
        {
            "name": "Avadhpuri Mora Dhaam (Janakpur Byahan Jaiyo)",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Tilak & Mayan Ceremony",
            "notes": "Traditional Bagheli Banra comparing the groom to Lord Ram leaving Ayodhya to marry Sita in Janakpur.",
            "lyrics_preview": "अवधपुर मोरा धाम, जनकपुर ब्याहन जइयो रे... बन्ना मोर सलोना..."
        },
        {
            "name": "Aaju Janakpur Marwa Suhaawan Lage",
            "artist": "Urmila Shrivastava / Folk Chorus",
            "type": "Evergreen Folk Classic",
            "ritual": "Mandap Chhavan & Pujan",
            "notes": "Sung when the sacred bamboo and mango leaves mandap is erected in the courtyard.",
            "lyrics_preview": "आजु जनकपुर मड़वा सुहावन लागे, हरी-हरी दूबि चढ़ावे सखी मिल..."
        },
        {
            "name": "Gauri Ke Lala Shubh Din Aaye",
            "artist": "Rambabu Jha / Mandali",
            "type": "Devotional Vivah Classic",
            "ritual": "Ganesh Sthapana & Kuldevi Pujan",
            "notes": "Invoking Lord Ganesha and Kuldevi for an obstacle-free (Nirvighna) wedding.",
            "lyrics_preview": "गौरी के लाला शुभ दिन आये, प्रथम मनाइये गणेश मनाइये..."
        },
        {
            "name": "Bhaiya Ka Tilak Chadhawa Aaye",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Tilak Ceremony",
            "notes": "Paternal relatives applying Chandan and Akshat on the groom's forehead.",
            "lyrics_preview": "भइया का तिलक चढ़ावा आये, माथे तिलक सोहै मोरे बीरन के..."
        },
        {
            "name": "Sat Seek Ke Pinjadwa (Bagheli Banra)",
            "artist": "Traditional Rewa Artists",
            "type": "Evergreen Folk Classic",
            "ritual": "Traditional Mayan Blessing",
            "notes": "One of the oldest Bagheli folk meters describing the graceful home and upbringing of the groom.",
            "lyrics_preview": "सात सीक के पिंजड़वा, बोले सुगनवा रे... बन्ना मोर प्यारा..."
        },
        {
            "name": "Mangal Baaje Baghelkhand Mein",
            "artist": "Abhishek Rashik, Anuradha Bagheli",
            "type": "Trending (2024-2026)",
            "ritual": "Modern Acoustic Mayan Entrance",
            "notes": "Modern acoustic production with Shehnai and Dholak, ideal for family morning tea.",
            "lyrics_preview": "मंगल बाजे बघेलखण्ड में, रेवा नगरिया आनंद छायो..."
        },
        {
            "name": "Tilak Chadhau Mor Banna",
            "artist": "Subham Tiwari & Vandana",
            "type": "Trending (2024-2026)",
            "ritual": "Reel & Youngster Favorite",
            "notes": "Trending across Vindhyan Instagram reels and weddings in Rewa and Satna.",
            "lyrics_preview": "तिलक चढ़ाऊ मोर बन्ना, सोने के थाली में अक्षत सजाई..."
        }
    ],
    "02 - Haldi & Nahchhu": [
        {
            "name": "Hardi Chadhawa Mor Bhaiya Ke",
            "artist": "Rewa Mahila Mandali",
            "type": "Evergreen Folk Classic",
            "ritual": "Tel-Baan & Haldi",
            "notes": "Elder aunts and bhabhis applying fragrant haldi and mustard oil ubtan.",
            "lyrics_preview": "हरदी चढ़ावा मोर भइया के, अँगना में मंगल गाईं सखी..."
        },
        {
            "name": "Piyari Range Mor Dulha Ke Chola",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Yellow Cloth / Haldi Kurta",
            "notes": "Celebrates the auspicious yellow robe and turmeric glow of the groom.",
            "lyrics_preview": "पियरी रंगे मोर दूल्हा के चोला, केसर हरदी सोहै रे..."
        },
        {
            "name": "Nahchhu Gaavahu Sakhiyaan Mil Ke",
            "artist": "Rewa Folk Troupe",
            "type": "Evergreen Folk Classic",
            "ritual": "Nahchhu Ceremonial Bath",
            "notes": "Auspicious folk ritual of nail paring and ceremonial bath before donning wedding silks.",
            "lyrics_preview": "नहछू गावहु सखियां मिल के, नउवा बोलाइ लियो रे..."
        },
        {
            "name": "Chauka Purao Dhani Maai",
            "artist": "Urmila Shrivastava",
            "type": "Evergreen Folk Classic",
            "ritual": "Chauka Pujan",
            "notes": "Drawing flour and turmeric alpana patterns on the floor for the groom's seat.",
            "lyrics_preview": "चौका पुराओ धनि माई, चौक पुराई के आसन लगाओ..."
        },
        {
            "name": "Hari-Hari Dubhi Chadhai",
            "artist": "Traditional Rewa Mandali",
            "type": "Evergreen Folk Classic",
            "ritual": "Baan Ritual",
            "notes": "Grass tips dipped in oil and milk touched to the groom's feet, knees, and shoulders.",
            "lyrics_preview": "हरी-हरी दूबि चढ़ाई, तेलवा चढ़ाईं सखी मिल के..."
        },
        {
            "name": "Hardi Lagau Sajna Ke (Bagheli Mix)",
            "artist": "Avinash Tiwari & Anuradha",
            "type": "Trending (2024-2026)",
            "ritual": "Upbeat Haldi Celebration",
            "notes": "Lively Bagheli beat that gets cousins and bhabhis dancing while applying haldi.",
            "lyrics_preview": "हरदी लगाऊँ सजना के, रीवा सहरिया में शोर मचायो..."
        },
        {
            "name": "Banna Ke Ang Ang Hardi Sajai",
            "artist": "Kanha Mishra, Shradha Khare",
            "type": "Trending (2024-2026)",
            "ritual": "Haldi Floral Shower Drop",
            "notes": "High energy modern Vindhyan beat perfect for the marigold flower shower.",
            "lyrics_preview": "बन्ना के अंग अंग हरदी सजाई, फूल बरसाओ रीवा नगरिया..."
        }
    ],
    "03 - Mehndi & Ladies Sangeet": [
        {
            "name": "Kahe Par Rachahu Mehndiya A Sakhi",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Traditional Dholak Mehndi",
            "notes": "The soothing signature Bagheli henna folk song asking about love and beauty.",
            "lyrics_preview": "काहे पर रचहुँ मेंहदियां ए सखी, हथवा में रची है सुहाग..."
        },
        {
            "name": "Haathan Mehndi Chamke Banna Tore Naam Ki",
            "artist": "Rewa Mahila Sangrah",
            "type": "Evergreen Folk Classic",
            "ritual": "Henna Pattern Blessings",
            "notes": "Sung by mothers and grandmothers during intricate bridal/family henna application.",
            "lyrics_preview": "हाथन मेंहदी चमके बन्ना तोरे नाम की, रंग गाढ़ा चढ़ायो रे..."
        },
        {
            "name": "Sasural Mein Bhabhi Tore Raj Ho",
            "artist": "Urmila Shrivastava",
            "type": "Evergreen Folk Classic",
            "ritual": "Family Celebration",
            "notes": "Affectionate teasing welcoming the new sister-in-law into the family dynasty.",
            "lyrics_preview": "ससुराल में भाभी तोरे राज हो, खुशियां मनावै सारा खानदान..."
        },
        {
            "name": "Banna Bada Albela (Dholak Geet)",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Gentle Sangeet Teasing",
            "notes": "Playful, dignified banter about the groom's handsome poise and charm.",
            "lyrics_preview": "बन्ना बड़ा अलबेला, रीवा का छैला, नैना कजरारे सोहैं..."
        },
        {
            "name": "Mhare Aanganwa Chhan Chhan Baaje",
            "artist": "Traditional Mandali",
            "type": "Evergreen Folk Classic",
            "ritual": "Women Sangeet Dance",
            "notes": "Folk dholak rhythm describing tinkling anklets (painjaniya) in the ancestral home.",
            "lyrics_preview": "म्‍हारे अँगना छन-छन बाजे पैंजनिया, नाचे सकल परिवार..."
        },
        {
            "name": "Mehndi Rachi Vindhya Nagari",
            "artist": "Avinash Tiwari & Shradha Khare",
            "type": "Trending (2024-2026)",
            "ritual": "Stage Performance Choreography",
            "notes": "Vindhya folk-pop hit ideal for sister/bhabhi sangeet stage dance routine.",
            "lyrics_preview": "मेंहदी रची विंध्य नगरी, ढोलक बाजे रीवा शहर..."
        },
        {
            "name": "Banno Rani Chamke Ri (Bagheli Pop)",
            "artist": "Kanha Mishra & Anuradha",
            "type": "Trending (2024-2026)",
            "ritual": "Youngster Sangeet Finale",
            "notes": "Fast tempo festive beat with modern keyboard and Bagheli acoustic dholak.",
            "lyrics_preview": "बन्नो रानी चमके री, रूप अनूप सलोना लागे..."
        }
    ],
    "04 - Baraat & Ghodi Chadai": [
        {
            "name": "Ghodi Chadhe Mor Veer Kunwar",
            "artist": "Rewa Folk Troupe",
            "type": "Evergreen Folk Classic",
            "ritual": "Ghodi Chadai",
            "notes": "Sister feeding grains to the mare and holding the reins as the brother mounts.",
            "lyrics_preview": "घोड़ी चढ़े मोर वीर कुंवर, बहिन दुलारे घोड़ी के लगाम..."
        },
        {
            "name": "Sehra Sajao Mor Banna Ke Maathe",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Sehra Bandi",
            "notes": "Golden sehra tied with blessings from mother, father, chacha, and mama.",
            "lyrics_preview": "सेहरा सजाओ मोर बन्ना के माथे, मोती झरे झलमल-झलमल..."
        },
        {
            "name": "Banna Chale Sasural Byahan",
            "artist": "Traditional Artists",
            "type": "Evergreen Folk Classic",
            "ritual": "Baraat Departure",
            "notes": "Procession moving forward amidst brass bands and dholak beats.",
            "lyrics_preview": "बन्ना चले ससुराल ब्याहन, संग में बराती हजार..."
        },
        {
            "name": "Dwar Pe Aaye Mor Dulhe Raja (Agwani)",
            "artist": "Urmila Shrivastava",
            "type": "Evergreen Folk Classic",
            "ritual": "Dwarachar & Toran Welcome",
            "notes": "Mother-in-law performing traditional aarti and welcoming the groom at the venue gate.",
            "lyrics_preview": "द्वार पे आये मोर दूल्हे राजा, आरती उतारो सासू माई..."
        },
        {
            "name": "Rewa Ka Banna Aaye Re (DJ Mix)",
            "artist": "Avinash Tiwari ft. DJ Rewa",
            "type": "Trending (2024-2026)",
            "ritual": "Heavy DJ Baraat Drop",
            "notes": "The #1 trending Bagheli DJ truck track across Rewa, Satna, and Sidhi baraats.",
            "lyrics_preview": "रीवा का बन्ना आये रे, ढोल बजाओ नगाड़ा बजाओ..."
        },
        {
            "name": "Dulha Bane Mor Veer (Baraat Blast)",
            "artist": "Kanha Mishra, Subham Tiwari",
            "type": "Trending (2024-2026)",
            "ritual": "Baraat Street Dance",
            "notes": "Dhol-tasha fusion track for cousins and friends dancing ahead of the groom's car/chariot.",
            "lyrics_preview": "दूल्हा बने मोर वीर, नाचे सारी विंध्य की धरती..."
        }
    ],
    "05 - Pheras, Varmala & Bidaai": [
        {
            "name": "Varmala Pehirao Mor Banna",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Jaimala Ceremony",
            "notes": "Garland exchange moment highlighting traditional reverence and joy.",
            "lyrics_preview": "वरमाला पहिराओ मोर बन्ना, सीताराम की जोड़ी लागे..."
        },
        {
            "name": "Sindoor Daan Karahu Mor Raja",
            "artist": "Urmila Shrivastava",
            "type": "Evergreen Folk Classic",
            "ritual": "Sindoor Daan & Mangalsutra",
            "notes": "Deeply emotional and sacred moment when the groom fills the bride's maang.",
            "lyrics_preview": "सिंदूर दान करहु मोर राजा, अमर रहे सुहाग तोरा..."
        },
        {
            "name": "Saat Phere Lehu Banna-Banni",
            "artist": "Traditional Pundit Mandali",
            "type": "Evergreen Folk Classic",
            "ritual": "Saat Pheras / Agni Pradakshina",
            "notes": "Gentle acoustic accompaniment honoring the seven sacred vows around the holy fire.",
            "lyrics_preview": "सात फेरे लेहु बन्ना-बन्नी, धर्म अर्थ काम मोक्ष संगी..."
        },
        {
            "name": "Babul Ka Ghar Chhode Dulhaniya (Bidaai)",
            "artist": "Vandana Mishra",
            "type": "Evergreen Folk Classic",
            "ritual": "Bidaai Farewell",
            "notes": "Touching separation song sung as the bride bids tearful farewell to childhood home.",
            "lyrics_preview": "बाबुल का घर छोड़े दुल्हिनिया, नयना सजल भये माई-बाप के..."
        },
        {
            "name": "Kahe Ko Byahe Bides (Bagheli Folk)",
            "artist": "Rewa Lok Mandali",
            "type": "Evergreen Folk Classic",
            "ritual": "Doli Departure",
            "notes": "Classic Amir Khusro folk adaptation deeply rooted in Vindhya folk singing traditions.",
            "lyrics_preview": "काहे को ब्याहे बिदेस, हम तो बाबुल तोरे बागन की कोयलिया..."
        },
        {
            "name": "Sajani Bida Hoti Jaai (Emotional)",
            "artist": "Anuradha Bagheli",
            "type": "Trending (2024-2026)",
            "ritual": "Modern Vidaai Ballad",
            "notes": "Gentle flute and sarangi fusion ballad capturing the dignity and love of parents.",
            "lyrics_preview": "सजनी बिदा होती जाई, दुआ देवे सारा गांव..."
        },
        {
            "name": "Banna Banni Ke Shubh Vivah",
            "artist": "Subham Tiwari & Vandana",
            "type": "Trending (2024-2026)",
            "ritual": "Wedding Blessings Finale",
            "notes": "Joyful blessing anthem played as the newlyweds take blessings from elders.",
            "lyrics_preview": "बन्ना-बन्नी के शुभ विवाह, जुग-जुग जियो सुहागन रानी..."
        }
    ]
}

def make_yt_link(name, artist):
    q = f"{name} {artist} Bagheli Vivah Geet".strip()
    return f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(q)}"

for folder_name, songs in bagheli_detailed_data.items():
    folder_path = os.path.join(BAGHELI_DIR, folder_name)
    readme_path = os.path.join(folder_path, "README.md")
    
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(f"# 🪔 {folder_name} - Bagheli Wedding Songs\n")
        f.write(f"**Region:** Rewa & Vindhya (बघेलखण्ड) | **Total Songs:** {len(songs)}\n\n")
        f.write("---\n\n")
        f.write("| # | Song Name (बघेली बोल) | Type | Artist(s) | Ritual Context | Lyrics Preview | Direct YouTube Link |\n")
        f.write("|---|------------------------|------|-----------|----------------|----------------|---------------------|\n")
        
        for idx, s in enumerate(songs, 1):
            yt = make_yt_link(s['name'], s['artist'])
            f.write(f"| {idx} | **{s['name']}** | {s['type']} | {s['artist']} | {s['ritual']} | *{s['lyrics_preview']}* | [▶ Play on YouTube]({yt}) |\n")
            
        f.write("\n---\n")
        f.write("### 📌 Cultural & DJ Guidance for this ceremony:\n")
        for s in songs:
            f.write(f"- **{s['name']}**: {s['notes']}\n")

print("Created individual ceremony READMEs with lyrics and cultural context!")

