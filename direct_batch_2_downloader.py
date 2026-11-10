import os
import re
import yt_dlp

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# Batch 2: Curated High-Priority Wedding Anthems across all Ceremonies
curated_batch_2 = [
    # Mayan & Devotional
    ("10 December - Mayan & Haldi\\Traditional Mayan Songs", "Maiya Yashoda - Kavita Krishnamurthy", "ytsearch1:Maiya Yashoda Yeh Tera Kanhaiya Hum Saath Saath Hain Official Song"),
    ("10 December - Mayan & Haldi\\Traditional Mayan Songs", "Chhote Chhote Bhaiyon Ke - Kavita Krishnamurthy", "ytsearch1:Chhote Chhote Bhaiyon Ke Bade Bhaiya Hum Saath Saath Hain Official Song"),
    ("10 December - Mayan & Haldi\\Traditional Mayan Songs", "Chaap Tilak - Rahul Vaidya", "https://www.youtube.com/watch?v=qXgq2wE8-w0"),
    ("10 December - Mayan & Haldi\\Traditional Mayan Songs", "Shri Ram Janki Vivah Mangalgeet - Sharma Bandhu", "ytsearch1:Shri Ram Janki Vivah Mangalgeet Sharma Bandhu"),

    # Haldi Ceremony
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "Laal Peeli Akhiyaan - Shadab Faridi", "ytsearch1:Laal Peeli Akhiyaan Teri Baaton Mein Aisa Uljha Jiya Official Video"),
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "Balam Pichkari - Vishal Dadlani", "ytsearch1:Balam Pichkari Yeh Jawaani Hai Deewani Official Video"),
    ("10 December - Mayan & Haldi\\Haldi Bollywood Songs", "Rangisari - Kanishk Seth", "ytsearch1:Rangisari Jugjugg Jeeyo Official Video"),
    ("10 December - Mayan & Haldi\\Punjabi Haldi Songs", "Wah Wah Vattna - Traditional", "ytsearch1:Wah Wah Vattna Katore Da Traditional Punjabi Folk Song"),
    ("10 December - Mayan & Haldi\\Punjabi Haldi Songs", "Suhe Ve Cheere Waleya - Surinder Kaur", "ytsearch1:Suhe Ve Cheere Waleya Surinder Kaur Official"),
    ("10 December - Mayan & Haldi\\Punjabi Haldi Songs", "Born to Shine - Diljit Dosanjh", "ytsearch1:Born to Shine Diljit Dosanjh Official Video"),
    ("10 December - Mayan & Haldi\\Family Celebration", "Yeh Ladka Hai Deewana - Udit Narayan", "ytsearch1:Yeh Ladka Hai Deewana Kuch Kuch Hota Hai Official Video"),

    # Mehndi & Sangeet
    ("11 December - Mehndi & Sangeet\\Mehndi Songs", "Mehendi Rang Layee - Udit Narayan", "ytsearch1:Mehendi Rang Layee Chal Mere Bhai Official Video"),
    ("11 December - Mehndi & Sangeet\\Ladies Sangeet", "Didi Tera Devar Deewana - Lata Mangeshkar", "ytsearch1:Didi Tera Devar Deewana Hum Aapke Hain Koun Official Video"),
    ("11 December - Mehndi & Sangeet\\Ladies Sangeet", "Ghoomar - Shreya Ghoshal", "ytsearch1:Ghoomar Padmaavat Official Video"),
    ("11 December - Mehndi & Sangeet\\Couple Dance", "Tum Hi Ho - Arijit Singh", "ytsearch1:Tum Hi Ho Aashiqui 2 Official Video"),
    ("11 December - Mehndi & Sangeet\\Couple Dance", "Raataan Lambiyan - Jubin Nautiyal", "ytsearch1:Raataan Lambiyan Shershaah Official Video"),
    ("11 December - Mehndi & Sangeet\\Couple Dance", "Heeriye - Jasleen Royal", "ytsearch1:Heeriye Jasleen Royal Arijit Singh Official Video"),
    ("11 December - Mehndi & Sangeet\\Couple Dance", "Main Rang Sharbaton Ka - Atif Aslam", "ytsearch1:Main Rang Sharbaton Ka Official Video"),
    ("11 December - Mehndi & Sangeet\\Friends Dance", "Tera Yaar Hoon Main - Arijit Singh", "ytsearch1:Tera Yaar Hoon Main Sonu Ke Titu Ki Sweety Official Video"),
    ("11 December - Mehndi & Sangeet\\Friends Dance", "Deewangi Deewangi - Om Shanti Om", "ytsearch1:Deewangi Deewangi Om Shanti Om Official Video"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "Ishq Mitaye - Mohit Chauhan", "ytsearch1:Ishq Mitaye Main Hoon Punjab Amar Singh Chamkila Official Video"),
    ("11 December - Mehndi & Sangeet\\DJ Party", "The Punjaabban Song - Gippy Grewal", "ytsearch1:The Punjaabban Song JugJugg Jeeyo Official Video"),

    # Baraat & Wedding Day
    ("12 December - Wedding\\Baraat & Groom Entry", "Saajanji Ghar Aaye - Kumar Sanu", "ytsearch1:Saajanji Ghar Aaye Kuch Kuch Hota Hai Official Video"),
    ("12 December - Wedding\\Baraat & Groom Entry", "Tenu Leke - Sonu Nigam", "ytsearch1:Tenu Leke Salaam-e-Ishq Official Video"),
    ("12 December - Wedding\\Baraat & Groom Entry", "Tunak Tunak Tun - Daler Mehndi", "ytsearch1:Tunak Tunak Tun Daler Mehndi Official Video"),
    ("12 December - Wedding\\Baraat & Groom Entry", "Ishq Tera Tadpave - Sukhbir", "ytsearch1:Ishq Tera Tadpave Sukhbir Official Video Oh Ho Ho Ho"),
    ("12 December - Wedding\\Varmala & Jaimala", "Satranga - Arijit Singh", "ytsearch1:Satranga Animal Arijit Singh Official Video"),
    ("12 December - Wedding\\Varmala & Jaimala", "Aaj Sajeya - Goldie Sohel", "ytsearch1:Aaj Sajeya Goldie Sohel Official Video"),
    ("12 December - Wedding\\Bride Entry", "Taare Hain Baaraati - Kumar Sanu", "ytsearch1:Taare Hain Baaraati Chandani Hai Ye Baraat Virasat Official Video"),
    ("12 December - Wedding\\Pheras Instrumental", "Mangal Parinay Shehnai - Bismillah Khan", "ytsearch1:Mangal Parinay Bismillah Khan Shehnai Official"),
    ("12 December - Wedding\\Pheras Instrumental", "Raag Yaman - Hariprasad Chaurasia", "ytsearch1:Raag Yaman Hariprasad Chaurasia Flute Classical"),
    ("12 December - Wedding\\Reception", "Chaleya - Arijit Singh", "ytsearch1:Chaleya Jawan Arijit Singh Official Video"),
    ("12 December - Wedding\\Vidaai", "Babul Ki Duayen Leti Ja - Mohammed Rafi", "ytsearch1:Babul Ki Duayen Leti Ja Neel Kamal Mohammed Rafi Official Video"),
    ("12 December - Wedding\\Vidaai", "Babul Jo Tumne Sikhaya - Lata Mangeshkar", "ytsearch1:Babul Jo Tumne Sikhaya Hum Aapke Hain Koun Official Video"),

    # Bagheli Rewa Traditions
    ("Bagheli Wedding Songs\\01 - Mayan & Mandap Pujan", "Aaju Janakpur Marwa Suhaawan - Bagheli", "ytsearch1:Aaju Janakpur Marwa Suhaawan Lage Bagheli Vivah Geet"),
    ("Bagheli Wedding Songs\\02 - Haldi & Nahchhu", "Hardi Chadhawa Mor Bhaiya Ke - Bagheli", "ytsearch1:Hardi Chadhawa Mor Bhaiya Ke Bagheli Vivah Geet"),
    ("Bagheli Wedding Songs\\03 - Mehndi & Ladies Sangeet", "Kahe Par Rachahu Mehndiya - Bagheli", "ytsearch1:Kahe Par Rachahu Mehndiya Vandana Mishra Bagheli"),
    ("Bagheli Wedding Songs\\04 - Baraat & Ghodi Chadai", "Banna Chale Sasural Byahan - Bagheli", "ytsearch1:Banna Chale Sasural Byahan Bagheli Vivah Geet"),
]

def main():
    print(f"🚀 Batch 2: Downloading {len(curated_batch_2)} songs directly...")
    success = 0
    
    for folder, name, query in curated_batch_2:
        target_dir = os.path.join(BASE_DIR, folder)
        os.makedirs(target_dir, exist_ok=True)
        
        # Check if already downloaded
        exists = any(f.startswith(name) for f in os.listdir(target_dir)) if os.path.exists(target_dir) else False
        if exists:
            print(f"⏩ ALREADY EXISTS: {name}")
            success += 1
            continue
            
        print(f"⬇️ DOWNLOADING: {name}...")
        ydl_opts = {
            'format': 'bestaudio/best',
            'outtmpl': os.path.join(target_dir, f"{name}.%(ext)s"),
            'quiet': True,
            'no_warnings': True,
            'retries': 2,
            'default_search': 'ytsearch'
        }
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([query])
            print(f"✅ SAVED: {name}")
            success += 1
        except Exception as e:
            print(f"❌ FAILED: {name} -> {e}")

    print("\n" + "="*50)
    print(f"Batch 2 Complete: {success}/{len(curated_batch_2)} tracks saved!")
    print("="*50)

if __name__ == "__main__":
    main()

