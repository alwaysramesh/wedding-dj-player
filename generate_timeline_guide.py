import os
import openpyxl

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"

# Master Ceremony Timeline & Audio Management Protocol
timeline_content = """# 🎧 Master DJ Timeline, Audio Engineering & Family Protocol
**Brother's Grand Wedding | 10, 11 & 12 December 2026**
**Music Library Reference:** `Brother Wedding Music/`
**Master Excel Sheet:** `Brother_Wedding_Music_Collection.xlsx`

---

## 📌 1. Master Ceremony Schedule & DJ Cue Sheets

### 🟡 DAY 1 – 10 DECEMBER 2026: MAYAN & HALDI

#### A. Mayan / Tilak Ceremony (Morning: 09:30 AM – 11:30 AM)
- **Ceremony Focus:** Ganpati Sthapana, Kuldevi Pujan, Tilak Chadhaai, Banna blessings.
- **Sound Profile:** Acoustic, traditional, Shehnai & Sitar. Volume: 55–65 dB (Pundit ji's microphone priority).
- **Audio Cues:**
  1. *09:30 AM (Guests Arriving & Sthapana):* `Mangal Dhun (Shehnai)` - Ustad Bismillah Khan
  2. *09:50 AM (Ganesh Vandana Ritual):* `Gauri Ke Lala Ko Manao` / `Mangal Bhawan Amangal Haari`
  3. *10:20 AM (Tilak Ritual by Elders):* `Aaj Janakpur Mein Marwa` / `Chaap Tilak Sab Cheeni`
  4. *11:00 AM (Family Blessings & Aarti):* `Shri Ram Janki Vivah Mangalgeet`
  5. *11:15 AM (Pre-Haldi Joyous Transition):* `Banna Re Banna` / `Tujh Mein Rab Dikhta Hai`

#### B. Haldi Ceremony & Vatna (Morning/Afternoon: 11:30 AM – 02:30 PM)
- **Ceremony Focus:** Tel-Baan, Vatna application, phoolon ki haldi, joyful family dance.
- **Sound Profile:** Mid-bass, uplifting rhythm, Punjabi folk dholki & modern feel-good pop. Volume: 70–82 dB.
- **Audio Cues:**
  1. *11:30 AM (Groom Seat Arrival):* `Wah Wah Vattna Katore Da` / `Mathe Te Chamkan Waal`
  2. *11:45 AM (Elders Applying Haldi):* `Suhe Ve Cheere Waleya` / `Latthe Di Chadar` / `Sasural Genda Phool`
  3. *12:15 PM (Sisters & Bhabhis Haldi Application):* `Haldi Lagao Sajna Ko` / `Haldi Hai Aayi (2025)`
  4. *12:45 PM (Flower Shower / Phoolon Ki Haldi):* `Rangisari` / `Navrai Majhi`
  5. *01:15 PM (Dance Floor Opens for All):* `London Thumakda` / `Laal Peeli Akhiyaan` / `Nachde Ne Saare`
  6. *01:45 PM (High-Energy Family Bhangra):* `Gallan Goodiyaan` / `Gud Naal Ishq Mitha` / `What Jhumka?`

---

### 🟢 DAY 2 – 11 DECEMBER 2026: MEHNDI, LADIES SANGEET & DJ NIGHT

#### A. Traditional Mehndi & Chooda Ceremony (Afternoon: 02:00 PM – 05:30 PM)
- **Ceremony Focus:** Henna application, traditional dholki singing, chooda blessings.
- **Sound Profile:** Melodious acoustic, soothing strings, joyful female vocals. Volume: 60–70 dB.
- **Audio Cues:**
  1. *02:00 PM (Mehndi Artists Setup & Welcome):* `Mehndi Hai Rachne Wali` / `Maine Payal Hai Chhankai`
  2. *02:45 PM (Henna Application in Full Swing):* `Mehendi Rang Layee` / `Likh Ke Mehndi Se Sajna Ka Naam`
  3. *03:30 PM (Ladies Dholki & Folk Circle):* `Chitta Kukkad` / `Kala Sha Kala` / `Lo Chali Main`
  4. *04:15 PM (Sweet Romantic Melodies):* `Dekhha Tenu (2024)` / `Rangrez` / `Sweetheart`
  5. *05:00 PM (Sangeet Teaser / Ladies Casual Dance):* `Bole Chudiyan` / `Punjabi Wedding Song`

#### B. Grand Ladies Sangeet & Choreographed Performances (Evening: 07:00 PM – 10:00 PM)
- **Ceremony Focus:** Family stage performances, brother-sister dance, couple waltz, parent blessings.
- **Sound Profile:** High-definition live stage audio, punchy vocal clarity. Volume: 80–90 dB.
- **Performance Sequence Cues:**
  1. *07:00 PM (Grand Stage Introduction):* `Deewangi Deewangi` (Family group entry)
  2. *07:20 PM (Parents & Elders Special):* `Maahi Ve` / `Say Shava Shava`
  3. *07:45 PM (Brother & Groom Squad Dance):* `Veer Ji Viyahn Aaye` / `Ainvayi Ainvayi` / `Sadi Gali`
  4. *08:15 PM (Sisters & Bhabhis Performance):* `Mere Haathon Mein Nau Nau` / `Ghoomar` / `Didi Tera Devar`
  5. *08:45 PM (Brother & Bhabhi / Couple Special Waltz):*
     - First Dance: `Tum Kya Mile` or `Heeriye` (Slow romantic waltz)
     - Second Dance: `Sajni (Laapataa Ladies)` (Soulful duet)
     - Romantic Crescendo: `O Maahi` / `Ve Haaniyaan`
  6. *09:30 PM (Grand Finale Family Group):* `Gallan Goodiyaan` / `Sauda Khara Khara`

#### C. DJ Night & Punjabi Dance Explosion (Night: 10:00 PM – Late)
- **Sound Profile:** Club bass (35Hz–70Hz punch), subwoofers active, laser lighting sync. Volume: 92–102 dB.
- **Dance Floor Bangers:**
  - *Phase 1 (Punjabi Anthems):* `Born to Shine` (Diljit), `Morni Banke`, `Ishq Tera Tadpave`
  - *Phase 2 (Bollywood Dance Hits):* `Aayi Nai (Stree 2)`, `Kala Chashma`, `The Punjaabban Song`
  - *Phase 3 (Late Night Climax):* `Vibe`, `Badtameez Dil`, `Sauda Khara Khara`

---

### 🔴 DAY 3 – 12 DECEMBER 2026: THE GRAND WEDDING DAY

#### A. Sehra Bandi & Ghodi Chadhai (Evening: 05:00 PM – 06:30 PM)
- **Ceremony Focus:** Turban tying by family, sister tying Sehra, mounting the Ghodi.
- **Sound Profile:** Traditional brass band & auspicious royal vocals. Volume: 75–85 dB.
- **Audio Cues:**
  1. *05:00 PM (Sehra Bandi Ritual):* `Dulhe Ka Sehra Suhana Lagta Hai`
  2. *05:30 PM (Bapu / Chacha / Mama Turban Blessings):* `Banna Re Baga Me Jhula` / `Banna Re Banni Re`
  3. *06:00 PM (Groom Mounting the Ghodi):* `Ghodi Chadeya Bhena Da Veer` / `Veer Ji Viyahn Aaye`
  4. *06:15 PM (Brass Band Procession Launch):* `Aaj Mere Yaar Ki Shaadi Hai (Brass Band Version)`

#### B. The Grand Royal Baraat (Evening: 06:30 PM – 08:30 PM)
- **Ceremony Focus:** Dhol beats, mobile sound system, friends & family dancing to venue entrance.
- **Sound Profile:** Maximum punch, synced with live Punjabi Dholis. Volume: 95–105 dB.
- **Audio Cues:**
  1. *Baraat Launch:* `Aaj Mere Yaar Ki Shaadi Hai` (Mohammed Rafi)
  2. *Dhol Jugalbandi Drop:* `Tunak Tunak Tun` (Daler Mehndi) / `Ishq Tera Tadpave` (Sukhbir)
  3. *High Voltage Bollywood Baraat:* `Malhari` / `Tenu Leke` / `Mundian To Bach Ke`
  4. *2024–2026 Baraat Hits:* `Bhool Bhulaiyaa 3 Title Track` / `Pushpa Pushpa` / `Mere Mehboob Mere Sanam`
  5. *Arrival at Gate (Toran / Swagat):* `Saajanji Ghar Aaye` / `Nagada Sang Dhol`

#### C. Royal Varmala / Jaimala Ceremony (Night: 08:30 PM – 09:15 PM)
- **Ceremony Focus:** Grand stage reveal, floral shower, cold pyrotechnics, garland exchange.
- **Sound Profile:** Cinematic, grand symphonic, romantic crescendo. Volume: 80–88 dB.
- **Audio Cues:**
  1. *Bridal Walkway (Phoolon Ki Chaadar):*
     - Option A (Traditional & Soulful): `Din Shagna Da` (Jasleen Royal)
     - Option B (Modern Royal): `Kudmayi (Bridal Cut)` / `Madhanya`
     - Option C (Emotional): `O Piya` (Parineeti Chopra)
  2. *Stage Meeting Moment:* `Jashn-E-Bahara (Flute Instrumental)`
  3. *Garland Exchange (At the exact second of Varmala + Cold Pyros):*
     - **CUE:** `Kudmayi` (Drop at "Kudmayi hoyi...") OR `Mangalyam Tantunanena` (A.R. Rahman)
  4. *Celebratory Floral Shower:* `Dekha Tenu (2024)` / `Aaj Sajeya` / `Pehla Nasha (Strings)`

#### D. Sacred Pheras & Vedic Rituals (Late Evening / Night: 09:30 PM – 11:30 PM)
- **Ceremony Focus:** Havan, Kanyadaan, Gathbandhan, Saat Pheras, Sindoor Daan, Mangalsutra.
- **Sound Profile:** Ambient, divine Shehnai, Bansuri and Vedic chanting. Volume: 50–60 dB (Very soft background).
- **Audio Cues:**
  1. *Mandap Entry & Gathbandhan:* `Mangal Parinay (Raga Rageshri)` - Ustad Bismillah Khan
  2. *Kanyadaan Moment:* `Kanyadaan (Hamsa Narayan)` - Shehnai
  3. *Saat Pheras (Seven Vows):* `Saptha Padi Mantras` / `Raag Yaman (Sitar & Bansuri)`
  4. *Sindoor Daan & Mangalsutra:* `Sindoor Daan Chaiti` - Ustad Bismillah Khan
  5. *Ritual Completion & Family Blessings:* `Vivah Sampann Riti` / `Mangal Dhun Va Vivah Dhun`

#### E. Reception & Dinner (Night: 09:30 PM – Midnight)
- **Ceremony Focus:** Meeting guests, photography, dinner banquet, celebratory mood.
- **Audio Cues:**
  - Background Lounge: `Tum Kya Mile`, `Chaleya`, `Sajni`, `Kesariya`, `Ve Haaniyaan`

#### F. The Vidaai / Doli (Late Night / Early Morning)
- **Ceremony Focus:** Bride farewell, emotional family parting, father-daughter embrace.
- **Sound Profile:** Deeply emotional, heartfelt, touching yet dignified. Volume: 60–70 dB.
- **Audio Cues:**
  1. *Doli Preparation:* `Dilbaro` (Raazi)
  2. *Father & Mother Farewell Embrace:* `Babul Ki Duayen Leti Ja` (Mohammed Rafi) / `Babul Jo Tumne Sikhaya`
  3. *Car Departure:* `Madhanya` (Neha Bhasin) / `Sajni (Vidaai Acoustic)` / `Ghar (Jab Harry Met Sejal)`

---

## 📋 2. Sound Setup & Technical Equipment Checklist

| Zone / Ceremony | Recommended PA & Audio Setup | Microphones | Critical Notes |
|-----------------|------------------------------|-------------|----------------|
| **Mayan & Haldi (Outdoor / Courtyard)** | 2x 15" Powered Tops (JBL / RCF) + 1x 18" Sub | 2x Wireless Handhelds | Keep sub-bass moderate; vocal clarity essential for Pundit ji & singers. |
| **Sangeet Stage (Banquet / Lawn)** | Line Array (4 tops per side) + 4x Dual 18" Subs | 4x Wireless Lapels / Handhelds | Monitor wedges for dancers; zero delay in stage sound. |
| **Baraat Sound Truck** | Mobile Heavy DJ Rig (4x Dual 18" Bass bins + 4 Tops) | 1x Long-range Wireless | Coordinate sync between Dhol players and truck operator. |
| **Mandap / Pheras (Sacred Canopy)** | 2x Column PA Speakers (Bose L1 / JBL EON One) | 1x Lapel for Pandit + 1 Handheld | Crystal-clear acoustic projection with zero screech or feedback. |

---

## 🔒 3. Family Dignity Protocol for the DJ Crew

1. **Strictly Blacklisted Audio:**
   - No Bollywood item numbers with suggestive slang.
   - No songs promoting alcohol, weapons, or violent machismo inside ceremonial areas.
   - No insulting "samdhi / in-law" parody tracks that cause embarrassment.
2. **Ceremony Decorum:**
   - During Pheras and Kanyadaan, the DJ must lower all track levels whenever the Pandit chants mantras.
   - During Varmala, the cold pyro drop must be timed precisely with the musical crescendo.
"""

with open(os.path.join(BASE_DIR, "Master_Wedding_DJ_Timeline_Guide.md"), "w", encoding="utf-8") as f:
    f.write(timeline_content)

print("Generated Master_Wedding_DJ_Timeline_Guide.md successfully!")
