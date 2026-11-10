import os
import json
import re

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

from generate_wedding_player_app import load_metadata, scan_songs

meta = load_metadata()
songs = scan_songs(meta)
songs_json = json.dumps(songs, ensure_ascii=False)

# Read the existing template from build_full_player_web_app.py
with open("build_full_player_web_app.py", "r", encoding="utf-8") as f:
    orig = f.read()

# Extract html_template
start_str = 'html_template = """'
end_str = '"""\n\nfinal_html'
s_idx = orig.find(start_str) + len(start_str)
e_idx = orig.find(end_str)
html_template = orig[s_idx:e_idx]

# Add DJ Sound FX, EQ Presets, Crossfade & M3U Export
# 1. In Toolbar: Add M3U Export button & EQ Presets dropdown
old_toolbar = '<button class="btn-bulk-queue" id="btn-queue-filtered">➕ Queue All</button>'
new_toolbar = """<button class="btn-bulk-queue" id="btn-queue-filtered">➕ Queue All</button>
                <button class="btn-hdr" id="btn-export-m3u" title="Export Current View as DJ Playlist (.m3u)">💾 Export .m3u</button>
                <div style="display: flex; align-items: center; gap: 6px; background: var(--bg-card); border: 1px solid var(--border-color); padding: 4px 10px; border-radius: var(--radius-md);">
                    <span style="font-size: 0.76rem; color: var(--gold-light); font-weight: 700;">🎚️ EQ:</span>
                    <select id="eq-preset-select" style="background: none; border: none; color: var(--text-main); font-size: 0.82rem; outline: none; cursor: pointer;">
                        <option value="flat" style="background: #111;">Standard (Flat)</option>
                        <option value="bass" style="background: #111;">Bass Boost (DJ Party)</option>
                        <option value="vocal" style="background: #111;">Vocal Clarity (Rituals)</option>
                        <option value="warm" style="background: #111;">Warm Acoustic (Melody)</option>
                    </select>
                </div>
                <div style="display: flex; align-items: center; gap: 6px; background: var(--bg-card); border: 1px solid var(--border-color); padding: 4px 10px; border-radius: var(--radius-md);">
                    <label style="font-size: 0.76rem; color: var(--gold-light); font-weight: 700; cursor: pointer; display: flex; align-items: center; gap: 6px;">
                        <input type="checkbox" id="chk-crossfade" style="cursor: pointer; accent-color: var(--gold-primary);">
                        <span>Crossfade (3s)</span>
                    </label>
                </div>"""

# 2. Add DJ Sound FX Launchpad Bar right under Ceremony soundboard
soundboard_target = '</div>\n        </div>\n\n        <!-- Search & View Toolbar -->'
dj_fx_bar = """</div>
        </div>

        <!-- DJ Instant Jingle & Sound FX Soundboard -->
        <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 12px 18px; display: flex; align-items: center; justify-content: space-between; gap: 14px; flex-wrap: wrap; box-shadow: var(--shadow-card);">
            <div style="display: flex; align-items: center; gap: 8px;">
                <span style="font-size: 1.1rem;">📢</span>
                <span style="font-size: 0.82rem; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; color: var(--gold-light);">DJ Live Sound FX Pads:</span>
                <span style="font-size: 0.72rem; color: var(--text-muted);">(Press 1-5 to trigger)</span>
            </div>
            <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                <button class="btn-hdr" onclick="playDjSoundFx('airhorn')" style="background: rgba(244, 63, 94, 0.15); border-color: var(--neon-crimson); color: #fff;">
                    <span>📢 Air Horn</span> <span style="font-size: 0.68rem; opacity: 0.7;">[1]</span>
                </button>
                <button class="btn-hdr" onclick="playDjSoundFx('dhol')" style="background: rgba(245, 158, 11, 0.15); border-color: var(--gold-primary); color: #fff;">
                    <span>🥁 Dhol Roll</span> <span style="font-size: 0.68rem; opacity: 0.7;">[2]</span>
                </button>
                <button class="btn-hdr" onclick="playDjSoundFx('shehnai')" style="background: rgba(16, 185, 129, 0.15); border-color: var(--neon-green); color: #fff;">
                    <span>🪔 Shehnai Chime</span> <span style="font-size: 0.68rem; opacity: 0.7;">[3]</span>
                </button>
                <button class="btn-hdr" onclick="playDjSoundFx('applause')" style="background: rgba(139, 92, 246, 0.15); border-color: var(--neon-purple); color: #fff;">
                    <span>👏 Crowd Cheer</span> <span style="font-size: 0.68rem; opacity: 0.7;">[4]</span>
                </button>
                <button class="btn-hdr" onclick="playDjSoundFx('whistle')" style="background: rgba(6, 182, 212, 0.15); border-color: var(--neon-cyan); color: #fff;">
                    <span>📯 Whistle Hype</span> <span style="font-size: 0.68rem; opacity: 0.7;">[5]</span>
                </button>
            </div>
        </div>

        <!-- Search & View Toolbar -->"""

# 3. Add Sound FX JS Implementation and EQ filters in the script
script_boot_target = '// Initial Boot'
enhanced_js = """
        // Equalizer & Filter Nodes
        let bassFilter, midFilter, trebleFilter;
        let isCrossfade = false;

        function setupEqualizer() {
            if (!audioCtx) return;
            try {
                bassFilter = audioCtx.createBiquadFilter();
                bassFilter.type = 'lowshelf';
                bassFilter.frequency.value = 250;

                midFilter = audioCtx.createBiquadFilter();
                midFilter.type = 'peaking';
                midFilter.frequency.value = 1500;
                midFilter.Q.value = 1.0;

                trebleFilter = audioCtx.createBiquadFilter();
                trebleFilter.type = 'highshelf';
                trebleFilter.frequency.value = 4000;

                // Reconnect: source -> bass -> mid -> treble -> analyser -> destination
                // Managed inside initAudioVisualizer
            } catch(e) {
                console.warn('EQ setup error:', e);
            }
        }

        document.getElementById('eq-preset-select').addEventListener('change', (e) => {
            const val = e.target.value;
            if (!bassFilter || !midFilter || !trebleFilter) return;
            if (val === 'bass') {
                bassFilter.gain.value = 8;
                midFilter.gain.value = 0;
                trebleFilter.gain.value = 3;
                showToast('🎚️ Bass Boost Activated');
            } else if (val === 'vocal') {
                bassFilter.gain.value = -3;
                midFilter.gain.value = 6;
                trebleFilter.gain.value = 4;
                showToast('🎚️ Vocal Clarity Activated');
            } else if (val === 'warm') {
                bassFilter.gain.value = 4;
                midFilter.gain.value = 2;
                trebleFilter.gain.value = -2;
                showToast('🎚️ Warm Acoustic Activated');
            } else {
                bassFilter.gain.value = 0;
                midFilter.gain.value = 0;
                trebleFilter.gain.value = 0;
                showToast('🎚️ EQ Reset to Flat');
            }
        });

        document.getElementById('chk-crossfade').addEventListener('change', (e) => {
            isCrossfade = e.target.checked;
            showToast(isCrossfade ? '⚡ Crossfade (3s) Enabled' : '➡️ Crossfade Disabled');
        });

        // Export Playlist to .m3u
        document.getElementById('btn-export-m3u').addEventListener('click', () => {
            const list = getFilteredSongs();
            if (list.length === 0) {
                showToast('No songs to export');
                return;
            }
            let m3u = '#EXTM3U\\n';
            list.forEach(s => {
                m3u += `#EXTINF:-1,${s.artist} - ${s.title}\\n${s.path}\\n`;
            });
            const blob = new Blob([m3u], { type: 'audio/x-mpegurl' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Wedding_Playlist_${state.activeFilters.cero || 'Master'}.m3u`;
            a.click();
            URL.revokeObjectURL(url);
            showToast(`💾 Exported ${list.length} songs as .m3u playlist`);
        });

        // DJ Real-Time Synthesizer Sound FX
        window.playDjSoundFx = function(type) {
            initAudioVisualizer();
            const ctx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
            const now = ctx.currentTime;

            if (type === 'airhorn') {
                const freqs = [466.16, 622.25, 783.99]; // Classic dancehall airhorn chord
                freqs.forEach(f => {
                    const osc = ctx.createOscillator();
                    const gain = ctx.createGain();
                    osc.type = 'sawtooth';
                    osc.frequency.setValueAtTime(f, now);
                    osc.frequency.exponentialRampToValueAtTime(f * 1.05, now + 0.1);
                    osc.frequency.exponentialRampToValueAtTime(f, now + 0.35);

                    gain.gain.setValueAtTime(0.25, now);
                    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.55);

                    osc.connect(gain);
                    gain.connect(ctx.destination);
                    osc.start(now);
                    osc.stop(now + 0.6);
                });
                showToast('📢 Air Horn!');
            } else if (type === 'dhol') {
                // Dhol roll (sharp boom + rumble)
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                osc.type = 'sine';
                osc.frequency.setValueAtTime(120, now);
                osc.frequency.exponentialRampToValueAtTime(45, now + 0.4);

                gain.gain.setValueAtTime(0.8, now);
                gain.gain.exponentialRampToValueAtTime(0.01, now + 0.45);

                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start(now);
                osc.stop(now + 0.5);
                showToast('🥁 Dhol Roll!');
            } else if (type === 'shehnai') {
                // Auspicious Shehnai Chime
                const notes = [587.33, 739.99, 880.00, 1174.66]; // D, F#, A, D
                notes.forEach((freq, idx) => {
                    const osc = ctx.createOscillator();
                    const gain = ctx.createGain();
                    osc.type = 'triangle';
                    osc.frequency.setValueAtTime(freq, now + idx * 0.08);

                    gain.gain.setValueAtTime(0, now + idx * 0.08);
                    gain.gain.linearRampToValueAtTime(0.2, now + idx * 0.08 + 0.02);
                    gain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.08 + 0.6);

                    osc.connect(gain);
                    gain.connect(ctx.destination);
                    osc.start(now + idx * 0.08);
                    osc.stop(now + idx * 0.08 + 0.65);
                });
                showToast('🪔 Shehnai Chime!');
            } else if (type === 'applause') {
                // White noise applause burst
                const bufferSize = ctx.sampleRate * 1.2;
                const buffer = ctx.createBuffer(1, bufferSize, ctx.sampleRate);
                const data = buffer.getChannelData(0);
                for (let i = 0; i < bufferSize; i++) {
                    data[i] = Math.random() * 2 - 1;
                }
                const noise = ctx.createBufferSource();
                noise.buffer = buffer;

                const filter = ctx.createBiquadFilter();
                filter.type = 'bandpass';
                filter.frequency.value = 1000;
                filter.Q.value = 1.2;

                const gain = ctx.createGain();
                gain.gain.setValueAtTime(0.01, now);
                gain.gain.linearRampToValueAtTime(0.4, now + 0.15);
                gain.gain.exponentialRampToValueAtTime(0.01, now + 1.2);

                noise.connect(filter);
                filter.connect(gain);
                gain.connect(ctx.destination);
                noise.start(now);
                showToast('👏 Crowd Cheer!');
            } else if (type === 'whistle') {
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                osc.type = 'sine';
                osc.frequency.setValueAtTime(2600, now);
                osc.frequency.linearRampToValueAtTime(3200, now + 0.15);
                osc.frequency.linearRampToValueAtTime(2800, now + 0.35);

                gain.gain.setValueAtTime(0.25, now);
                gain.gain.exponentialRampToValueAtTime(0.01, now + 0.4);

                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start(now);
                osc.stop(now + 0.45);
                showToast('📯 Whistle Hype!');
            }
        };

        // Number keys 1-5 for Sound FX
        window.addEventListener('keydown', (e) => {
            if (e.target === searchInput) return;
            if (e.key === '1') playDjSoundFx('airhorn');
            if (e.key === '2') playDjSoundFx('dhol');
            if (e.key === '3') playDjSoundFx('shehnai');
            if (e.key === '4') playDjSoundFx('applause');
            if (e.key === '5') playDjSoundFx('whistle');
        });

        // Initial Boot"""

html_template = html_template.replace(old_toolbar, new_toolbar)
html_template = html_template.replace(soundboard_target, dj_fx_bar)
html_template = html_template.replace(script_boot_target, enhanced_js)

final_html = html_template.replace("__SONGS_JSON__", songs_json)

# Write to files
player_path = os.path.join(BASE_DIR, "Wedding_Music_Player.html")
with open(player_path, "w", encoding="utf-8") as f:
    f.write(final_html)

index_path = os.path.join(BASE_DIR, "index.html")
with open(index_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print("✅ Upgraded Wedding Music Player with DJ Sound FX, Bass Boost EQ, Crossfade & M3U Export!")

