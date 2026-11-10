import os
import json
import re
import urllib.parse
import openpyxl

BASE_DIR = r"c:\Users\LOTUS IT SOLUTION\Downloads\Ramesh_S\Brother Wedding Music"
EXCEL_FILE = os.path.join(BASE_DIR, "Brother_Wedding_Music_Collection.xlsx")

from generate_wedding_player_app import load_metadata, scan_songs

meta = load_metadata()
songs = scan_songs(meta)

songs_json = json.dumps(songs, ensure_ascii=False)

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Wedding Music Console • Ramesh Brother's Wedding (10-12 Dec 2026)</title>
    <style>
        :root {
            --bg-main: #0a0c14;
            --bg-card: #131722;
            --bg-card-hover: #1c2233;
            --bg-player: rgba(15, 18, 28, 0.94);
            --gold-primary: #f59e0b;
            --gold-light: #fbbf24;
            --gold-dark: #b45309;
            --gold-glow: rgba(245, 158, 11, 0.35);
            --neon-crimson: #f43f5e;
            --neon-cyan: #06b6d4;
            --neon-green: #10b981;
            --neon-purple: #8b5cf6;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border-color: rgba(245, 158, 11, 0.16);
            --border-hover: rgba(245, 158, 11, 0.45);
            --radius-md: 10px;
            --radius-lg: 16px;
            --radius-full: 9999px;
            --shadow-card: 0 8px 24px rgba(0, 0, 0, 0.45);
            --shadow-glow: 0 0 25px var(--gold-glow);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            user-select: none;
            -webkit-user-select: none;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-main);
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            overflow-x: hidden;
            background-image: 
                radial-gradient(circle at 10% 15%, rgba(245, 158, 11, 0.08) 0%, transparent 40%),
                radial-gradient(circle at 90% 85%, rgba(139, 92, 246, 0.06) 0%, transparent 45%);
            padding-bottom: 120px;
        }

        /* Scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #0d1017;
        }
        ::-webkit-scrollbar-thumb {
            background: #2a334a;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: var(--gold-primary);
        }

        /* Header */
        header {
            position: sticky;
            top: 0;
            z-index: 100;
            background: rgba(10, 12, 20, 0.92);
            backdrop-filter: blur(14px);
            border-bottom: 1px solid var(--border-color);
            padding: 12px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
        }

        .brand-container {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .brand-logo {
            width: 44px;
            height: 44px;
            border-radius: 12px;
            background: linear-gradient(135deg, #b45309, #f59e0b, #fbbf24);
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 0 18px var(--gold-glow);
            font-size: 22px;
            color: #111;
        }

        .brand-text h1 {
            font-size: 1.15rem;
            font-weight: 800;
            letter-spacing: 0.5px;
            background: linear-gradient(90deg, #fff, #fbbf24);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .brand-text p {
            font-size: 0.78rem;
            color: var(--text-muted);
            letter-spacing: 0.2px;
        }

        .header-stats-pill {
            background: rgba(245, 158, 11, 0.1);
            border: 1px solid rgba(245, 158, 11, 0.3);
            color: var(--gold-light);
            padding: 6px 14px;
            border-radius: var(--radius-full);
            font-size: 0.8rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .header-stats-pill .live-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: var(--neon-green);
            box-shadow: 0 0 8px var(--neon-green);
            animation: pulse-dot 2s infinite;
        }

        @keyframes pulse-dot {
            0%, 100% { transform: scale(1); opacity: 1; }
            50% { transform: scale(1.3); opacity: 0.5; }
        }

        .header-actions {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .btn-hdr {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 8px 14px;
            border-radius: var(--radius-md);
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
        }

        .btn-hdr:hover {
            background: var(--bg-card-hover);
            border-color: var(--gold-primary);
            color: var(--gold-light);
            box-shadow: 0 0 12px rgba(245, 158, 11, 0.2);
        }

        .btn-hdr.active {
            background: var(--gold-primary);
            color: #0b0d14;
            border-color: var(--gold-primary);
            box-shadow: 0 0 15px var(--gold-glow);
        }

        /* Main Container */
        .app-container {
            max-width: 1600px;
            margin: 0 auto;
            padding: 20px 24px;
            width: 100%;
            display: flex;
            flex-direction: column;
            gap: 20px;
        }

        /* Navigation Filters */
        .filter-section {
            display: flex;
            flex-direction: column;
            gap: 14px;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            padding: 18px 20px;
            box-shadow: var(--shadow-card);
        }

        .filter-row {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 8px;
        }

        .filter-label {
            font-size: 0.76rem;
            text-transform: uppercase;
            font-weight: 800;
            letter-spacing: 1px;
            color: var(--gold-light);
            width: 110px;
            flex-shrink: 0;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .filter-pill {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid rgba(255, 255, 255, 0.1);
            color: var(--text-muted);
            padding: 6px 14px;
            border-radius: var(--radius-full);
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .filter-pill:hover {
            background: rgba(255, 255, 255, 0.08);
            color: var(--text-main);
            border-color: rgba(245, 158, 11, 0.3);
        }

        .filter-pill.active {
            background: linear-gradient(135deg, var(--gold-dark), var(--gold-primary));
            border-color: var(--gold-light);
            color: #0b0d14;
            font-weight: 700;
            box-shadow: 0 0 14px var(--gold-glow);
        }

        .filter-pill .count {
            font-size: 0.72rem;
            background: rgba(0, 0, 0, 0.25);
            padding: 2px 6px;
            border-radius: 10px;
        }

        .filter-pill.active .count {
            background: rgba(0, 0, 0, 0.4);
            color: #fff;
        }

        /* Ceremony Quick Launch Soundboard */
        .soundboard-section {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .section-title {
            font-size: 0.88rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: var(--gold-light);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .soundboard-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
            gap: 10px;
        }

        .ceremony-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 12px 14px;
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            flex-direction: column;
            gap: 6px;
            position: relative;
            overflow: hidden;
        }

        .ceremony-card:hover {
            background: var(--bg-card-hover);
            border-color: var(--gold-primary);
            transform: translateY(-2px);
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.4);
        }

        .ceremony-card.active {
            border-color: var(--gold-light);
            background: linear-gradient(145deg, #1b233a, #151a2a);
            box-shadow: 0 0 16px var(--gold-glow);
        }

        .ceremony-card.active::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
            background: var(--gold-primary);
        }

        .cero-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .cero-icon {
            font-size: 1.3rem;
        }

        .cero-count {
            font-size: 0.72rem;
            color: var(--gold-light);
            background: rgba(245, 158, 11, 0.12);
            padding: 2px 7px;
            border-radius: var(--radius-full);
            font-weight: 700;
        }

        .cero-name {
            font-size: 0.86rem;
            font-weight: 700;
            color: var(--text-main);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .cero-play-all {
            margin-top: 4px;
            background: rgba(245, 158, 11, 0.1);
            border: 1px solid rgba(245, 158, 11, 0.2);
            color: var(--gold-light);
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 0.72rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 4px;
            transition: all 0.15s;
        }

        .cero-play-all:hover {
            background: var(--gold-primary);
            color: #0b0d14;
        }

        /* Search & Toolbar */
        .toolbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            flex-wrap: wrap;
        }

        .search-box {
            flex: 1;
            min-width: 280px;
            position: relative;
        }

        .search-input {
            width: 100%;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 12px 16px 12px 42px;
            border-radius: var(--radius-md);
            font-size: 0.92rem;
            outline: none;
            transition: all 0.2s;
        }

        .search-input:focus {
            border-color: var(--gold-primary);
            box-shadow: 0 0 16px var(--gold-glow);
            background: #161c2a;
        }

        .search-icon {
            position: absolute;
            left: 14px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            pointer-events: none;
        }

        .search-clear {
            position: absolute;
            right: 12px;
            top: 50%;
            transform: translateY(-50%);
            background: none;
            border: none;
            color: var(--text-muted);
            cursor: pointer;
            font-size: 16px;
            display: none;
        }

        .search-clear:hover {
            color: #fff;
        }

        .view-controls {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .btn-view {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            padding: 8px 12px;
            border-radius: var(--radius-md);
            font-size: 0.85rem;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .btn-view.active {
            background: var(--gold-primary);
            color: #0b0d14;
            border-color: var(--gold-primary);
            font-weight: 700;
        }

        .btn-bulk-play {
            background: linear-gradient(135deg, var(--gold-dark), var(--gold-primary));
            border: 1px solid var(--gold-light);
            color: #0b0d14;
            padding: 10px 18px;
            border-radius: var(--radius-md);
            font-size: 0.88rem;
            font-weight: 800;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 0 16px var(--gold-glow);
            transition: all 0.2s;
        }

        .btn-bulk-play:hover {
            transform: scale(1.02);
            box-shadow: 0 0 24px var(--gold-glow);
        }

        .btn-bulk-queue {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 10px 14px;
            border-radius: var(--radius-md);
            font-size: 0.88rem;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s;
        }

        .btn-bulk-queue:hover {
            border-color: var(--gold-primary);
            color: var(--gold-light);
        }

        /* Song List / Table / Grid Views */
        .song-container {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }

        /* Table View */
        .song-table {
            width: 100%;
            border-collapse: collapse;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            overflow: hidden;
            box-shadow: var(--shadow-card);
        }

        .song-table th {
            text-align: left;
            padding: 12px 16px;
            font-size: 0.74rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: var(--text-muted);
            border-bottom: 1px solid rgba(255, 255, 255, 0.08);
            background: rgba(0, 0, 0, 0.2);
        }

        .song-row {
            border-bottom: 1px solid rgba(255, 255, 255, 0.04);
            transition: all 0.15s ease;
            cursor: pointer;
        }

        .song-row:hover {
            background: var(--bg-card-hover);
        }

        .song-row.playing {
            background: rgba(245, 158, 11, 0.08);
            border-left: 4px solid var(--gold-primary);
        }

        .song-row td {
            padding: 12px 16px;
            font-size: 0.88rem;
            vertical-align: middle;
        }

        .song-row .col-idx {
            width: 50px;
            text-align: center;
            color: var(--text-muted);
            font-weight: 700;
            font-size: 0.82rem;
        }

        .song-row .col-title {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .song-title-text {
            font-weight: 700;
            color: var(--text-main);
            font-size: 0.94rem;
        }

        .song-row.playing .song-title-text {
            color: var(--gold-light);
        }

        .song-artist-text {
            font-size: 0.78rem;
            color: var(--text-muted);
        }

        .song-tag {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 0.72rem;
            font-weight: 700;
            background: rgba(255, 255, 255, 0.06);
            color: var(--text-muted);
            border: 1px solid rgba(255, 255, 255, 0.08);
        }

        .tag-day {
            background: rgba(245, 158, 11, 0.1);
            color: var(--gold-light);
            border-color: rgba(245, 158, 11, 0.25);
        }

        .tag-lang {
            background: rgba(6, 182, 212, 0.1);
            color: var(--neon-cyan);
            border-color: rgba(6, 182, 212, 0.25);
        }

        .tag-cero {
            background: rgba(244, 63, 94, 0.1);
            color: var(--neon-crimson);
            border-color: rgba(244, 63, 94, 0.25);
        }

        .row-actions {
            display: flex;
            align-items: center;
            gap: 8px;
            justify-content: flex-end;
        }

        .btn-act {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid rgba(255, 255, 255, 0.08);
            color: var(--text-muted);
            width: 32px;
            height: 32px;
            border-radius: var(--radius-md);
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.15s;
        }

        .btn-act:hover {
            background: var(--bg-card-hover);
            color: var(--text-main);
            border-color: var(--gold-primary);
        }

        .btn-act.fav-active {
            color: var(--neon-crimson);
            border-color: var(--neon-crimson);
            background: rgba(244, 63, 94, 0.15);
        }

        /* Grid Cards View */
        .song-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
            gap: 14px;
        }

        .song-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 14px 16px;
            display: flex;
            flex-direction: column;
            gap: 10px;
            transition: all 0.2s ease;
            cursor: pointer;
            position: relative;
        }

        .song-card:hover {
            background: var(--bg-card-hover);
            border-color: var(--gold-primary);
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.5);
        }

        .song-card.playing {
            border-color: var(--gold-light);
            background: linear-gradient(145deg, #1b233a, #131722);
            box-shadow: 0 0 18px var(--gold-glow);
        }

        .card-top {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .card-play-btn {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            background: var(--gold-primary);
            color: #0b0d14;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            flex-shrink: 0;
            box-shadow: 0 0 12px var(--gold-glow);
            transition: all 0.2s;
        }

        .song-card:hover .card-play-btn {
            transform: scale(1.08);
            background: var(--gold-light);
        }

        .card-info {
            flex: 1;
            overflow: hidden;
        }

        .card-title {
            font-size: 0.94rem;
            font-weight: 700;
            color: var(--text-main);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .song-card.playing .card-title {
            color: var(--gold-light);
        }

        .card-artist {
            font-size: 0.78rem;
            color: var(--text-muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            margin-top: 2px;
        }

        .card-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            align-items: center;
        }

        .card-bottom {
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            padding-top: 8px;
            font-size: 0.74rem;
            color: var(--text-muted);
        }

        /* Folder Tree View */
        .folder-tree {
            display: flex;
            flex-direction: column;
            gap: 14px;
        }

        .folder-group {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            overflow: hidden;
            box-shadow: var(--shadow-card);
        }

        .folder-header {
            padding: 14px 18px;
            background: rgba(0, 0, 0, 0.25);
            display: flex;
            align-items: center;
            justify-content: space-between;
            cursor: pointer;
            transition: all 0.2s;
            border-bottom: 1px solid rgba(255, 255, 255, 0.06);
        }

        .folder-header:hover {
            background: rgba(245, 158, 11, 0.08);
        }

        .folder-title-left {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .folder-icon {
            font-size: 1.3rem;
            color: var(--gold-light);
        }

        .folder-name {
            font-size: 0.98rem;
            font-weight: 800;
            color: var(--text-main);
        }

        .folder-stats {
            font-size: 0.78rem;
            color: var(--text-muted);
        }

        .folder-actions {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .subfolder-container {
            padding: 10px 18px 16px;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .subfolder-block {
            background: rgba(0, 0, 0, 0.18);
            border: 1px solid rgba(255, 255, 255, 0.05);
            border-radius: var(--radius-md);
            padding: 10px 14px;
        }

        .subfolder-head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 8px;
        }

        .subfolder-title {
            font-size: 0.88rem;
            font-weight: 700;
            color: var(--gold-light);
            display: flex;
            align-items: center;
            gap: 6px;
        }

        /* Persistent Bottom DJ Player */
        .bottom-player {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            height: 100px;
            background: var(--bg-player);
            backdrop-filter: blur(20px);
            border-top: 1px solid var(--border-hover);
            box-shadow: 0 -8px 30px rgba(0, 0, 0, 0.7);
            z-index: 1000;
            display: flex;
            align-items: center;
            padding: 0 24px;
            gap: 20px;
        }

        /* Player Left: Current Track */
        .player-track-info {
            width: 320px;
            display: flex;
            align-items: center;
            gap: 14px;
            flex-shrink: 0;
        }

        .vinyl-container {
            width: 60px;
            height: 60px;
            border-radius: 50%;
            background: radial-gradient(circle, #333 20%, #111 60%, #444 65%, #111 90%);
            border: 2px solid var(--gold-primary);
            box-shadow: 0 0 14px var(--gold-glow);
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            flex-shrink: 0;
            animation: spin-vinyl 4s linear infinite;
            animation-play-state: paused;
        }

        .vinyl-container.spinning {
            animation-play-state: running;
        }

        .vinyl-center {
            width: 20px;
            height: 20px;
            border-radius: 50%;
            background: var(--gold-light);
            border: 2px solid #111;
        }

        @keyframes spin-vinyl {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }

        .track-text {
            overflow: hidden;
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .track-title {
            font-size: 0.96rem;
            font-weight: 800;
            color: var(--text-main);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .track-artist {
            font-size: 0.78rem;
            color: var(--gold-light);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .track-meta-pill {
            font-size: 0.68rem;
            color: var(--text-muted);
            margin-top: 2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        /* Player Center: Master Controls & Scrubber */
        .player-center {
            flex: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 6px;
            max-width: 720px;
            margin: 0 auto;
        }

        .transport-controls {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .btn-ctrl {
            background: none;
            border: none;
            color: var(--text-muted);
            cursor: pointer;
            font-size: 18px;
            transition: all 0.15s;
            display: flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            position: relative;
        }

        .btn-ctrl:hover {
            color: var(--text-main);
            background: rgba(255, 255, 255, 0.08);
        }

        .btn-ctrl.active {
            color: var(--gold-light);
        }

        .btn-ctrl.active::after {
            content: '';
            position: absolute;
            bottom: 4px;
            width: 4px;
            height: 4px;
            border-radius: 50%;
            background: var(--gold-primary);
        }

        .btn-play-master {
            width: 52px;
            height: 52px;
            border-radius: 50%;
            background: linear-gradient(135deg, var(--gold-primary), var(--gold-light));
            border: none;
            color: #0b0d14;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 0 22px var(--gold-glow);
            transition: all 0.2s;
            font-size: 22px;
        }

        .btn-play-master:hover {
            transform: scale(1.08);
            box-shadow: 0 0 30px var(--gold-glow);
        }

        .scrubber-container {
            width: 100%;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .time-text {
            font-size: 0.74rem;
            color: var(--text-muted);
            font-family: monospace;
            width: 42px;
            text-align: center;
        }

        .seek-bar-track {
            flex: 1;
            height: 6px;
            background: rgba(255, 255, 255, 0.15);
            border-radius: 3px;
            cursor: pointer;
            position: relative;
            overflow: hidden;
            transition: height 0.15s;
        }

        .seek-bar-track:hover {
            height: 8px;
        }

        .seek-progress {
            height: 100%;
            background: linear-gradient(90deg, var(--gold-primary), var(--gold-light));
            border-radius: 3px;
            width: 0%;
            position: relative;
        }

        /* Player Right: Volume & Tools */
        .player-tools {
            width: 320px;
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 12px;
            flex-shrink: 0;
        }

        .volume-container {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .vol-slider {
            width: 90px;
            -webkit-appearance: none;
            appearance: none;
            height: 5px;
            border-radius: 3px;
            background: rgba(255, 255, 255, 0.18);
            outline: none;
            cursor: pointer;
        }

        .vol-slider::-webkit-slider-thumb {
            -webkit-appearance: none;
            appearance: none;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: var(--gold-light);
            cursor: pointer;
            box-shadow: 0 0 8px var(--gold-glow);
        }

        /* Audio Visualizer Canvas */
        .visualizer-canvas {
            width: 80px;
            height: 36px;
            background: rgba(0, 0, 0, 0.2);
            border-radius: 6px;
            border: 1px solid rgba(245, 158, 11, 0.2);
        }

        /* Slide-out Queue Panel */
        .queue-panel {
            position: fixed;
            top: 0;
            right: -420px;
            width: 400px;
            height: calc(100vh - 100px);
            background: rgba(15, 18, 28, 0.98);
            backdrop-filter: blur(24px);
            border-left: 1px solid var(--border-color);
            box-shadow: -8px 0 30px rgba(0, 0, 0, 0.8);
            z-index: 900;
            transition: right 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            display: flex;
            flex-direction: column;
        }

        .queue-panel.open {
            right: 0;
        }

        .queue-header {
            padding: 16px 20px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .queue-title {
            font-size: 1rem;
            font-weight: 800;
            color: var(--gold-light);
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .queue-list {
            flex: 1;
            overflow-y: auto;
            padding: 10px 14px;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        .queue-item {
            background: var(--bg-card);
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: var(--radius-md);
            padding: 10px 12px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            transition: all 0.15s;
        }

        .queue-item:hover {
            background: var(--bg-card-hover);
            border-color: var(--gold-primary);
        }

        .queue-item.current {
            border-color: var(--gold-light);
            background: rgba(245, 158, 11, 0.12);
        }

        .queue-item-info {
            overflow: hidden;
            flex: 1;
        }

        .queue-item-title {
            font-size: 0.86rem;
            font-weight: 700;
            color: var(--text-main);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .queue-item-artist {
            font-size: 0.74rem;
            color: var(--text-muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        /* Fullscreen DJ Stage Mode Overlay */
        .dj-stage-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: #07090e;
            z-index: 2000;
            display: none;
            flex-direction: column;
            justify-content: space-between;
            padding: 30px 50px;
            background-image: 
                radial-gradient(circle at 50% 20%, rgba(245, 158, 11, 0.14) 0%, transparent 60%),
                radial-gradient(circle at 50% 90%, rgba(244, 63, 94, 0.1) 0%, transparent 70%);
        }

        .dj-stage-overlay.active {
            display: flex;
        }

        .stage-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .stage-brand {
            font-size: 1.2rem;
            font-weight: 800;
            color: var(--gold-light);
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .stage-center {
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            gap: 16px;
            max-width: 900px;
            margin: 0 auto;
        }

        .stage-vinyl {
            width: 220px;
            height: 220px;
            border-radius: 50%;
            background: radial-gradient(circle, #333 20%, #111 60%, #444 65%, #111 90%);
            border: 4px solid var(--gold-primary);
            box-shadow: 0 0 50px var(--gold-glow);
            display: flex;
            align-items: center;
            justify-content: center;
            animation: spin-vinyl 4s linear infinite;
        }

        .stage-vinyl-center {
            width: 70px;
            height: 70px;
            border-radius: 50%;
            background: var(--gold-light);
            border: 4px solid #111;
        }

        .stage-song-title {
            font-size: 2.2rem;
            font-weight: 900;
            background: linear-gradient(90deg, #ffffff, var(--gold-light));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            line-height: 1.2;
        }

        .stage-song-artist {
            font-size: 1.3rem;
            color: var(--gold-primary);
            font-weight: 600;
        }

        .stage-badges {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .stage-controls {
            display: flex;
            align-items: center;
            gap: 28px;
            margin-top: 10px;
        }

        .stage-btn-play {
            width: 80px;
            height: 80px;
            border-radius: 50%;
            background: linear-gradient(135deg, var(--gold-primary), var(--gold-light));
            border: none;
            color: #0b0d14;
            font-size: 32px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 0 35px var(--gold-glow);
            transition: all 0.2s;
        }

        .stage-btn-play:hover {
            transform: scale(1.1);
        }

        .stage-canvas {
            width: 100%;
            max-width: 800px;
            height: 60px;
            margin-top: 10px;
        }

        /* Keyboard Help Modal */
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: rgba(0, 0, 0, 0.75);
            backdrop-filter: blur(8px);
            z-index: 3000;
            display: none;
            align-items: center;
            justify-content: center;
        }

        .modal-overlay.open {
            display: flex;
        }

        .modal-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            width: 90%;
            max-width: 560px;
            box-shadow: var(--shadow-card);
            overflow: hidden;
        }

        .modal-head {
            padding: 16px 22px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .modal-body {
            padding: 22px;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .shortcut-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 0.9rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            padding-bottom: 8px;
        }

        .key-badge {
            background: rgba(245, 158, 11, 0.15);
            border: 1px solid var(--gold-primary);
            color: var(--gold-light);
            padding: 3px 10px;
            border-radius: 6px;
            font-family: monospace;
            font-weight: 700;
        }

        /* Toast Notifications */
        .toast-container {
            position: fixed;
            top: 80px;
            right: 24px;
            z-index: 4000;
            display: flex;
            flex-direction: column;
            gap: 8px;
            pointer-events: none;
        }

        .toast {
            background: rgba(19, 23, 34, 0.95);
            border: 1px solid var(--gold-primary);
            color: var(--text-main);
            padding: 10px 18px;
            border-radius: var(--radius-md);
            font-size: 0.86rem;
            font-weight: 600;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.6);
            display: flex;
            align-items: center;
            gap: 8px;
            animation: toast-in 0.3s ease, toast-out 0.3s ease 2.7s forwards;
        }

        @keyframes toast-in {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }

        @keyframes toast-out {
            from { transform: translateX(0); opacity: 1; }
            to { transform: translateX(100%); opacity: 0; }
        }

        /* Responsive */
        @media (max-width: 992px) {
            .player-tools {
                display: none;
            }
            .filter-label {
                width: 100%;
            }
        }

        @media (max-width: 768px) {
            header {
                padding: 10px 16px;
            }
            .app-container {
                padding: 12px 14px;
            }
            .soundboard-grid {
                grid-template-columns: repeat(2, 1fr);
            }
            .song-table th:nth-child(4),
            .song-table th:nth-child(5),
            .song-table td:nth-child(4),
            .song-table td:nth-child(5) {
                display: none;
            }
            .player-track-info {
                width: 160px;
            }
            .vinyl-container {
                display: none;
            }
        }
    </style>
</head>
<body>

    <!-- Audio Element -->
    <audio id="audio-engine" preload="auto"></audio>

    <!-- Header -->
    <header>
        <div class="brand-container">
            <div class="brand-logo">👑</div>
            <div class="brand-text">
                <h1>RAMESH WEDDING DJ CONSOLE</h1>
                <p>Brother's Wedding • 10, 11 & 12 December 2026 • 508 Offline Songs (4.08 GB)</p>
            </div>
        </div>

        <div class="header-stats-pill">
            <span class="live-dot"></span>
            <span>100% Offline Master Library</span>
        </div>

        <div class="header-actions">
            <button class="btn-hdr" id="btn-shuffle-all" title="Shuffle Play All Songs">🔀 Shuffle All</button>
            <button class="btn-hdr" id="btn-toggle-shortcuts" title="Keyboard Shortcuts">⌨️ Keys</button>
            <button class="btn-hdr" id="btn-fullscreen-stage" title="Fullscreen DJ Stage Mode">⛶ DJ Stage</button>
        </div>
    </header>

    <!-- Main Container -->
    <div class="app-container">

        <!-- Filter Section -->
        <div class="filter-section">
            <!-- Language Filters -->
            <div class="filter-row">
                <div class="filter-label">🌍 Language:</div>
                <div class="filter-pill active" data-filter-type="lang" data-filter-val="all">All Languages <span class="count">508</span></div>
                <div class="filter-pill" data-filter-type="lang" data-filter-val="Hindi & Bollywood">🎬 Hindi & Bollywood <span class="count">421</span></div>
                <div class="filter-pill" data-filter-type="lang" data-filter-val="Punjabi">🥁 Punjabi Hits <span class="count">16</span></div>
                <div class="filter-pill" data-filter-type="lang" data-filter-val="Bagheli (Rewa)">🌾 Bagheli (Rewa) <span class="count">32</span></div>
                <div class="filter-pill" data-filter-type="lang" data-filter-val="Bhojpuri">🪕 Bhojpuri Celebrations <span class="count">39</span></div>
            </div>

            <!-- Day Filters -->
            <div class="filter-row">
                <div class="filter-label">📅 Day / Event:</div>
                <div class="filter-pill active" data-filter-type="day" data-filter-val="all">All Events <span class="count">508</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="day1">🟡 Day 1: 10 Dec (Mayan & Haldi) <span class="count">115</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="day2">🟢 Day 2: 11 Dec (Mehndi & Sangeet) <span class="count">155</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="day3">🔴 Day 3: 12 Dec (Wedding Day) <span class="count">167</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="bagheli">🌾 Bagheli Traditions (Rewa) <span class="count">32</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="bhojpuri">🪕 Bhojpuri Classics & Trending <span class="count">39</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="favorites">⭐ Favourites <span class="count" id="fav-count-badge">0</span></div>
                <div class="filter-pill" data-filter-type="day" data-filter-val="recent">🕒 Recently Played <span class="count" id="recent-count-badge">0</span></div>
            </div>
        </div>

        <!-- Ceremony Quick Launch Soundboard -->
        <div class="soundboard-section">
            <div class="section-title">
                <span>🎛️ Ceremony Soundboard (One-Click Launchpads)</span>
                <span style="font-size: 0.74rem; color: var(--text-muted); font-weight: normal;">Click pad to filter, or press ▶ to Play All in Ceremony</span>
            </div>
            <div class="soundboard-grid" id="soundboard-container">
                <!-- Dynamically populated soundboard cards -->
            </div>
        </div>

        <!-- Search & View Toolbar -->
        <div class="toolbar">
            <div class="search-box">
                <span class="search-icon">🔍</span>
                <input type="text" class="search-input" id="search-input" placeholder="Search by song name, singer, movie, ceremony, or lyrics... (Press '/' to search)">
                <button class="search-clear" id="search-clear">✖</button>
            </div>

            <div class="view-controls">
                <button class="btn-bulk-play" id="btn-play-filtered">▶ Play Current View (<span id="filtered-count">508</span>)</button>
                <button class="btn-bulk-queue" id="btn-queue-filtered">➕ Queue All</button>

                <div style="display: flex; gap: 4px; margin-left: 8px;">
                    <button class="btn-view active" id="view-mode-table" title="Table View">📋 Table</button>
                    <button class="btn-view" id="view-mode-grid" title="Cards Grid View">🗂️ Cards</button>
                    <button class="btn-view" id="view-mode-folder" title="Physical Folder Tree Explorer">📁 Folders</button>
                </div>
            </div>
        </div>

        <!-- Song List Display Container -->
        <div id="songs-render-container">
            <!-- Dynamic view injected here (Table, Grid, or Folder Tree) -->
        </div>

    </div>

    <!-- Bottom Persistent DJ Player Bar -->
    <div class="bottom-player">
        <!-- Track Info -->
        <div class="player-track-info">
            <div class="vinyl-container" id="vinyl-turntable">
                <div class="vinyl-center"></div>
            </div>
            <div class="track-text">
                <div class="track-title" id="player-title">Select a song to start</div>
                <div class="track-artist" id="player-artist">Brother's Wedding Collection</div>
                <div class="track-meta-pill" id="player-meta">100% Offline Audio • 508 Songs</div>
            </div>
            <button class="btn-act" id="player-fav-btn" title="Add to Favourites" style="margin-left: auto;">🤍</button>
        </div>

        <!-- Center Controls & Scrubber -->
        <div class="player-center">
            <div class="transport-controls">
                <button class="btn-ctrl" id="btn-shuffle" title="Toggle Shuffle (S)">🔀</button>
                <button class="btn-ctrl" id="btn-prev" title="Previous Song (Left Arrow)">⏮</button>
                <button class="btn-play-master" id="btn-play-master" title="Play / Pause (Space)">▶</button>
                <button class="btn-ctrl" id="btn-next" title="Next Song (Right Arrow)">⏭</button>
                <button class="btn-ctrl" id="btn-repeat" title="Repeat Mode: Off (R)">🔁</button>
            </div>

            <div class="scrubber-container">
                <span class="time-text" id="time-current">00:00</span>
                <div class="seek-bar-track" id="seek-track">
                    <div class="seek-progress" id="seek-progress"></div>
                </div>
                <span class="time-text" id="time-total">00:00</span>
            </div>
        </div>

        <!-- Right Tools & Volume -->
        <div class="player-tools">
            <canvas class="visualizer-canvas" id="canvas-visualizer" width="80" height="36"></canvas>
            
            <div class="volume-container">
                <button class="btn-ctrl" id="btn-mute" title="Mute (M)" style="font-size: 16px;">🔊</button>
                <input type="range" class="vol-slider" id="volume-slider" min="0" max="1" step="0.01" value="0.85">
            </div>

            <button class="btn-hdr" id="btn-toggle-queue" title="View Queue">
                <span>📑 Queue</span>
                <span class="count" id="queue-badge" style="background: var(--gold-primary); color: #000; padding: 2px 6px; border-radius: 10px; font-size: 0.7rem; font-weight: 800;">0</span>
            </button>
        </div>
    </div>

    <!-- Slide-out Queue Panel -->
    <div class="queue-panel" id="queue-panel">
        <div class="queue-header">
            <div class="queue-title">
                <span>📑 Upcoming Queue</span>
                <span class="count" id="queue-header-count" style="font-size: 0.78rem; background: rgba(245, 158, 11, 0.2); padding: 2px 8px; border-radius: 12px; color: var(--gold-light);">0</span>
            </div>
            <div style="display: flex; gap: 8px;">
                <button class="btn-hdr" id="btn-clear-queue" style="font-size: 0.74rem; padding: 4px 8px;">Clear</button>
                <button class="btn-hdr" id="btn-close-queue" style="font-size: 0.74rem; padding: 4px 8px;">✖</button>
            </div>
        </div>
        <div class="queue-list" id="queue-items-container">
            <!-- Queue items -->
        </div>
    </div>

    <!-- Fullscreen DJ Stage Mode Overlay -->
    <div class="dj-stage-overlay" id="dj-stage-overlay">
        <div class="stage-top">
            <div class="stage-brand">
                <span>👑 RAMESH WEDDING DJ STAGE CONSOLE</span>
                <span style="font-size: 0.85rem; color: var(--gold-light); background: rgba(245, 158, 11, 0.15); padding: 4px 12px; border-radius: 20px;">10-12 DECEMBER 2026</span>
            </div>
            <button class="btn-hdr" id="btn-close-stage">✖ Exit Fullscreen (Esc)</button>
        </div>

        <div class="stage-center">
            <div class="stage-vinyl" id="stage-vinyl">
                <div class="stage-vinyl-center"></div>
            </div>

            <div class="stage-song-title" id="stage-title">Select a Song</div>
            <div class="stage-song-artist" id="stage-artist">Brother's Wedding Collection</div>

            <div class="stage-badges">
                <span class="song-tag tag-day" id="stage-day-badge">Day 1: Mayan & Haldi</span>
                <span class="song-tag tag-cero" id="stage-cero-badge">Haldi Ceremony</span>
                <span class="song-tag tag-lang" id="stage-lang-badge">Hindi & Bollywood</span>
            </div>

            <canvas class="stage-canvas" id="stage-canvas"></canvas>

            <div class="scrubber-container" style="max-width: 700px; margin-top: 10px;">
                <span class="time-text" id="stage-time-current">00:00</span>
                <div class="seek-bar-track" id="stage-seek-track">
                    <div class="seek-progress" id="stage-seek-progress"></div>
                </div>
                <span class="time-text" id="stage-time-total">00:00</span>
            </div>

            <div class="stage-controls">
                <button class="btn-ctrl" id="stage-btn-shuffle" style="font-size: 24px;">🔀</button>
                <button class="btn-ctrl" id="stage-btn-prev" style="font-size: 26px;">⏮</button>
                <button class="stage-btn-play" id="stage-btn-play">▶</button>
                <button class="btn-ctrl" id="stage-btn-next" style="font-size: 26px;">⏭</button>
                <button class="btn-ctrl" id="stage-btn-repeat" style="font-size: 24px;">🔁</button>
            </div>
        </div>

        <div style="display: flex; justify-content: center; gap: 20px; color: var(--text-muted); font-size: 0.84rem;">
            <span>Press <kbd style="color: var(--gold-light);">Space</kbd> to Play/Pause</span>
            <span>•</span>
            <span>Press <kbd style="color: var(--gold-light);">Right Arrow</kbd> for Next</span>
            <span>•</span>
            <span>Press <kbd style="color: var(--gold-light);">Left Arrow</kbd> for Previous</span>
            <span>•</span>
            <span>Press <kbd style="color: var(--gold-light);">Esc</kbd> to Exit Stage</span>
        </div>
    </div>

    <!-- Keyboard Shortcuts Modal -->
    <div class="modal-overlay" id="modal-shortcuts">
        <div class="modal-card">
            <div class="modal-head">
                <h3 style="color: var(--gold-light); font-size: 1.05rem;">⌨️ DJ Keyboard Shortcuts</h3>
                <button class="btn-hdr" id="btn-close-modal" style="padding: 4px 8px;">✖</button>
            </div>
            <div class="modal-body">
                <div class="shortcut-item">
                    <span>Play / Pause Master</span>
                    <span class="key-badge">Space</span>
                </div>
                <div class="shortcut-item">
                    <span>Next Track</span>
                    <span class="key-badge">→ (Right Arrow)</span>
                </div>
                <div class="shortcut-item">
                    <span>Previous Track / Restart</span>
                    <span class="key-badge">← (Left Arrow)</span>
                </div>
                <div class="shortcut-item">
                    <span>Volume Up (+5%)</span>
                    <span class="key-badge">↑ (Up Arrow)</span>
                </div>
                <div class="shortcut-item">
                    <span>Volume Down (-5%)</span>
                    <span class="key-badge">↓ (Down Arrow)</span>
                </div>
                <div class="shortcut-item">
                    <span>Toggle Mute / Unmute</span>
                    <span class="key-badge">M</span>
                </div>
                <div class="shortcut-item">
                    <span>Toggle Shuffle Mode</span>
                    <span class="key-badge">S</span>
                </div>
                <div class="shortcut-item">
                    <span>Cycle Repeat Mode (Off / All / One)</span>
                    <span class="key-badge">R</span>
                </div>
                <div class="shortcut-item">
                    <span>Toggle Fullscreen DJ Stage Mode</span>
                    <span class="key-badge">F</span>
                </div>
                <div class="shortcut-item">
                    <span>Instant Search Focus</span>
                    <span class="key-badge">/</span>
                </div>
                <div class="shortcut-item">
                    <span>Close Modal / Exit Fullscreen</span>
                    <span class="key-badge">Escape</span>
                </div>
            </div>
        </div>
    </div>

    <!-- Toast Notifications Container -->
    <div class="toast-container" id="toast-container"></div>

    <!-- Client-Side Engine -->
    <script>
        // 508 Offline Songs Catalog
        const ALL_SONGS = __SONGS_JSON__;

        // App State
        const state = {
            currentSongIndex: -1,
            isPlaying: false,
            isShuffle: false,
            repeatMode: 'off', // 'off', 'all', 'one'
            volume: 0.85,
            isMuted: false,
            viewMode: 'table', // 'table', 'grid', 'folder'
            activeFilters: {
                lang: 'all',
                day: 'all',
                cero: 'all',
                search: ''
            },
            queue: [], // Array of song IDs
            favorites: new Set(JSON.parse(localStorage.getItem('wedding_dj_favs') || '[]')),
            recent: JSON.parse(localStorage.getItem('wedding_dj_recent') || '[]')
        };

        // DOM Elements
        const audio = document.getElementById('audio-engine');
        const btnPlayMaster = document.getElementById('btn-play-master');
        const btnPrev = document.getElementById('btn-prev');
        const btnNext = document.getElementById('btn-next');
        const btnShuffle = document.getElementById('btn-shuffle');
        const btnRepeat = document.getElementById('btn-repeat');
        const btnMute = document.getElementById('btn-mute');
        const volumeSlider = document.getElementById('volume-slider');
        const seekTrack = document.getElementById('seek-track');
        const seekProgress = document.getElementById('seek-progress');
        const timeCurrent = document.getElementById('time-current');
        const timeTotal = document.getElementById('time-total');
        const playerTitle = document.getElementById('player-title');
        const playerArtist = document.getElementById('player-artist');
        const playerMeta = document.getElementById('player-meta');
        const playerFavBtn = document.getElementById('player-fav-btn');
        const vinylTurntable = document.getElementById('vinyl-turntable');
        const searchInput = document.getElementById('search-input');
        const searchClear = document.getElementById('search-clear');
        const filteredCount = document.getElementById('filtered-count');
        const soundboardContainer = document.getElementById('soundboard-container');
        const songsRenderContainer = document.getElementById('songs-render-container');
        const queuePanel = document.getElementById('queue-panel');
        const btnToggleQueue = document.getElementById('btn-toggle-queue');
        const btnCloseQueue = document.getElementById('btn-close-queue');
        const btnClearQueue = document.getElementById('btn-clear-queue');
        const queueBadge = document.getElementById('queue-badge');
        const queueHeaderCount = document.getElementById('queue-header-count');
        const queueItemsContainer = document.getElementById('queue-items-container');
        const djStageOverlay = document.getElementById('dj-stage-overlay');
        const btnFullscreenStage = document.getElementById('btn-fullscreen-stage');
        const btnCloseStage = document.getElementById('btn-close-stage');
        const stageBtnPlay = document.getElementById('stage-btn-play');
        const stageBtnPrev = document.getElementById('stage-btn-prev');
        const stageBtnNext = document.getElementById('stage-btn-next');
        const stageBtnShuffle = document.getElementById('stage-btn-shuffle');
        const stageBtnRepeat = document.getElementById('stage-btn-repeat');
        const stageTitle = document.getElementById('stage-title');
        const stageArtist = document.getElementById('stage-artist');
        const stageDayBadge = document.getElementById('stage-day-badge');
        const stageCeroBadge = document.getElementById('stage-cero-badge');
        const stageLangBadge = document.getElementById('stage-lang-badge');
        const stageVinyl = document.getElementById('stage-vinyl');
        const stageTimeCurrent = document.getElementById('stage-time-current');
        const stageTimeTotal = document.getElementById('stage-time-total');
        const stageSeekTrack = document.getElementById('stage-seek-track');
        const stageSeekProgress = document.getElementById('stage-seek-progress');
        const modalShortcuts = document.getElementById('modal-shortcuts');
        const btnToggleShortcuts = document.getElementById('btn-toggle-shortcuts');
        const btnCloseModal = document.getElementById('btn-close-modal');
        const favCountBadge = document.getElementById('fav-count-badge');
        const recentCountBadge = document.getElementById('recent-count-badge');
        const btnPlayFiltered = document.getElementById('btn-play-filtered');
        const btnQueueFiltered = document.getElementById('btn-queue-filtered');
        const btnShuffleAll = document.getElementById('btn-shuffle-all');

        // Ceremony Configuration for Soundboard
        const CEREMONIES = [
            { name: 'Mayan & Mandap', icon: '🪔', filter: 'Mayan & Mandap' },
            { name: 'Haldi Ceremony', icon: '✨', filter: 'Haldi Ceremony' },
            { name: 'Mehndi', icon: '🌿', filter: 'Mehndi' },
            { name: 'Ladies Sangeet', icon: '💃', filter: 'Ladies Sangeet' },
            { name: 'Couple Dance', icon: '💑', filter: 'Couple Dance' },
            { name: 'Family Dance', icon: '👨‍👩‍👧‍👦', filter: 'Family Dance' },
            { name: 'Friends Dance', icon: '🕺', filter: 'Friends Dance' },
            { name: 'DJ Party', icon: '🔊', filter: 'DJ Party' },
            { name: 'Baraat & Groom Entry', icon: '🥁', filter: 'Baraat & Groom Entry' },
            { name: 'Ghodi & Procession', icon: '🐎', filter: 'Ghodi & Procession' },
            { name: 'Bride Entry', icon: '👰', filter: 'Bride Entry' },
            { name: 'Varmala & Jaimala', icon: '🌸', filter: 'Varmala & Jaimala' },
            { name: 'Pheras & Instrumental', icon: '🔥', filter: 'Pheras & Instrumental' },
            { name: 'Wedding Rituals', icon: '🙏', filter: 'Wedding Rituals' },
            { name: 'Reception', icon: '🥂', filter: 'Reception' },
            { name: 'Vidaai', icon: '🕊️', filter: 'Vidaai' },
            { name: 'Funny & Entertaining', icon: '🎉', filter: 'Funny & Entertaining' }
        ];

        // Format Seconds to MM:SS
        function formatTime(seconds) {
            if (isNaN(seconds) || seconds < 0) return '00:00';
            const m = Math.floor(seconds / 60);
            const s = Math.floor(seconds % 60);
            return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
        }

        // Toast Notification
        function showToast(message) {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = 'toast';
            toast.innerHTML = `<span>🎵</span> <span>${message}</span>`;
            container.appendChild(toast);
            setTimeout(() => {
                toast.remove();
            }, 3000);
        }

        // Audio Context & Visualizer
        let audioCtx, analyser, dataArray;
        let isAudioContextInit = false;

        function initAudioVisualizer() {
            if (isAudioContextInit) return;
            try {
                const AudioContext = window.AudioContext || window.webkitAudioContext;
                if (!AudioContext) return;
                audioCtx = new AudioContext();
                analyser = audioCtx.createAnalyser();
                const source = audioCtx.createMediaElementSource(audio);
                source.connect(analyser);
                analyser.connect(audioCtx.destination);
                analyser.fftSize = 64;
                const bufferLength = analyser.frequencyBinCount;
                dataArray = new Uint8Array(bufferLength);
                isAudioContextInit = true;
                drawVisualizer();
            } catch (e) {
                console.warn('AudioContext visualization not available in current protocol:', e);
            }
        }

        function drawVisualizer() {
            requestAnimationFrame(drawVisualizer);
            if (!analyser || !dataArray) return;
            analyser.getByteFrequencyData(dataArray);

            // Bottom Player Canvas
            const canvas = document.getElementById('canvas-visualizer');
            if (canvas) {
                const ctx = canvas.getContext('2d');
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                const barWidth = (canvas.width / 16) - 1;
                for (let i = 0; i < 16; i++) {
                    const barHeight = (dataArray[i * 2] / 255) * canvas.height;
                    ctx.fillStyle = '#f59e0b';
                    ctx.fillRect(i * (barWidth + 1), canvas.height - barHeight, barWidth, barHeight);
                }
            }

            // Stage Mode Canvas
            const stageCanvas = document.getElementById('stage-canvas');
            if (stageCanvas && djStageOverlay.classList.contains('active')) {
                const sCtx = stageCanvas.getContext('2d');
                sCtx.clearRect(0, 0, stageCanvas.width, stageCanvas.height);
                const barWidth = (stageCanvas.width / 32) - 2;
                for (let i = 0; i < 32; i++) {
                    const barHeight = (dataArray[i] / 255) * stageCanvas.height;
                    const grad = sCtx.createLinearGradient(0, stageCanvas.height, 0, 0);
                    grad.addColorStop(0, '#f59e0b');
                    grad.addColorStop(1, '#fbbf24');
                    sCtx.fillStyle = grad;
                    sCtx.fillRect(i * (barWidth + 2), stageCanvas.height - barHeight, barWidth, barHeight);
                }
            }
        }

        // Filter Songs Logic
        function getFilteredSongs() {
            return ALL_SONGS.filter(song => {
                // Language Filter
                if (state.activeFilters.lang !== 'all' && song.language !== state.activeFilters.lang) {
                    return false;
                }

                // Day Filter
                if (state.activeFilters.day !== 'all') {
                    if (state.activeFilters.day === 'favorites') {
                        if (!state.favorites.has(song.id)) return false;
                    } else if (state.activeFilters.day === 'recent') {
                        if (!state.recent.includes(song.id)) return false;
                    } else if (state.activeFilters.day !== song.day_key) {
                        return false;
                    }
                }

                // Ceremony Filter
                if (state.activeFilters.cero !== 'all') {
                    if (song.function !== state.activeFilters.cero) return false;
                }

                // Search Filter
                if (state.activeFilters.search) {
                    const query = state.activeFilters.search.toLowerCase();
                    const text = (song.title + ' ' + song.artist + ' ' + song.movie + ' ' + song.function + ' ' + song.folder).toLowerCase();
                    if (!text.includes(query)) return false;
                }

                return true;
            });
        }

        // Update UI Counters & Soundboard
        function updateUI() {
            favCountBadge.textContent = state.favorites.size;
            recentCountBadge.textContent = state.recent.length;
            const currentList = getFilteredSongs();
            filteredCount.textContent = currentList.length;
            renderSongs();
        }

        // Play Song Function
        function playSongByIndex(index, startPlayback = true) {
            if (index < 0 || index >= ALL_SONGS.length) return;
            state.currentSongIndex = index;
            const song = ALL_SONGS[index];

            // Set audio source (encodeURI for spaces and characters)
            audio.src = encodeURI(song.path);
            audio.load();

            // Update UI text
            playerTitle.textContent = song.title;
            playerArtist.textContent = song.artist + (song.movie ? ' • ' + song.movie : '');
            playerMeta.textContent = `${song.day} • ${song.function} • ${song.language}`;
            
            // Stage text
            stageTitle.textContent = song.title;
            stageArtist.textContent = song.artist + (song.movie ? ' (' + song.movie + ')' : '');
            stageDayBadge.textContent = song.day;
            stageCeroBadge.textContent = song.function;
            stageLangBadge.textContent = song.language;

            // Favorite button update
            updateFavButton(song.id);

            // Record in Recent
            state.recent = [song.id, ...state.recent.filter(id => id !== song.id)].slice(0, 50);
            localStorage.setItem('wedding_dj_recent', JSON.stringify(state.recent));
            recentCountBadge.textContent = state.recent.length;

            if (startPlayback) {
                audio.play().then(() => {
                    state.isPlaying = true;
                    updatePlaybackState();
                    initAudioVisualizer();
                }).catch(err => {
                    console.warn('Playback error or blocked:', err);
                    state.isPlaying = false;
                    updatePlaybackState();
                });
            }

            highlightCurrentSong();
            renderQueue();
        }

        function togglePlay() {
            if (state.currentSongIndex === -1) {
                const list = getFilteredSongs();
                if (list.length > 0) {
                    const firstIdx = ALL_SONGS.findIndex(s => s.id === list[0].id);
                    playSongByIndex(firstIdx, true);
                }
                return;
            }

            if (audio.paused) {
                audio.play().then(() => {
                    state.isPlaying = true;
                    updatePlaybackState();
                    initAudioVisualizer();
                }).catch(err => console.error(err));
            } else {
                audio.pause();
                state.isPlaying = false;
                updatePlaybackState();
            }
        }

        function nextSong() {
            // Check Queue first
            if (state.queue.length > 0) {
                const nextId = state.queue.shift();
                const idx = ALL_SONGS.findIndex(s => s.id === nextId);
                if (idx !== -1) {
                    playSongByIndex(idx, true);
                    renderQueue();
                    return;
                }
            }

            const list = getFilteredSongs();
            if (list.length === 0) return;

            if (state.isShuffle) {
                const randomIdx = Math.floor(Math.random() * list.length);
                const masterIdx = ALL_SONGS.findIndex(s => s.id === list[randomIdx].id);
                playSongByIndex(masterIdx, true);
                return;
            }

            // Normal Sequential
            const currentFilteredPos = list.findIndex(s => s.id === ALL_SONGS[state.currentSongIndex]?.id);
            if (currentFilteredPos !== -1 && currentFilteredPos < list.length - 1) {
                const nextMasterIdx = ALL_SONGS.findIndex(s => s.id === list[currentFilteredPos + 1].id);
                playSongByIndex(nextMasterIdx, true);
            } else if (state.repeatMode === 'all') {
                const firstMasterIdx = ALL_SONGS.findIndex(s => s.id === list[0].id);
                playSongByIndex(firstMasterIdx, true);
            } else {
                showToast('End of playlist');
            }
        }

        function prevSong() {
            if (audio.currentTime > 3) {
                audio.currentTime = 0;
                return;
            }

            const list = getFilteredSongs();
            if (list.length === 0) return;

            const currentFilteredPos = list.findIndex(s => s.id === ALL_SONGS[state.currentSongIndex]?.id);
            if (currentFilteredPos > 0) {
                const prevMasterIdx = ALL_SONGS.findIndex(s => s.id === list[currentFilteredPos - 1].id);
                playSongByIndex(prevMasterIdx, true);
            } else {
                audio.currentTime = 0;
            }
        }

        function updatePlaybackState() {
            if (state.isPlaying) {
                btnPlayMaster.innerHTML = '⏸';
                stageBtnPlay.innerHTML = '⏸';
                vinylTurntable.classList.add('spinning');
                stageVinyl.classList.add('spinning');
            } else {
                btnPlayMaster.innerHTML = '▶';
                stageBtnPlay.innerHTML = '▶';
                vinylTurntable.classList.remove('spinning');
                stageVinyl.classList.remove('spinning');
            }
            highlightCurrentSong();
        }

        function highlightCurrentSong() {
            document.querySelectorAll('.song-row').forEach(row => {
                const sId = parseInt(row.getAttribute('data-id'));
                if (state.currentSongIndex !== -1 && sId === ALL_SONGS[state.currentSongIndex].id) {
                    row.classList.add('playing');
                } else {
                    row.classList.remove('playing');
                }
            });

            document.querySelectorAll('.song-card').forEach(card => {
                const sId = parseInt(card.getAttribute('data-id'));
                if (state.currentSongIndex !== -1 && sId === ALL_SONGS[state.currentSongIndex].id) {
                    card.classList.add('playing');
                } else {
                    card.classList.remove('playing');
                }
            });
        }

        function updateFavButton(songId) {
            if (state.favorites.has(songId)) {
                playerFavBtn.innerHTML = '❤️';
                playerFavBtn.classList.add('fav-active');
            } else {
                playerFavBtn.innerHTML = '🤍';
                playerFavBtn.classList.remove('fav-active');
            }
        }

        function toggleFavorite(songId) {
            if (state.favorites.has(songId)) {
                state.favorites.delete(songId);
                showToast('Removed from Favourites');
            } else {
                state.favorites.add(songId);
                showToast('Added to Favourites ⭐');
            }
            localStorage.setItem('wedding_dj_favs', JSON.stringify(Array.from(state.favorites)));
            if (state.currentSongIndex !== -1 && ALL_SONGS[state.currentSongIndex].id === songId) {
                updateFavButton(songId);
            }
            updateUI();
        }

        // Queue Controls
        function addToQueue(songId) {
            state.queue.push(songId);
            renderQueue();
            const song = ALL_SONGS.find(s => s.id === songId);
            showToast(`Queued: ${song.title}`);
        }

        function playNextInQueue(songId) {
            state.queue.unshift(songId);
            renderQueue();
            const song = ALL_SONGS.find(s => s.id === songId);
            showToast(`Up Next: ${song.title}`);
        }

        function removeFromQueue(qIndex) {
            state.queue.splice(qIndex, 1);
            renderQueue();
        }

        function clearQueue() {
            state.queue = [];
            renderQueue();
            showToast('Queue cleared');
        }

        function renderQueue() {
            queueBadge.textContent = state.queue.length;
            queueHeaderCount.textContent = state.queue.length;

            if (state.queue.length === 0) {
                queueItemsContainer.innerHTML = '<div style="color: var(--text-muted); text-align: center; padding: 40px 20px; font-size: 0.88rem;">Queue is empty.<br>Click ➕ on any song to queue it.</div>';
                return;
            }

            queueItemsContainer.innerHTML = state.queue.map((sId, qIdx) => {
                const song = ALL_SONGS.find(s => s.id === sId);
                if (!song) return '';
                return `
                    <div class="queue-item">
                        <div class="queue-item-info" onclick="playSongByIndex(${ALL_SONGS.findIndex(s => s.id === sId)}, true); state.queue.splice(${qIdx}, 1); renderQueue();">
                            <div class="queue-item-title">${song.title}</div>
                            <div class="queue-item-artist">${song.artist} • ${song.function}</div>
                        </div>
                        <button class="btn-act" onclick="removeFromQueue(${qIdx})" title="Remove" style="width: 26px; height: 26px; font-size: 12px;">✖</button>
                    </div>
                `;
            }).join('');
        }

        // Render Soundboard
        function renderSoundboard() {
            soundboardContainer.innerHTML = CEREMONIES.map(cero => {
                const count = ALL_SONGS.filter(s => s.function === cero.filter).length;
                const isActive = state.activeFilters.cero === cero.filter;
                return `
                    <div class="ceremony-card ${isActive ? 'active' : ''}" onclick="filterByCeremony('${cero.filter}')">
                        <div class="cero-top">
                            <span class="cero-icon">${cero.icon}</span>
                            <span class="cero-count">${count} Songs</span>
                        </div>
                        <div class="cero-name">${cero.name}</div>
                        <button class="cero-play-all" onclick="event.stopPropagation(); playAllInCeremony('${cero.filter}')">
                            <span>▶ Play All</span>
                        </button>
                    </div>
                `;
            }).join('');
        }

        function filterByCeremony(ceroName) {
            if (state.activeFilters.cero === ceroName) {
                state.activeFilters.cero = 'all';
            } else {
                state.activeFilters.cero = ceroName;
            }
            renderSoundboard();
            updateUI();
        }

        function playAllInCeremony(ceroName) {
            state.activeFilters.cero = ceroName;
            renderSoundboard();
            const list = ALL_SONGS.filter(s => s.function === ceroName);
            if (list.length > 0) {
                const masterIdx = ALL_SONGS.findIndex(s => s.id === list[0].id);
                playSongByIndex(masterIdx, true);
                showToast(`Playing all ${list.length} songs in ${ceroName}`);
            }
            updateUI();
        }

        // Render Songs Based on View Mode
        function renderSongs() {
            const list = getFilteredSongs();

            if (list.length === 0) {
                songsRenderContainer.innerHTML = `
                    <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 60px 20px; text-align: center;">
                        <div style="font-size: 3rem; margin-bottom: 12px;">🔍</div>
                        <h3 style="color: var(--gold-light); margin-bottom: 8px;">No Songs Found</h3>
                        <p style="color: var(--text-muted); font-size: 0.9rem;">Try selecting another ceremony, clearing your search query, or checking language filters.</p>
                    </div>
                `;
                return;
            }

            if (state.viewMode === 'table') {
                renderTableView(list);
            } else if (state.viewMode === 'grid') {
                renderGridView(list);
            } else if (state.viewMode === 'folder') {
                renderFolderView();
            }
        }

        // Table View
        function renderTableView(list) {
            let rowsHtml = list.map((song, idx) => {
                const isPlaying = state.currentSongIndex !== -1 && ALL_SONGS[state.currentSongIndex].id === song.id;
                const isFav = state.favorites.has(song.id);
                const masterIdx = ALL_SONGS.findIndex(s => s.id === song.id);

                return `
                    <tr class="song-row ${isPlaying ? 'playing' : ''}" data-id="${song.id}" onclick="playSongByIndex(${masterIdx}, true)">
                        <td class="col-idx">
                            ${isPlaying && state.isPlaying ? '🔊' : (idx + 1)}
                        </td>
                        <td class="col-title">
                            <span class="song-title-text">${song.title}</span>
                            <span class="song-artist-text">${song.artist}${song.movie ? ' • ' + song.movie : ''}</span>
                        </td>
                        <td>
                            <span class="song-tag tag-day">${song.day}</span>
                        </td>
                        <td>
                            <span class="song-tag tag-cero">${song.function}</span>
                        </td>
                        <td>
                            <span class="song-tag tag-lang">${song.language}</span>
                        </td>
                        <td>
                            <span style="font-size: 0.78rem; color: var(--text-muted);">${song.size_mb} MB</span>
                        </td>
                        <td>
                            <div class="row-actions" onclick="event.stopPropagation()">
                                <button class="btn-act ${isFav ? 'fav-active' : ''}" onclick="toggleFavorite(${song.id})" title="Favourite">
                                    ${isFav ? '❤️' : '🤍'}
                                </button>
                                <button class="btn-act" onclick="playNextInQueue(${song.id})" title="Play Next">⏭</button>
                                <button class="btn-act" onclick="addToQueue(${song.id})" title="Add to Queue">➕</button>
                            </div>
                        </td>
                    </tr>
                `;
            }).join('');

            songsRenderContainer.innerHTML = `
                <table class="song-table">
                    <thead>
                        <tr>
                            <th style="width: 50px; text-align: center;">#</th>
                            <th>Song Title & Artist</th>
                            <th>Day / Event</th>
                            <th>Ceremony</th>
                            <th>Language</th>
                            <th>Size</th>
                            <th style="text-align: right; width: 140px;">Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rowsHtml}
                    </tbody>
                </table>
            `;
        }

        // Grid View
        function renderGridView(list) {
            let cardsHtml = list.map(song => {
                const isPlaying = state.currentSongIndex !== -1 && ALL_SONGS[state.currentSongIndex].id === song.id;
                const isFav = state.favorites.has(song.id);
                const masterIdx = ALL_SONGS.findIndex(s => s.id === song.id);

                return `
                    <div class="song-card ${isPlaying ? 'playing' : ''}" data-id="${song.id}" onclick="playSongByIndex(${masterIdx}, true)">
                        <div class="card-top">
                            <div class="card-play-btn">
                                ${isPlaying && state.isPlaying ? '⏸' : '▶'}
                            </div>
                            <div class="card-info">
                                <div class="card-title">${song.title}</div>
                                <div class="card-artist">${song.artist}</div>
                            </div>
                        </div>

                        <div class="card-tags">
                            <span class="song-tag tag-day">${song.day}</span>
                            <span class="song-tag tag-cero">${song.function}</span>
                            <span class="song-tag tag-lang">${song.language}</span>
                        </div>

                        <div class="card-bottom">
                            <span>${song.movie || 'Wedding Anthem'} • ${song.size_mb} MB</span>
                            <div class="row-actions" onclick="event.stopPropagation()">
                                <button class="btn-act ${isFav ? 'fav-active' : ''}" onclick="toggleFavorite(${song.id})" title="Favourite">
                                    ${isFav ? '❤️' : '🤍'}
                                </button>
                                <button class="btn-act" onclick="playNextInQueue(${song.id})" title="Play Next">⏭</button>
                                <button class="btn-act" onclick="addToQueue(${song.id})" title="Add to Queue">➕</button>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');

            songsRenderContainer.innerHTML = `<div class="song-grid">${cardsHtml}</div>`;
        }

        // Physical Folder Tree Explorer View
        function renderFolderView() {
            const folders = {};
            ALL_SONGS.forEach(s => {
                if (!folders[s.folder]) folders[s.folder] = {};
                if (!folders[s.folder][s.subfolder]) folders[s.folder][s.subfolder] = [];
                folders[s.folder][s.subfolder].push(s);
            });

            let html = '<div class="folder-tree">';
            for (const [topFolder, subfolders] of Object.entries(folders)) {
                let totalFolderSongs = 0;
                for (const arr of Object.values(subfolders)) totalFolderSongs += arr.length;

                html += `
                    <div class="folder-group">
                        <div class="folder-header">
                            <div class="folder-title-left">
                                <span class="folder-icon">📁</span>
                                <div>
                                    <div class="folder-name">${topFolder}</div>
                                    <div class="folder-stats">${totalFolderSongs} Songs Across ${Object.keys(subfolders).length} Subfolders</div>
                                </div>
                            </div>
                            <div class="folder-actions">
                                <button class="btn-hdr" onclick="playAllInTopFolder('${topFolder}')">▶ Play Folder (${totalFolderSongs})</button>
                                <button class="btn-hdr" onclick="queueAllInTopFolder('${topFolder}')">➕ Queue</button>
                            </div>
                        </div>
                        <div class="subfolder-container">
                `;

                for (const [subName, sList] of Object.entries(subfolders)) {
                    html += `
                        <div class="subfolder-block">
                            <div class="subfolder-head">
                                <span class="subfolder-title">📂 ${subName || 'General'} (${sList.length} Songs)</span>
                                <div style="display: flex; gap: 6px;">
                                    <button class="cero-play-all" onclick="playSubfolderSongs('${topFolder}', '${subName}')">▶ Play</button>
                                    <button class="cero-play-all" onclick="queueSubfolderSongs('${topFolder}', '${subName}')">➕ Queue</button>
                                </div>
                            </div>
                            <div style="display: flex; flex-wrap: wrap; gap: 6px;">
                                ${sList.map(s => `
                                    <div class="song-pill" onclick="playSongByIndex(${ALL_SONGS.findIndex(item => item.id === s.id)}, true)" style="background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08); padding: 4px 10px; border-radius: 6px; font-size: 0.8rem; cursor: pointer; display: flex; align-items: center; gap: 6px;">
                                        <span>🎵</span>
                                        <span style="font-weight: 600;">${s.title}</span>
                                        <span style="color: var(--text-muted); font-size: 0.72rem;">${s.artist}</span>
                                    </div>
                                `).join('')}
                            </div>
                        </div>
                    `;
                }

                html += '</div></div>';
            }
            html += '</div>';

            songsRenderContainer.innerHTML = html;
        }

        function playAllInTopFolder(topFolder) {
            const list = ALL_SONGS.filter(s => s.folder === topFolder);
            if (list.length > 0) {
                const masterIdx = ALL_SONGS.findIndex(s => s.id === list[0].id);
                playSongByIndex(masterIdx, true);
                showToast(`Playing all ${list.length} songs in ${topFolder}`);
            }
        }

        function queueAllInTopFolder(topFolder) {
            const list = ALL_SONGS.filter(s => s.folder === topFolder);
            list.forEach(s => state.queue.push(s.id));
            renderQueue();
            showToast(`Added ${list.length} songs to Queue`);
        }

        function playSubfolderSongs(topFolder, subName) {
            const list = ALL_SONGS.filter(s => s.folder === topFolder && s.subfolder === subName);
            if (list.length > 0) {
                const masterIdx = ALL_SONGS.findIndex(s => s.id === list[0].id);
                playSongByIndex(masterIdx, true);
                showToast(`Playing ${list.length} songs from ${subName}`);
            }
        }

        function queueSubfolderSongs(topFolder, subName) {
            const list = ALL_SONGS.filter(s => s.folder === topFolder && s.subfolder === subName);
            list.forEach(s => state.queue.push(s.id));
            renderQueue();
            showToast(`Queued ${list.length} songs from ${subName}`);
        }

        // Audio Engine Event Listeners
        audio.addEventListener('timeupdate', () => {
            const cur = audio.currentTime;
            const dur = audio.duration;
            timeCurrent.textContent = formatTime(cur);
            stageTimeCurrent.textContent = formatTime(cur);

            if (!isNaN(dur) && dur > 0) {
                const pct = (cur / dur) * 100;
                seekProgress.style.width = pct + '%';
                stageSeekProgress.style.width = pct + '%';
                timeTotal.textContent = formatTime(dur);
                stageTimeTotal.textContent = formatTime(dur);
            }
        });

        audio.addEventListener('ended', () => {
            if (state.repeatMode === 'one') {
                audio.currentTime = 0;
                audio.play();
            } else {
                nextSong();
            }
        });

        audio.addEventListener('play', () => {
            state.isPlaying = true;
            updatePlaybackState();
        });

        audio.addEventListener('pause', () => {
            state.isPlaying = false;
            updatePlaybackState();
        });

        audio.addEventListener('error', (e) => {
            console.error('Audio load error:', e);
            showToast('⚠️ Audio loading error. Trying fallback...');
        });

        // Scrubber Seeking
        seekTrack.addEventListener('click', (e) => {
            const rect = seekTrack.getBoundingClientRect();
            const clickX = e.clientX - rect.left;
            const pct = clickX / rect.width;
            if (!isNaN(audio.duration)) {
                audio.currentTime = pct * audio.duration;
            }
        });

        stageSeekTrack.addEventListener('click', (e) => {
            const rect = stageSeekTrack.getBoundingClientRect();
            const clickX = e.clientX - rect.left;
            const pct = clickX / rect.width;
            if (!isNaN(audio.duration)) {
                audio.currentTime = pct * audio.duration;
            }
        });

        // Volume Controls
        volumeSlider.addEventListener('input', (e) => {
            state.volume = parseFloat(e.target.value);
            audio.volume = state.volume;
            state.isMuted = state.volume === 0;
            btnMute.innerHTML = state.isMuted ? '🔇' : (state.volume > 0.5 ? '🔊' : '🔉');
        });

        btnMute.addEventListener('click', () => {
            if (state.isMuted) {
                audio.volume = state.volume || 0.85;
                state.isMuted = false;
                btnMute.innerHTML = '🔊';
                volumeSlider.value = state.volume;
            } else {
                audio.volume = 0;
                state.isMuted = true;
                btnMute.innerHTML = '🔇';
                volumeSlider.value = 0;
            }
        });

        // Shuffle & Repeat
        btnShuffle.addEventListener('click', toggleShuffleMode);
        stageBtnShuffle.addEventListener('click', toggleShuffleMode);

        function toggleShuffleMode() {
            state.isShuffle = !state.isShuffle;
            btnShuffle.classList.toggle('active', state.isShuffle);
            stageBtnShuffle.classList.toggle('active', state.isShuffle);
            showToast(state.isShuffle ? '🔀 Shuffle On' : '➡️ Shuffle Off');
        }

        btnRepeat.addEventListener('click', cycleRepeatMode);
        stageBtnRepeat.addEventListener('click', cycleRepeatMode);

        function cycleRepeatMode() {
            if (state.repeatMode === 'off') {
                state.repeatMode = 'all';
                btnRepeat.innerHTML = '🔁';
                btnRepeat.classList.add('active');
                stageBtnRepeat.innerHTML = '🔁';
                stageBtnRepeat.classList.add('active');
                showToast('🔁 Repeat All');
            } else if (state.repeatMode === 'all') {
                state.repeatMode = 'one';
                btnRepeat.innerHTML = '🔂';
                btnRepeat.classList.add('active');
                stageBtnRepeat.innerHTML = '🔂';
                stageBtnRepeat.classList.add('active');
                showToast('🔂 Repeat One');
            } else {
                state.repeatMode = 'off';
                btnRepeat.innerHTML = '🔁';
                btnRepeat.classList.remove('active');
                stageBtnRepeat.innerHTML = '🔁';
                stageBtnRepeat.classList.remove('active');
                showToast('➡️ Repeat Off');
            }
        }

        // Search Input
        searchInput.addEventListener('input', (e) => {
            state.activeFilters.search = e.target.value.trim();
            searchClear.style.display = state.activeFilters.search ? 'block' : 'none';
            updateUI();
        });

        searchClear.addEventListener('click', () => {
            searchInput.value = '';
            state.activeFilters.search = '';
            searchClear.style.display = 'none';
            updateUI();
        });

        // View Mode Switcher
        document.getElementById('view-mode-table').addEventListener('click', () => setViewMode('table'));
        document.getElementById('view-mode-grid').addEventListener('click', () => setViewMode('grid'));
        document.getElementById('view-mode-folder').addEventListener('click', () => setViewMode('folder'));

        function setViewMode(mode) {
            state.viewMode = mode;
            document.querySelectorAll('.btn-view').forEach(btn => btn.classList.remove('active'));
            document.getElementById(`view-mode-${mode}`).classList.add('active');
            renderSongs();
        }

        // Filter Pills Event Listeners
        document.querySelectorAll('.filter-pill').forEach(pill => {
            pill.addEventListener('click', () => {
                const type = pill.getAttribute('data-filter-type');
                const val = pill.getAttribute('data-filter-val');

                document.querySelectorAll(`.filter-pill[data-filter-type="${type}"]`).forEach(p => p.classList.remove('active'));
                pill.classList.add('active');

                if (type === 'lang') state.activeFilters.lang = val;
                if (type === 'day') state.activeFilters.day = val;

                updateUI();
            });
        });

        // Bulk Buttons
        btnPlayFiltered.addEventListener('click', () => {
            const list = getFilteredSongs();
            if (list.length > 0) {
                const masterIdx = ALL_SONGS.findIndex(s => s.id === list[0].id);
                playSongByIndex(masterIdx, true);
                showToast(`Playing ${list.length} songs in current view`);
            }
        });

        btnQueueFiltered.addEventListener('click', () => {
            const list = getFilteredSongs();
            list.forEach(s => state.queue.push(s.id));
            renderQueue();
            showToast(`Added ${list.length} songs to Queue`);
        });

        btnShuffleAll.addEventListener('click', () => {
            state.isShuffle = true;
            btnShuffle.classList.add('active');
            stageBtnShuffle.classList.add('active');
            const list = getFilteredSongs();
            if (list.length > 0) {
                const randomPos = Math.floor(Math.random() * list.length);
                const masterIdx = ALL_SONGS.findIndex(s => s.id === list[randomPos].id);
                playSongByIndex(masterIdx, true);
                showToast(`🔀 Shuffled and playing all ${list.length} songs`);
            }
        });

        // Queue Drawer
        btnToggleQueue.addEventListener('click', () => {
            queuePanel.classList.toggle('open');
        });
        btnCloseQueue.addEventListener('click', () => {
            queuePanel.classList.remove('open');
        });
        btnClearQueue.addEventListener('click', clearQueue);

        // Fullscreen Stage Mode
        btnFullscreenStage.addEventListener('click', () => {
            djStageOverlay.classList.add('active');
            if (document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
        });

        btnCloseStage.addEventListener('click', () => {
            djStageOverlay.classList.remove('active');
            if (document.exitFullscreen) {
                document.exitFullscreen().catch(() => {});
            }
        });

        // Shortcuts Modal
        btnToggleShortcuts.addEventListener('click', () => modalShortcuts.classList.add('open'));
        btnCloseModal.addEventListener('click', () => modalShortcuts.classList.remove('open'));
        modalShortcuts.addEventListener('click', (e) => {
            if (e.target === modalShortcuts) modalShortcuts.classList.remove('open');
        });

        // Master Buttons
        btnPlayMaster.addEventListener('click', togglePlay);
        stageBtnPlay.addEventListener('click', togglePlay);
        btnNext.addEventListener('click', nextSong);
        stageBtnNext.addEventListener('click', nextSong);
        btnPrev.addEventListener('click', prevSong);
        stageBtnPrev.addEventListener('click', prevSong);
        playerFavBtn.addEventListener('click', () => {
            if (state.currentSongIndex !== -1) {
                toggleFavorite(ALL_SONGS[state.currentSongIndex].id);
            }
        });

        // Global Keyboard Shortcuts
        window.addEventListener('keydown', (e) => {
            // Ignore when typing in search input
            if (e.target === searchInput) {
                if (e.key === 'Escape') {
                    searchInput.blur();
                }
                return;
            }

            switch (e.code) {
                case 'Space':
                    e.preventDefault();
                    togglePlay();
                    break;
                case 'ArrowRight':
                    e.preventDefault();
                    nextSong();
                    break;
                case 'ArrowLeft':
                    e.preventDefault();
                    prevSong();
                    break;
                case 'ArrowUp':
                    e.preventDefault();
                    state.volume = Math.min(1, state.volume + 0.05);
                    audio.volume = state.volume;
                    volumeSlider.value = state.volume;
                    showToast(`🔊 Volume ${Math.round(state.volume * 100)}%`);
                    break;
                case 'ArrowDown':
                    e.preventDefault();
                    state.volume = Math.max(0, state.volume - 0.05);
                    audio.volume = state.volume;
                    volumeSlider.value = state.volume;
                    showToast(`🔉 Volume ${Math.round(state.volume * 100)}%`);
                    break;
                case 'KeyM':
                    btnMute.click();
                    break;
                case 'KeyS':
                    toggleShuffleMode();
                    break;
                case 'KeyR':
                    cycleRepeatMode();
                    break;
                case 'KeyF':
                    if (djStageOverlay.classList.contains('active')) {
                        btnCloseStage.click();
                    } else {
                        btnFullscreenStage.click();
                    }
                    break;
                case 'Slash':
                    e.preventDefault();
                    searchInput.focus();
                    break;
                case 'Escape':
                    if (djStageOverlay.classList.contains('active')) btnCloseStage.click();
                    if (modalShortcuts.classList.contains('open')) modalShortcuts.classList.remove('open');
                    if (queuePanel.classList.contains('open')) queuePanel.classList.remove('open');
                    break;
            }
        });

        // Initial Boot
        audio.volume = state.volume;
        renderSoundboard();
        updateUI();
        renderQueue();

        console.log('👑 Ramesh Wedding DJ Console initialized successfully with', ALL_SONGS.length, 'offline tracks.');
    </script>
</body>
</html>
"""

final_html = html_template.replace("__SONGS_JSON__", songs_json)

# Write Wedding_Music_Player.html
player_path = os.path.join(BASE_DIR, "Wedding_Music_Player.html")
with open(player_path, "w", encoding="utf-8") as f:
    f.write(final_html)

# Write index.html
index_path = os.path.join(BASE_DIR, "index.html")
with open(index_path, "w", encoding="utf-8") as f:
    f.write(final_html)

# Create 1-Click Windows Batch Launchers
bat_path = os.path.join(BASE_DIR, "Start_Wedding_Player.bat")
bat_content = """@echo off
title Ramesh Wedding Music Player
echo ========================================================
echo   👑 RAMESH BROTHER'S WEDDING MUSIC PLAYER CONSOLE
echo   10, 11 & 12 December 2026 - 508 Offline Audio Tracks
echo ========================================================
echo.
echo Launching Wedding Music Player in default web browser...
start "" "%~dp0Wedding_Music_Player.html"
exit
"""
with open(bat_path, "w", encoding="utf-8") as f:
    f.write(bat_content)

# Create Local HTTP Server Launcher for zero-CORS / network streaming to tablets/phones
server_bat_path = os.path.join(BASE_DIR, "Start_Local_Server.bat")
server_bat_content = """@echo off
title Ramesh Wedding Music Server (Local DJ Station)
echo ========================================================
echo   👑 RAMESH WEDDING MUSIC LOCAL HTTP SERVER
echo   10, 11 & 12 December 2026 - 508 Offline Audio Tracks
echo ========================================================
echo.
echo Starting local DJ server on port 8080...
echo You can connect laptops, tablets, and phones on the same Wi-Fi!
echo.
start "" "http://localhost:8080/Wedding_Music_Player.html"
python -m http.server 8080 --directory "%~dp0"
pause
"""
with open(server_bat_path, "w", encoding="utf-8") as f:
    f.write(server_bat_content)

print(f"✅ Web Player successfully created at: {player_path}")
print(f"✅ Also saved as index.html at: {index_path}")
print(f"✅ Windows Launcher created at: {bat_path}")
print(f"✅ Local Server Launcher created at: {server_bat_path}")
