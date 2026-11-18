import os
import sys
import json
import re
import socket
import threading
import time
import urllib.parse
import urllib.request
from http.server import SimpleHTTPRequestHandler
from socketserver import ThreadingTCPServer

PORT = 8080
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Global DJ in-memory state
dj_state = {
    "current_song_id": 1,
    "is_playing": False,
    "current_time": 0,
    "duration": 0,
    "volume": 1.0,
    "is_ducked": False,
    "bass_boost": False,
    "repeat_mode": "off",
    "is_shuffle": False,
    "queue": [],
    "last_update": time.time()
}

# Crowd requests
crowd_requests = []
request_id_counter = 1

# YouTube Search Cache
yt_search_cache = {}

def search_youtube_first_id(query):
    """Searches YouTube and returns the first matching video ID in real-time."""
    q_norm = query.strip().lower()
    if q_norm in yt_search_cache:
        return yt_search_cache[q_norm]

    search_term = query if any(w in q_norm for w in ['song', 'music', 'official', 'remix', 'video']) else f"{query} official song"
    url = "https://www.youtube.com/results?search_query=" + urllib.parse.quote(search_term)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    try:
        with urllib.request.urlopen(req, timeout=6) as response:
            content = response.read().decode("utf-8", errors="ignore")
            video_ids = re.findall(r'"videoId":"([a-zA-Z0-9_-]{11})"', content)
            if video_ids:
                vid = video_ids[0]
                yt_search_cache[q_norm] = vid
                return vid
    except Exception as e:
        sys.stderr.write(f"YouTube search error for {query}: {e}\\n")
    return None

# Active SSE client queues
sse_clients = set()
sse_lock = threading.Lock()

def broadcast_event(event_type, data):
    """Broadcasts a JSON event to all connected SSE clients."""
    payload = f"event: {event_type}\ndata: {json.dumps(data)}\n\n"
    with sse_lock:
        dead_clients = set()
        for q in sse_clients:
            try:
                q.put_nowait(payload)
            except Exception:
                dead_clients.add(q)
        for dead in dead_clients:
            sse_clients.discard(dead)

def get_local_ip():
    """Detects active local Wi-Fi or LAN IPv4 address."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # Doesn't need to be reachable, just triggers routing table
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        try:
            ip = socket.gethostbyname(socket.gethostname())
        except Exception:
            ip = '127.0.0.1'
    finally:
        s.close()
    return ip

class WeddingDJRequestHandler(SimpleHTTPRequestHandler):
    """
    Advanced HTTP Request Handler for Wedding DJ Console:
    - Full HTTP 206 Partial Content (Range requests) for flawless audio scrubbing on iOS Safari and Chrome.
    - Server-Sent Events (SSE) stream at /api/events for real-time mobile sync.
    - REST API for mobile remote control and guest requests.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def log_message(self, format, *args):
        # Suppress spammy asset logging, log API and errors only
        path = getattr(self, 'path', '')
        if '/api/' in path:
            sys.stderr.write(f"[{time.strftime('%H:%M:%S')}] API {format % args}\n")
        elif not any(path.endswith(ext) for ext in ['.webm', '.m4a', '.png', '.jpg', '.css', '.js']):
            sys.stderr.write(f"[{time.strftime('%H:%M:%S')}] {format % args}\n")

    def end_headers(self):
        # Allow CORS and caching headers
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, DELETE')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Range')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # 1. API: Server-Sent Events (SSE)
        if path == '/api/events':
            self.handle_sse()
            return

        # 2. API: Current DJ State
        if path == '/api/state':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(json.dumps(dj_state).encode('utf-8'))
            return

        # 3. API: List Crowd Requests
        if path == '/api/requests':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(json.dumps({"requests": crowd_requests}).encode('utf-8'))
            return

        # 4. API: Local IP info
        if path == '/api/ip':
            local_ip = get_local_ip()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                "local_ip": local_ip,
                "port": PORT,
                "laptop_url": f"http://localhost:{PORT}",
                "mobile_url": f"http://{local_ip}:{PORT}"
            }).encode('utf-8'))
            return

        # Favicon handler
        if path == '/favicon.ico':
            self.send_response(200)
            self.send_header('Content-Type', 'image/svg+xml')
            self.send_header('Cache-Control', 'max-age=86400')
            self.end_headers()
            self.wfile.write("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>👑</text></svg>".encode('utf-8'))
            return

        # 5. API: YouTube Instant Search
        if path == '/api/search_yt':
            query_params = urllib.parse.parse_qs(parsed.query)
            q = query_params.get('q', [''])[0].strip()
            if not q:
                self.send_error(400, "Missing query")
                return

            vid = search_youtube_first_id(q)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'max-age=3600')
            self.end_headers()
            if vid:
                self.wfile.write(json.dumps({"status": "ok", "videoId": vid, "query": q}).encode('utf-8'))
            else:
                self.wfile.write(json.dumps({"status": "not_found", "query": q}).encode('utf-8'))
            return

        # 5. Default static files with HTTP Range Request Support
        self.serve_file_with_range()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        content_len = int(self.headers.get('Content-Length', 0))
        body_bytes = self.rfile.read(content_len) if content_len > 0 else b'{}'
        
        try:
            data = json.loads(body_bytes.decode('utf-8')) if body_bytes else {}
        except Exception:
            data = {}

        # 1. Control Action (from Pocket Remote or Master Console)
        if path == '/api/control':
            action = data.get('action')
            if not action:
                self.send_error(400, "Missing action")
                return

            # Apply state changes if applicable
            if action in ['play', 'pause', 'toggle', 'next', 'prev', 'seek', 'jump', 'volume', 'duck', 'fx', 'play_song', 'queue_add']:
                if action == 'play': dj_state['is_playing'] = True
                elif action == 'pause': dj_state['is_playing'] = False
                elif action == 'toggle': dj_state['is_playing'] = not dj_state['is_playing']
                elif action == 'volume': dj_state['volume'] = float(data.get('value', 1.0))
                elif action == 'duck': dj_state['is_ducked'] = bool(data.get('value', False))
                elif action == 'bass_boost': dj_state['bass_boost'] = bool(data.get('value', False))
                elif action == 'play_song':
                    dj_state['current_song_id'] = int(data.get('song_id', 1))
                    dj_state['is_playing'] = True
                    dj_state['current_time'] = 0
                elif action == 'queue_add':
                    dj_state['queue'].append(int(data.get('song_id')))

                dj_state['last_update'] = time.time()
                # Broadcast command to all clients
                broadcast_event('control', {"action": action, "payload": data, "state": dj_state})

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok", "action": action, "state": dj_state}).encode('utf-8'))
            return

        # 2. Master DJ State Sync (Master laptop broadcasts its active state)
        if path == '/api/sync_state':
            dj_state.update(data)
            dj_state['last_update'] = time.time()
            broadcast_event('state_sync', dj_state)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
            return

        # 3. Crowd Request Submission (From Guest UI)
        if path == '/api/requests':
            global request_id_counter
            song_id = data.get('song_id')
            guest_name = data.get('guest_name', 'Wedding Guest').strip() or 'Wedding Guest'
            note = data.get('note', '').strip()
            song_title = data.get('song_title', f'Song #{song_id}')
            cero = data.get('ceremony', '')

            req_obj = {
                "id": request_id_counter,
                "song_id": song_id,
                "song_title": song_title,
                "guest_name": guest_name,
                "note": note,
                "ceremony": cero,
                "time": time.strftime('%H:%M'),
                "status": "pending"
            }
            request_id_counter += 1
            crowd_requests.append(req_obj)

            # Limit queue to last 50 requests
            if len(crowd_requests) > 50:
                crowd_requests.pop(0)

            # Broadcast new request alert to DJ console
            broadcast_event('new_request', req_obj)

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok", "request": req_obj}).encode('utf-8'))
            return

        self.send_error(404, "Endpoint not found")

    def do_DELETE(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path.startswith('/api/requests/'):
            try:
                req_id = int(path.split('/')[-1])
                global crowd_requests
                crowd_requests = [r for r in crowd_requests if r['id'] != req_id]
                broadcast_event('request_removed', {"id": req_id})
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
                return
            except Exception as e:
                self.send_error(400, str(e))
                return

        self.send_error(404, "Not found")

    def handle_sse(self):
        """Streams Server-Sent Events to keep mobile remotes & DJ console in lockstep."""
        import queue
        q = queue.Queue(maxsize=100)
        with sse_lock:
            sse_clients.add(q)

        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('Connection', 'keep-alive')
        self.end_headers()

        # Send initial state
        init_payload = f"event: state_sync\ndata: {json.dumps(dj_state)}\n\n"
        try:
            self.wfile.write(init_payload.encode('utf-8'))
            self.wfile.flush()
        except Exception:
            with sse_lock:
                sse_clients.discard(q)
            return

        try:
            while True:
                try:
                    msg = q.get(timeout=20.0)
                    self.wfile.write(msg.encode('utf-8'))
                    self.wfile.flush()
                except queue.Empty:
                    # Heartbeat comment to keep connection alive
                    self.wfile.write(b": heartbeat\n\n")
                    self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, Exception):
            pass
        finally:
            with sse_lock:
                sse_clients.discard(q)

    def serve_file_with_range(self):
        """
        Custom file server that supports HTTP 206 Partial Content (Byte Ranges).
        Crucial for instant scrubbing, audio seek, and mobile Safari/Chrome playback.
        """
        # Resolve path
        path = self.translate_path(self.path)
        if os.path.isdir(path):
            parts = urllib.parse.urlsplit(self.path)
            if not parts.path.endswith('/'):
                # redirect browser - do not treat as error
                self.send_response(301)
                new_parts = (parts[0], parts[1], parts[2] + '/', parts[3], parts[4])
                new_url = urllib.parse.urlunsplit(new_parts)
                self.send_header("Location", new_url)
                self.end_headers()
                return
            for index in "index.html", "index.htm":
                index_path = os.path.join(path, index)
                if os.path.exists(index_path):
                    path = index_path
                    break
            else:
                super().do_GET()
                return

        if not os.path.exists(path) or not os.path.isfile(path):
            self.send_error(404, "File not found")
            return

        ctype = self.guess_type(path)
        if path.endswith('.webm'): ctype = 'audio/webm'
        elif path.endswith('.m4a'): ctype = 'audio/mp4'
        elif path.endswith('.m3u'): ctype = 'audio/x-mpegurl'
        elif path.endswith('.json'): ctype = 'application/json'
        elif path.endswith('.webmanifest'): ctype = 'application/manifest+json'

        fsize = os.path.getsize(path)
        range_header = self.headers.get('Range')

        if not range_header:
            # Standard 200 OK delivery
            try:
                f = open(path, 'rb')
            except OSError:
                self.send_error(404, "File not accessible")
                return
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(fsize))
            self.send_header("Accept-Ranges", "bytes")
            self.send_header("Last-Modified", self.date_time_string(os.path.getmtime(path)))
            self.end_headers()
            self.copy_file_chunks(f, self.wfile)
            f.close()
            return

        # Process Range: bytes=START-END
        match = re.match(r'^bytes=(\d*)-(\d*)$', range_header.strip())
        if not match:
            self.send_error(416, "Requested Range Not Satisfiable")
            return

        start_str, end_str = match.groups()
        if start_str:
            start = int(start_str)
            end = int(end_str) if end_str else fsize - 1
        elif end_str:
            start = fsize - int(end_str)
            end = fsize - 1
        else:
            start = 0
            end = fsize - 1

        if start >= fsize or end >= fsize or start > end:
            self.send_response(416)
            self.send_header("Content-Range", f"bytes */{fsize}")
            self.end_headers()
            return

        length = end - start + 1
        try:
            f = open(path, 'rb')
            f.seek(start)
        except OSError:
            self.send_error(404, "File not accessible")
            return

        self.send_response(206)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Range", f"bytes {start}-{end}/{fsize}")
        self.send_header("Content-Length", str(length))
        self.send_header("Accept-Ranges", "bytes")
        self.send_header("Last-Modified", self.date_time_string(os.path.getmtime(path)))
        self.end_headers()

        self.copy_file_chunks(f, self.wfile, length)
        f.close()

    def copy_file_chunks(self, source, output, max_bytes=None):
        """Copies file in 64KB chunks up to max_bytes."""
        bufsize = 64 * 1024
        bytes_sent = 0
        try:
            while True:
                to_read = bufsize
                if max_bytes is not None:
                    remaining = max_bytes - bytes_sent
                    if remaining <= 0: break
                    to_read = min(bufsize, remaining)
                chunk = source.read(to_read)
                if not chunk: break
                output.write(chunk)
                bytes_sent += len(chunk)
        except (BrokenPipeError, ConnectionResetError):
            pass

def run_server():
    local_ip = get_local_ip()
    server_address = ('0.0.0.0', PORT)
    ThreadingTCPServer.allow_reuse_address = True
    httpd = ThreadingTCPServer(server_address, WeddingDJRequestHandler)

    print("=" * 72)
    print("      👑 RAMESH BROTHER WEDDING DJ PRO CONSOLE - WORLD-CLASS SERVER")
    print("                     10, 11 & 12 DECEMBER 2026")
    print("=" * 72)
    print(f"\n  [✔] Range Requests (206 Partial Content): ENABLED (Buttery Smooth Seek)")
    print(f"  [✔] Server-Sent Events (SSE Stream):       ENABLED (Real-time Remote Sync)")
    print(f"  [✔] Live Crowd Request Kiosk API:          ENABLED (/api/requests)")
    print("-" * 72)
    print(f"  💻 LAPTOP / STAGE MASTER:   http://localhost:{PORT}")
    print(f"  📱 MOBILE POCKET REMOTE:   http://{local_ip}:{PORT}")
    print(f"  🎉 CROWD REQUEST KIOSK:    http://{local_ip}:{PORT}?mode=crowd")
    print("-" * 72)
    print("  Press Ctrl+C to stop the server anytime.\n")

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down Wedding DJ Server gracefully...")
        httpd.server_close()

if __name__ == '__main__':
    run_server()
