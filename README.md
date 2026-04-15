# Krutidev Studio

A full-stack Python web app for converting and editing Krutidev-encoded Hindi Excel files — with a real-time collaborative spreadsheet editor and a free online text converter.

**Live:** [krutidevstudio.com](https://krutidevstudio.com)

---

## Features

- **Krutidev ↔ Unicode text converter** — real-time, browser-side, no signup
- **Excel batch converter** — upload `.xlsx`/`.xls`, convert all Krutidev cells to Unicode (or reverse), download with full formatting preserved (borders, merged cells, column widths, row heights)
- **Online spreadsheet editor** — Handsontable-powered grid with Krutidev font rendering, cell formatting (bold, italic, underline, bg color, font size, alignment), merge cells, search & replace
- **Shareable links** — generate `/sheet/<uuid>` links with view/edit access control and optional password protection
- **Real-time collaboration** — multiple users edit simultaneously via WebSocket (Socket.IO) with live cursor tracking
- **Font download** — serves KrutiDev 010 `.ttf` directly from `/font/download`
- **SEO-ready** — sitemap, robots.txt, structured data (JSON-LD), Open Graph, hreflang

---

## Tech Stack

| Layer | Tech |
|---|---|
| Backend | Python 3.11 · Flask 3.1 · Flask-SocketIO 5.6 |
| WSGI server | Gunicorn + eventlet worker (production) |
| Database | SQLite with WAL mode (`data/sheets.db`) |
| Excel | openpyxl 3.1 |
| Frontend | Handsontable (non-commercial) · Socket.IO client 4.7 |
| Font | KrutiDev 010 TTF |
| Container | Docker + Docker Compose |

---

## Quick Start (Local Dev)

```bash
git clone https://github.com/rajatpoddar/krutidev.git
cd krutidev

python3 -m venv venv
source venv/bin/activate        # macOS/Linux
# venv\Scripts\activate         # Windows

pip install -r requirements.txt
python app.py
```

Open **http://localhost:8765**

> The app falls back to `threading` async mode when `eventlet` is not installed in the venv — this is fine for local development.

---

## Deploy with Docker (Production)

**1. Create your `.env` file:**

```bash
cp .env.example .env
# Open .env and set a strong SECRET_KEY (32+ random chars)
```

**2. Build and start:**

```bash
docker compose up -d --build
```

**3. Check status:**

```bash
docker compose ps
docker compose logs -f krutidev
```

**4. Other useful commands:**

```bash
docker compose restart           # restart without rebuild
docker compose up -d --build     # rebuild after code changes
docker compose down              # stop and remove containers
```

**Backup the database:**

```bash
docker compose exec krutidev cp /app/data/sheets.db /app/data/sheets.db.bak
```

The app runs on port **8765**. Put Nginx in front for HTTPS — make sure to pass WebSocket upgrade headers:

```nginx
location / {
    proxy_pass         http://127.0.0.1:8765;
    proxy_http_version 1.1;
    proxy_set_header   Upgrade $http_upgrade;
    proxy_set_header   Connection "upgrade";
    proxy_set_header   Host $host;
    proxy_set_header   X-Real-IP $remote_addr;
}
```

---

## Font Setup

The KrutiDev 010 font is included at `static/fonts/KrutiDev010.ttf` and served via `/font/download`. If you need to replace it, drop the `.ttf` file at:

```
static/fonts/KrutiDev010.ttf
```

---

## Project Structure

```
krutidev/
├── app.py              # Flask routes + SocketIO handlers
├── converter.py        # Krutidev ↔ Unicode conversion engine
├── db.py               # SQLite persistence (CRUD for shared docs)
├── requirements.txt    # Pinned Python dependencies
├── Dockerfile
├── docker-compose.yml
├── .env.example
├── test_converter.py   # Converter unit tests (13 cases)
├── templates/
│   ├── index.html      # Homepage + text converter + Excel tools
│   ├── editor.html     # Spreadsheet editor
│   └── 404.html
├── static/
│   ├── style.css       # Homepage styles
│   ├── editor.css      # Editor styles
│   ├── editor.js       # Editor logic (Handsontable, Socket.IO, save/share)
│   └── fonts/
│       └── KrutiDev010.ttf
├── data/
│   └── sheets.db       # SQLite database (auto-created)
└── uploads/            # Temporary Excel files (auto-cleaned after download)
```

---

## API Routes

| Method | Route | Description |
|---|---|---|
| `GET` | `/` | Homepage — text converter + Excel tools |
| `GET` | `/editor` | New blank spreadsheet editor |
| `GET` | `/sheet/<id>` | Open a shared document |
| `GET` | `/font/download` | Download KrutiDev 010 TTF |
| `GET` | `/sitemap.xml` | XML sitemap |
| `GET` | `/robots.txt` | robots.txt |
| `POST` | `/upload` | Upload Excel → start async conversion job |
| `GET` | `/progress/<job_id>` | Poll conversion job progress |
| `GET` | `/download/<job_id>` | Download converted file (auto-deletes after) |
| `POST` | `/api/convert-text` | Convert text Krutidev ↔ Unicode |
| `POST` | `/api/load` | Parse Excel file → JSON sheets data |
| `POST` | `/api/save` | JSON sheets data → Excel download |
| `POST` | `/api/share` | Create or update a shared document |
| `GET` | `/api/doc/<id>` | Load shared document data |
| `DELETE` | `/api/doc/<id>` | Delete a shared document |
| `GET` | `/api/docs` | List all documents (last 50) |
| `GET` | `/api/doc/<id>/download` | Download shared document as Excel |

---

## WebSocket Events (Socket.IO)

| Event | Direction | Description |
|---|---|---|
| `join` | client → server | Join a document room |
| `users_update` | server → client | Broadcast active users list |
| `cell_change` | bidirectional | Sync a cell edit to all room members |
| `cursor_move` | bidirectional | Broadcast cursor position |
| `title_change` | bidirectional | Sync document title |

---

## Running Tests

```bash
python test_converter.py
```

Expected output: `13/13 passed`

---

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `SECRET_KEY` | `kd-editor-secret-2024` | Flask session secret — **change in production** |
| `PORT` | `8765` | Port the app listens on (local dev only; Docker uses gunicorn) |

---

## Notes

- **1 gunicorn worker** is intentional — Socket.IO with eventlet requires a single worker process for shared in-memory state.
- **SQLite WAL mode** is enabled for better concurrent read performance.
- Uploaded files are automatically deleted from `uploads/` after the converted file is downloaded.
- The `_source_files` in-memory store is capped at 50 entries to prevent unbounded memory growth.
