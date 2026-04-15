# Krutidev Studio

A full-stack Python web app for working with Krutidev-encoded Hindi Excel files.

## Features

- **Krutidev → Unicode converter** — batch convert `.xlsx`/`.xls` files
- **Spreadsheet editor** — Google Sheets-style grid with Krutidev font rendering
- **Shareable links** — generate `/sheet/<uuid>` links with view/edit access control and optional password
- **Real-time collaboration** — multiple users edit simultaneously via WebSocket (Socket.IO)
- **Dashboard** — SaaS-style UI with sidebar, recent sheets, dark mode

## Tech Stack

| Layer | Tech |
|---|---|
| Backend | Python · Flask · Flask-SocketIO |
| Database | SQLite (via `db.py`) |
| Excel | openpyxl · pandas |
| Frontend | Handsontable · Socket.IO client |
| Font | KrutiDev 010 (user-supplied `.ttf`) |

## Setup

```bash
git clone https://github.com/rajatpoddar/krutidev.git
cd krutidev

python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux

pip install -r requirements.txt
python app.py
```

Open **http://localhost:5000**

## Font Setup

The Krutidev font is not included (licensing). Place `KrutiDev010.ttf` in:

```
static/fonts/KrutiDev010.ttf
```

Download from [indiatyping.com](https://www.indiatyping.com/index.php/download/krutidev-font) or extract from `C:\Windows\Fonts\` on a Windows system with Kruti Dev installed.

## Routes

| Route | Description |
|---|---|
| `GET /` | Dashboard |
| `GET /editor` | New blank editor |
| `GET /sheet/<id>` | Open shared document |
| `GET /convert` | Batch Krutidev → Unicode converter |
| `POST /api/load` | Parse Excel → JSON |
| `POST /api/save` | JSON → Excel download |
| `POST /api/share` | Create/update shared document |
| `GET /api/doc/<id>` | Load document data |
| `GET /api/docs` | List all documents |
| `GET /api/doc/<id>/download` | Download document as Excel |

## Project Structure

```
krutidev/
├── app.py              # Flask + SocketIO routes
├── converter.py        # Krutidev → Unicode engine
├── db.py               # SQLite persistence
├── requirements.txt
├── templates/
│   ├── dashboard.html
│   ├── editor.html
│   ├── index.html
│   └── 404.html
└── static/
    ├── dashboard.css / dashboard.js
    ├── editor.css / editor.js
    ├── style.css / script.js
    └── fonts/          # Place KrutiDev010.ttf here
```
