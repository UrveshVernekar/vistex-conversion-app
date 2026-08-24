# Vistex Agreement Format Converter Portal

A full-stack web application for converting Vistex Agreement format Excel/CSV files. Supports authentication, automatic format detection, and simultaneous `.xlsx` + `.csv` output downloads.

## Features

- **Format Conversion Engine**:
  - **Script 1**: Upload Format (Vertical Blocks) ➔ Flat Output Table (`.xlsx` + `.csv`)
  - **Script 2**: Flat Output Table ➔ Reverse Upload Format (`.xlsx` + `.csv`)
- **Auto-Detection**: Automatically determines input file format (Excel or CSV).
- **Authentication**: Username & Password authentication portal.
- **Simultaneous Dual Download**: Downloads `.zip` archive containing both `.xlsx` and `.csv` files automatically.
- **Network Ready**: Easily hosted on internal server (`0.0.0.0:5000`).

---

## Installation & Setup

1. **Clone Repository**:
   ```bash
   git clone https://github.com/UrveshVernekar/vistex-conversion-app.git
   cd vistex-conversion-app
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

---

## Running the Server

Start the web server:
```bash
python run_server.py
```

- **Local Access**: `http://localhost:5000`
- **Network Access**: `http://<your-server-ip>:5000`

### Default Login Credentials
- **Username**: `admin`
- **Password**: `vistex2026`

---

## Repository Structure

```text
├── app.py              # FastAPI web server & route handlers
├── converter.py        # Conversion wrapper & auto-detection logic
├── formats.py          # Header schemas & format rules
├── script1.py          # Script 1 conversion logic (Upload -> Flat Output)
├── script2.py          # Script 2 conversion logic (Flat Output -> Reverse Upload)
├── run_server.py       # Server launcher (0.0.0.0:5000)
├── requirements.txt    # Python package dependencies
├── templates/          # Jinja2 HTML templates
│   ├── login.html
│   └── dashboard.html
└── static/             # CSS styles & visual assets
    └── style.css
```
