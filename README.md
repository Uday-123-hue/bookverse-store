# BookVerse — Multi-Category Online Bookstore & Live Business Intelligence Platform

**Tagline:** *Discover Stories. Explore Knowledge. Understand Readers.*

BookVerse is a full-stack Python Flask e-commerce bookstore paired with a live Pandas + OpenPyXL + Plotly Business Intelligence (BI) dashboard, server-side session cart, Verified Reader review system, automated 8-sheet Excel reporting, and optional Firebase Realtime Database REST synchronization.

## Tech Stack
- **Backend:** Python, Flask, Flask-WTF (CSRF), Werkzeug (password hashing)
- **Data Engineering & BI:** Pandas, OpenPyXL, Plotly
- **Frontend:** HTML5, Jinja2, Tailwind CSS CDN (zero custom JavaScript files)
- **Persistence:** JSON, CSV backups, 8-sheet Excel workbook (`orders.xlsx`), Firebase Realtime Database REST API (optional, graceful local fallback)
- **Deployment:** GitHub + Render + Gunicorn

## Local Setup
1. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   .\venv\Scripts\Activate.ps1   # Windows PowerShell