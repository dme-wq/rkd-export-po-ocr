# RKD PO OCR Extraction System 🚀

An AI-powered Purchase Order data extraction web application built for **RKD Furnishings**. It extracts purchase order header details and line item items from multi-page PDFs, images, and Excel spreadsheets using Gemini AI, saves data to Google Sheets, and sends missing field requests via WhatsApp.

---

## 🌟 Key Features
- **Modern Glassmorphic Design**: Clean UI with Dark & Light mode toggle.
- **Multi-File Drag & Drop Upload**: Upload multiple PDFs, Images, or Excel files at once.
- **AI OCR Extraction**: Automatically extracts PO numbers, buyer/vendor details, line items, and quantities using Google Gemini 2.0.
- **Interactive Editing**: Add/Remove line items with real-time total price calculations.
- **WhatsApp Missing Info Alerts**: 1-click notification to request missing PO details.
- **Tabulator Database View**: Searchable, filterable history of all saved PO entries with CSV export.
- **Vercel & GitHub Ready**: Universal Hybrid API bridge for hosting on Vercel or running directly inside Google Apps Script.

---

## 📋 Setup & Deployment Guide

### Step 1: Deploy Google Apps Script Web App
1. Open [Google Apps Script](https://script.google.com/).
2. Copy the contents of `Code.gs` and `Index.html` into your Apps Script editor.
3. Click **Deploy** -> **New deployment**.
4. Select Type: **Web app**.
5. Set:
   - **Execute as**: *Me*
   - **Who has access**: *Anyone*
6. Click **Deploy** and copy your **Web App URL** (looks like `https://script.google.com/macros/s/.../exec`).

---

### Step 2: Push Project to GitHub
Open your terminal in this directory and run:

```bash
git init
git add .
git commit -m "Initial commit - Modern WebApp with Vercel support"
git branch -M main
git remote add origin YOUR_GITHUB_REPOSITORY_URL
git push -u origin main
```

---

### Step 3: Deploy to Vercel
1. Log in to [Vercel](https://vercel.com).
2. Click **Add New** -> **Project**.
3. Import your GitHub repository (`rkd-po-ocr-system`).
4. Keep framework preset as **Other** / **Static Site**.
5. Click **Deploy**.
6. Once deployed, open your Vercel URL.
7. Click the ⚙️ **Sliders / Settings icon** in the top navigation bar and paste your **Google Apps Script Web App URL**.

---

## 💻 Tech Stack
- **Frontend**: HTML5, Vanilla CSS3 (Glassmorphic Design System), Vanilla JavaScript (ES6+), FontAwesome 6, SweetAlert2, Tabulator.
- **Backend API**: Google Apps Script (v8 engine), Google Drive API, Google Sheets API, Google Gemini AI API, Maytapi WhatsApp API.
- **Deployment**: Vercel & GitHub Pages.
