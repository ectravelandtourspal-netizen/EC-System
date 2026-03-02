# Booking Management Web App

This is a static multi-page frontend connected to a Google Apps Script backend.

## Public Deployment (GitHub Pages)

### 1) Verify backend URL
Open `script.js` and make sure `API_URL` is your latest deployed Apps Script Web App URL (`.../exec`).

Current format:

```js
const API_URL = "https://script.google.com/macros/s/.../exec";
```

### 2) Deploy Apps Script for public access
In Google Apps Script:
1. **Deploy** -> **Manage deployments** -> **Edit** (or **New deployment**)
2. Type: **Web app**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Deploy and copy the `/exec` URL
6. Paste that URL into `script.js` as `API_URL`

### 3) Push this project to GitHub
From this project folder, run:

```bash
git init
git add .
git commit -m "Initial public release"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

### 4) Enable GitHub Pages
In your GitHub repository:
1. Go to **Settings** -> **Pages**
2. Source: **Deploy from a branch**
3. Branch: **main**
4. Folder: **/(root)**
5. Save

Your public site URL will be:

`https://<your-username>.github.io/<your-repo>/`

## Project Files

- `index.html` - Home / quick actions
- `new-bookings.html` - New booking WhatsApp send flow
- `confirmed-whatsapp.html` - Confirmed guest WhatsApp send flow
- `bookings-dashboard.html` - Dashboard view
- `payment-receipt-encoding.html` - Receipt encoding view
- `commissions-payment-receipt-encoding.html` - Bulk commissions receipt encoding (Travel Date + Coupon + SUM AA)
- `style.css` - Shared styles
- `script.js` - Shared frontend logic and API calls
- `google-apps-script.gs` - Backend for Google Apps Script

## Notes

- If updates do not appear, hard refresh your browser (`Ctrl+F5`).
- If API calls fail publicly, redeploy Apps Script and confirm access is set to **Anyone**.
