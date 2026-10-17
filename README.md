# PJP Visit Calendar — SURAJ

Monthly PJP (visit plan) calendar: har din ka place / remarks bharo, holidays mark karo,
place-wise days + counters dekho, aur ek click mein **Excel (.xlsx) report** download karo.

Is repo mein do versions hain:

| Version | Kahan chalta hai | Data kahan save hota hai | Excel export |
|---|---|---|---|
| **Phone app** (`index.html`) | Kisi bhi phone browser mein (GitHub Pages link) — offline bhi chalta hai | Usi phone ke browser mein | Seedha phone mein download |
| **Google Apps Script** (`apps-script/`) | Google Web App link | Google Sheet (`PJP_SURAJ_Data`) | Google Drive se download |

---

## 1) Phone app — GitHub Pages par (recommended, sabse aasaan)

1. GitHub par repo kholo → **Settings → Pages**.
2. *Build and deployment* → **Source: Deploy from a branch**.
3. Branch: `main` (ya jis branch par ye files hain), folder: **/ (root)** → **Save**.
4. 1–2 minute baad link milega: `https://rishu37-ux.github.io/PJP-Suraj/`
5. Phone par ye link kholo:
   - **Android (Chrome):** ⋮ menu → *Add to Home screen* / *Install app*
   - **iPhone (Safari):** Share → *Add to Home Screen*

   Ab ye normal app ki tarah home screen se khulega, aur internet ke bina bhi chalega.

**Excel export:** upar *Export Excel* dabao → `PJP_<Name>_<Month>_<Year>.xlsx` seedha phone ke
Downloads mein aa jayega (same formatting: title, holidays red, place-wise summary, totals).

**Backup zaroor lo:** Phone app ka data sirf usi phone/browser mein rehta hai. Neeche
*Backup Data* se `.json` file save kar lo; naye phone par *Restore Backup* se wapas le aao.
Browser ka data/cache clear karoge to data chala jayega — isliye backup rakho.

> Repo **private** hai to free GitHub plan par Pages nahi chalega — repo public karo, ya
> Apps Script version use karo.

## 2) Google Apps Script version (data Google Sheet mein)

1. <https://script.google.com> → **New project**.
2. `Code.gs` mein `apps-script/Code.gs` ka poora code paste karo.
3. **+ → HTML** file banao, naam `Index` rakho, usme `apps-script/Index.html` paste karo.
4. Function dropdown se `setupApp` select karke **Run** karo (permissions allow karo).
   Isse Drive mein `PJP_SURAJ_Data` sheet ban jayegi.
5. **Deploy → New deployment → Web app** — *Execute as: Me*, *Who has access: Only myself*
   (ya jo chahiye) → Deploy. Jo URL mile use phone par kholo aur home screen par add karo.

Code update karne ke baad: **Deploy → Manage deployments → Edit → Version: New version**.

## Files

```
index.html            Phone app (UI)
local-backend.js      Phone app ka data + Excel export logic
vendor/xlsx.min.js    xlsx-js-style (styled Excel banane ki library, Apache-2.0)
sw.js, manifest.webmanifest, icon*  Offline + "Add to Home screen" support
apps-script/Code.gs   Google Apps Script backend
apps-script/Index.html Google Apps Script UI
```

## Original files mein jo fixes kiye

- `render()` / `saveProfile()` mein `document.getElementById('name')` exist nahi karta tha → JS error. Fix kiya.
- `resetDatabaseSpreadsheet()` galat property key (`DATA_SPREADSHEET_ID`) delete kar raha tha → ab `PJP_DATA_SPREADSHEET_ID`.
- Excel footer mein naam hardcoded `SURAJ` tha → ab profile ka naam aata hai.
- Duplicate `ensureCoreSheets_`, `populateBulkPlaceDropdown`, `closeBulkVisit` hataye.
- *Fill Visit* aur *Mark Holiday* popups ko baaki popups jaisa design diya (pehle bina padding ke the).
- Holiday remove fail hone par error message dikhata hai.
