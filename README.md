# Exium MUPS - Sr. / MIO Award Catalogue Choice Portal

Standalone, production-grade Award Choice Web Portal for Sr. / MIO Award Achievers across Bangladesh to select their preferred gift voucher from 6 authorized retail brands.

🌐 **Live Portal Link:** **[https://didarul2911-netizen.github.io/exium-award-campaign-2026/](https://didarul2911-netizen.github.io/exium-award-campaign-2026/)**

---

## 📂 Project Structure

```
G:\Exium\2026\Award\May-Jun\
│
├── index.html                               # Production Web Portal (Single-file, self-contained)
├── Award_Choice_Portal.html                 # Direct duplicate for backup/alternative hosting
│
├── Exium MUPS Logo.png                      # Official Header Logo
├── Brand Logo\                              # Official Brand Logos directory
│   ├── Aarong Logo.png
│   ├── Apex Logo.png
│   ├── Bata logo.jpg
│   ├── Best Buy Logo.png
│   ├── Infinity Logo.png
│   └── Cats Eye Logo.jpeg
│
├── FF list.xlsx                             # Official Field Force Master Hierarchy (1,856 territories)
├── Exium_Award_Choice_Master_2026.xlsx      # Master Excel file mapped with current FF list & transfers
├── Exium Award Achiever List_May 2026.xlsx  # Raw source file for May 2026 (721 Achievers)
├── Exium Award Achiever List_June 2026.xlsx # Raw source file for June 2026 (820 Achievers)
│
├── build_award_portal.py                    # Modular Ingestion Engine (Updates portal with new months)
├── generate_master_award_excel.py           # Master Excel generator & formatter with FF mapping
├── sync_excel.py                            # Updates Master Excel from exported JSON / live sync
├── Google_Apps_Script_Award_Backend.js      # Backend script for Google Sheets sync
└── README.md                                # System documentation and deployment guide
```

---

## 🌟 Key Features & Updates

### 1. Current FF List Integration & Transfer Tracking
- Every Award Achiever is mapped to their **CURRENT** Zone, Region, Regional Head, and Territory as defined in `FF list.xlsx`.
- **Transferred MIO Detection:**
  - If an MIO was transferred (their territory or region at the time of the award differs from their current FF posting), the card displays an eye-catching transfer note:
    > **🔄 Transferred MIO:** Award achieved while in: **[Previous Area] ([Previous Code])** • Region: **[Previous Region]** (Regional Head: **[Previous RH]**) • Zone: **[Previous Zone]**

### 2. Strict Role-Based Authentication & Access
- **🏢 Regional Head Login:**
  - Dropdown displays: `Region Name (SAP Region Code)` (e.g. `Bogura GEN-A-1 (16020)`).
  - Shows `Regional Head` and `Eligible Award Achievers count` upon selecting region.
  - Password: **ONLY** the `SAP Region Code` is accepted.
- **🌐 Zonal Head Login:**
  - Dropdown displays: `Zone Name - Zonal Head (SAP Zone Code)`.
  - Password: **ONLY** the `SAP Zone Code` (e.g. `17005`) is accepted.
  - Opens the **Zonal Dashboard**:
    - Zone KPI Summary (Total Achievers, Completed, Pending, % Complete).
    - Region-by-Region breakdown table under that Zone with direct links to view cards.
    - Voucher Brand popularity breakdown within that Zone.
- **🛡️ Admin Panel:**
  - Password: **`Exium MUPS`**.
  - Opens National Dashboard with full KPI stats, voucher purchase order distribution, CSV download, and Google Apps Script configuration.
- **🔍 Direct MIO Code Search:**
  - Quick lookup by SAP MIO Code, Area Code, or Name to directly open the achiever's card.

### 3. 6 Authorized Brand Voucher Options
Each achiever selects **1 voucher option** per award won:
1. 🛍️ **Infinity Gift Voucher** (Infinity Mega Mall - Modern lifestyle & fashion)
2. 👞 **Apex Gift Voucher** (Apex Footwear - Premium leather & footwear)
3. 👟 **Bata Gift Voucher** (Bata Bangladesh - Footwear & accessories)
4. 🥻 **Aarong Gift Voucher** (Aarong / BRAC - Ethnic heritage & lifestyle)
5. 🏠 **Best Buy Gift Voucher** (RFL Best Buy - Household & electronics retail)
6. 🐱 **Cats Eye Gift Voucher** (Cats Eye - Contemporary smart casuals & fashion)

All 6 official logos and the Exium MUPS header logo are **base64-embedded directly into the HTML**, ensuring high visual fidelity with zero broken images even when offline.

---

## ☁️ Google Apps Script Backend Setup

1. Create a Google Sheet named **`Exium_Award_Choice_Master_2026`** with two tabs: **`Award_May_2026`** and **`Award_June_2026`**.
2. Copy rows from `Exium_Award_Choice_Master_2026.xlsx` into these tabs.
3. Open **Extensions > Apps Script** and paste code from [Google_Apps_Script_Award_Backend.js](file:///g:/Exium/2026/Award/May-Jun/Google_Apps_Script_Award_Backend.js).
4. Click **Deploy > New Deployment**:
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the Web App URL and paste it in the portal's **Admin Panel > Endpoint URL**.

---

## 🚀 Adding September 2026 and Future Months

1. Place the new month's Excel file in `G:\Exium\2026\Award\May-Jun\`.
2. Add the month block in `build_award_portal.py`.
3. Run:
   ```bash
   python build_award_portal.py
   ```
4. The Web Portal will automatically add the new month's tab and all its achiever cards mapped to their current FF positions!
