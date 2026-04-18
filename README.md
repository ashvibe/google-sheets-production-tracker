# 🏭 Arora Production Tracker

A Google Sheets + Apps Script production management system built for **Arora Bisiklet**. Tracks orders, chassis deliveries, production orders (ARP), and line output — all from a single spreadsheet with sidebar forms and a live dashboard.

> 🇹🇷 Türkçe dokümantasyon için: [README.tr.md](./README.tr.md)

---

## ✨ Features

- **Order Management** — Multi-line orders per order number, auto-incrementing IDs
- **Chassis Tracking** — Per-supplier tracking with delivery log (each shipment on its own row)
- **ARP (Production Orders)** — Assign models to lines, track planned vs. actual output
- **Line Tracking** — Daily production entry per line (A/B/C/D)
- **Dashboard** — KPI cards, weekly line production table (day-by-day), order progress bars
- **Sidebar Forms** — User-friendly forms instead of raw cell editing
- **Backup & Archive** — One-click Google Drive backup; archive line data without losing history
- **Admin Reset** — Password-protected system reset (auto-backup before wipe)
- **Settings Page** — Add models, suppliers, and lines without touching code

---

## 🗂️ Sheet Structure

| Sheet | Purpose | Primary User |
|-------|---------|--------------|
| 📊 DASHBOARD | KPI overview, weekly line chart, order progress | Management |
| 📋 ORDERS | All orders with status and chassis info | Management |
| 🔩 CHASSIS TRACKING | Supplier-based chassis delivery log | Procurement |
| ⚙️ PRODUCTION ORDERS (ARP) | Line assignments and output tracking | Production Manager |
| 📈 LINE TRACKING | Daily production entries per line | Line Operators |
| ⚙️ SETTINGS | Model, supplier, and line lists | Admin |

---

## 🚀 Setup

### Requirements
- A Google account
- Google Sheets (free)

### Installation

1. Open a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete the default code
4. Paste the contents of `en/ProductionTracker.gs`
5. Click **Save** (💾)
6. Run `kurulumYap` from the function dropdown
7. Grant permissions when prompted
8. Return to your sheet — the **🏭 Production Tracker** menu will appear

---

## 📋 Usage

### Adding an Order (Management)
**Menu → Add New Order**
- Choose *New Order* (auto-assigns next order number) or *Add Item to Existing Order*
- Select model, enter quantity, set dates
- Automatically added to both Orders and Chassis Tracking sheets

### Chassis Entry (Procurement)
**Menu → Chassis Entry**
- Select model, supplier, enter received quantity and dates
- Summary row updates automatically (total received, remaining, status)
- Every shipment logged separately in the Entry Log

### Creating an ARP (Production Manager)
**Menu → Create New ARP**
- Select order, model, line, planned quantity
- ARP number auto-assigned (ARP-0001, ARP-0002...)
- Order status updates to "In Production"

### Daily Line Entry (Operators)
**Menu → Line Entry**
- Select date, line, model, produced quantity
- Optional: link to ARP number

### Dashboard Refresh
**Menu → Refresh Dashboard**
- Updates all KPI cards
- Rebuilds weekly production table
- Refreshes order progress bars

---

## 💾 Backup & Archive

| Action | Menu Item | What it does |
|--------|-----------|--------------|
| Manual backup | 📥 Backup (Save to Drive) | Copies entire file to Google Drive with timestamp |
| Archive line data | 🧹 Archive & Clear Line Tracking | Backs up first, then clears line tracking rows |
| System reset | 🔧 Setup / Reset System (Admin) | Password protected; auto-backup before wipe |

**Default admin password:** `ARORA00`

---

## ⚙️ Configuration

### Adding Models
1. **Menu → Settings**
2. Add new model names to the bottom of **Column A**
3. Click **✅ Done** — all dropdowns update automatically

### Adding Suppliers
1. **Menu → Settings**
2. Add supplier names to **Column C**

### Changing Admin Password
In the script, find `"ARORA00"` and replace with your password.

---

## 📁 Repository Structure

```
arora-production-tracker/
├── en/
│   └── ProductionTracker.gs     # English version
├── tr/
│   └── UretimTakip.gs           # Turkish version
├── docs/
│   └── screenshots/             # UI screenshots
├── README.md                    # English documentation (this file)
└── README.tr.md                 # Turkish documentation
```

---

## 🔑 ARP Numbering

Production orders follow the format `ARP-0001`, `ARP-0002`, etc.

*ARP = Arora Production* — inspired by the TSM (Temsa Service Management) naming convention used in the industry.

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

*Built with Google Apps Script. No external dependencies.*
