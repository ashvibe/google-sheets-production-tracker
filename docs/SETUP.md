# Quick Setup Guide

## 1. Customize Before First Run

Before running `kurulumYap`, open the script and update:

### Models (`_kurAyarlar` function)
```javascript
const modeller = [
  "YOUR MODEL 1",
  "YOUR MODEL 2",
  // ...
];
```

### Suppliers
```javascript
const tedarikciler = ["Your Supplier A", "Your Supplier B"];
```

### Admin Password
```javascript
// Find this line and set your own password:
if (r.getResponseText().trim() !== "YOUR_ADMIN_PASSWORD") {
```

## 2. Run Setup
- Apps Script → select `kurulumYap` → ▶️ Run
- Grant permissions
- Done ✅

## 3. Add Your Data
All data entry is via **Menu → 🏭 Production Tracker**

No need to edit cells directly.
