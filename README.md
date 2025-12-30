# 🍦 CFN Ice Cream Parlour POS System

👉 **[⬇️ Download CFN POS v1.0 (Windows EXE)](https://github.com/USERNAME/REPO/releases/download/v1.0/CFN.exe)**

---

## 🧾 Project Description

**CFN Ice Cream Parlour POS System** is a complete billing and stock management solution designed for small to medium ice cream shops.  
It supports **multi-table orders**, **category-based menus**, **admin inventory control**, **PDF bill generation**, and **thermal printer integration**.

This software works **offline**, stores data locally using **SQLite**, and is optimized for **Windows-based POS setups**.

---

## ✨ Key Features

### 🧮 Billing & POS
- Table-based billing (supports multiple tables)
- Add / remove items dynamically
- Quantity control (+ / –)
- Real-time total calculation
- Customer name & mobile number support
- Bill preview before printing

---

### 🖨️ Printing
- Thermal printer support (58mm / 80mm)
- ESC/POS direct printing (Windows)
- Kitchen Order Ticket (KOT) printing
- Automatic paper cut after print
- PDF bill generation (A4 size)

---

### 📦 Inventory & Stock Management
- Admin panel to manage items
- Category & subcategory support
- Stock auto-deduction on billing
- Low-stock visibility
- CSV import for bulk item upload

---

### 🍽️ Kitchen Management (KOT)
- Send orders to kitchen
- Kitchen queue view
- Mark orders as ready
- Print kitchen order slips

---

### 📊 Sales & Reports
- Date-wise sales report
- Grand total calculation
- Export sales data to CSV
- Bill history with customer details

---

### 🧑‍💼 Admin Panel
- Add / update / delete items
- Manage price, stock, category & subcategory
- Import items from CSV
- Real-time sync with POS screen

---

## 🛠️ Tech Stack

| Layer      | Technology            |
|-----------|------------------------|
| Language  | Python 3               |
| GUI       | Tkinter                |
| Database  | SQLite                 |
| Printing  | win32print (ESC/POS)   |
| PDF       | ReportLab              |
| OS        | Windows                |
