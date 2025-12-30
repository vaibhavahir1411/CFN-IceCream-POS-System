import os
import sqlite3
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
from datetime import datetime, date
import platform
from tkinter import simpledialog
import calendar
import csv
import platform
import win32print
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

DB_FILE = "items.db"
NUM_TABLES = 10

# ---------- Receipt / Thermal Printer settings ----------
# Common char widths: 58mm ~ 32-36 chars; 80mm ~ 42-48 chars, depending on printer font.
RECEIPT_PAPER = "58mm"  # "58mm" or "80mm"
RECEIPT_CHARS_PER_LINE = 32 


# ---------- Theme & Sizes ----------
BG          = "#F4F6F8"
CARD_BG     = "#FFFFFF"
BORDER      = "#E1E6EC"
BORDER_DARK = "#C9D2DC"
ACCENT      = "#1976D2"
ACCENT_SOFT = "#E3F2FD"
GREEN       = "#138A36"
TEXT        = "#1F2937"
MUTED       = "#6B7280"

# Font sizes that will scale
TITLE_SIZE  = 20
HEAD_SIZE   = 12
TEXT_SIZE   = 12
SMALL_SIZE  = 10

# Base dimensions that will scale
CARD_W      = 250
CARD_H      = 118
GRID_COLS   = 3
CART_MINSZ  = 360

# cart column widths
COL_ITEM_MIN   = 220
COL_QTY_W      = 160
COL_PRICE_MIN  = 90
COL_TOTAL_MIN  = 110

# Scale factor for different screen sizes
SCALE_FACTOR = 1.0

# ---------- Family pack helpers ----------
FAMILY_PACK_PRICE = 250.0
FAMILY_PACK_SUFFIX = " — Family Pack 500 ml"

# ---------- DB helpers ----------
def db_query(sql, params=(), one=False):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    if one:
        return dict(rows[0]) if rows else None
    return [dict(r) for r in rows]

def db_exec(sql, params=()):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute(sql, params)
    conn.commit()
    last = cur.lastrowid
    conn.close()
    return last

def ensure_subcategory_column():
    """Add items.subcategory TEXT if missing."""
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(items)")
    cols = [r[1].lower() for r in cur.fetchall()]
    if "subcategory" not in cols:
        cur.execute("ALTER TABLE items ADD COLUMN subcategory TEXT")
        conn.commit()
    conn.close()

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS items(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT,
            price REAL NOT NULL DEFAULT 0,
            stock INTEGER NOT NULL DEFAULT 0
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_items_name ON items(LOWER(name))")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sales(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            total REAL NOT NULL,
            customer_name TEXT,
            customer_mobile TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sale_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_id INTEGER NOT NULL,
            item_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            line_total REAL NOT NULL,
            FOREIGN KEY(sale_id) REFERENCES sales(id),
            FOREIGN KEY(item_id) REFERENCES items(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS kitchen_orders(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_no INTEGER NOT NULL,
            ts TEXT NOT NULL,
            table_no INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending'
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS kitchen_order_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ko_id INTEGER NOT NULL,
            item_name TEXT NOT NULL,
            qty INTEGER NOT NULL,
            FOREIGN KEY(ko_id) REFERENCES kitchen_orders(id)
        )
    """)
    conn.commit()
    conn.close()

    ensure_subcategory_column()
    
   


def get_categories_from_items():
    rows = db_query("SELECT DISTINCT category FROM items WHERE COALESCE(category,'')<>'' ORDER BY category")
    return [r["category"] for r in rows]

# ---------- Admin Panel ----------
class AdminPanel(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent.root)
        self.parent = parent
        self.title("Admin Panel — Manage Items")
        self.configure(bg=BG)
        self.scale_factor = SCALE_FACTOR * 0.9
        title_font = ("Segoe UI", int(TITLE_SIZE * self.scale_factor), "bold")
        head_font = ("Segoe UI", int(HEAD_SIZE * self.scale_factor), "bold")
        text_font = ("Segoe UI", int(TEXT_SIZE * self.scale_factor))
        small_font = ("Segoe UI", int(SMALL_SIZE * self.scale_factor))
        try:
            if platform.system().lower().startswith("win"):
                self.state('zoomed')
            else:
                self.attributes('-zoomed', True)
        except Exception:
            self.geometry("1200x700")
        self.minsize(int(1050 * self.scale_factor), int(600 * self.scale_factor))

        wrap = tk.Frame(self, bg=BG)
        wrap.pack(fill="both", expand=True, padx=12, pady=12)
        wrap.columnconfigure(0, weight=5)
        wrap.columnconfigure(1, weight=3)
        wrap.rowconfigure(2, weight=1)

        tk.Label(wrap, text="Items", font=head_font, bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w")
        search_bar = tk.Frame(wrap, bg=BG)
        search_bar.grid(row=1, column=0, sticky="new", pady=(2, 6))
        tk.Label(search_bar, text="Search:", font=text_font, bg=BG).pack(side="left")
        self.sv_search = tk.StringVar()
        e = tk.Entry(search_bar, textvariable=self.sv_search, font=text_font, width=30)
        e.pack(side="left", padx=6)
        e.bind("<KeyRelease>", lambda _e: self.refresh_table())
        tk.Button(search_bar, text="Clear", command=self._clear_search).pack(side="left")

        table_wrap = tk.Frame(wrap, bg=BG)
        table_wrap.grid(row=2, column=0, sticky="nsew")
        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_wrap,
            columns=("id", "name", "category", "subcategory", "price", "stock"),
            show="headings"
        )
        self.tree.grid(row=0, column=0, sticky="nsew")
        for col, txt, w, anchor in [
            ("id", "ID", int(60 * self.scale_factor), "e"),
            ("name", "Name", int(320 * self.scale_factor), "w"),
            ("category", "Category", int(160 * self.scale_factor), "w"),
            ("subcategory", "Subcategory", int(180 * self.scale_factor), "w"),
            ("price", "Price", int(100 * self.scale_factor), "e"),
            ("stock", "Stock", int(80 * self.scale_factor), "e"),
        ]:
            self.tree.heading(col, text=txt, anchor="w")
            self.tree.column(col, width=w, anchor=anchor, stretch=True)

        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", lambda _: self._load_selected())

        right = tk.Frame(wrap, bg=BG)
        right.grid(row=0, column=1, rowspan=3, sticky="nsew", padx=(16, 0))
        right.columnconfigure(1, weight=1)
        tk.Label(right, text="Details", font=head_font, bg=BG, fg=TEXT).grid(row=0, column=0, columnspan=2, sticky="w")

        self.var_id = tk.StringVar()
        self.var_name = tk.StringVar()
        self.var_cat = tk.StringVar()
        self.var_subcat = tk.StringVar()
        self.var_price = tk.StringVar()
        self.var_stock = tk.StringVar()

        tk.Label(right, text="(ID)", font=small_font, fg=MUTED, bg=BG).grid(row=1, column=0, sticky="w", pady=(6, 0))
        tk.Label(right, textvariable=self.var_id, font=small_font, fg=MUTED, bg=BG).grid(row=1, column=1, sticky="w", pady=(6, 0))

        self._row_entry(right, "Name", self.var_name, 2, text_font)

        tk.Label(right, text="Category", font=text_font, bg=BG).grid(row=3, column=0, sticky="w", pady=6)
        self.cat_combo = ttk.Combobox(right, textvariable=self.var_cat, font=text_font, state="normal")
        self._refresh_cat_values()
        self.cat_combo.grid(row=3, column=1, sticky="ew", pady=6)
        self.cat_combo.bind("<KeyRelease>", self._cat_autocomplete)

        self._row_entry(right, "Subcategory (optional)", self.var_subcat, 4, text_font)
        self._row_entry(right, "Price (₹)", self.var_price, 5, text_font)
        self._row_entry(right, "Stock", self.var_stock, 6, text_font)

        btns = tk.Frame(right, bg=BG)
        btns.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        tk.Button(btns, text="New", command=self._new_item).pack(side="left")
        tk.Button(btns, text="Save", command=self._save_item, bg=ACCENT, fg="white",
                  activebackground=ACCENT_SOFT).pack(side="left", padx=6)
        tk.Button(btns, text="Delete", command=self._delete_item).pack(side="left")
        tk.Button(btns, text="Import CSV", command=self._import_csv, bg=GREEN, fg="white",
                  activebackground="#2E7D32").pack(side="left", padx=6)
        tk.Button(btns, text="Close", command=self.destroy).pack(side="right")

        self.refresh_table()

    def _row_entry(self, parent, label, var, rowi, font):
        tk.Label(parent, text=label, font=font, bg=BG).grid(row=rowi, column=0, sticky="w", pady=6)
        ent = tk.Entry(parent, textvariable=var, font=font)
        ent.grid(row=rowi, column=1, sticky="ew", pady=6)
        return ent

    def _refresh_cat_values(self):
        self.cat_combo['values'] = get_categories_from_items() + ["Family Pack"]

    def _cat_autocomplete(self, event=None):
        typed = self.var_cat.get().strip()
        cats = get_categories_from_items() + ["Family Pack"]
        if typed:
            matches = [c for c in cats if c.lower().startswith(typed.lower())]
        else:
            matches = cats
        self.cat_combo['values'] = matches
        if matches:
            try:
                self.cat_combo.event_generate("<Down>")
            except Exception:
                pass

    def _clear_search(self):
        self.sv_search.set("")
        self.refresh_table()

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        q = self.sv_search.get().strip()
        if q:
            like = f"%{q}%"
            rows = db_query(
                "SELECT id,name,category,COALESCE(subcategory,'') AS subcategory,price,stock FROM items "
                "WHERE LOWER(name) LIKE LOWER(?) OR LOWER(category) LIKE LOWER(?) OR LOWER(COALESCE(subcategory,'')) LIKE LOWER(?) "
                "ORDER BY stock ASC",
                (like, like, like)
            )
        else:
            rows = db_query("SELECT id,name,category,COALESCE(subcategory,'') AS subcategory,price,stock FROM items ORDER BY stock ASC")
        for r in rows:
            self.tree.insert(
                "", "end",
                values=(r["id"], r["name"], r["category"], r["subcategory"], f"{float(r['price']):.2f}", r["stock"])
            )
        self._refresh_cat_values()

    def _load_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        _id, name, cat, subcat, price, stock = vals
        self.var_id.set(_id)
        self.var_name.set(name)
        self.var_cat.set(cat)
        self.var_subcat.set(subcat)
        self.var_price.set(str(price))
        self.var_stock.set(str(stock))

    def _new_item(self):
        self.var_id.set("")
        self.var_name.set("")
        self.var_cat.set("")
        self.var_subcat.set("")
        self.var_price.set("0")
        self.var_stock.set("0")
        self._refresh_cat_values()

    def _save_item(self):
        try:
            name = self.var_name.get().strip()
            if not name:
                messagebox.showwarning("Validation", "Name is required.")
                return
            cat = self.var_cat.get().strip()
            subcat = self.var_subcat.get().strip() or None
            price = float(self.var_price.get())
            stock = int(float(self.var_stock.get()))
        except ValueError:
            messagebox.showwarning("Validation", "Price must be a number and Stock must be an integer.")
            return

        _id = self.var_id.get().strip()
        if _id:
            db_exec("UPDATE items SET name=?, category=?, subcategory=?, price=?, stock=? WHERE id=?",
                    (name, cat, subcat, price, stock, int(_id)))
            messagebox.showinfo("Saved", "Item updated.")
        else:
            new_id = db_exec("INSERT INTO items(name, category, subcategory, price, stock) VALUES (?,?,?,?,?)",
                             (name, cat, subcat, price, stock))
            self.var_id.set(str(new_id))
            messagebox.showinfo("Saved", "Item added.")


        self.refresh_table()
        self.parent._reload_left_nav_from_db()
        self.parent._apply_filters_and_load()

    def _delete_item(self):
        _id = self.var_id.get().strip()
        if not _id:
            messagebox.showwarning("Delete", "Select an item to delete.")
            return
        if not messagebox.askyesno("Confirm", f"Delete item ID #{_id}?"):
            return
        db_exec("DELETE FROM items WHERE id=?", (int(_id),))
        self._new_item()
        self.refresh_table()
        self.parent._reload_left_nav_from_db()
        self.parent._apply_filters_and_load()

    def _import_csv(self):
        """Import items from CSV with columns: Name, Category, Price, Stock, (optional) Subcategory"""
        file_path = filedialog.askopenfilename(
            title="Select CSV File to Import",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            imported_count = 0
            updated_count = 0
            errors = []
            with open(file_path, 'r', encoding='utf-8') as csvfile:
                sample = csvfile.read(1024); csvfile.seek(0)
                sniffer = csv.Sniffer()
                try:
                    dialect = sniffer.sniff(sample)
                    delimiter = dialect.delimiter
                except Exception:
                    delimiter = ','
                reader = csv.DictReader(csvfile, delimiter=delimiter)
                fieldnames = [field.strip() for field in reader.fieldnames]
                required_cols = ['name', 'category', 'price', 'stock']
                optional_cols = ['subcategory']
                header_mapping = {}
                for req_col in required_cols:
                    found = None
                    for field in fieldnames:
                        if field.lower() == req_col:
                            found = field; break
                    if not found:
                        messagebox.showerror(
                            "Invalid CSV",
                            f"Missing required column: '{req_col}'\n"
                            f"Found: {', '.join(fieldnames)}\n"
                            f"Required: Name, Category, Price, Stock (optional: Subcategory)"
                        )
                        return
                    header_mapping[req_col] = found
                for opt_col in optional_cols:
                    found = None
                    for field in fieldnames:
                        if field.lower() == opt_col:
                            found = field; break
                    header_mapping[opt_col] = found  # may be None

                for row_num, row in enumerate(reader, start=2):
                    try:
                        name = str(row[header_mapping['name']]).strip()
                        category = str(row[header_mapping['category']]).strip()
                        subcategory = None
                        if header_mapping['subcategory'] is not None:
                            sub_str = str(row[header_mapping['subcategory']]).strip()
                            subcategory = sub_str if sub_str else None
                        price_str = str(row[header_mapping['price']]).strip()
                        stock_str = str(row[header_mapping['stock']]).strip()
                        if not name:
                            errors.append(f"Row {row_num}: Name is required")
                            continue
                        try:
                            price_clean = price_str.replace('₹', '').replace('$', '').replace(',', '')
                            price = float(price_clean)
                            if price < 0:
                                errors.append(f"Row {row_num}: Price cannot be negative"); continue
                        except ValueError:
                            errors.append(f"Row {row_num}: Invalid price '{price_str}'"); continue
                        try:
                            stock = int(float(stock_str))
                            if stock < 0:
                                errors.append(f"Row {row_num}: Stock cannot be negative"); continue
                        except ValueError:
                            errors.append(f"Row {row_num}: Invalid stock '{stock_str}'"); continue

                        existing = db_query("SELECT id FROM items WHERE LOWER(name)=LOWER(?)", (name,), one=True)
                        if existing:
                            db_exec("UPDATE items SET category=?, subcategory=?, price=?, stock=? WHERE id=?",
                                    (category, subcategory, price, stock, existing['id']))
                            updated_count += 1
                        else:
                            db_exec("INSERT INTO items(name, category, subcategory, price, stock) VALUES (?,?,?,?,?)",
                                    (name, category, subcategory, price, stock))
                            imported_count += 1
                    except Exception as e:
                        errors.append(f"Row {row_num}: {str(e)}"); continue

            result_msg = f"Import completed!\n\nNew items added: {imported_count}\nItems updated: {updated_count}"
            if errors:
                error_msg = "\n".join(errors[:10])
                if len(errors) > 10:
                    error_msg += f"\n... and {len(errors)-10} more errors"
                result_msg += f"\n\nErrors:\n{error_msg}"
            if imported_count > 0 or updated_count > 0:
                messagebox.showinfo("Import Results", result_msg)
                
                self.refresh_table()
                self.parent._reload_left_nav_from_db()
                self.parent._apply_filters_and_load()
                self._new_item()
            else:
                messagebox.showwarning("Import Results", result_msg)
        except FileNotFoundError:
            messagebox.showerror("Error", "File not found.")
        except UnicodeDecodeError:
            messagebox.showerror("Error", "Unable to read file. Ensure UTF-8 CSV.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while importing:\n{str(e)}")

# ---------- Main POS App ----------
class POSApp:
    def _generate_pdf_bill(self, sale_id):
        """Generate PDF bill and save to Bills folder"""
        sale = db_query("SELECT * FROM sales WHERE id=?", (sale_id,), one=True)
        lines = db_query("""
            SELECT i.name, i.price, si.quantity, si.line_total
            FROM sale_items si JOIN items i ON i.id=si.item_id
            WHERE si.sale_id=?
        """, (sale_id,))
        
        # Create Bills folder if it doesn't exist
        bills_folder = "Bills"
        if not os.path.exists(bills_folder):
            os.makedirs(bills_folder)
        
        # Create filename with date and bill number
        sale_date = sale['ts'][:10]  # Extract date part (YYYY-MM-DD)
        filename = f"Bill_{sale_id}_{sale_date.replace('-', '')}.pdf"
        filepath = os.path.join(bills_folder, filename)
        
        # Create PDF document
        doc = SimpleDocTemplate(filepath, pagesize=A4, 
                            leftMargin=0.5*inch, rightMargin=0.5*inch,
                            topMargin=0.5*inch, bottomMargin=0.5*inch)
        
        # Container for the 'Flowable' objects
        elements = []
        
        # Define styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=12,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=6,
            alignment=TA_CENTER
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            alignment=TA_LEFT
        )
        
        # Add title
        title = Paragraph("Cream Fresh Natural Icecream", title_style)
        elements.append(title)
        elements.append(Spacer(1, 12))
        
        # Add bill info
        bill_info = f"<b>Bill #{sale_id}</b><br/>{sale['ts']}"
        bill_para = Paragraph(bill_info, header_style)
        elements.append(bill_para)
        
        # Add customer info if available
        customer_name = sale.get('customer_name', '').strip()
        customer_mobile = sale.get('customer_mobile', '').strip()
        
        if customer_name or customer_mobile:
            customer_info = ""
            if customer_name:
                customer_info += f"<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Customer:</b> {customer_name}<br/>"
            if customer_mobile:
                customer_info += f"<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mobile:</b> {customer_mobile}"
            customer_para = Paragraph(customer_info, normal_style)
            elements.append(customer_para)
        
        elements.append(Spacer(1, 20))
        
        # Create table data
        table_data = [['Item', 'Qty', 'Rate', 'Amount']]
        
        total_amount = 0.0
        for line in lines:
            name = line['name']
            qty = str(line['quantity'])
            rate = f"{float(line['price']):.2f}"
            amount = f"{float(line['line_total']):.2f}"
            total_amount += float(line['line_total'])
            
            table_data.append([name, qty, rate, amount])
        
        # Add total row
        table_data.append(['', '', 'TOTAL:', f"{total_amount:.2f}"])
        
        # Create table
        table = Table(table_data, colWidths=[3*inch, 0.8*inch, 1*inch, 1.2*inch])
        
        # Add table styling
        table.setStyle(TableStyle([
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            
            # Data rows styling
            ('BACKGROUND', (0, 1), (-1, -2), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -2), colors.black),
            ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -2), 10),
            
            # Total row styling
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -1), (-1, -1), 12),
            
            # Table borders
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            
            # Alignment
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),   # Item names left-aligned
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'), # Numbers right-aligned
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        # Add thank you message
        thank_you = Paragraph("<b>Thank you for visiting!</b>", header_style)
        elements.append(thank_you)
        
        # Build PDF
        try:
            doc.build(elements)
            return filepath
        except Exception as e:
            raise Exception(f"Failed to generate PDF: {str(e)}")
    
    def _save_customer_name(self, *args):
       
        customer_name = self.customer_name_var.get().strip()
        self.customer_names[self.active_table] = customer_name  

    def _save_customer_mobile(self, *args):
        customer_mobile = self.customer_mobile_var.get().strip()
        self.customer_mobile_numbers[self.active_table] = customer_mobile 
        
    def __init__(self, root):
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        global SCALE_FACTOR
        if screen_width <= 1366:
            SCALE_FACTOR = 0.85
        elif screen_width <= 1600:
            SCALE_FACTOR = 0.95
        else:
            SCALE_FACTOR = 1.0

        self.card_w = int(CARD_W * SCALE_FACTOR)
        self.card_h = int(CARD_H * SCALE_FACTOR)
        self.cart_minsz = int(CART_MINSZ * SCALE_FACTOR)
        self.col_item_min = int(COL_ITEM_MIN * SCALE_FACTOR)
        self.col_qty_w = int(COL_QTY_W * SCALE_FACTOR)
        self.col_price_min = int(COL_PRICE_MIN * SCALE_FACTOR)
        self.col_total_min = int(COL_TOTAL_MIN * SCALE_FACTOR)
        self.title_font = ("Segoe UI", int(TITLE_SIZE * SCALE_FACTOR), "bold")
        self.head_font = ("Segoe UI", int(HEAD_SIZE * SCALE_FACTOR), "bold")
        self.text_font = ("Segoe UI", int(TEXT_SIZE * SCALE_FACTOR))
        self.small_font = ("Segoe UI", int(SMALL_SIZE * SCALE_FACTOR))
        self.mono_font = ("Consolas", int(TEXT_SIZE * SCALE_FACTOR))

        init_db()

        self.root = root
        self.root.title("Ice Cream Parlour POS System")
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.85)
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(int(1000 * SCALE_FACTOR), int(600 * SCALE_FACTOR))
        self.root.configure(bg=BG)
        self._style()

        self.carts = {i: [] for i in range(1, NUM_TABLES + 1)}
        self.active_table = 1
        self.search_text = tk.StringVar()  # center search
        self.customer_name_var = tk.StringVar()
        self.customer_names = {i: "" for i in range(1, NUM_TABLES + 1)}
        self.customer_mobile_var = tk.StringVar()
        self.customer_mobile_numbers = {i: "" for i in range(1, NUM_TABLES + 1)}  # Initialize with empty mobile numbers
        self.customer_name_var.trace_add("write", self._save_customer_name)
        self.customer_mobile_var.trace_add("write", self._save_customer_mobile)

        # Left nav state
        self.expanded_cats = set()
        self.expanded_subcats = set()
        self.left_entries = []  # list of dicts with keys: type, cat, subcat, item_id, item_name, label

        # Active filter
        self.active_filter = {"mode": "all"}  # modes: all, cat, subcat, item

        self._header()
        self._body()
        self._statusbar()
        self._reload_left_nav_from_db()
        self._apply_filters_and_load()
        self._tick_clock()

    # ---------- UI Sections ----------
    def _header(self):
        top = tk.Frame(self.root, bg=BG)
        top.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(top, text="CFN Ice Cream Parlour", font=self.title_font, fg=TEXT, bg=BG).pack(side="left")

        mid = tk.Frame(top, bg=BG)
        mid.pack(side="left", padx=16)
        tk.Label(mid, text="Tables:", font=self.head_font, fg=TEXT, bg=BG).pack(side="left", padx=(0, 6))
        self.tbl_buttons = []
        for i in range(1, NUM_TABLES + 1):
            b = tk.Button(mid, text=str(i), width=3, font=self.text_font,
                          bg="white", activebackground=ACCENT_SOFT, relief="raised",
                          command=lambda t=i: self._switch_table(t))
            b.pack(side="left", padx=3)
            self.tbl_buttons.append(b)
        self._refresh_table_buttons()
        self.clock_lbl = tk.Label(top, text="", font=self.text_font, fg=MUTED, bg=BG)
        self.clock_lbl.pack(side="right")

    def _body(self):
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=12, pady=8)
        body.columnconfigure(0, weight=0, minsize=int(240 * SCALE_FACTOR))
        body.columnconfigure(1, weight=3)
        body.columnconfigure(2, weight=2, minsize=self.cart_minsz)
        body.rowconfigure(0, weight=1)

        # LEFT: dynamic categories/subcategories/items
        self.left_panel = tk.Frame(body, bg=BG)
        left = self.left_panel
        left.grid(row=0, column=0, sticky="nsew")
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(left, text="Browse", font=self.head_font, fg=TEXT, bg=BG).grid(row=0, column=0, sticky="w")
        self.cat_list = tk.Listbox(left, height=18, exportselection=False, font=self.text_font, bd=1, relief="solid")
        self.cat_list.grid(row=1, column=0, sticky="nsew", pady=(6, 4))
        self.cat_list.bind("<<ListboxSelect>>", self._on_left_select)

        # CENTER: menu grid + search
        center = tk.Frame(body, bg=BG)
        center.grid(row=0, column=1, sticky="nsew", padx=(10, 10))
        center.rowconfigure(2, weight=1)
        center.columnconfigure(0, weight=1)
        header_row = tk.Frame(center, bg=BG)
        header_row.grid(row=0, column=0, sticky="ew")
        tk.Label(header_row, text="Menu Items", font=self.head_font, fg=TEXT, bg=BG).pack(side="left")

        search_box = tk.Frame(center, bg=BG)
        search_box.grid(row=1, column=0, sticky="ew", pady=(2, 6))
        tk.Label(search_box, text="Search:", font=self.text_font, bg=BG).pack(side="left")
        se = tk.Entry(search_box, textvariable=self.search_text, font=self.text_font, width=28)
        se.pack(side="left", padx=6)
        se.bind("<KeyRelease>", lambda e: self._apply_filters_and_load())
        tk.Button(search_box, text="Clear", command=lambda: self._clear_main_search()).pack(side="left")

        self.menu_canvas = tk.Canvas(center, highlightthickness=0, bg=BG)
        scroll = ttk.Scrollbar(center, orient="vertical", command=self.menu_canvas.yview)
        self.menu_canvas.configure(yscrollcommand=scroll.set)
        self.menu_canvas.grid(row=2, column=0, sticky="nsew")
        scroll.grid(row=2, column=1, sticky="ns")
        self.grid_host = tk.Frame(self.menu_canvas, bg=BG)
        self.grid_host.bind("<Configure>", lambda e: self.menu_canvas.configure(scrollregion=self.menu_canvas.bbox("all")))
        self.menu_canvas.create_window((0, 0), window=self.grid_host, anchor="nw")

        def _mw(e):
            if self._cursor_in(self.menu_canvas):
                self.menu_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
            elif self._cursor_in(self.cart_canvas):
                self.cart_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        self.root.bind_all("<MouseWheel>", _mw)

        # RIGHT: cart + customer
        right = tk.Frame(body, bg=BG)
        right.grid(row=0, column=2, sticky="nsew")
        cust = tk.Frame(right, bg=BG)
        cust.pack(fill="x", pady=(0, 6))
        tk.Label(cust, text="Customer Name (optional):", font=self.text_font, bg=BG).grid(row=0, column=0, sticky="w")
        tk.Entry(cust, textvariable=self.customer_name_var, font=self.text_font).grid(row=0, column=1, sticky="ew", padx=(6,0))
        tk.Label(cust, text="Mobile (optional):", font=self.text_font, bg=BG).grid(row=1, column=0, sticky="w", pady=(4,0))
        tk.Entry(cust, textvariable=self.customer_mobile_var, font=self.text_font).grid(row=1, column=1, sticky="ew", padx=(6,0))
        cust.grid_columnconfigure(1, weight=1)

        self.cart_title = tk.Label(right, text=f"Cart (Table {self.active_table})", font=self.head_font, fg=TEXT, bg=BG)
        self.cart_title.pack(anchor="w")

        hdr = tk.Frame(right, bg=BG, bd=0, relief="flat")
        hdr.pack(fill="x", pady=(2, 0))
        hdr.grid_columnconfigure(0, weight=1,  minsize=self.col_item_min)
        hdr.grid_columnconfigure(1,           minsize=self.col_qty_w)
        hdr.grid_columnconfigure(2,           minsize=self.col_price_min)
        hdr.grid_columnconfigure(3, weight=2, minsize=self.col_total_min)

        tk.Label(hdr, text="Item",  anchor="w", font=self.small_font, fg=MUTED, bg=BG)\
            .grid(row=0, column=0, sticky="w", padx=6)
        tk.Label(hdr, text="Qty",   anchor="center", font=self.small_font, fg=MUTED, bg=BG)\
            .grid(row=0, column=1)
        tk.Label(hdr, text="Price", anchor="e", font=self.small_font, fg=MUTED, bg=BG)\
            .grid(row=0, column=2, sticky="e")
        tk.Label(hdr, text="Total", anchor="e", font=self.small_font, fg=MUTED, bg=BG)\
            .grid(row=0, column=3, sticky="e", padx=(0, 10))
        tk.Frame(right, height=1, bg=BORDER_DARK).pack(fill="x", pady=(2, 0))

        cart_container = tk.Frame(right, bg=BG, height=int(250 * SCALE_FACTOR))
        cart_container.pack(fill="both", expand=True)
        cart_container.pack_propagate(False)
        cart_box = tk.Frame(cart_container, bg="white", bd=1, relief="solid", highlightbackground=BORDER, highlightcolor=BORDER)
        cart_box.pack(fill="both", expand=True)
        self.cart_canvas = tk.Canvas(cart_box, highlightthickness=0, bg="white")
        cart_scroll = ttk.Scrollbar(cart_box, orient="vertical", command=self.cart_canvas.yview)
        self.cart_rows = tk.Frame(self.cart_canvas, bg="white")
        self.cart_rows.bind("<Configure>", lambda e: self.cart_canvas.configure(scrollregion=self.cart_canvas.bbox("all")))
        self.cart_rows.grid_columnconfigure(0, weight=1)
        self.cart_canvas.create_window((0, 0), window=self.cart_rows, anchor="nw")
        self.cart_canvas.configure(yscrollcommand=cart_scroll.set)
        self.cart_canvas.pack(side="left", fill="both", expand=True)
        cart_scroll.pack(side="right", fill="y")

        footer = tk.Frame(right, bg=BG)
        footer.pack(fill="x", pady=(6, 6))
        footer.grid_columnconfigure(0, weight=1,  minsize=self.col_item_min)
        footer.grid_columnconfigure(1,           minsize=self.col_qty_w)
        footer.grid_columnconfigure(2,           minsize=self.col_price_min)
        footer.grid_columnconfigure(3, weight=2, minsize=self.col_total_min)

        tk.Label(footer, text="Total:", font=self.head_font, fg=TEXT, bg=BG, anchor="e")\
            .grid(row=0, column=2, sticky="e", padx=(3, 3))
        self.total_val = tk.Label(footer, text="₹0.00", font=self.head_font, fg=TEXT, bg=BG, anchor="e")
        self.total_val.grid(row=0, column=3, sticky="e", padx=(3, 8))

        tk.Button(right, text="Send to Kitchen (KOT)", font=self.text_font, command=self._send_to_kitchen)\
            .pack(fill="x", pady=(0, 6))
        tk.Button(right, text="Checkout (Active Table)", font=self.text_font, bg=ACCENT, fg="white",
                  activebackground=ACCENT_SOFT, command=self._checkout).pack(fill="x")
        tk.Button(right, text="Cancel All (Active Table)", font=self.text_font, command=self._cancel_all).pack(fill="x", pady=(6, 0))
        tk.Button(right, text="Kitchen Queue", font=self.text_font, command=self._show_kitchen_queue).pack(fill="x", pady=(6, 0))
        tk.Button(right, text="Sales (Pick Date)", font=self.text_font, command=self._show_sales).pack(fill="x", pady=(6, 0))
        tk.Button(right, text="Admin Panel", font=self.text_font, command=self._open_admin).pack(fill="x", pady=(6, 6))

    def _statusbar(self):
        sb = tk.Frame(self.root, bd=1, relief="sunken", bg="white")
        sb.pack(fill="x", side="bottom")
        tk.Label(sb, text="F1=New Sale   ESC=Cancel All (Active Table)", anchor="w",
                 font=self.small_font, bg="white", fg=MUTED).pack(side="left", padx=10)
        self.root.bind("<F1>", lambda e: self._new_sale())
        self.root.bind("<Escape>", lambda e: self._cancel_all())

    # ---------- Style ----------
    def _style(self):
        s = ttk.Style()
        try:
            s.theme_use("vista")
        except:
            pass
        s.configure("TScrollbar", troughcolor="white", background=BORDER)

    # ---------- Left Nav: Build from DB ----------
    def _reload_left_nav_from_db(self):
        """Build left_entries (All, Categories, Subcategories, Items) from DB respecting expanded state."""
        # Snapshot expansion sets; they persist across reloads
        cats = db_query("SELECT DISTINCT category FROM items WHERE COALESCE(category,'')<>'' ORDER BY category")
        entries = []
        # Always start with 'All'
        entries.append({"type": "all", "label": "All"})
        for rc in cats:
            cat = rc["category"]
            cat_label = f"{'▾ ' if cat in self.expanded_cats else '▸ '} {cat}"
            entries.append({"type": "category", "cat": cat, "label": cat_label})
            if cat in self.expanded_cats:
                # Items directly under category (no subcategory)
                # items_no_sub = db_query(
                #     "SELECT id,name FROM items WHERE category=? AND (subcategory IS NULL OR TRIM(subcategory)='') ORDER BY name",
                #     (cat,))
                # for it in items_no_sub:
                #     entries.append({
                #         "type": "item",
                #         "cat": cat,
                #         "subcat": None,
                #         "item_id": it["id"],
                #         "item_name": it["name"],
                #         "label": f"     · {it['name']}"
                #     })
                # Subcategories under this category
                subs = db_query(
                    "SELECT DISTINCT subcategory FROM items WHERE category=? AND COALESCE(TRIM(subcategory),'')<>'' ORDER BY subcategory",
                    (cat,))
                for rs in subs:
                   sub = rs["subcategory"]
                   entries.append({"type": "subcategory", "cat": cat, "subcat": sub, "label": f"  ▸ {sub}"})

        self.left_entries = entries
        # Render to listbox
        self.cat_list.delete(0, "end")
        for e in self.left_entries:
            self.cat_list.insert("end", e["label"])
        # Keep selection consistent with active filter
        self._select_left_by_active_filter()

    def _select_left_by_active_filter(self):
        idx = 0
        mode = self.active_filter.get("mode", "all")
        if mode == "all":
            # 'All' is index 0
            self.cat_list.select_clear(0, "end")
            self.cat_list.select_set(0)
            return
        for i, e in enumerate(self.left_entries):
            if mode == "cat" and e["type"] == "category" and e.get("cat") == self.active_filter.get("cat"):
                idx = i
                break
            if mode == "subcat" and e["type"] == "subcategory" and \
               e.get("cat") == self.active_filter.get("cat") and e.get("subcat") == self.active_filter.get("subcat"):
                idx = i
                break
            if mode == "item" and e["type"] == "item" and e.get("item_id") == self.active_filter.get("item_id"):
                idx = i
                break
        self.cat_list.select_clear(0, "end")
        self.cat_list.select_set(idx)

    def _on_left_select(self, event=None):
        sel = self.cat_list.curselection()
        if not sel:
            return
        entry = self.left_entries[sel[0]]
        etype = entry["type"]

        if etype == "all":
            self.active_filter = {"mode": "all"}
            self._apply_filters_and_load()
            return

        if etype == "category":
            cat = entry["cat"]

            # check if this category has any subcategories
            has_subs = db_query(
                "SELECT 1 AS x FROM items WHERE category=? AND COALESCE(TRIM(subcategory),'')<>'' LIMIT 1",
                (cat,),
                one=True
            )

            if has_subs:
                # keep old behavior (expand/collapse + reload) for categories that DO have subcats
                if cat in self.expanded_cats:
                    self.expanded_cats.remove(cat)
                else:
                    self.expanded_cats.add(cat)
                self.active_filter = {"mode": "cat", "cat": cat}
                self._reload_left_nav_from_db()
                self._apply_filters_and_load()
            else:
                # NEW: for categories with NO subcats, just filter center grid (no expand/reload)
                self.active_filter = {"mode": "cat", "cat": cat}
                self._apply_filters_and_load()
            return


        if etype == "subcategory":
            self.active_filter = {"mode": "subcat", "cat": entry["cat"], "subcat": entry["subcat"]}
            # no expand/collapse, no left reload—just filter center grid
            self._apply_filters_and_load()
            return

        if etype == "item":
            # Select only that item (center grid will show just it)
            self.active_filter = {"mode": "item", "item_id": entry["item_id"]}
            self._apply_filters_and_load()
            return

    # ---------- Helpers ----------
    def _tick_clock(self):
        self.clock_lbl.config(text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self._tick_clock)

    def _cursor_in(self, w):
        x, y = w.winfo_pointerxy()
        return w.winfo_rootx() <= x <= w.winfo_rootx() + w.winfo_width() and \
               w.winfo_rooty() <= y <= w.winfo_rooty() + w.winfo_height()

    def _clear_main_search(self):
        self.search_text.set("")
        self._apply_filters_and_load()

    # ---------- Center: Items load/render ----------
    def _apply_filters_and_load(self):
       
        q = self.search_text.get().strip()
        args, where = [], []

        mode = self.active_filter.get("mode", "all")
        if mode == "all":
            pass
        elif mode == "cat":
            cat = self.active_filter.get("cat")
            where.append("category=?"); args.append(cat)
        elif mode == "subcat":
            cat = self.active_filter.get("cat"); sub = self.active_filter.get("subcat")
            where.append("category=? AND COALESCE(subcategory,'')=?"); args.extend([cat, sub])
        elif mode == "item":
            iid = self.active_filter.get("item_id")
            where.append("id=?"); args.append(iid)

        if q:
            where.append("(LOWER(name) LIKE LOWER(?) OR LOWER(category) LIKE LOWER(?) OR LOWER(COALESCE(subcategory,'')) LIKE LOWER(?))")
            args.extend([f"%{q}%", f"%{q}%", f"%{q}%"])

        sql = "SELECT * FROM items"
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY name"
        items = db_query(sql, tuple(args))
        self._render_items(items)

    def _render_items(self, items):
        for w in self.grid_host.winfo_children():
            w.destroy()
        available_width = self.grid_host.winfo_width()
        if available_width < 1:
            available_width = self.root.winfo_width() - int(400 * SCALE_FACTOR)
        cols = max(GRID_COLS, min(GRID_COLS, available_width // (self.card_w + 16)))
        for i in range(cols):
            self.grid_host.columnconfigure(i, weight=1, minsize=self.card_w)

        r = c = 0
        for it in items:
            card = tk.Frame(self.grid_host, bd=1, relief="solid", bg=CARD_BG, highlightbackground=BORDER)
            card.grid(row=r, column=c, padx=8, pady=8, sticky="nsew")
            card.configure(width=self.card_w, height=self.card_h)
            card.grid_propagate(False)

            name_lbl = tk.Label(card, text=it["name"], font=self.head_font, fg=TEXT, bg=CARD_BG,
                                wraplength=self.card_w-18, justify="left", anchor="w")
            name_lbl.pack(fill="x", anchor="w", padx=8, pady=(8, 2))

            price_lbl = tk.Label(card, text=f"₹{float(it['price']):.2f}",
                                 font=("Segoe UI", int(14 * SCALE_FACTOR), "bold"), fg=GREEN, bg=CARD_BG, anchor="w")
            price_lbl.pack(fill="x", anchor="w", padx=8)

            info_parts = [f"Stock: {it['stock']}"]
            info = tk.Label(card, text=" | ".join(info_parts), font=self.small_font, fg=MUTED, bg=CARD_BG, anchor="w")
            info.pack(fill="x", anchor="w", padx=8, pady=(2, 6))

            def _on_click(e, r=it):
                self._add_item(r)
                return "break"
            for child in (card, name_lbl, price_lbl, info):
                try:
                    child.configure(takefocus=0, highlightthickness=0, bd=1 if child is card else 0)
                except Exception:
                    pass
                child.bind("<Button-1>", _on_click)

            c += 1
            if c >= cols:
                c = 0; r += 1
        # Update scroll region after adding all items
        self.root.after_idle(lambda: self.menu_canvas.configure(scrollregion=self.menu_canvas.bbox("all")))

    # ---------- Cart ops ----------
    def _active_cart(self):
        return self.carts[self.active_table]

    def _add_item(self, row):
        cart = self._active_cart()
        iid, name, price = row["id"], row["name"], float(row["price"])
        for idx, (ciid, nm, pr, qty) in enumerate(cart):
            if ciid == iid:
                cart[idx] = (ciid, nm, pr, qty + 1)
                self._refresh_cart()
                return
        cart.append((iid, name, price, 1))
        self._refresh_cart()

    def _refresh_cart(self):
        for w in self.cart_rows.winfo_children():
            w.destroy()
        cart = self._active_cart()
        total = 0.0
        for idx, (iid, name, price, qty) in enumerate(cart):
            amt = price * qty
            total += amt
            rf = tk.Frame(self.cart_rows, bg="white")
            rf.grid(row=idx, column=0, sticky="ew")
            rf.grid_columnconfigure(0, weight=1,  minsize=self.col_item_min)
            rf.grid_columnconfigure(1,           minsize=self.col_qty_w)
            rf.grid_columnconfigure(2,           minsize=self.col_price_min)
            rf.grid_columnconfigure(3, weight=2, minsize=self.col_total_min)
            if idx > 0:
                tk.Frame(rf, height=1, bg=BORDER).grid(row=0, column=0, columnspan=4, sticky="ew")
            row = tk.Frame(rf, bg="white")
            row.grid(row=1, column=0, columnspan=4, sticky="ew", padx=6, pady=4)
            tk.Label(row, text=name, width=28, anchor="w", font=self.text_font, bg="white", fg=TEXT)\
                .grid(row=0, column=0, sticky="w")
            tk.Frame(row, width=1, bg=BORDER).grid(row=0, column=1, sticky="ns", padx=(6, 6))
            tk.Frame(row, width=1, bg=BORDER).grid(row=0, column=3, sticky="ns", padx=(6, 6))
            tk.Frame(row, width=1, bg=BORDER).grid(row=0, column=5, sticky="ns", padx=(6, 6))
            ctr = tk.Frame(row, bg="white")
            ctr.grid(row=0, column=2, sticky="ew")
            tk.Button(ctr, text="-", width=2, font=self.text_font, command=lambda i=idx: self._dec(i)).pack(side="left")
            tk.Label(ctr, text=str(qty), width=4, anchor="center", font=self.text_font, bg="white").pack(side="left", padx=4)
            tk.Button(ctr, text="+", width=2, font=self.text_font, command=lambda i=idx: self._inc(i)).pack(side="left")
            tk.Button(ctr, text="✕", width=2, font=self.text_font, command=lambda i=idx: self._rm(i)).pack(side="left", padx=(6, 0))
            tk.Label(row, text=f"₹{price:.2f}", width=10, anchor="e", font=self.text_font, bg="white", fg=TEXT)\
                .grid(row=0, column=4, sticky="e")
            tk.Label(row, text=f"₹{amt:.2f}",   width=12, anchor="e", font=self.text_font, bg="white", fg=TEXT)\
                .grid(row=0, column=6, sticky="e")
            tk.Frame(rf, height=1, bg=BORDER).grid(row=2, column=0, columnspan=4, sticky="ew")

        self.total_val.config(text=f"₹{total:.2f}")
        self.cart_title.config(text=f"Cart (Table {self.active_table})")

    def _inc(self, idx):
        cart = self._active_cart()
        iid, nm, pr, qty = cart[idx]
        cart[idx] = (iid, nm, pr, qty + 1)
        self._refresh_cart()

    def _dec(self, idx):
        cart = self._active_cart()
        iid, nm, pr, qty = cart[idx]
        if qty > 1:
            cart[idx] = (iid, nm, pr, qty - 1)
        else:
            cart.pop(idx)
        self._refresh_cart()

    def _rm(self, idx):
        cart = self._active_cart()
        cart.pop(idx)
        self._refresh_cart()

    # ---------- KOT ----------
    def _next_order_no_today(self):
        today = date.today().isoformat()
        row = db_query("SELECT MAX(order_no) AS mx FROM kitchen_orders WHERE substr(ts,1,10)=?", (today,), one=True)
        return (row["mx"] or 0) + 1

    def _send_to_kitchen(self):
        cart = self._active_cart()
        if not cart:
            messagebox.showwarning("Empty", "No items in cart to send.")
            return
        ko_no = self._next_order_no_today()
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ko_id = db_exec("INSERT INTO kitchen_orders(order_no, ts, table_no, status) VALUES (?,?,?,?)",
                        (ko_no, ts, self.active_table, "pending"))
        conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
        for _, name, _, qty in cart:
            cur.execute("INSERT INTO kitchen_order_items(ko_id,item_name,qty) VALUES (?,?,?)", (ko_id, name, qty))
        conn.commit(); conn.close()
        messagebox.showinfo("Sent", f"Order #{ko_no} sent to kitchen.\nTable {self.active_table}")

    def _show_kitchen_queue(self):
        win = tk.Toplevel(self.root); win.title("Kitchen Queue (Pending)")
        top = tk.Frame(win); top.pack(fill="both", expand=True, padx=10, pady=10)
        cols = ("Order#", "Time", "Table", "Status")
        tree = ttk.Treeview(top, columns=cols, show="headings", height=16)
        for c, w, a in [("Order#", int(90 * SCALE_FACTOR), "e"), ("Time", int(180 * SCALE_FACTOR), "w"),
                        ("Table", int(70 * SCALE_FACTOR), "e"), ("Status", int(100 * SCALE_FACTOR), "w")]:
            tree.heading(c, text=c); tree.column(c, width=w, anchor=a)
        tree.pack(side="left", fill="both", expand=True)
        side = tk.Frame(top); side.pack(side="left", fill="y", padx=(8,0))
        items_box = tk.Text(side, width=int(40 * SCALE_FACTOR), height=16, font=self.text_font, state="disabled")
        items_box.pack(fill="y")

        def load_queue():
            tree.delete(*tree.get_children())
            rows = db_query("SELECT id, order_no, ts, table_no, status FROM kitchen_orders "
                            "WHERE status='pending' ORDER BY id ASC")
            for r in rows:
                tree.insert("", "end", iid=str(r["id"]),
                            values=(r["order_no"], r["ts"], r["table_no"], r["status"]))

        def on_select(_e=None):
            sel = tree.selection()
            items_box.config(state="normal"); items_box.delete("1.0","end")
            if not sel:
                items_box.config(state="disabled"); return
            ko_id = int(sel[0])
            rows = db_query("SELECT item_name, qty FROM kitchen_order_items WHERE ko_id=?", (ko_id,))
            for r in rows:
                items_box.insert("end", f"{r['item_name']}  x{r['qty']}\n")
            items_box.config(state="disabled")

        def mark_ready():
            sel = tree.selection()
            if not sel: return
            ko_id = int(sel[0])
            db_exec("UPDATE kitchen_orders SET status='ready' WHERE id=?", (ko_id,))
            load_queue(); items_box.config(state="normal"); items_box.delete("1.0","end"); items_box.config(state="disabled")

        # def print_kot():
        #     sel = tree.selection()
        #     if not sel: return
        #     ko_id = int(sel[0])
        #     hdr = db_query("SELECT order_no, ts, table_no FROM kitchen_orders WHERE id=?", (ko_id,), one=True)
        #     lines = db_query("SELECT item_name, qty FROM kitchen_order_items WHERE ko_id=?", (ko_id,))
        #     out = [f"KITCHEN ORDER #{hdr['order_no']}   {hdr['ts']}   Table {hdr['table_no']}",
        #            "-"*40]
        #     for r in lines:
        #         out.append(f"{r['item_name'][:30]:30} x{r['qty']:>2}")
        #     text = "\n".join(out)
        #     try:
        #         if platform.system().lower().startswith("win"):
        #             with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as f:
        #                 f.write(text); path = f.name
        #             os.startfile(path, "print")
        #         else:
        #             preview = tk.Toplevel(self.root); preview.title("KOT Preview")
        #             t = tk.Text(preview, width=int(50 * SCALE_FACTOR), height=20, font=self.mono_font); t.pack(fill="both", expand=True)
        #             t.insert("end", text); t.config(state="disabled")
        #     except Exception as e:
        #         messagebox.showerror("Print error", str(e))

        # tree.bind("<<TreeviewSelect>>", on_select)
        # btns = tk.Frame(win); btns.pack(fill="x", padx=10, pady=(0,10))
        # tk.Button(btns, text="Refresh", command=load_queue).pack(side="left")
        # tk.Button(btns, text="Mark Ready", command=mark_ready).pack(side="left", padx=6)
        # tk.Button(btns, text="Print KOT", command=print_kot).pack(side="left", padx=6)
        # load_queue()

        def print_kot():
            sel = tree.selection()
            if not sel: return
            ko_id = int(sel[0])
            hdr = db_query("SELECT order_no, ts, table_no FROM kitchen_orders WHERE id=?", (ko_id,), one=True)
            lines = db_query("SELECT item_name, qty FROM kitchen_order_items WHERE ko_id=?", (ko_id,))
            
            # Build KOT text using thermal printer format (similar to bill format)
            out = []
            out.append("    Cream Fresh Natural Icecream")
            out.append(f"    KITCHEN ORDER #{hdr['order_no']}")
            out.append(f"    {hdr['ts']}")
            out.append(f"    TABLE: {hdr['table_no']}")
            out.append("--------------------------------")  # 32 chars
            out.append("Item                         Qty")
            out.append("--------------------------------")
            
            for r in lines:
                item_name = str(r["item_name"])[:24]  # Truncate to 24 chars to leave space for qty
                qty = str(r["qty"])
                line = item_name.ljust(24) + " " + qty.rjust(3)
                out.append(line)
            
            out.append("--------------------------------")
            out.append("")
            out.append("*** KITCHEN COPY ***")
            
            # Add 3 empty lines to push content above cut
            out.append("")
            out.append("")
            out.append("")
            
            # Build the final KOT text
            text_kot = "\n".join(out)
            
            # ESC/POS full cut command
            cut_command = b'\x1D\x56\x00'
            
            try:
                if platform.system().lower().startswith("win"):
                    printer_name = "POS-58"  # Replace with your actual working printer name
                    hPrinter = win32print.OpenPrinter(printer_name)
                    try:
                        hJob = win32print.StartDocPrinter(hPrinter, 1, ("KOT Print Job", None, "RAW"))
                        win32print.StartPagePrinter(hPrinter)
                        win32print.WritePrinter(hPrinter, text_kot.encode('utf-8'))
                        win32print.WritePrinter(hPrinter, cut_command)  # Trigger paper cut
                        win32print.EndPagePrinter(hPrinter)
                        win32print.EndDocPrinter(hPrinter)
                    finally:
                        win32print.ClosePrinter(hPrinter)
                    messagebox.showinfo("Print", "KOT sent to printer.")
            except Exception as e:
                messagebox.showerror("Print error", str(e))
        tree.bind("<<TreeviewSelect>>", on_select)
        btns = tk.Frame(win); btns.pack(fill="x", padx=10, pady=(0,10))
        tk.Button(btns, text="Refresh", command=load_queue).pack(side="left")
        tk.Button(btns, text="Mark Ready", command=mark_ready).pack(side="left", padx=6)
        tk.Button(btns, text="Print KOT", command=print_kot).pack(side="left", padx=6)
        load_queue()    
    # ---------- Checkout ----------
    def _checkout(self):
        cart = self._active_cart()
        if not cart:
            messagebox.showwarning("Empty", "No items in cart.")
            return

        total = sum(pr * qty for _, _, pr, qty in cart)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Get the customer name and mobile from the respective variables
        customer_name = self.customer_name_var.get().strip()
        customer_mobile = self.customer_mobile_var.get().strip()

        # Save the sale with the customer details
        sale_id = db_exec("INSERT INTO sales(ts, total, customer_name, customer_mobile) VALUES (?,?,?,?)", 
                        (ts, total, customer_name, customer_mobile))

        # Save the sale items
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        for iid, nm, pr, qty in cart:
            cur.execute("INSERT INTO sale_items(sale_id, item_id, quantity, line_total) VALUES (?,?,?,?)",
                        (sale_id, iid, qty, pr * qty))
            cur.execute("UPDATE items SET stock = stock - ? WHERE id=?", (qty, iid))
        conn.commit()
        conn.close()

        # Clear the cart and refresh
        self.carts[self.active_table].clear()
        self._refresh_cart()
        self._apply_filters_and_load()
        
        try:
            # Generate PDF bill
            pdf_path = self._generate_pdf_bill(sale_id)
            
            # Show success message with PDF path
            # messagebox.showinfo("Bill Generated", f"Sale completed!\nPDF bill saved at: {pdf_path}")
            
            # Ask if user wants to open the PDF
            # if messagebox.askyesno("Open PDF", "Would you like to open the PDF bill?"):
            #     if platform.system().lower().startswith("win"):
            #         os.startfile(pdf_path)
            #     else:
            #         os.system(f"xdg-open '{pdf_path}'")
                    
        except Exception as e:
            messagebox.showerror("PDF Generation Error", f"Sale completed but PDF generation failed:\n{str(e)}")

        # Show the bill preview
        self._show_bill_preview(sale_id)

        # --- NEW: clear customer details after checkout ---
        self.customer_name_var.set("")
        self.customer_mobile_var.set("")
    # (trace_add handlers will also update per-table stored values)


    # ---------- Bill Preview ----------
    def _wrap_name_lines(self, text, width):
        words = text.split()
        if not words:
            yield ""
            return
        line = words[0]
        for w in words[1:]:
            if len(line) + 1 + len(w) <= width:
                line += " " + w
            else:
                yield line
                line = w
        yield line
        RECEIPT_PAPER = "58mm"
        RECEIPT_CHARS_PER_LINE = 32

    def _receipt_layout(self):
        """Compute receipt column widths based on RECEIPT_CHARS_PER_LINE."""
        w_line = RECEIPT_CHARS_PER_LINE
        # Keep Qty small, Amount larger; leave a small spacing of 2 between name and the numbers.
        w_qty  = 3
        w_rate = 7 #if w_line <= 42 else 12
        w_amt  = 8 #if w_line <= 42 else 14
        # Remaining goes to name column minus 2 spaces for padding
        w_name = max(10, w_line - (w_qty + w_rate + w_amt + 4))
        return 32, 16, 3, 4, 4

    def _wrap_cols(self, text, width):
        words = (text or "").split()
        if not words:
            return [""]
        out, line = [], words[0]
        for w in words[1:]:
            if len(line) + 1 + len(w) <= width:
                line += " " + w
            else:
                out.append(line)
                line = w
        out.append(line)
        return out


    # def _show_bill_preview(self, sale_id: int):
    #     sale = db_query("SELECT * FROM sales WHERE id=?", (sale_id,), one=True)
    #     lines = db_query("""
    #         SELECT i.name, i.price, si.quantity, si.line_total
    #         FROM sale_items si JOIN items i ON i.id=si.item_id
    #         WHERE si.sale_id=?
    #     """, (sale_id,))

    #     win = tk.Toplevel(self.root)
    #     win.title(f"Bill #{sale_id}")
    #     txt = tk.Text(win, width=int(64 * SCALE_FACTOR), height=int(28 * SCALE_FACTOR), font=self.mono_font)
    #     txt.pack(fill="both", expand=True)

    #     W_LINE = int(56 * SCALE_FACTOR)
    #     W_NAME = int(28 * SCALE_FACTOR)
    #     W_QTY  = int(3 * SCALE_FACTOR)
    #     W_RATE = int(8 * SCALE_FACTOR)
    #     W_AMT  = int(12 * SCALE_FACTOR)

    #     cust_name = (self.customer_name_var.get() or "").strip()

    #     txt.insert("end", "         Cream Fresh Natural Icecream\n")
    #     txt.insert("end", f"         Bill #{sale_id}    {sale['ts']}\n")
    #     if cust_name:
    #         txt.insert("end", f"         Customer: {cust_name}\n")
    #     txt.insert("end", "-" * W_LINE + "\n")
    #     txt.insert("end", f"{'Item':{W_NAME}} {'Qty':>{W_QTY}} {'Rate':>{W_RATE}} {'Amount':>{W_AMT}}\n")
    #     txt.insert("end", "-" * W_LINE + "\n")

    #     for li in lines:
    #         name = li["name"]; qty = li["quantity"]; rate = float(li["price"]); amt = float(li["line_total"])
    #         wrapped = list(self._wrap_name_lines(name, W_NAME))
    #         txt.insert("end", f"{wrapped[0]:{W_NAME}} {qty:>{W_QTY}} {rate:>{W_RATE}.2f} {amt:>{W_AMT}.2f}\n")
    #         for cont in wrapped[1:]:
    #             txt.insert("end", f"{cont:{W_NAME}} {'':>{W_QTY}} {'':>{W_RATE}} {'':>{W_AMT}}\n")

    #     txt.insert("end", "-" * W_LINE + "\n")
    #     txt.insert("end", f"{'TOTAL':>{W_NAME + W_QTY + W_RATE + 2}} {float(sale['total']):>{W_AMT}.2f}\n")
    #     txt.config(state="disabled")

    #     act = tk.Frame(win); act.pack(pady=8)
    #     tk.Button(act, text="Print", font=self.text_font,
    #               command=lambda: self._print_text(sale_id, sale, lines, cust_name)).grid(row=0, column=0, padx=6)
    #     tk.Button(act, text="Close", font=self.text_font, command=win.destroy).grid(row=0, column=1, padx=6)

    def _show_bill_preview(self, sale_id: int):
        sale = db_query("SELECT * FROM sales WHERE id=?", (sale_id,), one=True)
        lines = db_query("""
            SELECT i.name, i.price, si.quantity, si.line_total
            FROM sale_items si JOIN items i ON i.id=si.item_id
            WHERE si.sale_id=?
        """, (sale_id,))

        w_line, w_name, w_qty, w_rate, w_amt = self._receipt_layout()

        win = tk.Toplevel(self.root)
        win.title(f"Bill #{sale_id} — {RECEIPT_PAPER}")
        txt = tk.Text(win, width=max(40, w_line+2), height=int(28 * SCALE_FACTOR), font=self.mono_font)
        txt.pack(fill="both", expand=True)

        cust_name = (self.customer_name_var.get() or "").strip()
        cust_mobile = (self.customer_mobile_var.get() or "").strip()

        header = []
        header.append("    Cream Fresh Natural Icecream")
        header.append(f"    Bill #{sale_id}    {sale['ts']}")
        if cust_name:
            header.append(f"    Customer: {cust_name}")
        if cust_mobile:
            header.append(f"    Mobile:   {cust_mobile}")
        header.append("-" * w_line)
        header.append(f"{'Item':{w_name}} {'Qty':>{w_qty}} {'Rate':>{w_rate}} {'Amount':>{w_amt}}")
        header.append("-" * w_line)

        txt.insert("end", "\n".join(header) + "\n")

        for li in lines:
            name = li["name"]; qty = li["quantity"]; rate = float(li["price"]); amt = float(li["line_total"])
            wrapped = self._wrap_cols(name, w_name)
            # first line with numbers
            txt.insert("end", f"{wrapped[0]:{w_name}} {qty:>{w_qty}} {rate:>{w_rate}.2f} {amt:>{w_amt}.2f}\n")
            # continuation lines without numbers
            for cont in wrapped[1:]:
                txt.insert("end", f"{cont:{w_name}} {'':>{w_qty}} {'':>{w_rate}} {'':>{w_amt}}\n")

        txt.insert("end", "-" * w_line + "\n")
        txt.insert("end", f"{'TOTAL':>{w_name + w_qty + w_rate + 2}} {float(sale['total']):>{w_amt}.2f}\n")
        txt.config(state="disabled")

        act = tk.Frame(win); act.pack(pady=8)
        tk.Button(act, text=f"Print to {RECEIPT_PAPER}", font=self.text_font,
                command=lambda: self._print_text(sale_id, sale, lines, cust_name, cust_mobile)).grid(row=0, column=0, padx=6)
        tk.Button(act, text="Close", font=self.text_font, command=win.destroy).grid(row=0, column=1, padx=6)
        
        


    def _print_text(self, sale_id, sale, lines, cust_name, cust_mobile):
        out = []
        out.append("Cream Fresh Natural Icecream")
        out.append("Bill #" + str(sale_id) + "  " + str(sale['ts']))

        if cust_name:
            out.append("Customer: " + str(cust_name))
        if cust_mobile:
            out.append("Mobile: " + str(cust_mobile))

        out.append("--------------------------------")  # 32 chars
        out.append("Item                 Qty Rate Amt")
        out.append("--------------------------------")

        for li in lines:
            name = str(li["name"])[:16]  # Truncate to 16 chars
            qty = str(li["quantity"])
            rate = str(int(float(li["price"])))
            amt = str(int(float(li["line_total"])))

            line = name.ljust(16) + " " + qty.rjust(3) + " " + rate.rjust(4) + " " + amt.rjust(4)
            out.append(line)

        out.append("--------------------------------")
        total_line = "TOTAL".rjust(25) + " " + str(int(float(sale['total']))).rjust(4)
        out.append(total_line)

        out.append("")  # One clean space
        out.append("Thank you for visiting!")
        
        # Add 3 empty lines to push total above cut
        out.append("")
        out.append("")
        out.append("")

        # Build the final bill text
        text_bill = "\n".join(out)

        # ESC/POS full cut command
        cut_command = b'\x1D\x56\x00'

        try:
            if platform.system().lower().startswith("win"):
                printer_name = "POS-58"  # Replace with your actual working printer name
                hPrinter = win32print.OpenPrinter(printer_name)
                try:
                    hJob = win32print.StartDocPrinter(hPrinter, 1, ("Receipt Print Job", None, "RAW"))
                    win32print.StartPagePrinter(hPrinter)
                    win32print.WritePrinter(hPrinter, text_bill.encode('utf-8'))
                    win32print.WritePrinter(hPrinter, cut_command)  # Trigger paper cut
                    win32print.EndPagePrinter(hPrinter)
                    win32print.EndDocPrinter(hPrinter)
                finally:
                    win32print.ClosePrinter(hPrinter)
                messagebox.showinfo("Print", "Sent to printer.")
        except Exception as e:
            messagebox.showerror("Print error", str(e))


    def _show_plain_text(self, title, content):
        w = tk.Toplevel(self.root); w.title(title)
        t = tk.Text(w, width=int(64 * SCALE_FACTOR), height=int(28 * SCALE_FACTOR), font=self.mono_font)
        t.pack(fill="both", expand=True)
        t.insert("end", content); t.config(state="disabled")

    # ---------- Sales ----------


    def _show_sales(self):
        win = tk.Toplevel(self.root)
        win.title("Sales — Pick a Date")

        head = tk.Frame(win)
        head.pack(fill="x", padx=10, pady=8)
        tk.Label(head, text="Date (YYYY-MM-DD):", font=self.text_font).pack(side="left")
        dvar = tk.StringVar(value=date.today().isoformat())
        ent = tk.Entry(head, textvariable=dvar, font=self.text_font, width=12)
        ent.pack(side="left", padx=6)
        
        # Button to load sales data based on selected date
        btn = tk.Button(head, text="Load", font=self.text_font, command=lambda: load(dvar.get()))
        btn.pack(side="left", padx=6)

        # Export button to export sales data to CSV
        export_btn = tk.Button(win, text="Export to CSV", font=self.text_font, command=self._export_sales_to_csv)
        export_btn.pack(side="left", padx=10, pady=10)

        tree = ttk.Treeview(win, columns=("ID", "Time", "Customer Name", "Mobile", "Total"), show="headings", height=18)
        tree.heading("ID", text="Bill #")
        tree.column("ID", width=int(80 * SCALE_FACTOR), anchor="e")
        tree.heading("Time", text="Time")
        tree.column("Time", width=int(200 * SCALE_FACTOR), anchor="w")
        tree.heading("Customer Name", text="Customer Name")
        tree.column("Customer Name", width=int(150 * SCALE_FACTOR), anchor="w")
        tree.heading("Mobile", text="Mobile")
        tree.column("Mobile", width=int(120 * SCALE_FACTOR), anchor="w")
        tree.heading("Total", text="Total")
        tree.column("Total", width=int(120 * SCALE_FACTOR), anchor="e")
        tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        total_lbl = tk.Label(win, text="", font=self.head_font)
        total_lbl.pack(anchor="e", padx=14, pady=(0, 10))

        def load(dstr):
            try:
                _ = datetime.strptime(dstr, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Invalid date", "Use YYYY-MM-DD")
                return
            rows = db_query("SELECT id, ts, customer_name, customer_mobile, total FROM sales WHERE substr(ts,1,10)=? ORDER BY id DESC", (dstr,))
            tree.delete(*tree.get_children())
            g = 0.0
            for r in rows:
                a = float(r["total"])
                g += a
                tree.insert("", "end", values=(r["id"], r["ts"], r["customer_name"], r["customer_mobile"], f"{a:.2f}"))
            total_lbl.config(text=f"Grand Total: ₹{g:.2f}")

        load(dvar.get())

    def _export_sales_to_csv(self):
        """Export sales data to a CSV file based on the selected date range."""
        # Ask user to select the save location for the CSV file
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not file_path:
            return  # User canceled the save dialog

        # Prompt the user for the date range (from date and to date)
        start_date = simpledialog.askstring("Start Date", "Enter start date (YYYY-MM-DD):")
        end_date = simpledialog.askstring("End Date", "Enter end date (YYYY-MM-DD):")
        
        # Validate dates
        if not start_date or not end_date:
            messagebox.showerror("Invalid input", "Please enter valid start and end dates.")
            return

        # Fetch sales data from the database within the date range
        rows = db_query("SELECT id, ts, customer_name, customer_mobile, total FROM sales "
                        "WHERE ts BETWEEN ? AND ? ORDER BY id DESC", (start_date, end_date))

        # Write data to CSV file
        try:
            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                # Write header
                writer.writerow(["Bill #", "Time", "Customer Name", "Mobile", "Total"])
                # Write data rows
                for row in rows:
                    writer.writerow([row["id"], row["ts"], row["customer_name"], row["customer_mobile"], f"{row['total']:.2f}"])
            messagebox.showinfo("Export Complete", f"Sales data has been successfully exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred while exporting the data:\n{str(e)}")
    def _get_week_start_end(self, year, week_number):
        """Get the start and end dates for a given week of the year."""
        from datetime import datetime, timedelta

        # Calculate the start date of the week
        first_day_of_year = datetime(year, 1, 1)
        days_to_add = (week_number - 1) * 7
        start_date = first_day_of_year + timedelta(days=days_to_add)
        start_date = start_date - timedelta(days=start_date.weekday())  # Adjust to Monday of the week

        # End date is 6 days after the start date
        end_date = start_date + timedelta(days=6)

        return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")



    # ---------- Admin ----------
    def _open_admin(self):
        AdminPanel(self)

    # ---------- Tables ----------
    def _switch_table(self, t):
        self.active_table = t
        self._refresh_table_buttons()
        self.customer_name_var.set(self.customer_names.get(self.active_table, ""))
        self.customer_mobile_var.set(self.customer_mobile_numbers.get(self.active_table, ""))
        self._refresh_cart()

    def _refresh_table_buttons(self):
        for i, b in enumerate(self.tbl_buttons, start=1):
            if i == self.active_table:
                b.config(relief="sunken", bg=ACCENT_SOFT, fg=ACCENT)
            else:
                b.config(relief="raised", bg="white", fg="black")

    def _new_sale(self):
        self._cancel_all()

    def _cancel_all(self):
        if not self._active_cart():
            return
        if messagebox.askyesno("Confirm", f"Clear all items on Table {self.active_table}?"):
            self.carts[self.active_table].clear()
            self._refresh_cart()

# ---------- Run ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = POSApp(root)
    root.mainloop()