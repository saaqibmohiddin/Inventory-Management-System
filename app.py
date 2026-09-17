import os
import sqlite3
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

# ============================================================
# OPTIONAL PACKAGES
# ============================================================

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.styles import (
        getSampleStyleSheet,
        ParagraphStyle
    )
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle
    )
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# ============================================================
# SETTINGS
# ============================================================

DB_NAME = "inventory_system.db"

COMPANY_NAME = "MY INVENTORY STORE"
COMPANY_ADDRESS = "India"
COMPANY_PHONE = "Phone: 9876543210"

INVOICE_FOLDER = "invoices"


# ============================================================
# DATABASE CONNECTION
# ============================================================

def db_connect():
    conn = sqlite3.connect(
        DB_NAME,
        timeout=15
    )

    # Important:
    # The database does NOT use restrictive foreign keys
    # between history and products.
    conn.execute("PRAGMA foreign_keys = OFF")

    return conn


# ============================================================
# INITIALIZE DATABASE
# ============================================================

def initialize_database():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # ----------------------------------------------------
        # USERS
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL
            )
        """)

        cursor.execute("""
            INSERT OR IGNORE INTO users
            (
                username,
                password
            )
            VALUES
            (
                'admin',
                'admin123'
            )
        """)


        # ----------------------------------------------------
        # PRODUCTS
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT 'General',
                barcode TEXT NOT NULL DEFAULT '',
                selling_price REAL NOT NULL DEFAULT 0,
                cost_price REAL NOT NULL DEFAULT 0,
                quantity INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # SALES
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL,
                customer_name TEXT NOT NULL DEFAULT '',
                product_id INTEGER,
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                selling_price REAL NOT NULL,
                cost_price REAL NOT NULL,
                subtotal REAL NOT NULL,
                sale_date TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # PURCHASES
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                cost_price REAL NOT NULL,
                total_cost REAL NOT NULL,
                purchase_date TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # STOCK MOVEMENTS
        # ----------------------------------------------------

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stock_movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                product_name TEXT NOT NULL,
                movement_type TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                previous_quantity INTEGER NOT NULL,
                new_quantity INTEGER NOT NULL,
                movement_date TEXT NOT NULL
            )
        """)


        # ----------------------------------------------------
        # BARCODE UNIQUE INDEX
        # ----------------------------------------------------

        cursor.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS
            idx_products_barcode
            ON products(barcode)
            WHERE barcode <> ''
        """)


        conn.commit()

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ============================================================
# LOGIN
# ============================================================

def check_login(username, password):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                username
            FROM users
            WHERE username = ?
            AND password = ?
        """, (
            username,
            password
        ))

        return cursor.fetchone()

    finally:

        conn.close()


# ============================================================
# PRODUCT FUNCTIONS
# ============================================================

def get_all_products():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity
            FROM products
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================

def get_product(product_id):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        return cursor.fetchone()

    finally:

        conn.close()


# ============================================================

def add_product_db(
    name,
    category,
    barcode,
    selling_price,
    cost_price,
    quantity
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        barcode = barcode.strip()

        # Check duplicate barcode
        if barcode:

            cursor.execute("""
                SELECT id
                FROM products
                WHERE barcode = ?
            """, (
                barcode,
            ))

            existing = cursor.fetchone()

            if existing:

                return (
                    False,
                    "Barcode already exists."
                )


        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        cursor.execute("""
            INSERT INTO products
            (
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            name.strip(),
            category.strip(),
            barcode,
            selling_price,
            cost_price,
            quantity,
            now
        ))


        product_id = cursor.lastrowid


        # Stock history
        cursor.execute("""
            INSERT INTO stock_movements
            (
                product_id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            name.strip(),
            "ADD",
            quantity,
            0,
            quantity,
            now
        ))


        conn.commit()

        return (
            True,
            "Product added successfully."
        )


    except sqlite3.IntegrityError as error:

        conn.rollback()

        if "barcode" in str(error).lower():

            return (
                False,
                "Barcode already exists."
            )

        return (
            False,
            str(error)
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================

def update_product_db(
    product_id,
    name,
    category,
    barcode,
    selling_price,
    cost_price,
    quantity
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        old_product = cursor.fetchone()

        if not old_product:

            return (
                False,
                "Product not found."
            )


        old_quantity = int(
            old_product[1]
        )

        barcode = barcode.strip()


        # Check duplicate barcode
        if barcode:

            cursor.execute("""
                SELECT id
                FROM products
                WHERE barcode = ?
                AND id <> ?
            """, (
                barcode,
                product_id
            ))

            existing = cursor.fetchone()

            if existing:

                return (
                    False,
                    "Barcode already exists."
                )


        cursor.execute("""
            UPDATE products
            SET
                name = ?,
                category = ?,
                barcode = ?,
                selling_price = ?,
                cost_price = ?,
                quantity = ?
            WHERE id = ?
        """, (
            name.strip(),
            category.strip(),
            barcode,
            selling_price,
            cost_price,
            quantity,
            product_id
        ))


        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # Save stock movement when quantity changes
        if old_quantity != quantity:

            if quantity > old_quantity:

                movement_type = "UPDATE + STOCK"

            else:

                movement_type = "UPDATE - STOCK"


            cursor.execute("""
                INSERT INTO stock_movements
                (
                    product_id,
                    product_name,
                    movement_type,
                    quantity,
                    previous_quantity,
                    new_quantity,
                    movement_date
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                product_id,
                name.strip(),
                movement_type,
                abs(quantity - old_quantity),
                old_quantity,
                quantity,
                now
            ))


        conn.commit()

        return (
            True,
            "Product updated successfully."
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# DELETE PRODUCT
# ============================================================

def delete_product_db(product_id):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # Find product
        cursor.execute("""
            SELECT
                name
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if not product:

            return (
                False,
                "Product not found."
            )


        product_name = product[0]


        # IMPORTANT:
        # We only delete the product.
        # Sales, purchases and stock history remain.
        # Therefore there is no foreign-key error.

        cursor.execute("""
            DELETE FROM products
            WHERE id = ?
        """, (
            product_id,
        ))


        if cursor.rowcount != 1:

            conn.rollback()

            return (
                False,
                "Product could not be deleted."
            )


        conn.commit()


        return (
            True,
            f"{product_name} deleted successfully."
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# SEARCH / FILTER
# ============================================================

def filter_products_db(
    search_text="",
    category="All Categories",
    stock_status_filter="All Stock"
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        query = """
            SELECT
                id,
                name,
                category,
                barcode,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE 1=1
        """

        parameters = []


        # Search
        if search_text.strip():

            search_pattern = (
                "%"
                +
                search_text.strip()
                +
                "%"
            )

            query += """
                AND (
                    name LIKE ?
                    OR category LIKE ?
                    OR barcode LIKE ?
                )
            """

            parameters.extend([
                search_pattern,
                search_pattern,
                search_pattern
            ])


        # Category
        if (
            category
            and
            category != "All Categories"
        ):

            query += """
                AND category = ?
            """

            parameters.append(
                category
            )


        # Stock
        if stock_status_filter == "Out of Stock":

            query += """
                AND quantity = 0
            """

        elif stock_status_filter == "Low Stock":

            query += """
                AND quantity > 0
                AND quantity <= 5
            """

        elif stock_status_filter == "Normal":

            query += """
                AND quantity > 5
            """


        query += """
            ORDER BY id DESC
        """


        cursor.execute(
            query,
            parameters
        )

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# DASHBOARD
# ============================================================

def dashboard_data():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # Total products
        cursor.execute("""
            SELECT COUNT(*)
            FROM products
        """)

        total_products = (
            cursor.fetchone()[0]
        )


        # Total stock
        cursor.execute("""
            SELECT COALESCE(
                SUM(quantity),
                0
            )
            FROM products
        """)

        total_stock = (
            cursor.fetchone()[0]
        )


        # Low stock
        cursor.execute("""
            SELECT COUNT(*)
            FROM products
            WHERE quantity > 0
            AND quantity <= 5
        """)

        low_stock = (
            cursor.fetchone()[0]
        )


        # Out of stock
        cursor.execute("""
            SELECT COUNT(*)
            FROM products
            WHERE quantity = 0
        """)

        out_of_stock = (
            cursor.fetchone()[0]
        )


        # Inventory value
        cursor.execute("""
            SELECT COALESCE(
                SUM(
                    selling_price * quantity
                ),
                0
            )
            FROM products
        """)

        inventory_value = (
            cursor.fetchone()[0]
        )


        return (
            total_products,
            total_stock,
            low_stock,
            out_of_stock,
            inventory_value
        )

    finally:

        conn.close()


# ============================================================
# STOCK STATUS
# ============================================================

def get_stock_status(quantity):

    quantity = int(quantity)

    if quantity == 0:

        return "Out of Stock"

    if quantity <= 5:

        return "Low Stock"

    return "Normal"


# ============================================================
# SALES
# ============================================================

def next_invoice_number():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        today = datetime.now().strftime(
            "%Y%m%d"
        )

        cursor.execute("""
            SELECT COUNT(*)
            FROM sales
            WHERE sale_date LIKE ?
        """, (
            today + "%",
        ))

        count = (
            cursor.fetchone()[0]
            +
            1
        )

        return (
            f"INV-{today}-{count:04d}"
        )

    finally:

        conn.close()


# ============================================================

def make_sale(
    product_id,
    quantity,
    customer_name
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if not product:

            return (
                False,
                "Product not found."
            )


        (
            product_name,
            selling_price,
            cost_price,
            current_stock
        ) = product


        if quantity <= 0:

            return (
                False,
                "Quantity must be greater than 0."
            )


        if quantity > current_stock:

            return (
                False,
                f"Not enough stock.\n"
                f"Available stock: {current_stock}"
            )


        invoice_number = (
            next_invoice_number()
        )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        subtotal = (
            float(selling_price)
            *
            quantity
        )


        new_stock = (
            current_stock
            -
            quantity
        )


        # Update stock
        cursor.execute("""
            UPDATE products
            SET quantity = ?
            WHERE id = ?
        """, (
            new_stock,
            product_id
        ))


        # Save sale
        cursor.execute("""
            INSERT INTO sales
            (
                invoice_number,
                customer_name,
                product_id,
                product_name,
                quantity,
                selling_price,
                cost_price,
                subtotal,
                sale_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            invoice_number,
            customer_name,
            product_id,
            product_name,
            quantity,
            selling_price,
            cost_price,
            subtotal,
            now
        ))


        sale_id = cursor.lastrowid


        # Stock history
        cursor.execute("""
            INSERT INTO stock_movements
            (
                product_id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            "SALE",
            quantity,
            current_stock,
            new_stock,
            now
        ))


        conn.commit()


        return (
            True,
            {
                "sale_id": sale_id,
                "invoice": invoice_number,
                "customer": customer_name,
                "product": product_name,
                "quantity": quantity,
                "selling_price": float(
                    selling_price
                ),
                "cost_price": float(
                    cost_price
                ),
                "subtotal": subtotal,
                "date": now
            }
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# SALES HISTORY
# ============================================================

def get_sales_history():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                invoice_number,
                customer_name,
                product_name,
                quantity,
                selling_price,
                subtotal,
                cost_price,
                sale_date
            FROM sales
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# SALES SUMMARY
# ============================================================

def sales_summary():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        # Total sales
        cursor.execute("""
            SELECT COALESCE(
                SUM(subtotal),
                0
            )
            FROM sales
        """)

        total_sales = (
            cursor.fetchone()[0]
        )


        # Today's sales
        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        cursor.execute("""
            SELECT COALESCE(
                SUM(subtotal),
                0
            )
            FROM sales
            WHERE sale_date LIKE ?
        """, (
            today + "%",
        ))

        today_sales = (
            cursor.fetchone()[0]
        )


        # Transactions
        cursor.execute("""
            SELECT COUNT(*)
            FROM sales
        """)

        transactions = (
            cursor.fetchone()[0]
        )


        # Profit
        cursor.execute("""
            SELECT COALESCE(
                SUM(
                    subtotal
                    -
                    (
                        quantity
                        *
                        cost_price
                    )
                ),
                0
            )
            FROM sales
        """)

        profit = (
            cursor.fetchone()[0]
        )


        return (
            today_sales,
            total_sales,
            transactions,
            profit
        )

    finally:

        conn.close()


# ============================================================
# PURCHASES
# ============================================================

def make_purchase(
    product_id,
    quantity,
    purchase_cost
):

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                quantity,
                cost_price
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if not product:

            return (
                False,
                "Product not found."
            )


        (
            product_name,
            old_quantity,
            old_cost
        ) = product


        if quantity <= 0:

            return (
                False,
                "Quantity must be greater than 0."
            )


        if purchase_cost <= 0:

            return (
                False,
                "Purchase cost must be greater than 0."
            )


        new_quantity = (
            old_quantity
            +
            quantity
        )


        # Weighted average cost
        new_cost = (
            (
                old_quantity
                *
                float(old_cost)
            )
            +
            (
                quantity
                *
                float(purchase_cost)
            )
        ) / new_quantity


        total_cost = (
            quantity
            *
            float(purchase_cost)
        )


        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # Update product
        cursor.execute("""
            UPDATE products
            SET
                quantity = ?,
                cost_price = ?
            WHERE id = ?
        """, (
            new_quantity,
            new_cost,
            product_id
        ))


        # Save purchase
        cursor.execute("""
            INSERT INTO purchases
            (
                product_id,
                product_name,
                quantity,
                cost_price,
                total_cost,
                purchase_date
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            quantity,
            purchase_cost,
            total_cost,
            now
        ))


        # Stock history
        cursor.execute("""
            INSERT INTO stock_movements
            (
                product_id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            "PURCHASE",
            quantity,
            old_quantity,
            new_quantity,
            now
        ))


        conn.commit()


        return (
            True,
            {
                "product": product_name,
                "quantity": quantity,
                "cost": float(
                    purchase_cost
                ),
                "total": total_cost,
                "date": now
            }
        )


    except Exception as error:

        conn.rollback()

        return (
            False,
            str(error)
        )


    finally:

        conn.close()


# ============================================================
# PURCHASE HISTORY
# ============================================================

def get_purchase_history():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                product_name,
                quantity,
                cost_price,
                total_cost,
                purchase_date
            FROM purchases
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# STOCK HISTORY
# ============================================================

def get_stock_history():

    conn = db_connect()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                product_name,
                movement_type,
                quantity,
                previous_quantity,
                new_quantity,
                movement_date
            FROM stock_movements
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ============================================================
# OPEN FILE
# ============================================================

def open_file(path):

    try:

        if platform.system() == "Windows":

            os.startfile(
                os.path.abspath(path)
            )

        elif platform.system() == "Darwin":

            subprocess.Popen([
                "open",
                path
            ])

        else:

            subprocess.Popen([
                "xdg-open",
                path
            ])

    except Exception as error:

        messagebox.showerror(
            "Open File Error",
            str(error)
        )


# ============================================================
# CREATE PDF INVOICE
# ============================================================

def generate_invoice_pdf(data):

    if not REPORTLAB_AVAILABLE:

        raise RuntimeError(
            "ReportLab is not installed.\n\n"
            "Run:\n"
            "pip install reportlab"
        )


    os.makedirs(
        INVOICE_FOLDER,
        exist_ok=True
    )


    safe_invoice = ""

    for char in data["invoice"]:

        if (
            char.isalnum()
            or char in "-_"
        ):

            safe_invoice += char

        else:

            safe_invoice += "_"


    path = os.path.join(
        INVOICE_FOLDER,
        safe_invoice + ".pdf"
    )


    styles = getSampleStyleSheet()


    title_style = ParagraphStyle(
        "InvoiceTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=20
    )


    right_style = ParagraphStyle(
        "Right",
        parent=styles["Normal"],
        alignment=TA_RIGHT
    )


    document = SimpleDocTemplate(
        path,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )


    story = []


    story.append(
        Paragraph(
            COMPANY_NAME,
            title_style
        )
    )


    story.append(
        Paragraph(
            f"{COMPANY_ADDRESS} | "
            f"{COMPANY_PHONE}",
            styles["Normal"]
        )
    )


    story.append(
        Spacer(1, 15)
    )


    story.append(
        Paragraph(
            f"<b>Invoice:</b> "
            f"{data['invoice']}",
            styles["Normal"]
        )
    )


    story.append(
        Paragraph(
            f"<b>Date:</b> "
            f"{data['date']}",
            styles["Normal"]
        )
    )


    story.append(
        Paragraph(
            f"<b>Customer:</b> "
            f"{data['customer']}",
            styles["Normal"]
        )
    )


    story.append(
        Spacer(1, 15)
    )


    table_data = [
        [
            "Product",
            "Qty",
            "Selling Price",
            "Subtotal"
        ],
        [
            data["product"],
            str(data["quantity"]),
            f"Rs. "
            f"{data['selling_price']:,.2f}",
            f"Rs. "
            f"{data['subtotal']:,.2f}"
        ]
    ]


    table = Table(
        table_data,
        colWidths=[
            230,
            55,
            100,
            100
        ]
    )


    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),
            (
                "ALIGN",
                (1, 1),
                (-1, -1),
                "RIGHT"
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                8
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                8
            )
        ])
    )


    story.append(table)


    story.append(
        Spacer(1, 20)
    )


    story.append(
        Paragraph(
            f"<b>Grand Total: "
            f"Rs. {data['subtotal']:,.2f}</b>",
            right_style
        )
    )


    story.append(
        Spacer(1, 30)
    )


    story.append(
        Paragraph(
            "Thank you for your purchase!",
            ParagraphStyle(
                "Thanks",
                parent=styles["Normal"],
                alignment=TA_CENTER
            )
        )
    )


    document.build(story)


    return path


# ============================================================
# MAIN APPLICATION
# ============================================================

class InventoryApp:

    CATEGORIES = [
        "Electronics",
        "Accessories",
        "Furniture",
        "Stationery",
        "Clothing",
        "Food",
        "Other",
        "General"
    ]


    def __init__(self, root):

        self.root = root

        self.root.title(
            "Inventory Management System"
        )

        self.root.geometry(
            "1280x780"
        )

        self.root.minsize(
            1100,
            700
        )


        self.current_user = None

        self.selected_product_id = None

        self.last_invoice_path = None

        self.last_sale_data = None


        self.style = ttk.Style()

        try:

            self.style.theme_use(
                "clam"
            )

        except Exception:

            pass


        initialize_database()

        self.show_login()


    # ========================================================
    # CLEAR WINDOW
    # ========================================================

    def clear_window(self):

        for widget in self.root.winfo_children():

            widget.destroy()


    # ========================================================
    # LOGIN SCREEN
    # ========================================================

    def show_login(self):

        self.clear_window()

        self.root.geometry(
            "600x500"
        )


        outer = ttk.Frame(
            self.root,
            padding=40
        )

        outer.pack(
            fill="both",
            expand=True
        )


        ttk.Label(
            outer,
            text="INVENTORY MANAGEMENT SYSTEM",
            font=(
                "Arial",
                22,
                "bold"
            )
        ).pack(
            pady=(30, 40)
        )


        card = ttk.LabelFrame(
            outer,
            text="Login",
            padding=30
        )

        card.pack()


        ttk.Label(
            card,
            text="Username"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            pady=8
        )


        self.login_username = (
            tk.StringVar(
                value="admin"
            )
        )


        username_entry = ttk.Entry(
            card,
            textvariable=self.login_username,
            width=30
        )

        username_entry.grid(
            row=1,
            column=0,
            pady=(0, 12)
        )


        ttk.Label(
            card,
            text="Password"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            pady=8
        )


        self.login_password = (
            tk.StringVar(
                value="admin123"
            )
        )


        password_entry = ttk.Entry(
            card,
            textvariable=self.login_password,
            show="*",
            width=30
        )

        password_entry.grid(
            row=3,
            column=0,
            pady=(0, 20)
        )
        ttk.Button(
            card,
            text="CHANGE PASSWORD",
            command=self.change_password
        ).grid(
            row=5,
            column=0,
            sticky="ew",
            pady=8
        )


        

        username_entry.focus()


        password_entry.bind(
            "<Return>",
            lambda event: self.login()
        )


    # ========================================================
    # LOGIN
    # ========================================================

    def login(self):

        username = (
            self.login_username
            .get()
            .strip()
        )

        password = (
            self.login_password
            .get()
        )


        if not username:

            messagebox.showwarning(
                "Login",
                "Please enter username."
            )

            return


        if not password:

            messagebox.showwarning(
                "Login",
                "Please enter password."
            )

            return


        user = check_login(
            username,
            password
        )


        if not user:

            messagebox.showerror(
                "Login Failed",
                "Invalid username or password."
            )

            return


        self.current_user = user[1]

        self.show_main()
    def change_password(self):
        from tkinter import messagebox
        messagebox.showinfo("Change Password", "Change Password option is working.")
       


    # ========================================================
    # MAIN WINDOW
    # ========================================================

    def show_main(self):

        self.clear_window()

        self.root.geometry(
            "1280x780"
        )


        header = ttk.Frame(
            self.root,
            padding=10
        )

        header.pack(
            fill="x"
        )


        ttk.Label(
            header,
            text=COMPANY_NAME,
            font=(
                "Arial",
                22,
                "bold"
            )
        ).pack(
            side="left"
        )


        ttk.Button(
            header,
            text="Logout",
            command=self.logout
        ).pack(
            side="right"
        )


        ttk.Label(
            header,
            text=(
                f"Logged in: "
                f"{self.current_user}"
            )
        ).pack(
            side="right",
            padx=15
        )


        self.notebook = ttk.Notebook(
            self.root
        )

        self.notebook.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 10)
        )


        self.inventory_tab = ttk.Frame(
            self.notebook
        )

        self.sales_tab = ttk.Frame(
            self.notebook
        )

        self.purchase_tab = ttk.Frame(
            self.notebook
        )

        self.alerts_tab = ttk.Frame(
            self.notebook
        )

        self.reports_tab = ttk.Frame(
            self.notebook
        )

        self.stock_history_tab = ttk.Frame(
            self.notebook
        )


        self.notebook.add(
            self.inventory_tab,
            text="Inventory"
        )

        self.notebook.add(
            self.sales_tab,
            text="Sales / Billing"
        )

        self.notebook.add(
            self.purchase_tab,
            text="Purchases"
        )

        self.notebook.add(
            self.alerts_tab,
            text="Alerts"
        )

        self.notebook.add(
            self.reports_tab,
            text="Reports"
        )

        self.notebook.add(
            self.stock_history_tab,
            text="Stock History"
        )


        self.build_inventory_tab()

        self.build_sales_tab()

        self.build_purchase_tab()

        self.build_alerts_tab()

        self.build_reports_tab()

        self.build_stock_history_tab()


        self.refresh_all()


    # ========================================================
    # LOGOUT
    # ========================================================

    def logout(self):

        confirm = messagebox.askyesno(
            "Logout",
            "Are you sure you want to logout?"
        )

        if confirm:

            self.current_user = None

            self.show_login()


    # ========================================================
    # INVENTORY TAB
    # ========================================================

    def build_inventory_tab(self):

        dashboard = ttk.Frame(
            self.inventory_tab,
            padding=10
        )

        dashboard.pack(
            fill="x"
        )


        self.total_products_var = (
            tk.StringVar()
        )

        self.total_stock_var = (
            tk.StringVar()
        )

        self.low_stock_var = (
            tk.StringVar()
        )

        self.out_stock_var = (
            tk.StringVar()
        )

        self.inventory_value_var = (
            tk.StringVar()
        )


        cards = [
            (
                "TOTAL PRODUCTS",
                self.total_products_var
            ),
            (
                "TOTAL STOCK",
                self.total_stock_var
            ),
            (
                "LOW STOCK",
                self.low_stock_var
            ),
            (
                "OUT OF STOCK",
                self.out_stock_var
            ),
            (
                "INVENTORY VALUE",
                self.inventory_value_var
            )
        ]


        for i, (
            title,
            variable
        ) in enumerate(cards):

            dashboard.columnconfigure(
                i,
                weight=1
            )


            frame = ttk.LabelFrame(
                dashboard,
                text=title,
                padding=12
            )

            frame.grid(
                row=0,
                column=i,
                padx=5,
                sticky="nsew"
            )


            ttk.Label(
                frame,
                textvariable=variable,
                font=(
                    "Arial",
                    14,
                    "bold"
                )
            ).pack()


        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        filter_frame = ttk.Frame(
            self.inventory_tab,
            padding=(10, 0)
        )

        filter_frame.pack(
            fill="x"
        )


        ttk.Label(
            filter_frame,
            text="Search:"
        ).pack(
            side="left"
        )


        self.search_var = (
            tk.StringVar()
        )


        ttk.Entry(
            filter_frame,
            textvariable=self.search_var,
            width=28
        ).pack(
            side="left",
            padx=6
        )


        ttk.Button(
            filter_frame,
            text="SEARCH",
            command=self.apply_inventory_filter
        ).pack(
            side="left"
        )


        ttk.Label(
            filter_frame,
            text="Category:"
        ).pack(
            side="left",
            padx=(15, 3)
        )


        self.category_filter_var = (
            tk.StringVar(
                value="All Categories"
            )
        )


        ttk.Combobox(
            filter_frame,
            textvariable=self.category_filter_var,
            values=[
                "All Categories"
            ] + self.CATEGORIES,
            state="readonly",
            width=18
        ).pack(
            side="left"
        )


        ttk.Label(
            filter_frame,
            text="Stock:"
        ).pack(
            side="left",
            padx=(15, 3)
        )


        self.stock_filter_var = (
            tk.StringVar(
                value="All Stock"
            )
        )


        ttk.Combobox(
            filter_frame,
            textvariable=self.stock_filter_var,
            values=[
                "All Stock",
                "Normal",
                "Low Stock",
                "Out of Stock"
            ],
            state="readonly",
            width=15
        ).pack(
            side="left"
        )


        ttk.Button(
            filter_frame,
            text="FILTER",
            command=self.apply_inventory_filter
        ).pack(
            side="left",
            padx=5
        )


        ttk.Button(
            filter_frame,
            text="RESET",
            command=self.reset_inventory_filter
        ).pack(
            side="left"
        )


        # ----------------------------------------------------
        # PRODUCT FORM
        # ----------------------------------------------------

        form = ttk.LabelFrame(
            self.inventory_tab,
            text="Product Details",
            padding=10
        )

        form.pack(
            fill="x",
            padx=10,
            pady=8
        )


        self.product_name_var = (
            tk.StringVar()
        )

        self.product_barcode_var = (
            tk.StringVar()
        )

        self.product_price_var = (
            tk.StringVar()
        )

        self.product_cost_var = (
            tk.StringVar()
        )

        self.product_quantity_var = (
            tk.StringVar()
        )

        self.product_category_var = (
            tk.StringVar(
                value="General"
            )
        )


        fields = [
            (
                "Product Name",
                self.product_name_var
            ),
            (
                "Barcode",
                self.product_barcode_var
            ),
            (
                "Selling Price",
                self.product_price_var
            ),
            (
                "Cost Price",
                self.product_cost_var
            ),
            (
                "Quantity",
                self.product_quantity_var
            )
        ]


        for i, (
            label,
            variable
        ) in enumerate(fields):

            ttk.Label(
                form,
                text=label
            ).grid(
                row=0,
                column=i,
                padx=5,
                sticky="w"
            )


            ttk.Entry(
                form,
                textvariable=variable,
                width=18
            ).grid(
                row=1,
                column=i,
                padx=5,
                pady=4
            )


        ttk.Label(
            form,
            text="Category"
        ).grid(
            row=0,
            column=5,
            padx=5,
            sticky="w"
        )


        ttk.Combobox(
            form,
            textvariable=self.product_category_var,
            values=self.CATEGORIES,
            state="readonly",
            width=18
        ).grid(
            row=1,
            column=5,
            padx=5,
            pady=4
        )


        # ----------------------------------------------------
        # BUTTONS
        # ----------------------------------------------------

        buttons = ttk.Frame(form)

        buttons.grid(
            row=2,
            column=0,
            columnspan=6,
            pady=(8, 0)
        )


        ttk.Button(
            buttons,
            text="ADD PRODUCT",
            command=self.add_product
        ).pack(
            side="left",
            padx=4
        )


        ttk.Button(
            buttons,
            text="UPDATE",
            command=self.update_product
        ).pack(
            side="left",
            padx=4
        )


        ttk.Button(
            buttons,
            text="DELETE",
            command=self.delete_product
        ).pack(
            side="left",
            padx=4
        )


        ttk.Button(
            buttons,
            text="CLEAR",
            command=self.clear_product_form
        ).pack(
            side="left",
            padx=4
        )


        ttk.Button(
            buttons,
            text="EXPORT EXCEL",
            command=self.export_inventory
        ).pack(
            side="left",
            padx=4
        )


        # ----------------------------------------------------
        # PRODUCT TABLE
        # ----------------------------------------------------

        table_frame = ttk.Frame(
            self.inventory_tab,
            padding=10
        )

        table_frame.pack(
            fill="both",
            expand=True
        )


        columns = (
            "id",
            "name",
            "category",
            "barcode",
            "selling_price",
            "cost_price",
            "quantity",
            "status"
        )


        self.product_table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )


        headings = {
            "id": "ID",
            "name": "Name",
            "category": "Category",
            "barcode": "Barcode",
            "selling_price": "Selling Price",
            "cost_price": "Cost Price",
            "quantity": "Quantity",
            "status": "Status"
        }


        widths = {
            "id": 55,
            "name": 190,
            "category": 130,
            "barcode": 140,
            "selling_price": 110,
            "cost_price": 100,
            "quantity": 80,
            "status": 110
        }


        for column in columns:

            self.product_table.heading(
                column,
                text=headings[column]
            )

            self.product_table.column(
                column,
                width=widths[column],
                anchor="center"
            )


        self.product_table.tag_configure(
            "out_of_stock",
            background="#ffcccc"
        )

        self.product_table.tag_configure(
            "low_stock",
            background="#fff2b2"
        )

        self.product_table.tag_configure(
            "normal_stock",
            background="white"
        )


        scroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.product_table.yview
        )


        self.product_table.configure(
            yscrollcommand=scroll.set
        )


        self.product_table.pack(
            side="left",
            fill="both",
            expand=True
        )


        scroll.pack(
            side="right",
            fill="y"
        )


        self.product_table.bind(
            "<<TreeviewSelect>>",
            self.select_product
        )


    # ========================================================
    # VALIDATE PRODUCT
    # ========================================================

    def validate_product_form(self):

        name = (
            self.product_name_var
            .get()
            .strip()
        )

        category = (
            self.product_category_var
            .get()
            .strip()
        )

        barcode = (
            self.product_barcode_var
            .get()
            .strip()
        )

        price_text = (
            self.product_price_var
            .get()
            .strip()
        )

        cost_text = (
            self.product_cost_var
            .get()
            .strip()
        )

        quantity_text = (
            self.product_quantity_var
            .get()
            .strip()
        )


        if not name:

            messagebox.showerror(
                "Invalid Input",
                "Product name cannot be empty."
            )

            return None


        if not category:

            messagebox.showerror(
                "Invalid Input",
                "Please select a category."
            )

            return None


        try:

            price = float(
                price_text
            )

            if price < 0:

                raise ValueError

        except ValueError:

            messagebox.showerror(
                "Invalid Input",
                "Selling price must be a valid number."
            )

            return None


        try:

            cost = float(
                cost_text
            )

            if cost < 0:

                raise ValueError

        except ValueError:

            messagebox.showerror(
                "Invalid Input",
                "Cost price must be a valid number."
            )

            return None


        try:

            quantity = int(
                quantity_text
            )

            if quantity < 0:

                raise ValueError

        except ValueError:

            messagebox.showerror(
                "Invalid Input",
                "Quantity must be a whole number 0 or greater."
            )

            return None


        return (
            name,
            category,
            barcode,
            price,
            cost,
            quantity
        )


    # ========================================================
    # ADD PRODUCT
    # ========================================================

    def add_product(self):

        data = (
            self.validate_product_form()
        )

        if data is None:

            return


        success, message = (
            add_product_db(*data)
        )


        if success:

            messagebox.showinfo(
                "Success",
                message
            )

            self.clear_product_form()

            self.refresh_all()

        else:

            messagebox.showerror(
                "Add Product Failed",
                message
            )


    # ========================================================
    # SELECT PRODUCT
    # ========================================================

    def select_product(self, event=None):

        selected = (
            self.product_table.selection()
        )

        if not selected:

            return


        values = (
            self.product_table.item(
                selected[0],
                "values"
            )
        )


        if not values:

            return


        try:

            product_id = int(
                values[0]
            )

        except Exception:

            return


        product = get_product(
            product_id
        )


        if not product:

            return


        self.selected_product_id = (
            product[0]
        )


        self.product_name_var.set(
            product[1]
        )

        self.product_category_var.set(
            product[2]
        )

        self.product_barcode_var.set(
            product[3]
        )

        self.product_price_var.set(
            str(product[4])
        )

        self.product_cost_var.set(
            str(product[5])
        )

        self.product_quantity_var.set(
            str(product[6])
        )


    # ========================================================
    # UPDATE PRODUCT
    # ========================================================

    def update_product(self):

        if self.selected_product_id is None:

            messagebox.showwarning(
                "Select Product",
                "Please select a product first."
            )

            return


        data = (
            self.validate_product_form()
        )


        if data is None:

            return


        success, message = (
            update_product_db(
                self.selected_product_id,
                *data
            )
        )


        if success:

            messagebox.showinfo(
                "Success",
                message
            )

            self.clear_product_form()

            self.refresh_all()

        else:

            messagebox.showerror(
                "Update Failed",
                message
            )


    # ========================================================
    # DELETE PRODUCT
    # ========================================================

    def delete_product(self):

        # Try table selection if no ID
        if self.selected_product_id is None:

            selected = (
                self.product_table.selection()
            )

            if selected:

                values = (
                    self.product_table.item(
                        selected[0],
                        "values"
                    )
                )

                if values:

                    try:

                        self.selected_product_id = (
                            int(values[0])
                        )

                    except Exception:

                        self.selected_product_id = None


        if self.selected_product_id is None:

            messagebox.showwarning(
                "Select Product",
                "Please select a product from the table first."
            )

            return


        product = get_product(
            self.selected_product_id
        )


        if not product:

            messagebox.showerror(
                "Delete Failed",
                "Product no longer exists."
            )

            self.refresh_all()

            return


        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Product: {product[1]}\n\n"
            "Are you sure you want to delete this product?"
        )


        if not confirm:

            return


        success, message = (
            delete_product_db(
                self.selected_product_id
            )
        )


        if success:

            messagebox.showinfo(
                "Delete Successful",
                message
            )

            self.clear_product_form()

            self.refresh_all()

        else:

            messagebox.showerror(
                "Delete Failed",
                message
            )


    # ========================================================
    # CLEAR
    # ========================================================

    def clear_product_form(self):

        self.selected_product_id = None

        self.product_name_var.set("")

        self.product_barcode_var.set("")

        self.product_price_var.set("")

        self.product_cost_var.set("")

        self.product_quantity_var.set("")

        self.product_category_var.set(
            "General"
        )


        if hasattr(
            self,
            "product_table"
        ):

            for item in (
                self.product_table.selection()
            ):

                self.product_table.selection_remove(
                    item
                )


    # ========================================================
    # REFRESH PRODUCT TABLE
    # ========================================================

    def refresh_products(self, rows=None):

        if rows is None:

            rows = get_all_products()


        for item in (
            self.product_table.get_children()
        ):

            self.product_table.delete(
                item
            )


        for product in rows:

            quantity = int(
                product[6]
            )

            status = get_stock_status(
                quantity
            )


            if status == "Out of Stock":

                tag = "out_of_stock"

            elif status == "Low Stock":

                tag = "low_stock"

            else:

                tag = "normal_stock"


            self.product_table.insert(
                "",
                "end",
                values=(
                    product[0],
                    product[1],
                    product[2],
                    product[3],
                    f"Rs. {float(product[4]):,.2f}",
                    f"Rs. {float(product[5]):,.2f}",
                    quantity,
                    status
                ),
                tags=(tag,)
            )


    # ========================================================
    # SEARCH / FILTER
    # ========================================================

    def apply_inventory_filter(self):

        rows = filter_products_db(
            self.search_var.get(),
            self.category_filter_var.get(),
            self.stock_filter_var.get()
        )

        self.refresh_products(
            rows
        )


    # ========================================================

    def reset_inventory_filter(self):

        self.search_var.set("")

        self.category_filter_var.set(
            "All Categories"
        )

        self.stock_filter_var.set(
            "All Stock"
        )

        self.refresh_products()


    # ========================================================
    # DASHBOARD REFRESH
    # ========================================================

    def refresh_inventory_dashboard(self):

        (
            total_products,
            total_stock,
            low_stock,
            out_of_stock,
            inventory_value
        ) = dashboard_data()


        self.total_products_var.set(
            str(total_products)
        )

        self.total_stock_var.set(
            str(total_stock)
        )

        self.low_stock_var.set(
            str(low_stock)
        )

        self.out_stock_var.set(
            str(out_of_stock)
        )

        self.inventory_value_var.set(
            f"Rs. {inventory_value:,.2f}"
        )


    # ========================================================
    # INVENTORY EXCEL
    # ========================================================

    def export_inventory(self):

        if not OPENPYXL_AVAILABLE:

            messagebox.showerror(
                "Excel Error",
                "OpenPyXL is not installed.\n\n"
                "Run:\n"
                "pip install openpyxl"
            )

            return


        rows = get_all_products()


        if not rows:

            messagebox.showwarning(
                "No Data",
                "There are no products to export."
            )

            return


        path = filedialog.asksaveasfilename(
            title="Save Inventory Report",
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx"
                )
            ],
            initialfile="inventory_report.xlsx"
        )


        if not path:

            return


        try:

            workbook = Workbook()

            sheet = workbook.active

            sheet.title = "Inventory"


            headers = [
                "ID",
                "Product Name",
                "Category",
                "Barcode",
                "Selling Price",
                "Cost Price",
                "Quantity",
                "Status",
                "Stock Value"
            ]


            sheet.append(
                headers
            )


            for cell in sheet[1]:

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )


            total = 0


            for product in rows:

                stock_value = (
                    float(product[5])
                    *
                    int(product[6])
                )


                total += stock_value


                sheet.append([
                    product[0],
                    product[1],
                    product[2],
                    product[3],
                    product[4],
                    product[5],
                    product[6],
                    get_stock_status(
                        product[6]
                    ),
                    stock_value
                ])


            sheet.append([])


            sheet.append([
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "TOTAL STOCK VALUE",
                total
            ])


            for column in sheet.columns:

                maximum = 0

                letter = (
                    column[0]
                    .column_letter
                )


                for cell in column:

                    try:

                        maximum = max(
                            maximum,
                            len(
                                str(
                                    cell.value
                                )
                            )
                        )

                    except Exception:

                        pass


                sheet.column_dimensions[
                    letter
                ].width = min(
                    maximum + 2,
                    35
                )


            workbook.save(
                path
            )


            messagebox.showinfo(
                "Export Successful",
                f"Inventory report saved to:\n{path}"
            )


        except Exception as error:

            messagebox.showerror(
                "Export Error",
                str(error)
            )


    # ========================================================
    # SALES TAB
    # ========================================================

    def build_sales_tab(self):

        dashboard = ttk.Frame(
            self.sales_tab,
            padding=10
        )

        dashboard.pack(
            fill="x"
        )


        self.today_sales_var = (
            tk.StringVar()
        )

        self.total_sales_var = (
            tk.StringVar()
        )

        self.transaction_var = (
            tk.StringVar()
        )

        self.profit_var = (
            tk.StringVar()
        )


        cards = [
            (
                "TODAY'S SALES",
                self.today_sales_var
            ),
            (
                "TOTAL SALES",
                self.total_sales_var
            ),
            (
                "TRANSACTIONS",
                self.transaction_var
            ),
            (
                "TOTAL PROFIT",
                self.profit_var
            )
        ]


        for i, (
            title,
            variable
        ) in enumerate(cards):

            dashboard.columnconfigure(
                i,
                weight=1
            )


            frame = ttk.LabelFrame(
                dashboard,
                text=title,
                padding=12
            )

            frame.grid(
                row=0,
                column=i,
                padx=5,
                sticky="nsew"
            )


            ttk.Label(
                frame,
                textvariable=variable,
                font=(
                    "Arial",
                    13,
                    "bold"
                )
            ).pack()


        # ----------------------------------------------------
        # SALES FORM
        # ----------------------------------------------------

        form = ttk.LabelFrame(
            self.sales_tab,
            text="Create Sale / Bill",
            padding=10
        )

        form.pack(
            fill="x",
            padx=10,
            pady=5
        )


        self.sale_customer_var = (
            tk.StringVar()
        )

        self.sale_product_var = (
            tk.StringVar()
        )

        self.sale_quantity_var = (
            tk.StringVar()
        )

        self.sale_total_var = (
            tk.StringVar(
                value="Rs. 0.00"
            )
        )


        ttk.Label(
            form,
            text="Customer Name"
        ).grid(
            row=0,
            column=0,
            sticky="w"
        )


        ttk.Entry(
            form,
            textvariable=self.sale_customer_var,
            width=25
        ).grid(
            row=1,
            column=0,
            padx=5
        )


        ttk.Label(
            form,
            text="Product"
        ).grid(
            row=0,
            column=1,
            sticky="w"
        )


        self.sale_product_combo = ttk.Combobox(
            form,
            textvariable=self.sale_product_var,
            state="readonly",
            width=45
        )


        self.sale_product_combo.grid(
            row=1,
            column=1,
            padx=5
        )


        self.sale_product_combo.bind(
            "<<ComboboxSelected>>",
            self.calculate_sale_total
        )


        ttk.Label(
            form,
            text="Quantity"
        ).grid(
            row=0,
            column=2,
            sticky="w"
        )


        quantity_entry = ttk.Entry(
            form,
            textvariable=self.sale_quantity_var,
            width=12
        )


        quantity_entry.grid(
            row=1,
            column=2,
            padx=5
        )


        quantity_entry.bind(
            "<KeyRelease>",
            self.calculate_sale_total
        )


        ttk.Label(
            form,
            text="Total"
        ).grid(
            row=0,
            column=3,
            sticky="w"
        )


        ttk.Label(
            form,
            textvariable=self.sale_total_var,
            font=(
                "Arial",
                12,
                "bold"
            )
        ).grid(
            row=1,
            column=3,
            padx=10
        )


        ttk.Button(
            form,
            text="SELL / CREATE BILL",
            command=self.sell_product
        ).grid(
            row=1,
            column=4,
            padx=8
        )


        ttk.Button(
            form,
            text="OPEN LAST INVOICE",
            command=self.open_last_invoice
        ).grid(
            row=1,
            column=5,
            padx=5
        )


        ttk.Button(
            form,
            text="PRINT LAST BILL",
            command=self.print_last_invoice
        ).grid(
            row=1,
            column=6,
            padx=5
        )


        # ----------------------------------------------------
        # SALES TABLE
        # ----------------------------------------------------

        table_frame = ttk.Frame(
            self.sales_tab,
            padding=10
        )

        table_frame.pack(
            fill="both",
            expand=True
        )


        columns = (
            "id",
            "invoice",
            "customer",
            "product",
            "quantity",
            "price",
            "total",
            "cost",
            "date"
        )


        self.sales_table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings"
        )


        headings = {
            "id": "ID",
            "invoice": "Invoice",
            "customer": "Customer",
            "product": "Product",
            "quantity": "Qty",
            "price": "Selling Price",
            "total": "Total",
            "cost": "Cost Price",
            "date": "Date & Time"
        }


        widths = {
            "id": 50,
            "invoice": 145,
            "customer": 130,
            "product": 160,
            "quantity": 60,
            "price": 100,
            "total": 100,
            "cost": 90,
            "date": 150
        }


        for column in columns:

            self.sales_table.heading(
                column,
                text=headings[column]
            )

            self.sales_table.column(
                column,
                width=widths[column],
                anchor="center"
            )


        scroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.sales_table.yview
        )


        self.sales_table.configure(
            yscrollcommand=scroll.set
        )


        self.sales_table.pack(
            side="left",
            fill="both",
            expand=True
        )


        scroll.pack(
            side="right",
            fill="y"
        )


        bottom = ttk.Frame(
            self.sales_tab,
            padding=10
        )

        bottom.pack(
            fill="x"
        )


        ttk.Button(
            bottom,
            text="EXPORT SALES TO EXCEL",
            command=self.export_sales
        ).pack(
            side="left"
        )


        ttk.Button(
            bottom,
            text="REFRESH",
            command=self.refresh_sales
        ).pack(
            side="left",
            padx=8
        )


    # ========================================================
    # SALES PRODUCT LIST
    # ========================================================

    def refresh_sale_products(self):

        products = get_all_products()

        values = []


        for product in products:

            if int(product[6]) > 0:

                values.append(
                    f"{product[0]} - "
                    f"{product[1]} | "
                    f"Rs. {float(product[4]):,.2f} | "
                    f"Stock: {product[6]}"
                )


        self.sale_product_combo[
            "values"
        ] = values


        if (
            self.sale_product_var.get()
            not in values
        ):

            self.sale_product_var.set("")


    # ========================================================

    def get_sale_product_id(self):

        value = (
            self.sale_product_var
            .get()
        )


        if not value:

            return None


        try:

            return int(
                value.split(
                    " - "
                )[0]
            )

        except Exception:

            return None


    # ========================================================

    def calculate_sale_total(
        self,
        event=None
    ):

        product_id = (
            self.get_sale_product_id()
        )


        if product_id is None:

            self.sale_total_var.set(
                "Rs. 0.00"
            )

            return


        try:

            quantity = int(
                self.sale_quantity_var
                .get()
            )

        except Exception:

            self.sale_total_var.set(
                "Rs. 0.00"
            )

            return


        product = get_product(
            product_id
        )


        if not product or quantity <= 0:

            self.sale_total_var.set(
                "Rs. 0.00"
            )

            return


        total = (
            float(product[4])
            *
            quantity
        )


        self.sale_total_var.set(
            f"Rs. {total:,.2f}"
        )


    # ========================================================
    # SELL PRODUCT
    # ========================================================

    def sell_product(self):

        product_id = (
            self.get_sale_product_id()
        )


        if product_id is None:

            messagebox.showwarning(
                "Sale",
                "Please select a product."
            )

            return


        try:

            quantity = int(
                self.sale_quantity_var
                .get()
            )

        except Exception:

            messagebox.showerror(
                "Sale",
                "Enter a valid quantity."
            )

            return


        if quantity <= 0:

            messagebox.showerror(
                "Sale",
                "Quantity must be greater than 0."
            )

            return


        customer = (
            self.sale_customer_var
            .get()
            .strip()
        )


        if not customer:

            customer = (
                "Walk-in Customer"
            )


        product = get_product(
            product_id
        )


        if not product:

            messagebox.showerror(
                "Sale",
                "Product not found."
            )

            return


        subtotal = (
            float(product[4])
            *
            quantity
        )


        profit = (
            float(product[4])
            -
            float(product[5])
        ) * quantity


        confirm = messagebox.askyesno(
            "Confirm Sale",
            f"Customer: {customer}\n"
            f"Product: {product[1]}\n"
            f"Quantity: {quantity}\n"
            f"Selling Price: "
            f"Rs. {float(product[4]):,.2f}\n"
            f"Subtotal: "
            f"Rs. {subtotal:,.2f}\n"
            f"Profit: "
            f"Rs. {profit:,.2f}\n\n"
            "Confirm sale?"
        )


        if not confirm:

            return


        success, result = make_sale(
            product_id,
            quantity,
            customer
        )


        if not success:

            messagebox.showerror(
                "Sale Failed",
                result
            )

            return


        self.last_sale_data = result


        try:

            self.last_invoice_path = (
                generate_invoice_pdf(
                    result
                )
            )

        except Exception as error:

            self.last_invoice_path = None

            messagebox.showwarning(
                "Invoice Warning",
                "Sale was completed, "
                "but PDF invoice could not "
                f"be generated.\n\n{error}"
            )


        messagebox.showinfo(
            "Sale Successful",
            f"Sale completed successfully!\n\n"
            f"Invoice: {result['invoice']}\n"
            f"Customer: {result['customer']}\n"
            f"Product: {result['product']}\n"
            f"Quantity: {result['quantity']}\n"
            f"Total: "
            f"Rs. {result['subtotal']:,.2f}"
        )


        self.sale_customer_var.set("")

        self.sale_product_var.set("")

        self.sale_quantity_var.set("")

        self.sale_total_var.set(
            "Rs. 0.00"
        )


        self.refresh_all()


    # ========================================================
    # SALES HISTORY
    # ========================================================

    def refresh_sales(self):

        for item in (
            self.sales_table
            .get_children()
        ):

            self.sales_table.delete(
                item
            )


        for sale in get_sales_history():

            self.sales_table.insert(
                "",
                "end",
                values=(
                    sale[0],
                    sale[1],
                    sale[2],
                    sale[3],
                    sale[4],
                    f"Rs. {float(sale[5]):,.2f}",
                    f"Rs. {float(sale[6]):,.2f}",
                    f"Rs. {float(sale[7]):,.2f}",
                    sale[8]
                )
            )


        (
            today,
            total,
            transactions,
            profit
        ) = sales_summary()


        self.today_sales_var.set(
            f"Rs. {today:,.2f}"
        )

        self.total_sales_var.set(
            f"Rs. {total:,.2f}"
        )

        self.transaction_var.set(
            str(transactions)
        )

        self.profit_var.set(
            f"Rs. {profit:,.2f}"
        )


    # ========================================================
    # OPEN INVOICE
    # ========================================================

    def open_last_invoice(self):

        if not self.last_invoice_path:

            messagebox.showinfo(
                "Invoice",
                "No invoice has been generated "
                "in this session."
            )

            return


        if not os.path.exists(
            self.last_invoice_path
        ):

            messagebox.showerror(
                "Invoice",
                "Invoice file no longer exists."
            )

            return


        open_file(
            self.last_invoice_path
        )


    # ========================================================
    # PRINT INVOICE
    # ========================================================

    def print_last_invoice(self):

        if not self.last_invoice_path:

            messagebox.showinfo(
                "Print",
                "No invoice has been generated "
                "in this session."
            )

            return


        if not os.path.exists(
            self.last_invoice_path
        ):

            messagebox.showerror(
                "Print",
                "Invoice file no longer exists."
            )

            return


        try:

            if platform.system() == "Windows":

                os.startfile(
                    os.path.abspath(
                        self.last_invoice_path
                    ),
                    "print"
                )

            else:

                subprocess.Popen([
                    "lp",
                    self.last_invoice_path
                ])


        except Exception as error:

            messagebox.showerror(
                "Print Error",
                str(error)
            )


    # ========================================================
    # EXPORT SALES
    # ========================================================

    def export_sales(self):

        if not OPENPYXL_AVAILABLE:

            messagebox.showerror(
                "Excel Error",
                "OpenPyXL is not installed.\n\n"
                "Run:\n"
                "pip install openpyxl"
            )

            return


        rows = get_sales_history()


        if not rows:

            messagebox.showwarning(
                "No Sales",
                "There are no sales to export."
            )

            return


        path = filedialog.asksaveasfilename(
            title="Save Sales Report",
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx"
                )
            ],
            initialfile="sales_report.xlsx"
        )


        if not path:

            return


        try:

            workbook = Workbook()

            sheet = workbook.active

            sheet.title = "Sales Report"


            headers = [
                "Sale ID",
                "Invoice",
                "Customer",
                "Product",
                "Quantity",
                "Selling Price",
                "Total Sales",
                "Cost Price",
                "Profit",
                "Sale Date & Time"
            ]


            sheet.append(
                headers
            )


            for cell in sheet[1]:

                cell.font = Font(
                    bold=True
                )


            total_profit = 0


            for sale in rows:

                profit = (
                    float(sale[6])
                    -
                    (
                        int(sale[4])
                        *
                        float(sale[7])
                    )
                )


                total_profit += profit


                sheet.append([
                    sale[0],
                    sale[1],
                    sale[2],
                    sale[3],
                    sale[4],
                    sale[5],
                    sale[6],
                    sale[7],
                    profit,
                    sale[8]
                ])


            sheet.append([])


            sheet.append([
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "TOTAL PROFIT",
                total_profit
            ])


            workbook.save(
                path
            )


            messagebox.showinfo(
                "Export Successful",
                f"Sales report saved to:\n{path}"
            )


        except Exception as error:

            messagebox.showerror(
                "Export Error",
                str(error)
            )


    # ========================================================
    # PURCHASE TAB
    # ========================================================

    def build_purchase_tab(self):

        form = ttk.LabelFrame(
            self.purchase_tab,
            text="Purchase / Restock",
            padding=15
        )

        form.pack(
            fill="x",
            padx=10,
            pady=10
        )


        self.purchase_product_var = (
            tk.StringVar()
        )

        self.purchase_quantity_var = (
            tk.StringVar()
        )

        self.purchase_cost_var = (
            tk.StringVar()
        )

        self.purchase_total_var = (
            tk.StringVar(
                value="Rs. 0.00"
            )
        )


        ttk.Label(
            form,
            text="Product"
        ).grid(
            row=0,
            column=0,
            sticky="w"
        )


        self.purchase_product_combo = ttk.Combobox(
            form,
            textvariable=self.purchase_product_var,
            state="readonly",
            width=50
        )


        self.purchase_product_combo.grid(
            row=1,
            column=0,
            padx=5
        )


        ttk.Label(
            form,
            text="Quantity"
        ).grid(
            row=0,
            column=1,
            sticky="w"
        )


        quantity_entry = ttk.Entry(
            form,
            textvariable=self.purchase_quantity_var,
            width=12
        )


        quantity_entry.grid(
            row=1,
            column=1,
            padx=5
        )


        quantity_entry.bind(
            "<KeyRelease>",
            self.calculate_purchase_total
        )


        ttk.Label(
            form,
            text="Purchase Cost / Unit"
        ).grid(
            row=0,
            column=2,
            sticky="w"
        )


        cost_entry = ttk.Entry(
            form,
            textvariable=self.purchase_cost_var,
            width=15
        )


        cost_entry.grid(
            row=1,
            column=2,
            padx=5
        )


        cost_entry.bind(
            "<KeyRelease>",
            self.calculate_purchase_total
        )


        ttk.Label(
            form,
            text="Total Cost"
        ).grid(
            row=0,
            column=3,
            sticky="w"
        )


        ttk.Label(
            form,
            textvariable=self.purchase_total_var,
            font=(
                "Arial",
                12,
                "bold"
            )
        ).grid(
            row=1,
            column=3,
            padx=10
        )


        ttk.Button(
            form,
            text="ADD STOCK",
            command=self.add_purchase
        ).grid(
            row=1,
            column=4,
            padx=10
        )


        table_frame = ttk.Frame(
            self.purchase_tab,
            padding=10
        )

        table_frame.pack(
            fill="both",
            expand=True
        )


        columns = (
            "id",
            "product",
            "quantity",
            "cost",
            "total",
            "date"
        )


        self.purchase_table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings"
        )


        headings = {
            "id": "ID",
            "product": "Product",
            "quantity": "Quantity",
            "cost": "Cost / Unit",
            "total": "Total Cost",
            "date": "Purchase Date & Time"
        }


        for column in columns:

            self.purchase_table.heading(
                column,
                text=headings[column]
            )


        self.purchase_table.column(
            "id",
            width=60,
            anchor="center"
        )

        self.purchase_table.column(
            "product",
            width=260
        )

        self.purchase_table.column(
            "quantity",
            width=100,
            anchor="center"
        )

        self.purchase_table.column(
            "cost",
            width=120,
            anchor="center"
        )

        self.purchase_table.column(
            "total",
            width=120,
            anchor="center"
        )

        self.purchase_table.column(
            "date",
            width=180,
            anchor="center"
        )


        scroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.purchase_table.yview
        )


        self.purchase_table.configure(
            yscrollcommand=scroll.set
        )


        self.purchase_table.pack(
            side="left",
            fill="both",
            expand=True
        )


        scroll.pack(
            side="right",
            fill="y"
        )


    # ========================================================
    # PURCHASE PRODUCT LIST
    # ========================================================

    def refresh_purchase_products(self):

        products = get_all_products()

        values = []


        for product in products:

            values.append(
                f"{product[0]} - "
                f"{product[1]} | "
                f"Current Stock: {product[6]}"
            )


        self.purchase_product_combo[
            "values"
        ] = values


        if (
            self.purchase_product_var.get()
            not in values
        ):

            self.purchase_product_var.set("")


    # ========================================================

    def get_purchase_product_id(self):

        value = (
            self.purchase_product_var
            .get()
        )


        if not value:

            return None


        try:

            return int(
                value.split(
                    " - "
                )[0]
            )

        except Exception:

            return None


    # ========================================================

    def calculate_purchase_total(
        self,
        event=None
    ):

        try:

            quantity = int(
                self.purchase_quantity_var
                .get()
            )

            cost = float(
                self.purchase_cost_var
                .get()
            )


            if (
                quantity > 0
                and
                cost > 0
            ):

                total = (
                    quantity
                    *
                    cost
                )


                self.purchase_total_var.set(
                    f"Rs. {total:,.2f}"
                )

            else:

                self.purchase_total_var.set(
                    "Rs. 0.00"
                )


        except Exception:

            self.purchase_total_var.set(
                "Rs. 0.00"
            )


    # ========================================================
    # ADD PURCHASE
    # ========================================================

    def add_purchase(self):

        product_id = (
            self.get_purchase_product_id()
        )


        if product_id is None:

            messagebox.showwarning(
                "Purchase",
                "Please select a product."
            )

            return


        try:

            quantity = int(
                self.purchase_quantity_var
                .get()
            )

        except Exception:

            messagebox.showerror(
                "Purchase",
                "Enter a valid quantity."
            )

            return


        try:

            cost = float(
                self.purchase_cost_var
                .get()
            )

        except Exception:

            messagebox.showerror(
                "Purchase",
                "Enter a valid purchase cost."
            )

            return


        if quantity <= 0:

            messagebox.showerror(
                "Purchase",
                "Quantity must be greater than 0."
            )

            return


        if cost <= 0:

            messagebox.showerror(
                "Purchase",
                "Cost must be greater than 0."
            )

            return


        product = get_product(
            product_id
        )


        if not product:

            messagebox.showerror(
                "Purchase",
                "Product not found."
            )

            return


        total = (
            quantity
            *
            cost
        )


        confirm = messagebox.askyesno(
            "Confirm Purchase",
            f"Product: {product[1]}\n"
            f"Quantity: {quantity}\n"
            f"Cost / Unit: "
            f"Rs. {cost:,.2f}\n"
            f"Total Cost: "
            f"Rs. {total:,.2f}\n\n"
            "Add this stock?"
        )


        if not confirm:

            return


        success, result = make_purchase(
            product_id,
            quantity,
            cost
        )


        if success:

            messagebox.showinfo(
                "Purchase Successful",
                f"Stock added successfully!\n\n"
                f"Product: "
                f"{result['product']}\n"
                f"Quantity: "
                f"{result['quantity']}\n"
                f"Total Cost: "
                f"Rs. {result['total']:,.2f}"
            )


            self.purchase_product_var.set("")

            self.purchase_quantity_var.set("")

            self.purchase_cost_var.set("")

            self.purchase_total_var.set(
                "Rs. 0.00"
            )


            self.refresh_all()


        else:

            messagebox.showerror(
                "Purchase Failed",
                result
            )


    # ========================================================
    # PURCHASE HISTORY
    # ========================================================

    def refresh_purchases(self):

        for item in (
            self.purchase_table
            .get_children()
        ):

            self.purchase_table.delete(
                item
            )


        for row in get_purchase_history():

            self.purchase_table.insert(
                "",
                "end",
                values=(
                    row[0],
                    row[1],
                    row[2],
                    f"Rs. {float(row[3]):,.2f}",
                    f"Rs. {float(row[4]):,.2f}",
                    row[5]
                )
            )


    # ========================================================
    # ALERTS
    # ========================================================

    def build_alerts_tab(self):

        top = ttk.Frame(
            self.alerts_tab,
            padding=10
        )

        top.pack(
            fill="x"
        )


        ttk.Button(
            top,
            text="REFRESH ALERTS",
            command=self.refresh_alerts
        ).pack(
            side="left"
        )


        ttk.Label(
            top,
            text="Low Stock = 1 to 5 | Out of Stock = 0"
        ).pack(
            side="left",
            padx=20
        )


        frame = ttk.Frame(
            self.alerts_tab,
            padding=10
        )

        frame.pack(
            fill="both",
            expand=True
        )


        columns = (
            "id",
            "name",
            "category",
            "barcode",
            "quantity",
            "status"
        )


        self.alert_table = ttk.Treeview(
            frame,
            columns=columns,
            show="headings"
        )


        headings = {
            "id": "ID",
            "name": "Product",
            "category": "Category",
            "barcode": "Barcode",
            "quantity": "Quantity",
            "status": "Status"
        }


        for column in columns:

            self.alert_table.heading(
                column,
                text=headings[column]
            )


        self.alert_table.column(
            "id",
            width=60,
            anchor="center"
        )

        self.alert_table.column(
            "name",
            width=220
        )

        self.alert_table.column(
            "category",
            width=150
        )

        self.alert_table.column(
            "barcode",
            width=160
        )

        self.alert_table.column(
            "quantity",
            width=100,
            anchor="center"
        )

        self.alert_table.column(
            "status",
            width=140,
            anchor="center"
        )


        self.alert_table.tag_configure(
            "out",
            background="#ffcccc"
        )

        self.alert_table.tag_configure(
            "low",
            background="#fff2b2"
        )


        self.alert_table.pack(
            fill="both",
            expand=True
        )


    # ========================================================

    def refresh_alerts(self):

        for item in (
            self.alert_table
            .get_children()
        ):

            self.alert_table.delete(
                item
            )


        for product in get_all_products():

            status = get_stock_status(
                product[6]
            )


            if status == "Normal":

                continue


            if status == "Out of Stock":

                tag = "out"

            else:

                tag = "low"


            self.alert_table.insert(
                "",
                "end",
                values=(
                    product[0],
                    product[1],
                    product[2],
                    product[3],
                    product[6],
                    status
                ),
                tags=(tag,)
            )


    # ========================================================
    # REPORTS
    # ========================================================

    def build_reports_tab(self):

        top = ttk.Frame(
            self.reports_tab,
            padding=15
        )

        top.pack(
            fill="x"
        )


        ttk.Button(
            top,
            text="REFRESH REPORTS",
            command=self.refresh_reports
        ).pack(
            side="left"
        )


        ttk.Button(
            top,
            text="EXPORT INVENTORY",
            command=self.export_inventory
        ).pack(
            side="left",
            padx=8
        )


        ttk.Button(
            top,
            text="EXPORT SALES",
            command=self.export_sales
        ).pack(
            side="left"
        )


        self.report_text = tk.Text(
            self.reports_tab,
            height=20,
            font=(
                "Consolas",
                12
            )
        )


        self.report_text.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=10
        )


        self.report_text.configure(
            state="disabled"
        )


    # ========================================================

    def refresh_reports(self):

        (
            products,
            stock,
            low,
            out,
            inventory_value
        ) = dashboard_data()


        (
            today_sales,
            total_sales,
            transactions,
            profit
        ) = sales_summary()


        purchase_rows = (
            get_purchase_history()
        )


        total_purchase = sum(
            float(row[4])
            for row in purchase_rows
        )


        report = f"""

{COMPANY_NAME}
============================================================

INVENTORY SUMMARY
------------------------------------------------------------
Total Products       : {products}
Total Stock          : {stock}
Low Stock Products   : {low}
Out of Stock         : {out}
Inventory Value      : Rs. {inventory_value:,.2f}

SALES SUMMARY
------------------------------------------------------------
Today's Sales        : Rs. {today_sales:,.2f}
Total Sales          : Rs. {total_sales:,.2f}
Transactions         : {transactions}
Total Profit         : Rs. {profit:,.2f}

PURCHASE SUMMARY
------------------------------------------------------------
Total Purchase Cost  : Rs. {total_purchase:,.2f}

SYSTEM
------------------------------------------------------------
Database             : {DB_NAME}
Invoice Folder       : {os.path.abspath(INVOICE_FOLDER)}

LOGIN
------------------------------------------------------------
Username             : admin
Password             : admin123

"""


        self.report_text.configure(
            state="normal"
        )


        self.report_text.delete(
            "1.0",
            "end"
        )


        self.report_text.insert(
            "1.0",
            report
        )


        self.report_text.configure(
            state="disabled"
        )


    # ========================================================
    # STOCK HISTORY
    # ========================================================

    def build_stock_history_tab(self):

        top = ttk.Frame(
            self.stock_history_tab,
            padding=10
        )

        top.pack(
            fill="x"
        )


        ttk.Button(
            top,
            text="REFRESH HISTORY",
            command=self.refresh_stock_history
        ).pack(
            side="left"
        )


        frame = ttk.Frame(
            self.stock_history_tab,
            padding=10
        )

        frame.pack(
            fill="both",
            expand=True
        )


        columns = (
            "id",
            "product",
            "type",
            "quantity",
            "previous",
            "new",
            "date"
        )


        self.stock_table = ttk.Treeview(
            frame,
            columns=columns,
            show="headings"
        )


        headings = {
            "id": "ID",
            "product": "Product",
            "type": "Movement Type",
            "quantity": "Quantity",
            "previous": "Previous Stock",
            "new": "New Stock",
            "date": "Date & Time"
        }


        for column in columns:

            self.stock_table.heading(
                column,
                text=headings[column]
            )


        self.stock_table.column(
            "id",
            width=60,
            anchor="center"
        )

        self.stock_table.column(
            "product",
            width=230
        )

        self.stock_table.column(
            "type",
            width=150
        )

        self.stock_table.column(
            "quantity",
            width=90,
            anchor="center"
        )

        self.stock_table.column(
            "previous",
            width=110,
            anchor="center"
        )

        self.stock_table.column(
            "new",
            width=90,
            anchor="center"
        )

        self.stock_table.column(
            "date",
            width=180,
            anchor="center"
        )


        scroll = ttk.Scrollbar(
            frame,
            orient="vertical",
            command=self.stock_table.yview
        )


        self.stock_table.configure(
            yscrollcommand=scroll.set
        )


        self.stock_table.pack(
            side="left",
            fill="both",
            expand=True
        )


        scroll.pack(
            side="right",
            fill="y"
        )


    # ========================================================

    def refresh_stock_history(self):

        for item in (
            self.stock_table
            .get_children()
        ):

            self.stock_table.delete(
                item
            )


        for row in get_stock_history():

            self.stock_table.insert(
                "",
                "end",
                values=(
                    row[0],
                    row[1],
                    row[2],
                    row[3],
                    row[4],
                    row[5],
                    row[6]
                )
            )


    # ========================================================
    # REFRESH EVERYTHING
    # ========================================================

    def refresh_all(self):

        self.refresh_products()

        self.refresh_inventory_dashboard()

        self.refresh_sale_products()

        self.refresh_sales()

        self.refresh_purchase_products()

        self.refresh_purchases()

        self.refresh_alerts()

        self.refresh_reports()

        self.refresh_stock_history()


# ============================================================
# START PROGRAM
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    application = InventoryApp(
        root
    )

    root.mainloop()