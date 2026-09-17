import sqlite3
from datetime import datetime


# ==========================================================
# DATABASE SETTINGS
# ==========================================================

DB_NAME = "inventory.db"


# ==========================================================
# DATABASE CONNECTION
# ==========================================================

def get_connection():
    conn = sqlite3.connect(DB_NAME, timeout=10)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ==========================================================
# HELPER
# ==========================================================

def _table_exists(cursor, table_name):
    cursor.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type = 'table'
        AND name = ?
    """, (table_name,))

    return cursor.fetchone() is not None


def _table_columns(cursor, table_name):

    if not _table_exists(cursor, table_name):
        return []

    cursor.execute(
        f"PRAGMA table_info({table_name})"
    )

    return [row[1] for row in cursor.fetchall()]


def _column_exists(cursor, table_name, column_name):

    return column_name in _table_columns(
        cursor,
        table_name
    )


# ==========================================================
# INITIALIZE DATABASE
# ==========================================================

def initialize_database():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        # ==================================================
        # USERS TABLE
        # ==================================================

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL
            )
        """)

        # Default login
        cursor.execute("""
            INSERT OR IGNORE INTO users
            (username, password)
            VALUES (?, ?)
        """, (
            "admin",
            "admin123"
        ))


        # ==================================================
        # PRODUCTS TABLE
        # ==================================================

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT 'General',
                selling_price REAL NOT NULL DEFAULT 0,
                cost_price REAL NOT NULL DEFAULT 0,
                quantity INTEGER NOT NULL DEFAULT 0
            )
        """)

        # Upgrade old products table

        columns = _table_columns(
            cursor,
            "products"
        )

        if "category" not in columns:

            cursor.execute("""
                ALTER TABLE products
                ADD COLUMN category TEXT
                NOT NULL DEFAULT 'General'
            """)

        if "selling_price" not in columns:

            cursor.execute("""
                ALTER TABLE products
                ADD COLUMN selling_price REAL
                NOT NULL DEFAULT 0
            """)

            # If old database used "price"
            # copy it into selling_price.

            if "price" in columns:

                cursor.execute("""
                    UPDATE products
                    SET selling_price = price
                    WHERE selling_price = 0
                """)

        if "cost_price" not in columns:

            cursor.execute("""
                ALTER TABLE products
                ADD COLUMN cost_price REAL
                NOT NULL DEFAULT 0
            """)

        if "quantity" not in columns:

            cursor.execute("""
                ALTER TABLE products
                ADD COLUMN quantity INTEGER
                NOT NULL DEFAULT 0
            """)


        # ==================================================
        # SALES TABLE
        # ==================================================

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                price REAL NOT NULL,
                total REAL NOT NULL,
                sale_date TEXT NOT NULL,
                cost_price REAL NOT NULL DEFAULT 0,
                customer_name TEXT DEFAULT ''
            )
        """)

        sales_columns = _table_columns(
            cursor,
            "sales"
        )

        if "product_id" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN product_id INTEGER
            """)

        if "product_name" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN product_name TEXT
            """)

        if "quantity" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN quantity INTEGER
                NOT NULL DEFAULT 0
            """)

        if "price" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN price REAL
                NOT NULL DEFAULT 0
            """)

        if "total" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN total REAL
                NOT NULL DEFAULT 0
            """)

        if "sale_date" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN sale_date TEXT
            """)

        if "cost_price" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN cost_price REAL
                NOT NULL DEFAULT 0
            """)

        if "customer_name" not in sales_columns:

            cursor.execute("""
                ALTER TABLE sales
                ADD COLUMN customer_name TEXT
                DEFAULT ''
            """)


        # ==================================================
        # PURCHASES TABLE
        # ==================================================

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

        purchase_columns = _table_columns(
            cursor,
            "purchases"
        )

        if "product_id" not in purchase_columns:

            cursor.execute("""
                ALTER TABLE purchases
                ADD COLUMN product_id INTEGER
            """)

        if "product_name" not in purchase_columns:

            cursor.execute("""
                ALTER TABLE purchases
                ADD COLUMN product_name TEXT
            """)

        if "quantity" not in purchase_columns:

            cursor.execute("""
                ALTER TABLE purchases
                ADD COLUMN quantity INTEGER
                NOT NULL DEFAULT 0
            """)

        if "cost_price" not in purchase_columns:

            cursor.execute("""
                ALTER TABLE purchases
                ADD COLUMN cost_price REAL
                NOT NULL DEFAULT 0
            """)

        if "total_cost" not in purchase_columns:

            cursor.execute("""
                ALTER TABLE purchases
                ADD COLUMN total_cost REAL
                NOT NULL DEFAULT 0
            """)

        if "purchase_date" not in purchase_columns:

            cursor.execute("""
                ALTER TABLE purchases
                ADD COLUMN purchase_date TEXT
            """)


        conn.commit()

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ==========================================================
# LOGIN
# ==========================================================

def check_login(username, password):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT id, username
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


# ==========================================================
# PRODUCTS
# ==========================================================

def get_all_products():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                name,
                category,
                selling_price,
                cost_price,
                quantity
            FROM products
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ==========================================================
# GET PRODUCT BY ID
# ==========================================================

def get_product_by_id(product_id):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                name,
                category,
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


# ==========================================================
# ADD PRODUCT
# ==========================================================

def add_product(
    name,
    category,
    price,
    cost_price,
    quantity
):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            INSERT INTO products
            (
                name,
                category,
                selling_price,
                cost_price,
                quantity
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            name,
            category,
            price,
            cost_price,
            quantity
        ))

        product_id = cursor.lastrowid

        conn.commit()

        return product_id

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ==========================================================
# UPDATE PRODUCT
# ==========================================================

def update_product(
    product_id,
    name,
    category,
    price,
    cost_price,
    quantity
):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            UPDATE products
            SET
                name = ?,
                category = ?,
                selling_price = ?,
                cost_price = ?,
                quantity = ?
            WHERE id = ?
        """, (
            name,
            category,
            price,
            cost_price,
            quantity,
            product_id
        ))

        conn.commit()

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ==========================================================
# DELETE PRODUCT
# ==========================================================
#
# IMPORTANT:
# This function fixes:
#
# 1. FOREIGN KEY constraint failed
# 2. no such column: product_id
#
# It checks the database structure before touching
# sales/purchases.
#
# ==========================================================

def delete_product(product_id):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        # --------------------------------------------------
        # Check product
        # --------------------------------------------------

        cursor.execute("""
            SELECT name
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if product is None:

            raise Exception(
                "Product not found."
            )

        product_name = product[0]


        # --------------------------------------------------
        # SALES
        # --------------------------------------------------

        if _table_exists(
            cursor,
            "sales"
        ):

            sales_columns = _table_columns(
                cursor,
                "sales"
            )

            if "product_id" in sales_columns:

                # Check if product_id is NOT NULL
                cursor.execute(
                    "PRAGMA table_info(sales)"
                )

                table_info = cursor.fetchall()

                product_id_not_null = False

                for column in table_info:

                    if column[1] == "product_id":

                        product_id_not_null = (
                            column[3] == 1
                        )

                        break

                if product_id_not_null:

                    # Old database does not allow NULL.
                    # Delete dependent sales records.

                    cursor.execute("""
                        DELETE FROM sales
                        WHERE product_id = ?
                    """, (
                        product_id,
                    ))

                else:

                    # Keep sales history and remove
                    # product reference.

                    cursor.execute("""
                        UPDATE sales
                        SET product_id = NULL
                        WHERE product_id = ?
                    """, (
                        product_id,
                    ))


        # --------------------------------------------------
        # PURCHASES
        # --------------------------------------------------

        if _table_exists(
            cursor,
            "purchases"
        ):

            purchase_columns = _table_columns(
                cursor,
                "purchases"
            )

            if "product_id" in purchase_columns:

                cursor.execute(
                    "PRAGMA table_info(purchases)"
                )

                table_info = cursor.fetchall()

                product_id_not_null = False

                for column in table_info:

                    if column[1] == "product_id":

                        product_id_not_null = (
                            column[3] == 1
                        )

                        break

                if product_id_not_null:

                    cursor.execute("""
                        DELETE FROM purchases
                        WHERE product_id = ?
                    """, (
                        product_id,
                    ))

                else:

                    cursor.execute("""
                        UPDATE purchases
                        SET product_id = NULL
                        WHERE product_id = ?
                    """, (
                        product_id,
                    ))


        # --------------------------------------------------
        # Check OTHER foreign-key tables
        # --------------------------------------------------

        cursor.execute("""
            SELECT name
            FROM sqlite_master
            WHERE type = 'table'
        """)

        all_tables = [
            row[0]
            for row in cursor.fetchall()
        ]

        for table_name in all_tables:

            if table_name in (
                "products",
                "sales",
                "purchases",
                "users",
                "sqlite_sequence"
            ):
                continue

            columns = _table_columns(
                cursor,
                table_name
            )

            if "product_id" not in columns:
                continue

            # Check foreign keys
            cursor.execute(
                f"PRAGMA foreign_key_list({table_name})"
            )

            foreign_keys = cursor.fetchall()

            references_products = False

            for fk in foreign_keys:

                # fk[2] = referenced table
                # fk[3] = local column

                if (
                    fk[2] == "products"
                    and fk[3] == "product_id"
                ):

                    references_products = True
                    break

            if not references_products:
                continue


            # Check whether product_id allows NULL

            cursor.execute(
                f"PRAGMA table_info({table_name})"
            )

            table_info = cursor.fetchall()

            product_id_not_null = False

            for column in table_info:

                if column[1] == "product_id":

                    product_id_not_null = (
                        column[3] == 1
                    )

                    break


            if product_id_not_null:

                cursor.execute(
                    f"""
                    DELETE FROM {table_name}
                    WHERE product_id = ?
                    """,
                    (
                        product_id,
                    )
                )

            else:

                cursor.execute(
                    f"""
                    UPDATE {table_name}
                    SET product_id = NULL
                    WHERE product_id = ?
                    """,
                    (
                        product_id,
                    )
                )


        # --------------------------------------------------
        # Finally delete product
        # --------------------------------------------------

        cursor.execute("""
            DELETE FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        if cursor.rowcount == 0:

            raise Exception(
                "Product could not be deleted."
            )

        conn.commit()

        return True


    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()


# ==========================================================
# SEARCH PRODUCTS
# ==========================================================

def search_products(search_text):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        text = f"%{search_text}%"

        cursor.execute("""
            SELECT
                id,
                name,
                category,
                selling_price,
                cost_price,
                quantity
            FROM products
            WHERE
                name LIKE ?
                OR category LIKE ?
            ORDER BY id DESC
        """, (
            text,
            text
        ))

        return cursor.fetchall()

    finally:

        conn.close()


# ==========================================================
# TOTAL PRODUCTS
# ==========================================================

def get_total_products():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COUNT(*)
            FROM products
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# TOTAL STOCK
# ==========================================================

def get_total_stock():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COALESCE(
                SUM(quantity),
                0
            )
            FROM products
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# LOW STOCK
# ==========================================================

def get_low_stock_count():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COUNT(*)
            FROM products
            WHERE quantity BETWEEN 1 AND 5
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# OUT OF STOCK
# ==========================================================

def get_out_of_stock_count():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COUNT(*)
            FROM products
            WHERE quantity = 0
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# INVENTORY VALUE
# ==========================================================

def get_total_inventory_value():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COALESCE(
                SUM(
                    selling_price * quantity
                ),
                0
            )
            FROM products
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# SELL PRODUCT
# ==========================================================

def sell_product(
    product_id,
    sell_quantity,
    customer_name=""
):

    conn = get_connection()
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

        if product is None:

            return False, "Product not found."

        (
            product_name,
            selling_price,
            cost_price,
            current_quantity
        ) = product


        if sell_quantity <= 0:

            return (
                False,
                "Sale quantity must be greater than 0."
            )


        if sell_quantity > current_quantity:

            return (
                False,
                f"Not enough stock. "
                f"Available stock: {current_quantity}"
            )


        total = (
            float(selling_price)
            *
            int(sell_quantity)
        )

        new_quantity = (
            current_quantity
            -
            sell_quantity
        )

        sale_date = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )


        # Update stock

        cursor.execute("""
            UPDATE products
            SET quantity = ?
            WHERE id = ?
        """, (
            new_quantity,
            product_id
        ))


        # Insert sale

        cursor.execute("""
            INSERT INTO sales
            (
                product_id,
                product_name,
                quantity,
                price,
                total,
                sale_date,
                cost_price,
                customer_name
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            product_id,
            product_name,
            sell_quantity,
            selling_price,
            total,
            sale_date,
            cost_price,
            customer_name
        ))


        sale_id = cursor.lastrowid

        conn.commit()

        return True, (
            total,
            sale_id
        )


    except Exception as error:

        conn.rollback()

        return False, str(error)

    finally:

        conn.close()


# ==========================================================
# SALES HISTORY
# ==========================================================

def get_sales_history():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                product_name,
                quantity,
                price,
                total,
                sale_date,
                cost_price,
                customer_name
            FROM sales
            ORDER BY id DESC
        """)

        return cursor.fetchall()

    finally:

        conn.close()


# ==========================================================
# GET SALE BY ID
# ==========================================================

def get_sale_by_id(sale_id):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                id,
                product_id,
                product_name,
                quantity,
                price,
                total,
                sale_date,
                cost_price,
                customer_name
            FROM sales
            WHERE id = ?
        """, (
            sale_id,
        ))

        return cursor.fetchone()

    finally:

        conn.close()


# ==========================================================
# TOTAL SALES
# ==========================================================

def get_total_sales():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COALESCE(
                SUM(total),
                0
            )
            FROM sales
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# TODAY'S SALES
# ==========================================================

def get_today_sales():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        cursor.execute("""
            SELECT COALESCE(
                SUM(total),
                0
            )
            FROM sales
            WHERE sale_date LIKE ?
        """, (
            f"{today}%",
        ))

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# TOTAL SALE COUNT
# ==========================================================

def get_total_sale_count():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COUNT(*)
            FROM sales
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# TOTAL PROFIT
# ==========================================================

def get_total_profit():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COALESCE(
                SUM(
                    total
                    -
                    (quantity * cost_price)
                ),
                0
            )
            FROM sales
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# RESTOCK PRODUCT
# ==========================================================

def restock_product(
    product_id,
    restock_quantity,
    purchase_cost
):

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT
                name,
                cost_price,
                quantity
            FROM products
            WHERE id = ?
        """, (
            product_id,
        ))

        product = cursor.fetchone()

        if product is None:

            return False, "Product not found."


        (
            product_name,
            old_cost,
            old_quantity
        ) = product


        if restock_quantity <= 0:

            return (
                False,
                "Restock quantity must be greater than 0."
            )


        if purchase_cost <= 0:

            return (
                False,
                "Purchase cost must be greater than 0."
            )


        old_cost = float(old_cost)
        old_quantity = int(old_quantity)
        purchase_cost = float(purchase_cost)


        # Weighted average cost

        existing_value = (
            old_quantity
            *
            old_cost
        )

        new_purchase_value = (
            restock_quantity
            *
            purchase_cost
        )

        total_quantity = (
            old_quantity
            +
            restock_quantity
        )


        if total_quantity > 0:

            new_average_cost = (
                existing_value
                +
                new_purchase_value
            ) / total_quantity

        else:

            new_average_cost = purchase_cost


        new_quantity = (
            old_quantity
            +
            restock_quantity
        )


        total_cost = (
            restock_quantity
            *
            purchase_cost
        )


        purchase_date = datetime.now().strftime(
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
            new_average_cost,
            product_id
        ))


        # Add purchase history

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
            restock_quantity,
            purchase_cost,
            total_cost,
            purchase_date
        ))


        conn.commit()

        return True, total_cost


    except Exception as error:

        conn.rollback()

        return False, str(error)

    finally:

        conn.close()


# ==========================================================
# PURCHASE HISTORY
# ==========================================================

def get_purchase_history():

    conn = get_connection()
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


# ==========================================================
# TODAY'S PURCHASES
# ==========================================================

def get_today_purchases():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        today = datetime.now().strftime(
            "%Y-%m-%d"
        )

        cursor.execute("""
            SELECT COALESCE(
                SUM(total_cost),
                0
            )
            FROM purchases
            WHERE purchase_date LIKE ?
        """, (
            f"{today}%",
        ))

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# TOTAL PURCHASES
# ==========================================================

def get_total_purchases():

    conn = get_connection()
    cursor = conn.cursor()

    try:

        cursor.execute("""
            SELECT COALESCE(
                SUM(total_cost),
                0
            )
            FROM purchases
        """)

        return cursor.fetchone()[0]

    finally:

        conn.close()


# ==========================================================
# START DATABASE
# ==========================================================

initialize_database()