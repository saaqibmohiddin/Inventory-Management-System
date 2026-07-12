import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Table
from reportlab.lib import colors
from reportlab.platypus import TableStyle
# -----------------------------
# Database Connection
# -----------------------------
conn = sqlite3.connect("inventory.db")
cursor = conn.cursor()

# Create table if it doesn't exist
cursor.execute("""
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    price REAL NOT NULL,
    quantity INTEGER NOT NULL
)
""")
conn.commit()

# -----------------------------
# Main Window
# -----------------------------
root = tk.Tk()
root.title("Inventory Management System")
root.geometry("700x500")

# -----------------------------
# Labels
# -----------------------------
tk.Label(root, text="Product Name").grid(row=0, column=0, padx=10, pady=10)
tk.Label(root, text="Price").grid(row=1, column=0, padx=10, pady=10)
tk.Label(root, text="Quantity").grid(row=2, column=0, padx=10, pady=10)
tk.Label(root, text="Search").grid(row=0, column=2, padx=10)
total_label = tk.Label(root, text="Total Products: 0")
total_label.grid(row=4, column=2, columnspan=2, pady=10)
# -----------------------------
# Entry Boxes
# -----------------------------
name_entry = tk.Entry(root)
price_entry = tk.Entry(root)
quantity_entry = tk.Entry(root)
search_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=10, pady=10)
price_entry.grid(row=1, column=1, padx=10, pady=10)
quantity_entry.grid(row=2, column=1, padx=10, pady=10)
search_entry.grid(row=0, column=3, padx=10, pady=10)
# -----------------------------
# Functions
# -----------------------------
def search_product():
    search = search_entry.get()

    cursor.execute(
        "SELECT * FROM products WHERE name LIKE ?",
        ('%' + search + '%',)
    )

    rows = cursor.fetchall()

    listbox.delete(*listbox.get_children())

    for row in rows:
        listbox.insert("", tk.END, values=row)
def sort_products():
    rows = cursor.execute(
        "SELECT * FROM products ORDER BY name ASC"
    ).fetchall()

    listbox.delete(*listbox.get_children())

    for row in rows:
        listbox.insert("", tk.END, values=row)
    total = len(rows)
    total_label.config(text=f"Total Products:{total}")
def export_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    ws.append(["ID", "Name", "Price", "Quantity"])

    rows = cursor.execute("SELECT * FROM products").fetchall()

    for row in rows:
        ws.append(row)

    wb.save("products.xlsx")
    messagebox.showinfo("Success", "Products exported to products.xlsx")
    def view_products():
     rows = cursor.execute("SELECT * FROM products").fetchall()

    # Clear old data
    for item in listbox.get_children():
        listbox.delete(item)

    # Insert fresh data
    for row in rows:
        listbox.insert("", tk.END, values=row)

    total = len(rows)
    total_label.config(text=f"Total Products: {total}")
def export_to_pdf():

def export_to_pdf():
    doc = SimpleDocTemplate("products.pdf")

    data = [["ID", "Name", "Price", "Quantity"]]

    rows = cursor.execute("SELECT * FROM products").fetchall()

    for row in rows:
        data.append(list(row))

    table = Table(data)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
    ]))

    doc.build([table])

    messagebox.showinfo("Success", "Products exported to products.pdf")


def print_bill():
    export_to_pdf()
    messagebox.showinfo("Bill", "Print Bill")









def view_products():
    rows = cursor.execute("SELECT * FROM products").fetchall()

    # Clear old data
    for item in listbox.get_children():
        listbox.delete(item)

    # Insert fresh data
    for row in rows:
        listbox.insert("", tk.END, values=row)
    total = len(rows)
    total_label.config(text=f"Total Products: {total}")


def add_product():
    """Add a new product."""
    name = name_entry.get().strip()
    price = price_entry.get().strip()
    quantity = quantity_entry.get().strip()

    if name == "" or price == "" or quantity == "":
        messagebox.showerror("Error", "Please fill all fields")
        return

    try:
        price = float(price)
        quantity = int(quantity)
    except ValueError:
        messagebox.showerror("Error", "Price must be a number and Quantity must be an integer.")
        return

    cursor.execute(
        "INSERT INTO products (name, price, quantity) VALUES (?, ?, ?)",
        (name, price, quantity)
    )
    conn.commit()

    messagebox.showinfo("Success", "Product Added Successfully")

    # Clear entry boxes
    name_entry.delete(0, tk.END)
    price_entry.delete(0, tk.END)
    quantity_entry.delete(0, tk.END)

    view_products()


def delete_product():
    """Delete the selected product."""
    selected = listbox.focus()

    if not selected:
        messagebox.showerror("Error", "Please select a product to delete.")
        return

    item = listbox.item(selected)
    product_id = item["values"][0]

    cursor.execute("DELETE FROM products WHERE id=?", (product_id,))
    conn.commit()

    messagebox.showinfo("Success", "Product Deleted Successfully")
    view_products()


def update_product():
    selected = listbox.focus()

    if not selected:
        messagebox.showerror("Error", "Please select a product")
        return

    item = listbox.item(selected)
    product_id = item["values"][0]

    name = name_entry.get()
    price = price_entry.get()
    quantity = quantity_entry.get()

    if name == "" or price == "" or quantity == "":
        messagebox.showerror("Error", "Please fill all fields")
        return

    cursor.execute(
        "UPDATE products SET name=?, price=?, quantity=? WHERE id=?",
        (name, price, quantity, product_id)
    )
    conn.commit()
    view_products()
    messagebox.showinfo("Success", "Product Updated Successfully")

def clear_fields():
    name_entry.delete(0,tk.END)
    price_entry.delete(0,tk.END)
    quantity_entry.delete(0,tk.END)

# -----------------------------
# Buttons
# -----------------------------
add_btn = tk.Button(root, text="Add Product", command=add_product)
add_btn.grid(row=3, column=0, pady=10)

delete_btn = tk.Button(root, text="Delete Product", command=delete_product)
delete_btn.grid(row=3, column=1, pady=10)
update_btn = tk.Button(root, text="Update Product", command=update_product)
update_btn.grid(row=3, column=2, pady=10)
clear_btn = tk.Button(root, text="Clear", command=clear_fields)
clear_btn.grid(row=3, column=3, pady=10)
search_btn = tk.Button(root, text="Search", command=search_product)
search_btn.grid(row=1, column=3, padx=10, pady=10)
sort_btn = tk.Button(root, text="Sort", command=sort_products)
sort_btn.grid(row=2, column=4, padx=10, pady=10)
export_btn = tk.Button(root, text="Export Excel", command=export_to_excel)
export_btn.grid(row=3, column=5, padx=10, pady=10)

pdf_btn = tk.Button(root, text="Export PDF", command=export_to_pdf)
pdf_btn.grid(row=3, column=6, padx=10, pady=10)
# -----------------------------
# Product Table
# -----------------------------
listbox = ttk.Treeview(
    root,
    columns=("ID", "Name", "Price", "Quantity"),
    show="headings",
    height=12
)

listbox.heading("ID", text="ID")
listbox.heading("Name", text="Name")
listbox.heading("Price", text="Price")
listbox.heading("Quantity", text="Quantity")

listbox.column("ID", width=60, anchor="center")
listbox.column("Name", width=220)
listbox.column("Price", width=120, anchor="center")
listbox.column("Quantity", width=120, anchor="center")

listbox.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# -----------------------------
# Load Existing Products
# -----------------------------
view_products()

# -----------------------------
# Run Application
# -----------------------------
root.mainloop()

# Close database connection
conn.close()