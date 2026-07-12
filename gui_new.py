import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

# Database Connection
conn = sqlite3.connect("inventory.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS products(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    price REAL,
    quantity INTEGER
)
""")

conn.commit()

# Main Window
root = tk.Tk()
root.title("Inventory Management System")
root.geometry("1000x600")
root.configure(bg="white")
tk.Label(root, text="Product Name").grid(row=0, column=0, padx=10, pady=10)
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Price").grid(row=1, column=0, padx=10, pady=10)
price_entry = tk.Entry(root)
price_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Quantity").grid(row=2, column=0, padx=10, pady=10)
quantity_entry = tk.Entry(root)
quantity_entry.grid(row=2, column=1, padx=10, pady=10)
tk.Label(root, text="Search").grid(row=3, column=0, padx=10, pady=10)

search_entry = tk.Entry(root)
search_entry.grid(row=3, column=1, padx=10, pady=10)
def add_product():
    name = name_entry.get()
    price = price_entry.get()
    quantity = quantity_entry.get()

    # Validation
    if not name or not price or not quantity:
        messagebox.showerror("Error", "Please fill all fields.")
        return

    try:
        price = float(price)
        quantity = int(quantity)
    except ValueError:
        messagebox.showerror("Error", "Price must be a number and Quantity must be an integer.")
        return

    # Paste the cursor.execute() code here
    cursor.execute(
        "INSERT INTO products(name, price, quantity) VALUES (?, ?, ?)",
        (name, price, quantity)
    )

    conn.commit()

    messagebox.showinfo("Success", "Product Added Successfully!")

    name_entry.delete(0, tk.END)
    price_entry.delete(0, tk.END)
    quantity_entry.delete(0, tk.END)

def search_product():
            search = search_entry.get()

            for item in tree.get_children():
                tree.delete(item)

            rows = cursor.execute(
                "SELECT * FROM products WHERE name LIKE ?",
                ('%' + search + '%',)
            ).fetchall()

            for row in rows:
                tree.insert("", tk.END, values=row)

def delete_product():
    selected = tree.selection()

    if not selected:
        messagebox.showerror("Error", "Please select a product.")
        return

    item = tree.item(selected[0])
    product_id = item["values"][0]

    confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this product?")

    if not confirm:
        return
    cursor.execute("DELETE FROM products WHERE id = ?", (product_id,))
    conn.commit()

    tree.delete(selected[0])

    messagebox.showinfo("Success", "Product Deleted Successfully!")

def update_product():
    selected = tree.selection()

    if not selected:
        messagebox.showerror("Error", "Please select a product.")
        return

    item = tree.item(selected[0])
    product_id = item["values"][0]

    name = name_entry.get()
    price = price_entry.get()
    quantity = quantity_entry.get()

    cursor.execute(
        "UPDATE products SET name=?, price=?, quantity=? WHERE id=?",
    (name, price, quantity, product_id)
)

    conn.commit()
    for item in tree.get_children():
        tree.delete(item)

    rows = cursor.execute("SELECT * FROM products").fetchall()

    for row in rows:
        tree.insert("", tk.END, values=row)

    messagebox.showinfo("Success", "Product Updated Successfully!")
def clear_fields():
    name_entry.delete(0, tk.END)
    price_entry.delete(0, tk.END)
    quantity_entry.delete(0, tk.END)
    search_entry.delete(0, tk.END)
    tree.selection_remove(tree.selection())
clear_btn = tk.Button(root, text="Clear", command=clear_fields)
clear_btn.grid(row=3, column=5, padx=10, pady=10)

# Close the database when the window closes
add_btn = tk.Button(root, text="Add Product", command=add_product)
add_btn.grid(row=3, column=1, pady=10)
delete_btn = tk.Button(root, text="Delete", command=delete_product)
delete_btn.grid(row=3, column=3, padx=10, pady=10)
update_btn = tk.Button(root, text="Update", command=update_product)
update_btn.grid(row=3, column=4, padx=10, pady=10)
search_btn = tk.Button(root, text="Search", command=search_product)
search_btn.grid(row=3, column=2, padx=10, pady=10)



columns = ("ID", "Name", "Price", "Quantity")

tree = ttk.Treeview(root, columns=columns, show="headings")

tree.heading("ID", text="ID")
tree.heading("Name", text="Name")
tree.heading("Price", text="Price")
tree.heading("Quantity", text="Quantity")

tree.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

def load_products():
    for item in tree.get_children():
        tree.delete(item)

    rows = cursor.execute("SELECT * FROM products").fetchall()

    for row in rows:
        tree.insert("", tk.END, values=row)

load_products()
def select_product(event):
    selected = tree.selection()

    if not selected:
        return

    item = tree.item(selected[0])
    values = item["values"]

    name_entry.delete(0, tk.END)
    price_entry.delete(0, tk.END)
    quantity_entry.delete(0, tk.END)

    name_entry.insert(0, values[1])
    price_entry.insert(0, values[2])
    quantity_entry.insert(0, values[3])

tree.bind("<<TreeviewSelect>>", select_product)

root.mainloop()