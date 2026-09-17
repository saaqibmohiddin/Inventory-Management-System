# Inventory Management System

## Project Description

The Inventory Management System is a desktop application developed using Python, Tkinter, and SQLite. It helps users manage product information such as product name, price, and quantity.

## Features

- Add new products
- View all products
- Search products
- Update product information
- Delete products
- Clear input fields
- Validate price and quantity
- Prevent negative prices
- Prevent negative quantities
- Export product data to Excel

## Technologies Used

- Python
- Tkinter
- SQLite
- openpyxl

## Project Structure

```text
Inventory Management System
│
├── database.py
├── gui.py
├── gui_new.py
├── inventory_management.py
├── inventory.db
├── inventory.json
├── products.pdf
├── products.xlsx
└── README.md
```
## How to Run

1. Install Python.
2. Install the required library:
   `pip install openpyxl`
3. Open the project folder in VS Code.
4. Run the application:
   `python gui_new.py`