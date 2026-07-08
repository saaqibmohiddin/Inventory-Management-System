import json
import os

FILENAME = "inventory.json"

# Load products from file
def load_products():
    if os.path.exists(FILENAME):
        with open(FILENAME, "r") as file:
            return json.load(file)
    return []

# Save products to file
def save_products(products):
    with open(FILENAME, "w") as file:
        json.dump(products, file, indent=4)

# Add Product
def add_product(products):
    name = input("Enter Product Name: ")
    price = float(input("Enter Product Price: "))
    quantity = int(input("Enter Product Quantity: "))

    product = {
        "name": name,
        "price": price,
        "quantity": quantity
    }

    products.append(product)
    save_products(products)
    print("Product Added Successfully!")

# View Products
def view_products(products):
    if not products:
        print("No Products Available.")
        return

    print("\n----- Product List -----")
    for product in products:
        print(f"Name: {product['name']}")
        print(f"Price: ₹{product['price']}")
        print(f"Quantity: {product['quantity']}")
        print("-" * 25)

# Search Product
def search_product(products):
    name = input("Enter Product Name to Search: ")

    for product in products:
        if product["name"].lower() == name.lower():
            print("Product Found!")
            print(product)
            return

    print("Product Not Found!")

# Update Product
def update_product(products):
    name = input("Enter Product Name to Update: ")

    for product in products:
        if product["name"].lower() == name.lower():
            product["price"] = float(input("Enter New Price: "))
            product["quantity"] = int(input("Enter New Quantity: "))
            save_products(products)
            print("Product Updated Successfully!")
            return

    print("Product Not Found!")

# Delete Product
def delete_product(products):
    name = input("Enter Product Name to Delete: ")

    for product in products:
        if product["name"].lower() == name.lower():
            products.remove(product)
            save_products(products)
            print("Product Deleted Successfully!")
            return

    print("Product Not Found!")

# Sell Product
def sell_product(products):
    name = input("Enter Product Name: ")
    qty = int(input("Enter Quantity Sold: "))

    for product in products:
        if product["name"].lower() == name.lower():
            if product["quantity"] >= qty:
                product["quantity"] -= qty
                save_products(products)
                print("Sale Successful!")
                print("Remaining Stock:", product["quantity"])
            else:
                print("Not Enough Stock!")
            return

    print("Product Not Found!")

# Main Program
def main():
    products = load_products()

    while True:
        print("\n========== Inventory Management System ==========")
        print("1. Add Product")
        print("2. View Products")
        print("3. Search Product")
        print("4. Update Product")
        print("5. Delete Product")
        print("6. Sell Product")
        print("7. Exit")

        choice = input("Enter Your Choice: ")

        if choice == "1":
            add_product(products)
        elif choice == "2":
            view_products(products)
        elif choice == "3":
            search_product(products)
        elif choice == "4":
            update_product(products)
        elif choice == "5":
            delete_product(products)
        elif choice == "6":
            sell_product(products)
        elif choice == "7":
            print("Thank You for Using Inventory Management System!")
            break
        else:
            print("Invalid Choice! Please Try Again.")

main()